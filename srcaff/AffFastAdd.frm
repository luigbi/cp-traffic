VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFastAdd 
   Caption         =   "frmFastAdd"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AffFastAdd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame frcFromFile 
      Height          =   7695
      Left            =   6720
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton cmcBrowse 
         Caption         =   "Browse"
         Height          =   300
         Left            =   4485
         TabIndex        =   20
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox txtFile 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   3600
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFastAddFile 
         Height          =   6705
         Left            =   60
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   720
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   11827
         _Version        =   393216
         Rows            =   3
         Cols            =   11
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
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
         _Band(0).Cols   =   11
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lblFastAddStatus 
         Height          =   375
         Left            =   5760
         TabIndex        =   47
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lbcFile 
         Caption         =   "File"
         Height          =   315
         Left            =   60
         TabIndex        =   21
         Top             =   255
         Width           =   540
      End
   End
   Begin VB.Frame frcManualHeader 
      Height          =   3855
      Left            =   120
      TabIndex        =   23
      Top             =   360
      Width           =   9375
      Begin VB.OptionButton rbcGetFrom 
         Caption         =   "List of All Stations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   7920
         TabIndex        =   40
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton rbcGetFrom 
         Caption         =   "Agreements Now on"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   7920
         TabIndex        =   39
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton rbcGetFrom 
         Caption         =   "External File..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   7920
         TabIndex        =   38
         Top             =   1605
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtStartDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1710
         TabIndex        =   37
         Top             =   2085
         Width           =   1455
      End
      Begin VB.TextBox txtEndDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   36
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ComboBox cboStationMarket 
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
         ItemData        =   "AffFastAdd.frx":08CA
         Left            =   1710
         List            =   "AffFastAdd.frx":08CC
         Sorted          =   -1  'True
         TabIndex        =   35
         Top             =   2580
         Width           =   2760
      End
      Begin VB.ComboBox cboVehicle 
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
         Left            =   4560
         Sorted          =   -1  'True
         TabIndex        =   34
         Top             =   2985
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtBrowse 
         BackColor       =   &H00C0FFFF&
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
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2985
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.ListBox lbcCreateVehicle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         ItemData        =   "AffFastAdd.frx":08CE
         Left            =   1710
         List            =   "AffFastAdd.frx":08D0
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   240
         Width           =   5745
      End
      Begin VB.Frame frcMulticast 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   3435
         Width           =   8295
         Begin VB.OptionButton rbcMulticast 
            Caption         =   "All possible Multicast Stations"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3870
            TabIndex        =   30
            Top             =   -75
            Width           =   2700
         End
         Begin VB.OptionButton rbcMulticast 
            Caption         =   "All prior Multicast Stations"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1590
            TabIndex        =   29
            Top             =   -75
            Width           =   2280
         End
         Begin VB.Label lacMulticast 
            Caption         =   "Generate for"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   2235
         End
      End
      Begin VB.CheckBox ckcDelivery 
         Caption         =   "Copy Delivery Service to Stations/Vehicles w/o Agreements"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7920
         TabIndex        =   27
         Top             =   1920
         Visible         =   0   'False
         Width           =   4680
      End
      Begin VB.ComboBox cboGetStationsFrom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "AffFastAdd.frx":08D2
         Left            =   1710
         List            =   "AffFastAdd.frx":08DF
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2985
         Width           =   2760
      End
      Begin VB.ComboBox cboCopyDeliveryService 
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
         ItemData        =   "AffFastAdd.frx":092F
         Left            =   5880
         List            =   "AffFastAdd.frx":0931
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2580
         Width           =   3375
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   4560
         TabIndex        =   26
         Top             =   2985
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label lblCreate 
         Caption         =   "Create (one or more) Agreements for"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblStart 
         Caption         =   "Start Date"
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
         Left            =   120
         TabIndex        =   45
         Top             =   2205
         Width           =   975
      End
      Begin VB.Label lblEnd 
         Caption         =   "End Date (optional)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   44
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblPledge 
         Caption         =   "Get Pledges from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   2595
         Width           =   1695
      End
      Begin VB.Label lblStations 
         Caption         =   "Get Stations from"
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
         Left            =   120
         TabIndex        =   42
         Top             =   2985
         Width           =   1695
      End
      Begin VB.Label lblCopy 
         Alignment       =   2  'Center
         Caption         =   "Delivery Service"
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
         Left            =   4560
         TabIndex        =   41
         Top             =   2640
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   1800
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frcManual 
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   9375
      Begin VB.Frame frcAll 
         Caption         =   "Active as-of"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4080
         TabIndex        =   10
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
         Begin VB.PictureBox pbcIncludeToExclude 
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
            Left            =   960
            Picture         =   "AffFastAdd.frx":0933
            ScaleHeight     =   180
            ScaleWidth      =   120
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   810
            Width           =   120
         End
         Begin VB.CommandButton cmdAll 
            Caption         =   "Add All     "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   1020
         End
         Begin VB.TextBox txtActiveDate 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.CommandButton cmdMoveLeft 
         Height          =   375
         Left            =   4320
         Picture         =   "AffFastAdd.frx":0A0D
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   360
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdMoveRight 
         Height          =   375
         Left            =   4320
         Picture         =   "AffFastAdd.frx":0B83
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdExclude 
         Height          =   3345
         Left            =   60
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   420
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   5900
         _Version        =   393216
         Rows            =   3
         Cols            =   9
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdInclude 
         Height          =   3345
         Left            =   5400
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   420
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   5900
         _Version        =   393216
         Rows            =   3
         Cols            =   9
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
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
      Begin VB.Label lblInclude 
         Caption         =   "Stations to Include"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5400
         TabIndex        =   17
         Top             =   200
         Width           =   3915
      End
      Begin VB.Label lblExclude 
         Caption         =   "Stations to Exclude"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         TabIndex        =   16
         Top             =   200
         Width           =   3915
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar2 
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   8145
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9000
      Top             =   8160
   End
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   8955
      Top             =   7980
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4530
      Top             =   8085
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGen 
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2820
      TabIndex        =   0
      Top             =   8145
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5250
      TabIndex        =   1
      Top             =   8145
      Width           =   1455
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8400
      Top             =   8040
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8685
      FormDesignWidth =   9600
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSvExclude 
      Height          =   675
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7950
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1191
      _Version        =   393216
      Rows            =   3
      Cols            =   10
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
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSvInclude 
      Height          =   630
      Left            =   6630
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7980
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1111
      _Version        =   393216
      Rows            =   3
      Cols            =   10
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
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7935
      Left            =   120
      TabIndex        =   5
      Top             =   45
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   13996
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Manual"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "From File"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFastAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmFastAdd - Allow users to mass clone agreements based off of an
'*               existing agreement
'*
'*  Created March, 2004 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc. 2004
'******************************************************

Option Explicit
Option Compare Text

'support for the type ahead
Private imVehicleInChg As Integer
Private imCreateVehInChg As Integer
Private imStationMarketInChg As Integer
Private imCreateVehBSMode As Integer
Private imStationMarketBSMode As Integer
Private imVehicleBSMode As Integer

Private tmOverlapInfo() As AGMNTOVERLAPINFO
Private tmDat() As DAT
Private tmFastAddAttCount() As FASTADDATTCOUNT

'misc. vars
Private bmShowDates As Boolean
Private smAdminEMail As String
Private imVefCode As Integer
Private lmBaseAttCode As Long
Private imBaseShttCode As Integer
Private imAddShttCode As Integer
Private smPledgeType As String
Private lmAttCode As Long
Private smCurDate As String
Private smCurTime As String
Private imWebType As Integer
Private imUnivisionType As Integer
'Private slWebEMail As String
'Private slWebPW As String
Private smStationName As String
Private imStationHasTimeZoneDefined As Integer
Private imOKToConvertTimeZones As Integer
Private smPledgeByEvent As String
'Private imAgreeType As Integer
Private bmTrmntAgrmnt As Boolean

'ADO vars
Private attBase_rst As ADODB.Recordset
Private shtt_rst As ADODB.Recordset
Private dat_rst As ADODB.Recordset
Private adrst As ADODB.Recordset
Private rst_Pet As ADODB.Recordset
Private rst_Gsf As ADODB.Recordset
Private rst_Lst As ADODB.Recordset

Private imLastExcludeColSorted As Integer
Private imLastExcludeSort As Integer
Private lmLastExcludeClickedRow As Long

Private imLastIncludeColSorted As Integer
Private imLastIncludeSort As Integer
Private lmLastIncludeClickedRow As Long

Private Const CALLLETTERSINDEX = 0
Private Const MARKETINDEX = 1
Private Const FORMATINDEX = 2
Private Const OWNERINDEX = 3
Private Const ZONEINDEX = 4
Private Const DATERANGEINDEX = 5
Private Const RANKINDEX = 6
Private Const SHTTCODEINDEX = 7
Private Const SORTINDEX = 8

Private Const FILELINE = 0
Private Const FILECALLLETTERS = 1
Private Const FILEVEHICLE = 2
Private Const FILESTARTDATE = 3
Private Const FILEPLEDGESFROM = 4
Private Const FILEDLVYSVCMODE = 5
Private Const FILEGENFOR = 6
Private Const FILESTATUS = 7
Private Const FILEVEFCODE = 8
Private Const FILECALLLETTERSCODE = 9
Private Const FILEPLEDGESFROMCODE = 10

'Private Const SELECTEDINDEX = 9  'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - removed Selected index, using Native grid row, RowSel properties
Dim bmCreateChanged As Boolean
Dim imCurrentStationList As Integer
Dim imClearSearchTime As Integer
Dim imFileOkay As Integer

Private Function mAdjOverlapAgmnts(llOnAir As Long, llOffAir As Long, llDropDate As Long) As Integer
    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim llEndDate As Long
    Dim llTestEndDate As Long
    Dim slDrop As String
    Dim ilAgreeType As Integer
    Dim ilExported As Integer
    ReDim tlCpttArray(0 To 0) As CPTTARRAY
    Dim rst_Cptt As ADODB.Recordset
    Dim rst_Temp As ADODB.Recordset
    Dim ilUpper As Integer
    Dim slEndDate As String
    Dim ilRet As Integer
    Dim ilShtt As Integer
    Dim slCallLetters As String
    Dim slVehicleName As String
    
    ReDim tmFastAddAttCount(0 To 0) As FASTADDATTCOUNT
    'manual = 0, web site = 1, Univision = 2
    If imUnivisionType = True Then
        ilAgreeType = 2
    ElseIf imWebType = True Then
        ilAgreeType = 1
    Else
        ilAgreeType = 0
    End If
    
    If UBound(tmOverlapInfo) <= LBound(tmOverlapInfo) Then
        mAdjOverlapAgmnts = True
        Exit Function
    End If
    
    On Error GoTo ErrHand
    
    If llDropDate < llOffAir Then
        llEndDate = llDropDate
    Else
        llEndDate = llOffAir
    End If
    For ilLoop = LBound(tmOverlapInfo) To UBound(tmOverlapInfo) - 1 Step 1
        'Determine if Agreement should be Delected or terminated
        If tmOverlapInfo(ilLoop).lDropDate < tmOverlapInfo(ilLoop).lOffAirDate Then
            llTestEndDate = tmOverlapInfo(ilLoop).lDropDate
        Else
            llTestEndDate = tmOverlapInfo(ilLoop).lOffAirDate
        End If

        If (llOnAir <= tmOverlapInfo(ilLoop).lOnAirDate) And (tmOverlapInfo(ilLoop).lOnAirDate <= llEndDate) Then
            slCallLetters = gGetCallLettersByAttCode(tmOverlapInfo(ilLoop).lAttCode)
            slVehicleName = gGetVehNameByVefCode(gGetVehCodeFromAttCode(CStr(tmOverlapInfo(ilLoop).lAttCode)))
            mSaveAttCount tmOverlapInfo(ilLoop).iShfCode
            'Delete old agreement as it is starts after new agreement
            SQLQuery = "SELECT cpttStartDate, cpttCode FROM Cptt WHERE (cpttAtfCode = " & tmOverlapInfo(ilLoop).lAttCode & ")"
            Set rst_Cptt = gSQLSelectCall(SQLQuery)
            ilUpper = 0
            While Not rst_Cptt.EOF
                tlCpttArray(ilUpper).lCpttCode = rst_Cptt!cpttCode
                tlCpttArray(ilUpper).sCpttStartDate = rst_Cptt!CpttStartDate
                ilUpper = ilUpper + 1
                ReDim Preserve tlCpttArray(0 To ilUpper)
                rst_Cptt.MoveNext
            Wend
            
            For ilIdx = 0 To ilUpper - 1 Step 1
                ilExported = gCheckIfSpotsHaveBeenExported(imVefCode, tlCpttArray(ilIdx).sCpttStartDate, ilAgreeType)
                If (igChangedNewErased = 1 Or igChangedNewErased = 2) And (ilExported = False) Then
                    SQLQuery = "DELETE FROM Cptt WHERE (cpttCode = " & tlCpttArray(ilIdx).lCpttCode & ")"
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "FastAdd-mAdjOverlapAgmnts"
                        mAdjOverlapAgmnts = False
                        Exit Function
                    End If
                    gLogMsg "Deleting CPTT2: " & slCallLetters & " running on: " & slVehicleName & " for the week of: " & Format$(tlCpttArray(ilIdx).sCpttStartDate, "mm/dd/yyyy"), "FastAddVerbose.Txt", False
                    slEndDate = DateAdd("d", 6, tlCpttArray(ilIdx).sCpttStartDate)
                    SQLQuery = "DELETE FROM Ast WHERE (astAtfCode = " & tmOverlapInfo(ilLoop).lAttCode
                    SQLQuery = SQLQuery & " AND astFeedDate >= '" & Format$(tlCpttArray(ilIdx).sCpttStartDate, sgSQLDateForm) & "'"
                    SQLQuery = SQLQuery & " AND astFeedDate <= '" & Format$(slEndDate, sgSQLDateForm) & "')"
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "FastAdd-mAdjOverlapAgmnts"
                        mAdjOverlapAgmnts = False
                        Exit Function
                    End If
                    gLogMsg "Deleting AST2: " & slCallLetters & " running on: " & slVehicleName & " for the period: " & Format$(tlCpttArray(ilIdx).sCpttStartDate, "mm/dd/yyyy") & " - " & Format$(slEndDate, "mm/dd/yyyy"), "FastAddVerbose.Txt", False
                    If (ilAgreeType = 1) Or (ilAgreeType = 2) Then
                        ilRet = gAlertAdd("R", "S", imVefCode, tlCpttArray(ilIdx).sCpttStartDate)
                    End If
                End If
            Next ilIdx
            
            'Delete old agreement as it is starts after new agreement
            ' JD 12-18-2006 Added new function to properly remove an agreement.
            'D.S. 4/27/18 added if statement and uncommented nested if for gDeleteAgreement
            If bmTrmntAgrmnt Then
                If Not gDeleteAgreement(tmOverlapInfo(ilLoop).lAttCode, "FastAddSummary.Txt") Then
                    gLogMsg "FAIL: mAdjOverlapAgmnts - Unable to delete att code " & tmOverlapInfo(ilLoop).lAttCode, "AffErrorLog.Txt", False
                End If
            End If
'            cnn.BeginTrans
'            SQLQuery = "DELETE FROM dat WHERE (datAtfCode = " & tmOverlapInfo(ilLoop).lattCode & ")"
'            cnn.Execute SQLQuery, rdExecDirect
'            SQLQuery = "DELETE FROM Att WHERE (AttCode = " & tmOverlapInfo(ilLoop).lattCode & ")"
'            cnn.Execute SQLQuery, rdExecDirect
'            cnn.CommitTrans
        ElseIf (tmOverlapInfo(ilLoop).lOnAirDate < llOnAir) And (llTestEndDate >= llOnAir) Then
            mSaveAttCount tmOverlapInfo(ilLoop).iShfCode
            'Terminate the agreement on llOnAir minus 1 as it starts prior to new agreement
            slDrop = Format$(llOnAir, "m/d/yy")
'            cnn.BeginTrans

            SQLQuery = "SELECT cpttStartDate, cpttCode FROM cptt WHERE (cpttAtfCode = " & tmOverlapInfo(ilLoop).lAttCode & " And "
            SQLQuery = SQLQuery & "cpttStartDate >= '" & Format$(slDrop, sgSQLDateForm) & "'" & ")"
            Set rst_Cptt = gSQLSelectCall(SQLQuery)
            ilUpper = 0
            While Not rst_Cptt.EOF
                tlCpttArray(ilUpper).lCpttCode = rst_Cptt!cpttCode
                tlCpttArray(ilUpper).sCpttStartDate = rst_Cptt!CpttStartDate
                ilUpper = ilUpper + 1
                ReDim Preserve tlCpttArray(0 To ilUpper)
                rst_Cptt.MoveNext
            Wend
                
            For ilIdx = 0 To ilUpper - 1 Step 1
                ilExported = gCheckIfSpotsHaveBeenExported(imVefCode, tlCpttArray(ilIdx).sCpttStartDate, ilAgreeType)
                'igChangedNewErased values  1 = changed, 2 = new, 3 = erased
                'If they are changing an agreement and it's already been exported then don't delete the CPTTs
                If (igChangedNewErased = 1 Or igChangedNewErased = 2) And (ilExported = False) Then
                    SQLQuery = "DELETE FROM Cptt WHERE (cpttCode = " & tlCpttArray(ilIdx).lCpttCode & ")"
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "FastAdd-mAdjOverlapAgmnts"
                        mAdjOverlapAgmnts = False
                        Exit Function
                    End If
                    gLogMsg "Deleting CPTT3: " & slCallLetters & " running on: " & slVehicleName & " for the week of: " & Format$(tlCpttArray(ilIdx).sCpttStartDate, "mm/dd/yyyy"), "FastAddVerbose.Txt", False
                    slEndDate = DateAdd("d", 6, tlCpttArray(ilIdx).sCpttStartDate)
                    SQLQuery = "DELETE FROM Ast WHERE (astAtfCode = " & tmOverlapInfo(ilLoop).lAttCode
                    SQLQuery = SQLQuery & " AND astFeedDate >= '" & Format$(tlCpttArray(ilIdx).sCpttStartDate, sgSQLDateForm) & "'"
                    SQLQuery = SQLQuery & " AND astFeedDate <= '" & Format$(slEndDate, sgSQLDateForm) & "')"
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "FastAdd-mAdjOverlapAgmnts"
                        mAdjOverlapAgmnts = False
                        Exit Function
                    End If
                    gLogMsg "Deleting AST3: " & slCallLetters & " running on: " & slVehicleName & " for the period: " & Format$(tlCpttArray(ilIdx).sCpttStartDate, "mm/dd/yyyy") & " - " & Format$(slEndDate, "mm/dd/yyyy"), "FastAddVerbose.Txt", False
                    'cnn.CommitTrans

                    If (ilAgreeType = 1) Or (ilAgreeType = 2) Then
                        ilRet = gAlertAdd("R", "S", imVefCode, tlCpttArray(ilIdx).sCpttStartDate)
                    End If
                End If
            Next ilIdx
            slDrop = DateAdd("d", -1, slDrop)
            SQLQuery = "UPDATE att SET "
            If bmShowDates Then
                SQLQuery = SQLQuery & "attDropDate = '" & Format$(slDrop, sgSQLDateForm) & "',"
            Else
                SQLQuery = SQLQuery & "attOffAir = '" & Format$(slDrop, sgSQLDateForm) & "',"
            End If
            SQLQuery = SQLQuery & "attEnterDate = '" & Format(gNow(), sgSQLDateForm) & "',"
            SQLQuery = SQLQuery & "attEnterTime = '" & Format(gNow(), sgSQLTimeForm) & "', "
            SQLQuery = SQLQuery & "attSentToXDSStatus = '" & "M" & "'"
            SQLQuery = SQLQuery & " WHERE attCode = " & tmOverlapInfo(ilLoop).lAttCode & ""
            'cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "FastAdd-mAdjOverlapAgmnts"
                mAdjOverlapAgmnts = False
                Exit Function
            End If
            'cnn.CommitTrans
        End If
    Next ilLoop
    If UBound(tmFastAddAttCount) > 0 Then
        For ilLoop = 0 To UBound(tmFastAddAttCount) - 1 Step 1
            If tmFastAddAttCount(ilLoop).iShttCount > 1 Then
                ilShtt = gBinarySearchStationInfoByCode(tmFastAddAttCount(ilLoop).iShttCode)
                If ilShtt > 0 Then
                    gLogMsg Trim$(tgStationInfoByCode(ilShtt).sCallLetters) & " had " & tmFastAddAttCount(ilLoop).iShttCount & " Agreements affected", "FastAddVerbose.Txt", False
                End If
            End If
        Next ilLoop
    End If
    mAdjOverlapAgmnts = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFastAdd-mAdjOverlapAgmnts"
    mAdjOverlapAgmnts = False
    Exit Function
End Function

'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - replace Radio buttons to help UI
Private Sub cboGetStationsFrom_Click()
    Select Case cboGetStationsFrom.ListIndex
        Case 0: 'station active on vehicle
            If imCurrentStationList <> 1 Then
                rbcGetFrom(1).Value = True: rbcGetFrom_Click (1): txtBrowse.Text = ""
                mClearGrid grdExclude
                mClearGrid grdInclude
                If cboVehicle.Text <> "" Then cboVehicle_Click
            End If
        Case 1: 'all stations
            If imCurrentStationList <> 0 Then
                mClearGrid grdExclude
                mClearGrid grdInclude
                rbcGetFrom(0).Value = True: rbcGetFrom_Click (0): txtBrowse.Text = ""
            End If
        Case 2: 'external file
            If imCurrentStationList <> 2 Then
                mClearGrid grdExclude
                mClearGrid grdInclude
                cboGetStationsFrom.Enabled = False
                rbcGetFrom(2).Value = True

                If txtBrowse.Text = "" Then
                    rbcGetFrom_Click (2)
                End If
                cboGetStationsFrom.Enabled = True
            End If
    End Select
    
    If cboGetStationsFrom.ListCount = 4 Then
        cboGetStationsFrom.RemoveItem (3) 'remove the ""
    End If
    mGenOK
End Sub

Private Sub cboStationMarket_Change()
    Dim llRow As Long
    Dim slName As String
    Dim ilLen As Integer
    
    On Error GoTo ErrHand
    'If imStationMarketInChg Then
    '    Exit Sub
    'End If
    imStationMarketInChg = True
    Screen.MousePointer = vbHourglass
    slName = LTrim$(cboStationMarket.Text)
    ilLen = Len(slName)
    If imStationMarketBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imStationMarketBSMode = False
    End If
    
    llRow = SendMessageByString(cboStationMarket.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        cboStationMarket.ListIndex = llRow
        cboStationMarket.SelStart = ilLen
        cboStationMarket.SelLength = Len(cboStationMarket.Text)
    End If
    rbcGetFrom(0).Value = False
    rbcGetFrom(1).Value = False
    rbcGetFrom(2).Value = False
    
    mClearGrid grdExclude
    mClearGrid grdInclude
    
    'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - reset StationsFrom Selection, to force reload of Vehicles, etc... due to Station market changed
    cboGetStationsFrom.AddItem ""
    cboGetStationsFrom.ListIndex = 3
    imCurrentStationList = -1
    txtBrowse.Visible = False
    cboVehicle.Visible = False
    
    Screen.MousePointer = vbDefault
    'imStationMarketInChg = False
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-cboStationMarket_Change: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub cboStationMarket_Click()
    Dim att_rst As ADODB.Recordset
    On Error GoTo ErrHand

    If cboStationMarket.Text <> "" Then
        'lmBaseAttCode = CLng(cboStationMarket.ItemData(cboStationMarket.ListIndex))
        'If imBaseAttCode <> -1 Then
        '    SQLQuery = "Select * from att"
        '    SQLQuery = SQLQuery + " WHERE (attCode = " & lmBaseAttCode & ")"
        '    Set att_rst = gSQLSelectCall(SQLQuery)
        '    If Not att_rst.EOF Then
        '        imBaseShttCode = att_rst!attShfCode
        '    Else
        '        imBaseShttCode = 0
        '    End If
        'Else
        '    imBaseShttCode = -1
        'End If
        imBaseShttCode = cboStationMarket.ItemData(cboStationMarket.ListIndex)
    End If
    rbcGetFrom(0).Value = False
    rbcGetFrom(1).Value = False
    rbcGetFrom(2).Value = False
    
    'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
    If TabStrip1.SelectedItem.Index = 1 Then
        mClearGrid grdExclude
        mClearGrid grdInclude
        cboStationMarket_Change
        mGenOK
    End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFastAdd-cboStationMarket_Click"
    Exit Sub
End Sub

Private Sub cboStationMarket_KeyDown(KeyCode As Integer, Shift As Integer)
    imStationMarketBSMode = False
End Sub

Private Sub cboStationMarket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboStationMarket.SelLength <> 0 Then
            imStationMarketBSMode = True
        End If
    End If
End Sub

Private Sub cboVehicle_Change()
    Dim llRow As Long
    Dim slName As String
    Dim ilLen As Integer
    
    On Error GoTo ErrHand
    If imVehicleInChg Then
        mGenOK
        Exit Sub
    End If
    imVehicleInChg = True
    
    slName = LTrim$(cboVehicle.Text)
    ilLen = Len(slName)
    If imVehicleBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imVehicleBSMode = False
    End If
    
    llRow = SendMessageByString(cboVehicle.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        cboVehicle.ListIndex = llRow
        cboVehicle.SelStart = ilLen
        cboVehicle.SelLength = Len(cboVehicle.Text)
    End If
    Screen.MousePointer = vbDefault
    imVehicleInChg = False
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-cboVehicle_Change: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub cboVehicle_Click()
    On Error GoTo ErrHand
    'lbcStations(0).Clear
    'lbcStations(1).Clear
    mClearGrid grdExclude
    mClearGrid grdInclude
    If cboVehicle.Text <> "" Then
        Screen.MousePointer = vbHourglass
        mShowSelectiveStations
        Screen.MousePointer = vbDefault
    Else
        Exit Sub
    End If
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-cboVehicle_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub cboVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
    imVehicleBSMode = False
    mGenOK
End Sub

Private Sub cboVehicle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboVehicle.SelLength <> 0 Then
            imVehicleBSMode = True
        End If
    End If
    mGenOK
End Sub

'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
Private Sub cmcBrowse_Click()
    Dim slCurDir As String
    lblFastAddStatus.Caption = ""
    slCurDir = CurDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    ' Set filters
    CommonDialog1.Filter = "CSV Files (*.csv)|*.csv"
    ' Specify default filter
    CommonDialog1.FilterIndex = 1
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file
    txtFile.Text = Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    
    Screen.MousePointer = vbHourglass
    imFileOkay = False
    lblFastAddStatus.Caption = "Loading && Verifying File...."
    If mLoadAndVerifyFastAddFile = True Then
        If grdFastAddFile.Rows >= 1 Then
            cmdGen.Enabled = True
            lblFastAddStatus.Caption = "File Verified"
        End If
    Else
        lblFastAddStatus.Caption = "Failed Validation"
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    txtFile.Text = ""
    cmdGen.Enabled = False
    Exit Sub
End Sub

Private Sub cmdAll_Click()
    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim ilExclCount As Integer
    Dim ilInclCount As Integer
    Dim slStr As String
    Dim ilPos As Integer
    Dim slDates As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llActiveDate As Long
    On Error GoTo ErrHand
    
    If (txtActiveDate.Visible) And (Trim$(txtActiveDate.Text) <> "") Then
        If Not gIsDate(Trim$(txtActiveDate.Text)) Then
            txtActiveDate.SetFocus
            gMsgBox "Active date is not a valid date", vbCritical
            Exit Sub
        End If
        'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - removed Selected index, using Native grid row, RowSel properties
        grdExclude.Visible = False
        grdInclude.Visible = False
        Screen.MousePointer = vbHourglass
        llActiveDate = DateValue(gAdjYear(txtActiveDate.Text))
        'For ilLoop = 0 To lbcStations(0).ListCount - 1 Step 1
        '    slStr = lbcStations(0).List(ilLoop)
        '    ilPos = InStrRev(slStr, " ", -1, vbTextCompare)
        '    If ilPos > 0 Then
        '        slStr = Mid(slStr, ilPos + 1)
        '        ilPos = InStr(1, slStr, "-", vbTextCompare)
        '        If ilPos > 0 Then
        '            slStartDate = Left(slStr, ilPos - 1)
        '            slEndDate = Mid(slStr, ilPos + 1)
        '            If (slEndDate = "TFN") Then
        '                lbcStations(0).Selected(ilLoop) = True
        '            Else
        '                If llActiveDate <= DateValue(gAdjYear(slEndDate)) Then
        '                    lbcStations(0).Selected(ilLoop) = True
        '                End If
        '            End If
        '        End If
        '    End If
        'Next ilLoop
        For ilLoop = grdExclude.Rows - 1 To grdExclude.FixedRows Step -1
            If grdExclude.TextMatrix(ilLoop, CALLLETTERSINDEX) <> "" Then
                slStr = grdExclude.TextMatrix(ilLoop, DATERANGEINDEX)
                ilPos = InStr(1, slStr, "-", vbTextCompare)
                If ilPos > 0 Then
                    slStartDate = Left(slStr, ilPos - 1)
                    slEndDate = Mid(slStr, ilPos + 1)
                    If (slEndDate = "TFN") Then
                        'grdExclude.TextMatrix(ilLoop, SELECTEDINDEX) = "1"
                        grdExclude.Row = ilLoop
                        cmdMoveRight_Click
                    Else
                        If llActiveDate <= DateValue(gAdjYear(slEndDate)) Then
                            'grdExclude.TextMatrix(ilLoop, SELECTEDINDEX) = "1"
                            grdExclude.Row = ilLoop
                            cmdMoveRight_Click
                        End If
                    End If
                End If
            End If
        Next ilLoop
        'cmdMoveRight_Click
    Else
        'For ilLoop = 0 To lbcStations(0).ListCount - 1 Step 1
        '    lbcStations(1).AddItem lbcStations(0).List(ilLoop)
        '    lbcStations(1).ItemData(lbcStations(1).NewIndex) = lbcStations(0).ItemData(ilLoop)
        'Next ilLoop
        'lbcStations(0).Clear
        'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - removed Selected index, using Native grid row, RowSel properties
        grdExclude.Visible = False
        grdInclude.Visible = False
        For ilLoop = grdExclude.Rows - 1 To grdExclude.FixedRows Step -1
            If grdExclude.TextMatrix(ilLoop, CALLLETTERSINDEX) <> "" Then
                'grdExclude.TextMatrix(ilLoop, SELECTEDINDEX) = "1"
                grdExclude.Row = ilLoop
                cmdMoveRight_Click
            End If
        Next ilLoop
        'cmdMoveRight_Click
    End If
    
    grdExclude.Visible = True
    grdInclude.Visible = True
    Screen.MousePointer = vbDefault
    
    grdExclude.RowSel = grdExclude.Row
    grdExclude.Col = 0
    grdExclude.ColSel = SORTINDEX
    grdExclude.TopRow = grdExclude.Row
    
    grdInclude.RowSel = grdInclude.Row
    grdInclude.Col = 0
    grdInclude.ColSel = SORTINDEX
    grdInclude.TopRow = grdInclude.Row
    
    'lbcStations(0).ListIndex = -1
 '''   cmdAll.Visible = False
 '''   txtActiveDate.Visible = False
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-cmdAll_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload frmFastAdd
End Sub

Private Sub cmdGen_Click()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilShtt As Integer
    Dim slDate As String
    Dim att_rst As ADODB.Recordset
    Dim slMessage As String 'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness
    On Error GoTo ErrHand
    
    'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
    Dim llGenLoop As Long
    Dim llGenLoopMax As Long
    Dim ilImportErrors As Integer
    llGenLoopMax = 1
    If TabStrip1.SelectedItem.Index = 2 Then 'Fast Add from File
        llGenLoopMax = grdFastAddFile.Rows - 1
        mClearGrid grdInclude
    End If
    ilImportErrors = False
    
    'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
    'Loop through Fast Add File OR do the one Manual entry
    For llGenLoop = 1 To llGenLoopMax
        If TabStrip1.SelectedItem.Index = 2 Then 'Fast Add from File
            lblFastAddStatus.Caption = "Processing Line " & grdFastAddFile.TextMatrix(llGenLoop, FILELINE) & ": " & Int((llGenLoop - 1) / llGenLoopMax * 100) & "%"
            'FILESTATUS
            grdFastAddFile.TextMatrix(llGenLoop, FILESTATUS) = "Processing..."

            'FILEVEHICLE
            mSelectFromList lbcCreateVehicle, grdFastAddFile.TextMatrix(llGenLoop, FILEVEHICLE)
            If lbcCreateVehicle.Text = "" Then
                grdFastAddFile.TextMatrix(llGenLoop, FILESTATUS) = "No Vehicle found for " & grdFastAddFile.TextMatrix(llGenLoop, FILEVEHICLE)
                gLogMsg "Vehicle not found for " & grdFastAddFile.TextMatrix(llGenLoop, FILEVEHICLE), "FastAddVerbose.Txt", False
                GoTo SkipAdd
            End If
            
            'FILESTARTDATE
            txtStartDate.Text = grdFastAddFile.TextMatrix(llGenLoop, FILESTARTDATE)
            txtEndDate.Text = ""
            
            'Load list of "Get Pledges from" stations
            mGetStaMark llGenLoop
            
            'FILEPLEDGESFROM
            mSelectFromCombo cboStationMarket, grdFastAddFile.TextMatrix(llGenLoop, FILEPLEDGESFROM)
            If cboStationMarket.Text = "" Then
                grdFastAddFile.TextMatrix(llGenLoop, FILESTATUS) = "No Pledges found for " & grdFastAddFile.TextMatrix(llGenLoop, FILEPLEDGESFROM)
                gLogMsg "No Pledges found for " & grdFastAddFile.TextMatrix(llGenLoop, FILEPLEDGESFROM), "FastAddVerbose.Txt", False
                ilImportErrors = True
                GoTo SkipAdd
            End If
            
            'FILEDLVYSVCMODE
            cboCopyDeliveryService.ListIndex = Val(grdFastAddFile.TextMatrix(llGenLoop, FILEDLVYSVCMODE) - 1)
            
            'FILEGENFOR
            If LCase(grdFastAddFile.TextMatrix(llGenLoop, FILEGENFOR)) = "prior" Then
                rbcMulticast(0).Value = True
            Else
                rbcMulticast(1).Value = True
            End If
            
            'FILECALLLETTERS
            mSelectCallLetters Val(grdFastAddFile.TextMatrix(llGenLoop, FILECALLLETTERSCODE)), grdFastAddFile.TextMatrix(llGenLoop, FILECALLLETTERS)
        End If
    
        'The left list box
        If grdInclude.TextMatrix(grdInclude.FixedRows, SHTTCODEINDEX) = "" Then
            grdFastAddFile.TextMatrix(llGenLoop, FILESTATUS) = "No Stations to Process"
            gMsgBox "No Stations to Process"
            'Exit Sub
            ilImportErrors = True
            GoTo AddEnd
        End If
        
        'Start Date
        '11/15/14: Remove backslash if last character
        If right(txtStartDate.Text, 1) = "/" Then
            txtStartDate.Text = Left(txtStartDate.Text, Len(txtStartDate.Text) - 1)
        End If
        If txtStartDate.Text = "" Then
            If TabStrip1.SelectedItem.Index = 1 Then txtStartDate.SetFocus
            ilImportErrors = True
            grdFastAddFile.TextMatrix(llGenLoop, FILESTATUS) = "You must enter a Start Date."
            gMsgBox "You must enter a Start Date."
            'Exit Sub
            ilImportErrors = True
            GoTo AddEnd
        End If
        
        If Not gIsDate(Trim$(txtStartDate.Text)) Then
            If TabStrip1.SelectedItem.Index = 1 Then txtStartDate.SetFocus
            grdFastAddFile.TextMatrix(llGenLoop, FILESTATUS) = "Start date is not a valid date"
            gMsgBox "Start date is not a valid date", vbCritical
            'Exit Sub
            ilImportErrors = True
            GoTo AddEnd
        End If
        
        'End Date
        '11/15/14: Remove backslash if last character
        If right(txtEndDate.Text, 1) = "/" Then
            txtEndDate.Text = Left(txtEndDate.Text, Len(txtEndDate.Text) - 1)
        End If
        If txtEndDate.Text <> "" Then
            If Not gIsDate(Trim$(txtEndDate.Text)) Then
                If TabStrip1.SelectedItem.Index = 1 Then txtEndDate.SetFocus
                grdFastAddFile.TextMatrix(llGenLoop, FILESTATUS) = "End date is not a valid date"
                gMsgBox "End date is not a valid date", vbCritical
                'Exit Sub
                ilImportErrors = True
                GoTo AddEnd
            End If
        End If
        'D.S. 02/12/19 TTP 9208 Force a start date to fall on a Monday like Agreements do
        If Weekday(Trim$(txtStartDate.Text)) <> vbMonday Then
            txtStartDate.Text = DateValue(gObtainPrevMonday(gAdjYear(Format$(txtStartDate.Text, "m/d/yy"))))
        End If
        If Trim$(smAdminEMail) = "" And imWebType = True Then
            grdFastAddFile.TextMatrix(llGenLoop, FILESTATUS) = "System Administrator E-Mail Address Missing"
            gMsgBox "System Administrator E-Mail Address Missing", vbCritical
            'Exit Sub
            ilImportErrors = True
            GoTo AddEnd
        End If
        
        '11/5/14
        If (rbcMulticast(0).Value) = False And (rbcMulticast(1).Value = False) Then
            grdFastAddFile.TextMatrix(llGenLoop, FILESTATUS) = """" & "Create Multicast for" & """" & " not specified"
            gMsgBox """" & "Create Multicast for" & """" & " not specified", vbCritical
            'Exit Sub
            ilImportErrors = True
            GoTo AddEnd
        End If
        Screen.MousePointer = vbHourglass
        'Log user choices
        If TabStrip1.SelectedItem.Index = 1 Then
            gLogMsg "", "FastAddVerbose.Txt", False
        End If
        gLogMsg "Attempting to Create Agreements for:", "FastAddVerbose.Txt", False
        'For ilLoop = 0 To lbcStations(1).ListCount - 1 Step 1
        '    gLogMsg lbcStations(1).List(ilLoop), "FastAddVerbose.Txt", False
        'Next ilLoop
        slMessage = ""
        For ilLoop = grdInclude.FixedRows To grdInclude.Rows - 1 Step 1
            slMessage = slMessage & Trim$(grdInclude.TextMatrix(ilLoop, CALLLETTERSINDEX)) ' & vbCrLf
            'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - it is slow to open/close the file 1000's of times
            'gLogMsg Trim$(grdInclude.TextMatrix(ilLoop, CALLLETTERSINDEX)), "FastAddVerbose.Txt", False
        Next ilLoop
        gLogMsg slMessage, "FastAddVerbose.Txt", False
        gLogMsg "Agreements will be modeled from:", "FastAddVerbose.Txt", False
        '7/24/13: Change from combo box to multi-select list
        'gLogMsg "Vehicle: " & cboCreateVehicle.Text, "FastAddVerbose.Txt", False
        gLogMsg "Station: " & cboStationMarket.Text, "FastAddVerbose.Txt", False
        gLogMsg "With a Start Date of: " & txtStartDate.Text, "FastAddVerbose.Txt", False
        If txtEndDate.Text = "" Then
            gLogMsg "With an End Date of: TFN", "FastAddVerbose.Txt", False
        Else
            gLogMsg "With an End Date of: " & txtEndDate.Text, "FastAddVerbose.Txt", False
        End If
        If rbcGetFrom(0).Value = True Then
            gLogMsg "Stations selection was based off a listing of all stations", "FastAddVerbose.Txt", False
        End If
        If rbcGetFrom(1).Value = True Then
            gLogMsg "Stations selection was based off of agreements in: " & cboVehicle.Text, "FastAddVerbose.Txt", False
        End If
        If rbcGetFrom(2).Value = True Then
            gLogMsg "Stations selection was based off an external file", "FastAddVerbose.Txt", False
        End If
        
        slDate = Format(gNow(), "yyyy-mm-dd")
        imVefCode = -1
        Screen.MousePointer = vbHourglass
        '7/24/13: Change from combo box to multi-select list
        For ilLoop = 0 To lbcCreateVehicle.ListCount - 1 Step 1
            If lbcCreateVehicle.Selected(ilLoop) Then
                If imVefCode = -1 Then
                    mSaveStationList
                Else
                    mRestoreStationList
                End If
                gLogMsg "Vehicle: " & Trim$(lbcCreateVehicle.List(ilLoop)), "FastAddVerbose.Txt", False
                imVefCode = lbcCreateVehicle.ItemData(ilLoop)
    
                ''6/14/14: Verify that all multicast station are matched-up in the Include list.
                'gAlignMulticastStations imVefCode, "S", lbcStations(1), lbcStations(0)
                '11/5/14: Test if Prior or All selected
                If rbcMulticast(1).Value Then
                    mAlignAllMulticastStations
                Else
                    mAlignMulticastStations imVefCode, "S"
                End If
                
                SQLQuery = "Select attCode from att"
                SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode & ""
                'SQLQuery = SQLQuery + " AND (attOffAir  >= '" & slDate & "')"
                'SQLQuery = SQLQuery + " AND (attDropDate >= '" & slDate & "')"
                SQLQuery = SQLQuery + " AND (attOffAir  = '" & "2069-12-31" & "')"
                SQLQuery = SQLQuery + " AND (attDropDate = '" & "2069-12-31" & "')"
                SQLQuery = SQLQuery + " AND attShfCode = " & imBaseShttCode & ")"
                Set att_rst = gSQLSelectCall(SQLQuery)
                If Not att_rst.EOF Then
                    lmBaseAttCode = att_rst!attCode
                    'Start the pre-validations process from a list of all stations
                    mPreValidateStations
                    If igFastAddContinue Then
                        'If lbcStations(1).ListCount > 0 Then
                        If Trim$(grdInclude.TextMatrix(grdInclude.FixedRows, SHTTCODEINDEX)) <> "" Then
                            ilRet = mGetBaseAgreementInfo
                            If ilRet = False Then
                                'Exit Sub
                                ilImportErrors = True
                                grdFastAddFile.TextMatrix(llGenLoop, FILESTATUS) = "Failed to get Base Agreement Info"
                                GoTo AddEnd
                            End If
                            'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - Give the user a clue
                            ProgressBar2.Value = 0
                            ProgressBar2.Visible = True
                            mBuildAgreements
                            cmdCancel.Caption = "Done"
                            ProgressBar2.Value = 100
                            If TabStrip1.SelectedItem.Index = 1 Then
                                MsgBox "Generate Complete"
                            End If
                        Else
                            If TabStrip1.SelectedItem.Index = 1 Then
                                gLogMsg "Nothing to Process", "FastAddVerbose.Txt", False
                                gMsgBox "Nothing to Process"
                            End If
                        End If
                    End If
                Else
                    ilVef = gBinarySearchVef(CLng(imVefCode))
                    ilShtt = gBinarySearchStationInfoByCode(imBaseShttCode)
                    If (ilVef <> -1) And (ilShtt <> -1) Then
                        If TabStrip1.SelectedItem.Index = 1 Then
                            gLogMsg "Agreement not found for " & Trim$(tgVehicleInfo(ilVef).sVehicle) & " " & Trim$(tgStationInfoByCode(ilShtt).sCallLetters), "FastAddVerbose.Txt", False
                            gMsgBox "Agreement not found for " & Trim$(tgVehicleInfo(ilVef).sVehicle) & " " & Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
                        End If
                        If TabStrip1.SelectedItem.Index = 2 Then
                            ilImportErrors = True
                            grdFastAddFile.TextMatrix(llGenLoop, FILESTATUS) = "Agreement not found for " & Trim$(tgVehicleInfo(ilVef).sVehicle) & " " & Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
                        End If
                    End If
                End If
            End If
        Next ilLoop
        
SkipAdd:

        gLogMsg "", "FastAddVerbose.Txt", False
        
        'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
        If TabStrip1.SelectedItem.Index = 2 Then
            If grdFastAddFile.TextMatrix(llGenLoop, FILESTATUS) = "Processing..." Then
                grdFastAddFile.TextMatrix(llGenLoop, FILESTATUS) = "Agreement Added"
            End If
        End If
    Next llGenLoop

AddEnd:
    Screen.MousePointer = vbDefault
    ProgressBar2.Value = 0
    ProgressBar2.Visible = False
    If TabStrip1.SelectedItem.Index = 2 Then
        lblFastAddStatus.Caption = "Finished Adding Agreements"
        cmdGen.Enabled = False
        
        'if there's issues found, Remove the "Agreement Added" rows to show only BAD rows
        If ilImportErrors = True Then
            For ilLoop = grdFastAddFile.Rows - 1 To 1 Step -1
                If grdFastAddFile.TextMatrix(ilLoop, FILESTATUS) = "OK: no issues were detected" Then
                    grdFastAddFile.TextMatrix(ilLoop, FILESTATUS) = "Agreement Not Added (Skipped due to error)"
                End If
            Next ilLoop
            lblFastAddStatus.Caption = "Errors encountered adding Agreements:"
            MsgBox "Errors were encountered while adding Agreements." & vbCrLf & "Check the Fast add file grid for issues that may have prevented Agreement(s) from being added.", vbExclamation + vbOKOnly, "Fast Add File"
        End If
        
        imFileOkay = False
        txtStartDate.Text = ""
        mClearGrid grdExclude
        mClearGrid grdInclude
    End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-cmdGen_Click: "
        
        If TabStrip1.SelectedItem.Index = 2 Then
            grdFastAddFile.TextMatrix(llGenLoop, FILESTATUS) = gMsg
        End If
        
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub cmdMoveLeft_Click()
    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim ilExclCount As Integer
    Dim ilInclCount As Integer
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llRowSel As Long
    On Error GoTo ErrHand
    'For ilLoop = 0 To lbcStations(1).ListCount - 1 Step 1
    '    If lbcStations(1).Selected(ilLoop) Then
    '        lbcStations(0).AddItem lbcStations(1).List(ilLoop)
    '        lbcStations(0).ItemData(lbcStations(0).NewIndex) = lbcStations(1).ItemData(ilLoop)
    '    End If
    'Next ilLoop
    If grdInclude.Row = 0 Then grdInclude.Row = 1
    llRow = grdExclude.FixedRows
    For ilLoop = grdExclude.Rows - 1 To grdExclude.FixedRows Step -1
        If grdExclude.TextMatrix(ilLoop, CALLLETTERSINDEX) <> "" Then
            llRow = ilLoop + 1
            Exit For
        End If
    Next ilLoop
    'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - removed Selected index, using Native grid row, RowSel properties
'    For ilLoop = grdInclude.Rows - 1 To grdInclude.FixedRows Step -1
'        If grdInclude.TextMatrix(ilLoop, SELECTEDINDEX) = "1" Then
'            mMoveGridRow grdInclude, grdExclude, CLng(ilLoop), llRow
'        End If
'    Next ilLoop
    llRowSel = grdInclude.RowSel
    llCurrentRow = grdInclude.Row
    If llRowSel > llCurrentRow Then
        llRowSel = grdInclude.Row
        llCurrentRow = grdInclude.RowSel
    End If
    
    If llRowSel <> llCurrentRow Then
        grdExclude.Visible = False
        grdInclude.Visible = False
        Screen.MousePointer = vbHourglass
    End If
    
    For ilLoop = llCurrentRow To llRowSel Step -1
        If grdInclude.TextMatrix(ilLoop, CALLLETTERSINDEX) <> "" Then
            mMoveGridRow grdInclude, grdExclude, CLng(ilLoop), llRow
        End If
    Next ilLoop
    'ilInclCount = lbcStations(0).ListCount
    'ilExclCount = lbcStations(1).ListCount
    'For ilLoop = 0 To ilInclCount - 1 Step 1
    '    For ilIdx = 0 To ilExclCount - 1 Step 1
    '        If lbcStations(0).List(ilLoop) = lbcStations(1).List(ilIdx) Then
    '            lbcStations(1).RemoveItem (ilIdx)
    '            ilExclCount = ilExclCount - 1
    '            mGenOK
    '            Exit For
    '        End If
    '    Next ilIdx
    'Next ilLoop
    
    'unselect range, and leave a row selected
        
    If llRowSel <> llCurrentRow Then
        grdInclude.Row = grdInclude.Row
        grdInclude.RowSel = grdInclude.Row
        
        grdExclude.Visible = True
        grdInclude.Visible = True
        Screen.MousePointer = vbDefault
    End If
    grdInclude.Col = 0
    grdInclude.ColSel = SORTINDEX
    grdExclude.TopRow = grdExclude.Rows - 1
    'lbcStations(1).ListIndex = -1
    'lbcStations(0).ListIndex = -1
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-cmdMoveLeft_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub cmdMoveRight_Click()
    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim ilExclCount As Integer
    Dim ilInclCount As Integer
    Dim ilLen As Integer
    Dim ilMaxLen As Integer
    Dim slTemp As String
    Dim llRow As Long
    Dim ilCount As Integer
    Dim llCurrentRow As Long
    Dim llRowSel As Long
    
    On Error GoTo ErrHand
    If grdExclude.Row = 0 Then grdExclude.Row = 1
    If rbcGetFrom(2).Value = True Then
        ilCount = 0
        For llRow = grdExclude.FixedRows To grdExclude.Rows - 1 Step 1
            If grdExclude.TextMatrix(llRow, SHTTCODEINDEX) <> "" Then
                ilCount = ilCount + 1
                If ilCount > 1 Then
                    Exit For
                End If
            End If
        Next llRow
        'If lbcStations(0).ListIndex = 1 Then
        If ilCount = 1 Then
            gMsgBox "This station cannot be moved until the station information is entered into the station area."
        Else
            gMsgBox "These stations cannot be moved until the stations information is entered into the station area."
        End If
        mGenOK
        Exit Sub
    End If
    
    'For ilLoop = 0 To lbcStations(0).ListCount - 1 Step 1
    '    If lbcStations(0).Selected(ilLoop) Then
    '        lbcStations(1).AddItem lbcStations(0).List(ilLoop)
    '        lbcStations(1).ItemData(lbcStations(1).NewIndex) = lbcStations(0).ItemData(ilLoop)
    '    End If
    'Next ilLoop
    
    llRow = grdInclude.FixedRows
    For ilLoop = grdInclude.Rows - 1 To grdInclude.FixedRows Step -1
        If grdInclude.TextMatrix(ilLoop, CALLLETTERSINDEX) <> "" Then
            llRow = ilLoop + 1
            Exit For
        End If
    Next ilLoop
    'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - removed Selected index, using Native grid row, RowSel properties
    'For ilLoop = grdExclude.Rows - 1 To grdExclude.FixedRows Step -1
    '    If grdExclude.TextMatrix(ilLoop, SELECTEDINDEX) = "1" Then
    '        mMoveGridRow grdExclude, grdInclude, CLng(ilLoop), llRow
    '    End If
    'Next ilLoop
    llRowSel = grdExclude.RowSel
    llCurrentRow = grdExclude.Row
    If llRowSel > llCurrentRow Then
        llRowSel = grdExclude.Row
        llCurrentRow = grdExclude.RowSel
    End If
    If llRowSel <> llCurrentRow Then
        grdExclude.Visible = False
        grdInclude.Visible = False
        Screen.MousePointer = vbHourglass
    End If
    For ilLoop = llCurrentRow To llRowSel Step -1
        If grdExclude.TextMatrix(ilLoop, CALLLETTERSINDEX) <> "" Then
            mMoveGridRow grdExclude, grdInclude, CLng(ilLoop), llRow
        End If
    Next ilLoop
        
    'unselect range, and leave a row selected
    If llRowSel <> llCurrentRow Then
        grdExclude.Row = grdExclude.Row
        grdExclude.RowSel = grdExclude.Row
        Screen.MousePointer = vbDefault
        grdExclude.Visible = True
        grdInclude.Visible = True
    End If
    grdInclude.TopRow = grdInclude.Rows - 1
    grdExclude.Col = 0
    grdExclude.ColSel = SORTINDEX
    
    'ilInclCount = lbcStations(1).ListCount
    'ilExclCount = lbcStations(0).ListCount
    'For ilLoop = 0 To ilInclCount - 1 Step 1
    '    For ilIdx = 0 To ilExclCount - 1 Step 1
    '        If lbcStations(1).List(ilLoop) = lbcStations(0).List(ilIdx) Then
    '            lbcStations(0).RemoveItem (ilIdx)
    '            ilExclCount = ilExclCount - 1
    '            mGenOK
    '            Exit For
    '        End If
    '    Next ilIdx
    'Next ilLoop
   '
   ' For ilLoop = 0 To lbcStations(1).ListCount - 1 Step 1
   '     slTemp = lbcStations(1).List(ilLoop)
   '     'create horz. scrool bar if the text is wider than the list box
   '     ilLen = Me.TextWidth(slTemp)
   '     If Me.ScaleMode = vbTwips Then
   '         ilLen = ilLen / Screen.TwipsPerPixelX  ' if twips change to pixels
   '     End If
   '     If ilLen > ilMaxLen Then
   '         ilMaxLen = ilLen
    '    End If
    'Next ilLoop
    'SendMessageByNum lbcStations(1).hwnd, LB_SETHORIZONTALEXTENT, ilMaxLen, 0
    
    
    
    'lbcStations(1).ListIndex = -1
    'lbcStations(0).ListIndex = -1
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-cmdMoveRight_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub Form_Initialize()
    'D.S. 5/22/18 moved from unload
    Me.Width = Screen.Width / 1.05   '1.05  '1.15
    Me.Height = Screen.Height / 1.3 '15    '1.45    '1.25
    Me.Top = (Screen.Height - Me.Height) / 1.4
    Me.Left = (Screen.Width - Me.Width) / 2
    mSetGridColumns
    mSetGridTitles
    gGrid_IntegralHeight grdExclude
    gGrid_FillWithRows grdExclude
    gGrid_IntegralHeight grdInclude
    gGrid_FillWithRows grdInclude
    
    'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
    gGrid_IntegralHeight grdFastAddFile
    gGrid_FillWithRows grdFastAddFile
    
    grdExclude.Row = 1
    grdExclude.RowSel = 1
    grdInclude.Row = 1
    grdInclude.RowSel = 1
    
    grdExclude.Col = 0
    grdExclude.ColSel = SORTINDEX
    grdInclude.Col = 0
    grdInclude.ColSel = SORTINDEX
    
    imCurrentStationList = -1
    
    mClearGrid grdExclude
    mClearGrid grdInclude
    
    'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
    mClearGrid grdFastAddFile
    
    gSetFonts frmFastAdd
    gCenterForm frmFastAdd
    
End Sub

Private Sub Form_Load()
    Dim ilRet As Integer
    On Error GoTo ErrHand
    cmdGen.Enabled = False
    imCreateVehBSMode = False
    imCreateVehInChg = False
    imStationMarketBSMode = False
    'imStationMarketInChg = False
    imVehicleBSMode = False
    imVehicleInChg = False
    imCurrentStationList = 1
    
    frmFastAdd.Caption = "Affiliate Fast Add - " & sgClientName
    gLogMsg "", "FastAddVerbose.Txt", False
    gLogMsg "   *** Starting FastAdd Program   ***", "FastAddVerbose.Txt", False
    
    'TTP 11070 - Fast Add Turn on Delivery Service for Stop/Started Agreements
    cboCopyDeliveryService.Clear
    cboCopyDeliveryService.AddItem "1. Do not copy model station service to stations w/o agreements"
    cboCopyDeliveryService.AddItem "2. Copy model station service to stations w/o agreements"
    cboCopyDeliveryService.AddItem "3. Copy model station service to all stations"
    cboCopyDeliveryService.ListIndex = 0
    
    'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
    frcManual.Left = 120
    frcManual.Top = 4080
    frcFromFile.Left = 120
    frcFromFile.Top = 360
    
    ilRet = mInit
    If Not ilRet Then
        Exit Sub
    End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-Form_Load: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub

Private Function mInit()
    Dim ilLoop As Integer
    On Error GoTo ErrHand
    mInit = True
    imLastExcludeColSorted = -1
    imLastExcludeSort = -1
    lmLastExcludeClickedRow = -1
    
    imLastIncludeColSorted = -1
    imLastIncludeSort = -1
    lmLastIncludeClickedRow = -1
    
    'Load the list of vehicles for the Create area
    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        '11/4/09-  Show Log and Conventional vehicle.  Let client pick which they want agreements to be used for
        'Temporarily include only for Special user until testing is complete
        'If tgVehicleInfo(ilLoop).sVehType = "C" Or tgVehicleInfo(ilLoop).sVehType = "A" Or tgVehicleInfo(ilLoop).sVehType = "G" Or tgVehicleInfo(ilLoop).sVehType = "I" Then
        'If tgVehicleInfo(ilLoop).sVehType = "C" Or tgVehicleInfo(ilLoop).sVehType = "A" Or tgVehicleInfo(ilLoop).sVehType = "G" Or tgVehicleInfo(ilLoop).sVehType = "I" Or ((Len(sgSpecialPassword) = 4) And (tgVehicleInfo(ilLoop).sVehType = "L")) Then
        If tgVehicleInfo(ilLoop).sVehType = "C" Or tgVehicleInfo(ilLoop).sVehType = "A" Or tgVehicleInfo(ilLoop).sVehType = "G" Or tgVehicleInfo(ilLoop).sVehType = "I" Or (tgVehicleInfo(ilLoop).sVehType = "L") Then
            'If (tgVehicleInfo(ilLoop).sOLAExport <> "Y") Then
                '7/24/13: Change from combo box to multi-select list
                'cboCreateVehicle.AddItem Trim$(tgVehicleInfo(ilLoop).sVehicle)
                'cboCreateVehicle.ItemData(cboCreateVehicle.NewIndex) = tgVehicleInfo(ilLoop).icode
                lbcCreateVehicle.AddItem Trim$(tgVehicleInfo(ilLoop).sVehicle)
                lbcCreateVehicle.ItemData(lbcCreateVehicle.NewIndex) = tgVehicleInfo(ilLoop).iCode
            'End If
        End If
    Next ilLoop

    'Load the list of vehicles frothe Create area
    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        'If (tgVehicleInfo(ilLoop).sOLAExport <> "Y") Then
            cboVehicle.AddItem Trim$(tgVehicleInfo(ilLoop).sVehicle)
            cboVehicle.ItemData(cboVehicle.NewIndex) = tgVehicleInfo(ilLoop).iCode
        'End If
    Next ilLoop
    smAdminEMail = ""
    
    SQLQuery = "SELECT * FROM Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If rst!siteAdminArttCode > 0 Then
            SQLQuery = "SELECT * FROM ARTT Where arttCode = " & rst!siteAdminArttCode
            Set adrst = gSQLSelectCall(SQLQuery)
            If Not adrst.EOF Then
                smAdminEMail = Trim$(adrst!arttEmail)
            End If
            bmShowDates = False
            If rst!siteShowContrDate = "Y" Then
                bmShowDates = True
            End If
        End If
    End If
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd - mInit: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    mInit = False
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    gLogMsg "", "FastAddVerbose.Txt", False
    gLogMsg "   *** Ending FastAdd Program   ***", "FastAddVerbose.Txt", False
    attBase_rst.Close
    shtt_rst.Close
    dat_rst.Close
    rst_Pet.Close
    adrst.Close
    rst_Gsf.Close
    rst_Lst.Close
    Erase tmDat
    Erase tmOverlapInfo
    Erase tmFastAddAttCount
    'D.S. 5/22/18
    'unload frmFastAdd
    Set frmFastAdd = Nothing
End Sub

Private Sub grdExclude_DblClick()
    If grdExclude.Row > 0 Then cmdMoveRight_Click
End Sub

'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - Quick Find by typing on Exclude Grid (Keyboard handler)
Private Sub grdExclude_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 8: 'backspace
            If txtSearch.Text = "" Then
                Exit Sub
            Else
                txtSearch.Text = Mid(txtSearch.Text, 1, Len(txtSearch.Text) - 1)
            End If
        Case 13: 'enter
            cmdMoveRight_Click
            txtSearch.Text = ""
        Case 46: 'delete
            txtSearch.Text = ""
            txtSearch.Visible = False
        Case 189: 'dash
            txtSearch.Text = txtSearch.Text & "-"
        Case Else:
            If KeyCode >= 65 And KeyCode <= 90 Then 'a - z
                txtSearch.Text = txtSearch.Text & Chr$(KeyCode)
            Else
                'Do nothing
            End If
    End Select
End Sub

Private Sub grdExclude_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
'    Dim ilFound As Integer
'    Dim slStr As String
'
    If Y < grdExclude.RowHeight(0) Then
        grdExclude.Col = grdExclude.MouseCol
        mSortExcludeCol grdExclude.Col
        grdExclude.Row = 0
        grdExclude.Col = SHTTCODEINDEX
        Exit Sub
    End If
    llCurrentRow = grdExclude.MouseRow
    llCol = grdExclude.MouseCol
    If llCurrentRow < grdExclude.FixedRows Then
        Exit Sub
    End If
'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - get rid of this Selected index stuff (Slow on huge lists), use native Row and RowSel
    If llCurrentRow >= grdExclude.FixedRows Then
'        If grdExclude.TextMatrix(llCurrentRow, CALLLETTERSINDEX) <> "" Then
'            'grdExclude.TopRow = lmScrollTop
'            llTopRow = grdExclude.TopRow
'            If (Shift And CTRLMASK) > 0 Then
'                If grdExclude.TextMatrix(grdExclude.Row, SHTTCODEINDEX) <> "" Then
'                    If grdExclude.TextMatrix(grdExclude.Row, SELECTEDINDEX) <> "1" Then
'                        grdExclude.TextMatrix(grdExclude.Row, SELECTEDINDEX) = "1"
'                    Else
'                        grdExclude.TextMatrix(grdExclude.Row, SELECTEDINDEX) = "0"
'                    End If
'                    mPaintRowColor grdExclude, grdExclude.Row
'                End If
'            Else
'                For llRow = grdExclude.FixedRows To grdExclude.Rows - 1 Step 1
'                    If grdExclude.TextMatrix(llRow, CALLLETTERSINDEX) <> "" Then
'                        grdExclude.TextMatrix(llRow, SELECTEDINDEX) = "0"
'                        If grdExclude.TextMatrix(llRow, SHTTCODEINDEX) <> "" Then
'                            If (lmLastExcludeClickedRow = -1) Or ((Shift And SHIFTMASK) <= 0) Then
'                                If llRow = llCurrentRow Then
'                                    grdExclude.TextMatrix(llRow, SELECTEDINDEX) = "1"
'                                Else
'                                    grdExclude.TextMatrix(llRow, SELECTEDINDEX) = "0"
'                                End If
'                            ElseIf lmLastExcludeClickedRow < llCurrentRow Then
'                                If (llRow >= lmLastExcludeClickedRow) And (llRow <= llCurrentRow) Then
'                                    grdExclude.TextMatrix(llRow, SELECTEDINDEX) = "1"
'                                End If
'                            Else
'                                If (llRow >= llCurrentRow) And (llRow <= lmLastExcludeClickedRow) Then
'                                    grdExclude.TextMatrix(llRow, SELECTEDINDEX) = "1"
'                                End If
'                            End If
'                            'mPaintRowColor grdExclude, llRow
'                        End If
'                    End If
'                Next llRow
'                grdExclude.TopRow = llTopRow
'                grdExclude.Row = llCurrentRow
'            End If
            lmLastExcludeClickedRow = llCurrentRow
'        End If
    End If
End Sub

Private Sub grdInclude_DblClick()
    If grdInclude.Row > 0 Then cmdMoveLeft_Click
End Sub

Private Sub grdInclude_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdInclude.RowHeight(0) Then
        grdInclude.Col = grdInclude.MouseCol
        mSortIncludeCol grdInclude.Col
        grdInclude.Row = 0
        grdInclude.Col = SHTTCODEINDEX
        Exit Sub
    End If
    llCurrentRow = grdInclude.MouseRow
    llCol = grdInclude.MouseCol
    If llCurrentRow < grdInclude.FixedRows Then
        Exit Sub
    End If
    'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - get rid of this Selected index stuff (Slow on huge lists), use native Row and RowSel
    If llCurrentRow >= grdInclude.FixedRows Then
        If grdInclude.TextMatrix(llCurrentRow, CALLLETTERSINDEX) <> "" Then
            'grdInclude.TopRow = lmScrollTop
            llTopRow = grdInclude.TopRow
'            If (Shift And CTRLMASK) > 0 Then
'                If grdInclude.TextMatrix(grdInclude.Row, SHTTCODEINDEX) <> "" Then
'                    'If grdInclude.TextMatrix(grdInclude.Row, SELECTEDINDEX) <> "1" Then
'                    '    grdInclude.TextMatrix(grdInclude.Row, SELECTEDINDEX) = "1"
'                    'Else
'                    '    grdInclude.TextMatrix(grdInclude.Row, SELECTEDINDEX) = "0"
'                    'End If
'                    'mPaintRowColor grdInclude, grdInclude.Row
'                End If
'            Else
'                For llRow = grdInclude.FixedRows To grdInclude.Rows - 1 Step 1
'                    If grdInclude.TextMatrix(llRow, CALLLETTERSINDEX) <> "" Then
'                        grdInclude.TextMatrix(llRow, SELECTEDINDEX) = "0"
'                        If grdInclude.TextMatrix(llRow, SHTTCODEINDEX) <> "" Then
'                            If (lmLastIncludeClickedRow = -1) Or ((Shift And SHIFTMASK) <= 0) Then
'                                If llRow = llCurrentRow Then
'                                    grdInclude.TextMatrix(llRow, SELECTEDINDEX) = "1"
'                                Else
'                                    grdInclude.TextMatrix(llRow, SELECTEDINDEX) = "0"
'                                End If
'                            ElseIf lmLastIncludeClickedRow < llCurrentRow Then
'                                If (llRow >= lmLastIncludeClickedRow) And (llRow <= llCurrentRow) Then
'                                    grdInclude.TextMatrix(llRow, SELECTEDINDEX) = "1"
'                                End If
'                            Else
'                                If (llRow >= llCurrentRow) And (llRow <= lmLastIncludeClickedRow) Then
'                                    grdInclude.TextMatrix(llRow, SELECTEDINDEX) = "1"
'                                End If
'                            End If
'                            mPaintRowColor grdInclude, llRow
'                        End If
'                    End If
'                Next llRow
                'grdInclude.TopRow = llTopRow
                'grdInclude.Row = llCurrentRow
'            End If
            lmLastIncludeClickedRow = llCurrentRow
        End If
    End If
End Sub

Private Sub lbcCreateVehicle_Click()
    'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
    If TabStrip1.SelectedItem.Index = 2 Then
        Exit Sub
    End If
    tmcDelay.Enabled = False
    If lbcCreateVehicle.SelCount > 0 Then
        'tmcDelay.Enabled = True
        bmCreateChanged = True
    Else
        'lbcStations(0).Clear
        'lbcStations(1).Clear
        mClearGrid grdExclude
        mClearGrid grdInclude
        mGenOK
    End If
End Sub

Private Sub lbcStations_DblClick(Index As Integer)
    On Error GoTo ErrHand
'    If rbcGetFrom(2).Value = True Then
'        If lbcStations(0).ListIndex = 1 Then
'            gMsgBox "This station cannot be moved until the station information is entered into the station area."
'        Else
'            gMsgBox "These stations cannot be moved until the stations information is entered into the station area."
'        End If
'        Exit Sub
'    End If
'
'    If lbcStations(0).ListIndex >= 0 Then
'        lbcStations(1).AddItem lbcStations(0).Text
'        lbcStations(1).ItemData(lbcStations(1).NewIndex) = lbcStations(0).ItemData(lbcStations(0).ListIndex)
'        lbcStations(0).RemoveItem (lbcStations(0).ListIndex)
'        lbcStations(1).ListIndex = -1
'        lbcStations(0).ListIndex = -1
'    End If
'
'    If lbcStations(1).ListIndex >= 0 Then
'        lbcStations(0).AddItem lbcStations(1).Text
'        lbcStations(0).ItemData(lbcStations(0).NewIndex) = lbcStations(1).ItemData(lbcStations(1).ListIndex)
'        lbcStations(1).RemoveItem (lbcStations(1).ListIndex)
'        lbcStations(0).ListIndex = -1
'        lbcStations(1).ListIndex = -1
'        If rbcGetFrom(1).Value = True And lbcStations(0).ListCount > 0 Then
'            cmdAll.Visible = True
'            If Trim$(txtActiveDate.Text) = "" Then
'                txtActiveDate.Text = txtStartDate.Text
'            End If
'            txtActiveDate.Visible = True
'        End If
'
'    End If
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-lbcStations_DblClick: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - Reload if List of Vehicles to Create agreement for Changed
Private Sub lbcCreateVehicle_LostFocus()
    If bmCreateChanged = True Then
        tmcDelay_Timer
        bmCreateChanged = False
    End If
End Sub

Private Sub rbcGetFrom_Click(Index As Integer)
    On Error GoTo ErrHand
    'bcStations(0).Clear
    'lbcStations(1).Clear
    lblInclude.Caption = "Stations to Include"
    lblInclude.ForeColor = vbBlack
    lblExclude.ForeColor = vbBlack
    lblExclude.Caption = "Stations to Exclude"
    mSetGridColumns
    If rbcGetFrom(0).Value = True Then
        cboVehicle.ListIndex = -1
        cboVehicle.Visible = False
        'cmdAll.Visible = False
        frcAll.Visible = False
        'txtActiveDate.Visible = False
        txtBrowse.Visible = False
        'mShowAllStations
        If imCurrentStationList <> Index Then
            mClearGrid grdExclude
            mClearGrid grdInclude
            
            mShowAllStationsNew
            imCurrentStationList = Index
        End If
        mGenOK
        Exit Sub
    End If
    
    If rbcGetFrom(1).Value = True Then
        '7/24/13: Change from combo box to multi-select list
        'If cboCreateVehicle.Text = "" Then
'        If lbcCreateVehicle.SelCount <= 0 Then
'            gMsgBox "Please select from the ""Create one or more Agreements for"" field before continuing."
'            rbcGetFrom(1).Value = False
'            'boCreateVehicle.SetFocus
'            lbcCreateVehicle.SetFocus
'            rbcGetFrom(imCurrentStationList).Value = True
'            Exit Sub
'        End If
'
'        If txtStartDate.Text = "" Then
'            gMsgBox "Please enter the ""Start Date"" before continuing."
'            txtStartDate.SetFocus
'            rbcGetFrom(1).Value = False
'            rbcGetFrom(imCurrentStationList).Value = True
'            Exit Sub
'        End If
        
        cboVehicle.Visible = True
        txtBrowse.Visible = False
        imCurrentStationList = Index
    End If
    
    If rbcGetFrom(2).Value = True Then
        cboVehicle.ListIndex = -1
        cboVehicle.Visible = False
        'cmdAll.Visible = False
        frcAll.Visible = False
        'txtActiveDate.Visible = False
        mBrowse
        If txtBrowse.Text <> "" Then
            txtBrowse.Visible = True
            imCurrentStationList = Index
            mSetGridColumns
        Else
            rbcGetFrom(2).Value = False
            txtBrowse.Visible = False
            '7/8/21 - JW - Fix TTP 10051 / 10243 per Jason Email: Wed 7/7/21 3:24 PM
            cboGetStationsFrom.AddItem ""
            cboGetStationsFrom.ListIndex = 3
            imCurrentStationList = -1

            'rbcGetFrom(imCurrentStationList).Value = True
        End If
    End If
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    '7/8/21 - JW - Fix TTP 10051 / 10243 - dont show an error for user clicking Cancel on the Dialog: Error 32755 = "Cancel was selected."
    If (Err.Number <> 0) And (gMsg = "") And Err.Number <> 32755 Then
        gMsg = "A general error has occured in frmFastAdd-rbcGetFrom_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Public Sub mBrowse()
    Dim slTemp As String
    Dim slCurDir As String
    slCurDir = CurDir
    
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 3
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    slTemp = sgDatabaseName
    txtBrowse.Text = Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    mExternalFile
    imLastIncludeColSorted = -1
    imLastIncludeSort = -1
    mSortIncludeCol CALLLETTERSINDEX
    mGenOK
    'D.S. 5/22/18 added code
    If UBound(tgStaNameAndCode) > LBound(tgStaNameAndCode) Then
        bgFastAddCancelButton = True
        frmFastAddWarning.Show vbModal
        If igFastAddContinue = False Then
            rbcGetFrom(2).Value = False
            txtBrowse.Text = ""
            lblInclude.Caption = "Stations to Include"
            lblInclude.ForeColor = vbBlack
            lblExclude.ForeColor = vbBlack
            lblExclude.Caption = "Stations to Exclude"
            
            mClearGrid grdExclude
            mClearGrid grdInclude
        End If
    End If
    'If igFastAddContinue Then
    '    cmdGen_Click
    'End If
    
    ChDir slCurDir
    
    Exit Sub

ErrHandler:
    'User pressed the Cancel button
    '7/8/21 - JW - Fix TTP 10051 / 10243 per Jason Email: Wed 7/7/21 3:24 PM
    txtBrowse.Text = ""
    Exit Sub
End Sub

Private Sub mGetStaMark(Optional llRowNum As Long = 0)
    Dim shtt_att_rst As ADODB.Recordset
    Dim att_rst As ADODB.Recordset
    Dim slDate As String
    Dim slEndDate As String
    Dim slRange As String
    Dim llRow As Long
    Dim slName As String
    Dim slVefCode As String
    Dim ilLoop As Integer
    Dim llVehicleSelCount As Long
    Dim llAgreementCount As Long
    Dim ilPrevShttCode As Integer
    Dim slCallLetters As String
    Dim slMarket As String
    
    On Error GoTo ErrHand
    slDate = Format(gNow(), "yyyy-mm-dd")
    
    '7/24/13: Change from combo box to multi-select list
    'slName = cboCreateVehicle.Text
    'llRow = SendMessageByString(cboCreateVehicle.hwnd, CB_FINDSTRING, -1, slName)
    'imVefCode = CInt(cboCreateVehicle.ItemData(llRow))
    slVefCode = ""
    llVehicleSelCount = 0
    For ilLoop = 0 To lbcCreateVehicle.ListCount - 1 Step 1
        If lbcCreateVehicle.Selected(ilLoop) Then
            llVehicleSelCount = llVehicleSelCount + 1
            If slVefCode = "" Then
                slVefCode = lbcCreateVehicle.ItemData(ilLoop)
                slName = lbcCreateVehicle.List(ilLoop)
            Else
                slVefCode = slVefCode & "," & lbcCreateVehicle.ItemData(ilLoop)
            End If
        End If
    Next ilLoop
    
    'TTP 11074 - Fast Add: SQL errorwhen entering start date before selecting a vehicle
    If Trim(slVefCode) = "" Then
        Exit Sub
    End If
    
    slVefCode = "(" & slVefCode & ")"
    
    'SQLQuery = "SELECT attCode, attShfCode, attDropDate, attOffAir, AttOnAir, shttCallLetters, shttMarket"
    'SQLQuery = SQLQuery + " FROM att, shtt"
    SQLQuery = "SELECT attCode, attShfCode, attDropDate, attOffAir, AttOnAir, shttCallLetters, mktName, vefName"
    'SQLQuery = SQLQuery + " FROM att LEFT OUTER JOIN shtt on attShfCode = shttCode LEFT OUTER JOIN mkt on shttMktCode = mktCode LEFT OUTER JOIN vef_Vehicles on attVefCode = vefCode"
    SQLQuery = SQLQuery + " FROM att JOIN shtt on attShfCode = shttCode JOIN vef_Vehicles on attVefCode = vefCode LEFT OUTER JOIN mkt on shttMktCode = mktCode "
    '7/24/13: Change from combo box to multi-select list
    'SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode & ")"
    SQLQuery = SQLQuery + " WHERE (attVefCode in " & slVefCode & ")"
    'SQLQuery = SQLQuery + " AND shttCode = attShfCode"
    'SQLQuery = SQLQuery + " AND (attOffAir  >= '" & slDate & "')"
    'SQLQuery = SQLQuery + " AND (attDropDate >= '" & slDate & "')"
    SQLQuery = SQLQuery + " AND (attOffAir  = '" & "2069-12-31" & "')"
    SQLQuery = SQLQuery + " AND (attDropDate = '" & "2069-12-31" & "')"
    SQLQuery = SQLQuery + " Order by shttCallLetters, vefName"
    
    Set att_rst = gSQLSelectCall(SQLQuery)
'Debug.Print SQLQuery
    
    If att_rst.EOF Then
        cboStationMarket.Clear
        cboStationMarket.ForeColor = vbRed
        '7/24/13: Change from combo box to multi-select list
        'cboStationMarket.Text = "No stations are pledged for: " & cboCreateVehicle.Text
        If lbcCreateVehicle.ListCount = 1 Then
            cboStationMarket.Text = "No stations are pledged for: " & slName
            'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
            If TabStrip1.SelectedItem.Index = 2 Then
                If llRowNum > 0 Then
                    grdFastAddFile.TextMatrix(llRowNum, FILESTATUS) = "No stations are pledged for: " & slName
                End If
            End If
        Else
            cboStationMarket.Text = "No stations are pledged for selected vehicles"
        End If
        Exit Sub
    End If
    
    cboStationMarket.ForeColor = vbBlack
    ilPrevShttCode = -1
    llAgreementCount = 0
    If Not att_rst.EOF Then
        cboStationMarket.Clear
        While Not att_rst.EOF
            If (ilPrevShttCode) <> -1 And (ilPrevShttCode <> att_rst!attshfcode) Then
                If llAgreementCount = llVehicleSelCount Then
                    If llVehicleSelCount = 1 Then
                        If slMarket <> "" Then
                            cboStationMarket.AddItem Trim$(slCallLetters) & " , " & Trim$(slMarket) & " " & slRange
                        Else
                            cboStationMarket.AddItem Trim$(slCallLetters) & " " & slRange
                        End If
                        cboStationMarket.ItemData(cboStationMarket.NewIndex) = ilPrevShttCode
                    Else
                        If slMarket <> "" Then
                            cboStationMarket.AddItem Trim$(slCallLetters) & " , " & Trim$(slMarket)
                        Else
                            cboStationMarket.AddItem Trim$(slCallLetters)
                        End If
                        cboStationMarket.ItemData(cboStationMarket.NewIndex) = ilPrevShttCode
                    End If
                End If
                llAgreementCount = 0
            End If
            llAgreementCount = llAgreementCount + 1
            slCallLetters = Trim$(att_rst!shttCallLetters)
            If IsNull(att_rst!mktName) Then
                slMarket = ""
            Else
                slMarket = Trim$(att_rst!mktName)
            End If
            ilPrevShttCode = att_rst!attshfcode
            If DateValue(gAdjYear(att_rst!attDropDate)) < DateValue(gAdjYear(att_rst!attOffAir)) Then
                slEndDate = Format$(att_rst!attDropDate, sgShowDateForm)
            Else
                slEndDate = Format$(att_rst!attOffAir, sgShowDateForm)
            End If
            If (DateValue(gAdjYear(att_rst!attOnAir)) = DateValue("1/1/1970")) Then 'Or (att_rst!attOnAir = "1/1/70") Then    'Placeholder value to prevent using Nulls/outer joins
                slRange = ""
            Else
                slRange = Format$(Trim$(att_rst!attOnAir), sgShowDateForm)
            End If
            If (DateValue(gAdjYear(slEndDate)) = DateValue("12/31/2069") Or DateValue(gAdjYear(slEndDate)) = DateValue("12/31/69")) Then  'Or (att_rst!attOffAir = "12/31/69") Then
                If slRange <> "" Then
                    slRange = slRange & "-TFN"
                End If
            Else
                If slRange <> "" Then
                    slRange = slRange & "-" & slEndDate    'att_rst!attOffAir
                Else
                    slRange = "Thru " & slEndDate 'att_rst!attOffAir
                End If
            End If
        
''            SQLQuery = "SELECT shttCallLetters, shttMarket"
''            SQLQuery = SQLQuery + " FROM shtt"
''            SQLQuery = SQLQuery + " WHERE (shttCode = " & att_rst!attShfCode & ")"
''            Set shtt_att_rst = gSQLSelectCall(SQLQuery)
''            cboStationMarket.AddItem Trim$(shtt_att_rst!shttCallLetters) & " , " & Trim$(shtt_att_rst!shttMarket) & " " & slRange
''            cboStationMarket.ItemData(cboStationMarket.NewIndex) = att_rst!attCode
            ''cboStationMarket.AddItem Trim$(att_rst!shttCallLetters) & " , " & Trim$(att_rst!shttMarket) & " " & slRange
            'If IsNull(att_rst!mktName) = True Then
            '    cboStationMarket.AddItem Trim$(att_rst!shttCallLetters) & " " & Trim$(att_rst!vefName) & " " & slRange
            'Else
            '    cboStationMarket.AddItem Trim$(att_rst!shttCallLetters) & " " & Trim$(att_rst!vefName) & " , " & Trim$(att_rst!mktName) & " " & slRange
            'End If
            'cboStationMarket.ItemData(cboStationMarket.NewIndex) = att_rst!attCode
            att_rst.MoveNext
        Wend
        If (ilPrevShttCode) <> -1 Then
            If llAgreementCount = llVehicleSelCount Then
                If llVehicleSelCount = 1 Then
                    If slMarket <> "" Then
                        cboStationMarket.AddItem Trim$(slCallLetters) & " , " & Trim$(slMarket) & " " & slRange
                    Else
                        cboStationMarket.AddItem Trim$(slCallLetters) & " " & slRange
                    End If
                    cboStationMarket.ItemData(cboStationMarket.NewIndex) = ilPrevShttCode
                Else
                    If slMarket <> "" Then
                        cboStationMarket.AddItem Trim$(slCallLetters) & " , " & Trim$(slMarket)
                    Else
                        cboStationMarket.AddItem Trim$(slCallLetters)
                    End If
                    cboStationMarket.ItemData(cboStationMarket.NewIndex) = ilPrevShttCode
                End If
            End If
        End If
    End If
    
    If cboStationMarket.ListCount = 0 Then
        'cboStationMarket.Enabled = False
        'ckcDelivery.Enabled = False
    Else
        cboStationMarket.Enabled = True
        'ckcDelivery.Enabled = True
    End If
    
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFastAdd-mGetStaMark"
End Sub

'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - replaced with mShowAllStations
'Private Sub mShowAllStations()
'    Dim ilLoop As Integer
'    Dim ilAddStation As Integer
'    Dim llRow As Long
'    Dim llIndex As Long
''    Dim att_rst As ADODB.Recordset
'
'    llRow = grdExclude.FixedRows
'    For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
'        ilAddStation = True
'        If imBaseShttCode = tgStationInfo(ilLoop).iCode Then
''            'Test if agreement exist
''            SQLQuery = "Select * from att"
''            SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode & ""
''            SQLQuery = SQLQuery + " AND attShfCode = " & imBaseShttCode & ")"
''            Set att_rst = gSQLSelectCall(SQLQuery)
''            If Not att_rst.EOF Then
'                ilAddStation = False
''            End If
'        End If
'        If ilAddStation Then
'            If tgStationInfo(ilLoop).sUsedForATT = "Y" Then
'                'lbcStations(0).AddItem Trim$(tgStationInfo(ilLoop).sCallLetters) & ", " & Trim$(tgStationInfo(ilLoop).sMarket)
'                'lbcStations(0).ItemData(lbcStations(0).NewIndex) = tgStationInfo(ilLoop).iCode
'                mAddStationToGrid grdExclude, llRow, Trim$(tgStationInfo(ilLoop).sCallLetters), ""
'            End If
'        End If
'    Next ilLoop
'    imLastExcludeColSorted = -1
'    imLastExcludeSort = -1
'    mSortExcludeCol CALLLETTERSINDEX
'    mGenOK
'    Exit Sub
'
'ErrHand:
'    Screen.MousePointer = vbDefault
'    gMsg = ""
'    If (Err.Number <> 0) And (gMsg = "") Then
'        gMsg = "A general error has occured in frmFastAdd-mShowAllStations: "
'        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
'        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
'    End If
'End Sub

Private Sub mShowSelectiveStations()
    Dim shtt_rst As ADODB.Recordset
    Dim att_rst As ADODB.Recordset
    Dim ilFromVefCode As Integer
    Dim ilCreateVefCode As Integer
    Dim slDate As String
    Dim llStartDate As Long
    Dim llTodayDate As Long
    Dim slName As String
    Dim llRow As Long
    Dim slEndDate As String
    Dim slRange As String
    Dim ilShtt As Integer
    Dim blFound As Boolean
    Dim ilLen As Integer
    Dim ilMaxLen As Integer
    Dim slTemp As String
    Dim ilLoop As Integer
    
    On Error GoTo ErrHand
    'lbcStations(0).Clear
    mClearGrid grdExclude
    
    slName = cboVehicle.Text
    llRow = SendMessageByString(cboVehicle.hwnd, CB_FINDSTRING, -1, slName)
    ilFromVefCode = CInt(cboVehicle.ItemData(llRow))
    
    '7/24/13: Change from combo box to multi-select list
    'slName = cboCreateVehicle.Text
    'llRow = SendMessageByString(cboCreateVehicle.hwnd, CB_FINDSTRING, -1, slName)
    'ilCreateVefCode = CInt(cboCreateVehicle.ItemData(llRow))
    If lbcCreateVehicle.SelCount = 1 Then
        For ilLoop = 0 To lbcCreateVehicle.ListCount - 1 Step 1
            If lbcCreateVehicle.Selected(ilLoop) Then
                ilCreateVefCode = lbcCreateVehicle.ItemData(ilLoop)
                Exit For
            End If
        Next ilLoop
    Else
        ilCreateVefCode = -1
    End If
    
    If ilFromVefCode <> ilCreateVefCode Then
        'If the vehicle that the agreements are being created for is different
        'from the vehicle that the stations are being selected from, show all
        'stations except those that are ended prior to the start date of the new
        'agreements or prior to today's date, whichever is earlier
        llTodayDate = DateValue(gAdjYear(Format$(gNow(), "m/d/yy")))
        llStartDate = 0
        If txtStartDate.Text <> "" Then
            llStartDate = DateValue(gAdjYear(txtStartDate.Text))
        End If
        
        If llTodayDate <= llStartDate Or txtStartDate.Text = "" Then
            slDate = Format(gNow(), "yyyy-mm-dd")
        Else
            slDate = Format(txtStartDate.Text, "yyyy-mm-dd")
        End If
        
        SQLQuery = "SELECT attCode, attShfCode, attVefCode, attDropDate, attOffAir, AttOnAir"
        SQLQuery = SQLQuery + " FROM att"
        SQLQuery = SQLQuery + " WHERE (attVefCode = '" & ilFromVefCode & "')"
        SQLQuery = SQLQuery + " AND (attOffAir  >= '" & slDate & "')"
        SQLQuery = SQLQuery + " AND (attAgreeEnd >= '" & slDate & "')"
        Set att_rst = gSQLSelectCall(SQLQuery)
    Else
        'If the vehicle that the agreements are being created for is the same as
        'the vehicle from which the stations are selected, then exclude nothing,
        'as they all have to have end dates prior to the start date of the new
        'agreements, or they will be rejected due to overlap errors.

        SQLQuery = "SELECT DISTINCT attCode, attShfCode, attVefCode, attDropDate, attOffAir, AttOnAir"
        SQLQuery = SQLQuery + " FROM att"
        SQLQuery = SQLQuery + " WHERE (attVefCode = '" & ilFromVefCode & "')"
        Set att_rst = gSQLSelectCall(SQLQuery)

    End If
    
    If att_rst.EOF Then
        ''7/24/13: Change from combo box to multi-select list
        ''lbcStations(0).Text = "No station are pledged for: " & cboCreateVehicle.Text
        'lbcStations(0).Text = "No station are pledged for: " & cboVehicle.Text
        grdExclude.TextMatrix(grdExclude.FixedRows, CALLLETTERSINDEX) = "No Station"
        mGenOK
        Exit Sub
    End If
    llRow = grdExclude.FixedRows
    While Not att_rst.EOF
        If (lmBaseAttCode = att_rst!attCode) Or ((imBaseShttCode = att_rst!attshfcode) And (imVefCode = att_rst!attvefCode)) Then
        Else
            If DateValue(gAdjYear(att_rst!attDropDate)) < DateValue(gAdjYear(att_rst!attOffAir)) Then
                slEndDate = Format$(att_rst!attDropDate, sgShowDateForm)
            Else
                slEndDate = Format$(att_rst!attOffAir, sgShowDateForm)
            End If
            If (DateValue(gAdjYear(att_rst!attOnAir)) = DateValue("1/1/1970")) Then 'Or (att_rst!attOnAir = "1/1/70") Then    'Placeholder value to prevent using Nulls/outer joins
                slRange = ""
            Else
                slRange = Format$(Trim$(att_rst!attOnAir), sgShowDateForm)
            End If
            If (DateValue(gAdjYear(slEndDate)) = DateValue("12/31/2069") Or DateValue(gAdjYear(slEndDate)) = DateValue("12/31/69")) Then  'Or (att_rst!attOffAir = "12/31/69") Then
                If slRange <> "" Then
                    slRange = slRange & "-TFN"
                End If
            Else
                If slRange <> "" Then
                    slRange = slRange & "-" & slEndDate    'att_rst!attOffAir
                Else
                    slRange = "Thru " & slEndDate 'att_rst!attOffAir
                End If
            End If
            blFound = False
            'For ilShtt = 0 To lbcStations(0).ListCount - 1 Step 1
            '    If lbcStations(0).ItemData(ilShtt) = att_rst!attshfCode Then
            '        blFound = True
            '        lbcStations(0).List(ilShtt) = lbcStations(0).List(ilShtt) & ", " & slRange
            '        Exit For
            '    End If
            'Next ilShtt
            For ilShtt = grdExclude.FixedRows To grdExclude.Rows - 1 Step 1
                If Val(grdExclude.TextMatrix(ilShtt, SHTTCODEINDEX)) = att_rst!attshfcode Then
                    blFound = True
                    'lbcStations(0).List(ilShtt) = lbcStations(0).List(ilShtt) & ", " & slRange
                    grdExclude.TextMatrix(ilShtt, DATERANGEINDEX) = slRange
                    Exit For
                End If
            Next ilShtt
            '11/16/12:  Disaloow creating agreement for station modeling from
            If imBaseShttCode = att_rst!attshfcode Then
                blFound = True
            End If
            If Not blFound Then
                'SQLQuery = "SELECT shttCallLetters, shttMarket"
                'SQLQuery = SQLQuery + " FROM shtt"
                SQLQuery = "SELECT shttCallLetters, mktName"
                SQLQuery = SQLQuery + " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode"
                SQLQuery = SQLQuery + " WHERE (shttCode = " & att_rst!attshfcode & ")"
                Set shtt_rst = gSQLSelectCall(SQLQuery)
            
                ''lbcStations(0).AddItem Trim$(shtt_rst!shttCallLetters) & " , " & Trim$(shtt_rst!shttMarket) & " " & slRange
                'If IsNull(shtt_rst!mktName) = True Then
                '    lbcStations(0).AddItem Trim$(shtt_rst!shttCallLetters) & " " & slRange
                'Else
                '    lbcStations(0).AddItem Trim$(shtt_rst!shttCallLetters) & " , " & Trim$(shtt_rst!mktName) & " " & slRange
                'End If
                'lbcStations(0).ItemData(lbcStations(0).NewIndex) = att_rst!attshfCode
                mAddStationToGrid grdExclude, llRow, Trim$(shtt_rst!shttCallLetters), slRange

            End If
        End If
        att_rst.MoveNext
    Wend
    
    'ilLen = 0
    'ilMaxLen = 0
    'If lbcStations(0).ListCount > 0 Then
    '    For ilLoop = 0 To lbcStations(0).ListCount - 1 Step 1
    '        slTemp = lbcStations(0).List(ilLoop)
    '        'create horz. scrool bar if the text is wider than the list box
    '        ilLen = Me.TextWidth(slTemp)
    '        If Me.ScaleMode = vbTwips Then
    '            ilLen = ilLen / Screen.TwipsPerPixelX  ' if twips change to pixels
    '        End If
    '        If ilLen > ilMaxLen Then
    '            ilMaxLen = ilLen
    '        End If
    '    Next ilLoop
    '    SendMessageByNum lbcStations(0).hwnd, LB_SETHORIZONTALEXTENT, ilMaxLen, 0
    '
    '    cmdAll.Visible = True
    '    If Trim$(txtActiveDate.Text) = "" Then
    '        txtActiveDate.Text = txtStartDate.Text
    '    End If
    '    txtActiveDate.Visible = True
    'End If
    If grdExclude.TextMatrix(grdExclude.FixedRows, CALLLETTERSINDEX) <> "" Then
        'cmdAll.Visible = True
        frcAll.Visible = True
        If Trim$(txtActiveDate.Text) = "" Then
            txtActiveDate.Text = txtStartDate.Text
        End If
        txtActiveDate.Visible = True
    End If
    imLastExcludeColSorted = -1
    imLastExcludeSort = -1
    mSortExcludeCol CALLLETTERSINDEX
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFastAdd-mShowSelectiveStations"
End Sub

Private Sub mExternalFile()
    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject
    Dim slRetString As String
    Dim slLocation As String
    Dim slTemp As String
    Dim ilLineNumber As Integer
    Dim ilPos As Integer
    
    On Error GoTo ErrHand
    
    ReDim tgStaNameAndCode(0 To 0)
    slLocation = Trim$(txtBrowse.Text)
    If fs.FILEEXISTS(slLocation) Then
        Set tlTxtStream = fs.OpenTextFile(slLocation, ForReading, False)
    Else
        gMsgBox "** No Data Available **"
        mGenOK
        Exit Sub
    End If
        
    'D.S. 5/22/18 moved below
    'lblExclude.Caption = "Missing From Station File - Please Add Before Continuing        "
    'lblExclude.ForeColor = vbRed
    'lblInclude.Caption = "Found in Station File - OK"
    ilLineNumber = 0
    Do While tlTxtStream.AtEndOfStream <> True
        slRetString = tlTxtStream.ReadLine
        slRetString = UCase(slRetString)
        ilPos = InStr(1, slRetString, ",")
        If ilPos > 0 Then
            slRetString = Left(slRetString, ilPos - 1)
        End If
        ilLineNumber = ilLineNumber + 1
        slTemp = mTestCallLetters(slRetString)
        If slTemp <> "" Then
            gLogMsg "On line number " & Str(ilLineNumber) & " " & "On line number " & Str(ilLineNumber) & " " & slTemp, "FastAddVerbose.Txt", False
            tgStaNameAndCode(UBound(tgStaNameAndCode)).sStationName = Trim$(slRetString)
            tgStaNameAndCode(UBound(tgStaNameAndCode)).sInfo = "On line number " & Str(ilLineNumber) & " " & slTemp
            ReDim Preserve tgStaNameAndCode(0 To UBound(tgStaNameAndCode) + 1)
        End If
    Loop
    
    tlTxtStream.Close
    'D.S. 5/22/18 added code
    lblInclude.Caption = "Found in Station File - OK"
    If UBound(tgStaNameAndCode) > LBound(tgStaNameAndCode) Then
        lblExclude.Caption = "Missing From Station File - Please Add Before Continuing        "
        lblExclude.ForeColor = vbRed
    Else
        lblInclude.ForeColor = vbBlack
        lblExclude.ForeColor = vbBlack
        lblExclude.Caption = "Stations to Exclude"
    End If

    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-mExternalFile: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    
End Sub

Private Function mTestCallLetters(sCallLetters As String) As String
    Dim slTemp As String
    Dim slCallLetters As String
    Dim ilPos As Integer
    On Error GoTo ErrHand
    slCallLetters = UCase(sCallLetters)
    mTestCallLetters = ""
    
    ''Test for a dash in either the 4th or 5th position
    'slTemp = Mid$(slCallLetters, 4, 1)
    'If slTemp <> "-" Then
    '    slTemp = Mid$(slCallLetters, 5, 1)
    '    If slTemp <> "-" Then
    '        mTestCallLetters = "No dash ""-"" character was found or it's not in a valid position."
    '        mGenOK
    '        Exit Function
    '    End If
    'End If
    
    ilPos = InStr(1, slCallLetters, "-", vbBinaryCompare)
    If ilPos <= 0 Then
        mTestCallLetters = "No dash ""-"" character was found."
        mGenOK
        Exit Function
    End If
    ''Test for AM or FM n either the 5th and 6th position or the 6th and 7th position
    'slTemp = Mid$(slCallLetters, 5, 2)
    'If slTemp <> "AM" And slTemp <> "FM" And slTemp <> "HD" Then
    '    slTemp = Mid$(slCallLetters, 6, 2)
    '    If slTemp <> "AM" And slTemp <> "FM" And slTemp <> "HD" Then
    '        mTestCallLetters = "No AM or FM or HD characters were found or they are not in a valid postion."
    '        mGenOK
    '        Exit Function
    '    End If
    'End If
    slTemp = Mid$(slCallLetters, ilPos + 1, 1)
    If slTemp <> "A" And slTemp <> "F" And slTemp <> "H" Then
        mTestCallLetters = "No -A or -F or -H characters were found."
        mGenOK
        Exit Function
    End If
    ilPos = InStr(1, slCallLetters, "-", vbTextCompare)
    slTemp = Mid$(slCallLetters, 1, ilPos + 2)
    mAddExtFileSta slTemp
    mGenOK
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-mTestCallLetters: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    
End Function

Private Sub mAddExtFileSta(sCallLetters As String)
    Dim shtt_rst As ADODB.Recordset
'    Dim att_rst As ADODB.Recordset
    Dim ilVefCode As Integer
    Dim ilAddStation As Integer
    Dim ilShtt As Integer
    Dim llRow As Long
    Dim llERow As Long
    
    On Error GoTo ErrHand
    llRow = grdInclude.FixedRows
    llERow = grdExclude.FixedRows
    SQLQuery = "SELECT shttCallLetters, shttCode"
    SQLQuery = SQLQuery + " FROM shtt"
    SQLQuery = SQLQuery + " WHERE (shttCallLetters = '" & sCallLetters & "')"
    Set shtt_rst = gSQLSelectCall(SQLQuery)
    
    If Not shtt_rst.EOF Then
        'Test that agreement dose not already exist
        ilAddStation = True
        ''11/16/12: Check if name previously added
        'For ilShtt = 0 To lbcStations(1).ListCount - 1 Step 1
        '    If lbcStations(1).ItemData(ilShtt) = Val(shtt_rst!shttCode) Then
        '        ilAddStation = False
        '        Exit For
        '    End If
        'Next ilShtt
        For ilShtt = grdInclude.FixedRows To grdInclude.Rows - 1 Step 1
            If grdInclude.TextMatrix(ilShtt, SHTTCODEINDEX) = shtt_rst!shttCode Then
                ilAddStation = False
                Exit For
            End If
        Next ilShtt
        If imBaseShttCode = Val(shtt_rst!shttCode) Then
            'Test if agreement exist
'            SQLQuery = "Select * from att"
'            SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode & ""
'            SQLQuery = SQLQuery + " AND attShfCode = " & imBaseShttCode & ")"
'            Set att_rst = gSQLSelectCall(SQLQuery)
'            If Not att_rst.EOF Then
                ilAddStation = False
'            End If
        End If
        If ilAddStation Then
            'lbcStations(1).AddItem Trim$(sCallLetters)
            'lbcStations(1).ItemData(lbcStations(1).NewIndex) = Val(shtt_rst!shttCode)
            mAddStationToGrid grdInclude, llRow, Trim$(sCallLetters), ""
        End If
    Else
        'lbcStations(0).AddItem Trim$(sCallLetters)
        mAddStationToGrid grdExclude, llERow, Trim$(sCallLetters), ""
        gLogMsg "Warning: " & sCallLetters & "Call Letters were not in the database.", "FastAddVerbose.Txt", False
        tgStaNameAndCode(UBound(tgStaNameAndCode)).sStationName = Trim$(sCallLetters)
        tgStaNameAndCode(UBound(tgStaNameAndCode)).sInfo = "Call Letters were not found in the database - Please Add"
        ReDim Preserve tgStaNameAndCode(0 To UBound(tgStaNameAndCode) + 1)
    End If
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-mAddExtFileSta: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
Resume Next
End Sub

Private Sub mGenOK()
    Dim ilRet As Integer
    On Error GoTo ErrHand
    
    'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
    If TabStrip1.SelectedItem.Index = 2 Then
        If imFileOkay = True Then
            cmdGen.Enabled = True
        Else
            cmdGen.Enabled = False
        End If
        Exit Sub
    End If
    
    ilRet = InStr(1, cboStationMarket.Text, "No stations", vbTextCompare)
    If ilRet = 1 Then
        Exit Sub
    End If
    
    'One of these options must be set
    If rbcMulticast(0).Value = False And rbcMulticast(1).Value = False Then
        cmdGen.Enabled = False
        igFastAddContinue = False
        Exit Sub
    End If
    If txtStartDate.Text = "" Then
        cmdGen.Enabled = False
        igFastAddContinue = False
        Exit Sub
    End If
   
    '7/24/13: Change from combo box to multi-select list
    'If cboCreateVehicle.Text <> "" And cboStationMarket.Text <> "" Then
    If (lbcCreateVehicle.ListCount > 0) And (cboStationMarket.Text <> "") Then
        'If lbcStations(1).ListCount > 0 Then
        If grdInclude.TextMatrix(grdInclude.FixedRows, CALLLETTERSINDEX) <> "" Then
            If (rbcGetFrom(0).Value = True) Or (rbcGetFrom(1).Value = True And cboVehicle.Text <> "") Or (rbcGetFrom(2).Value = True And txtBrowse.Text <> "") Then
                cmdGen.Enabled = True
            End If
        Else
            cmdGen.Enabled = False
            igFastAddContinue = False
        End If
    '7/24/13: Change from combo box to multi-select list
    Else
        cmdGen.Enabled = False
        igFastAddContinue = False
    End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-mGenOK: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Function mGetBaseAgreementInfo() As Integer
    Dim ilUpper As Integer
    Dim ilRet As Integer
    Dim shtt_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    mGetBaseAgreementInfo = False
    'Get the agreement info from ATT that we will be modeling from
    SQLQuery = "Select * from att "
    SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode & ""
    SQLQuery = SQLQuery + " AND attCode = " & lmBaseAttCode & ")"
    Set attBase_rst = gSQLSelectCall(SQLQuery)
    
    imWebType = False
    imUnivisionType = False
    'imAgreeType = attBase_rst!attExportType
    'If attBase_rst!attExportType = 1 Then
    If attBase_rst!attExportToWeb = "Y" Then
        imWebType = True
    End If
    
    'If attBase_rst!attExportType = 2 Then
    If attBase_rst!attExportToUnivision = "Y" Then
        imUnivisionType = True
    End If
    
    imBaseShttCode = attBase_rst!attshfcode
    
    'Determine if the station has a time zone defined
    SQLQuery = "Select shttTimeZone, shttCallLetters from shtt"
    SQLQuery = SQLQuery + " WHERE shttCode = " & imBaseShttCode & ""
    Set shtt_rst = gSQLSelectCall(SQLQuery)
    
    imStationHasTimeZoneDefined = True
    If Trim$(shtt_rst!shttTimeZone) = "" Then
        imStationHasTimeZoneDefined = False
        ilRet = gMsgBox("No Time Zone is defined for " & Trim$(shtt_rst!shttCallLetters) & ".  Times will not be converted to local station's time zones.  Do you wish to continue?", vbYesNo)
        If ilRet = vbNo Then
            Exit Function
        End If
    End If
    
    'Get the avails and pledge info from DAT table that we will be modeling from
    ilRet = mGetDat(lmBaseAttCode)
    
    smPledgeType = Trim$(attBase_rst!attPledgeType)
    'Determine if the agreement we are modelling from is a good candidate
    If Not mIsAgrmtOKToModelFrom(attBase_rst!attPledgeType) Then
        Exit Function
    End If
    
    mGenOK
    mGetBaseAgreementInfo = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFastAdd-mGetBaseAgreementInfo"
End Function

Private Function mDateOverlap(lAttCode As Long, lOnAir As Long, lOffAir As Long, lDropDate As Long) As Integer
    Dim sOnAir As String
    Dim sOffAir As String
    Dim sDropDate As String
    Dim sEndDate As String
    Dim lEndDate As Long
    Dim ilUpper As Integer
    Dim ilRet As Integer
    
    'Test if Agreement date overlap- if so disallow agreement being saved by returning TRUE
    On Error GoTo ErrHand
    If lDropDate < lOffAir Then
        lEndDate = lDropDate
    Else
        lEndDate = lOffAir
    End If
    On Error GoTo ErrHand
    SQLQuery = "SELECT attCode, attOnAir, attOffAir, attDropDate FROM att"
    SQLQuery = SQLQuery + " WHERE (attShfCode = " & imAddShttCode & " AND attVefCode = " & imVefCode & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        'Test dates
        If (lAttCode <> rst!attCode) Then
            sOnAir = Format$(rst!attOnAir, "mm/dd/yyyy")
            sOffAir = Format$(rst!attOffAir, "mm/dd/yyyy")
            sDropDate = Format$(rst!attDropDate, "mm/dd/yyyy")
            If DateValue(gAdjYear(sDropDate)) < DateValue(gAdjYear(sOffAir)) Then
                sEndDate = sDropDate
            Else
                sEndDate = sOffAir
            End If
            If (lEndDate >= DateValue(gAdjYear(sOnAir))) And (DateValue(gAdjYear(sEndDate)) >= lOnAir) Then
                ilUpper = UBound(tmOverlapInfo)
                tmOverlapInfo(ilUpper).lAttCode = rst!attCode
                tmOverlapInfo(ilUpper).lOnAirDate = DateValue(gAdjYear(sOnAir))
                tmOverlapInfo(ilUpper).lOffAirDate = DateValue(gAdjYear(sOffAir))
                tmOverlapInfo(ilUpper).lDropDate = DateValue(gAdjYear(sDropDate))
                tmOverlapInfo(ilUpper).iShfCode = imAddShttCode
                ReDim Preserve tmOverlapInfo(0 To ilUpper + 1) As AGMNTOVERLAPINFO
                'mDateOverlap = True
                'mGenOK
                'Exit Function
            End If
        End If
        rst.MoveNext
    Wend
    mDateOverlap = False
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFastAdd-mDateOverlap"
    mDateOverlap = True
    Exit Function
End Function

Private Function mAddCPTT(iNewRec As Integer, sOnAir As String, sOffAir As String, sDropDate As String) As Integer
    'This module is for adding/deleting CPTT records
    Dim sLLD As String 'Last Log date
    Dim lSvCPTTStart As Long
    Dim lSvCPTTEnd As Long
    Dim lCPTTStart As Long
    Dim lCPTTEnd As Long
    Dim lSDate As Long
    Dim lEDate As Long
    Dim lDate As Long
    Dim iCycle As Integer
    Dim sTime As String
    Dim iWkDay As Integer
    Dim sMsg As String
    Dim ilRet As Integer
    Dim slCallLetters As String
    Dim slVehicleName As String
    Dim rst_TestWk As ADODB.Recordset
    Dim temp_rst As ADODB.Recordset
    Dim slServiceAgreement As String
    Dim slStr As String
    
    On Error GoTo ErrHand
        
    sTime = Format("12:00AM", "hh:mm:ss")
    sMsg = ""
    
    If DateValue(gAdjYear(sOnAir)) = DateValue("1/1/1970") Then
        mAddCPTT = True
        Exit Function
    End If
    
    If iNewRec Then
        SQLQuery = "SELECT vpfLLD, vpfLNoDaysCycle"
        SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
        SQLQuery = SQLQuery + " WHERE (vpfvefKCode =" & imVefCode & ")"
        
        Set rst = gSQLSelectCall(SQLQuery)
        If rst.EOF Then
            mAddCPTT = True
            Exit Function
        End If
        If IsNull(rst!vpfLLD) Then
            mAddCPTT = True
            Exit Function
        End If
        If Not gIsDate(rst!vpfLLD) Then
            'sLLD = "1/1/1970"
            'iWkDay = vbMonday  'Monday
            mAddCPTT = True
            Exit Function
        Else
            sLLD = Format$(rst!vpfLLD, "mm/dd/yyyy")
            If Trim$(sLLD) = "" Then
                mAddCPTT = True
                Exit Function
            End If
            iWkDay = Weekday(Format$(DateValue(gAdjYear(sLLD)) + 1, "m/d/yyyy"))
        End If
        
        '6/7/19
        slServiceAgreement = "N"
        slStr = "Select attServiceAgreement from att where attCode = " & lmAttCode
        Set temp_rst = gSQLSelectCall(slStr)
        If temp_rst.EOF = False Then
            slServiceAgreement = temp_rst!attServiceAgreement
        End If
        
        
        iCycle = 7
        If DateValue(gAdjYear(sOnAir)) <= DateValue(gAdjYear(sLLD)) Then
            If DateValue(gAdjYear(sDropDate)) < DateValue(gAdjYear(sLLD)) Then
                'Add CPTT between OnAir And DropDate
                lSDate = DateValue(gAdjYear(sOnAir))
                lEDate = DateValue(gAdjYear(sDropDate))
            Else
                'Add CPTT between OnAir and LLD
                lSDate = DateValue(gAdjYear(sOnAir))
                lEDate = DateValue(gAdjYear(sLLD))
            End If
            lSDate = DateValue(gObtainPrevMonday(gAdjYear(Format$(lSDate, "m/d/yy"))))
            slCallLetters = gGetCallLettersByShttCode(imAddShttCode)
            slVehicleName = gGetVehNameByVefCode(imVefCode)
            If lSDate <= lEDate Then
                sMsg = "Added weeks: " & Format$(lSDate, sgShowDateForm) & "-" & Format$(lEDate, sgShowDateForm)
                For lDate = lSDate To lEDate Step iCycle
                    'D.S. 10/25/04
                    'Before we add the new cptt recs we have to clean out any old for that week
                    SQLQuery = "DELETE FROM cptt WHERE"
                    SQLQuery = SQLQuery & " cpttVefCode = " & imVefCode
                    SQLQuery = SQLQuery & " And cpttShfCode = " & imAddShttCode
                    SQLQuery = SQLQuery & " And cpttStartDate = " & "'" & Format$(lDate, sgSQLDateForm) & "'"
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "FastAddVerbose.Txt", "FastAdd-mAddCPTT"
                        mAddCPTT = False
                        Exit Function
                    End If
                    gLogMsg "Deleting CPTT1: " & slCallLetters & " running on: " & slVehicleName & " for the week of: " & Format$(lDate, "mm/dd/yyyy"), "FastAddVerbose.Txt", False
                
                    '3/31/14: Check if LST exist
                    SQLQuery = "SELECT Count(lstCode) From LST"
                    SQLQuery = SQLQuery & " WHERE (lstLogVefCode = " & imVefCode
                    SQLQuery = SQLQuery & " AND (lstLogDate >= " & "'" & Format(lDate, sgSQLDateForm) & "'"
                    SQLQuery = SQLQuery & " AND  lstLogDate <= " & "'" & Format(lDate + 6, sgSQLDateForm) & "'))"
                    Set rst_Lst = gSQLSelectCall(SQLQuery)
                    
                    If rst_Lst(0).Value > 0 Then
                
                        SQLQuery = "INSERT INTO cptt(cpttAtfCode, cpttShfCode, cpttVefCode, "
                        SQLQuery = SQLQuery & "cpttCreateDate, cpttStartDate, "
                        SQLQuery = SQLQuery & "cpttStatus, cpttUsfCode)"
                        SQLQuery = SQLQuery & " VALUES "
                        SQLQuery = SQLQuery & "(" & lmAttCode & ", " & imAddShttCode & ", " & imVefCode & ", "
                        SQLQuery = SQLQuery & "'" & Format$(smCurDate, sgSQLDateForm) & "', '" & Format(lDate, sgSQLDateForm) & "', "
                        If slServiceAgreement = "Y" Then
                            SQLQuery = SQLQuery & "" & 1 & ", " & igUstCode & ")"
                        Else
                            SQLQuery = SQLQuery & "" & 0 & ", " & igUstCode & ")"
                        End If
                        'cnn.Execute SQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/11/16: Replaced GoSub
                            'GoSub ErrHand:
                            Screen.MousePointer = vbDefault
                            gHandleError "FastAddVerbose.Txt", "FastAdd-mAddCPTT"
                            mAddCPTT = False
                            Exit Function
                        End If
                        gFileChgdUpdate "cptt.mkd", True
                        
                        'If we've added new agreements that are either Univision or Web type so show an alert
                        If (imWebType = True) Or (imUnivisionType = True) Then
                            ilRet = gAlertAdd("R", "S", imVefCode, Format(lDate, sgSQLDateForm))
                        End If
                        If lDate + 6 <= lEDate Then
                            gSetStationSpotBuilder "F", imVefCode, imAddShttCode, lDate, lDate + 6
                        Else
                            gSetStationSpotBuilder "F", imVefCode, imAddShttCode, lDate, lEDate
                        End If
                    End If
                Next lDate
            End If
        End If
    End If
    If sMsg <> "" Then
        gLogMsg smStationName & " " & sMsg, "FastAddVerbose.Txt", False
    End If
    mAddCPTT = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFastAdd-mAddCPTT"
    mAddCPTT = False
End Function

Private Sub mBuildAgreements()
    Dim slStart As String
    Dim slEnd As String
    Dim slAgmntStart As String
    Dim slAgmntEnd As String
    Dim llEarliestDate As Long
    Dim llLatestDate As Long
    Dim slFdStTime As String
    Dim slFdEdTime As String
    Dim slPdStTime As String
    Dim slPdEdTime As String
    Dim llAttCode As Long
    Dim llLatestAttCode As Long
    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim ilRet As Integer
    Dim sOnAir As String
    Dim sOffAir As String
    Dim sDropDate As String
    Dim sEndDate As String
    Dim ilOverlap As Integer
    Dim ilGenWarningMsg As Integer
    Dim ilFound As Integer
    'Att fields
    Dim ilLoad As Integer
    Dim ilComp As Integer
    Dim ilBarCode As Integer
    Dim slNotice As String
    Dim ilCarryCmml As Integer
    Dim ilNoCDs As Integer
    Dim slACName As String
    Dim slACPhone As String
    Dim slGenLog As String
    Dim slGenCP As String
    Dim slGenOther As String
    Dim ilPrintCP As Integer
    Dim slComments As String
    Dim llAgreementID As Long
    Dim ilSigned As Integer
    Dim slSignDate As String
    Dim slWebPW As String
    Dim slWebEMail As String
    Dim ilSendLogEMail As Integer
    Dim slSuppressNotice As String
    Dim slNCR As String         '7-7-09
    Dim slFormerNCR As String
    Dim slExportToMarketron As String
    Dim ilMktRepUstCode As Integer
    Dim ilServRepUstCode As Integer
    Dim slVehProgStartTime As String
    Dim slVehProgEndTime As String
    Dim slCDStartTime As String
    Dim llTemp As Long
    
    Dim slLabelID As String
    Dim slLabelShipInfo As String
    Dim slMulticast As String
    Dim slRadarClearType As String
    Dim llArttCode As Long
    Dim slForbidSplitLive As String
    Dim llXDReceiverID As Long
    Dim slIDCReceiverID As String
    Dim slVoiceTracked As String
    Dim slAudioDelivery As String
    Dim slServiceAgreement As String
    Dim slXDSSendNotCarry As String
    
    Dim slPledgeByEvent As String
   
    ReDim llAttCodeRef(0 To 0) As Long
    Dim llAtt As Long
    
    Dim llGroupID As Long
    
    Dim att_rst As ADODB.Recordset
    Dim shtt_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    smCurDate = Format(gNow(), "mm/dd/yyyy")
    smCurTime = Format(gNow(), sgShowTimeWSecForm)
    slStart = gAdjYear(Trim$(txtStartDate.Text))
    
    If txtEndDate.Text = "" Then
        slEnd = "12/31/2069"
    Else
        slEnd = gAdjYear(Trim$(txtEndDate.Text))
    End If
    
    '6/30/20: Moved from below
    If Not mAdjOverlapAgmnts(DateValue(gAdjYear(slStart)), DateValue(gAdjYear(slEnd)), DateValue(gAdjYear(slEnd))) Then
        Exit Sub
    End If
    
    slPledgeByEvent = mGetPledgeByEvent()
    
    Screen.MousePointer = vbHourglass
     ''Adding a new agreement
     'For ilLoop = 0 To lbcStations(1).ListCount - 1 Step 1
     '   imAddShttCode = lbcStations(1).ItemData(ilLoop)
     '   smStationName = lbcStations(1).List(ilLoop)
     For ilLoop = grdInclude.FixedRows To grdInclude.Rows - 1 Step 1
        ProgressBar2.Value = ilLoop * (100 / (grdInclude.Rows))
        
        If grdInclude.TextMatrix(ilLoop, SHTTCODEINDEX) <> "" Then
            imAddShttCode = Val(grdInclude.TextMatrix(ilLoop, SHTTCODEINDEX))
            smStationName = grdInclude.TextMatrix(ilLoop, CALLLETTERSINDEX)
            ReDim llAttCodeRef(0 To 0) As Long
            For ilOverlap = LBound(tmOverlapInfo) To UBound(tmOverlapInfo) - 1 Step 1
                If imAddShttCode = tmOverlapInfo(ilOverlap).iShfCode Then
                    llAttCodeRef(UBound(llAttCodeRef)) = tmOverlapInfo(ilOverlap).lAttCode
                    ReDim Preserve llAttCodeRef(0 To UBound(llAttCodeRef) + 1) As Long
                End If
            Next ilOverlap
            If UBound(llAttCodeRef) <= LBound(llAttCodeRef) Then
                llEarliestDate = -1
                llLatestAttCode = -1
                SQLQuery = "SELECT attCode, attOnAir, attOffAir, attDropDate FROM att"
                SQLQuery = SQLQuery + " WHERE (attShfCode = " & imAddShttCode & " AND attVefCode = " & imVefCode & ")"
                Set rst = gSQLSelectCall(SQLQuery)
                While Not rst.EOF
                    sOnAir = Format$(rst!attOnAir, "mm/dd/yyyy")
                    sOffAir = Format$(rst!attOffAir, "mm/dd/yyyy")
                    sDropDate = Format$(rst!attDropDate, "mm/dd/yyyy")
                    If DateValue(gAdjYear(sDropDate)) < DateValue(gAdjYear(sOffAir)) Then
                        sEndDate = sDropDate
                    Else
                        sEndDate = sOffAir
                    End If
                    If llEarliestDate = -1 Then
                        llEarliestDate = DateValue(gAdjYear(sOnAir))
                        llLatestDate = DateValue(gAdjYear(sEndDate))
                        llLatestAttCode = rst!attCode
                    Else
                        If DateValue(gAdjYear(sOnAir)) < llEarliestDate Then
                            llEarliestDate = DateValue(gAdjYear(sOnAir))
                        End If
                        If DateValue(gAdjYear(sEndDate)) > llLatestDate Then
                            llLatestDate = DateValue(gAdjYear(sEndDate))
                            llLatestAttCode = rst!attCode
                        End If
                    End If
                    rst.MoveNext
                Wend
                llAttCodeRef(UBound(llAttCodeRef)) = llLatestAttCode
                ReDim Preserve llAttCodeRef(0 To UBound(llAttCodeRef) + 1) As Long
            End If
            'This section only applies to modeling from agreements that were of website type
            slWebEMail = ""
            slWebPW = ""
            ilGenWarningMsg = False
            If imWebType = True Then
                slWebEMail = "Not Defined"
                slWebPW = gGeneratePassword(4, 1)
                SQLQuery = "Select shttWebEmail, shttWebPW from shtt where shttCode = " & imAddShttCode
                Set shtt_rst = gSQLSelectCall(SQLQuery)
                If Not shtt_rst.EOF Then
                    If Trim$(shtt_rst!shttWebEmail) <> "" Then
                        slWebEMail = Trim$(shtt_rst!shttWebEmail)
    '                Else
    '                    slWebEmail = "Not Defined"
                    End If
                    If Trim$(shtt_rst!shttWebPW) <> "" Then
                        slWebPW = Trim$(shtt_rst!shttWebPW)
    '                Else
    '                    slWebPW = gGeneratePassword(4, 1)
                    End If
                End If
                        
                If slWebEMail = "Not Defined" Or Trim$(shtt_rst!shttWebPW) = "" Then
                    ilGenWarningMsg = True
                    If slWebEMail = "Not Defined" Then
                        slWebEMail = smAdminEMail
                    End If
                End If
            End If
            
            For llAtt = 0 To UBound(llAttCodeRef) - 1 Step 1
                slAgmntStart = slStart
                slAgmntEnd = slEnd
                ilLoad = attBase_rst!attLoad
                ilComp = attBase_rst!attComp
                ilBarCode = attBase_rst!attBarCode
                slNotice = attBase_rst!attNotice
                ilCarryCmml = attBase_rst!attCarryCmml
                ilNoCDs = attBase_rst!attNoCDs
                slACName = attBase_rst!attACName
                slACPhone = attBase_rst!attACPhone
                slGenLog = attBase_rst!attGenLog
                slGenCP = attBase_rst!attGenCP
                slGenOther = attBase_rst!attGenOther
                ilPrintCP = attBase_rst!attPrintCP
                slComments = attBase_rst!attComments
                llAgreementID = attBase_rst!attAgreementID
                ilSigned = attBase_rst!attSigned
                slSignDate = Format$(attBase_rst!attSignDate, sgShowDateForm)
                ilSendLogEMail = attBase_rst!attSendLogEmail
                slSuppressNotice = attBase_rst!attSuppressNotice
                slNCR = "N"         '7-7-09 default to compliant agreement
                slFormerNCR = "N"       '7-7-09 default to not a former non-compliant agreement
                slExportToMarketron = "N"
                
                'D.S. Added 10/25/10
                slLabelID = ""
                slLabelShipInfo = ""
                slMulticast = ""
                slRadarClearType = ""
                llArttCode = 0
                slForbidSplitLive = "N"
                llXDReceiverID = 0
                slVoiceTracked = "N"
                slIDCReceiverID = ""
                slAudioDelivery = "N"
                slServiceAgreement = "N"
                slXDSSendNotCarry = "N"
                
                ilMktRepUstCode = 0
                slVehProgStartTime = "12:00:00 AM"
                slVehProgEndTime = "12:00:00 AM"
                'ReDim tlDat(0 To 0) As DAT
                'ReDim tmDat(0 To 0) As DAT
                If llAttCodeRef(llAtt) > 0 Then
                    SQLQuery = "SELECT * FROM att"
                    SQLQuery = SQLQuery + " WHERE (attCode = " & llAttCodeRef(llAtt) & ")"
                    Set rst = gSQLSelectCall(SQLQuery)
                    If Not rst.EOF Then
                        ilLoad = rst!attLoad
                        ilComp = rst!attComp
                        ilBarCode = rst!attBarCode
                        slNotice = rst!attNotice
                        ilCarryCmml = rst!attCarryCmml
                        ilNoCDs = rst!attNoCDs
                        slACName = rst!attACName
                        slACPhone = rst!attACPhone
                        slGenLog = rst!attGenLog
                        slGenCP = rst!attGenCP
                        slGenOther = rst!attGenOther
                        ilPrintCP = rst!attPrintCP
                        slComments = rst!attComments
                        llAgreementID = rst!attAgreementID
                        ilSigned = rst!attSigned
                        slSignDate = Format$(rst!attSignDate, sgShowDateForm)
                        If (imWebType = True) And (Trim$(rst!attWebPW) <> "") Then
                            slWebPW = rst!attWebPW
                            ilGenWarningMsg = False
                        End If
                        If (imWebType = True) And (Trim$(rst!attWebEmail) <> "") Then
                            slWebEMail = rst!attWebEmail
                            ilGenWarningMsg = False
                        End If
                        ilSendLogEMail = rst!attSendLogEmail
                        slSuppressNotice = rst!attSuppressNotice
                        slNCR = rst!attNCR        '7-7-09
                        slFormerNCR = rst!attFormerNCR      '7-7-09
                        slExportToMarketron = rst!attExportToMarketron
                        sOnAir = Format$(rst!attOnAir, "mm/dd/yyyy")
                        sOffAir = Format$(rst!attOffAir, "mm/dd/yyyy")
                        sDropDate = Format$(rst!attDropDate, "mm/dd/yyyy")
                        If DateValue(gAdjYear(sDropDate)) < DateValue(gAdjYear(sOffAir)) Then
                            sEndDate = sDropDate
                        Else
                            sEndDate = sOffAir
                        End If
                        If DateValue(gAdjYear(sEndDate)) > DateValue(gAdjYear(slStart)) Then
                            If DateValue(gAdjYear(sOnAir)) > DateValue(gAdjYear(slStart)) Then
                                slAgmntStart = sOnAir
                            End If
                            If DateValue(gAdjYear(sEndDate)) < DateValue(gAdjYear(slEnd)) Then
                                slAgmntEnd = sEndDate
                            End If
                        End If
                        slLabelID = rst!attLabelID
                        slLabelShipInfo = rst!attLabelShipInfo
                        slMulticast = rst!attMulticast
                        slRadarClearType = rst!attRadarClearType
                        llArttCode = rst!attArttCode
                        slForbidSplitLive = rst!attForbidSplitLive
                        llXDReceiverID = rst!attXDReceiverId
                        slServiceAgreement = rst!attServiceAgreement
                        slXDSSendNotCarry = rst!attXDSSendNotCarry
                        slVoiceTracked = rst!attVoiceTracked
                        slAudioDelivery = rst!attAudioDelivery
                        slIDCReceiverID = rst!attIDCReceiverID
                        ilMktRepUstCode = rst!attMktRepUstCode
                        ilServRepUstCode = rst!attServRepUstCode
                        slVehProgStartTime = Format$(rst!attVehProgStartTime, sgShowTimeWSecForm)
                        slVehProgEndTime = Format$(rst!attVehProgEndTime, sgShowTimeWSecForm)
                        'D.S. 09/23/04 Commented out mGetDat call and If statement below
                        'I have no idea why you would want to get the DAT records from the old
                        'agreements when we are about to replace them with the ones from the modeled
                        'agreement, but I could be wrong!
                        'ilRet = mGetDat(llAttCodeRef(llAtt), tlDat())
                        'If Not ilRet Then
                        '    ReDim tlDat(0 To 0) As DAT
                        'End If
                    End If
                End If
                
                '11/5/14:
                If rbcMulticast(1).Value Then
                    llGroupID = gGetStaMulticastGroupID(imAddShttCode)
                    If llGroupID > 0 Then
                        slMulticast = "Y"
                    End If
                End If
                
                ilRet = gDetermineAgreementTimes(imAddShttCode, imVefCode, Format$(slAgmntStart, "m/d/yy"), Format$(slAgmntEnd, "m/d/yy"), Format$(slAgmntEnd, "m/d/yy"), slCDStartTime, slVehProgStartTime, slVehProgEndTime)
                           
                            
                slWebPW = gFixQuote(slWebPW)
                slWebEMail = gFixQuote(slWebEMail)
                slACName = gFixQuote(slACName)
                slComments = gFixQuote(slComments)
                'D.S. 8/2/05
                llTemp = gFindAttHole()
                If llTemp = -1 Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                SQLQuery = "INSERT INTO att(attCode, attShfCode, attVefCode, attAgreeStart, "
                SQLQuery = SQLQuery & "attAgreeEnd, attOnAir, attOffAir, attSigned, attSignDate, "
                SQLQuery = SQLQuery & "attLoad, attTimeType, attComp, attBarCode, attDropDate, "
                SQLQuery = SQLQuery & "attUsfCode, attEnterDate, attEnterTime, attNotice, "
                SQLQuery = SQLQuery & "attCarryCmml, attNoCDs, attSendTape, attACName, "
                SQLQuery = SQLQuery & "attACPhone, attGenLog, attGenCP, attPostingType, attPrintCP, "
                SQLQuery = SQLQuery & "attExportType, attLogType, attPostType, attWebPW, attWebEmail, "
                SQLQuery = SQLQuery & "attSendLogEMail, attSuppressNotice, attComments, attGenOther, "
                SQLQuery = SQLQuery & "attStartTime,attNCR,attFormerNCR, "
                SQLQuery = SQLQuery & "attLabelID, attLabelShipInfo,attMulticast, attRadarClearType, attArttCode, attForbidSplitLive, attXDReceiverID, attVoiceTracked, "
                
                SQLQuery = SQLQuery & "attWebInterface, "
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
                SQLQuery = SQLQuery & "attXDSSendNotCarry, "
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
                SQLQuery = SQLQuery & "(" & llTemp & ", " & imAddShttCode & ", " & imVefCode & ", '" & Format$(slAgmntStart, sgSQLDateForm) & "', '" & Format$(slAgmntEnd, sgSQLDateForm) & "', "
                SQLQuery = SQLQuery & "'" & Format$(slAgmntStart, sgSQLDateForm) & "', '" & Format$(slAgmntEnd, sgSQLDateForm) & "', " & ilSigned & ", "
                SQLQuery = SQLQuery & "'" & Format$(slSignDate, sgSQLDateForm) & "', " & ilLoad & ", " & attBase_rst!attTimeType & ", "
                SQLQuery = SQLQuery & ilComp & ", " & ilBarCode & ", '" & Format$(slAgmntEnd, sgSQLDateForm) & "', "
                SQLQuery = SQLQuery & igUstCode & ", '" & Format$(smCurDate, sgSQLDateForm) & "', '" & Format$(smCurTime, sgSQLTimeForm) & "', '" & slNotice & "', "
                SQLQuery = SQLQuery & ilCarryCmml & ", " & ilNoCDs & ", " & attBase_rst!attSendTape & ", '" & slACName & "', "
                SQLQuery = SQLQuery & "'" & slACPhone & "', '" & slGenLog & "', '" & slGenCP & "', " & attBase_rst!attPostingType & ", " & ilPrintCP & ", "
                SQLQuery = SQLQuery & attBase_rst!attExportType & ", " & attBase_rst!attLogType & ", " & attBase_rst!attPostType & ", '" & slWebPW & "', '" & slWebEMail & "', "
                SQLQuery = SQLQuery & ilSendLogEMail & ", '" & slSuppressNotice & "', '" & slComments & "', '" & slGenOther & "', '" & Format$(attBase_rst!attStartTime, sgSQLTimeForm) & "', " & "'" & slNCR & "', '" & slFormerNCR & "', "
                SQLQuery = SQLQuery & "'" & slLabelID & "', '" & slLabelShipInfo & "', '" & slMulticast & "', '" & slRadarClearType & "', " & llArttCode & ", '" & slForbidSplitLive & "'," & llXDReceiverID & ", '" & slVoiceTracked & "', "
                SQLQuery = SQLQuery & "'" & attBase_rst!attWebInterface & "', "
                SQLQuery = SQLQuery & "'" & "N" & "', "
                SQLQuery = SQLQuery & ilMktRepUstCode & ", "
                SQLQuery = SQLQuery & ilServRepUstCode & ", "
                SQLQuery = SQLQuery & "'" & Format$(slVehProgStartTime, sgSQLTimeForm) & "', "
                SQLQuery = SQLQuery & "'" & Format$(slVehProgEndTime, sgSQLTimeForm) & "', "
                SQLQuery = SQLQuery & "'" & attBase_rst!attExportToWeb & "', "
                SQLQuery = SQLQuery & "'" & attBase_rst!attExportToUnivision & "', "
                SQLQuery = SQLQuery & "'" & attBase_rst!attExportToMarketron & "', "
                SQLQuery = SQLQuery & "'" & attBase_rst!attExportToCBS & "', "
                SQLQuery = SQLQuery & "'" & attBase_rst!attExportToClearCh & "', "
                SQLQuery = SQLQuery & "'" & attBase_rst!attPledgeType & "', "
                SQLQuery = SQLQuery & attBase_rst!attNoAirPlays & ", "
                SQLQuery = SQLQuery & attBase_rst!attDesignVersion & ", "
                SQLQuery = SQLQuery & "'" & slIDCReceiverID & "', "
                SQLQuery = SQLQuery & "'" & "N" & "', "
                SQLQuery = SQLQuery & "'" & slAudioDelivery & "', "
                SQLQuery = SQLQuery & "'" & attBase_rst!attExportToJelli & "', "
                '3/23/15: Add Send Delays to XDS
                SQLQuery = SQLQuery & "'" & attBase_rst!attSendDelayToXDS & "', "
                SQLQuery = SQLQuery & "'" & slXDSSendNotCarry & "', "
                SQLQuery = SQLQuery & "'" & slServiceAgreement & "', "
                '4-3-19
                SQLQuery = SQLQuery & "'" & attBase_rst!attExcludeFillSpot & "', "
                SQLQuery = SQLQuery & "'" & attBase_rst!attExcludeCntrTypeQ & "', "
                SQLQuery = SQLQuery & "'" & attBase_rst!attExcludeCntrTypeR & "', "
                SQLQuery = SQLQuery & "'" & attBase_rst!attExcludeCntrTypeT & "', "
                SQLQuery = SQLQuery & "'" & attBase_rst!attExcludeCntrTypeM & "', "
                SQLQuery = SQLQuery & "'" & attBase_rst!attExcludeCntrTypeS & "', "
                SQLQuery = SQLQuery & "'" & attBase_rst!attExcludeCntrTypeV & "', "
                                
                SQLQuery = SQLQuery & "'" & "" & "'"
                SQLQuery = SQLQuery & ")"
        
                cnn.BeginTrans
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/11/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "FastAddVerbose.Txt", "FastAdd-mBuildAgreements"
                    cnn.RollbackTrans
                    Exit Sub
                End If
                
                If llTemp = 0 Then
                    SQLQuery = "SELECT MAX(attCode) from att"
                    Set att_rst = gSQLSelectCall(SQLQuery)
                    llAttCode = att_rst(0).Value
                Else
                    llAttCode = llTemp
                End If
                
                'D.S. 09/23/04 commented out if else below and added the single call to mInsertDat below
                'If UBound(tlDat) > LBound(tlDat) Then
                '    ilRet = mInsertDAT(llAttCode, tlDat())
                'Else
                '    ilRet = mInsertDAT(llAttCode, tgDat())
                'End If
                If slPledgeByEvent <> "Y" Then
                    ilRet = mInsertDAT(llAttCode)
                Else
                    ilRet = mAddPet(imVefCode, imAddShttCode, llAttCode, llAttCodeRef(llAtt))
                End If
                '7701  8599
                
                'SQLQuery = "insert into VAT_Vendor_Agreement (vatAttCode,vatWvtVendorId) ( select " & llAttCode & ", vatWvtVendorId from VAT_Vendor_Agreement where vatAttCode = " & lmBaseAttCode & ")"
                'TTP 11070 - Fast Add Turn on Delivery Service for Stop/Started Agreements
                '1. None: works like having the existing "copy delivery service" checkbox unchecked
                'If llAttCodeRef(llAtt) > 0 Then
                If llAttCodeRef(llAtt) > 0 And cboCopyDeliveryService.ListIndex < 2 Then
                    SQLQuery = "insert into VAT_Vendor_Agreement (vatAttCode,vatWvtVendorId) ( select " & llAttCode & ", vatWvtVendorId from VAT_Vendor_Agreement where vatAttCode = " & llAttCodeRef(llAtt) & ")"
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "FastAddVerbose.Txt", "FastAdd-mBuildAgreements"
                        Exit Sub
                    End If
                
                'TTP 11070 - Fast Add Turn on Delivery Service for Stop/Started Agreements
                'ElseIf ckcDelivery.Value = vbChecked Then
                ElseIf cboCopyDeliveryService.ListIndex = 1 Then
                    '2. Copy delivery service of modeled from agreement to new agreements for previously unaffiliated stations,
                    ' and for stations with current agreements, use the vendor settings from the current agreement: what it currently does with the "copy delivery service" checkbox checked
                    SQLQuery = "insert into VAT_Vendor_Agreement (vatAttCode,vatWvtVendorId) ( select " & llAttCode & ", vatWvtVendorId from VAT_Vendor_Agreement where vatAttCode = " & lmBaseAttCode & ")"
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "FastAddVerbose.Txt", "FastAdd-mBuildAgreements"
                        Exit Sub
                    End If
                
                'TTP 11070 - Fast Add Turn on Delivery Service for Stop/Started Agreements
                ElseIf cboCopyDeliveryService.ListIndex = 2 Then
                    '3. Copy delivery service of modeled from agreement to all new agreements
                    '(what USRN is asking for - this is for a case where a vehicle is entirely switching to a new service)
                    SQLQuery = "insert into VAT_Vendor_Agreement (vatAttCode,vatWvtVendorId) ( select " & llAttCode & ", vatWvtVendorId from VAT_Vendor_Agreement where vatAttCode = " & lmBaseAttCode & ")"
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "FastAddVerbose.Txt", "FastAdd-mBuildAgreements"
                        Exit Sub
                    End If
                End If
                
                lmAttCode = llAttCode
                
                ilRet = mSetUsedForAtt(imAddShttCode, slAgmntEnd)
                ilRet = mCheckHistDate(imAddShttCode, slAgmntStart)
                ilRet = mAddCPTT(True, slAgmntStart, slAgmntEnd, slAgmntEnd)
                If ilRet Then
                    cnn.CommitTrans
                End If
            Next llAtt
            If ilGenWarningMsg Then
                gLogMsg "Warning:  No Password was defined for station " & smStationName, "FastAddVerbose.Txt", False
                gLogMsg "Warning:  Password " & slWebPW & " was generated for station: " & smStationName, "FastAddVerbose.Txt", False
                gLogMsg "Warning:  No Email address was defined for station" & smStationName, "FastAddVerbose.Txt", False
                gLogMsg "Warning:  " & slWebEMail & " was used until a valid address can be defined for: " & smStationName, "FastAddVerbose.Txt", False
                SQLQuery = "Update SHTT Set shttWebPW = '" & slWebPW & "', shttWebEmail = '" & slWebEMail & "' Where shttCode = " & imAddShttCode
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/11/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "FastAddVerbose.Txt", "FastAdd-mBuildAgreements"
                    Exit Sub
                End If
            End If
        End If
    Next ilLoop
    '6/30/20: Moved above
    'If Not mAdjOverlapAgmnts(DateValue(gAdjYear(slStart)), DateValue(gAdjYear(slEnd)), DateValue(gAdjYear(slEnd))) Then
    '    Exit Sub
    'End If
    '11/26/17
    gFileChgdUpdate "shtt.mkd", True
    
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFastAdd-mBuildAgreements"
End Sub

Private Function mPreValidateStations()
    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim slStart As String
    Dim slEnd As String
    Dim slStationName As String
    Dim slVehicleName As String
    Dim ilRet As Integer
    Dim ilOverlap As Integer
    Dim ilVef As Integer
    Dim llRow As Long
    Dim llToRow As Long
        
    On Error GoTo ErrHand
    ReDim tgStaNameAndCode(0 To 0)
    ReDim tmFastAddAttCount(0 To 0) As FASTADDATTCOUNT
    
    igFastAddContinue = True
    slStart = Trim$(txtStartDate.Text)
    If txtEndDate.Text = "" Then
        slEnd = "12/31/2069"
    Else
        slEnd = Trim$(txtEndDate.Text)
    End If
    ilVef = gBinarySearchVef(CLng(imVefCode))
    If ilVef <> -1 Then
        slVehicleName = Trim$(tgVehicleInfo(ilVef).sVehicle)
    Else
        slVehicleName = "VefCode: " & imVefCode
    End If
    ReDim tmOverlapInfo(0 To 0) As AGMNTOVERLAPINFO
    ilIdx = 0
    'Check to see if any of the selected stations has pre-existing agreements
    'that have overlapping dates
'    For ilLoop = 0 To lbcStations(1).ListCount - 1 Step 1
'        imAddShttCode = lbcStations(1).ItemData(ilLoop)
'        slStationName = lbcStations(1).List(ilLoop)
'        If mDateOverlap(lmBaseAttCode, DateValue(gAdjYear(slStart)), DateValue(gAdjYear(slEnd)), DateValue(gAdjYear(slEnd))) Then
''            Screen.MousePointer = vbDefault
''            'gMsgBox slStationName & " On/Off dates overlap with previously defined Agreement, Can't Save", vbOKOnly, "Save"
''            gLogMsg "Warning: " & slStationName & " On/Off dates overlap with previously defined Agreement, Can't Save", "FastAddVerbose.Txt", False
''            tgStaNameAndCode(ilIdx).iStationCode = imAddShttCode
''            tgStaNameAndCode(ilIdx).sStationName = slStationName
''            tgStaNameAndCode(ilIdx).sInfo = " On/Off dates overlap with previously defined Agreement"
''            ilIdx = ilIdx + 1
''            ReDim Preserve tgStaNameAndCode(0 To ilIdx)
'            igFastAddContinue = False
'            Exit Function
'        End If
'    Next ilLoop
    For llRow = grdInclude.FixedRows To grdInclude.Rows - 1 Step 1
        If grdInclude.TextMatrix(llRow, CALLLETTERSINDEX) <> "" Then
            imAddShttCode = grdInclude.TextMatrix(llRow, SHTTCODEINDEX)
            slStationName = Trim$(grdInclude.TextMatrix(llRow, CALLLETTERSINDEX))
            If mDateOverlap(lmBaseAttCode, DateValue(gAdjYear(slStart)), DateValue(gAdjYear(slEnd)), DateValue(gAdjYear(slEnd))) Then
                igFastAddContinue = False
                Exit Function
            End If
        End If
    Next llRow
    If UBound(tmOverlapInfo) > LBound(tmOverlapInfo) Then
        bmTrmntAgrmnt = True
        'D.S. 5/22/18
        'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
        If TabStrip1.SelectedItem.Index = 2 Then
            ilRet = vbYes
        Else
            ilRet = gMsgBox(slVehicleName & ": Agreement Dates Overlap, Continue with Save by Terminating Overlapped Agreement(s)?", vbYesNo)
        End If
        
        If ilRet = vbNo Then
            Screen.MousePointer = vbDefault
            bmTrmntAgrmnt = False
''            'gMsgBox slStationName & " On/Off dates overlap with previously defined Agreement, Can't Save", vbOKOnly, "Save"
''            gLogMsg "Warning: " & slStationName & " On/Off dates overlap with previously defined Agreement, Can't Save", "FastAddVerbose.Txt", False
''            tgStaNameAndCode(ilIdx).iStationCode = imAddShttCode
''            tgStaNameAndCode(ilIdx).sStationName = slStationName
''            tgStaNameAndCode(ilIdx).sInfo = " On/Off dates overlap with previously defined Agreement"
''            ilIdx = ilIdx + 1
''            ReDim Preserve tgStaNameAndCode(0 To ilIdx)
            'For ilLoop = 0 To lbcStations(1).ListCount - 1 Step 1
            '    imAddShttCode = lbcStations(1).ItemData(ilLoop)
            '    slStationName = lbcStations(1).List(ilLoop)
            For llRow = grdInclude.FixedRows To grdInclude.Rows - 1 Step 1
                If grdInclude.TextMatrix(llRow, CALLLETTERSINDEX) <> "" Then
                    imAddShttCode = grdInclude.TextMatrix(llRow, SHTTCODEINDEX)
                    slStationName = Trim$(grdInclude.TextMatrix(llRow, CALLLETTERSINDEX))
                    For ilOverlap = LBound(tmOverlapInfo) To UBound(tmOverlapInfo) - 1 Step 1
                        If tmOverlapInfo(ilOverlap).iShfCode = imAddShttCode Then
                            'gMsgBox slStationName & " On/Off dates overlap with previously defined Agreement, Can't Save", vbOKOnly, "Save"
                            gLogMsg "Warning: " & slVehicleName & "- " & slStationName & " On/Off dates overlap with previously defined Agreement", "FastAddVerbose.Txt", False
                            tgStaNameAndCode(ilIdx).iStationCode = imAddShttCode
                            tgStaNameAndCode(ilIdx).sStationName = slStationName
                            tgStaNameAndCode(ilIdx).sInfo = " On/Off dates overlap with previously defined Agreement"
                            ilIdx = ilIdx + 1
                            ReDim Preserve tgStaNameAndCode(0 To ilIdx)
                        End If
                    Next ilOverlap
                End If
            Next llRow
            'D.S. 5/22/18
            ReDim tmOverlapInfo(0 To 0)
        End If
    End If
    
    ''Move stations found with overlapping agreements back over to the excluded list
    'For ilLoop = 0 To UBound(tgStaNameAndCode) - 1 Step 1
    '    For ilIdx = 0 To lbcStations(1).ListCount - 1 Step 1
    '        If lbcStations(1).ItemData(ilIdx) = tgStaNameAndCode(ilLoop).iStationCode Then
    '            slStationName = lbcStations(1).List(ilIdx)
    '            gLogMsg "Warning: " & slVehicleName & "- " & slStationName & " is being moved from the Include Stations List to the Exclude Stations List.", "FastAddVerbose.Txt", False
    '            'gMsgBox "Station " & slStationName & " is being moved from the Include Stations List to the Exclude Stations List."
    '            lbcStations(0).AddItem lbcStations(1).List(ilIdx)
    '            lbcStations(0).ItemData(lbcStations(0).NewIndex) = lbcStations(1).ItemData(ilIdx)
    '            lbcStations(1).RemoveItem (ilIdx)
    '            Exit For
    '        End If
    '    Next ilIdx
    'Next ilLoop
    
    llToRow = grdExclude.FixedRows
    For llRow = grdExclude.Rows - 1 To grdExclude.FixedRows Step -1
        If grdExclude.TextMatrix(llRow, CALLLETTERSINDEX) <> "" Then
            llToRow = llToRow + 1
            Exit For
        End If
    Next llRow
    For ilLoop = 0 To UBound(tgStaNameAndCode) - 1 Step 1
        'For ilIdx = 0 To lbcStations(1).ListCount - 1 Step 1
        '    If lbcStations(1).ItemData(ilIdx) = tgStaNameAndCode(ilLoop).iStationCode Then
        '        slStationName = lbcStations(1).List(ilIdx)
        For llRow = grdInclude.Rows - 1 To grdInclude.FixedRows Step -1
            If grdInclude.TextMatrix(llRow, CALLLETTERSINDEX) <> "" Then
                imAddShttCode = grdInclude.TextMatrix(llRow, SHTTCODEINDEX)
                slStationName = Trim$(grdInclude.TextMatrix(llRow, CALLLETTERSINDEX))
                If imAddShttCode = tgStaNameAndCode(ilLoop).iStationCode Then
                    gLogMsg "Warning: " & slVehicleName & "- " & slStationName & " is being moved from the Include Stations List to the Exclude Stations List.", "FastAddVerbose.Txt", False
                    ''gMsgBox "Station " & slStationName & " is being moved from the Include Stations List to the Exclude Stations List."
                    'lbcStations(0).AddItem lbcStations(1).List(ilIdx)
                    'lbcStations(0).ItemData(lbcStations(0).NewIndex) = lbcStations(1).ItemData(ilIdx)
                    'lbcStations(1).RemoveItem (ilIdx)
                    mMoveGridRow grdInclude, grdExclude, llRow, llToRow
                    Exit For
                End If
            End If
        Next llRow
    Next ilLoop
    
    If UBound(tgStaNameAndCode) > 0 Then
        bgFastAddCancelButton = True
        frmFastAddWarning.Show vbModal
    End If
    
    If TabStrip1.SelectedItem.Index = 1 Then
        mGenOK
    End If
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-mPreValidateStations: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    Resume Next
End Function

'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - determine if Gen Button can Enable with click of Multicast choice
Private Sub rbcMulticast_Click(Index As Integer)
    If TabStrip1.SelectedItem.Index = 2 Then
        Exit Sub
    End If
    mGenOK
End Sub

'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = 1 Then
        frcManual.Visible = True
        frcFromFile.Visible = False
        frcManualHeader.Visible = True
    End If
    If TabStrip1.SelectedItem.Index = 2 Then
        frcManual.Visible = False
        frcFromFile.Visible = True
        frcManualHeader.Visible = False
    End If
    mGenOK
End Sub

'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - Quick Find by typing on Exclude Grid
Private Sub Timer1_Timer()
    imClearSearchTime = imClearSearchTime - 1
    If imClearSearchTime < 0 Then
        If txtSearch.Text <> "" Then txtSearch.Text = "": txtSearch.Visible = False
        imClearSearchTime = 0
    End If
End Sub

Private Sub tmcDelay_Timer()
    tmcDelay.Enabled = False
    lbcCreateVehicle.Enabled = False
    Screen.MousePointer = vbHourglass
    mGetStaMark
    Screen.MousePointer = vbDefault
    lbcCreateVehicle.Enabled = True
    mGenOK
End Sub

'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - set caption to help user
Private Sub txtActiveDate_Change()
    If txtActiveDate.Text = "" Then
        cmdAll.Caption = "Add All"
    Else
        cmdAll.Caption = "Add Active"
    End If
End Sub

Private Sub txtActiveDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - Quick Find by typing on Exclude Grid
Private Sub txtSearch_Change()
    imClearSearchTime = 5
    txtSearch.Visible = True
    If txtSearch.Text = "" Then
        txtSearch.Visible = False
    Else
        mFindMatch txtSearch.Text, grdExclude
    End If
End Sub

'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - Quick Find by typing on Exclude Grid
Private Sub txtSearch_GotFocus()
    grdExclude.SetFocus
End Sub

Private Sub txtStartDate_Change()
    If TabStrip1.SelectedItem.Index = 2 Then
        Exit Sub
    End If
    bmCreateChanged = True
    mGenOK
End Sub

Private Sub txtStartDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - reload if Start Date changed
Private Sub txtStartDate_LostFocus()
    If bmCreateChanged = True Then
        tmcDelay_Timer
        bmCreateChanged = False
    End If
End Sub

Private Function mGetDat(llAttCode As Long) As Integer
    Dim ilUpper As Integer
    On Error GoTo ErrHand
    ReDim tmDat(0 To 0) As DAT
    ilUpper = 0
    SQLQuery = "SELECT * "
    SQLQuery = SQLQuery + " FROM dat"
    SQLQuery = SQLQuery + " WHERE (datAtfCode = " & llAttCode & ")"
    SQLQuery = SQLQuery & " ORDER BY datFdStTime"
    Set dat_rst = gSQLSelectCall(SQLQuery)
    
    If Not dat_rst.EOF Then
        While Not dat_rst.EOF
            tmDat(ilUpper).iStatus = 1
            tmDat(ilUpper).lCode = dat_rst!datCode    '(0).Value
            tmDat(ilUpper).lAtfCode = dat_rst!datAtfCode  '(1).Value
            tmDat(ilUpper).iShfCode = dat_rst!datShfCode  '(2).Value
            tmDat(ilUpper).iVefCode = dat_rst!datVefCode  '(3).Value
            'tmDat(ilUpper).iDACode = dat_rst!datDACode    '(4).Value
            tmDat(ilUpper).iFdDay(0) = dat_rst!datFdMon   '(5).Value
            tmDat(ilUpper).iFdDay(1) = dat_rst!datFdTue   '(6).Value
            tmDat(ilUpper).iFdDay(2) = dat_rst!datFdWed   '(7).Value
            tmDat(ilUpper).iFdDay(3) = dat_rst!datFdThu   '(8).Value
            tmDat(ilUpper).iFdDay(4) = dat_rst!datFdFri   '(9).Value
            tmDat(ilUpper).iFdDay(5) = dat_rst!datFdSat   '(10).Value
            tmDat(ilUpper).iFdDay(6) = dat_rst!datFdSun   '(11).Value
            If Second(dat_rst!datFdStTime) = 0 Then
                tmDat(ilUpper).sFdSTime = Format$(CStr(dat_rst!datFdStTime), sgShowTimeWOSecForm)
            Else
                tmDat(ilUpper).sFdSTime = Format$(CStr(dat_rst!datFdStTime), sgShowTimeWSecForm)
            End If
            If Second(dat_rst!datFdEdTime) = 0 Then
                tmDat(ilUpper).sFdETime = Format$(CStr(dat_rst!datFdEdTime), sgShowTimeWOSecForm)
            Else
                tmDat(ilUpper).sFdETime = Format$(CStr(dat_rst!datFdEdTime), sgShowTimeWSecForm)
            End If
            tmDat(ilUpper).iFdStatus = dat_rst!datFdStatus    '(14).Value
            tmDat(ilUpper).iPdDay(0) = dat_rst!datPdMon   '(15).Value
            tmDat(ilUpper).iPdDay(1) = dat_rst!datPdTue   '(16).Value
            tmDat(ilUpper).iPdDay(2) = dat_rst!datPdWed   '(17).Value
            tmDat(ilUpper).iPdDay(3) = dat_rst!datPdThu   '(18).Value
            tmDat(ilUpper).iPdDay(4) = dat_rst!datPdFri   '(19).Value
            tmDat(ilUpper).iPdDay(5) = dat_rst!datPdSat   '(20).Value
            tmDat(ilUpper).iPdDay(6) = dat_rst!datPdSun   '(21).Value
            tmDat(ilUpper).sPdDayFed = dat_rst!datPdDayFed
            If (tmDat(ilUpper).iFdStatus <= 1) Or (tmDat(ilUpper).iFdStatus = 9) Or (tmDat(ilUpper).iFdStatus = 10) Then
                If Second(dat_rst!datPdStTime) = 0 Then
                    tmDat(ilUpper).sPdSTime = Format$(CStr(dat_rst!datPdStTime), sgShowTimeWOSecForm)
                Else
                    tmDat(ilUpper).sPdSTime = Format$(CStr(dat_rst!datPdStTime), sgShowTimeWSecForm)
                End If
                If Second(dat_rst!datPdEdTime) = 0 Then
                    tmDat(ilUpper).sPdETime = Format$(CStr(dat_rst!datPdEdTime), sgShowTimeWOSecForm)
                Else
                    tmDat(ilUpper).sPdETime = Format$(CStr(dat_rst!datPdEdTime), sgShowTimeWSecForm)
                End If
            Else
                tmDat(ilUpper).sPdSTime = ""
                tmDat(ilUpper).sPdETime = ""
            End If
            tmDat(ilUpper).iAirPlayNo = dat_rst!datAirPlayNo
            tmDat(ilUpper).sEstimatedTime = "N" 'dat_rst!datEstimatedTime
            '7/15/14
            tmDat(ilUpper).sEmbeddedOrROS = dat_rst!datEmbeddedOrROS
            ilUpper = ilUpper + 1
            ReDim Preserve tmDat(0 To ilUpper) As DAT
            dat_rst.MoveNext
        Wend
    End If
    mGetDat = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFastAdd-mGetDat"
    mGetDat = False
    Exit Function
End Function

Private Function mInsertDAT(llAttCode As Long) As Integer
    Dim slFdStTime As String
    Dim slFdEdTime As String
    Dim slPdStTime As String
    Dim slPdEdTime As String
    Dim llFdTime As Long
    Dim llPdTime As Long
    Dim ilIdx As Integer
    Dim ilTimeOffset As Integer
    Dim llBaseTime As Long
    Dim llFdNewTime As Long
    Dim llPdNewTime As Long
    Dim ilDayIdx As Integer
    Dim ilOKToConvertTimeZones As Integer
    Dim ilLockedDays(0 To 6) As Integer
        
    On Error GoTo ErrHand
    
    ReDim tlDat(0 To UBound(tmDat)) As DAT
    
    'Copy the contents of tmDat into tlDat so we maintain tmDat integrity. We don't want to update the
    'values of tmDat.  It's only loaded once from the agreement we are modelling from.
    For ilIdx = 0 To UBound(tmDat)
        tlDat(ilIdx) = tmDat(ilIdx)
    Next ilIdx

    If imOKToConvertTimeZones Then
        ilTimeOffset = mGetTimeOffSet()
    End If
    For ilIdx = 0 To UBound(tlDat) - 1 Step 1
        '9/29/11:  This code was added to help determine how to handle the new design on the pledge tab.  The Same rule for specifing Estimate
        '          time will be used to determine if pledge should be adjusted
        '9/30/11:  The rule has been changed to the following: Adjust Feed if the length of time is less then 5 min.
        '                                                      Adjust the Pledge if the length of time is less then 5 min
        ilOKToConvertTimeZones = imOKToConvertTimeZones
        'If (ilOKToConvertTimeZones) And (smPledgeType = "") Then
        '    If tgStatusTypes(tlDat(ilIdx).iFdStatus).iPledged = 1 Then  '1=Delay
        '        llFdTime = gTimeToLong(tlDat(ilIdx).sFdETime, True) - gTimeToLong(tlDat(ilIdx).sFdSTime, False)
        '        llPdTime = gTimeToLong(tlDat(ilIdx).sPdETime, True) - gTimeToLong(tlDat(ilIdx).sPdSTime, False)
        '        If llPdTime > llFdTime Then
        '            ilOKToConvertTimeZones = False
        '        End If
        '    End If
        'End If
        llFdTime = gTimeToLong(tlDat(ilIdx).sFdETime, True) - gTimeToLong(tlDat(ilIdx).sFdSTime, False)
        llPdTime = gTimeToLong(tlDat(ilIdx).sPdETime, True) - gTimeToLong(tlDat(ilIdx).sPdSTime, False)

        If ilOKToConvertTimeZones Then
            'Time and date adjustments are necessary for this type of agreement
            llBaseTime = gTimeToLong(tlDat(ilIdx).sFdSTime, False)
            '9/30/11: Made Feed and Pledge independent
            'llNewTime = llBaseTime + (ilTimeOffset * 3600)
            If (smPledgeType <> "") Or ((smPledgeType = "") And (llFdTime <= AVAIL_OR_DP_TIME)) Then
                llFdNewTime = llBaseTime + (ilTimeOffset * 3600)
            Else
                llFdNewTime = llBaseTime
            End If
            '9/30/11:  Added Pledge
            If (smPledgeType <> "") Then
                llPdNewTime = llBaseTime + (ilTimeOffset * 3600)
            ElseIf ((smPledgeType = "") And (llPdTime <= AVAIL_OR_DP_TIME)) Then
                llPdNewTime = gTimeToLong(tlDat(ilIdx).sPdSTime, False) + (ilTimeOffset * 3600)
            Else
                llPdNewTime = llBaseTime
            End If
            
            'The time offset is causing the spot to fall into the previous day
            If llFdNewTime < 0 Then
                'init all days to be not locked
                For ilDayIdx = 0 To 6 Step 1
                    ilLockedDays(ilDayIdx) = False
                Next ilDayIdx
    
                For ilDayIdx = 0 To 6 Step 1
                    If tlDat(ilIdx).iFdDay(ilDayIdx) = 1 Then
                        If ilDayIdx > 0 Then
                            tlDat(ilIdx).iFdDay(ilDayIdx - 1) = 1
                            ilLockedDays(ilDayIdx - 1) = True
                            If Not ilLockedDays(ilDayIdx) Then
                                tlDat(ilIdx).iFdDay(ilDayIdx) = 0
                            End If
                        Else
                            tlDat(ilIdx).iFdDay(6) = 1
                            ilLockedDays(6) = True
                            If Not ilLockedDays(0) Then
                                tlDat(ilIdx).iFdDay(ilDayIdx) = 0
                            End If
                        End If
                    End If
                Next ilDayIdx
            
            '9/30/11:  Move End If here so that Pledge will be adjusted indepedent of Feed
            End If

                For ilDayIdx = 0 To 6 Step 1
                    ilLockedDays(ilDayIdx) = False
                Next ilDayIdx
                    
                '9/30/11
                'If tlDat(ilIdx).iFdStatus = 0 Then
                If ((smPledgeType <> "") And (tlDat(ilIdx).iFdStatus = 0) And (llFdNewTime < 0)) Or ((smPledgeType = "") And (llPdNewTime < 0)) Then
                    For ilDayIdx = 0 To 6 Step 1
                        If tlDat(ilIdx).iPdDay(ilDayIdx) = 1 Then
                            If ilDayIdx > 0 Then
                                tlDat(ilIdx).iPdDay(ilDayIdx - 1) = 1
                                ilLockedDays(ilDayIdx - 1) = True
                                If Not ilLockedDays(ilDayIdx) Then
                                    tlDat(ilIdx).iPdDay(ilDayIdx) = 0
                                End If
                            Else
                                tlDat(ilIdx).iPdDay(6) = 1
                                ilLockedDays(6) = True
                                If Not ilLockedDays(0) Then
                                    tlDat(ilIdx).iPdDay(ilDayIdx) = 0
                                End If
                            End If
                        End If
                    Next ilDayIdx
                End If
            '9/30/11:  Move End If up so that Pledge will be adjusted indepedent of Feed
            'End If
        
            'The time offset is causing the spot to fall into the next day
            If llFdNewTime > 86400 Then
                'init all days to be not locked
                For ilDayIdx = 0 To 6 Step 1
                    ilLockedDays(ilDayIdx) = False
                Next ilDayIdx
    
                For ilDayIdx = 0 To 6 Step 1
                    If tlDat(ilIdx).iFdDay(ilDayIdx) = 1 Then
                        If ilDayIdx < 6 Then
                            'tlDat(ilIdx).iFdDay(ilDayIdx + 1) = 1
                            If tlDat(ilIdx).iFdDay(ilDayIdx + 1) = 0 Then
                                tlDat(ilIdx).iFdDay(ilDayIdx + 1) = 8
                            End If
                            ilLockedDays(ilDayIdx + 1) = True
                            If Not ilLockedDays(ilDayIdx) Then
                                tlDat(ilIdx).iFdDay(ilDayIdx) = 0
                            End If
                        Else
                            tlDat(ilIdx).iFdDay(0) = 1
                            ilLockedDays(0) = True
                            If Not ilLockedDays(6) Then
                                tlDat(ilIdx).iFdDay(ilDayIdx) = 0
                            End If
                        End If
                    End If
                Next ilDayIdx
                
                For ilDayIdx = 0 To 6 Step 1
                    If tlDat(ilIdx).iFdDay(ilDayIdx) > 0 Then
                        tlDat(ilIdx).iFdDay(ilDayIdx) = 1
                    End If
                Next ilDayIdx
                
            '9/30/11:  Move End If here so that Pledge will be adjusted indepedent of Feed
            End If
            
                For ilDayIdx = 0 To 6 Step 1
                    ilLockedDays(ilDayIdx) = False
                Next ilDayIdx
                
                'If tlDat(ilIdx).iFdStatus = 0 Then
                If ((smPledgeType <> "") And (tlDat(ilIdx).iFdStatus = 0) And (llFdNewTime > 86400)) Or ((smPledgeType = "") And (llPdNewTime > 86400)) Then
                    For ilDayIdx = 0 To 6 Step 1
                        If tlDat(ilIdx).iPdDay(ilDayIdx) = 1 Then
                            If ilDayIdx < 6 Then
                                'tlDat(ilIdx).iPdDay(ilDayIdx + 1) = 1
                                If tlDat(ilIdx).iPdDay(ilDayIdx + 1) = 0 Then
                                    tlDat(ilIdx).iPdDay(ilDayIdx + 1) = 8
                                End If
                                ilLockedDays(ilDayIdx + 1) = True
                                If Not ilLockedDays(ilDayIdx) Then
                                    tlDat(ilIdx).iPdDay(ilDayIdx) = 0
                                End If
                            Else
                                tlDat(ilIdx).iPdDay(0) = 1
                                ilLockedDays(0) = True
                                If Not ilLockedDays(6) Then
                                    tlDat(ilIdx).iPdDay(ilDayIdx) = 0
                                End If
                            End If
                        End If
                    Next ilDayIdx
                    
                    For ilDayIdx = 0 To 6 Step 1
                        If tlDat(ilIdx).iPdDay(ilDayIdx) > 0 Then
                            tlDat(ilIdx).iPdDay(ilDayIdx) = 1
                        End If
                    Next ilDayIdx

                End If
            '9/30/11:  Move End If up so that Pledge will be adjusted indepedent of Feed
            'End If
            
            'Feed Times
            '9/30/11:  The rule has been changed to the following: Adjust Feed if the length of time is less then 5 min.
            '                                                      Adjust the Pledge if the length of time is less then 5 min
            'slFdStTime = Format$(DateAdd("h", ilTimeOffset, tlDat(ilIdx).sFdSTime), sgShowTimeWSecForm)
            'slFdEdTime = Format$(DateAdd("h", ilTimeOffset, tlDat(ilIdx).sFdETime), sgShowTimeWSecForm)
            If (smPledgeType <> "") Or ((smPledgeType = "") And (llFdTime <= AVAIL_OR_DP_TIME)) Then
                slFdStTime = Format$(DateAdd("h", ilTimeOffset, tlDat(ilIdx).sFdSTime), sgShowTimeWSecForm)
                slFdEdTime = Format$(DateAdd("h", ilTimeOffset, tlDat(ilIdx).sFdETime), sgShowTimeWSecForm)
            Else
                slFdStTime = Format$(tlDat(ilIdx).sFdSTime, sgShowTimeWSecForm)
                slFdEdTime = Format$(tlDat(ilIdx).sFdETime, sgShowTimeWSecForm)
            End If
            'D.S. 12/16/04 Per Jim, if the status is not then don't change the pledge times
            'Pledge Start Times
            If Len(Trim$(tlDat(ilIdx).sPdSTime)) = 0 Then
                slPdStTime = slFdStTime
            Else
                '9/30/11:  The rule has been changed to the following: Adjust Feed if the length of time is less then 5 min.
                '                                                      Adjust the Pledge if the length of time is less then 5 min
                ''If the pledge status is not live then we don't want to adjust the time
                'If tlDat(ilIdx).iFdStatus = 0 Then
                '    slPdStTime = Format$(DateAdd("h", ilTimeOffset, tlDat(ilIdx).sPdSTime), sgShowTimeWSecForm)
                'Else
                '    slPdStTime = Format$(DateAdd("h", 0, tlDat(ilIdx).sPdSTime), sgShowTimeWSecForm)
                'End If
                If ((smPledgeType <> "") And (tlDat(ilIdx).iFdStatus = 0)) Or ((smPledgeType = "") And (llPdTime <= AVAIL_OR_DP_TIME)) Then
                    slPdStTime = Format$(DateAdd("h", ilTimeOffset, tlDat(ilIdx).sPdSTime), sgShowTimeWSecForm)
                Else
                    slPdStTime = Format$(DateAdd("h", 0, tlDat(ilIdx).sPdSTime), sgShowTimeWSecForm)
                End If
            End If
            
            'Pledge End Times
            If Len(Trim$(tlDat(ilIdx).sPdETime)) = 0 Then
                slPdEdTime = slPdStTime
            Else
                '9/30/11:  The rule has been changed to the following: Adjust Feed if the length of time is less then 5 min.
                '                                                      Adjust the Pledge if the length of time is less then 5 min
                ''If the pledge status is not live then we don't want to adjust the time
                'If tlDat(ilIdx).iFdStatus = 0 Then
                '    slPdEdTime = Format$(DateAdd("h", ilTimeOffset, tlDat(ilIdx).sPdETime), sgShowTimeWSecForm)
                'Else
                '    slPdEdTime = Format$(DateAdd("h", 0, tlDat(ilIdx).sPdETime), sgShowTimeWSecForm)
                'End If
                If ((smPledgeType <> "") And (tlDat(ilIdx).iFdStatus = 0)) Or ((smPledgeType = "") And (llPdTime <= AVAIL_OR_DP_TIME)) Then
                    slPdEdTime = Format$(DateAdd("h", ilTimeOffset, tlDat(ilIdx).sPdETime), sgShowTimeWSecForm)
                Else
                    slPdEdTime = Format$(DateAdd("h", 0, tlDat(ilIdx).sPdETime), sgShowTimeWSecForm)
                End If
            End If
        Else
            'No time or date adjustments are necessary for this type of agreement
            slFdStTime = Format$(tlDat(ilIdx).sFdSTime, sgShowTimeWSecForm)
            slFdEdTime = Format$(tlDat(ilIdx).sFdETime, sgShowTimeWSecForm)
            If Len(Trim$(tlDat(ilIdx).sPdSTime)) = 0 Then
                slPdStTime = slFdStTime
            Else
                slPdStTime = Format$(tlDat(ilIdx).sPdSTime, sgShowTimeWSecForm)
            End If
            If Len(Trim$(tlDat(ilIdx).sPdETime)) = 0 Then
                slPdEdTime = slPdStTime
            Else
                slPdEdTime = Format$(tlDat(ilIdx).sPdETime, sgShowTimeWSecForm)
            End If
        End If
            
        tlDat(ilIdx).lCode = 0
        'SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, datDACode, "
        SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, "
        SQLQuery = SQLQuery & "datFdMon, datFdTue, datFdWed, datFdThu, "
        SQLQuery = SQLQuery & "datFdFri, datFdSat, datFdSun, datFdStTime, datFdEdTime, datFdStatus, "
        SQLQuery = SQLQuery & "datPdMon, datPdTue, datPdWed, datPdThu, datPdFri, "
        SQLQuery = SQLQuery & "datPdSat, datPdSun, datPdDayFed, datPdStTime, datPdEdTime, datAirPlayNo, datEstimatedTime, datEmbeddedOrROS)"
        SQLQuery = SQLQuery & " VALUES (" & tlDat(ilIdx).lCode & ", " & llAttCode & ", " & imAddShttCode & ", " & imVefCode
        SQLQuery = SQLQuery & "," '& tlDat(0).iDACode & ","
        SQLQuery = SQLQuery & tlDat(ilIdx).iFdDay(0) & ", " & tlDat(ilIdx).iFdDay(1) & ","
        SQLQuery = SQLQuery & tlDat(ilIdx).iFdDay(2) & ", " & tlDat(ilIdx).iFdDay(3) & ","
        SQLQuery = SQLQuery & tlDat(ilIdx).iFdDay(4) & ", " & tlDat(ilIdx).iFdDay(5) & ","
        SQLQuery = SQLQuery & tlDat(ilIdx).iFdDay(6) & ", "
        SQLQuery = SQLQuery & "'" & Format$(slFdStTime, sgSQLTimeForm) & "','" & Format$(slFdEdTime, sgSQLTimeForm) & "',"
        SQLQuery = SQLQuery & tlDat(ilIdx).iFdStatus & ","
        SQLQuery = SQLQuery & tlDat(ilIdx).iPdDay(0) & ", " & tlDat(ilIdx).iPdDay(1) & ","
        SQLQuery = SQLQuery & tlDat(ilIdx).iPdDay(2) & ", " & tlDat(ilIdx).iPdDay(3) & ","
        SQLQuery = SQLQuery & tlDat(ilIdx).iPdDay(4) & ", " & tlDat(ilIdx).iPdDay(5) & ","
        SQLQuery = SQLQuery & tlDat(ilIdx).iPdDay(6) & ", " & "'" & tlDat(ilIdx).sPdDayFed & "', "
        'SQLQuery = SQLQuery & "'" & Format$(slPdStTime, sgSQLTimeForm) & "','" & Format$(slPdEdTime, sgSQLTimeForm) & "')"
        SQLQuery = SQLQuery & "'" & Format$(slPdStTime, sgSQLTimeForm) & "','" & Format$(slPdEdTime, sgSQLTimeForm) & "',"
        SQLQuery = SQLQuery & tlDat(ilIdx).iAirPlayNo & ", "
        SQLQuery = SQLQuery & "'" & tlDat(ilIdx).sEstimatedTime & "', '" & tlDat(ilIdx).sEmbeddedOrROS & "')"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "FastAddVerbose.Txt", "FastAdd-mInsertDAT"
            mInsertDAT = False
            Exit Function
        End If
        SQLQuery = "SELECT MAX(datCode) from dat"
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            tlDat(ilIdx).lCode = rst(0).Value
        End If
    Next ilIdx
    mInsertDAT = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFastAdd-mInsertDat"
    mInsertDAT = False
    Exit Function
End Function

Private Function mGetTimeOffSet() As Integer
    'If the time zone is different from the station we are modeling from
    'then we need to make the adjust before inserting into the Dat file.
    Dim ilBaseTimeAdj As Integer
    Dim ilAddTimeAdj As Integer
    Dim ilNewTimeAdj As Integer
    Dim ilLoop As Integer
    Dim ilVef As Integer
    Dim iZone As Integer
    
    On Error GoTo ErrHand
    
    If (imAddShttCode > 0) And (imVefCode > 0) Then
        For ilLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(ilLoop).iCode = imAddShttCode Then
                For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
                    If tgVehicleInfo(ilVef).iCode = imVefCode Then
                        For iZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                            If StrComp(tgStationInfo(ilLoop).sZone, tgVehicleInfo(ilVef).sZone(iZone), 1) = 0 Then
                                ilBaseTimeAdj = tgVehicleInfo(ilVef).iVehLocalAdj(iZone)
                                Exit For
                            End If
                        Next iZone
                        Exit For
                    End If
                Next ilVef
                Exit For
            End If
        Next ilLoop
    End If

    If (imBaseShttCode > 0) And (imVefCode > 0) Then
        For ilLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(ilLoop).iCode = imBaseShttCode Then
                For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
                    If tgVehicleInfo(ilVef).iCode = imVefCode Then
                        For iZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                            If StrComp(tgStationInfo(ilLoop).sZone, tgVehicleInfo(ilVef).sZone(iZone), 1) = 0 Then
                                ilAddTimeAdj = tgVehicleInfo(ilVef).iVehLocalAdj(iZone)
                                Exit For
                            End If
                        Next iZone
                        Exit For
                    End If
                Next ilVef
                Exit For
            End If
        Next ilLoop
    End If

    ilNewTimeAdj = ilBaseTimeAdj - ilAddTimeAdj
    mGetTimeOffSet = ilNewTimeAdj
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-mGetTimeOffSet: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    mGetTimeOffSet = False
    Exit Function
End Function

Private Function mIsAgrmtOKToModelFrom(slPledgeType As String) As Integer
    'Dim ilAgreementType As Integer
    Dim ilLoop As Integer
    Dim slDatTime As String
    Dim slAttTime As String
    Dim llFdTime As Long
    Dim llPdTime As Long
    
    On Error GoTo ErrHand
    
    mIsAgrmtOKToModelFrom = False
    'ilAgreementType = tmDat(0).iDACode
    
    'Agreement by Dayparts
    'If ilAgreementType = 0 Then
    If slPledgeType = "D" Then
        For ilLoop = 0 To UBound(tmDat) - 1 Step 1
            If tmDat(ilLoop).iFdStatus = 0 Then
                gMsgBox "This agreement cannot be modelled from as it contains at least one status that is 1-Aired Live"
                Exit Function
            End If
        Next ilLoop
        imOKToConvertTimeZones = False
    End If
    
    'Agreement by Avails
    'If ilAgreementType = 1 Then
    If slPledgeType = "A" Then
        'For ilLoop = 0 To UBound(tmDat) - 1 Step 1
        '    If tmDat(ilLoop).iFdStatus = 1 Or tmDat(ilLoop).iFdStatus = 9 Or tmDat(ilLoop).iFdStatus = 10 Then
        '        gMsgBox "This agreement cannot be modelled from as it contains at least one status that is either 2-Aired Delay B'cast, 10-Delay Cmml/Prg or 11-Air Cmml"
        '        Exit Function
        '    End If
        'Next ilLoop
        imOKToConvertTimeZones = True
    End If
    
    'Agreement by CD/Tape
    'If ilAgreementType = 2 Then
    If slPledgeType = "C" Then
        slDatTime = Hour(Trim$(tmDat(0).sFdSTime))
        slAttTime = Hour(Trim$(attBase_rst!attStartTime))
        If slAttTime = slDatTime Then
            For ilLoop = 0 To UBound(tmDat) - 1 Step 1
                If tmDat(ilLoop).iFdStatus = 1 Or tmDat(ilLoop).iFdStatus = 9 Or tmDat(ilLoop).iFdStatus = 10 Then
                    gMsgBox "This agreement cannot be modelled from as it contains at least one status that is either 2-Aired Delay B'cast, 10-Delay Cmml/Prg or 11-Air Cmml"
                    Exit Function
                End If
            Next ilLoop
        Else
            For ilLoop = 0 To UBound(tmDat) - 1 Step 1
                If tmDat(ilLoop).iFdStatus = 0 Then
                    gMsgBox "This agreement cannot be modelled from as it contains at least one status that is 1-Aired Live"
                    Exit Function
                End If
            Next ilLoop
        End If
        imOKToConvertTimeZones = False
    End If
    If Trim$(slPledgeType) = "" Then
        '9/29/11:  This code was added to help determine how to handle the new design on the pledge tab.  The Same rule for specifing Estimate
        '          time will be used to determine if pledge should be adjusted
        '9/30/11:  The rule has been changed to the following: Adjust Feed if the length of time is less then 5 min.
        '                                                      Adjust the Pledge if the length of time is less then 5 min
        'imOKToConvertTimeZones = False
        'For ilLoop = 0 To UBound(tmDat) - 1 Step 1
        '    If tgStatusTypes(tmDat(ilLoop).iFdStatus).iPledged <> 1 Then    '1=Delay
        '        imOKToConvertTimeZones = True
        '        Exit For
        '    Else
        '        llFdTime = gTimeToLong(tmDat(ilLoop).sFdETime, True) - gTimeToLong(tmDat(ilLoop).sFdSTime, False)
        '        llPdTime = gTimeToLong(tmDat(ilLoop).sPdETime, True) - gTimeToLong(tmDat(ilLoop).sPdSTime, False)
        '        If llPdTime <= llFdTime Then
        '            imOKToConvertTimeZones = True
        '            Exit For
        '        End If
        '    End If
        'Next ilLoop
        imOKToConvertTimeZones = True
    End If
    
    mIsAgrmtOKToModelFrom = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-mIsAgrmtOKToModelFrom: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    mIsAgrmtOKToModelFrom = False
    Exit Function
End Function

Private Sub mSaveAttCount(ilShfCode As Integer)
    Dim blFound As Boolean
    Dim ilShtt As Integer
    blFound = False
    For ilShtt = 0 To UBound(tmFastAddAttCount) - 1 Step 1
        If tmFastAddAttCount(ilShtt).iShttCode = ilShfCode Then
            blFound = True
            tmFastAddAttCount(ilShtt).iShttCount = tmFastAddAttCount(ilShtt).iShttCount + 1
            Exit For
        End If
    Next ilShtt
    If Not blFound Then
        ilShtt = UBound(tmFastAddAttCount)
        tmFastAddAttCount(ilShtt).iShttCode = ilShfCode
        tmFastAddAttCount(ilShtt).iShttCount = 1
        ReDim Preserve tmFastAddAttCount(0 To ilShtt + 1) As FASTADDATTCOUNT
    End If
End Sub

Private Function mSetUsedForAtt(ilShttCode As Integer, slAgmntEnd As String) As Integer
    Dim llOffDate As Long
    Dim llDropDate As Long
    Dim llNowDate As Long
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand:
    
    llNowDate = DateValue(gAdjYear(Format(gNow(), sgShowDateForm)))
    If slAgmntEnd <> "" Then
        llOffDate = gDateValue(slAgmntEnd)
        llDropDate = llOffDate
    Else
        llOffDate = llNowDate + 1
        llDropDate = llOffDate
    End If
    If (llNowDate <= llOffDate) And (llNowDate <= llDropDate) Then
        slSQLQuery = "UPDATE shtt SET shttUsedForAtt = 'Y', shttAgreementExist = 'Y'"
        slSQLQuery = slSQLQuery & " WHERE shttCode = " & ilShttCode
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "FastAdd-mSetUsedForAtt"
            mSetUsedForAtt = False
            Exit Function
        End If
    Else
        slSQLQuery = "UPDATE shtt SET shttAgreementExist = 'Y'"
        slSQLQuery = slSQLQuery & " WHERE shttCode = " & ilShttCode
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "FastAdd-mSetUsedForAtt"
            mSetUsedForAtt = False
            Exit Function
        End If
    End If
    mSetUsedForAtt = True
    Exit Function
    
ErrHand:
    gHandleError "AffErrorLog.txt", "frmFastAdd-mSetUsedForAtt"
    mSetUsedForAtt = False
End Function

Private Function mCheckHistDate(ilShttCode As Integer, slAgmntStart As String) As Integer
    Dim blUpdateDate As Boolean
    Dim slHistDate As String
    Dim slSQLQuery As String
    On Error GoTo ErrHand
    
    If slAgmntStart = "" Then
        mCheckHistDate = True
        Exit Function
    End If
    slSQLQuery = "UPDATE shtt SET shttHistStartDate = '" & Format(slAgmntStart, sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery & " WHERE shttHistStartDate > '" & Format(slAgmntStart, sgSQLDateForm) & "' AND shttCode = " & ilShttCode
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "FastAdd-mCheckHistDate"
        mCheckHistDate = False
        Exit Function
    End If
    mCheckHistDate = True
    Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmFastAdd-mCheckHistDate"
    mCheckHistDate = False
End Function

Private Function mGetPledgeByEvent() As String
    Dim ilVff As Integer
    mGetPledgeByEvent = "N"
    If ((Asc(sgSpfSportInfo) And USINGSPORTS) <> USINGSPORTS) Then
        Exit Function
    End If
    If imVefCode <= 0 Then
        Exit Function
    End If
    ilVff = gBinarySearchVff(imVefCode)
    If ilVff <> -1 Then
        If Trim$(tgVffInfo(ilVff).sPledgeByEvent) = "" Then
            mGetPledgeByEvent = "N"
        Else
            mGetPledgeByEvent = Trim$(tgVffInfo(ilVff).sPledgeByEvent)
        End If
    End If
End Function

Private Function mAddPet(ilVefCode As Integer, ilShttCode As Integer, llNewATTCode As Long, llOldAttCode As Long)
    Dim slDeclaredStatus As String
    Dim llPetCode As Long
    Dim blPetBuilt As Boolean
    
    On Error GoTo ErrHand
    blPetBuilt = False
    If llOldAttCode > 0 Then
        SQLQuery = "SELECT petCode, petGsfCode, petDeclaredStatus, petClearStatus"
        SQLQuery = SQLQuery + " FROM pet"
        SQLQuery = SQLQuery & " WHERE (petAttCode = " & llOldAttCode & ")"
        SQLQuery = SQLQuery + " ORDER BY petGsfCode"
        Set rst_Pet = gSQLSelectCall(SQLQuery)
        If Not rst_Pet.EOF Then
            blPetBuilt = True
            While Not rst_Pet.EOF
                SQLQuery = "Insert Into pet ( "
                SQLQuery = SQLQuery & "petCode, "
                SQLQuery = SQLQuery & "petAttCode, "
                SQLQuery = SQLQuery & "petVefCode, "
                SQLQuery = SQLQuery & "petShttCode, "
                SQLQuery = SQLQuery & "petGsfCode, "
                SQLQuery = SQLQuery & "petDeclaredStatus, "
                SQLQuery = SQLQuery & "petClearStatus, "
                SQLQuery = SQLQuery & "petUstCode, "
                SQLQuery = SQLQuery & "petEnteredDate, "
                SQLQuery = SQLQuery & "petEnteredTime, "
                SQLQuery = SQLQuery & "petUnused "
                SQLQuery = SQLQuery & ") "
                SQLQuery = SQLQuery & "Values ( "
                SQLQuery = SQLQuery & "Replace" & ", "
                SQLQuery = SQLQuery & llNewATTCode & ", "
                SQLQuery = SQLQuery & ilVefCode & ", "
                SQLQuery = SQLQuery & ilShttCode & ", "
                SQLQuery = SQLQuery & rst_Pet!petGsfCode & ", "
                SQLQuery = SQLQuery & "'" & gFixQuote(rst_Pet!petDeclaredStatus) & "', "
                SQLQuery = SQLQuery & "'" & gFixQuote(rst_Pet!petClearStatus) & "', "
                SQLQuery = SQLQuery & igUstCode & ", "
                SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
                SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLTimeForm) & "', "
                SQLQuery = SQLQuery & "'" & gFixQuote("") & "' "
                SQLQuery = SQLQuery & ") "
                llPetCode = gInsertAndReturnCode(SQLQuery, "pet", "petCode", "Replace")
                rst_Pet.MoveNext
            Wend
        End If
    End If
    If Not blPetBuilt Then
        SQLQuery = "SELECT * FROM GSF_Game_Schd WHERE (gsfVefCode = " & ilVefCode & ")"
        Set rst_Gsf = gSQLSelectCall(SQLQuery)
        Do While Not rst_Gsf.EOF
            
            If gDateValue(Format$(rst_Gsf!gsfAirDate, sgShowDateForm)) <= gDateValue(Format(gNow(), "m/d/yy")) Then
                slDeclaredStatus = "N"
            Else
                slDeclaredStatus = "U"
            End If
            SQLQuery = "Insert Into pet ( "
            SQLQuery = SQLQuery & "petCode, "
            SQLQuery = SQLQuery & "petAttCode, "
            SQLQuery = SQLQuery & "petVefCode, "
            SQLQuery = SQLQuery & "petShttCode, "
            SQLQuery = SQLQuery & "petGsfCode, "
            SQLQuery = SQLQuery & "petDeclaredStatus, "
            SQLQuery = SQLQuery & "petClearStatus, "
            SQLQuery = SQLQuery & "petUstCode, "
            SQLQuery = SQLQuery & "petEnteredDate, "
            SQLQuery = SQLQuery & "petEnteredTime, "
            SQLQuery = SQLQuery & "petUnused "
            SQLQuery = SQLQuery & ") "
            SQLQuery = SQLQuery & "Values ( "
            SQLQuery = SQLQuery & "Replace" & ", "
            SQLQuery = SQLQuery & llNewATTCode & ", "
            SQLQuery = SQLQuery & ilVefCode & ", "
            SQLQuery = SQLQuery & ilShttCode & ", "
            SQLQuery = SQLQuery & rst_Gsf!gsfCode & ", "
            SQLQuery = SQLQuery & "'" & gFixQuote(slDeclaredStatus) & "', "
            SQLQuery = SQLQuery & "'" & gFixQuote("U") & "', "
            SQLQuery = SQLQuery & igUstCode & ", "
            SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLTimeForm) & "', "
            SQLQuery = SQLQuery & "'" & gFixQuote("") & "' "
            SQLQuery = SQLQuery & ") "
            llPetCode = gInsertAndReturnCode(SQLQuery, "pet", "petCode", "Replace")
            rst_Gsf.MoveNext
        Loop
    End If
    mAddPet = True
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Fast Add-mAddPet"
    mAddPet = False
End Function

Private Sub mSaveStationList()
    Dim ilLoop As Integer
    Dim llRow As Long
    
    'lbcStations(2).Clear
    'lbcStations(3).Clear
    'For ilLoop = 0 To lbcStations(0).ListCount - 1 Step 1
    '    lbcStations(2).AddItem lbcStations(0).List(ilLoop)
    '    lbcStations(2).ItemData(lbcStations(2).NewIndex) = lbcStations(0).ItemData(ilLoop)
    'Next ilLoop
    'For ilLoop = 0 To lbcStations(1).ListCount - 1 Step 1
    '    lbcStations(3).AddItem lbcStations(1).List(ilLoop)
    '    lbcStations(3).ItemData(lbcStations(3).NewIndex) = lbcStations(1).ItemData(ilLoop)
    'Next ilLoop
    mClearGrid grdSvExclude
    mClearGrid grdSvInclude
        
    'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - make grid's invisible to help performance
    Screen.MousePointer = vbHourglass
    DoEvents
    grdSvExclude.Visible = False
    grdSvInclude.Visible = False
    grdExclude.Visible = False
    grdInclude.Visible = False
    For llRow = grdExclude.FixedRows To grdExclude.Rows - 1 Step 1
        If grdExclude.TextMatrix(llRow, SHTTCODEINDEX) <> "" Then
            mCopyGridRow grdExclude, grdSvExclude, llRow, llRow
        End If
    Next llRow
    For llRow = grdInclude.FixedRows To grdInclude.Rows - 1 Step 1
        If grdInclude.TextMatrix(llRow, SHTTCODEINDEX) <> "" Then
            mCopyGridRow grdInclude, grdSvInclude, llRow, llRow
        End If
    Next llRow
    
    grdSvExclude.Visible = False
    grdSvInclude.Visible = False
    grdExclude.Visible = True
    grdInclude.Visible = True
End Sub

Private Sub mRestoreStationList()
    Dim ilLoop As Integer
    Dim llRow As Long
    'lbcStations(0).Clear
    'lbcStations(1).Clear
    'For ilLoop = 0 To lbcStations(2).ListCount - 1 Step 1
    '    lbcStations(0).AddItem lbcStations(2).List(ilLoop)
    '    lbcStations(0).ItemData(lbcStations(0).NewIndex) = lbcStations(2).ItemData(ilLoop)
    'Next ilLoop
    'For ilLoop = 0 To lbcStations(3).ListCount - 1 Step 1
    '    lbcStations(1).AddItem lbcStations(3).List(ilLoop)
    '    lbcStations(1).ItemData(lbcStations(1).NewIndex) = lbcStations(3).ItemData(ilLoop)
    'Next ilLoop
    mClearGrid grdExclude
    mClearGrid grdInclude
    For llRow = grdSvExclude.FixedRows To grdSvExclude.Rows - 1 Step 1
        If grdSvExclude.TextMatrix(llRow, SHTTCODEINDEX) <> "" Then
            mCopyGridRow grdSvExclude, grdExclude, llRow, llRow
        End If
    Next llRow
    For llRow = grdSvInclude.FixedRows To grdSvInclude.Rows - 1 Step 1
        If grdSvInclude.TextMatrix(llRow, SHTTCODEINDEX) <> "" Then
            mCopyGridRow grdSvInclude, grdInclude, llRow, llRow
        End If
    Next llRow
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    grdExclude.ColWidth(SHTTCODEINDEX) = 0
    grdExclude.ColWidth(RANKINDEX) = 0
    grdExclude.ColWidth(SORTINDEX) = 0
    'grdExclude.ColWidth(SELECTEDINDEX) = 0
    If rbcGetFrom(1).Value Then
        grdExclude.ColWidth(CALLLETTERSINDEX) = grdExclude.Width * 0.15
        grdExclude.ColWidth(MARKETINDEX) = grdExclude.Width * 0.17
        grdExclude.ColWidth(FORMATINDEX) = grdExclude.Width * 0.17
        grdExclude.ColWidth(OWNERINDEX) = grdExclude.Width * 0.17
        grdExclude.ColWidth(ZONEINDEX) = grdExclude.Width * 0.1
        grdExclude.ColWidth(DATERANGEINDEX) = grdExclude.Width * 0.2
    Else
        grdExclude.ColWidth(CALLLETTERSINDEX) = grdExclude.Width * 0.15
        grdExclude.ColWidth(MARKETINDEX) = grdExclude.Width * 0.23
        grdExclude.ColWidth(FORMATINDEX) = grdExclude.Width * 0.23
        grdExclude.ColWidth(OWNERINDEX) = grdExclude.Width * 0.23
        grdExclude.ColWidth(ZONEINDEX) = grdExclude.Width * 0.1
        grdExclude.ColWidth(DATERANGEINDEX) = 0
    End If
    grdExclude.ColWidth(CALLLETTERSINDEX) = grdExclude.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To DATERANGEINDEX Step 1
        If ilCol <> CALLLETTERSINDEX Then
            grdExclude.ColWidth(CALLLETTERSINDEX) = grdExclude.ColWidth(CALLLETTERSINDEX) - grdExclude.ColWidth(ilCol)
        End If
    Next ilCol
    For ilCol = 0 To SORTINDEX Step 1
        grdInclude.ColWidth(ilCol) = grdExclude.ColWidth(ilCol)
    Next ilCol


    'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
    grdFastAddFile.ColWidth(FILELINE) = 1290
    grdFastAddFile.ColWidth(FILECALLLETTERS) = 1785
    grdFastAddFile.ColWidth(FILEVEHICLE) = 6870
    grdFastAddFile.ColWidth(FILESTARTDATE) = 1935
    grdFastAddFile.ColWidth(FILEPLEDGESFROM) = 2145
    grdFastAddFile.ColWidth(FILEDLVYSVCMODE) = 2025
    grdFastAddFile.ColWidth(FILEGENFOR) = 2370
    grdFastAddFile.ColWidth(FILESTATUS) = 7080
    grdFastAddFile.ColWidth(FILEVEFCODE) = 0
    grdFastAddFile.ColWidth(FILECALLLETTERSCODE) = 0
    grdFastAddFile.ColWidth(FILEPLEDGESFROMCODE) = 0

    'Align columns to left
    gGrid_AlignAllColsLeft grdExclude
    gGrid_AlignAllColsLeft grdInclude
    gGrid_AlignAllColsLeft grdFastAddFile
End Sub

Private Sub mSetGridTitles()
    Dim llCol As Long
    
    'Set column titles
    grdExclude.TextMatrix(0, CALLLETTERSINDEX) = "Call Letters"
    grdExclude.TextMatrix(0, MARKETINDEX) = "Market"
    grdExclude.TextMatrix(0, FORMATINDEX) = "Format"
    grdExclude.TextMatrix(0, OWNERINDEX) = "Owner"
    grdExclude.TextMatrix(0, ZONEINDEX) = "Zone"
    grdExclude.TextMatrix(0, DATERANGEINDEX) = "Date Range"
    
    grdInclude.TextMatrix(0, CALLLETTERSINDEX) = grdExclude.TextMatrix(0, CALLLETTERSINDEX)
    grdInclude.TextMatrix(0, MARKETINDEX) = grdExclude.TextMatrix(0, MARKETINDEX)
    grdInclude.TextMatrix(0, FORMATINDEX) = grdExclude.TextMatrix(0, FORMATINDEX)
    grdInclude.TextMatrix(0, OWNERINDEX) = grdExclude.TextMatrix(0, OWNERINDEX)
    grdInclude.TextMatrix(0, ZONEINDEX) = grdExclude.TextMatrix(0, ZONEINDEX)
    grdInclude.TextMatrix(0, DATERANGEINDEX) = grdExclude.TextMatrix(0, DATERANGEINDEX)
    
    grdExclude.Row = 0
    grdInclude.Row = 0
    For llCol = CALLLETTERSINDEX To DATERANGEINDEX Step 1
        grdExclude.Col = llCol
        grdExclude.CellBackColor = LIGHTBLUE
        grdInclude.Col = llCol
        grdInclude.CellBackColor = LIGHTBLUE
    Next llCol
    'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - use native grid selection
    grdExclude.Row = 1
    grdExclude.Col = 0
    grdExclude.ColSel = SORTINDEX
    grdInclude.Row = 1
    grdInclude.Col = 0
    grdInclude.ColSel = SORTINDEX

    'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
    grdFastAddFile.TextMatrix(0, FILELINE) = "Line"
    grdFastAddFile.TextMatrix(0, FILECALLLETTERS) = "Call Letters"
    grdFastAddFile.TextMatrix(0, FILEVEHICLE) = "Vehicle"
    grdFastAddFile.TextMatrix(0, FILESTARTDATE) = "Start Date"
    grdFastAddFile.TextMatrix(0, FILEPLEDGESFROM) = "Pledges From"
    grdFastAddFile.TextMatrix(0, FILEDLVYSVCMODE) = "Dlv Svc Mode"
    grdFastAddFile.TextMatrix(0, FILEGENFOR) = "Gen For"
    grdFastAddFile.TextMatrix(0, FILESTATUS) = "Status"
    grdFastAddFile.TextMatrix(0, FILEVEFCODE) = "VefCode"
    grdFastAddFile.TextMatrix(0, FILECALLLETTERSCODE) = "CallLetterSHTT"
    grdFastAddFile.TextMatrix(0, FILEPLEDGESFROMCODE) = "PledgesSHTT"
    grdFastAddFile.ScrollBars = flexScrollBarBoth
End Sub

Private Sub mSortExcludeCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
    For llRow = grdExclude.FixedRows To grdExclude.Rows - 1 Step 1
        slStr = Trim$(grdExclude.TextMatrix(llRow, CALLLETTERSINDEX))
        If slStr <> "" Then
            If ilCol = DATERANGEINDEX Then
                slStr = grdExclude.TextMatrix(llRow, DATERANGEINDEX)
                ilPos = InStr(1, slStr, "-")
                If ilPos > 0 Then
                    slStr = Left(slStr, ilPos - 1)
                End If
                slSort = Trim$(Str$(gDateValue(slStr)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = MARKETINDEX Then
                slSort = grdExclude.TextMatrix(llRow, RANKINDEX)
                If slSort = "" Then
                    slSort = "999"
                Else
                    Do While Len(slSort) < 3
                        slSort = "0" & slSort
                    Loop
                End If
            Else
                slSort = UCase$(Trim$(grdExclude.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdExclude.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastExcludeColSorted) Or ((ilCol = imLastExcludeColSorted) And (imLastExcludeSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdExclude.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdExclude.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastExcludeColSorted Then
        imLastExcludeColSorted = SORTINDEX
    Else
        imLastExcludeColSorted = -1
        imLastExcludeSort = -1
    End If
    gGrid_SortByCol grdExclude, CALLLETTERSINDEX, SORTINDEX, imLastExcludeColSorted, imLastExcludeSort
    imLastExcludeColSorted = ilCol
End Sub

Private Sub mSortIncludeCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
    For llRow = grdInclude.FixedRows To grdInclude.Rows - 1 Step 1
        slStr = Trim$(grdInclude.TextMatrix(llRow, CALLLETTERSINDEX))
        If slStr <> "" Then
            If ilCol = DATERANGEINDEX Then
                slStr = grdInclude.TextMatrix(llRow, DATERANGEINDEX)
                ilPos = InStr(1, slStr, "-")
                If ilPos > 0 Then
                    slStr = Left(slStr, ilPos - 1)
                End If
                slSort = Trim$(Str$(gDateValue(slStr)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = MARKETINDEX Then
                slSort = grdInclude.TextMatrix(llRow, RANKINDEX)
                If slSort = "" Then
                    slSort = "999"
                Else
                    Do While Len(slSort) < 3
                        slSort = "0" & slSort
                    Loop
                End If
            Else
                slSort = UCase$(Trim$(grdInclude.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdInclude.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastIncludeColSorted) Or ((ilCol = imLastIncludeColSorted) And (imLastIncludeSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdInclude.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdInclude.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastIncludeColSorted Then
        imLastIncludeColSorted = SORTINDEX
    Else
        imLastIncludeColSorted = -1
        imLastIncludeSort = -1
    End If
    gGrid_SortByCol grdInclude, CALLLETTERSINDEX, SORTINDEX, imLastIncludeColSorted, imLastIncludeSort
    imLastIncludeColSorted = ilCol
End Sub

Private Sub mClearGrid(grd As MSHFlexGrid)
    Dim llRow As Long
    Dim llCol As Long
    
    'Blank rows within grid
     
    'gGrid_Clear grdAirPlay, False
    'Set color within cells
'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - removed Selected index, using Native grid row, RowSel properties
'    For llRow = grd.FixedRows To grd.Rows - 1 Step 1
'        For llCol = CALLLETTERSINDEX To SELECTEDINDEX Step 1
'            grd.TextMatrix(llRow, llCol) = ""
'            grd.Row = llRow
'            grd.Col = llCol
'            grd.CellBackColor = vbWhite
'        Next llCol
'    Next llRow
    
    'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - clearing cell by cell is slow.  Resize grid, then clear last row
    grd.Rows = 2
    
    mSetGridTitles
    grd.TopRow = grd.FixedRows
    grd.Row = 1
    For llCol = CALLLETTERSINDEX To grd.Cols - 1
        grd.TextMatrix(1, llCol) = ""
        grd.Col = llCol
        grd.CellBackColor = vbWhite
        grd.CellForeColor = vbBlack
    Next llCol
End Sub

Private Sub mAddStationToGrid(grdAdd As MSHFlexGrid, llRow As Long, slCallLetters As String, slRange As String)
    Dim llShtt As Long
    Dim llIndex As Integer
    Dim llMkt As Long
    
    If llRow >= grdAdd.Rows Then
        grdAdd.AddItem "", llRow
    Else
        If grdAdd.TextMatrix(llRow, CALLLETTERSINDEX) <> "" Then
            grdAdd.AddItem "", llRow
        End If
    End If
    llShtt = gBinarySearchStation(Trim$(slCallLetters))
    If llShtt < 0 Then
        grdAdd.TextMatrix(llRow, CALLLETTERSINDEX) = Trim$(slCallLetters)
        grdAdd.Row = llRow
        grdAdd.Col = CALLLETTERSINDEX
        grdAdd.CellForeColor = vbRed
        grdAdd.TextMatrix(llRow, SHTTCODEINDEX) = -1
        llRow = llRow + 1
        Exit Sub
    End If
    grdAdd.TextMatrix(llRow, CALLLETTERSINDEX) = Trim$(tgStationInfo(llShtt).sCallLetters)
    grdAdd.TextMatrix(llRow, DATERANGEINDEX) = slRange
    grdAdd.TextMatrix(llRow, MARKETINDEX) = Trim$(tgStationInfo(llShtt).sMarket)
    llMkt = gBinarySearchMkt(CLng(tgStationInfo(llShtt).iMktCode))
    If llMkt >= 0 Then
        grdAdd.TextMatrix(llRow, RANKINDEX) = tgMarketInfo(llMkt).iRank
    Else
        grdAdd.TextMatrix(llRow, RANKINDEX) = ""
    End If
    llIndex = gBinarySearchFmt(CLng(tgStationInfo(llShtt).iFormatCode))
    If llIndex >= 0 Then
        grdAdd.TextMatrix(llRow, FORMATINDEX) = Trim$(tgFormatInfo(llIndex).sName)
    Else
        grdAdd.TextMatrix(llRow, FORMATINDEX) = ""
    End If
    llIndex = gBinarySearchOwner(CLng(tgStationInfo(llShtt).lOwnerCode))
    If llIndex >= 0 Then
        grdAdd.TextMatrix(llRow, OWNERINDEX) = Trim(tgOwnerInfo(llIndex).sName)
    Else
        grdAdd.TextMatrix(llRow, OWNERINDEX) = ""
    End If
    grdAdd.TextMatrix(llRow, ZONEINDEX) = Trim$(tgStationInfo(llShtt).sZone)
    grdAdd.TextMatrix(llRow, SHTTCODEINDEX) = tgStationInfo(llShtt).iCode
    llRow = llRow + 1
End Sub

Private Sub mPaintRowColor(grdPaint As MSHFlexGrid, llRow As Long)
    Dim llCol As Long
    grdPaint.Row = llRow
    'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - SLOW ON BIG LISTS (19k + items)
    'For llCol = CALLLETTERSINDEX To DATERANGEINDEX Step 1
        'grdPaint.Col = llCol
        'If grdPaint.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
        '    If grdPaint.CellBackColor = GRAY Then grdPaint.CellBackColor = vbWhite
        'Else
        '    If grdPaint.CellBackColor = vbWhite Then grdPaint.CellBackColor = GRAY    'vbBlue
        'End If
    'Next llCol
    grdPaint.SelectionMode = flexSelectionByRow
    grdPaint.Row = llRow
End Sub

Private Sub mMoveGridRow(grdFrom As MSHFlexGrid, grdTo As MSHFlexGrid, llFromRow As Long, llToRow As Long)
    Dim ilCount As Integer
    Dim llRow As Long
    'If llToRow >= grdTo.Rows Then
    If llToRow >= grdTo.Rows Then
        'grdTo.AddItem "", llToRow
        grdTo.AddItem ""
    Else
        If grdTo.TextMatrix(llToRow, CALLLETTERSINDEX) <> "" Then
            'grdTo.AddItem "", llToRow
            grdTo.AddItem ""
        End If
    End If
    grdTo.TextMatrix(llToRow, CALLLETTERSINDEX) = grdFrom.TextMatrix(llFromRow, CALLLETTERSINDEX)
    grdTo.TextMatrix(llToRow, DATERANGEINDEX) = grdFrom.TextMatrix(llFromRow, DATERANGEINDEX)
    grdTo.TextMatrix(llToRow, MARKETINDEX) = grdFrom.TextMatrix(llFromRow, MARKETINDEX)
    grdTo.TextMatrix(llToRow, RANKINDEX) = grdFrom.TextMatrix(llFromRow, RANKINDEX)
    grdTo.TextMatrix(llToRow, FORMATINDEX) = grdFrom.TextMatrix(llFromRow, FORMATINDEX)
    grdTo.TextMatrix(llToRow, OWNERINDEX) = grdFrom.TextMatrix(llFromRow, OWNERINDEX)
    grdTo.TextMatrix(llToRow, ZONEINDEX) = grdFrom.TextMatrix(llFromRow, ZONEINDEX)
    grdTo.TextMatrix(llToRow, SHTTCODEINDEX) = grdFrom.TextMatrix(llFromRow, SHTTCODEINDEX)
    'grdTo.TextMatrix(llToRow, SELECTEDINDEX) = ""
    ilCount = 0
'    For llRow = grdFrom.FixedRows To grdFrom.Rows - 1 Step 1
'        If grdFrom.TextMatrix(llRow, SHTTCODEINDEX) <> "" Then
'            ilCount = ilCount + 1
'            If ilCount > 1 Then
'                Exit For
'            End If
'        End If
'    Next llRow
    ilCount = grdFrom.Rows - 1

    If ilCount > 1 Then
        grdFrom.RemoveItem llFromRow
    Else
        mClearGrid grdFrom
    End If
End Sub

Private Sub mCopyGridRow(grdFrom As MSHFlexGrid, grdTo As MSHFlexGrid, llFromRow As Long, llToRow As Long)
    'If llToRow >= grdTo.Rows Then
    If llToRow >= grdTo.Rows Then
        grdTo.AddItem "", llToRow
    Else
        If grdTo.TextMatrix(llToRow, CALLLETTERSINDEX) <> "" Then
            grdTo.AddItem "", llToRow
        End If
    End If
    grdTo.TextMatrix(llToRow, CALLLETTERSINDEX) = grdFrom.TextMatrix(llFromRow, CALLLETTERSINDEX)
    grdTo.TextMatrix(llToRow, DATERANGEINDEX) = grdFrom.TextMatrix(llFromRow, DATERANGEINDEX)
    grdTo.TextMatrix(llToRow, MARKETINDEX) = grdFrom.TextMatrix(llFromRow, MARKETINDEX)
    grdTo.TextMatrix(llToRow, RANKINDEX) = grdFrom.TextMatrix(llFromRow, RANKINDEX)
    grdTo.TextMatrix(llToRow, FORMATINDEX) = grdFrom.TextMatrix(llFromRow, FORMATINDEX)
    grdTo.TextMatrix(llToRow, OWNERINDEX) = grdFrom.TextMatrix(llFromRow, OWNERINDEX)
    grdTo.TextMatrix(llToRow, ZONEINDEX) = grdFrom.TextMatrix(llFromRow, ZONEINDEX)
    grdTo.TextMatrix(llToRow, SHTTCODEINDEX) = grdFrom.TextMatrix(llFromRow, SHTTCODEINDEX)
    'grdTo.TextMatrix(llToRow, SELECTEDINDEX) = grdTo.TextMatrix(llFromRow, SELECTEDINDEX)
End Sub

'Taken from:gAlignAllMulticastStations
Sub mAlignAllMulticastStations()
    Dim ilLoop1 As Integer
    Dim ilShttCode1 As Integer
    Dim llGroupID As Long
    Dim ilLoop2 As Integer
    Dim ilShttCode2 As Integer
    Dim ilFound As Integer
    Dim llRow As Long
    Dim shtt_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
        
    For ilLoop1 = grdInclude.Rows - 1 To grdInclude.FixedRows Step -1
        If grdInclude.TextMatrix(ilLoop1, SHTTCODEINDEX) <> "" Then
            ilShttCode1 = Val(grdInclude.TextMatrix(ilLoop1, SHTTCODEINDEX))
            llGroupID = gGetStaMulticastGroupID(ilShttCode1)
            If llGroupID > 0 Then
                SQLQuery = "Select shttCode FROM shtt where shttMultiCastGroupID = " & llGroupID
                Set shtt_rst = gSQLSelectCall(SQLQuery)
                While Not shtt_rst.EOF
                    ilFound = False
                    ilShttCode2 = shtt_rst!shttCode
                    For ilLoop2 = grdInclude.FixedRows To grdInclude.Rows - 1 Step 1
                        If grdInclude.TextMatrix(ilLoop2, SHTTCODEINDEX) <> "" Then
                            If ilShttCode2 = Val(grdInclude.TextMatrix(ilLoop2, SHTTCODEINDEX)) Then
                                ilFound = True
                                Exit For
                            End If
                        End If
                    Next ilLoop2
                    If Not ilFound Then
                        llRow = grdInclude.FixedRows
                        For ilLoop2 = grdInclude.Rows - 1 To grdInclude.FixedRows Step -1
                            If grdInclude.TextMatrix(ilLoop2, CALLLETTERSINDEX) <> "" Then
                                llRow = ilLoop2 + 1
                                Exit For
                            End If
                        Next ilLoop2
                        For ilLoop2 = 0 To grdExclude.Rows - 1 Step 1
                            If grdExclude.TextMatrix(ilLoop2, SHTTCODEINDEX) <> "" Then
                                If ilShttCode2 = Val(grdExclude.TextMatrix(ilLoop2, SHTTCODEINDEX)) Then
                                    'Move
                                    mMoveGridRow grdExclude, grdInclude, CLng(ilLoop2), llRow
                                    Exit For
                                End If
                            End If
                        Next ilLoop2
                    End If
                    shtt_rst.MoveNext
                Wend
            End If
        End If
    Next ilLoop1
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "gAlignAllMulticastStations"
    Resume Next
ErrHand1:
    gHandleError "AffErrorLog.txt", "gAlignAllMulticastStations"
    Return
End Sub

'Taken from: gAlignMulticastStations
Sub mAlignMulticastStations(ilVefCode As Integer, slStationOrAgreementItemData As String, Optional llAttStartDate As Long = -1, Optional llAttEndDate As Long = -1)
    '9/12/18: Added date as an option.  It was added to handle call from Spot Utilty
    Dim ilLoop1 As Integer
    Dim ilShttCode1 As Integer
    Dim llGroupID As Long
    Dim slMulticast1 As String
    Dim llLastDate1 As Long
    Dim ilLoop2 As Integer
    Dim ilShttCode2 As Integer
    Dim slMulticast2 As String
    Dim llLastDate2 As Long
    Dim ilShttCode3 As Integer
    Dim llOnAir As Long
    Dim ilFound As Integer
    Dim ilLoopIndex As Integer
    Dim llRow As Long
    Dim att_rst As ADODB.Recordset
    Dim att_rst2 As ADODB.Recordset
    Dim shtt_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
        
    For ilLoop1 = grdInclude.Rows - 1 To grdInclude.FixedRows Step -1
        If grdInclude.TextMatrix(ilLoop1, SHTTCODEINDEX) <> "" Then
            ilShttCode1 = grdInclude.TextMatrix(ilLoop1, SHTTCODEINDEX)
            If gIsMulticast(ilShttCode1) Then
                'Determine if agreement exist, if not then can determine if any other station needs to be multicast with it
                llLastDate1 = 0
                llOnAir = 0
                slMulticast1 = ""
                SQLQuery = "SELECT attCode, attOnAir, attOffAir, attDropDate, attMulticast FROM att"
                SQLQuery = SQLQuery + " WHERE (attShfCode = " & ilShttCode1 & " AND attVefCode = " & ilVefCode & ")"
                Set att_rst = gSQLSelectCall(SQLQuery)
                While Not att_rst.EOF
                    '9/12/18: Added date as an option.
                    llOnAir = gDateValue(att_rst!attOnAir)
                    If gDateValue(att_rst!attOffAir) <= gDateValue(att_rst!attDropDate) Then
                        If gDateValue(att_rst!attOffAir) > llLastDate1 Then
                            llLastDate1 = gDateValue(att_rst!attOffAir)
                            slMulticast1 = att_rst!attMulticast
                        End If
                    Else
                        If gDateValue(att_rst!attDropDate) > llLastDate1 Then
                            llLastDate1 = gDateValue(att_rst!attDropDate)
                            slMulticast1 = att_rst!attMulticast
                        End If
                    End If
                    att_rst.MoveNext
                Wend
                '9/12/18: Added date as an option.
                If (slMulticast1 = "Y") And ((llAttEndDate >= llOnAir) Or (llAttEndDate = -1)) And ((llAttStartDate <= llLastDate1) Or (llAttStartDate = -1)) Then
                    'Obtain list of other multicast stations
                    llGroupID = gGetStaMulticastGroupID(ilShttCode1)
                    SQLQuery = "Select shttCode FROM shtt where shttMultiCastGroupID = " & llGroupID
                    Set shtt_rst = gSQLSelectCall(SQLQuery)
                    While Not shtt_rst.EOF
                        ilShttCode2 = shtt_rst!shttCode
                        llLastDate2 = 0
                        slMulticast2 = ""
                        SQLQuery = "SELECT attCode, attOnAir, attOffAir, attDropDate, attMulticast FROM att"
                        SQLQuery = SQLQuery + " WHERE (attShfCode = " & ilShttCode2 & " AND attVefCode = " & ilVefCode & ")"
                        Set att_rst = gSQLSelectCall(SQLQuery)
                        While Not att_rst.EOF
                            If gDateValue(att_rst!attOffAir) <= gDateValue(att_rst!attDropDate) Then
                                If gDateValue(att_rst!attOffAir) > llLastDate2 Then
                                    llLastDate2 = gDateValue(att_rst!attOffAir)
                                    slMulticast2 = att_rst!attMulticast
                                End If
                            Else
                                If gDateValue(att_rst!attDropDate) > llLastDate2 Then
                                    llLastDate2 = gDateValue(att_rst!attDropDate)
                                    slMulticast2 = att_rst!attMulticast
                                End If
                            End If
                            att_rst.MoveNext
                        Wend
                        If (slMulticast2 = "Y") And (llLastDate1 = llLastDate2) Then
                            ilFound = False
                            For ilLoop2 = grdInclude.FixedRows To grdInclude.Rows - 1 Step 1
                                If grdInclude.TextMatrix(ilLoop2, SHTTCODEINDEX) <> "" Then
                                    If ilShttCode2 = Val(grdInclude.TextMatrix(ilLoop2, SHTTCODEINDEX)) Then
                                        SQLQuery = "SELECT attCode, attOnAir, attOffAir, attDropDate, attMulticast FROM att"
                                        SQLQuery = SQLQuery + " WHERE (attShfCode = " & ilShttCode2 & " AND attVefCode = " & ilVefCode & ")"
                                        Set att_rst = gSQLSelectCall(SQLQuery)
                                        While Not att_rst.EOF
                                            If gDateValue(att_rst!attOffAir) <= gDateValue(att_rst!attDropDate) Then
                                                If llLastDate2 = gDateValue(att_rst!attOffAir) Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Else
                                                If llLastDate2 = gDateValue(att_rst!attDropDate) Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            End If
                                            att_rst.MoveNext
                                        Wend
                                    End If
                                End If
                            Next ilLoop2
                            If Not ilFound Then
                                'See if in other list: if so move it;
                                ilFound = False
                                For ilLoop2 = 0 To grdExclude.Rows - 1 Step 1
                                    If grdExclude.TextMatrix(ilLoop2, SHTTCODEINDEX) <> "" Then
                                        If ilShttCode2 = Val(grdExclude.TextMatrix(ilLoop2, SHTTCODEINDEX)) Then
                                            SQLQuery = "SELECT attCode, attOnAir, attOffAir, attDropDate, attMulticast FROM att"
                                            SQLQuery = SQLQuery + " WHERE (attShfCode = " & ilShttCode2 & " AND attVefCode = " & ilVefCode & ")"
                                            Set att_rst = gSQLSelectCall(SQLQuery)
                                            While Not att_rst.EOF
                                                If gDateValue(att_rst!attOffAir) <= gDateValue(att_rst!attDropDate) Then
                                                    If llLastDate2 = gDateValue(att_rst!attOffAir) Then
                                                        ilFound = True
                                                        ilLoopIndex = ilLoop2
                                                        Exit For
                                                    End If
                                                Else
                                                    If llLastDate2 = gDateValue(att_rst!attDropDate) Then
                                                        ilFound = True
                                                        ilLoopIndex = ilLoop2
                                                        Exit For
                                                    End If
                                                End If
                                                att_rst.MoveNext
                                            Wend
                                        End If
                                    End If
                                Next ilLoop2
                                If ilFound Then
                                    'Move
                                    llRow = grdInclude.FixedRows
                                    For ilLoop2 = grdInclude.Rows - 1 To grdInclude.FixedRows Step -1
                                        If grdInclude.TextMatrix(ilLoop2, CALLLETTERSINDEX) <> "" Then
                                            llRow = ilLoop2 + 1
                                            Exit For
                                        End If
                                    Next ilLoop2
                                    mMoveGridRow grdExclude, grdInclude, CLng(ilLoopIndex), llRow
                                End If
                            End If
                        End If
                        shtt_rst.MoveNext
                    Wend
                End If
            End If
        End If
    Next ilLoop1
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "gAlignMulticastStations"
    Resume Next
ErrHand1:
    gHandleError "AffErrorLog.txt", "gAlignMulticastStations"
    Return
End Sub

'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - Quick Find by typing on Exclude Grid
Sub mFindMatch(sSearch As String, grd As MSHFlexGrid)
    Dim ilLoop As Integer
    For ilLoop = grd.FixedRows To grd.Rows - 1
        If Mid(grd.TextMatrix(ilLoop, CALLLETTERSINDEX), 1, Len(sSearch)) = sSearch Then
            grd.Row = ilLoop: grd.TopRow = ilLoop
            Exit Sub
        End If
    Next ilLoop
End Sub

'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - this uses tgStationInfo, no lookups for ownwer, etc.. tgStationInfo updated to include these values
Private Sub mShowAllStationsNew()
    Dim ilLoop As Integer
    Dim ilAddStation As Integer
    Dim llRow As Long
    Dim llIndex As Long
    Dim slRank As String
    Dim slFormat As String
    Dim slOwner As String
    Dim ilDoe As Integer
    
    ProgressBar1.Value = 0
    ProgressBar1.Max = 100
    ProgressBar1.Visible = True
    grdExclude.Rows = grdExclude.FixedRows + 1
    lblExclude.Caption = "Stations to Exclude"
    Screen.MousePointer = vbHourglass
    DoEvents
    grdExclude.Visible = False
    'llRow = grdExclude.FixedRows
    For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        ilAddStation = True
        If imBaseShttCode = tgStationInfo(ilLoop).iCode Then ilAddStation = False
        If tgStationInfo(ilLoop).sUsedForATT <> "Y" Then ilAddStation = False
        'slRank = "": slFormat = "": slOwner = ""
        If ilAddStation Then
            If grdExclude.TextMatrix(1, CALLLETTERSINDEX) = "" Then
                'already a blank row avail
                grdExclude.TextMatrix(1, CALLLETTERSINDEX) = Trim$(tgStationInfo(ilLoop).sCallLetters)
                grdExclude.TextMatrix(1, MARKETINDEX) = Trim$(tgStationInfo(ilLoop).sMarket)
                grdExclude.TextMatrix(1, FORMATINDEX) = Trim$(tgStationInfo(ilLoop).sFormat)
                grdExclude.TextMatrix(1, OWNERINDEX) = Trim$(tgStationInfo(ilLoop).sOwner)
                grdExclude.TextMatrix(1, ZONEINDEX) = Trim$(tgStationInfo(ilLoop).sZone)
                grdExclude.TextMatrix(1, DATERANGEINDEX) = ""
                grdExclude.TextMatrix(1, RANKINDEX) = Trim$(tgStationInfo(ilLoop).sRank)
                grdExclude.TextMatrix(1, SHTTCODEINDEX) = tgStationInfo(ilLoop).iCode
            Else
                'add row
                grdExclude.AddItem Trim$(tgStationInfo(ilLoop).sCallLetters) & vbTab & Trim$(tgStationInfo(ilLoop).sMarket) & vbTab & Trim$(tgStationInfo(ilLoop).sFormat) & vbTab & Trim$(tgStationInfo(ilLoop).sOwner) & vbTab & Trim$(tgStationInfo(ilLoop).sZone) & vbTab & "" & vbTab & Trim$(tgStationInfo(ilLoop).sRank) & vbTab & tgStationInfo(ilLoop).iCode
            End If
        End If
        ilDoe = ilDoe + 1
        If ilDoe >= 2500 Then
            ProgressBar1.Value = ilLoop * (100 / UBound(tgStationInfo))
            ilDoe = 0
        End If
        
    Next ilLoop
    lblExclude.Caption = "Stations to Exclude"
    
    imLastExcludeColSorted = -1
    imLastExcludeSort = -1
    mSortExcludeCol CALLLETTERSINDEX
    grdExclude.Visible = True
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    ProgressBar1.Visible = False
    mGenOK
    Exit Sub
    
ErrHand:
    Me.Enabled = True
    ProgressBar1.Visible = False
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-mShowAllStations: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

'TTP 11066 - Fast Add: add ability to read in import file that can select Fast Add criteria with different agreement start dates
Function mLoadAndVerifyFastAddFile() As Boolean
    mLoadAndVerifyFastAddFile = False
    Dim ilRet As Integer
    Dim hlFrom As Integer
    Dim slFromFile As String
    Dim slLine As String
    Dim slFields(6) As String
    Dim CallLettersCol As Integer
    Dim VehicleCol As Integer
    Dim StartDateCol As Integer
    Dim PledgesFromCol As Integer
    Dim CopyDeliveryModeCol As Integer
    Dim MulticastGenForCol As Integer
    Dim ilLoop As Integer
    Dim slError As String
    Dim llCurrentRow As Long
    Dim ilShttCode As Integer
    Dim ilVefCode As Integer
    Dim ilRowOkay As Integer

    imFileOkay = False
    llCurrentRow = 1
    grdFastAddFile.Redraw = False
    mClearGrid grdFastAddFile
    mSetGridColumns
    slFromFile = txtFile
    
    If TabStrip1.SelectedItem.Index = 2 Then
        cmdGen.Enabled = False
    End If
    If slFromFile = "" Then Exit Function
    
    ilRet = gFileOpen(slFromFile, "Input Access Read", hlFrom)
    If ilRet <> 0 Then
        MsgBox "Open " & slFromFile & " error#" & Str$(ilRet), vbCritical + vbOKOnly, "Fast Add From File"
        Exit Function
    End If

    Screen.MousePointer = vbHourglass
    grdFastAddFile.MousePointer = flexHourglass
    Me.Refresh
    
    'Open File
    ilRet = 0
    'On Error GoTo mReadFileGlobalErr:
    Line Input #hlFrom, slLine
    slLine = Replace(slLine, "﻿", "")
    
    'Verify Header
    If Len(slLine) > 0 Then
        gParseCDFields slLine, True, slFields()
    End If
    CallLettersCol = -1
    VehicleCol = -1
    StartDateCol = -1
    PledgesFromCol = -1
    CopyDeliveryModeCol = -1
    MulticastGenForCol = -1
    For ilLoop = 0 To UBound(slFields)
        If Trim(slFields(ilLoop)) = "callletters" Then CallLettersCol = ilLoop
        If Trim(slFields(ilLoop)) = "vehicle" Then VehicleCol = ilLoop
        If Trim(slFields(ilLoop)) = "startdate" Then StartDateCol = ilLoop
        If Trim(slFields(ilLoop)) = "pledgesfrom" Then PledgesFromCol = ilLoop
        If Trim(slFields(ilLoop)) = "copydeliverymode" Then CopyDeliveryModeCol = ilLoop
        If Trim(slFields(ilLoop)) = "multicastgenfor" Then MulticastGenForCol = ilLoop
    Next ilLoop
    If CallLettersCol = -1 Or VehicleCol = -1 Or StartDateCol = -1 Or PledgesFromCol = -1 Or CopyDeliveryModeCol = -1 Or MulticastGenForCol = -1 Then
        slError = "Bad File Header!" & vbCrLf & "Please make sure the file header has the following columns:" & vbCrLf & "CallLetters, Vehicle, StartDate, PledgesFrom, CopyDeliveryMode, MulticastGenFor"
        GoTo mReadFileGlobalErr
    End If
    imFileOkay = True
    
    On Error GoTo 0
    'Load records
    Do While Not EOF(hlFrom)
        ilRet = 0
        Line Input #hlFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        
        If Trim(Replace(slLine, ",", "")) = "" Then
            Exit Do
        End If
        
        If Len(slLine) > 0 Then
            ilRowOkay = True
            gParseCDFields slLine, False, slFields()
            If grdFastAddFile.TextMatrix(grdFastAddFile.Rows - 1, FILECALLLETTERS) <> "" Then grdFastAddFile.AddItem ""
            grdFastAddFile.TextMatrix(llCurrentRow, FILELINE) = llCurrentRow + 1
            
            For ilLoop = 0 To UBound(slFields)
                If ilLoop = CallLettersCol Then grdFastAddFile.TextMatrix(llCurrentRow, FILECALLLETTERS) = slFields(ilLoop)
                If ilLoop = VehicleCol Then grdFastAddFile.TextMatrix(llCurrentRow, FILEVEHICLE) = slFields(ilLoop)
                If ilLoop = StartDateCol Then grdFastAddFile.TextMatrix(llCurrentRow, FILESTARTDATE) = slFields(ilLoop)
                If ilLoop = PledgesFromCol Then grdFastAddFile.TextMatrix(llCurrentRow, FILEPLEDGESFROM) = slFields(ilLoop)
                If ilLoop = CopyDeliveryModeCol Then grdFastAddFile.TextMatrix(llCurrentRow, FILEDLVYSVCMODE) = slFields(ilLoop)
                If ilLoop = MulticastGenForCol Then grdFastAddFile.TextMatrix(llCurrentRow, FILEGENFOR) = slFields(ilLoop)
            Next ilLoop
            
            'Verify Gen For
            If Not mVerifyMulticastGenFor(grdFastAddFile.TextMatrix(llCurrentRow, FILEGENFOR)) Then
                grdFastAddFile.TextMatrix(llCurrentRow, FILESTATUS) = "Multicast Gen For invalid"
                grdFastAddFile.Row = llCurrentRow
                grdFastAddFile.Col = FILEGENFOR
                grdFastAddFile.CellBackColor = RGB(255, 232, 232)
                ilRowOkay = False
            End If
            'Verify Delivery Mode
            If ilRowOkay Then
                If Not mVerifyCopyDeliveryMode(grdFastAddFile.TextMatrix(llCurrentRow, FILEDLVYSVCMODE)) Then
                    grdFastAddFile.TextMatrix(llCurrentRow, FILESTATUS) = "Delivery service mode invalid"
                    grdFastAddFile.Row = llCurrentRow
                    grdFastAddFile.Col = FILEDLVYSVCMODE
                    grdFastAddFile.CellBackColor = RGB(255, 232, 232)
                    ilRowOkay = False
                End If
            End If
            'Verify is a Date
            If ilRowOkay Then
                If Not mVerifyDate(grdFastAddFile.TextMatrix(llCurrentRow, FILESTARTDATE)) Then
                    grdFastAddFile.TextMatrix(llCurrentRow, FILESTATUS) = "Start Date invalid"
                    grdFastAddFile.Row = llCurrentRow
                    grdFastAddFile.Col = FILESTARTDATE
                    grdFastAddFile.CellBackColor = RGB(255, 232, 232)
                    ilRowOkay = False
                End If
            End If
            'Convert to Monday
            If ilRowOkay Then
                If Weekday(DateValue(grdFastAddFile.TextMatrix(llCurrentRow, FILESTARTDATE))) <> 1 Then
                    grdFastAddFile.TextMatrix(llCurrentRow, FILESTARTDATE) = gObtainPrevMonday(grdFastAddFile.TextMatrix(llCurrentRow, FILESTARTDATE))
                End If
            End If
            'Verify Call Letters
            If ilRowOkay Then
                ilShttCode = mVerifyCallLetters(grdFastAddFile.TextMatrix(llCurrentRow, FILECALLLETTERS))
                If ilShttCode = -1 Then
                    grdFastAddFile.TextMatrix(llCurrentRow, FILESTATUS) = "Call Letters not found"
                    ilRowOkay = False
                    grdFastAddFile.Row = llCurrentRow
                    grdFastAddFile.Col = FILECALLLETTERS
                    grdFastAddFile.CellBackColor = RGB(255, 232, 232)
                Else
                    grdFastAddFile.TextMatrix(llCurrentRow, FILECALLLETTERSCODE) = ilShttCode
                End If
            End If
            'Verify Pledges From Not Same as Call Letters
            If ilRowOkay Then
                If Trim(LCase(grdFastAddFile.TextMatrix(llCurrentRow, FILECALLLETTERS))) = Trim(LCase(grdFastAddFile.TextMatrix(llCurrentRow, FILEPLEDGESFROM))) Then
                    grdFastAddFile.TextMatrix(llCurrentRow, FILESTATUS) = "Same Call Letters and Pledges from"
                    ilRowOkay = False
                    grdFastAddFile.Row = llCurrentRow
                    grdFastAddFile.Col = FILECALLLETTERS
                    grdFastAddFile.CellBackColor = RGB(255, 232, 232)
                    grdFastAddFile.Col = FILEPLEDGESFROM
                    grdFastAddFile.CellBackColor = RGB(255, 232, 232)
                End If
            End If
            'Verify Pledges From Call Letters
            If ilRowOkay Then
                ilShttCode = mVerifyCallLetters(grdFastAddFile.TextMatrix(llCurrentRow, FILEPLEDGESFROM))
                If ilShttCode = -1 Then
                    grdFastAddFile.TextMatrix(llCurrentRow, FILESTATUS) = "Call Letters not found (Pledges from)"
                    ilRowOkay = False
                    grdFastAddFile.Row = llCurrentRow
                    grdFastAddFile.Col = FILEPLEDGESFROM
                    grdFastAddFile.CellBackColor = RGB(255, 232, 232)
                Else
                    grdFastAddFile.TextMatrix(llCurrentRow, FILEPLEDGESFROMCODE) = ilShttCode
                End If
            End If
            'Verify Vehicle Name
            If ilRowOkay Then
                ilVefCode = mVerifyVehicle(grdFastAddFile.TextMatrix(llCurrentRow, FILEVEHICLE))
                If ilVefCode = -1 Then
                    grdFastAddFile.TextMatrix(llCurrentRow, FILESTATUS) = "Vehicle not found"
                    ilRowOkay = False
                    grdFastAddFile.Row = llCurrentRow
                    grdFastAddFile.Col = FILEVEHICLE
                    grdFastAddFile.CellBackColor = RGB(255, 232, 232)
                Else
                    grdFastAddFile.TextMatrix(llCurrentRow, FILEVEFCODE) = ilVefCode
                End If
            End If
            'Verify Active or Future pledges exist for Vehicle / Pledges from
            If ilRowOkay Then
                If Not mVerifyHasPledges(Val(grdFastAddFile.TextMatrix(llCurrentRow, FILEVEFCODE)), Val(grdFastAddFile.TextMatrix(llCurrentRow, FILEPLEDGESFROMCODE))) Then
                    grdFastAddFile.TextMatrix(llCurrentRow, FILESTATUS) = "Pledges not found for Vehicle"
                    grdFastAddFile.Row = llCurrentRow
                    grdFastAddFile.Col = FILEVEHICLE
                    grdFastAddFile.CellBackColor = RGB(255, 232, 232)
                    grdFastAddFile.Col = FILEPLEDGESFROM
                    grdFastAddFile.CellBackColor = RGB(255, 232, 232)
                    ilRowOkay = False
                End If
            End If
            If ilRowOkay = False Then imFileOkay = False
            llCurrentRow = llCurrentRow + 1
        End If
    Loop
    
    'Verify No Duplicates
    If mCheckForDuplicates = False Then
        imFileOkay = False
    End If
    
    'if there's issues found, Remove the "OK" rows to show only BAD rows
    If imFileOkay = False Then
        For llCurrentRow = grdFastAddFile.Rows - 1 To 1 Step -1
            If grdFastAddFile.TextMatrix(llCurrentRow, FILESTATUS) = "OK: no issues were detected" Then
                grdFastAddFile.RemoveItem (llCurrentRow)
            End If
        Next llCurrentRow
    End If
    grdFastAddFile.TopRow = 1
    grdFastAddFile.Row = 1
    grdFastAddFile.Col = 0
    mLoadAndVerifyFastAddFile = imFileOkay
    
mReadFileGlobalErr:
    grdFastAddFile.MousePointer = flexDefault
    grdFastAddFile.Redraw = True
    Screen.MousePointer = vbDefault
    
    Close #hlFrom
    If Err = 0 And slError <> "" Then
        MsgBox "Error Validating Fast Add File: " & slError, vbCritical + vbOKOnly, "Fast Add from File"
    End If
    If Err <> 0 Then
        MsgBox "Error Validating Fast Add File: " & Err & " - " & Error(Err), vbCritical + vbOKOnly, "Fast Add from File"
    End If
End Function

Function mVerifyCopyDeliveryMode(slvalue As String) As Boolean
    mVerifyCopyDeliveryMode = False
    If Val(slvalue) < 4 And Val(slvalue) > 0 Then mVerifyCopyDeliveryMode = True
End Function

Function mVerifyMulticastGenFor(slvalue As String) As Boolean
    mVerifyMulticastGenFor = False
    If LCase(Trim(slvalue)) = "prior" Then mVerifyMulticastGenFor = True
    If LCase(Trim(slvalue)) = "possible" Then mVerifyMulticastGenFor = True
End Function

Function mVerifyDate(slDate As String) As Boolean
    Dim sDateParts() As String
    mVerifyDate = False
    If gIsDate(slDate) Then
        If Year(DateValue(slDate)) < 1970 Then Exit Function
        If Year(DateValue(slDate)) > 2070 Then Exit Function
        sDateParts = Split(slDate, "/")
        If UBound(sDateParts) <> 2 Then Exit Function
        If Val(sDateParts(0)) < 1 Or Val(sDateParts(0)) > 12 Then Exit Function
        If Val(sDateParts(1)) < 1 Or Val(sDateParts(1)) > 31 Then Exit Function
        mVerifyDate = True
    End If
End Function

Function mVerifyCallLetters(slCallLetters As String) As Integer
    Dim ilLoop As Integer
    mVerifyCallLetters = -1
    For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If LCase(Trim(slCallLetters)) = LCase(Trim(tgStationInfo(ilLoop).sCallLetters)) Then
            mVerifyCallLetters = tgStationInfo(ilLoop).iCode
            Exit For
        End If
    Next ilLoop
End Function

Function mVerifyVehicle(slVehicleName As String) As Integer
    Dim ilLoop As Integer
    mVerifyVehicle = -1
    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        If LCase(Trim(slVehicleName)) = LCase(Trim(tgVehicleInfo(ilLoop).sVehicleName)) Then
            mVerifyVehicle = tgVehicleInfo(ilLoop).iCode
            Exit For
        End If
    Next ilLoop
End Function

Function mVerifyHasPledges(ilVef As Integer, ilShtt As Integer)
    Dim slSQLQuery As String
    Dim rst As ADODB.Recordset
    mVerifyHasPledges = True

    slSQLQuery = "Select attCode from att"
    slSQLQuery = slSQLQuery + " WHERE attVefCode = " & ilVef
    slSQLQuery = slSQLQuery + " AND attShfCode = " & ilShtt
    slSQLQuery = slSQLQuery + " AND attOffAir  = '" & "2069-12-31" & "'"
    slSQLQuery = slSQLQuery + " AND attDropDate = '" & "2069-12-31" & "'"
    
    Set rst = gSQLSelectCall(slSQLQuery)
    If rst.EOF Then
        mVerifyHasPledges = False
    End If
End Function

Function mCheckForDuplicates() As Boolean
    mCheckForDuplicates = True
    Dim llRow As Long
    Dim llRow2 As Long
    Dim CurrentRowValue As String
    Dim CheckRowValue As String
    Dim ilFound As Integer
    For llRow = 1 To grdFastAddFile.Rows - 1
        CurrentRowValue = grdFastAddFile.TextMatrix(llRow, FILECALLLETTERS) & grdFastAddFile.TextMatrix(llRow, FILEVEHICLE)
        ilFound = False
        For llRow2 = 1 To grdFastAddFile.Rows - 1
            If llRow <> llRow2 Then
                CheckRowValue = grdFastAddFile.TextMatrix(llRow2, FILECALLLETTERS) & grdFastAddFile.TextMatrix(llRow2, FILEVEHICLE)
                If LCase(CurrentRowValue) = LCase(CheckRowValue) And grdFastAddFile.TextMatrix(llRow2, FILESTATUS) = "" Then
                    grdFastAddFile.TextMatrix(llRow2, FILESTATUS) = "Duplicate Call Letter/Vehicle, Line:" & llRow + 1
                    ilFound = True
                    mCheckForDuplicates = False
                    grdFastAddFile.Row = llRow2
                    grdFastAddFile.Col = FILEVEHICLE
                    grdFastAddFile.CellBackColor = RGB(255, 232, 232)
                    grdFastAddFile.Col = FILECALLLETTERS
                    grdFastAddFile.CellBackColor = RGB(255, 232, 232)
                End If
            End If
            If ilFound = True Then grdFastAddFile.TextMatrix(llRow, FILESTATUS) = " "
        Next llRow2
        If ilFound = False Then
            If Trim(grdFastAddFile.TextMatrix(llRow, FILESTATUS)) = "" Then
                grdFastAddFile.TextMatrix(llRow, FILESTATUS) = "OK: no issues were detected"
            End If
        End If
    Next llRow
    For llRow = grdFastAddFile.Rows - 1 To 1 Step -1
        If Trim(grdFastAddFile.TextMatrix(llRow, FILESTATUS)) = "" Then
            grdFastAddFile.TextMatrix(llRow, FILESTATUS) = "OK: no issues were detected"
        End If
    Next llRow
End Function

Sub mSelectFromList(olList As ListBox, slItem As String)
    Dim ilLoop As Integer
    For ilLoop = 0 To olList.ListCount - 1
        If Trim(LCase(olList.List(ilLoop))) = Trim(LCase(slItem)) Then
            olList.Selected(ilLoop) = True
        Else
            olList.Selected(ilLoop) = False
        End If
    Next ilLoop
End Sub

Sub mSelectFromCombo(olCombo As ComboBox, slItem As String)
    Dim ilLoop As Integer
    For ilLoop = 0 To olCombo.ListCount - 1
        If Mid(Trim(LCase(olCombo.List(ilLoop))), 1, Len(Trim(LCase(slItem)))) = Trim(LCase(slItem)) Then
            olCombo.ListIndex = ilLoop
            Exit For
        End If
    Next ilLoop
End Sub

Sub mSelectCallLetters(ilShttCode As Integer, slCallLetters As String)
    Dim ilLoop As Integer
    Dim ilLoop2 As Integer
    'cboGetStationsFrom.ListIndex = 1
    'grdInclude.TextMatrix(1, SHTTCODEINDEX) = ilShttCode
    'grdInclude.TextMatrix(1, CALLLETTERSINDEX) = slCallLetters
    
    'JW: v81 TTP 11066 - test results (Thu 6/27/24 11:24 AM) - ISSUE #1 & #2
    mClearGrid grdExclude
    mClearGrid grdInclude
    mShowAllStationsNew
    For ilLoop = grdExclude.FixedRows To grdExclude.Rows - 1
        If LCase(grdExclude.TextMatrix(ilLoop, CALLLETTERSINDEX)) = LCase(Trim(slCallLetters)) Then
            For ilLoop2 = grdInclude.Rows - 1 To 1 Step -1
                If grdInclude.TextMatrix(ilLoop2, CALLLETTERSINDEX) = "" Then
                    mMoveGridRow grdExclude, grdInclude, CLng(ilLoop), CLng(ilLoop2)
                    Exit For
                End If
            Next ilLoop2
            Exit For
        End If
    Next ilLoop
    If grdInclude.TextMatrix(1, CALLLETTERSINDEX) = "" Then
        grdInclude.RemoveItem (1)
    End If
End Sub
    
