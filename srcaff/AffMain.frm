VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Affiliate"
   ClientHeight    =   3495
   ClientLeft      =   165
   ClientTop       =   150
   ClientWidth     =   10650
   Icon            =   "AffMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmcStartUp 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4800
      Top             =   2130
   End
   Begin VB.PictureBox pbcMsgArea 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   10590
      TabIndex        =   0
      Top             =   0
      Width           =   10650
      Begin VB.CommandButton cmcReport 
         Caption         =   "REPORT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8715
         TabIndex        =   13
         Top             =   0
         Width           =   675
      End
      Begin VB.CommandButton cmcTraffic 
         Caption         =   "TRAFFIC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8040
         TabIndex        =   12
         Top             =   0
         Width           =   675
      End
      Begin VB.CommandButton cmcRadar 
         Caption         =   "RADAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7365
         TabIndex        =   11
         ToolTipText     =   "RADAR"
         Top             =   0
         Width           =   675
      End
      Begin VB.CommandButton cmcUser 
         Caption         =   "USER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6705
         TabIndex        =   10
         ToolTipText     =   "User Options"
         Top             =   0
         Width           =   675
      End
      Begin VB.CommandButton cmcSite 
         Caption         =   "SITE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6030
         TabIndex        =   9
         ToolTipText     =   "Site Options"
         Top             =   0
         Width           =   675
      End
      Begin VB.CommandButton cmcExport 
         Caption         =   "EXPORT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5340
         TabIndex        =   8
         ToolTipText     =   "Export Center"
         Top             =   0
         Width           =   765
      End
      Begin VB.CommandButton cmcManagement 
         Caption         =   "MANAGEMENT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4335
         TabIndex        =   7
         ToolTipText     =   "Affiliate Management"
         Top             =   0
         Width           =   1110
      End
      Begin VB.CommandButton cmcPostBuy 
         Caption         =   "POST-BUY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3600
         TabIndex        =   6
         ToolTipText     =   "Post-buy Planning"
         Top             =   0
         Width           =   810
      End
      Begin VB.CommandButton cmcAffIdavit 
         Caption         =   "AFFIDAVIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2880
         TabIndex        =   5
         ToolTipText     =   "Affiliate Affidavits"
         Top             =   0
         Width           =   810
      End
      Begin VB.CommandButton cmcLog 
         Caption         =   "LOG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2220
         TabIndex        =   4
         ToolTipText     =   "Network Log"
         Top             =   0
         Width           =   675
      End
      Begin VB.CommandButton cmcEMail 
         Caption         =   "EMAILS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1545
         TabIndex        =   3
         ToolTipText     =   "Vehicle Emails"
         Top             =   0
         Width           =   675
      End
      Begin VB.CommandButton cmcAgreement 
         Caption         =   "AGREEMENT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   660
         TabIndex        =   2
         ToolTipText     =   "Affiliate Agreements"
         Top             =   0
         Width           =   960
      End
      Begin VB.CommandButton cmcStation 
         Caption         =   "STATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Station Information"
         Top             =   0
         Width           =   675
      End
   End
   Begin VB.Timer tmcWebConnectIssue 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   240
      Top             =   2160
   End
   Begin VB.Timer tmcMonitor 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   885
      Top             =   2280
   End
   Begin VB.Timer tmcClock 
      Interval        =   60000
      Left            =   1530
      Top             =   2430
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2220
      Top             =   2340
   End
   Begin VB.Timer tmcCheckAlert 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3030
      Top             =   2100
   End
   Begin VB.Timer tmcFlashAlert 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3360
      Top             =   2580
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   2010
      Top             =   525
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup"
         Begin VB.Menu mnuCompleteSystemBackup 
            Caption         =   "Complete System Backup"
         End
      End
      Begin VB.Menu mnuImport 
         Caption         =   "&Import"
         Begin VB.Menu mnuImportCSISpots 
            Caption         =   "&Counterpoint Remote Traffic Spots"
         End
         Begin VB.Menu mnuImportWWOSpots 
            Caption         =   "WWO Traffic Spots"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuImportAiredStationSpots 
            Caption         =   "Univision &Aired Station Spots"
         End
         Begin VB.Menu mnuWebImportAiredStationSpot 
            Caption         =   "&Counterpoint Affidavit"
         End
         Begin VB.Menu mnuImportStation 
            Caption         =   "&Station Data"
         End
         Begin VB.Menu mnuImportSpots 
            Caption         =   "S&pot Data"
            Begin VB.Menu mnuImportLogSpots 
               Caption         =   "&Log"
            End
            Begin VB.Menu mnuImportAffiliateSpots 
               Caption         =   "&Affiliate"
            End
            Begin VB.Menu mnuImportMYLSpots 
               Caption         =   "&MYL"
            End
         End
         Begin VB.Menu mnuImportAgree 
            Caption         =   "&Agreement Data"
            Begin VB.Menu mnuImportAirDates 
               Caption         =   "&Air Dates"
            End
            Begin VB.Menu mnuImportAgreePledge 
               Caption         =   "&Pledge Times"
            End
            Begin VB.Menu mnuImportCPs 
               Caption         =   "&CP's"
            End
         End
         Begin VB.Menu mnuImportAffAE 
            Caption         =   "&Affiliate A/E Assignment"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuImportCSVAffidavit 
            Caption         =   "CSV Affidavit"
         End
         Begin VB.Menu mnuImportMarketron 
            Caption         =   "Mar&ketron"
         End
         Begin VB.Menu mnuImportWegenerCompel 
            Caption         =   "&Wegener-Compel"
         End
         Begin VB.Menu mnuImportIpump 
            Caption         =   "Wegener-IPump"
         End
         Begin VB.Menu mnuImportXDigital 
            Caption         =   "&X-Digital"
         End
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export"
         Begin VB.Menu mnuExportSpec 
            Caption         =   "Specifications"
         End
         Begin VB.Menu mnuExportSchdStationSpots 
            Caption         =   "Univision &Scheduled Station Spots"
         End
         Begin VB.Menu mnuWebExportSchdStationSpots 
            Caption         =   "&Counterpoint Affidavit"
         End
         Begin VB.Menu mnuExportISCI 
            Caption         =   "&ISCI"
         End
         Begin VB.Menu mnuExportStarGuide 
            Caption         =   "&StarGuide"
         End
         Begin VB.Menu mnuExportCnCSpots 
            Caption         =   "Clearance n Compensation"
         End
         Begin VB.Menu mnuExportRCS4 
            Caption         =   "RCS 4 Digit Cart #'s"
         End
         Begin VB.Menu mnuExportRCS5 
            Caption         =   "RCS 5 Digit Cart #'s"
         End
         Begin VB.Menu mnuExportLabelInfo 
            Caption         =   "Label Info"
         End
         Begin VB.Menu mnuExportRadar 
            Caption         =   "RA&DAR"
         End
         Begin VB.Menu mnuExportXDigital 
            Caption         =   "&X-Digital"
         End
         Begin VB.Menu mnuExportISCIXref 
            Caption         =   "&ISCI Cross Reference"
         End
         Begin VB.Menu mnuExportWegener 
            Caption         =   "W&egener Compel"
         End
         Begin VB.Menu mnuExportIPump 
            Caption         =   "Wegener &IPump"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuExportMarketron 
            Caption         =   "Mar&ketron"
         End
         Begin VB.Menu mnuExportStationInformation 
            Caption         =   "Station Information"
         End
         Begin VB.Menu mnuExportIDC 
            Caption         =   "&IDC"
         End
         Begin VB.Menu mnuExportStationCompensation 
            Caption         =   "Station Compensation"
         End
         Begin VB.Menu mnuExportOverdue 
            Caption         =   "Affidavit Overdue"
         End
      End
      Begin VB.Menu mnuMerge 
         Caption         =   "&Merge"
         Begin VB.Menu mnuMergeFormats 
            Caption         =   "&Formats"
         End
         Begin VB.Menu mnuMergerMarkets 
            Caption         =   "&Markets"
         End
      End
      Begin VB.Menu mnuManageFormats 
         Caption         =   "Manage &Formats"
         Begin VB.Menu mnuManageFormatsNew 
            Caption         =   "&New and Change"
         End
         Begin VB.Menu mnuManageFormatsXRef 
            Caption         =   "&Update Cross References"
         End
      End
      Begin VB.Menu mnuUtilities 
         Caption         =   "&Utilities"
         Begin VB.Menu mnuAstCheckUtil 
            Caption         =   "AST Check Utility"
         End
         Begin VB.Menu mnuUtilitiesAvgWksDelinq 
            Caption         =   "Calculate Avg Weeks Delinquent in Posting"
         End
         Begin VB.Menu mnuEmailFormat 
            Caption         =   "Check Email Format"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuUtilitiesDuplSHTTFix 
            Caption         =   "Clean-up Station Info"
         End
         Begin VB.Menu mnuUtilitiesCPTTAgree 
            Caption         =   "CPTT Agree between Affiliate and Web"
         End
         Begin VB.Menu mnuUtilitiesCPTTCheck 
            Caption         =   "CPTT Check"
         End
         Begin VB.Menu mnuComplianceTracer 
            Caption         =   "Compliance Tracer"
         End
         Begin VB.Menu mnuUtilitiesDuplBkoutFix 
            Caption         =   "Duplicate Blackout Fix"
         End
         Begin VB.Menu mnuUtilitiesDuplCPTTFix 
            Caption         =   "Duplicate CPTT Fix"
         End
         Begin VB.Menu mnuUtilitiesReImportWebSpots 
            Caption         =   "Re-Import Affiliate Spots"
         End
         Begin VB.Menu mnuUtilitiesResetCompliant 
            Caption         =   "Reset Program Times and Compliant Counts"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuUtilitiesSetCompliant 
            Caption         =   "Set Compliance"
         End
         Begin VB.Menu mnuUtilitiesSetMG 
            Caption         =   "Set MG's"
         End
         Begin VB.Menu mnuUtilitiesSetPrgTimes 
            Caption         =   "Set Program Times"
         End
         Begin VB.Menu mnuUtilitiesSpotCount 
            Caption         =   "Spot Count Tie-out"
         End
         Begin VB.Menu mnuUtilitiesWebPosting 
            Caption         =   "Unpost and or Delete Spot"
         End
         Begin VB.Menu mnuViewSql 
            Caption         =   "View Sql"
         End
         Begin VB.Menu mnuUtilitiesWegenerCheckUtility 
            Caption         =   "Wegener Check Utility"
         End
      End
      Begin VB.Menu mnuWebVendors 
         Caption         =   "&Vendor Setup"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSSB 
         Caption         =   "&Station Spot Builder Status"
      End
      Begin VB.Menu mnuFileRQS 
         Caption         =   "&Report Queue Status"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrtSetup 
         Caption         =   "&Printer Setup"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAccessories 
      Caption         =   "&Accessories"
      Begin VB.Menu mnuAccessoriesMessages 
         Caption         =   "&Messages"
      End
      Begin VB.Menu mnuAccessoriesDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccessoriesUsersStatus 
         Caption         =   "&Users Status"
      End
      Begin VB.Menu mnuAccessoriesDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccessoriesDate 
         Caption         =   "Counterpoint Date"
      End
      Begin VB.Menu mnuAccessoriesDash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccessoriesViewBlocks 
         Caption         =   "View Blocks"
      End
   End
   Begin VB.Menu mnuGroupName 
      Caption         =   "&Group Name"
      Begin VB.Menu mnuGroupNameDMAMarket 
         Caption         =   "&DMA Market"
      End
      Begin VB.Menu mnuGroupNameFormat 
         Caption         =   "&Format"
      End
      Begin VB.Menu mnuGroupNameMSAMarket 
         Caption         =   "&MSA Market"
      End
      Begin VB.Menu mnuGroupNameTimeZone 
         Caption         =   "&Time Zone"
      End
      Begin VB.Menu mnuGroupNameState 
         Caption         =   "&State"
      End
      Begin VB.Menu mnuGroupNameVehicle 
         Caption         =   "&Vehicle"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpGoToWeb 
         Caption         =   "&Counterpoint Web Site"
         Begin VB.Menu mnuDocumentation 
            Caption         =   "&Documentation"
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnuWebHome 
            Caption         =   "&Home Page"
         End
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Affiliate..."
      End
   End
   Begin VB.Menu mnuAlert 
      Caption         =   ""
      Begin VB.Menu mnuAlertView 
         Caption         =   "View"
      End
   End
   Begin VB.Menu mnuBlank1 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "css"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuCSSName 
         Caption         =   "Contract Spot Scheduler"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCSSStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "ssb"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mnuSSBName 
         Caption         =   "Station Spot Builder"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSSBStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "aeq"
      Index           =   2
      Visible         =   0   'False
      Begin VB.Menu mnuAEQName 
         Caption         =   "Affiliate Export Queue"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAEQStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "asi"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mnuASIName 
         Caption         =   "Affidavit Spot Import"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuASIStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "amb"
      Index           =   4
      Visible         =   0   'False
      Begin VB.Menu mnuAMBName 
         Caption         =   "Affiliate Measurement Builder"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAMBStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "arq"
      Index           =   5
      Visible         =   0   'False
      Begin VB.Menu mnuARQName 
         Caption         =   "Affiliate Report Queue"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuARQStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "asg"
      Index           =   6
      Visible         =   0   'False
      Begin VB.Menu mnuASGName 
         Caption         =   "Avail Summary Generation"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuASGStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "sc"
      Index           =   7
      Visible         =   0   'False
      Begin VB.Menu mnuSCName 
         Caption         =   "Set Credit"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSCStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "ce"
      Index           =   8
      Visible         =   0   'False
      Begin VB.Menu mnuCEName 
         Caption         =   "Corporate Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCEStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "sfe"
      Index           =   9
      Visible         =   0   'False
      Begin VB.Menu mnuSFEName 
         Caption         =   "Sales Force Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSFEStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "me"
      Index           =   10
      Visible         =   0   'False
      Begin VB.Menu mnuMEName 
         Caption         =   "Matrix Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuMEStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "epe"
      Index           =   11
      Visible         =   0   'False
      Begin VB.Menu mnuEPEName 
         Caption         =   "Efficio Projection Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEPEStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "ere"
      Index           =   12
      Visible         =   0   'False
      Begin VB.Menu mnuEREName 
         Caption         =   "Efficio Revenue Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEREStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "gpe"
      Index           =   13
      Visible         =   0   'False
      Begin VB.Menu mnuGPEName 
         Caption         =   "Get Paid Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGPEStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "bd"
      Index           =   14
      Visible         =   0   'False
      Begin VB.Menu mnuBDName 
         Caption         =   "Backup Database"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBDStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "te"
      Index           =   15
      Visible         =   0   'False
      Begin VB.Menu mnuTeName 
         Caption         =   "Tableau  Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTeStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "wvm"
      Enabled         =   0   'False
      Index           =   16
      Visible         =   0   'False
      Begin VB.Menu mnuWviName 
         Caption         =   "Web Vendor Manager"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWviStatus 
         Caption         =   "See Alerts-Web Vendors"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "pb"
      Index           =   17
      Visible         =   0   'False
      Begin VB.Menu mnuPbName 
         Caption         =   "Programmatic Buy"
      End
      Begin VB.Menu mnuPbStatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "cai"
      Index           =   18
      Visible         =   0   'False
      Begin VB.Menu mnuCAIName 
         Caption         =   "Compel Auto Import"
      End
      Begin VB.Menu mnuCAIStatus 
         Caption         =   "Not Working, Get Help!"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "re"
      Index           =   19
      Visible         =   0   'False
      Begin VB.Menu mnuREName 
         Caption         =   "RAB Export"
      End
      Begin VB.Menu mnuREstatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMonitor 
      Caption         =   "cre"
      Index           =   20
      Visible         =   0   'False
      Begin VB.Menu mnuCREName 
         Caption         =   "Custom Revenue Export"
      End
      Begin VB.Menu mnuCREstatus 
         Caption         =   "Not Working, Get Help!"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuDate 
      Caption         =   ""
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmMain - MDI parent form
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private dmShellRet As Double

Const clMaxName As Long = 255
Const GW_CHILD = 5
Const GW_HWNDNEXT = 2
Const GW_HWNDFIRST = 0
Const WM_SETTEXT = &HC

Private Declare Function GetWindow Lib "user32" (ByVal hwnd _
    As Long, ByVal wCmd As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

    
Dim hmUlf As Integer
Dim tmUlfSrchKey As LONGKEY0
Dim tmUlfSrchKey1 As ULFKEY1
Dim tmUlf As ULF
Dim imUlfRecLen As Integer

Dim hmCef As Integer    'Comment file handle
Dim tmCef As CEF        'CEF record image
Dim tmCefSrchKey As LONGKEY0    'CEF key record image
Dim imCefRecLen As Integer        'CEF record length

Dim imCountdown As Integer
Dim lmCountdownEnteredTime As Long
Dim imPrevAlertBlockChk As Integer

Private rst_Uaf As ADODB.Recordset

Private rst_Urf As ADODB.Recordset
Private rst_Ust As ADODB.Recordset

Private sgDateBrannerMsg As String

Dim tmAuf As AUF
Dim hmAlert As Long
Dim lmMenuID As Long
Dim hmRedAlertBitmap As Long
Dim hmWhiteAlertBitmap As Long

Dim hmMonitor As Long
'Dim hmMonitorBitmap(0 To 13) As Long
'Dim imMonitor As Integer
Dim tmTmf As TMF
Dim imTmfRecLen As Integer
Dim tmTmfSrchKey1 As TMFKEY1
Dim lmDate1970 As Long

Dim imShowVersionNo As Integer
Dim smWallpaper As String

Private tmFormatLinkInfo() As FORMATLINKINFO

Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
''Dan M 6/18/10 for help going to web
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim imLastHourGGChecked As Integer
'8793
Private emMonitoringAllowed As MonitorAllowed
Private smMonitorID As String
Private Enum MonitorAllowed
    MonitorOff = 0
    MonitorNormal = 1
    MonitorNoEmail = 2
End Enum
Dim bmActiveOnce As Boolean

Private Function mCopyPictureImage(SourceImage As PictureBox) As Long
    Dim bm As BITMAP
    Dim newbm&
    Dim tdc&, oldbm&
    Dim di&

    ' First get the information about the image bitmap
    di = GetObjectAPI(SourceImage.Image, Len(bm), bm)
    bm.bmBits = 0
    ' Create a new bitmap with the same structure and size
    ' of the image bitmap
    newbm = CreateBitmapIndirect(bm)

    ' Create a temporary memory device context to use
    tdc = CreateCompatibleDC(SourceImage.hDC)
    ' Select in the newly created bitmap
    oldbm = SelectObject(tdc, newbm)

    ' Now copy the bitmap from the persistant bitmap in
    ' picture 2 (note that picture2 has AutoRedraw set TRUE
    di = BitBlt(tdc, 0, 0, bm.bmWidth, bm.bmHeight, SourceImage.hDC, 0, 0, SRCCOPY)
    ' Select out the bitmap and delete the memory DC
    oldbm = SelectObject(tdc, oldbm)
    di = DeleteDC(tdc)

    ' And return the new bitmap
    mCopyPictureImage = newbm
End Function

Private Sub mShowAlert(hlBitmap As Long)
    Dim ilRet As Integer
    Dim llMenuID As Long
    Dim llMenuIndex As Long
    
    'If (gWegenerExport = False) And (gOLAExport = False) Then
    '    llMenuIndex = 3
    'Else
    '    llMenuIndex = 4
    'End If
    
    ' Get a handle to the top level menu
    hmAlert = GetMenu(frmMain.hwnd)
    
    'Get AlertMenu Index
    llMenuIndex = GetAlertMenuIndex(hmAlert, 0)
    lmMenuID = GetMenuItemID(hmAlert, llMenuIndex)
    
    'If hmBitmap = 0 Then hmBitmap = CopyPictureImage(frmDirectory!pbcAlert)
    ' And replace it with a bitmap.
    ilRet = ModifyMenuBynum(hmAlert, llMenuIndex, MF_BITMAP Or MF_BYPOSITION Or MF_ENABLED, llMenuID, hlBitmap)
End Sub

Private Sub cmcAffIdavit_Click()
    frmDirectory.cmdCPReturns_Click
End Sub

Private Sub cmcAgreement_Click()
    frmDirectory.cmdAgreements_Click
End Sub

Private Sub cmcEMail_Click()
    frmDirectory.cmdEMail_Click
End Sub

Private Sub cmcExport_Click()
    frmDirectory.cmdExports_Click
End Sub

Private Sub cmcLog_Click()
    frmDirectory.cmdPostLog_Click
End Sub

Private Sub cmcManagement_Click()
    frmDirectory.cmcManagement_Click
End Sub

Private Sub cmcPostBuy_Click()
    frmDirectory.cmcPostBuy_Click
End Sub

Private Sub cmcRadar_Click()
    frmDirectory.cmdRadar_Click
End Sub

Private Sub cmcReport_Click()
    igReportSource = 1
    frmReports.Show
End Sub

Private Sub cmcSite_Click()
    frmDirectory.cmdSite_Click
End Sub

Private Sub cmcStation_Click()
    frmDirectory.cmdStation_Click
End Sub

Private Sub cmcTraffic_Click()
    Dim ilShell As Integer
    Dim slCommandStr As String
    Dim ilPos As Integer
    Dim slDate As String
    Dim blStart As Boolean
    Dim ilRet As Integer
    
    Dim hlWndTraffic As Long
    Dim hlWndChild As Long
    Dim slTestName As String * 255
    Dim llNumChars As Long

    
    If igTestSystem Then
        slCommandStr = "Affiliate^Test\" & sgUserNameToPassToTraffic & "\" & sgUserPasswordToPassToTraffic & "\" & "FROMAFFILIATE"
    Else
        slCommandStr = "Affiliate^Prod\" & sgUserNameToPassToTraffic & "\" & sgUserPasswordToPassToTraffic & "\" & "FROMAFFILIATE"
    End If
    
    slDate = Trim$(sgNowDate)
    If slDate <> "" Then
        slDate = " /D:" & slDate
        slCommandStr = slCommandStr & slDate
    End If
    'slCommandStr = slCommandStr & " /FROMAFFILIATE"
    On Error GoTo LoadErr
    ilRet = 0
    'AppActivate "CSI Traffic"
    'If ilRet = 1 Then
    '    dmShellRet = Shell(sgExeDirectory & "Traffic2.Exe " & slCommandStr, vbNormalFocus)
    'Else
    hlWndTraffic = FindWindow(vbNullString, "TrafficAffiliateCom")
    If hlWndTraffic = 0 Or ilRet = 1 Then
        ilRet = 0
        AppActivate "CSI Traffic"
        If ilRet = 1 Then
            dmShellRet = Shell(sgExeDirectory & "Traffic2.Exe " & slCommandStr, vbNormalFocus)
        End If
    Else
        hlWndChild = GetWindow(hlWndTraffic, GW_CHILD)
        Do While hlWndChild <> 0
            llNumChars = GetClassName(hlWndChild, slTestName, (clMaxName))
            If InStr(1, slTestName, "TextBox") > 0 Then
                SendMessageByString& hlWndChild, WM_SETTEXT, 0, "Set Window size"
                ilRet = 0
                AppActivate "CSI Traffic"
                If ilRet = 1 Then
                    dmShellRet = Shell(sgExeDirectory & "Traffic2.Exe " & slCommandStr, vbNormalFocus)
                End If
                Exit Do
            End If
            hlWndChild = GetWindow(hlWndChild, GW_HWNDNEXT)
        Loop
    End If
    Exit Sub
LoadErr:
    ilRet = 1
    Resume Next


End Sub

Private Sub cmcUser_Click()
    frmDirectory.cmdOptions_Click
End Sub

Private Sub MDIForm_Activate()
    If bmActiveOnce = True Then Exit Sub
    mInitMonitorMenu
    
    'draw background after it's been sized
    frmMstPict.mMstPictSetMsg

    bmActiveOnce = True
End Sub

Private Sub MDIForm_Load()

    Dim slUserName As String
    Dim ilRet As Integer
    Dim slMessage As String
    Dim slDateTime As String
    Dim ilValue10 As Integer
    Dim slRevision As String
        
    imLastHourGGChecked = -1    'Hour(Now)
    
    imPrevAlertBlockChk = 0
    '5/17/13: Moved to Login
    ''Start the Pervasive API engine
    'If Not mOpenPervasiveAPI Then
    '    tmcTerminate.Enabled = True
    '    Exit Sub
    'End If
   '
    'If Not gCheckDDFDates() Then
    '    tmcTerminate.Enabled = True
    '    Exit Sub
    'End If
    
    'Test if setDDFFields needs to be run
    If Not igAutoImport And Not igCompelAutoImport Then
        ilRet = gPopAll
    End If
    
    '11/8/14: Restrict Remote Export
    If (StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) <> 0) And (StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) <> 0) Then
        If bgRemoteExport Then
            mnuFile.Enabled = False
        End If
    End If
    
    mSetCommentSource
    If Not mCheckSetDDFFields() Then
        tmcTerminate.Enabled = True
        Exit Sub
    End If
    imShowVersionNo = igShowVersionNo
    smWallpaper = sgWallpaper


    ilRet = gAlertCheckBlock(tmAuf)
    If ilRet = 1 Then   'Block user
        slMessage = "System is shut down until further notice"
        MsgBox slMessage, vbCritical, "Shut down"
        tmcTerminate.Enabled = True
        Exit Sub
    ElseIf ilRet = 2 Then   'Block Initiated by this user
    End If
    
    If lgEMailCefCode > 0 Then
        ilRet = mOpenCEFFile()
        ilRet = mGetCefComment(lgEMailCefCode, sgEMail)
        ilRet = mCloseCEFFile()
    End If
    
    'Returns NULLs proper to the field datatype to prevent errors
        
    ilRet = mSignOnOff(True)
    
    If igSQLSpec = 0 Then   'Pervasive 7
        SQLQuery = "Set Stringnull = ''"
        'cnn.Execute SQLQuery          ', rdExecDirect
        Call gSQLWaitNoMsgBox(SQLQuery, False)
        SQLQuery = "Set Binarynull = 0"
        'cnn.Execute SQLQuery          ', rdExecDirect
        Call gSQLWaitNoMsgBox(SQLQuery, False)
    End If
    
    If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
        slUserName = "System"
    Else
        slUserName = sgUserName
    End If
    
    slRevision = App.Revision
    If Len(slRevision) = 3 Then
        slRevision = "0" & slRevision
    End If
    sgDateBrannerMsg = slUserName & " on Counterpoint ®, " & "V" & App.Major & "." & App.Minor & " B" & slRevision & " " & App.FileDescription & " for " & sgClientName
    slDateTime = " " & Format$(gNow(), "ddd, m/d/yy h:mm AM/PM")
    mnuDate.Caption = slDateTime & " " & sgDateBrannerMsg
    frmMain.Caption = "CSI Affiliate"    'slUserName & " on Counterpoint ®, " & "V" & App.Major & "." & App.Minor & " B" & slRevision & " " & App.FileDescription & " for " & sgClientName
    'Dan M 8/31/10 test for updates in case user is skipping 'affiliat.exe'
    If mUpdatesAvailable() Then
    '5676
        gLogMsg "An update is available.  Please run Affiliat.exe or Traffic.exe to install updates", "UpdateErrors.txt", False
        'gMsgBox "An update is available.  Please run Affiliat.exe or Traffic.exe to install updates", , "Updates Available"
    End If
        '5666
    If Len(sgShortDate) = 0 Then
        sgShortDate = gRegistryGetShortDate()
        If Len(sgShortDate) > 0 And sgShortDate <> MYSHORTDATE Then
            gRegistrySetShortDate (MYSHORTDATE)
        End If
    End If

    'Dan M 3/9/10 open csiNetReporter
    'Dan 4/22/11 stop using modular level and use global
    'mMakeNetReport
'    gCallNetReporter StartReports
'    'Dan M 8/31/10 test that csiNetReporter started
'     If Not bgReportModuleRunning Then
'        gMsgBox "A component is missing that is needed for reports to run properly.  Please contact Counterpoint.", , "CsiNetReporter Missing"
'     End If
' dan M as per Dick 12/05/11, moved below.
'    mnuImportAiredStationSpots.Enabled = gUsingUnivision
'    mnuExportSchdStationSpots.Enabled = gUsingUnivision
'
'    mnuWebImportAiredStationSpot.Enabled = gUsingWeb
'    mnuWebExportSchdStationSpots.Enabled = gUsingWeb
    
    'If (gUsingUnivision = False) And (gUsingWeb = False) Or ((sgExportISCI <> "A") And (sgExportISCI <> "N")) Then
    If gISCIExport = False Then
        mnuExportISCI.Enabled = False
    End If
    If sgRCSExportCart4 = "N" Then
        mnuExportRCS4.Enabled = False
    End If
    If sgRCSExportCart5 = "N" Then
        mnuExportRCS5.Enabled = False
    End If
    If igTestSystem Then
        frmMain.BackColor = &HC0C0C0
    End If
    If ((Asc(sgSpfUsingFeatures5) And RADAR) <> RADAR) Or (sgUstWin(12) <> "I") Then
        mnuExportRadar.Enabled = False
    End If
    'If ((Asc(sgSpfUsingFeatures2) And SPLITCOPY) <> SPLITCOPY) Then
    If gUsingXDigital = False Then
        mnuImportXDigital.Enabled = False
    End If
    If gUsingXDigital = False Then
        mnuExportXDigital.Enabled = False
    End If
    
    If (gWegenerExport = False) And (gOLAExport = False) Then
        mnuGroupName.Visible = False
    End If
    
    If (gWegenerExport = False) Then
        mnuExportWegener.Enabled = False
        mnuImportWegenerCompel.Enabled = False
    End If
        
    'If (gOLAExport = False) Then
    '    mnuExportOLA.Enabled = False
    'End If
    If sgUsingStationID = "Y" Then
        mnuImportStation.Caption = "Update/Add Stations"
    ElseIf sgUsingStationID = "A" Then
        If UBound(tgStationInfo) > LBound(tgStationInfo) Then
            mnuImportStation.Caption = "Continue Adding Stations"
        Else
            mnuImportStation.Caption = "Add Initial Stations"
        End If
    Else
        If UBound(tgStationInfo) > LBound(tgStationInfo) Then
            mnuImportStation.Caption = "Update Existing Stations"
        Else
            mnuImportStation.Caption = "Add Initial Stations"
        End If
    End If
    If ((Asc(sgSpfUsingFeatures5) And STATIONINTERFACE) <> STATIONINTERFACE) Then
        mnuExportMarketron.Enabled = False
        mnuExportSchdStationSpots.Enabled = False
        mnuWebExportSchdStationSpots.Enabled = False
        mnuImportAiredStationSpots.Enabled = False
        mnuImportMarketron.Enabled = False
        mnuWebImportAiredStationSpot.Enabled = False
        If Not gUsingUnivision Then
            mnuImportAiredStationSpots.Visible = False
            mnuExportSchdStationSpots.Visible = False
        End If
    Else
        mnuImportAiredStationSpots.Enabled = gUsingUnivision
        mnuExportSchdStationSpots.Enabled = gUsingUnivision
        If Not gUsingUnivision Then
            mnuImportAiredStationSpots.Visible = False
            mnuExportSchdStationSpots.Visible = False
        End If
        mnuWebImportAiredStationSpot.Enabled = gUsingWeb
        mnuWebExportSchdStationSpots.Enabled = gUsingWeb
    End If
    

    sgSplitState = "M"
    SQLQuery = "Select safFeatures1, safFeatures3, safFeatures5, safFeatures6 From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set rst = gSQLSelectCall(SQLQuery, "frmMain: Load")
    If Not rst.EOF Then
        ilValue10 = Asc(rst!safFeatures1)
        If (ilValue10 And COMPENSATION) <> COMPENSATION Then
            mnuExportStationCompensation.Enabled = False
        End If
        ilValue10 = Asc(rst!safFeatures3)
        If (ilValue10 And SPLITCOPYLICENSE) = SPLITCOPYLICENSE Then
            sgSplitState = "L"
        ElseIf (ilValue10 And SPLITCOPYPHYSICAL) = SPLITCOPYPHYSICAL Then
            sgSplitState = "P"
        End If
        'D.S. 3/21/18
        If (Asc(rst!safFeatures5) And PROGRAMMATICALLOWED) <> PROGRAMMATICALLOWED Then
            mnuMonitor(17).Visible = False
        End If
        
        'If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
            If (Asc(rst!safFeatures5) And CSVAFFIDAVITIMPORT) <> CSVAFFIDAVITIMPORT Then
                mnuImportCSVAffidavit.Visible = False
            End If
        'Else
        '    mnuImportCSVAffidavit.Visible = False
        'End If
        If (Asc(rst!safFeatures6) And OVERDUEEXPORT) <> OVERDUEEXPORT Then
            mnuExportOverdue.Enabled = False
        End If
    End If
    If igGGFlag = 0 Then
        mnuBackup.Enabled = False
        mnuImport.Enabled = False
        mnuExport.Enabled = False
        mnuMerge.Enabled = False
        mnuManageFormats.Enabled = False
        mnuUtilities.Enabled = False
        mnuGroupName.Enabled = False
    End If
    
    If (Not bgReportQueue) Then
        mnuFileBar2.Visible = False
        mnuFileRQS.Visible = False
    End If
    
    
    mLoadPicture
    
    'mSetJobButtons     'Move to timer control
    
    Load frmDirectory
    
    lgShellAndWaitID = 0
    Load AffiliateTrafficCom
    'Uncomment this to make active the alert as to backup programming not running
'    If igNumMoBehind > 0 Then
'        gMsgBox "Your automatic spot archiving program is not running." & Chr(13) & Chr(10) & "There are " & igNumMoBehind & " month(s) waiting to be archived." & Chr(13) & Chr(10) & "Please call Counterpoint.", vbOKOnly, "Warning"
'    End If
    
    Unload frmLogin
    
    frmDirectory!pbcRedAlert.Width = 64
    frmDirectory!pbcRedAlert.Height = 18
    frmDirectory!pbcWhiteAlert.Width = 64
    frmDirectory!pbcWhiteAlert.Height = 18
    
    hmRedAlertBitmap = mCopyPictureImage(frmDirectory!pbcRedAlert)
    hmWhiteAlertBitmap = mCopyPictureImage(frmDirectory!pbcWhiteAlert)

    frmDirectory!pbcMonitor.Width = 26
    frmDirectory!pbcMonitor.Height = 18
    frmDirectory!pbcMonitor.FontSize = 8
    frmDirectory!pbcMonitor.FontName = "Arial Narrow"
    
    

    slDateTime = " " & Format$(gNow(), "ddd, m/d/yy h:mm AM/PM")
    mnuDate.Caption = slDateTime & " " & sgDateBrannerMsg

    ilRet = gAlertForceCheck()
    'D.S. 09/04/02
    'This call won't work here.  I moved it to frmLogin CmdOK click
    'mLoadPicture
    'ttp 5552
   ' If gIsInternalGuide() Then
    mnuEmailFormat.Visible = True
    mnuEmailFormat.Enabled = True
   ' End If
'   '6079
    mnuExportIPump.Visible = True
    mnuImportIpump.Visible = True
   '8156 don't need anymore
'    SQLQuery = "Select spfUsingFeatures10 From SPF_Site_Options"
'    Set rst = gSQLSelectCall(SQLQuery)
'    If Not rst.EOF Then
'        ilValue10 = Asc(rst!spfUsingFeatures10)
'        If (ilValue10 And WEGENERIPUMP) = WEGENERIPUMP Then
'            mnuExportIPump.Enabled = True
'            mnuImportIpump.Enabled = True
'        Else
'            mnuExportIPump.Enabled = False
'            mnuImportIpump.Enabled = False
'        End If
'    End If
    '7701
    If Not mConvertToVAT() Then
        mUpdateVendors
    End If
    '7912
    mnuImportXDigital.Visible = False
    '7967  8129
    '  this is wvi monitor: mnumonitor(16).Visible = False
'    If gVendorToWebAllowed(tmcWebConnectIssue) Then
'        gGrabWebVendorIssuesFromWeb True, tmcWebConnectIssue
'    End If
    gVendorToWebAllowed tmcWebConnectIssue
    '8156
    gAdjustAllowedExportsImports
'Dan while testing here.
    If Not gIsInternalGuide() Then
        If Not gIsTrueGuide() Then
            mnuImportIpump.Visible = False
        End If
        mnuImportIpump.Enabled = False
    End If
    '7/27/20: Start timer to adjust display buttons
    tmcStartUp.Enabled = True

End Sub

Private Sub MDIForm_Resize()
    If lgShellAndWaitID <> 0 And frmMain.WindowState <> vbMinimized Then
        frmMain.WindowState = vbMinimized
    End If

    If frmMain.WindowState = vbMaximized Then
        frmMstPict.Move 0, 0, Me.Width - 240, Me.Height - 1800
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim ilRet As Integer
    Dim slCommandLine As String
    On Error Resume Next
    mUnloadForms
    
    If (sgSQLTrace = "Y") And (hgSQLTrace >= 0) Then
        gLogMsgWODT "W", hgSQLTrace, "SQL Overall Time: " & gTimeString(lgTtlTimeSQL / 1000, True)
        gLogMsgWODT "C", hgSQLTrace, ""
    End If
    
    ilRet = mSignOnOff(False)
    Erase tgCifCpfInfo1
    Erase tgCrfInfo1
    Erase lgUserLogUlfCode
    Erase tgCopyRotInfo
    Erase tgGameInfo
    Erase tgStationInfoByCode
    Erase tgCpfInfo
    Erase tgMarketInfo
    Erase tgMSAMarketInfo
    Erase tgTerritoryInfo
    Erase tgCityInfo
    Erase tgCountyInfo
    Erase tgAreaInfo
    Erase tgMonikerInfo
    Erase tgOperatorInfo
    Erase tgMarketRepInfo
    Erase tgServiceRepInfo
    Erase tgAffAEInfo
    Erase tgSellingVehicleInfo
    Erase tgVpfOptions
    Erase tgLstInfo
    Erase tgAttInfo1
    Erase tgShttInfo1
    Erase tgCpttInfo
    Erase sgAufsKey
    Erase tgRBofRec
    Erase tgSplitNetLastFill
    Erase tgAvailNamesInfo
    Erase tgMediaCodesInfo
    Erase tgTitleInfo
    Erase tgOwnerInfo
    Erase tgFormatInfo
    Erase tgVffInfo
    Erase tgTeamInfo
    Erase tgLangInfo
    Erase tgTimeZoneInfo
    Erase tgStateInfo
    Erase tgSubtotalGroupInfo
    Erase tgAttExpMon
    Erase tgReportNames
    Erase tgRff
    Erase tgRffExtended
    Erase tgUstInfo
    Erase tgDeptInfo
    
    
    Erase tgStationInfo
    Erase tgVehicleInfo
    Erase tgRnfInfo
    Erase tgAdvtInfo
    Erase sgStationImportTitles
    
    '9/11/06: Split Network stuff
    Erase tgRBofRec
    Erase tgSplitNetLastFill
    
    Erase tmFormatLinkInfo

    If hmAlert <> 0 Then
        Call DeleteObject(hmAlert)
    End If
    On Error Resume Next
    rstAlertUlf.Close
    rst_Uaf.Close
    On Error Resume Next
    rstAlert.Close
    
    ilRet = gCloseMKDFile(hgTmf, "Tmf.btr")
    
    '5666
    If Len(sgShortDate) > 0 And sgShortDate <> MYSHORTDATE Then
        gRegistrySetShortDate (sgShortDate)
    End If
    'gLogMsg "Closing Pervasive API Engine. User: " & gGetComputerName(), "WebExportLog.Txt", False
    mClosePervasiveAPI
    cnn.Close
    'gLogMsg "Pervasive API Engine Closed Successfully. User: " & gGetComputerName(), "WebExportLog.Txt", False
   ' Dan M 12/18/09 test to see if need to close csinetreporter
    'Dan M 12/09/09 don't close csiNetReporter
    'Dan M 11/23/09  close csiNetReporter
On Error Resume Next
'    If bgReportModuleRunning Then
'        gCallNetReporter FinishReports
'    'dan M 4/22/11 use global
''        slCommandLine = mBuildAlternateAsNeeded
''        slCommandLine = slCommandLine & " /Q"
''        Shell slCommandLine
'    End If
''    slCommandLine = mBuildAlternateAsNeeded
''    slCommandLine = slCommandLine & " /Q"
''    Shell slCommandLine
'
'   ' Shell sgExeDirectory & "csinetreporter.exe /Q"
    Set frmMain = Nothing
End Sub
'Private Function mBuildAlternateAsNeeded() As String
'    Dim ilRet As Integer
'    Dim slStr As String
'    Dim slTempPath As String
'
'    If InStr(1, sgExeDirectory, "Test", vbTextCompare) > 0 Then
'        slTempPath = sgExeDirectory & "csinetreporteralternate.exe "
'    Else
'        slTempPath = sgExeDirectory & "csinetreporter.exe "
'    End If
'    ilRet = 0
'    On Error GoTo FileErr
'    slStr = FileDateTime(slTempPath)
'    On Error GoTo 0
'    If ilRet = 1 Then
'        If InStr(1, sgExeDirectory, "Test", vbTextCompare) > 0 Then
'            slTempPath = "csinetreporteralternate.exe "
'        Else
'            slTempPath = "csinetreporter.exe "
'        End If
'    End If
'    mBuildAlternateAsNeeded = slTempPath
'    Exit Function
'FileErr:
'    ilRet = 1
'    Resume Next
'End Function
Private Sub mnuAccessoriesMessages_Click()
    frmMessages.Show vbModal
End Sub

Private Sub mnuAccessoriesUsersStatus_Click()
    frmUsersLog.Show vbModal
End Sub

Private Sub mnuAccessoriesViewBlocks_Click()
    frmViewBlocks.Show vbModal
End Sub

Private Sub mnuAlert_Click()
    frmAlertVw.Show vbModal
End Sub

Private Sub mnuAstCheckUtil_Click()
    igPasswordOk = False
    frmAstCheckUtil.Show vbModal
End Sub

Private Sub mnuCompleteSystemBackup_Click()
    BUZip.Show vbModal
End Sub

Private Sub mnuComplianceTracer_Click()
    'If sgUserName = "Guide" Then
        frmComplianceTracer.Show vbModal
    'Else
    '    frmProgressMsg.SetMessage 1, vbCrLf & "            Under Construction!" & vbCrLf & vbCrLf & "              Check Back Soon!"
    '    frmProgressMsg.Caption = "Compliance Tracer"
    '    frmProgressMsg.Show vbModal
    'End If
End Sub

Private Sub mnuEmailFormat_Click()
    frmAffEmailFormat.Show vbModal
End Sub

Private Sub mnuExportCnCSpots_Click()
    igExportSource = 1
    igExportTypeNumber = 6
    frmExportCnCSpots.Show vbModal
End Sub

Private Sub mnuExportIDC_Click()
    igExportSource = 1
    igExportTypeNumber = 7
    FrmExportIDC.Show vbModal
End Sub

Private Sub mnuExportIPump_Click()
    igExportSource = 1
    igExportTypeNumber = 13
    FrmExportiPump.Show vbModal
End Sub

Private Sub mnuExportOverdue_Click()
    frmExportOverdue.Show vbModeless
End Sub

Private Sub mnuFileSSB_Click()
    frmAstBuildQueueStatus.Show vbModeless
End Sub

Private Sub mnuExportStationCompensation_Click()
    frmExportStationComp.Show vbModal
End Sub

Private Sub mnuImportCSVAffidavit_Click()
    frmImportCSVAffidavit.Show vbModal
End Sub

Private Sub mnuImportIpump_Click()
    frmImportiPump.Show vbModal
End Sub

Private Sub mnuExportISCI_Click()
    igExportSource = 1
    igExportTypeNumber = 8
    frmExportISCI.Show vbModal
End Sub


Private Sub mnuExportISCIXref_Click()
    igExportSource = 1
    igExportTypeNumber = 9
    frmExportISCIXRef.Show vbModal
End Sub

Private Sub mnuExportLabelInfo_Click()
    frmExportLabelInfo.Show vbModal
End Sub

Private Sub mnuExportMarketron_Click()
    igExportSource = 1
    igExportTypeNumber = 1
    FrmExportMarketron.Show vbModal
End Sub

'Private Sub mnuExportOLA_Click()
'    FrmExportOLA.Show vbModal
'End Sub

Private Sub mnuExportRadar_Click()
    frmRadarExport.Show vbModal
End Sub

Private Sub mnuExportRCS4_Click() '
    igRCSExportBy = 4
    igExportSource = 1
    igExportTypeNumber = 4
    frmExportRCS.Show vbModal
End Sub

Private Sub mnuExportRCS5_Click()
    igRCSExportBy = 5
    igExportSource = 1
    igExportTypeNumber = 5
    frmExportRCS.Show vbModal
End Sub

Private Sub mnuExportSchdStationSpots_Click()
    igExportSource = 1
    igExportTypeNumber = 2
    frmExportSchdSpot.Show vbModal
End Sub

Private Sub mnuExportSpec_Click()
    frmExportSpec.Show vbModal
End Sub

Private Sub mnuExportStarGuide_Click()
    igExportSource = 1
    igExportTypeNumber = 10
    frmExportStarGuide.Show vbModal
End Sub

Private Sub mnuExportWegener_Click()
    igExportSource = 1
    igExportTypeNumber = 11
    FrmExportWegener.Show vbModal
End Sub

Private Sub mnuExportXDigital_Click()
    igExportSource = 1
    igExportTypeNumber = 12
    FrmExportXDigital.Show vbModal
End Sub

Private Sub mnuFileRQS_Click()
    frmReportQueueStatus.Show vbModeless
End Sub

Private Sub mnuGroupNameFormat_Click()
    sgFormatCall = "M"
    sgFormatName = ""
    frmGroupNameFormat.Show vbModal
End Sub

Private Sub mnuGroupNameDMAMarket_Click()
    sgGNMarketCall = "M"
    frmGroupNameMarket.Show vbModal
End Sub

Private Sub mnuGroupNameMSAMarket_Click()
    sgGNMarketCall = "M"
    frmGroupNameMSAMarket.Show vbModal
End Sub

Private Sub mnuGroupNameState_Click()
    sgStateCall = "M"
    frmGroupNameState.Show vbModal
End Sub

Private Sub mnuGroupNameTimeZone_Click()
    sgTimeZoneCall = "M"
    frmGroupNameTimeZone.Show vbModal
End Sub

Private Sub mnuGroupNameVehicle_Click()
    sgVehicleCall = "M"
    frmGroupNameVehicle.Show vbModal
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub
'Dan removed 6/16/10
'Private Sub mnuHelpContents_Click()
'
'
'    Dim nRet As Integer
'
'
'    'if there is no helpfile for this project display a message to the user
'    'you can set the HelpFile for your application in the
'    'Project Properties dialog
'    If Len(App.HelpFile) = 0 Then
'        gMsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
'    Else
'        On Error Resume Next
'        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
'        If Err Then
'            gMsgBox Err.Description
'        End If
'    End If
'End Sub

'Dan removed 6/16/10
'Private Sub mnuHelpSearch_Click()
'
'
'    Dim nRet As Integer
'
'
'    'if there is no helpfile for this project display a message to the user
'    'you can set the HelpFile for your application in the
'    'Project Properties dialog
'    If Len(App.HelpFile) = 0 Then
'        gMsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
'    Else
'        On Error Resume Next
'        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
'        If Err Then
'            gMsgBox Err.Description
'        End If
'    End If
'End Sub
Private Sub mnuFileExit_Click()
    On Error Resume Next
    Unload frmLogin
    Unload frmMain
End Sub

Private Sub mnuHelpGoToWeb_Click()
    'Dan M 6/16/10 go to website
    'ShellExecute 0&, vbNullString, "www.counterpoint.net/clientsLogin.php?error=4", vbNullString, vbNullString, vbNormalFocus
End Sub
Private Sub mnuDocumentation_Click()
    '10482
    'WebConnect.Show vbModal
    ShellExecute 0&, vbNullString, "www.counterpoint.net/clients/DocumentationNew.php?CsiNo=4307&CsiCo=" & Trim(sgClientName), vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub mnuImportWegenerCompel_Click()
    frmImportWegener.Show vbModal
End Sub

Private Sub mnuImportXDigital_Click()
    frmImportXDS.Show vbModal
End Sub

Private Sub mnuUtilitiesAvgWksDelinq_Click()
    frmAvgWksDelinq.Show vbModal
End Sub

Private Sub mnuUtilitiesDuplBkoutFix_Click()
    frmCheckBkout.Show vbModal
End Sub

Private Sub mnuUtilitiesReImportWebSpots_Click()
    igPasswordOk = False
    sgPasswordAddition = "Disallow View"
    CSPWord.Show vbModal
    sgPasswordAddition = ""
    If igPasswordOk Then
        frmReImportWebSpots.Show vbModal
    End If
End Sub

Private Sub mnuUtilitiesCPTTAgree_Click()
    frmCPTTAgree.Show vbModal
End Sub


Private Sub mnuUtilitiesSetCompliant_Click()
    frmSetCompliants.Show vbModal
End Sub

Private Sub mnuUtilitiesSetMG_Click()
    frmSetMG.Show vbModal
End Sub

Private Sub mnuUtilitiesSetPrgTimes_Click()
    sgSetFieldCallSource = "M"
    frmSetPrgTimes.Show vbModal
End Sub

Private Sub mnuUtilitiesSpotCount_Click()
    igPasswordOk = False
    sgPasswordAddition = "Disallow View"
    CSPWord.Show vbModal
    sgPasswordAddition = ""
    If igPasswordOk Then
        frmSpotCountSpec.Show vbModal
    End If
End Sub

' JD Added 02-22-24
Private Sub mnuUtilitiesWegenerCheckUtility_Click()
    Dim slPath As String
    
    slPath = sgExeDirectory & "WegenerUtility.exe"
    If gFileExist(slPath) = FILEEXISTS Then
        Shell slPath & " " & sgDatabaseName & " , " & sgDatabaseName, vbNormalFocus
    Else
        gMsgBox "The Wegener Check Utility does not exist on your system.  Contact Counterpoint for installion instructions.", vbOKOnly, "WegenerUtility"
    End If
End Sub

Private Sub mnuViewSql_Click()
    '6240
    Dim slPath As String
    
    slPath = sgExeDirectory & "ViewSql.exe"
    '8886
    'If Dir(slPath) <> "" Then
    If gFileExist(slPath) = FILEEXISTS Then
        Shell slPath & " " & sgDatabaseName & " , " & sgDatabaseName, vbNormalFocus
    Else
        gMsgBox "ViewSql does not exist on your system.  Contact Counterpoint for installion instructions.", vbOKOnly, "ViewSql"
    End If
End Sub

Private Sub mnuWebHome_Click()
    WebConnect.bmHomePage = True
    WebConnect.Show vbModal
End Sub
Private Sub mnuImportAffAE_Click()
    frmImportAE.Show vbModal
End Sub
Private Sub mnuExportStationInformation_Click()
    frmExportStationInformation.Show vbModal
End Sub
Private Sub mnuImportAffiliateSpots_Click()
    igImportSelection = 3
    frmImportCSV.Show vbModal
End Sub

Private Sub mnuImportAgreePledge_Click()
    igImportSelection = 5
    frmImportCSV.Show vbModal
End Sub

Private Sub mnuImportAirDates_Click()
    igImportSelection = 4
    frmImportCSV.Show vbModal
End Sub

Private Sub mnuImportAiredStationSpots_Click()
    frmImportAiredSpot.Show vbModal
End Sub


Private Sub mnuImportCPs_Click()
    igImportSelection = 6
    frmImportCSV.Show vbModal
End Sub

Private Sub mnuImportCSISpots_Click()
    frmImportCSISpot.Show vbModal
End Sub

Private Sub mnuImportCSV_Click()
    igImportSelection = 0
    frmImportCSV.Show vbModal
End Sub

Private Sub mnuImportCSV2_Click()
    igImportSelection = 8
    frmImportCSV.Show vbModal
End Sub

Private Sub mnuImportLogSpots_Click()
    igImportSelection = 2
    frmImportCSV.Show vbModal
End Sub

Private Sub mnuImportMarketron_Click()
     frmImportMarketron.Show vbModal
End Sub

Private Sub mnuImportMYLSpots_Click()
    igImportSelection = 7
    frmImportCSV.Show vbModal
End Sub

Private Sub mnuImportOracle_Click()
    igImportSelection = 1
    frmImportCSV.Show vbModal
End Sub


Private Sub mnuImportStation_Click()
    frmImportUpdateStations.Show vbModal
End Sub

Private Sub mnuImportWWOSpots_Click()
    frmImportWWOSpot.Show vbModal
End Sub

Private Sub mnuManageFormatsNew_Click()
    sgFormatCall = "N"
    sgFormatName = ""
    frmGroupNameFormat.Show vbModal
End Sub

Private Sub mnuManageFormatsXRef_Click()
    mManageFormatsXRef
End Sub

Private Sub mnuMergeFormats_Click()
    Dim ilRet As Integer
    
    ilRet = MsgBox("Backup of database must be done before merge, has it been done", vbQuestion + vbYesNo, "Merge Formats")
    If ilRet = vbNo Then
        Exit Sub
    End If
    ilRet = MsgBox("Warning! Before running Merge, no Traffic users can be in Copy Region screen or Network Region screen, and no Affiliate users can be in Station screen.  Otherwise disaster will result. Is it safe to proceed?", vbQuestion + vbYesNo, "Merge Formats")
    If ilRet = vbNo Then
        Exit Sub
    End If
    igMergeType = 1
    frmMerge.Show vbModal
End Sub

Private Sub mnuMergerMarkets_Click()
    Dim ilRet As Integer
    
    ilRet = MsgBox("Backup of database must be done before merge, has it been done", vbQuestion + vbYesNo, "Merge Markets")
    If ilRet = vbNo Then
        Exit Sub
    End If
    ilRet = MsgBox("Warning! Before running Merge, no Traffic users can be in Copy Region screen or Network Region screen, and no Affiliate users can be in Station screen.  Otherwise disaster will result. Is it safe to proceed?", vbQuestion + vbYesNo, "Merge Markets")
    If ilRet = vbNo Then
        Exit Sub
    End If
    igMergeType = 0
    frmMerge.Show vbModal
End Sub

Private Sub mnuUtilitiesBar_Click()

End Sub

Private Sub mnuUtilitiesCPTTCheck_Click()
    frmCPTTCheck.Show vbModal
End Sub

Private Sub mnuUtilitiesDuplCPTTFix_Click()
    frmDuplCPTTFix.Show vbModal
End Sub

Private Sub mnuUtilitiesDuplSHTTFix_Click()
    frmDuplSHTTFix.Show vbModal
End Sub

Private Sub mnuUtilitiesResetCompliant_Click()
    sgSetFieldCallSource = "M"
    frmSetDDFFields.Show vbModal
End Sub

Private Sub mnuUtilitiesWebPosting_Click()
    igPasswordOk = False
    'CSPWord.Show vbModal
    frmSpotUtil.Show vbModal
End Sub

Private Sub mnuWebImportAiredStationSpot_Click()
    frmWebImportAiredSpot.Show vbModal
End Sub

Private Sub mnuPrtSetup_Click()
    'cdcSetup.Flags = cdlPDPrintSetup
    cdcSetup.ShowPrinter
End Sub

Public Sub mLoadPicture()
        
    Dim slCommand As String
    Dim ilPos As Integer
    Dim ilRet As Integer
    
    slCommand = Command$
    
    'Check For Remote User - If so don't display the bitmap file
    ilPos = InStr(1, slCommand, "/RemoteUser", 1)

    
    If (igTestSystem) Or (igShowVersionNo = 1) Or (igShowVersionNo = 2) Or (Trim$(sgWallpaper) <> "") Then
        frmMstPict.Show vbModeless
        frmMstPict.Enabled = False
        frmMain.BackColor = &HC0C0C0
    Else
        'If ilPos = 0 Or slCommand = "debug" Then
        If (ilPos = 0) Or (InStr(1, slCommand, "Debug", vbTextCompare) > 0) Then
            On Error GoTo mRetryPicture:
            If Screen.Height / Screen.TwipsPerPixelY <= 480 Then
                If gFileExist(sgLogoDirectory & "CSI640A.jpg") Then
                    frmMain.Picture = LoadPicture(sgLogoDirectory & "CSI640A.jpg")
                ElseIf gFileExist(sgLogoDirectory & "CSI640A.gif") Then
                    frmMain.Picture = LoadPicture(sgLogoDirectory & "CSI640A.gif")
                ElseIf gFileExist(sgLogoDirectory & "CSI640A.Bmp") Then
                        frmMain.Picture = LoadPicture(sgLogoDirectory & "CSI640A.Bmp")
                End If
            ElseIf Screen.Height / Screen.TwipsPerPixelY <= 600 Then
                If gFileExist(sgLogoDirectory & "CSI800A.jpg") Then
                    frmMain.Picture = LoadPicture(sgLogoDirectory & "CSI800A.jpg")
                ElseIf gFileExist(sgLogoDirectory & "CSI800A.gif") Then
                    frmMain.Picture = LoadPicture(sgLogoDirectory & "CSI800A.gif")
                ElseIf gFileExist(sgLogoDirectory & "CSI800A.Bmp") Then
                    frmMain.Picture = LoadPicture(sgLogoDirectory & "CSI800A.Bmp")
                End If
            
            ElseIf Screen.Height / Screen.TwipsPerPixelY <= 768 Then
                If gFileExist(sgLogoDirectory & "CSI1024A.jpg") Then
                    frmMain.Picture = LoadPicture(sgLogoDirectory & "CSI1024A.jpg")
                ElseIf gFileExist(sgLogoDirectory & "CSI1024A.gif") Then
                    frmMain.Picture = LoadPicture(sgLogoDirectory & "CSI1024A.gif")
                ElseIf gFileExist(sgLogoDirectory & "CSI1024A.Bmp") Then
                    frmMain.Picture = LoadPicture(sgLogoDirectory & "CSI1024A.Bmp")
                End If
            Else
                If gFileExist(sgLogoDirectory & "CSI1280A.jpg") Then
                    frmMain.Picture = LoadPicture(sgLogoDirectory & "CSI1280A.jpg")
                ElseIf gFileExist(sgLogoDirectory & "CSI1280A.gif") Then
                    frmMain.Picture = LoadPicture(sgLogoDirectory & "CSI1280A.gif")
                ElseIf gFileExist(sgLogoDirectory & "CSI1280A.Bmp") Then
                    frmMain.Picture = LoadPicture(sgLogoDirectory & "CSI1280A.Bmp")
                End If
            End If
        Else
            'gFadeForm frmMain, False, False, True
        End If
    End If
    Exit Sub
mRetryPicture:
    ilRet = 1
    Resume Next
mNoPicture:
    'gFadeForm frmMain, False, False, True
    On Error GoTo 0
    Resume Next

'        If Screen.Height / Screen.TwipsPerPixelY <= 480 Then
'            frmMain.Picture = LoadPicture(sgLogoDirectory & "CSI640A.Bmp")
'        ElseIf Screen.Height / Screen.TwipsPerPixelY <= 600 Then
'            frmMain.Picture = LoadPicture(sgLogoDirectory & "CSI800A.Bmp")
'        Else
'            frmMain.Picture = LoadPicture(sgLogoDirectory & "CSI1024A.Bmp")
'        End If
'    End If
'
'    Exit Sub
'
'mNoPicture:
'    'gFadeForm frmMain, False, False, True
'    On Error GoTo 0
'    Resume Next
End Sub

Private Sub mnuWebExportSchdStationSpots_Click()
    Dim slAllowWebExports As String
    
    'Test to see if web exports are allowed.
    slAllowWebExports = "Y"
    Call gLoadOption(sgWebServerSection, "AllowWebExports", slAllowWebExports)
    
    If UCase(slAllowWebExports) = "N" Or UCase(slAllowWebExports) = "NO" Then
        gMsgBox "Web Exports have been temporarily turned off. " & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "If you believe you've received this message in error, please contact your administrator."
    Else
        igExportSource = 1
        igExportTypeNumber = 3
        sgWebExport = "B"
        frmWebExportSchdSpot.Show vbModal
    End If
End Sub

Private Sub mnuWebVendors_Click()
    igPasswordOk = False
    CSPWord.Show vbModal
    frmWebVendors.Show vbModal
End Sub

Private Sub tmcCheckAlert_Timer()
    Dim ilRet As Integer
    
    ilRet = gAlertCheck()
End Sub

Private Sub tmcClock_Timer()
    Dim ilRet As Integer
    Dim slMessage As String
    Dim llAufCode As Long
    Dim llCefCode As Long
    Dim slDateTime As String
    
    mCheckGG
    
    ilRet = gAlertCheckBlock(tmAuf)
    If ilRet = 1 Then   'Block user
        imPrevAlertBlockChk = ilRet
        frmMstPict.Cls
        If (igShowVersionNo <> -1) Or (tmAuf.lEnteredTime <> lmCountdownEnteredTime) Then
            If tmAuf.lCefCode > 0 Then
                hmCef = CBtrvTable(ONEHANDLE)
                ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                If ilRet = BTRV_ERR_NONE Then
                    tmCefSrchKey.lCode = tmAuf.lCefCode
                    tmCef.sComment = ""
                    imCefRecLen = Len(tmCef)    '1009
                    ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        'If tmCef.iStrLen > 0 Then
                        '    slMessage = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                        'End If
                        slMessage = gStripChr0(tmCef.sComment)
                    End If
                Else
                    slMessage = "System will be shut down in " & Trim$(Str$(tmAuf.iCountdown)) & " minutes"
                End If
                btrDestroy hmCef
            Else
                slMessage = "System will be shut down in " & Trim$(Str$(tmAuf.iCountdown)) & " minutes"
            End If
            igShowVersionNo = -1
            sgWallpaper = "Shutdown: " & tmAuf.iCountdown & "m"
            gWriteBkgd
            imCountdown = tmAuf.iCountdown
            lmCountdownEnteredTime = tmAuf.lEnteredTime
            MsgBox slMessage, vbInformation, "Notification"
        Else
            imCountdown = imCountdown - 1
            If imCountdown >= 0 Then
                igShowVersionNo = -1
                sgWallpaper = "Shutdown: " & imCountdown & "m"
                gWriteBkgd
            Else
                igShowVersionNo = -1
                sgWallpaper = "Shutdown"
                gWriteBkgd
            End If
        End If
    ElseIf ilRet = 2 Then   'Block Initiated by this user
        imPrevAlertBlockChk = ilRet
        frmMstPict.Cls
        If (igShowVersionNo <> -1) Or (tmAuf.lEnteredTime <> lmCountdownEnteredTime) Then
            igShowVersionNo = -1
            sgWallpaper = "Shutdown: " & tmAuf.iCountdown & "m"
            gWriteBkgd
            imCountdown = tmAuf.iCountdown
            lmCountdownEnteredTime = tmAuf.lEnteredTime
        Else
            imCountdown = imCountdown - 1
            If imCountdown >= 0 Then
                igShowVersionNo = -1
                sgWallpaper = "Shutdown: " & imCountdown & "m"
                gWriteBkgd
            Else
                igShowVersionNo = -1
                sgWallpaper = "Shutdown"
                gWriteBkgd
            End If
        End If
    Else
        If igShowVersionNo = -1 Then
            frmMstPict.Cls
            igShowVersionNo = imShowVersionNo
            sgWallpaper = smWallpaper
            gWriteBkgd
            If imPrevAlertBlockChk <> 2 Then
                MsgBox "System Shut Down has been stopped!", vbInformation, "Notification"
            End If
            imPrevAlertBlockChk = ilRet
        End If
    End If
    
    Do
        llAufCode = gAlertCheckNotice(lgUlfCode, llCefCode)
        If llAufCode > 0 Then
            If llCefCode > 0 Then
                hmCef = CBtrvTable(ONEHANDLE)
                ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                If ilRet = BTRV_ERR_NONE Then
                    tmCefSrchKey.lCode = llCefCode
                    tmCef.sComment = ""
                    imCefRecLen = Len(tmCef)    '1009
                    ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        slMessage = gStripChr0(tmCef.sComment)
                        'If tmCef.iStrLen > 0 Then
                        If slMessage <> "" Then
                            'slMessage = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                            MsgBox slMessage, vbInformation, "Notification"
                        End If
                    End If
                End If
                btrDestroy hmCef
            End If
            ilRet = gAlertClear("A", "N", "", 0, Trim$(Str$(llAufCode)))
        Else
            Exit Do
        End If
    Loop While llAufCode > 0
    
    slDateTime = " " & Format$(gNow(), "ddd, m/d/yy h:mm AM/PM")
    mnuDate.Caption = slDateTime & " " & sgDateBrannerMsg
    
    mCheckActiveLog
    
End Sub

Private Sub tmcFlashAlert_Timer()
    If igGGFlag = 0 Then
        mnuAlert.Visible = False
        tmcFlashAlert.Enabled = False
        Exit Sub
    End If

    igAlertFlash = igAlertFlash + 1
    If igAlertFlash = 10 Then
        If ((igAlertInterval <> 0) And (igAlertInterval <= igAlertTimer)) Then
            tmcFlashAlert.Enabled = False
            mnuAlert.Visible = False
            DoEvents
            mnuAlert.Visible = True
            mShowAlert hmRedAlertBitmap
            tmcCheckAlert.Enabled = True
        Else
            mnuAlert.Visible = False
            DoEvents
            mnuAlert.Visible = True
            mShowAlert hmRedAlertBitmap
            tmcFlashAlert.Interval = 60000  'Every minute
            igAlertFlash = 0
            igAlertTimer = igAlertTimer + 1
        End If
        Exit Sub
    Else
        tmcFlashAlert.Interval = 2000  'Every 2 seconds
        If igAlertFlash Mod 2 = 0 Then
            mnuAlert.Visible = False
            DoEvents
            mnuAlert.Visible = True
            mShowAlert hmRedAlertBitmap
            
        Else
            mnuAlert.Visible = False
            DoEvents
            mnuAlert.Visible = True
            mShowAlert hmWhiteAlertBitmap
        End If
        DoEvents
    End If
End Sub

Private Sub mUnloadForms()
    'On Error Resume Next
    'Unload frmAgmnt
    'Unload frmContact
    'Unload frmCP
    'Unload frmCPReturns
    'Unload frmOptions
    'Unload frmPostLog
    'Unload frmSiteOptions
    'Unload frmStation
    'Unload frmWebEMail
    'Unload frmStationSearch
    'On Error GoTo 0
    Dim ilLoop As Integer
    For ilLoop = Forms.Count - 1 To 0 Step -1
        Unload Forms(ilLoop)
    Next ilLoop
End Sub

Private Sub tmcMonitor_Timer()
    Dim llDateMonitorChecked As Long
    Dim llTimeMonitorChecked As Long
    Dim ilTask As Integer
    Dim slDate As String
    Dim slTime As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim llRunningDate As Long
    Dim llRunningTime As Long
    Dim ll1stStartDate As Long
    Dim ll1stStartTime As Long
    Dim llStartDate As Long
    Dim llStartTime As Long
    Dim llTestStartDate As Long
    Dim llTestStartTime As Long
    Dim ll1stEndDate As Long
    Dim ll1stEndTime As Long
    Dim llEndDate As Long
    Dim llEndTime As Long
    Dim llTestEndDate As Long
    Dim llTestEndTime As Long
    Dim blFound As Boolean
    Dim llNextDate As Long
    Dim ilDay As Integer
    Dim llCurDate As Long
    Dim llCurTime As Long
    Dim slCurDate As String
    Dim slCurTime As String
    Dim llInvDate As Long
    Dim llEMailDate As Long
    Dim llEMailTime As Long
    Dim rst_Spf As ADODB.Recordset
    '8793
    Dim slDebug As String
    
    llDateMonitorChecked = gDateValue(Format(Now, "m/d/yy"))
    llTimeMonitorChecked = gTimeToLong(Format(Now, "h:mm:ssAM/PM"), True)
    If (lgDateMonitorChecked <> llDateMonitorChecked) Or ((llTimeMonitorChecked - lgTimeMonitorChecked >= MONITORRUNINTERVAL)) Then
        For ilTask = 0 To UBound(tgTaskInfo) Step 1
            '8793
            slDebug = ""
            If tgTaskInfo(ilTask).iMenuIndex > 0 Then
                tmTmfSrchKey1.sTaskCode = Trim$(tgTaskInfo(ilTask).sTaskCode)
                ilRet = btrGetEqual(hgTmf, tmTmf, imTmfRecLen, tmTmfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    'gUnpackDate tmTmf.iRunningDate(0), tmTmf.iRunningDate(1), slDate
                    'gUnpackTime tmTmf.iRunningTime(0), tmTmf.iRunningTime(1), "M", "1", slTime
                    'Select Case smZone
                    '    Case "E"
                    '        slStr = slDate & " " & slTime
                    '    Case "C"
                    '        slStr = DateAdd("h", 1, slDate & " " & slTime)
                    '    Case "M"
                    '        slStr = DateAdd("h", 2, slDate & " " & slTime)
                    '    Case "P"
                    '        slStr = DateAdd("h", 3, slDate & " " & slTime)
                    'End Select
                    'slDate = Format$(slStr, "m/d/yy")
                    'slTime = Format$(slStr, "h:mm:ssAM/PM")
                    'If (gDateValue(slDate) <> lgMonitorRunningDate) Or (gTimeToLong(slTime, False) <> lgMonitorRunningTime) Then
                    '7967
                    If tmTmf.sService = "W" Then
                        '8129
                        tmTmf.sStatus = gMonitorVendor(tmcWebConnectIssue)
'                        If Not gMonitorVendor(tmcWebConnectIssue) Then
'                            tmTmf.sStatus = "E"
'                        End If
                    End If
                    gUnpackDateLong tmTmf.iRunningDate(0), tmTmf.iRunningDate(1), llRunningDate
                    gUnpackTimeLong tmTmf.iRunningTime(0), tmTmf.iRunningTime(1), True, llRunningTime
                    gUnpackDateLong tmTmf.iStartRunDate(0), tmTmf.iStartRunDate(1), llStartDate
                    '8129 this is Dan forces to yellow
                    If tmTmf.sService = "W" And tmTmf.sStatus = "Y" Then
                        llRunningDate = lmDate1970
                    End If
                    If (llRunningDate = lmDate1970) And (llStartDate = lmDate1970) Then
                        If tgTaskInfo(ilTask).lColor <> LIGHTYELLOW Then
                            tgTaskInfo(ilTask).lColor = LIGHTYELLOW
                            mSetMonitorInfo ilTask
                        End If
                    Else
                        If (llRunningDate <> tgTaskInfo(ilTask).lRunningDate) Or (llRunningTime <> tgTaskInfo(ilTask).lRunningTime) Then
                            If tgTaskInfo(ilTask).lColor <> DARKGREEN Then
                                tgTaskInfo(ilTask).lColor = DARKGREEN
                                mSetMonitorInfo ilTask
                            End If
                            tgTaskInfo(ilTask).lRunningDate = llRunningDate
                            tgTaskInfo(ilTask).lRunningTime = llRunningTime
                            tgTaskInfo(ilTask).lElapsedTime = 0
                        Else
                            If (tmTmf.sStatus = "E") Or (tmTmf.sStatus = "A") Then
                                tgTaskInfo(ilTask).lElapsedTime = tgTaskInfo(ilTask).lElapsedTime + MONITORRUNINTERVAL
                                If tgTaskInfo(ilTask).lElapsedTime >= MONITORELAPSEDINTERVAL Then
                                    slDebug = "DEBUG 1- elapsed time = " & tgTaskInfo(ilTask).lElapsedTime
                                    If tgTaskInfo(ilTask).lColor <> vbRed Then
                                        tgTaskInfo(ilTask).lColor = vbRed
                                        mSetMonitorInfo ilTask
                                    End If
                                    tgTaskInfo(ilTask).lElapsedTime = 0
                                '7967
                                ElseIf tmTmf.sService = "W" Then
                                    If tgTaskInfo(ilTask).lColor <> vbRed Then
                                        tgTaskInfo(ilTask).lColor = vbRed
                                        mSetMonitorInfo ilTask
                                    End If
                                    tgTaskInfo(ilTask).lElapsedTime = 0
                                End If
                            ElseIf tmTmf.sRunMode = "C" Then
                                tgTaskInfo(ilTask).lElapsedTime = tgTaskInfo(ilTask).lElapsedTime + MONITORRUNINTERVAL
                                If tgTaskInfo(ilTask).lElapsedTime >= MONITORELAPSEDINTERVAL Then
                                    slDebug = "DEBUG 2- elapsed time = " & tgTaskInfo(ilTask).lElapsedTime
                                    If tgTaskInfo(ilTask).lColor <> vbRed Then
                                        '7967 exclude web service
                                        If tmTmf.sService <> "W" Then
                                            tgTaskInfo(ilTask).lColor = vbRed
                                            mSetMonitorInfo ilTask
                                        End If
                                    End If
                                    tgTaskInfo(ilTask).lElapsedTime = 0
                                End If
                            Else
                                'Periodic
                                slCurDate = Format(Now, "m/d/yy")
                                slCurTime = Format(Now, "h:mm:ssAM/PM")
                                Select Case sgTimeZone
                                    Case "E"
                                        slStr = slCurDate & " " & slCurTime
                                    Case "C"
                                        slStr = DateAdd("h", 1, slCurDate & " " & slCurTime)
                                    Case "M"
                                        slStr = DateAdd("h", 2, slCurDate & " " & slCurTime)
                                    Case "P"
                                        slStr = DateAdd("h", 3, slCurDate & " " & slCurTime)
                                End Select
                                llCurDate = gDateValue(Format$(slStr, "m/d/yy"))
                                llCurTime = gTimeToLong(Format$(slStr, "h:mm:ssAM/PM"), False)
                                gUnpackDateLong tmTmf.i1stStartRunDate(0), tmTmf.i1stStartRunDate(1), ll1stStartDate
                                gUnpackTimeLong tmTmf.i1stStartRunTime(0), tmTmf.i1stStartRunTime(1), False, ll1stStartTime
                                gUnpackDateLong tmTmf.i1stEndRunDate(0), tmTmf.i1stEndRunDate(1), ll1stEndDate
                                gUnpackTimeLong tmTmf.i1stEndRunTime(0), tmTmf.i1stEndRunTime(1), True, ll1stEndTime
                                gUnpackDateLong tmTmf.iStartRunDate(0), tmTmf.iStartRunDate(1), llStartDate
                                gUnpackTimeLong tmTmf.iStartRunTime(0), tmTmf.iStartRunTime(1), False, llStartTime
                                gUnpackDateLong tmTmf.iEndRunDate(0), tmTmf.iEndRunDate(1), llEndDate
                                gUnpackTimeLong tmTmf.iEndRunTime(0), tmTmf.iEndRunTime(1), True, llEndTime
                                'If ((llStartDate = llEndDate) And (llEndTime >= llStartTime)) Or ((llStartDate + 1 = llEndDate) And (llEndTime < llStartTime)) Then
                                If (ll1stStartDate = llStartDate) And (ll1stStartTime >= llStartTime) Then
                                    'has started on today but not completed
                                    llTestStartDate = ll1stStartDate
                                    llTestEndDate = ll1stEndDate
                                    llTestStartTime = ll1stStartTime
                                    llTestEndTime = ll1stEndTime
                                Else
                                    'last run was yesterday
                                    llTestStartDate = llStartDate
                                    llTestEndDate = llEndDate
                                    llTestStartTime = llStartTime
                                    llTestEndTime = llEndTime
                                End If
                                If ((llTestStartDate = llTestEndDate) And (llTestEndTime >= llTestStartTime)) Or ((llTestStartDate + 1 = llTestEndDate) And (llTestEndTime < llTestStartTime)) Then
                                    'Determine next date that it should run
                                    If Trim(tmTmf.sMonthPeriod) = "" Then
                                        blFound = False
                                        If llTestEndTime >= llTestStartTime Then
                                            llNextDate = llTestEndDate + 1
                                        Else
                                            llNextDate = llTestEndDate
                                        End If
                                        Do
                                            ilDay = gWeekDayLong(llNextDate)
                                            Select Case ilDay
                                                Case 0
                                                    If tmTmf.sMo = "Y" Then
                                                        blFound = True
                                                    End If
                                                Case 1
                                                    If tmTmf.sTu = "Y" Then
                                                        blFound = True
                                                    End If
                                                Case 2
                                                    If tmTmf.sWe = "Y" Then
                                                        blFound = True
                                                    End If
                                                Case 3
                                                    If tmTmf.sTh = "Y" Then
                                                        blFound = True
                                                    End If
                                                Case 4
                                                    If tmTmf.sFr = "Y" Then
                                                        blFound = True
                                                    End If
                                                Case 5
                                                    If tmTmf.sSa = "Y" Then
                                                        blFound = True
                                                    End If
                                                Case 6
                                                    If tmTmf.sSu = "Y" Then
                                                        blFound = True
                                                    End If
                                            End Select
                                            If blFound Then
                                                Exit Do
                                            End If
                                            llNextDate = llNextDate + 1
                                        Loop While Not blFound
                                    Else
                                        Select Case tmTmf.sMonthPeriod
                                            Case "SE"
                                                slDate = gObtainEndStd(Format(llTestStartDate + 1, "m/d/yy"))
                                            Case "CE"
                                                slDate = gObtainEndCal(Format(llTestStartDate + 1, "m/d/yy"))
                                            Case "IE"
                                                SQLQuery = "Select spfBLastStdMnth From SPF_Site_Options"
                                                Set rst_Spf = gSQLSelectCall(SQLQuery, "frmMain: tmcMonitor")
                                                slDate = gObtainEndStd(Format(gDateValue(Format(rst_Spf!spfBLastStdMnth, "m/d/yy")) + 1, "m/d/yy"))
                                        End Select
                                        llNextDate = gDateValue(DateAdd("d", tmTmf.iDaysAfter, slDate))
                                    End If
                                    If (llCurDate < llNextDate) Or (llCurDate = llNextDate) And (llCurTime < llTestStartTime) Then
                                        If tgTaskInfo(ilTask).lColor <> DARKGREEN Then
                                            tgTaskInfo(ilTask).lColor = DARKGREEN
                                            mSetMonitorInfo ilTask
                                        End If
                                    Else
                                        slDebug = "DEBUG 3- current date = " & Format(llCurDate, sgSQLDateForm) & " next date = " & Format(llNextDate, sgSQLDateForm) & " current time = " & gFormatTimeLong(llCurTime, "A", "1") & " test start time = " & gFormatTimeLong(llTestStartTime, "A", "1")
                                        If tgTaskInfo(ilTask).lColor <> vbRed Then
                                            tgTaskInfo(ilTask).lColor = vbRed
                                            mSetMonitorInfo ilTask
                                        End If
                                    End If
                                ElseIf (llTestStartDate > llTestEndDate) Or ((llTestStartDate = llTestEndDate) And (llTestStartTime > llTestEndTime)) Then
                                    tgTaskInfo(ilTask).lElapsedTime = tgTaskInfo(ilTask).lElapsedTime + MONITORRUNINTERVAL
                                    If tgTaskInfo(ilTask).lElapsedTime >= MONITORELAPSEDINTERVALPERIOD Then
                                        slDebug = "DEBUG 4- test start date = " & Format(llTestStartDate, sgSQLDateForm) & " test end date = " & Format(llTestEndDate, sgSQLDateForm) & " test start time = " & gFormatTimeLong(llTestStartTime, "A", "1") & " test end time = " & gFormatTimeLong(llTestEndTime, "A", "1")
                                        If tgTaskInfo(ilTask).lColor <> vbRed Then
                                            tgTaskInfo(ilTask).lColor = vbRed
                                            mSetMonitorInfo ilTask
                                        End If
                                        tgTaskInfo(ilTask).lElapsedTime = 0
                                    ElseIf tgTaskInfo(ilTask).lColor <> vbRed Then
                                        If tgTaskInfo(ilTask).lColor <> LIGHTYELLOW Then
                                            tgTaskInfo(ilTask).lColor = LIGHTYELLOW
                                            mSetMonitorInfo ilTask
                                        End If
                                    End If
                                End If
                            End If
                            If tgTaskInfo(ilTask).lColor = vbRed Then
                                slCurDate = Format(Now, "m/d/yy")
                                slCurTime = Format(Now, "h:mm:ssAM/PM")
                                Select Case sgTimeZone
                                    Case "E"
                                        slStr = slCurDate & " " & slCurTime
                                    Case "C"
                                        slStr = DateAdd("h", 1, slCurDate & " " & slCurTime)
                                    Case "M"
                                        slStr = DateAdd("h", 2, slCurDate & " " & slCurTime)
                                    Case "P"
                                        slStr = DateAdd("h", 3, slCurDate & " " & slCurTime)
                                End Select
                                llCurDate = gDateValue(Format$(slStr, "m/d/yy"))
                                llCurTime = gTimeToLong(Format$(slStr, "h:mm:ssAM/PM"), False)
                                gUnpackDateLong tmTmf.iEMailSentDate(0), tmTmf.iEMailSentDate(1), llEMailDate
                                If (llEMailDate <> llCurDate) Then
                                    gUnpackDateLong tmTmf.iEMailReqDate(0), tmTmf.iEMailReqDate(1), llEMailDate
                                    If (llEMailDate <> llCurDate) Then
                                        gPackDate Format(llCurDate, "m/d/yy"), tmTmf.iEMailReqDate(0), tmTmf.iEMailReqDate(1)
                                        gPackTimeLong llCurTime, tmTmf.iEMailReqTime(0), tmTmf.iEMailReqTime(1)
                                        ilRet = btrUpdate(hgTmf, tmTmf, imTmfRecLen)
                                    Else
                                        gUnpackTimeLong tmTmf.iEMailReqTime(0), tmTmf.iEMailReqTime(1), False, llEMailTime
                                        If llCurTime - llEMailTime > MONITOREMAILINTERVAL And emMonitoringAllowed <> MonitorNoEmail Then
                                            gPackDate Format(llCurDate, "m/d/yy"), tmTmf.iEMailSentDate(0), tmTmf.iEMailSentDate(1)
                                            ilRet = btrUpdate(hgTmf, tmTmf, imTmfRecLen)
                                            '8793
                                            If Len(slDebug) = 0 Then
                                                slDebug = "Hit no debug points!"
                                            End If
                                            slDebug = smMonitorID & " User: " & sgUserName & " Time Zone: " & sgTimeZone & " Debug Info: Running Date = " & Format$(llRunningDate, sgSQLDateForm) & " Running Time = " & gFormatTimeLong(llRunningTime, "A", "1") & " " & slDebug
                                            'Dan generate email--don't send if signed in as 'csi'
                                            If (StrComp(sgUserName, "Guide", 1) = 0) And Not bgLimitedGuide Then
                                                'testing only!
                                               ' mTempSendServiceEmail Trim$(tgTaskInfo(ilTask).sTaskName) & " program failure", "The Monitor Program in Affiliate has detected a problem for " & sgClientName & " at " & Now & slDebug
                                            Else
                                                'gSendServiceEmail Trim$(tgTaskInfo(ilTask).sTaskName) & " program failure", "The Monitor Program in Affiliate has detected a problem for " & sgClientName & " at " & gNow()
                                               '8763
                                                'gSendServiceEmail Trim$(tgTaskInfo(ilTask).sTaskName) & " program failure", "The Monitor Program in Affiliate has detected a problem for " & sgClientName & " at " & Now
                                                gSendServiceEmail Trim$(tgTaskInfo(ilTask).sTaskName) & " program failure", "The Monitor Program in Affiliate has detected a problem for " & sgClientName & " at " & Now & slDebug
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next ilTask
        lgDateMonitorChecked = llDateMonitorChecked
        lgTimeMonitorChecked = llTimeMonitorChecked
    End If
End Sub


Private Sub tmcStartUp_Timer()
    tmcStartUp.Enabled = False
    mSetJobButtons
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    mnuFileExit_Click
End Sub
Private Function mSignOnOff(ilSignOn As Integer) As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim slPCName As String
    Dim slMACAddr As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim slDBType As String
    Dim slDate As String
    
    hmUlf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmUlf, "", sgDBPath & "Ulf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mSignOnOff = False
        Exit Function
    End If
    imUlfRecLen = Len(tmUlf)  'btrRecordLength(hlUrf)  'Get and save record length
    slPCName = Trim$(gGetComputerName())
    slMACAddr = Trim$(gGetMACs_AdaptInfo())
    slNowDate = Format$(Now, "m/d/yy")
    slNowTime = Format$(Now, "h:mm:ssAM/PM")
    If igTestSystem Then
        tmUlfSrchKey1.sDBType = "T"
    Else
        tmUlfSrchKey1.sDBType = "P"
    End If
    slDBType = tmUlfSrchKey1.sDBType
    ilFound = False
    tmUlfSrchKey1.sSystemType = "A"
    tmUlfSrchKey1.iurfCode = 0
    tmUlfSrchKey1.iUstCode = igUstCode
    tmUlfSrchKey1.iUieCode = 0
    ilRet = btrGetEqual(hmUlf, tmUlf, imUlfRecLen, tmUlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, 0)
    Do While ilRet = BTRV_ERR_NONE
        If (tmUlf.sDBType = slDBType) And (tmUlf.sSystemType = "A") And (tmUlf.iUstCode = igUstCode) And (StrComp(Trim$(tmUlf.sPCName), slPCName, vbTextCompare) = 0) And (StrComp(Trim$(tmUlf.sPCMACAddr), slMACAddr, vbTextCompare) = 0) Then
            ilFound = True
            If ilSignOn Then
                'Update Sign On time
                gUnpackDate tmUlf.iSignOffDate(0), tmUlf.iSignOffDate(1), slDate
                If gDateValue(slDate) = gDateValue("12/31/2069") Then
                    'User Aborted, set UAF
                    mSetUAFToAborted tmUlf.lCode
                End If
                'Update Sign On time
                gPackDate slNowDate, tmUlf.iSignOnDate(0), tmUlf.iSignOnDate(1)
                gPackTime slNowTime, tmUlf.iSignOnTime(0), tmUlf.iSignOnTime(1)
                gPackDate "12/31/2069", tmUlf.iSignOffDate(0), tmUlf.iSignOffDate(1)
                gPackTime "12am", tmUlf.iSignOffTime(0), tmUlf.iSignOffTime(1)
                gPackDate slNowDate, tmUlf.iActiveLogDate(0), tmUlf.iActiveLogDate(1)
                gPackTime slNowTime, tmUlf.iActiveLogTime(0), tmUlf.iActiveLogTime(1)
            Else
                gPackDate slNowDate, tmUlf.iSignOffDate(0), tmUlf.iSignOffDate(1)
                gPackTime slNowTime, tmUlf.iSignOffTime(0), tmUlf.iSignOffTime(1)
                gPackDate slNowDate, tmUlf.iActiveLogDate(0), tmUlf.iActiveLogDate(1)
                gPackTime slNowTime, tmUlf.iActiveLogTime(0), tmUlf.iActiveLogTime(1)
            End If
            ilRet = btrUpdate(hmUlf, tmUlf, imUlfRecLen)
            Exit Do
        End If
        ilRet = btrGetNext(hmUlf, tmUlf, imUlfRecLen, BTRV_LOCK_NONE, 0)
    Loop
    If (Not ilFound) And (ilSignOn) Then
        'Create record with Sign On time
        tmUlf.lCode = 0
        If igTestSystem Then
            tmUlf.sDBType = "T"
        Else
            tmUlf.sDBType = "P"
        End If
        tmUlf.sSystemType = "A"
        tmUlf.iurfCode = 0
        tmUlf.iUstCode = igUstCode
        tmUlf.iUieCode = 0
        gPackDate slNowDate, tmUlf.iSignOnDate(0), tmUlf.iSignOnDate(1)
        gPackTime slNowTime, tmUlf.iSignOnTime(0), tmUlf.iSignOnTime(1)
        gPackDate "12/31/2069", tmUlf.iSignOffDate(0), tmUlf.iSignOffDate(1)
        gPackTime "12am", tmUlf.iSignOffTime(0), tmUlf.iSignOffTime(1)
        gPackDate slNowDate, tmUlf.iActiveLogDate(0), tmUlf.iActiveLogDate(1)
        gPackTime slNowTime, tmUlf.iActiveLogTime(0), tmUlf.iActiveLogTime(1)
        tmUlf.sPCName = slPCName
        tmUlf.sPCMACAddr = slMACAddr
        tmUlf.sTimeZone = Left$(gGetLocalTZName(), 1)
        tmUlf.iTrafJobNo = -1
        tmUlf.iTrafListNo = 0
        tmUlf.iTrafRnfCode = 0
        tmUlf.iAffTaskNo = -1
        tmUlf.iAffSubtaskNo = 0
        tmUlf.iAffRptNo = 0
        ilRet = btrInsert(hmUlf, tmUlf, imUlfRecLen, INDEXKEY0)
    End If
    If ilFound Then
        lgActiveUlfCode = tmUlf.lCode
        sgActiveLogDate = slNowDate
        sgActiveLogTime = slNowTime
    End If
    btrDestroy hmUlf
    lgUlfCode = tmUlf.lCode
    igNoDaysRetainUAF = 5
    If ilSignOn Then
        '4/2/11: Add setting and call.  Note: The call in _Load will be ignored
        SQLQuery = "Select safNoDaysRetainUAF From SAF_Schd_Attributes WHERE safVefCode = 0"
        Set rst = gSQLSelectCall(SQLQuery, "rmMain: mSignOnOff")
        If Not rst.EOF Then
            igNoDaysRetainUAF = rst!safNoDaysRetainUAF
        End If
        igLogActivityStatus = 32123 'Start Log Activity
        gUserActivityLog "L", "AffMain.Frm"
        'If (Not igDemoMode) And (Len(sgSpecialPassword) <> 4) Then
        '    mnuExportSpec.Visible = False
        'Else
        '    mnuExportSpec.Visible = True
        'End If
    Else
        '4/2/11: Add setting and call.
        If igLogActivityStatus = 32123 Then
            igLogActivityStatus = -32123    'End Log Activity
            gUserActivityLog "", ""
        End If
    End If
    mSignOnOff = True
End Function

Private Sub mCheckActiveLog()
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim ilRet As Integer
    Dim llMinuteDiff As Long
    
    If sgActiveLogDate = "" Then
        Exit Sub
    End If
    If Not gIsDate(sgActiveLogDate) Then
        Exit Sub
    End If
    slNowDate = Format$(Now, "m/d/yy")
    slNowTime = Format$(Now, "h:mm:ssAM/PM")
    llMinuteDiff = DateDiff("n", sgActiveLogDate & " " & sgActiveLogTime, slNowDate & " " & slNowTime)
    If llMinuteDiff > 60 Then
        hmUlf = CBtrvTable(TWOHANDLES)
        ilRet = btrOpen(hmUlf, "", sgDBPath & "Ulf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet = BTRV_ERR_NONE Then
            imUlfRecLen = Len(tmUlf)  'btrRecordLength(hlUrf)  'Get and save record length
            tmUlfSrchKey.lCode = lgActiveUlfCode
            ilRet = btrGetEqual(hmUlf, tmUlf, imUlfRecLen, tmUlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                gPackDate slNowDate, tmUlf.iActiveLogDate(0), tmUlf.iActiveLogDate(1)
                gPackTime slNowTime, tmUlf.iActiveLogTime(0), tmUlf.iActiveLogTime(1)
                ilRet = btrUpdate(hmUlf, tmUlf, imUlfRecLen)
            End If
        End If
        sgActiveLogDate = slNowDate
        sgActiveLogTime = slNowTime
        btrDestroy hmUlf
    End If
End Sub

Private Sub mnuAccessoriesDate_Click()
    Dim slDate As String
    Dim slTime As String
    Dim slMsg As String
    Dim slInputDate As String
    Dim ilRet As Integer
    Dim llNowDate As Long
    Dim slDateTime As String
    
    slDate = Format$(gNow(), "m/d/yy")
    slTime = Format$(gNow(), "h:mm:ss AM/PM")
    slMsg = "Counterpoint Date: " & slDate & "; Counterpoint Time: " & slTime
    If (StrComp(sgUserName, "Counterpoint", 1) = 0) Or (StrComp(sgUserName, "Guide", 1) = 0) Then
        'slInputDate = InputBox$(slMsg, "Date/Time", slDate)
        sgGenMsg = slMsg
        sgCMCTitle(0) = "Done"
        sgCMCTitle(1) = "Cancel"
        sgCMCTitle(2) = ""
        sgCMCTitle(3) = ""
        igDefCMC = 0
        igEditBox = 1
        sgEditValue = slDate
        frmGenMsg.Show vbModal
        If igAnsCMC = 0 Then
            slInputDate = sgEditValue
            If Len(slInputDate) <> 0 Then
                If gIsDate(slInputDate) Then
                    If DateValue(gAdjYear(slDate)) <> DateValue(gAdjYear(slInputDate)) Then
                        'Add year if missing
                        sgNowDate = gAdjYear(slInputDate)
                    Else
                        '7/24/08:Retain date if not changed
                    End If
                End If
            Else
                sgNowDate = ""
            End If
        End If
        slDateTime = " " & Format$(gNow(), "ddd, m/d/yy h:mm AM/PM")
        mnuDate.Caption = slDateTime & " " & sgDateBrannerMsg
    Else
        MsgBox slMsg & "; User: " & sgUserName, vbOKOnly + vbInformation, "Date/Time"
    End If
End Sub
'Private Sub mMakeNetReport()
'    Dim slCommandLine As String
'    'best to make these global
'    Const MYINTERFACEVERSION As String = "/Version1.0"
'    Const DEBUGMODE = "/D"
'
'    'Dan M 12/16/09  run csiNetReporterAlternate IF 'test' is in the folder name.
'    slCommandLine = mBuildAlternateAsNeeded
'    If (Len(sgSpecialPassword) = 4) Then
'        slCommandLine = slCommandLine & DEBUGMODE & " "
'    End If
'    slCommandLine = slCommandLine & MYINTERFACEVERSION & " "
'    If LenB(sgStartupDirectory) Then
'        slCommandLine = slCommandLine & " """ & sgStartupDirectory & """ "
'        slCommandLine = slCommandLine & " /PreRun"
'    End If
'    On Error GoTo errornoexe
'        Shell slCommandLine 'batch
'        bgReportModuleRunning = True
'    Exit Sub
'errornoexe:
'        bgReportModuleRunning = False
'End Sub

Private Sub mManageFormatsXRef()
    Dim llLoop As Long

    Screen.MousePointer = vbHourglass
    On Error GoTo ErrHand
    If Not gPopFormats() Then
        Screen.MousePointer = vbDefault
        MsgBox "Unable to Load Existing Format Names", vbApplicationModal + vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If
    If UBound(tgFormatInfo) <= LBound(tgFormatInfo) Then
        Screen.MousePointer = vbDefault
        MsgBox "Define format names prior to Adjusting Cross References", vbApplicationModal + vbExclamation + vbOKOnly, "Formats"
        Exit Sub
    End If
    If Not mPopFormatLink() Then
        Screen.MousePointer = vbDefault
        MsgBox "Unable to Load Existing Format Cross Refences", vbApplicationModal + vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If
    
    If UBound(tmFormatLinkInfo) <= LBound(tmFormatLinkInfo) Then
        Screen.MousePointer = vbDefault
        MsgBox "Update Station Information via Import prior to adjusting Cross References", vbApplicationModal + vbExclamation + vbOKOnly, "Cross References"
        Exit Sub
    End If
    ReDim tgNewNamesImported(0 To 0) As NEWNAMESIMPORTED
    For llLoop = LBound(tmFormatLinkInfo) To UBound(tmFormatLinkInfo) - 1 Step 1
        tgNewNamesImported(UBound(tgNewNamesImported)).sNewName = tmFormatLinkInfo(llLoop).sExtFormatName
        tgNewNamesImported(UBound(tgNewNamesImported)).lUpdateStationIndex = llLoop
        tgNewNamesImported(UBound(tgNewNamesImported)).lReplaceCode = tmFormatLinkInfo(llLoop).iIntFmtCode
        tgNewNamesImported(UBound(tgNewNamesImported)).iCount = 1
        ReDim Preserve tgNewNamesImported(0 To UBound(tgNewNamesImported) + 1) As NEWNAMESIMPORTED
    Next llLoop
    igNewNamesImportedType = 5
    frmCategoryMatching.Show vbModal
    If igNewNamesImportedReturn = 0 Then
        Erase tmFormatLinkInfo
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    For llLoop = LBound(tgNewNamesImported) To UBound(tgNewNamesImported) - 1 Step 1
        If tmFormatLinkInfo(tgNewNamesImported(llLoop).lUpdateStationIndex).iIntFmtCode <> CInt(tgNewNamesImported(llLoop).lReplaceCode) Then
            SQLQuery = "UPDATE flt"
            SQLQuery = SQLQuery & " SET fltIntFmtCode = " & CInt(tgNewNamesImported(llLoop).lReplaceCode)
            SQLQuery = SQLQuery & " WHERE fltCode = " & tmFormatLinkInfo(tgNewNamesImported(llLoop).lUpdateStationIndex).iCode
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                Exit For
            End If
        End If
    Next llLoop
    Erase tmFormatLinkInfo
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    gHandleError "AffErorLog.txt", "frmMain-mManageFormatsXRef"
    Exit Sub
End Sub

Private Function mPopFormatLink()
    Dim flt_rst As ADODB.Recordset
    Dim ilUpper As Integer
    Dim llMax As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(fltCode) from flt"
    Set rst = gSQLSelectCall(SQLQuery, "frmMain: mPopFormatLink")
    If IsNull(rst(0).Value) Then
        ReDim tmFormatLinkInfo(0 To 0) As FORMATLINKINFO
        mPopFormatLink = True
        Exit Function
    End If
    
    llMax = rst(0).Value
    ReDim tmFormatLinkInfo(0 To llMax) As FORMATLINKINFO
    
    SQLQuery = "Select fltCode, fltExtFormatName, fltIntFmtCode from flt "
    Set flt_rst = gSQLSelectCall(SQLQuery, "frmMain: mPopFormatLink")
    ilUpper = 0
    While Not flt_rst.EOF
        tmFormatLinkInfo(ilUpper).iCode = flt_rst!fltCode
        tmFormatLinkInfo(ilUpper).sExtFormatName = flt_rst!fltExtFormatName
        tmFormatLinkInfo(ilUpper).iIntFmtCode = flt_rst!fltIntFmtCode
        ilUpper = ilUpper + 1
        flt_rst.MoveNext
    Wend

    ReDim Preserve tmFormatLinkInfo(0 To ilUpper) As FORMATLINKINFO

   
   mPopFormatLink = True
   flt_rst.Close
   Exit Function

ErrHand:
    gHandleError "AffErorLog.txt", "frmMain-mPopFormatLink"
    mPopFormatLink = False
    Exit Function
End Function
Private Function mUpdatesAvailable() As Boolean
    Dim myInfo As csiInstallerInfo
    Dim ilPos As Integer
    Dim slCsiSystemPath As String
    Dim blRet As Boolean
    Dim myFile As FileSystemObject
    Dim ilVersion As Integer
    
    ilVersion = mGetCsiVersion()
    Set myFile = New FileSystemObject
    Set myInfo = New csiInstallerInfo
    ilPos = InStr(1, sgExeDirectory, "\exe", vbTextCompare)
    'dan 11/16/11 look in more places
    slCsiSystemPath = Mid$(sgExeDirectory, 1, ilPos) & "Setup\ClientInstall\System32"
    If Not myFile.FolderExists(slCsiSystemPath) Then
        ilPos = InStrRev(sgDBPath, "\", Len(sgDBPath) - 1)
        If ilPos > 0 Then
            slCsiSystemPath = Mid$(sgDBPath, 1, ilPos) & "Setup\ClientInstall\System32"
        End If
        If Not myFile.FolderExists(slCsiSystemPath) Then
            ilPos = InStrRev(sgImportDirectory, "\", Len(sgImportDirectory) - 1)
            If ilPos > 0 Then
                slCsiSystemPath = Mid$(sgImportDirectory, 1, ilPos) & "Setup\ClientInstall\System32"
            End If
            If Not myFile.FolderExists(slCsiSystemPath) Then
                gLogMsg "The path " & slCsiSystemPath & " doesn't exist.  Cannot test if updates are available in traffic-mUpdatesAvailable.", "UpdateErrors.txt", False
                mUpdatesAvailable = False
                GoTo Cleanup
            End If
        End If
    End If
    'If myInfo.Start(slCsiSystemPath) Then
    If myInfo.Start(slCsiSystemPath, ilVersion) Then
        If Not myInfo.CompareCheckers() Then
            If myInfo.FailureReason = NOERROR Then
                blRet = True
            Else
                gLogMsg "Problem in frmMain:mUpdatesAvailable.  " & myInfo.FailureMessage, "UpdateErrors.txt", False
            End If
        End If
    Else
        gLogMsg "The path " & slCsiSystemPath & " doesn't exist.  Cannot test if updates are available in traffic-mUpdatesAvailable.", "UpdateErrors.txt", False
    End If
    mUpdatesAvailable = blRet
Cleanup:
    Set myInfo = Nothing
    Set myFile = Nothing
End Function
Private Function mGetCsiVersion() As Integer
    mGetCsiVersion = CInt(App.Major & App.Minor)
End Function
Private Function mCheckSetDDFFields() As Integer
    Dim rst As ADODB.Recordset
    mCheckSetDDFFields = True
    On Error GoTo ErrHand
    SQLQuery = "Select siteDDF092710 from site"
    Set rst = gSQLSelectCall(SQLQuery, "frmMain: mCheckSetDDFFields")
    If Not rst.EOF Then
        If Trim$(rst!siteDDF092710) = "" Or rst!siteDDF092710 = "N" Then
            sgSetFieldCallSource = "S"
            SQLQuery = "Update site Set "
            SQLQuery = SQLQuery & "siteDDF092710 = '" & "P" & "' "
            SQLQuery = SQLQuery & " Where siteCode = " & 1
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand:
                gHandleError "AffErrorLog.txt", "Main-mCheckSetDDFFields"
                mCheckSetDDFFields = False
                On Error Resume Next
                rst.Close
                Exit Function
            End If
            frmSetDDFFields.Show vbModal
            SQLQuery = "Select siteDDF092710 from site"
            Set rst = gSQLSelectCall(SQLQuery, "frmMain: mCheckSetDDFFields")
            If Not rst.EOF Then
                If rst!siteDDF092710 <> "Y" Then
                    mCheckSetDDFFields = False
                End If
            End If
        ElseIf rst!siteDDF092710 = "P" Then
            MsgBox "Affiliate Conversion in progress, Affiliate can't be run at this time", vbOKOnly + vbInformation
            mCheckSetDDFFields = False
        End If
    End If
    On Error Resume Next
    rst.Close
    Exit Function
ErrHand:
    gHandleError "AffErorLog.txt", "frmMain-mCheckSetDDFFields"
    mCheckSetDDFFields = False
End Function

Private Sub mSetCommentSource()
    Dim ilLoop As Integer
    Dim rst_cst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "SELECT Count(cstCode) FROM CST"
    Set rst_cst = gSQLSelectCall(SQLQuery, "frmMain: mSetCommentSource")
    If rst_cst(0).Value = 0 Then
        For ilLoop = 1 To 5 Step 1
    
            SQLQuery = "Insert Into cst ( "
            SQLQuery = SQLQuery & "cstCode, "
            SQLQuery = SQLQuery & "cstName, "
            SQLQuery = SQLQuery & "cstDefault, "
            SQLQuery = SQLQuery & "cstSortCode, "
            SQLQuery = SQLQuery & "cstUnused "
            SQLQuery = SQLQuery & ") "
            SQLQuery = SQLQuery & "Values ( "
            SQLQuery = SQLQuery & ilLoop & ", "
            Select Case ilLoop
                Case 1
                    SQLQuery = SQLQuery & "'" & gFixQuote("Call: Outgoing") & "', "
                    SQLQuery = SQLQuery & "'" & gFixQuote("Y") & "', "
                Case 2
                    SQLQuery = SQLQuery & "'" & gFixQuote("Call: Incoming") & "', "
                    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
                Case 3
                    SQLQuery = SQLQuery & "'" & gFixQuote("E-Mail: Outgoing") & "', "
                    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
                Case 4
                    SQLQuery = SQLQuery & "'" & gFixQuote("Mass E-Mail: Outgoing") & "', "
                    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
                Case 5
                    SQLQuery = SQLQuery & "'" & gFixQuote("E-Mail: Incoming") & "', "
                    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
            End Select
            SQLQuery = SQLQuery & ilLoop & ", "
            SQLQuery = SQLQuery & "'" & "" & "' "
            SQLQuery = SQLQuery & ") "
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand:
                gHandleError "AffErrorLog.txt", "Main-mSetCommentSource"
                On Error Resume Next
                rst_cst.Close
                Exit Sub
            End If
        Next ilLoop
    End If
    On Error Resume Next
    rst_cst.Close
    Exit Sub
ErrHand:
    gHandleError "AffErorLog.txt", "frmMain-mSetCommentSource"
End Sub
Private Sub mSetUAFToAborted(llUlfCode As Long)
    Dim ilRet As Integer
    Dim slStatus As String
    Dim llUaf As Long
    ReDim llUafCode(0 To 0) As Long
    
    On Error GoTo ErrHand:
    slStatus = "I"
    SQLQuery = "SELECT * FROM Uaf_User_Activity"
    SQLQuery = SQLQuery & " WHERE uafUlfCode = " & llUlfCode
    SQLQuery = SQLQuery & " AND uafStatus = '" & slStatus & "'"
    SQLQuery = SQLQuery & " ORDER BY uafStartDate DESC, uafStartTime Desc"
    Set rst_Uaf = gSQLSelectCall(SQLQuery, "frmMain: mSetUAFToAborted")
    Do While Not rst_Uaf.EOF
        llUafCode(UBound(llUafCode)) = rst_Uaf!uafCode
        ReDim Preserve llUafCode(0 To UBound(llUafCode) + 1) As Long
        rst_Uaf.MoveNext
    Loop
    For llUaf = 0 To UBound(llUafCode) - 1 Step 1
        SQLQuery = "UPDATE Uaf_User_Activity"
        SQLQuery = SQLQuery & " SET uafStatus = 'A'" & ", "
        SQLQuery = SQLQuery & " uafEndDate = '" & Format(Now, sgSQLDateForm) & "',"
        SQLQuery = SQLQuery & " uafEndTime = '" & Format(Now, sgSQLTimeForm) & "'"
        SQLQuery = SQLQuery & " WHERE uafCode = " & llUafCode(llUaf)
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/13/16: Replaced GoSub
            'GoSub ErrHand:
            gHandleError "AffErrorLog.txt", "Main-mSetUAFToAborted"
            Exit For
        End If
    Next llUaf
    On Error Resume Next
    rst_Uaf.Close
    Exit Sub
ErrHand:
    gHandleError "AffErorLog.txt", "frmMain-mSetUAFToAborted"
    rst_Uaf.Close
End Sub

Private Sub mInitMonitorMenu()
    Dim ilTask As Integer
    Dim ilMenuIndex As Integer
    Dim ilRet As Integer
    
    ' Get a handle to the top level menu
    lmDate1970 = gDateValue("1/1/1970")
    ilRet = gOpenMKDFile(hgTmf, "Tmf.btr")
    imTmfRecLen = Len(tmTmf)
    sgTimeZone = Left$(gGetLocalTZName(), 1)
    hmMonitor = GetMenu(frmMain.hwnd)
    gInitTaskInfo
    
    ' Get a handle to the top level menu
    hmAlert = GetMenu(frmMain.hwnd)
    
    '3/27/15: When changed to Different look on the Job screen, for some reason the index changed.
    If mnuGroupName.Visible = False Then
        'ilMenuIndex = 5
        ilMenuIndex = GetAlertMenuIndex(hmAlert, 0) + 3
    Else
        'ilMenuIndex = 6
        ilMenuIndex = GetAlertMenuIndex(hmAlert, 0) + 2
    End If
    
    '8793
    If mMonitoringAllowed() Then
        For ilTask = 0 To UBound(tgTaskInfo) Step 1
            tmTmfSrchKey1.sTaskCode = Trim$(tgTaskInfo(ilTask).sTaskCode)
            ilRet = btrGetEqual(hgTmf, tmTmf, imTmfRecLen, tmTmfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                If (tmTmf.sService <> "N") And (Trim$(tmTmf.sService) <> "") Then
                    gUnpackDateLong tmTmf.iRunningDate(0), tmTmf.iRunningDate(1), tgTaskInfo(ilTask).lRunningDate
                    gUnpackTimeLong tmTmf.iRunningTime(0), tmTmf.iRunningTime(1), True, tgTaskInfo(ilTask).lRunningTime
                    mnuMonitor(ilTask).Visible = True
                    tgTaskInfo(ilTask).iMenuIndex = ilMenuIndex
                    tgTaskInfo(ilTask).lColor = LIGHTYELLOW
                    tgTaskInfo(ilTask).lElapsedTime = 0
                    mSetMonitorInfo ilTask
                    ilMenuIndex = ilMenuIndex + 1
                End If
            End If
        Next ilTask
        lgDateMonitorChecked = gDateValue(Format(Now, "m/d/yy"))
        lgTimeMonitorChecked = gTimeToLong(Format(Now, "h:mm:ssAM/PM"), False)
        tmcMonitor.Interval = 100 * MONITORTIMEINTERVAL
        tmcMonitor.Enabled = True
    End If
End Sub
Private Sub mSetMonitorColor(llBgColor As Long)
    frmDirectory!pbcMonitor.BackColor = llBgColor
    If (llBgColor = DARKGREEN) Or (llBgColor = vbRed) Then
        frmDirectory!pbcMonitor.ForeColor = vbWhite
    Else
        frmDirectory!pbcMonitor.ForeColor = vbBlack
    End If
End Sub

Private Sub mShowMonitor(ilMenuNumber As Integer)
    Dim ilRet As Integer
    Dim hlMonitorBitmap As Long
    Dim llMenuID As Long

    hlMonitorBitmap = mCopyPictureImage(frmDirectory!pbcMonitor)
    llMenuID = GetMenuItemID(hmMonitor, ilMenuNumber)

    'If hmBitmap = 0 Then hmBitmap = CopyPictureImage(frmDirectory!pbcAlert)
    ' And replace it with a bitmap.
    ilRet = ModifyMenuBynum(hmMonitor, ilMenuNumber, MF_BITMAP Or MF_BYPOSITION Or MF_ENABLED, llMenuID, hlMonitorBitmap)
End Sub

Private Sub mSetMonitorInfo(ilTask As Integer)
    Dim ilMenuIndex  As Integer
    
    ' Get a handle to the top level menu
    hmAlert = GetMenu(frmMain.hwnd)
    
    'Get AlertMenu Index
    ilMenuIndex = GetAlertMenuIndex(hmAlert, 0)
    
    mSetMonitorColor tgTaskInfo(ilTask).lColor
    frmDirectory!pbcMonitor.CurrentX = 15
    frmDirectory!pbcMonitor.CurrentY = 0
    frmDirectory!pbcMonitor.Print Trim$(tgTaskInfo(ilTask).sTaskCode)
    
    'TTP 10248
    If mnuAlert.Visible = False Then
        mShowMonitor tgTaskInfo(ilTask).iMenuIndex - (Abs(ilMenuIndex - 5))
    Else
        mShowMonitor tgTaskInfo(ilTask).iMenuIndex - (Abs(ilMenuIndex - 4))
    End If
    
    mSetSubmenuStatus ilTask
End Sub

Private Sub mSetSubmenuStatus(ilTask As Integer)
    Select Case Trim$(tgTaskInfo(ilTask).sTaskCode)
        Case "CSS"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuCSSStatus.Visible = True
            Else
                mnuCSSStatus.Visible = False
            End If
        Case "SSB"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuSSBStatus.Visible = True
            Else
                mnuSSBStatus.Visible = False
            End If
        Case "AEQ"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuAEQStatus.Visible = True
            Else
                mnuAEQStatus.Visible = False
            End If
        Case "ASI"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuASIStatus.Visible = True
            Else
                mnuASIStatus.Visible = False
            End If
        Case "ARQ"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuARQStatus.Visible = True
            Else
                mnuARQStatus.Visible = False
            End If
        Case "ASG"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuASGStatus.Visible = True
            Else
                mnuASGStatus.Visible = False
            End If
    
        Case "SC"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuSCStatus.Visible = True
            Else
                mnuSCStatus.Visible = False
            End If
        Case "CE"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuCEStatus.Visible = True
            Else
                mnuCEStatus.Visible = False
            End If
        Case "SFE"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuSFEStatus.Visible = True
            Else
                mnuSFEStatus.Visible = False
            End If
        Case "ME"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuMEStatus.Visible = True
            Else
                mnuMEStatus.Visible = False
            End If
        Case "EPE"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuEPEStatus.Visible = True
            Else
                mnuEPEStatus.Visible = False
            End If
        Case "ERE"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuEREStatus.Visible = True
            Else
                mnuEREStatus.Visible = False
            End If
        Case "GPE"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuGPEStatus.Visible = True
            Else
                mnuGPEStatus.Visible = False
            End If
        Case "BD"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuBDStatus.Visible = True
            Else
                mnuBDStatus.Visible = False
            End If
        Case "AMB"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuAMBStatus.Visible = True
            Else
                mnuAMBStatus.Visible = False
            End If
        '7967
        Case "WVM"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuWviStatus.Visible = True
            Else
                mnuWviStatus.Visible = False
            End If
        'D.S. 03/21/18
        Case "PB"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuPbStatus.Visible = True
            Else
                mnuPbStatus.Visible = False
            End If
        Case "CAI"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuCAIStatus.Visible = True
            Else
                mnuCAIStatus.Visible = False
            End If
        Case "RE"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuREstatus.Visible = True
            Else
                mnuREstatus.Visible = False
            End If
        Case "CRE"
            If tgTaskInfo(ilTask).lColor = vbRed Then
                mnuCREstatus.Visible = True
            Else
                mnuCREstatus.Visible = False
            End If
    End Select
End Sub
Private Sub mCheckGG()
    Dim c As Integer
    Dim slName As String
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim slStr As String
    Dim llDate As Long
    Dim llNow As Long
    
    Dim gg_rst As ADODB.Recordset
    
    If Not frmDirectory.Visible Then
        Exit Sub
    End If
    If imLastHourGGChecked = Hour(Now) Then
        Exit Sub
    End If
    imLastHourGGChecked = Hour(Now)
    
    If bgInternalGuide Then
        Exit Sub
    End If
    
    SQLQuery = "Select safName From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set gg_rst = gSQLSelectCall(SQLQuery, "frmMain: mCheckGG")
    If Not gg_rst.EOF Then
        slName = Trim$(gg_rst!safName)
        ilField1 = Asc(slName)
        slStr = Mid$(slName, 2, 5)
        llDate = Val(slStr)
        llNow = gDateValue(Format$(gNow(), "m/d/yy"))
        ilField2 = Asc(Mid$(slName, 11, 1))
        If (ilField1 = 0) And (ilField2 = 1) Then
            If llDate <= llNow Then
                ilField2 = 0
            End If
        End If
        If (ilField1 = 0) And (ilField2 = 0) Then
        
            igGGFlag = 0
            
            For c = 0 To UBound(sgUstWin)
                sgUstWin(c) = "H"
            Next c
            sgUstClear = "N"
            sgUstActivityLog = "H"
            sgUstDelete = "N"
            sgUstPledge = "N"
            sgUstAllowCmmtChg = "N"
            sgUstAllowCmmtDelete = "N"
            sgExptSpotAlert = "N"
            sgExptISCIAlert = "N"
            sgTrafLogAlert = "N"
            sgChgExptPriority = "N"
            sgExptSpec = "N"
            sgChgRptPriority = "N"
            
            sgSpfUseCartNo = "Y"
            sgSpfRemoteUsers = "N"
            sgSpfUsingFeatures2 = Chr$(0)
            sgSpfUsingFeatures5 = Chr$(0)
            sgSpfUsingFeatures9 = Chr$(0)
            sgSpfSportInfo = Chr$(0)
            sgSpfUseProdSptScr = "A"
            gUsingWeb = False
            gUsingUnivision = False
            gISCIExport = False
            sgRCSExportCart4 = "N"
            sgRCSExportCart5 = "N"
            gUsingXDigital = False
            gWegenerExport = False
            gOLAExport = False
            gWegenerExport = False
            
            mnuBackup.Enabled = False
            mnuImport.Enabled = False
            mnuExport.Enabled = False
            mnuMerge.Enabled = False
            mnuManageFormats.Enabled = False
            mnuUtilities.Enabled = False
            mnuGroupName.Enabled = False
            
            frmDirectory!cmcManagement.Enabled = False
            frmDirectory!cmcPostBuy.Enabled = False
            frmDirectory!cmdAffAE.Enabled = False
            frmDirectory!cmdAffTimes.Enabled = False
            frmDirectory!cmdAgreements.Enabled = False
            frmDirectory!cmdContact.Enabled = False
            frmDirectory!cmdCP.Enabled = False
            frmDirectory!cmdCPReturns.Enabled = False
            frmDirectory!cmdEMail.Enabled = False
            frmDirectory!cmdExports.Enabled = False
            frmDirectory!cmdLog.Enabled = False
            frmDirectory!cmdOptions.Enabled = False
            frmDirectory!cmdPostLog.Enabled = False
            frmDirectory!cmdPreLog.Enabled = False
            frmDirectory!cmdRadar.Enabled = False
            frmDirectory!cmdSite.Enabled = False
            frmDirectory!cmdStation.Enabled = False
            frmDirectory!cmdReports.Enabled = False

            frmDirectory!lacManagement.Enabled = False
            frmDirectory!lacPostBuy.Enabled = False
            frmDirectory!lacAgreements.Enabled = False
            frmDirectory!lacCPReturns.Enabled = False
            frmDirectory!lacEMail.Enabled = False
            frmDirectory!lacExports.Enabled = False
            frmDirectory!lacOptions.Enabled = False
            frmDirectory!lacPostLog.Enabled = False
            frmDirectory!lacRadar.Enabled = False
            frmDirectory!lacSite.Enabled = False
            frmDirectory!lacStation.Enabled = False

            frmDirectory!lacManagement.FontBold = False
            frmDirectory!lacPostBuy.FontBold = False
            frmDirectory!lacAgreements.FontBold = False
            frmDirectory!lacCPReturns.FontBold = False
            frmDirectory!lacEMail.FontBold = False
            frmDirectory!lacExports.FontBold = False
            frmDirectory!lacOptions.FontBold = False
            frmDirectory!lacPostLog.FontBold = False
            frmDirectory!lacRadar.FontBold = False
            frmDirectory!lacSite.FontBold = False
            frmDirectory!lacStation.FontBold = False

            frmDirectory!lacReports.Enabled = False
        End If
        gSetRptGGFlag slName
    End If
    If (igGGFlag <> 0) Or (igRptGGFlag <> 0) Then
        frmDirectory!cmdReports.Enabled = True
        frmDirectory!lacReports.Enabled = True
    End If
    gg_rst.Close
End Sub
Private Sub mUpdateVendors()
    Dim tlVendors() As VendorInfo
    Dim slSql As String
    Dim ilIndex As Integer
    Dim ilCode As Integer
    
    tlVendors = gGetAvailableVendors()
    For ilIndex = 0 To UBound(tlVendors) - 1
        ilCode = tlVendors(ilIndex).iIdCode
        slSql = "Select wvtName,wvtApprovalPassword,wvtExportMethod,wvtSendUpdatesOnly,wvtImportMethod,wvtHierarchy,wvtDeliveryType FROM WVT_Vendor_Table WHERE wvtVendorID = " & ilCode
        Set rst = gSQLSelectCall(slSql, "frmMain: mUpdateVendors")
        Do While Not rst.EOF
            With tlVendors(ilIndex)
                If Trim$(gFixQuote(.sName)) <> Trim$(rst!wvtName) Or Trim$(gFixQuote(.sApprovalPassword)) <> Trim$(rst!wvtApprovalPassword) Or .iExportMethod <> rst!wvtExportMethod Or .iImportMethod <> rst!wvtImportMethod Or .iHierarchy <> rst!wvthierarchy Or .sDeliveryType <> rst!wvtDeliveryType Or .sSendUpdatesOnly <> Trim$(rst!wvtSendUpdatesOnly) Then
                    slSql = "Update WVT_Vendor_Table SET wvtName = '" & Trim$(gFixQuote(.sName)) & "', wvtApprovalPassword = '" & Trim$(gFixQuote(.sApprovalPassword)) & "', wvtExportMethod = " & .iExportMethod & ", wvtImportMethod = " & .iImportMethod & ", wvtHierarchy = " & .iHierarchy & ", wvtDeliveryType = '" & .sDeliveryType & "', wvtSendUpdatesOnly = '" & .sSendUpdatesOnly & "' "
                    slSql = slSql & " WHERE wvtVendorID = " & ilCode
                     If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub ErrHand:
                        gHandleError "AffErrorLog.txt", "Main-mUpdateVendors"
                        Exit Sub
                     End If
                End If
            End With
            rst.MoveNext
        Loop
    Next ilIndex
    Exit Sub
ErrHandler:
    gHandleError "", "frmMain-mUpdateVendors"
End Sub
Private Function mConvertToVAT() As Boolean
    Dim slSql As String
    Dim blError As Boolean
    Dim tlVendors() As VendorInfo
    Dim c As Integer
    Dim llAgreementsChanged As Long
    Dim blRet As Boolean
    Dim llTotalChange As Long
    Dim hlLog As Integer
    '8146
    Dim llAgreementVendorChanged() As Long
    Dim slError As String
    
    slError = ""
    blRet = False
    blError = False
    llTotalChange = 0
    slSql = "Select count(*) as amount FROM VAT_Vendor_Agreement"
    Set rst = gSQLSelectCall(slSql, "frmMain: mConvertToVAT")
    If Not rst.EOF Then
        If rst!amount = 0 Then
            blRet = True
            '7907
            frmProgressMsg.Show
            frmProgressMsg.SetMessage 0, "Updating agreements, please wait..."
            tlVendors = gGetAvailableVendors()
            ReDim llAgreementVendorChanged(UBound(tlVendors) - 1)
            For c = 0 To UBound(tlVendors) - 1
                DoEvents
                llAgreementsChanged = gConvertToVat(tlVendors(c).iIdCode)
                'add to vendor table?
                If llAgreementsChanged > 0 Then
                    llAgreementVendorChanged(c) = llAgreementsChanged
                    llTotalChange = llTotalChange + llAgreementsChanged
                    frmProgressMsg.SetMessage 0, "Updating agreements, please wait..." & vbCrLf & llTotalChange & " agreements changed..."
                    'maybe being too cautious...make sure vendor doesn't exist before adding.
                    slSql = "select wvtVendorId from WVT_Vendor_Table where wvtvendorid = " & tlVendors(c).iIdCode
                    Set rst = gSQLSelectCall(slSql, "frmMain: mConvertToVAT")
                    If rst.EOF Then
                        slSql = mWvtInsertSql(tlVendors(c))
                        If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                            '6/12/16: Replaced GoSub
                            'GoSub ErrHandler:
                            gHandleError "VendorConversion.txt", "frmMain-mConvertToVAT"
                            blError = True
                            blRet = False
                            GoTo Cleanup
                        End If
                    End If
                ElseIf llAgreementsChanged < 0 Then
                    slError = slError & " " & tlVendors(c).sName
                End If
            Next c
            '8145
            If llTotalChange > 0 Then
                gLogMsgWODT "OA", hlLog, sgMsgDirectory & "VendorConversion.txt"
                gLogMsgWODT "W", hlLog, "Vendor Conversion Result List, Started: " & gNow()
                gLogMsgWODT "W", hlLog, "Converted " & llTotalChange & " agreements to vat table"
                For c = 0 To UBound(llAgreementVendorChanged)
                    If llAgreementVendorChanged(c) > 0 Then
                        gLogMsgWODT "W", hlLog, "  " & tlVendors(c).sName & ": " & llAgreementVendorChanged(c)
                    End If
                Next c
        '            Else
        '                gLogMsgWODT "W", hlLog, "No agreements needed to be converted."
                gLogMsgWODT "C", hlLog, ""
            End If
        End If
    End If
Cleanup:
    Unload frmProgressMsg
    Erase tlVendors
    Erase llAgreementVendorChanged
    If blError Then
        gMsgBox "There was a problem with Vendor Conversion:  could not insert to wvt.  Contact Counterpoint with " & sgMsgDirectory & "VendorConversion.txt"
    ElseIf Len(slError) > 0 Then
        gMsgBox "There was a problem with Vendor Conversion: could not insert to Vat.  Contact Counterpoint with " & sgMsgDirectory & "VendorConversion.txt"
        gLogMsgWODT "OA", hlLog, sgMsgDirectory & "VendorConversion.txt"
        gLogMsgWODT "W", hlLog, "Vendor Conversion Result List, Started: " & gNow()
        gLogMsgWODT "W", hlLog, "Problem with insertions to Vat: " & slError
        gLogMsgWODT "C", hlLog, ""
    ElseIf llTotalChange > 0 Then
        gMsgBox "Vendor Conversion ran.  Please review '" & sgMsgDirectory & "VendorConversion.txt' to see which exports were affected."
    End If
    mConvertToVAT = blRet
    Exit Function
ErrHandler:
    gHandleError "VendorConversion.txt", "frmMain-mConvertToVAT"
    blError = True
    blRet = False
    GoTo Cleanup
End Function
Private Function mWvtInsertSql(tlMyVendor As VendorInfo) As String
    Dim slRet As String

On Error GoTo errbox
    With tlMyVendor
        slRet = "Insert into WVT_Vendor_Table (wvtVendorID, wvtName,wvtApprovalPassword,wvtVendorUserName, wvtVendorPassword, wvtExportMethod, "
        slRet = slRet & "wvtAddress, wvtStationUserName,wvtStationPassword, wvtSendUpdatesOnly, wvtImportMethod, wvtHierarchy, wvtDeliveryType, wvtIsOverridable) "
        slRet = slRet & " VALUES ( " & .iIdCode & ", '" & Trim$(gFixQuote(.sName)) & "','" & Trim$(gFixQuote(.sApprovalPassword)) & "','" & Trim$(gFixQuote(.sVendorUserName)) & "','" & Trim$(gFixQuote(.sVendorPassword)) & "', "
        slRet = slRet & .iExportMethod & ",'" & Trim$(gFixQuote(.sAddress)) & "','" & Trim$(gFixQuote(.sStationUserName)) & "','" & Trim$(gFixQuote(.sStationPassword)) & "','" & .sSendUpdatesOnly & "'," & .iImportMethod & " , " & .iHierarchy & ", '" & .sDeliveryType & "','" & .sSendUpdatesOnly & "')"
    End With
    mWvtInsertSql = slRet
    Exit Function
errbox:
    mWvtInsertSql = ""
End Function

Private Sub tmcWebConnectIssue_Timer()
    '7967  in past, checking with web did 20 retries.  Now only try once and come here to retry if that fails
    Static ilMinutes As Integer
    
    ilMinutes = ilMinutes + 1
    '8133 10 minute span
    If ilMinutes = 10 Then
        ilMinutes = 0
        If Not bgVendorToWebAllowed Then
            If gVendorToWebAllowed() Then
                '8129
'                If gGrabWebVendorIssuesFromWeb(False) Then
'                    'force task monitor timer to run test to web
'                    igWVImportElapsed = 0
'                    dgWvImportLast = 0
'                    ' set by gVendorToWebAllowed bgVendorToWebAllowed = True
'                    tmcWebConnectIssue.Enabled = False
'                Else
'                    gLogMsg "VendorToWeb-retry of gGrabWebVendorIssuesFromWeb failed", "AffErrorLog.Txt", False
'                End If
                'force task monitor timer to run test to web
                igWVImportElapsed = 0
                dgWvImportLast = 0
                tmcWebConnectIssue.Enabled = False
            Else
                gLogMsg "VendorToWeb-retry of gVendorToWebAllowed failed", "AffErrorLog.Txt", False
            End If
        End If
    End If
End Sub
Public Sub gAllowedExportsImportsInMenu(blIsOn As Boolean, ilVendor As Vendors)
    '8156
    
    Select Case ilVendor
        Case Vendors.iDc
            mnuExportIDC.Enabled = blIsOn
        Case Vendors.NetworkConnect
            mnuExportMarketron.Enabled = blIsOn
            mnuImportMarketron.Enabled = blIsOn
        Case Vendors.Wegener_Compel
            mnuExportWegener.Enabled = blIsOn
            mnuImportWegenerCompel.Enabled = blIsOn
        Case Vendors.Wegener_IPump
            mnuExportIPump.Enabled = blIsOn
            'dan 9/8/16 not currently allowed
            'mnuImportIpump.Enabled = blIsOn
        Case Vendors.XDS_Break, Vendors.XDS_ISCI
            mnuExportXDigital.Enabled = blIsOn
    End Select
End Sub
Private Function mMonitoringAllowed() As Boolean
    'also sets emMonitoringAllowed for testing in timer, and sets the monitor id value to be output in email
    Dim blRet As Boolean
    Dim slSection As String
    
    blRet = True
    If igTestSystem <> True Then
        slSection = "Locations"
    Else
        slSection = "TestLocations"
    End If
    
    gLoadOption slSection, "MonitorID", smMonitorID
    emMonitoringAllowed = MonitorNormal
    If Len(Trim$(smMonitorID)) > 0 Then
        Select Case UCase(smMonitorID)
            Case "OFF", "ICONSOFF"
                emMonitoringAllowed = MonitorOff
                blRet = False
            Case "EMAILOFF"
                emMonitoringAllowed = MonitorNoEmail
        End Select
        smMonitorID = " MonitorID: " & smMonitorID
    Else
        smMonitorID = " MonitorID is not set"
    End If
    mMonitoringAllowed = blRet
End Function

Private Sub mTempSendServiceEmail(slSubject As String, slBody As String)

    If slSubject = "" Then
        slSubject = "Automated message from a client"
    End If
    Set ogEmailer = New CEmail

    With ogEmailer
        .FromAddress = "AClient@Counterpoint.net"
        .FromName = Trim$(sgClientName)
        .AddTOAddress "dan.dbmich@gmail.com", "Dan"
        .Subject = slSubject
        .Message = slBody
        .SetHost "smtpauth.hosting.earthlink.net", 587, "emailsender@counterpoint.net", "Csi44Sic", True
        If Not .Send() Then
            gLogMsg "Email could not be sent from frmMain-mTempSendServiceEmail: ", "AffErrorLog.Txt", False
        End If
    End With
    Set ogEmailer = Nothing
End Sub

Private Sub mSetJobButtons()
    Dim slSQLQuery As String
    Dim blShowTraffic As Boolean
    Dim slUserName As String
    Dim slUserPassword As String
    
    cmcAgreement.Left = cmcStation.Left + cmcStation.Width
    cmcEMail.Left = cmcAgreement.Left + cmcAgreement.Width
    cmcLog.Left = cmcEMail.Left + cmcEMail.Width
    cmcAffIdavit.Left = cmcLog.Left + cmcLog.Width
    cmcPostBuy.Left = cmcAffIdavit.Left + cmcAffIdavit.Width
    cmcManagement.Left = cmcPostBuy.Left + cmcPostBuy.Width
    cmcExport.Left = cmcManagement.Left + cmcManagement.Width
    cmcSite.Left = cmcExport.Left + cmcExport.Width
    cmcUser.Left = cmcSite.Left + cmcSite.Width
    cmcRadar.Left = cmcUser.Left + cmcUser.Width
    blShowTraffic = False
    If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
        blShowTraffic = True
    Else
        slSQLQuery = "Select ustName, ustPassword From ust Where "
        slSQLQuery = slSQLQuery & " ustCode = " & igUstCode
        Set rst_Ust = gSQLSelectCall(slSQLQuery)
        If Not rst_Ust.EOF Then
            slSQLQuery = "Select urfName, urfPassword From urf_User_options Where "
            slSQLQuery = slSQLQuery & " urfDelete <> 'Y' "
            Set rst_Urf = gSQLSelectCall(slSQLQuery)
            Do While Not rst_Urf.EOF
                slUserName = Trim$(UCase(gDecryptField(Trim$(rst_Urf!urfName))))
                slUserPassword = Trim$(UCase(gDecryptField(Trim$(rst_Urf!urfPassword))))
                If (slUserName = UCase(sgUserName)) And (slUserPassword = UCase(Trim$(rst_Ust!ustpassword))) Then
                    blShowTraffic = True
                    Exit Do
                End If
                rst_Urf.MoveNext
            Loop
            rst_Urf.Close
        End If
        rst_Ust.Close
    End If
    If blShowTraffic Then
        cmcTraffic.Left = cmcRadar.Left + cmcRadar.Width
        cmcTraffic.Width = 3 * cmcTraffic.Width
        cmcReport.Left = cmcTraffic.Left + cmcTraffic.Width
        cmcReport.Width = pbcMsgArea.Width - cmcTraffic.Left - cmcTraffic.Width + 15
    Else
        cmcTraffic.Visible = False
        cmcReport.Left = cmcRadar.Left + cmcRadar.Width
        cmcReport.Width = pbcMsgArea.Width - cmcRadar.Left - cmcRadar.Width + 15
    End If
End Sub

Function GetAlertMenuIndex(hMenu As Long, spaces As Integer) As Integer
    'Menu's (through the user32 API) are 0 based.
    'this function finds the "Alert" menu index. This is key'd off of finding the "Help" menu, and Adding 1.
    'The Menu's to the Left of Help (File, Accessories, Group Name) may or may not be Visible; and this affects the indexes
    Dim num As Integer
    Dim i As Integer
    Dim LENGTH As Long
    Dim sub_hmenu As Long
    Dim sub_name As String
    
    num = GetMenuItemCount(hMenu)
    For i = 0 To num - 1
        'Save this menu's info.
        sub_hmenu = GetSubMenu(hMenu, i)
        sub_name = Space$(256)
        LENGTH = GetMenuString(hMenu, i, sub_name, Len(sub_name), MF_BYPOSITION)
        sub_name = Left$(sub_name, LENGTH)
        If sub_hmenu <> 0 And spaces = 0 Then
            'Debug.Print i & "=Main " & sub_name
            If sub_name = "&Help" Then
                GetAlertMenuIndex = i + 1
                Exit Function
            End If
        Else
            'Debug.Print Space$(5) & "Sub-" & spaces / 5 & " " & sub_name
        End If
        'Get its child menu's names.
        GetAlertMenuIndex sub_hmenu, spaces + 5
    Next i
    
End Function
