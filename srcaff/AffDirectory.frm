VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmDirectory 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9000
   ClientLeft      =   2625
   ClientTop       =   1530
   ClientWidth     =   13500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "AffDirectory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   13500
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pbcDirectory 
      BorderStyle     =   0  'None
      Height          =   9000
      Left            =   0
      Picture         =   "AffDirectory.frx":08CA
      ScaleHeight     =   9000
      ScaleWidth      =   13515
      TabIndex        =   13
      Top             =   0
      Width           =   13520
      Begin VB.CommandButton cmdReports 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   11160
         Picture         =   "AffDirectory.frx":18C12C
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.CommandButton cmdOptions 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   4515
         Picture         =   "AffDirectory.frx":191916
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   7810
         Width           =   3345
      End
      Begin VB.CommandButton cmdCPReturns 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   4515
         Picture         =   "AffDirectory.frx":1995B0
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3900
         Width           =   3345
      End
      Begin VB.CommandButton cmcPostBuy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   8355
         Picture         =   "AffDirectory.frx":19ED9A
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3900
         Width           =   3345
      End
      Begin VB.CommandButton cmdRadar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   8355
         Picture         =   "AffDirectory.frx":1A4584
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   7810
         Width           =   3345
      End
      Begin VB.CommandButton cmdExports 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   4515
         Picture         =   "AffDirectory.frx":1AC21E
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   5910
         Width           =   3345
      End
      Begin VB.CommandButton cmdSite 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   720
         Picture         =   "AffDirectory.frx":1B1A08
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   7810
         Width           =   3345
      End
      Begin VB.CommandButton cmdPostLog 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   720
         Picture         =   "AffDirectory.frx":1B96A2
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3900
         Width           =   3345
      End
      Begin VB.CommandButton cmcManagement 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   720
         Picture         =   "AffDirectory.frx":1BEE8C
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   5910
         Width           =   3345
      End
      Begin VB.CommandButton cmdAgreements 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   4560
         Picture         =   "AffDirectory.frx":1C4676
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2100
         Width           =   3345
      End
      Begin VB.CommandButton cmdEMail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   8355
         Picture         =   "AffDirectory.frx":1C9E60
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2100
         Width           =   3345
      End
      Begin VB.CommandButton cmdStation 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   720
         Picture         =   "AffDirectory.frx":1CF64A
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2100
         Width           =   3345
      End
      Begin VB.Label lacSetupGroup 
         BackStyle       =   0  'Transparent
         Caption         =   "SETUP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   720
         TabIndex        =   55
         Top             =   7210
         Width           =   3375
      End
      Begin VB.Label lacManagementGroup 
         BackStyle       =   0  'Transparent
         Caption         =   "MANAGEMENT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   720
         TabIndex        =   54
         Top             =   5310
         Width           =   3855
      End
      Begin VB.Label lacComplianceGroup 
         BackStyle       =   0  'Transparent
         Caption         =   "COMPLIANCE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   720
         TabIndex        =   53
         Top             =   3300
         Width           =   3495
      End
      Begin VB.Label lacAffiliationsGroup 
         BackStyle       =   0  'Transparent
         Caption         =   "AFFILIATIONS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   720
         TabIndex        =   52
         Top             =   1500
         Width           =   3615
      End
      Begin VB.Label lblRadarDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Define the RADAR program schedule"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   8355
         TabIndex        =   51
         Top             =   8460
         Width           =   3375
      End
      Begin VB.Label lblUserOptDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set up Affiliate users and permissions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   4560
         TabIndex        =   50
         Top             =   8460
         Width           =   3735
      End
      Begin VB.Label lblSiteOptDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set up system options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   720
         TabIndex        =   49
         Top             =   8460
         Width           =   3495
      End
      Begin VB.Label lblExportCenterDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Configure and queue exports to run in the background"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   4560
         TabIndex        =   48
         Top             =   6555
         Width           =   3375
      End
      Begin VB.Label lblAffMgmtDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View affiliate compliance for 52 weeks on one screen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   720
         TabIndex        =   47
         Top             =   6555
         Width           =   3375
      End
      Begin VB.Label lblPostbuyPlanningDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gather data about where spots aired for an advertiser"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   8355
         TabIndex        =   46
         Top             =   4530
         Width           =   3375
      End
      Begin VB.Label lblAffAffidavitsDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View affiliate logs and perform manual posting"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   4560
         TabIndex        =   45
         Top             =   4530
         Width           =   3375
      End
      Begin VB.Label lblNetworkLogDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View spots created during log generation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   720
         TabIndex        =   44
         Top             =   4530
         Width           =   3375
      End
      Begin VB.Label lblEmailsByVehicleDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Send affiliates mass and individual emails"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   8355
         TabIndex        =   43
         Top             =   2740
         Width           =   3375
      End
      Begin VB.Label lblAffAgreementsDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create, modify, and terminate agreements"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   4560
         TabIndex        =   42
         Top             =   2740
         Width           =   3375
      End
      Begin VB.Label lblStationsDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View and edit station information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   720
         TabIndex        =   41
         Top             =   2740
         Width           =   3375
      End
      Begin VB.Image cmcCSLogo 
         Height          =   990
         Left            =   120
         Picture         =   "AffDirectory.frx":1D4E34
         Stretch         =   -1  'True
         Top             =   90
         Width           =   7650
      End
      Begin VB.Label lacJobs 
         BackStyle       =   0  'Transparent
         Caption         =   "Jobs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12735
         TabIndex        =   27
         Top             =   210
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label imcOutline 
         Enabled         =   0   'False
         Height          =   570
         Left            =   5280
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label lacReports 
         BackStyle       =   0  'Transparent
         Caption         =   "&REPORTS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11355
         TabIndex        =   25
         Top             =   405
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lacRadar 
         BackStyle       =   0  'Transparent
         Caption         =   "RA&DAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   8355
         TabIndex        =   24
         Top             =   7815
         Width           =   3105
      End
      Begin VB.Label lacEMail 
         BackStyle       =   0  'Transparent
         Caption         =   "Emails by &Vehicle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   8355
         TabIndex        =   23
         Top             =   2100
         Width           =   3105
      End
      Begin VB.Label lacOptions 
         BackStyle       =   0  'Transparent
         Caption         =   "&User Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4515
         TabIndex        =   22
         Top             =   7815
         Width           =   3105
      End
      Begin VB.Label lacSite 
         BackStyle       =   0  'Transparent
         Caption         =   "S&ite Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   720
         TabIndex        =   21
         Top             =   7815
         Width           =   3105
      End
      Begin VB.Label lacCPReturns 
         BackStyle       =   0  'Transparent
         Caption         =   "Affiliate Affi&davits"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4515
         TabIndex        =   20
         Top             =   3900
         Width           =   3105
      End
      Begin VB.Label lacExports 
         BackStyle       =   0  'Transparent
         Caption         =   "&Export Center"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4515
         TabIndex        =   19
         Top             =   5910
         Width           =   2985
      End
      Begin VB.Label lacPostLog 
         BackStyle       =   0  'Transparent
         Caption         =   "Network L&og"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   720
         TabIndex        =   18
         Top             =   3900
         Width           =   3105
      End
      Begin VB.Label lacAgreements 
         BackStyle       =   0  'Transparent
         Caption         =   "Affiliate &Agreements"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4560
         TabIndex        =   17
         Top             =   2100
         Width           =   3105
      End
      Begin VB.Label lacStation 
         BackStyle       =   0  'Transparent
         Caption         =   "&Stations"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   720
         TabIndex        =   16
         Top             =   2100
         Width           =   3105
      End
      Begin VB.Label lacPostBuy 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Post-&Buy Planning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   8355
         TabIndex        =   15
         Top             =   3900
         Width           =   3105
      End
      Begin VB.Label lacManagement 
         BackStyle       =   0  'Transparent
         Caption         =   "Affiliate &Management"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   720
         TabIndex        =   14
         Top             =   5910
         Width           =   3105
      End
   End
   Begin VB.PictureBox pbcMonitor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   12300
      ScaleHeight     =   210
      ScaleWidth      =   330
      TabIndex        =   12
      Top             =   7095
      Visible         =   0   'False
      Width           =   360
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   435
      Left            =   12600
      TabIndex        =   11
      Top             =   4620
      Visible         =   0   'False
      Width           =   465
      ExtentX         =   820
      ExtentY         =   767
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton cmdAffAE 
      Caption         =   "Affiliate A/E"
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
      Left            =   11625
      TabIndex        =   1
      Top             =   4155
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox pbcRedAlert 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   11925
      Picture         =   "AffDirectory.frx":5C92F2
      ScaleHeight     =   270
      ScaleWidth      =   945
      TabIndex        =   10
      Top             =   3270
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox pbcWhiteAlert 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   12195
      Picture         =   "AffDirectory.frx":5C9774
      ScaleHeight     =   270
      ScaleWidth      =   945
      TabIndex        =   9
      Top             =   3045
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "Log"
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
      Left            =   11910
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdPreLog 
      Caption         =   "Pre-Log"
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
      Left            =   11790
      TabIndex        =   2
      Top             =   1125
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox lbcLookup2 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "AffDirectory.frx":5C9BF6
      Left            =   11520
      List            =   "AffDirectory.frx":5C9BF8
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.ListBox lbcLookup1 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "AffDirectory.frx":5C9BFA
      Left            =   12855
      List            =   "AffDirectory.frx":5C9BFC
      TabIndex        =   7
      Top             =   6705
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.ListBox lbcStationInfo 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "AffDirectory.frx":5C9BFE
      Left            =   11625
      List            =   "AffDirectory.frx":5C9C05
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   6780
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdAffTimes 
      Caption         =   "Affiliate Times"
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
      Height          =   375
      Left            =   11730
      TabIndex        =   5
      Top             =   3570
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdCP 
      Caption         =   "C.P."
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
      Left            =   11970
      TabIndex        =   4
      Top             =   1485
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdContact 
      Caption         =   "Contact"
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
      Left            =   11820
      TabIndex        =   0
      Top             =   1905
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image imcExport 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   12480
      Picture         =   "AffDirectory.frx":5C9C19
      Top             =   5160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lacToolTip 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12480
      TabIndex        =   28
      ToolTipText     =   "Disallowed"
      Top             =   7440
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   12930
      Picture         =   "AffDirectory.frx":5CA4E3
      Top             =   6120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcTrashClosed 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   11745
      Picture         =   "AffDirectory.frx":5CA7ED
      Top             =   5850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcTrashOpened 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   12465
      Picture         =   "AffDirectory.frx":5CAAF7
      Top             =   6135
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   11970
      Picture         =   "AffDirectory.frx":5CAE01
      Top             =   4755
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcPrinter 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   11925
      Picture         =   "AffDirectory.frx":5CB6CB
      Top             =   5385
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmDirectory - opening box with buttons to direct user to program features
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private rst_Eqt As ADODB.Recordset


Public Sub cmcManagement_Click()
    imcOutline.Visible = False
    sgStationSearchCallSource = "M"
    If bgManagementVisible Then
        If frmStationSearch.WindowState = vbMinimized Then
            frmStationSearch.WindowState = vbNormal
        End If
        frmStationSearch.SetFocus
    Else
        frmStationSearch.Show
    End If
End Sub

Private Sub cmcManagement_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMoveOutline lacManagement
    If cmcManagement.Enabled Then
        If cmcManagement.FontBold = False Then
            cmcManagement.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmcManagement.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
        Exit Sub
    End If
    cmcManagement.ToolTipText = ""
End Sub

Public Sub cmcPostBuy_Click()
    imcOutline.Visible = False
    sgStationSearchCallSource = "P"
    If bgPostBuyVisible Then
        If frmStationSearch.WindowState = vbMinimized Then
            frmStationSearch.WindowState = vbNormal
        End If
        frmStationSearch.SetFocus
    Else
        frmStationSearch.Show
    End If
End Sub

Private Sub cmcPostBuy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMoveOutline lacPostBuy
    If cmcPostBuy.Enabled Then
        If cmcPostBuy.FontBold = False Then
            cmcPostBuy.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmcPostBuy.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
        Exit Sub
    End If
    cmcPostBuy.ToolTipText = ""
End Sub

Private Sub cmdAffAE_Click()
    frmAffRep.Show
End Sub

Private Sub cmdAffAE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAffAE.Enabled Then
        If cmdAffAE.FontBold = False Then
            cmdAffAE.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmdAffAE.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
        Exit Sub
    End If
    cmdAffAE.ToolTipText = ""
End Sub

Private Sub cmdAffTimes_Click()
    'Using CP Return instead
    'frmPosting.Show
End Sub

Public Sub cmdAgreements_Click()
    imcOutline.Visible = False
    If bgAgreementVisible Then
        If frmAgmnt.WindowState = vbMinimized Then
            frmAgmnt.WindowState = vbNormal
        End If
        frmAgmnt.SetFocus
    Else
        sgAgreementCallSource = "D"
        igTCShttCode = 0
        sgTCCallLetters = ""
        lgTCAttCode = 0
        'gRemoteTestForNewWebPW
        frmAgmnt.Show
        'frmAgmnt!cboPSSort.SetFocus
    End If
End Sub

Private Sub cmdAgreements_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMoveOutline lacAgreements
    If cmdAgreements.Enabled Then
        If cmdAgreements.FontBold = False Then
            cmdAgreements.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmdAgreements.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
        Exit Sub
    End If
    cmdAgreements.ToolTipText = ""
End Sub

Private Sub cmdContact_Click()
    frmContact.Show
    'frmContact!cboWeeks.SetFocus
End Sub

Private Sub cmdContact_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdContact.Enabled Then
        If cmdContact.FontBold = False Then
            cmdContact.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmdContact.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
        Exit Sub
    End If
    cmdContact.ToolTipText = ""
End Sub

Private Sub cmdCP_Click()
    igCPOrLog = 0
    frmCP.Show
End Sub

Private Sub cmdCP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdCP.Enabled Then
        If cmdCP.FontBold = False Then
            cmdCP.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmdCP.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
        Exit Sub
    End If
    cmdCP.ToolTipText = ""
End Sub

Public Sub cmdCPReturns_Click()
    imcOutline.Visible = False
    If bgAffidavitVisible Then
        If frmCPReturns.WindowState = vbMinimized Then
            frmCPReturns.WindowState = vbNormal
        End If
        frmCPReturns.SetFocus
    Else
        frmCPReturns.Show
    End If
End Sub

Private Sub cmdCPReturns_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMoveOutline lacCPReturns
    If cmdCPReturns.Enabled Then
        If cmdCPReturns.FontBold = False Then
            cmdCPReturns.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmdCPReturns.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
        Exit Sub
    End If
    cmdCPReturns.ToolTipText = ""
End Sub

Public Sub cmdEMail_Click()
    imcOutline.Visible = False
    If bgEMailVisible Then
        If frmWebEMail.WindowState = vbMinimized Then
            frmWebEMail.WindowState = vbNormal
        End If
        frmWebEMail.SetFocus
    Else
        frmWebEMail.Show
    End If
End Sub

Private Sub cmdEMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMoveOutline lacEMail
    If cmdEMail.Enabled Then
        If cmdEMail.FontBold = False Then
            cmdEMail.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmdEMail.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
        Exit Sub
    End If
    cmdEMail.ToolTipText = ""
End Sub

Public Sub cmdExports_Click()
    imcOutline.Visible = False
    mGetServiceInfo
    Sleep 1000
    If bgExportVisible = True Then
        If frmExport.WindowState = vbMinimized Then
            frmExport.WindowState = vbNormal
        End If
        frmExport.SetFocus
    Else
        frmExport.Show
    End If
End Sub

Private Sub cmdExports_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMoveOutline lacExports
    If cmdExports.Enabled Then
        If cmdExports.FontBold = False Then
            cmdExports.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmdExports.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
        Exit Sub
    End If
    cmdExports.ToolTipText = ""
End Sub

Private Sub cmdLog_Click()
    igCPOrLog = 1
    frmCP.Show
End Sub

Private Sub cmdLog_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdLog.Enabled Then
        If cmdLog.FontBold = False Then
            cmdLog.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmdLog.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
        Exit Sub
    End If
    cmdLog.ToolTipText = ""
End Sub

Private Sub cmdOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMoveOutline lacOptions
    If cmdOptions.Enabled Then
        If cmdOptions.FontBold = False Then
            cmdOptions.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmdOptions.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
        Exit Sub
    End If
    cmdOptions.ToolTipText = ""
End Sub

Public Sub cmdPostLog_Click()
    imcOutline.Visible = False
    igPreOrPost = 1
    If bgLogVisible Then
        If frmPostLog.WindowState = vbMinimized Then
            frmPostLog.WindowState = vbNormal
        End If
        frmPostLog.SetFocus
    Else
        frmPostLog.Show
    End If
End Sub

Private Sub cmdPostLog_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMoveOutline lacPostLog
    If cmdPostLog.Enabled Then
        If cmdPostLog.FontBold = False Then
            cmdPostLog.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmdPostLog.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
        Exit Sub
    End If
    cmdPostLog.ToolTipText = ""
End Sub

Private Sub cmdPreLog_Click()
    igPreOrPost = 0
    frmPostLog.Show
End Sub

Private Sub cmdPreLog_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdPreLog.Enabled Then
        If cmdPreLog.FontBold = False Then
            cmdPreLog.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmdPreLog.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
        Exit Sub
    End If
    cmdPreLog.ToolTipText = ""
End Sub

Public Sub cmdRadar_Click()
    imcOutline.Visible = False
    If bgRadarVisible Then
        If frmRadarProgSchd.WindowState = vbMinimized Then
            frmRadarProgSchd.WindowState = vbNormal
        End If
        frmRadarProgSchd.SetFocus
    Else
        frmRadarProgSchd.Show
    End If
End Sub

Private Sub cmdRadar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMoveOutline lacRadar
    If cmdRadar.Enabled Then
        If cmdRadar.FontBold = False Then
            cmdRadar.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmdRadar.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
        Exit Sub
    End If
    cmdRadar.ToolTipText = ""
End Sub

Private Sub cmdReports_Click()
    imcOutline.Visible = False
    igReportSource = 1
    frmReports.Show
End Sub

Private Sub cmdReports_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMoveOutline lacReports
End Sub

Public Sub cmdSite_Click()
    imcOutline.Visible = False
    On Error GoTo err1
    If bgSiteVisible Then
        If frmSiteOptions.WindowState = vbMinimized Then
            frmSiteOptions.WindowState = vbNormal
        End If
        frmSiteOptions.SetFocus
    Else
        igPasswordOk = False
        CSPWord.Show vbModal
        frmSiteOptions.Show
    End If
    If (gUsingUnivision = False) And (gUsingWeb = False) Then
        frmMain.mnuExportISCI.Enabled = False
    Else
        frmMain.mnuExportISCI.Enabled = True
    End If
err1:
End Sub

Private Sub cmdSite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMoveOutline lacSite
    If cmdSite.Enabled Then
        If cmdSite.FontBold = False Then
            cmdSite.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmdSite.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
    End If
End Sub

Public Sub cmdStation_Click()
    imcOutline.Visible = False
    If bgStationVisible Then
        If frmStation.WindowState = vbMinimized Then
            frmStation.WindowState = vbNormal
        End If
        frmStation.SetFocus
    Else
        sgStationCallSource = "D"
        igTCShttCode = 0
        sgTCCallLetters = ""
        'Get all of the latest passwords and email addresses from the web
        gRemoteTestForNewEmail
        gRemoteTestForNewWebPW
        frmStation.Show
        'frmStation!cboStations.SetFocus
    End If
End Sub

Public Sub cmdOptions_Click()
    imcOutline.Visible = False
    If bgUserVisible Then
        If frmOptions.WindowState = vbMinimized Then
            frmOptions.WindowState = vbNormal
        End If
        frmOptions.SetFocus
    Else
        frmOptions.Show
    End If
End Sub


Private Sub cmdStation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMoveOutline lacStation
    If cmdStation.Enabled Then
        If cmdStation.FontBold = False Then
            cmdStation.ToolTipText = "View Only"
            imcOutline.BackColor = vbYellow
        Else
            cmdStation.ToolTipText = "Updates Allowed"
            imcOutline.BackColor = vbGreen
        End If
        Exit Sub
    End If
    cmdStation.ToolTipText = ""
End Sub

Private Sub Form_Activate()
    Dim ilPos As Integer
    'If (Len(sgSpecialPassword) = 4) Then
    '    cmdStationSearch.Visible = True
    'End If
    If sgUstWin(1) = "I" Then
        cmdStation.Enabled = True
        lacStation.Enabled = True
        cmdStation.ToolTipText = "Updates Allowed"
        frmMain.cmcStation.Enabled = True
        frmMain.cmcStation.FontBold = True
    ElseIf sgUstWin(1) = "V" Then
        cmdStation.Enabled = True
        cmdStation.FontBold = False
        lacStation.Enabled = True
        lacStation.FontBold = False
        cmdStation.ToolTipText = "View Only"
        frmMain.cmcStation.Enabled = True
        frmMain.cmcStation.FontBold = False
    Else
        cmdStation.Enabled = False
        lacStation.Enabled = False
        lacStation.FontBold = False
        cmdStation.ToolTipText = "Disallowed"
        frmMain.cmcStation.Enabled = False
        frmMain.cmcStation.FontBold = False
    End If
    If sgUstWin(2) = "I" Then
        cmdAgreements.Enabled = True
        lacAgreements.Enabled = True
        cmdAgreements.ToolTipText = "Updates Allowed"
        frmMain.cmcAgreement.Enabled = True
        frmMain.cmcAgreement.FontBold = True
    ElseIf sgUstWin(2) = "V" Then
        cmdAgreements.Enabled = True
        cmdAgreements.FontBold = False
        lacAgreements.Enabled = True
        lacAgreements.FontBold = False
        cmdAgreements.ToolTipText = "View Only"
        frmMain.cmcAgreement.Enabled = True
        frmMain.cmcAgreement.FontBold = False
    Else
        cmdAgreements.Enabled = False
        lacAgreements.Enabled = False
        lacAgreements.FontBold = False
        cmdAgreements.ToolTipText = "Disallowed"
        frmMain.cmcAgreement.Enabled = False
        frmMain.cmcAgreement.FontBold = False
    End If
    'If sgUstWin(3) = "I" Then
    '    cmdPreLog.Enabled = True
    'ElseIf sgUstWin(3) = "V" Then
    '    cmdPreLog.Enabled = True
    '    cmdPreLog.FontBold = False
    'Else
    '    cmdPreLog.Enabled = False
    'End If
    cmdPreLog.Visible = False
    'If sgUstWin(4) = "I" Then
    '    cmdLog.Enabled = True
    'ElseIf sgUstWin(4) = "V" Then
    '    cmdLog.Enabled = True
    '    cmdLog.FontBold = False
    'Else
    '    cmdLog.Enabled = False
    'End If
    cmdLog.Visible = False
    If sgUstWin(5) = "I" Then
        cmdPostLog.Enabled = True
        lacPostLog.Enabled = True
        cmdPostLog.ToolTipText = "Updates Allowed"
        frmMain.cmcLog.Enabled = True
        frmMain.cmcLog.FontBold = True
    ElseIf sgUstWin(5) = "V" Then
        cmdPostLog.Enabled = True
        cmdPostLog.FontBold = False
        lacPostLog.Enabled = True
        lacPostLog.FontBold = False
        cmdPostLog.ToolTipText = "View Only"
        frmMain.cmcLog.Enabled = True
        frmMain.cmcLog.FontBold = False
    Else
        cmdPostLog.Enabled = False
        lacPostLog.Enabled = False
        lacPostLog.FontBold = False
        cmdPostLog.ToolTipText = "Disallowed"
        frmMain.cmcLog.Enabled = False
        frmMain.cmcLog.FontBold = False
    End If
    'If sgUstWin(6) = "I" Then
    '    cmdCP.Enabled = True
    'ElseIf sgUstWin(6) = "V" Then
    '    cmdCP.Enabled = True
    '    cmdCP.FontBold = False
    'Else
    '    cmdCP.Enabled = False
    'End If
    cmdCP.Visible = False
    If sgUstWin(7) = "I" Then
        cmdCPReturns.Enabled = True
        lacCPReturns.Enabled = True
        cmdCPReturns.ToolTipText = "Updates Allowed"
        frmMain.cmcAffIdavit.Enabled = True
        frmMain.cmcAffIdavit.FontBold = True
    ElseIf sgUstWin(7) = "V" Then
        cmdCPReturns.Enabled = True
        cmdCPReturns.FontBold = False
        lacCPReturns.Enabled = True
        lacCPReturns.FontBold = False
        cmdCPReturns.ToolTipText = "View Only"
        frmMain.cmcAffIdavit.Enabled = True
        frmMain.cmcAffIdavit.FontBold = False
    Else
        cmdCPReturns.Enabled = False
        lacCPReturns.Enabled = False
        lacCPReturns.FontBold = False
        cmdCPReturns.ToolTipText = "Disallowed"
        frmMain.cmcAffIdavit.Enabled = False
        frmMain.cmcAffIdavit.FontBold = False
    End If
    'If sgUstWin(8) = "I" Then
    '    cmdContact.Enabled = True
    'ElseIf sgUstWin(8) = "V" Then
    '    cmdContact.Enabled = True
    '    cmdContact.FontBold = False
    'Else
    '    cmdContact.Enabled = False
    'End If
    cmdContact.Visible = False
    If sgUstWin(9) = "I" Then
        cmdOptions.Enabled = True
        lacOptions.Enabled = True
        cmdOptions.ToolTipText = "Updates Allowed"
        frmMain.cmcUser.Enabled = True
        frmMain.cmcUser.FontBold = True
    ElseIf sgUstWin(9) = "V" Then
        cmdOptions.Enabled = True
        cmdOptions.FontBold = False
        lacOptions.Enabled = True
        lacOptions.FontBold = False
        cmdOptions.ToolTipText = "View Only"
        frmMain.cmcUser.Enabled = True
        frmMain.cmcUser.FontBold = False
    Else
        cmdOptions.Enabled = False
        lacOptions.Enabled = False
        lacOptions.FontBold = False
        cmdOptions.ToolTipText = "Disallowed"
        frmMain.cmcUser.Enabled = False
        frmMain.cmcUser.FontBold = False
    End If
    If sgUstWin(10) = "I" Then
        cmdSite.Enabled = True
        lacSite.Enabled = True
        cmdSite.ToolTipText = "Updates Allowed"
        frmMain.cmcSite.Enabled = True
        frmMain.cmcSite.FontBold = True
    ElseIf sgUstWin(10) = "V" Then
        cmdSite.Enabled = True
        cmdSite.FontBold = False
        lacSite.Enabled = True
        lacSite.FontBold = False
        cmdSite.ToolTipText = "View Only"
        frmMain.cmcSite.Enabled = True
        frmMain.cmcSite.FontBold = False
    Else
        cmdSite.Enabled = False
        lacSite.Enabled = False
        lacSite.FontBold = False
        cmdSite.ToolTipText = "Disallowed"
        frmMain.cmcSite.Enabled = False
        frmMain.cmcSite.FontBold = False
    End If
    If (gUsingWeb = False) Then
    'E-Mail only works with Web unless we install e-mail program on a server ofter then the web server.
    'If ((Asc(sgSpfUsingFeatures5) And STATIONINTERFACE) <> STATIONINTERFACE) Then
        cmdEMail.Enabled = False
        lacEMail.Enabled = False
        lacEMail.FontBold = False
        cmdEMail.ToolTipText = "Disallowed"
        frmMain.cmcEMail.Enabled = False
        frmMain.cmcEMail.FontBold = False
    Else
        'cmdEMail.Enabled = True
        If sgUstWin(11) = "I" Then
            cmdEMail.Enabled = True
            lacEMail.Enabled = True
            cmdEMail.ToolTipText = "Updates Allowed"
            frmMain.cmcEMail.Enabled = True
            frmMain.cmcEMail.FontBold = True
        ElseIf sgUstWin(11) = "V" Then
            cmdEMail.Enabled = True
            cmdEMail.FontBold = False
            lacEMail.Enabled = True
            lacEMail.FontBold = False
            cmdEMail.ToolTipText = "View Only"
            frmMain.cmcEMail.Enabled = True
            frmMain.cmcEMail.FontBold = False
        Else
            cmdEMail.Enabled = False
            lacEMail.Enabled = False
            lacEMail.FontBold = False
            cmdEMail.ToolTipText = "Disallowed"
            frmMain.cmcEMail.Enabled = False
            frmMain.cmcEMail.FontBold = False
        End If
    End If
    If ((Asc(sgSpfUsingFeatures5) And RADAR) <> RADAR) Then
        cmdRadar.Enabled = False
        lacRadar.Enabled = False
        lacRadar.FontBold = False
        cmdRadar.ToolTipText = "Disallowed"
        ' Dan M flag for user options directory screen
        bgNoRADAR = True
        frmMain.cmcRadar.Enabled = False
        frmMain.cmcRadar.FontBold = False
    Else
        If sgUstWin(12) = "I" Then
            cmdRadar.Enabled = True
            lacRadar.Enabled = True
            cmdRadar.ToolTipText = "Updates Allowed"
            frmMain.cmcRadar.Enabled = True
            frmMain.cmcRadar.FontBold = True
        ElseIf sgUstWin(12) = "V" Then
            cmdRadar.Enabled = True
            cmdRadar.FontBold = False
            lacRadar.Enabled = True
            lacRadar.FontBold = False
            cmdRadar.ToolTipText = "View Only"
            frmMain.cmcRadar.Enabled = True
            frmMain.cmcRadar.FontBold = False
        Else
            cmdRadar.Enabled = False
            lacRadar.Enabled = False
            lacRadar.FontBold = False
            cmdRadar.ToolTipText = "Disallowed"
            frmMain.cmcRadar.Enabled = False
            frmMain.cmcRadar.FontBold = False
        End If
    End If
    'Jim:1/21/11- Gray out Affiliate A/E
    'If sgUstWin(13) = "I" Then
    '    cmdAffAE.Enabled = True
    'ElseIf sgUstWin(13) = "V" Then
    '    cmdAffAE.Enabled = True
    'Else
        cmdAffAE.Enabled = False
    'End If
    cmdAffAE.Visible = False
    'If ((Asc(sgSpfUsingFeatures9) And AFFILIATECRM) <> AFFILIATECRM) Then
    '    cmdStationSearch.Enabled = False
    'Else
        If sgUstWin(0) = "I" Then
            cmcManagement.Enabled = True
            lacManagement.Enabled = True
            cmcManagement.ToolTipText = "Updates Allowed"
            frmMain.cmcManagement.Enabled = True
            frmMain.cmcManagement.FontBold = True
        ElseIf sgUstWin(0) = "V" Then
            cmcManagement.Enabled = True
            cmcManagement.FontBold = False
            lacManagement.Enabled = True
            lacManagement.FontBold = False
            cmcManagement.ToolTipText = "View Only"
            frmMain.cmcManagement.Enabled = True
            frmMain.cmcManagement.FontBold = False
        Else
            cmcManagement.Enabled = False
            lacManagement.Enabled = False
            lacManagement.FontBold = False
            cmcManagement.ToolTipText = "Disallowed"
            frmMain.cmcManagement.Enabled = False
            frmMain.cmcManagement.FontBold = False
        End If
    'End If
    If sgUstWin(13) = "I" Then
        cmcPostBuy.Enabled = True
        lacPostBuy.Enabled = True
        cmcPostBuy.ToolTipText = "Updates Allowed"
        frmMain.cmcPostBuy.Enabled = True
        frmMain.cmcPostBuy.FontBold = True
    ElseIf sgUstWin(13) = "V" Then
        cmcPostBuy.Enabled = True
        cmcPostBuy.FontBold = False
        lacPostBuy.Enabled = True
        lacPostBuy.FontBold = False
        cmcPostBuy.ToolTipText = "View Only"
        frmMain.cmcPostBuy.Enabled = True
        frmMain.cmcPostBuy.FontBold = False
    Else
        cmcPostBuy.Enabled = False
        lacPostBuy.Enabled = False
        lacPostBuy.FontBold = False
        cmcPostBuy.ToolTipText = "Disallowed"
        frmMain.cmcPostBuy.Enabled = False
        frmMain.cmcPostBuy.FontBold = False
    End If
    
    ''6/30/14: Allow Reports
    ''If igGGFlag = 0 Then
    ''    cmdReports.Enabled = False
    ''End If
    'ilPos = InStr(1, sgCommand, "/Reports", vbTextCompare)
    'If (igGGFlag = 0) And (ilPos = 0) Then
    '    cmdReports.Enabled = False
    '    lacReports.Enabled = False
    '    cmdReports.ToolTipText = "Disallowed"
    'End If
    If (igGGFlag = 0) And (igRptGGFlag = 0) Then
        cmdReports.Enabled = False
        lacReports.Enabled = False
        lacReports.FontBold = False
        cmdReports.ToolTipText = "Disallowed"
        imcOutline.BackColor = vbRed
        frmMain.cmcReport.Enabled = False
        frmMain.cmcReport.FontBold = False
    End If
    
    'Dan 6/02/10 user and site buttons enabled if limited guide
    mLimitedGuideButtons
    
    '6/30/14: Disallow any activity except Reports
    If (StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) <> 0) And (StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) <> 0) Then
        If igGGFlag = 0 Then
            cmcManagement.Enabled = False
            cmcPostBuy.Enabled = False
            cmdAffAE.Enabled = False
            cmdAffTimes.Enabled = False
            cmdAgreements.Enabled = False
            cmdContact.Enabled = False
            cmdCP.Enabled = False
            cmdCPReturns.Enabled = False
            cmdEMail.Enabled = False
            cmdExports.Enabled = False
            cmdExports.Enabled = False
            cmdLog.Enabled = False
            cmdOptions.Enabled = False
            cmdPostLog.Enabled = False
            cmdPreLog.Enabled = False
            cmdRadar.Enabled = False
            cmdSite.Enabled = False
            cmdStation.Enabled = False
        
            lacManagement.Enabled = False
            lacPostBuy.Enabled = False
            lacAgreements.Enabled = False
            lacCPReturns.Enabled = False
            lacEMail.Enabled = False
            lacExports.Enabled = False
            lacOptions.Enabled = False
            lacPostLog.Enabled = False
            lacRadar.Enabled = False
            lacSite.Enabled = False
            lacStation.Enabled = False
        
            lacManagement.FontBold = False
            lacPostBuy.FontBold = False
            lacAgreements.FontBold = False
            lacCPReturns.FontBold = False
            lacEMail.FontBold = False
            lacExports.FontBold = False
            lacOptions.FontBold = False
            lacPostLog.FontBold = False
            lacRadar.FontBold = False
            lacSite.FontBold = False
            lacStation.FontBold = False
            
            frmMain.cmcStation.Enabled = False
            frmMain.cmcStation.FontBold = False
            frmMain.cmcAgreement.Enabled = False
            frmMain.cmcAgreement.FontBold = False
            frmMain.cmcLog.Enabled = False
            frmMain.cmcLog.FontBold = False
            frmMain.cmcAffIdavit.Enabled = False
            frmMain.cmcAffIdavit.FontBold = False
            frmMain.cmcUser.Enabled = False
            frmMain.cmcUser.FontBold = False
            frmMain.cmcSite.Enabled = False
            frmMain.cmcSite.FontBold = False
            frmMain.cmcEMail.Enabled = False
            frmMain.cmcEMail.FontBold = False
            frmMain.cmcRadar.Enabled = False
            frmMain.cmcRadar.FontBold = False
            frmMain.cmcManagement.Enabled = False
            frmMain.cmcManagement.FontBold = False
            frmMain.cmcPostBuy.Enabled = False
            frmMain.cmcPostBuy.FontBold = False
            
        End If
    End If

    'frmDirectory.Refresh
    
    If igEmailNeedsConv Then
        'frmEmailConv.Show vbModal
        igEmailNeedsConv = False
    End If
    
    'If (Not igDemoMode) And (Len(sgSpecialPassword) <> 4) Then
    '    cmdExports.Visible = False
    'Else
        'cmdPostLog.Width = cmdEMail.Width
        'cmdCPReturns.Width = cmdReports.Width
        'cmdCPReturns.Left = cmdReports.Left
        If sgUstWin(14) = "I" Then
            cmdExports.Enabled = True
            lacExports.Enabled = True
            cmdExports.ToolTipText = "Updates Allowed"
            frmMain.cmcExport.Enabled = True
            frmMain.cmcExport.FontBold = True
        ElseIf sgUstWin(14) = "V" Then
            cmdExports.Enabled = True
            cmdExports.FontBold = False
            lacExports.Enabled = True
            lacExports.FontBold = False
            cmdExports.ToolTipText = "View Only"
            frmMain.cmcExport.Enabled = True
            frmMain.cmcExport.FontBold = False
        Else
            cmdExports.Enabled = False
            lacExports.Enabled = False
            lacExports.FontBold = False
            cmdExports.ToolTipText = "Disallowed"
            frmMain.cmcExport.Enabled = False
            frmMain.cmcExport.FontBold = False
        End If
        
        '6/30/14: Disallow
        If (StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) <> 0) And (StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) <> 0) Then
            If igGGFlag = 0 Then
                cmdExports.Enabled = False
                lacExports.Enabled = False
                lacExports.FontBold = False
                cmdExports.ToolTipText = "Disallowed"
                frmMain.cmcExport.Enabled = False
                frmMain.cmcExport.FontBold = False
            End If
        End If
    'End If
    
End Sub

Private Sub Form_Click()
    imcOutline.Visible = False
End Sub

Private Sub Form_Load()
    lacAffiliationsGroup.ForeColor = RGB(131, 60, 11)
    lacComplianceGroup.ForeColor = RGB(97, 144, 160)
    lacManagementGroup.ForeColor = RGB(125, 171, 130)
    lacSetupGroup.ForeColor = RGB(172, 146, 160)
    
    If App.PrevInstance Then
        If Not igCompelAutoImport Then
            MsgBox "Only one copy of Affiliate can be run at a time, sorry", vbOKOnly + vbInformation, "Counterpoint"
        Else
            gLogMsg "Only one copy of Affiliate can be run at a time, sorry", "WegenerImportResult.Txt", False
        End If
    End If

   ScaleMode = 3       'set screen mode to pixels
   'Me.Width = Screen.Width / 1.5
   'Me.Height = Screen.Height / 2.1
   'Me.Top = (Screen.Height - Me.Height) / 2.2
   'Me.Left = (Screen.Width - Me.Width) / 2
   'frmDirectory.Refresh
   'Me.Width = 9870
   'Me.Height = 6405
    If Screen.Height <= 720 * 15 Then
        Me.Top = (Screen.Height - Me.Height) / 3.3
        Me.Left = (Screen.Width - Me.Width) / 2.3
    Else
        Me.Top = (Screen.Height - Me.Height) / 2.2
        Me.Left = (Screen.Width - Me.Width) / 2
    End If
    If bgLimitedGuide Then
         mLimitGuide
    End If
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Visible = False
    mDisallowToolTip X, Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst_Eqt.Close
    
    Set frmDirectory = Nothing
End Sub
Private Sub mLimitedGuideButtons()
    
    If bgLimitedGuide Then
        cmdOptions.Enabled = True
        cmdSite.Enabled = True
    End If
End Sub
Private Sub mLimitGuide()
Dim control As control
Dim slVisibleMenuItems(6) As String
Dim c As Integer
Dim blSkip As Boolean

    cmdReports.Enabled = False
    'frmMain.mnuAccessories.Visible = False
    'frmMain.mnuFile.Visible = False
    mSetVisibleMenuList slVisibleMenuItems
    For Each control In frmMain.Controls
        If TypeOf control Is Menu Then
            blSkip = False
        On Error Resume Next
            For c = 0 To UBound(slVisibleMenuItems)
                If control.Name = slVisibleMenuItems(c) Then
                    blSkip = True
                    Exit For
                End If
            Next c
            If Not blSkip Then
                control.Visible = False
            End If
        End If
     Next
End Sub
Private Sub mSetVisibleMenuList(ByRef slVisibleMenuItems() As String)
slVisibleMenuItems(0) = "mnuFile"
slVisibleMenuItems(1) = "mnuFileExit"
slVisibleMenuItems(2) = "mnuHelp"
slVisibleMenuItems(3) = "mnuHelpSearch"
slVisibleMenuItems(4) = "mnuHelpAbout"
slVisibleMenuItems(5) = "mnuHelpBar1"
slVisibleMenuItems(6) = "mnuHelpContents"
    
    
    
End Sub
Private Sub mGetServiceInfo()
    Dim ilRet As Integer
    Dim llServiceDate As Long
    Dim llServiceTime As Long
    
    On Error GoTo ErrHand
    SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE eqtType = 'T'"
    Set rst_Eqt = gSQLSelectCall(SQLQuery)
    If Not rst_Eqt.EOF Then
        lgLastServiceDate = gDateValue(Format(rst_Eqt!eqtDateEntered, sgShowDateForm))
        lgLastServiceTime = gTimeToLong(Format(rst_Eqt!eqtTimeEntered, sgShowTimeWSecForm), True)
    Else
        lgLastServiceDate = gDateValue("1/1/1970")
    End If
    Exit Sub
ErrHand:
    lgLastServiceDate = gDateValue("1/1/1970")
    Exit Sub
End Sub

Private Sub pbcDirectory_Click()
    imcOutline.Visible = False
End Sub

Private Sub pbcDirectory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Visible = False
End Sub

Private Sub mMoveOutline(lacCtrl As Label)
    imcOutline.Move lacCtrl.Left - 90, lacCtrl.Top - 90, lacCtrl.Width + 180, lacCtrl.Height + 90
    imcOutline.Visible = True
End Sub

Private Sub mDisallowToolTip(X As Single, Y As Single)
    Dim Ctrl As control
    Dim ilLoop As Integer
    
    For Each Ctrl In frmDirectory.Controls
        If TypeOf Ctrl Is CommandButton Then
            If (Not Ctrl.Enabled) And (Ctrl.Visible) Then
                If (X >= Ctrl.Left) And (X <= Ctrl.Left + Ctrl.Width) Then
                    If (Y >= Ctrl.Top) And (Y <= Ctrl.Top + Ctrl.Height) Then
                        lacToolTip.Move Ctrl.Left, Ctrl.Top, Ctrl.Width
                        lacToolTip.Visible = True
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next Ctrl
    lacToolTip.Visible = False
End Sub
