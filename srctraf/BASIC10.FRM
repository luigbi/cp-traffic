VERSION 5.00
Begin VB.Form Basic10 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   135
   ClientWidth     =   13500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   LinkMode        =   1  'Source
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9000
   ScaleWidth      =   13500
   Begin VB.PictureBox plcBasic10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   9000
      Left            =   0
      Picture         =   "BASIC10.frx":0000
      ScaleHeight     =   9000
      ScaleWidth      =   13515
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13520
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Height          =   690
         Index           =   4
         Left            =   3870
         Picture         =   "BASIC10.frx":18B862
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "Move, Exchange, Miss, Reschedule, Close or Freeze spots"
         Top             =   4620
         Width           =   2565
      End
      Begin VB.Timer tmcClick 
         Enabled         =   0   'False
         Interval        =   400
         Left            =   0
         Top             =   7320
      End
      Begin VB.Timer tmcFilemakerToOrder 
         Enabled         =   0   'False
         Interval        =   400
         Left            =   0
         Top             =   8280
      End
      Begin VB.Timer tmcFilemaker 
         Enabled         =   0   'False
         Interval        =   400
         Left            =   0
         Top             =   7800
      End
      Begin VB.PictureBox pbcHor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2520
         Left            =   13560
         ScaleHeight     =   2520
         ScaleWidth      =   13500
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   6360
         Visible         =   0   'False
         Width           =   13500
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "&Budgets"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   10
            Left            =   600
            TabIndex        =   31
            Top             =   2400
            Width           =   1275
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "&Station Feed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   12
            Left            =   8220
            TabIndex        =   28
            Top             =   3390
            Visible         =   0   'False
            Width           =   2040
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "&Feed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   13
            Left            =   8235
            TabIndex        =   27
            Top             =   2025
            Width           =   810
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
            Left            =   10320
            TabIndex        =   25
            Top             =   45
            Width           =   570
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "Co&mmissions"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   11
            Left            =   6660
            TabIndex        =   21
            Top             =   4755
            Width           =   2070
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "Collectio&ns"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   9
            Left            =   4620
            TabIndex        =   19
            Top             =   4755
            Width           =   1815
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "&Invoices"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   8
            Left            =   2430
            TabIndex        =   17
            Top             =   4755
            Width           =   1335
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "&Post Log"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   7
            Left            =   645
            TabIndex        =   15
            Top             =   4755
            Width           =   1470
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "Log&s"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   6
            Left            =   6660
            TabIndex        =   13
            Top             =   3390
            Width           =   930
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "&Copy"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   5
            Left            =   4620
            TabIndex        =   11
            Top             =   3390
            Width           =   855
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "Spo&ts"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   4
            Left            =   3225
            TabIndex        =   10
            Top             =   3390
            Width           =   930
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "Progra&mming"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   645
            TabIndex        =   8
            Top             =   3390
            Width           =   2085
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "&Orders"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   6660
            TabIndex        =   6
            Top             =   2010
            Width           =   1065
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "Propos&als"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   4620
            TabIndex        =   4
            Top             =   2025
            Width           =   1545
         End
         Begin VB.Label lacTask 
            BackStyle       =   0  'Transparent
            Caption         =   "Rat&e Cards"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   2430
            TabIndex        =   2
            Top             =   2025
            Width           =   1725
         End
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "&Feed"
         Height          =   690
         Index           =   12
         Left            =   13155
         TabIndex        =   29
         Tag             =   "Add, Change, Cancel or Display contracts"
         Top             =   4485
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Height          =   690
         Index           =   11
         Left            =   9900
         Picture         =   "BASIC10.frx":190B54
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "Enter transactions for cash received"
         Top             =   7140
         Width           =   2565
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Height          =   690
         Index           =   9
         Left            =   6885
         Picture         =   "BASIC10.frx":195E46
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "Enter transactions for cash received"
         Top             =   7140
         Width           =   2565
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Height          =   690
         Index           =   8
         Left            =   3870
         Picture         =   "BASIC10.frx":19B138
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "Create invoices and update Accounts Receivables"
         Top             =   7140
         Width           =   2565
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Height          =   690
         Index           =   7
         Left            =   840
         Picture         =   "BASIC10.frx":1A042A
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "Alter spots to agree with aired log"
         Top             =   7140
         Width           =   2565
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Height          =   690
         Index           =   6
         Left            =   9900
         Picture         =   "BASIC10.frx":1A571C
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "Create the Log for a day"
         Top             =   4620
         Width           =   2565
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Height          =   690
         Index           =   5
         Left            =   6885
         Picture         =   "BASIC10.frx":1AAA0E
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "Commands intended for the Continuity department"
         Top             =   4620
         Width           =   2565
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   3
         Left            =   840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BASIC10.frx":1AFD00
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "Add, change, or delete all Program material for logs"
         Top             =   4620
         Width           =   2565
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Height          =   690
         Index           =   2
         Left            =   9900
         Picture         =   "BASIC10.frx":1B4FF2
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "Add, Change, Cancel or Display contracts"
         Top             =   2130
         Width           =   2565
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Index           =   1
         Left            =   6885
         Picture         =   "BASIC10.frx":1BC3B4
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "Create or modify an Agency Proposal"
         Top             =   2130
         Width           =   2565
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Height          =   690
         Index           =   0
         Left            =   3870
         Picture         =   "BASIC10.frx":1C16A6
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "Create new Rate cards or modify existing ones"
         Top             =   2130
         Width           =   2565
      End
      Begin VB.PictureBox pbcClickFocus 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   -195
         ScaleHeight     =   165
         ScaleWidth      =   75
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3855
         Width           =   75
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Height          =   690
         Index           =   10
         Left            =   840
         Picture         =   "BASIC10.frx":1C6998
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "Create new Rate cards or modify existing ones"
         Top             =   2130
         Width           =   2565
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "&Feed"
         Height          =   690
         Index           =   13
         Left            =   13155
         TabIndex        =   26
         Tag             =   "Add, Change, Cancel or Display contracts"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Image cmcDashboard 
         Height          =   600
         Left            =   11760
         Picture         =   "BASIC10.frx":1CBC8A
         Stretch         =   -1  'True
         ToolTipText     =   "Dashboard"
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lacSalesGroup 
         BackStyle       =   0  'Transparent
         Caption         =   "SALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   850
         TabIndex        =   47
         Top             =   1490
         Width           =   3375
      End
      Begin VB.Label lacTrafficGroup 
         BackStyle       =   0  'Transparent
         Caption         =   "TRAFFIC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   850
         TabIndex        =   46
         Top             =   3980
         Width           =   3375
      End
      Begin VB.Label lacAccountingGroup 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNTING"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   850
         TabIndex        =   45
         Top             =   6500
         Width           =   3735
      End
      Begin VB.Label lblOrdersDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create, revise, and schedule air time, podcast, and NTR orders"
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
         Height          =   735
         Left            =   9900
         TabIndex        =   44
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label lblProposalsDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add and change air time, podcast, and NTR proposals"
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
         Left            =   6885
         TabIndex        =   43
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label lblRatecardDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Define vehicles, dayparts, and rates"
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
         Left            =   3870
         TabIndex        =   42
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label lblBudgetDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create and edit budgets and actuals"
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
         Left            =   840
         TabIndex        =   41
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label lblLogsDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generate preliminary, final, and reprint logs"
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
         Left            =   9900
         TabIndex        =   40
         Top             =   5400
         Width           =   2535
      End
      Begin VB.Label lblCopyDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Manage copy inventory and create copy rotations"
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
         Left            =   6885
         TabIndex        =   39
         Top             =   5400
         Width           =   2535
      End
      Begin VB.Label lblSpotsDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View and traffic spots in future weeks"
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
         Left            =   3870
         TabIndex        =   38
         Top             =   5400
         Width           =   2535
      End
      Begin VB.Label lblProgrammingDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Define and schedule program libraries"
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
         Left            =   840
         TabIndex        =   37
         Top             =   5400
         Width           =   2535
      End
      Begin VB.Label lblPostLogDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reconcile spots in the past"
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
         Left            =   840
         TabIndex        =   36
         Top             =   7920
         Width           =   2535
      End
      Begin VB.Label lblInvoicesDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generate preliminary, final, and reprint invoices"
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
         Left            =   3870
         TabIndex        =   35
         Top             =   7920
         Width           =   2535
      End
      Begin VB.Label lblCollectionsDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Perform accounts receivable activity"
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
         Left            =   6885
         TabIndex        =   34
         Top             =   7920
         Width           =   2535
      End
      Begin VB.Label lblCommissionsDesc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Access commission reports or the sub-company commission table"
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
         Height          =   735
         Left            =   9900
         TabIndex        =   33
         Top             =   7920
         Width           =   2415
      End
      Begin VB.Label imcOutline 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   570
         Left            =   6840
         TabIndex        =   30
         Top             =   2040
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Image cmcStartList 
         Height          =   600
         Left            =   12600
         Picture         =   "BASIC10.frx":258064
         Stretch         =   -1  'True
         ToolTipText     =   "Show Lists Screen"
         Top             =   240
         Width           =   600
      End
      Begin VB.Image cmcCSLogo 
         Height          =   945
         Left            =   120
         Picture         =   "BASIC10.frx":3180A6
         Stretch         =   -1  'True
         Top             =   120
         Width           =   7650
      End
   End
End
Attribute VB_Name = "Basic10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of BASIC10.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Basic10.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Basic 10 task buttons screen code
Option Explicit
Option Compare Text

Dim imTaskIndex As Integer
Dim imPrevTaskIndex As Integer

Const JOBRATECARD = 1
Const JOBPROPOSALS = 2
Const JOBCONTRACTS = 3
Const JOBPROGRAMS = 4
Const JOBSPOTS = 5
Const JOBCOPY = 6
Const JOBLOGS = 7
Const JOBAFFIDAVITS = 8
Const JOBINVOICES = 9
Const JOBCOLLECTION = 10
Const JOBBUDGETS = 12
Const JOBSLSPCOMMS = 13
Const JOBAFFILIATE = 14
Dim tmSaf As SAF
Dim hmSaf As Integer
Dim imSafRecLen As Integer


Private Sub cmcCSLogo_Click()
    Dim ilAns As Integer
    Dim llNow As Long
    Dim slStr As String
    Dim slNow As String
    Dim ilLoop As Integer
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim llDate As Long
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilValue As Integer
    Dim slNxtValue As String
    Dim slInValue As String
    Dim slDays As String
    Dim ilPattern As Integer
    Dim ilRet As Integer
    Dim slDate As String
    Dim slRptDays As String

    If igGGFlag = -1 Then
        ilAns = MsgBox(sgCSVersion, 0, "Counterpoint Software")
    Else
        hmSaf = CBtrvTable(TWOHANDLES)
        ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet = BTRV_ERR_NONE Then
            imSafRecLen = Len(tmSaf)  'btrRecordLength(hlUrf)  'Get and save record length
            ilRet = btrGetFirst(hmSaf, tmSaf, imSafRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                ilField1 = Asc(tmSaf.sName)
                slStr = Mid$(tmSaf.sName, 2, 5)
                If Left$(slStr, 1) <> "9" Then
                    llDate = Val(slStr)
                Else
                    llDate = gDateValue(Format$(gNow(), "m/d/yy"))
                End If
                llNow = gDateValue(Format$(gNow(), "m/d/yy"))
                ilField2 = Asc(Mid$(tmSaf.sName, 11, 1))
                sgGenMsg = sgCSVersion
                If llNow < llDate Then
                    sgGenMsg = sgGenMsg & "," & str$(llDate - llNow)
                End If
                slStr = Mid$(tmSaf.sName, 7, 4)
                'If slStr = "0000" Then
                    'ilValue = Int(10000 * Rnd(-ilValue) + 1)
                    ilMin = Asc("A")
                    ilMax = Asc("Z")
                    slStr = ""
                    For ilLoop = 0 To 3 Step 1
                        Randomize
                        slStr = slStr & Trim$(Chr$(Int((ilMax - ilMin + 1) * Rnd + ilMin)))
                    Next ilLoop
                'End If
                For ilLoop = 0 To 3 Step 1
                    ilValue = ilValue + (ilLoop + 1) * Asc(Mid$(slStr, ilLoop + 1, 1))
                Next ilLoop
                slNxtValue = ""
                For ilLoop = 0 To 3 Step 1
                    ilValue = Int(9000 * Rnd(-ilValue) + 1000)
                    slNxtValue = slNxtValue & Trim$(Chr$(Asc("A") + ilValue Mod 25))
                Next ilLoop
                sgGenMsg = sgGenMsg & sgCR & sgLF & "Key Code: " & slStr
                sgCMCTitle(0) = "Ok"
                sgCMCTitle(1) = "Cancel"
                sgCMCTitle(2) = ""
                sgCMCTitle(3) = ""
                igDefCMC = 1
                igEditBox = 1
                sgEditValue = ""
                On Error Resume Next
                GenMsg.Show vbModal
                If (igAnsCMC = 0) And (Len(sgEditValue) >= 9) Then    'Ok, test value in sgEditValue
                    'decode input
                    sgEditValue = UCase(sgEditValue)
                    ilPattern = ((Asc(sgEditValue) - Asc("A")) Mod 4) + 1
                    Select Case ilPattern
                        Case 1  'v1d1v2d2v3d3v4d4
                            slInValue = Mid$(sgEditValue, 2, 1) & Mid$(sgEditValue, 4, 1) & Mid$(sgEditValue, 6, 1) & Mid$(sgEditValue, 8, 1)
                            slDays = Mid$(sgEditValue, 3, 1) & Mid$(sgEditValue, 5, 1) & Mid$(sgEditValue, 7, 1) & Mid$(sgEditValue, 9, 1)
                        Case 2  'v1d4v2d3v3d2v4d1
                            slInValue = Mid$(sgEditValue, 2, 1) & Mid$(sgEditValue, 4, 1) & Mid$(sgEditValue, 6, 1) & Mid$(sgEditValue, 8, 1)
                            slDays = Mid$(sgEditValue, 9, 1) & Mid$(sgEditValue, 7, 1) & Mid$(sgEditValue, 5, 1) & Mid$(sgEditValue, 3, 1)
                        Case 3  'd4v4d3v3d2v2d1v1
                            slInValue = Mid$(sgEditValue, 9, 1) & Mid$(sgEditValue, 7, 1) & Mid$(sgEditValue, 5, 1) & Mid$(sgEditValue, 3, 1)
                            slDays = Mid$(sgEditValue, 8, 1) & Mid$(sgEditValue, 6, 1) & Mid$(sgEditValue, 4, 1) & Mid$(sgEditValue, 2, 1)
                        Case 4  'd1v4d2v3d3v2d4v1
                            slInValue = Mid$(sgEditValue, 9, 1) & Mid$(sgEditValue, 7, 1) & Mid$(sgEditValue, 5, 1) & Mid$(sgEditValue, 3, 1)
                            slDays = Mid$(sgEditValue, 2, 1) & Mid$(sgEditValue, 4, 1) & Mid$(sgEditValue, 6, 1) & Mid$(sgEditValue, 8, 1)
                    End Select
                    slDate = Trim$(str$(Asc(Mid$(slDays, 1, 1)) - Asc("A") + 1))
                    slDate = slDate & "/" & Trim$(str$(Asc(Mid$(slDays, 2, 1)) - Asc("A"))) & Trim$(str$(Asc(Mid$(slDays, 3, 1)) - Asc("A")))
                    If Len(Trim$(str$(Asc(Mid$(slDays, 4, 1)) - Asc("A")))) = 1 Then
                        slDate = slDate & "/0" & Trim$(str$(Asc(Mid$(slDays, 4, 1)) - Asc("A")))
                    Else
                        slDate = slDate & "/" & Trim$(str$(Asc(Mid$(slDays, 4, 1)) - Asc("A")))
                    End If
                    slRptDays = ""
                    If Len(sgEditValue) > 9 Then
                        For ilLoop = 10 To Len(sgEditValue) Step 1
                            slRptDays = slRptDays & str$(Asc(Mid$(sgEditValue, ilLoop, 1)) - Asc("A"))
                        Next ilLoop
                    End If
                    If (slNxtValue = slInValue) And (gValidDate(slDate)) Then
                        llNow = gDateValue(Format$(gNow(), "m/d/yy"))
                        If llNow + 730 < gDateValue(slDate) Then
                            slStr = Chr(1)
                            'slStr = slStr & "99999" & slInValue
                            slStr = slStr & "9"
                            For ilLoop = 0 To 3 Step 1
                                slStr = slStr & Trim$(str$(Int((9 - 1 + 1) * Rnd + 1)))
                            Next ilLoop
                            slStr = slStr & slInValue
                            slStr = slStr & Chr(1)
                            slStr = slStr & "9"
                            For ilLoop = 0 To 3 Step 1
                                slStr = slStr & Trim$(str$(Int((9 - 1 + 1) * Rnd + 1)))
                            Next ilLoop
                            tmSaf.sName = slStr & slRptDays
                            ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
                        ElseIf llNow < gDateValue(slDate) Then
                            slStr = Chr(0)
                            llNow = gDateValue(slDate)
                            slNow = Trim$(str$(llNow))
                            Do While Len(slNow) < 5
                                slNow = "0" & slNow
                            Loop
                            slStr = slStr & slNow & slInValue
                            slStr = slStr & Chr(1)
                            tmSaf.sName = slStr & Mid$(tmSaf.sName, 12, 5) & slRptDays
                            ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
                        End If
                    End If
                End If
                gSetRptGGFlag tmSaf
            End If
            btrDestroy hmSaf
        End If
    End If
End Sub

'TTP 10937 - Contract Dashboard: new screen to show contract header information
Sub cmcDashboard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If igAuxWndShowing Then
        DashboardVw.SetFocus
    Else
        DashboardVw.Show vbModal
    End If
    
    DoEvents
    Sleep 50
    If Not igJobShowing(CONTRACTSJOB) Then
        If igDashboardCntrStatus = 1 Then
            sgCntrScreen = "Proposals"  'Change Complete
            Contract.Show vbModeless  'This is required to activate form if previousely activated
        ElseIf igDashboardCntrStatus = 2 Then   'Change Rev Complete
            sgCntrScreen = "Proposals"
            Contract.Show vbModeless  'This is required to activate form if previousely activated
        ElseIf igDashboardCntrStatus = 3 Then   'View mode
            sgCntrScreen = "Proposals"
            Contract.Show vbModeless  'This is required to activate form if previousely activated
        ElseIf igDashboardCntrStatus = 4 Then   'Change Not Schedule
            sgCntrScreen = "Orders"
            Contract.Show vbModeless  'This is required to activate form if previousely activated
        ElseIf igDashboardCntrStatus = 5 Then   'View mode
            sgCntrScreen = "Orders"
            Contract.Show vbModeless  'This is required to activate form if previousely activated
        End If
    Else
        igDashboardCntrStatus = -1
    End If
End Sub

Private Sub cmcStartList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift Then
        Unload Basic10
    Else
        If igAuxWndShowing Then
            AuxWnd.SetFocus
        Else
            AuxWnd.Show vbModeless
        End If
    End If
End Sub

Private Sub cmcTask_Click(Index As Integer)
    If imPrevTaskIndex <> -1 Then
        'lacTask(imPrevTaskIndex).ForeColor = vbBlack
        imPrevTaskIndex = -1
    End If
    If Not igSiteSet Then
        MsgBox "Please Define Site Options", vbOKOnly + vbInformation, "Counterpoint"
        Exit Sub
    End If
    If Not igVehicleDefined Then
        MsgBox "Please Define Vehicle", vbOKOnly + vbInformation, "Counterpoint"
        Exit Sub
    End If
    If (Index = PROPOSALSJOB) Or (Index = CONTRACTSJOB) Then
        If Not gRecExistForFile("Rcf.btr") Then
            MsgBox "Please Define Rate Card", vbOKOnly + vbInformation, "Counterpoint"
            Exit Sub
        End If
        'If Not gWinRoom(igNoLJWinRes(Index) + igNoExeWinRes(CREVNOEXE)) Then
        '    Exit Sub
        'End If
    ElseIf (Index = COPYJOB) Then
        If tgSpf.sUseCartNo <> "N" Then
            If Not gRecExistForFile("Mcf.btr") Then
                MsgBox "Please Define Media", vbOKOnly + vbInformation, "Counterpoint"
                Exit Sub
            End If
        End If
        'If Not gWinRoom(igNoLJWinRes(Index) + igNoExeWinRes(CREVNOEXE)) Then
        '    Exit Sub
        'End If
    ElseIf (Index = LOGSJOB) Then
        'If Not gWinRoom(igNoLJWinRes(Index)) Then
        '    Exit Sub
        'End If
        'If tgSpf.sGUseAffSys <> "Y" Then
        '    hmSpf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
        '    ilRet = btrOpen(hmSpf, "", sgDBPath & "Spf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE) 'BTRV_LOCK_NONE)
        '    If ilRet <> BTRV_ERR_NONE Then
        '        Exit Sub
        '    End If
        '    imSpfRecLen = Len(tmSpf)
        '    'Until multi-user version of btrieve installed- test field
        '    Do
        '        ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        '        If (Asc(tmSpf.sLkLog) = 255) And (tmSpf.iLkLogUrfCode <> tgUrf(0).iCode) Then
        '            MsgBox "Log Function in Use", vbOkOnly + vbInformation, "Counterpoint"
        '            ilRet = btrClose(hmSpf)
        '            btrDestroy hmSpf
        '            Exit Sub
        '        End If
        '        tmSpf.sLkLog = Chr(255)
        '        tmSpf.iLkLogUrfCode = tgUrf(0).iCode
        '        ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
        '    Loop While ilRet = BTRV_ERR_CONFLICT
        '    ilRet = btrClose(hmSpf)
        '    btrDestroy hmSpf
        'End If
    ElseIf (Index = BUDGETSJOB) Or (Index = SLSPCOMMSJOB) Then
        If Not gRecExistForFile("Sof.btr") Then
            MsgBox "Please Define Sales Offices", vbOKOnly + vbInformation, "Counterpoint"
            Exit Sub
        End If
        If Not gRecExistForFile("Slf.btr") Then
            MsgBox "Please Define Salespeople", vbOKOnly + vbInformation, "Counterpoint"
            Exit Sub
        End If
        'If Not gWinRoom(igNoLJWinRes(Index)) Then
        '    Exit Sub
        'End If
    ElseIf (Index = FEEDJOB) Then
        If tgSpf.sSystemType = "R" Then
            If Not gRecExistForFile("Fnf.btr") Then
                MsgBox "Please Define Feed Names", vbOKOnly + vbInformation, "Counterpoint"
                Exit Sub
            End If
        End If
    Else
        'If Not gWinRoom(igNoLJWinRes(Index)) Then
        '    Exit Sub
        'End If
    End If
    If tmcClick.Enabled = False Then
        imTaskIndex = Index
        tmcClick.Enabled = True
    End If
End Sub
Private Sub cmcTask_GotFocus(Index As Integer)
    gCtrlGotFocus cmcTask(Index)
    Select Case Index
        Case RATECARDSJOB 'Rate Card
            gShowHelpMess tgJobHelp(), JOBRATECARD
        Case PROPOSALSJOB 'Order
            gShowHelpMess tgJobHelp(), JOBPROPOSALS
        Case CONTRACTSJOB 'Contract
            gShowHelpMess tgJobHelp(), JOBCONTRACTS
        Case PROGRAMMINGJOB 'Program
            gShowHelpMess tgJobHelp(), JOBPROGRAMS
        Case SPOTSJOB 'Spots
            gShowHelpMess tgJobHelp(), JOBSPOTS
        Case COPYJOB 'Copy
            gShowHelpMess tgJobHelp(), JOBCOPY
        Case LOGSJOB 'Log
            gShowHelpMess tgJobHelp(), JOBLOGS
        Case POSTLOGSJOB 'Invoicing
            gShowHelpMess tgJobHelp(), JOBAFFIDAVITS
        Case INVOICESJOB 'Spot Posting
            gShowHelpMess tgJobHelp(), JOBINVOICES
        Case COLLECTIONSJOB 'Cash Posting
            gShowHelpMess tgJobHelp(), JOBCOLLECTION
        Case BUDGETSJOB 'Reports
            gShowHelpMess tgJobHelp(), JOBBUDGETS
        Case SLSPCOMMSJOB 'Reports
            gShowHelpMess tgJobHelp(), JOBSLSPCOMMS
        Case STATIONFEEDJOB 'Reports
'            gShowHelpMess tgJobHelp(), JOBAFFILIATE
        'Case REPORTSJOB 'Reports
        '    gShowHelpMess tgJobHelp(), JOBREPORTS
        Case Else
            imcOutline.Visible = False
    End Select
End Sub

Private Sub cmcTask_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move cmcTask(Index).Left - 90, cmcTask(Index).Top - 90, cmcTask(Index).Width, cmcTask(Index).height
'    cmcTask(Index).ToolTipText = ""
    If imPrevTaskIndex <> -1 Then
        'lacTask(imPrevTaskIndex).ForeColor = vbBlack
        imPrevTaskIndex = -1
    End If
    Select Case Index
        Case RATECARDSJOB 'Rate Card
            If igWinStatus(RATECARDSJOB) = 0 Then
                cmcTask(Index).ToolTipText = "Disallowed"
            Else
                If igWinStatus(RATECARDSJOB) = 1 Then
                    cmcTask(Index).ToolTipText = "View Only"
                    imPrevTaskIndex = Index
                    imcOutline.BackColor = vbYellow
                    'lacTask(imPrevTaskIndex).ForeColor = vbYellow
                Else
                    cmcTask(Index).ToolTipText = "Updates Allowed"
                    imcOutline.BackColor = &H80FF80
                End If
            End If
        Case PROPOSALSJOB 'Order
            If tgSpf.sGUsePropSys = "Y" Then
                If igWinStatus(PROPOSALSJOB) = 0 Then
                    cmcTask(Index).ToolTipText = "Disallowed"
                Else
                    If igWinStatus(PROPOSALSJOB) = 1 Then
                        cmcTask(Index).ToolTipText = "View Only"
                        imPrevTaskIndex = Index
                        'lacTask(imPrevTaskIndex).ForeColor = vbYellow
                    Else
                        cmcTask(Index).ToolTipText = "Updates Allowed"
                        imcOutline.BackColor = &H80FF80
                    End If
                End If
            Else
                cmcTask(Index).ToolTipText = "Disallowed"
            End If
        Case CONTRACTSJOB 'Contract
            If igWinStatus(CONTRACTSJOB) = 0 Then
                cmcTask(Index).ToolTipText = "Disallowed"
            Else
                If igWinStatus(CONTRACTSJOB) = 1 Then
                    cmcTask(Index).ToolTipText = "View Only"
                    imPrevTaskIndex = Index
                    imcOutline.BackColor = vbYellow
                    'lacTask(imPrevTaskIndex).ForeColor = vbYellow
                Else
                    cmcTask(Index).ToolTipText = "Updates Allowed"
                    imcOutline.BackColor = &H80FF80
                End If
            End If
        Case PROGRAMMINGJOB 'Program
            If igWinStatus(PROGRAMMINGJOB) = 0 Then
                cmcTask(Index).ToolTipText = "Disallowed"
            Else
                If igWinStatus(PROGRAMMINGJOB) = 1 Then
                    cmcTask(Index).ToolTipText = "View Only"
                    imPrevTaskIndex = Index
                    imcOutline.BackColor = vbYellow
                    'lacTask(imPrevTaskIndex).ForeColor = vbYellow
                Else
                    cmcTask(Index).ToolTipText = "Updates Allowed"
                    imcOutline.BackColor = &H80FF80
                End If
            End If
        Case SPOTSJOB 'Spots
            If igWinStatus(SPOTSJOB) = 0 Then
                cmcTask(Index).ToolTipText = "Disallowed"
            Else
                If igWinStatus(SPOTSJOB) = 1 Then
                    cmcTask(Index).ToolTipText = "View Only"
                    imcOutline.BackColor = vbYellow
                    imPrevTaskIndex = Index
                    'lacTask(imPrevTaskIndex).ForeColor = vbYellow
                Else
                    cmcTask(Index).ToolTipText = "Updates Allowed"
                    imcOutline.BackColor = &H80FF80
                End If
            End If
        Case COPYJOB 'Copy
            If igWinStatus(COPYJOB) = 0 Then
                cmcTask(Index).ToolTipText = "Disallowed"
            Else
                If igWinStatus(COPYJOB) = 1 Then
                    cmcTask(Index).ToolTipText = "View Only"
                    imcOutline.BackColor = vbYellow
                    imPrevTaskIndex = Index
                    'lacTask(imPrevTaskIndex).ForeColor = vbYellow
                Else
                    cmcTask(Index).ToolTipText = "Updates Allowed"
                    imcOutline.BackColor = &H80FF80
                End If
            End If
        Case LOGSJOB 'Log
            If igWinStatus(LOGSJOB) = 0 Then
                cmcTask(Index).ToolTipText = "Disallowed"
            Else
                If igWinStatus(LOGSJOB) = 1 Then
                    cmcTask(Index).ToolTipText = "View Only"
                    imcOutline.BackColor = vbYellow
                    imPrevTaskIndex = Index
                    'lacTask(imPrevTaskIndex).ForeColor = vbYellow
                Else
                    cmcTask(Index).ToolTipText = "Updates Allowed"
                    imcOutline.BackColor = &H80FF80
                End If
            End If
        Case POSTLOGSJOB 'Invoicing
            If igWinStatus(POSTLOGSJOB) = 0 Then
                cmcTask(Index).ToolTipText = "Disallowed"
            Else
                If igWinStatus(POSTLOGSJOB) = 1 Then
                    cmcTask(Index).ToolTipText = "View Only"
                    imPrevTaskIndex = Index
                    imcOutline.BackColor = vbYellow
                    'lacTask(imPrevTaskIndex).ForeColor = vbYellow
                Else
                    cmcTask(Index).ToolTipText = "Updates Allowed"
                    imcOutline.BackColor = &H80FF80
                End If
            End If
        Case INVOICESJOB 'Spot Posting
            If igWinStatus(INVOICESJOB) = 0 Then
                cmcTask(Index).ToolTipText = "Disallowed"
            Else
                If igWinStatus(INVOICESJOB) = 1 Then
                    cmcTask(Index).ToolTipText = "View Only"
                    imPrevTaskIndex = Index
                    imcOutline.BackColor = vbYellow
                    'lacTask(imPrevTaskIndex).ForeColor = vbYellow
                Else
                    cmcTask(Index).ToolTipText = "Updates Allowed"
                    imcOutline.BackColor = &H80FF80
                End If
            End If
        Case COLLECTIONSJOB 'Cash Posting
            If igWinStatus(COLLECTIONSJOB) = 0 Then
                cmcTask(Index).ToolTipText = "Disallowed"
            Else
                If igWinStatus(COLLECTIONSJOB) = 1 Then
                    cmcTask(Index).ToolTipText = "View Only"
                    imPrevTaskIndex = Index
                    imcOutline.BackColor = vbYellow
                    'lacTask(imPrevTaskIndex).ForeColor = vbYellow
                Else
                    cmcTask(Index).ToolTipText = "Updates Allowed"
                    imcOutline.BackColor = &H80FF80
                End If
            End If
        Case BUDGETSJOB 'Reports
            If tgSpf.sGUsePropSys = "Y" Then
                If igWinStatus(BUDGETSJOB) = 0 Then
                    cmcTask(Index).ToolTipText = "Disallowed"
                Else
                    If igWinStatus(BUDGETSJOB) = 1 Then
                        cmcTask(Index).ToolTipText = "View Only"
                        imcOutline.BackColor = vbYellow
                        imPrevTaskIndex = Index
                        'lacTask(imPrevTaskIndex).ForeColor = vbYellow
                    Else
                        cmcTask(Index).ToolTipText = "Updates Allowed"
                        imcOutline.BackColor = &H80FF80
                    End If
                End If
            Else
                cmcTask(Index).ToolTipText = "Disallowed"
            End If
        Case SLSPCOMMSJOB 'Reports
            If igWinStatus(SLSPCOMMSJOB) = 0 Then
                cmcTask(Index).ToolTipText = "Disallowed"
            Else
                If igWinStatus(SLSPCOMMSJOB) = 1 Then
                    cmcTask(Index).ToolTipText = "View Only"
                    imcOutline.BackColor = vbYellow
                    imPrevTaskIndex = Index
                    'lacTask(imPrevTaskIndex).ForeColor = vbYellow
                Else
                    cmcTask(Index).ToolTipText = "Updates Allowed"
                    imcOutline.BackColor = &H80FF80
                End If
            End If
        Case STATIONFEEDJOB 'Reports
'            If tgSpf.sSystemType <> "R" Then
'                If tgSpf.sGUseAffFeed = "Y" Then
'                    If igWinStatus(STATIONFEEDJOB) = 0 Then
'                        cmcTask(Index).ToolTipText = "Disallowed"
'                    Else
'                        If igWinStatus(STATIONFEEDJOB) = 1 Then
'                            cmcTask(Index).ToolTipText = "View Only"
'                            imPrevTaskIndex = Index
'                            'lacTask(imPrevTaskIndex).ForeColor = vbYellow
'                        Else
'                            cmcTask(Index).ToolTipText = "Updates Allowed"
'                        End If
'                    End If
'                Else
'                    cmcTask(Index).ToolTipText = "Disallowed"
'                End If
'            Else
                cmcTask(Index).ToolTipText = "Disallowed"
'            End If
        Case FEEDJOB 'Invoicing
            If tgSpf.sSystemType = "R" Then
                '12/24/12: Removed FeedSpot because of Out of memory
                'If igWinStatus(FEEDJOB) = 0 Then
                '    cmcTask(Index).ToolTipText = "Disallowed"
                'Else
                '    If igWinStatus(FEEDJOB) = 1 Then
                '        cmcTask(Index).ToolTipText = "View Only"
                '        imPrevTaskIndex = Index
                '        'lacTask(imPrevTaskIndex).ForeColor = vbYellow
                '    Else
                '        cmcTask(Index).ToolTipText = "Updates Allowed"
                '    End If
                'End If
                '12/24/12: End Point
            Else
                cmcTask(Index).ToolTipText = "Disallowed"
            End If
        'Case REPORTSJOB 'Reports
    End Select
    imcOutline.Visible = True
End Sub

Private Sub Form_Activate()
    Dim ilLoop As Integer
    igBasic10Showing = True
    gShowPartBranner
    mSetControls
    imTaskIndex = 100
    tmcClick.Enabled = True
    If igGGFlag = 0 Then
        pbcHor.Enabled = False
        For ilLoop = RATECARDSJOB To REPORTSJOB - 1 Step 1
            cmcTask(ilLoop).Enabled = False
            lacTask(ilLoop).Enabled = False
            lacTask(ilLoop).ForeColor = vbBlack
        Next ilLoop
        'cmcReport.Enabled = False
    End If
    Me.KeyPreview = True
    Me.ZOrder 0 'Send to front
    pbcHor.Visible = False
    DoEvents
    pbcHor.Visible = True
    
    'TTP 10937 - Contract Dashboard: new screen to show contract header information
    If igTrafficFirstActivate Then
        If cmcDashboard.Enabled = True And cmcDashboard.Visible = True And tgUrf(0).iDashOnStart = 1 Then
            DoEvents
            cmcDashboard_MouseUp 0, 0, 0, 0
        End If
    End If
    igTrafficFirstActivate = False
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    'ElseIf KeyCode = KEYF5 Then
    '    AuxWnd.Show vbModeless
    End If
    If (KeyCode = vbKey8) And (Shift = 4) Then
        igBasicOrAux = -1
        igBasic10Showing = False
        Unload Basic10
    End If
End Sub

Private Sub Form_Load()
    imFirstActivate = True
    lacSalesGroup.ForeColor = RGB(131, 60, 11)
    lacTrafficGroup.ForeColor = RGB(97, 144, 160)
    lacAccountingGroup.ForeColor = RGB(125, 171, 130)
    
    If igBasicOrAux = 2 Then
        Unload AuxWnd
    End If
    igBasic10Showing = True
    igBasicOrAux = 1    'Indicate that basic 10 loaded
    imTaskIndex = -1
    imPrevTaskIndex = -1
    'Basic10.Caption = "Jobs"
    'cmcCSLogo.Picture = IconTraf!imcCSLogo.Picture
    plcBasic10.Left = 0
    plcBasic10.Top = 0
    Basic10.height = plcBasic10.height + fgPanelAdj ' \ 2 + 30
    Basic10.Width = plcBasic10.Width + fgPanelAdj ' \ 2 + 30
        
    If lgCurrBPP <= 8 Then
        pbcHor.Picture = Traffic!pbc256Job
    End If
    'cmcCSLogo.Picture = IconTraf!imcCSLogo.Picture
    gCenterForm Basic10
    
    'TTP 10937 - Contract Dashboard: new screen to show contract header information
    If (tgSpf.sGUsePropSys <> "Y" Or igWinStatus(PROPOSALSJOB) = 0) And igWinStatus(CONTRACTSJOB) = 0 Then
        cmcDashboard.Enabled = False
        cmcDashboard.Visible = False
    End If

End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mShowCredit                     *
'*                                                     *
'*             Created:10/27/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Show credit alert              *
'*                                                     *
'*******************************************************
Private Sub mShowCredit()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilErr                         slSpfDate                                               *
'******************************************************************************************

    Dim ilRet As Integer
    Dim slDate As String
    Dim slStr As String
    slDate = Format$(gNow(), "m/d/yy")  'Test if credit has been shown for this date
    'If gDateValue(slDate) <> lgDateCreditShown Then
    If Not igShowCreditShown Then
        gSpfRead
        'gSetUsingTraffic
        '6/14/07:  Remove date test
        'gUnpackDate tgSpf.iRCreditDate(0), tgSpf.iRCreditDate(1), slSpfDate
        'If gDateValue(slSpfDate) = gDateValue(slDate) Then
            'Screen.MousePointer = vbHourglass
            'Traffic.Enabled = False
            'ilShell = 0
            'ilErr = False
            'On Error GoTo mShowCreditErr
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName
            Else
                slStr = "Traffic^Prod\" & sgUserName
            End If
            'ilShell = Shell(sgExePath & "ShoCrdit.Exe " & slStr, 1)
            'Screen.MousePointer = vbDefault
            'If ilErr Then
            '    Exit Sub
            'End If
            'While GetModuleUsage(ilShell) > 0
            '    ilRet = DoEvents()
            'Wend
            'Traffic.Enabled = True

            'Temporaryly removed until ShoCrdit Added to v5.0
            ilRet = 0
            On Error GoTo mShowCreditErr
            'Modeless required so that Collect can be call
            ShoCrdit.Show vbModeless    'vbModal
            If ilRet = 0 Then
                'lgDateCreditShown = gDateValue(slDate)
                igShowCreditShown = True
            End If
            On Error GoTo 0
            'cmcTask(COLLECTIONSJOB).SetFocus
        'End If
    End If
    Exit Sub
mShowCreditErr:
    ilRet = 0
    Resume Next
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    ilRet = btrClose(hmSaf)
    btrDestroy hmSaf
    
    igBasic10Showing = False
    Set Basic10 = Nothing
End Sub

Private Sub pbcClickFocus_GotFocus()
    imcOutline.Visible = False
    If imPrevTaskIndex <> -1 Then
        'lacTask(imPrevTaskIndex).ForeColor = vbBlack
        imPrevTaskIndex = -1
    End If
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub plcBasic10_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcBasic10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilLoop As Integer
    imcOutline.Visible = False
    If imPrevTaskIndex <> -1 Then
        'lacTask(imPrevTaskIndex).ForeColor = vbBlack
        imPrevTaskIndex = -1
    End If
    
    mDisallowToolTip X, Y

End Sub

Private Sub tmcClick_Timer()

    Dim slStr As String        'General string
    Dim ilLoop As Integer
    Dim ilRptType As Integer
    Dim slCommand As String
    Dim ilAskMkt As Integer
    Dim ilShowSelection As Integer

    tmcClick.Enabled = False
    gShowHelpMess tgJobHelp(), -1   'Blank help
    If imTaskIndex = PROPOSALSJOB Then
        If tgSpf.sGUsePropSys = "Y" Then
            If igWinStatus(PROPOSALSJOB) = 0 Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    ElseIf imTaskIndex = CONTRACTSJOB Then
        If igWinStatus(CONTRACTSJOB) = 0 Then
            Exit Sub
        End If
    End If
    '*** Market Code
    If (tgSpf.sMktBase = "Y") And ((imTaskIndex = PROPOSALSJOB) Or (imTaskIndex = CONTRACTSJOB) Or (imTaskIndex = LOGSJOB) Or (imTaskIndex = SPOTSJOB)) Then
        ilAskMkt = True
        If imTaskIndex = CONTRACTSJOB Then
            If igJobShowing(CONTRACTSJOB) Then
                ilAskMkt = False
            End If
        End If
        If imTaskIndex = PROPOSALSJOB Then
            If igJobShowing(PROPOSALSJOB) Then
                ilAskMkt = False
            End If
        End If
        If imTaskIndex = LOGSJOB Then
            If igJobShowing(LOGSJOB) Then
                ilAskMkt = False
            End If
        End If
        If imTaskIndex = SPOTSJOB Then
            If igJobShowing(SPOTSJOB) Then
                ilAskMkt = False
            End If
        End If
        If ilAskMkt Then
            LogMkt.Show vbModal
            DoEvents
            If UBound(igMktCode) <= 0 Then
                plcBasic10.Visible = False
                plcBasic10.Visible = True
                imTaskIndex = -1
                Exit Sub
            End If
            If (imTaskIndex = PROPOSALSJOB) Or (imTaskIndex = CONTRACTSJOB) Then
                ReDim igCntrMktCode(0 To UBound(igMktCode)) As Integer
                For ilLoop = 0 To UBound(igMktCode) Step 1
                    igCntrMktCode(ilLoop) = igMktCode(ilLoop)
                Next ilLoop
            ElseIf (imTaskIndex = SPOTSJOB) Then
                ReDim igSpotMktCode(0 To UBound(igMktCode)) As Integer
                For ilLoop = 0 To UBound(igMktCode) Step 1
                    igSpotMktCode(ilLoop) = igMktCode(ilLoop)
                Next ilLoop
            Else
                ReDim igLogMktCode(0 To UBound(igMktCode)) As Integer
                For ilLoop = 0 To UBound(igMktCode) Step 1
                    igLogMktCode(ilLoop) = igMktCode(ilLoop)
                Next ilLoop
            End If
        End If
    Else
        'ReDim igMktCode(0 To 0) As Integer
        'ReDim igCntrMktCode(0 To 0) As Integer
        'ReDim igLogMktCode(0 To 0) As Integer
    End If
    '*** Market Code
    Select Case imTaskIndex
        Case RATECARDSJOB  'Rate Card
'            ''Screen.MousePointer = vbHourGlass  'Wait
'            'RateCard.Caption = "Rate Card"  'This reference causes the form to load
'            'RateCard.Show vbModeless  'This is required to activate form if previousely activated
'            ''Screen.MousePointer = vbDefault    'Default
'            'Traffic!cmcBasicWnd.Value = False
'            igRptCallType = RATECARDSJOB    'Using contractjob instead spotsjob
'            ilRptType = 0
'            Screen.MousePointer = vbHourglass  'Wait
'            igChildDone = False 'edcLinkDestDoneMsg.Text = ""
'            Traffic!edcLinkSrceHelpMsg.Text = ""
'            If igTestSystem Then
'                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            Else
'                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            End If
'            lgShellRet = Shell(sgExePath & "RateCard.Exe " & slStr, 1)
'            Traffic.Enabled = False
'            Screen.MousePointer = vbDefault  'Wait
'            Do While Not igChildDone
'                DoEvents
'            Loop
'            slStr = sgDoneMsg
'            Traffic.Enabled = True
'            Traffic!edcLinkSrceHelpMsg.Text = "Ok"
'            For ilLoop = 0 To 10
'                DoEvents
'            Next ilLoop
'            'Read in Spf if Reallocation set date
'            gSpfRead
'            'Re-Read User to get Urf.sRCView
'            sgUrfStamp = ""
'            ilRet = csiSetStamp("URF", sgUrfStamp)
'            gUrfRead Basic10, sgUserName, True, tgUrf(), False  'Obtain user records (and set sgCPName and sgSUName)
'            If igRUDirectMode = 1 Then
'                tgUrf(0).iRemoteUserID = 0
'                'Disallow Contract Projections in Direct Mode (1/4/2000)
'                igWinStatus(PROPOSALSJOB) = 1
'            ElseIf (igRUDirectMode = 2) And (tgUrf(0).iRemoteUserID > 0) Then
'                igWinStatus(PROPOSALSJOB) = 1
'            End If
            imcOutline.Visible = False
'            Traffic!mnuRemoteUser.Enabled = False
            On Error Resume Next
            'If igJobShowing(RATECARDSJOB) Then
            '    RateCard.SetFocus
            'Else
            '    RateCard.Show vbModeless  'This is required to activate form if previousely activated
            'End If
            gShellToRateCard
'            'Screen.MousePointer = vbDefault    'Default
'            Traffic!cmcBasicWnd.Value = False
'
'            'Read in Spf if Reallocation set date
'            gSpfRead
'            'Re-Read User to get Urf.sRCView
'            sgUrfStamp = ""
'            ilRet = csiSetStamp("URF", sgUrfStamp)
'            gUrfRead Basic10, sgUserName, True, tgUrf(), False  'Obtain user records (and set sgCPName and sgSUName)
'            If igRUDirectMode = 1 Then
'                tgUrf(0).iRemoteUserID = 0
'                'Disallow Contract Projections in Direct Mode (1/4/2000)
'                igWinStatus(PROPOSALSJOB) = 1
'            ElseIf (igRUDirectMode = 2) And (tgUrf(0).iRemoteUserID > 0) Then
'                igWinStatus(PROPOSALSJOB) = 1
'            End If

        Case PROPOSALSJOB  'Proposal
            imcOutline.Visible = False
'            Traffic!mnuRemoteUser.Enabled = False
            'Screen.MousePointer = vbHourGlass  'Wait
            'Contract.Caption = "Proposals"   'This reference causes the form to load
            'Contract.plcScreen = "Proposals"
            On Error Resume Next
            If igJobShowing(CONTRACTSJOB) Then
                Contract.SetFocus
            Else
                igAlertCntrStatus = -1
                sgCntrScreen = "Proposals"
                Contract.Show vbModeless  'This is required to activate form if previousely activated
            End If
            'Screen.MousePointer = vbDefault    'Default
'            Traffic!cmcBasicWnd.Value = False
        Case CONTRACTSJOB  'Contract
            imcOutline.Visible = False
'            Traffic!mnuRemoteUser.Enabled = False
            'Screen.MousePointer = vbHourGlass  'Wait
            'Contract.Caption = "Orders"  'This reference causes the form to load
            'Contract.plcScreen = "Orders"
            On Error Resume Next
            If igJobShowing(CONTRACTSJOB) Then
                Contract.SetFocus
            Else
                '8723
                If ogContractCreator.isFilemakerActive Then
                    FileMakerToOrder.Show vbModal
                Else
                    sgCntrScreen = "Orders"
                    igAlertCntrStatus = -1
                    Contract.Show vbModeless  'This is required to activate form if previousely activated
                End If
            End If
            'Screen.MousePointer = vbDefault    'Default
'            Traffic!cmcBasicWnd.Value = False
        Case PROGRAMMINGJOB  'Program
'            ''Screen.MousePointer = vbHourGlass  'Wait
'            'Program.Caption = "Programming" 'This reference causes the form to load
'            'Program.Show vbModeless  'This is required to activate form if previousely activated
'            ''Screen.MousePointer = vbDefault    'Default
'            'Traffic!cmcBasicWnd.Value = False
'            igRptCallType = PROGRAMMINGJOB    'Using contractjob instead spotsjob
'            ilRptType = 0
'            Screen.MousePointer = vbHourglass  'Wait
'            igChildDone = False 'edcLinkDestDoneMsg.Text = ""
'            Traffic!edcLinkSrceHelpMsg.Text = ""
'            If igTestSystem Then
'                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            Else
'                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            End If
'            lgShellRet = Shell(sgExePath & "Program.Exe " & slStr, 1)
'            Traffic.Enabled = False
'            Screen.MousePointer = vbDefault  'Wait
'            Do While Not igChildDone
'                DoEvents
'            Loop
'            slStr = sgDoneMsg
'            Traffic.Enabled = True
'            Traffic!edcLinkSrceHelpMsg.Text = "Ok"
'            For ilLoop = 0 To 10
'                DoEvents
'            Next ilLoop

            imcOutline.Visible = False
'            Traffic!mnuRemoteUser.Enabled = False
            On Error Resume Next
            If igJobShowing(PROGRAMMINGJOB) Then
                Program.SetFocus
            Else
                Program.Show vbModeless  'This is required to activate form if previousely activated
            End If
'            'Screen.MousePointer = vbDefault    'Default
'            Traffic!cmcBasicWnd.Value = False
'
'            Screen.MousePointer = vbHourglass
'            sgMVefStamp = "~"   'Force Read
'            ilRet = gObtainVef()
'            sgVpfStamp = "~"    'Force read
'            ilRet = gVpfRead(AuxWnd)
'            gObtainVirtVehList
'            Screen.MousePointer = vbDefault    'Default
        Case SPOTSJOB  'Spots
            'sgSpotCallType = "A"    '"T"
            ''Spots.Caption = "Spots" 'This reference causes the form to load
            ''Spots.Show vbModeless  'This is required to activate form if previousely activated
            ''Traffic!cmcBasicWnd.Value = False
            'igRptCallType = SPOTSJOB    'Using contractjob instead spotsjob
            'ilRptType = 0
            'Screen.MousePointer = vbHourglass  'Wait
            'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
            'Traffic!edcLinkSrceHelpMsg.Text = ""
            'If igTestSystem Then
            '    slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
            'Else
            '    slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
            'End If
            'lgShellRet = Shell(sgExePath & "Spots.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            'slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic!edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            imcOutline.Visible = False
'            Traffic!mnuRemoteUser.Enabled = False
            If tgSpf.sSystemType = "R" Then
                sgSpotCallType = "T"
            Else
                sgSpotCallType = "A"
            End If
            On Error Resume Next
            If igJobShowing(SPOTSJOB) Then
                Spots.SetFocus
            Else
                Spots.Show vbModeless  'This is required to activate form if previousely activated
            End If
            'Screen.MousePointer = vbDefault    'Default
'            Traffic!cmcBasicWnd.Value = False
        Case COPYJOB  'Copy
'            ''Screen.MousePointer = vbHourGlass  'Wait
'            'Copy.Caption = "Copy"   'This reference causes the form to load
'            'Copy.Show vbModeless  'This is required to activate form if previousely activated
'            ''Screen.MousePointer = vbDefault    'Default
'            'Traffic!cmcBasicWnd.Value = False
'            igRptCallType = COPYJOB    'Using contractjob instead spotsjob
'            ilRptType = 0
'            Screen.MousePointer = vbHourglass  'Wait
'            igChildDone = False 'edcLinkDestDoneMsg.Text = ""
'            Traffic!edcLinkSrceHelpMsg.Text = ""
'            If igTestSystem Then
'                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            Else
'                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            End If
'            lgShellRet = Shell(sgExePath & "Copy.Exe " & slStr, 1)
'            Traffic.Enabled = False
'            Screen.MousePointer = vbDefault  'Wait
'            Do While Not igChildDone
'                DoEvents
'            Loop
'            slStr = sgDoneMsg
'            Traffic.Enabled = True
'            Traffic!edcLinkSrceHelpMsg.Text = "Ok"
'            For ilLoop = 0 To 10
'                DoEvents
'            Next ilLoop

            imcOutline.Visible = False
'            Traffic!mnuRemoteUser.Enabled = False
            On Error Resume Next
            If igJobShowing(COPYJOB) Then
                Copy.SetFocus
            Else
                Copy.Show vbModeless  'This is required to activate form if previousely activated
            End If
            'Screen.MousePointer = vbDefault    'Default
'            Traffic!cmcBasicWnd.Value = False

        Case LOGSJOB  'Log
'            ''Screen.MousePointer = vbHourGlass  'Wait
'            'Logs.Caption = "Logs"   'This reference causes the form to load
'            'Logs.Show vbModeless  'This is required to activate form if previousely activated
'            ''Screen.MousePointer = vbDefault    'Default
'            'Traffic!cmcBasicWnd.Value = False
'            'If Not gWinRoom(igNoLJWinRes(Index) + igNoExeWinRes(LOGSJOB)) Then
'            '    Exit Sub
'            'End If
'            igRptCallType = LOGSJOB    'Using contractjob instead spotsjob
'            ilRptType = 0
'            Screen.MousePointer = vbHourglass  'Wait
'            igChildDone = False 'edcLinkDestDoneMsg.Text = ""
'            Traffic!edcLinkSrceHelpMsg.Text = ""
'            If igTestSystem Then
'                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            Else
'                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            End If
'            lgShellRet = Shell(sgExePath & "Logs.Exe " & slStr, 1)
'            Traffic.Enabled = False
'            Screen.MousePointer = vbDefault  'Wait
'            Do While Not igChildDone
'                DoEvents
'            Loop
'            slStr = sgDoneMsg
'            Traffic.Enabled = True
'            Traffic!edcLinkSrceHelpMsg.Text = "Ok"
'            For ilLoop = 0 To 10
'                DoEvents
'            Next ilLoop
'            'hmSpf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
'            'ilRet = btrOpen(hmSpf, "", sgDBPath & "Spf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE) 'BTRV_LOCK_NONE)
'            'If ilRet <> BTRV_ERR_NONE Then
'            '    Exit Sub
'            'End If
'            'imSpfRecLen = Len(tmSpf)
'            ''Until multi-user version of btrieve installed- test field
'            'Do
'            '    ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'            '    tmSpf.sLkLog = " "
'            '    ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
'            'Loop While ilRet = BTRV_ERR_CONFLICT
'            'ilRet = btrClose(hmSpf)
'            'btrDestroy hmSpf
'            Screen.MousePointer = vbHourglass
'            sgMVefStamp = "~"   'Force Read
'            ilRet = gObtainVef()
'            sgVpfStamp = "~"    'Force read
'            ilRet = gVpfRead(Basic10)
'            gObtainVirtVehList
'            Screen.MousePointer = vbDefault    'Default

            imcOutline.Visible = False
'            Traffic!mnuRemoteUser.Enabled = False
            On Error Resume Next
            If igJobShowing(LOGSJOB) Then
                Logs.SetFocus
            Else
                Logs.Show vbModeless  'This is required to activate form if previousely activated
            End If
'            'Screen.MousePointer = vbDefault    'Default
'            Traffic!cmcBasicWnd.Value = False
'            Screen.MousePointer = vbHourglass
'            sgMVefStamp = "~"   'Force Read
'            ilRet = gObtainVef()
'            sgVpfStamp = "~"    'Force read
'            ilRet = gVpfRead(Basic10)
'            gObtainVirtVehList
'            Screen.MousePointer = vbDefault    'Default
        Case POSTLOGSJOB  'Spot posting
'            ''Screen.MousePointer = vbHourGlass  'Wait
'            'PostLog.Caption = "Post Log" 'This reference causes the form to load
'            'PostLog.Show vbModeless  'This is required to activate form if previousely activated
'            ''Screen.MousePointer = vbDefault    'Default
'            'Traffic!cmcBasicWnd.Value = False
'            igRptCallType = POSTLOGSJOB    'Using contractjob instead spotsjob
'            ilRptType = 0
'            Screen.MousePointer = vbHourglass  'Wait
'            igChildDone = False 'edcLinkDestDoneMsg.Text = ""
'            Traffic!edcLinkSrceHelpMsg.Text = ""
'            If igTestSystem Then
'                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            Else
'                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            End If
'            lgShellRet = Shell(sgExePath & "PostLog.Exe " & slStr, 1)
'            Traffic.Enabled = False
'            Screen.MousePointer = vbDefault  'Wait
'            Do While Not igChildDone
'                DoEvents
'            Loop
'            slStr = sgDoneMsg
'            Traffic.Enabled = True
'            Traffic!edcLinkSrceHelpMsg.Text = "Ok"
'            For ilLoop = 0 To 10
'                DoEvents
'            Next ilLoop


            imcOutline.Visible = False
'            Traffic!mnuRemoteUser.Enabled = False
            On Error Resume Next
            If igJobShowing(POSTLOGSJOB) Then
                If igPostType = 0 Then
                    PostLog.SetFocus
                ElseIf igPostType = 1 Then
                    PostRep.SetFocus
                ElseIf igPostType = 2 Then
                    PostRep.SetFocus
                ElseIf igPostType = 3 Then
                    PostRep.SetFocus
                ElseIf igPostType = 4 Then
                    PostRep.SetFocus
                ElseIf igPostType = 5 Then
                    PostRep.SetFocus
                ElseIf igPostType = 6 Then
                    PostRep.SetFocus
                End If
            Else
                gAnyClustersDef
                gAnyRepDef
                'Determine which Post Log screen
                If (sgMktClusterDef = "Y") Or (sgRepDef = "Y") Then
                    'If tgSpf.sPostCalAff <> "D" Then
                    If (tgSpf.sPostCalAff <> "N") Or ((Asc(tgSpf.sUsingFeatures8) And REPBYDT) = REPBYDT) Then
                        ilShowSelection = True
                    Else
                        ilShowSelection = False
                        'igPostType = 5
                        'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        '    If (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S") Then
                        '        ilShowSelection = True
                        '        Exit For
                        '    End If
                        'Next ilLoop
                    End If
                    If ilShowSelection Then
                        PLogMkt.Show vbModal
                    End If
                    DoEvents
                    If igPostType < 0 Then
                        plcBasic10.Visible = False
                        plcBasic10.Visible = True
                        imTaskIndex = -1
                        Exit Sub
                    End If
                Else
                    igPostType = 0
                End If
                If igPostType = 0 Then
                    PostLog.Show vbModeless  'This is required to activate form if previousely activated
                ElseIf igPostType = 1 Then
                    PostRep.Show vbModeless  'This is required to activate form if previousely activated
                ElseIf igPostType = 2 Then
                    PostRep.Show vbModeless  'This is required to activate form if previousely activated
                ElseIf igPostType = 3 Then
                    PostRep.Show vbModeless  'This is required to activate form if previousely activated
                ElseIf igPostType = 4 Then
                    PostRep.Show vbModeless  'This is required to activate form if previousely activated
                ElseIf igPostType = 5 Then
                    PostRep.Show vbModeless  'This is required to activate form if previousely activated
                ElseIf igPostType = 6 Then
                    PostRep.Show vbModeless  'This is required to activate form if previousely activated
                End If
            End If
            'Screen.MousePointer = vbDefault    'Default
'            Traffic!cmcBasicWnd.Value = False

        Case INVOICESJOB  'Invoicing
'            'Screen.MousePointer = vbHourGlass  'Wait
'            'Invoice.Caption = "Invoices"   'This reference causes the form to load
'            'Invoice.Show vbModeless  'This is required to activate form if previousely activated
'            'Screen.MousePointer = vbDefault    'Default
'            'Traffic!cmcBasicWnd.Value = False
'            'If Not gWinRoom(igNoLJWinRes(Index) + igNoExeWinRes(INVOICESJOB)) Then
'            '    Exit Sub
'            'End If
'            igRptCallType = INVOICESJOB    'Using contractjob instead spotsjob
'            ilRptType = 0
'            Screen.MousePointer = vbHourglass  'Wait
'            igChildDone = False 'edcLinkDestDoneMsg.Text = ""
'            Traffic!edcLinkSrceHelpMsg.Text = ""
'            If igTestSystem Then
'                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            Else
'                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            End If
'            lgShellRet = Shell(sgExePath & "Invoice.Exe " & slStr, 1)
'            Traffic.Enabled = False
'            Screen.MousePointer = vbDefault  'Wait
'            Do While Not igChildDone
'                DoEvents
'            Loop
'            slStr = sgDoneMsg
'            Traffic.Enabled = True
'            Traffic!edcLinkSrceHelpMsg.Text = "Ok"
'            For ilLoop = 0 To 10
'                DoEvents
'            Next ilLoop

            imcOutline.Visible = False
'            Traffic!mnuRemoteUser.Enabled = False
            On Error Resume Next
            'If igJobShowing(INVOICESJOB) Then
            '    Invoice.SetFocus
            'Else
            '    gSpfRead
            '    'gSetUsingTraffic
            '    Invoice.Show vbModeless  'This is required to activate form if previousely activated
            'End If
            gShellToInvoice 'Basic10
'            'Screen.MousePointer = vbDefault    'Default
'            Traffic!cmcBasicWnd.Value = False
'
'            sgSpfStamp = "~"    'Force read
'            gSpfRead
            'Screen.MousePointer = vbDefault
        Case COLLECTIONSJOB  'Cash Posting
'            ''Screen.MousePointer = vbHourGlass  'Wait
'            'Collect.Caption = "Collections" 'This reference causes the form to load
'            'Collect.Show vbModeless  'This is required to activate form if previousely activated
'            ''Screen.MousePointer = vbDefault    'Default
'            'Traffic!cmcBasicWnd.Value = False
'            igRptCallType = COLLECTIONSJOB    'Using contractjob instead spotsjob
'            ilRptType = 0
'            Screen.MousePointer = vbHourglass  'Wait
'            igChildDone = False 'edcLinkDestDoneMsg.Text = ""
'            Traffic!edcLinkSrceHelpMsg.Text = ""
'            If igTestSystem Then
'                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            Else
'                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            End If
'            lgShellRet = Shell(sgExePath & "Collect.Exe " & slStr, 1)
'            Traffic.Enabled = False
'            Screen.MousePointer = vbDefault  'Wait
'            Do While Not igChildDone
'                DoEvents
'            Loop
'            slStr = sgDoneMsg
'            Traffic.Enabled = True
'            Traffic!edcLinkSrceHelpMsg.Text = "Ok"
'            For ilLoop = 0 To 10
'                DoEvents
'            Next ilLoop

            imcOutline.Visible = False
'            Traffic!mnuRemoteUser.Enabled = False
            On Error Resume Next
            If igJobShowing(COLLECTIONSJOB) Then
                Collect.SetFocus
            Else
                Collect.Show vbModeless  'This is required to activate form if previousely activated
            End If
'            'Screen.MousePointer = vbDefault    'Default
'            Traffic!cmcBasicWnd.Value = False
'
'
'            sgSpfStamp = "~"    'Force read
'            gSpfRead
        Case STATIONFEEDJOB
''            igRptCallType = STATIONFEEDJOB    'Using contractjob instead spotsjob
''            ilRptType = 0
''            Screen.MousePointer = vbHourglass  'Wait
''            igChildDone = False 'edcLinkDestDoneMsg.Text = ""
''            Traffic!edcLinkSrceHelpMsg.Text = ""
''            If igTestSystem Then
''                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
''            Else
''                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
''            End If
''            lgShellRet = Shell(sgExePath & "ExpStnFd.Exe " & slStr, 1)
''            Traffic.Enabled = False
''            Screen.MousePointer = vbDefault  'Wait
''            Do While Not igChildDone
''                DoEvents
''            Loop
''            slStr = sgDoneMsg
''            Traffic.Enabled = True
''            Traffic!edcLinkSrceHelpMsg.Text = "Ok"
''            For ilLoop = 0 To 10
''                DoEvents
''            Next ilLoop
'            imcOutline.Visible = False
''            Traffic!mnuRemoteUser.Enabled = False
'            On Error Resume Next
'            If igJobShowing(STATIONFEEDJOB) Then
'                ExpStnFd.SetFocus
'            Else
'                ExpStnFd.Show vbModeless  'This is required to activate form if previousely activated
'            End If
''            ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
''            ChDir sgCurDir
''            'Screen.MousePointer = vbDefault    'Default
''            Traffic!cmcBasicWnd.Value = False
        Case FEEDJOB
            '12/24/12: Removed FeedSpot because of out of memory
            'imcOutline.Visible = False
            'On Error Resume Next
            'If igJobShowing(FEEDJOB) Then
            '    FeedSpot.SetFocus
            'Else
            '    FeedSpot.Show vbModeless  'This is required to activate form if previousely activated
            'End If
            '12/24/12:end remove point
'        Case REPORTSJOB 'Reports- make work as if coming from contracts
''            Unload IconTraf
''            If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
''                Exit Sub
''            End If
'            imcOutline.Visible = False
'            igRptCallType = -2  'CONTRACTSJOB
''            'If Contract.plcScreen = "Orders" Then
'                ilRptType = 0   '1
''            'Else
''            '    ilRptType = 0
''            'End If
''            Screen.MousePointer = vbHourglass  'Wait
''            sgSpfStamp = "~"    'Force read
''            gSpfRead
''            igChildDone = False 'edcLinkDestDoneMsg.Text = ""
''            Traffic!edcLinkSrceHelpMsg.Text = ""
'            If igTestSystem Then
'                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType))
'            Else
'                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType))
'            End If
''            'lgShellRet = Shell(sgExePath & "RptSelCt.Exe " & slStr, 1)
''            lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
''            Traffic.Enabled = False
''            Screen.MousePointer = vbDefault  'Wait
''            Do While Not igChildDone
''                DoEvents
''            Loop
''            slStr = sgDoneMsg
''            Traffic.Enabled = True
''            Traffic!edcLinkSrceHelpMsg.Text = "Ok"
''            For ilLoop = 0 To 10
''                DoEvents
''            Next ilLoop
'
'            sgCommandStr = slStr
'            sgSpfStamp = "~"    'Force read
'            gSpfRead
'            'Traffic!mnuRemoteUser.Enabled = False
'            'so report list works the same when called from job screen buttons
'            On Error Resume Next
'            RptList.Show vbModal    'vbModeless  'This is required to activate form if previousely activated
'            'Screen.MousePointer = vbDefault    'Default
'            'Traffic!cmcBasicWnd.Value = False


            'Screen.MousePointer = vbDefault
        Case BUDGETSJOB  'Rate Card
'            ''Screen.MousePointer = vbHourGlass  'Wait
'            'RateCard.Caption = "Rate Card"  'This reference causes the form to load
'            'RateCard.Show vbModeless  'This is required to activate form if previousely activated
'            ''Screen.MousePointer = vbDefault    'Default
'            'Traffic!cmcBasicWnd.Value = False
'            igRptCallType = BUDGETSJOB    'Using contractjob instead spotsjob
'            ilRptType = 0
'            Screen.MousePointer = vbHourglass  'Wait
'            igChildDone = False 'edcLinkDestDoneMsg.Text = ""
'            Traffic!edcLinkSrceHelpMsg.Text = ""
'            If igTestSystem Then
'                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            Else
'                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(ilRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode))
'            End If
'            lgShellRet = Shell(sgExePath & "Budget.Exe " & slStr, 1)
'            Traffic.Enabled = False
'            Screen.MousePointer = vbDefault  'Wait
'            Do While Not igChildDone
'                DoEvents
'            Loop
'            slStr = sgDoneMsg
'            Traffic.Enabled = True
'            Traffic!edcLinkSrceHelpMsg.Text = "Ok"
'            For ilLoop = 0 To 10
'                DoEvents
'            Next ilLoop

            imcOutline.Visible = False
'            Traffic!mnuRemoteUser.Enabled = False
            On Error Resume Next
            'If igJobShowing(BUDGETSJOB) Then
            '    Budget.SetFocus
            'Else
            '    Budget.Show vbModeless  'This is required to activate form if previousely activated
            'End If
            gShellToBugdet Basic10
            'Screen.MousePointer = vbDefault    'Default
'            Traffic!cmcBasicWnd.Value = False
'
'
'            Screen.MousePointer = vbHourglass
'            sgMVefStamp = "~"   'Force Read
'            ilRet = gObtainVef()
'            sgVpfStamp = "~"    'Force read
'            ilRet = gVpfRead(Basic10)
'            gObtainVirtVehList
'            Screen.MousePointer = vbDefault    'Default
        Case SLSPCOMMSJOB  'Salesperson Commission
'            ''Screen.MousePointer = vbHourGlass  'Wait
'            'RateCard.Caption = "Rate Card"  'This reference causes the form to load
'            'RateCard.Show vbModeless  'This is required to activate form if previousely activated
'            ''Screen.MousePointer = vbDefault    'Default
'            'Traffic!cmcBasicWnd.Value = False
'            Unload IconTraf
            imcOutline.Visible = False
            igRptCallType = SLSPCOMMSJOB    'Using contractjob instead spotsjob
            igRptType = 0
'            Screen.MousePointer = vbHourglass  'Wait
'            igChildDone = False 'edcLinkDestDoneMsg.Text = ""
'            Traffic!edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(ilRptType)) & "\" & Trim$(str$(tgUrf(0).iCode))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(ilRptType)) & "\" & Trim$(str$(tgUrf(0).iCode))
            End If
'            'lgShellRet = Shell(sgExePath & "RptSelCt.Exe " & slStr, 1)
'            If tgSpf.sSubCompany = "Y" Then
'                lgShellRet = Shell(sgExePath & "SlspComm.Exe " & slStr, 1)
'            Else
'                lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
'            End If
'            Traffic.Enabled = False
'            Screen.MousePointer = vbDefault  'Wait
'            Do While Not igChildDone
'                DoEvents
'            Loop
'            slStr = sgDoneMsg
'            Traffic.Enabled = True
'            Traffic!edcLinkSrceHelpMsg.Text = "Ok"
'            For ilLoop = 0 To 10
'                DoEvents
'            Next ilLoop

            sgCommandStr = slStr
            'Traffic!mnuRemoteUser.Enabled = False
            On Error Resume Next
            If tgSpf.sSubCompany = "Y" Then
                If igJobShowing(SLSPCOMMSJOB) Then
                    SlspComm.SetFocus
                Else
                    SlspComm.Show vbModeless  'This is required to activate form if previousely activated
                End If
            Else
                RptList.Show vbModal    'vbModeless  'This is required to activate form if previousely activated
            End If
            'Screen.MousePointer = vbDefault    'Default
            'Traffic!cmcBasicWnd.Value = False

        Case 100
            slCommand = Command$
            If (igWinStatus(COLLECTIONSJOB) = 2) And (StrComp(Trim$(slCommand), "Debug", vbTextCompare) <> 0) Then
                mShowCredit
                ''Test if user is to see advertiser/agency with credit problems
                'slDate = Format$(gNow(), "m/d/yy")
                'If gDateValue(slDate) <> lgDateCreditShown Then
                '    gSpfRead
                '    gUnpackDate tgSpf.iRCreditDate(0), tgSpf.iRCreditDate(1), slSpfDate
                '    If gDateValue(slSpfDate) = gDateValue(slDate) Then
                '        Screen.MousePointer = vbHourGlass  'Wait
                '        igChildDone = False 'edcLinkDestDoneMsg.Text = ""
                '        Traffic!edcLinkSrceHelpMsg.Text = ""
                '        If igTestSystem Then
                '            slStr = "Traffic^Test\" & sgUserName
                '        Else
                '            slStr = "Traffic^Prod\" & sgUserName
                '        End If
                '        lgShellRet = Shell(sgExePath & "ShoCrdit.Exe " & slStr, 1)
                '        Traffic.Enabled = False
                '        Screen.MousePointer = vbDefault  'Wait
                '        Do While Not igChildDone
                '            DoEvents
                '        Loop
                '        slStr = sgDoneMsg
                '        Traffic.Enabled = True
                '        Traffic!edcLinkSrceHelpMsg.Text = "Ok"
                '        For ilLoop = 0 To 10
                '            DoEvents
                '        Next ilLoop
                '
                '        lgDateCreditShown = gDateValue(slDate)
                '    End If
                'End If
            End If
    End Select
    imcOutline.Visible = False
    imTaskIndex = -1
End Sub
Private Sub plcBasic10_Paint()
    'plcBasic10.CurrentX = cmcCSLogo.Left + (3 * cmcCSLogo.Width) / 2
    'plcBasic10.CurrentY = cmcCSLogo.Top + cmcCSLogo.Height / 3
    'plcBasic10.Print "Counterpoint Software ®"
End Sub

Private Sub mSetControls()
    Dim ilLoop As Integer
    
    '7/22/11: igWinStatus now set in gSetWinStatusFromSPF for Using Traffic, Using Proposal and Using Affilaite
    gSetWinStatusFromSPF igWinStatus()
    
    For ilLoop = RATECARDSJOB To FEEDJOB Step 1
        If ilLoop = PROPOSALSJOB Then
            'If tgSpf.sGUsePropSys = "Y" Then
                If (igWinStatus(ilLoop) = 0) Or (tgSpf.sGUsePropSys <> "Y") Then
                    cmcTask(ilLoop - RATECARDSJOB).Enabled = False
                    lacTask(ilLoop - RATECARDSJOB).Enabled = False
                    lacTask(ilLoop - RATECARDSJOB).FontBold = False
                    igWinStatus(ilLoop) = 0
                Else
                    cmcTask(ilLoop - RATECARDSJOB).Enabled = True
                End If
            'Else
            '    cmcTask(ilLoop - RATECARDSJOB).Enabled = False
            '    lacTask(ilLoop - RATECARDSJOB).Enabled = False
            'End If
        ElseIf ilLoop = STATIONFEEDJOB Then
'            If tgSpf.sSystemType <> "R" Then
'                cmcTask(ilLoop - RATECARDSJOB).Visible = True
'                lacTask(ilLoop - RATECARDSJOB).Visible = True
'                If tgSpf.sGUseAffFeed = "Y" Then
'                    If igWinStatus(ilLoop) = 0 Then
'                        cmcTask(ilLoop - RATECARDSJOB).Enabled = False
'                        lacTask(ilLoop - RATECARDSJOB).Enabled = False
'                    Else
'                        cmcTask(ilLoop - RATECARDSJOB).Enabled = True
'                    End If
'                Else
                    cmcTask(ilLoop - RATECARDSJOB).Enabled = False
                    lacTask(ilLoop - RATECARDSJOB).Enabled = False
                    lacTask(ilLoop - RATECARDSJOB).FontBold = False
'                End If
'            Else
'                cmcTask(ilLoop - RATECARDSJOB).Visible = False
'                lacTask(ilLoop - RATECARDSJOB).Visible = False
'            End If
        ElseIf ilLoop = FEEDJOB Then
            '12/24/12: Removed FeedSpot because of Out of Memory
            'If tgSpf.sSystemType = "R" Then
            '    cmcTask(ilLoop - RATECARDSJOB).Visible = True
            '    lacTask(ilLoop - RATECARDSJOB).Visible = True
            '    If igWinStatus(ilLoop) = 0 Then
            '        cmcTask(ilLoop - RATECARDSJOB).Enabled = False
            '        lacTask(ilLoop - RATECARDSJOB).Enabled = False
            '    Else
            '        cmcTask(ilLoop - RATECARDSJOB).Enabled = True
            '    End If
            'Else
                cmcTask(ilLoop - RATECARDSJOB).Visible = False
                lacTask(ilLoop - RATECARDSJOB).Visible = False
            'End If
            '12/24/12: End point
        Else
            If igWinStatus(ilLoop) = 0 Then
                cmcTask(ilLoop - RATECARDSJOB).Enabled = False
                lacTask(ilLoop - RATECARDSJOB).Enabled = False
                lacTask(ilLoop - RATECARDSJOB).FontBold = False
            Else
                cmcTask(ilLoop - RATECARDSJOB).Enabled = True
            End If
        End If
        If cmcTask(ilLoop - RATECARDSJOB).Enabled Then
            lacTask(ilLoop - RATECARDSJOB).Enabled = True
            If igWinStatus(ilLoop) = 1 Then
                lacTask(ilLoop - RATECARDSJOB).FontBold = False
                lacTask(ilLoop - RATECARDSJOB).ForeColor = vbBlack
            Else
                lacTask(ilLoop - RATECARDSJOB).ForeColor = vbBlack
            End If
        Else
            lacTask(ilLoop - RATECARDSJOB).ForeColor = vbBlack
        End If
    Next ilLoop
    'If tgSpf.sGUsePropSys = "Y" Then
        If (igWinStatus(BUDGETSJOB) = 0) Or (tgSpf.sGUsePropSys <> "Y") Then
            cmcTask(BUDGETSJOB).Enabled = False
            lacTask(BUDGETSJOB).Enabled = False
            lacTask(BUDGETSJOB).ForeColor = vbBlack
            lacTask(BUDGETSJOB).FontBold = False
            igWinStatus(ilLoop) = 0
        Else
            cmcTask(BUDGETSJOB).Enabled = True
            If igWinStatus(BUDGETSJOB) = 1 Then
                lacTask(BUDGETSJOB).FontBold = False
                lacTask(BUDGETSJOB).ForeColor = vbBlack
            Else
                lacTask(BUDGETSJOB).ForeColor = vbBlack
            End If
        End If
    'Else
    '    cmcTask(BUDGETSJOB).Enabled = False
    '    lacTask(BUDGETSJOB).Enabled = False
    '    lacTask(BUDGETSJOB).ForeColor = vbBlack
    'End If
End Sub
Private Sub mDisallowToolTip(X As Single, Y As Single)
    Dim Ctrl As Control
    Dim ilLoop As Integer
    
    For Each Ctrl In Basic10.Controls
        If Ctrl.Enabled = False Then
            If TypeOf Ctrl Is CommandButton Then
                If (X >= Ctrl.Left) And (X <= Ctrl.Left + Ctrl.Width) Then
                    If (Y >= Ctrl.Top) And (Y <= Ctrl.Top + Ctrl.height) Then
                        plcBasic10.ToolTipText = "Disallowed"
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next Ctrl
    plcBasic10.ToolTipText = ""

End Sub

Private Sub tmcFilemaker_Timer()
    tmcFilemaker.Enabled = False
    FileMakerToOrder.Show vbModal
End Sub

Private Sub tmcFilemakerToOrder_Timer()
    tmcFilemakerToOrder.Enabled = False
    sgCntrScreen = "Orders"
    igAlertCntrStatus = -1
    Contract.Show vbModeless  'This is required to activate form if previousely activated
End Sub
