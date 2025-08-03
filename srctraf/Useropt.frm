VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form UserOpt 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7470
   ClientLeft      =   -19725
   ClientTop       =   5985
   ClientWidth     =   10425
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7470
   ScaleWidth      =   10425
   Begin VB.ComboBox cbcSelect 
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
      Left            =   5805
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   4440
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   9690
      Top             =   4635
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".Txt"
      Filter          =   "*.Txt|*.Txt|*.Doc|*.Doc|*.Asc|*.Asc"
      FilterIndex     =   1
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.CommandButton cmcRptSet 
      Appearance      =   0  'Flat
      Caption         =   "Report Se&t"
      Height          =   285
      Left            =   7065
      TabIndex        =   49
      Top             =   7110
      Width           =   1050
   End
   Begin VB.CommandButton cmcPassword 
      Appearance      =   0  'Flat
      Caption         =   "&Password"
      Height          =   285
      Left            =   5985
      TabIndex        =   48
      Top             =   7110
      Width           =   1050
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
      Left            =   9120
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   4725
      Visible         =   0   'False
      Width           =   525
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
      Height          =   75
      Left            =   270
      ScaleHeight     =   75
      ScaleWidth      =   60
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   6585
      Width           =   60
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   30
      ScaleHeight     =   270
      ScaleWidth      =   1140
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   0
      Width           =   1140
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   1665
      TabIndex        =   44
      Top             =   7110
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2745
      TabIndex        =   45
      Top             =   7110
      Width           =   1050
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   3825
      TabIndex        =   46
      Top             =   7110
      Width           =   1050
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
      Height          =   285
      Left            =   4905
      TabIndex        =   47
      Top             =   7110
      Width           =   1050
   End
   Begin VB.PictureBox plcMain 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   210
      ScaleHeight     =   4050
      ScaleWidth      =   9960
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2820
      Width           =   10020
      Begin VB.Frame frcGeneral 
         Caption         =   "Selected Fields"
         ForeColor       =   &H00000000&
         Height          =   3240
         Left            =   2385
         TabIndex        =   42
         Top             =   675
         Visible         =   0   'False
         Width           =   6690
         Begin VB.VScrollBar vbcSelFields 
            Height          =   1245
            Left            =   6480
            Max             =   3
            TabIndex        =   59
            Top             =   1860
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Allowed to Create New Contracts"
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
            Height          =   225
            Index           =   22
            Left            =   75
            TabIndex        =   162
            Top             =   2895
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Allowed to Remove Attachments"
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
            Height          =   225
            Index           =   21
            Left            =   3375
            TabIndex        =   161
            Top             =   2655
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Allowed to Add Attachments"
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
            Height          =   225
            Index           =   20
            Left            =   75
            TabIndex        =   150
            Top             =   2655
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Advanced Avails Allowed"
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
            Height          =   225
            Index           =   19
            Left            =   3375
            TabIndex        =   160
            Top             =   2415
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Allowed to Change Acquisition Cost"
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
            Height          =   225
            Index           =   18
            Left            =   75
            TabIndex        =   149
            Top             =   2415
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Set Contract Verification"
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
            Height          =   225
            Index           =   17
            Left            =   3375
            TabIndex        =   159
            Top             =   2175
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Activity Log"
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
            Height          =   225
            Index           =   16
            Left            =   75
            TabIndex        =   148
            Top             =   2175
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Allow Today's Date Change"
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
            Height          =   225
            Index           =   15
            Left            =   3375
            TabIndex        =   158
            Top             =   1935
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Allow Display of Final Invoices"
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
            Height          =   225
            Index           =   14
            Left            =   75
            TabIndex        =   147
            Top             =   1935
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Change Billed Contract Prices"
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
            Height          =   225
            Index           =   13
            Left            =   3375
            TabIndex        =   157
            Top             =   1695
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Flight Button"
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
            Height          =   225
            Index           =   12
            Left            =   75
            TabIndex        =   146
            Top             =   1695
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Change Contract Prices"
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
            Height          =   225
            Index           =   11
            Left            =   3375
            TabIndex        =   156
            Top             =   1455
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Access Copy Regions"
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
            Height          =   225
            Index           =   10
            Left            =   3375
            TabIndex        =   155
            Top             =   1215
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Change Advertiser/Agencies Credit Rating"
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
            Height          =   225
            Index           =   9
            Left            =   3375
            TabIndex        =   154
            Top             =   975
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Reference Reservation Orders"
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
            Height          =   225
            Index           =   8
            Left            =   3375
            TabIndex        =   153
            Top             =   735
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Change Contracts in Past with Unbilled Spots"
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
            Height          =   225
            Index           =   7
            Left            =   3375
            TabIndex        =   152
            Top             =   495
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Change Billed Spots in Post Log"
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
            Height          =   225
            Index           =   6
            Left            =   3375
            TabIndex        =   151
            Top             =   255
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hide Spots"
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
            Height          =   225
            Index           =   5
            Left            =   75
            TabIndex        =   145
            Top             =   1455
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Merge Operation on Screen"
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
            Height          =   225
            Index           =   4
            Left            =   75
            TabIndex        =   144
            Top             =   1215
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Payment Rating in Advertiser and Agency"
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
            Height          =   225
            Index           =   3
            Left            =   75
            TabIndex        =   143
            Top             =   975
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Credit Restrictions in Advertiser and Agency"
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
            Height          =   225
            Index           =   2
            Left            =   75
            TabIndex        =   142
            Top             =   735
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Change Fill Spot Invoice Status in Post Log"
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
            Height          =   225
            Index           =   1
            Left            =   75
            TabIndex        =   141
            Top             =   495
            Width           =   3210
         End
         Begin VB.Label plcSelectedFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Today's Rate Card Grid Level"
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
            Height          =   225
            Index           =   0
            Left            =   75
            TabIndex        =   140
            Top             =   255
            Width           =   3210
         End
      End
      Begin VB.Frame frcAlerts 
         Caption         =   "Alerts"
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
         Height          =   4110
         Left            =   2880
         TabIndex        =   61
         Top             =   -75
         Visible         =   0   'False
         Width           =   6090
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Receive Impression Import Email Notifications"
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
            Height          =   225
            Index           =   15
            Left            =   405
            TabIndex        =   178
            Top             =   3810
            Width           =   5295
         End
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Receive Digital Contract Email Notifications"
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
            Height          =   225
            Index           =   14
            Left            =   405
            TabIndex        =   177
            Top             =   3570
            Width           =   5295
         End
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Show Rep-Net Messages"
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
            Height          =   225
            Index           =   13
            Left            =   405
            TabIndex        =   176
            Top             =   3330
            Width           =   5295
         End
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Allowed to Initiate Shutdown"
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
            Height          =   225
            Index           =   12
            Left            =   405
            TabIndex        =   175
            Top             =   3090
            Width           =   5295
         End
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Show Contracts Affecting Spots Prior to Last Log Date"
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
            Height          =   225
            Index           =   11
            Left            =   405
            TabIndex        =   174
            Top             =   2850
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Show Contracts with Credit Exceeded"
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
            Height          =   225
            Index           =   10
            Left            =   405
            TabIndex        =   173
            Top             =   2610
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Show Proposals with Credit Check Denied"
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
            Height          =   225
            Index           =   9
            Left            =   405
            TabIndex        =   172
            Top             =   2370
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Show Proposals with Credit Check Approved"
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
            Height          =   225
            Index           =   8
            Left            =   405
            TabIndex        =   171
            Top             =   2130
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Show proposals with Insufficient Avails"
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
            Height          =   225
            Index           =   7
            Left            =   405
            TabIndex        =   170
            Top             =   1890
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Show Proposals Affected by Change in Research Data"
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
            Height          =   225
            Index           =   6
            Left            =   405
            TabIndex        =   169
            Top             =   1650
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Show Proposals Affected by Rate Card Change"
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
            Height          =   225
            Index           =   5
            Left            =   405
            TabIndex        =   168
            Top             =   1410
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Show Holds on Alert Screen"
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
            Height          =   225
            Index           =   4
            Left            =   405
            TabIndex        =   167
            Top             =   1170
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Show Contracts Which Require Scheduling"
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
            Height          =   225
            Index           =   3
            Left            =   405
            TabIndex        =   166
            Top             =   930
            Width           =   5295
         End
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Show Proposal When Set to Complete"
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
            Height          =   225
            Index           =   2
            Left            =   405
            TabIndex        =   165
            Top             =   690
            Width           =   5295
         End
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Show Proposal When Set to Unapproved"
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
            Height          =   225
            Index           =   1
            Left            =   405
            TabIndex        =   164
            Top             =   450
            Width           =   5295
         End
         Begin VB.Label plcAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Show Reprint Log/CP"
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
            Height          =   225
            Index           =   0
            Left            =   405
            TabIndex        =   163
            Top             =   210
            Width           =   5295
         End
      End
      Begin VB.Frame frcSet 
         Caption         =   "Set Jobs To"
         Height          =   720
         Left            =   135
         TabIndex        =   200
         Top             =   2730
         Width           =   1965
         Begin VB.OptionButton rbcSet 
            Caption         =   "As Is"
            Height          =   225
            Index           =   3
            Left            =   1005
            TabIndex        =   204
            Top             =   450
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton rbcSet 
            Caption         =   "Red"
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   203
            Top             =   450
            Width           =   825
         End
         Begin VB.OptionButton rbcSet 
            Caption         =   "Yellow"
            Height          =   225
            Index           =   1
            Left            =   1005
            TabIndex        =   202
            Top             =   225
            Width           =   900
         End
         Begin VB.OptionButton rbcSet 
            Caption         =   "Green"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   201
            Top             =   225
            Width           =   810
         End
      End
      Begin VB.Frame frcPDF 
         Caption         =   "PDF Setup"
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
         Height          =   2760
         Left            =   7065
         TabIndex        =   83
         Top             =   3900
         Visible         =   0   'False
         Width           =   6120
         Begin VB.TextBox edcPDF 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
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
            Index           =   0
            Left            =   3330
            MaxLength       =   1
            TabIndex        =   87
            Top             =   570
            Width           =   540
         End
         Begin VB.TextBox edcPDF 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
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
            Left            =   5460
            MaxLength       =   2
            TabIndex        =   89
            Top             =   870
            Width           =   540
         End
         Begin VB.TextBox edcPDF 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
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
            Index           =   2
            Left            =   3315
            MaxLength       =   1
            TabIndex        =   91
            Top             =   1170
            Width           =   540
         End
         Begin VB.TextBox edcPDF 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
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
            Index           =   3
            Left            =   5460
            MaxLength       =   2
            TabIndex        =   93
            Top             =   1470
            Width           =   540
         End
         Begin VB.CommandButton cmcSetup 
            Appearance      =   0  'Flat
            Caption         =   "View Printer Setup"
            Height          =   285
            Left            =   210
            TabIndex        =   96
            Top             =   2310
            Width           =   1770
         End
         Begin VB.CommandButton cmcTestPDF 
            Appearance      =   0  'Flat
            Caption         =   "Test Switch To PDF"
            Height          =   285
            Left            =   2130
            TabIndex        =   97
            Top             =   2310
            Width           =   1770
         End
         Begin VB.CommandButton cmcTestDefault 
            Appearance      =   0  'Flat
            Caption         =   "Test Switch To Default "
            Height          =   285
            Left            =   4035
            TabIndex        =   98
            Top             =   2310
            Width           =   1995
         End
         Begin VB.TextBox edcPDF 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
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
            Index           =   4
            Left            =   495
            MaxLength       =   1
            TabIndex        =   85
            Top             =   270
            Width           =   375
         End
         Begin VB.TextBox edcPDF 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
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
            Index           =   5
            Left            =   4080
            MaxLength       =   2
            TabIndex        =   95
            Top             =   1770
            Width           =   540
         End
         Begin VB.Label lacPDF 
            Appearance      =   0  'Flat
            Caption         =   "First Letter of PDF Printer Name"
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
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   86
            Top             =   600
            Width           =   3135
         End
         Begin VB.Label lacPDF 
            Appearance      =   0  'Flat
            Caption         =   "# of Times First Letter Entered to Select PDF Printer Name"
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
            Height          =   210
            Index           =   1
            Left            =   150
            TabIndex        =   88
            Top             =   900
            Width           =   5175
         End
         Begin VB.Label lacPDF 
            Appearance      =   0  'Flat
            Caption         =   "First Letter of Default Printer Name"
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
            Height          =   210
            Index           =   2
            Left            =   150
            TabIndex        =   90
            Top             =   1200
            Width           =   2985
         End
         Begin VB.Label lacPDF 
            Appearance      =   0  'Flat
            Caption         =   "# of Times First Letter Entered to Select Default Printer Name"
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
            Height          =   210
            Index           =   3
            Left            =   150
            TabIndex        =   92
            Top             =   1500
            Width           =   5340
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            X1              =   105
            X2              =   5985
            Y1              =   2175
            Y2              =   2175
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FFFFFF&
            X1              =   105
            X2              =   5985
            Y1              =   2190
            Y2              =   2190
         End
         Begin VB.Label lacPDF 
            Appearance      =   0  'Flat
            Caption         =   "Alt          to Select Printer Name Drop Down Box"
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
            Height          =   210
            Index           =   4
            Left            =   150
            TabIndex        =   84
            Top             =   300
            Width           =   4230
         End
         Begin VB.Label lacPDF 
            Appearance      =   0  'Flat
            Caption         =   "# Enter Keys Required to Activate Ok button"
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
            Height          =   210
            Index           =   5
            Left            =   150
            TabIndex        =   94
            Top             =   1800
            Width           =   3960
         End
      End
      Begin VB.ComboBox cbcModel 
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
         Left            =   5325
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   90
         Width           =   4410
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
         ForeColor       =   &H00000000&
         Height          =   2340
         Left            =   135
         TabIndex        =   32
         Top             =   375
         Width           =   1980
         Begin VB.OptionButton rbcSelect 
            Caption         =   "PDF Setup"
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
            Height          =   270
            Index           =   6
            Left            =   120
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1920
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.OptionButton rbcSelect 
            Caption         =   "Contract Types"
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
            Height          =   270
            Index           =   5
            Left            =   120
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1635
            Width           =   1635
         End
         Begin VB.OptionButton rbcSelect 
            Caption         =   "Contract Status's"
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
            Height          =   270
            Index           =   4
            Left            =   120
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1365
            Width           =   1785
         End
         Begin VB.OptionButton rbcSelect 
            Caption         =   "Alerts"
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
            Height          =   270
            Index           =   3
            Left            =   120
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1095
            Width           =   1635
         End
         Begin VB.OptionButton rbcSelect 
            Caption         =   "Selected Fields"
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
            Height          =   270
            Index           =   2
            Left            =   120
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   825
            Width           =   1635
         End
         Begin VB.OptionButton rbcSelect 
            Caption         =   "Lists"
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
            Height          =   270
            Index           =   1
            Left            =   120
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   555
            Width           =   1635
         End
         Begin VB.OptionButton rbcSelect 
            Caption         =   "Jobs"
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
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   285
            Value           =   -1  'True
            Width           =   1635
         End
      End
      Begin VB.ListBox lbcModel 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   345
         Sorted          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Frame frcJobs 
         Caption         =   "Jobs"
         ForeColor       =   &H00000000&
         Height          =   2715
         Left            =   7530
         TabIndex        =   41
         Top             =   3345
         Width           =   6420
         Begin VB.Frame frcSports 
            BorderStyle     =   0  'None
            Height          =   510
            Left            =   225
            TabIndex        =   210
            Top             =   2130
            Width           =   4275
            Begin VB.OptionButton rbcSports 
               Caption         =   "No"
               Height          =   210
               Index           =   1
               Left            =   2880
               TabIndex        =   212
               Top             =   210
               Width           =   645
            End
            Begin VB.OptionButton rbcSports 
               Caption         =   "Yes"
               Height          =   210
               Index           =   0
               Left            =   2160
               TabIndex        =   211
               Top             =   210
               Width           =   645
            End
            Begin VB.Label lacSportsOr 
               Caption         =   "  Or"
               Height          =   195
               Left            =   1665
               TabIndex        =   213
               Top             =   -15
               Width           =   345
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00FFFFFF&
               X1              =   30
               X2              =   4200
               Y1              =   75
               Y2              =   75
            End
            Begin VB.Line Line8 
               BorderColor     =   &H00808080&
               X1              =   30
               X2              =   4200
               Y1              =   60
               Y2              =   60
            End
            Begin VB.Label lacSports 
               Caption         =   "Sports Proposal Only"
               Height          =   240
               Left            =   300
               TabIndex        =   214
               Top             =   195
               Width           =   1740
            End
         End
         Begin VB.Frame frcLiveLog 
            BorderStyle     =   0  'None
            Height          =   510
            Left            =   225
            TabIndex        =   205
            Top             =   1635
            Width           =   4275
            Begin VB.OptionButton rbcLiveLog 
               Caption         =   "Yes"
               Height          =   210
               Index           =   0
               Left            =   1890
               TabIndex        =   207
               Top             =   210
               Width           =   645
            End
            Begin VB.OptionButton rbcLiveLog 
               Caption         =   "No"
               Height          =   210
               Index           =   1
               Left            =   2610
               TabIndex        =   206
               Top             =   210
               Width           =   645
            End
            Begin VB.Label lacLiveLogOr 
               Caption         =   "  Or"
               Height          =   195
               Left            =   1665
               TabIndex        =   208
               Top             =   -15
               Width           =   345
            End
            Begin VB.Label lacLiveLog 
               Caption         =   "Live Log Only"
               Height          =   240
               Left            =   570
               TabIndex        =   209
               Top             =   195
               Width           =   1245
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00808080&
               X1              =   30
               X2              =   4200
               Y1              =   60
               Y2              =   60
            End
            Begin VB.Line Line6 
               BorderColor     =   &H00FFFFFF&
               X1              =   30
               X2              =   4200
               Y1              =   75
               Y2              =   75
            End
         End
         Begin VB.Label plcJobs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Feed"
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
            Height          =   225
            Index           =   13
            Left            =   4920
            TabIndex        =   190
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label plcJobs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Commission"
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
            Height          =   225
            Index           =   11
            Left            =   3750
            TabIndex        =   199
            Top             =   1185
            Width           =   1140
         End
         Begin VB.Label plcJobs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Collections"
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
            Height          =   225
            Index           =   9
            Left            =   2580
            TabIndex        =   198
            Top             =   1185
            Width           =   1140
         End
         Begin VB.Label plcJobs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Invoices"
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
            Height          =   225
            Index           =   8
            Left            =   1410
            TabIndex        =   197
            Top             =   1185
            Width           =   1140
         End
         Begin VB.Label plcJobs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Post Log"
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
            Height          =   225
            Index           =   7
            Left            =   240
            TabIndex        =   196
            Top             =   1185
            Width           =   1140
         End
         Begin VB.Label plcJobs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Station Feed"
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
            Height          =   225
            Index           =   12
            Left            =   4920
            TabIndex        =   195
            Top             =   750
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label plcJobs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Logs"
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
            Height          =   225
            Index           =   6
            Left            =   3750
            TabIndex        =   194
            Top             =   750
            Width           =   1140
         End
         Begin VB.Label plcJobs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Copy"
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
            Height          =   225
            Index           =   5
            Left            =   2580
            TabIndex        =   193
            Top             =   750
            Width           =   1140
         End
         Begin VB.Label plcJobs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Spots"
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
            Height          =   225
            Index           =   4
            Left            =   1410
            TabIndex        =   192
            Top             =   750
            Width           =   1140
         End
         Begin VB.Label plcJobs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Programming"
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
            Height          =   225
            Index           =   3
            Left            =   240
            TabIndex        =   191
            Top             =   750
            Width           =   1140
         End
         Begin VB.Label plcJobs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Orders"
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
            Height          =   225
            Index           =   2
            Left            =   3750
            TabIndex        =   189
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label plcJobs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Proposals"
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
            Height          =   225
            Index           =   1
            Left            =   2580
            TabIndex        =   188
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label plcJobs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Rate Card"
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
            Height          =   225
            Index           =   0
            Left            =   1410
            TabIndex        =   187
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label plcJobs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Budgets"
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
            Height          =   225
            Index           =   10
            Left            =   240
            TabIndex        =   186
            Top             =   300
            Width           =   1140
         End
      End
      Begin VB.Frame frcTypes 
         Caption         =   "Contract Types"
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
         Height          =   2355
         Left            =   8145
         TabIndex        =   40
         Top             =   2910
         Visible         =   0   'False
         Width           =   6075
         Begin VB.Label plcTypes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Allow to Select Programmatic Buys"
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
            Height          =   225
            Index           =   6
            Left            =   390
            TabIndex        =   185
            Top             =   1755
            Width           =   5325
         End
         Begin VB.Label plcTypes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Allowed to Select Promos"
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
            Height          =   225
            Index           =   5
            Left            =   390
            TabIndex        =   184
            Top             =   1515
            Width           =   5325
         End
         Begin VB.Label plcTypes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Allowed to Select PSAs"
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
            Height          =   225
            Index           =   4
            Left            =   390
            TabIndex        =   183
            Top             =   1275
            Width           =   5325
         End
         Begin VB.Label plcTypes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Allowed to Select Per Inquiries"
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
            Height          =   225
            Index           =   3
            Left            =   390
            TabIndex        =   182
            Top             =   1035
            Width           =   5325
         End
         Begin VB.Label plcTypes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Allowed to Select Direct Responses"
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
            Height          =   225
            Index           =   2
            Left            =   390
            TabIndex        =   181
            Top             =   795
            Width           =   5325
         End
         Begin VB.Label plcTypes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Allow to Select Remnants"
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
            Height          =   225
            Index           =   1
            Left            =   390
            TabIndex        =   180
            Top             =   555
            Width           =   5325
         End
         Begin VB.Label plcTypes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Allowed to Select Reservations"
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
            Height          =   225
            Index           =   0
            Left            =   390
            TabIndex        =   179
            Top             =   315
            Width           =   5325
         End
      End
      Begin VB.Frame frcStatus 
         Caption         =   "Contract Status's"
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
         Height          =   2295
         Left            =   8355
         TabIndex        =   60
         Top             =   2730
         Visible         =   0   'False
         Width           =   6120
         Begin VB.PictureBox plcReviseCntr 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   105
            ScaleHeight     =   255
            ScaleWidth      =   5655
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   1935
            Width           =   5655
            Begin VB.OptionButton rbcReviseCntr 
               Caption         =   "No"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   4410
               TabIndex        =   99
               Top             =   0
               Width           =   510
            End
            Begin VB.OptionButton rbcReviseCntr 
               Caption         =   "Yes"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   3720
               TabIndex        =   82
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.PictureBox plcStatus 
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
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   4
            Left            =   105
            ScaleHeight     =   240
            ScaleWidth      =   5280
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   1395
            Width           =   5280
            Begin VB.CheckBox ckcHStatus 
               Caption         =   "Order"
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
               Left            =   1395
               TabIndex        =   80
               Top             =   0
               Width           =   975
            End
         End
         Begin VB.PictureBox plcStatus 
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
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   3
            Left            =   105
            ScaleHeight     =   240
            ScaleWidth      =   5250
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   1125
            Width           =   5250
            Begin VB.CheckBox ckcDStatus 
               Caption         =   "Working"
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
               Left            =   1395
               TabIndex        =   79
               Top             =   0
               Width           =   1035
            End
         End
         Begin VB.PictureBox plcStatus 
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
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   2
            Left            =   105
            ScaleHeight     =   240
            ScaleWidth      =   5850
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   855
            Width           =   5850
            Begin VB.CheckBox ckcIStatus 
               Caption         =   "Order"
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
               Index           =   3
               Left            =   4950
               TabIndex        =   78
               Top             =   0
               Width           =   795
            End
            Begin VB.CheckBox ckcIStatus 
               Caption         =   "Hold"
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
               Index           =   2
               Left            =   4155
               TabIndex        =   77
               Top             =   0
               Width           =   795
            End
            Begin VB.CheckBox ckcIStatus 
               Caption         =   "Completed"
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
               Left            =   2850
               TabIndex        =   76
               Top             =   0
               Width           =   1275
            End
            Begin VB.CheckBox ckcIStatus 
               Caption         =   "Rejected"
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
               Left            =   1395
               TabIndex        =   75
               Top             =   0
               Width           =   1110
            End
         End
         Begin VB.PictureBox plcStatus 
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
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   105
            ScaleHeight     =   240
            ScaleWidth      =   5895
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   585
            Width           =   5895
            Begin VB.CheckBox ckcCStatus 
               Caption         =   "Rejected"
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
               Left            =   2850
               TabIndex        =   72
               Top             =   0
               Width           =   1125
            End
            Begin VB.CheckBox ckcCStatus 
               Caption         =   "Order"
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
               Index           =   3
               Left            =   4950
               TabIndex        =   74
               Top             =   0
               Width           =   810
            End
            Begin VB.CheckBox ckcCStatus 
               Caption         =   "Hold"
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
               Index           =   2
               Left            =   4155
               TabIndex        =   73
               Top             =   0
               Width           =   735
            End
            Begin VB.CheckBox ckcCStatus 
               Caption         =   "Unapproved"
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
               Left            =   1395
               TabIndex        =   71
               Top             =   0
               Width           =   1410
            End
         End
         Begin VB.PictureBox plcStatus 
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
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   105
            ScaleHeight     =   240
            ScaleWidth      =   5865
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   315
            Width           =   5865
            Begin VB.CheckBox ckcWStatus 
               Caption         =   "Order"
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
               Index           =   3
               Left            =   4950
               TabIndex        =   70
               Top             =   0
               Width           =   810
            End
            Begin VB.CheckBox ckcWStatus 
               Caption         =   "Hold"
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
               Index           =   2
               Left            =   4155
               TabIndex        =   69
               Top             =   0
               Width           =   855
            End
            Begin VB.CheckBox ckcWStatus 
               Caption         =   "Completed"
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
               Left            =   2850
               TabIndex        =   68
               Top             =   0
               Width           =   1230
            End
            Begin VB.CheckBox ckcWStatus 
               Caption         =   "Rejected"
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
               Left            =   1395
               TabIndex        =   67
               Top             =   0
               Width           =   1155
            End
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            X1              =   120
            X2              =   6000
            Y1              =   1815
            Y2              =   1815
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FFFFFF&
            X1              =   120
            X2              =   6000
            Y1              =   1830
            Y2              =   1830
         End
      End
      Begin VB.Frame frcLists 
         Caption         =   "Lists"
         ForeColor       =   &H00000000&
         Height          =   3135
         Left            =   2175
         TabIndex        =   43
         Top             =   330
         Visible         =   0   'False
         Width           =   7410
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ad Server Items"
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
            Height          =   225
            Index           =   36
            Left            =   3720
            TabIndex        =   216
            Top             =   945
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Split Net"
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
            Height          =   225
            Index           =   35
            Left            =   6030
            TabIndex        =   124
            Top             =   1185
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tax Table"
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
            Height          =   225
            Index           =   34
            Left            =   4875
            TabIndex        =   138
            Top             =   2040
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Feed Names"
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
            Height          =   225
            Index           =   33
            Left            =   3720
            TabIndex        =   107
            Top             =   285
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "User"
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
            Height          =   225
            Index           =   24
            Left            =   4515
            TabIndex        =   132
            Top             =   2775
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Site"
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
            Height          =   225
            Index           =   22
            Left            =   3345
            TabIndex        =   133
            Top             =   2775
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Announcers"
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
            Height          =   225
            Index           =   21
            Left            =   255
            TabIndex        =   130
            Top             =   2760
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Media Def."
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
            Height          =   225
            Index           =   20
            Left            =   1410
            TabIndex        =   131
            Top             =   2760
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tran Type"
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
            Height          =   225
            Index           =   26
            Left            =   255
            TabIndex        =   139
            Top             =   2280
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "EDI Services"
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
            Height          =   225
            Index           =   14
            Left            =   255
            TabIndex        =   134
            Top             =   2040
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Lock Boxes"
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
            Height          =   225
            Index           =   13
            Left            =   2565
            TabIndex        =   136
            Top             =   2040
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Invoice Sorts"
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
            Height          =   225
            Index           =   12
            Left            =   1410
            TabIndex        =   135
            Top             =   2040
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "NTR Types"
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
            Height          =   225
            Index           =   11
            Left            =   3720
            TabIndex        =   137
            Top             =   2040
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Missed/Cancel Reason"
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
            Height          =   225
            Index           =   10
            Left            =   4875
            TabIndex        =   108
            Top             =   285
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Genre Names"
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
            Height          =   225
            Index           =   19
            Left            =   4875
            TabIndex        =   129
            Top             =   1605
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Feed Types"
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
            Height          =   225
            Index           =   18
            Left            =   3720
            TabIndex        =   128
            Top             =   1605
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Avail Names"
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
            Height          =   225
            Index           =   17
            Left            =   255
            TabIndex        =   125
            Top             =   1605
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Event Names"
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
            Height          =   225
            Index           =   16
            Left            =   1410
            TabIndex        =   126
            Top             =   1605
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Event Types"
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
            Height          =   225
            Index           =   15
            Left            =   2565
            TabIndex        =   127
            Top             =   1605
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Boilerplate"
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
            Height          =   225
            Index           =   23
            Left            =   255
            TabIndex        =   114
            Top             =   945
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Revenus Sets"
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
            Height          =   225
            Index           =   4
            Left            =   6030
            TabIndex        =   118
            Top             =   945
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Competitors"
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
            Height          =   225
            Index           =   32
            Left            =   1410
            TabIndex        =   115
            Top             =   945
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Custom Demo"
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
            Height          =   225
            Index           =   31
            Left            =   2565
            TabIndex        =   116
            Top             =   945
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Research"
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
            Height          =   225
            Index           =   30
            Left            =   4875
            TabIndex        =   117
            Top             =   945
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Salespeople"
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
            Height          =   225
            Index           =   9
            Left            =   255
            TabIndex        =   119
            Top             =   1185
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sales Teams"
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
            Height          =   225
            Index           =   8
            Left            =   4875
            TabIndex        =   123
            Top             =   1185
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sales Offices"
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
            Height          =   225
            Index           =   7
            Left            =   1410
            TabIndex        =   120
            Top             =   1185
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sales Regions"
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
            Height          =   225
            Index           =   6
            Left            =   2565
            TabIndex        =   121
            Top             =   1185
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sales Sources"
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
            Height          =   225
            Index           =   5
            Left            =   3720
            TabIndex        =   122
            Top             =   1185
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Exclusions"
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
            Height          =   225
            Index           =   25
            Left            =   2565
            TabIndex        =   111
            Top             =   525
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Potential Codes"
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
            Height          =   225
            Index           =   29
            Left            =   255
            TabIndex        =   109
            Top             =   525
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Bus. Categories"
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
            Height          =   225
            Index           =   28
            Left            =   2565
            TabIndex        =   106
            Top             =   285
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Product Protect"
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
            Height          =   225
            Index           =   3
            Left            =   1410
            TabIndex        =   110
            Top             =   525
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Advertisers"
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
            Height          =   225
            Index           =   2
            Left            =   255
            TabIndex        =   104
            Top             =   285
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Agencies"
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
            Height          =   225
            Index           =   1
            Left            =   1410
            TabIndex        =   105
            Top             =   285
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vehicle Groups"
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
            Height          =   225
            Index           =   27
            Left            =   4875
            TabIndex        =   113
            Top             =   525
            Width           =   1140
         End
         Begin VB.Label plcLists 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vehicles"
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
            Height          =   225
            Index           =   0
            Left            =   3720
            TabIndex        =   112
            Top             =   525
            Width           =   1140
         End
      End
      Begin VB.Label lacModel 
         Appearance      =   0  'Flat
         Caption         =   "Model from User/Vehicle"
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
         Left            =   3090
         TabIndex        =   30
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox plcName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   45
      ScaleHeight     =   2325
      ScaleWidth      =   10125
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   420
      Width           =   10185
      Begin VB.CommandButton cmcErasePassword 
         Caption         =   "Erase Password"
         Height          =   210
         Left            =   9330
         TabIndex        =   215
         Top             =   2040
         Visible         =   0   'False
         Width           =   600
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
         Left            =   1440
         TabIndex        =   13
         Top             =   1995
         Width           =   7680
      End
      Begin VB.TextBox edcCity 
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1650
         Width           =   4455
      End
      Begin VB.TextBox edcPhone 
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
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   9
         Top             =   1305
         Width           =   2685
      End
      Begin VB.ComboBox cbcVehicle 
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
         Left            =   7350
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   450
         Width           =   2610
      End
      Begin VB.ComboBox cbcHub 
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
         ItemData        =   "Useropt.frx":0000
         Left            =   4725
         List            =   "Useropt.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   960
         Width           =   1275
      End
      Begin VB.CheckBox ckcBlockRU 
         Alignment       =   1  'Right Justify
         Caption         =   "Block RU  "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4350
         TabIndex        =   20
         Top             =   1350
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox edcGroupNo 
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
         MaxLength       =   4
         TabIndex        =   17
         Top             =   630
         Width           =   690
      End
      Begin VB.TextBox edcRemoteID 
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
         MaxLength       =   4
         TabIndex        =   15
         Top             =   300
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.PictureBox plcState 
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
         Height          =   225
         Left            =   3255
         ScaleHeight     =   225
         ScaleWidth      =   2040
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   15
         Visible         =   0   'False
         Width           =   2040
         Begin VB.OptionButton rbcState 
            Caption         =   "Dormant"
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
            Left            =   915
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   0
            Width           =   1035
         End
         Begin VB.OptionButton rbcState 
            Caption         =   "Active"
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
            Left            =   30
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   0
            Width           =   870
         End
      End
      Begin VB.ComboBox cbcRptSet 
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
         Left            =   7350
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   120
         Width           =   2610
      End
      Begin VB.TextBox edcRept 
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
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   5
         Top             =   630
         Width           =   2685
      End
      Begin VB.ComboBox cbcSalesperson 
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
         Left            =   7350
         TabIndex        =   28
         Top             =   1110
         Width           =   2610
      End
      Begin VB.ComboBox cbcDefVeh 
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
         Left            =   7350
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   780
         Width           =   2610
      End
      Begin VB.TextBox edcPassword 
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
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   7
         Top             =   960
         Width           =   2685
      End
      Begin VB.TextBox edcName 
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
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   3
         Top             =   300
         Width           =   2685
      End
      Begin VB.Label lacEMail 
         Appearance      =   0  'Flat
         Caption         =   "E-Mail Address"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   90
         TabIndex        =   12
         Top             =   2025
         Width           =   1350
      End
      Begin VB.Label lacCity 
         Appearance      =   0  'Flat
         Caption         =   "City"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   90
         TabIndex        =   10
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label lacPhone 
         Appearance      =   0  'Flat
         Caption         =   "Phone"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   90
         TabIndex        =   8
         Top             =   1335
         Width           =   900
      End
      Begin VB.Label lacAVehicle 
         Appearance      =   0  'Flat
         Caption         =   "Vehicle Access"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6060
         TabIndex        =   23
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label lacHub 
         Appearance      =   0  'Flat
         Caption         =   "Hub"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4350
         TabIndex        =   18
         Top             =   975
         Width           =   375
      End
      Begin VB.Label lacNRemoteID 
         Appearance      =   0  'Flat
         Caption         =   "Remote ID"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4350
         TabIndex        =   14
         Top             =   330
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lacGroupNo 
         Appearance      =   0  'Flat
         Caption         =   "Group#"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4350
         TabIndex        =   16
         Top             =   645
         Width           =   900
      End
      Begin VB.Label lacRptSet 
         Appearance      =   0  'Flat
         Caption         =   "Report Set"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6060
         TabIndex        =   21
         Top             =   150
         Width           =   1245
      End
      Begin VB.Label lacNRept 
         Appearance      =   0  'Flat
         Caption         =   "Name on Report"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   90
         TabIndex        =   4
         Top             =   660
         Width           =   1350
      End
      Begin VB.Label lacNSalesperson 
         Appearance      =   0  'Flat
         Caption         =   "Salesperson"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6060
         TabIndex        =   27
         Top             =   1140
         Width           =   1245
      End
      Begin VB.Label lacNVehicle 
         Appearance      =   0  'Flat
         Caption         =   "default Vehicle"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6060
         TabIndex        =   25
         Top             =   810
         Width           =   1245
      End
      Begin VB.Label lacNPassword 
         Appearance      =   0  'Flat
         Caption         =   "Password"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   90
         TabIndex        =   6
         Top             =   990
         Width           =   900
      End
      Begin VB.Label lacNName 
         Appearance      =   0  'Flat
         Caption         =   "Sign on Name"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   90
         TabIndex        =   2
         Top             =   330
         Width           =   1245
      End
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
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
      Left            =   1065
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   5835
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
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
      Left            =   1005
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   5355
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcHiddenPassword 
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
      Height          =   300
      Left            =   195
      TabIndex        =   58
      Top             =   5340
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
      Height          =   285
      Left            =   7620
      TabIndex        =   51
      Top             =   5400
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmcHub 
      Appearance      =   0  'Flat
      Caption         =   "&Hub"
      Height          =   285
      Left            =   8145
      TabIndex        =   50
      Top             =   7110
      Width           =   1050
   End
   Begin VB.Label lacCode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   45
      TabIndex        =   103
      Top             =   255
      Width           =   840
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   30
      Top             =   5655
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "UserOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Useropt.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: UserOpt.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the User Option input screen code
Option Explicit
Option Compare Text
Dim tmUrf() As URF         'URF record image
Dim tmWkUrf As URF          'Urf work image
Dim imUrfIndex As Integer   'Index into tmUrf
Dim tmSrchKey As INTKEY0    'URF key record image
Dim imUrfRecLen As Integer        'URF record length
Dim imIncludeDormant As Integer
Dim smPrgmmaticAllow As String
'Report Sets
Dim tmSnf As SNF            'SNF record image
Dim hmSnf As Integer        'SNF Handle
Dim imSnfRecLen As Integer      'SNF record length
'E-mail
Dim tmCef As CEF            'CEF record image
Dim hmCef As Integer        'CEF Handle
Dim imCefRecLen As Integer      'CEF record length
Dim tmCefSrchKey As LONGKEY0    'CEF key record image
Dim smEMail As String
Dim lmEMailCefCode As Long
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imSelChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imDVSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imRSSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imSPSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imVehSelectedIndex As Integer  'Index of selected record (0 if new)
Dim hmUrf As Integer        'User Options file handle
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imWin(0 To 69) As Integer   '0=Hide; 1=View; 2=Input
Dim imWinMap(0 To 69) As Integer
'10959 added to selected from 21 to 22
Dim imSelectedFields(0 To 22) As Integer '0=Hide; 1=View;   ' Dan M 4/10/09 added ChangeCsiDate and AllowFinalInvDisplay, only hide and input
                        '2=Input (Merge, HideSpots, ChgBilled only allow hide and input, ChgCntr only allow hide and input)
'Dim smSelFields(0 To 8) As String
Dim imTypeFields(0 To 6) As Integer '0=Hide; 1=View; 2=Input
Dim imAlerts(0 To 15) As Integer '0=Yes; 1=No
Dim imComboBoxIndex As Integer
Dim smVehicle As String    'Original name selected- used to check if name changed
Dim smRptSet As String    'Original name selected- used to check if name changed
Dim smHub As String
Dim tmHubCode() As SORTCODE
Dim smHubCodeTag As String
Dim smSalesPerson As String    'Original name selected- used to check if name changed
Dim smVeh As String    'Original name selected- used to check if name changed
Dim smDefVeh As String    'Original name selected- used to check if name changed
Dim imAltered As Integer    'Indicates if any field has been altered
Dim imNewRec As Integer     'True=New record,False=change record
Dim imSvUrfIndex() As Integer
Dim imModelCode() As Integer    'Code number for all records retained in model
Dim imStartX As Integer     'X coordinate start paint location
Dim imStartY As Integer     'Y coordinate start paint location
Dim imIAdj As Integer       'X adjustment for centering I
Dim imIgnoreChg As Integer  'Ignore changes to fields during Move Record to Controls
Dim imFirstTime As Integer  'If first time and new-show defaults
Dim imInSelect As Integer   'Avoid MovRecToCtrl being called twice
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imFirstFocus As Integer
Dim smCurrentPassword As String
Dim smLastModel As String
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imRemoteID() As Integer
Dim imUpdateAllowed As Integer    'User can update records

Private Sub cbcDefVeh_Change()
    If imChgMode = False Then
        imChgMode = True
        If cbcDefVeh.Text <> "" Then
            gManLookAhead cbcDefVeh, imBSMode, imComboBoxIndex
        End If
        imDVSelectedIndex = cbcDefVeh.ListIndex
        imChgMode = False
    End If
    mSetCommands
End Sub
Private Sub cbcDefVeh_Click()
    imComboBoxIndex = cbcDefVeh.ListIndex
    imDVSelectedIndex = cbcDefVeh.ListIndex
    mSetCommands 'Process change as change event is not generated
End Sub
Private Sub cbcDefVeh_DropDown()
    If (imDVSelectedIndex = -1) And (cbcDefVeh.ListCount > 0) Then
        cbcDefVeh.ListIndex = 0
    End If
End Sub
Private Sub cbcDefVeh_GotFocus()
    Dim ilLoop As Integer
    Dim slRecCode As String
    Dim slNameCode As String  'name and code
    Dim ilRet As Integer    'Return call status
    Dim slName As String
    Dim slCode As String    'code number

    imIgnoreChg = YES
    If StrComp(cbcVehicle.Text, "[All Vehicles]", 1) <> 0 Then
        If cbcDefVeh.ListCount <= 0 Then
            cbcDefVeh.AddItem cbcVehicle.Text
        Else
            slName = cbcVehicle.Text
            gFindMatch slName, 0, cbcDefVeh
            If gLastFound(cbcDefVeh) = -1 Then
                cbcDefVeh.AddItem cbcVehicle.Text
            End If
        End If
    End If
    If (imFirstTime = YES) And (imNewRec) Then
'        mInitCtrlFields
'        mMoveCtrlToRec
        If cbcDefVeh.ListIndex < 0 Then
            'slRecCode = Trim$(str$(tmUrf(LBound(tmUrf)).iDefVeh))
            If tmUrf(LBound(tmUrf)).iDefVeh <> 0 Then
                '1/6/10:  Removed reference to tgVehicle since it was never populated and all other references have been previously removed
                'For ilLoop = 0 To UBound(tgVehicle) - 1 Step 1  'Traffic!lbcVehicle.ListCount - 1 Step 1
                '    slNameCode = tgVehicle(ilLoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
                '    ilRet = gParseItem(slNameCode, 1, "\", slName)
                '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                '    If slRecCode = Trim$(slCode) Then
                '        gFindMatch slName, 0, cbcDefVeh
                '        If gLastFound(cbcDefVeh) >= 0 Then
                '            cbcDefVeh.ListIndex = gLastFound(cbcDefVeh)
                '        End If
                '        Exit For
                '    End If
                'Next ilLoop
                For ilLoop = 0 To cbcDefVeh.ListCount - 1 Step 1
                    If tmUrf(LBound(tmUrf)).iDefVeh = cbcDefVeh.ItemData(ilLoop) Then
                        cbcDefVeh.ListIndex = ilLoop
                        Exit For
                    End If
                Next ilLoop
            End If
        End If
        If cbcDefVeh.ListIndex < 0 Then
            If cbcDefVeh.ListCount > 0 Then
                cbcDefVeh.ListIndex = 0
            Else
                cbcDefVeh.ListIndex = -1
            End If
        End If
        cmcUpdate.Enabled = True
    End If
    imFirstTime = NO
    gCtrlGotFocus cbcDefVeh
    imIgnoreChg = NO
End Sub
Private Sub cbcDefVeh_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcDefVeh_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcDefVeh.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcDefVeh_LostFocus()
'    gSetIndexFromText cbcDefVeh
End Sub

Private Sub cbcHub_Change()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilRet                                                   *
'******************************************************************************************

    If imChgMode = False Then
        imChgMode = True
        If cbcSalesperson.Text <> "" Then
            gManLookAhead cbcHub, imBSMode, imComboBoxIndex
        End If
        imChgMode = False
    End If
    mSetCommands
End Sub

Private Sub cbcHub_Click()
    imComboBoxIndex = cbcSalesperson.ListIndex
    mSetCommands
End Sub

Private Sub cbcHub_GotFocus()
    gCtrlGotFocus cbcHub
End Sub

Private Sub cbcHub_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub cbcHub_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcHub.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub cbcModel_Change()
    Dim ilRet As Integer
    Dim ilRecLen As Integer    'URF record length
    Dim hlUrf As Integer    'User Option file handle
    Dim tlUrf As URF        'Local record image of user record
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilMapIndex As Integer
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        imChgMode = True    'Set change mode to avoid infinite loop
        ilRet = gOptionLookAhead(cbcModel, imBSMode, slStr)
        If ilRet = 0 Then
            smLastModel = cbcModel.Text
            hlUrf = CBtrvTable(ONEHANDLE)
            ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            On Error GoTo cbcModelErr
            gBtrvErrorMsg ilRet, "cbcModel_Change (btrOpen):" & "Urf.Btr", UserOpt
            On Error GoTo 0
            ilRecLen = Len(tlUrf)  'btrRecordLength(hlUrf)  'Get and save record length
            tmSrchKey.iCode = imModelCode(cbcModel.ListIndex)
            ilRet = btrGetEqual(hlUrf, tlUrf, ilRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            gUrfDecrypt tlUrf
            On Error GoTo cbcModelErr
            gBtrvErrorMsg ilRet, "cbcModel_Change (btrGetEqual):" & "Urf.Btr", UserOpt
            On Error GoTo 0
            btrExtClear hlUrf   'Clear any previous extend operation
            ilRet = btrClose(hlUrf)
            btrDestroy hlUrf
            'Move control setting to record
            For ilLoop = LBound(imWin) To UBound(imWin) Step 1
                If tlUrf.sWin(ilLoop) = "H" Then 'Hide
                    imWin(ilLoop) = 0
                ElseIf tlUrf.sWin(ilLoop) = "V" Then    'View only
                    imWin(ilLoop) = 1
                Else                    'Input
                    imWin(ilLoop) = 2
                End If
            Next ilLoop
            'Grid
            If tlUrf.sGrid = "H" Then 'Hide
                imSelectedFields(0) = 0
            ElseIf tlUrf.sGrid = "V" Then    'View only
                imSelectedFields(0) = 1
            Else                    'Input
                imSelectedFields(0) = 2
            End If
            'Price
            If tlUrf.sPrice = "H" Then 'Hide
                imSelectedFields(1) = 0
            ElseIf tlUrf.sPrice = "V" Then    'View only
                imSelectedFields(1) = 1
            Else                    'Input
                imSelectedFields(1) = 2
            End If
            'Credit
            If tlUrf.sCredit = "H" Then 'Hide
                imSelectedFields(2) = 0
            ElseIf tlUrf.sCredit = "V" Then    'View only
                imSelectedFields(2) = 1
            Else                    'Input
                imSelectedFields(2) = 2
            End If
            'Payment rating
            If tlUrf.sPayRate = "H" Then 'Hide
                imSelectedFields(3) = 0
            ElseIf tlUrf.sPayRate = "V" Then    'View only
                imSelectedFields(3) = 1
            Else                    'Input
                imSelectedFields(3) = 2
            End If
            'Merge
            If tlUrf.sMerge = "H" Then 'Hide
                imSelectedFields(4) = 0
            ElseIf tlUrf.sMerge = "V" Then    'View only
                imSelectedFields(4) = 1
            Else                    'Input
                imSelectedFields(4) = 2
            End If
            'Hide Spots
            If tlUrf.sHideSpots = "H" Then 'Hide
                imSelectedFields(5) = 0
            ElseIf tlUrf.sHideSpots = "V" Then    'View only
                imSelectedFields(5) = 1
            Else                    'Input
                imSelectedFields(5) = 2
            End If
            'Change Billed
            If tlUrf.sChgBilled = "H" Then 'Hide
                imSelectedFields(6) = 0
            ElseIf tlUrf.sChgBilled = "V" Then    'View only
                imSelectedFields(6) = 1
            Else                    'Input
                imSelectedFields(6) = 2
            End If
            'Change Contract
            If tlUrf.sChgCntr = "H" Then 'Hide
                imSelectedFields(7) = 0
            ElseIf tlUrf.sChgCntr = "V" Then    'View only
                imSelectedFields(7) = 1
            Else                    'Input
                imSelectedFields(7) = 2
            End If
            'Reference Reservation Contracts
            If tlUrf.sRefResvType = "H" Then 'Hide
                imSelectedFields(8) = 0
            ElseIf tlUrf.sRefResvType = "V" Then    'View only
                imSelectedFields(8) = 1
            Else                    'Input
                imSelectedFields(8) = 2
            End If
            'Credit Rating
            If tlUrf.sChgCrRt = "H" Then 'Hide
                imSelectedFields(9) = 0
            ElseIf tlUrf.sChgCrRt = "V" Then    'View only
                imSelectedFields(9) = 1
            Else                    'Input
                imSelectedFields(9) = 2
            End If
            ''Compute Button
            'If tlUrf.sUseComputeCMC = "H" Then 'Hide
            '    imSelectedFields(10) = 0
            'ElseIf tlUrf.sUseComputeCMC = "V" Then    'View only
            '    imSelectedFields(10) = 1
            'Else                    'Input
            '    imSelectedFields(10) = 2
            'End If
            'Region Copy
            If tlUrf.sRegionCopy = "H" Then 'Hide
                imSelectedFields(10) = 0
            ElseIf tlUrf.sRegionCopy = "V" Then    'View only
                imSelectedFields(10) = 1
            Else                    'Input
                imSelectedFields(10) = 2
            End If
            'Contract prices
            If tlUrf.sChgPrices = "H" Then 'Hide
                imSelectedFields(11) = 0
            ElseIf tlUrf.sChgPrices = "V" Then    'View only
                imSelectedFields(11) = 1
            Else                    'Input
                imSelectedFields(11) = 2
            End If
            'Flight
            If tlUrf.sActFlightButton = "I" Then 'Input
                imSelectedFields(12) = 2
            ElseIf tlUrf.sActFlightButton = "V" Then    'View only
                imSelectedFields(12) = 1
            Else                    'Hide
                imSelectedFields(12) = 0
            End If
            'change billed Contract prices
            If tlUrf.sChgLnBillPrice = "I" Then 'Input
                imSelectedFields(13) = 2
            ElseIf tlUrf.sChgLnBillPrice = "V" Then    'View only
                imSelectedFields(13) = 1
            Else                    'Hide
                imSelectedFields(13) = 0
            End If
            ' Dan M 4/10/09  14 and 15 are I or H only
            If tlUrf.sAllowInvDisplay = "I" Then 'Input
                imSelectedFields(14) = 2
           ' ElseIf tlUrf.sAllowInvDisplay = "V" Then    'View only
            '    imSelectedFields(14) = 1
            Else                    'Hide
                imSelectedFields(14) = 0
            End If
            If tlUrf.sChangeCSIDate = "I" Then 'Input
                imSelectedFields(15) = 2
           ' ElseIf tlUrf.sChangeCSIDate = "V" Then    'View only
           '     imSelectedFields(15) = 1
            Else                    'Hide
                imSelectedFields(15) = 0
            End If
            If tlUrf.sActivityLog = "V" Then    'View only
                imSelectedFields(16) = 1
            Else                    'Hide
                imSelectedFields(16) = 0
            End If
            If tlUrf.sCntrVerify = "I" Then
                imSelectedFields(17) = 2
            Else                    'Hide
                imSelectedFields(17) = 0
            End If
            If tlUrf.sChgAcq = "I" Then 'Input
                imSelectedFields(18) = 2
            Else                    'View
                imSelectedFields(18) = 1
            End If
            'If ((Asc(tgSaf(0).sFeatures6) And ADVANCEAVAILS) = ADVANCEAVAILS) Then
            If (tgSaf(0).sAdvanceAvail = "Y") Then
                If tlUrf.sAdvanceAvails = "I" Then
                    imSelectedFields(19) = 2
                Else                    'Hide
                    imSelectedFields(19) = 0
                End If
            Else
                imSelectedFields(19) = 0
            End If
            'Select Reservation
            If tlUrf.sResvType = "H" Then 'Hide
                imTypeFields(0) = 0
            ElseIf tlUrf.sResvType = "V" Then    'View only
                imTypeFields(0) = 1
            Else                    'Input
                imTypeFields(0) = 2
            End If
            'Select Remnant
            If tlUrf.sRemType = "H" Then 'Hide
                imTypeFields(1) = 0
            ElseIf tlUrf.sRemType = "V" Then    'View only
                imTypeFields(1) = 1
            Else                    'Input
                imTypeFields(1) = 2
            End If
            'Select DR
            If tlUrf.sDRType = "H" Then 'Hide
                imTypeFields(2) = 0
            ElseIf tlUrf.sDRType = "V" Then    'View only
                imTypeFields(2) = 1
            Else                    'Input
                imTypeFields(2) = 2
            End If
            'Select PI
            If tlUrf.sPIType = "H" Then 'Hide
                imTypeFields(3) = 0
            ElseIf tlUrf.sPIType = "V" Then    'View only
                imTypeFields(3) = 1
            Else                    'Input
                imTypeFields(3) = 2
            End If
            'Select PSA
            If tlUrf.sPSAType = "H" Then 'Hide
                imTypeFields(4) = 0
            ElseIf tlUrf.sPSAType = "V" Then    'View only
                imTypeFields(4) = 1
            Else                    'Input
                imTypeFields(4) = 2
            End If
            'Select Promo
            If tlUrf.sPromoType = "H" Then 'Hide
                imTypeFields(5) = 0
            ElseIf tlUrf.sPromoType = "V" Then    'View only
                imTypeFields(5) = 1
            Else                    'Input
                imTypeFields(5) = 2
            End If
            If tlUrf.sPrgmmaticAlert = "I" Then 'Hide
                imTypeFields(6) = 2
            ElseIf tlUrf.sPrgmmaticAlert = "V" Then    'View only
                imTypeFields(6) = 1
            Else                    'Input
                imTypeFields(6) = 0
            End If
            'tlUrf.ir
            'Show Alerts
            If tlUrf.sReprintLogAlert = "Y" Then 'Yes
                imAlerts(0) = 0
            Else                    'Input
                imAlerts(0) = 1
            End If
            'Show Incomplete
            If tlUrf.sIncompAlert = "Y" Then 'Yes
                imAlerts(1) = 0
            Else                    'Input
                imAlerts(1) = 1
            End If
            'Show Complete
            If tlUrf.sCompAlert = "Y" Then 'Yes
                imAlerts(2) = 0
            Else                    'Input
                imAlerts(2) = 1
            End If
            'Show Req Schd Contract
            If tlUrf.sSchAlert = "Y" Then 'Yes
                imAlerts(3) = 0
            Else                    'Input
                imAlerts(3) = 1
            End If
            'Show Hold
            If tlUrf.sHoldAlert = "Y" Then 'Yes
                imAlerts(4) = 0
            Else                    'Input
                imAlerts(4) = 1
            End If
            'Show Rate Card Chg
            If tlUrf.sRateCardAlert = "Y" Then 'Yes
                imAlerts(5) = 0
            Else                    'Input
                imAlerts(5) = 1
            End If
            'Show Research Chg
            If tlUrf.sResearchAlert = "Y" Then 'Yes
                imAlerts(6) = 0
            Else                    'Input
                imAlerts(6) = 1
            End If
            'Show Insuff Avail
            If tlUrf.sAvailAlert = "Y" Then 'Yes
                imAlerts(7) = 0
            Else                    'Input
                imAlerts(7) = 1
            End If
            'Show Credit Approved
            If tlUrf.sCrdChkAlert = "Y" Then 'Yes
                imAlerts(8) = 0
            Else                    'Input
                imAlerts(8) = 1
            End If
            'Show Credit Denied
            If tlUrf.sDeniedAlert = "Y" Then 'Yes
                imAlerts(9) = 0
            Else                    'Input
                imAlerts(9) = 1
            End If
            'Show Credit Exceeded
            If tlUrf.sCrdLimitAlert = "Y" Then 'Yes
                imAlerts(10) = 0
            Else                    'Input
                imAlerts(10) = 1
            End If
            'Show Affect LLD
            If tlUrf.sMoveAlert = "Y" Then 'Yes
                imAlerts(11) = 0
            Else                    'Input
                imAlerts(11) = 1
            End If
            'Allowed to Initiate Shutdown
            If tlUrf.sAllowedToBlock = "Y" Then 'Yes
                imAlerts(12) = 0
            Else                    'Input
                imAlerts(12) = 1
            End If
            'Show Rep-Net Messages
            If tlUrf.sShowNRMsg = "Y" Then 'Yes
                imAlerts(13) = 0
            Else                    'Input
                imAlerts(13) = 1
            End If
            
            '' Megaphone JJB
             'Email Digital Contracts
            If tlUrf.sDigitalCntrAlert = "Y" Then 'Yes
                imAlerts(14) = 0
            Else                    'Input
                imAlerts(14) = 1
            End If
            
             'Email Digital Impressions
            If tlUrf.sDigitalImpAlert = "Y" Then 'Yes
                imAlerts(15) = 0
            Else                    'Input
                imAlerts(15) = 1
            End If
            ''''''''''''''''
            
            If tlUrf.sWorkToDead = "Y" Then
                ckcWStatus(0).Value = vbChecked
            Else
                ckcWStatus(0).Value = vbUnchecked
            End If
            If tlUrf.sWorkToComp = "Y" Then
                ckcWStatus(1).Value = vbChecked
            Else
                ckcWStatus(1).Value = vbUnchecked
            End If
            If tlUrf.sWorkToHold = "Y" Then
                ckcWStatus(2).Value = vbChecked
            Else
                ckcWStatus(2).Value = vbUnchecked
            End If
            If tlUrf.sWorkToOrder = "Y" Then
                ckcWStatus(3).Value = vbChecked
            Else
                ckcWStatus(3).Value = vbUnchecked
            End If
            If tlUrf.sCompToIncomp = "Y" Then
                ckcCStatus(0).Value = vbChecked
            Else
                ckcCStatus(0).Value = vbUnchecked
            End If
            If tlUrf.sCompToDead = "Y" Then
                ckcCStatus(1).Value = vbChecked
            Else
                ckcCStatus(1).Value = vbUnchecked
            End If
            If tlUrf.sCompToHold = "Y" Then
                ckcCStatus(2).Value = vbChecked
            Else
                ckcCStatus(2).Value = vbUnchecked
            End If
            If tlUrf.sCompToOrder = "Y" Then
                ckcCStatus(3).Value = vbChecked
            Else
                ckcCStatus(3).Value = vbUnchecked
            End If
            If tlUrf.sIncompToDead = "Y" Then
                ckcIStatus(0).Value = vbChecked
            Else
                ckcIStatus(0).Value = vbUnchecked
            End If
            If tlUrf.sIncompToComp = "Y" Then
                ckcIStatus(1).Value = vbChecked
            Else
                ckcIStatus(1).Value = vbUnchecked
            End If
            If tlUrf.sIncompToHold = "Y" Then
                ckcIStatus(2).Value = vbChecked
            Else
                ckcIStatus(2).Value = vbUnchecked
            End If
            If tlUrf.sIncompToOrder = "Y" Then
                ckcIStatus(3).Value = vbChecked
            Else
                ckcIStatus(3).Value = vbUnchecked
            End If
            If tlUrf.sDeadToWork = "Y" Then
                ckcDStatus(0).Value = vbChecked
            Else
                ckcDStatus(0).Value = vbUnchecked
            End If
            If tlUrf.sHoldToOrder = "Y" Then
                ckcHStatus(0).Value = vbChecked
            Else
                ckcHStatus(0).Value = vbUnchecked
            End If
            If tlUrf.sReviseCntr = "N" Then
                rbcReviseCntr(1).Value = True
            Else
                rbcReviseCntr(0).Value = True
            End If
            For ilLoop = RATECARDSJOB To REPORTSJOB - 1 Step 1
                ilIndex = imWinMap(ilLoop)
                If ilIndex >= 0 Then
                    'pbcJobs(ilIndex).Cls
                    'pbcJobs(ilIndex).CurrentX = imStartX
                    'pbcJobs(ilIndex).CurrentY = imStartY
                    'If tlUrf.sWin(ilLoop) = "I" Then
                    '    pbcJobs(ilIndex).CurrentX = imStartX + imIAdj
                    'End If
                    'pbcJobs(ilIndex).Print tlUrf.sWin(ilLoop)
                    If plcJobs(ilIndex).Enabled Then
                        If tlUrf.sWin(ilLoop) = "H" Then
                            plcJobs(ilIndex).BackColor = Red
                        ElseIf tlUrf.sWin(ilLoop) = "V" Then
                            plcJobs(ilIndex).BackColor = Yellow
                        Else
                            plcJobs(ilIndex).BackColor = GREEN
                        End If
                    Else
                        plcJobs(ilIndex).BackColor = Red
                    End If
                End If
            Next ilLoop
            'If vbcSelFields.Value = 0 Then
            '    vbcSelFields_Change
            'Else
            '    vbcSelFields.Value = 0
            'End If
            For ilLoop = LBound(imSelectedFields) To UBound(imSelectedFields) Step 1
                'pbcSelectedFields_Paint ilLoop
                ilMapIndex = ilLoop
                If imSelectedFields(ilLoop) = 0 Then
                    plcSelectedFields(ilMapIndex).BackColor = Red    'Print "H"
                ElseIf imSelectedFields(ilLoop) = 1 Then
                    plcSelectedFields(ilMapIndex).BackColor = Yellow 'Print "V"
                ElseIf imSelectedFields(ilLoop) = 2 Then
                    plcSelectedFields(ilMapIndex).BackColor = GREEN   'Print "I"
                Else
                    plcSelectedFields(ilMapIndex).BackColor = GRAY   'Print ""
                End If
            Next ilLoop
            For ilLoop = LBound(imTypeFields) To UBound(imTypeFields) Step 1
                'pbcTypes_Paint ilLoop
                ilMapIndex = ilLoop
                If imTypeFields(ilLoop) = 0 Then
                    plcTypes(ilMapIndex).BackColor = Red 'Print "H"
                ElseIf imTypeFields(ilLoop) = 1 Then
                    plcTypes(ilMapIndex).BackColor = Yellow  'Print "V"
                ElseIf imTypeFields(ilLoop) = 2 Then
                    plcTypes(ilMapIndex).BackColor = GREEN   'Print "I"
                Else
                    plcTypes(ilMapIndex).BackColor = GRAY    'Print ""
                End If
            Next ilLoop
            For ilLoop = LBound(imAlerts) To UBound(imAlerts) Step 1
                'pbcAlerts_Paint ilLoop
                ilMapIndex = ilLoop
                If imAlerts(ilLoop) = 0 Then
                    plcAlerts(ilMapIndex).BackColor = GREEN
                ElseIf imAlerts(ilLoop) = 1 Then
                    plcAlerts(ilMapIndex).BackColor = Red
                Else
                    plcAlerts(ilMapIndex).BackColor = GRAY
                End If
            Next ilLoop
            '    Select Case ilLoop
            '        Case 0  'Grid
            '            If tlUrf.sGrid = "I" Then
            '                pbcSelectedFields(ilLoop).CurrentX = imStartX + imIAdj
            '            End If
            '            pbcSelectedFields(ilLoop).Print tlUrf.sGrid
            '        Case 1  'Price
            '            If tlUrf.sPrice = "I" Then
            '                pbcSelectedFields(ilLoop).CurrentX = imStartX + imIAdj
            '            End If
            '            pbcSelectedFields(ilLoop).Print tlUrf.sPrice
            '        Case 2  'Credit
            '            If tlUrf.sCredit = "I" Then
            '                pbcSelectedFields(ilLoop).CurrentX = imStartX + imIAdj
            '            End If
            '            pbcSelectedFields(ilLoop).Print tlUrf.sCredit
            '        Case 3  'PayRating
            '            If tlUrf.sPayRate = "I" Then
            '                pbcSelectedFields(ilLoop).CurrentX = imStartX + imIAdj
            '            End If
            '            pbcSelectedFields(ilLoop).Print tlUrf.sPayRate
            '        Case 4  'Merge
            '            If tlUrf.sMerge = "I" Then
            '                pbcSelectedFields(ilLoop).CurrentX = imStartX + imIAdj
            '            End If
            '            pbcSelectedFields(ilLoop).Print tlUrf.sMerge
            '    End Select
            'Next ilLoop
            For ilLoop = VEHICLESLIST To USERLIST Step 1
                ilIndex = imWinMap(ilLoop)
                If ilIndex >= 0 Then
                    'pbcLists_Paint ilIndex
                    If imWin(ilLoop) = 0 Then
                        plcLists(ilIndex).BackColor = Red
                    ElseIf imWin(ilLoop) = 1 Then
                        plcLists(ilIndex).BackColor = Yellow
                    Else
                        plcLists(ilIndex).BackColor = GREEN
                    End If
                End If
            Next ilLoop
            'Force change so altered flag will be set in mSetChg
            For ilLoop = LBound(imWin) To UBound(imWin) Step 1
                tmWkUrf.sWin(ilLoop) = ""
            Next ilLoop
        End If
        imChgMode = False
    End If
    mSetCommands
    Exit Sub
cbcModelErr:
    On Error GoTo 0
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    imIgnoreChg = NO
    imTerminate = True
    imChgMode = False
    Exit Sub
End Sub
Private Sub cbcModel_Click()
    cbcModel_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcModel_DropDown()
    If (cbcModel.Text = "") And (smLastModel <> "") Then
        gFindMatch smLastModel, 0, cbcModel
        If gLastFound(cbcModel) >= 0 Then
            cbcModel.ListIndex = gLastFound(cbcModel)
        End If
    End If
End Sub
Private Sub cbcModel_GotFocus()
    imIgnoreChg = YES
    If (imFirstTime = YES) And (imNewRec) Then
'        mInitCtrlFields
'        mMoveCtrlToRec
        cmcUpdate.Enabled = True
    End If
    If (cbcModel.Text = "") And (smLastModel <> "") Then
        gFindMatch smLastModel, 0, cbcModel
        If gLastFound(cbcModel) >= 0 Then
            cbcModel.ListIndex = gLastFound(cbcModel)
        End If
    End If
    imFirstTime = NO
    gCtrlGotFocus cbcModel
    imIgnoreChg = NO
End Sub
Private Sub cbcModel_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcModel_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcModel.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcRptSet_Change()
    Dim slStr As String
    Dim ilRet As Integer
    If imChgMode = False Then
        imChgMode = True
        ilRet = gOptionLookAhead(cbcRptSet, imBSMode, slStr)
        If ilRet = 1 Then
            cbcRptSet.ListIndex = 0
        End If
        imRSSelectedIndex = cbcRptSet.ListIndex
        imChgMode = False
    End If
    mSetCommands
End Sub
Private Sub cbcRptSet_Click()
    cbcRptSet_Change
End Sub
Private Sub cbcRptSet_GotFocus()
    gCtrlGotFocus cbcRptSet
End Sub
Private Sub cbcRptSet_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcRptSet_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcRptSet.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcSalesperson_Change()
    If imChgMode = False Then
        imChgMode = True
        If cbcSalesperson.Text <> "" Then
            gManLookAhead cbcSalesperson, imBSMode, imComboBoxIndex
        End If
        imSPSelectedIndex = cbcSalesperson.ListIndex
        imChgMode = False
    End If
    mSetCommands
End Sub
Private Sub cbcSalesperson_Click()
    imComboBoxIndex = cbcSalesperson.ListIndex
    imSPSelectedIndex = cbcSalesperson.ListIndex
    mSetCommands 'Process change as change event is not generated
End Sub
Private Sub cbcSalesperson_DropDown()
    mSlfPop
    If imTerminate Then
        Exit Sub
    End If
    If (imSPSelectedIndex = -1) And (cbcSalesperson.ListCount >= 1) Then
        cbcSalesperson.ListIndex = 0
    End If
End Sub
Private Sub cbcSalesperson_GotFocus()
    imIgnoreChg = YES
    If (imFirstTime = YES) And (imNewRec) Then
'        mInitCtrlFields
'        mMoveCtrlToRec
        cmcUpdate.Enabled = True
    End If
    imFirstTime = NO
    mSlfPop
    If imTerminate Then
        Exit Sub
    End If
    gCtrlGotFocus cbcSalesperson
    imIgnoreChg = NO
End Sub
Private Sub cbcSalesperson_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcSalesperson_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSalesperson.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcSelect_Change()
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim slName As String  'Current name selected from combo box
    Dim ilType As Integer
    Dim ilIndex As Integer
    imIgnoreChg = YES
    If imSelChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imSelChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    imInSelect = True
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex 'required by mVefPop which is called by rbcOption
        imUrfIndex = -1
        slName = cbcSelect.List(cbcSelect.ListIndex)
        sgUrfStamp = ""
        ilRet = csiSetStamp("URF", sgUrfStamp)
        gUrfRead UserOpt, slName, False, tmUrf(), imIncludeDormant
        If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
            lacCode.Caption = Str$(tmUrf(0).iCode)
        Else
            lacCode.Caption = ""
        End If
        plcName.Enabled = False
'        rbcOption(0).Enabled = True
'        'Set add to true so count can be checked
'        rbcOption(0).Value = True
'        If cbcVehicle.ListCount <= 0 Then
'            rbcOption(0).Enabled = False
'        End If
'        rbcOption(1).Enabled = True
'        If rbcOption(1).Value Then
'            rbcOption_Click 1
'        Else
'            rbcOption(1).Value = True   'Force to change mode
'        End If
        cbcVehicle.Enabled = True
        cbcDefVeh.Enabled = True
        cbcSalesperson.Enabled = True
        If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
            cbcHub.Visible = True
            lacHub.Visible = True
        Else
            cbcHub.Visible = False
            lacHub.Visible = False
        End If
        frcLiveLog.Enabled = True
        frcSports.Enabled = True
'        If rbcOption(0).Value Then
'        If UBound(tmUrf) <= 0 Then
'            ilType = 1
'            imNewRec = True
'        Else
            ilType = 0
            imNewRec = False
'        End If
'            mVehPop cbcVehicle, ilType    'When option set- this is done
'            If cbcVehicle.ListCount > 0 Then
'                cbcVehicle.ListIndex = 0
'            Else
'                cbcVehicle.Text = ""
'            End If
'            mVehPop cbcDefVeh, 0
        If (cbcSelect.Text = sgCPName) Or (cbcSelect.Text = sgSUName) Then
            imUrfIndex = 0
            gInitSuperUser tmUrf(0)
            plcName.Enabled = True
'            rbcOption(0).Enabled = False
'            rbcOption(1).Enabled = False
            cbcVehicle.Enabled = False
            cbcDefVeh.Enabled = False
            cbcSalesperson.Enabled = False
            cbcHub.Visible = False
            lacHub.Visible = False
            frcLiveLog.Enabled = False
            frcSports.Enabled = False
            imNewRec = False
        Else
            imUrfIndex = 0
        End If
'        If tgUrf(0).iCode > 2 Then
'            rbcOption(0).Enabled = False
'            rbcOption(1).Enabled = False
'        End If
    Else
        lacCode.Caption = ""
        imSelectedIndex = 0 'required by mVefPop which is called by rbcOption
'        rbcOption(0).Enabled = True
'        rbcOption(1).Enabled = False
'        rbcOption(0).Value = True   'Force to add mode
        cbcVehicle.Enabled = True
        cbcDefVeh.Enabled = True
'        cbcDefVeh.Clear
        cbcSalesperson.Enabled = True
'            cbcSalesperson.Clear
        If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
            cbcHub.Visible = True
            lacHub.Visible = True
        Else
            cbcHub.Visible = False
            lacHub.Visible = False
        End If
        If ilRet = 1 Then
            cbcSelect.ListIndex = 0
        End If
        ilRet = 1   'Clear fields as no match name found
        imNewRec = True
'            mVehPop cbcVehicle, ilType 'To fill list box before gotfocus
'            If cbcVehicle.ListCount > 0 Then
'                cbcVehicle.ListIndex = 0
'            Else
'                cbcVehicle.Text = ""
'            End If
    End If
    rbcSet(3).Value = True
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
'        mVehPop cbcDefVeh, 0, True
        mMoveRecToCtrl
    Else
        imSelectedIndex = 0
        mClearCtrlFields
'        mVehPop cbcDefVeh, 0, True
        If slStr <> "[New]" Then
            edcName.Text = slStr
        End If
    End If
    If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
        ilIndex = imWinMap(SITELIST)
        plcLists(ilIndex).Enabled = True
        ilIndex = imWinMap(USERLIST)
        plcLists(ilIndex).Enabled = True
    Else
        ilIndex = imWinMap(SITELIST)
        plcLists(ilIndex).Enabled = False
        ilIndex = imWinMap(USERLIST)
        plcLists(ilIndex).Enabled = False
    End If
    If (Trim$(tmWkUrf.sName) = sgCPName) Or (Trim$(tmWkUrf.sName) = sgSUName) Then
        frcJobs.Enabled = False
        frcLists.Enabled = False
        frcGeneral.Enabled = False
        cbcModel.Enabled = False
        cmcPassword.Enabled = False
    ElseIf (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
        frcJobs.Enabled = True
        frcLists.Enabled = True
        frcGeneral.Enabled = True
        cbcModel.Enabled = True
        cmcPassword.Enabled = False
    Else    'Any other user
        frcJobs.Enabled = False
        frcLists.Enabled = False
        frcGeneral.Enabled = False
        cbcModel.Enabled = False
        cmcPassword.Enabled = True
    End If
    imInSelect = False
    imSelChgMode = False
    imIgnoreChg = YES
'    mSetCommands
    imBypassSetting = False
    If ilRet = 0 Then
        mSetCommands
    End If
    'Dan M limit regular 'guide'
     mLimitGuide
    imIgnoreChg = NO
    Exit Sub

    On Error GoTo 0
    imIgnoreChg = NO
    imTerminate = True
    imSelChgMode = False
    Exit Sub
End Sub
Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcSelect_DropDown()
    mPopulate  'Populate user selection combo box
    If imTerminate Then
        Exit Sub
    End If
End Sub
Private Sub cbcSelect_GotFocus()
    Dim slSvText As String   'Save so list box can be reset
    If imTerminate Then
        Exit Sub
    End If
    imFirstTime = YES
    slSvText = cbcSelect.Text
'    ilSvIndex = cbcSelect.ListIndex
    mPopulate  'Populate user selection combo box
    If imTerminate Then
        Exit Sub
    End If
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
    End If
    If cbcSelect.ListCount <= 1 Then 'Test if any user exist
        If StrComp(cbcSelect.List(0), "[New]", 1) = 0 Then
            cbcSelect.ListIndex = 0
            mClearCtrlFields
            ''rbcOption(0).Value = True
            'cbcVehicle.SetFocus
            Exit Sub
        Else
            cbcSelect.ListIndex = 0
            mSetCommands
            'If plcName.Enabled Then
            '    cbcVehicle.SetFocus
            'Else
                cmcDone.SetFocus
            'End If
            Exit Sub
        End If
    End If
    gCtrlGotFocus cbcSelect
    If (slSvText = "") Or (slSvText = "[New]") Then
        cbcSelect.ListIndex = 0
        cbcSelect_Change    'Call change so picture area repainted
    Else
        gFindMatch slSvText, 1, cbcSelect
        If gLastFound(cbcSelect) > 0 Then
'            If (ilSvIndex <> gLastFound(cbcSelect)) Or (ilSvIndex <> cbcSelect.ListIndex) Then
            If (slSvText <> cbcSelect.List(gLastFound(cbcSelect))) Or imPopReqd Then
                cbcSelect.ListIndex = gLastFound(cbcSelect)
                cbcSelect_Change    'Call change so picture area repainted
                imPopReqd = False
            End If
        Else
            cbcSelect.ListIndex = 0
            mClearCtrlFields
            cbcSelect_Change    'Call change so picture area repainted
        End If
    End If
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
Private Sub cbcSelect_LostFocus()
'    gSetIndexFromText cbcSelect
End Sub
Private Sub cbcVehicle_Change()
    If imChgMode = False Then
        imChgMode = True
        If cbcVehicle.Text <> "" Then
            gManLookAhead cbcVehicle, imBSMode, imComboBoxIndex
        End If
        imVehSelectedIndex = cbcVehicle.ListIndex
        If imVehSelectedIndex > 0 Then
            edcGroupNo.Text = ""
            cbcDefVeh.ListIndex = imVehSelectedIndex
            cbcDefVeh.Enabled = False
        Else
            cbcDefVeh.Enabled = True
        End If
        imChgMode = False
    End If
    mSetCommands
End Sub
Private Sub cbcVehicle_Click()
    imComboBoxIndex = cbcVehicle.ListIndex
    imVehSelectedIndex = cbcVehicle.ListIndex
    cbcVehicle_Change
End Sub
Private Sub cbcVehicle_DropDown()
    If (imVehSelectedIndex = -1) And (cbcVehicle.ListCount > 0) Then
        cbcVehicle.ListIndex = 0
    End If
End Sub
Private Sub cbcVehicle_GotFocus()
''    gSetIndexFromText cbcSelect
'    If rbcOption(0).Value Then
'        ilType = 1
'    Else
'        ilType = 0
'    End If
'    mVehPop cbcVehicle, ilType, False 'This is required if Vehicle changed since populating it
'    If imTerminate Then
'        Exit Sub
'    End If
    imFirstTime = YES
    If (cbcVehicle.Text = "") And (cbcVehicle.ListCount > 0) Then
        cbcVehicle.ListIndex = 0
    End If
    gCtrlGotFocus cbcVehicle
    imComboBoxIndex = cbcVehicle.ListIndex
End Sub
Private Sub cbcVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcVehicle_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcVehicle.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub ckcBlockRU_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcBlockRU.Value = vbChecked Then
        Value = True
    End If
    'End of Coded added
    mSetCommands
End Sub
Private Sub ckcCStatus_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcCStatus(Index).Value = vbChecked Then
        Value = True
    End If
    'End of Coded added
    mSetCommands
End Sub
Private Sub ckcDStatus_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcDStatus(Index).Value = vbChecked Then
        Value = True
    End If
    'End of Coded added
    mSetCommands
End Sub
Private Sub ckcHStatus_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcHStatus(Index).Value = vbChecked Then
        Value = True
    End If
    'End of Coded added
    mSetCommands
End Sub
Private Sub ckcIStatus_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcIStatus(Index).Value = vbChecked Then
        Value = True
    End If
    'End of Coded added
    mSetCommands
End Sub
Private Sub ckcWStatus_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcWStatus(Index).Value = vbChecked Then
        Value = True
    End If
    'End of Coded added
    mSetCommands
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
    Dim slMess As String
    Dim ilRes As Integer
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    mSetChg
    If imAltered Then
        If Not imNewRec Then
            slMess = "Save Changes to " & cbcSelect.Text
        Else
            slMess = "Add " & Trim$(edcName.Text)
        End If
        ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
        If ilRes = vbCancel Then
            If plcName.Enabled Then
                If edcPassword.Enabled Then
                    edcPassword.SetFocus
                Else
                    edcRept.SetFocus
                End If
            Else
                cbcSelect.SetFocus
            End If
            Exit Sub
        End If
        If ilRes = vbYes Then
            If Len(Trim$(edcName.Text)) = 0 Then
                Beep
                edcName.SetFocus
                Exit Sub
            End If
'            If (edcPassword.Enabled) And Len(Trim$(edcPassword.Text)) < 4 Then
'                Beep
'                edcPassword.SetFocus
'                Exit Sub
'            End If
            cmcUpdate_Click
        End If
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcErase_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slName                                                                                *
'******************************************************************************************

    Dim ilRet As Integer
    Dim tlUrf As URF    'Position to record so it can be updated
    Dim slSyncDate As String
    Dim slSyncTime As String

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If sgCPName = cbcSelect.Text Then
        Exit Sub
    End If
    If sgSUName = cbcSelect.Text Then
        Exit Sub
    End If
    imIgnoreChg = YES
    If (Not imNewRec) And (imUrfIndex > -1) Then
        gGetSyncDateTime slSyncDate, slSyncTime
        tmSrchKey.iCode = tmUrf(imUrfIndex).iCode
        ilRet = btrGetEqual(hmUrf, tlUrf, imUrfRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        gUrfDecrypt tlUrf
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrGetEqual)", UserOpt
        On Error GoTo 0
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        ilRet = MsgBox("OK to remove " & Trim$(tlUrf.sName) & " from using " & cbcVehicle.Text, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        tmUrf(imUrfIndex).sDelete = "Y"
        gPackDate slSyncDate, tmUrf(imUrfIndex).iSyncDate(0), tmUrf(imUrfIndex).iSyncDate(1)
        gPackTime slSyncTime, tmUrf(imUrfIndex).iSyncTime(0), tmUrf(imUrfIndex).iSyncTime(1)
        gUrfEncrypt tmUrf(imUrfIndex)
        ilRet = btrUpdate(hmUrf, tmUrf(imUrfIndex), imUrfRecLen)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrUpdate)", UserOpt
        On Error GoTo 0
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        gUrfDecrypt tmUrf(imUrfIndex)
    End If
    mSetCommands
    mModelPop
    'If cbcVehicle.ListCount > 1 Then
    '    slName = Trim$(tlUrf.sName)
    '    sgUrfStamp = ""
    '    ilRet = csiSetStamp("URF", sgUrfStamp)
    '    gUrfRead UserOpt, slName, False, tmUrf(), imIncludeDormant
    '    cbcVehicle.SetFocus
    'Else
        If cbcSelect.Enabled Then
            cbcSelect.SetFocus
        End If
    'End If
    imIgnoreChg = NO
    Exit Sub
cmcEraseErr:
    On Error GoTo 0
    imTerminate = True
    imIgnoreChg = NO
    Resume Next
End Sub
Private Sub cmcErase_GotFocus()
    gCtrlGotFocus cmcErase
End Sub

Private Sub cmcErasePassword_Click()
    edcPassword.Text = ""
    'guide doesn't get message
    If Not sgUserName = sgSUName Then
        MsgBox "Your password will be erased upon saving.  The next time you enter traffic, you will be asked for a new password.", vbOKOnly, "User Password"
    End If
End Sub

Private Sub cmcHub_Click()
    mHubBranch
End Sub

Private Sub cmcPassword_Click()
    Dim ilRet As Integer
    Dim tlUrf As URF    'Position to record so it can be updated
    Dim ilLoop As Integer
    Dim slDate As String
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilUrfIndex As Integer

    ilUrfIndex = imUrfIndex
    If (ilUrfIndex < 0) Or (ilUrfIndex > UBound(tmUrf)) Then
        ilUrfIndex = 0
    End If
    tgPWUrf = tmUrf(ilUrfIndex)
    CSINewPW.Show vbModal
    edcHiddenPassword.Text = sgPWResult
    If edcHiddenPassword.Text <> "" Then
        gGetSyncDateTime slSyncDate, slSyncTime
        'Update password
        tmUrf(ilUrfIndex).sOldPassword(2) = tmUrf(ilUrfIndex).sOldPassword(1)
        tmUrf(ilUrfIndex).sOldPassword(1) = tmUrf(ilUrfIndex).sOldPassword(0)
        tmUrf(ilUrfIndex).sOldPassword(0) = tmUrf(ilUrfIndex).sPassword
        tmUrf(ilUrfIndex).sPassword = edcHiddenPassword.Text
        tmSrchKey.iCode = tmUrf(ilUrfIndex).iCode
        ilRet = btrGetEqual(hmUrf, tlUrf, imUrfRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        gUrfDecrypt tlUrf
        On Error GoTo cmcPasswordErr
        gBtrvErrorMsg ilRet, "cmcPassword (btrGetEqual)", UserOpt
        On Error GoTo 0
        slDate = Format$(gNow(), "m/d/yy")
        gPackDate slDate, tmUrf(ilUrfIndex).iPasswordDate(0), tmUrf(ilUrfIndex).iPasswordDate(1)
        gPackDate slSyncDate, tmUrf(ilUrfIndex).iSyncDate(0), tmUrf(ilUrfIndex).iSyncDate(1)
        gPackTime slSyncTime, tmUrf(ilUrfIndex).iSyncTime(0), tmUrf(ilUrfIndex).iSyncTime(1)
        gUrfEncrypt tmUrf(ilUrfIndex)
        ilRet = btrUpdate(hmUrf, tmUrf(ilUrfIndex), imUrfRecLen)
        On Error GoTo cmcPasswordErr
        gBtrvErrorMsg ilRet, "cmcPassword (btrUpdate)", UserOpt
        On Error GoTo 0
        gUrfDecrypt tmUrf(ilUrfIndex)
        tmWkUrf = tmUrf(ilUrfIndex)
        
'D.S. 09-09-15 get new Password here and call API  read in password from CEF maybe make a function to handle it in EDS pass it the tmURF ?
        
        For ilLoop = LBound(tmUrf) To UBound(tmUrf) Step 1
            If ilLoop <> ilUrfIndex Then
                tmSrchKey.iCode = tmUrf(ilLoop).iCode
                ilRet = btrGetEqual(hmUrf, tmUrf(ilLoop), imUrfRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                gUrfDecrypt tmUrf(ilLoop)
                On Error GoTo cmcPasswordErr
                gBtrvErrorMsg ilRet, "cmcPassword (btrGetEqual)", UserOpt
                On Error GoTo 0
                'tmUrf(ilLoop).sName = tmUrf(ilUrfIndex).sName
                tmUrf(ilLoop).sOldPassword(2) = tmUrf(ilUrfIndex).sOldPassword(2)
                tmUrf(ilLoop).sOldPassword(1) = tmUrf(ilUrfIndex).sOldPassword(1)
                tmUrf(ilLoop).sOldPassword(0) = tmUrf(ilUrfIndex).sOldPassword(0)
                tmUrf(ilLoop).sPassword = tmUrf(ilUrfIndex).sPassword
                gPackDate slDate, tmUrf(ilLoop).iPasswordDate(0), tmUrf(ilLoop).iPasswordDate(1)
                'tmUrf(ilLoop).sRept = tmUrf(ilUrfIndex).sRept
                'tmUrf(ilLoop).iDefVeh = tmUrf(ilUrfIndex).iDefVeh
                'tmUrf(ilLoop).iSlfCode = tmUrf(ilUrfIndex).iSlfCode
                ''Update all list task which are vehicle independent
                ''General
                'tmUrf(ilLoop).sWin(ADVERTISERSLIST) = tmUrf(ilUrfIndex).sWin(ADVERTISERSLIST)
                'tmUrf(ilLoop).sWin(AGENCIESLIST) = tmUrf(ilUrfIndex).sWin(AGENCIESLIST)
                'tmUrf(ilLoop).sWin(COMPETITIVESLIST) = tmUrf(ilUrfIndex).sWin(COMPETITIVESLIST)
                'tmUrf(ilLoop).sWin(EXCLUSIONSLIST) = tmUrf(ilUrfIndex).sWin(EXCLUSIONSLIST)
                ''Sales
                'tmUrf(ilLoop).sWin(SALESSOURCESLIST) = tmUrf(ilUrfIndex).sWin(SALESSOURCESLIST)
                'tmUrf(ilLoop).sWin(SALESREGIONSLIST) = tmUrf(ilUrfIndex).sWin(SALESREGIONSLIST)
                'tmUrf(ilLoop).sWin(SALESTEAMSLIST) = tmUrf(ilUrfIndex).sWin(SALESTEAMSLIST)
                'tmUrf(ilLoop).sWin(SALESPEOPLELIST) = tmUrf(ilUrfIndex).sWin(SALESPEOPLELIST)
                'tmUrf(ilLoop).sWin(REVENUESETSLIST) = tmUrf(ilUrfIndex).sWin(REVENUESETSLIST)
                'tmUrf(ilLoop).sWin(BOILERPLATESLIST) = tmUrf(ilUrfIndex).sWin(BOILERPLATESLIST)
                ''Programming
                'tmUrf(ilLoop).sWin(EVENTTYPESLIST) = tmUrf(ilUrfIndex).sWin(EVENTTYPESLIST)
                'tmUrf(ilLoop).sWin(AVAILNAMESLIST) = tmUrf(ilUrfIndex).sWin(AVAILNAMESLIST)
                'tmUrf(ilLoop).sWin(FEEDTYPESLIST) = tmUrf(ilUrfIndex).sWin(FEEDTYPESLIST)
                'tmUrf(ilLoop).sWin(GENRESLIST) = tmUrf(ilUrfIndex).sWin(GENRESLIST)
                ''Copy
                'tmUrf(ilLoop).sWin(MEDIADEFINITIONSLIST) = tmUrf(ilUrfIndex).sWin(MEDIADEFINITIONSLIST)
                'tmUrf(ilLoop).sWin(ANNOUNCERNAMESLIST) = tmUrf(ilUrfIndex).sWin(ANNOUNCERNAMESLIST)
                ''Accounting
                'tmUrf(ilLoop).sWin(MISSEDREASONSLIST) = tmUrf(ilUrfIndex).sWin(MISSEDREASONSLIST)
                'tmUrf(ilLoop).sWin(ITEMBILLINGTYPESLIST) = tmUrf(ilUrfIndex).sWin(ITEMBILLINGTYPESLIST)
                'tmUrf(ilLoop).sWin(INVOICESORTLIST) = tmUrf(ilUrfIndex).sWin(INVOICESORTLIST)
                'tmUrf(ilLoop).sWin(LOCKBOXESLIST) = tmUrf(ilUrfIndex).sWin(LOCKBOXESLIST)
                'tmUrf(ilLoop).sWin(EDISERVICESLIST) = tmUrf(ilUrfIndex).sWin(EDISERVICESLIST)
                ''Option
                'tmUrf(ilLoop).sWin(SITELIST) = tmUrf(ilUrfIndex).sWin(SITELIST)
                'tmUrf(ilLoop).sWin(USERLIST) = tmUrf(ilUrfIndex).sWin(USERLIST)
                'tmUrf(ilLoop).sCredit = tmUrf(ilUrfIndex).sCredit
                'tmUrf(ilLoop).sPayRate = tmUrf(ilUrfIndex).sPayRate
                'tmUrf(ilLoop).sMerge = tmUrf(ilUrfIndex).sMerge
                gPackDate slSyncDate, tmUrf(ilLoop).iSyncDate(0), tmUrf(ilLoop).iSyncDate(1)
                gPackTime slSyncTime, tmUrf(ilLoop).iSyncTime(0), tmUrf(ilLoop).iSyncTime(1)
                gUrfEncrypt tmUrf(ilLoop)
                ilRet = btrUpdate(hmUrf, tmUrf(ilLoop), imUrfRecLen)
                On Error GoTo cmcPasswordErr
                gBtrvErrorMsg ilRet, "cmcPassword (btrUpdate)", UserOpt
                On Error GoTo 0
                gUrfDecrypt tmUrf(ilLoop)
            End If
        Next ilLoop
    End If
    Exit Sub
cmcPasswordErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub cmcRptSet_Click()
    sgRptSetName = Trim$(cbcRptSet.Text)
    RptSets.Show vbModal
    mSnfPop
    If igRptSetReturn = 0 Then  'Done press
        gFindMatch sgRptSetName, 1, cbcRptSet
        sgRptSetName = ""
        If gLastFound(cbcRptSet) > 0 Then
            imChgMode = True
            cbcRptSet.ListIndex = gLastFound(cbcRptSet)
            imRSSelectedIndex = cbcRptSet.ListIndex
            imChgMode = False
        Else
            imChgMode = True
            cbcRptSet.ListIndex = 0
            imChgMode = False
            cbcRptSet.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub cmcRptSet_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcSetup_Click()
    'cdcSetup.flags = cdlPDPrintSetup
    'cdcSetup.Action = 5    'DLG_PRINT
    cdcSetup.flags = cdlPDPrintSetup
    cdcSetup.ShowPrinter
End Sub
Private Sub cmcTestDefault_Click()
    Dim slPDFDrvChar As String * 1       '1st character of PDF Print Driver Name
    Dim ilPDFDnArrowCnt As Integer       'Number of down arrows to select PDF Driver from printer setup box
    Dim slPrtDrvChar As String * 1       '1st character of Default Print Driver Name
    Dim ilPrtDnArrowCnt As Integer       'Number of down arrows to select Default Driver from printer setup box
    Dim slPrtNameAltKey As String * 1    'Alt Key for defining Printer Name
    Dim ilPrtNoEnterKeys As Integer     'Number of times the enter key must be entered to leave the Printer Setup screen

    slPrtNameAltKey = tgUrf(0).sPrtNameAltKey
    slPDFDrvChar = tgUrf(0).sPDFDrvChar
    ilPDFDnArrowCnt = tgUrf(0).iPDFDnArrowCnt
    slPrtDrvChar = tgUrf(0).sPrtDrvChar
    ilPrtDnArrowCnt = tgUrf(0).iPrtDnArrowCnt
    ilPrtNoEnterKeys = tgUrf(0).iPrtNoEnterKeys

    tgUrf(0).sPrtNameAltKey = Trim$(edcPDF(4).Text)
    tgUrf(0).sPDFDrvChar = Trim$(edcPDF(0).Text)
    tgUrf(0).iPDFDnArrowCnt = Val(edcPDF(1).Text)
    tgUrf(0).sPrtDrvChar = Trim$(edcPDF(2).Text)
    tgUrf(0).iPrtDnArrowCnt = Val(edcPDF(3).Text)
    tgUrf(0).iPrtNoEnterKeys = Val(edcPDF(5).Text)
    gSwitchToPDF cdcSetup, 1
    tgUrf(0).sPrtNameAltKey = slPrtNameAltKey
    tgUrf(0).sPDFDrvChar = slPDFDrvChar
    tgUrf(0).iPDFDnArrowCnt = ilPDFDnArrowCnt
    tgUrf(0).sPrtDrvChar = slPrtDrvChar
    tgUrf(0).iPrtDnArrowCnt = ilPrtDnArrowCnt
    tgUrf(0).iPrtNoEnterKeys = ilPrtNoEnterKeys
End Sub
Private Sub cmcTestPDF_Click()
    Dim slPDFDrvChar As String * 1       '1st character of PDF Print Driver Name
    Dim ilPDFDnArrowCnt As Integer       'Number of down arrows to select PDF Driver from printer setup box
    Dim slPrtDrvChar As String * 1       '1st character of Default Print Driver Name
    Dim ilPrtDnArrowCnt As Integer       'Number of down arrows to select Default Driver from printer setup box
    Dim slPrtNameAltKey As String * 1    'Alt Key for defining Printer Name
    Dim ilPrtNoEnterKeys As Integer     'Number of times the enter key must be entered to leave the Printer Setup screen

    slPrtNameAltKey = tgUrf(0).sPrtNameAltKey
    slPDFDrvChar = tgUrf(0).sPDFDrvChar
    ilPDFDnArrowCnt = tgUrf(0).iPDFDnArrowCnt
    slPrtDrvChar = tgUrf(0).sPrtDrvChar
    ilPrtDnArrowCnt = tgUrf(0).iPrtDnArrowCnt
    ilPrtNoEnterKeys = tgUrf(0).iPrtNoEnterKeys

    tgUrf(0).sPrtNameAltKey = Trim$(edcPDF(4).Text)
    tgUrf(0).sPDFDrvChar = Trim$(edcPDF(0).Text)
    tgUrf(0).iPDFDnArrowCnt = Val(edcPDF(1).Text)
    tgUrf(0).sPrtDrvChar = Trim$(edcPDF(2).Text)
    tgUrf(0).iPrtDnArrowCnt = Val(edcPDF(3).Text)
    tgUrf(0).iPrtNoEnterKeys = Val(edcPDF(5).Text)
    gSwitchToPDF cdcSetup, 0
    tgUrf(0).sPrtNameAltKey = slPrtNameAltKey
    tgUrf(0).sPDFDrvChar = slPDFDrvChar
    tgUrf(0).iPDFDnArrowCnt = ilPDFDnArrowCnt
    tgUrf(0).sPrtDrvChar = slPrtDrvChar
    tgUrf(0).iPrtDnArrowCnt = ilPrtDnArrowCnt
    tgUrf(0).iPrtNoEnterKeys = ilPrtNoEnterKeys
End Sub
Private Sub cmcUndo_Click()
    If imNewRec Then
'        If UBound(tmUrf) > 0 Then
'            ilBound = UBound(tmUrf) - 1
'            ReDim Preserve tmUrf(0 To ilBound)
'        End If
        imIgnoreChg = YES
        mSetCommands
        imIgnoreChg = NO
        cbcSelect.ListIndex = 0
        cbcSelect.SetFocus
    Else
        mMoveRecToCtrl
        edcName.SetFocus
        mSetCommands
    End If
End Sub
Private Sub cmcUndo_GotFocus()
    gCtrlGotFocus cmcUndo
End Sub
Private Sub cmcUpdate_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVefCode                                                                             *
'******************************************************************************************

    Dim ilRet As Integer
    Dim slMsg As String
    Dim tlUrf As URF    'Position to record so it can be updated
    Dim ilLoop As Integer
    Dim slName As String
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilUrf As Integer
    Dim ilCode As Integer
    Dim slStr As String
    Dim ilMatchOld As Integer
    Dim ilUrfCode As Integer

    ilUrfCode = 0
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If (edcPassword.Enabled) Then
        slStr = Trim(edcPassword.Text)
        If (Trim$(tmWkUrf.sPassword) <> slStr) And ((Not imNewRec) Or ((imNewRec) And (slStr <> ""))) Then
            If (Asc(tgSpf.sUsingFeatures2) And STRONGPASSWORD) = STRONGPASSWORD Then
                'Dan allow password to be saved as blank 9-27-09
                'If Not gStrongPassword(slStr) Then
                If Not gStrongPassword(slStr) And Not LenB(slStr) = 0 Then
                    MsgBox "New Password Violates Password Rules", vbOKOnly, "Error"
                    edcPassword.SetFocus
                    Exit Sub
                End If
                ilMatchOld = False
                If StrComp(slStr, tmWkUrf.sOldPassword(0), vbTextCompare) = 0 Then
                    ilMatchOld = True
                End If
                If StrComp(slStr, tmWkUrf.sOldPassword(1), vbTextCompare) = 0 Then
                    ilMatchOld = True
                End If
                If StrComp(slStr, tmWkUrf.sOldPassword(2), vbTextCompare) = 0 Then
                    ilMatchOld = True
                End If
                If ilMatchOld Then
                    MsgBox "New Password cannot match Old Passwords", vbOKOnly, "Error"
                    edcPassword.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
    'Test that Remote User ID is Unique unless it is zero
    If Trim$(edcRemoteID.Text) <> "" Then
        If (tmWkUrf.iRemoteUserID <> Val(edcRemoteID.Text)) Then
            For ilLoop = 0 To UBound(imRemoteID) - 1 Step 1
                If Val(edcRemoteID.Text) = imRemoteID(ilLoop) Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Remote ID in Use", vbOKOnly + vbInformation, "Update"
                    edcRemoteID.SetFocus
                    Exit Sub
                End If
            Next ilLoop
        End If
    End If
    If Not mProgrammaticUserOk() Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    gGetSyncDateTime slSyncDate, slSyncTime
    If sgCPName = cbcSelect.Text Then
        sgCPName = Trim$(edcName.Text)
    End If
    If sgSUName = cbcSelect.Text Then
        sgSUName = Trim$(edcName.Text)
    End If
    If (imNewRec) Or (imUrfIndex < 0) Then
        imNewRec = True
        mCreateUrf  'This sets imUrfCode
    End If
    If sgUrfName = Trim$(tmUrf(imUrfIndex).sName) Then
        sgUrfName = Trim$(edcName.Text)
    End If
    mMoveCtrlToRec
    If imNewRec Then
        ilCode = 0
    Else
        ilCode = tmUrf(imUrfIndex).iCode
    End If
    'For ilLoop = 1 To UBound(tgPopUrf) - 1 Step 1
    For ilLoop = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
        'If (tmUrf(imUrfIndex).iVefCode = tgPopUrf(ilLoop).iVefCode) And (ilCode <> tgPopUrf(ilLoop).iCode) Then
        If (ilCode <> tgPopUrf(ilLoop).iCode) And (tgPopUrf(ilLoop).sDelete <> "Y") Then
            If StrComp(tgPopUrf(ilLoop).sName, tmUrf(imUrfIndex).sName, vbTextCompare) = 0 Then
                Screen.MousePointer = vbDefault
                MsgBox Trim$(edcName.Text) & " previously defined", vbOKOnly + vbInformation, "Update"
                edcName.SetFocus
                Exit Sub
            End If
        End If
    Next ilLoop
    If imNewRec Then
        'ilVefCode = tmUrf(imUrfIndex).iVefCode
        tmUrf(imUrfIndex).iCode = 0
        If imUrfIndex = LBound(tmUrf) Then
            tmUrf(imUrfIndex).sClnMoYr = "M"    'Calendar by month
            tmUrf(imUrfIndex).sClnType = "S"    'Calendar type-standard
            tmUrf(imUrfIndex).sClnLayout = "A"  'Calendar- Across
            tmUrf(imUrfIndex).iClnLeft = 100    'Calendar left
            tmUrf(imUrfIndex).iClnTop = 200 'Calendar top
            tmUrf(imUrfIndex).iClcLeft = 2775   'Traffic.Width - 2775 - 100   '2625 width of calc form
            tmUrf(imUrfIndex).iClcTop = 200
        Else
            tmUrf(imUrfIndex).sClnMoYr = tmUrf(LBound(tmUrf)).sClnMoYr    'Calendar by month
            tmUrf(imUrfIndex).sClnType = tmUrf(LBound(tmUrf)).sClnType    'Calendar type-standard
            tmUrf(imUrfIndex).sClnLayout = tmUrf(LBound(tmUrf)).sClnLayout  'Calendar- Across
            tmUrf(imUrfIndex).iClnLeft = tmUrf(LBound(tmUrf)).iClnLeft   'Calendar left
            tmUrf(imUrfIndex).iClnTop = tmUrf(LBound(tmUrf)).iClnTop 'Calendar top
            tmUrf(imUrfIndex).iClcLeft = tmUrf(LBound(tmUrf)).iClcLeft
            tmUrf(imUrfIndex).iClcTop = tmUrf(LBound(tmUrf)).iClcTop
        End If
        tmUrf(imUrfIndex).sDelete = "N"
        tmUrf(imUrfIndex).iRemoteID = tgUrf(0).iRemoteUserID
        tmUrf(imUrfIndex).iAutoCode = tmUrf(imUrfIndex).iCode
        gUrfEncrypt tmUrf(imUrfIndex)
        ilRet = btrInsert(hmUrf, tmUrf(imUrfIndex), imUrfRecLen, INDEXKEY0)
        If ilRet = BTRV_ERR_DUPLICATE_KEY Then
            Screen.MousePointer = vbDefault
            MsgBox Trim$(edcName.Text) & " for " & Trim$(cbcVehicle.Text) & " previously added", vbOKOnly + vbInformation, "Update"
            edcName.SetFocus
            Exit Sub
        End If
        On Error GoTo cmcUpdateErr
        gBtrvErrorMsg ilRet, "cmcUpdate (btrInsert)", UserOpt
        On Error GoTo 0
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        gUrfDecrypt tmUrf(imUrfIndex)
        Do
            tmUrf(imUrfIndex).iRemoteID = tgUrf(0).iRemoteUserID
            tmUrf(imUrfIndex).iAutoCode = tmUrf(imUrfIndex).iCode
            gPackDate slSyncDate, tmUrf(imUrfIndex).iSyncDate(0), tmUrf(imUrfIndex).iSyncDate(1)
            gPackTime slSyncTime, tmUrf(imUrfIndex).iSyncTime(0), tmUrf(imUrfIndex).iSyncTime(1)
            gUrfEncrypt tmUrf(imUrfIndex)
            ilRet = btrUpdate(hmUrf, tmUrf(imUrfIndex), imUrfRecLen)
            slMsg = "mSaveRec (btrUpdate:User)"
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo cmcUpdateErr
        gBtrvErrorMsg ilRet, "cmcUpdate (btrUpdate)", UserOpt
        On Error GoTo 0
        gUrfDecrypt tmUrf(imUrfIndex)
    Else
        tmSrchKey.iCode = tmUrf(imUrfIndex).iCode
        ilRet = btrGetEqual(hmUrf, tlUrf, imUrfRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        gUrfDecrypt tlUrf
        On Error GoTo cmcUpdateErr
        gBtrvErrorMsg ilRet, "cmcUpdate (btrGetEqual)", UserOpt
        On Error GoTo 0
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        gPackDate slSyncDate, tmUrf(imUrfIndex).iSyncDate(0), tmUrf(imUrfIndex).iSyncDate(1)
        gPackTime slSyncTime, tmUrf(imUrfIndex).iSyncTime(0), tmUrf(imUrfIndex).iSyncTime(1)
        gUrfEncrypt tmUrf(imUrfIndex)
        ilRet = btrUpdate(hmUrf, tmUrf(imUrfIndex), imUrfRecLen)
        If ilRet = BTRV_ERR_DUPLICATE_KEY Then
            Screen.MousePointer = vbDefault
            MsgBox Trim$(edcName.Text) & " for " & Trim$(cbcVehicle.Text) & " name previously existed", vbOKOnly + vbInformation, "Update"
            edcName.SetFocus
            Exit Sub
        End If
        On Error GoTo cmcUpdateErr
        gBtrvErrorMsg ilRet, "cmcUpdate (btrUpdate)", UserOpt
        On Error GoTo 0
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        gUrfDecrypt tmUrf(imUrfIndex)
    End If
    tmWkUrf = tmUrf(imUrfIndex)
    For ilLoop = LBound(tmUrf) To UBound(tmUrf) Step 1
        If ilLoop <> imUrfIndex Then
            tmSrchKey.iCode = tmUrf(ilLoop).iCode
            ilRet = btrGetEqual(hmUrf, tmUrf(ilLoop), imUrfRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            gUrfDecrypt tmUrf(ilLoop)
            On Error GoTo cmcUpdateErr
            gBtrvErrorMsg ilRet, "cmcUpdate (btrGetEqual)", UserOpt
            On Error GoTo 0
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            If tmUrf(ilLoop).sDelete <> "Y" Then
                tmUrf(ilLoop).sName = tmUrf(imUrfIndex).sName
                tmUrf(ilLoop).sOldPassword(2) = tmUrf(imUrfIndex).sOldPassword(2)
                tmUrf(ilLoop).sOldPassword(1) = tmUrf(imUrfIndex).sOldPassword(1)
                tmUrf(ilLoop).sOldPassword(0) = tmUrf(imUrfIndex).sOldPassword(0)
                tmUrf(ilLoop).sPassword = tmUrf(imUrfIndex).sPassword
                tmUrf(ilLoop).iPasswordDate(0) = tmUrf(imUrfIndex).iPasswordDate(0)
                tmUrf(ilLoop).iPasswordDate(1) = tmUrf(imUrfIndex).iPasswordDate(1)
                tmUrf(ilLoop).iRemoteUserID = 0 'tmUrf(imUrfIndex).iRemoteUserID
                tmUrf(ilLoop).iGroupNo = tmUrf(imUrfIndex).iGroupNo
                tmUrf(ilLoop).iMnfHubCode = tmUrf(imUrfIndex).iMnfHubCode
                tmUrf(ilLoop).sBlockRU = "N"    'tmUrf(imUrfIndex).sBlockRU
                tmUrf(ilLoop).sRept = tmUrf(imUrfIndex).sRept
                tmUrf(ilLoop).iVefCode = tmUrf(imUrfIndex).iVefCode
                tmUrf(ilLoop).iDefVeh = tmUrf(imUrfIndex).iDefVeh
                tmUrf(ilLoop).iSlfCode = tmUrf(imUrfIndex).iSlfCode
                tmUrf(ilLoop).sPDFDrvChar = tmUrf(imUrfIndex).sPDFDrvChar
                tmUrf(ilLoop).iPDFDnArrowCnt = tmUrf(imUrfIndex).iPDFDnArrowCnt
                tmUrf(ilLoop).sPrtDrvChar = tmUrf(imUrfIndex).sPrtDrvChar
                tmUrf(ilLoop).iPrtDnArrowCnt = tmUrf(imUrfIndex).iPrtDnArrowCnt
                'Update all list task which are vehicle independent
                'General
                tmUrf(ilLoop).sWin(ADVERTISERSLIST) = tmUrf(imUrfIndex).sWin(ADVERTISERSLIST)
                tmUrf(ilLoop).sWin(AGENCIESLIST) = tmUrf(imUrfIndex).sWin(AGENCIESLIST)
                tmUrf(ilLoop).sWin(COMPETITIVESLIST) = tmUrf(imUrfIndex).sWin(COMPETITIVESLIST)
                tmUrf(ilLoop).sWin(EXCLUSIONSLIST) = tmUrf(imUrfIndex).sWin(EXCLUSIONSLIST)
                tmUrf(ilLoop).sWin(VEHICLEGROUPSLIST) = tmUrf(imUrfIndex).sWin(VEHICLEGROUPSLIST)
                tmUrf(ilLoop).sWin(BUSCATEGORIESLIST) = tmUrf(imUrfIndex).sWin(BUSCATEGORIESLIST)
                tmUrf(ilLoop).sWin(POTENTIALCODESLIST) = tmUrf(imUrfIndex).sWin(POTENTIALCODESLIST)
                tmUrf(ilLoop).sWin(MISSEDREASONSLIST) = tmUrf(imUrfIndex).sWin(MISSEDREASONSLIST)
                tmUrf(ilLoop).sWin(FEEDNAMELIST) = tmUrf(imUrfIndex).sWin(FEEDNAMELIST)
                'Sales
                tmUrf(ilLoop).sWin(SALESSOURCESLIST) = tmUrf(imUrfIndex).sWin(SALESSOURCESLIST)
                tmUrf(ilLoop).sWin(SALESREGIONSLIST) = tmUrf(imUrfIndex).sWin(SALESREGIONSLIST)
                tmUrf(ilLoop).sWin(SALESTEAMSLIST) = tmUrf(imUrfIndex).sWin(SALESTEAMSLIST)
                tmUrf(ilLoop).sWin(SALESPEOPLELIST) = tmUrf(imUrfIndex).sWin(SALESPEOPLELIST)
                tmUrf(ilLoop).sWin(REVENUESETSLIST) = tmUrf(imUrfIndex).sWin(REVENUESETSLIST)
                tmUrf(ilLoop).sWin(BOILERPLATESLIST) = tmUrf(imUrfIndex).sWin(BOILERPLATESLIST)
                tmUrf(ilLoop).sWin(RESEARCHLIST) = tmUrf(imUrfIndex).sWin(RESEARCHLIST)
                tmUrf(ilLoop).sWin(DEMOSLIST) = tmUrf(imUrfIndex).sWin(DEMOSLIST)
                tmUrf(ilLoop).sWin(COMPETITORSLIST) = tmUrf(imUrfIndex).sWin(COMPETITORSLIST)
                tmUrf(ilLoop).sWin(SPLITNETSLIST) = tmUrf(imUrfIndex).sWin(SPLITNETSLIST)
                tmUrf(ilLoop).sWin(PODITEMSLIST) = tmUrf(imUrfIndex).sWin(PODITEMSLIST)
                'Programming
                tmUrf(ilLoop).sWin(EVENTTYPESLIST) = tmUrf(imUrfIndex).sWin(EVENTTYPESLIST)
                tmUrf(ilLoop).sWin(AVAILNAMESLIST) = tmUrf(imUrfIndex).sWin(AVAILNAMESLIST)
                tmUrf(ilLoop).sWin(FEEDTYPESLIST) = tmUrf(imUrfIndex).sWin(FEEDTYPESLIST)
                tmUrf(ilLoop).sWin(GENRESLIST) = tmUrf(imUrfIndex).sWin(GENRESLIST)
                'Copy
                tmUrf(ilLoop).sWin(MEDIADEFINITIONSLIST) = tmUrf(imUrfIndex).sWin(MEDIADEFINITIONSLIST)
                tmUrf(ilLoop).sWin(ANNOUNCERNAMESLIST) = tmUrf(imUrfIndex).sWin(ANNOUNCERNAMESLIST)
                'Accounting
                tmUrf(ilLoop).sWin(TAXTABLELIST) = tmUrf(imUrfIndex).sWin(TAXTABLELIST)
                tmUrf(ilLoop).sWin(ITEMBILLINGTYPESLIST) = tmUrf(imUrfIndex).sWin(ITEMBILLINGTYPESLIST)
                tmUrf(ilLoop).sWin(INVOICESORTLIST) = tmUrf(imUrfIndex).sWin(INVOICESORTLIST)
                tmUrf(ilLoop).sWin(LOCKBOXESLIST) = tmUrf(imUrfIndex).sWin(LOCKBOXESLIST)
                tmUrf(ilLoop).sWin(EDISERVICESLIST) = tmUrf(imUrfIndex).sWin(EDISERVICESLIST)
                tmUrf(ilLoop).sWin(TRANSACTIONSLIST) = tmUrf(imUrfIndex).sWin(TRANSACTIONSLIST)
                'Option
                tmUrf(ilLoop).sWin(SITELIST) = tmUrf(imUrfIndex).sWin(SITELIST)
                tmUrf(ilLoop).sWin(USERLIST) = tmUrf(imUrfIndex).sWin(USERLIST)
                tmUrf(ilLoop).sCredit = tmUrf(imUrfIndex).sCredit
                tmUrf(ilLoop).sPayRate = tmUrf(imUrfIndex).sPayRate
                tmUrf(ilLoop).sMerge = tmUrf(imUrfIndex).sMerge
                tmUrf(ilLoop).sChgCrRt = tmUrf(imUrfIndex).sChgCrRt
                tmUrf(ilLoop).sBouChk = tmUrf(imUrfIndex).sBouChk
                tmUrf(ilLoop).sReprintLogAlert = tmUrf(imUrfIndex).sReprintLogAlert
                tmUrf(ilLoop).sIncompAlert = tmUrf(imUrfIndex).sIncompAlert
                tmUrf(ilLoop).sCompAlert = tmUrf(imUrfIndex).sCompAlert
                tmUrf(ilLoop).sSchAlert = tmUrf(imUrfIndex).sSchAlert
                tmUrf(ilLoop).sHoldAlert = tmUrf(imUrfIndex).sHoldAlert
                tmUrf(ilLoop).sRateCardAlert = tmUrf(imUrfIndex).sRateCardAlert
                tmUrf(ilLoop).sResearchAlert = tmUrf(imUrfIndex).sResearchAlert
                tmUrf(ilLoop).sAvailAlert = tmUrf(imUrfIndex).sAvailAlert
                tmUrf(ilLoop).sCrdChkAlert = tmUrf(imUrfIndex).sCrdChkAlert
                tmUrf(ilLoop).sDeniedAlert = tmUrf(imUrfIndex).sDeniedAlert
                tmUrf(ilLoop).sCrdLimitAlert = tmUrf(imUrfIndex).sCrdLimitAlert
                tmUrf(ilLoop).sMoveAlert = tmUrf(imUrfIndex).sMoveAlert
                tmUrf(ilLoop).sAllowedToBlock = tmUrf(imUrfIndex).sAllowedToBlock
                tmUrf(ilLoop).sShowNRMsg = tmUrf(imUrfIndex).sShowNRMsg
                tmUrf(ilLoop).sWorkToDead = tmUrf(imUrfIndex).sWorkToDead
                tmUrf(ilLoop).sWorkToComp = tmUrf(imUrfIndex).sWorkToComp
                tmUrf(ilLoop).sWorkToHold = tmUrf(imUrfIndex).sWorkToHold
                tmUrf(ilLoop).sWorkToOrder = tmUrf(imUrfIndex).sWorkToOrder
                tmUrf(ilLoop).sCompToIncomp = tmUrf(imUrfIndex).sCompToIncomp
                tmUrf(ilLoop).sCompToDead = tmUrf(imUrfIndex).sCompToDead
                tmUrf(ilLoop).sCompToHold = tmUrf(imUrfIndex).sCompToHold
                tmUrf(ilLoop).sCompToOrder = tmUrf(imUrfIndex).sCompToOrder
                tmUrf(ilLoop).sIncompToDead = tmUrf(imUrfIndex).sIncompToDead
                tmUrf(ilLoop).sIncompToComp = tmUrf(imUrfIndex).sIncompToComp
                tmUrf(ilLoop).sIncompToHold = tmUrf(imUrfIndex).sIncompToHold
                tmUrf(ilLoop).sDeadToWork = tmUrf(imUrfIndex).sDeadToWork
                tmUrf(ilLoop).sHoldToOrder = tmUrf(imUrfIndex).sHoldToOrder
                tmUrf(ilLoop).sReviseCntr = tmUrf(imUrfIndex).sReviseCntr
                tmUrf(ilLoop).sResvType = tmUrf(imUrfIndex).sResvType
                tmUrf(ilLoop).sRemType = tmUrf(imUrfIndex).sRemType
                tmUrf(ilLoop).sDRType = tmUrf(imUrfIndex).sDRType
                tmUrf(ilLoop).sPIType = tmUrf(imUrfIndex).sPIType
                tmUrf(ilLoop).sPSAType = tmUrf(imUrfIndex).sPSAType
                tmUrf(ilLoop).sPromoType = tmUrf(imUrfIndex).sPromoType
                tmUrf(ilLoop).sRefResvType = tmUrf(imUrfIndex).sRefResvType
                tmUrf(ilLoop).sUseComputeCMC = tmUrf(imUrfIndex).sUseComputeCMC
                tmUrf(ilLoop).sRegionCopy = tmUrf(imUrfIndex).sRegionCopy
                tmUrf(ilLoop).sChgPrices = tmUrf(imUrfIndex).sChgPrices
                tmUrf(ilLoop).sChgLnBillPrice = tmUrf(imUrfIndex).sChgLnBillPrice
                tmUrf(ilLoop).sActFlightButton = tmUrf(imUrfIndex).sActFlightButton
                gPackDate slSyncDate, tmUrf(ilLoop).iSyncDate(0), tmUrf(ilLoop).iSyncDate(1)
                gPackTime slSyncTime, tmUrf(ilLoop).iSyncTime(0), tmUrf(ilLoop).iSyncTime(1)
                gUrfEncrypt tmUrf(ilLoop)
                ilRet = btrUpdate(hmUrf, tmUrf(ilLoop), imUrfRecLen)
                On Error GoTo cmcUpdateErr
                gBtrvErrorMsg ilRet, "cmcUpdate (btrUpdate)", UserOpt
                On Error GoTo 0
                If imTerminate Then
                    cmcCancel_Click
                    Exit Sub
                End If
                gUrfDecrypt tmUrf(ilLoop)
                'For ilUrf = 1 To UBound(tgPopUrf) - 1 Step 1
                For ilUrf = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
                    If tgPopUrf(ilUrf).iCode = tmUrf(ilLoop).iCode Then
                        tgPopUrf(ilUrf) = tmUrf(ilLoop)
                        Exit For
                    End If
                Next ilUrf
            End If
        Else
            If imNewRec Then
                tgPopUrf(UBound(tgPopUrf)) = tmUrf(imUrfIndex)
                'ReDim Preserve tgPopUrf(1 To UBound(tgPopUrf) + 1) As URF
                ReDim Preserve tgPopUrf(0 To UBound(tgPopUrf) + 1) As URF
                ilUrfCode = UBound(tgPopUrf) - 1
            Else
                'For ilUrf = 1 To UBound(tgPopUrf) - 1 Step 1
                For ilUrf = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
                    If tgPopUrf(ilUrf).iCode = tmUrf(ilLoop).iCode Then
                        tgPopUrf(ilUrf) = tmUrf(ilLoop)
                        ilUrfCode = ilUrf
                        Exit For
                    End If
                Next ilUrf
            End If
        End If
    Next ilLoop
    imIgnoreChg = YES
    mModelPop
    mSetCommands
    slName = Trim$(tmUrf(imUrfIndex).sName)
    sgUrfStamp = ""
    ilRet = csiSetStamp("URF", sgUrfStamp)
    gUrfRead UserOpt, slName, False, tmUrf(), imIncludeDormant
    cbcSelect.Tag = ""
    cbcSelect.Clear
    If imNewRec Then
        mPopulate
        gFindMatch slName, 0, cbcSelect
        If gLastFound(cbcSelect) >= 0 Then
            cbcSelect.ListIndex = gLastFound(cbcSelect)
        End If

    Else
        mPopulate
        gFindMatch slName, 0, cbcSelect
        If gLastFound(cbcSelect) >= 0 Then
            cbcSelect.ListIndex = gLastFound(cbcSelect)
        End If
        'If (StrComp(slName, sgCPName, 1) = 0) Or (StrComp(slName, sgSUName, 1) = 0) Then
        If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
            cbcSelect.SetFocus
        End If
    End If
    
    '07-13-15 Send the EDS either a sigle new or updated user.  Note if there is no email address defined it will not be sent to EDS
    '08-25-15 Verified new and update
    If bgEDSIsActive Then
        ilRet = gAddOrUpdateSingleNetworkUser(ilUrfCode)
    End If
    
    imIgnoreChg = NO
    Screen.MousePointer = vbDefault
    Exit Sub
cmcUpdateErr:
    On Error GoTo 0
    imTerminate = True
    Resume Next
End Sub

Private Sub cmcUpdate_GotFocus()
    If Len(Trim$(edcName.Text)) = 0 Then
        Beep
        edcName.SetFocus
        Exit Sub
    End If
'    If (edcPassword.Enabled) And Len(Trim$(edcPassword.Text)) < 4 Then
'        Beep
'        edcPassword.SetFocus
'        Exit Sub
'    End If
    gCtrlGotFocus cmcUpdate
End Sub

Private Sub edcCity_Change()
    mSetCommands
End Sub

Private Sub edcCity_GotFocus()
    imIgnoreChg = YES
    If (imFirstTime = YES) And (imNewRec) Then
'        mInitCtrlFields
'        mMoveCtrlToRec
        cmcUpdate.Enabled = True
    End If
    imFirstTime = NO
    gCtrlGotFocus edcCity
    imIgnoreChg = NO
End Sub

Private Sub edcCity_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcEMail_Change()
    mSetCommands
End Sub

Private Sub edcEMail_GotFocus()
    imIgnoreChg = YES
    If (imFirstTime = YES) And (imNewRec) Then
'        mInitCtrlFields
'        mMoveCtrlToRec
        cmcUpdate.Enabled = True
    End If
    imFirstTime = NO
    gCtrlGotFocus edcEMail
    imIgnoreChg = NO
End Sub

Private Sub edcEMail_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcGroupNo_Change()
    Dim slStr As String

    slStr = Trim$(edcGroupNo.Text)
    If (slStr <> "") And (Val(slStr) <> 0) Then
        cbcVehicle.ListIndex = 0
        cbcDefVeh.Enabled = True
    End If
    mSetCommands
End Sub
Private Sub edcGroupNo_GotFocus()
    imIgnoreChg = YES
    If (imFirstTime = YES) And (imNewRec) Then
'        mInitCtrlFields
'        mMoveCtrlToRec
        cmcUpdate.Enabled = True
    End If
    imFirstTime = NO
    gCtrlGotFocus edcGroupNo
    imIgnoreChg = NO
End Sub
Private Sub edcGroupNo_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim slStr As String
    If (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcGroupNo.Text
    slStr = Left$(slStr, edcGroupNo.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcGroupNo.SelStart - edcGroupNo.SelLength)
    If gCompNumberStr(slStr, "100") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcName_Change()
    mSetCommands
End Sub
Private Sub edcName_GotFocus()
    imIgnoreChg = YES
    If (imFirstTime = YES) And (imNewRec) Then
'        mInitCtrlFields
'        mMoveCtrlToRec
        cmcUpdate.Enabled = True
        If edcName.Text <> "" Then
            If edcPassword.Enabled Then
                edcPassword.SetFocus
            Else
                edcRept.SetFocus
            End If
        End If
    End If
    imFirstTime = NO
    gCtrlGotFocus edcName
    imIgnoreChg = NO
End Sub
Private Sub edcName_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcPassword_Change()
    mSetCommands
End Sub
Private Sub edcPassword_GotFocus()
    imIgnoreChg = YES
    If (imFirstTime = YES) And (imNewRec) Then
'        mInitCtrlFields
'        mMoveCtrlToRec
        cmcUpdate.Enabled = True
    End If
    imFirstTime = NO
    gCtrlGotFocus edcPassword
    imIgnoreChg = NO
End Sub
Private Sub edcPassword_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    If (Len(edcPassword.Text) = 0) And (KeyAscii = KEYASTERISK) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcPDF_Change(Index As Integer)
    mSetCommands
End Sub
Private Sub edcPDF_GotFocus(Index As Integer)
    imIgnoreChg = YES
    If (imFirstTime = YES) And (imNewRec) Then
'        mInitCtrlFields
'        mMoveCtrlToRec
        cmcUpdate.Enabled = True
    End If
    imFirstTime = NO
    gCtrlGotFocus edcPDF(Index)
    imIgnoreChg = NO
End Sub
Private Sub edcPDF_KeyPress(Index As Integer, KeyAscii As Integer)
    If (Index = 1) Or (Index = 3) Or (Index = 5) Then
        'Filter characters (allow only BackSpace, numbers 0 thru 9
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub edcPhone_Change()
    mSetCommands
End Sub

Private Sub edcPhone_GotFocus()
    imIgnoreChg = YES
    If (imFirstTime = YES) And (imNewRec) Then
'        mInitCtrlFields
'        mMoveCtrlToRec
        cmcUpdate.Enabled = True
    End If
    imFirstTime = NO
    gCtrlGotFocus edcPhone
    imIgnoreChg = NO
End Sub

Private Sub edcPhone_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcRemoteID_Change()
    mSetCommands
End Sub
Private Sub edcRemoteID_GotFocus()
    imIgnoreChg = YES
    If (imFirstTime = YES) And (imNewRec) Then
'        mInitCtrlFields
'        mMoveCtrlToRec
        cmcUpdate.Enabled = True
    End If
    imFirstTime = NO
    gCtrlGotFocus edcRemoteID
    imIgnoreChg = NO
End Sub
Private Sub edcRemoteID_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim slStr As String
    If (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcRemoteID.Text
    slStr = Left$(slStr, edcRemoteID.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcRemoteID.SelStart - edcRemoteID.SelLength)
    If gCompNumberStr(slStr, "2000") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcRept_Change()
    mSetCommands
End Sub
Private Sub edcRept_GotFocus()
    imIgnoreChg = YES
    If (imFirstTime = YES) And (imNewRec) Then
'        mInitCtrlFields
'        mMoveCtrlToRec
        cmcUpdate.Enabled = True
    End If
    imFirstTime = NO
    gCtrlGotFocus edcRept
    imIgnoreChg = NO
End Sub
Private Sub edcRept_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
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
    'dan m 10/14/11. everyone can update. don't bother enabling here, as msetCommands is called (over and over) to do the same thing.
    imUpdateAllowed = True
    edcPassword.PasswordChar = "*"
    If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
        edcPassword.PasswordChar = ""
    End If
'    If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
'        plcMain.Enabled = True
'        plcName.Enabled = True
'        imUpdateAllowed = True
'        edcPassword.PasswordChar = ""
'    Else
'        edcPassword.PasswordChar = "*"
'        If igWinStatus(USERLIST) = 1 Then
'            plcMain.Enabled = False
'            plcName.Enabled = False
'            imUpdateAllowed = False
'        Else
'            plcMain.Enabled = True
'            plcName.Enabled = True
'            imUpdateAllowed = True
'        End If
'    End If
    If imIncludeDormant Then
        plcState.Visible = True
    Else
        plcState.Visible = False
    End If
    If imUpdateAllowed Then
        frcSet.Visible = True
    Else
        frcSet.Visible = False
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    UserOpt.Refresh
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

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    If igLogActivityStatus = 32123 Then
        igLogActivityStatus = -32123
        gUserActivityLog "", ""
    End If
    
    btrExtClear hmUrf   'Clear any previous extend operation
    ilRet = btrClose(hmUrf)
    btrDestroy hmUrf
    btrExtClear hmSnf   'Clear any previous extend operation
    ilRet = btrClose(hmSnf)
    btrDestroy hmSnf
    btrExtClear hmCef   'Clear any previous extend operation
    ilRet = btrClose(hmCef)
    btrDestroy hmCef
'    Erase tgRnfList
    Erase tmHubCode
    Erase tgSnfCode
    Erase tmUrf
    Erase imSvUrfIndex
    Erase imModelCode
    Erase imRemoteID
    
    Set UserOpt = Nothing   'Remove data segment

    End

End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: clear controls modular         *
'*                                                     *
'*******************************************************
Private Sub mClearCtrlFields()
'
'   mClearCtrlFields
'   Where:
'
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    imIgnoreChg = YES

    If imIncludeDormant Then
        rbcState(0).Value = True
    End If
    tmWkUrf.sName = ""  'clear any previous name
    edcName.Text = ""
    edcPassword.Text = ""
    edcRemoteID.Text = ""
    edcGroupNo.Text = ""
    ckcBlockRU.Value = vbUnchecked
    edcRept.Text = ""
    edcPhone.Text = ""
    edcCity.Text = ""
    edcEMail.Text = ""
    smEMail = ""
    lmEMailCefCode = 0
    cbcRptSet.ListIndex = -1
    smRptSet = ""
    cbcHub.ListIndex = -1
    smHub = ""
    cbcVehicle.ListIndex = -1
    imVehSelectedIndex = -1
    smVeh = ""
    cbcDefVeh.ListIndex = -1
    imDVSelectedIndex = -1
    smDefVeh = ""
    cbcSalesperson.ListIndex = -1
    imSPSelectedIndex = -1
    smSalesPerson = ""
    cbcModel.ListIndex = -1
    edcPDF(0).Text = ""
    edcPDF(1).Text = ""
    edcPDF(2).Text = ""
    edcPDF(3).Text = ""
    edcPDF(4).Text = ""
    edcPDF(5).Text = ""
    For ilLoop = RATECARDSJOB To REPORTSJOB - 1 Step 1
        ilIndex = imWinMap(ilLoop)
        If ilIndex >= 0 Then
            'pbcJobs(ilIndex).Cls
            If plcJobs(ilIndex).Enabled Then
                plcJobs(ilIndex).BackColor = GRAY
            Else
                plcJobs(ilIndex).BackColor = Red
            End If
        End If
    Next ilLoop
    For ilLoop = LBound(imSelectedFields) To UBound(imSelectedFields) Step 1
        'pbcSelectedFields(ilLoop).Cls
        plcSelectedFields(ilLoop).BackColor = GRAY
    Next ilLoop
    'General
    For ilLoop = VEHICLESLIST To USERLIST Step 1
        ilIndex = imWinMap(ilLoop)
        If ilIndex >= 0 Then
            'pbcLists(ilIndex).Cls
            plcLists(ilIndex).BackColor = GRAY
        End If
    Next ilLoop
    For ilLoop = LBound(imTypeFields) To UBound(imTypeFields) Step 1
        plcTypes(ilLoop).BackColor = GRAY
    Next ilLoop
    For ilLoop = LBound(imAlerts) To UBound(imAlerts) Step 1
        plcAlerts(ilLoop).BackColor = GRAY
    Next ilLoop
    imIgnoreChg = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateUrf                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create an URF record           *
'*                                                     *
'*******************************************************
Private Sub mCreateUrf()
    Dim ilBound As Integer
    If imUrfIndex < 0 Then
        If cbcSelect.ListIndex > 0 Then
            ilBound = UBound(tmUrf) + 1
        Else
            ilBound = 0
        End If
        ReDim Preserve tmUrf(0 To ilBound)
        imUrfIndex = ilBound
    End If
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
'
'   ANmInit
'   Where:
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilLastAlert As Integer

    Screen.MousePointer = vbHourglass
    imIgnoreChg = YES
    imFirstActivate = True
    imTerminate = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    'mInitBox
    UserOpt.Height = cmcDone.Top + 5 * cmcDone.Height / 3 - 60
    gCenterStdAlone UserOpt
    UserOpt.Top = UserOpt.Top - 120
    imStartX = 60
    imStartY = -20
    imIAdj = 35
    plcName.Move 45, 420
    'plcMain.Move plcName.Left, plcName.Top + plcName.Height + 15
    plcMain.Move 210, plcName.Top + plcName.Height + 30
    frcSelect.Move 135, 375, frcSelect.Width, 1965
    frcJobs.Move frcSelect.Left + frcSelect.Width + (plcMain.Width - frcJobs.Width - frcSelect.Left - frcSelect.Width) \ 2, frcSelect.Top
    If (Asc(tgSpf.sSportInfo) And USINGSPORTS) <> USINGSPORTS Then
        frcSports.Visible = False
    End If
    If (Asc(tgSpf.sUsingFeatures) And USINGLIVELOG) <> USINGLIVELOG Then 'Using Live Log
        frcLiveLog.Visible = False
        If frcSports.Visible = True Then
            frcSports.Move frcLiveLog.Left, frcLiveLog.Top
        End If
    End If

    frcLists.Move frcSelect.Left + frcSelect.Width + (plcMain.Width - frcLists.Width - frcSelect.Left - frcSelect.Width) \ 2, frcSelect.Top
    frcGeneral.Move frcSelect.Left + frcSelect.Width + (plcMain.Width - frcGeneral.Width - frcSelect.Left - frcSelect.Width) \ 2, frcSelect.Top
    frcAlerts.Move frcSelect.Left + frcSelect.Width + (plcMain.Width - frcAlerts.Width - frcSelect.Left - frcSelect.Width) \ 2, frcSelect.Top
    frcStatus.Move frcSelect.Left + frcSelect.Width + (plcMain.Width - frcStatus.Width - frcSelect.Left - frcSelect.Width) \ 2, frcSelect.Top
    frcTypes.Move frcSelect.Left + frcSelect.Width + (plcMain.Width - frcTypes.Width - frcSelect.Left - frcSelect.Width) \ 2, frcSelect.Top
    frcPDF.Move frcSelect.Left + frcSelect.Width + (plcMain.Width - frcPDF.Width - frcSelect.Left - frcSelect.Width) \ 2, frcSelect.Top
    'UserOpt.Show
    Screen.MousePointer = vbHourglass

    ReDim imRemoteID(0 To 0) As Integer
    'imcHelp.Picture = Traffic!imcHelp.Picture
    For ilLoop = LBound(imWinMap) To UBound(imWinMap) Step 1
        imWinMap(ilLoop) = -1
    Next ilLoop
    'Set to index value of the control
    imWinMap(RATECARDSJOB) = 0
    imWinMap(PROPOSALSJOB) = 1
    imWinMap(CONTRACTSJOB) = 2
    imWinMap(PROGRAMMINGJOB) = 3
    imWinMap(SPOTSJOB) = 4
    imWinMap(COPYJOB) = 5
    imWinMap(LOGSJOB) = 6
    imWinMap(POSTLOGSJOB) = 7
    imWinMap(INVOICESJOB) = 8
    imWinMap(COLLECTIONSJOB) = 9
    imWinMap(BUDGETSJOB) = 10
    imWinMap(SLSPCOMMSJOB) = 11
    imWinMap(STATIONFEEDJOB) = 12
    imWinMap(FEEDJOB) = 13
    'General
    imWinMap(VEHICLESLIST) = 0
    imWinMap(VEHICLEGROUPSLIST) = 27
    imWinMap(AGENCIESLIST) = 1
    imWinMap(ADVERTISERSLIST) = 2
    imWinMap(COMPETITIVESLIST) = 3
    imWinMap(BUSCATEGORIESLIST) = 28
    imWinMap(POTENTIALCODESLIST) = 29
    imWinMap(EXCLUSIONSLIST) = 25
    imWinMap(FEEDNAMELIST) = 33
    imWinMap(MISSEDREASONSLIST) = 10
    'Sales
    imWinMap(SALESSOURCESLIST) = 5
    imWinMap(SALESREGIONSLIST) = 6
    imWinMap(SALESOFFICESLIST) = 7
    imWinMap(SALESTEAMSLIST) = 8
    imWinMap(SALESPEOPLELIST) = 9
    imWinMap(RESEARCHLIST) = 30
    imWinMap(DEMOSLIST) = 31
    imWinMap(COMPETITORSLIST) = 32
    imWinMap(REVENUESETSLIST) = 4
    imWinMap(BOILERPLATESLIST) = 23
    imWinMap(SPLITNETSLIST) = 35
    imWinMap(PODITEMSLIST) = 36
    'Programming
    imWinMap(EVENTTYPESLIST) = 15
    imWinMap(EVENTNAMESLIST) = 16
    imWinMap(AVAILNAMESLIST) = 17
    imWinMap(FEEDTYPESLIST) = 18
    imWinMap(GENRESLIST) = 19
    'Copy
    imWinMap(MEDIADEFINITIONSLIST) = 20
    imWinMap(ANNOUNCERNAMESLIST) = 21
    'Accounting
    imWinMap(TAXTABLELIST) = 34
    imWinMap(ITEMBILLINGTYPESLIST) = 11
    imWinMap(INVOICESORTLIST) = 12
    imWinMap(LOCKBOXESLIST) = 13
    imWinMap(EDISERVICESLIST) = 14
    imWinMap(TRANSACTIONSLIST) = 26
    'Options
    imWinMap(SITELIST) = 22
    imWinMap(USERLIST) = 24

    'If Not Radio, then hide Feed Name
    If tgSpf.sSystemType <> "R" Then
        plcLists(33).Visible = False
        'plcJobs(12).Visible = True
        plcJobs(12).Visible = False
        plcJobs(13).Visible = False
    Else
        plcLists(33).Visible = True
        plcJobs(12).Visible = False
        '12/24/12: Removed FeedSpot because of Out of Memory errors
        'plcJobs(13).Visible = True
        plcJobs(13).Visible = False
        '12/24/12: End Point
    End If
    If (Asc(tgSaf(0).sFeatures5) And PROGRAMMATICALLOWED) = PROGRAMMATICALLOWED Then 'Programmatic Buy Allowed
        plcTypes(6).Visible = True
        smPrgmmaticAllow = "Y"
    Else
        plcTypes(6).Visible = False
        smPrgmmaticAllow = "N"
    End If
    For ilLoop = LBound(imAlerts) To UBound(imAlerts) Step 1
        If plcAlerts(ilLoop).Visible = False Then
            For ilIndex = UBound(imAlerts) To ilLoop + 1 Step -1
                plcAlerts(ilIndex).Top = plcAlerts(ilIndex - 1).Top
            Next ilIndex
        Else
            ilLastAlert = ilLoop
        End If
    Next ilLoop
    frcAlerts.Height = plcAlerts(ilLastAlert).Top + 2 * plcAlerts(ilLastAlert).Height

    hmUrf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    ReDim tmUrf(0) As URF
    imUrfRecLen = Len(tmUrf(0)) 'btrRecordLength(hlUrf)  'Get and save record length
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen):" & "Urf.Btr", UserOpt
    On Error GoTo 0
    hmSnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSnf, "", sgDBPath & "Snf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    ReDim tgSnfCode(0) As SNFCODE
    imSnfRecLen = Len(tmSnf) 'btrRecordLength(hlUrf)  'Get and save record length
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen):" & "Snf.Btr", UserOpt
    On Error GoTo 0

    hmCef = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCefRecLen = Len(tmCef) 'btrRecordLength(hlUrf)  'Get and save record length
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen):" & "Cef.Btr", UserOpt
    On Error GoTo 0
    If tgUrf(0).iCode <= 2 Then
        imIncludeDormant = True
    Else
        imIncludeDormant = False
    End If
    imFirstFocus = True
    imPopReqd = False
    ilRet = gObtainVef()
    cbcVehicle.Clear
    cbcDefVeh.Clear
    mVehPop
    cbcSalesperson.Clear 'Force population
    mSlfPop
    Screen.MousePointer = vbHourglass
    mSnfPop
    cbcModel.Clear
    mModelPop
    If imTerminate Then
        Exit Sub
    End If
    mHubPop
    Screen.MousePointer = vbHourglass
    imUrfIndex = -1
    ReDim imSvUrfIndex(0 To 0) As Integer
    imSvUrfIndex(0) = -1
    For ilLoop = LBound(imWin) To UBound(imWin) Step 1
        If (ilLoop = SITELIST) Or (ilLoop = USERLIST) Then
            imWin(ilLoop) = 0
        Else
            imWin(ilLoop) = 1
        End If
    Next ilLoop
    For ilLoop = LBound(imSelectedFields) To UBound(imSelectedFields) Step 1
        If (ilLoop = 4) Or (ilLoop = 5) Or (ilLoop = 6) Or (ilLoop = 7) Or (ilLoop = 8) Or (ilLoop = 14) Or (ilLoop = 15) Or (ilLoop = 16) Or (ilLoop = 17) Or (ilLoop = 19) Then
            imSelectedFields(ilLoop) = 0
        Else
            imSelectedFields(ilLoop) = 1
        End If
    Next ilLoop
    For ilLoop = LBound(imTypeFields) To UBound(imTypeFields) Step 1
        imTypeFields(ilLoop) = 1
    Next ilLoop
    For ilLoop = LBound(imAlerts) To UBound(imAlerts) Step 1
        imAlerts(ilLoop) = 1    'No
    Next ilLoop
    ckcWStatus(0).Value = vbUnchecked
    ckcWStatus(1).Value = vbUnchecked
    ckcWStatus(2).Value = vbUnchecked
    ckcWStatus(3).Value = vbUnchecked
    ckcCStatus(0).Value = vbUnchecked
    ckcCStatus(1).Value = vbUnchecked
    ckcCStatus(2).Value = vbUnchecked
    ckcCStatus(3).Value = vbUnchecked
    ckcIStatus(0).Value = vbUnchecked
    ckcIStatus(1).Value = vbUnchecked
    ckcIStatus(2).Value = vbUnchecked
    ckcIStatus(3).Value = vbUnchecked
    ckcDStatus(0).Value = vbUnchecked
    ckcHStatus(0).Value = vbUnchecked
    rbcReviseCntr(1).Value = False
    rbcReviseCntr(0).Value = False
    If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
        cbcHub.Visible = True
        cmcHub.Visible = True
    Else
        cbcHub.Visible = False
        cmcHub.Visible = False
    End If
    imAltered = False
    imInSelect = False
    smVehicle = ""
    smSalesPerson = ""
    smVeh = ""
    smDefVeh = ""
    imBypassSetting = False
'    rbcOption(1).Value = True   'Show previously defined Vehicles
    'UserOpt.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    'gCenterModalForm UserOpt
    imIgnoreChg = YES
    'mLoadNewSelected
    mSetCommands
    imIgnoreChg = NO
    'smSelFields(0) = "Today's Rate Card Grid Level"
    'smSelFields(1) = "Spot Prices in Post Log"
    'smSelFields(2) = "Credit Restrictions in Advertiser and Agency"
    'smSelFields(3) = "Payment Rating in Advertiser and Agency"
    'smSelFields(4) = "Merge Operation on Screens"
    'smSelFields(5) = "Hide Spots"
    'smSelFields(6) = "Change Billed Spots in Post Log"
    'smSelFields(7) = "Change Contracts in Past with Unbilled Spots"
    'vbcSelFields.Value = 0
    'vbcSelFields_Change
    smLastModel = ""
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitCtrlFields                 *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInitCtrlFields()
'
'   mInitCtrlFields
'   Where:
'
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilMapIndex As Integer
    imIgnoreChg = YES
    If imUrfIndex <= 0 Then
        edcPassword.Text = ""
        edcRemoteID.Text = ""
        edcGroupNo.Text = ""
        ckcBlockRU.Value = vbUnchecked
        edcRept.Text = ""
        edcPhone.Text = ""
        edcCity.Text = ""
        edcEMail.Text = ""
        smEMail = ""
        lmEMailCefCode = 0
        cbcRptSet.ListIndex = -1
        smRptSet = ""
        cbcHub.ListIndex = -1
        smHub = ""
        cbcVehicle.ListIndex = -1
        imVehSelectedIndex = -1
        smVeh = ""
        cbcDefVeh.ListIndex = -1
        imDVSelectedIndex = -1
        smDefVeh = ""
        cbcSalesperson.ListIndex = -1
        imSPSelectedIndex = -1
        smSalesPerson = ""
        cbcModel.ListIndex = -1
    End If
    edcPDF(0).Text = ""
    edcPDF(1).Text = ""
    edcPDF(2).Text = ""
    edcPDF(3).Text = ""
    edcPDF(4).Text = ""
    edcPDF(5).Text = ""
    For ilLoop = LBound(imWin) To UBound(imWin) Step 1
        If (ilLoop = SITELIST) Or (ilLoop = USERLIST) Then
            imWin(ilLoop) = 0
        Else
            imWin(ilLoop) = 1
        End If
    Next ilLoop
    For ilLoop = LBound(imSelectedFields) To UBound(imSelectedFields) Step 1
        imSelectedFields(ilLoop) = 1
        If (ilLoop = 4) Or (ilLoop = 5) Or (ilLoop = 6) Or (ilLoop = 7) Or (ilLoop = 8) Or (ilLoop = 14) Or (ilLoop = 15) Or (ilLoop = 16) Or (ilLoop = 17) Or (ilLoop = 19) Then   'Merge
            imSelectedFields(ilLoop) = 0
        End If
    Next ilLoop
    For ilLoop = LBound(imTypeFields) To UBound(imTypeFields) Step 1
        imTypeFields(ilLoop) = 1    'View
    Next ilLoop
    For ilLoop = LBound(imAlerts) To UBound(imAlerts) Step 1
        imAlerts(ilLoop) = 1    'No
    Next ilLoop
    ckcWStatus(0).Value = vbUnchecked
    ckcWStatus(1).Value = vbUnchecked
    ckcWStatus(2).Value = vbUnchecked
    ckcWStatus(3).Value = vbUnchecked
    ckcCStatus(0).Value = vbUnchecked
    ckcCStatus(1).Value = vbUnchecked
    ckcCStatus(2).Value = vbUnchecked
    ckcCStatus(3).Value = vbUnchecked
    ckcIStatus(0).Value = vbUnchecked
    ckcIStatus(1).Value = vbUnchecked
    ckcIStatus(2).Value = vbUnchecked
    ckcIStatus(3).Value = vbUnchecked
    ckcDStatus(0).Value = vbUnchecked
    ckcHStatus(0).Value = vbUnchecked
    rbcReviseCntr(1).Value = False
    rbcReviseCntr(0).Value = False
    For ilLoop = RATECARDSJOB To REPORTSJOB - 1 Step 1
        ilIndex = imWinMap(ilLoop)
        If ilIndex >= 0 Then
            'pbcJobs(ilIndex).Cls
            'pbcJobs(ilIndex).CurrentX = imStartX
            'pbcJobs(ilIndex).CurrentY = imStartY
            If plcJobs(ilIndex).Enabled Then
                plcJobs(ilIndex).BackColor = Yellow 'Print "V"
            Else
                plcJobs(ilIndex).BackColor = Red
            End If
        End If
    Next ilLoop
    For ilLoop = LBound(imSelectedFields) To UBound(imSelectedFields) Step 1
        'pbcSelectedFields_Paint ilLoop
        ilMapIndex = ilLoop
        If imSelectedFields(ilLoop) = 0 Then
            plcSelectedFields(ilMapIndex).BackColor = Red    'Print "H"
        ElseIf imSelectedFields(ilLoop) = 1 Then
            plcSelectedFields(ilMapIndex).BackColor = Yellow 'Print "V"
        ElseIf imSelectedFields(ilLoop) = 2 Then
            plcSelectedFields(ilMapIndex).BackColor = GREEN   'Print "I"
        Else
            plcSelectedFields(ilMapIndex).BackColor = GRAY   'Print ""
        End If
    Next ilLoop
    For ilLoop = LBound(imTypeFields) To UBound(imTypeFields) Step 1
        'pbcTypes_Paint ilLoop
        ilMapIndex = ilLoop
        If imTypeFields(ilLoop) = 0 Then
            plcTypes(ilMapIndex).BackColor = Red 'Print "H"
        ElseIf imTypeFields(ilLoop) = 1 Then
            plcTypes(ilMapIndex).BackColor = Yellow  'Print "V"
        ElseIf imTypeFields(ilLoop) = 2 Then
            plcTypes(ilMapIndex).BackColor = GREEN   'Print "I"
        Else
            plcTypes(ilMapIndex).BackColor = GRAY    'Print ""
        End If
    Next ilLoop
    For ilLoop = LBound(imAlerts) To UBound(imAlerts) Step 1
        'pbcAlerts_Paint ilLoop
        ilMapIndex = ilLoop
        If imAlerts(ilLoop) = 0 Then
            plcAlerts(ilMapIndex).BackColor = GREEN
        ElseIf imAlerts(ilLoop) = 1 Then
            plcAlerts(ilMapIndex).BackColor = Red
        Else
            plcAlerts(ilMapIndex).BackColor = GRAY
        End If
    Next ilLoop
    'If vbcSelFields.Value = 0 Then
    '    vbcSelFields_Change
    'Else
    '    vbcSelFields.Value = 0
    'End If
    'General
    For ilLoop = VEHICLESLIST To USERLIST Step 1
        ilMapIndex = imWinMap(ilLoop)
        If ilMapIndex >= 0 Then
            'pbcLists(ilIndex).Cls
            'pbcLists(ilIndex).CurrentX = imStartX
            'pbcLists(ilIndex).CurrentY = imStartY
            If (ilLoop = SITELIST) Or (ilLoop = USERLIST) Then
                'pbcLists(ilIndex).Print "H"
                plcLists(ilMapIndex).BackColor = Red
            Else
                'pbcLists(ilIndex).Print "V"
                plcLists(ilMapIndex).BackColor = Yellow
            End If
        End If
    Next ilLoop
    imIgnoreChg = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mModelPop                       *
'*                                                     *
'*             Created:5/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Model combo control   *
'*                                                     *
'*******************************************************
Private Sub mModelPop()
    Dim hlUrf As Integer    'User Option file handle
    Dim tlUrf As URF        'Local record image of user record
    Dim slStamp As String   'Current time stamp
    Dim ilRet As Integer    '
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim slUserVehName As String
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer  'List box index, used with imModelCode to save the code number
    Dim ilUrf As Integer

    slStamp = gFileDateTime(sgDBPath & "Urf.btr")
    If cbcModel.Tag <> "" Then
        If slStamp = cbcModel.Tag Then
            Exit Sub
        End If
    End If
    cbcModel.Tag = slStamp
'    If (tgSpf.sSSellNet = "Y") Or (tgSpf.sSDelNet = "Y") Then
'        ilFilter(0) = NOFILTER
'        slFilter(0) = ""
'        ilOffset(0) = 0
'    Else
'        ilFilter(0) = CHARFILTER
'        slFilter(0) = "C"
'        ilOffset(0) = gFieldOffset("Vef", "VefType") '167
'    End If
'    'ilRet = gIMoveListBox(UserOpt, cbcModel, Traffic!lbcVehicle, "Vef.btr", gFieldOffset("Vef", "VefName"), 20, ilFilter(), slFilter(), ilOffset())
'    ilRet = gIMoveListBox(UserOpt, cbcModel, tgVehicle(), sgVehicleTag, "Vef.btr", gFieldOffset("Vef", "VefName"), 20, ilFilter(), slFilter(), ilOffset())
'    If ilRet <> CP_MSG_NOPOPREQ Then
'        On Error GoTo mModelPopErr
'        gCPErrorMsg ilRet, "mModelPop (gIMoveListBox: Vehicle)", UserOpt
'        On Error GoTo 0
'    End If
'    hlUrf = CBtrvTable(ONEHANDLE)
'    ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mModelPopErr
'    gBtrvErrorMsg ilRet, "mModelPop (btrOpen):" & "Urf.Btr", UserOpt
'    On Error GoTo 0
'    ilRecLen = Len(tlUrf)  'btrRecordLength(hlUrf)  'Get and save record length
'    llNoRec = btrRecords(hlUrf) 'Obtain number of records
    ilRet = gObtainUrf()
    ReDim imModelCode(0 To UBound(tgPopUrf) - 1) As Integer
    ReDim imRemoteID(0 To 0) As Integer
    ilIndex = 0
    cbcModel.Clear
'    ilRet = btrGetFirst(hlUrf, tlUrf, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'    Do While ilRet = BTRV_ERR_NONE
'        gUrfDecrypt tlUrf
    'For ilUrf = 1 To UBound(tgPopUrf) - 1 Step 1
    For ilUrf = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
        tlUrf = tgPopUrf(ilUrf)
        If (tlUrf.iCode > 2) And (tlUrf.sDelete <> "Y") Then 'Bypass counterpoint and Guide
            ilFound = False
            slUserVehName = Trim$(tlUrf.sName)
            If tlUrf.iVefCode > 0 Then
                ilLoop = gBinarySearchVef(tlUrf.iVefCode)
                If ilLoop <> -1 Then
                    ilFound = True
                    slUserVehName = slUserVehName & "/" & Trim$(tgMVef(ilLoop).sName)
                End If
            Else
                ilFound = True
                slUserVehName = slUserVehName & "/" & "[All Vehicles]"
            End If
            If ilFound Then
                cbcModel.AddItem slUserVehName & "|" & Trim$(Str$(tlUrf.iCode))
                'imModelCode(cbcModel.NewIndex) = tlUrf.iCode
                ilIndex = ilIndex + 1
            End If
            ilFound = False
            For ilLoop = 0 To UBound(imRemoteID) - 1 Step 1
                If tlUrf.iRemoteUserID = imRemoteID(ilLoop) Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                imRemoteID(UBound(imRemoteID)) = tlUrf.iRemoteUserID
                ReDim Preserve imRemoteID(0 To UBound(imRemoteID) + 1) As Integer
            End If
        End If
'        ilRet = btrGetNext(hlUrf, tlUrf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'    Loop
    Next ilUrf
    For ilLoop = 0 To cbcModel.ListCount - 1 Step 1
        slNameCode = cbcModel.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "|", slName)
        ilRet = gParseItem(slNameCode, 2, "|", slCode)
        imModelCode(ilLoop) = Val(slCode)
        cbcModel.List(ilLoop) = slName
    Next ilLoop
'    ilRet = btrClose(hlUrf)
'    btrDestroy hlUrf
    Exit Sub

    On Error GoTo 0
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    imTerminate = True
    Exit Sub
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
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    Dim ilPasswordAltered As Integer
    Dim slDate As String
    If Trim$(tmWkUrf.sPassword) <> edcPassword.Text Then
        ilPasswordAltered = True
        'ilRet = ChangeUserPassword(Trim(tmWkUrf.sName), Trim(edcPassword.Text))
        'D.S. 09-09-15 make call to API - tmUrf(imUrfIndex).sPassword  If Trim$(edcEMail.Text) <> "" Then
    Else
        ilPasswordAltered = False
    End If
    'If imNewRec Then
    '    mCreateUrf
    'End If
    tmUrf(imUrfIndex) = tmWkUrf
    If imIncludeDormant Then
        If rbcState(1).Value Then
            tmUrf(imUrfIndex).sDelete = "Y"
        Else
            tmUrf(imUrfIndex).sDelete = "N"
        End If
    End If
    tmUrf(imUrfIndex).sName = Trim$(edcName.Text)
    'dan M 10/14/11 password should be enabled for everyone..user can change password for himself.
    'If (edcPassword.Enabled) Then
        tmUrf(imUrfIndex).sPassword = Trim$(edcPassword.Text)
        If ilPasswordAltered Then
            slDate = Format$(gNow(), "m/d/yy")
            gPackDate slDate, tmUrf(imUrfIndex).iPasswordDate(0), tmUrf(imUrfIndex).iPasswordDate(1)
        End If
    'End If
    'If Trim$(edcRemoteID.Text) <> "" Then
    '    tmUrf(imUrfIndex).iRemoteUserID = Val(Trim$(edcRemoteID.Text))
    'Else
        tmUrf(imUrfIndex).iRemoteUserID = 0
    'End If
    If Trim$(edcGroupNo.Text) <> "" Then
        tmUrf(imUrfIndex).iGroupNo = Val(Trim$(edcGroupNo.Text))
    Else
        tmUrf(imUrfIndex).iGroupNo = 0
    End If
    'If ckcBlockRU.Value Then
    '    tmUrf(imUrfIndex).sBlockRU = "Y"
    'Else
        tmUrf(imUrfIndex).sBlockRU = "N"
    'End If
    tmUrf(imUrfIndex).sRept = Trim$(edcRept.Text)
    tmUrf(imUrfIndex).sPhoneNo = Trim$(edcPhone.Text)
    tmUrf(imUrfIndex).sCity = Trim$(edcCity.Text)
    tmCefSrchKey.lCode = lmEMailCefCode
    If tmCefSrchKey.lCode <> 0 Then
        tmCef.sComment = ""
        imCefRecLen = Len(tmCef)    '1009
        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmCef.lCode = 0
        End If
    Else
        tmCef.lCode = 0
    End If
    'tmCef.iStrLen = Len(Trim$(edcEMail.Text))
    tmCef.sComment = Trim$(edcEMail.Text) & Chr$(0) '& Chr$(0) 'sgTB
    imCefRecLen = Len(tmCef)    '5 + Len(Trim$(tmCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
    'If imCefRecLen - 2 > 7 Then '-2 so the control character at the end is not counted
    If Trim$(edcEMail.Text) <> "" Then
        If tmCef.lCode = 0 Then
            tmCef.lCode = 0 'Autoincrement
            ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
        Else
            ilRet = btrUpdate(hmCef, tmCef, imCefRecLen)
        End If
    Else
        If tmCef.lCode <> 0 Then
            ilRet = btrDelete(hmCef)
        End If
        tmCef.lCode = 0
    End If
    tmUrf(imUrfIndex).lEMailCefCode = tmCef.lCode
    lmEMailCefCode = tmCef.lCode
    If Trim$(edcPDF(4).Text) <> "" Then
        tmUrf(imUrfIndex).sPrtNameAltKey = Trim$(edcPDF(4).Text)
        If Trim$(edcPDF(0).Text) <> "" Then
            tmUrf(imUrfIndex).sPDFDrvChar = Trim$(edcPDF(0).Text)
            tmUrf(imUrfIndex).iPDFDnArrowCnt = Val(edcPDF(1).Text)
        Else
            tmUrf(imUrfIndex).sPDFDrvChar = ""
            tmUrf(imUrfIndex).iPDFDnArrowCnt = 0
        End If
        If Trim$(edcPDF(2).Text) <> "" Then
            tmUrf(imUrfIndex).sPrtDrvChar = Trim$(edcPDF(2).Text)
            tmUrf(imUrfIndex).iPrtDnArrowCnt = Val(edcPDF(3).Text)
        Else
            tmUrf(imUrfIndex).sPrtDrvChar = ""
            tmUrf(imUrfIndex).iPrtDnArrowCnt = 0
        End If
        If Trim$(edcPDF(5).Text) <> "" Then
            tmUrf(imUrfIndex).iPrtNoEnterKeys = Val(edcPDF(5).Text)
        Else
            tmUrf(imUrfIndex).iPrtNoEnterKeys = 1
        End If
    Else
        tmUrf(imUrfIndex).sPrtNameAltKey = ""
        tmUrf(imUrfIndex).sPDFDrvChar = ""
        tmUrf(imUrfIndex).iPDFDnArrowCnt = 0
        tmUrf(imUrfIndex).sPrtDrvChar = ""
        tmUrf(imUrfIndex).iPrtDnArrowCnt = 0
        tmUrf(imUrfIndex).iPrtNoEnterKeys = 0
    End If
    If (sgCPName <> cbcSelect.Text) Then
        'Show Alerts
        If imAlerts(0) = 0 Then 'Yes
            tmUrf(imUrfIndex).sReprintLogAlert = "Y"
        Else
            tmUrf(imUrfIndex).sReprintLogAlert = "N"
        End If
        'Show Incomplete
        If imAlerts(1) = 0 Then 'Yes
            tmUrf(imUrfIndex).sIncompAlert = "Y"
        Else
            tmUrf(imUrfIndex).sIncompAlert = "N"
        End If
        'Show Complete
        If imAlerts(2) = 0 Then 'Yes
            tmUrf(imUrfIndex).sCompAlert = "Y"
        Else
            tmUrf(imUrfIndex).sCompAlert = "N"
        End If
        'Show Req Schd
        If imAlerts(3) = 0 Then 'Yes
            tmUrf(imUrfIndex).sSchAlert = "Y"
        Else
            tmUrf(imUrfIndex).sSchAlert = "N"
        End If
        'Show Hold
        If imAlerts(4) = 0 Then 'Yes
            tmUrf(imUrfIndex).sHoldAlert = "Y"
        Else
            tmUrf(imUrfIndex).sHoldAlert = "N"
        End If
        'Show Rate Card Chg
        If imAlerts(5) = 0 Then 'Yes
            tmUrf(imUrfIndex).sRateCardAlert = "Y"
        Else
            tmUrf(imUrfIndex).sRateCardAlert = "N"
        End If
        'Show Research
        If imAlerts(6) = 0 Then 'Yes
            tmUrf(imUrfIndex).sResearchAlert = "Y"
        Else
            tmUrf(imUrfIndex).sResearchAlert = "N"
        End If
        'Show Insuff Avails
        If imAlerts(7) = 0 Then 'Yes
            tmUrf(imUrfIndex).sAvailAlert = "Y"
        Else
            tmUrf(imUrfIndex).sAvailAlert = "N"
        End If
        'Show Credit Approved
        If imAlerts(8) = 0 Then 'Yes
            tmUrf(imUrfIndex).sCrdChkAlert = "Y"
        Else
            tmUrf(imUrfIndex).sCrdChkAlert = "N"
        End If
        'Show Credit Denied
        If imAlerts(9) = 0 Then 'Yes
            tmUrf(imUrfIndex).sDeniedAlert = "Y"
        Else
            tmUrf(imUrfIndex).sDeniedAlert = "N"
        End If
        'Show Credit Limit
        If imAlerts(10) = 0 Then 'Yes
            tmUrf(imUrfIndex).sCrdLimitAlert = "Y"
        Else
            tmUrf(imUrfIndex).sCrdLimitAlert = "N"
        End If
        'Show Spot Affecting LLD
        If imAlerts(11) = 0 Then 'Yes
            tmUrf(imUrfIndex).sMoveAlert = "Y"
        Else
            tmUrf(imUrfIndex).sMoveAlert = "N"
        End If
        'Allowed to Initiate Shutdown
        If imAlerts(12) = 0 Then 'Yes
            tmUrf(imUrfIndex).sAllowedToBlock = "Y"
        Else
            tmUrf(imUrfIndex).sAllowedToBlock = "N"
        End If
        'Show Rep-Net Messages
        If imAlerts(13) = 0 Then 'Yes
            tmUrf(imUrfIndex).sShowNRMsg = "Y"
        Else
            tmUrf(imUrfIndex).sShowNRMsg = "N"
        End If
        
        '' Megaphone
         'Email Digital Contracts
        If imAlerts(14) = 0 Then 'Yes
            tmUrf(imUrfIndex).sDigitalCntrAlert = "Y"
        Else
            tmUrf(imUrfIndex).sDigitalCntrAlert = "N"
        End If
        
         'Email Digital Impressions
        If imAlerts(15) = 0 Then 'Yes
            tmUrf(imUrfIndex).sDigitalImpAlert = "Y"
        Else
            tmUrf(imUrfIndex).sDigitalImpAlert = "N"
        End If
        '''''''''''''''''''
        
    End If
    If (sgCPName = cbcSelect.Text) Or (sgSUName = cbcSelect.Text) Then
        tmUrf(imUrfIndex).iVefCode = 0
        tmUrf(imUrfIndex).iDefVeh = 0
        tmUrf(imUrfIndex).iSlfCode = 0
        tmUrf(imUrfIndex).iMnfHubCode = 0
        tmUrf(imUrfIndex).sGrid = "I"
        tmUrf(imUrfIndex).sPrice = "I"
        tmUrf(imUrfIndex).sCredit = "I"
        tmUrf(imUrfIndex).sPayRate = "I"
        tmUrf(imUrfIndex).sMerge = "I"
        tmUrf(imUrfIndex).sHideSpots = "I"
        tmUrf(imUrfIndex).sChgBilled = "I"
        tmUrf(imUrfIndex).sChgCntr = "I"
        For ilLoop = LBound(imWin) To UBound(imWin) Step 1
            tmUrf(imUrfIndex).sWin(ilLoop) = "I"
        Next ilLoop
    Else
        If (cbcVehicle.Text = "") Or (cbcVehicle.ListIndex <= 0) Then
            tmUrf(imUrfIndex).iVefCode = 0
        Else
            ilFound = False
            If cbcVehicle.ListIndex > 0 Then
                ilLoop = gBinarySearchVef(cbcVehicle.ItemData(cbcVehicle.ListIndex))
                If ilLoop <> -1 Then
                    'slCode = Trim$(slCode)
                    tmUrf(imUrfIndex).iVefCode = tgMVef(ilLoop).iCode   '(slCode)
                Else
                    tmUrf(imUrfIndex).iVefCode = 0
                End If
            Else
                tmUrf(imUrfIndex).iVefCode = 0
            End If
        End If

        If (cbcDefVeh.Text = "") Or (cbcDefVeh.ListIndex <= 0) Then
            tmUrf(imUrfIndex).iDefVeh = 0
        Else
            ilFound = False
'            For ilLoop = 0 To UBound(tgVehicle) - 1 Step 1  'Traffic!lbcVehicle.ListCount - 1 Step 1
'                slNameCode = tgVehicle(ilLoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
'                ilRet = gParseItem(slNameCode, 1, "\", slVehName)
'                On Error GoTo mMoveCtrlToRecErr
'                gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", UserOpt
'                On Error GoTo 0
'                ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                On Error GoTo mMoveCtrlToRecErr
'                gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", UserOpt
'                On Error GoTo 0
'                If cbcDefVeh.Text = Trim$(slVehName) Then
'                    ilFound = True
'                    Exit For
'                End If
'            Next ilLoop
'            If ilFound Then
'                slCode = Trim$(slCode)
'                tmUrf(imUrfIndex).iDefVeh = CInt(slCode)
'            Else
'                tmUrf(imUrfIndex).iDefVeh = 0
'            End If
            If cbcDefVeh.ListIndex > 0 Then
                ilLoop = gBinarySearchVef(cbcDefVeh.ItemData(cbcDefVeh.ListIndex))
                If ilLoop <> -1 Then
                    'slCode = Trim$(slCode)
                    tmUrf(imUrfIndex).iDefVeh = tgMVef(ilLoop).iCode   '(slCode)
                Else
                    tmUrf(imUrfIndex).iDefVeh = 0
                End If
            Else
                tmUrf(imUrfIndex).iDefVeh = 0
            End If
        End If
        If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
            If cbcHub.ListIndex > 0 Then
                slNameCode = tmHubCode(cbcHub.ListIndex - 1).sKey  'lbcLogVehCode.List(gLastFound(lbcLogVeh) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmUrf(imUrfIndex).iMnfHubCode = Val(slCode)
            Else
                tmUrf(imUrfIndex).iMnfHubCode = 0
            End If
        Else
            tmUrf(imUrfIndex).iMnfHubCode = 0
        End If
        If (cbcRptSet.Text = "") Or (imRSSelectedIndex <= 0) Then
            tmUrf(imUrfIndex).iSnfCode = 0
        Else
            tmUrf(imUrfIndex).iSnfCode = tgSnfCode(imRSSelectedIndex - 1).tSnf.iCode
        End If
        If (cbcSalesperson.Text = "") Or (imSPSelectedIndex <= 0) Then
            tmUrf(imUrfIndex).iSlfCode = 0
        Else
            slNameCode = tgSalesperson(imSPSelectedIndex - 1).sKey 'Traffic!lbcSalesperson.List(imSPSelectedIndex)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2: Salesperson)", UserOpt
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmUrf(imUrfIndex).iSlfCode = CInt(slCode)
        End If
        'Grid
        If imSelectedFields(0) = 0 Then 'Hide
            tmUrf(imUrfIndex).sGrid = "H"
        ElseIf imSelectedFields(0) = 1 Then    'View only
            tmUrf(imUrfIndex).sGrid = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sGrid = "I"
        End If
        'Price
        If imSelectedFields(1) = 0 Then 'Hide
            tmUrf(imUrfIndex).sPrice = "H"
        ElseIf imSelectedFields(1) = 1 Then    'View only
            tmUrf(imUrfIndex).sPrice = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sPrice = "I"
        End If
        If imSelectedFields(2) = 0 Then 'Hide
            tmUrf(imUrfIndex).sCredit = "H"
        ElseIf imSelectedFields(2) = 1 Then    'View only
            tmUrf(imUrfIndex).sCredit = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sCredit = "I"
        End If
        If imSelectedFields(3) = 0 Then 'Hide
            tmUrf(imUrfIndex).sPayRate = "H"
        ElseIf imSelectedFields(3) = 1 Then    'View only
            tmUrf(imUrfIndex).sPayRate = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sPayRate = "I"
        End If
        If imSelectedFields(4) = 0 Then 'Hide
            tmUrf(imUrfIndex).sMerge = "H"
        ElseIf imSelectedFields(4) = 1 Then    'View only
            tmUrf(imUrfIndex).sMerge = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sMerge = "I"
        End If
        If imSelectedFields(5) = 0 Then 'Hide
            tmUrf(imUrfIndex).sHideSpots = "H"
        ElseIf imSelectedFields(5) = 1 Then    'View only
            tmUrf(imUrfIndex).sHideSpots = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sHideSpots = "I"
        End If
        If imSelectedFields(6) = 0 Then 'Hide
            tmUrf(imUrfIndex).sChgBilled = "H"
        ElseIf imSelectedFields(6) = 1 Then    'View only
            tmUrf(imUrfIndex).sChgBilled = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sChgBilled = "I"
        End If
        If imSelectedFields(7) = 0 Then 'Hide
            tmUrf(imUrfIndex).sChgCntr = "H"
        ElseIf imSelectedFields(7) = 1 Then    'View only
            tmUrf(imUrfIndex).sChgCntr = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sChgCntr = "I"
        End If
        If imSelectedFields(8) = 0 Then 'Hide
            tmUrf(imUrfIndex).sRefResvType = "H"
        ElseIf imSelectedFields(8) = 1 Then    'View only
            tmUrf(imUrfIndex).sRefResvType = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sRefResvType = "I"
        End If
        'Change Credit Rating
        If imSelectedFields(9) = 0 Then 'Hide
            tmUrf(imUrfIndex).sChgCrRt = "H"
        ElseIf imSelectedFields(9) = 1 Then    'View only
            tmUrf(imUrfIndex).sChgCrRt = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sChgCrRt = "I"
        End If
        ''Compute Button
        'If imSelectedFields(10) = 0 Then 'Hide
        '    tmUrf(imUrfIndex).sUseComputeCMC = "H"
        'ElseIf imSelectedFields(10) = 1 Then    'View only
        '    tmUrf(imUrfIndex).sUseComputeCMC = "V"
        'Else                    'Input
        '    tmUrf(imUrfIndex).sUseComputeCMC = "I"
        'End If
        'Region Copy
        If imSelectedFields(10) = 0 Then 'Hide
            tmUrf(imUrfIndex).sRegionCopy = "H"
        ElseIf imSelectedFields(10) = 1 Then    'View only
            tmUrf(imUrfIndex).sRegionCopy = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sRegionCopy = "I"
        End If
        'Contract Price
        If imSelectedFields(11) = 0 Then 'Hide
            tmUrf(imUrfIndex).sChgPrices = "H"
        ElseIf imSelectedFields(11) = 1 Then    'View only
            tmUrf(imUrfIndex).sChgPrices = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sChgPrices = "I"
        End If
        'Flight
        If imSelectedFields(12) = 2 Then 'Input
            tmUrf(imUrfIndex).sActFlightButton = "I"
        ElseIf imSelectedFields(12) = 1 Then    'View only
            tmUrf(imUrfIndex).sActFlightButton = "V"
        Else                    'Hide
            tmUrf(imUrfIndex).sActFlightButton = "H"
        End If
        'Change Billed Contract Price
        If ((Asc(tgSpf.sUsingFeatures4) And CHGBILLEDPRICE) <> CHGBILLEDPRICE) Or (imSelectedFields(13) = 1) Or (imSelectedFields(11) <> 2) Then
            imSelectedFields(13) = 0
        End If
        If imSelectedFields(13) = 2 Then    'Input
            tmUrf(imUrfIndex).sChgLnBillPrice = "I"
        ElseIf imSelectedFields(13) = 1 Then    'View only
            tmUrf(imUrfIndex).sChgLnBillPrice = "V"
        Else                    'Hide
            tmUrf(imUrfIndex).sChgLnBillPrice = "H"
        End If
        If imSelectedFields(14) = 2 Then 'Input
            tmUrf(imUrfIndex).sAllowInvDisplay = "I"
'        ElseIf imSelectedFields(14) = 1 Then    'View only
'            tmUrf(imUrfIndex).sAllowInvDisplay = "V"
        Else                    'Hide
            tmUrf(imUrfIndex).sAllowInvDisplay = "H"
        End If
        If imSelectedFields(15) = 2 Then 'Input
            tmUrf(imUrfIndex).sChangeCSIDate = "I"
'        ElseIf imSelectedFields(15) = 1 Then    'View only
'            tmUrf(imUrfIndex).sChangeCSIDate = "V"
        Else                    'Hide
            tmUrf(imUrfIndex).sChangeCSIDate = "H"
        End If
        If imSelectedFields(16) = 1 Then    'View only
            tmUrf(imUrfIndex).sActivityLog = "V"
        Else                    'Hide
            tmUrf(imUrfIndex).sActivityLog = "H"
        End If
        If imSelectedFields(17) = 2 Then    'View only
            tmUrf(imUrfIndex).sCntrVerify = "I"
        Else                    'Hide
            tmUrf(imUrfIndex).sCntrVerify = "H"
        End If
        If imSelectedFields(18) = 2 Then    'View only
            tmUrf(imUrfIndex).sChgAcq = "I"
        Else                    'Hide
            tmUrf(imUrfIndex).sChgAcq = "V"
        End If
        'If ((Asc(tgSaf(0).sFeatures6) And ADVANCEAVAILS) = ADVANCEAVAILS) Then
        If (tgSaf(0).sAdvanceAvail = "Y") Then
            If imSelectedFields(19) = 2 Then
                tmUrf(imUrfIndex).sAdvanceAvails = "I"
            Else                    'Hide
                tmUrf(imUrfIndex).sAdvanceAvails = "H"
            End If
        Else                    'Hide
            tmUrf(imUrfIndex).sAdvanceAvails = "H"
        End If
        
        'JJB 2024-01-24 for Contract Attachments SOW
        tmUrf(imUrfIndex).sAddAttach = IIF(imSelectedFields(20) = 2, "Y", "N")
        tmUrf(imUrfIndex).sRemoveAttach = IIF(imSelectedFields(21) = 2, "Y", "N")
        '10959
        tmUrf(imUrfIndex).sContractCreation = IIF(imSelectedFields(22) = 2, "Y", "N")
        'Select Reservation
        If imTypeFields(0) = 0 Then 'Hide
            tmUrf(imUrfIndex).sResvType = "H"
        ElseIf imTypeFields(0) = 1 Then    'View only
            tmUrf(imUrfIndex).sResvType = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sResvType = "I"
        End If
        'Select Remnant
        If imTypeFields(1) = 0 Then 'Hide
            tmUrf(imUrfIndex).sRemType = "H"
        ElseIf imTypeFields(1) = 1 Then    'View only
            tmUrf(imUrfIndex).sRemType = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sRemType = "I"
        End If
        'Select DR
        If imTypeFields(2) = 0 Then 'Hide
            tmUrf(imUrfIndex).sDRType = "H"
        ElseIf imTypeFields(2) = 1 Then    'View only
            tmUrf(imUrfIndex).sDRType = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sDRType = "I"
        End If
        'Select PI
        If imTypeFields(3) = 0 Then 'Hide
            tmUrf(imUrfIndex).sPIType = "H"
        ElseIf imTypeFields(3) = 1 Then    'View only
            tmUrf(imUrfIndex).sPIType = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sPIType = "I"
        End If
        'Select PSA
        If imTypeFields(4) = 0 Then 'Hide
            tmUrf(imUrfIndex).sPSAType = "H"
        ElseIf imTypeFields(4) = 1 Then    'View only
            tmUrf(imUrfIndex).sPSAType = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sPSAType = "I"
        End If
        'Select Promo
        If imTypeFields(5) = 0 Then 'Hide
            tmUrf(imUrfIndex).sPromoType = "H"
        ElseIf imTypeFields(5) = 1 Then    'View only
            tmUrf(imUrfIndex).sPromoType = "V"
        Else                    'Input
            tmUrf(imUrfIndex).sPromoType = "I"
        End If
        If (Asc(tgSaf(0).sFeatures5) And PROGRAMMATICALLOWED) <> PROGRAMMATICALLOWED Then
            tmUrf(imUrfIndex).sPrgmmaticAlert = "H"
        Else
            If imTypeFields(6) = 2 Then 'Yes
                tmUrf(imUrfIndex).sPrgmmaticAlert = "I"
            ElseIf imTypeFields(6) = 1 Then    'View only
                tmUrf(imUrfIndex).sPrgmmaticAlert = "V"
            Else
                tmUrf(imUrfIndex).sPrgmmaticAlert = "H"
            End If
        End If
'        'Show Alerts
'        If imAlerts(0) = 0 Then 'Yes
'            tmUrf(imUrfIndex).sReprintLogAlert = "Y"
'        Else
'            tmUrf(imUrfIndex).sReprintLogAlert = "N"
'        End If
'        'Show Incomplete
'        If imAlerts(1) = 0 Then 'Yes
'            tmUrf(imUrfIndex).sIncompAlert = "Y"
'        Else
'            tmUrf(imUrfIndex).sIncompAlert = "N"
'        End If
'        'Show Complete
'        If imAlerts(2) = 0 Then 'Yes
'            tmUrf(imUrfIndex).sCompAlert = "Y"
'        Else
'            tmUrf(imUrfIndex).sCompAlert = "N"
'        End If
'        'Show Req Schd
'        If imAlerts(3) = 0 Then 'Yes
'            tmUrf(imUrfIndex).sSchAlert = "Y"
'        Else
'            tmUrf(imUrfIndex).sSchAlert = "N"
'        End If
'        'Show Hold
'        If imAlerts(4) = 0 Then 'Yes
'            tmUrf(imUrfIndex).sHoldAlert = "Y"
'        Else
'            tmUrf(imUrfIndex).sHoldAlert = "N"
'        End If
'        'Show Rate Card Chg
'        If imAlerts(5) = 0 Then 'Yes
'            tmUrf(imUrfIndex).sRateCardAlert = "Y"
'        Else
'            tmUrf(imUrfIndex).sRateCardAlert = "N"
'        End If
'        'Show Research
'        If imAlerts(6) = 0 Then 'Yes
'            tmUrf(imUrfIndex).sResearchAlert = "Y"
'        Else
'            tmUrf(imUrfIndex).sResearchAlert = "N"
'        End If
'        'Show Insuff Avails
'        If imAlerts(7) = 0 Then 'Yes
'            tmUrf(imUrfIndex).sAvailAlert = "Y"
'        Else
'            tmUrf(imUrfIndex).sAvailAlert = "N"
'        End If
'        'Show Credit Approved
'        If imAlerts(8) = 0 Then 'Yes
'            tmUrf(imUrfIndex).sCrdChkAlert = "Y"
'        Else
'            tmUrf(imUrfIndex).sCrdChkAlert = "N"
'        End If
'        'Show Credit Denied
'        If imAlerts(9) = 0 Then 'Yes
'            tmUrf(imUrfIndex).sDeniedAlert = "Y"
'        Else
'            tmUrf(imUrfIndex).sDeniedAlert = "N"
'        End If
'        'Show Credit Limit
'        If imAlerts(10) = 0 Then 'Yes
'            tmUrf(imUrfIndex).sCrdLimitAlert = "Y"
'        Else
'            tmUrf(imUrfIndex).sCrdLimitAlert = "N"
'        End If
'        'Show Spot Affecting LLD
'        If imAlerts(11) = 0 Then 'Yes
'            tmUrf(imUrfIndex).sMoveAlert = "Y"
'        Else
'            tmUrf(imUrfIndex).sMoveAlert = "N"
'        End If
        If ckcWStatus(0).Value = vbChecked Then
            tmUrf(imUrfIndex).sWorkToDead = "Y"
        Else
            tmUrf(imUrfIndex).sWorkToDead = "N"
        End If
        If ckcWStatus(1).Value = vbChecked Then
            tmUrf(imUrfIndex).sWorkToComp = "Y"
        Else
            tmUrf(imUrfIndex).sWorkToComp = "N"
        End If
        If ckcWStatus(2).Value = vbChecked Then
            tmUrf(imUrfIndex).sWorkToHold = "Y"
        Else
            tmUrf(imUrfIndex).sWorkToHold = "N"
        End If
        If ckcWStatus(3).Value = vbChecked Then
            tmUrf(imUrfIndex).sWorkToOrder = "Y"
        Else
            tmUrf(imUrfIndex).sWorkToOrder = "N"
        End If
        If ckcCStatus(0).Value = vbChecked Then
            tmUrf(imUrfIndex).sCompToIncomp = "Y"
        Else
            tmUrf(imUrfIndex).sCompToIncomp = "N"
        End If
        If ckcCStatus(1).Value = vbChecked Then
            tmUrf(imUrfIndex).sCompToDead = "Y"
        Else
            tmUrf(imUrfIndex).sCompToDead = "N"
        End If
        If ckcCStatus(2).Value = vbChecked Then
            tmUrf(imUrfIndex).sCompToHold = "Y"
        Else
            tmUrf(imUrfIndex).sCompToHold = "N"
        End If
        If ckcCStatus(3).Value = vbChecked Then
            tmUrf(imUrfIndex).sCompToOrder = "Y"
        Else
            tmUrf(imUrfIndex).sCompToOrder = "N"
        End If
        If ckcIStatus(0).Value = vbChecked Then
            tmUrf(imUrfIndex).sIncompToDead = "Y"
        Else
            tmUrf(imUrfIndex).sIncompToDead = "N"
        End If
        If ckcIStatus(1).Value = vbChecked Then
            tmUrf(imUrfIndex).sIncompToComp = "Y"
        Else
            tmUrf(imUrfIndex).sIncompToComp = "N"
        End If
        If ckcIStatus(2).Value = vbChecked Then
            tmUrf(imUrfIndex).sIncompToHold = "Y"
        Else
            tmUrf(imUrfIndex).sIncompToHold = "N"
        End If
        If ckcIStatus(3).Value = vbChecked Then
            tmUrf(imUrfIndex).sIncompToOrder = "Y"
        Else
            tmUrf(imUrfIndex).sIncompToOrder = "N"
        End If
        If ckcDStatus(0).Value = vbChecked Then
            tmUrf(imUrfIndex).sDeadToWork = "Y"
        Else
            tmUrf(imUrfIndex).sDeadToWork = "N"
        End If
        If ckcHStatus(0).Value = vbChecked Then
            tmUrf(imUrfIndex).sHoldToOrder = "Y"
        Else
            tmUrf(imUrfIndex).sHoldToOrder = "N"
        End If
        If rbcReviseCntr(1).Value Then
            tmUrf(imUrfIndex).sReviseCntr = "N"
        Else
            tmUrf(imUrfIndex).sReviseCntr = "Y"
        End If
        If rbcLiveLog(0).Value Then
            tmUrf(imUrfIndex).sLiveLogPostOnly = "Y"
        Else
            tmUrf(imUrfIndex).sLiveLogPostOnly = "N"
        End If
        If rbcSports(0).Value Then
            tmUrf(imUrfIndex).sSportPropOnly = "Y"
        Else
            tmUrf(imUrfIndex).sSportPropOnly = "N"
        End If
        For ilLoop = LBound(imWin) To UBound(imWin) Step 1
            If imWin(ilLoop) = 0 Then 'Hide
                tmUrf(imUrfIndex).sWin(ilLoop) = "H"
            ElseIf imWin(ilLoop) = 1 Then    'View only
                tmUrf(imUrfIndex).sWin(ilLoop) = "V"
            Else                    'Input
                tmUrf(imUrfIndex).sWin(ilLoop) = "I"
            End If
        Next ilLoop
    End If
    Exit Sub
mMoveCtrlToRecErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slRecCode As String
    Dim slNameCode As String  'name and code
    Dim ilRet As Integer    'Return call status
    Dim slCode As String    'code number
    Dim ilUrfIndex As Integer
    Dim ilSvIgnoreChg As Integer
    Dim ilMapIndex As Integer
    ilSvIgnoreChg = imIgnoreChg
    imIgnoreChg = YES
    ilUrfIndex = imUrfIndex
    If (ilUrfIndex < 0) Or (ilUrfIndex > UBound(tmUrf)) Then
'        mClearCtrlFields
        If (cbcSelect.ListIndex < 0) Or (cbcSelect.Text = "[New]") Then
            mInitCtrlFields
            If cbcSelect.Text <> "[New]" Then
                edcName.Text = cbcSelect.Text
            End If
            imIgnoreChg = ilSvIgnoreChg
            Exit Sub
        End If
        ilUrfIndex = 0  'Temporary
        tmWkUrf = tmUrf(ilUrfIndex)    'Move in previous values for password, name, default fac
    Else
        tmWkUrf = tmUrf(ilUrfIndex)
    End If
    imIgnoreChg = YES
    If imIncludeDormant Then
        If tmWkUrf.sDelete <> "Y" Then
            rbcState(0).Value = True
        Else
            rbcState(1).Value = True
        End If
    End If
    edcName.Text = Trim$(tmWkUrf.sName)

    'Jim 12/12/06: Show [All Vehicles]
    cbcVehicle.ListIndex = 0    '-1
    imVehSelectedIndex = 0      '-1
    slRecCode = Trim$(Str$(tmWkUrf.iVefCode))
    smVeh = ""
    If tmWkUrf.iVefCode <> 0 Then
        For ilLoop = 0 To cbcVehicle.ListCount - 1 Step 1
            If tmWkUrf.iVefCode = cbcVehicle.ItemData(ilLoop) Then
                cbcVehicle.ListIndex = ilLoop
                smVeh = cbcVehicle.List(ilLoop)
            End If
        Next ilLoop
    Else
        cbcVehicle.ListIndex = 0    '-1
        imVehSelectedIndex = 0      '-1
    End If

    cbcDefVeh.ListIndex = -1
    imDVSelectedIndex = -1
    slRecCode = Trim$(Str$(tmWkUrf.iDefVeh))
    smDefVeh = ""
    If tmWkUrf.iDefVeh <> 0 Then
'        For ilLoop = 0 To UBound(tgVehicle) - 1 Step 1  'Traffic!lbcVehicle.ListCount - 1 Step 1
'            slNameCode = tgVehicle(ilLoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
'            ilRet = gParseItem(slNameCode, 1, "\", slName)
'            On Error GoTo mMoveRecToCtrlErr
'            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 1)", UserOpt
'            On Error GoTo 0
'            ilRet = gParseItem(slNameCode, 2, "\", slCode)
'            On Error GoTo mMoveRecToCtrlErr
'            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", UserOpt
'            On Error GoTo 0
'            If slRecCode = Trim$(slCode) Then
'                gFindMatch slName, 1, cbcDefVeh
'                If gLastFound(cbcDefVeh) > 0 Then
'                    cbcDefVeh.ListIndex = gLastFound(cbcDefVeh)
'                    smDefVeh = cbcDefVeh.List(gLastFound(cbcDefVeh))
'                End If
'                Exit For
'            End If
'        Next ilLoop
        For ilLoop = 0 To cbcVehicle.ListCount - 1 Step 1
            If tmWkUrf.iDefVeh = cbcVehicle.ItemData(ilLoop) Then
                cbcDefVeh.ListIndex = ilLoop
                smDefVeh = cbcDefVeh.List(ilLoop)
            End If
        Next ilLoop
    Else
        cbcDefVeh.ListIndex = -1
        imDVSelectedIndex = -1
    End If
    edcPassword.Enabled = True
    cmcPassword.Enabled = False
    edcPassword.PasswordChar = ""
    If (Trim$(tmWkUrf.sName) = sgCPName) And (Trim$(tgUrf(0).sName) = sgCPName) Then
        edcPassword.Text = Trim$(tmUrf(ilUrfIndex).sPassword)
        smCurrentPassword = edcPassword.Text
    ElseIf (Trim$(tmWkUrf.sName) = sgSUName) And (Trim$(tgUrf(0).sName) = sgSUName) Then
        edcPassword.Text = Trim$(tmWkUrf.sPassword)
        smCurrentPassword = edcPassword.Text
'    ElseIf (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Or (Trim$(tmWkUrf.sName) = Trim$(tgUrf(0).sName)) Or imNewRec Then
    'Disallow user from seeing his password as he might leave the screen with the password showing
    ElseIf (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Or imNewRec Then
        edcPassword.Text = Trim$(tmWkUrf.sPassword)
        smCurrentPassword = edcPassword.Text
    Else    'basic user
        smCurrentPassword = Trim$(tmWkUrf.sPassword)
        ' Dan M 7/23/09 added text to password box, but protected.
        edcPassword.Text = smCurrentPassword
        'edcPassword.Text = ""
        edcPassword.PasswordChar = "*"
        'dan m 10/14/11 user can change his password
      '  edcPassword.Enabled = False
       ' cmcPassword.Enabled = True
    End If
    'edcRemoteID.Text = Trim$(Str$(tmWkUrf.iRemoteUserID))
    edcGroupNo.Text = Trim$(Str$(tmWkUrf.iGroupNo))
    'If tmWkUrf.sBlockRU = "Y" Then
    '    ckcBlockRU.Value = vbChecked
    'Else
    '    ckcBlockRU.Value = vbUnchecked
    'End If
    smHub = ""
    cbcHub.ListIndex = -1
    If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
        If tmWkUrf.iMnfHubCode > 0 Then
            For ilLoop = 0 To UBound(tmHubCode) - 1 Step 1 'lbcVehGpCode.ListCount - 1 Step 1
                slNameCode = tmHubCode(ilLoop).sKey   'lbcVehGpCode.List(ilVef)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) = tmWkUrf.iMnfHubCode Then
                    cbcHub.ListIndex = ilLoop + 1
                    smHub = cbcHub.List(ilLoop + 1)
                    Exit For
                End If
            Next ilLoop
        Else
            cbcHub.ListIndex = 0
            smHub = cbcHub.List(0)
        End If
    End If

    edcRept.Text = Trim$(tmWkUrf.sRept)
    edcPhone.Text = Trim$(tmWkUrf.sPhoneNo)
    edcCity.Text = Trim$(tmWkUrf.sCity)
    lmEMailCefCode = 0
    tmCefSrchKey.lCode = tmWkUrf.lEMailCefCode
    lmEMailCefCode = tmWkUrf.lEMailCefCode
    If tmCefSrchKey.lCode <> 0 Then
        tmCef.sComment = ""
        imCefRecLen = Len(tmCef)    '1009
        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            'If tmCef.iStrLen > 0 Then
            '    edcEMail.Text = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
            'Else
            '    edcEMail.Text = ""
            'End If
            edcEMail.Text = gStripChr0(tmCef.sComment)
        Else
            edcEMail.Text = ""
        End If
    Else
        edcEMail.Text = ""
    End If
    smEMail = edcEMail.Text
    smRptSet = ""
    If tmWkUrf.iSnfCode <> 0 Then
        For ilLoop = 0 To UBound(tgSnfCode) - 1 Step 1  'Traffic!lbcSalesperson.ListCount - 1 Step 1
            If tmWkUrf.iSnfCode = tgSnfCode(ilLoop).tSnf.iCode Then
                cbcRptSet.ListIndex = ilLoop + 1
                smRptSet = cbcRptSet.List(ilLoop + 1)
                Exit For
            End If
        Next ilLoop
    Else
        cbcRptSet.ListIndex = -1
        imRSSelectedIndex = -1
    End If
    slRecCode = Trim$(Str$(tmWkUrf.iSlfCode))
    smSalesPerson = ""
    If tmWkUrf.iSlfCode <> 0 Then
        For ilLoop = 0 To UBound(tgSalesperson) - 1 Step 1  'Traffic!lbcSalesperson.ListCount - 1 Step 1
            slNameCode = tgSalesperson(ilLoop).sKey    'Traffic!lbcSalesperson.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", UserOpt
            On Error GoTo 0
            If slRecCode = slCode Then
                cbcSalesperson.ListIndex = ilLoop + 1
                smSalesPerson = cbcSalesperson.List(ilLoop + 1)
                Exit For
            End If
        Next ilLoop
    Else
        cbcSalesperson.ListIndex = -1
        imSPSelectedIndex = -1
    End If
'    If (imUrfIndex < 0) Or (imUrfIndex > UBound(tmUrf)) Then
'        imIgnoreChg = ilSvIgnoreChg
'        Exit Sub
'    End If
    If Trim$(tmWkUrf.sPrtNameAltKey) <> "" Then
        edcPDF(4).Text = tmWkUrf.sPrtNameAltKey
    Else
        edcPDF(4).Text = ""
    End If
    If Trim$(tmWkUrf.sPDFDrvChar) <> "" Then
        edcPDF(0).Text = tmWkUrf.sPDFDrvChar
        edcPDF(1).Text = Trim$(Str$(tmWkUrf.iPDFDnArrowCnt))
    Else
        edcPDF(0).Text = ""
        edcPDF(1).Text = ""
    End If
    If Trim$(tmWkUrf.sPrtDrvChar) <> "" Then
        edcPDF(2).Text = tmWkUrf.sPrtDrvChar
        edcPDF(3).Text = Trim$(Str$(tmWkUrf.iPrtDnArrowCnt))
    Else
        edcPDF(2).Text = ""
        edcPDF(3).Text = ""
    End If
    If tmWkUrf.iPrtNoEnterKeys > 0 Then
        edcPDF(5).Text = Trim$(Str$(tmWkUrf.iPrtNoEnterKeys))
    Else
        edcPDF(5).Text = ""
    End If
    For ilLoop = LBound(imWin) To UBound(imWin) Step 1
        'Test if blank- fix if records added manually via BU
        If (tmWkUrf.sWin(ilLoop) = "") Or (tmWkUrf.sWin(ilLoop) = " ") Then
            If (ilLoop = SITELIST) Or (ilLoop = USERLIST) Then
                tmWkUrf.sWin(ilLoop) = "H"
            Else
                tmWkUrf.sWin(ilLoop) = "V"
            End If
        End If
        If tmWkUrf.sWin(ilLoop) = "H" Then 'Hide
            imWin(ilLoop) = 0
        ElseIf tmWkUrf.sWin(ilLoop) = "V" Then    'View only
            imWin(ilLoop) = 1
        Else                    'Input
            imWin(ilLoop) = 2
        End If
    Next ilLoop
    'Copy of gSetWinStatusFromSPF
    mSetWinStatusFromSPF
    For ilLoop = RATECARDSJOB To REPORTSJOB - 1 Step 1
        ilIndex = imWinMap(ilLoop)
        If ilIndex >= 0 Then
            'pbcJobs(ilIndex).Cls
            'pbcJobs(ilIndex).CurrentX = imStartX
            'pbcJobs(ilIndex).CurrentY = imStartY
            'If tmWkUrf.sWin(ilLoop) = "I" Then
            '    pbcJobs(ilIndex).CurrentX = imStartX + imIAdj
            'End If
            'pbcJobs(ilIndex).Print tmWkUrf.sWin(ilLoop)
            '9/26/15: Even if disabled show status color
            'If plcJobs(ilIndex).Enabled Then
                If tmWkUrf.sWin(ilLoop) = "H" Then
                    plcJobs(ilIndex).BackColor = Red
                ElseIf tmWkUrf.sWin(ilLoop) = "V" Then
                    plcJobs(ilIndex).BackColor = Yellow
                ElseIf tmWkUrf.sWin(ilLoop) = "I" Then
                    plcJobs(ilIndex).BackColor = GREEN
                Else
                    plcJobs(ilIndex).BackColor = GRAY
                End If
            'Else
            '    plcJobs(ilIndex).BackColor = RED
            'End If
        End If
    Next ilLoop
    'General
    For ilLoop = VEHICLESLIST To USERLIST Step 1
        ilIndex = imWinMap(ilLoop)
        If ilIndex >= 0 Then
            'pbcLists(ilIndex).Cls
            'pbcLists(ilIndex).CurrentX = imStartX
            'pbcLists(ilIndex).CurrentY = imStartY
            'If tmWkUrf.sWin(ilLoop) = "I" Then
            '    pbcLists(ilIndex).CurrentX = imStartX + imIAdj
            'End If
            'pbcLists(ilIndex).Print tmWkUrf.sWin(ilLoop)
            If tmWkUrf.sWin(ilLoop) = "H" Then
                plcLists(ilIndex).BackColor = Red    'Print "H"
            ElseIf tmWkUrf.sWin(ilLoop) = "V" Then
                plcLists(ilIndex).BackColor = Yellow
            ElseIf tmWkUrf.sWin(ilLoop) = "I" Then
                plcLists(ilIndex).BackColor = GREEN
            Else
                plcLists(ilIndex).BackColor = GRAY
            End If
        End If
    Next ilLoop
    'Grid
    If tmWkUrf.sGrid = "H" Then 'Hide
        imSelectedFields(0) = 0
    ElseIf tmWkUrf.sGrid = "V" Then    'View only
        imSelectedFields(0) = 1
    ElseIf tmWkUrf.sGrid = "I" Then                    'Input
        imSelectedFields(0) = 2
    Else
        imSelectedFields(0) = -1
    End If
    'Price
    If tmWkUrf.sPrice = "H" Then 'Hide
        imSelectedFields(1) = 0
    ElseIf tmWkUrf.sPrice = "V" Then    'View only
        imSelectedFields(1) = 1
    ElseIf tmWkUrf.sPrice = "I" Then                    'Input
        imSelectedFields(1) = 2
    Else                            'Not Defined
        imSelectedFields(1) = -1
    End If
    If tmWkUrf.sCredit = "H" Then 'Hide
        imSelectedFields(2) = 0
    ElseIf tmWkUrf.sCredit = "V" Then    'View only
        imSelectedFields(2) = 1
    ElseIf tmWkUrf.sCredit = "I" Then 'Input
        imSelectedFields(2) = 2
    Else                            'Not Defined
        imSelectedFields(2) = -1
    End If
    If tmWkUrf.sPayRate = "H" Then 'Hide
        imSelectedFields(3) = 0
    ElseIf tmWkUrf.sPayRate = "V" Then    'View only
        imSelectedFields(3) = 1
    ElseIf tmWkUrf.sPayRate = "I" Then 'Input
        imSelectedFields(3) = 2
    Else                    'Not Defined
        imSelectedFields(3) = -1
    End If
    If tmWkUrf.sMerge = "H" Then 'Hide
        imSelectedFields(4) = 0
    ElseIf tmWkUrf.sMerge = "V" Then    'View only
        imSelectedFields(4) = 1
    ElseIf tmWkUrf.sMerge = "I" Then 'Input
        imSelectedFields(4) = 2
    Else                    'Not Defined
        imSelectedFields(4) = -1
    End If
    If tmWkUrf.sHideSpots = "H" Then 'Hide
        imSelectedFields(5) = 0
    ElseIf tmWkUrf.sHideSpots = "V" Then    'View only
        imSelectedFields(5) = 1
    ElseIf tmWkUrf.sHideSpots = "I" Then 'Input
        imSelectedFields(5) = 2
    Else                    'Not Defined
        imSelectedFields(5) = -1
    End If
    If tmWkUrf.sChgBilled = "H" Then 'Hide
        imSelectedFields(6) = 0
    ElseIf tmWkUrf.sChgBilled = "V" Then    'View only
        imSelectedFields(6) = 1
    ElseIf tmWkUrf.sChgBilled = "I" Then 'Input
        imSelectedFields(6) = 2
    Else                    'Not Defined
        imSelectedFields(6) = -1
    End If
    If tmWkUrf.sChgCntr = "H" Then 'Hide
        imSelectedFields(7) = 0
    ElseIf tmWkUrf.sChgCntr = "V" Then    'View only
        imSelectedFields(7) = 1
    ElseIf tmWkUrf.sChgCntr = "I" Then 'Input
        imSelectedFields(7) = 2
    Else                    'Not Defined
        imSelectedFields(7) = -1
    End If
    If tmWkUrf.sRefResvType = "H" Then 'Hide
        imSelectedFields(8) = 0
    ElseIf tmWkUrf.sRefResvType = "V" Then    'View only
        imSelectedFields(8) = 1
    Else                    'Input
        imSelectedFields(8) = 2
    End If
    If tmWkUrf.sChgCrRt = "H" Then 'Hide
        imSelectedFields(9) = 0
    ElseIf tmWkUrf.sChgCrRt = "V" Then    'View only
        imSelectedFields(9) = 1
    ElseIf tmWkUrf.sChgCrRt = "I" Then 'Input
        imSelectedFields(9) = 2
    Else                    'Not Defined
        imSelectedFields(9) = -1
    End If
    'If tmWkUrf.sUseComputeCMC = "H" Then 'Hide
    '    imSelectedFields(10) = 0
    'ElseIf tmWkUrf.sUseComputeCMC = "V" Then    'View only
    '    imSelectedFields(10) = 1
    'ElseIf tmWkUrf.sUseComputeCMC = "I" Then 'Input
    '    imSelectedFields(10) = 2
    'Else                    'Not Defined
    '    imSelectedFields(10) = -1
    'End If
    If tmWkUrf.sRegionCopy = "H" Then 'Hide
        imSelectedFields(10) = 0
    ElseIf tmWkUrf.sRegionCopy = "V" Then    'View only
        imSelectedFields(10) = 1
    ElseIf tmWkUrf.sRegionCopy = "I" Then 'Input
        imSelectedFields(10) = 2
    Else                    'Not Defined
        imSelectedFields(10) = -1
    End If
    'Change prices
    If tmWkUrf.sChgPrices = "H" Then 'Hide
        imSelectedFields(11) = 0
    ElseIf tmWkUrf.sChgPrices = "V" Then    'View only
        imSelectedFields(11) = 1
    ElseIf tmWkUrf.sChgPrices = "I" Then 'Input
        imSelectedFields(11) = 2
    Else                    'Not Defined
        imSelectedFields(11) = -1
    End If
    'Change Flight
    If tmWkUrf.sActFlightButton = "H" Then 'Hide
        imSelectedFields(12) = 0
    ElseIf tmWkUrf.sActFlightButton = "V" Then    'View only
        imSelectedFields(12) = 1
    ElseIf tmWkUrf.sActFlightButton = "I" Then 'Input
        imSelectedFields(12) = 2
    Else                    'Not Defined
        imSelectedFields(12) = -1
    End If
    'Change billed contract prices
    If tmWkUrf.sChgLnBillPrice = "H" Then 'Hide
        imSelectedFields(13) = 0
    ElseIf tmWkUrf.sChgLnBillPrice = "V" Then    'View only
        imSelectedFields(13) = 1
    ElseIf tmWkUrf.sChgLnBillPrice = "I" Then 'Input
        imSelectedFields(13) = 2
    Else                    'Not Defined
        imSelectedFields(13) = -1
    End If
    If ((Asc(tgSpf.sUsingFeatures4) And CHGBILLEDPRICE) <> CHGBILLEDPRICE) Or (imSelectedFields(13) = 1) Or (imSelectedFields(11) <> 2) Then
        imSelectedFields(13) = 0
    End If
         ' Dan M 4/10/09  14 and 15 are I or H only
     If tmWkUrf.sAllowInvDisplay = "I" Then 'Input
         imSelectedFields(14) = 2
    ' ElseIf tlUrf.sAllowInvDisplay = "V" Then    'View only
     '    imSelectedFields(14) = 1
     Else                    'Hide
         imSelectedFields(14) = 0
     End If
     If tmWkUrf.sChangeCSIDate = "I" Then 'Input
         imSelectedFields(15) = 2
    ' ElseIf tlUrf.sChangeCSIDate = "V" Then    'View only
    '     imSelectedFields(15) = 1
     Else                    'Hide
         imSelectedFields(15) = 0
    End If
    If tmWkUrf.sActivityLog = "V" Then    'View only
        imSelectedFields(16) = 1
    Else
        imSelectedFields(16) = 0
    End If
    If tmWkUrf.sCntrVerify = "I" Then
        imSelectedFields(17) = 2
    Else
        imSelectedFields(17) = 0
    End If
    If tmWkUrf.sChgAcq = "I" Then
        imSelectedFields(18) = 2
    Else
        imSelectedFields(18) = 1
    End If
    'If ((Asc(tgSaf(0).sFeatures6) And ADVANCEAVAILS) = ADVANCEAVAILS) Then
    If (tgSaf(0).sAdvanceAvail = "Y") Then
        If tmWkUrf.sAdvanceAvails = "I" Then
            imSelectedFields(19) = 2
        Else
            imSelectedFields(19) = 0
        End If
    Else
        imSelectedFields(19) = 0
    End If
   
    'JJB 2024-01-24 for Contract Attachments SOW
    imSelectedFields(20) = IIF(tmWkUrf.sAddAttach = "Y", 2, 0)
    imSelectedFields(21) = IIF(tmWkUrf.sRemoveAttach = "Y", 2, 0)
   '10959 default is "Y"
   imSelectedFields(22) = IIF(tmWkUrf.sContractCreation = "N", 0, 2)
    'For ilLoop = LBound(imSelectedFields) To UBound(imSelectedFields) Step 1  '4 Step 1
    '    pbcSelectedFields(ilLoop).Cls
    'Next ilLoop
    For ilLoop = LBound(imSelectedFields) To UBound(imSelectedFields) Step 1
    '    pbcSelectedFields_Paint ilLoop
        ilMapIndex = ilLoop
        If imSelectedFields(ilLoop) = 0 Then
            plcSelectedFields(ilMapIndex).BackColor = Red    'Print "H"
        ElseIf imSelectedFields(ilLoop) = 1 Then
            plcSelectedFields(ilMapIndex).BackColor = Yellow 'Print "V"
        ElseIf imSelectedFields(ilLoop) = 2 Then
            plcSelectedFields(ilMapIndex).BackColor = GREEN   'Print "I"
        Else
            plcSelectedFields(ilMapIndex).BackColor = GRAY   'Print ""
        End If
    Next ilLoop
    'If vbcSelFields.Value = 0 Then
    '    vbcSelFields_Change
    'Else
    '    vbcSelFields.Value = 0
    'End If
    'Select Reservation
    If tmWkUrf.sResvType = "H" Then 'Hide
        imTypeFields(0) = 0
    ElseIf tmWkUrf.sResvType = "V" Then    'View only
        imTypeFields(0) = 1
    ElseIf tmWkUrf.sResvType = "I" Then    'Input
        imTypeFields(0) = 2
    Else                    'Not Defined
        imTypeFields(0) = -1
    End If
    'Select Remnant
    If tmWkUrf.sRemType = "H" Then 'Hide
        imTypeFields(1) = 0
    ElseIf tmWkUrf.sRemType = "V" Then    'View only
        imTypeFields(1) = 1
    ElseIf tmWkUrf.sRemType = "I" Then    'Input
        imTypeFields(1) = 2
    Else                    'Not Defined
        imTypeFields(1) = -1
    End If
    'Select DR
    If tmWkUrf.sDRType = "H" Then 'Hide
        imTypeFields(2) = 0
    ElseIf tmWkUrf.sDRType = "V" Then    'View only
        imTypeFields(2) = 1
    ElseIf tmWkUrf.sDRType = "I" Then    'Input
        imTypeFields(2) = 2
    Else                    'Not Defined
        imTypeFields(2) = -1
    End If
    'Select PI
    If tmWkUrf.sPIType = "H" Then 'Hide
        imTypeFields(3) = 0
    ElseIf tmWkUrf.sPIType = "V" Then    'View only
        imTypeFields(3) = 1
    ElseIf tmWkUrf.sPIType = "I" Then    'Input
        imTypeFields(3) = 2
    Else                    'Not Defined
        imTypeFields(3) = -1
    End If
    'Select PSA
    If tmWkUrf.sPSAType = "H" Then 'Hide
        imTypeFields(4) = 0
    ElseIf tmWkUrf.sPSAType = "V" Then    'View only
        imTypeFields(4) = 1
    ElseIf tmWkUrf.sPSAType = "I" Then    'Input
        imTypeFields(4) = 2
    Else                    'Not Defined
        imTypeFields(4) = -1
    End If
    'Select Promo
    If tmWkUrf.sPromoType = "H" Then 'Hide
        imTypeFields(5) = 0
    ElseIf tmWkUrf.sPromoType = "V" Then    'View only
        imTypeFields(5) = 1
    ElseIf tmWkUrf.sPromoType = "I" Then    'Input
        imTypeFields(5) = 2
    Else                    'Not Defined
        imTypeFields(5) = -1
    End If
    'Programmatic Buy
    If tmWkUrf.sPrgmmaticAlert = "I" Then 'Hide
        imTypeFields(6) = 2
    ElseIf tmWkUrf.sPrgmmaticAlert = "V" Then    'View only
        imTypeFields(6) = 1
    ElseIf tmWkUrf.sPrgmmaticAlert = "H" Then    'Input
        imTypeFields(6) = 0
    Else                    'Not Defined
        imTypeFields(6) = -1
    End If
    'For ilLoop = LBound(imTypeFields) To UBound(imTypeFields) Step 1  '4 Step 1
    '    pbcTypes(ilLoop).Cls
    'Next ilLoop
    For ilLoop = LBound(imTypeFields) To UBound(imTypeFields) Step 1
    '    pbcTypes_Paint ilLoop
        ilMapIndex = ilLoop
        If imTypeFields(ilLoop) = 0 Then
            plcTypes(ilMapIndex).BackColor = Red 'Print "H"
        ElseIf imTypeFields(ilLoop) = 1 Then
            plcTypes(ilMapIndex).BackColor = Yellow  'Print "V"
        ElseIf imTypeFields(ilLoop) = 2 Then
            plcTypes(ilMapIndex).BackColor = GREEN   'Print "I"
        Else
            plcTypes(ilMapIndex).BackColor = GRAY    'Print ""
        End If
    Next ilLoop
    'Show Alerts
    If tmWkUrf.sReprintLogAlert = "N" Then 'No
        imAlerts(0) = 1
    Else                    'Input
        imAlerts(0) = 0
    End If
    'Show Incomp
    If tmWkUrf.sIncompAlert = "N" Then 'No
        imAlerts(1) = 1
    Else                    'Input
        imAlerts(1) = 0
    End If
    'Show Complete
    If tmWkUrf.sCompAlert = "N" Then 'No
        imAlerts(2) = 1
    Else                    'Input
        imAlerts(2) = 0
    End If
    'Show Req Schd
    If tmWkUrf.sSchAlert = "N" Then 'No
        imAlerts(3) = 1
    Else                    'Input
        imAlerts(3) = 0
    End If
    'Show Hold
    If tmWkUrf.sHoldAlert = "N" Then 'No
        imAlerts(4) = 1
    Else                    'Input
        imAlerts(4) = 0
    End If
    'Show Rate Card Chg
    If tmWkUrf.sRateCardAlert = "N" Then 'No
        imAlerts(5) = 1
    Else                    'Input
        imAlerts(5) = 0
    End If
    'Show Research
    If tmWkUrf.sResearchAlert = "N" Then 'No
        imAlerts(6) = 1
    Else                    'Input
        imAlerts(6) = 0
    End If
    'Show Insuff. Avail
    If tmWkUrf.sAvailAlert = "N" Then 'No
        imAlerts(7) = 1
    Else                    'Input
        imAlerts(7) = 0
    End If
    'Show Credit Approved
    If tmWkUrf.sCrdChkAlert = "N" Then 'No
        imAlerts(8) = 1
    Else                    'Input
        imAlerts(8) = 0
    End If
    'Show Credit Denied
    If tmWkUrf.sDeniedAlert = "N" Then 'No
        imAlerts(9) = 1
    Else                    'Input
        imAlerts(9) = 0
    End If
    'Show Credit exceeded
    If tmWkUrf.sCrdLimitAlert = "N" Then 'No
        imAlerts(10) = 1
    Else                    'Input
        imAlerts(10) = 0
    End If
    'Show Affect prior to LLD
    If tmWkUrf.sMoveAlert = "N" Then 'No
        imAlerts(11) = 1
    Else                    'Input
        imAlerts(11) = 0
    End If
    'Allowed to initiate Shutdown
    If tmWkUrf.sAllowedToBlock = "Y" Then 'No
        imAlerts(12) = 0
    Else                    'Input
        imAlerts(12) = 1
    End If
    'Show Rep-Net Messages
    If tmWkUrf.sShowNRMsg = "Y" Then 'No
        imAlerts(13) = 0
    Else                    'Input
        imAlerts(13) = 1
    End If
    
    '' Megaphone JJB
    'Email Digital Contracts
    If tmWkUrf.sDigitalCntrAlert = "Y" Then
        imAlerts(14) = 0
    Else                    'Input
        imAlerts(14) = 1
    End If
    
    'Email Digital Impressions
    If tmWkUrf.sDigitalImpAlert = "Y" Then
        imAlerts(15) = 0
    Else                    'Input
        imAlerts(15) = 1
    End If
    ''''''''''
    
    'For ilLoop = LBound(imAlerts) To UBound(imAlerts) Step 1  '4 Step 1
    '    pbcAlerts(ilLoop).Cls
    'Next ilLoop
    For ilLoop = LBound(imAlerts) To UBound(imAlerts) Step 1
    '    pbcAlerts_Paint ilLoop
        ilMapIndex = ilLoop
        If imAlerts(ilLoop) = 0 Then
            plcAlerts(ilMapIndex).BackColor = GREEN
        ElseIf imAlerts(ilLoop) = 1 Then
            plcAlerts(ilMapIndex).BackColor = Red
        Else
            plcAlerts(ilMapIndex).BackColor = GRAY
        End If
    Next ilLoop
    If tmWkUrf.sWorkToDead = "Y" Then
        ckcWStatus(0).Value = vbChecked
    Else
        ckcWStatus(0).Value = vbUnchecked
    End If
    If tmWkUrf.sWorkToComp = "Y" Then
        ckcWStatus(1).Value = vbChecked
    Else
        ckcWStatus(1).Value = vbUnchecked
    End If
    If tmWkUrf.sWorkToHold = "Y" Then
        ckcWStatus(2).Value = vbChecked
    Else
        ckcWStatus(2).Value = vbUnchecked
    End If
    If tmWkUrf.sWorkToOrder = "Y" Then
        ckcWStatus(3).Value = vbChecked
    Else
        ckcWStatus(3).Value = vbUnchecked
    End If
    If tmWkUrf.sCompToIncomp = "Y" Then
        ckcCStatus(0).Value = vbChecked
    Else
        ckcCStatus(0).Value = vbUnchecked
    End If
    If tmWkUrf.sCompToDead = "Y" Then
        ckcCStatus(1).Value = vbChecked
    Else
        ckcCStatus(1).Value = vbUnchecked
    End If
    If tmWkUrf.sCompToHold = "Y" Then
        ckcCStatus(2).Value = vbChecked
    Else
        ckcCStatus(2).Value = vbUnchecked
    End If
    If tmWkUrf.sCompToOrder = "Y" Then
        ckcCStatus(3).Value = vbChecked
    Else
        ckcCStatus(3).Value = vbUnchecked
    End If
    If tmWkUrf.sIncompToDead = "Y" Then
        ckcIStatus(0).Value = vbChecked
    Else
        ckcIStatus(0).Value = vbUnchecked
    End If
    If tmWkUrf.sIncompToComp = "Y" Then
        ckcIStatus(1).Value = vbChecked
    Else
        ckcIStatus(1).Value = vbUnchecked
    End If
    If tmWkUrf.sIncompToHold = "Y" Then
        ckcIStatus(2).Value = vbChecked
    Else
        ckcIStatus(2).Value = vbUnchecked
    End If
    If tmWkUrf.sIncompToOrder = "Y" Then
        ckcIStatus(3).Value = vbChecked
    Else
        ckcIStatus(3).Value = vbUnchecked
    End If
    If tmWkUrf.sDeadToWork = "Y" Then
        ckcDStatus(0).Value = vbChecked
    Else
        ckcDStatus(0).Value = vbUnchecked
    End If
    If tmWkUrf.sHoldToOrder = "Y" Then
        ckcHStatus(0).Value = vbChecked
    Else
        ckcHStatus(0).Value = vbUnchecked
    End If
    If tmWkUrf.sReviseCntr = "N" Then
        rbcReviseCntr(1).Value = True
    Else
        rbcReviseCntr(0).Value = True
    End If
    If tmWkUrf.sLiveLogPostOnly = "Y" Then
        rbcLiveLog(0).Value = True
    Else
        rbcLiveLog(1).Value = True
    End If
    If tmWkUrf.sSportPropOnly = "Y" Then
        rbcSports(0).Value = True
    Else
        rbcSports(1).Value = True
    End If
    If (imUrfIndex < 0) Or (imUrfIndex > UBound(tmUrf)) Then
        'Force change so altered flag will be set in mSetChg
        For ilLoop = LBound(imWin) To UBound(imWin) Step 1
            tmWkUrf.sWin(ilLoop) = ""
        Next ilLoop
    End If
    imIgnoreChg = ilSvIgnoreChg
    Exit Sub
mMoveRecToCtrlErr:
    On Error GoTo 0
    imTerminate = True
    imIgnoreChg = ilSvIgnoreChg
    Exit Sub
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
'    Dim slCommand As String
'    Dim slStr As String
'    Dim ilRet As Integer
'    Dim slTestSystem As String
'    Dim ilTestSystem As Integer
'    slCommand = sgCommandStr    'Command$
'    'If StrComp(slCommand, "Debug", 1) = 0 Then
'    '    igStdAloneMode = True 'Switch from/to stand alone mode
'    '    sgCallAppName = ""
'    '    slStr = "Guide"
'    '    ilTestSystem = False
'    '    imShowHelpMsg = False
'    'Else
'    '    igStdAloneMode = False  'Switch from/to stand alone mode
'        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
'        If Trim$(slStr) = "" Then
'            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
'            'End
'            imTerminate = True
'            Exit Sub
'        End If
'        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
'        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
'        If StrComp(slTestSystem, "Test", 1) = 0 Then
'            ilTestSystem = True
'        Else
'            ilTestSystem = False
'        End If
'    '    imShowHelpMsg = True
'    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
'    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
'    '        imShowHelpMsg = False
'    '    End If
'        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
'    'End If
'    'gInitStdAlone UserOpt, slStr, ilTestSystem
'    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
'    igUrfCallSource = Val(slStr)
'    'If igStdAloneMode Then
'    '    igUrfCallSource = CALLNONE
'    'End If
''    If igUrfCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
''        ilRet = gParseItem(slCommand, 4, "\", slStr)
''        If ilRet = CP_MSG_NONE Then
''            sgUrfName = slStr
''        Else
''            sgUrfName = ""
''        End If
''    End If
'    sgUrfName = sgUserName
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
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:5/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate selection control     *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim ilLoop As Integer
    imPopReqd = False
    If tgUrf(0).iCode <= 2 Then
        ilRet = gPopUserListBox(UserOpt, cbcSelect, True)
    Else
        ilRet = gPopUserListBox(UserOpt, cbcSelect, False)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        imPopReqd = True
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", UserOpt
        On Error GoTo 0
        If tgUrf(0).iCode <> 1 Then
            gFindMatch sgCPName, 0, cbcSelect
            If gLastFound(cbcSelect) >= 0 Then
                cbcSelect.RemoveItem gLastFound(cbcSelect)
            End If
            If tgUrf(0).iCode <> 2 Then
                gFindMatch sgSUName, 0, cbcSelect
                If gLastFound(cbcSelect) >= 0 Then
                    cbcSelect.RemoveItem gLastFound(cbcSelect)
                End If
                'Remove all but this user
                For ilLoop = cbcSelect.ListCount - 1 To 0 Step -1
                    If StrComp(Trim$(tgUrf(0).sName), Trim$(cbcSelect.List(ilLoop)), 1) <> 0 Then
                        cbcSelect.RemoveItem ilLoop
                    End If
                Next ilLoop
            End If
        End If
        If tgUrf(0).iCode <= 2 Then
            cbcSelect.AddItem "[New]", 0  'Force as first item on list
        End If
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetChg                         *
'*                                                     *
'*             Created:5/21/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set clear and set altered flag *
'*                                                     *
'*******************************************************
Private Sub mSetChg()
    Dim ilLoop As Integer
    imAltered = False
    If imIgnoreChg = YES Then   'Bypass changing during move record to controls
        Exit Sub
    End If
'    If (imUrfIndex < 0) Or (imUrfIndex > UBound(tmUrf)) Then
'        Exit Sub
'    End If
    If imIncludeDormant Then
        If (tmWkUrf.sDelete <> "Y") Then
            If rbcState(1).Value Then
                imAltered = True
                Exit Sub
            End If
        Else
            If rbcState(0).Value Then
                imAltered = True
                Exit Sub
            End If
        End If
    End If
    If Trim$(tmWkUrf.sName) <> Trim$(edcName.Text) Then
        imAltered = True
        Exit Sub
    End If
    'dan 10/14/11 now always enabled.
   ' If (edcPassword.Enabled) Then
        If Trim$(tmWkUrf.sPassword) <> edcPassword.Text Then
            imAltered = True
            Exit Sub
        End If
    'End If
    'If (tmWkUrf.iRemoteUserID <> Val(edcRemoteID.Text)) Then
    '    imAltered = True
    '    Exit Sub
    'End If
    If (tmWkUrf.iGroupNo <> Val(edcGroupNo.Text)) Then
        imAltered = True
        Exit Sub
    End If
    'If (tmWkUrf.sBlockRU = "Y") And (ckcBlockRU.Value = vbUnchecked) Then
    '    imAltered = True
    '    Exit Sub
    'End If
    'If ((tmWkUrf.sBlockRU = "N") Or (tmWkUrf.sBlockRU = " ")) And (ckcBlockRU.Value = vbChecked) Then
    '    imAltered = True
    '    Exit Sub
    'End If
    If Trim$(tmWkUrf.sRept) <> edcRept.Text Then
        imAltered = True
        Exit Sub
    End If
    If Trim$(tmWkUrf.sPhoneNo) <> Trim$(edcPhone.Text) Then
        imAltered = True
        Exit Sub
    End If
    If Trim$(tmWkUrf.sCity) <> Trim$(edcCity.Text) Then
        imAltered = True
        Exit Sub
    End If
    If Trim$(smEMail) <> Trim$(edcEMail.Text) Then
        imAltered = True
        Exit Sub
    End If
'    If smVehicle <> cbcVehicle.Text Then
'        imAltered = True
'        Exit Sub
'    End If
    If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
        If smHub <> cbcHub.Text Then
            imAltered = True
            Exit Sub
        End If
    End If
    If smRptSet <> cbcRptSet.Text Then
        imAltered = True
        Exit Sub
    End If
    If smSalesPerson <> cbcSalesperson.Text Then
        imAltered = True
        Exit Sub
    End If
    If smVeh <> cbcVehicle.Text Then
        imAltered = True
        Exit Sub
    End If
    If smDefVeh <> cbcDefVeh.Text Then
        imAltered = True
        Exit Sub
    End If
    If (Trim$(tmWkUrf.sPrtNameAltKey) <> Trim$(edcPDF(4).Text)) Then
        imAltered = True
        Exit Sub
    End If
    If (Trim$(tmWkUrf.sPDFDrvChar) <> Trim$(edcPDF(0).Text)) Then
        imAltered = True
        Exit Sub
    End If
    If (tmWkUrf.iPDFDnArrowCnt <> Val(edcPDF(1).Text)) Then
        imAltered = True
        Exit Sub
    End If
    If (Trim$(tmWkUrf.sPrtDrvChar) <> Trim$(edcPDF(2).Text)) Then
        imAltered = True
        Exit Sub
    End If
    If (tmWkUrf.iPrtDnArrowCnt <> Val(edcPDF(3).Text)) Then
        imAltered = True
        Exit Sub
    End If
    If (tmWkUrf.iPrtNoEnterKeys <> Val(edcPDF(5).Text)) Then
        imAltered = True
        Exit Sub
    End If
    'Grid
    If imSelectedFields(0) = 0 Then 'Hide
        If tmWkUrf.sGrid <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(0) = 1 Then    'View only
        If tmWkUrf.sGrid <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(0) = 2 Then    'Input
        If tmWkUrf.sGrid <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Price
    If imSelectedFields(1) = 0 Then 'Hide
        If tmWkUrf.sPrice <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(1) = 1 Then    'View only
        If tmWkUrf.sPrice <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(1) = 2 Then                    'Input
        If tmWkUrf.sPrice <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(2) = 0 Then 'Hide
        If tmWkUrf.sCredit <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(2) = 1 Then    'View only
        If tmWkUrf.sCredit <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(2) = 2 Then                    'Input
        If tmWkUrf.sCredit <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(3) = 0 Then 'Hide
        If tmWkUrf.sPayRate <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(3) = 1 Then    'View only
        If tmWkUrf.sPayRate <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(3) = 2 Then                    'Input
        If tmWkUrf.sPayRate <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(4) = 0 Then 'Hide
        If tmWkUrf.sMerge <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(4) = 1 Then    'View only
        If tmWkUrf.sMerge <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(4) = 2 Then                    'Input
        If tmWkUrf.sMerge <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(5) = 0 Then 'Hide
        If tmWkUrf.sHideSpots <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(5) = 1 Then    'View only
        If tmWkUrf.sHideSpots <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(5) = 2 Then                    'Input
        If tmWkUrf.sHideSpots <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(6) = 0 Then 'Hide
        If tmWkUrf.sChgBilled <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(6) = 1 Then    'View only
        If tmWkUrf.sChgBilled <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(6) = 2 Then                    'Input
        If tmWkUrf.sChgBilled <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(7) = 0 Then 'Hide
        If tmWkUrf.sChgCntr <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(7) = 1 Then    'View only
        If tmWkUrf.sChgCntr <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(7) = 2 Then                    'Input
        If tmWkUrf.sChgCntr <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(8) = 0 Then 'Hide
        If tmWkUrf.sRefResvType <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(8) = 1 Then    'View only
        If tmWkUrf.sRefResvType <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(8) = 2 Then                    'Input
        If tmWkUrf.sRefResvType <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(9) = 0 Then 'Hide
        If tmWkUrf.sChgCrRt <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(9) = 1 Then    'View only
        If tmWkUrf.sChgCrRt <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(9) = 2 Then                    'Input
        If tmWkUrf.sChgCrRt <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'If imSelectedFields(10) = 0 Then 'Hide
    '    If tmWkUrf.sUseComputeCMC <> "H" Then
    '        imAltered = True
    '        Exit Sub
    '    End If
    'ElseIf imSelectedFields(10) = 1 Then    'View only
    '    If tmWkUrf.sUseComputeCMC <> "V" Then
    '        imAltered = True
    '        Exit Sub
    '    End If
    'ElseIf imSelectedFields(10) = 2 Then                    'Input
    '    If tmWkUrf.sUseComputeCMC <> "I" Then
    '        imAltered = True
    '        Exit Sub
    '    End If
    'End If
    If imSelectedFields(10) = 0 Then 'Hide
        If tmWkUrf.sRegionCopy <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(10) = 1 Then    'View only
        If tmWkUrf.sRegionCopy <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(10) = 2 Then                    'Input
        If tmWkUrf.sRegionCopy <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(11) = 0 Then 'Hide
        If tmWkUrf.sChgPrices <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(11) = 1 Then    'View only
        If tmWkUrf.sChgPrices <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(11) = 2 Then                    'Input
        If tmWkUrf.sChgPrices <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(12) = 0 Then 'Hide
        If tmWkUrf.sActFlightButton <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(12) = 1 Then    'View only
        If tmWkUrf.sActFlightButton <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(12) = 2 Then                    'Input
        If tmWkUrf.sActFlightButton <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Change Billed contract prices
    If imSelectedFields(13) = 0 Then 'Hide
        If tmWkUrf.sChgLnBillPrice <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(13) = 1 Then    'View only
        If tmWkUrf.sChgLnBillPrice <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(13) = 2 Then                    'Input
        If tmWkUrf.sChgLnBillPrice <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(14) = 0 Then 'Hide
        If tmWkUrf.sAllowInvDisplay <> "H" Then
            imAltered = True
            Exit Sub
        End If
    'ElseIf imSelectedFields(14) = 1 Then    'View only
    '    If tmWkUrf.sAllowInvDisplay <> "V" Then
    '        imAltered = True
    '        Exit Sub
    '    End If
    ElseIf imSelectedFields(14) = 2 Then                    'Input
        If tmWkUrf.sAllowInvDisplay <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(15) = 0 Then 'Hide
        If tmWkUrf.sChangeCSIDate <> "H" Then
            imAltered = True
            Exit Sub
        End If
    'ElseIf imSelectedFields(15) = 1 Then    'View only
    '    If tmWkUrf.sChangeCSIDate <> "V" Then
    '        imAltered = True
    '        Exit Sub
    '    End If
    ElseIf imSelectedFields(15) = 2 Then                    'Input
        If tmWkUrf.sChangeCSIDate <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(16) = 0 Then 'Hide
        If tmWkUrf.sActivityLog <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(16) = 1 Then    'View only
        If tmWkUrf.sActivityLog <> "V" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(17) = 0 Then 'Hide
        If tmWkUrf.sCntrVerify <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(17) = 2 Then
        If tmWkUrf.sCntrVerify <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(18) = 1 Then 'View
        If tmWkUrf.sChgAcq <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(18) = 2 Then
        If tmWkUrf.sChgAcq <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    If imSelectedFields(19) = 0 Then 'Hide
        If tmWkUrf.sAdvanceAvails <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imSelectedFields(19) = 2 Then
        If tmWkUrf.sAdvanceAvails <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Select Reservation
    If imTypeFields(0) = 0 Then 'Hide
        If tmWkUrf.sResvType <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imTypeFields(0) = 1 Then    'View only
        If tmWkUrf.sResvType <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imTypeFields(0) = 2 Then                    'Input
        If tmWkUrf.sResvType <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Select Remnant
    If imTypeFields(1) = 0 Then 'Hide
        If tmWkUrf.sRemType <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imTypeFields(1) = 1 Then    'View only
        If tmWkUrf.sRemType <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imTypeFields(1) = 2 Then                    'Input
        If tmWkUrf.sRemType <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Select DR
    If imTypeFields(2) = 0 Then 'Hide
        If tmWkUrf.sDRType <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imTypeFields(2) = 1 Then    'View only
        If tmWkUrf.sDRType <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imTypeFields(2) = 2 Then                    'Input
        If tmWkUrf.sDRType <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Select PI
    If imTypeFields(3) = 0 Then 'Hide
        If tmWkUrf.sPIType <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imTypeFields(3) = 1 Then    'View only
        If tmWkUrf.sPIType <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imTypeFields(3) = 2 Then                    'Input
        If tmWkUrf.sPIType <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Select PSA
    If imTypeFields(4) = 0 Then 'Hide
        If tmWkUrf.sPSAType <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imTypeFields(4) = 1 Then    'View only
        If tmWkUrf.sPSAType <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imTypeFields(4) = 2 Then                    'Input
        If tmWkUrf.sPSAType <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Select Promo
    If imTypeFields(5) = 0 Then 'Hide
        If tmWkUrf.sPromoType <> "H" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imTypeFields(5) = 1 Then    'View only
        If tmWkUrf.sPromoType <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imTypeFields(5) = 2 Then                    'Input
        If tmWkUrf.sPromoType <> "I" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Programmatic Buy
    If imTypeFields(6) = 2 Then 'Hide
        If tmWkUrf.sPrgmmaticAlert <> "I" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imTypeFields(6) = 1 Then    'View only
        If tmWkUrf.sPrgmmaticAlert <> "V" Then
            imAltered = True
            Exit Sub
        End If
    ElseIf imTypeFields(6) = 0 Then
        If tmWkUrf.sPrgmmaticAlert <> "H" Then
            imAltered = True
            Exit Sub
        End If
    End If
   
    'Show Alerts
    If imAlerts(0) = 0 Then 'Yes
        If tmWkUrf.sReprintLogAlert <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'No
        If tmWkUrf.sReprintLogAlert <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Show Incomplete
    If imAlerts(1) = 0 Then 'Yes
        If tmWkUrf.sIncompAlert <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'No
        If tmWkUrf.sIncompAlert <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Show Complete
    If imAlerts(2) = 0 Then 'Yes
        If tmWkUrf.sCompAlert <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'No
        If tmWkUrf.sCompAlert <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Show Scheduling Req
    If imAlerts(3) = 0 Then 'Hide
        If tmWkUrf.sSchAlert <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'Input
        If tmWkUrf.sSchAlert <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Show Hold
    If imAlerts(4) = 0 Then 'Hide
        If tmWkUrf.sHoldAlert <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'Input
        If tmWkUrf.sHoldAlert <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Show Rate Card Chg
    If imAlerts(5) = 0 Then 'Hide
        If tmWkUrf.sRateCardAlert <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'Input
        If tmWkUrf.sRateCardAlert <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Show Research
    If imAlerts(6) = 0 Then 'Yes
        If tmWkUrf.sResearchAlert <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'No
        If tmWkUrf.sResearchAlert <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Show Insuff Avail
    If imAlerts(7) = 0 Then 'Yes
        If tmWkUrf.sAvailAlert <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'No
        If tmWkUrf.sAvailAlert <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Show Credit Approved
    If imAlerts(8) = 0 Then 'Yes
        If tmWkUrf.sCrdChkAlert <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'No
        If tmWkUrf.sCrdChkAlert <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Show Credit Denied
    If imAlerts(9) = 0 Then 'Hide
        If tmWkUrf.sDeniedAlert <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'Input
        If tmWkUrf.sDeniedAlert <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Show Credit Exceeded
    If imAlerts(10) = 0 Then 'Hide
        If tmWkUrf.sCrdLimitAlert <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'Input
        If tmWkUrf.sCrdLimitAlert <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Show Affect prior to LLD
    If imAlerts(11) = 0 Then 'Hide
        If tmWkUrf.sMoveAlert <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'Input
        If tmWkUrf.sMoveAlert <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Allowed to Initiate Shutdown
    If imAlerts(12) = 0 Then 'Hide
        If tmWkUrf.sAllowedToBlock <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'Input
        If tmWkUrf.sAllowedToBlock <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If
    'Rep-Net Messages
    If imAlerts(13) = 0 Then 'Hide
        If tmWkUrf.sShowNRMsg <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'Input
        If tmWkUrf.sShowNRMsg <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If

    '' Megaphone JJB
    'Email Digital Contracts
    If imAlerts(14) = 0 Then 'Hide
        If tmWkUrf.sDigitalCntrAlert <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'Input
        If tmWkUrf.sDigitalCntrAlert <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If
    
    'Email Digital Impressions
    If imAlerts(15) = 0 Then 'Hide
        If tmWkUrf.sDigitalImpAlert <> "Y" Then
            imAltered = True
            Exit Sub
        End If
    Else                    'Input
        If tmWkUrf.sDigitalImpAlert <> "N" Then
            imAltered = True
            Exit Sub
        End If
    End If
    '''''''''''''''''''''''''''''''''''''
    If tmWkUrf.sWorkToDead = "Y" Then
        If ckcWStatus(0).Value <> vbChecked Then
            imAltered = True
            Exit Sub
        End If
    Else
        If ckcWStatus(0).Value <> vbUnchecked Then
            imAltered = True
            Exit Sub
        End If
    End If
    If tmWkUrf.sWorkToComp = "Y" Then
        If ckcWStatus(1).Value <> vbChecked Then
            imAltered = True
            Exit Sub
        End If
    Else
        If ckcWStatus(1).Value <> vbUnchecked Then
            imAltered = True
            Exit Sub
        End If
    End If
    If tmWkUrf.sWorkToHold = "Y" Then
        If ckcWStatus(2).Value <> vbChecked Then
            imAltered = True
            Exit Sub
        End If
    Else
        If ckcWStatus(2).Value <> vbUnchecked Then
            imAltered = True
            Exit Sub
        End If
    End If
    If tmWkUrf.sWorkToOrder = "Y" Then
        If ckcWStatus(3).Value <> vbChecked Then
            imAltered = True
            Exit Sub
        End If
    Else
        If ckcWStatus(3).Value <> vbUnchecked Then
            imAltered = True
            Exit Sub
        End If
    End If
    If tmWkUrf.sCompToIncomp = "Y" Then
        If ckcCStatus(0).Value <> vbChecked Then
            imAltered = True
            Exit Sub
        End If
    Else
        If ckcCStatus(0).Value <> vbUnchecked Then
            imAltered = True
            Exit Sub
        End If
    End If
    If tmWkUrf.sCompToDead = "Y" Then
        If ckcCStatus(1).Value <> vbChecked Then
            imAltered = True
            Exit Sub
        End If
    Else
        If ckcCStatus(1).Value <> vbUnchecked Then
            imAltered = True
            Exit Sub
        End If
    End If
    If tmWkUrf.sCompToHold = "Y" Then
        If ckcCStatus(2).Value <> vbChecked Then
            imAltered = True
            Exit Sub
        End If
    Else
        If ckcCStatus(2).Value <> vbUnchecked Then
            imAltered = True
            Exit Sub
        End If
    End If
    If tmWkUrf.sCompToOrder = "Y" Then
        If ckcCStatus(3).Value <> vbChecked Then
            imAltered = True
            Exit Sub
        End If
    Else
        If ckcCStatus(3).Value <> vbUnchecked Then
            imAltered = True
            Exit Sub
        End If
    End If
    If tmWkUrf.sIncompToDead = "Y" Then
        If ckcIStatus(0).Value <> vbChecked Then
            imAltered = True
            Exit Sub
        End If
    Else
        If ckcIStatus(0).Value <> vbUnchecked Then
            imAltered = True
            Exit Sub
        End If
    End If
    If tmWkUrf.sIncompToComp = "Y" Then
        If ckcIStatus(1).Value <> vbChecked Then
            imAltered = True
            Exit Sub
        End If
    Else
        If ckcIStatus(1).Value <> vbUnchecked Then
            imAltered = True
            Exit Sub
        End If
    End If
    If tmWkUrf.sIncompToHold = "Y" Then
        If ckcIStatus(2).Value <> vbChecked Then
            imAltered = True
            Exit Sub
        End If
    Else
        If ckcIStatus(2).Value <> vbUnchecked Then
            imAltered = True
            Exit Sub
        End If
    End If
    If tmWkUrf.sIncompToOrder = "Y" Then
        If ckcIStatus(3).Value <> vbChecked Then
            imAltered = True
            Exit Sub
        End If
    Else
        If ckcIStatus(3).Value <> vbUnchecked Then
            imAltered = True
            Exit Sub
        End If
    End If
    If tmWkUrf.sDeadToWork = "Y" Then
        If ckcDStatus(0).Value <> vbChecked Then
            imAltered = True
            Exit Sub
        End If
    Else
        If ckcDStatus(0).Value <> vbUnchecked Then
            imAltered = True
            Exit Sub
        End If
    End If
    If tmWkUrf.sHoldToOrder = "Y" Then
        If ckcHStatus(0).Value <> vbChecked Then
            imAltered = True
            Exit Sub
        End If
    Else
        If ckcHStatus(0).Value <> vbUnchecked Then
            imAltered = True
            Exit Sub
        End If
    End If
    If tmWkUrf.sReviseCntr = "N" Then
        If rbcReviseCntr(1).Value <> True Then
            imAltered = True
            Exit Sub
        End If
    Else
        If rbcReviseCntr(0).Value <> True Then
            imAltered = True
            Exit Sub
        End If
    End If
    For ilLoop = LBound(imWin) To UBound(imWin) Step 1
        If imWin(ilLoop) = 0 Then 'Hide
            If tmWkUrf.sWin(ilLoop) <> "H" Then
                imAltered = True
                Exit Sub
            End If
        ElseIf imWin(ilLoop) = 1 Then    'View only
            If tmWkUrf.sWin(ilLoop) <> "V" Then
                imAltered = True
                Exit Sub
            End If
        Else                    'Input
            If tmWkUrf.sWin(ilLoop) <> "I" Then
                imAltered = True
                Exit Sub
            End If
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:5/21/93       By:D. LeVine      *
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
    If imBypassSetting Then
        Exit Sub
    End If
    mSetChg
    If imAltered And imUpdateAllowed Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    'If (sgCPName = cbcSelect.Text) Or (sgSUName = cbcSelect.Text) Or imNewRec Then
    '    cmcErase.Enabled = False
    'Else
    '    cmcErase.Enabled = True
    'End If

    'Revert button set if any field changed
    If imAltered Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
    If (cbcSelect.Text = sgCPName) Or (Trim$(tgUrf(0).sName) = sgCPName) Then
        If Not plcName.Enabled Then
            plcName.Enabled = True
        End If
        cbcRptSet.Enabled = True
        cmcRptSet.Enabled = True
        'If tgSpf.sRemoteUsers = "Y" Then
        '    edcRemoteID.Enabled = True
            edcGroupNo.Enabled = True
        'Else
        '    edcRemoteID.Enabled = False
        '    edcGroupNo.Enabled = False
        'End If
    ElseIf (cbcSelect.Text = sgSUName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
        If Not plcName.Enabled Then
            plcName.Enabled = True
        End If
        cbcRptSet.Enabled = True
        cmcRptSet.Enabled = True
        'If tgSpf.sRemoteUsers = "Y" Then
        '    edcRemoteID.Enabled = True
            edcGroupNo.Enabled = True
        'Else
        '    edcRemoteID.Enabled = False
        '    edcGroupNo.Enabled = False
        'End If
    Else
        If (((imSelectedIndex >= 0) Or (cbcSelect.Text <> "")) And (imVehSelectedIndex >= 0)) And (igWinStatus(USERLIST) = 2) Then
            If Not plcName.Enabled Then
                plcName.Enabled = True
            End If
            If Not plcMain.Enabled Then
                plcMain.Enabled = True
            End If
        Else
           ' plcName.Enabled = False
           plcName.Enabled = True
           'but don't allow some things to be changed:
           cbcRptSet.Enabled = False
           cbcVehicle.Enabled = False
           cbcDefVeh.Enabled = False
           cbcRptSet.Enabled = False
           cbcSalesperson.Enabled = False
            With cmcErasePassword
                .Enabled = True
                .Caption = "Erase Password"
                .Move 1440, 960, 2685, 300
                .Visible = True
            End With

            'Dan M 7/23/09  regular user can look at restrictions, but cannot change them
          '  plcMain.Enabled = False
          'then added below:
          plcMain.Enabled = True
          frcAlerts.Enabled = False
          frcJobs.Enabled = False
          frcLists.Enabled = False
          frcStatus.Enabled = False
          frcSet.Enabled = False
          frcTypes.Enabled = False
          frcGeneral.Enabled = False
          
        End If
        cbcRptSet.Enabled = False
        cmcRptSet.Enabled = False
        'edcRemoteID.Enabled = False
        edcGroupNo.Enabled = False
    End If
    If Not imAltered Then
        'If Not plcSelect.Enabled Then
        '    plcSelect.Enabled = True
        'End If
        If Not cbcSelect.Enabled Then
            cbcSelect.Enabled = True
        End If
    Else
        'If plcSelect.Enabled Then
        '    plcSelect.Enabled = False
        'End If
        If cbcSelect.Enabled Then
            cbcSelect.Enabled = False
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSlfPop                         *
'*                                                     *
'*             Created:5/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Salesperson combo     *
'*                      control                        *
'*                                                     *
'*******************************************************
Private Sub mSlfPop()
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = imSPSelectedIndex
    If ilIndex > 0 Then
        slName = cbcSalesperson.List(ilIndex)
    End If
    'ilRet = gPopSalespersonBox(UserOpt, 0, True, True, cbcSalesperson, Traffic!lbcSalesperson, igSlfFirstNameFirst)
    ilRet = gPopSalespersonBox(UserOpt, 0, True, True, cbcSalesperson, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSlfPopErr
        gCPErrorMsg ilRet, "mSlfPop (gIMoveListBox: Salesperson)", UserOpt
        On Error GoTo 0
        cbcSalesperson.AddItem "[None]", 0
        If ilIndex > 0 Then
            gFindMatch slName, 1, cbcSalesperson
            If gLastFound(cbcSalesperson) > 0 Then
                cbcSalesperson.ListIndex = gLastFound(cbcSalesperson)
            Else
                cbcSalesperson.ListIndex = -1
            End If
        Else
            cbcSalesperson.ListIndex = ilIndex
        End If
    End If
    Exit Sub
mSlfPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSnfPop                         *
'*                                                     *
'*             Created:5/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Report Set Name combo *
'*                      control                        *
'*                                                     *
'*******************************************************
Private Sub mSnfPop()
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    ilIndex = cbcRptSet.ListIndex
    If ilIndex >= 0 Then
        slName = cbcRptSet.List(ilIndex)
    End If
    gObtainSNF hmSnf, False
    cbcRptSet.Clear
    For ilLoop = 0 To UBound(tgSnfCode) - 1 Step 1
        cbcRptSet.AddItem Trim$(tgSnfCode(ilLoop).tSnf.sName)
    Next ilLoop
    cbcRptSet.AddItem "[None]", 0
    If ilIndex >= 0 Then
        gFindMatch slName, 0, cbcRptSet
        If gLastFound(cbcRptSet) >= 0 Then
            cbcRptSet.ListIndex = gLastFound(cbcRptSet)
        Else
            cbcRptSet.ListIndex = -1
        End If
    Else
        'cbcRptSet.ListIndex = ilIndex
    End If
    Exit Sub
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
'   ANmTerminate
'   Where:
'
    Dim ilRet As Integer
    If igStdAloneMode <> True Then
        sgUrfStamp = ""
        ilRet = csiSetStamp("URF", sgUrfStamp)
        gUrfRead UserOpt, sgUrfName, False, tmUrf(), imIncludeDormant
        sgUserName = sgUrfName
    End If

    igParentRestarted = False

    sgDoneMsg = Trim$(Str$(igUrfCallSource)) & "\" & sgUrfName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload UserOpt
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:5/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
'
'   mVehPop
'   Where:
'       cbcVeh (I)- control to be populated
'
    Dim ilLoop As Integer

    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        cbcVehicle.AddItem Trim$(tgMVef(ilLoop).sName)
        cbcVehicle.ItemData(cbcVehicle.NewIndex) = tgMVef(ilLoop).iCode
        cbcDefVeh.AddItem Trim$(tgMVef(ilLoop).sName)
        cbcDefVeh.ItemData(cbcVehicle.NewIndex) = tgMVef(ilLoop).iCode
    Next ilLoop
    cbcVehicle.AddItem "[All Vehicles]", 0
    cbcDefVeh.AddItem "[None]", 0

End Sub



Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcAlerts_Click(Index As Integer)
    Dim ilIndex As Integer
    If rbcLiveLog(0).Value Then
        Exit Sub
    End If
    If (imFirstTime = YES) And (imNewRec) Then
        cmcUpdate.Enabled = True
    End If
    imFirstTime = NO
    rbcSet(3).Value = True
    ilIndex = Index '+ vbcSelFields.Value
    imAlerts(ilIndex) = imAlerts(ilIndex) + 1
    If imAlerts(ilIndex) > 1 Then
        imAlerts(ilIndex) = 0
    End If
    If imAlerts(ilIndex) = 0 Then
        plcAlerts(Index).BackColor = GREEN
    ElseIf imAlerts(ilIndex) = 1 Then
        plcAlerts(Index).BackColor = Red
    Else
        plcAlerts(Index).BackColor = GRAY
    End If
    mSetCommands
End Sub
Private Sub plcJobs_Click(Index As Integer)
    If rbcLiveLog(0).Value Then
        Exit Sub
    End If
    If rbcSports(0).Value Then
        If (Index + RATECARDSJOB <> PROPOSALSJOB) And (Index + RATECARDSJOB <> PROGRAMMINGJOB) Then
            Exit Sub
        End If
    End If
    imIgnoreChg = YES
    If (imFirstTime = YES) And (imNewRec) Then
        cmcUpdate.Enabled = True
    End If
    imFirstTime = NO
    imIgnoreChg = NO
    rbcSet(3).Value = True
    imWin(Index + RATECARDSJOB) = imWin(Index + RATECARDSJOB) + 1
    If imWin(Index + RATECARDSJOB) > 2 Then
        imWin(Index + RATECARDSJOB) = 0
    End If
    If imWin(Index + RATECARDSJOB) = 0 Then
        plcJobs(Index).BackColor = Red
    ElseIf imWin(Index + RATECARDSJOB) = 1 Then
        plcJobs(Index).BackColor = Yellow
    Else
        plcJobs(Index).BackColor = GREEN
    End If
    mSetCommands
End Sub
Private Sub plcLists_Click(Index As Integer)
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    If rbcLiveLog(0).Value Then
        Exit Sub
    End If
   imIgnoreChg = YES
    If (imFirstTime = YES) And (imNewRec) Then
        cmcUpdate.Enabled = True
    End If
    imFirstTime = NO
    imIgnoreChg = NO
    rbcSet(3).Value = True
    ilIndex = -1
    For ilLoop = VEHICLESLIST To USERLIST Step 1
        If imWinMap(ilLoop) = Index Then
            ilIndex = ilLoop
            Exit For
        End If
    Next ilLoop
    If ilIndex >= 0 Then
        imWin(ilIndex) = imWin(ilIndex) + 1
        If ilIndex = USERLIST Then
            If imWin(ilIndex) > 1 Then
                imWin(ilIndex) = 0
            End If
        Else
            If imWin(ilIndex) > 2 Then
                imWin(ilIndex) = 0
            End If
        End If
        If imWin(ilIndex) = 0 Then
            plcLists(Index).BackColor = Red    'Print "H"
        ElseIf imWin(ilIndex) = 1 Then
            plcLists(Index).BackColor = Yellow  'Print "V"
        Else
            plcLists(Index).BackColor = GREEN  'Print "I"
        End If
        mSetCommands
    End If
End Sub
Private Sub plcMain_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcName_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcSelectedFields_Click(Index As Integer)
    Dim ilIndex As Integer
    If rbcLiveLog(0).Value Then
        Exit Sub
    End If
    imIgnoreChg = YES
    If (imFirstTime = YES) And (imNewRec) Then
        cmcUpdate.Enabled = True
    End If
    imFirstTime = NO
    imIgnoreChg = NO
    rbcSet(3).Value = True
    ilIndex = Index '+ vbcSelFields.Value
    imSelectedFields(ilIndex) = imSelectedFields(ilIndex) + 1
    'If merge, bypass view  Dan M 4/10/09 added changeCSIDisplay and AllowInvoiceDisplay
    If ((ilIndex = 4) Or (ilIndex = 5) Or (ilIndex = 6) Or (ilIndex = 7) Or (ilIndex = 8) Or (ilIndex = 12) Or (ilIndex = 14) Or (ilIndex = 15) Or (ilIndex = 17) Or (ilIndex = 19) Or (ilIndex = 20) Or (ilIndex = 21) Or (ilIndex = 22)) And (imSelectedFields(ilIndex) = 1) Then
        imSelectedFields(ilIndex) = 2
    End If
    'Bypass Input
    If ((ilIndex = 16)) And (imSelectedFields(ilIndex) = 2) Then
        imSelectedFields(ilIndex) = 0
    End If
    If imSelectedFields(ilIndex) > 2 Then
        imSelectedFields(ilIndex) = 0
    End If
    If ((ilIndex = 11) Or (ilIndex = 18)) And (imSelectedFields(ilIndex) = 0) Then
        imSelectedFields(ilIndex) = 1
    End If
    If (ilIndex = 11) And (imSelectedFields(ilIndex) <> 2) Then
        imSelectedFields(13) = 0
        plcSelectedFields(13).BackColor = Red    'Print "H"
    End If
    If (ilIndex = 13) Then
        If ((Asc(tgSpf.sUsingFeatures4) And CHGBILLEDPRICE) <> CHGBILLEDPRICE) Or (imSelectedFields(11) <> 2) Then
            imSelectedFields(ilIndex) = 0
        Else
            If imSelectedFields(ilIndex) = 1 Then
                imSelectedFields(ilIndex) = 2
            End If
        End If
    End If
    'If (Index = 19) And ((Asc(tgSaf(0).sFeatures6) And ADVANCEAVAILS) <> ADVANCEAVAILS) Then
    If (Index = 19) And (tgSaf(0).sAdvanceAvail <> "Y") Then
        imSelectedFields(ilIndex) = 0
    End If
    If imSelectedFields(ilIndex) = 0 Then
        plcSelectedFields(Index).BackColor = Red    'Print "H"
    ElseIf imSelectedFields(ilIndex) = 1 Then
        plcSelectedFields(Index).BackColor = Yellow 'Print "V"
    ElseIf imSelectedFields(ilIndex) = 2 Then
        plcSelectedFields(Index).BackColor = GREEN   'Print "I"
    Else
        plcSelectedFields(Index).BackColor = GRAY   'Print ""
    End If
    mSetCommands
End Sub

Private Sub plcTypes_Click(Index As Integer)
    Dim ilIndex As Integer
    If rbcLiveLog(0).Value Then
        Exit Sub
    End If
    imIgnoreChg = YES
    If (imFirstTime = YES) And (imNewRec) Then
        cmcUpdate.Enabled = True
    End If
    imFirstTime = NO
    imIgnoreChg = NO
    rbcSet(3).Value = True
    ilIndex = Index '+ vbcSelFields.Value
    imTypeFields(ilIndex) = imTypeFields(ilIndex) + 1
    If imTypeFields(ilIndex) > 2 Then
        imTypeFields(ilIndex) = 0
    End If
    If imTypeFields(ilIndex) = 0 Then
        plcTypes(Index).BackColor = Red 'Print "H"
    ElseIf imTypeFields(ilIndex) = 1 Then
        plcTypes(Index).BackColor = Yellow  'Print "V"
    ElseIf imTypeFields(ilIndex) = 2 Then
        plcTypes(Index).BackColor = GREEN   'Print "I"
    Else
        plcTypes(Index).BackColor = GRAY    'Print ""
    End If
    mSetCommands
End Sub

Private Sub rbcLiveLog_Click(Index As Integer)
    Dim ilSetToValue As Integer
    Dim llColor As Long
    If rbcLiveLog(0).Value Then
        rbcSports(1).Value = True
        ilSetToValue = 0
        llColor = Red
        mSet 1, ilSetToValue, llColor
    End If
    mSetCommands
End Sub

Private Sub rbcReviseCntr_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcReviseCntr(Index).Value
    'End of Coded added
    mSetCommands
End Sub
Private Sub rbcSelect_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelect(Index).Value
    'End of Coded added
    If Value Then
        frcJobs.Visible = False
        frcLists.Visible = False
        frcGeneral.Visible = False
        frcAlerts.Visible = False
        frcStatus.Visible = False
        frcTypes.Visible = False
        frcPDF.Visible = False
        Select Case Index
            Case 0
                frcJobs.Visible = True
                frcSet.Caption = "Set Jobs To"
            Case 1
                frcLists.Visible = True
                frcSet.Caption = "Set Lists To"
            Case 2
                frcGeneral.Visible = True
                frcSet.Caption = "Set Fields To"
            Case 3
                frcAlerts.Visible = True
                frcSet.Caption = "Set Alerts To"
            Case 4
                frcStatus.Visible = True
            Case 5
                frcTypes.Visible = True
                frcSet.Caption = "Set Types To"
            Case 6
                frcPDF.Visible = True
        End Select
        If (Index = 4) Or (Index = 6) Then
            frcSet.Visible = False
        Else
            If imUpdateAllowed Then
                frcSet.Visible = True
            Else
                frcSet.Visible = False
            End If
            If Index = 3 Then
                rbcSet(1).Enabled = False
            Else
                rbcSet(1).Enabled = True
            End If
        End If
        rbcSet(3).Value = True
    End If
End Sub

Private Sub rbcSet_Click(Index As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        ilIndex                       ilSetIndex                *
'*                                                                                        *
'******************************************************************************************

    Dim ilSetToValue As Integer
    Dim llColor As Long

    If Not imUpdateAllowed Then
        rbcSet(3).Value = True
        Exit Sub
    End If
    If rbcSet(0).Value And rbcSports(0).Value Then
        rbcSet(3).Value = True
        Exit Sub
    End If
    If (rbcLiveLog(0).Value) Then
        rbcSet(3).Value = True
        Exit Sub
    End If
    If (rbcSet(Index).Value) And (Index <> 3) Then
        If Index = 0 Then
            ilSetToValue = 2
            llColor = GREEN
        ElseIf Index = 1 Then
            ilSetToValue = 1
            llColor = Yellow
        ElseIf Index = 2 Then
            ilSetToValue = 0
            llColor = Red
        End If
        mSet 0, ilSetToValue, llColor
        mSetCommands
    End If
End Sub

Private Sub rbcSports_Click(Index As Integer)
    Dim ilSetToValue As Integer
    Dim llColor As Long
    If rbcSports(0).Value Then
        rbcLiveLog(1).Value = True
        ilSetToValue = 0
        llColor = Red
        mSet 2, ilSetToValue, llColor
    End If
    mSetCommands
End Sub

Private Sub rbcState_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcState(Index).Value
    'End of Coded added
    If imIgnoreChg = YES Then
        Exit Sub
    End If
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub vbcSelFields_Change()
    'For ilLoop = vbcSelFields.Value To vbcSelFields.Value + 4 Step 1
    '    plcSelectedFields(ilLoop - vbcSelFields.Value).Caption = smSelFields(ilLoop)
    '    pbcSelectedFields(ilLoop - vbcSelFields.Value).CurrentY = imStartY
    '    pbcSelectedFields(ilLoop - vbcSelFields.Value).CurrentX = imStartX
    '    pbcSelectedFields(ilLoop - vbcSelFields.Value).Cls
    '    pbcSelectedFields(ilLoop - vbcSelFields.Value).CurrentY = imStartY
    '    pbcSelectedFields(ilLoop - vbcSelFields.Value).CurrentX = imStartX
    '    If imSelectedFields(ilLoop) = 2 Then
    '        slStr = "I"
    '        pbcSelectedFields(ilLoop - vbcSelFields.Value).CurrentX = imStartX + imIAdj
    '    ElseIf imSelectedFields(ilLoop) = 1 Then
    '        slStr = "V"
    '    ElseIf imSelectedFields(ilLoop) = 0 Then
    '        slStr = "H"
    '    Else
    '        slStr = " "
    '    End If
    '    pbcSelectedFields(ilLoop - vbcSelFields.Value).Print slStr
    'Next ilLoop
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "User Options"
End Sub
Private Sub plcMain_Paint()
    plcMain.CurrentX = 0
    plcMain.CurrentY = 0
    plcMain.Print "Restrictions"
End Sub
Private Sub plcReviseCntr_Paint()
    plcReviseCntr.CurrentX = 0
    plcReviseCntr.CurrentY = 0
    plcReviseCntr.Print "Allowed to Revise Existing Holds or Orders"
End Sub
Private Sub plcStatus_Paint(Index As Integer)
    plcStatus(Index).CurrentX = 0
    plcStatus(Index).CurrentY = 0
    Select Case Index
        Case 0
            plcStatus(Index).Print "Working to"
        Case 1
            plcStatus(Index).Print "Completed to"
        Case 2
            plcStatus(Index).Print "Unapproved to"
        Case 3
            plcStatus(Index).Print "Rejected to"
        Case 4
            plcStatus(Index).Print "Hold to"
    End Select
End Sub
Private Sub plcName_Paint()
    plcName.CurrentX = 0
    plcName.CurrentY = 0
    plcName.Print "Name Information"
End Sub



Private Sub mSet(ilFrom As Integer, ilInSetToValue As Integer, llInColor As Long)
    'ilFrom (I)- 0=from rbcSet; 1=From rbcLiveLog; 2=From rbcSports
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilSetIndex As Integer
    Dim ilSetToValue As Integer
    Dim llColor As Long

    ilSetToValue = ilInSetToValue
    llColor = llInColor
    If rbcSelect(0).Value Or (ilFrom <> 0) Then   'Job
        For ilLoop = plcJobs.LBound To plcJobs.UBound Step 1
            If ((ilFrom = 2) And (ilLoop + RATECARDSJOB <> PROPOSALSJOB) And (ilLoop + RATECARDSJOB <> PROGRAMMINGJOB)) Or (ilFrom <> 2) Then
                If plcJobs(ilLoop).Enabled Then
                    imWin(ilLoop + RATECARDSJOB) = ilSetToValue
                    plcJobs(ilLoop).BackColor = llColor
                End If
            End If
        Next ilLoop
        mSetWinStatusFromSPF
    End If
    If rbcSelect(1).Value Or (ilFrom = 1) Then  'List
        For ilLoop = plcLists.LBound To plcLists.UBound Step 1
            ilSetIndex = -1
            For ilIndex = VEHICLESLIST To USERLIST Step 1
                If imWinMap(ilIndex) = ilLoop Then
                    ilSetIndex = ilIndex
                    Exit For
                End If
            Next ilIndex
            If ilSetIndex <> -1 Then
                If (ilSetIndex = USERLIST) Then
                    If ((Trim$(tmWkUrf.sName) = sgCPName) And (Trim$(tgUrf(0).sName) = sgCPName)) Or ((Trim$(tmWkUrf.sName) = sgSUName) And (Trim$(tgUrf(0).sName) = sgSUName)) Then
                        imWin(ilSetIndex) = 2
                        plcLists(ilLoop).BackColor = GREEN
                    ElseIf ilSetToValue = 2 Then
                        imWin(ilSetIndex) = 1
                        plcLists(ilLoop).BackColor = Yellow
                    Else
                        imWin(ilSetIndex) = ilSetToValue
                        plcLists(ilLoop).BackColor = llColor
                    End If
                Else
                    imWin(ilSetIndex) = ilSetToValue
                    plcLists(ilLoop).BackColor = llColor
                End If
            End If
        Next ilLoop
    End If
    If rbcSelect(2).Value Or (ilFrom = 1) Then  'Select Fields
        For ilLoop = LBound(imSelectedFields) To UBound(imSelectedFields) Step 1
            If ((ilLoop = 4) Or (ilLoop = 5) Or (ilLoop = 6) Or (ilLoop = 7) Or (ilLoop = 8) Or (ilLoop = 12) Or (ilLoop = 14) Or (ilLoop = 15) Or (ilLoop = 17) Or (ilLoop = 19)) And (ilSetToValue = 1) Then
                imSelectedFields(ilLoop) = 0
                plcSelectedFields(ilLoop).BackColor = Red
            ElseIf (ilLoop = 16) And (ilSetToValue = 2) Then
                imSelectedFields(ilLoop) = 0
                plcSelectedFields(ilLoop).BackColor = Red
            ElseIf (ilLoop = 11) And (ilSetToValue = 0) Then
                imSelectedFields(ilLoop) = 1
                plcSelectedFields(ilLoop).BackColor = Yellow
            ElseIf ilLoop = 13 Then
                If ((Asc(tgSpf.sUsingFeatures4) And CHGBILLEDPRICE) <> CHGBILLEDPRICE) Or (imSelectedFields(11) <> 2) Then
                    imSelectedFields(ilLoop) = 0
                    plcSelectedFields(ilLoop).BackColor = Red
                Else
                    If ilSetToValue = 1 Then
                        imSelectedFields(ilLoop) = 0
                        plcSelectedFields(ilLoop).BackColor = Red
                    Else
                        imSelectedFields(ilLoop) = ilSetToValue
                        plcSelectedFields(ilLoop).BackColor = llColor
                    End If
                End If
            Else
                imSelectedFields(ilLoop) = ilSetToValue
                plcSelectedFields(ilLoop).BackColor = llColor
            End If
        Next ilLoop
    End If
    If rbcSelect(3).Value Or (ilFrom = 1) Then  'Alerts
        If ilSetToValue = 1 Then
            ilSetToValue = 1
            llColor = Red
        ElseIf ilSetToValue = 0 Then
            ilSetToValue = 1
            llColor = Red
        Else
            ilSetToValue = 0
            llColor = GREEN
        End If
        For ilLoop = LBound(imAlerts) To UBound(imAlerts) Step 1
            imAlerts(ilLoop) = ilSetToValue
            plcAlerts(ilLoop).BackColor = llColor
        Next ilLoop
    End If
    If rbcSelect(4).Value Or (ilFrom = 1) Then  'Contract Status
    End If
    If rbcSelect(5).Value Or (ilFrom = 1) Then  'Contract Types
        For ilLoop = LBound(imTypeFields) To UBound(imTypeFields) Step 1
            imTypeFields(ilLoop) = ilSetToValue
            plcTypes(ilLoop).BackColor = llColor
        Next ilLoop
    End If
End Sub

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
Private Sub mHubPop()
'
'   mVehGpPop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = cbcHub.ListIndex
    If ilIndex >= 0 Then
        slName = cbcHub.List(ilIndex)
    End If
    ilfilter(0) = CHARFILTER
    slFilter(0) = "W"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    'ilRet = gIMoveListBox(Vehicle, lbcDemo, lbcDemoCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(UserOpt, cbcHub, tmHubCode(), smHubCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mHubPopErr
        gCPErrorMsg ilRet, "mHubPop (gIMoveListBox)", UserOpt
        On Error GoTo 0
        cbcHub.AddItem "[All]", 0  'Force as first item on list
        If ilIndex >= 0 Then
            gFindMatch slName, 0, cbcHub
            If gLastFound(cbcHub) >= 0 Then
                cbcHub.ListIndex = gLastFound(cbcHub)
            Else
                cbcHub.ListIndex = -1
            End If
        Else
            cbcHub.ListIndex = ilIndex
        End If
    End If
    Exit Sub
mHubPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mHubBranch                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      Hub and process               *
'*                      communication back from        *
'*                      Hub                           *
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
Private Function mHubBranch() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         ilParse                                                 *
'******************************************************************************************

'
'   ilRet = mHubBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim slStr As String
    Dim slHub As String
    Dim ilUpdateAllowed As Integer

    slHub = cbcHub.Text
    sgMnfCallType = "W"
    igMNmCallSource = USERLIST
    sgMNmName = ""
    ilUpdateAllowed = imUpdateAllowed

    If igTestSystem Then
        slStr = "UserOpt^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    Else
        slStr = "UserOpt^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    End If
    sgCommandStr = slStr
    MultiNm.Show vbModal
    mHubBranch = True
    imUpdateAllowed = ilUpdateAllowed
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        cbcHub.Clear
        smHubCodeTag = ""
        mHubPop
        If imTerminate Then
            mHubBranch = False
            Exit Function
        End If
        gFindMatch slHub, 0, cbcHub
        sgMNmName = ""
        If gLastFound(cbcHub) >= 0 Then
            imChgMode = True
            cbcHub.ListIndex = gLastFound(cbcHub)
            imChgMode = False
            mHubBranch = False
        Else
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
        Exit Function
    End If
    Exit Function
End Function
Private Sub mLimitGuide()
    If (cbcSelect.Text = sgSUName) Then  'guide looking at self
            cmcErasePassword.Visible = False    'Show password
            cbcRptSet.Enabled = False           'disable report set
            rbcState(1).Enabled = False         'can't make dormant
            plcMain.Visible = False             'can't change main settings
    Else
        If sgUserName = sgSUName Then
                 plcMain.Visible = True
             cbcRptSet.Enabled = True
             rbcState(1).Enabled = True
             If cbcSelect.Text = "[New]" Then
                 cmcErasePassword.Enabled = False
                 cmcErasePassword.Caption = ""
             End If
            If Not bgInternalGuide Then
                With cmcErasePassword
                    .Enabled = True
                    .Caption = "Erase Password"
                    .Move 1440, 960, 2685, 300
                    .Visible = True
                End With
            End If ' not Internal Guide?
       End If   'Guide, Internal or not?
    End If  'Guide looking at Guide?
End Sub
Private Sub mLoadNewSelected()
'Load plcSelectedFields(14)
'Load plcSelectedFields(15)
'plcSelectedFields(14).Caption = "Allow Display of Final Invoices"
'plcSelectedFields(15).Caption = "Allow Today's Date Change"
'plcSelectedFields(14).Move plcSelectedFields(12).Left, plcSelectedFields(12).Top + plcSelectedFields(12).Height + 20
'plcSelectedFields(15).Move plcSelectedFields(13).Left, plcSelectedFields(14).Top
'plcSelectedFields(14).Visible = True
'plcSelectedFields(15).Visible = True

End Sub

Private Sub mSetWinStatusFromSPF()
    If (tgSpf.sGUsePropSys = "Y") And (tgSpf.sUsingTraffic = "Y") Then
        Exit Sub
    End If
    If (tgSpf.sGUsePropSys <> "Y") And (tgSpf.sUsingTraffic = "Y") Then
        mDisableJobCtrl BUDGETSJOB
        mDisableJobCtrl PROPOSALSJOB
        Exit Sub
    End If
    If (tgSpf.sGUsePropSys = "Y") And (tgSpf.sUsingTraffic <> "Y") And (tgSpf.sGUseAffSys <> "Y") Then
        mDisableJobCtrl CONTRACTSJOB
        mDisableJobCtrl PROGRAMMINGJOB
        mDisableJobCtrl SPOTSJOB
        mDisableJobCtrl COPYJOB
        mDisableJobCtrl LOGSJOB
'        mDisableJobCtrl STATIONFEEDJOB
        mDisableJobCtrl POSTLOGSJOB
        mDisableJobCtrl INVOICESJOB
        mDisableJobCtrl COLLECTIONSJOB
        mDisableJobCtrl SLSPCOMMSJOB
        Exit Sub
    End If
    If (tgSpf.sGUsePropSys = "Y") And (tgSpf.sUsingTraffic <> "Y") And (tgSpf.sGUseAffSys = "Y") Then
        mDisableJobCtrl CONTRACTSJOB
        mDisableJobCtrl SPOTSJOB
        mDisableJobCtrl COPYJOB
        mDisableJobCtrl LOGSJOB
'        mDisableJobCtrl STATIONFEEDJOB
        mDisableJobCtrl POSTLOGSJOB
        mDisableJobCtrl INVOICESJOB
        mDisableJobCtrl COLLECTIONSJOB
        mDisableJobCtrl SLSPCOMMSJOB
        Exit Sub
    End If
    If (tgSpf.sGUsePropSys <> "Y") And (tgSpf.sUsingTraffic <> "Y") And (tgSpf.sGUseAffSys = "Y") Then
        mDisableJobCtrl BUDGETSJOB
        mDisableJobCtrl RATECARDSJOB
        mDisableJobCtrl PROPOSALSJOB
        mDisableJobCtrl CONTRACTSJOB
        mDisableJobCtrl SPOTSJOB
        mDisableJobCtrl COPYJOB
        mDisableJobCtrl LOGSJOB
'        mDisableJobCtrl STATIONFEEDJOB
        mDisableJobCtrl POSTLOGSJOB
        mDisableJobCtrl INVOICESJOB
        mDisableJobCtrl COLLECTIONSJOB
        mDisableJobCtrl SLSPCOMMSJOB
        Exit Sub
    End If
    
End Sub


Private Sub mDisableJobCtrl(ilIndex As Integer)
    plcJobs(imWinMap(ilIndex)).BackColor = vbRed
    plcJobs(imWinMap(ilIndex)).Enabled = False
End Sub

Private Function mProgrammaticUserOk() As Boolean
    Dim slSQLQuery As String
    Dim rst_Temp As ADODB.Recordset
        
    mProgrammaticUserOk = True
    If smPrgmmaticAllow = "N" Then
        Exit Function
    End If
    If imTypeFields(6) <> 2 Then '0=Hidden;1=view
        Exit Function
    End If
    slSQLQuery = "Select urfCode From URF_User_Options Where urfPrgmmaticAlert = 'I'"
    Set rst_Temp = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Temp.EOF
        If Not imNewRec Then
            If (tmUrf(imUrfIndex).iCode <> rst_Temp!urfCode) Then
                MsgBox "Only one user can be setup as 'Allow to select Programmatic Buys' with color Green"
                mProgrammaticUserOk = False
                rst_Temp.Close
                Exit Function
            End If
        Else
            MsgBox "Only one user can be setup as 'Allow to select Programmatic Buys' with color Green"
            mProgrammaticUserOk = False
            rst_Temp.Close
            Exit Function
        End If
        rst_Temp.MoveNext
    Loop
    '8/29/18: Removing this requirement as a Programmatic Salesperson will be added
    'If (cbcSalesperson.Text = "") Or (imSPSelectedIndex <= 0) Then
    '    MsgBox "Programmatic Buy user defined as 'Allow to select Programmatic Buys' with color Green must also be defined as Salesperson"
    '    mProgrammaticUserOk = False
    '    rst_Temp.Close
    '    Exit Function
    'End If
    rst_Temp.Close
End Function

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
