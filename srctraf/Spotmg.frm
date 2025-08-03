VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form SpotMG 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5835
   ClientLeft      =   1140
   ClientTop       =   2190
   ClientWidth     =   9480
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5835
   ScaleWidth      =   9480
   Begin VB.TextBox edcMForN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
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
      Height          =   165
      Left            =   1155
      TabIndex        =   57
      Text            =   "M for N MGs"
      Top             =   150
      Width           =   1080
   End
   Begin VB.PictureBox plc1MG1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5220
      Left            =   270
      ScaleHeight     =   5190
      ScaleWidth      =   9015
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   465
      Width           =   9045
      Begin VB.CheckBox ckcSpotType 
         Caption         =   "N/C, .00, Bonus"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   1305
         TabIndex        =   19
         Top             =   45
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CheckBox ckcSpotType 
         Caption         =   "$"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   720
         TabIndex        =   18
         Top             =   45
         Value           =   1  'Checked
         Width           =   480
      End
      Begin VB.CheckBox ckcAirWeek 
         Caption         =   "Line Veh/Wks Only"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6960
         TabIndex        =   64
         Top             =   4695
         Width           =   1980
      End
      Begin VB.CheckBox ckcCntrVehOnly 
         Caption         =   "Contract Vehicles Only"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4575
         TabIndex        =   11
         Top             =   4695
         Value           =   1  'Checked
         Width           =   2265
      End
      Begin VB.CheckBox ckcDaysTimes 
         Caption         =   "Use Line Days/Times"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4995
         TabIndex        =   14
         Top             =   4935
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.PictureBox plcAired 
         ForeColor       =   &H00000000&
         Height          =   3930
         Left            =   4635
         ScaleHeight     =   3870
         ScaleWidth      =   4305
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   315
         Width           =   4365
         Begin VB.PictureBox pbcLbcAired 
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
            Height          =   3540
            Left            =   15
            ScaleHeight     =   3540
            ScaleWidth      =   4275
            TabIndex        =   59
            Top             =   270
            Width           =   4275
         End
         Begin VB.HScrollBar hbcAiredWk 
            Height          =   240
            Left            =   930
            Min             =   1
            TabIndex        =   33
            Top             =   15
            Value           =   1
            Width           =   3360
         End
         Begin VB.ListBox lbcAired 
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
            Height          =   3600
            Left            =   0
            MultiSelect     =   2  'Extended
            TabIndex        =   32
            Top             =   255
            Width           =   4305
         End
         Begin VB.Label plcAiredWk 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            TabIndex        =   34
            Top             =   15
            Width           =   930
         End
      End
      Begin VB.CommandButton cmc1MG1Done 
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
         Left            =   8190
         TabIndex        =   17
         Top             =   4335
         Width           =   735
      End
      Begin VB.PictureBox plcMissed 
         ForeColor       =   &H00000000&
         Height          =   3930
         Left            =   75
         ScaleHeight     =   3870
         ScaleWidth      =   4305
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   315
         Width           =   4365
         Begin VB.PictureBox pbcLbcMissed 
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
            Height          =   3570
            Left            =   15
            ScaleHeight     =   3570
            ScaleWidth      =   4260
            TabIndex        =   58
            Top             =   270
            Width           =   4260
         End
         Begin VB.ListBox lbcMissed 
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
            Height          =   3600
            Left            =   0
            MultiSelect     =   2  'Extended
            TabIndex        =   29
            Top             =   255
            Width           =   4305
         End
         Begin VB.HScrollBar hbcMissedWk 
            Height          =   240
            Left            =   930
            Min             =   1
            TabIndex        =   28
            Top             =   15
            Value           =   1
            Width           =   3360
         End
         Begin VB.Label plcMissedWk 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            TabIndex        =   30
            Top             =   15
            Width           =   930
         End
      End
      Begin VB.PictureBox plcMove 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3495
         ScaleHeight     =   195
         ScaleWidth      =   3180
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   60
         Width           =   3180
         Begin VB.OptionButton rbcMove 
            Caption         =   "All"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1770
            TabIndex        =   26
            Top             =   0
            Value           =   -1  'True
            Width           =   540
         End
         Begin VB.OptionButton rbcMove 
            Caption         =   "One Spot"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   0
            Width           =   1170
         End
         Begin VB.OptionButton rbcMove 
            Caption         =   "Ask"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   2355
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   0
            Width           =   690
         End
      End
      Begin VB.PictureBox plcSort 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   90
         ScaleHeight     =   195
         ScaleWidth      =   2970
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   4230
         Width           =   2970
         Begin VB.OptionButton rbcSort 
            Caption         =   "Vehicle"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1920
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton rbcSort 
            Caption         =   "Advertiser"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   690
            TabIndex        =   21
            Top             =   0
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmc1MG1Outside 
         Appearance      =   0  'Flat
         Caption         =   "&Outside"
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
         Left            =   7245
         TabIndex        =   16
         Top             =   4335
         Width           =   840
      End
      Begin VB.CommandButton cmc1MG1MG 
         Appearance      =   0  'Flat
         Caption         =   "&MGs"
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
         Left            =   6270
         TabIndex        =   15
         Top             =   4335
         Width           =   855
      End
      Begin VB.PictureBox plcAC 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         ScaleHeight     =   195
         ScaleWidth      =   6345
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   4470
         Width           =   6345
         Begin VB.OptionButton rbcAC 
            Caption         =   "None"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   5310
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   780
         End
         Begin VB.OptionButton rbcAC 
            Caption         =   "Not Same Break"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   3585
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   1755
         End
         Begin VB.OptionButton rbcAC 
            Caption         =   "Vehicle Rules"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   2085
            TabIndex        =   6
            Top             =   0
            Value           =   -1  'True
            Width           =   1545
         End
      End
      Begin VB.CheckBox ckcExcl 
         Caption         =   "Use Program Exclusions"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   75
         TabIndex        =   9
         Top             =   4695
         Value           =   1  'Checked
         Width           =   2430
      End
      Begin VB.CheckBox ckcAvail 
         Caption         =   "Honor Avail Names"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2520
         TabIndex        =   10
         Top             =   4695
         Value           =   1  'Checked
         Width           =   1965
      End
      Begin VB.CheckBox ckcPkgVeh 
         Caption         =   "Use Package Vehicles"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   75
         TabIndex        =   12
         Top             =   4935
         Value           =   1  'Checked
         Width           =   2265
      End
      Begin VB.CheckBox ckcOverride 
         Caption         =   "Use Package Overrides"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2520
         TabIndex        =   13
         Top             =   4935
         Width           =   2385
      End
      Begin VB.ComboBox cbcLen 
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
         Left            =   8130
         TabIndex        =   4
         Top             =   30
         Width           =   870
      End
      Begin VB.Label lacSpotType 
         Caption         =   "Spots:"
         Height          =   195
         Left            =   75
         TabIndex        =   65
         Top             =   60
         Width           =   615
      End
      Begin VB.Label lacLen 
         Appearance      =   0  'Flat
         Caption         =   "Spot Length"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   7065
         TabIndex        =   35
         Top             =   60
         Width           =   1050
      End
   End
   Begin VB.PictureBox plcMMGN 
      ForeColor       =   &H80000008&
      Height          =   5070
      Left            =   330
      ScaleHeight     =   5010
      ScaleWidth      =   9000
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   465
      Width           =   9060
      Begin VB.CommandButton cmcMGs 
         Appearance      =   0  'Flat
         Caption         =   "&MGs"
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
         Left            =   2475
         TabIndex        =   50
         Top             =   4605
         Width           =   1050
      End
      Begin VB.PictureBox plcVehDP 
         ForeColor       =   &H00000000&
         Height          =   3795
         Left            =   4515
         ScaleHeight     =   3735
         ScaleWidth      =   4425
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   315
         Width           =   4485
         Begin VB.PictureBox pbcLbcVehDp 
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
            Height          =   825
            Index           =   1
            Left            =   30
            ScaleHeight     =   825
            ScaleWidth      =   4365
            TabIndex        =   62
            Top             =   2880
            Width           =   4365
         End
         Begin VB.PictureBox pbcLbcVehDp 
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
            Height          =   2475
            Index           =   0
            Left            =   15
            ScaleHeight     =   2475
            ScaleWidth      =   4380
            TabIndex        =   61
            Top             =   270
            Width           =   4380
         End
         Begin VB.ListBox lbcVehDP 
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
            Index           =   0
            Left            =   0
            TabIndex        =   48
            Top             =   255
            Width           =   4425
         End
         Begin VB.HScrollBar hbcVehDPWk 
            Height          =   240
            Left            =   930
            Min             =   1
            TabIndex        =   47
            Top             =   15
            Value           =   1
            Width           =   3495
         End
         Begin VB.ListBox lbcVehDP 
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
            Left            =   0
            TabIndex        =   46
            Top             =   2865
            Width           =   4425
         End
         Begin VB.Label plcVehDpWk 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            TabIndex        =   49
            Top             =   15
            Width           =   930
         End
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
         Left            =   4980
         TabIndex        =   44
         Top             =   4605
         Width           =   1050
      End
      Begin VB.CommandButton cmcReplace 
         Appearance      =   0  'Flat
         Caption         =   "&Outside"
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
         Left            =   3750
         TabIndex        =   43
         Top             =   4605
         Width           =   1050
      End
      Begin VB.PictureBox plcScreen 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   60
         ScaleHeight     =   240
         ScaleWidth      =   3975
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   60
         Width           =   3975
      End
      Begin VB.PictureBox plcSpot 
         ForeColor       =   &H00000000&
         Height          =   3795
         Left            =   90
         ScaleHeight     =   3735
         ScaleWidth      =   4290
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   315
         Width           =   4350
         Begin VB.PictureBox pbcLbcSpots 
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
            Height          =   825
            Index           =   1
            Left            =   15
            ScaleHeight     =   825
            ScaleWidth      =   4230
            TabIndex        =   63
            Top             =   2880
            Width           =   4230
         End
         Begin VB.PictureBox pbcLbcSpots 
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
            Height          =   2490
            Index           =   0
            Left            =   15
            ScaleHeight     =   2490
            ScaleWidth      =   4245
            TabIndex        =   60
            Top             =   270
            Width           =   4245
         End
         Begin VB.ListBox lbcSpots 
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
            Left            =   0
            TabIndex        =   40
            Top             =   2865
            Width           =   4290
         End
         Begin VB.HScrollBar hbcSpotWk 
            Height          =   240
            Left            =   930
            Min             =   1
            TabIndex        =   39
            Top             =   15
            Value           =   1
            Width           =   3345
         End
         Begin VB.ListBox lbcSpots 
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
            Index           =   0
            Left            =   0
            Sorted          =   -1  'True
            TabIndex        =   38
            Top             =   255
            Width           =   4290
         End
         Begin VB.Label plcSpotWk 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            TabIndex        =   41
            Top             =   15
            Width           =   930
         End
      End
      Begin VB.Label lacAdjAud 
         Appearance      =   0  'Flat
         Caption         =   "Audience Adjustments to Date"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   5130
         TabIndex        =   56
         Top             =   60
         Width           =   2640
      End
      Begin VB.Label lacVehDPWk 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   4350
         TabIndex        =   55
         Top             =   4305
         Width           =   2820
      End
      Begin VB.Label lacSpotWk 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   54
         Top             =   4305
         Width           =   2820
      End
      Begin VB.Label lacVehDPWk 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   7890
         TabIndex        =   53
         Top             =   4305
         Width           =   840
      End
      Begin VB.Label lacVehDPWk 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   7035
         TabIndex        =   52
         Top             =   4305
         Width           =   840
      End
      Begin VB.Label lacSpotWk 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   2985
         TabIndex        =   51
         Top             =   4305
         Width           =   1005
      End
   End
   Begin ComctlLib.TabStrip plcTabSelection 
      Height          =   5685
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   10028
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "1 For 1 MG"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "M For N MGs"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   -15
      ScaleHeight     =   135
      ScaleWidth      =   45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4845
      Width           =   45
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   8895
      Top             =   5055
   End
   Begin VB.Label lacAdjAud 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   7875
      TabIndex        =   1
      Top             =   0
      Width           =   915
   End
End
Attribute VB_Name = "SpotMG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Spotmg.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmGhfSrchKey0                 tmGsfSrchKey0                                           *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SpotMG.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Price Grid Calculate input screen code
Option Explicit
Option Compare Text
'Contract line
Dim hmCHF As Integer        'Contract line file handle
Dim tmChf As CHF            'CHF record image
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim imCHFRecLen As Integer     'CHF record length
Dim tmChfAdvtExt() As CHFADVTEXT
'Contract line
Dim hmClf As Integer        'Contract line file handle
Dim tmClf() As CLFLIST            'CLF record image
Dim tmClfSrchKey As CLFKEY0 'CLF key record image
Dim imClfRecLen As Integer     'CLF record length
Dim tmIClf As CLF
Dim tmTCff(0 To 1) As CFF       '5-9-06
Dim tmPclf As CLF   'Package line
'Contract line flights
Dim hmCff As Integer        'Contract line file handle
Dim tmCff() As CFFLIST            'CLF record image
Dim tmICff(0 To 2) As CFF

Dim tmFCff() As CFF
Dim tmPCff As CFF
Dim tmCffSrchKey As CFFKEY0 'CLF key record image
Dim imCffRecLen As Integer     'CLF record length

Dim hmGhf As Integer
Dim tmGhf As GHF        'GHF record image
Dim tmGhfSrchKey1 As GHFKEY1    'GHF key record image
Dim imGhfRecLen As Integer        'GHF record length

Dim hmGsf As Integer
Dim tmGsf() As GSF        'GSF record image
Dim tmGsfSrchKey1 As GSFKEY1    'GSF key record image
Dim imGsfRecLen As Integer        'GSF record length

Dim hmCgf As Integer
Dim tmCgf As CGF
Dim tmCgfCff() As CFF
Dim imCgfRecLen As Integer
Dim tmCgfSrchKey1 As CGFKEY1    'CntrNo; CntRevNo; PropVer

Dim imBkQH As Integer
Dim imPriceLevel As Integer
Dim lmSepLength As Long 'Separation length for advertiser
Dim lmCompTime As Long  'Competitive time for vehicle
'Spot detail record information
Dim hmSdf As Integer        'Spot detail file handle
Dim tmSdf As SDF            'SDF record image
Dim tmSdfSrchKey0 As SDFKEY0
Dim tmSdfSrchKey3 As LONGKEY0
Dim imSdfRecLen As Integer  'SDF record length
Dim lmSdfRecPos() As Long
Dim lmTSdfRecPos() As Long
Dim lmBkDates() As Long
Dim lmTBkDates() As Long
'Spot summary
Dim hmSsf As Integer        'Spot summary file handle
Dim tmSsf As SSF
Dim tmSsfSrchKey As SSFKEY0 'SSF key record image
Dim imSsfRecLen As Integer  'SSF record length
Dim lmSsfRecPos As Long
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim tmSpotMove() As SPOTMOVE
Dim tmAvailIndex() As AVAILINDEX
Dim tmVcf0() As VCF
Dim tmVcf6() As VCF
Dim tmVcf7() As VCF
Dim tmRdf As RDF
Dim smRdfInOut As String
Dim imRdfAnfCode As Integer
'Spot MG record
Dim hmSmf As Integer        'Spot MG file handle
Dim tmSmf As SMF            'SMF record image
Dim imSmfRecLen As Integer  'SMF record length
' Spot Tracking File (only only if spots can be moved from Todays date+1 to Last log date)
Dim hmStf As Integer        'Spot tracking file handle
Dim tmStf As STF            'STF record image
Dim imStfRecLen As Integer  'STF record length
'Vehicle- required only for Multi User LOCK- don't allow two user into same vehicle
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF               'VEF record image
Dim imVefRecLen As Integer         'VEF record length
Dim imVpfIndex As Integer
Dim imVefCode As Integer
Dim imVehGame As Integer    'Flag indicating if spot vehicle was game vehicle
Dim lmLastLogDate As Long
Dim tmMGBookInfo() As MGBOOKINFO
Dim tmTMGBookInfo() As MGBOOKINFO
'MultiName
Dim hmMnf As Integer            'MultiName file handle
Dim tmMnf As MNF               'MNF record image
Dim imMnfRecLen As Integer         'MNF record length
'research Data
Dim hmDrf As Integer            'MultiName file handle
Dim tmDrf As DRF               'MNF record image
Dim imDrfRecLen As Integer         'MNF record length
'Plus Data
Dim hmDpf As Integer            'MultiName file handle
'Research Estimate
Dim hmDef As Integer
Dim hmRaf As Integer
'Record Lock
Dim hmRlf As Integer
'Feed
Dim hmFsf As Integer
Dim tmFsf As FSF            'FSF record image
Dim tmFSFSrchKey As LONGKEY0 'FSF key record image
Dim imFsfRecLen As Integer     'FSF record length
'Feed Name
Dim hmFnf As Integer

'Product
Dim hmPrf As Integer

Dim hmSxf As Integer

'Advertiser
Dim hmAdf As Integer            'Advertiser file handle
Dim tmAdf As ADF               'ADF record image
Dim tmAdfSrchKey As INTKEY0     'ADF key record image
Dim imAdfRecLen As Integer         'ADF record length
'Rotation
Dim hmCrf As Integer            'Rotation file handle
Dim tmCrf As CRF               'CRF record image
Dim imCrfRecLen As Integer         'CRF record length
'Field Areas
Dim imBlankDate As Integer
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imRcfIndex As Integer
Dim lmMissedWkDate As Long
Dim lmAiredWkDate As Long
Dim lmSpotWkDate As Long
Dim lmVehDPWkDate As Long
Dim lmEarliestMonDate As Long
Dim imSelectDelay As Integer
Dim imDelayType As Integer
Dim tmSdfExt() As SDFEXT    'Spot scheduled for a contract
Dim tmSdfExtSort() As SDFEXTSORT
Dim smSchStatus As String
Dim imMGGen As Integer
Dim imIgnoreChg As Integer
Dim imSelectIndex As Integer
Dim imSvIndex As Integer
Dim smAud As String
Dim smScreenCaption As String
'Dim imListField(1 To 7) As Integer
Dim imListField(0 To 7) As Integer
Dim imLBCtrls As Integer

'Dim imListFieldChar(1 To 6) As Integer
Dim imListFieldSpots(0 To 6) As Integer
'Dim imListFieldSpotsChar(1 To 5) As Integer
Dim imListFieldVehDp(0 To 7) As Integer
'Dim imListFieldVehDpChar(1 To 6) As Integer
Dim imTabSelection As Integer
Dim imUpdateAllowed As Integer

Dim imLengths() As Integer
'Dim lmTBStartTime(1 To 49) As Long  'Allowed times if time buy (-1 indicates end of times) and library times if library buy
'Dim lmTBEndTime(1 To 49) As Long
Dim lmTBStartTime(0 To 48) As Long  'Allowed times if time buy (-1 indicates end of times) and library times if library buy
Dim lmTBEndTime(0 To 48) As Long
''Required to be compatible with general schedule routines
''The array are not used by spots except for compatiblity
'Dim imHour(1 To 24) As Integer     'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
'Dim imDay(1 To 7) As Integer       'Day count (Index 1 for monday; Index 2 for tuesday;..)
'Dim imQH(1 To 4) As Integer        'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
''Actual for the day or week be processed- this will be a subset from
''imC---- or imP----
'Dim imAHour(1 To 24) As Integer    'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
'Dim imADay(1 To 7) As Integer      'Day count (Index 1 for monday; Index 2 for tuesday;..)
'Dim imAQH(1 To 4) As Integer       'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
'Dim imSkip(1 To 24, 1 To 4, 0 To 6) As Integer  '-1=Skip all test;0=All test;
Dim imHour(0 To 23) As Integer     'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
Dim imDay(0 To 6) As Integer       'Day count (Index 1 for monday; Index 2 for tuesday;..)
Dim imQH(0 To 3) As Integer        'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
'Actual for the day or week be processed- this will be a subset from
'imC---- or imP----
Dim imAHour(0 To 23) As Integer    'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
Dim imADay(0 To 6) As Integer      'Day count (Index 1 for monday; Index 2 for tuesday;..)
Dim imAQH(0 To 3) As Integer       'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
Dim imSkip(0 To 23, 0 To 3, 0 To 6) As Integer  '-1=Skip all test;0=All test;
                                    'Bit 0=Skip insert;
                                    'Bit 1=Skip move;
                                    'Bit 2=Skip competitive pack;
                                    'Bit 3=Skip Preempt

                                    'Bit 0=Skip insert;
                                    'Bit 1=Skip move;
                                    'Bit 2=Skip competitive pack;
                                    'Bit 3=Skip Preempt
Private Sub cbcLen_Change()
    If (imIgnoreChg = True) Or (plcAiredWk.Caption = "") Then
        imIgnoreChg = False
        Exit Sub
    End If
    igMGSpotLen = Val(cbcLen.List(cbcLen.ListIndex))
    tmcClick.Enabled = False
    'DoEvents
    imSelectDelay = True
    imDelayType = 2
    tmcClick.Interval = 3000    '2 seconds
    tmcClick.Enabled = True
End Sub
Private Sub cbcLen_Click()
    cbcLen_Change
End Sub
Private Sub cbcLen_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcDaysTimes_Click()
    If ckcDaysTimes.Value = vbChecked Then
        ckcOverride.Enabled = False
    Else
        ckcOverride.Enabled = True
    End If
End Sub

Private Sub ckcOverride_Click()
    If ckcOverride.Value = vbChecked Then
        ckcDaysTimes.Enabled = False
    Else
        ckcDaysTimes.Enabled = True
    End If
End Sub

Private Sub ckcSpotType_Click(Index As Integer)
    tmcClick.Enabled = False
    If (ckcSpotType(0).Value = vbChecked) And (ckcSpotType(1).Value = vbChecked) Then
        ckcAirWeek.Value = vbUnchecked
    Else
        If (ckcSpotType(0).Value = vbChecked) Then
            ckcAirWeek.Value = vbChecked
        Else
            ckcAirWeek.Value = vbUnchecked
        End If
    End If
    If Trim$(plcMissedWk.Caption) <> "" Then
        tmcClick.Enabled = False
        'DoEvents
        imSelectDelay = True
        imDelayType = 3
        tmcClick.Interval = 3000    '2 seconds
        tmcClick.Enabled = True
    End If
End Sub

Private Sub cmc1MG1Done_Click()
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If imMGGen Then
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmc1MG1MG_Click()
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    imMGGen = True
    smSchStatus = "G"
    mSchSpots
    imMGGen = False
End Sub
Private Sub cmc1MG1Outside_Click()
    imMGGen = True
    smSchStatus = "O"
    mSchSpots
    imMGGen = False
End Sub
Private Sub cmcDone_Click()
    mTerminate
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
    If (igWinStatus(SPOTSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
        If plcMMGN.Visible Then
            plcMMGN.Visible = False
            plcMMGN.Visible = True
        Else
            plc1MG1.Visible = False
            plc1MG1.Visible = True
        End If
    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        mTerminate
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tmGsf
    Erase tmFCff
    Erase tmCgfCff
    Erase tmSdfExt
    Erase tmSdfExtSort
    Erase tmClf
    Erase tmCff
    Erase tgVefRdfInfo
    Erase tgCntSpot
    Erase tgSpotLinks
    Erase tgClfSpot
    Erase tgCffSpot
    Erase lmSdfRecPos
    Erase lmTSdfRecPos
    Erase tmMGBookInfo
    Erase tmTMGBookInfo
    Erase lmBkDates
    Erase tmAvailIndex
    Erase imLengths
    Erase tmVcf0
    Erase tmVcf6
    Erase tmVcf7
    Erase tmSpotMove
    btrDestroy hmSxf
    btrDestroy hmRlf
    btrDestroy hmCrf
    btrDestroy hmCHF
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmSdf
    btrDestroy hmSsf
    btrDestroy hmSmf
    btrDestroy hmStf
    btrDestroy hmVef
    btrDestroy hmMnf
    btrDestroy hmDrf
    btrDestroy hmDpf
    btrDestroy hmDef
    btrDestroy hmRaf
    btrDestroy hmFsf
    btrDestroy hmGhf
    btrDestroy hmGsf
    btrDestroy hmCgf
    btrDestroy hmAdf
    
    Set SpotMG = Nothing   'Remove data segment
    
End Sub

Private Sub hbcAiredWk_Change()
    Dim slDate As String
    Dim llDate As Long
    If imIgnoreChg = True Then
        imIgnoreChg = False
        Exit Sub
    End If
    If (tmcClick.Enabled And (imDelayType <> 2)) Then
        Exit Sub
    End If
    DoEvents
    llDate = 7 * (hbcAiredWk.Value - 1) + lmEarliestMonDate
    slDate = Format(llDate, "m/d/yy")
    slDate = gFormatDate(slDate)
    plcAiredWk.Caption = slDate
    tmcClick.Enabled = False
    If hbcAiredWk.Value = hbcVehDPWk.Value Then
        hbcVehDPWk_Change
    Else
        hbcVehDPWk.Value = hbcAiredWk.Value
    End If
End Sub
Private Sub hbcAiredWk_GotFocus()
    If (tmcClick.Enabled And (imDelayType <> 2)) Then
        Exit Sub
    End If
End Sub
Private Sub hbcAiredWk_Scroll()
    Dim slDate As String
    Dim llDate As Long
    llDate = 7 * (hbcAiredWk.Value - 1) + lmEarliestMonDate
    slDate = Format(llDate, "m/d/yy")
    slDate = gFormatDate(slDate)
    If (plcAiredWk.Caption = "") Or (imBlankDate = True) Then
        imBlankDate = True
        plcAiredWk.Caption = slDate
        tmcClick.Enabled = False
        If hbcAiredWk.Value = hbcVehDPWk.Value Then
            hbcVehDPWk_Change
        Else
            hbcVehDPWk.Value = hbcAiredWk.Value
        End If
    Else
        plcAiredWk.Caption = slDate
        tmcClick.Enabled = False
    End If
End Sub
Private Sub hbcMissedWk_Change()
    Dim slDate As String
    Dim llDate As Long
    If imIgnoreChg = True Then
        imIgnoreChg = False
        Exit Sub
    End If
    If (tmcClick.Enabled And (imDelayType <> 3)) Then
        Exit Sub
    End If
    DoEvents
    llDate = 7 * (hbcMissedWk.Value - 1) + lmEarliestMonDate
    slDate = Format(llDate, "m/d/yy")
    slDate = gFormatDate(slDate)
    plcMissedWk.Caption = slDate
    tmcClick.Enabled = False
    'DoEvents
    imSelectDelay = True
    imDelayType = 3
    tmcClick.Interval = 3000    '2 seconds
    tmcClick.Enabled = True
End Sub
Private Sub hbcMissedWk_GotFocus()
    If (tmcClick.Enabled And (imDelayType <> 3)) Then
        Exit Sub
    End If
End Sub
Private Sub hbcMissedWk_Scroll()
    Dim slDate As String
    Dim llDate As Long
    llDate = 7 * (hbcMissedWk.Value - 1) + lmEarliestMonDate
    slDate = Format(llDate, "m/d/yy")
    slDate = gFormatDate(slDate)
    If (plcMissedWk.Caption = "") Or (imBlankDate) Then
        imBlankDate = True
        tmcClick.Enabled = False
        plcMissedWk.Caption = slDate
        imSelectDelay = True
        imDelayType = 3
        tmcClick.Interval = 3000    '2 seconds
        tmcClick.Enabled = True
    Else
        plcMissedWk.Caption = slDate
        tmcClick.Enabled = False
    End If
End Sub
Private Sub hbcSpotWk_Change()
    Dim slDate As String
    Dim llDate As Long
    If imIgnoreChg = True Then
        imIgnoreChg = False
        Exit Sub
    End If
    If (tmcClick.Enabled And (imDelayType <> 1)) Then
        Exit Sub
    End If
    DoEvents
    llDate = 7 * (hbcSpotWk.Value - 1) + lmEarliestMonDate
    slDate = Format(llDate, "m/d/yy")
    slDate = gFormatDate(slDate)
    plcSpotWk.Caption = slDate
    tmcClick.Enabled = False
    'DoEvents
    imSelectDelay = True
    imDelayType = 1
    tmcClick.Interval = 3000    '2 seconds
    tmcClick.Enabled = True
End Sub
Private Sub hbcSpotWk_GotFocus()
    If (tmcClick.Enabled And (imDelayType <> 1)) Then
        Exit Sub
    End If
End Sub
Private Sub hbcSpotWk_Scroll()
    Dim slDate As String
    Dim llDate As Long
    llDate = 7 * (hbcSpotWk.Value - 1) + lmEarliestMonDate
    slDate = Format(llDate, "m/d/yy")
    slDate = gFormatDate(slDate)
    If (plcSpotWk.Caption = "") Or (imBlankDate) Then
        imBlankDate = True
        tmcClick.Enabled = False
        plcSpotWk.Caption = slDate
        imSelectDelay = True
        imDelayType = 1
        tmcClick.Interval = 3000    '2 seconds
        tmcClick.Enabled = True
    Else
        plcSpotWk.Caption = slDate
        tmcClick.Enabled = False
    End If
End Sub
Private Sub hbcVehDPWk_Change()
    Dim slDate As String
    Dim llDate As Long
    If imIgnoreChg = True Then
        imIgnoreChg = False
        Exit Sub
    End If
    If (tmcClick.Enabled And (imDelayType <> 2)) Then
        Exit Sub
    End If
    DoEvents
    llDate = 7 * (hbcVehDPWk.Value - 1) + lmEarliestMonDate
    slDate = Format(llDate, "m/d/yy")
    slDate = gFormatDate(slDate)
    plcVehDpWk.Caption = slDate
    plcAiredWk.Caption = slDate
    tmcClick.Enabled = False
    'DoEvents
    imSelectDelay = True
    imDelayType = 2
    tmcClick.Interval = 3000    '2 seconds
    tmcClick.Enabled = True
End Sub
Private Sub hbcVehDPWk_GotFocus()
    If (tmcClick.Enabled And (imDelayType <> 2)) Then
        Exit Sub
    End If
End Sub
Private Sub hbcVehDPWk_Scroll()
    Dim slDate As String
    Dim llDate As Long
    llDate = 7 * (hbcVehDPWk.Value - 1) + lmEarliestMonDate
    slDate = Format(llDate, "m/d/yy")
    slDate = gFormatDate(slDate)
    If (plcVehDpWk.Caption = "") Or (imBlankDate) Then
        imBlankDate = True
        tmcClick.Enabled = False
        plcVehDpWk.Caption = slDate
        plcAiredWk.Caption = slDate
        imSelectDelay = True
        imDelayType = 2
        tmcClick.Interval = 3000    '2 seconds
        tmcClick.Enabled = True
    Else
        plcVehDpWk.Caption = slDate
        plcAiredWk.Caption = slDate
        tmcClick.Enabled = False
    End If
End Sub

Private Sub lbcAired_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilFound                       ilLoop                                                  *
'******************************************************************************************

    'Any selected
    mSetCommands
    pbcLbcAired_Paint
End Sub

Private Sub lbcAired_Scroll()
    pbcLbcAired_Paint
End Sub

Private Sub lbcMissed_Click()
    mSetCommands
    pbcLbcMissed_Paint
End Sub

Private Sub lbcMissed_Scroll()
    pbcLbcMissed_Paint
End Sub

Private Sub lbcSpots_Click(Index As Integer)
    pbcLbcSpots_Paint Index
End Sub

Private Sub lbcSpots_DblClick(Index As Integer)
    Dim ilSelectIndex As Integer
    Dim slStr As String
    Dim llTotal As Long
    Dim ilLoop As Integer
    Dim ilRet As Integer
    ilSelectIndex = lbcSpots(Index).ListIndex
    slStr = lbcSpots(Index).List(ilSelectIndex)
    If Index = 0 Then
        lbcSpots(1).AddItem slStr
        lbcSpots(0).RemoveItem ilSelectIndex
    Else
        lbcSpots(0).AddItem slStr
        lbcSpots(1).RemoveItem ilSelectIndex
    End If
    llTotal = 0
    For ilLoop = 0 To lbcSpots(1).ListCount - 1 Step 1
        slStr = lbcSpots(1).List(ilLoop)
        ilRet = gParseItem(slStr, 4, "|", slStr)
        llTotal = llTotal + Val(Trim$(slStr))
    Next ilLoop
    If llTotal = 0 Then
        lacSpotWk(1).Caption = ""
    Else
        lacSpotWk(1).Caption = Trim$(str$(llTotal))
    End If
    mComputeReq
    If (lacVehDPWk(1).Caption <> "") And (lacSpotWk(1).Caption <> "") Then
        lacVehDPWk(2).Caption = gLongToStrDec(100 * Val(lacVehDPWk(1).Caption) / Val(lacSpotWk(1)), 0) & "%"
    Else
        lacVehDPWk(2).Caption = ""
    End If
    If lbcSpots(1).ListCount > 1 Then
        lacSpotWk(0).Caption = Trim$(str$(lbcSpots(1).ListCount)) & " Spots for Total Audience of"
    ElseIf lbcSpots(1).ListCount = 1 Then
        lacSpotWk(0).Caption = Trim$(str$(lbcSpots(1).ListCount)) & " Spot for Total Audience of"
    Else
        lacSpotWk(0).Caption = ""
    End If
    pbcLbcSpots_Paint 0
    pbcLbcSpots_Paint 1
End Sub
Private Sub lbcSpots_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llTotal As Long
    Dim slStr As String
    Dim ilRet As Integer
    imSelectIndex = lbcSpots(Index).ListIndex
    imSvIndex = Index
    If lacSpotWk(1).Caption = "" Then
        llTotal = 0
    Else
        llTotal = Val(lacSpotWk(1).Caption)
    End If
    If Index = 0 Then
        slStr = lbcSpots(Index).List(imSelectIndex)
        ilRet = gParseItem(slStr, 4, "|", slStr)
        smAud = slStr
        llTotal = llTotal + Val(Trim$(slStr))
    Else
        slStr = lbcSpots(Index).List(imSelectIndex)
        ilRet = gParseItem(slStr, 4, "|", slStr)
        smAud = slStr
        llTotal = llTotal - Val(Trim$(slStr))
    End If
    lacSpotWk(1).Caption = Trim$(str$(llTotal))
End Sub
Private Sub lbcSpots_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (imSvIndex = Index) And (imSelectIndex = lbcSpots(Index).ListIndex) Then
        Exit Sub
    End If
    If imSelectIndex < 0 Then
        Exit Sub
    End If
    lbcSpots_MouseUp Index, Button, Shift, X, Y
    lbcSpots_MouseDown Index, Button, Shift, X, Y
End Sub
Private Sub lbcSpots_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llTotal As Long
    If imSelectIndex < 0 Then
        Exit Sub
    End If
    If lacSpotWk(1).Caption = "" Then
        llTotal = 0
    Else
        llTotal = Val(lacSpotWk(1).Caption)
    End If
    If imSvIndex = 0 Then
        llTotal = llTotal - Val(smAud)
    Else
        llTotal = llTotal + Val(smAud)
    End If
    If llTotal = 0 Then
        lacSpotWk(1).Caption = ""
    Else
        lacSpotWk(1).Caption = Trim$(str$(llTotal))
    End If
    smAud = ""
    imSelectIndex = -1
    imSvIndex = -1
End Sub

Private Sub lbcSpots_Scroll(Index As Integer)
    pbcLbcSpots_Paint Index
End Sub

Private Sub lbcVehDP_Click(Index As Integer)
    pbcLbcVehDp_Paint Index
End Sub

Private Sub lbcVehDP_DblClick(Index As Integer)
    Dim ilSelectIndex As Integer
    Dim slStr As String
    Dim llTotal As Long
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slStr1 As String
    Dim slStr2 As String
    Dim slStr3 As String
    Dim slStr4 As String
    Dim slDate As String
    Dim slCode As String
    Dim ilIndex As Integer
    ilSelectIndex = lbcVehDP(Index).ListIndex
    If Index = 0 Then
        slStr = lbcVehDP(Index).List(ilSelectIndex)
        slDate = plcVehDpWk.Caption
        ilRet = gParseItem(slStr, 1, "|", slStr1)
        ilRet = gParseItem(slStr, 2, "|", slStr2)
        ilRet = gParseItem(slStr, 3, "|", slStr3)
        ilRet = gParseItem(slStr, 4, "|", slStr4)
        ilRet = gParseItem(slStr, 6, "|", slCode)
        ilIndex = Val(Trim$(slCode))
        slStr = slStr1 & "|" & slStr2 & "|" & slStr3 & "|" & slStr4 & "|" & slDate & "|" & str$(hbcVehDPWk.Value) & "|" & str$(tgVefRdfInfo(ilIndex).iVefIndex) & "|" & str$(tgVefRdfInfo(ilIndex).iRdfIndex)
        'lbcVehDP(1).AddItem gAlignStringByPixel(slStr, "|", imListFieldVehDp(), imListFieldVehDpChar())
        lbcVehDP(1).AddItem slStr
        'lbcVehDP(0).RemoveItem ilSelectIndex
    Else
        'lbcVehDP(0).AddItem slStr
        lbcVehDP(1).RemoveItem ilSelectIndex
    End If
    llTotal = 0
    For ilLoop = 0 To lbcVehDP(1).ListCount - 1 Step 1
        slStr = lbcVehDP(1).List(ilLoop)
        ilRet = gParseItem(slStr, 4, "|", slStr)
        llTotal = llTotal + Val(Trim$(slStr))
    Next ilLoop
    If llTotal = 0 Then
        lacVehDPWk(1).Caption = ""
        lacVehDPWk(2).Caption = ""
    Else
        lacVehDPWk(1).Caption = Trim$(str$(llTotal))
        If lacSpotWk(1).Caption <> "" Then
            lacVehDPWk(2).Caption = gLongToStrDec(100 * llTotal / Val(lacSpotWk(1)), 0) & "%"
        Else
            lacVehDPWk(2).Caption = ""
        End If
    End If
    mComputeReq
    If lbcVehDP(1).ListCount > 1 Then
        lacVehDPWk(0).Caption = Trim$(str$(lbcVehDP(1).ListCount)) & " MGs for Total Audience of"
    ElseIf lbcVehDP(1).ListCount = 1 Then
        lacVehDPWk(0).Caption = Trim$(str$(lbcVehDP(1).ListCount)) & " MG for Total Audience of"
    Else
        lacVehDPWk(0).Caption = ""
    End If
    'pbcLbcVehDp_Paint 0
    pbcLbcVehDp_Paint 1
End Sub
Private Sub lbcVehDP_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llTotal As Long
    Dim slStr As String
    Dim ilRet As Integer
    imSelectIndex = lbcVehDP(Index).ListIndex
    imSvIndex = Index
    If lacVehDPWk(1).Caption = "" Then
        llTotal = 0
    Else
        llTotal = Val(lacVehDPWk(1).Caption)
    End If
    If Index = 0 Then
        slStr = lbcVehDP(Index).List(imSelectIndex)
        ilRet = gParseItem(slStr, 4, "|", slStr)
        smAud = Trim$(slStr)
        llTotal = llTotal + Val(Trim$(slStr))
    Else
        slStr = lbcVehDP(Index).List(imSelectIndex)
        ilRet = gParseItem(slStr, 4, "|", slStr)
        smAud = Trim$(slStr)
        llTotal = llTotal - Val(Trim$(slStr))
    End If
    lacVehDPWk(1).Caption = Trim$(str$(llTotal))
End Sub
Private Sub lbcVehDP_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (imSvIndex = Index) And (imSelectIndex = lbcVehDP(Index).ListIndex) Then
        Exit Sub
    End If
    If imSelectIndex < 0 Then
        Exit Sub
    End If
    lbcVehDP_MouseUp Index, Button, Shift, X, Y
    lbcVehDP_MouseDown Index, Button, Shift, X, Y
End Sub
Private Sub lbcVehDP_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llTotal As Long
    If imSelectIndex < 0 Then
        Exit Sub
    End If
    If lacVehDPWk(1).Caption = "" Then
        llTotal = 0
    Else
        llTotal = Val(lacVehDPWk(1).Caption)
    End If
    If imSvIndex = 0 Then
        llTotal = llTotal - Val(smAud)
    Else
        llTotal = llTotal + Val(smAud)
    End If
    If llTotal = 0 Then
        lacVehDPWk(1).Caption = ""
    Else
        lacVehDPWk(1).Caption = Trim$(str$(llTotal))
    End If
    smAud = ""
    imSelectIndex = -1
    imSvIndex = -1
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAnyConflicts                   *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test if Ok to book spot        *
'*                                                     *
'*******************************************************
Private Function mAnyConflicts(ilAvailIndex As Integer, ilAdfCode As Integer, ilMnfComp0 As Integer, ilMnfComp1 As Integer) As Integer
    Dim ilSpotIndex As Integer
    Dim ilMatchComp As Integer
    Dim tlAvail As AVAILSS
    If rbcAC(2).Value Then
        'None
        mAnyConflicts = False
        Exit Function
    ElseIf rbcAC(0).Value Then
        'Vehicle rules
        If Not gAdvtTest(hmSsf, tmSsf, lmSsfRecPos, tmSpotMove(), imVpfIndex, lmSepLength, ilAvailIndex, tmChf.iAdfCode, tmChf.iMnfComp(0), tmChf.iMnfComp(1), 4, 0, "I", "N", imPriceLevel, True) Then
            mAnyConflicts = True
            Exit Function
        End If
        If Not gCompetitiveTest(lmCompTime, hmSsf, tmSsf, lmSsfRecPos, tmSpotMove(), imVpfIndex, tmIClf.iLen, tmChf.iMnfComp(0), tmChf.iMnfComp(1), ilAvailIndex, tmVcf0(), tmVcf6(), tmVcf7(), 4, 0, "I", "N", imPriceLevel, True) Then
            mAnyConflicts = True
            Exit Function
        End If
    Else
        'Not same break
       LSet tlAvail = tmSsf.tPas(ADJSSFPASBZ + ilAvailIndex)
        For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tlAvail.iNoSpotsThis Step 1
           LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
            ilMatchComp = False
            If (ilMnfComp0 = 0) And (ilMnfComp1 = 0) And (tmSpot.iMnfComp(0) = 0) And (tmSpot.iMnfComp(1) = 0) Then
                ilMatchComp = True
            Else
                If (ilMnfComp0 <> 0) And ((ilMnfComp0 = tmSpot.iMnfComp(0)) Or (ilMnfComp0 = tmSpot.iMnfComp(1))) Then
                    ilMatchComp = True
                End If
                If (ilMnfComp1 <> 0) And ((ilMnfComp1 = tmSpot.iMnfComp(0)) Or (ilMnfComp1 = tmSpot.iMnfComp(1))) Then
                    ilMatchComp = True
                End If
            End If
            If (tmSpot.iAdfCode = ilAdfCode) And (ilMatchComp) Then
                mAnyConflicts = True
                Exit Function
            End If
            If (ilMnfComp0 <> 0) And ((ilMnfComp0 = tmSpot.iMnfComp(0)) Or (ilMnfComp0 = tmSpot.iMnfComp(1))) Then
                mAnyConflicts = True
                Exit Function
            ElseIf (ilMnfComp1 <> 0) And ((ilMnfComp1 = tmSpot.iMnfComp(0)) Or (ilMnfComp1 = tmSpot.iMnfComp(1))) Then
                mAnyConflicts = True
                Exit Function
            End If
        Next ilSpotIndex
    End If
    mAnyConflicts = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAvailCount                     *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine avail count          *
'*                                                     *
'*******************************************************
Private Sub mAvailCount(hlSsf As Integer, hlSdf As Integer, ilVefCode As Integer, llStartDate As Long, llEndDate As Long, ilIncludeMissed As Integer)
'
'
'   lgSchChfCode(I)
'   igComp1Code(I)
'   igComp2Code(I)
'

    Dim ilType As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilDay As Integer
    Dim ilEvt As Integer
    Dim ilRet As Integer
    Dim ilSpot As Integer
    Dim llTime As Long
    Dim ilLoop As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilLen As Integer
    Dim ilUnits As Integer
    Dim ilAvailOk As Integer
    Dim ilVef As Integer
    Dim llLoopDate As Long
    Dim ilVpfIndex As Integer
    Dim ilRdfIndex As Integer
    Dim ilLtfCode As Integer
    Dim ilGsf As Integer
    Dim llGsfDate As Long

    ilVpfIndex = gBinarySearchVpf(ilVefCode)    'gVpfFind(SpotMG, ilVefCode)
    If ilVpfIndex < 0 Then
        Exit Sub
    End If
    ilType = 0
    ilVef = gBinarySearchVef(ilVefCode)
    If ilVef = -1 Then
        Exit Sub
    End If
    ReDim tmGsf(0 To 1) As GSF
    tmGsf(0).iGameNo = 0
    If tgMVef(ilVef).sType = "G" Then
        ilRet = mGhfGsfReadRec(ilVefCode, llStartDate, llEndDate)
    End If
    For llLoopDate = llStartDate To llEndDate Step 1
        For ilGsf = 0 To UBound(tmGsf) - 1 Step 1
            ilType = tmGsf(ilGsf).iGameNo
            If ilType <> 0 Then
                gUnpackDateLong tmGsf(ilGsf).iAirDate(0), tmGsf(ilGsf).iAirDate(1), llGsfDate
            Else
                llGsfDate = llLoopDate
            End If
            If llLoopDate = llGsfDate Then
                slDate = Format$(llLoopDate, "m/d/yy")
                ilDay = gWeekDayStr(slDate)
                gPackDate slDate, ilDate0, ilDate1
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                tmSsfSrchKey.iType = ilType
                tmSsfSrchKey.iVefCode = ilVefCode
                tmSsfSrchKey.iDate(0) = ilDate0
                tmSsfSrchKey.iDate(1) = ilDate1
                tmSsfSrchKey.iStartTime(0) = 0
                tmSsfSrchKey.iStartTime(1) = 0
                ilRet = gSSFGetGreaterOrEqual(hlSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
                    gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llDate
                    ilEvt = 1
                    Do While ilEvt <= tmSsf.iCount
                       LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                        If tmProg.iRecType = 1 Then    'Program (not working for nested prog)
                            ilLtfCode = tmProg.iLtfCode
                        ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 2) Then 'Contract Avails only
                            For ilVef = LBound(tgVefRdfInfo) To UBound(tgVefRdfInfo) - 1 Step 1
                                If ilVefCode = tgMVef(tgVefRdfInfo(ilVef).iVefIndex).iCode Then
                                    ilRdfIndex = tgVefRdfInfo(ilVef).iRdfIndex
                                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                                    'Determine which rate card program this is associated with
                                    ilAvailOk = False
                                    If (tgMRdf(ilRdfIndex).iLtfCode(0) <> 0) Or (tgMRdf(ilRdfIndex).iLtfCode(1) <> 0) Or (tgMRdf(ilRdfIndex).iLtfCode(2) <> 0) Then
                                        If (ilLtfCode = tgMRdf(ilRdfIndex).iLtfCode(0)) Or (ilLtfCode = tgMRdf(ilRdfIndex).iLtfCode(1)) Or (ilLtfCode = tgMRdf(ilRdfIndex).iLtfCode(1)) Then
                                            ilAvailOk = False    'True- code later
                                        End If
                                    Else
                                        For ilLoop = LBound(tgMRdf(ilRdfIndex).iStartTime, 2) To UBound(tgMRdf(ilRdfIndex).iStartTime, 2) Step 1 'Row
                                            If (tgMRdf(ilRdfIndex).iStartTime(0, ilLoop) <> 1) Or (tgMRdf(ilRdfIndex).iStartTime(1, ilLoop) <> 0) Then
                                                gUnpackTimeLong tgMRdf(ilRdfIndex).iStartTime(0, ilLoop), tgMRdf(ilRdfIndex).iStartTime(1, ilLoop), False, llStartTime
                                                gUnpackTimeLong tgMRdf(ilRdfIndex).iEndTime(0, ilLoop), tgMRdf(ilRdfIndex).iEndTime(1, ilLoop), True, llEndTime
                                                'If (llTime >= llStartTime) And (llTime < llEndTime) And (tgMRdf(ilRdfIndex).sWkDays(ilLoop, ilDay + 1) = "Y") Then
                                                If (llTime >= llStartTime) And (llTime < llEndTime) And (tgMRdf(ilRdfIndex).sWkDays(ilLoop, ilDay) = "Y") Then
                                                    ilAvailOk = True
                                                    Exit For
                                                End If
                                            End If
                                        Next ilLoop
                                    End If
                                    If ilAvailOk Then
                                        If tgMRdf(ilRdfIndex).sInOut = "I" Then   'Book into
                                            If tmAvail.ianfCode <> tgMRdf(ilRdfIndex).ianfCode Then
                                                ilAvailOk = False
                                            End If
                                        ElseIf tgMRdf(ilRdfIndex).sInOut = "O" Then   'Exclude
                                            If tmAvail.ianfCode = tgMRdf(ilRdfIndex).ianfCode Then
                                                ilAvailOk = False
                                            End If
                                        End If
                                    End If
                                    If ilAvailOk Then
                                        If tgVefRdfInfo(ilVef).iAvail = -9999 Then
                                            tgVefRdfInfo(ilVef).iAvail = 0
                                        End If
                                        ilLen = tmAvail.iLen
                                        ilUnits = tmAvail.iAvInfo And &H1F
                                        For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                           LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)
                                            If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                                If (tgVpf(ilVpfIndex).sSSellOut = "B") Or (tgVpf(ilVpfIndex).sSSellOut = "U") Then
                                                    ilUnits = ilUnits - 1
                                                    ilLen = ilLen - (tmSpot.iPosLen And &HFFF)
                                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                                    ilUnits = ilUnits - 1
                                                    ilLen = ilLen - (tmSpot.iPosLen And &HFFF)
                                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                                    'slSpotLen = Trim$(Str$(tmSpot.iPosLen And &HFFF))
                                                    'slStr = gDivStr(slSpotLen, "30.0")
                                                    'slUnits = gSubStr(slUnits, slSpotLen)
                                                End If
                                            End If
                                        Next ilSpot                             'loop from ssf file for # spots in avail
                                        If (tgVpf(ilVpfIndex).sSSellOut = "B") Or (tgVpf(ilVpfIndex).sSSellOut = "U") Or (tgVpf(ilVpfIndex).sSSellOut = "M") Then
                                            If (ilUnits > 0) And (ilLen > 0) Then
                                                If (igMGSpotLen = 30) Or (igMGSpotLen = 60) Then
                                                    Do While (ilLen >= igMGSpotLen) And (ilUnits >= 1)
                                                        tgVefRdfInfo(ilVef).iAvail = tgVefRdfInfo(ilVef).iAvail + 1
                                                        ilUnits = ilUnits - 1
                                                        ilLen = ilLen - igMGSpotLen
                                                    Loop
                                                Else
                                                    If (tmAvail.iLen) = igMGSpotLen Then
                                                        tgVefRdfInfo(ilVef).iAvail = tgVefRdfInfo(ilVef).iAvail + 1
                                                    End If
                                                End If
                                            End If
                                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                        End If
                                    End If                                          'Avail OK
                                End If
                            Next ilVef
                            ilEvt = ilEvt + tmAvail.iNoSpotsThis                'bypass spots
                        End If
                        ilEvt = ilEvt + 1   'Increment to next event
                    Loop                                                        'do while ilEvt <= tmSsf.iCount
                    imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    ilRet = gSSFGetNext(hlSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
        Next ilGsf
    Next llLoopDate
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBookSpot                       *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get 30sec count for week       *
'*                                                     *
'*******************************************************
Private Function mBookSpot(llSdfRecPos As Long, ilVefCode As Integer, ilRdfIndex As Integer, llSDate As Long, llEDate As Long, llSTime As Long, llETime As Long, ilPass As Integer) As Integer
'
'
    Dim slDate As String
    Dim slMonDate As String
    Dim llDate As Long
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilEvt As Integer
    Dim ilAvEvt As Integer
    Dim ilUnits As Integer
    Dim ilLen As Integer
    Dim ilPosition As Integer
    Dim llTime As Long
    Dim ilType As Integer
    Dim ilBkQH As Integer
    Dim slSchStatus As String
    Dim ilSpot As Integer
    Dim ilDay As Integer
    Dim ilAvailOk As Integer
    Dim ilExcl As Integer
    Dim ilAvail As Integer
    Dim ilRow As Integer
    Dim ilRdf As Integer
    Dim ilIndex As Integer
    Dim ilVpfIndex As Integer
    Dim slSdfAirWeek As String
    Dim llExclEndTime As Long
    Dim ilLenOk As Integer
    Dim ilTestEvt As Integer
    Dim llLockRecCode As Long
    Dim slUserName As String
    Dim ilBookOk As Integer
    Dim tlSdf As SDF
    Dim ilVef As Integer
    Dim ilGameNo As Integer
    Dim ilGsf As Integer
    Dim llGsfDate As Long
    Dim llMissedDate As Long

    ilPosition = 0  '-1
    ilType = 0
    'ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
    'If ilRet <> BTRV_ERR_NONE Then
    '    mBookSpot = -1
    '    Exit Function
    'End If
    If tmSdf.sSchStatus <> "M" Then
        mBookSpot = 1
        Exit Function
    End If
    If llETime < llSTime Then
        mBookSpot = 0
        Exit Function
    End If
    'tmChfSrchKey.lCode = tmSdf.lChfCode
    'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    'If ilRet <> BTRV_ERR_NONE Then
    '    mBookSpot = -1
    '    Exit Function
    'End If
    ilVef = gBinarySearchVef(ilVefCode)
    If ilVef = -1 Then
        mBookSpot = 0
        Exit Function
    End If
    ReDim tmGsf(0 To 1) As GSF
    tmGsf(0).iGameNo = 0
    If tgMVef(ilVef).sType = "G" Then
        ilRet = mGhfGsfReadRec(ilVefCode, llSDate, llEDate)
    End If
    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slSdfAirWeek
    llMissedDate = gDateValue(slSdfAirWeek)
    slSdfAirWeek = gObtainPrevMonday(slSdfAirWeek)
    mBookSpot = 0
    For llDate = llSDate To llEDate Step 1
        For ilGsf = 0 To UBound(tmGsf) - 1 Step 1
            ilType = tmGsf(ilGsf).iGameNo
            ilGameNo = tmGsf(ilGsf).iGameNo
            If ilType <> 0 Then
                gUnpackDateLong tmGsf(ilGsf).iAirDate(0), tmGsf(ilGsf).iAirDate(1), llGsfDate
            Else
                llGsfDate = llDate
            End If
            If llDate = llGsfDate Then
                llExclEndTime = 0
                ilDay = gWeekDayLong(llDate)
                slDate = Format$(llDate, "m/d/yy")
                If ilType = 0 Then
                    llLockRecCode = gCreateLockRec(hmRlf, "S", "S", 65536 * ilVefCode + llDate, False, slUserName)
                Else
                    llLockRecCode = gCreateLockRec(hmRlf, "S", "S", 65536 * ilVefCode + ilType, False, slUserName)
                End If
                If llLockRecCode > 0 Then
                    slMonDate = gObtainPrevMonday(slDate)
                    ReDim tmAvailIndex(0 To 0) As AVAILINDEX
                    imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    gPackDate slDate, ilLogDate0, ilLogDate1
                    tmSsfSrchKey.iType = ilType
                    tmSsfSrchKey.iVefCode = ilVefCode
                    tmSsfSrchKey.iDate(0) = ilLogDate0
                    tmSsfSrchKey.iDate(1) = ilLogDate1
                    tmSsfSrchKey.iStartTime(0) = 0
                    tmSsfSrchKey.iStartTime(1) = 0
                    ilRet = gSSFGetEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode) And (tmSsf.iDate(0) = ilLogDate0) And (tmSsf.iDate(1) = ilLogDate1)
                        ilRet = gSSFGetPosition(hmSsf, lmSsfRecPos)
                        'Set all fill spot to low priorty (2000)
                        ilEvt = 1
                        Do While ilEvt <= tmSsf.iCount
                           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                            'If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                            If (tmAvail.iRecType = 2) Then
                                For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                    ilEvt = ilEvt + 1
                                   LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                    If ((tmSpot.iRank And RANKMASK) > 1000) Then
                                        'Possible Fill Spot
                                        tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                        ilRet = btrGetEqual(hmSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet = BTRV_ERR_NONE Then
                                            If tlSdf.sSpotType = "X" Then
                                                'Don't need to reset as prior to booking spots the ssf is read back in
                                                tmSpot.iRank = (tmSpot.iRank And PRICELEVELMASK) + 2000
                                                LSet tmSsf.tPas(ADJSSFPASBZ + ilEvt) = tmSpot
                                            End If
                                        End If
                                    End If
                                Next ilSpot
                            End If
                            ilEvt = ilEvt + 1
                        Loop
                        ilEvt = 1
                        Do While ilEvt <= tmSsf.iCount
                           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                            'If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                            If (tmAvail.iRecType = 2) Then
                                gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                                If llTime >= llETime Then
                                    Exit Do
                                End If
                                If llTime >= llSTime Then
                                    ilExcl = False
                                    ilAvail = False
                                    'Test if closed avail
                                    ilAvailOk = True
                                    If ((tmAvail.iAvInfo And SSLOCK) = SSLOCK) Then
                                        ilAvailOk = False
                                    End If
                                    'Remove this test because if Honor Avails was unchecking then the spot will still not book into any avail
        '                            If ilAvailOk Then
        '                                If tgMRdf(ilRdfIndex).sInOut = "I" Then   'Book into
        '                                    If tmAvail.ianfCode <> tgMRdf(ilRdfIndex).ianfCode Then
        '                                        ilAvailOk = False
        '                                    End If
        '                                ElseIf tgMRdf(ilRdfIndex).sInOut = "O" Then   'Exclude
        '                                    If tmAvail.ianfCode = tgMRdf(ilRdfIndex).ianfCode Then
        '                                        ilAvailOk = False
        '                                    End If
        '                                End If
        '                            End If
                                    If smRdfInOut = "I" Then
                                        If (tmAvail.ianfCode <> imRdfAnfCode) Then
                                            ilAvail = True
                                        End If
                                    ElseIf smRdfInOut = "O" Then
                                        If (tmAvail.ianfCode = imRdfAnfCode) Then
                                            ilAvail = True
                                        End If
                                    Else    'Book into any avail which allows sustaining spots
                                        If (tmAvail.iAvInfo And SSSUSTAINING) <> SSSUSTAINING Then
                                            ilAvail = True
                                        End If
                                    End If
                                    If (ckcAvail.Value = vbChecked) And ilAvail Then
                                        ilAvailOk = False
                                    End If
                                    If llTime < llExclEndTime Then
                                        ilExcl = True
                                    Else
                                        llExclEndTime = 0
                                    End If
                                    If (ckcExcl.Value = vbChecked) And ilExcl Then
                                        ilAvailOk = False
                                    End If
                                    If tmIClf.sSoloAvail = "Y" Then
                                        If tmAvail.iLen <> tmSdf.iLen Then
                                            ilAvailOk = False
                                        Else
                                            If tmAvail.iNoSpotsThis > 0 Then
                                                ilAvailOk = False
                                            End If
                                        End If
                                    End If
                                    If tmIClf.iPosition = 1 Then
                                        If tmAvail.iNoSpotsThis > 0 Then
                                           LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + 1)
                                            '3/22/10
                                            'If (tmSpot.iPosLen And SS1STPOSITION) = SS1STPOSITION Then
                                            If (tmSpot.lBkInfo And SS1STPOSITION) = SS1STPOSITION Then
                                                ilAvailOk = False
                                            End If
                                        End If
                                    End If
                                    If ilAvailOk Then
                                        ilAvEvt = ilEvt
                                        'Test if within selected times
                                        ilLen = tmAvail.iLen
                                        ilUnits = tmAvail.iAvInfo And &H1F
                                        For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                            ilEvt = ilEvt + 1
                                           LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                            If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                                ilUnits = ilUnits - 1
                                                ilLen = ilLen - (tmSpot.iPosLen And &HFFF)
                                            End If
                                        Next ilSpot
                                        ilLenOk = False
                                        'ReDim tmSpotMove(1 To 1) As SPOTMOVE
                                        ReDim tmSpotMove(0 To 0) As SPOTMOVE
                                        If ilPass = 0 Then
                                            If (ilUnits > 0) And (tmSdf.iLen <= ilLen) Then
                                                If (tmSdf.iLen = 30) Or (tmSdf.iLen = 60) Or (tmSdf.iLen = tmAvail.iLen) Then
                                                    ilLenOk = True
                                                End If
                                            End If
                                        Else
                                            ilLenOk = False
                                            If (ilUnits <= 0) Or (tmSdf.iLen > ilLen) Then
                                                ilTestEvt = ilAvEvt
                                                For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                                    ilTestEvt = ilTestEvt + 1
                                                   LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilTestEvt)
                                                    If ((tmSpot.iRank And RANKMASK) = 2000) And (tmSdf.iLen >= (tmSpot.iPosLen And &HFFF)) Then
                                                        'Possible Fill Spot
                                                        'tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                                        'ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        'If ilRet = BTRV_ERR_NONE Then
                                                        '    If tmSdf.sSpotType = "X" Then
                                                                'ReDim tmSpotMove(1 To 2) As SPOTMOVE
                                                                'tmSpotMove(1).iSpotIndex = ilTestEvt
                                                                'tmSpotMove(1).lSpotSsfRecPos = lmSsfRecPos
                                                                'tmSpotMove(1).lSdfCode = tmSpot.lSdfCode
                                                                ReDim tmSpotMove(0 To 1) As SPOTMOVE
                                                                tmSpotMove(0).iSpotIndex = ilTestEvt
                                                                tmSpotMove(0).lSpotSsfRecPos = lmSsfRecPos
                                                                tmSpotMove(0).lSdfCode = tmSpot.lSdfCode
                                                                If Not mAnyConflicts(ilAvEvt, tmSdf.iAdfCode, tmChf.iMnfComp(0), tmChf.iMnfComp(1)) Then
                                                                    tmAvailIndex(UBound(tmAvailIndex)).lSsfRecPos = lmSsfRecPos
                                                                    tmAvailIndex(UBound(tmAvailIndex)).iAvEvt = ilAvEvt
                                                                    tmAvailIndex(UBound(tmAvailIndex)).lEvtTime = llTime
                                                                    tmAvailIndex(UBound(tmAvailIndex)).iAvail = ilAvail
                                                                    tmAvailIndex(UBound(tmAvailIndex)).iExcl = ilExcl
                                                                    'tmAvailIndex(UBound(tmAvailIndex)).lFill1SdfCode = tmSpotMove(1).lSdfCode
                                                                    tmAvailIndex(UBound(tmAvailIndex)).lFill1SdfCode = tmSpotMove(0).lSdfCode
                                                                    'If UBound(tmSpotMove) = 3 Then
                                                                    If UBound(tmSpotMove) = 2 Then
                                                                        'tmAvailIndex(UBound(tmAvailIndex)).lFill2SdfCode = tmSpotMove(2).lSdfCode
                                                                        tmAvailIndex(UBound(tmAvailIndex)).lFill2SdfCode = tmSpotMove(1).lSdfCode
                                                                    Else
                                                                        tmAvailIndex(UBound(tmAvailIndex)).lFill2SdfCode = 0
                                                                    End If
                                                                    ReDim Preserve tmAvailIndex(0 To UBound(tmAvailIndex) + 1) As AVAILINDEX
                                                                    Exit For
                                                                End If
                                                        '    End If
                                                        'End If
                                                    End If
                                                Next ilSpot
                                            End If
                                        End If
                                        If ilLenOk Then
                                            If Not mAnyConflicts(ilAvEvt, tmSdf.iAdfCode, tmChf.iMnfComp(0), tmChf.iMnfComp(1)) Then
                                                tmAvailIndex(UBound(tmAvailIndex)).lSsfRecPos = lmSsfRecPos
                                                tmAvailIndex(UBound(tmAvailIndex)).iAvEvt = ilAvEvt
                                                tmAvailIndex(UBound(tmAvailIndex)).lEvtTime = llTime
                                                tmAvailIndex(UBound(tmAvailIndex)).iAvail = ilAvail
                                                tmAvailIndex(UBound(tmAvailIndex)).iExcl = ilExcl
                                                tmAvailIndex(UBound(tmAvailIndex)).lFill1SdfCode = 0
                                                tmAvailIndex(UBound(tmAvailIndex)).lFill2SdfCode = 0
                                                If UBound(tmSpotMove) > LBound(tmSpotMove) Then
                                                    'tmAvailIndex(UBound(tmAvailIndex)).lFill1SdfCode = tmSpotMove(1).lSdfCode
                                                    tmAvailIndex(UBound(tmAvailIndex)).lFill1SdfCode = tmSpotMove(0).lSdfCode
                                                End If
                                                ReDim Preserve tmAvailIndex(0 To UBound(tmAvailIndex) + 1) As AVAILINDEX
                                            End If
                                        End If
                                        'ilEvt = ilEvt + tmAvail.iNoSpotsThis
                                    Else
                                        ilEvt = ilEvt + tmAvail.iNoSpotsThis
                                    End If
                                Else
                                    ilEvt = ilEvt + tmAvail.iNoSpotsThis
                                End If
                            ElseIf tmAvail.iRecType = 1 Then
                                If ckcExcl.Value = vbChecked Then
                                   LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                    If (tmChf.iMnfExcl(0) <> 0) And ((tmChf.iMnfExcl(0) = tmProg.iMnfExcl(0)) Or (tmChf.iMnfExcl(0) = tmProg.iMnfExcl(1))) Then
                                        gUnpackTimeLong tmProg.iEndTime(0), tmProg.iEndTime(1), True, llExclEndTime
                                    ElseIf (tmChf.iMnfExcl(1) <> 0) And ((tmChf.iMnfExcl(1) = tmProg.iMnfExcl(0)) Or (tmChf.iMnfExcl(1) = tmProg.iMnfExcl(1))) Then
                                        gUnpackTimeLong tmProg.iEndTime(0), tmProg.iEndTime(1), True, llExclEndTime
                                    End If
                                End If
                            End If
                            ilEvt = ilEvt + 1
                            DoEvents
                        Loop
                        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                        ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    'Reset Fill Rank

                    If UBound(tmAvailIndex) > LBound(tmAvailIndex) Then
                        ilIndex = Int((UBound(tmAvailIndex)) * Rnd + 1) - 1
                        lmSsfRecPos = tmAvailIndex(ilIndex).lSsfRecPos
                        ilAvEvt = tmAvailIndex(ilIndex).iAvEvt
                        ilAvail = tmAvailIndex(ilIndex).iAvail
                        ilExcl = tmAvailIndex(ilIndex).iExcl
                        ilRet = btrBeginTrans(hmSdf, 1000)
                        If ilRet <> BTRV_ERR_NONE Then
                            ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                            mBookSpot = -1
                            Exit Function
                        End If
                        'Add Test if not vioding contract then set status = to "S"
                        slSchStatus = smSchStatus
                        ilBkQH = 0
                        If tmChf.iPctTrade = 100 Then
                            ilBkQH = 1040
                        End If
                        If (tmIClf.iVefCode = ilVefCode) And (ilAvail = False) And (ilExcl = False) And (gDateValue(slMonDate) = gDateValue(slSdfAirWeek)) And (tmICff(1).sDelete <> "Y") Then
                            ilVpfIndex = gBinarySearchVpf(ilVefCode)    'gVpfFind(SpotMG, ilVefCode)
                            If ((tmIClf.iStartTime(0) <> 1) Or (tmIClf.iStartTime(1) <> 0)) And (tgVpf(ilVpfIndex).sGMedium <> "S") Then
                                gUnpackTimeLong tmIClf.iStartTime(0), tmIClf.iStartTime(1), False, llTime
                                If tmAvailIndex(ilIndex).lEvtTime >= llTime Then
                                    gUnpackTimeLong tmIClf.iEndTime(0), tmIClf.iEndTime(1), True, llTime
                                    If tmAvailIndex(ilIndex).lEvtTime < llTime Then
                                        If tmICff(1).sDyWk = "D" Then
                                            If llMissedDate = llDate Then
                                                If tmICff(1).iDay(ilDay) > 0 Then
                                                    slSchStatus = "S"
                                                    ilBkQH = imBkQH
                                                End If
                                            End If
                                        Else
                                            If tmICff(1).iDay(ilDay) > 0 Then
                                                slSchStatus = "S"
                                                ilBkQH = imBkQH
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                                '    If tgMRdf(ilRdf).iCode = tmIClf.iRdfcode Then
                                    ilRdf = gBinarySearchRdf(tmIClf.iRdfCode)
                                    If ilRdf <> -1 Then
                                        For ilRow = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                            If (tgMRdf(ilRdf).iStartTime(0, ilRow) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilRow) <> 0) Then
                                                gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilRow), tgMRdf(ilRdf).iStartTime(1, ilRow), False, llTime
                                                If tmAvailIndex(ilIndex).lEvtTime >= llTime Then
                                                    gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilRow), tgMRdf(ilRdf).iEndTime(1, ilRow), True, llTime
                                                    If tmAvailIndex(ilIndex).lEvtTime < llTime Then
                                                        If tmICff(1).sDyWk = "D" Then
                                                            If llMissedDate = llDate Then
                                                                If tmICff(1).iDay(ilDay) > 0 Then
                                                                    slSchStatus = "S"
                                                                    ilBkQH = imBkQH
                                                                    Exit For
                                                                End If
                                                            End If
                                                        Else
                                                            If tmICff(1).iDay(ilDay) > 0 Then
                                                                slSchStatus = "S"
                                                                ilBkQH = imBkQH
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                             End If
                                        Next ilRow
                                '        Exit For
                                    End If
                                'Next ilRdf
                            End If
                        End If
                        ilBookOk = True
                        If tmAvailIndex(ilIndex).lFill1SdfCode > 0 Then
                            lmSsfRecPos = 0
                            llExclEndTime = 0
                            tmSdfSrchKey3.lCode = tmAvailIndex(ilIndex).lFill1SdfCode
                            ilRet = btrGetEqual(hmSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                ilRet = gChgSchSpot("D", hmSdf, tlSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf, llExclEndTime, lmSsfRecPos, hmSxf, hmGsf, hmGhf)
                                If tmAvailIndex(ilIndex).lFill2SdfCode > 0 Then
                                    tmSdfSrchKey3.lCode = tmAvailIndex(ilIndex).lFill2SdfCode
                                    ilRet = btrGetEqual(hmSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet = BTRV_ERR_NONE Then
                                        ilRet = gChgSchSpot("D", hmSdf, tlSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf, llExclEndTime, lmSsfRecPos, hmSxf, hmGsf, hmGhf)
                                    Else
                                        ilBookOk = False
                                    End If
                                End If
                            Else
                                ilBookOk = False
                            End If
                        End If
                        'BookSpot Re-Read Ssf so handle is correct
                        If ilBookOk Then
                            ilRet = gBookSpot(slSchStatus, hmSdf, tmSdf, llSdfRecPos, ilBkQH, hmSsf, tmSsf, lmSsfRecPos, ilAvEvt, ilPosition, tmChf, tmIClf, tmRdf, imVpfIndex, hmSmf, tmSmf, hmClf, hmCrf, imPriceLevel, False, hmSxf, hmGsf)
                            If ilRet Then
                                igSpotFillReturn = 1
                                'mMakeTracer llSdfRecPos, "S"
                                ilRet = gMakeTracer(hmSdf, tmSdf, llSdfRecPos, hmStf, lmLastLogDate, "S", "M", tmSdf.iRotNo, hmGsf)
                                If ilRet Then
                                    mBookSpot = 1
                                    'tgCntSpot(ilCntSpotIndex).iNoTimesUsed = tgCntSpot(ilCntSpotIndex).iNoTimesUsed + 1
                                    ilRet = btrEndTrans(hmSdf)
                                    ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                    Exit Function
                                End If
                            End If
                        End If
                        mBookSpot = -1
                        ilCRet = btrAbortTrans(hmSdf)
                        ilCRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                        'ilRet = MsgBox("Task could not be completed successfully because of " & Str$(ilRet) & ", Redo Task", vbOkOnly + vbExclamation, "Spot")
                    End If
                    ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                End If
            End If
        Next ilGsf
    Next llDate
    mBookSpot = 0
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCntrPop                        *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mCntrPop(llSDate As Long, llEDate As Long, llSTime As Long, llETime As Long)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVehGame                                                                             *
'******************************************************************************************

'
'   mCntrPop
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim ilFound As Integer
    Dim slDate As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llDate As Long
    Dim llTDate As Long
    Dim ilLoop As Integer
    Dim ilClf As Integer
    Dim ilVef As Integer
    Dim ilRdf As Integer
    Dim slStr As String
    Dim slCntrStatus As String
    Dim ilHOType As Integer
    Dim ilUpper As Integer
    Dim slCntrNo As String
    Dim ilDay As Integer
    Dim ilSdf As Integer
    Dim ilTest As Integer
    Dim slCntrType As String
    Dim ilMnfDemo As Integer
    Dim ilMnfSocEco As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim ilVpfIndex As Integer
    Dim ilPrevLkIndex As Integer
    Dim ilPassDnfCode As Integer
    Dim ilPassVefCode As Integer
    Dim ilMkt As Integer
    Dim llAllowedSTime As Long
    Dim llAllowedETime As Long
    Dim ilAllowedTimeIndex As Integer
    ReDim ilAllowedDays(0 To 6) As Integer
    ReDim tgCntSpot(0 To 0) As CNTSPOT
    ReDim tgSpotLinks(0 To 0) As SPOTLINKS
    Dim llPopEst As Long
    Dim slAdvtName As String
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long

    slCntrStatus = "HO"
    slCntrType = "C"
    ilHOType = 1
    slStartDate = Format$(llSDate, "m/d/yy")
    slEndDate = Format$(llEDate, "m/d/yy")
    sgCntrForDateStamp = ""
    ilRet = gObtainCntrForDate(SpotMG, slStartDate, slEndDate, slCntrStatus, slCntrType, ilHOType, tmChfAdvtExt())
    If (ilRet <> CP_MSG_NOPOPREQ) And (ilRet <> CP_MSG_NONE) Then
        Exit Sub
    End If
    For ilLoop = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
        ilRet = gObtainChfClf(hmCHF, hmClf, tmChfAdvtExt(ilLoop).lCode, False, tmChf, tgClfSpot())
        If ilRet Then
            For ilClf = LBound(tgClfSpot) To UBound(tgClfSpot) - 1 Step 1
                tmIClf = tgClfSpot(ilClf).ClfRec
                ReDim lmSdfRecPos(0 To 0) As Long
                ilFound = True
                If (tgSpf.sMktBase = "Y") Then
                    ilFound = False
                    ilVef = gBinarySearchVef(tmIClf.iVefCode)
                    If ilVef <> -1 Then
                        For ilMkt = 0 To UBound(igSpotMktCode) - 1 Step 1
                            If tgMVef(ilVef).iMnfVehGp3Mkt = igSpotMktCode(ilMkt) Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilMkt
                    End If
                End If
                ilVef = gBinarySearchVef(tmIClf.iVefCode)
                If ilVef <> -1 Then
                    If imVehGame Then
                        If tgMVef(ilVef).sType <> "G" Then
                            ilFound = False
                        End If
                    Else
                        If tgMVef(ilVef).sType = "G" Then
                            ilFound = False
                        End If
                    End If
                End If
                'Bypass Partial buys at this time.  Reduce the work of adding code to scheduling the missed
                If tmIClf.lRafCode > 0 Then
                    ilFound = False
                End If
                If gIsImportInvoicedSpots(tmIClf.iVefCode) Then
                    ilFound = False
                End If
                If ilFound Then
                    'Test if any missed spots
                    For llDate = llSDate To llEDate Step 1
                        tmSdfSrchKey0.iVefCode = tmIClf.iVefCode
                        tmSdfSrchKey0.lChfCode = tmChf.lCode
                        tmSdfSrchKey0.iLineNo = tmIClf.iLine
                        tmSdfSrchKey0.lFsfCode = 0
                        gPackDate Format$(llDate, "m/d/yy"), tmSdfSrchKey0.iDate(0), tmSdfSrchKey0.iDate(1)
                        tmSdfSrchKey0.sSchStatus = "M"
                        tmSdfSrchKey0.iTime(0) = 0
                        tmSdfSrchKey0.iTime(1) = 0
                        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
                        'This code added as replacement for Ext operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = tmIClf.iVefCode) And (tmSdf.lChfCode = tmChf.lCode) And (tmSdf.iLineNo = tmIClf.iLine) And (tmSdf.sSchStatus = "M")
                            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llTDate
                            If (llTDate <> llDate) Then
                                Exit Do
                            End If
                            ilRet = btrGetPosition(hmSdf, lmSdfRecPos(UBound(lmSdfRecPos)))
                            ReDim Preserve lmSdfRecPos(0 To UBound(lmSdfRecPos) + 1) As Long
                            ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    Next llDate
                End If
                If UBound(lmSdfRecPos) > LBound(lmSdfRecPos) Then
                    'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                    '    If tmIClf.iRdfcode = tgMRdf(ilRdf).iCode Then
                        ilRdf = gBinarySearchRdf(tmIClf.iRdfCode)
                        If ilRdf <> -1 Then
                            tmRdf = tgMRdf(ilRdf)
                    '        Exit For
                        End If
                    'Next ilRdf
                    ilFound = False
                    tmCffSrchKey.lChfCode = tmChf.lCode
                    tmCffSrchKey.iClfLine = tmIClf.iLine
                    tmCffSrchKey.iCntRevNo = tmIClf.iCntRevNo
                    tmCffSrchKey.iPropVer = tmIClf.iPropVer
                    tmCffSrchKey.iStartDate(0) = 0
                    tmCffSrchKey.iStartDate(1) = 0
                    imCffRecLen = Len(tmICff(0))
                    ilRet = btrGetGreaterOrEqual(hmCff, tmICff(0), imCffRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmICff(0).lChfCode = tmChf.lCode) And (tmICff(0).iClfLine = tmIClf.iLine)
                        If (tmICff(0).iCntRevNo = tmIClf.iCntRevNo) And (tmICff(0).iPropVer = tmIClf.iPropVer) Then 'And (tmCff(2).sDelete <> "Y") Then
                            gUnpackDateLong tmICff(0).iStartDate(0), tmICff(0).iStartDate(1), llStartDate    'Week Start date
                            gUnpackDateLong tmICff(0).iEndDate(0), tmICff(0).iEndDate(1), llEndDate    'Week Start date
                            If (llEDate >= llStartDate) And (llSDate <= llEndDate) Then
                                ilFound = True
                                Exit Do
                            End If
                        End If
                        ilRet = btrGetNext(hmCff, tmICff(0), imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If ilFound Then
                        If tmICff(0).sDyWk = "D" Then
                            For ilDay = 0 To 6 Step 1
                                If tmICff(0).iDay(ilDay) > 0 Then
                                    ilAllowedDays(ilDay) = True
                                Else
                                    ilAllowedDays(ilDay) = False
                                End If
                            Next ilDay
                        Else
                            For ilDay = 0 To 6 Step 1
                                If (tmICff(0).iDay(ilDay) > 0) Or (tmICff(0).sXDay(ilDay) = "Y") Then
                                    ilAllowedDays(ilDay) = True
                                Else
                                    ilAllowedDays(ilDay) = False
                                End If
                            Next ilDay
                        End If
                        ilRet = False
                        For ilDay = 0 To 6 Step 1
                            'If ckcDay(ilDay).Value And ilAllowedDays(ilDay) Then
                            If ilAllowedDays(ilDay) Then
                                ilRet = True
                                Exit For
                            End If
                        Next ilDay
                    Else
                        ilRet = False
                    End If
                    If (ilRet) Then
                        If tmChf.iAdfCode <> tmAdf.iCode Then
                            tmAdfSrchKey.iCode = tmChf.iAdfCode
                            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        End If
                        'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        '    If tmIClf.iVefCode = tgMVef(ilVef).iCode Then
                            ilVef = gBinarySearchVef(tmIClf.iVefCode)
                            If ilVef <> -1 Then
                                ilVpfIndex = gBinarySearchVpf(tmIClf.iVefCode)    'gVpfFind(SpotMG, tmIClf.iVefCode)
                                slCntrNo = Trim$(str$(tmChf.lCntrNo))
                                Do While Len(slCntrNo) < 8
                                    slCntrNo = "0" & slCntrNo
                                Loop
                                ilUpper = UBound(tgCntSpot)
                                tgCntSpot(ilUpper).sType = tmChf.sType
                                tgCntSpot(ilUpper).iAdfCode = tmChf.iAdfCode
                                tgCntSpot(ilUpper).lChfCode = tmChf.lCode
                                tgCntSpot(ilUpper).iVefCode = tmIClf.iVefCode
                                tgCntSpot(ilUpper).iLnVefCode = tmIClf.iVefCode
                                tgCntSpot(ilUpper).iLineNo = tmIClf.iLine
                                tgCntSpot(ilUpper).iNoSSpots = 0
                                tgCntSpot(ilUpper).iNoGSpots = 0
                                tgCntSpot(ilUpper).iNoMSpots = 1
                                tgCntSpot(ilUpper).iNoESpots = 0
                                tgCntSpot(ilUpper).lSdfRecPos = 0
                                tgCntSpot(ilUpper).iRdfCode = tmIClf.iRdfCode
                                tgCntSpot(ilUpper).iSpotLkIndex = UBound(tgSpotLinks)
                                tgSpotLinks(UBound(tgSpotLinks)).iStatus = 0
                                tgSpotLinks(UBound(tgSpotLinks)).lSdfRecPos = lmSdfRecPos(0)
                                tgSpotLinks(UBound(tgSpotLinks)).iSpotLkIndex = -1
                                ilPrevLkIndex = UBound(tgSpotLinks)
                                ReDim Preserve tgSpotLinks(0 To UBound(tgSpotLinks) + 1) As SPOTLINKS
                                For ilSdf = 1 To UBound(lmSdfRecPos) - 1 Step 1
                                    tgCntSpot(ilUpper).iNoMSpots = tgCntSpot(ilUpper).iNoMSpots + 1
                                    tgSpotLinks(ilPrevLkIndex).iSpotLkIndex = UBound(tgSpotLinks)
                                    tgSpotLinks(UBound(tgSpotLinks)).iStatus = 0
                                    tgSpotLinks(UBound(tgSpotLinks)).lSdfRecPos = lmSdfRecPos(ilSdf)
                                    tgSpotLinks(UBound(tgSpotLinks)).iSpotLkIndex = -1
                                    ilPrevLkIndex = UBound(tgSpotLinks)
                                    ReDim Preserve tgSpotLinks(0 To UBound(tgSpotLinks) + 1) As SPOTLINKS
                                Next ilSdf
                                If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                                    slAdvtName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
                                Else
                                    slAdvtName = Trim$(tmAdf.sName)
                                End If
                                If rbcSort(0).Value Then
                                    tgCntSpot(ilUpper).sKey = "1" & "|" & slAdvtName & "|" & tgMVef(ilVef).sName
                                Else
                                    tgCntSpot(ilUpper).sKey = "2" & "|" & tgMVef(ilVef).sName & "|" & slAdvtName
                                End If
                                tgCntSpot(ilUpper).sLen = Trim$(str$(tmIClf.iLen))
                                tgCntSpot(ilUpper).sProduct = tmChf.sProduct
                                gUnpackDate tmIClf.iStartDate(0), tmIClf.iStartDate(1), slDate
                                slStr = slDate
                                gUnpackDate tmIClf.iEndDate(0), tmIClf.iEndDate(1), slDate
                                slStr = slStr & "-" & slDate
                                tgCntSpot(ilUpper).sDate = slStr
                                llOvStartTime = 0
                                llOvEndTime = 0
                                For ilTest = LBound(tgCntSpot(ilUpper).lAllowedSTime) To UBound(tgCntSpot(ilUpper).lAllowedSTime) Step 1
                                    tgCntSpot(ilUpper).lAllowedSTime(ilTest) = -1
                                    tgCntSpot(ilUpper).lAllowedETime(ilTest) = -1
                                Next ilTest
                                ilAllowedTimeIndex = LBound(tgCntSpot(ilUpper).lAllowedSTime)
                                If (tmRdf.iLtfCode(0) <> 0) Or (tmRdf.iLtfCode(1) <> 0) Or (tmRdf.iLtfCode(2) <> 0) Then
                                Else
                                    If ((tmIClf.iStartTime(0) = 1) And (tmIClf.iStartTime(1) = 0)) Or (tgVpf(ilVpfIndex).sGMedium = "S") Then
                                        For ilTest = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1
                                            If (tmRdf.iStartTime(0, ilTest) <> 1) Or (tmRdf.iStartTime(1, ilTest) <> 0) Then
                                                gUnpackTimeLong tmRdf.iStartTime(0, ilTest), tmRdf.iStartTime(1, ilTest), False, llAllowedSTime
                                                gUnpackTimeLong tmRdf.iEndTime(0, ilTest), tmRdf.iEndTime(1, ilTest), True, llAllowedETime
                                                mChkXMid llSTime, llETime, ilAllowedTimeIndex, llAllowedSTime, llAllowedETime
                                            End If
                                        Next ilTest
                                    Else
                                        gUnpackTimeLong tmIClf.iStartTime(0), tmIClf.iStartTime(1), False, llAllowedSTime
                                        llOvStartTime = llAllowedSTime
                                        gUnpackTimeLong tmIClf.iEndTime(0), tmIClf.iEndTime(1), True, llAllowedETime
                                        llOvEndTime = llAllowedETime
                                        mChkXMid llSTime, llETime, ilAllowedTimeIndex, llAllowedSTime, llAllowedETime
                                    End If
                                End If
                                If ilAllowedTimeIndex > LBound(tgCntSpot(ilUpper).lAllowedSTime) Then
                                    tgCntSpot(ilUpper).iNoTimesUsed = 0
                                    tgCntSpot(ilUpper).iMnfComp0 = tmChf.iMnfComp(0)
                                    tgCntSpot(ilUpper).iMnfComp1 = tmChf.iMnfComp(1)
                                    For ilDay = 0 To 6 Step 1
                                        tgCntSpot(ilUpper).iAllowedDays(ilDay) = ilAllowedDays(ilDay)
                                    Next ilDay
                                    tgCntSpot(ilUpper).lPrice = tmICff(0).lActPrice
                                    ilMnfDemo = tmChf.iMnfDemo(0)
                                    ilMnfSocEco = 0
                                    ilPassDnfCode = tgMVef(ilVef).iDnfCode
                                    ilPassVefCode = tgMVef(ilVef).iCode
                                    If tgSpf.sDemoEstAllowed = "Y" Then
                                        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, lmSdfRecPos(0), INDEXKEY0, BTRV_LOCK_NONE)
                                        If ilRet = BTRV_ERR_NONE Then
                                            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                                        Else
                                            llDate = llSDate
                                        End If
                                    Else
                                        llDate = llSDate
                                    End If
                                    ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilPassDnfCode, ilPassVefCode, ilMnfSocEco, ilMnfDemo, llDate, llDate, tmIClf.iRdfCode, llOvStartTime, llOvEndTime, ilAllowedDays(), tmIClf.sType, tmIClf.lRafCode, tgCntSpot(ilUpper).lAud, llPopEst, ilAudFromSource, llAudFromCode)
                                    If ((tmICff(0).sPriceType = "T") And (tmICff(0).lActPrice > 0) And (ckcSpotType(0).Value = vbChecked)) Or ((tmICff(0).sPriceType = "T") And (tmICff(0).lActPrice = 0) And (ckcSpotType(1).Value = vbChecked)) Or ((tmICff(0).sPriceType <> "T") And (ckcSpotType(1).Value = vbChecked)) Then
                                        ReDim Preserve tgCntSpot(0 To UBound(tgCntSpot) + 1) As CNTSPOT
                                    End If
                                End If
                        '        Exit For
                            End If
                        'Next ilVef
                    End If
                End If
            Next ilClf
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mComputeReq                     *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute the values for the Req *
'*                      column                         *
'*                                                     *
'*******************************************************
Private Sub mComputeReq()
    Dim ilLoop As Integer
    Dim llReq As Long
    Dim slStr As String
    Dim slvalue As String
    Dim llTotalSpot As Long
    Dim llTotalVehDP As Long
    Dim llTotal As Long
    Dim ilRet As Integer
    Dim slStr1 As String
    Dim slStr2 As String
    Dim slStr3 As String
    Dim slStr4 As String
    Dim slCode As String
    If lacVehDPWk(1).Caption <> "" Then
        llTotalVehDP = Val(lacVehDPWk(1).Caption)
    Else
        llTotalVehDP = 0
    End If
    If lacSpotWk(1).Caption <> "" Then
        llTotalSpot = Val(lacSpotWk(1).Caption)
    Else
        llTotalSpot = 0
    End If
    llTotal = llTotalSpot - llTotalVehDP
    For ilLoop = 0 To lbcVehDP(0).ListCount - 1 Step 1
        slStr = lbcVehDP(0).List(ilLoop)
        ilRet = gParseItem(slStr, 4, "|", slStr)
        slStr = Trim$(slStr)
        If (Val(slStr) > 0) And (llTotalSpot > 0) Then
            If llTotal > 0 Then
                llReq = (10 * llTotal) / Val(slStr)
                slvalue = gLongToStrDec(llReq, 1)
            Else
                slvalue = "0"
            End If
        Else
            slvalue = ""
        End If
        slStr = lbcVehDP(0).List(ilLoop)
        ilRet = gParseItem(slStr, 1, "|", slStr1)
        ilRet = gParseItem(slStr, 2, "|", slStr2)
        ilRet = gParseItem(slStr, 3, "|", slStr3)
        ilRet = gParseItem(slStr, 4, "|", slStr4)
        ilRet = gParseItem(slStr, 6, "|", slCode)
        'lbcVehDP(0).List(ilLoop) = gAlignStringByPixel(slStr1 & "|" & slStr2 & "|" & slStr3 & "|" & slStr4 & "|" & slValue & "|" & slCode, "|", imListFieldVehDp(), imListFieldVehDpChar())
        lbcVehDP(0).List(ilLoop) = slStr1 & "|" & slStr2 & "|" & slStr3 & "|" & slStr4 & "|" & slvalue & "|" & slCode
    Next ilLoop
    pbcLbcVehDp_Paint 0
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mHbcMissedWkChange              *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Advt missed date change        *
'*                                                     *
'*******************************************************
Private Sub mHbcMissedWkChange()
    Dim llDate As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llSDate As Long
    Dim llEDate As Long
    Dim slLine As String
    Dim ilLoop As Integer
    Dim slName As String
    Dim slDPName As String
    Dim ilRet As Integer
    Dim ilRdf As Integer
    Dim slPrice As String
    imBlankDate = False
    lbcMissed.Clear
    llDate = 7 * (hbcMissedWk.Value - 1) + lmEarliestMonDate
    slStartDate = Format(llDate, "m/d/yy")
    slEndDate = gObtainNextSunday(slStartDate)
    llSDate = gDateValue(slStartDate)
    llEDate = gDateValue(slEndDate)
    mCntrPop llSDate, llEDate, 0, 90000
    If UBound(tgCntSpot) - 1 > 0 Then
        ArraySortTyp fnAV(tgCntSpot(), 0), UBound(tgCntSpot), 0, LenB(tgCntSpot(0)), 0, LenB(tgCntSpot(0).sKey), 0
    End If
    For ilLoop = LBound(tgCntSpot) To UBound(tgCntSpot) - 1 Step 1
        slLine = tgCntSpot(ilLoop).sKey
        slPrice = gLongToStrDec(tgCntSpot(ilLoop).lPrice, 2)
        'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
        '    If tgCntSpot(ilLoop).iRdfcode = tgMRdf(ilRdf).iCode Then
            ilRdf = gBinarySearchRdf(tgCntSpot(ilLoop).iRdfCode)
            If ilRdf <> -1 Then
                slDPName = Trim$(tgMRdf(ilRdf).sName)
        '        Exit For
            End If
        'Next ilRdf
        If rbcSort(0).Value Then
            ilRet = gParseItem(slLine, 2, "|", slName)
            'lbcMissed.AddItem gAlignStringByPixel(Trim$(slName) & "|" & Trim$(tgCntSpot(ilLoop).sProduct) & "|" & Trim$(tgCntSpot(ilLoop).sLen) & "|" & slPrice & "|" & " " & "|" & "-1", "|", imListField(), imListFieldChar())
            lbcMissed.AddItem Trim$(slName) & "|" & Trim$(tgCntSpot(ilLoop).sProduct) & "|" & Trim$(tgCntSpot(ilLoop).sLen) & "|" & slPrice & "|" & " " & "|" & "-1"
            ilRet = gParseItem(slLine, 3, "|", slName)
            If tgCntSpot(ilLoop).lAud > 0 Then
                'lbcMissed.AddItem gAlignStringByPixel("  " & Trim$(slName) & "|" & slDPName & "| " & Str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & Trim$(Str$(tgCntSpot(ilLoop).lAud)) & "|" & " " & "|" & Str$(ilLoop), "|", imListField(), imListFieldChar())
                lbcMissed.AddItem "  " & Trim$(slName) & "|" & slDPName & "| " & str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & Trim$(str$(tgCntSpot(ilLoop).lAud)) & "|" & " " & "|" & str$(ilLoop)
            Else
                'lbcMissed.AddItem gAlignStringByPixel("  " & Trim$(slName) & "|" & slDPName & "| " & Str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & " " & "|" & " " & "|" & Str$(ilLoop), "|", imListField(), imListFieldChar())
                lbcMissed.AddItem "  " & Trim$(slName) & "|" & slDPName & "| " & str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & " " & "|" & " " & "|" & str$(ilLoop)
            End If
        Else
            ilRet = gParseItem(slLine, 2, "|", slName)
            If tgCntSpot(ilLoop).lAud > 0 Then
                'lbcMissed.AddItem gAlignStringByPixel(Trim$(slName) & "|" & slDPName & "| " & Str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & Trim$(Str$(tgCntSpot(ilLoop).lAud)) & "|" & " " & "|" & Str$(ilLoop), "|", imListField(), imListFieldChar())
                lbcMissed.AddItem Trim$(slName) & "|" & slDPName & "| " & str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & Trim$(str$(tgCntSpot(ilLoop).lAud)) & "|" & " " & "|" & str$(ilLoop)
            Else
                'lbcMissed.AddItem gAlignStringByPixel(Trim$(slName) & "|" & slDPName & "| " & Str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & " " & "|" & " " & "|" & Str$(ilLoop), "|", imListField(), imListFieldChar())
                lbcMissed.AddItem Trim$(slName) & "|" & slDPName & "| " & str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & " " & "|" & " " & "|" & str$(ilLoop)
            End If
            ilRet = gParseItem(slLine, 3, "|", slName)
            'lbcMissed.AddItem gAlignStringByPixel("  " & Trim$(slName) & "|" & Trim$(tgCntSpot(ilLoop).sProduct) & "|" & Trim$(tgCntSpot(ilLoop).sLen) & "|" & slPrice & "|" & " " & "|" & "-1", "|", imListField(), imListFieldChar())
            lbcMissed.AddItem "  " & Trim$(slName) & "|" & Trim$(tgCntSpot(ilLoop).sProduct) & "|" & Trim$(tgCntSpot(ilLoop).sLen) & "|" & slPrice & "|" & " " & "|" & "-1"
        End If
    Next ilLoop
    pbcLbcMissed_Paint
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mHbcSpotWkChange                *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Advt missed date change        *
'*                                                     *
'*******************************************************
Private Sub mHbcSpotWkChange()
    Dim llDate As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slNameCode As String
    'Dim ilLoop As Integer
    Dim llLoop As Long
    Dim slVehName As String
    Dim slDate As String
    Dim slTime As String
    'Dim ilIndex As Integer
    Dim llIndex As Long
    Dim llAvgAud As Long
    Dim ilRet As Integer
    Dim ilDnfCode As Integer
    Dim ilMnfDemo As Integer
    Dim ilMnfSocEco As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim ilClf As Integer
    Dim ilCff As Integer
    Dim llFSDate As Long
    Dim llFEDate As Long
    Dim ilDay As Integer
    Dim ilVef As Integer
    ReDim ilInputDay(0 To 6) As Integer
    Dim llPopEst As Long
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long

    imBlankDate = False
    lbcSpots(0).Clear
    llDate = 7 * (hbcSpotWk.Value - 1) + lmEarliestMonDate
    'If llDate < lgMGAllowDate Then
    '    llDate = lgMGAllowDate
    'End If
    slStartDate = Format(llDate, "m/d/yy")
    'slStartDate = gFormatDate(slStartDate)
    'plcSpotWk.Caption = slStartDate
    slEndDate = gObtainNextSunday(slStartDate)
    ilRet = gObtainCntrSpot(-1, False, lgChfMGCode, -1, "S", slStartDate, slEndDate, tmSdfExtSort(), tmSdfExt(), 0, False)
    'For ilLoop = 0 To UBound(tmSdfExtSort) - 1 Step 1
    For llLoop = 0 To UBound(tmSdfExtSort) - 1 Step 1
        slNameCode = tmSdfExtSort(llLoop).sKey   'Line#|Vehicle Name|Date|Time
        ilRet = gParseItem(slNameCode, 2, "|", slVehName)
        'ilIndex = tmSdfExtSort(ilLoop).iSdfExtIndex 'Val(slCode)
        llIndex = tmSdfExtSort(llLoop).lSdfExtIndex 'Val(slCode)
        gUnpackDate tmSdfExt(llIndex).iDate(0), tmSdfExt(llIndex).iDate(1), slDate
        llDate = gDateValue(slDate)
        gUnpackTime tmSdfExt(llIndex).iTime(0), tmSdfExt(llIndex).iTime(1), "A", "1", slTime
        If ((tmSdfExt(llIndex).sSchStatus = "S") And (llDate >= lgMGAllowDate)) Or (tmSdfExt(llIndex).sSchStatus = "M") Or (tmSdfExt(llIndex).sSchStatus = "U") Then
            If tmSdfExt(llIndex).iLen = igMGSpotLen Then
                'ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPosSrce, INDEXKEY0, BTRV_LOCK_NONE)
                ilMnfDemo = tmChf.iMnfDemo(0)
                ilMnfSocEco = 0
                llAvgAud = 0
                If ilMnfDemo > 0 Then
                    For ilClf = LBound(tmClf) To UBound(tmClf) - 1 Step 1
                        If tmClf(ilClf).ClfRec.iLine = tmSdfExt(llIndex).iLineNo Then
                            ilDnfCode = tmClf(ilClf).ClfRec.iDnfCode
                            If ilDnfCode <= 0 Then  'Get from vehicle
                                'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                '    If tgMVef(ilVef).iCode = tmSdfExt(ilIndex).iVefCode Then
                                    ilVef = gBinarySearchVef(tmSdfExt(llIndex).iVefCode)
                                    If ilVef <> -1 Then
                                        ilDnfCode = tgMVef(ilVef).iDnfCode
                                '        Exit For
                                    End If
                                'Next ilVef
                                'tmVefSrchKey.iCode = tmSdfExt(ilIndex).iVefCode
                                'ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                'ilDnfCode = tmVef.iDnfCode
                            End If
                            If ((tmClf(ilClf).ClfRec.iStartTime(0) <> 1) Or (tmClf(ilClf).ClfRec.iStartTime(1) <> 0)) And ((tmClf(ilClf).ClfRec.iEndTime(0) <> 1) Or (tmClf(ilClf).ClfRec.iEndTime(1) <> 0)) Then
                                gUnpackTimeLong tmClf(ilClf).ClfRec.iStartTime(0), tmClf(ilClf).ClfRec.iStartTime(1), False, llOvStartTime
                                gUnpackTimeLong tmClf(ilClf).ClfRec.iEndTime(0), tmClf(ilClf).ClfRec.iEndTime(1), True, llOvEndTime
                            Else
                                llOvStartTime = 0
                                llOvEndTime = 0
                            End If
                            For ilCff = LBound(tmCff) To UBound(tmCff) - 1 Step 1
                                If tmCff(ilCff).CffRec.iClfLine = tmClf(ilClf).ClfRec.iLine Then
                                    gUnpackDateLong tmCff(ilCff).CffRec.iStartDate(0), tmCff(ilCff).CffRec.iStartDate(1), llFSDate
                                    gUnpackDateLong tmCff(ilCff).CffRec.iEndDate(0), tmCff(ilCff).CffRec.iEndDate(1), llFEDate
                                    If (llDate >= llFSDate) And (llDate <= llFEDate) Then
                                        For ilDay = 0 To 6 Step 1
                                            If tmCff(ilCff).CffRec.iDay(ilDay) > 0 Then
                                                ilInputDay(ilDay) = True
                                            Else
                                                ilInputDay(ilDay) = False
                                            End If
                                        Next ilDay
                                        ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, tmClf(ilClf).ClfRec.iVefCode, ilMnfSocEco, ilMnfDemo, llDate, llDate, tmClf(ilClf).ClfRec.iRdfCode, llOvStartTime, llOvEndTime, ilInputDay(), tmClf(ilClf).ClfRec.sType, tmClf(ilClf).ClfRec.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                                        Exit For
                                    End If
                                End If
                            Next ilCff
                            Exit For
                        End If
                    Next ilClf
                End If
                'lbcSpots(0).AddItem gAlignStringByPixel(slVehName & "|" & slDate & "|" & slTime & "|" & Trim$(Str$(llAvgAud)) & "|" & Str$(hbcSpotWk.Value) & "|" & Str$(tmSdfExt(ilIndex).lRecPos), "|", imListFieldSpots(), imListFieldSpotsChar())
                lbcSpots(0).AddItem slVehName & "|" & slDate & "|" & slTime & "|" & Trim$(str$(llAvgAud)) & "|" & str$(hbcSpotWk.Value) & "|" & str$(tmSdfExt(llIndex).lRecPos)
            End If
        End If
    Next llLoop
    pbcLbcSpots_Paint 0
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mHbcVehDPWkChange               *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Advt missed date change        *
'*                                                     *
'*******************************************************
Private Sub mHbcVehDPWkChange()
    Dim llDate As Long
    Dim slDate As String
    Dim ilRcf As Integer
    Dim llRcfDate As Long
    Dim llTestDate As Long
    Dim ilRcfIndex As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilVef As Integer
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim ilMnfDemo As Integer
    Dim ilMnfSocEco As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim ilDay As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slRdfStr As String
    Dim slStr As String
    Dim ilTime As Integer
    Dim ilRow As Integer
    Dim llSec As Long
    Dim llMin As Long
    Dim llHour As Long
    Dim llTime As Long
    Dim slTime As String
    Dim ilYear As Integer
    Dim slPrevVeh As String
    Dim ilPassDnfCode As Integer
    Dim ilPassVefCode As Integer
    Dim ilPassRdfCode As Integer
    ReDim ilInputDay(0 To 6) As Integer
    Dim llPopEst As Long
    Dim ilMkt As Integer
    Dim llRafCode As Long
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long

    imBlankDate = False
    lbcVehDP(0).Clear
    lbcAired.Clear
    llDate = 7 * (hbcVehDPWk.Value - 1) + lmEarliestMonDate
    slDate = Format(llDate, "m/d/yy")
    llStartDate = llDate
    llEndDate = llDate + 6
    slStr = Format$(llEndDate, "m/d/yy")
    ilYear = Year(gAdjYear(slStr))
    'slDate = gFormatDate(slDate)
    'plcVehDPWk.Caption = slDate
    'Determine rate card to use
    llOvStartTime = 0
    llOvEndTime = 0
    ilMnfDemo = tmChf.iMnfDemo(0)
    ilMnfSocEco = 0
    ilRcfIndex = -1
    For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
        gUnpackDateLong tgMRcf(ilRcf).iStartDate(0), tgMRcf(ilRcf).iStartDate(1), llTestDate
        If ilRcfIndex = -1 Then
            If llTestDate <= llDate Then
                ilRcfIndex = ilRcf
                llRcfDate = llTestDate
            End If
        Else
            If (llTestDate <= llDate) And (llTestDate > llRcfDate) Then
                ilRcfIndex = ilRcf
                llRcfDate = llTestDate
            End If
        End If
    Next ilRcf
    'Find base dayparts
    If ilRcfIndex >= 0 Then
        If imRcfIndex <> ilRcfIndex Then
            For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                If (tgMRcf(ilRcfIndex).iCode = tgMRif(llRif).iRcfCode) And (tgMRif(llRif).iYear = ilYear) Then
                    'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                    '    If tgMRif(llRif).iRdfcode = tgMRdf(ilRdf).iCode Then
                        ilFound = True
                        If (tgSpf.sMktBase = "Y") Then
                            ilFound = False
                            ilVef = gBinarySearchVef(tgMRif(llRif).iVefCode)
                            If ilVef <> -1 Then
                                For ilMkt = 0 To UBound(igSpotMktCode) - 1 Step 1
                                    If tgMVef(ilVef).iMnfVehGp3Mkt = igSpotMktCode(ilMkt) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilMkt
                            End If
                        End If
                        If ilFound Then
                            If gIsImportInvoicedSpots(tgMRif(llRif).iVefCode) Then
                                ilFound = False
                            End If
                        End If
                        ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfCode)
                        If (ilRdf <> -1) And (ilFound) Then
                            If tgMRdf(ilRdf).sBase = "Y" Then
                                slRdfStr = Trim$(str$(tgMRdf(ilRdf).iSortCode))
                                Do While Len(slRdfStr) < 3
                                    slRdfStr = "0" & slRdfStr
                                Loop
                                ilTime = LBound(tgMRdf(ilRdf).iStartTime, 2)    '1
                                For ilRow = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                    If (tgMRdf(ilRdf).iStartTime(0, ilRow) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilRow) <> 0) Then
                                        ilTime = ilRow
                                        For ilIndex = 1 To 7 Step 1
                                            If tgMRdf(ilRdf).sWkDays(ilTime, ilIndex - 1) <> "Y" Then
                                                slRdfStr = slRdfStr & "B"
                                            Else
                                                slRdfStr = slRdfStr & "A"
                                            End If
                                        Next ilIndex
                                        Exit For
                                    End If
                                Next ilRow
                                llSec = tgMRdf(ilRdf).iStartTime(0, ilTime) \ 256 'Obtain seconds
                                llMin = tgMRdf(ilRdf).iStartTime(1, ilTime) And &HFF 'Obtain Minutes
                                llHour = tgMRdf(ilRdf).iStartTime(1, ilTime) \ 256 'Obtain month
                                llTime = 3600 * llHour + 60 * llMin + llSec
                                slTime = Trim$(str$(llTime))
                                Do While (Len(slTime) < 5)
                                    slTime = "0" & slTime
                                Loop
                                slRdfStr = slRdfStr & slTime
                                'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                '    If tgMRif(llRif).iVefCode = tgMVef(ilVef).iCode Then
                                    ilVef = gBinarySearchVef(tgMRif(llRif).iVefCode)
                                    If ilVef <> -1 Then
                                        If ((imVehGame) And (tgMVef(ilVef).sType = "G")) Or ((Not imVehGame) And (tgMVef(ilVef).sType <> "G")) Then
                                            tgVefRdfInfo(ilUpper).sKey = tgMVef(ilVef).sName & "|" & slRdfStr & "|" & tgMRdf(ilRdf).sName
                                            tgVefRdfInfo(ilUpper).iVefIndex = ilVef
                                            tgVefRdfInfo(ilUpper).iRdfIndex = ilRdf
                                            tgVefRdfInfo(ilUpper).iDnfCode = tgMVef(ilVef).iDnfCode
                                            tgVefRdfInfo(ilUpper).lAvgAud = 0
                                            For ilDay = 1 To 7 Step 1
                                                ilInputDay(ilDay - 1) = False
                                            Next ilDay
                                            For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                                If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                                    For ilDay = 1 To 7 Step 1
                                                        'If (tgMRdf(ilRdf).sWkDays(ilIndex, ilDay) = "Y") Then
                                                        If (tgMRdf(ilRdf).sWkDays(ilIndex, ilDay - 1) = "Y") Then
                                                            ilInputDay(ilDay - 1) = True
                                                        End If
                                                    Next ilDay
                                                End If
                                            Next ilIndex
                                            ilPassDnfCode = tgMVef(ilVef).iDnfCode
                                            ilPassVefCode = tgMVef(ilVef).iCode
                                            ilPassRdfCode = tgMRdf(ilRdf).iCode
                                            llRafCode = 0
                                            ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilPassDnfCode, ilPassVefCode, ilMnfSocEco, ilMnfDemo, llDate, llDate, ilPassRdfCode, llOvStartTime, llOvEndTime, ilInputDay(), "S", llRafCode, tgVefRdfInfo(ilUpper).lAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                                            ilUpper = ilUpper + 1
                                            ReDim Preserve tgVefRdfInfo(0 To ilUpper) As VEFRDFINFO
                                        End If
                                '        Exit For
                                    End If
                                'Next ilVef
                            End If
                    '        Exit For
                        End If
                    'Next ilRdf
                End If
            Next llRif
            If UBound(tgVefRdfInfo) - 1 > 0 Then
                ArraySortTyp fnAV(tgVefRdfInfo(), 0), UBound(tgVefRdfInfo), 0, LenB(tgVefRdfInfo(0)), 0, LenB(tgVefRdfInfo(0).sKey), 0
            End If
        End If
        For ilLoop = 0 To ilUpper - 1 Step 1
            tgVefRdfInfo(ilLoop).iAvail = -9999
        Next ilLoop
        'Get avails
        ReDim ilVefCode(0 To 0) As Integer
        ilUpper = 0
        For ilLoop = 0 To UBound(tgVefRdfInfo) - 1 Step 1
            ilFound = False
            For ilVef = 0 To UBound(ilVefCode) - 1 Step 1
                If ilVefCode(ilVef) = tgMVef(tgVefRdfInfo(ilLoop).iVefIndex).iCode Then
                    ilFound = True
                    Exit For
                End If
            Next ilVef
            If Not ilFound Then
                ilVefCode(ilUpper) = tgMVef(tgVefRdfInfo(ilLoop).iVefIndex).iCode
                ilUpper = ilUpper + 1
                ReDim Preserve ilVefCode(0 To ilUpper) As Integer
            End If
        Next ilLoop
        For ilVef = 0 To UBound(ilVefCode) - 1 Step 1
            mAvailCount hmSsf, hmSdf, ilVefCode(ilVef), llStartDate, llEndDate, False
        Next ilVef
        Erase ilVefCode
        slPrevVeh = ""
        For ilLoop = 0 To UBound(tgVefRdfInfo) - 1 Step 1
            If tgVefRdfInfo(ilLoop).iAvail <> -9999 Then
                'lbcVehDP(0).AddItem gAlignStringByPixel(Trim$(tgMVef(tgVefRdfInfo(ilLoop).iVefIndex).sName) & "|" & Trim$(tgMRdf(tgVefRdfInfo(ilLoop).iRdfIndex).sName) & "|" & Trim$(Str$(tgVefRdfInfo(ilLoop).iAvail)) & "|" & Str$(tgVefRdfInfo(ilLoop).lAvgAud) & "|" & " " & "|" & Str$(ilLoop), "|", imListFieldVehDp(), imListFieldVehDpChar())
                lbcVehDP(0).AddItem Trim$(tgMVef(tgVefRdfInfo(ilLoop).iVefIndex).sName) & "|" & Trim$(tgMRdf(tgVefRdfInfo(ilLoop).iRdfIndex).sName) & "|" & Trim$(str$(tgVefRdfInfo(ilLoop).iAvail)) & "|" & str$(tgVefRdfInfo(ilLoop).lAvgAud) & "|" & " " & "|" & str$(ilLoop)
                If slPrevVeh <> Trim$(tgMVef(tgVefRdfInfo(ilLoop).iVefIndex).sName) Then
                    slStr = Trim$(tgMVef(tgVefRdfInfo(ilLoop).iVefIndex).sName)
                Else
                    slStr = " "
                End If
                slPrevVeh = Trim$(tgMVef(tgVefRdfInfo(ilLoop).iVefIndex).sName)
                If tgVefRdfInfo(ilLoop).lAvgAud > 0 Then
                    'lbcAired.AddItem gAlignStringByPixel(slStr & "|" & Trim$(tgMRdf(tgVefRdfInfo(ilLoop).iRdfIndex).sName) & "| " & Str$(tgVefRdfInfo(ilLoop).iAvail) & "|" & Str$(tgVefRdfInfo(ilLoop).lAvgAud) & "|" & " " & "|" & Str$(ilLoop), "|", imListField(), imListFieldChar())
                    lbcAired.AddItem slStr & "|" & Trim$(tgMRdf(tgVefRdfInfo(ilLoop).iRdfIndex).sName) & "| " & str$(tgVefRdfInfo(ilLoop).iAvail) & "|" & str$(tgVefRdfInfo(ilLoop).lAvgAud) & "|" & " " & "|" & str$(ilLoop)
                Else
                    'lbcAired.AddItem gAlignStringByPixel(slStr & "|" & Trim$(tgMRdf(tgVefRdfInfo(ilLoop).iRdfIndex).sName) & "| " & Str$(tgVefRdfInfo(ilLoop).iAvail) & "|" & " " & "|" & " " & "|" & Str$(ilLoop), "|", imListField(), imListFieldChar())
                    lbcAired.AddItem slStr & "|" & Trim$(tgMRdf(tgVefRdfInfo(ilLoop).iRdfIndex).sName) & "| " & str$(tgVefRdfInfo(ilLoop).iAvail) & "|" & " " & "|" & " " & "|" & str$(ilLoop)
                End If
            End If
        Next ilLoop
    End If
    mSetCommands
    pbcLbcVehDp_Paint 0
    pbcLbcAired_Paint
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
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilVpf As Integer
    Dim ilLen As Integer
    Dim ilFound As Integer
    Dim ilLenMin As Integer
    Dim ilLenMax As Integer
    Dim ilVef As Integer
    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    imBlankDate = False
    imFirstActivate = True
    sgCntrForDateStamp = ""
    imTerminate = False
    imIgnoreChg = False
    imMGGen = False
    imRcfIndex = -1
    imSelectIndex = -1
    imVehGame = False
    ilVef = gBinarySearchVef(igMGVefCode)
    If ilVef <> -1 Then
        If tgMVef(ilVef).sType = "G" Then
            imVehGame = True
        End If
    End If
    plcTabSelection.Move 60, 45
    plc1MG1.Move 165, 435
    plcMMGN.Move plc1MG1.Left, plc1MG1.Top
    'plcTabSelection(0).Move 120, 60
    'plcTabSelection(1).Move plcTabSelection(0).Left + plcTabSelection(0).Width + 30, plcTabSelection(0).Top
    'plc1MG1.Move 120, plcTabSelection(0).Top + plcTabSelection(0).Height - 30 ', plcTabSelection(1).Left + plcTabSelection(1).Width + 30 + plcTabSelection(1).Width - plcTabSelection(0).Left
    'plcMMGN.Move plc1MG1.Left, plc1MG1.Top ', plcRptName.Width
    'lncLine(0).X1 = plcTabSelection(0).Left - plc1MG1.Left + 45
    'lncEdge(0).X1 = lncLine(0).X2 '+ 15
    'lncEdge(0).X2 = lncLine(0).X2 '+ 15
    'lncEdge(0).Y1 = 0
    'lncEdge(0).Y2 = 30
    'lncEdge(1).X1 = lncLine(0).X2 + 15 '+ 30
    'lncEdge(1).X2 = lncLine(0).X2 + 15 '+ 30
    'lncEdge(1).Y1 = 0
    'lncEdge(1).Y2 = 15
    'lncLine(1).X1 = plcTabSelection(1).Left - plc1MG1.Left + 45
    'lncEdge(2).X1 = lncLine(1).X2 + 15
    'lncEdge(2).X2 = lncLine(1).X2 + 15
    'lncEdge(2).Y1 = 0
    'lncEdge(2).Y2 = 30
    'lncEdge(3).X1 = lncLine(1).X2 + 30
    'lncEdge(3).X2 = lncLine(1).X2 + 30
    'lncEdge(3).Y1 = 0
    'lncEdge(3).Y2 = 15
    imListField(1) = 15
    imListField(2) = 16 * igAlignCharWidth
    imListField(3) = 32 * igAlignCharWidth
    imListField(4) = 38 * igAlignCharWidth
    imListField(5) = 50 * igAlignCharWidth
    imListField(6) = 100 * igAlignCharWidth
    imListField(7) = 130 * igAlignCharWidth
    imListFieldSpots(1) = 15
    imListFieldSpots(2) = 12 * igAlignCharWidth
    imListFieldSpots(3) = 22 * igAlignCharWidth
    imListFieldSpots(4) = 33 * igAlignCharWidth
    imListFieldSpots(5) = 100 * igAlignCharWidth
    imListFieldSpots(6) = 130 * igAlignCharWidth
    imListFieldVehDp(1) = 15
    imListFieldVehDp(2) = 16 * igAlignCharWidth
    imListFieldVehDp(3) = 29 * igAlignCharWidth
    imListFieldVehDp(4) = 35 * igAlignCharWidth
    imListFieldVehDp(5) = 41 * igAlignCharWidth
    imListFieldVehDp(6) = 100 * igAlignCharWidth
    imListFieldVehDp(7) = 130 * igAlignCharWidth
    'SpotMG.Height = plc1MG1.Top + plc1MG1.Height + 90
    SpotMG.Height = plcTabSelection.Top + plcTabSelection.Height + 120
    'gCenterModalForm SpotMG
    gCenterModalForm SpotMG
    'ReDim tmSdfExt(1 To 1) As SDFEXT
    ReDim tmSdfExt(0 To 0) As SDFEXT
    ReDim tmSdfExtSort(0 To 0) As SDFEXTSORT
    ReDim tgVefRdfInfo(0 To 0) As VEFRDFINFO
    'Contract
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", SpotMG
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)  'Get and save CHF record length
    'Line
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", SpotMG
    On Error GoTo 0
    ReDim tmClf(0 To 0) As CLFLIST
    imClfRecLen = Len(tmClf(0).ClfRec)  'Get and save CLF record length
    'Flight
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", SpotMG
    On Error GoTo 0
    ReDim tmCff(0 To 0) As CFFLIST
    imCffRecLen = Len(tmCff(0).CffRec)  'Get and save CLF record length
    'Games
    hmGhf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ghf.Btr)", SpotMG
    On Error GoTo 0
    imGhfRecLen = Len(tmGhf)  'Get and save CLF record length
    'Games
    hmGsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Gsf.Btr)", SpotMG
    On Error GoTo 0
    ReDim tmGsf(0 To 0) As GSF
    imGsfRecLen = Len(tmGsf(0))  'Get and save CFF record length
    hmGhf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ghf.Btr)", SpotMG
    On Error GoTo 0
    'Games
    hmCgf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCgf, "", sgDBPath & "Cgf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cgf.Btr)", SpotMG
    On Error GoTo 0
    imCgfRecLen = Len(tmCgf)  'Get and save CLF record length
   'MultiName
    hmMnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf.Btr)", SpotMG
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)  'Get and save MNF record length
    'Research Data
    hmDrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Drf.Btr)", SpotMG
    On Error GoTo 0
    imDrfRecLen = Len(tmDrf)  'Get and save MNF record length
    hmDpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dpf.Btr)", SpotMG
    On Error GoTo 0
    hmDef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Def.Btr)", SpotMG
    On Error GoTo 0
    hmRaf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Raf.Btr)", SpotMG
    On Error GoTo 0
    hmVef = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", SpotMG
    On Error GoTo 0
    imVefRecLen = Len(tmVef)  'Get and save VEF record length
    hmSdf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", SpotMG
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)  'Get and save SDF record length
    hmSsf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ssf.Btr)", SpotMG
    On Error GoTo 0
    ' Spot MG File
    hmSmf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf.Btr)", SpotMG
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)  'Get and save SMF record length
    'Spot tracking
    hmStf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmStf, "", sgDBPath & "Stf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Stf.Btr)", SpotMG
    On Error GoTo 0
    imStfRecLen = Len(tmStf)  'Get and save STF record length
    'Advertiser
    hmAdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", SpotMG
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)  'Get and save ADF record length
    'Feed
    hmFsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Fsf.Btr)", SpotMG
    On Error GoTo 0
    imFsfRecLen = Len(tmFsf)  'Get and save ADF record length
    'Rotation
    hmCrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Crf.Btr)", SpotMG
    On Error GoTo 0
    imCrfRecLen = Len(tmCrf)  'Get and save CRF record length
    hmRlf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmRlf, "", sgDBPath & "Rlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rlf.Btr)", SpotMG
    On Error GoTo 0
    hmSxf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSxf, "", sgDBPath & "Sxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sxf.Btr)", SpotMG
    On Error GoTo 0
    ReDim tmVcf0(0 To 0) As VCF
    ReDim tmVcf6(0 To 0) As VCF
    ReDim tmVcf7(0 To 0) As VCF
    ReDim imLengths(0 To 0) As Integer
    ilLenMin = 32000
    ilLenMax = 0
    For ilVpf = LBound(tgVpf) To UBound(tgVpf) Step 1
        'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If tgMVef(ilVef).iCode = tgVpf(ilVpf).iVefKCode Then
            ilVef = gBinarySearchVef(tgVpf(ilVpf).iVefKCode)
            If ilVef <> -1 Then
                If (tgMVef(ilVef).sState = "A") And (tgMVef(ilVef).sType = "C") Then
                    If ((imVehGame) And (tgMVef(ilVef).sType = "G")) Or ((Not imVehGame) And (tgMVef(ilVef).sType <> "G")) Then
                        For ilLen = LBound(tgVpf(ilVpf).iSLen) To UBound(tgVpf(ilVpf).iSLen) Step 1
                            If tgVpf(ilVpf).iSLen(ilLen) > 0 Then
                                ilFound = False
                                For ilLoop = LBound(imLengths) To UBound(imLengths) - 1 Step 1
                                    If imLengths(ilLoop) = tgVpf(ilVpf).iSLen(ilLen) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilLoop
                                If Not ilFound Then
                                    If tgVpf(ilVpf).iSLen(ilLen) < ilLenMin Then
                                        ilLenMin = tgVpf(ilVpf).iSLen(ilLen)
                                    End If
                                    If tgVpf(ilVpf).iSLen(ilLen) > ilLenMax Then
                                        ilLenMax = tgVpf(ilVpf).iSLen(ilLen)
                                    End If
                                    imLengths(UBound(imLengths)) = tgVpf(ilVpf).iSLen(ilLen)
                                    ReDim Preserve imLengths(0 To UBound(imLengths) + 1) As Integer
                                End If
                            End If
                        Next ilLen
                    End If
                End If
        '        Exit For
            End If
        'Next ilVef
    Next ilVpf
    For ilLoop = ilLenMin To ilLenMax Step 1
        For ilLen = LBound(imLengths) To UBound(imLengths) - 1 Step 1
            If ilLoop = imLengths(ilLen) Then
                cbcLen.AddItem Trim$(str$(imLengths(ilLen)))
                Exit For
            End If
        Next ilLen
    Next ilLoop
    imIgnoreChg = True
    For ilLen = 0 To cbcLen.ListCount - 1 Step 1
        If Val(cbcLen.List(ilLen)) = igMGSpotLen Then
            cbcLen.ListIndex = ilLen
        End If
    Next ilLen
    ReDim tgCntSpot(0 To 0) As CNTSPOT
    ReDim tgSpotLinks(0 To 0) As SPOTLINKS
    ilRet = gObtainRcfRifRdf()
    'plcMissedWk.Caption = sgMGStartDate
    lmMissedWkDate = gDateValue(sgMGStartDate)
    lmEarliestMonDate = lgMGEarliestDate    'lgMGAllowDate
    Do While gWeekDayLong(lmEarliestMonDate) <> 0
        lmEarliestMonDate = lmEarliestMonDate - 1
    Loop
    imIgnoreChg = True
    hbcMissedWk.Max = (lgMGLatestDate - lmEarliestMonDate) \ 7 + 1 '1 for adjusting
    'plcAiredWk.Caption = sgMGStartDate
    lmAiredWkDate = gDateValue(sgMGStartDate)
    imIgnoreChg = True
    hbcAiredWk.Max = (lgMGLatestDate - lmEarliestMonDate) \ 7 + 1 '1 for adjusting
    '3/9/99- Changed range to dates to be 13 in past/13 in future (I changed this so scroll is in middle)
    imIgnoreChg = True
    hbcMissedWk.Value = (lmMissedWkDate - lmEarliestMonDate) \ 7 + 1 '1   '(lmMissedWkDate - lmEarliestMonDate) \ 7 + 1
    imIgnoreChg = True
    hbcAiredWk.Value = (lmMissedWkDate - lmEarliestMonDate) \ 7 + 1 '1    '(lmMissedWkDate - lmEarliestMonDate) \ 7 + 1
    imIgnoreChg = False
    'plcSpotWk.Caption = sgMGStartDate
    lmSpotWkDate = gDateValue(sgMGStartDate)
    lmEarliestMonDate = lgMGEarliestDate    'lgMGAllowDate
    Do While gWeekDayLong(lmEarliestMonDate) <> 0
        lmEarliestMonDate = lmEarliestMonDate - 1
    Loop
    imIgnoreChg = True
    hbcSpotWk.Max = (lgMGLatestDate - lmEarliestMonDate) \ 7 + 1 '1 for adjusting
    'plcVehDPWk.Caption = sgMGStartDate
    lmVehDPWkDate = gDateValue(sgMGStartDate)
    imIgnoreChg = True
    hbcVehDPWk.Max = (lgMGLatestDate - lmEarliestMonDate) \ 7 + 1 '1 for adjusting
    imIgnoreChg = True
    hbcSpotWk.Value = (lmSpotWkDate - lmEarliestMonDate) \ 7 + 1
    imIgnoreChg = True
    hbcVehDPWk.Value = (lmSpotWkDate - lmEarliestMonDate) \ 7 + 1
    imIgnoreChg = False
    If (lgChfMGCode > 0) And ((tgSpf.sGUsePropSys = "Y") Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName)) Then
        tmChfSrchKey.lCode = lgChfMGCode
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, lgChfMGCode, False, tmChf, tmClf(), tmCff())
        If Not ilRet Then
            imTerminate = True
            Exit Sub
        End If
        tmAdfSrchKey.iCode = tmChf.iAdfCode
        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        'slStr = Trim$(tmAdf.sName)
        If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
            slStr = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
        Else
            slStr = Trim$(tmAdf.sName)
        End If
        ilPos = InStr(slStr, "&")
        If ilPos > 0 Then
            slStr = Left$(slStr, ilPos - 1) & "&&" & Mid$(slStr, ilPos + 1)
        End If
        smScreenCaption = "MG + for" & str$(tmChf.lCntrNo) & ", " & slStr & "," & str$(igMGSpotLen) & "sec"
        plcSpotWk.Caption = sgMGStartDate
        mHbcSpotWkChange
        'mHbcVehDPWkChange
        'mHbcMissedWkChange
        edcMForN.Visible = False
        plcTabSelection.TabIndex = 1
        plcScreen_Paint
    Else
        'mHbcMissedWkChange
        'mHbcVehDPWkChange
        'plcTabSelection.Enabled = False
        'plcTabSelection.Tabs.Remove (2)
        edcMForN.Visible = True
        plcTabSelection.TabIndex = 0
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
'*      Procedure Name:mReadChfClfRdfCffRec            *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read records                   *
'*                                                     *
'*******************************************************
Private Function mReadChfClfRdfCffRec(llChfCode As Long, ilLineNo As Integer, llFsfCode As Long, ilOrderedGameNo As Integer, slSpotDate As String, slLnStartDate As String, slLnEndDate As String, slNoSpots As String) As Integer
'
'   iRet = mReadChfClfRdfCffRec(llChfCode, ilLineNo, slMissedDate, SlStartDate, slEndDate, slNoSpots)
'   Where:
'       llChfCode (I) - Contract code
'       ilLineNo(I)- Line number
'       slMissedDate(I)- Missed date or date to find bracketing week
'       slLnStartdate(O)- line start date
'       slLnEndDate(O)- line end date
'       slNoSpots(O)- if "" then invalid week
'       tmICff(1)(O)- contains valid flight week (if sDelete = "Y", then week is invalid)
'       iRet (O)- True if all records read,
'                 False if error in read
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim llSpotDate As Long
    Dim ilNoSpots As Integer
    Dim ilVef As Integer

    slLnStartDate = ""
    slLnEndDate = ""
    slNoSpots = ""
    tmICff(1).sDelete = "Y"  'Set as flag that illegal week
    If mReadChfClfRdfRec(llChfCode, ilLineNo, llFsfCode) Then
        llStartDate = 0
        llEndDate = 0
        llSpotDate = gDateValue(slSpotDate)
        If llChfCode > 0 Then
            ilVef = gBinarySearchVef(tmIClf.iVefCode)
            If ilVef <> -1 Then
                If tgMVef(ilVef).sType <> "G" Then
                    tmCffSrchKey.lChfCode = llChfCode
                    tmCffSrchKey.iClfLine = ilLineNo
                    tmCffSrchKey.iCntRevNo = tmIClf.iCntRevNo
                    tmCffSrchKey.iPropVer = tmIClf.iPropVer
                    tmCffSrchKey.iStartDate(0) = 0
                    tmCffSrchKey.iStartDate(1) = 0
                    ilRet = btrGetGreaterOrEqual(hmCff, tmICff(2), imCffRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Else
                    tmCgfSrchKey1.lClfCode = tmIClf.lCode
                    ilRet = btrGetEqual(hmCgf, tmCgf, imCgfRecLen, tmCgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If tmIClf.lCode = tmCgf.lClfCode Then
                        gCgfToCff tmIClf, tmCgf, tmICff()
                        tmICff(2) = tmICff(0)   'tmICff(1)
                        If tmCgf.iGameNo <> ilOrderedGameNo Then
                            tmICff(2).iCntRevNo = -1 'Force to read next
                        End If
                    Else
                        tmICff(2).lChfCode = -1
                    End If
                End If
            Else
                mReadChfClfRdfCffRec = False
                Exit Function
            End If
        Else
            tmICff(2) = tmFCff(1)
            tmICff(2).lChfCode = llChfCode
            tmICff(2).iClfLine = ilLineNo
            ilRet = BTRV_ERR_NONE
        End If
        Do While (ilRet = BTRV_ERR_NONE) And (tmICff(2).lChfCode = llChfCode) And (tmICff(2).iClfLine = ilLineNo)
            If (tmICff(2).iCntRevNo = tmIClf.iCntRevNo) And (tmICff(2).iPropVer = tmIClf.iPropVer) Then 'And (tmICff(2).sDelete <> "Y") Then
                tmICff(2).sDelete = "N"  'Set flight as if not deleted (delete is set if line replaced)
                                        'Only if line is altered (not scheduled will this happen)
                gUnpackDate tmICff(2).iStartDate(0), tmICff(2).iStartDate(1), slStartDate    'Week Start date
                gUnpackDate tmICff(2).iEndDate(0), tmICff(2).iEndDate(1), slEndDate    'Week Start date
                If llStartDate = 0 Then
                    llStartDate = gDateValue(slStartDate)
                    llEndDate = gDateValue(slEndDate)
                Else
                    If gDateValue(slStartDate) < llStartDate Then
                        llStartDate = gDateValue(slStartDate)
                    End If
                    If gDateValue(slEndDate) > llEndDate Then
                        llEndDate = gDateValue(slEndDate)
                    End If
                End If
                If (llSpotDate >= gDateValue(slStartDate)) And (llSpotDate <= gDateValue(slEndDate)) Then
                    tmICff(1) = tmICff(2)
                    ilNoSpots = 0
                    'If (tmCff(1).iSpotsWk <> 0) Or (tmCff(1).iXSpotsWk <> 0) Then 'Weekly
                    If (tmICff(1).sDyWk <> "D") Then  'Weekly
                        ilNoSpots = tmICff(1).iSpotsWk + tmICff(1).iXSpotsWk
                    Else    'Daily
                        For ilLoop = 0 To 6 Step 1
                            ilNoSpots = ilNoSpots + tmICff(1).iDay(ilLoop)
                        Next ilLoop
                    End If
                    slNoSpots = Trim$(str$(ilNoSpots))
                    'Don't exit as end date of all flights must be determined
                End If
            End If
            If llChfCode > 0 Then
                If tgMVef(ilVef).sType <> "G" Then
                    ilRet = btrGetNext(hmCff, tmICff(2), imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Else
                    ilRet = btrGetNext(hmCgf, tmCgf, imCgfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    If tmIClf.lCode <> tmCgf.lClfCode Then
                        Exit Do
                    End If
                    gCgfToCff tmIClf, tmCgf, tmICff()
                    tmICff(2) = tmICff(0)   'tmICff(1)
                    If tmCgf.iGameNo <> ilOrderedGameNo Then
                        tmICff(2).iCntRevNo = -1 'Force to read next
                    End If
                End If
            Else
                Exit Do
            End If
        Loop
        If llStartDate > 0 Then
            slLnStartDate = Format$(llStartDate, "m/d/yy")
            slLnEndDate = Format$(llEndDate, "m/d/yy")
        End If
        mReadChfClfRdfCffRec = True
    Else
        mReadChfClfRdfCffRec = False
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadChfClfRdfRec               *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read records                   *
'*                                                     *
'*******************************************************
Private Function mReadChfClfRdfRec(llChfCode As Long, ilLineNo As Integer, llFsfCode As Long) As Integer
'
'   iRet = mReadChfClfRdfRec(llChfCode, ilLineNo)
'   Where:
'       llChfCode (I) - Contract code
'       ilLineNo(I)- Line number
'       iRet (O)- True if all records read,
'                 False if error in read
'
    Dim ilRet As Integer
    Dim ilRdf As Integer
    If llChfCode > 0 Then
        'If llChfCode <> tmChf.lCode Then
            tmChfSrchKey.lCode = llChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                mReadChfClfRdfRec = False
                Exit Function
            End If
        'End If
        'If (tmIClf.lChfCode <> llChfCode) Or (tmIClf.iLine <> ilLineNo) Then
            tmClfSrchKey.lChfCode = llChfCode
            tmClfSrchKey.iLine = ilLineNo
            tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
            tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
            ilRet = btrGetGreaterOrEqual(hmClf, tmIClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tmIClf.lChfCode = llChfCode) And (tmIClf.iLine = ilLineNo) And ((tmIClf.sSchStatus <> "M") And (tmIClf.sSchStatus <> "F"))  'And (tmIClf.sSchStatus = "A")
                ilRet = btrGetNext(hmClf, tmIClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        'Else
        '    ilRet = BTRV_ERR_NONE
        'End If
        If (ilRet = BTRV_ERR_NONE) And (tmIClf.lChfCode = llChfCode) And (tmIClf.iLine = ilLineNo) Then
            'If tmRdf.iCode <> tmIClf.iRdfCode Then
            '    tmRdfSrchKey.iCode = tmIClf.iRdfCode  ' Rate card program/time File Code
            '    ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            '    If ilRet <> BTRV_ERR_NONE Then
            '        mReadChfClfRdfRec = False
            '        Exit Function
            '    End If
            'End If
            'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
            '    If tmIClf.iRdfcode = tgMRdf(ilRdf).iCode Then
                ilRdf = gBinarySearchRdf(tmIClf.iRdfCode)
                If ilRdf <> -1 Then
                    tmRdf = tgMRdf(ilRdf)
                    mReadChfClfRdfRec = True
                    Exit Function
                End If
            'Next ilRdf
            mReadChfClfRdfRec = False
        Else
            mReadChfClfRdfRec = False
        End If
    Else
        tmFSFSrchKey.lCode = llFsfCode
        ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        gMoveFeedToCntr tmFsf, tmRdf, tmChf, tmIClf, tmFCff(), hmFnf, hmPrf
        mReadChfClfRdfRec = True
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSchSpots                       *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine spot to be booked and*
'*                      which vehicle/daypart          *
'*                                                     *
'*******************************************************
Private Sub mSchSpots()
    Dim ilLoop As Integer
    Dim slLine As String
    Dim slIndex As String
    Dim slStr As String
    Dim ilIndex As Integer
    Dim ilCount As Integer
    Dim ilRet As Integer
    Dim ilRdf As Integer
    Dim ilSpotLkIndex As Integer
    Dim ilNoToMove As Integer
    Dim ilUpper As Integer
    Dim llSTime As Long
    Dim llETime As Long
    Dim llTime As Long
    Dim llDate As Long
    Dim ilDate As Integer
    Dim ilRow As Integer
    Dim ilDay As Integer
    Dim ilDayIndex As Integer
    Dim ilNum As Integer
    Dim ilVef As Integer
    Dim ilParse As Integer
    Dim ilSdf As Integer
    Dim llNowDate As Long
    Dim llChfCode As Long
    Dim ilAdfCode As Integer
    Dim ilVehComp As Integer
    Dim llStartDateLen As Long
    Dim llEndDateLen As Long
    Dim ilLineNo As Integer
    Dim slMissedDate As String
    Dim llMissedDate As Long
    Dim ilMissedDay As Integer
    Dim slLnStartDate As String
    Dim slLnEndDate As String
    Dim ilVehOk As Integer
    Dim slNoSpots As String
    Dim slLength As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilPass As Integer   'First pass look for opening withoput removing files; second pass remove fills to make room
    Dim ilVehPass As Integer
    Dim ilVehSPass As Integer
    Dim ilVehEPass As Integer
    Dim ilLnRdf As Integer
    Dim ilTest As Integer
    Dim ilBooked As Integer
    Dim llOvSTime As Long
    Dim llOvETime As Long
    Dim ilFound As Integer
    Dim ilGameNo As Integer
    Dim tlSdf As SDF
    Dim tlCff As CFF
    ReDim ilCffDays(0 To 6) As Integer
    ReDim slField(0 To 6) As String

    Screen.MousePointer = vbHourglass
    llDate = 7 * (hbcAiredWk.Value - 1) + lmEarliestMonDate
    llNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
    'Build table of spots to be booked
    ReDim lmSdfRecPos(0 To 0) As Long
    For ilLoop = 0 To lbcMissed.ListCount - 1 Step 2
        If lbcMissed.Selected(ilLoop) Or lbcMissed.Selected(ilLoop + 1) Then
            If rbcSort(0).Value Then
                slLine = lbcMissed.List(ilLoop + 1)
            Else
                slLine = lbcMissed.List(ilLoop)
            End If
            ilRet = gParseItem(slLine, 6, "|", slIndex)
            ilIndex = Val(Trim$(slIndex))
            If ilIndex >= 0 Then
                If rbcMove(0).Value Then
                    'Move one
                    ilSpotLkIndex = tgCntSpot(ilIndex).iSpotLkIndex
                    Do While ilSpotLkIndex >= 0
                        If tgSpotLinks(ilSpotLkIndex).iStatus = 0 Then
                            lmSdfRecPos(UBound(lmSdfRecPos)) = tgSpotLinks(ilSpotLkIndex).lSdfRecPos
                            ReDim Preserve lmSdfRecPos(0 To UBound(lmSdfRecPos) + 1) As Long
                            Exit Do
                        End If
                        ilSpotLkIndex = tgSpotLinks(ilSpotLkIndex).iSpotLkIndex
                    Loop
                ElseIf rbcMove(1).Value Then
                    'Move all
                    ilSpotLkIndex = tgCntSpot(ilIndex).iSpotLkIndex
                    Do While ilSpotLkIndex >= 0
                        If tgSpotLinks(ilSpotLkIndex).iStatus = 0 Then
                            lmSdfRecPos(UBound(lmSdfRecPos)) = tgSpotLinks(ilSpotLkIndex).lSdfRecPos
                            ReDim Preserve lmSdfRecPos(0 To UBound(lmSdfRecPos) + 1) As Long
                        End If
                        ilSpotLkIndex = tgSpotLinks(ilSpotLkIndex).iSpotLkIndex
                    Loop
                Else
                    'Ask
                    ilCount = 0
                    ilSpotLkIndex = tgCntSpot(ilIndex).iSpotLkIndex
                    Do While ilSpotLkIndex >= 0
                        If tgSpotLinks(ilSpotLkIndex).iStatus = 0 Then
                            ilCount = ilCount + 1
                        End If
                        ilSpotLkIndex = tgSpotLinks(ilSpotLkIndex).iSpotLkIndex
                    Loop
                    If ilCount > 1 Then
                        sgGenMsg = ""
                        slLine = tgCntSpot(ilIndex).sKey
                        ilRet = gParseItem(slLine, 2, "|", slStr)
                        sgGenMsg = "For: " & slStr
                        ilRet = gParseItem(slLine, 3, "|", slStr)
                        sgGenMsg = sgGenMsg & " " & slStr
                        slStr = ""
                        'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                        '    If tgCntSpot(ilIndex).iRdfcode = tgMRdf(ilRdf).iCode Then
                            ilRdf = gBinarySearchRdf(tgCntSpot(ilIndex).iRdfCode)
                            If ilRdf <> -1 Then
                                slStr = Trim$(tgMRdf(ilRdf).sName)
                        '        Exit For
                            End If
                        'Next ilRdf
                        sgGenMsg = sgGenMsg & " " & slStr
                        sgGenMsg = sgGenMsg & ", schedule number of spots indicated below"
                        sgCMCTitle(0) = "Ok"
                        sgCMCTitle(1) = ""
                        sgCMCTitle(2) = ""
                        sgCMCTitle(3) = ""
                        igDefCMC = 0
                        igEditBox = 1
                        sgEditValue = Trim$(str$(ilCount))
                        GenMsg.Show vbModal
                        If Val(sgEditValue) < ilCount Then
                            ilCount = Val(sgEditValue)
                        End If
                    End If
                    If ilCount > 0 Then
                        ilSpotLkIndex = tgCntSpot(ilIndex).iSpotLkIndex
                        Do While ilSpotLkIndex >= 0
                            If tgSpotLinks(ilSpotLkIndex).iStatus = 0 Then
                                lmSdfRecPos(UBound(lmSdfRecPos)) = tgSpotLinks(ilSpotLkIndex).lSdfRecPos
                                ReDim Preserve lmSdfRecPos(0 To UBound(lmSdfRecPos) + 1) As Long
                                ilCount = ilCount - 1
                                If ilCount <= 0 Then
                                    Exit Do
                                End If
                            End If
                            ilSpotLkIndex = tgSpotLinks(ilSpotLkIndex).iSpotLkIndex
                        Loop
                    End If
                End If
            End If
        End If
    Next ilLoop
    pbcLbcMissed_Paint
    pbcLbcAired_Paint
    DoEvents
    If UBound(lmSdfRecPos) <= LBound(lmSdfRecPos) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    ReDim lmTSdfRecPos(0 To UBound(lmSdfRecPos)) As Long
    Randomize
    ilNoToMove = UBound(lmSdfRecPos)
    ilUpper = LBound(lmSdfRecPos)
    Do While ilNoToMove >= 1
        ilIndex = Int((ilNoToMove) * Rnd + 1) - 1
        lmTSdfRecPos(ilUpper) = lmSdfRecPos(ilIndex)
        For ilLoop = ilIndex To UBound(lmSdfRecPos) - 1 Step 1
            lmSdfRecPos(ilLoop) = lmSdfRecPos(ilLoop + 1)
        Next ilLoop
        ilUpper = ilUpper + 1
        ilNoToMove = ilNoToMove - 1
    Loop
    For ilLoop = LBound(lmTSdfRecPos) To UBound(lmTSdfRecPos) Step 1
        lmSdfRecPos(ilLoop) = lmTSdfRecPos(ilLoop)
    Next ilLoop
    ReDim tmMGBookInfo(0 To 0) As MGBOOKINFO
    For ilLoop = 0 To lbcAired.ListCount - 1 Step 1
        If lbcAired.Selected(ilLoop) Then
            slLine = lbcAired.List(ilLoop)
            ilRet = gParseItem(slLine, 6, "|", slIndex)
            ilIndex = Val(Trim$(slIndex))
            If ilIndex >= 0 Then
                tmMGBookInfo(UBound(tmMGBookInfo)).iVefIndex = tgVefRdfInfo(ilIndex).iVefIndex
                tmMGBookInfo(UBound(tmMGBookInfo)).iRdfIndex = tgVefRdfInfo(ilIndex).iRdfIndex
                ReDim Preserve tmMGBookInfo(0 To UBound(tmMGBookInfo) + 1) As MGBOOKINFO
            End If
        End If
    Next ilLoop
    If UBound(tmMGBookInfo) <= LBound(tmMGBookInfo) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If ckcCntrVehOnly.Value = vbChecked Then
        ilVehSPass = 0
        ilVehEPass = 1
    Else
        ilVehSPass = 1
        ilVehEPass = 1
    End If
    ReDim tmTMGBookInfo(0 To UBound(tmMGBookInfo)) As MGBOOKINFO
    For ilLoop = LBound(lmSdfRecPos) To UBound(lmSdfRecPos) - 1 Step 1

        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, lmSdfRecPos(ilLoop), INDEXKEY0, BTRV_LOCK_NONE)

        If (ilRet = BTRV_ERR_NONE) And (tmSdf.sSchStatus = "M") Then
            ilGameNo = tmSdf.iGameNo
            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slMissedDate
            llMissedDate = gDateValue(slMissedDate)
            ilMissedDay = gWeekDayLong(llMissedDate)
            ilRet = mReadChfClfRdfCffRec(tmSdf.lChfCode, tmSdf.iLineNo, tmSdf.lFsfCode, ilGameNo, slMissedDate, slLnStartDate, slLnEndDate, slNoSpots)

            llChfCode = 0   'Force to recompute value so ilBkQH gets commputed
            tmTCff(0) = tmICff(1)      '5-9-06
            gGetLineSchParameters hmSsf, tgSsf(), lgSsfDate(), lgSsfRecPos(), gDateValue(slMissedDate), tmSdf.iVefCode, tmChf.iAdfCode, ilGameNo, tmTCff(), tmIClf, tmRdf, lmSepLength, llStartDateLen, llEndDateLen, llChfCode, ilLineNo, ilAdfCode, ilVehComp, imHour(), imDay(), imQH(), imAHour(), imADay(), imAQH(), lmTBStartTime(), lmTBEndTime(), 0, imSkip(), tmChf.sType, tmChf.iPctTrade, imBkQH, False, imPriceLevel, False
            Randomize
            ilNoToMove = UBound(tmMGBookInfo)
            ilUpper = LBound(tmMGBookInfo)
            Do While ilNoToMove >= 1
                ilIndex = Int((ilNoToMove) * Rnd + 1) - 1
                tmTMGBookInfo(ilUpper) = tmMGBookInfo(ilIndex)
                For ilNum = ilIndex To UBound(tmMGBookInfo) - 1 Step 1
                    tmMGBookInfo(ilNum) = tmMGBookInfo(ilNum + 1)
                Next ilNum
                ilUpper = ilUpper + 1
                ilNoToMove = ilNoToMove - 1
            Loop
            For ilNum = LBound(tmTMGBookInfo) To UBound(tmTMGBookInfo) Step 1
                tmMGBookInfo(ilNum) = tmTMGBookInfo(ilNum)
            Next ilNum
            If (ckcCntrVehOnly.Value = vbChecked) Or ((ckcPkgVeh.Value = vbChecked) And (tmIClf.sType = "H")) Then
                ilRet = gObtainChfClf(hmCHF, hmClf, tmChf.lCode, False, tgChfSpot, tgClfSpot())
            End If
            smRdfInOut = ""
            imRdfAnfCode = 0
            'For ilRow = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
            '    If tgMRdf(ilRow).iCode = tmIClf.iRdfcode Then
                ilLnRdf = gBinarySearchRdf(tmIClf.iRdfCode)
                'If ilRow <> -1 Then
                If ilLnRdf <> -1 Then
                    smRdfInOut = tgMRdf(ilLnRdf).sInOut
                    imRdfAnfCode = tgMRdf(ilLnRdf).ianfCode
            '        Exit For
                End If
            'Next ilRow
            ilBooked = False
            For ilVehPass = ilVehSPass To ilVehEPass Step 1
                For ilIndex = LBound(tmMGBookInfo) To UBound(tmMGBookInfo) - 1 Step 1
                    ilVef = tmMGBookInfo(ilIndex).iVefIndex
                    imVefCode = tgMVef(ilVef).iCode

                    imVpfIndex = gBinarySearchVpfPlus(imVefCode)    'gVpfFind(SpotMG, imVefCode)
                    If (tgVpf(imVpfIndex).iLLD(0) <> 0) Or (tgVpf(imVpfIndex).iLLD(1) <> 0) Then
                        gUnpackDateLong tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), lmLastLogDate
                    Else
                        lmLastLogDate = -1
                    End If
                    If tgVpf(imVpfIndex).sSCompType = "T" Then
                        gUnpackLength tgVpf(imVpfIndex).iSCompLen(0), tgVpf(imVpfIndex).iSCompLen(1), "3", False, slLength
                        lmCompTime = CLng(gLengthToCurrency(slLength))
                    Else
                        lmCompTime = 0&
                    End If
                    ilRdf = tmMGBookInfo(ilIndex).iRdfIndex
                    If (ckcPkgVeh.Value = vbChecked) And (tmIClf.sType = "H") Then
                        ilVehOk = False
                        For ilRow = LBound(tgClfSpot) To UBound(tgClfSpot) - 1 Step 1
                            If tmIClf.iPkLineNo = tgClfSpot(ilRow).ClfRec.iPkLineNo Then
                                If imVefCode = tgClfSpot(ilRow).ClfRec.iVefCode Then
                                    ilVehOk = True
                                    Exit For
                                End If
                            End If
                        Next ilRow
                    Else
                        ilVehOk = True
                    End If
                    If ilVehOk Then
                        If (ckcCntrVehOnly.Value = vbChecked) Then
                            If (ckcPkgVeh.Value <> vbChecked) Or (tmIClf.sType <> "H") Then
                                ilVehOk = False
                                If ilVehPass = 0 Then
                                    If imVefCode = tmIClf.iVefCode Then
                                        ilVehOk = True
                                    End If
                                Else
                                    For ilRow = LBound(tgClfSpot) To UBound(tgClfSpot) - 1 Step 1
                                        If imVefCode = tgClfSpot(ilRow).ClfRec.iVefCode Then
                                            ilVehOk = True
                                            Exit For
                                        End If
                                    Next ilRow
                                End If
                            End If
                        End If
                    End If
                    If (ilVehOk) And (ckcAirWeek.Value = vbChecked) Then
                        'If checking only line of spot
                        If imVefCode = tmIClf.iVefCode Then
                            tlSdf = tmSdf
                            gPackDateLong llDate, tlSdf.iDate(0), tlSdf.iDate(1)
                            ilVehOk = gGetSpotFlight(tlSdf, tmIClf, hmCff, hmSmf, tlCff)
                        Else
                            ilVehOk = False
                        End If
                        'If checking against contract
                        'ilVehOk = False
                        'For ilRow = LBound(tgClfSpot) To UBound(tgClfSpot) - 1 Step 1
                        '    tlSdf = tmSdf
                        '    'Only sdfChfCode as date are obtained from sdf in gGetSpotFlight
                        '    gPackDateLong llDate, tlSdf.iDate(0), tlSdf.iDate(1)
                        '    ilVehOk = gGetSpotFlight(tlSdf, tgClfSpot(ilRow).ClfRec, hmCff, hmSmf, tlCff)
                        '    If ilVehOk Then
                        '        Exit For
                        '    End If
                        'Next ilRow
                    End If
                    If ilVehOk Then
                        For ilPass = 0 To 1 Step 1
                            For ilRow = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                If (tgMRdf(ilRdf).iStartTime(0, ilRow) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilRow) <> 0) Then
                                    gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilRow), tgMRdf(ilRdf).iStartTime(1, ilRow), False, llSTime
                                    gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilRow), tgMRdf(ilRdf).iEndTime(1, ilRow), True, llETime
                                    ilDayIndex = 0
                                    For ilDay = 1 To 7 Step 1
                                        'If tgMRdf(ilRdf).sWkDays(ilRow, ilDay) = "Y" Then
                                        If tgMRdf(ilRdf).sWkDays(ilRow, ilDay - 1) = "Y" Then
                                            ilCffDays(ilDayIndex) = True
                                        Else
                                            ilCffDays(ilDayIndex) = False
                                        End If
                                        ilDayIndex = ilDayIndex + 1
                                    Next ilDay
                                    If ckcDaysTimes.Value = vbChecked Then
                                        'Use override times if defined instead of Rate Card
                                        If (tmIClf.iStartTime(0) <> 1) Or (tmIClf.iStartTime(1) <> 0) Then
                                            gUnpackTimeLong tmIClf.iStartTime(0), tmIClf.iStartTime(1), False, llOvSTime
                                            gUnpackTimeLong tmIClf.iEndTime(0), tmIClf.iEndTime(1), True, llOvETime
                                            If (llOvETime < llSTime) Or (llOvSTime > llETime) Then
                                                llSTime = -1
                                                llETime = -2
                                            Else
                                                If llOvSTime > llSTime Then
                                                    llSTime = llOvSTime
                                                End If
                                                If llOvETime < llETime Then
                                                    llETime = llOvETime
                                                End If
                                            End If
                                        Else
                                            ilFound = False
                                            If ilLnRdf <> -1 Then
                                                For ilTest = LBound(tgMRdf(ilLnRdf).iStartTime, 2) To UBound(tgMRdf(ilLnRdf).iStartTime, 2) Step 1 'Row
                                                    If (tgMRdf(ilLnRdf).iStartTime(0, ilTest) <> 1) Or (tgMRdf(ilLnRdf).iStartTime(1, ilTest) <> 0) Then
                                                        gUnpackTimeLong tgMRdf(ilLnRdf).iStartTime(0, ilTest), tgMRdf(ilLnRdf).iStartTime(1, ilTest), False, llOvSTime
                                                        gUnpackTimeLong tgMRdf(ilLnRdf).iEndTime(0, ilTest), tgMRdf(ilLnRdf).iEndTime(1, ilTest), True, llOvETime
                                                        If (llOvETime >= llSTime) And (llOvSTime <= llETime) Then
                                                            If llOvSTime > llSTime Then
                                                                llSTime = llOvSTime
                                                            End If
                                                            If llOvETime < llETime Then
                                                                llETime = llOvETime
                                                            End If
                                                            ilFound = True
                                                            Exit For
                                                        End If
                                                    End If
                                                Next ilTest
                                            End If
                                            If Not ilFound Then
                                                llSTime = -1
                                                llETime = -2
                                            End If
                                        End If
                                        If llSTime <> -1 Then
                                            tmCffSrchKey.lChfCode = tmChf.lCode  'llChfCode
                                            tmCffSrchKey.iClfLine = tmIClf.iLine 'tlSdf.iLineNo using line so avg price can be obtained for package line which bill by airing
                                            tmCffSrchKey.iCntRevNo = tmIClf.iCntRevNo
                                            tmCffSrchKey.iPropVer = tmIClf.iPropVer
                                            tmCffSrchKey.iStartDate(0) = 0
                                            tmCffSrchKey.iStartDate(1) = 0
                                            imCffRecLen = Len(tmPCff)
                                            ilRet = btrGetGreaterOrEqual(hmCff, tmPCff, imCffRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                            Do While (ilRet = BTRV_ERR_NONE) And (tmPCff.lChfCode = tmChf.lCode) And (tmPCff.iClfLine = tmIClf.iLine)
                                                If (tmPCff.iCntRevNo = tmIClf.iCntRevNo) And (tmPCff.iPropVer = tmIClf.iPropVer) Then 'And (tmCff(2).sDelete <> "Y") Then
                                                    gUnpackDateLong tmPCff.iStartDate(0), tmPCff.iStartDate(1), llStartDate    'Week Start date
                                                    gUnpackDateLong tmPCff.iEndDate(0), tmPCff.iEndDate(1), llEndDate    'Week Start date
                                                    If (llMissedDate >= llStartDate) And (llMissedDate <= llEndDate) Then
                                                        If tmPCff.sDyWk = "D" Then
                                                            For ilDay = 0 To 6 Step 1
                                                                If tmPCff.iDay(ilDay) > 0 Then
                                                                '    ilCffDays(ilDay + 1) = True
                                                                    If ilDay <> ilMissedDay Then
                                                                        ilCffDays(ilDay) = False
                                                                    End If
                                                                Else
                                                                    ilCffDays(ilDay) = False
                                                                End If
                                                            Next ilDay
                                                        Else
                                                            For ilDay = 0 To 6 Step 1
                                                                If (tmPCff.iDay(ilDay) > 0) Or (tmPCff.sXDay(ilDay) = "Y") Then
                                                                '    ilCffDays(ilDay + 1) = True
                                                                Else
                                                                    ilCffDays(ilDay) = False
                                                                End If
                                                            Next ilDay
                                                        End If
                                                        Exit Do
                                                    End If
                                                End If
                                                ilRet = btrGetNext(hmCff, tmPCff, imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                            Loop
                                        Else
                                            For ilDay = 0 To 6 Step 1
                                                ilCffDays(ilDay) = False
                                            Next ilDay
                                        End If
                                    End If


                                    If (ckcOverride.Value = vbChecked) And (llSTime <> -1) Then
                                        If tmIClf.sType = "H" Then
                                            'Get vehicle name from package line
                                            tmClfSrchKey.lChfCode = tmChf.lCode
                                            tmClfSrchKey.iLine = tmIClf.iPkLineNo
                                            tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
                                            tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
                                            ilRet = btrGetGreaterOrEqual(hmClf, tmPclf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                            If ilRet = BTRV_ERR_NONE Then
                                                If (tmPclf.iStartTime(0) <> 1) Or (tmPclf.iStartTime(1) <> 0) Then
                                                    gUnpackTimeLong tmPclf.iStartTime(0), tmPclf.iStartTime(1), False, llTime
                                                    If llTime > llSTime Then
                                                        llSTime = llTime
                                                    End If
                                                    gUnpackTimeLong tmPclf.iEndTime(0), tmPclf.iEndTime(1), True, llTime
                                                    If llTime < llETime Then
                                                        llETime = llTime
                                                    End If
                                                End If
                                                tmCffSrchKey.lChfCode = tmChf.lCode  'llChfCode
                                                tmCffSrchKey.iClfLine = tmPclf.iLine 'tlSdf.iLineNo using line so avg price can be obtained for package line which bill by airing
                                                tmCffSrchKey.iCntRevNo = tmPclf.iCntRevNo
                                                tmCffSrchKey.iPropVer = tmPclf.iPropVer
                                                tmCffSrchKey.iStartDate(0) = 0
                                                tmCffSrchKey.iStartDate(1) = 0
                                                imCffRecLen = Len(tmPCff)
                                                ilRet = btrGetGreaterOrEqual(hmCff, tmPCff, imCffRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                                Do While (ilRet = BTRV_ERR_NONE) And (tmPCff.lChfCode = tmChf.lCode) And (tmPCff.iClfLine = tmPclf.iLine)
                                                    If (tmPCff.iCntRevNo = tmPclf.iCntRevNo) And (tmPCff.iPropVer = tmPclf.iPropVer) Then 'And (tmCff(2).sDelete <> "Y") Then
                                                        gUnpackDateLong tmPCff.iStartDate(0), tmPCff.iStartDate(1), llStartDate    'Week Start date
                                                        gUnpackDateLong tmPCff.iEndDate(0), tmPCff.iEndDate(1), llEndDate    'Week Start date
                                                        If (llMissedDate >= llStartDate) And (llMissedDate <= llEndDate) Then
                                                            If tmPCff.sDyWk = "D" Then
                                                                For ilDay = 0 To 6 Step 1
                                                                    If tmPCff.iDay(ilDay) > 0 Then
                                                                    '    ilCffDays(ilDay + 1) = True
                                                                    Else
                                                                        ilCffDays(ilDay) = False
                                                                    End If
                                                                Next ilDay
                                                            Else
                                                                For ilDay = 0 To 6 Step 1
                                                                    If (tmPCff.iDay(ilDay) > 0) Or (tmPCff.sXDay(ilDay) = "Y") Then
                                                                    '    ilCffDays(ilDay + 1) = True
                                                                    Else
                                                                        ilCffDays(ilDay) = False
                                                                    End If
                                                                Next ilDay
                                                            End If
                                                            Exit Do
                                                        End If
                                                    End If
                                                    ilRet = btrGetNext(hmCff, tmPCff, imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                Loop
                                            End If
                                        End If
                                    End If
                                    ReDim lmBkDates(0 To 0) As Long
                                    ilDayIndex = 0
                                    For ilDay = 1 To 7 Step 1
                                        If ilCffDays(ilDayIndex) Then
                                            If llNowDate < lmLastLogDate Then
                                                If (tgVpf(imVpfIndex).sMoveLLD = "Y") Then
                                                    If llDate + ilDay - 1 > llNowDate Then
                                                        lmBkDates(UBound(lmBkDates)) = llDate + ilDay - 1
                                                        ReDim Preserve lmBkDates(0 To UBound(lmBkDates) + 1) As Long
                                                    End If
                                                Else
                                                    If llDate + ilDay - 1 > lmLastLogDate Then
                                                        lmBkDates(UBound(lmBkDates)) = llDate + ilDay - 1
                                                        ReDim Preserve lmBkDates(0 To UBound(lmBkDates) + 1) As Long
                                                    End If
                                                End If
                                            Else
                                                If llDate + ilDay - 1 > llNowDate Then
                                                    lmBkDates(UBound(lmBkDates)) = llDate + ilDay - 1
                                                    ReDim Preserve lmBkDates(0 To UBound(lmBkDates) + 1) As Long
                                                End If
                                            End If
                                        End If
                                        ilDayIndex = ilDayIndex + 1
                                    Next ilDay
                                    ReDim lmTBkDates(0 To UBound(lmBkDates)) As Long
                                    Randomize
                                    ilNoToMove = UBound(lmBkDates)
                                    ilUpper = LBound(lmBkDates)
                                    Do While ilNoToMove >= 1
                                        ilNum = Int((ilNoToMove) * Rnd + 1) - 1
                                        lmTBkDates(ilUpper) = lmBkDates(ilNum)
                                        For ilDate = ilNum To UBound(lmBkDates) - 1 Step 1
                                            lmBkDates(ilDate) = lmBkDates(ilDate + 1)
                                        Next ilDate
                                        ilUpper = ilUpper + 1
                                        ilNoToMove = ilNoToMove - 1
                                    Loop
                                    For ilDate = LBound(lmTBkDates) To UBound(lmTBkDates) Step 1
                                        lmBkDates(ilDate) = lmTBkDates(ilDate)
                                    Next ilDate
                                    If llSTime <> -1 Then
                                        For ilDate = LBound(lmBkDates) To UBound(lmBkDates) - 1 Step 1
                                            ilRet = mBookSpot(lmSdfRecPos(ilLoop), imVefCode, tmMGBookInfo(ilIndex).iRdfIndex, lmBkDates(ilDate), lmBkDates(ilDate), llSTime, llETime, ilPass)
                                            If ilRet = -1 Then
                                                Screen.MousePointer = vbDefault
                                                Exit Sub
                                            ElseIf ilRet = 1 Then
                                                ilBooked = True
                                                igSpotMGReturn = 1
                                                'Reduce count
                                                For ilSdf = LBound(tgSpotLinks) To UBound(tgSpotLinks) - 1 Step 1
                                                    If tgSpotLinks(ilSdf).lSdfRecPos = lmSdfRecPos(ilLoop) Then
                                                        tgSpotLinks(ilSdf).iStatus = 1
                                                    End If
                                                Next ilSdf
                                                For ilSdf = 0 To lbcMissed.ListCount - 1 Step 2
                                                    If lbcMissed.Selected(ilSdf) Or lbcMissed.Selected(ilSdf + 1) Then
                                                        If rbcSort(0).Value Then
                                                            slLine = lbcMissed.List(ilSdf + 1)
                                                        Else
                                                            slLine = lbcMissed.List(ilSdf)
                                                        End If
                                                        ilRet = gParseItem(slLine, 6, "|", slIndex)
                                                        slIndex = Trim$(slIndex)
                                                        ilSpotLkIndex = tgCntSpot(Val(slIndex)).iSpotLkIndex
                                                        'If rbcSort(1).Value Then
                                                        '    slLine = lbcMissed.List(ilSdf)
                                                        'End If
                                                        Do While ilSpotLkIndex >= 0
                                                            If lmSdfRecPos(ilLoop) = tgSpotLinks(ilSpotLkIndex).lSdfRecPos Then
                                                                tgSpotLinks(ilSpotLkIndex).iStatus = 1
                                                                For ilParse = 1 To 6 Step 1
                                                                    ilRet = gParseItem(slLine, ilParse, "|", slField(ilParse))
                                                                Next ilParse
                                                                slField(3) = gSubStr(Trim$(slField(3)), "1")
                                                                slLine = "  " & Trim$(slField(1)) & "|" & slField(2) & "|  " & Trim$(slField(3)) & "|" & slField(4) & "|" & slField(5) & "|" & slField(6)
                                                                If rbcSort(0).Value Then
                                                                    'lbcMissed.List(ilSdf + 1) = gAlignStringByPixel(slLine, "|", imListField(), imListFieldChar())
                                                                    lbcMissed.List(ilSdf + 1) = slLine
                                                                Else
                                                                    'lbcMissed.List(ilSdf) = gAlignStringByPixel(slLine, "|", imListField(), imListFieldChar())
                                                                    lbcMissed.List(ilSdf) = slLine
                                                                End If
                                                                For ilVef = 0 To lbcAired.ListCount - 1 Step 1
                                                                    If lbcAired.Selected(ilVef) Then
                                                                        slLine = lbcAired.List(ilVef)
                                                                        ilRet = gParseItem(slLine, 6, "|", slIndex)
                                                                        If Val(Trim$(slIndex)) >= 0 Then
                                                                            If imVefCode = tgMVef(tgVefRdfInfo(Val(slIndex)).iVefIndex).iCode Then
                                                                                If tgMRdf(ilRdf).iCode = tgMRdf(tgVefRdfInfo(Val(slIndex)).iRdfIndex).iCode Then
                                                                                    For ilParse = 1 To 6 Step 1
                                                                                        ilRet = gParseItem(slLine, ilParse, "|", slField(ilParse))
                                                                                    Next ilParse
                                                                                    slField(3) = gSubStr(Trim$(slField(3)), "1")
                                                                                    If Val(slField(3)) < 0 Then
                                                                                        slField(3) = "0"
                                                                                    End If
                                                                                    slLine = slField(1) & "|" & slField(2) & "|  " & slField(3) & "|" & slField(4) & "|" & slField(5) & "|" & slField(6)
                                                                                    'lbcAired.List(ilVef) = gAlignStringByPixel(slLine, "|", imListField(), imListFieldChar())
                                                                                    lbcAired.List(ilVef) = slLine
                                                                                    Exit For
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next ilVef
                                                                Exit For
                                                            End If
                                                            ilSpotLkIndex = tgSpotLinks(ilSpotLkIndex).iSpotLkIndex
                                                        Loop
                                                    End If
                                                Next ilSdf
                                                Exit For
                                            End If
                                        Next ilDate
                                        If ilBooked Then
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next ilRow
                            If ilBooked Then
                                Exit For
                            End If
                        Next ilPass
                        If ilBooked Then
                            Exit For
                        End If
                    End If
                Next ilIndex
                If ilBooked Then
                    Exit For
                End If
            Next ilVehPass
        End If
    Next ilLoop
    pbcLbcMissed_Paint
    pbcLbcAired_Paint
    Screen.MousePointer = vbDefault
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
    tmcClick.Enabled = False

    sgCntrForDateStamp = ""
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload SpotMG
    igManUnload = NO
End Sub

Private Sub lbcVehDP_Scroll(Index As Integer)
    pbcLbcVehDp_Paint Index
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcLbcAired_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilAiredEnd As Integer
    Dim ilField As Integer
    Dim llWidth As Long
    Dim slFields(0 To 5) As String
    Dim llFgColor As Long
    Dim ilFieldIndex As Integer
    
    ilAiredEnd = lbcAired.TopIndex + lbcAired.Height \ fgListHtArial825
    If ilAiredEnd > lbcAired.ListCount Then
        ilAiredEnd = lbcAired.ListCount
    End If
    If lbcAired.ListCount <= lbcAired.Height \ fgListHtArial825 Then
        llWidth = lbcAired.Width - 30
    Else
        llWidth = lbcAired.Width - igScrollBarWidth - 30
    End If
    pbcLbcAired.Width = llWidth
    pbcLbcAired.Cls
    llFgColor = pbcLbcAired.ForeColor
    For ilLoop = lbcAired.TopIndex To ilAiredEnd - 1 Step 1
        pbcLbcAired.ForeColor = llFgColor
        If lbcAired.MultiSelect = 0 Then
            If lbcAired.ListIndex = ilLoop Then
                gPaintArea pbcLbcAired, CSng(0), CSng((ilLoop - lbcAired.TopIndex) * fgListHtArial825), CSng(pbcLbcAired.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcAired.ForeColor = vbWhite
            End If
        Else
            If lbcAired.Selected(ilLoop) Then
                gPaintArea pbcLbcAired, CSng(0), CSng((ilLoop - lbcAired.TopIndex) * fgListHtArial825), CSng(pbcLbcAired.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcAired.ForeColor = vbWhite
            End If
        End If
        slStr = lbcAired.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = imLBCtrls To UBound(imListField) - 1 Step 1
            pbcLbcAired.CurrentX = imListField(ilField)
            pbcLbcAired.CurrentY = (ilLoop - lbcAired.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            ilFieldIndex = ilFieldIndex + 1
            gAdjShowLen pbcLbcAired, slStr, imListField(ilField + 1) - imListField(ilField)
            pbcLbcAired.Print slStr
        Next ilField
        pbcLbcAired.ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub pbcLbcMissed_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilMissedEnd As Integer
    Dim ilField As Integer
    Dim llWidth As Long
    Dim slFields(0 To 5) As String
    Dim llFgColor As Long
    Dim ilFieldIndex As Integer
    ilMissedEnd = lbcMissed.TopIndex + lbcMissed.Height \ fgListHtArial825
    If ilMissedEnd > lbcMissed.ListCount Then
        ilMissedEnd = lbcMissed.ListCount
    End If
    If lbcMissed.ListCount <= lbcMissed.Height \ fgListHtArial825 Then
        llWidth = lbcMissed.Width - 30
    Else
        llWidth = lbcMissed.Width - igScrollBarWidth - 30
    End If
    pbcLbcMissed.Width = llWidth
    pbcLbcMissed.Cls
    llFgColor = pbcLbcMissed.ForeColor
    For ilLoop = lbcMissed.TopIndex To ilMissedEnd - 1 Step 1
        pbcLbcMissed.ForeColor = llFgColor
        If lbcMissed.MultiSelect = 0 Then
            If lbcMissed.ListIndex = ilLoop Then
                gPaintArea pbcLbcMissed, CSng(0), CSng((ilLoop - lbcMissed.TopIndex) * fgListHtArial825), CSng(pbcLbcMissed.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcMissed.ForeColor = vbWhite
            End If
        Else
            If lbcMissed.Selected(ilLoop) Then
                gPaintArea pbcLbcMissed, CSng(0), CSng((ilLoop - lbcMissed.TopIndex) * fgListHtArial825), CSng(pbcLbcMissed.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcMissed.ForeColor = vbWhite
            End If
        End If
        slStr = lbcMissed.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = imLBCtrls To UBound(imListField) - 1 Step 1
            pbcLbcMissed.CurrentX = imListField(ilField)
            pbcLbcMissed.CurrentY = (ilLoop - lbcMissed.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            ilFieldIndex = ilFieldIndex + 1
            gAdjShowLen pbcLbcMissed, slStr, imListField(ilField + 1) - imListField(ilField)
            pbcLbcMissed.Print slStr
        Next ilField
        pbcLbcMissed.ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub pbcLbcSpots_Paint(Index As Integer)
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilSpotsEnd As Integer
    Dim ilField As Integer
    Dim llWidth As Long
    Dim slFields(0 To 5) As String
    Dim llFgColor As Long
    Dim ilFieldIndex As Integer
    
    ilSpotsEnd = lbcSpots(Index).TopIndex + lbcSpots(Index).Height \ fgListHtArial825
    If ilSpotsEnd > lbcSpots(Index).ListCount Then
        ilSpotsEnd = lbcSpots(Index).ListCount
    End If
    If lbcSpots(Index).ListCount <= lbcSpots(Index).Height \ fgListHtArial825 Then
        llWidth = lbcSpots(Index).Width - 30
    Else
        llWidth = lbcSpots(Index).Width - igScrollBarWidth - 30
    End If
    pbcLbcSpots(Index).Width = llWidth
    pbcLbcSpots(Index).Cls
    llFgColor = pbcLbcSpots(Index).ForeColor
    For ilLoop = lbcSpots(Index).TopIndex To ilSpotsEnd - 1 Step 1
        pbcLbcSpots(Index).ForeColor = llFgColor
        If lbcSpots(Index).MultiSelect = 0 Then
            If lbcSpots(Index).ListIndex = ilLoop Then
                gPaintArea pbcLbcSpots(Index), CSng(0), CSng((ilLoop - lbcSpots(Index).TopIndex) * fgListHtArial825), CSng(pbcLbcSpots(Index).Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcSpots(Index).ForeColor = vbWhite
            End If
        Else
            If lbcSpots(Index).Selected(ilLoop) Then
                gPaintArea pbcLbcSpots(Index), CSng(0), CSng((ilLoop - lbcSpots(Index).TopIndex) * fgListHtArial825), CSng(pbcLbcSpots(Index).Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcSpots(Index).ForeColor = vbWhite
            End If
        End If
        slStr = lbcSpots(Index).List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = imLBCtrls To UBound(imListField) - 1 Step 1
            pbcLbcSpots(Index).CurrentX = imListField(ilField)
            pbcLbcSpots(Index).CurrentY = (ilLoop - lbcSpots(Index).TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            ilFieldIndex = ilFieldIndex + 1
            gAdjShowLen pbcLbcSpots(Index), slStr, imListField(ilField + 1) - imListField(ilField)
            pbcLbcSpots(Index).Print slStr
        Next ilField
        pbcLbcSpots(Index).ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub pbcLbcVehDp_Paint(Index As Integer)
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilVehDpEnd As Integer
    Dim ilField As Integer
    Dim llWidth As Long
    Dim slFields(0 To 5) As String
    Dim llFgColor As Long
    Dim ilFieldIndex As Integer
    ilVehDpEnd = lbcVehDP(Index).TopIndex + lbcVehDP(Index).Height \ fgListHtArial825
    If ilVehDpEnd > lbcVehDP(Index).ListCount Then
        ilVehDpEnd = lbcVehDP(Index).ListCount
    End If
    If lbcVehDP(Index).ListCount <= lbcVehDP(Index).Height \ fgListHtArial825 Then
        llWidth = lbcVehDP(Index).Width - 30
    Else
        llWidth = lbcVehDP(Index).Width - igScrollBarWidth - 30
    End If
    pbcLbcVehDp(Index).Width = llWidth
    pbcLbcVehDp(Index).Cls
    llFgColor = pbcLbcVehDp(Index).ForeColor
    For ilLoop = lbcVehDP(Index).TopIndex To ilVehDpEnd - 1 Step 1
        pbcLbcVehDp(Index).ForeColor = llFgColor
        If lbcVehDP(Index).MultiSelect = 0 Then
            If lbcVehDP(Index).ListIndex = ilLoop Then
                gPaintArea pbcLbcVehDp(Index), CSng(0), CSng((ilLoop - lbcVehDP(Index).TopIndex) * fgListHtArial825), CSng(pbcLbcVehDp(Index).Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcVehDp(Index).ForeColor = vbWhite
            End If
        Else
            If lbcVehDP(Index).Selected(ilLoop) Then
                gPaintArea pbcLbcVehDp(Index), CSng(0), CSng((ilLoop - lbcVehDP(Index).TopIndex) * fgListHtArial825), CSng(pbcLbcVehDp(Index).Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcVehDp(Index).ForeColor = vbWhite
            End If
        End If
        slStr = lbcVehDP(Index).List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = imLBCtrls To UBound(imListField) - 1 Step 1
            pbcLbcVehDp(Index).CurrentX = imListField(ilField)
            pbcLbcVehDp(Index).CurrentY = (ilLoop - lbcVehDP(Index).TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            ilFieldIndex = ilFieldIndex + 1
            gAdjShowLen pbcLbcVehDp(Index), slStr, imListField(ilField + 1) - imListField(ilField)
            pbcLbcVehDp(Index).Print slStr
        Next ilField
        pbcLbcVehDp(Index).ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub plcTabSelection_BeforeClick(Cancel As Integer)
    If (lgChfMGCode > 0) And ((tgSpf.sGUsePropSys = "Y") Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName)) Then
    Else
        If plcTabSelection.SelectedItem.Index = 1 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub plcTabSelection_Click()
    Dim ilIndex As Integer
    If imMGGen Then
        Exit Sub
    End If
    ilIndex = plcTabSelection.SelectedItem.Index - 1
    If ilIndex = 0 Then   '1 For 1
        plcMMGN.Visible = False
        plc1MG1.Visible = True
        plc1MG1.ZOrder vbBringToFront
    ElseIf ilIndex = 1 Then   'M for N
        plc1MG1.Visible = False
        plcMMGN.Visible = True
        plcMMGN.ZOrder vbBringToFront
    End If
    imTabSelection = ilIndex
End Sub
Private Sub rbcSort_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSort(Index).Value
    'End of coded added
    Dim ilLoop As Integer
    Dim slLine As String
    Dim slType As String
    Dim slAdvtName As String
    Dim slVehName As String
    Dim slName As String
    Dim slDPName As String
    Dim ilRet As Integer
    Dim ilRdf As Integer
    Dim slPrice As String
    If rbcSort(Index).Value Then
        lbcMissed.Clear
        For ilLoop = LBound(tgCntSpot) To UBound(tgCntSpot) - 1 Step 1
            slLine = tgCntSpot(ilLoop).sKey
            ilRet = gParseItem(slLine, 1, "|", slType)
            If slType = "1" Then
                ilRet = gParseItem(slLine, 2, "|", slAdvtName)
                ilRet = gParseItem(slLine, 3, "|", slVehName)
            Else
                ilRet = gParseItem(slLine, 2, "|", slVehName)
                ilRet = gParseItem(slLine, 3, "|", slAdvtName)
            End If
            If Index = 0 Then
                tgCntSpot(ilLoop).sKey = "1" & "|" & slAdvtName & "|" & slVehName
            Else
                tgCntSpot(ilLoop).sKey = "2" & "|" & slVehName & "|" & slAdvtName
            End If
        Next ilLoop
        If UBound(tgCntSpot) - 1 > 0 Then
            ArraySortTyp fnAV(tgCntSpot(), 0), UBound(tgCntSpot), 0, LenB(tgCntSpot(0)), 0, LenB(tgCntSpot(0).sKey), 0
        End If
        For ilLoop = LBound(tgCntSpot) To UBound(tgCntSpot) - 1 Step 1
            slLine = tgCntSpot(ilLoop).sKey
            slPrice = gLongToStrDec(tgCntSpot(ilLoop).lPrice, 2)
            'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
            '    If tgCntSpot(ilLoop).iRdfcode = tgMRdf(ilRdf).iCode Then
                ilRdf = gBinarySearchRdf(tgCntSpot(ilLoop).iRdfCode)
                If ilRdf <> -1 Then
                    slDPName = Trim$(tgMRdf(ilRdf).sName)
            '        Exit For
                End If
            'Next ilRdf
            If Index = 0 Then
                ilRet = gParseItem(slLine, 2, "|", slName)
                'lbcMissed.AddItem gAlignStringByPixel(Trim$(slName) & "|" & Trim$(tgCntSpot(ilLoop).sProduct) & "|" & Trim$(tgCntSpot(ilLoop).sLen) & "|" & slPrice & "|" & " " & "|" & "-1", "|", imListField(), imListFieldChar())
                lbcMissed.AddItem Trim$(slName) & "|" & Trim$(tgCntSpot(ilLoop).sProduct) & "|" & Trim$(tgCntSpot(ilLoop).sLen) & "|" & slPrice & "|" & " " & "|" & "-1"
                ilRet = gParseItem(slLine, 3, "|", slName)
                If tgCntSpot(ilLoop).lAud > 0 Then
                    'lbcMissed.AddItem gAlignStringByPixel("  " & Trim$(slName) & "|" & slDPName & "| " & Str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & Trim$(Str$(tgCntSpot(ilLoop).lAud)) & "|" & " " & "|" & Str$(ilLoop), "|", imListField(), imListFieldChar())
                    lbcMissed.AddItem "  " & Trim$(slName) & "|" & slDPName & "| " & str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & Trim$(str$(tgCntSpot(ilLoop).lAud)) & "|" & " " & "|" & str$(ilLoop)
                Else
                    'lbcMissed.AddItem gAlignStringByPixel("  " & Trim$(slName) & "|" & slDPName & "| " & Str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & " " & "|" & " " & "|" & Str$(ilLoop), "|", imListField(), imListFieldChar())
                    lbcMissed.AddItem "  " & Trim$(slName) & "|" & slDPName & "| " & str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & " " & "|" & " " & "|" & str$(ilLoop)
                End If
            Else
                ilRet = gParseItem(slLine, 2, "|", slName)
                If tgCntSpot(ilLoop).lAud > 0 Then
                    'lbcMissed.AddItem gAlignStringByPixel(Trim$(slName) & "|" & slDPName & "| " & Str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & Trim$(Str$(tgCntSpot(ilLoop).lAud)) & "|" & " " & "|" & Str$(ilLoop), "|", imListField(), imListFieldChar())
                    lbcMissed.AddItem Trim$(slName) & "|" & slDPName & "| " & str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & Trim$(str$(tgCntSpot(ilLoop).lAud)) & "|" & " " & "|" & str$(ilLoop)
                Else
                    'lbcMissed.AddItem gAlignStringByPixel(Trim$(slName) & "|" & slDPName & "| " & Str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & " " & "|" & " " & "|" & Str$(ilLoop), "|", imListField(), imListFieldChar())
                    lbcMissed.AddItem Trim$(slName) & "|" & slDPName & "| " & str$(tgCntSpot(ilLoop).iNoMSpots) & "|" & " " & "|" & " " & "|" & str$(ilLoop)
                End If
                ilRet = gParseItem(slLine, 3, "|", slName)
                'lbcMissed.AddItem gAlignStringByPixel("  " & Trim$(slName) & "|" & Trim$(tgCntSpot(ilLoop).sProduct) & "|" & Trim$(tgCntSpot(ilLoop).sLen) & "|" & slPrice & "|" & " " & "|" & "-1", "|", imListField(), imListFieldChar())
                lbcMissed.AddItem "  " & Trim$(slName) & "|" & Trim$(tgCntSpot(ilLoop).sProduct) & "|" & Trim$(tgCntSpot(ilLoop).sLen) & "|" & slPrice & "|" & " " & "|" & "-1"
            End If
        Next ilLoop
        pbcLbcMissed_Paint
    End If
    mSetCommands
End Sub
Private Sub tmcClick_Timer()
    If imSelectDelay Then
        tmcClick.Enabled = False
        imSelectDelay = False
        Screen.MousePointer = vbHourglass
        Select Case imDelayType
            Case 1  'Spot Week
                mHbcSpotWkChange
            Case 2  'Vehicle DP
                mHbcVehDPWkChange
            Case 3
                mHbcMissedWkChange
        End Select
        Screen.MousePointer = vbDefault
        imDelayType = -1
    End If
End Sub
Private Sub plcAC_Paint()
    plcAC.CurrentX = 0
    plcAC.CurrentY = 0
    plcAC.Print "Advertiser/Competitives"
End Sub
Private Sub plcSort_Paint()
    plcSort.CurrentX = 0
    plcSort.CurrentY = 0
    plcSort.Print "Sort by"
End Sub
Private Sub plcMove_Paint()
    plcMove.CurrentX = 0
    plcMove.CurrentY = 0
    plcMove.Print "Move"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smScreenCaption
End Sub

Private Sub mSetCommands()
    Dim ilFound As Integer
    Dim ilLoop As Integer

    ilFound = False
    For ilLoop = 0 To lbcAired.ListCount - 1 Step 1
        If lbcAired.Selected(ilLoop) Then
            ilFound = True
            Exit For
        End If
    Next ilLoop
    If ilFound Then
        ilFound = False
        For ilLoop = 0 To lbcMissed.ListCount - 1 Step 1
            If lbcMissed.Selected(ilLoop) Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
    End If
    If ilFound Then
        cmc1MG1MG.Enabled = True
        cmc1MG1Outside.Enabled = True
    Else
        cmc1MG1MG.Enabled = False
        cmc1MG1Outside.Enabled = False
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGhfGsfReadRec                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mGhfGsfReadRec(ilVefCode As Integer, llStartDate As Long, llEndDate As Long) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mGhfGsfReadRecErr                                                                     *
'******************************************************************************************

'
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpper As Integer
    Dim llDate As Long

    ReDim tmGsf(0 To 0) As GSF
    ilUpper = 0
    tmGhfSrchKey1.iVefCode = ilVefCode
    ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        tmGsfSrchKey1.lGhfCode = tmGhf.lCode
        tmGsfSrchKey1.iGameNo = 0
        ilRet = btrGetGreaterOrEqual(hmGsf, tmGsf(ilUpper), imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.lCode = tmGsf(ilUpper).lGhfCode)
            gUnpackDateLong tmGsf(ilUpper).iAirDate(0), tmGsf(ilUpper).iAirDate(1), llDate
            If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                ReDim Preserve tmGsf(0 To UBound(tmGsf) + 1) As GSF
                ilUpper = UBound(tmGsf)
            End If
            ilRet = btrGetNext(hmGsf, tmGsf(ilUpper), imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
    Else
        mGhfGsfReadRec = False
        Exit Function
    End If
    mGhfGsfReadRec = True
    Exit Function
mGhfGsfReadRecErr: 'VBC NR
    On Error GoTo 0
    mGhfGsfReadRec = False
    Exit Function
End Function

Private Sub mChkXMid(llSTime As Long, llETime As Long, ilAllowedTimeIndex As Integer, llAllowedSTime As Long, llAllowedETime As Long)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilDay                                                                                 *
'******************************************************************************************

    Dim ilUpper As Integer

    ilUpper = UBound(tgCntSpot)
    If llAllowedSTime <= llAllowedETime Then
        If (llETime >= llAllowedSTime) And (llSTime <= llAllowedETime) And (ilAllowedTimeIndex <= UBound(tgCntSpot(ilUpper).lAllowedSTime)) Then
            tgCntSpot(ilUpper).lAllowedSTime(ilAllowedTimeIndex) = llAllowedSTime
            tgCntSpot(ilUpper).lAllowedETime(ilAllowedTimeIndex) = llAllowedETime
            ilAllowedTimeIndex = ilAllowedTimeIndex + 1
        End If
    Else
        If (llETime >= llAllowedSTime) And (llSTime <= 86400) And (ilAllowedTimeIndex <= UBound(tgCntSpot(ilUpper).lAllowedSTime)) Then
            ilAllowedTimeIndex = ilAllowedTimeIndex + 1
        End If
        If (llETime >= 0) And (llSTime <= llAllowedETime) And (ilAllowedTimeIndex <= UBound(tgCntSpot(ilUpper).lAllowedSTime)) Then
            ilAllowedTimeIndex = ilAllowedTimeIndex + 1
        End If
    End If
End Sub
