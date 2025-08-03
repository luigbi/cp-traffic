VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form SpotAction 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6045
   ClientLeft      =   2310
   ClientTop       =   1980
   ClientWidth     =   9420
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6045
   ScaleWidth      =   9420
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   5790
      TabIndex        =   2
      Top             =   5700
      Width           =   945
   End
   Begin VB.PictureBox pbcAction 
      Height          =   4890
      Left            =   210
      ScaleHeight     =   4830
      ScaleWidth      =   8895
      TabIndex        =   4
      Top             =   675
      Visible         =   0   'False
      Width           =   8955
      Begin VB.CheckBox ckcVehicleNonSports 
         Caption         =   "All Vehicles"
         Height          =   210
         Left            =   2700
         TabIndex        =   76
         Top             =   2055
         Width           =   1665
      End
      Begin VB.Frame frcVehicle 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   240
         Left            =   2715
         TabIndex        =   71
         Top             =   2055
         Width           =   4650
         Begin VB.OptionButton rbcVehicle 
            Caption         =   "Both"
            Height          =   210
            Index           =   2
            Left            =   3510
            TabIndex        =   75
            Top             =   0
            Width           =   780
         End
         Begin VB.OptionButton rbcVehicle 
            Caption         =   "Sports"
            Height          =   210
            Index           =   1
            Left            =   2550
            TabIndex        =   74
            Top             =   0
            Width           =   1035
         End
         Begin VB.OptionButton rbcVehicle 
            Caption         =   "Non-Sports"
            Height          =   210
            Index           =   0
            Left            =   1200
            TabIndex        =   73
            Top             =   0
            Width           =   1305
         End
         Begin VB.Label lacVehcile 
            Caption         =   "All Vehicles"
            Height          =   240
            Left            =   90
            TabIndex        =   72
            Top             =   0
            Width           =   1140
         End
      End
      Begin VB.PictureBox plcTme 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1410
         Left            =   3390
         ScaleHeight     =   1380
         ScaleWidth      =   1095
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   405
         Visible         =   0   'False
         Width           =   1125
         Begin VB.PictureBox pbcTme 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   1305
            Left            =   45
            Picture         =   "SpotAction.frx":0000
            ScaleHeight     =   1305
            ScaleWidth      =   1020
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   45
            Width           =   1020
            Begin VB.Image imcTmeInv 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   480
               Left            =   360
               Picture         =   "SpotAction.frx":0CBE
               Top             =   765
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image imcTmeOutline 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   285
               Left            =   330
               Top             =   270
               Visible         =   0   'False
               Width           =   360
            End
         End
      End
      Begin VB.PictureBox plcCalendar 
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
         Height          =   1770
         Left            =   5055
         ScaleHeight     =   1740
         ScaleWidth      =   1965
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   -15
         Visible         =   0   'False
         Width           =   1995
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
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   45
            Width           =   285
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
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   45
            Width           =   285
         End
         Begin VB.PictureBox pbcCalendar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ClipControls    =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   1440
            Left            =   45
            Picture         =   "SpotAction.frx":0FC8
            ScaleHeight     =   1410
            ScaleWidth      =   1845
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   255
            Width           =   1875
            Begin VB.Label lacDate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   240
               Left            =   510
               TabIndex        =   65
               Top             =   405
               Visible         =   0   'False
               Width           =   300
            End
         End
         Begin VB.Label lacCalName 
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
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   330
            TabIndex        =   68
            Top             =   30
            Width           =   1305
         End
      End
      Begin VB.Frame frcRetainLock 
         Caption         =   "If Avail/Spot Locked"
         Height          =   885
         Left            =   15
         TabIndex        =   28
         Top             =   2280
         Width           =   2400
         Begin VB.CheckBox ckcRemoveSpot 
            Caption         =   "Pre-empt Spot"
            Height          =   210
            Left            =   120
            TabIndex        =   30
            Top             =   570
            Width           =   2160
         End
         Begin VB.CheckBox ckcRemoveLock 
            Caption         =   "Unlock Avail/Spot"
            Height          =   210
            Left            =   120
            TabIndex        =   29
            Top             =   300
            Width           =   1785
         End
      End
      Begin VB.Frame frcLInfo 
         Caption         =   "Lock Information"
         Height          =   645
         Left            =   15
         TabIndex        =   12
         Top             =   15
         Width           =   2400
         Begin VB.OptionButton rbcLInfo 
            Caption         =   "Spot"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   14
            Top             =   270
            Width           =   810
         End
         Begin VB.OptionButton rbcLInfo 
            Caption         =   "Avail"
            Height          =   255
            Index           =   0
            Left            =   135
            TabIndex        =   13
            Top             =   270
            Width           =   810
         End
      End
      Begin VB.Frame frcCInfo 
         Caption         =   "Cancel Information"
         Height          =   1155
         Left            =   15
         TabIndex        =   5
         Top             =   15
         Width           =   2400
         Begin VB.OptionButton rbcCCancel 
            Caption         =   "Canceled to Missed"
            Height          =   255
            Index           =   2
            Left            =   60
            TabIndex        =   8
            Top             =   810
            Width           =   2235
         End
         Begin VB.OptionButton rbcCCancel 
            Caption         =   "Missed to Canceled"
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   7
            Top             =   555
            Width           =   2235
         End
         Begin VB.OptionButton rbcCCancel 
            Caption         =   "Scheduled to Canceled"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   6
            Top             =   270
            Value           =   -1  'True
            Width           =   2235
         End
      End
      Begin VB.CommandButton cmcAdvtCntr 
         Appearance      =   0  'Flat
         Caption         =   "&Get Advertiser"
         Height          =   285
         Left            =   2745
         TabIndex        =   55
         Top             =   4380
         Width           =   1470
      End
      Begin VB.Frame frcPPInfo 
         Caption         =   "Product Protection"
         Height          =   660
         Left            =   6120
         TabIndex        =   56
         Top             =   -15
         Width           =   2700
         Begin VB.OptionButton rbcPPInclExcl 
            Caption         =   "Exclude"
            Height          =   270
            Index           =   1
            Left            =   1380
            TabIndex        =   58
            Top             =   255
            Width           =   1035
         End
         Begin VB.OptionButton rbcPPInclExcl 
            Caption         =   "Include"
            Height          =   270
            Index           =   0
            Left            =   165
            TabIndex        =   57
            Top             =   255
            Width           =   1035
         End
      End
      Begin VB.Frame frcPreemptInfo 
         Caption         =   "Pre-empt Information"
         Height          =   3165
         Left            =   15
         TabIndex        =   31
         Top             =   1680
         Width           =   2415
         Begin VB.CheckBox ckcPIPreemptMG 
            Caption         =   "Pre-empt MG's"
            Height          =   210
            Left            =   120
            TabIndex        =   35
            Top             =   2625
            Value           =   1  'Checked
            Width           =   1785
         End
         Begin VB.CheckBox ckcPIRetainReason 
            Caption         =   "Retain Current Reason"
            Height          =   210
            Left            =   120
            TabIndex        =   36
            Top             =   2895
            Value           =   1  'Checked
            Width           =   2205
         End
         Begin VB.ListBox lbcMissedCancelReason 
            Height          =   1740
            ItemData        =   "SpotAction.frx":3DE2
            Left            =   90
            List            =   "SpotAction.frx":3DE4
            TabIndex        =   32
            Top             =   225
            Width           =   2160
         End
         Begin VB.CheckBox ckcPILock 
            Caption         =   "Lock Pre-empt Range"
            Height          =   210
            Left            =   120
            TabIndex        =   34
            Top             =   2340
            Width           =   2100
         End
         Begin VB.CheckBox ckcPIReSchd 
            Caption         =   "Reschedule"
            Height          =   240
            Left            =   120
            TabIndex        =   33
            Top             =   2055
            Width           =   2130
         End
      End
      Begin VB.Frame frcToDateTime 
         Caption         =   "To Date/Time"
         Height          =   2070
         Left            =   15
         TabIndex        =   37
         Top             =   1710
         Width           =   2400
         Begin VB.TextBox edcMaxFrontLoad 
            Height          =   315
            Left            =   570
            MaxLength       =   2
            TabIndex        =   51
            Text            =   "1"
            Top             =   1635
            Width           =   330
         End
         Begin VB.TextBox edcTDDropDown 
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
            Index           =   0
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   39
            Top             =   270
            Width           =   885
         End
         Begin VB.CommandButton cmcTDDropDown 
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
            Index           =   0
            Left            =   2010
            Picture         =   "SpotAction.frx":3DE6
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   270
            Width           =   195
         End
         Begin VB.TextBox edcTDDropDown 
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
            Index           =   1
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   42
            Top             =   630
            Width           =   885
         End
         Begin VB.CommandButton cmcTDDropDown 
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
            Index           =   1
            Left            =   2010
            Picture         =   "SpotAction.frx":3EE0
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   630
            Width           =   195
         End
         Begin VB.TextBox edcTTDropDown 
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
            Index           =   0
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   45
            Top             =   975
            Width           =   885
         End
         Begin VB.CommandButton cmcTTDropDown 
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
            Index           =   0
            Left            =   2010
            Picture         =   "SpotAction.frx":3FDA
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   975
            Width           =   195
         End
         Begin VB.TextBox edcTTDropDown 
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
            Index           =   1
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   48
            Top             =   1320
            Width           =   885
         End
         Begin VB.CommandButton cmcTTDropDown 
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
            Index           =   1
            Left            =   2010
            Picture         =   "SpotAction.frx":40D4
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   1335
            Width           =   195
         End
         Begin VB.Label lacMaxFrontLoad 
            Appearance      =   0  'Flat
            Caption         =   "Max         Spots/Advt/Veh"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   180
            TabIndex        =   50
            Top             =   1665
            Width           =   2025
         End
         Begin VB.Label lacTStartDate 
            Appearance      =   0  'Flat
            Caption         =   "Start Date"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   180
            TabIndex        =   38
            Top             =   270
            Width           =   900
         End
         Begin VB.Label lacTEndDate 
            Appearance      =   0  'Flat
            Caption         =   "End Date"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   180
            TabIndex        =   41
            Top             =   615
            Width           =   900
         End
         Begin VB.Label lacTStartTime 
            Appearance      =   0  'Flat
            Caption         =   "Start Time"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   180
            TabIndex        =   44
            Top             =   960
            Width           =   900
         End
         Begin VB.Label lacTEndTime 
            Appearance      =   0  'Flat
            Caption         =   "End Time"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   180
            TabIndex        =   47
            Top             =   1305
            Width           =   900
         End
      End
      Begin VB.Frame frcFromDateTime 
         Caption         =   "From Date/Time"
         Height          =   1710
         Left            =   15
         TabIndex        =   15
         Top             =   -15
         Width           =   2400
         Begin VB.CommandButton cmcFTDropDown 
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
            Index           =   1
            Left            =   2010
            Picture         =   "SpotAction.frx":41CE
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1335
            Width           =   195
         End
         Begin VB.TextBox edcFTDropDown 
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
            Index           =   1
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   26
            Top             =   1320
            Width           =   885
         End
         Begin VB.CommandButton cmcFTDropDown 
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
            Index           =   0
            Left            =   2010
            Picture         =   "SpotAction.frx":42C8
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   975
            Width           =   195
         End
         Begin VB.TextBox edcFTDropDown 
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
            Index           =   0
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   23
            Top             =   975
            Width           =   885
         End
         Begin VB.CommandButton cmcFDDropDown 
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
            Index           =   1
            Left            =   2010
            Picture         =   "SpotAction.frx":43C2
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   630
            Width           =   195
         End
         Begin VB.TextBox edcFDDropDown 
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
            Index           =   1
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   20
            Top             =   630
            Width           =   885
         End
         Begin VB.CommandButton cmcFDDropDown 
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
            Index           =   0
            Left            =   2010
            Picture         =   "SpotAction.frx":44BC
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   270
            Width           =   195
         End
         Begin VB.TextBox edcFDDropDown 
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
            Index           =   0
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   17
            Top             =   270
            Width           =   885
         End
         Begin VB.Label lacFEndTime 
            Appearance      =   0  'Flat
            Caption         =   "End Time"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   180
            TabIndex        =   25
            Top             =   1305
            Width           =   900
         End
         Begin VB.Label lacFStartTime 
            Appearance      =   0  'Flat
            Caption         =   "Start Time"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   180
            TabIndex        =   22
            Top             =   960
            Width           =   900
         End
         Begin VB.Label lacFEndDate 
            Appearance      =   0  'Flat
            Caption         =   "End Date"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   180
            TabIndex        =   19
            Top             =   615
            Width           =   900
         End
         Begin VB.Label lacFStartDate 
            Appearance      =   0  'Flat
            Caption         =   "Start Date"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   180
            TabIndex        =   16
            Top             =   270
            Width           =   900
         End
      End
      Begin VB.CheckBox ckcAdvtCntr 
         Caption         =   "All Advertisers"
         Height          =   210
         Left            =   4350
         TabIndex        =   54
         Top             =   4425
         Width           =   1545
      End
      Begin VB.ListBox lbcVehicles 
         Height          =   1950
         ItemData        =   "SpotAction.frx":45B6
         Left            =   2715
         List            =   "SpotAction.frx":45B8
         MultiSelect     =   2  'Extended
         TabIndex        =   52
         Top             =   75
         Width           =   3255
      End
      Begin VB.CheckBox ckcProductProtection 
         Caption         =   "All"
         Height          =   210
         Left            =   6150
         TabIndex        =   61
         Top             =   4425
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.ListBox lbcProductProtection 
         Height          =   4050
         ItemData        =   "SpotAction.frx":45BA
         Left            =   6135
         List            =   "SpotAction.frx":45BC
         MultiSelect     =   2  'Extended
         TabIndex        =   60
         Top             =   255
         Visible         =   0   'False
         Width           =   2700
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAdvtCntr 
         Height          =   1950
         Left            =   2730
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2370
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3440
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
      Begin VB.Frame frcRInfo 
         Caption         =   "Remnant Information"
         Height          =   960
         Left            =   15
         TabIndex        =   9
         Top             =   15
         Width           =   2400
         Begin VB.OptionButton rbcRemnant 
            Caption         =   "Hidden to Missed"
            Height          =   285
            Index           =   1
            Left            =   60
            TabIndex        =   11
            Top             =   570
            Width           =   2055
         End
         Begin VB.OptionButton rbcRemnant 
            Caption         =   "Missed to Hidden"
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   10
            Top             =   285
            Width           =   1740
         End
      End
      Begin VB.Label lacProductProtection 
         Caption         =   "Product Protection"
         Height          =   210
         Left            =   6120
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   1800
      End
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
      Height          =   165
      Left            =   1860
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   62
      Top             =   -15
      Width           =   75
   End
   Begin VB.CommandButton cmcAction 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3540
      TabIndex        =   1
      Top             =   5700
      Visible         =   0   'False
      Width           =   1665
   End
   Begin ComctlLib.TabStrip tbcSelection 
      Height          =   5340
      Left            =   60
      TabIndex        =   3
      Top             =   285
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   9419
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   8
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Pre-Empt"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Reschedule"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Spot Cancel"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Hide Remnant"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Erase Fill"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Lock"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Unlock"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Front Load"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lacScreen 
      Caption         =   "Action"
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   750
   End
End
Attribute VB_Name = "SpotAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SpotAction.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Height adjustment factor

'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBypassAll As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imChgMode As Integer
Dim imBSMode As Integer

'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim imCalType As Integer
Dim imDateBox As Integer    '1=Start Date; 2=End Date; 3=Transmission Date
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar

Dim imTimeBox As Integer
Dim imBypassFocus As Integer

Dim imSetAll As Integer
Dim imAllClicked As Integer

Dim tmUserVehicle() As SORTCODE
Dim smUserVehicleTag As String

Dim tmMissedCancelCode() As SORTCODE
Dim smMissedCancelCodeTag As String

Dim tmProductProtectionCode() As SORTCODE
Dim smProductProtectionCodeTag As String

Dim imLbcVehicleHeight As Integer
Dim imLbcProductProtectionHeight As Integer
Dim imPreemptInfo As Integer

Dim smAction As String

'Spot detail record information
Dim hmSdf As Integer        'Spot detail file handle
Dim tmSdf As SDF            'SDF record image
Dim tmSdfAction() As SDF
Dim tmSdfSrchKey1 As SDFKEY1    'SDF record image (key 3)
Dim tmSdfSrchKey3 As LONGKEY0    'SDF record image (key 3)
Dim imSdfRecLen As Integer  'SDF record length
Dim tmSpotActionInfo() As SPOTACTIONINFO
Dim lmFLChfCode() As Long
'Contract header record information
Dim hmCHF As Integer        'Contract file handle
Dim imCHFRecLen As Integer  'CHF record length
Dim tmChf As CHF            'CHF record image
Dim tmChfSrchKey0 As LONGKEY0

'Contract Line record information
Dim hmClf As Integer        'Contract file handle
Dim imClfRecLen As Integer  'CHF record length
Dim tmClf As CLF            'CHF record image
Dim tmClfSrchKey As CLFKEY0

Dim hmCff As Integer
Dim tmCff As CFF

Dim hmSsf As Integer        'Spot summary file handle
Dim lmSsfDate As Long    'Dates of the days stored into tmSsf
Dim lmSsfRecPos As Long  'Record positions
Dim tmSsf As SSF         'Spot summary for one week (0 index for monday;
Dim imGameNo As Integer

'Game Schedule
Dim hmGsf As Integer
Dim tmGsf As GSF
Dim tmGsfSrchKey3 As GSFKEY3
Dim imGsfRecLen As Integer

Dim hmGhf As Integer

Dim hmSmf As Integer
Dim tmSmf As SMF

Dim hmRlf As Integer

Dim hmSxf As Integer

Dim lmChfCode() As Long
Dim imPPCode() As Integer
Dim imPPAll As Integer
Dim imAdvtCntrAll As Integer

Private imCtrlKey As Integer
Private imShiftKey As Integer

Private imLastAdvtCntrColSorted As Integer
Private imLastAdvtCntrSort As Integer
Private lmLastClickedRow As Long
Private lmScrollTop As Long

Private lmLastLogDate As Long
Private imVpfIndex As Integer

Private imLastTabSelected As Integer

Private bmFirstCallToVpfFind As Boolean

Dim imPreEmptIndex As Integer   'P:1
Dim imRescheduleIndex As Integer    'R:2
Dim imSpotCancelIndex As Integer    'C:3
Dim imHideRemnantIndex As Integer   'H:4 or -1
Dim imEraseFillIndex As Integer     'E:4 or 5
Dim imLockIndex As Integer          '5 or 6
Dim imUnlockIndex As Integer        '6 or 7
Dim imFrontLoadIndex As Integer     'F:7 or 8

Const ADVTINDEX = 0
Const CNTRINDEX = 1
Const SCHDINDEX = 2
Const MISSINDEX = 3
Const SORTINDEX = 4
Const SELECTEDINDEX = 5
Const CHFCODEINDEX = 6







Private Sub ckcAdvtCntr_Click()
    Dim llRow As Long
    Dim slvalue As String
    
    If imBypassAll Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbHourglass
    slvalue = ""
    If ckcAdvtCntr.Value = vbChecked Then
        slvalue = "1"
    End If
    grdAdvtCntr.Redraw = False
    For llRow = grdAdvtCntr.FixedRows To grdAdvtCntr.Rows - 1 Step 1
        If grdAdvtCntr.TextMatrix(llRow, ADVTINDEX) <> "" Then
            grdAdvtCntr.TextMatrix(llRow, SELECTEDINDEX) = slvalue
            mPaintRowColor llRow
        End If
    Next llRow
    grdAdvtCntr.Redraw = True
    If slvalue = "1" Then
        If ckcProductProtection.Value = vbChecked Then
            ckcProductProtection.Value = vbUnchecked
        Else
            ckcProductProtection_Click
        End If
    End If
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbDefault

End Sub

Private Sub ckcAdvtCntr_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcPILock_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcPIPreemptMG_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcPIReSchd_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcProductProtection_Click()
    Dim llValue As Integer
    Dim llRg As Long
    Dim ilValue As Integer
    Dim llRet As Long

    If lbcProductProtection.ListCount <= 0 Then
        Exit Sub
    End If

    ilValue = False
    If ckcProductProtection.Value = vbChecked Then
        ilValue = True
    End If

    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcProductProtection.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcProductProtection.hWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    If ilValue = True Then
        If ckcAdvtCntr.Value = vbChecked Then
            ckcAdvtCntr.Value = vbUnchecked
        Else
            ckcAdvtCntr_Click
        End If
    End If
End Sub

Private Sub ckcProductProtection_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcVehicleNonSports_Click()
    Dim llValue As Integer
    Dim llRg As Long
    Dim ilValue As Integer
    Dim llRet As Long
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer

    If lbcVehicles.ListCount <= 0 Then
        Exit Sub
    End If
    If Not imSetAll Then
        Exit Sub
    End If
    ilValue = False
    If ckcVehicleNonSports.Value = vbChecked Then
        ilValue = True
    End If
    imAllClicked = True
    llRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
    llRet = SendMessageByNum(lbcVehicles.hWnd, LB_SELITEMRANGE, ilValue, llRg)
    imAllClicked = False
    mClearGrid
End Sub

Private Sub ckcVehicleNonSports_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcVehicleSports_Click()
'    Dim llValue As Integer
'    Dim llRg As Long
'    Dim ilValue As Integer
'    Dim llRet As Long
'    Dim ilLoop As Integer
'    Dim slNameCode As String
'    Dim slCode As String
'    Dim ilVefCode As Integer
'    Dim ilIndex As Integer
'    Dim ilRet As Integer
'
'    If lbcVehicles.ListCount <= 0 Then
'        Exit Sub
'    End If
'    If Not imSetAll Then
'        Exit Sub
'    End If
'    ilValue = False
'    If ckcVehicleSports.Value = vbChecked Then
'        ilValue = True
'    End If
'    If (ckcVehicleNonSports.Value = vbChecked) And (ilValue = True) Then
'        imAllClicked = True
'        llRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
'        llRet = SendMessageByNum(lbcVehicles.hWnd, LB_SELITEMRANGE, ilValue, llRg)
'        imAllClicked = False
'    Else
'        imAllClicked = True
'        lbcVehicles.Visible = False
'        For ilLoop = 0 To lbcVehicles.ListCount - 1 Step 1
'            slNameCode = tmUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
'            ilRet = gParseItem(slNameCode, 2, "\", slCode)
'            ilVefCode = Val(slCode)
'            ilIndex = gBinarySearchVef(ilVefCode)
'            If ilIndex <> -1 Then
'                If tgMVef(ilIndex).sType = "G" Then
'                    lbcVehicles.Selected(ilLoop) = ilValue
'                End If
'            End If
'        Next ilLoop
'        lbcVehicles.Visible = True
'        imAllClicked = False
'    End If
'    mClearGrid
End Sub

Private Sub ckcVehicleSports_GotFocus()
'    cmcCancel.Caption = "&Cancel"
'    plcTme.Visible = False
'    plcCalendar.Visible = False
'    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcAction_Click()
    Dim ilRet As Integer
    Select Case tbcSelection.SelectedItem.Index
        Case imPreEmptIndex  'Pre-empt
            ilRet = mPreemptOperation()
        Case imRescheduleIndex  'Reschedule
            ilRet = mRescheduleOperation()
        Case imSpotCancelIndex  'Spot Cancel
            ilRet = mCancelOperation()
        Case imHideRemnantIndex  'Hide Remnant
            ilRet = mRemnantOperation()
        Case imEraseFillIndex 'Erase Fill
            ilRet = mEraseOperation()
        Case imLockIndex  'Lock
            ilRet = mLockOperation()
        Case imUnlockIndex 'Unlock
            ilRet = mUnlockOperation(0)
        Case imFrontLoadIndex  'Front Load
            ilRet = mFrontLoadOperation()
    End Select
    mClearGrid
    If ilRet Then
        cmcCancel.Caption = "&Done"
        igSpotActionReturn = 1
    End If
End Sub

Private Sub cmcAction_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcAdvtCntr_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilAdf As Integer
    Dim llRow As Long
    Dim llCol As Long

    
    mClearGrid
    If lbcVehicles.SelCount <= 0 Then
        MsgBox "No Vehicle(s) has been selected", vbOKOnly + vbApplicationModal, "Get Advertiser"
        Exit Sub
    End If
    If tbcSelection.SelectedItem.Index = imHideRemnantIndex Then
        If Not mCheckFromDateTime("H") Then
            Exit Sub
        End If
    Else
        If Not mCheckFromDateTime() Then
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbHourglass
    ilRet = mBuildSpots("A")
    If tbcSelection.SelectedItem.Index = imPreEmptIndex Then
        mFillAdvtCntrGrid "P"
    ElseIf tbcSelection.SelectedItem.Index = imRescheduleIndex Then
        mFillAdvtCntrGrid "R"
    ElseIf tbcSelection.SelectedItem.Index = imSpotCancelIndex Then
        mFillAdvtCntrGrid "C"
    ElseIf tbcSelection.SelectedItem.Index = imHideRemnantIndex Then
        mFillAdvtCntrGrid "H"
    ElseIf tbcSelection.SelectedItem.Index = imEraseFillIndex Then
        mFillAdvtCntrGrid "E"
    ElseIf tbcSelection.SelectedItem.Index = imFrontLoadIndex Then
        mFillAdvtCntrGrid "F"
    End If
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbDefault
End Sub

Private Sub cmcAdvtCntr_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcFDDropDown_Click(Index As Integer)
    plcCalendar.Visible = Not plcCalendar.Visible
    edcFDDropDown(Index).SelStart = 0
    edcFDDropDown(Index).SelLength = Len(edcFDDropDown(Index).Text)
    edcFDDropDown(Index).SetFocus
End Sub

Private Sub cmcFDDropDown_GotFocus(Index As Integer)
    Dim slStr As String
    
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    If imDateBox <> Index + 1 Then
        plcCalendar.Visible = False
        slStr = edcFDDropDown(Index).Text
        If gValidDate(slStr) Then
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Else
            lacDate.Visible = False
        End If
    End If
    imDateBox = Index + 1
    plcCalendar.Move frcFromDateTime.Left + edcFDDropDown(Index).Left, frcFromDateTime.Top + edcFDDropDown(Index).Top + edcFDDropDown(Index).height
    gCtrlGotFocus ActiveControl

End Sub

Private Sub cmcFTDropDown_Click(Index As Integer)
    plcTme.Visible = Not plcTme.Visible
    edcFTDropDown(Index).SelStart = 0
    edcFTDropDown(Index).SelLength = Len(edcFTDropDown(Index).Text)
    edcFTDropDown(Index).SetFocus
End Sub

Private Sub cmcFTDropDown_GotFocus(Index As Integer)
    cmcCancel.Caption = "&Cancel"
    plcCalendar.Visible = False
    If imTimeBox <> Index + 1 Then
        plcTme.Visible = False
    End If
    imTimeBox = Index + 1
    plcTme.Move frcFromDateTime.Left + edcFTDropDown(Index).Left, frcFromDateTime.Top + edcFTDropDown(Index).Top + edcFTDropDown(Index).height
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcTDDropDown_Click(Index As Integer)
    plcCalendar.Visible = Not plcCalendar.Visible
    edcTDDropDown(Index).SelStart = 0
    edcTDDropDown(Index).SelLength = Len(edcTDDropDown(Index).Text)
    edcTDDropDown(Index).SetFocus
End Sub

Private Sub cmcTDDropDown_GotFocus(Index As Integer)
    Dim slStr As String
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    If imDateBox <> Index + 3 Then
        plcCalendar.Visible = False
        slStr = edcTDDropDown(Index).Text
        If gValidDate(slStr) Then
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Else
            lacDate.Visible = False
        End If
    End If
    imDateBox = Index + 3
    plcCalendar.Move frcToDateTime.Left + edcTDDropDown(Index).Left, frcToDateTime.Top + edcTDDropDown(Index).Top + edcTDDropDown(Index).height
    gCtrlGotFocus ActiveControl

End Sub

Private Sub cmcTTDropDown_Click(Index As Integer)
    plcTme.Visible = Not plcTme.Visible
    edcTTDropDown(Index).SelStart = 0
    edcTTDropDown(Index).SelLength = Len(edcTTDropDown(Index).Text)
    edcTTDropDown(Index).SetFocus
End Sub

Private Sub cmcTTDropDown_GotFocus(Index As Integer)
    cmcCancel.Caption = "&Cancel"
    plcCalendar.Visible = False
    If imTimeBox <> Index + 3 Then
        plcTme.Visible = False
    End If
    imTimeBox = Index + 3
    plcTme.Move frcToDateTime.Left + edcTTDropDown(Index).Left, frcToDateTime.Top + edcTTDropDown(Index).Top + edcTTDropDown(Index).height
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcFDDropDown_Change(Index As Integer)
    Dim slStr As String
    
    slStr = edcFDDropDown(Index).Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    mClearGrid
End Sub

Private Sub edcFDDropDown_GotFocus(Index As Integer)
    Dim slStr As String
    
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    If imDateBox <> Index + 1 Then
        plcCalendar.Visible = False
    End If
    imDateBox = Index + 1
    If (edcFDDropDown(Index).Text = "") And (Index = 1) Then
        slStr = edcFDDropDown(0).Text
        If slStr <> "" Then
            edcFDDropDown(Index).Text = gObtainNextSunday(slStr)
        End If
    End If
    plcCalendar.Move frcFromDateTime.Left + edcFDDropDown(Index).Left, frcFromDateTime.Top + edcFDDropDown(Index).Top + edcFDDropDown(Index).height
    gCtrlGotFocus ActiveControl

End Sub

Private Sub edcFDDropDown_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcFDDropDown_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcFDDropDown(Index).SelLength <> 0 Then    'avoid deleting two characters
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

Private Sub edcFDDropDown_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcFDDropDown(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcFDDropDown(Index).Text = slDate
            End If
        End If
        edcFDDropDown(Index).SelStart = 0
        edcFDDropDown(Index).SelLength = Len(edcFDDropDown(Index).Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcFDDropDown(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcFDDropDown(Index).Text = slDate
            End If
        End If
        edcFDDropDown(Index).SelStart = 0
        edcFDDropDown(Index).SelLength = Len(edcFDDropDown(Index).Text)
    End If
End Sub

Private Sub edcFTDropDown_Change(Index As Integer)
    mClearGrid
End Sub

Private Sub edcFTDropDown_GotFocus(Index As Integer)
    cmcCancel.Caption = "&Cancel"
    plcCalendar.Visible = False
    If imTimeBox <> Index + 1 Then
        plcTme.Visible = False
    End If
    imTimeBox = Index + 1
    plcTme.Move frcFromDateTime.Left + edcFTDropDown(Index).Left, frcFromDateTime.Top + edcFTDropDown(Index).Top + edcFTDropDown(Index).height
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub

Private Sub edcFTDropDown_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcFTDropDown_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilKey As Integer
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
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If ActiveControl.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub edcFTDropDown_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcTme.Visible = Not plcTme.Visible
        End If
        edcFTDropDown(Index).SelStart = 0
        edcFTDropDown(Index).SelLength = Len(edcFTDropDown(Index).Text)
    End If
End Sub

Private Sub edcTDDropDown_Change(Index As Integer)
    Dim slStr As String
    
    slStr = edcTDDropDown(Index).Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint

End Sub

Private Sub edcTDDropDown_GotFocus(Index As Integer)
    Dim slStr As String
    
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    If imDateBox <> Index + 3 Then
        plcCalendar.Visible = False
    End If
    imDateBox = Index + 3
    If (edcTDDropDown(Index).Text = "") And (Index = 1) Then
        slStr = edcFDDropDown(0).Text
        If slStr <> "" Then
            slStr = gDecOneDay(slStr)
            If gDateValue(edcTDDropDown(0).Text) < gDateValue(slStr) Then
                edcTDDropDown(Index).Text = slStr
            End If
        End If
    End If
    plcCalendar.Move frcToDateTime.Left + edcTDDropDown(Index).Left, frcToDateTime.Top + edcTDDropDown(Index).Top + edcTDDropDown(Index).height
    gCtrlGotFocus ActiveControl

End Sub

Private Sub edcTDDropDown_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcTDDropDown_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcTDDropDown(Index).SelLength <> 0 Then    'avoid deleting two characters
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

Private Sub edcTDDropDown_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcTDDropDown(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcTDDropDown(Index).Text = slDate
            End If
        End If
        edcTDDropDown(Index).SelStart = 0
        edcTDDropDown(Index).SelLength = Len(edcTDDropDown(Index).Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcTDDropDown(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcTDDropDown(Index).Text = slDate
            End If
        End If
        edcTDDropDown(Index).SelStart = 0
        edcTDDropDown(Index).SelLength = Len(edcTDDropDown(Index).Text)
    End If
End Sub

Private Sub edcTTDropDown_GotFocus(Index As Integer)
    cmcCancel.Caption = "&Cancel"
    plcCalendar.Visible = False
    If imTimeBox <> Index + 3 Then
        plcTme.Visible = False
    End If
    imTimeBox = Index + 3
    plcTme.Move frcToDateTime.Left + edcTTDropDown(Index).Left, frcToDateTime.Top + edcTTDropDown(Index).Top + edcTTDropDown(Index).height
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub

Private Sub edcTTDropDown_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcTTDropDown_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilKey As Integer
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
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If ActiveControl.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub edcTTDropDown_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcTme.Visible = Not plcTme.Visible
        End If
        edcTTDropDown(Index).SelStart = 0
        edcTTDropDown(Index).SelLength = Len(edcTTDropDown(Index).Text)
    End If
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
'    gShowBranner
    Me.KeyPreview = True
    SpotAction.Refresh
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
    End If
End Sub

Private Sub Form_Load()
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.height) / (480 * 15 / Me.height))) / 100) / Me.height
        Me.height = (lgPercentAdjH * ((Screen.height) / (480 * 15 / Me.height))) / 100
    End If
    mInit
    If imTerminate Then
        'cmcCancel_Click
    End If
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
    Dim slDate As String
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbHourglass
    imTerminate = False
    imFirstActivate = True
    imChgMode = False
    imSetAll = True
    imAllClicked = False
    imDateBox = -1
    imBypassFocus = False
    imBypassAll = False
    bmFirstCallToVpfFind = True

    ReDim tmSpotActionInfo(0 To 0) As SPOTACTIONINFO
    lmLastClickedRow = -1
    lmScrollTop = grdAdvtCntr.FixedRows
    imLastAdvtCntrColSorted = -1
    imLastAdvtCntrSort = -1
    
    imLastTabSelected = -1

    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbHourglass
    'SpotAction.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    pbcAction.BorderStyle = 0
    pbcAction.Visible = False
    gCenterStdAlone SpotAction
    SpotAction.Top = SpotAction.Top - 150
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbHourglass
    
    hmSdf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", SpotAction
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", SpotAction
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", SpotAction
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", SpotAction
    On Error GoTo 0
    
    hmSsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ssf.Btr)", SpotAction
    On Error GoTo 0
    
    hmGsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Gsf.Btr)", SpotAction
    On Error GoTo 0
    imGsfRecLen = Len(tmGsf)
    
    hmGhf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ghf.Btr)", SpotAction
    On Error GoTo 0
    
    hmSmf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf.Btr)", SpotAction
    On Error GoTo 0
    
    hmRlf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRlf, "", sgDBPath & "Rlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rlf.Btr)", SpotAction
    On Error GoTo 0
    
    hmSxf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSxf, "", sgDBPath & "Sxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sxf.Btr)", SpotAction
    On Error GoTo 0
    
    mInitBox
    If sgSpotActionDefaultDate <> "" Then
        slDate = sgSpotActionDefaultDate
        If gDateValue(slDate) < gDateValue(Format$(gNow(), "m/d/yy")) Then
            slDate = Format$(gNow(), "m/d/yy")
            If ((Asc(tgSpf.sUsingFeatures4) And ALLOWMOVEONTODAY) = ALLOWMOVEONTODAY) Then
                edcFDDropDown(0).Text = slDate
            Else
                edcFDDropDown(0).Text = Format$(gDateValue(slDate) + 1, "m/d/yy")
            End If
        Else
            edcFDDropDown(0).Text = slDate
        End If
    Else
        slDate = Format$(gNow(), "m/d/yy")
        If ((Asc(tgSpf.sUsingFeatures4) And ALLOWMOVEONTODAY) = ALLOWMOVEONTODAY) Then
            edcFDDropDown(0).Text = slDate
        Else
            edcFDDropDown(0).Text = Format$(gDateValue(slDate) + 1, "m/d/yy")
        End If
    End If
    edcFDDropDown(1).Text = gObtainNextSunday(edcFDDropDown(0).Text)
    edcFTDropDown(0).Text = "12M"
    edcFTDropDown(1).Text = "12M"
    
    
    edcTTDropDown(0).Text = "12M"
    edcTTDropDown(1).Text = "12M"
    Screen.MousePointer = vbDefault
    
    If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) <> USINGSPORTS) Then
        mVehPop
    End If
    If imTerminate Then
        Exit Sub
    End If
    If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) <> USINGSPORTS) Then
        'ckcVehicleSports.Visible = False
        'ckcVehicleNonSports.Caption = "All Vehicles"
        ckcVehicleNonSports.Visible = True
        frcVehicle.Visible = False
    End If
    mMissedCancelPop
    If imTerminate Then
        Exit Sub
    End If
    mProductProtectionPop
    If imTerminate Then
        Exit Sub
    End If
    If tgSpf.sSchdRemnant <> "Y" Then
        tbcSelection.Tabs.Remove (4)
        imPreEmptIndex = 1
        imRescheduleIndex = 2
        imSpotCancelIndex = 3
        imHideRemnantIndex = -100
        imEraseFillIndex = 4
        imLockIndex = 5
        imUnlockIndex = 6
        imFrontLoadIndex = 7
    Else
        imPreEmptIndex = 1
        imRescheduleIndex = 2
        imSpotCancelIndex = 3
        imHideRemnantIndex = 4
        imEraseFillIndex = 5
        imLockIndex = 6
        imUnlockIndex = 7
        imFrontLoadIndex = 8
    End If
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbDefault
    Exit Sub
         
    On Error GoTo 0
    imTerminate = True
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
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
    Dim ilLoop
    Dim ilCol As Integer
    Dim llRet As Long

    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    plcCalendar.Move frcFromDateTime.Left + edcFDDropDown(0).Left, frcFromDateTime.Top + edcFDDropDown(0).Top + edcFDDropDown(0).height

    imBSMode = False
    imCalType = 0   'Standard

    cmcCancel.Top = SpotAction.height - cmcCancel.height - 120
    cmcAction.Top = cmcCancel.Top
    ckcProductProtection.Move 15, 0
    frcCInfo.Move 15, 0
    frcFromDateTime.Move 15, 0
    frcToDateTime.Move 15, frcFromDateTime.Top + frcFromDateTime.height + 60
    frcPreemptInfo.Move 15, frcToDateTime.Top + frcToDateTime.height + 60
    tbcSelection.Top = 285
    tbcSelection.Left = 60
    tbcSelection.Width = fmAdjFactorW * tbcSelection.Width
    tbcSelection.height = cmcCancel.Top - tbcSelection.Top - 120
    pbcAction.Width = fmAdjFactorW * pbcAction.Width
    pbcAction.height = tbcSelection.height - 510
    pbcAction.Left = tbcSelection.Left + 180
    pbcAction.Top = tbcSelection.Top + 400
    lbcProductProtection.Width = fmAdjFactorW * lbcProductProtection.Width
    lbcProductProtection.Left = pbcAction.Width - lbcProductProtection.Width - 180
    ckcProductProtection.Top = pbcAction.height - ckcAdvtCntr.height - 120
    ckcProductProtection.Left = lbcProductProtection.Left
    lacProductProtection.Move lbcProductProtection.Left, lbcVehicles.Top
    frcPPInfo.Move lbcProductProtection.Left, lbcVehicles.Top
    lbcProductProtection.Top = lacProductProtection.Top + lacProductProtection.height
    lbcProductProtection.height = ckcProductProtection.Top - lbcProductProtection.Top - 60
    lbcVehicles.Width = lbcProductProtection.Left - lbcVehicles.Left - 120
    lbcVehicles.height = lbcProductProtection.height / 2 - frcVehicle.height    'ckcVehicleNonSports.Height
    ckcVehicleNonSports.Left = lbcVehicles.Left
    ckcVehicleNonSports.Top = lbcVehicles.Top + lbcVehicles.height + 90
    'ckcVehicleSports.Top = ckcVehicleNonSports.Top
    'ckcVehicleSports.Left = ckcVehicleNonSports.Left + ckcVehicleNonSports.Width + 180
    frcVehicle.Left = lbcVehicles.Left
    frcVehicle.Top = lbcVehicles.Top + lbcVehicles.height + 90
    grdAdvtCntr.Width = lbcVehicles.Width
    grdAdvtCntr.Left = lbcVehicles.Left
    grdAdvtCntr.Top = frcVehicle.Top + frcVehicle.height + 90   'NonSports.Top + ckcVehicleNonSports.Height + 90
    ckcAdvtCntr.Top = pbcAction.height - ckcAdvtCntr.height - 120
    grdAdvtCntr.height = ckcAdvtCntr.Top - grdAdvtCntr.Top - 60
    cmcAdvtCntr.Move grdAdvtCntr.Left, ckcAdvtCntr.Top
    ckcAdvtCntr.Left = cmcAdvtCntr.Left + cmcAdvtCntr.Width + 120
    
    imLbcVehicleHeight = lbcVehicles.height
    imLbcProductProtectionHeight = lbcProductProtection.height
    imPreemptInfo = 2910    '2615    'frcPreemptInfo.Height
    cmcCancel.Left = SpotAction.Width / 2 - cmcCancel.Width / 2

    mSetGridColumns
    mSetGridTitles ""
    gGrid_IntegralHeight grdAdvtCntr, fgBoxGridH + 15
    gGrid_FillWithRows grdAdvtCntr, fgBoxGridH + 15

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
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbDefault
       
    igManUnload = YES
    Unload SpotAction
    igManUnload = NO
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gGetSchParameters
    Erase lmChfCode
    Erase imPPCode
    Erase tmSpotActionInfo
    Erase tmUserVehicle
    Erase tmMissedCancelCode
    Erase tmProductProtectionCode
    Erase tmSdfAction
    Erase lmFLChfCode
    btrDestroy hmSdf
    btrDestroy hmCHF
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmSsf
    btrDestroy hmGsf
    btrDestroy hmGhf
    btrDestroy hmSmf
    btrDestroy hmRlf
    btrDestroy hmSxf
    
    Set SpotAction = Nothing   'Remove data segment

End Sub

Private Sub frcCInfo_Click()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub frcFromDateTime_Click()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub frcLInfo_Click()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub frcPPInfo_Click()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub frcPreemptInfo_Click()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub frcToDateTime_Click()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub grdAdvtCntr_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub lacFEndDate_Click()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub lacFEndTime_Click()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub lacFStartDate_Click()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub lacFStartTime_Click()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub lacScreen_Click()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub lacTEndDate_Click()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub lacTEndTime_Click()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub lacTStartDate_Click()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub lacTStartTime_Click()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub



Private Sub lbcMissedCancelReason_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub lbcMissedCancelReason_Scroll()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub lbcProductProtection_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcProductProtection.Value = vbUnchecked
        imSetAll = True
    End If
    If ckcAdvtCntr.Value = vbChecked Then
        ckcAdvtCntr.Value = vbUnchecked
    Else
        ckcAdvtCntr_Click
    End If
End Sub

Private Sub lbcProductProtection_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub lbcProductProtection_Scroll()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub lbcVehicles_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcVehicleNonSports.Value = vbUnchecked
        'ckcVehicleSports.Value = vbUnchecked
        rbcVehicle(0).Value = False
        rbcVehicle(1).Value = False
        rbcVehicle(2).Value = False
        imSetAll = True
    End If
    mClearGrid
End Sub

Private Sub lbcVehicles_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub lbcVehicles_Scroll()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub


Private Sub pbcAction_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub pbcClickFocus_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub pbcTme_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeInv.Visible = True
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

Private Sub pbcTme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Select Case ilRowNo
                        Case 1
                            Select Case ilColNo
                                Case 1
                                    slKey = "7"
                                Case 2
                                    slKey = "8"
                                Case 3
                                    slKey = "9"
                            End Select
                        Case 2
                            Select Case ilColNo
                                Case 1
                                    slKey = "4"
                                Case 2
                                    slKey = "5"
                                Case 3
                                    slKey = "6"
                            End Select
                        Case 3
                            Select Case ilColNo
                                Case 1
                                    slKey = "1"
                                Case 2
                                    slKey = "2"
                                Case 3
                                    slKey = "3"
                            End Select
                        Case 4
                            Select Case ilColNo
                                Case 1
                                    slKey = "0"
                                Case 2
                                    slKey = "00"
                            End Select
                        Case 5
                            Select Case ilColNo
                                Case 1
                                    slKey = ":"
                                Case 2
                                    slKey = "AM"
                                Case 3
                                    slKey = "PM"
                            End Select
                    End Select
                    If imTimeBox = 1 Then
                        imBypassFocus = True    'Don't change select text
                        edcFTDropDown(0).SetFocus
                        gSendKeys edcFTDropDown(0), slKey
                    ElseIf imTimeBox = 2 Then
                        imBypassFocus = True    'Don't change select text
                        edcFTDropDown(1).SetFocus
                        gSendKeys edcFTDropDown(1), slKey
                    ElseIf imTimeBox = 3 Then
                        imBypassFocus = True    'Don't change select text
                        edcTTDropDown(0).SetFocus
                        gSendKeys edcTTDropDown(0), slKey
                    ElseIf imTimeBox = 4 Then
                        imBypassFocus = True    'Don't change select text
                        edcTTDropDown(1).SetFocus
                        gSendKeys edcTTDropDown(1), slKey
                    End If
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

Private Sub rbcCCancel_Click(Index As Integer)
    If rbcCCancel(0).Value Then
        frcRetainLock.Move frcLInfo.Left, frcFromDateTime.Top + frcFromDateTime.height + 60
        frcPreemptInfo.Move frcLInfo.Left, frcRetainLock.Top + frcRetainLock.height + 60, frcPreemptInfo.Width, 2055
        frcRetainLock.Visible = True
    Else
        frcPreemptInfo.Move frcLInfo.Left, frcFromDateTime.Top + frcFromDateTime.height + 60, frcPreemptInfo.Width, 2340
        frcRetainLock.Visible = False
    End If
    If rbcCCancel(2).Value Then
        frcPreemptInfo.Caption = "Missed Reason"
    Else
        frcPreemptInfo.Caption = "Cancel Reason"
    End If
    mClearGrid
    mSetGridTitles "C"
    mFillAdvtCntrGrid "C"
End Sub

Private Sub rbcCCancel_GotFocus(Index As Integer)
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub rbcLInfo_GotFocus(Index As Integer)
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub rbcPPInclExcl_GotFocus(Index As Integer)
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Sub rbcRemnant_Click(Index As Integer)
    mClearGrid
    mSetGridTitles "H"
    mFillAdvtCntrGrid "H"
End Sub

Private Sub rbcVehicle_Click(Index As Integer)
    Dim llValue As Integer
    Dim llRg As Long
    Dim ilValue As Integer
    Dim llRet As Long
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer
    
    If lbcVehicles.ListCount <= 0 Then
        Exit Sub
    End If
    If Not imSetAll Then
        Exit Sub
    End If
    
    If rbcVehicle(Index).Value Then
        Screen.MousePointer = vbHourglass
        gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbHourglass
        If Index <> 2 Then
            ilValue = False
            imAllClicked = True
            llRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcVehicles.hWnd, LB_SELITEMRANGE, ilValue, llRg)
            imAllClicked = False
        End If
        If Index = 0 Then
            imAllClicked = True
            lbcVehicles.Visible = False
            DoEvents
            ilValue = True
            For ilLoop = 0 To lbcVehicles.ListCount - 1 Step 1
                slNameCode = tmUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                ilIndex = gBinarySearchVef(ilVefCode)
                If ilIndex <> -1 Then
                    If tgMVef(ilIndex).sType <> "G" Then
                        lbcVehicles.Selected(ilLoop) = ilValue
                    End If
                End If
            Next ilLoop
            lbcVehicles.Visible = True
            DoEvents
            imAllClicked = False
        ElseIf Index = 1 Then
            imAllClicked = True
            lbcVehicles.Visible = False
            DoEvents
            ilValue = True
            For ilLoop = 0 To lbcVehicles.ListCount - 1 Step 1
                slNameCode = tmUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                ilIndex = gBinarySearchVef(ilVefCode)
                If ilIndex <> -1 Then
                    If tgMVef(ilIndex).sType = "G" Then
                        lbcVehicles.Selected(ilLoop) = ilValue
                    End If
                End If
            Next ilLoop
            lbcVehicles.Visible = True
            DoEvents
            imAllClicked = False
        Else
            ilValue = True
            imAllClicked = True
            llRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcVehicles.hWnd, LB_SELITEMRANGE, ilValue, llRg)
            imAllClicked = False
        End If
        mClearGrid
        Screen.MousePointer = vbDefault
        gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbDefault
    End If
End Sub

Private Sub rbcVehicle_GotFocus(Index As Integer)
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub tbcSelection_Click()
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbHourglass
    pbcAction.Visible = False
    frcLInfo.Visible = False
    frcCInfo.Visible = False
    frcFromDateTime.Visible = False
    frcToDateTime.Visible = False
    frcPreemptInfo.Visible = False
    lbcVehicles.Visible = False
    ckcVehicleNonSports.Visible = False
    'ckcVehicleSports.Visible = False
    frcVehicle.Visible = False
    rbcVehicle(1).Enabled = True
    rbcVehicle(2).Enabled = True
    grdAdvtCntr.Visible = False
    ckcAdvtCntr.Visible = False
    cmcAdvtCntr.Visible = False
    lacProductProtection.Visible = False
    frcPPInfo.Visible = False
    lbcProductProtection.Visible = False
    ckcProductProtection.Visible = False
    plcCalendar.Visible = False
    plcTme.Visible = False
    frcRetainLock.Visible = False
    frcRInfo.Visible = False
    
    Select Case tbcSelection.SelectedItem.Index
        Case imPreEmptIndex  'Pre-empt
            frcFromDateTime.Move frcLInfo.Left, frcLInfo.Top
            frcRetainLock.Move frcLInfo.Left, frcFromDateTime.Top + frcFromDateTime.height + 60
            frcPreemptInfo.Move frcLInfo.Left, frcRetainLock.Top + frcRetainLock.height + 60
            frcPreemptInfo.height = imPreemptInfo
            frcPreemptInfo.Caption = "Pre-empt Information"
            ckcPIReSchd.Visible = True
            ckcPIRetainReason.Visible = False
            lbcVehicles.height = imLbcVehicleHeight
            ckcVehicleNonSports.Top = lbcVehicles.Top + lbcVehicles.height + 90
            'ckcVehicleSports.Top = ckcVehicleNonSports.Top
            frcVehicle.Top = lbcVehicles.Top + lbcVehicles.height + 90
            lacProductProtection.Top = lbcVehicles.Top
            lbcProductProtection.Top = lacProductProtection.Top + lacProductProtection.height
            lbcProductProtection.height = imLbcProductProtectionHeight
            frcFromDateTime.Visible = True
            frcRetainLock.Visible = True
            frcPreemptInfo.Visible = True
            lbcVehicles.Visible = True
            'ckcVehicleNonSports.Visible = True
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) Then
                frcVehicle.Visible = True
            Else
                ckcVehicleNonSports.Visible = True
            End If
            grdAdvtCntr.Visible = True
            ckcAdvtCntr.Visible = True
            cmcAdvtCntr.Visible = True
            lacProductProtection.Visible = True
            lbcProductProtection.Visible = True
            'ckcProductProtection.Visible = True
            cmcAction.Caption = "Pre-empt"
            ckcRemoveSpot.Caption = "Pre-empt Spot"
            mClearGrid
            mSetGridTitles "P"
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) <> USINGSPORTS) Or (imLastTabSelected <> imFrontLoadIndex) Then
                mFillAdvtCntrGrid "P"
            End If
        Case imRescheduleIndex  'Reschedule
            frcFromDateTime.Move frcLInfo.Left, frcLInfo.Top
            lbcVehicles.height = imLbcVehicleHeight
            ckcVehicleNonSports.Top = lbcVehicles.Top + lbcVehicles.height + 90
            'ckcVehicleSports.Top = ckcVehicleNonSports.Top
            frcVehicle.Top = lbcVehicles.Top + lbcVehicles.height + 90
            lbcProductProtection.Top = frcPPInfo.Top + frcPPInfo.height + 60
            lbcProductProtection.height = imLbcProductProtectionHeight - frcPPInfo.height
            frcFromDateTime.Visible = True
            frcPPInfo.Visible = True
            lbcVehicles.Visible = True
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) Then
                frcVehicle.Visible = True
            Else
                ckcVehicleNonSports.Visible = True
            End If
            grdAdvtCntr.Visible = True
            ckcAdvtCntr.Visible = True
            cmcAdvtCntr.Visible = True
            lbcProductProtection.Visible = True
            'ckcProductProtection.Visible = True
            cmcAction.Caption = "Reschedule"
            mClearGrid
            mSetGridTitles "R"
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) <> USINGSPORTS) Or (imLastTabSelected <> imFrontLoadIndex) Then
                mFillAdvtCntrGrid "R"
            End If
        Case imSpotCancelIndex  'Spot Cancel
            frcFromDateTime.Move frcCInfo.Left, frcCInfo.Top + frcCInfo.height + 60
            If rbcCCancel(0).Value Then
                frcRetainLock.Move frcLInfo.Left, frcFromDateTime.Top + frcFromDateTime.height + 60
                frcPreemptInfo.Move frcLInfo.Left, frcRetainLock.Top + frcRetainLock.height + 60, frcPreemptInfo.Width, 2055
            Else
                frcPreemptInfo.Move frcLInfo.Left, frcFromDateTime.Top + frcFromDateTime.height + 60, frcPreemptInfo.Width, 2340
            End If
            ckcPIReSchd.Visible = False
            ckcPIRetainReason.Top = ckcPIReSchd.Top
            ckcPIRetainReason.Left = ckcPIReSchd.Left
            ckcPIRetainReason.Visible = True
            If rbcCCancel(2).Value Then
                frcPreemptInfo.Caption = "Missed Reason"
            Else
                frcPreemptInfo.Caption = "Cancel Reason"
            End If
            lbcVehicles.height = imLbcVehicleHeight
            ckcVehicleNonSports.Top = lbcVehicles.Top + lbcVehicles.height + 90
            'ckcVehicleSports.Top = ckcVehicleNonSports.Top
            frcVehicle.Top = lbcVehicles.Top + lbcVehicles.height + 90
            lacProductProtection.Top = lbcVehicles.Top
            lbcProductProtection.Top = lacProductProtection.Top + lacProductProtection.height
            lbcProductProtection.height = imLbcProductProtectionHeight
            frcCInfo.Visible = True
            frcFromDateTime.Visible = True
            If rbcCCancel(0).Value Then
                frcRetainLock.Visible = True
            End If
            frcPreemptInfo.Visible = True
            lbcVehicles.Visible = True
            'ckcVehicleNonSports.Visible = True
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) Then
                frcVehicle.Visible = True
            Else
                ckcVehicleNonSports.Visible = True
            End If
            grdAdvtCntr.Visible = True
            ckcAdvtCntr.Visible = True
            cmcAdvtCntr.Visible = True
            lacProductProtection.Visible = True
            lbcProductProtection.Visible = True
            'ckcProductProtection.Visible = True
            cmcAction.Caption = "Spot Cancel"
            ckcRemoveSpot.Caption = "Cancel Spot"
            mClearGrid
            mSetGridTitles "C"
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) <> USINGSPORTS) Or (imLastTabSelected <> imFrontLoadIndex) Then
                mFillAdvtCntrGrid "C"
            End If
        Case imHideRemnantIndex  'Hide Remnant
            frcRInfo.Move frcCInfo.Left, frcCInfo.Top
            frcFromDateTime.Move frcRInfo.Left, frcRInfo.Top + frcRInfo.height + 60
            lbcVehicles.height = imLbcVehicleHeight
            ckcVehicleNonSports.Top = lbcVehicles.Top + lbcVehicles.height + 90
            'ckcVehicleSports.Top = ckcVehicleNonSports.Top
            frcVehicle.Top = lbcVehicles.Top + lbcVehicles.height + 90
            lacProductProtection.Top = lbcVehicles.Top
            lbcProductProtection.Top = lacProductProtection.Top + lacProductProtection.height
            lbcProductProtection.height = imLbcProductProtectionHeight
            
            
            frcRInfo.Visible = True
            frcFromDateTime.Visible = True
            lbcVehicles.Visible = True
            'ckcVehicleNonSports.Visible = True
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) Then
                frcVehicle.Visible = True
            Else
                ckcVehicleNonSports.Visible = True
            End If
            grdAdvtCntr.Visible = True
            ckcAdvtCntr.Visible = True
            cmcAdvtCntr.Visible = True
            lacProductProtection.Visible = True
            lbcProductProtection.Visible = True
            'ckcProductProtection.Visible = True
            cmcAction.Caption = "Hide Remnant"
            mClearGrid
            mSetGridTitles "H"
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) <> USINGSPORTS) Or (imLastTabSelected <> imFrontLoadIndex) Then
                mFillAdvtCntrGrid "H"
            End If
            
        Case imEraseFillIndex  'Erase Fill
            frcFromDateTime.Move frcLInfo.Left, frcLInfo.Top
            frcRetainLock.Move frcLInfo.Left, frcFromDateTime.Top + frcFromDateTime.height + 60
            lbcVehicles.height = imLbcVehicleHeight 'imLbcProductProtectionHeight
            ckcVehicleNonSports.Top = lbcVehicles.Top + lbcVehicles.height + 90
            'ckcVehicleSports.Top = ckcVehicleNonSports.Top
            frcVehicle.Top = lbcVehicles.Top + lbcVehicles.height + 90
            frcFromDateTime.Visible = True
            frcRetainLock.Visible = True
            lbcVehicles.Visible = True
            'ckcVehicleNonSports.Visible = True
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) Then
                frcVehicle.Visible = True
            Else
                ckcVehicleNonSports.Visible = True
            End If
            grdAdvtCntr.Visible = True
            ckcAdvtCntr.Visible = True
            cmcAdvtCntr.Visible = True
            cmcAction.Caption = "Erase Fill"
            ckcRemoveSpot.Caption = "Erase Fill Spot"
            mClearGrid
            mSetGridTitles "E"
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) <> USINGSPORTS) Or (imLastTabSelected <> imFrontLoadIndex) Then
                mFillAdvtCntrGrid "E"
            End If
        Case imLockIndex  'Lock
            frcLInfo.Caption = "Lock Information"
            frcFromDateTime.Move frcLInfo.Left, frcLInfo.Top + frcLInfo.height + 60
            lbcVehicles.height = imLbcProductProtectionHeight
            ckcVehicleNonSports.Top = lbcVehicles.Top + lbcVehicles.height + 90
            'ckcVehicleSports.Top = ckcVehicleNonSports.Top
            frcVehicle.Top = lbcVehicles.Top + lbcVehicles.height + 90
            frcLInfo.Visible = True
            frcFromDateTime.Visible = True
            lbcVehicles.Visible = True
            'ckcVehicleNonSports.Visible = True
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) Then
                frcVehicle.Visible = True
            Else
                ckcVehicleNonSports.Visible = True
            End If
            cmcAction.Caption = "Lock"
        Case imUnlockIndex  'Unlock
            frcLInfo.Caption = "Unlock Information"
            frcFromDateTime.Move frcLInfo.Left, frcLInfo.Top + frcLInfo.height + 60
            lbcVehicles.height = imLbcProductProtectionHeight
            ckcVehicleNonSports.Top = lbcVehicles.Top + lbcVehicles.height + 90
            'ckcVehicleSports.Top = ckcVehicleNonSports.Top
            frcVehicle.Top = lbcVehicles.Top + lbcVehicles.height + 90
            frcLInfo.Visible = True
            frcFromDateTime.Visible = True
            lbcVehicles.Visible = True
            'ckcVehicleNonSports.Visible = True
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) Then
                frcVehicle.Visible = True
            Else
                ckcVehicleNonSports.Visible = True
            End If
            cmcAction.Caption = "Unlock"
        Case imFrontLoadIndex  'Front Load
            frcFromDateTime.Move frcLInfo.Left, frcLInfo.Top
            frcRetainLock.Move frcLInfo.Left, frcFromDateTime.Top + frcFromDateTime.height + 60
            frcToDateTime.Move frcLInfo.Left, frcRetainLock.Top + frcRetainLock.height + 60
            lbcVehicles.height = imLbcVehicleHeight
            ckcVehicleNonSports.Top = lbcVehicles.Top + lbcVehicles.height + 90
            'ckcVehicleSports.Top = ckcVehicleNonSports.Top
            frcVehicle.Top = lbcVehicles.Top + lbcVehicles.height + 90
            lacProductProtection.Top = lbcVehicles.Top
            lbcProductProtection.Top = lacProductProtection.Top + lacProductProtection.height
            lbcProductProtection.height = imLbcProductProtectionHeight
            frcFromDateTime.Visible = True
            frcRetainLock.Visible = True
            frcToDateTime.Visible = True
            lbcVehicles.Visible = True
            'ckcVehicleNonSports.Visible = True
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) Then
                frcVehicle.Visible = True
                rbcVehicle(1).Enabled = False
                rbcVehicle(2).Enabled = False
            Else
                ckcVehicleNonSports.Visible = True
            End If
            
            'No sports vehicles allowed with front load
            'If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) Then
            '    ckcVehicleSports.Visible = True
            'End If
            'ckcVehicleSports.Value = vbUnchecked
            rbcVehicle(1).Enabled = False
            rbcVehicle(2).Enabled = False
            grdAdvtCntr.Visible = True
            ckcAdvtCntr.Visible = True
            cmcAdvtCntr.Visible = True
            lacProductProtection.Visible = True
            lbcProductProtection.Visible = True
            'ckcProductProtection.Visible = True
            cmcAction.Caption = "Front Load"
            ckcRemoveSpot.Caption = "Move Spot"
            mClearGrid
            mSetGridTitles "F"
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) <> USINGSPORTS) Then
                mFillAdvtCntrGrid "F"
            End If
    End Select
    If cmcAction.Visible = False Then
        cmcAction.Left = SpotAction.Width / 2 - cmcAction.Width - cmcCancel.Width / 2
        cmcCancel.Left = SpotAction.Width / 2 + cmcCancel.Width / 2
        cmcAction.Visible = True
    End If
    pbcAction.Visible = True
    If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) Then
        If (imLastTabSelected = -1) Or (imLastTabSelected <> tbcSelection.SelectedItem.Index) And (imLastTabSelected = imFrontLoadIndex) Or (imLastTabSelected <> tbcSelection.SelectedItem.Index) And (tbcSelection.SelectedItem.Index = imFrontLoadIndex) Then
            imLastTabSelected = tbcSelection.SelectedItem.Index
            mVehPop
        End If
    End If
    imLastTabSelected = tbcSelection.SelectedItem.Index
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbDefault
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMissedPop                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Missed list box if    *
'*                      required                       *
'*                                                     *
'*******************************************************
Private Sub mMissedCancelPop()
'
'   mMissedPop
'   Where:
'
    ReDim ilfilter(0 To 1) As Integer
    ReDim slFilter(0 To 1) As String
    ReDim ilOffSet(0 To 1) As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDefaultName As String
    Dim slStr As String

    'Missed Reason
    slDefaultName = "Missed"
    slStr = ""
    'ReDim tgMRMnf(1 To 1) As MNF
    ReDim tgMRMnf(0 To 0) As MNF
    ilRet = gObtainMnfForType("M", slStr, tgMRMnf())
    For ilLoop = LBound(tgMRMnf) To UBound(tgMRMnf) - 1 Step 1
        If Trim$(tgMRMnf(ilLoop).sUnitType) = "Y" Then
            slDefaultName = Trim$(tgMRMnf(ilLoop).sName)
            Exit For
        End If
    Next ilLoop
    'Repopulate if required- if sales source changed by another user while in this screen
    ilfilter(0) = CHARFILTER
    slFilter(0) = "M"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilfilter(1) = CHARFILTERNOT
    slFilter(1) = "A"
    ilOffSet(1) = gFieldOffset("Mnf", "MnfCodeStn")
    'ilRet = gIMoveListBox(PostLog, lbcMissedCancelReason, lbcMissedCancelReasonCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(SpotAction, lbcMissedCancelReason, tmMissedCancelCode(), smMissedCancelCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mMissedPopErr
        gCPErrorMsg ilRet, "mMissedPop (gIMoveListBox)", SpotAction
        On Error GoTo 0
        imChgMode = True
        gFindMatch slDefaultName, 0, lbcMissedCancelReason
        If gLastFound(lbcMissedCancelReason) >= 0 Then
            lbcMissedCancelReason.ListIndex = gLastFound(lbcMissedCancelReason)
        Else
            lbcMissedCancelReason.ListIndex = -1
        End If
        imChgMode = False
    End If
    Exit Sub
mMissedPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mCompPop                        *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate competitive list      *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mProductProtectionPop()
'
'   mCompPop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slName As String      'Competitive name, saved to determine if changed
    Dim ilIndex As Integer      'Competitive name, saved to determine if changed
    ilfilter(0) = CHARFILTER
    slFilter(0) = "C"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilIndex = lbcProductProtection.ListIndex
    If ilIndex >= 0 Then
        slName = lbcProductProtection.List(ilIndex)
    End If
    ilRet = gIMoveListBox(SpotAction, lbcProductProtection, tmProductProtectionCode(), smProductProtectionCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mCompPopErr
        gCPErrorMsg ilRet, "mProductProtectionPop (gIMoveListBox)", SpotAction
        On Error GoTo 0
        imChgMode = True
        If ilIndex >= 0 Then
            gFindMatch slName, 0, lbcProductProtection
            If gLastFound(lbcProductProtection) >= 0 Then
                lbcProductProtection.ListIndex = gLastFound(lbcProductProtection)
            Else
                lbcProductProtection.ListIndex = -1
            End If
        Else
            lbcProductProtection.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mCompPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
    Dim ilRet As Integer
    Dim llFilter As Long
    Dim ilSetVehicle As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilLoop As Integer
    
    If lbcVehicles.ListCount <= 0 Then
        ilSetVehicle = True
    Else
        ilSetVehicle = False
    End If
    lbcVehicles.Clear
    smUserVehicleTag = ""
    ''ilRet = gPopUserVehicleBox(Spots, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, cbcVehicle, Traffic!lbcUserVehicle)
    'ilRet = gPopUserVehicleBox(Spots, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, cbcVehicle, tmUserVehicle(), smUserVehicleTag)
    If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) Then
        If imLastTabSelected = imFrontLoadIndex Then
            If tgSpf.sMktBase = "Y" Then
                llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHEXCLUDESPORT + VEHVIRTUAL + ACTIVEVEH + VEHBYMKT + VEHEXCLUDEIMPORTINVOICESPOTS + VEHEXCLUDEPODNOPRGM ' Airing and all conventional vehicles (except with Log) and Log, 1/25/21 exclude CPM vehicles, 2/8/21 - exclude podcast w/o program based on ad server.
                ilRet = gPopUserVehicleByMkt(SpotAction, llFilter, igSpotMktCode(), lbcVehicles, tmUserVehicle(), smUserVehicleTag)
            Else
                llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHEXCLUDESPORT + VEHVIRTUAL + ACTIVEVEH + VEHEXCLUDEIMPORTINVOICESPOTS + VEHEXCLUDEPODNOPRGM ' Airing and all conventional vehicles (except with Log) and Log, 1/25/21 exclude CPM vehicles, 2/8/21 - exclude podcast w/o program based on ad server.
                ilRet = gPopUserVehicleBox(SpotAction, llFilter, lbcVehicles, tmUserVehicle(), smUserVehicleTag)
            End If
        Else
            If tgSpf.sMktBase = "Y" Then
                llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHSPORT + VEHVIRTUAL + ACTIVEVEH + VEHBYMKT + VEHEXCLUDEIMPORTINVOICESPOTS + VEHEXCLUDEPODNOPRGM ' Airing and all conventional vehicles (except with Log) and Log, 1/25/21 exclude CPM vehicles, 2/8/21 - exclude podcast w/o program based on ad server.
                ilRet = gPopUserVehicleByMkt(SpotAction, llFilter, igSpotMktCode(), lbcVehicles, tmUserVehicle(), smUserVehicleTag)
            Else
                llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHSPORT + VEHVIRTUAL + ACTIVEVEH + VEHEXCLUDEIMPORTINVOICESPOTS + VEHEXCLUDEPODNOPRGM ' Airing and all conventional vehicles (except with Log) and Log, 1/25/21 exclude CPM vehicles, 2/8/21 - exclude podcast w/o program based on ad server.
                ilRet = gPopUserVehicleBox(SpotAction, llFilter, lbcVehicles, tmUserVehicle(), smUserVehicleTag)
            End If
        End If
    Else
        If tgSpf.sMktBase = "Y" Then
            llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + VEHBYMKT + VEHEXCLUDEIMPORTINVOICESPOTS + VEHEXCLUDEPODNOPRGM ' Airing and all conventional vehicles (except with Log) and Log, 1/25/21 exclude CPM vehicles, 2/8/21 - exclude podcast w/o program based on ad server.
            ilRet = gPopUserVehicleByMkt(SpotAction, llFilter, igSpotMktCode(), lbcVehicles, tmUserVehicle(), smUserVehicleTag)
        Else
            llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + VEHEXCLUDEIMPORTINVOICESPOTS + VEHEXCLUDEPODNOPRGM ' Airing and all conventional vehicles (except with Log) and Log, 1/25/21 exclude CPM vehicles, 2/8/21 - exclude podcast w/o program based on ad server.
            ilRet = gPopUserVehicleBox(SpotAction, llFilter, lbcVehicles, tmUserVehicle(), smUserVehicleTag)
        End If
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", SpotAction
        On Error GoTo 0
    End If
    If ilSetVehicle Then
        For ilLoop = 0 To UBound(tmUserVehicle) - 1 Step 1
            slNameCode = tmUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            If igSpotActionVefCode = ilVefCode Then
                lbcVehicles.Selected(ilLoop) = True
                cmcAdvtCntr_Click
                Exit For
            End If
        Next ilLoop
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    If imDateBox = 1 Then
        slStr = edcFDDropDown(0).Text
    ElseIf imDateBox = 2 Then
        slStr = edcFDDropDown(1).Text
    ElseIf imDateBox = 3 Then
        slStr = edcTDDropDown(0).Text
    ElseIf imDateBox = 4 Then
        slStr = edcTDDropDown(1).Text
    End If
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
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

Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                If imDateBox = 1 Then
                    edcFDDropDown(0).Text = Format$(llDate, "m/d/yy")
                    edcFDDropDown(0).SelStart = 0
                    edcFDDropDown(0).SelLength = Len(edcFDDropDown(0).Text)
                    imBypassFocus = True
                    edcFDDropDown(0).SetFocus
                    Exit Sub
                ElseIf imDateBox = 2 Then
                    edcFDDropDown(1).Text = Format$(llDate, "m/d/yy")
                    edcFDDropDown(1).SelStart = 0
                    edcFDDropDown(1).SelLength = Len(edcFDDropDown(1).Text)
                    imBypassFocus = True
                    edcFDDropDown(1).SetFocus
                    Exit Sub
                ElseIf imDateBox = 3 Then
                    edcTDDropDown(0).Text = Format$(llDate, "m/d/yy")
                    edcTDDropDown(0).SelStart = 0
                    edcTDDropDown(0).SelLength = Len(edcTDDropDown(0).Text)
                    imBypassFocus = True
                    edcTDDropDown(0).SetFocus
                    Exit Sub
                ElseIf imDateBox = 4 Then
                    edcTDDropDown(1).Text = Format$(llDate, "m/d/yy")
                    edcTDDropDown(1).SelStart = 0
                    edcTDDropDown(1).SelLength = Len(edcTDDropDown(1).Text)
                    imBypassFocus = True
                    edcTDDropDown(1).SetFocus
                    Exit Sub
                End If
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    If imDateBox = 1 Then
        edcFDDropDown(0).SetFocus
    ElseIf imDateBox = 2 Then
        edcFDDropDown(1).SetFocus
    ElseIf imDateBox = 3 Then
        edcTDDropDown(0).SetFocus
    ElseIf imDateBox = 4 Then
        edcTDDropDown(1).SetFocus
    End If
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub cmcCalDn_Click()
    If imDateBox = 1 Then
        imCalMonth = imCalMonth - 1
        If imCalMonth <= 0 Then
            imCalMonth = 12
            imCalYear = imCalYear - 1
        End If
        pbcCalendar_Paint
        edcFDDropDown(0).SelStart = 0
        edcFDDropDown(0).SelLength = Len(edcFDDropDown(0).Text)
        edcFDDropDown(0).SetFocus
    ElseIf imDateBox = 2 Then
        imCalMonth = imCalMonth - 1
        If imCalMonth <= 0 Then
            imCalMonth = 12
            imCalYear = imCalYear - 1
        End If
        pbcCalendar_Paint
        edcFDDropDown(1).SelStart = 0
        edcFDDropDown(1).SelLength = Len(edcFDDropDown(1).Text)
        edcFDDropDown(1).SetFocus
    ElseIf imDateBox = 3 Then
        imCalMonth = imCalMonth - 1
        If imCalMonth <= 0 Then
            imCalMonth = 12
            imCalYear = imCalYear - 1
        End If
        pbcCalendar_Paint
        edcTDDropDown(0).SelStart = 0
        edcTDDropDown(0).SelLength = Len(edcTDDropDown(0).Text)
        edcTDDropDown(0).SetFocus
    ElseIf imDateBox = 4 Then
        imCalMonth = imCalMonth - 1
        If imCalMonth <= 0 Then
            imCalMonth = 12
            imCalYear = imCalYear - 1
        End If
        pbcCalendar_Paint
        edcTDDropDown(1).SelStart = 0
        edcTDDropDown(1).SelLength = Len(edcTDDropDown(1).Text)
        edcTDDropDown(1).SetFocus
    End If
End Sub
Private Sub cmcCalUp_Click()
    If imDateBox = 1 Then
        imCalMonth = imCalMonth + 1
        If imCalMonth > 12 Then
            imCalMonth = 1
            imCalYear = imCalYear + 1
        End If
        pbcCalendar_Paint
        edcFDDropDown(0).SelStart = 0
        edcFDDropDown(0).SelLength = Len(edcFDDropDown(0).Text)
        edcFDDropDown(0).SetFocus
    ElseIf imDateBox = 2 Then
        imCalMonth = imCalMonth + 1
        If imCalMonth > 12 Then
            imCalMonth = 1
            imCalYear = imCalYear + 1
        End If
        pbcCalendar_Paint
        edcFDDropDown(1).SelStart = 0
        edcFDDropDown(1).SelLength = Len(edcFDDropDown(1).Text)
        edcFDDropDown(1).SetFocus
    ElseIf imDateBox = 3 Then
        imCalMonth = imCalMonth + 1
        If imCalMonth > 12 Then
            imCalMonth = 1
            imCalYear = imCalYear + 1
        End If
        pbcCalendar_Paint
        edcTDDropDown(0).SelStart = 0
        edcTDDropDown(0).SelLength = Len(edcTDDropDown(0).Text)
        edcTDDropDown(0).SetFocus
    ElseIf imDateBox = 4 Then
        imCalMonth = imCalMonth + 1
        If imCalMonth > 12 Then
            imCalMonth = 1
            imCalYear = imCalYear + 1
        End If
        pbcCalendar_Paint
        edcTDDropDown(1).SelStart = 0
        edcTDDropDown(1).SelLength = Len(edcTDDropDown(1).Text)
        edcTDDropDown(1).SetFocus
    End If
End Sub

Private Sub tbcSelection_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
End Sub

Private Function mCheckFromDateTime(Optional slAction As String = "")
    Dim slSDate As String
    Dim slEDate As String
    Dim slSTime As String
    Dim slETime As String
    Dim llSTime As Long
    Dim llETime As Long
    
    If imLastTabSelected = -1 Then
        mCheckFromDateTime = True
        Exit Function
    End If
    mCheckFromDateTime = False
    slSDate = Trim$(edcFDDropDown(0).Text)
    If slSDate = "" Then
        MsgBox "From Start Date is Not Defined", vbOKOnly + vbApplicationModal, "Start Date"
        edcFDDropDown(0).SetFocus
        Exit Function
    End If
    If Not gValidDate(slSDate) Then
        MsgBox "From Start Date is Not Valid", vbOKOnly + vbApplicationModal, "Start Date"
        edcFDDropDown(0).SetFocus
        Exit Function
    End If
    slEDate = Trim$(edcFDDropDown(1).Text)
    If slEDate = "" Then
        MsgBox "From End Date is Not Defined", vbOKOnly + vbApplicationModal, "End Date"
        edcFDDropDown(1).SetFocus
        Exit Function
    End If
    If Not gValidDate(slEDate) Then
        MsgBox "From End Date is Not Valid", vbOKOnly + vbApplicationModal, "End Date"
        edcFDDropDown(1).SetFocus
        Exit Function
    End If
    slSTime = edcFTDropDown(0).Text
    If slSTime = "" Then
        MsgBox "From Start Time is Not Defined", vbOKOnly + vbApplicationModal, "Start Time"
        edcFTDropDown(0).SetFocus
        Exit Function
    End If
    If Not gValidTime(slSTime) Then
        MsgBox "From Start Time is Not Valid", vbOKOnly + vbApplicationModal, "Start Time"
        edcFDDropDown(0).SetFocus
        Exit Function
    End If
    slETime = edcFTDropDown(1).Text
    If slETime = "" Then
        MsgBox "From End Time is Not Defined", vbOKOnly + vbApplicationModal, "End Time"
        edcFTDropDown(1).SetFocus
        Exit Function
    End If
    If Not gValidTime(slETime) Then
        MsgBox "From End Time is Not Valid", vbOKOnly + vbApplicationModal, "End Time"
        edcFDDropDown(1).SetFocus
        Exit Function
    End If
    If gDateValue(slEDate) < gDateValue(slSDate) Then
        MsgBox "From End Date must be after or equal to Start Date", vbOKOnly + vbApplicationModal, "Date"
        edcFDDropDown(0).SetFocus
        Exit Function
    End If
    If (slAction <> "H") Or (Not rbcRemnant(0).Value) Then
        If ((Asc(tgSpf.sUsingFeatures4) And ALLOWMOVEONTODAY) = ALLOWMOVEONTODAY) Then
            If gDateValue(slSDate) < gDateValue(Format$(gNow(), "m/d/yy")) Then
                MsgBox "From Start Date must be after or equal to " & Format$(gNow(), "m/d/yy"), vbOKOnly + vbApplicationModal, "Date"
                edcFDDropDown(0).SetFocus
                Exit Function
            End If
        Else
            If gDateValue(slSDate) <= gDateValue(Format$(gNow(), "m/d/yy")) Then
                MsgBox "From Start Date must be after " & Format$(gNow(), "m/d/yy"), vbOKOnly + vbApplicationModal, "Date"
                edcFDDropDown(0).SetFocus
                Exit Function
            End If
        End If
    End If
    llSTime = gTimeToLong(slSTime, False)
    llETime = gTimeToLong(slETime, True)
    If llETime < llSTime Then
        MsgBox "From End Time must be after or equal to Start Time", vbOKOnly + vbApplicationModal, "Time"
        edcFTDropDown(0).SetFocus
        Exit Function
    End If
    mCheckFromDateTime = True
End Function

Private Function mCheckToDateTime()
    Dim slSDate As String
    Dim slEDate As String
    Dim slSTime As String
    Dim slETime As String
    Dim llSTime As Long
    Dim llETime As Long

    mCheckToDateTime = False
    slSDate = Trim$(edcTDDropDown(0).Text)
    If slSDate = "" Then
        MsgBox "To Start Date is Not Defined", vbOKOnly + vbApplicationModal, "Start Date"
        edcFDDropDown(0).SetFocus
        Exit Function
    End If
    If Not gValidDate(slSDate) Then
        MsgBox "To Start Date is Not Valid", vbOKOnly + vbApplicationModal, "Start Date"
        edcFDDropDown(0).SetFocus
        Exit Function
    End If
    slEDate = Trim$(edcTDDropDown(1).Text)
    If slEDate = "" Then
        MsgBox "To End Date is Not Defined", vbOKOnly + vbApplicationModal, "End Date"
        edcFDDropDown(1).SetFocus
        Exit Function
    End If
    If Not gValidDate(slEDate) Then
        MsgBox "To End Date is Not Valid", vbOKOnly + vbApplicationModal, "End Date"
        edcFDDropDown(1).SetFocus
        Exit Function
    End If
    slSTime = edcTTDropDown(0).Text
    If slSTime = "" Then
        MsgBox "To Start Time is Not Defined", vbOKOnly + vbApplicationModal, "Start Time"
        edcFTDropDown(0).SetFocus
        Exit Function
    End If
    If Not gValidTime(slSTime) Then
        MsgBox "To Start Time is Not Valid", vbOKOnly + vbApplicationModal, "Start Time"
        edcFDDropDown(0).SetFocus
        Exit Function
    End If
    slETime = edcTTDropDown(1).Text
    If slETime = "" Then
        MsgBox "To End Time is Not Defined", vbOKOnly + vbApplicationModal, "End Time"
        edcFTDropDown(1).SetFocus
        Exit Function
    End If
    If Not gValidTime(slETime) Then
        MsgBox "To End Time is Not Valid", vbOKOnly + vbApplicationModal, "End Time"
        edcFDDropDown(1).SetFocus
        Exit Function
    End If
    If gDateValue(slEDate) < gDateValue(slSDate) Then
        MsgBox "To End Date must be after or equal to Start Date", vbOKOnly + vbApplicationModal, "Date"
        edcFDDropDown(0).SetFocus
        Exit Function
    End If
    llSTime = gTimeToLong(slSTime, False)
    llETime = gTimeToLong(slETime, True)
    If llETime < llSTime Then
        MsgBox "To End Time must be after or equal to Start Time", vbOKOnly + vbApplicationModal, "Time"
        edcFTDropDown(0).SetFocus
        Exit Function
    End If
    mCheckToDateTime = True
End Function

Private Sub grdAdvtCntr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdAdvtCntr.rowHeight(0) Then
        grdAdvtCntr.Col = grdAdvtCntr.MouseCol
        mAdvtCntrSortCol grdAdvtCntr.Col
        grdAdvtCntr.Row = 0
        grdAdvtCntr.Col = CHFCODEINDEX
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdAdvtCntr, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdAdvtCntr.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdAdvtCntr.FixedRows Then
        If grdAdvtCntr.TextMatrix(llCurrentRow, ADVTINDEX) <> "" Then
            grdAdvtCntr.TopRow = lmScrollTop
            llTopRow = grdAdvtCntr.TopRow
            If (Shift And CTRLMASK) > 0 Then
                If grdAdvtCntr.TextMatrix(grdAdvtCntr.Row, CHFCODEINDEX) <> "" Then
                    If grdAdvtCntr.TextMatrix(grdAdvtCntr.Row, SELECTEDINDEX) <> "1" Then
                        grdAdvtCntr.TextMatrix(grdAdvtCntr.Row, SELECTEDINDEX) = "1"
                    Else
                        grdAdvtCntr.TextMatrix(grdAdvtCntr.Row, SELECTEDINDEX) = "0"
                    End If
                    mPaintRowColor grdAdvtCntr.Row
                End If
            Else
                For llRow = grdAdvtCntr.FixedRows To grdAdvtCntr.Rows - 1 Step 1
                    If grdAdvtCntr.TextMatrix(llRow, ADVTINDEX) <> "" Then
                        grdAdvtCntr.TextMatrix(llRow, SELECTEDINDEX) = "0"
                        If grdAdvtCntr.TextMatrix(llRow, CHFCODEINDEX) <> "" Then
                            If (lmLastClickedRow = -1) Or ((Shift And SHIFTMASK) <= 0) Then
                                If llRow = llCurrentRow Then
                                    grdAdvtCntr.TextMatrix(llRow, SELECTEDINDEX) = "1"
                                Else
                                    grdAdvtCntr.TextMatrix(llRow, SELECTEDINDEX) = "0"
                                End If
                            ElseIf lmLastClickedRow < llCurrentRow Then
                                If (llRow >= lmLastClickedRow) And (llRow <= llCurrentRow) Then
                                    grdAdvtCntr.TextMatrix(llRow, SELECTEDINDEX) = "1"
                                End If
                            Else
                                If (llRow >= llCurrentRow) And (llRow <= lmLastClickedRow) Then
                                    grdAdvtCntr.TextMatrix(llRow, SELECTEDINDEX) = "1"
                                End If
                            End If
                            mPaintRowColor llRow
                        End If
                    End If
                Next llRow
                grdAdvtCntr.TopRow = llTopRow
                grdAdvtCntr.Row = llCurrentRow
            End If
            lmLastClickedRow = llCurrentRow
        End If
    End If
    If ckcProductProtection.Value = vbChecked Then
        ckcProductProtection.Value = vbUnchecked
    Else
        ckcProductProtection_Click
    End If
    imBypassAll = True
    ckcAdvtCntr.Value = vbUnchecked
    imBypassAll = False
End Sub

Private Sub grdAdvtCntr_Scroll()
    cmcCancel.Caption = "&Cancel"
    plcTme.Visible = False
    plcCalendar.Visible = False
    lmScrollTop = grdAdvtCntr.TopRow
End Sub

Private Sub grdAdvtCntr_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If
    If (Shift And SHIFTMASK) > 0 Then
        imShiftKey = True
    Else
        imShiftKey = False
    End If
End Sub

Private Sub grdAdvtCntr_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
    imShiftKey = False
End Sub
Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    grdAdvtCntr.ColWidth(SORTINDEX) = 0
    grdAdvtCntr.ColWidth(SELECTEDINDEX) = 0
    grdAdvtCntr.ColWidth(CHFCODEINDEX) = 0
    'grdAdvtCntr.ColWidth(ADVTINDEX) = grdAdvtCntr.Width * 0.07
    grdAdvtCntr.ColWidth(CNTRINDEX) = grdAdvtCntr.Width * 0.18
    grdAdvtCntr.ColWidth(SCHDINDEX) = grdAdvtCntr.Width * 0.13
    grdAdvtCntr.ColWidth(MISSINDEX) = grdAdvtCntr.Width * 0.13
    
    grdAdvtCntr.ColWidth(ADVTINDEX) = grdAdvtCntr.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To MISSINDEX Step 1
        If ilCol <> ADVTINDEX Then
            grdAdvtCntr.ColWidth(ADVTINDEX) = grdAdvtCntr.ColWidth(ADVTINDEX) - grdAdvtCntr.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdAdvtCntr
End Sub

Private Sub mSetGridTitles(slAction As String)
    'Set column titles
    grdAdvtCntr.TextMatrix(0, ADVTINDEX) = "Advertiser/Product"
    grdAdvtCntr.TextMatrix(0, CNTRINDEX) = "Contract"
    If slAction = "P" Then
        grdAdvtCntr.TextMatrix(0, SCHDINDEX) = "Schd"
        grdAdvtCntr.TextMatrix(0, MISSINDEX) = "Missed"
    ElseIf slAction = "R" Then
        grdAdvtCntr.TextMatrix(0, SCHDINDEX) = "Missed"
        grdAdvtCntr.TextMatrix(0, MISSINDEX) = "Schd"
    ElseIf slAction = "C" Then
        If rbcCCancel(0).Value Then
            grdAdvtCntr.TextMatrix(0, SCHDINDEX) = "Schd"
            grdAdvtCntr.TextMatrix(0, MISSINDEX) = "Cancel"
        ElseIf rbcCCancel(1).Value Then
            grdAdvtCntr.TextMatrix(0, SCHDINDEX) = "Missed"
            grdAdvtCntr.TextMatrix(0, MISSINDEX) = "Cancel"
        ElseIf rbcCCancel(2).Value Then
            grdAdvtCntr.TextMatrix(0, SCHDINDEX) = "Cancel"
            grdAdvtCntr.TextMatrix(0, MISSINDEX) = "Missed"
        Else
            grdAdvtCntr.TextMatrix(0, SCHDINDEX) = ""
            grdAdvtCntr.TextMatrix(0, MISSINDEX) = ""
        End If
    ElseIf slAction = "E" Then
        grdAdvtCntr.TextMatrix(0, SCHDINDEX) = "Fills"
    ElseIf slAction = "F" Then
        grdAdvtCntr.TextMatrix(0, SCHDINDEX) = "Schd"
        grdAdvtCntr.TextMatrix(0, MISSINDEX) = "Missed"
    ElseIf slAction = "H" Then
        If rbcRemnant(0).Value Then
            grdAdvtCntr.TextMatrix(0, SCHDINDEX) = "Missed"
            grdAdvtCntr.TextMatrix(0, MISSINDEX) = "Hidden"
        ElseIf rbcRemnant(1).Value Then
            grdAdvtCntr.TextMatrix(0, SCHDINDEX) = "Hidden"
            grdAdvtCntr.TextMatrix(0, MISSINDEX) = "Missed"
        Else
            grdAdvtCntr.TextMatrix(0, SCHDINDEX) = ""
            grdAdvtCntr.TextMatrix(0, MISSINDEX) = ""
        End If
    Else
        grdAdvtCntr.TextMatrix(0, SCHDINDEX) = ""
        grdAdvtCntr.TextMatrix(0, MISSINDEX) = ""
    End If
    'Set height of grid

End Sub

Private Sub mAdvtCntrSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slDate As String
    Dim slTime As String
    Dim slDays As String
    Dim slHours As String
    Dim slMinutes As String
    Dim ilChar As Integer
    
    
    For llRow = grdAdvtCntr.FixedRows To grdAdvtCntr.Rows - 1 Step 1
        slStr = Trim$(grdAdvtCntr.TextMatrix(llRow, ADVTINDEX))
        If slStr <> "" Then
            If ilCol = CNTRINDEX Then
                slSort = grdAdvtCntr.TextMatrix(llRow, CNTRINDEX)
                Do While Len(slSort) < 8
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = SCHDINDEX Then
                slSort = grdAdvtCntr.TextMatrix(llRow, SCHDINDEX)
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = MISSINDEX) Then
                slSort = grdAdvtCntr.TextMatrix(llRow, MISSINDEX)
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdAdvtCntr.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdAdvtCntr.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastAdvtCntrColSorted) Or ((ilCol = imLastAdvtCntrColSorted) And (imLastAdvtCntrSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdAdvtCntr.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdAdvtCntr.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastAdvtCntrColSorted Then
        imLastAdvtCntrColSorted = SORTINDEX
    Else
        imLastAdvtCntrColSorted = -1
        imLastAdvtCntrSort = -1
    End If
    gGrid_SortByCol grdAdvtCntr, ADVTINDEX, SORTINDEX, imLastAdvtCntrColSorted, imLastAdvtCntrSort
    imLastAdvtCntrColSorted = ilCol
End Sub
Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    imBypassAll = True
    ckcAdvtCntr.Value = vbUnchecked
    imBypassAll = False
    grdAdvtCntr.rowHeight(0) = fgBoxGridH + 15
    For llRow = grdAdvtCntr.FixedRows To grdAdvtCntr.Rows - 1 Step 1
        grdAdvtCntr.Row = llRow
        For llCol = 0 To CHFCODEINDEX Step 1
            grdAdvtCntr.Col = llCol
            grdAdvtCntr.CellBackColor = vbWhite
            grdAdvtCntr.TextMatrix(llRow, llCol) = ""
        Next llCol
        grdAdvtCntr.rowHeight(llRow) = fgBoxGridH + 15
    Next llRow
    lmLastClickedRow = -1
    imLastAdvtCntrColSorted = -1
    imLastAdvtCntrSort = -1
    lmScrollTop = grdAdvtCntr.FixedRows
End Sub
Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long
    
    grdAdvtCntr.Row = llRow
    For llCol = ADVTINDEX To MISSINDEX Step 1
        grdAdvtCntr.Col = llCol
        If grdAdvtCntr.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
            grdAdvtCntr.CellBackColor = vbWhite
        Else
            grdAdvtCntr.CellBackColor = GRAY    'vbBlue
        End If
    Next llCol

End Sub


Private Function mLockOperation() As Integer
    Dim ilRet As Integer
    Dim slSDate As String
    Dim llSDate As Long
    Dim slEDate As String
    Dim llEDate As Long
    Dim slSTime As String
    Dim slETime As String
    Dim ilSpotLock As Integer
    Dim ilAvailLock As Integer
    Dim ilVef As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim ilIndex As Integer
    
    If (tbcSelection.SelectedItem.Index = imLockIndex) Or (tbcSelection.SelectedItem.Index = imUnlockIndex) Then
        If Not mCheckFromDateTime() Then
            mLockOperation = False
            Exit Function
        End If
        If (rbcLInfo(0).Value = False) And (rbcLInfo(1).Value = False) Then
            MsgBox "Lock Avail or Spot is Not Defined", vbOKOnly + vbApplicationModal, "Lock"
            mLockOperation = False
            Exit Function
        End If
        If lbcVehicles.SelCount <= 0 Then
            MsgBox "No Vehicle(s) has been selected", vbOKOnly + vbApplicationModal, "Lock"
            mLockOperation = False
            Exit Function
        End If
        Screen.MousePointer = vbHourglass
        gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbHourglass
    End If
    slSDate = Trim$(edcFDDropDown(0).Text)
    llSDate = gDateValue(slSDate)
    slEDate = Trim$(edcFDDropDown(1).Text)
    llEDate = gDateValue(slEDate)
    slSTime = edcFTDropDown(0).Text
    slETime = edcFTDropDown(1).Text
    
    For ilVef = 0 To lbcVehicles.ListCount - 1 Step 1
        If lbcVehicles.Selected(ilVef) Then
            slNameCode = tmUserVehicle(ilVef).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            If (tbcSelection.SelectedItem.Index = imLockIndex) Or (tbcSelection.SelectedItem.Index = imUnlockIndex) Then
                If rbcLInfo(0).Value = True Then   'Lock Avail
                    ilSpotLock = -1
                    ilAvailLock = 1 'Lock
                Else    'Lock spots
                    ilAvailLock = -1
                    ilSpotLock = 1  'Lock
                End If
            Else
                ilSpotLock = -1
                ilAvailLock = 1 'Lock
            End If
            ilIndex = gBinarySearchVef(ilVefCode)
            If ilIndex <> -1 Then
                If tgMVef(ilIndex).sType = "G" Then
                    tmGsfSrchKey3.iVefCode = ilVefCode
                    tmGsfSrchKey3.iGameNo = 0
                    ilRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)   'Get last current record to obtain date
                    Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = ilVefCode)
                        gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llDate
                        If (llDate >= llSDate) And (llDate <= llEDate) Then
                            gUnpackDate tmGsf.iAirDate(0), tmGsf.iAirDate(1), slDate
                            gSetLockStatus ilVefCode, ilAvailLock, ilSpotLock, slDate, slDate, tmGsf.iGameNo, slSTime, slETime
                        End If
                        ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                Else
                    gSetLockStatus ilVefCode, ilAvailLock, ilSpotLock, slSDate, slEDate, 0, slSTime, slETime
                End If
            End If
        End If
    Next ilVef
    If (tbcSelection.SelectedItem.Index = imLockIndex) Or (tbcSelection.SelectedItem.Index = imUnlockIndex) Then
        Screen.MousePointer = vbDefault
        gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbDefault
    End If
    mLockOperation = True
End Function

Private Function mUnlockOperation(ilSource As Integer) As Integer
    Dim ilRet As Integer
    Dim slSDate As String
    Dim llSDate As Long
    Dim slEDate As String
    Dim llEDate As Long
    Dim slSTime As String
    Dim slETime As String
    Dim ilSpotLock As Integer
    Dim ilAvailLock As Integer
    Dim ilVef As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim ilIndex As Integer
    
    If ilSource = 0 Then
        If Not mCheckFromDateTime() Then
            mUnlockOperation = False
            Exit Function
        End If
        If (rbcLInfo(0).Value = False) And (rbcLInfo(1).Value = False) Then
            MsgBox "Unlock Avail or Spot is Not Defined", vbOKOnly + vbApplicationModal, "Unlock"
            mUnlockOperation = False
            Exit Function
        End If
        If lbcVehicles.SelCount <= 0 Then
            MsgBox "No Vehicle(s) has been selected", vbOKOnly + vbApplicationModal, "Unlock"
            mUnlockOperation = False
            Exit Function
        End If
    
        Screen.MousePointer = vbHourglass
        gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbHourglass
    End If
    slSDate = Trim$(edcFDDropDown(0).Text)
    llSDate = gDateValue(slSDate)
    slEDate = Trim$(edcFDDropDown(1).Text)
    llEDate = gDateValue(slEDate)
    slSTime = edcFTDropDown(0).Text
    slETime = edcFTDropDown(1).Text
    
    For ilVef = 0 To lbcVehicles.ListCount - 1 Step 1
        If lbcVehicles.Selected(ilVef) Then
            slNameCode = tmUserVehicle(ilVef).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            If ilSource = 0 Then
                If rbcLInfo(0).Value = True Then   'Unlock Avail
                    ilSpotLock = -1
                    ilAvailLock = 0 'Unlock
                Else    'Unlock spots
                    ilAvailLock = -1
                    ilSpotLock = 0  'Unlock
                End If
            Else
                ilAvailLock = 0
                ilSpotLock = 0
            End If
            ilIndex = gBinarySearchVef(ilVefCode)
            If ilIndex <> -1 Then
                If tgMVef(ilIndex).sType = "G" Then
                    tmGsfSrchKey3.iVefCode = ilVefCode
                    tmGsfSrchKey3.iGameNo = 0
                    ilRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)   'Get last current record to obtain date
                    Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = ilVefCode)
                        gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llDate
                        If (llDate >= llSDate) And (llDate <= llEDate) Then
                            gUnpackDate tmGsf.iAirDate(0), tmGsf.iAirDate(1), slDate
                            gSetLockStatus ilVefCode, ilAvailLock, ilSpotLock, slDate, slDate, tmGsf.iGameNo, slSTime, slETime
                        End If
                        ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                Else
                    gSetLockStatus ilVefCode, ilAvailLock, ilSpotLock, slSDate, slEDate, 0, slSTime, slETime
                End If
            End If
        End If
    Next ilVef
    If ilSource = 0 Then
        Screen.MousePointer = vbDefault
        gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbDefault
    End If
    mUnlockOperation = True
End Function

Private Function mRescheduleOperation() As Integer
    Dim ilRet As Integer
    
    If Not mCheckFromDateTime() Then
        mRescheduleOperation = False
        Exit Function
    End If
    
    ReDim lmChfCode(0 To 0) As Long
    ReDim imPPCode(0 To 0) As Integer
    imAdvtCntrAll = False
    imPPAll = False
    If lbcProductProtection.SelCount <= 0 Then
        mBuildChfCode
    Else
        mBuildPPCode
    End If
    If (UBound(lmChfCode) = 0) And (UBound(imPPCode) = 0) And (Not imAdvtCntrAll) And (Not imPPAll) Then
        MsgBox "Either Advertiser/Contract or Product Protection must been selected", vbOKOnly + vbApplicationModal, "Reschedule"
        mRescheduleOperation = False
        Exit Function
    End If
    If lbcVehicles.SelCount <= 0 Then
        MsgBox "No Vehicle(s) has been selected", vbOKOnly + vbApplicationModal, "Reschedule"
        mRescheduleOperation = False
        Exit Function
    End If
    If (imPPAll) Or (UBound(imPPCode) > 0) Then
        If (rbcPPInclExcl(0).Value = False) And (rbcPPInclExcl(1).Value = False) Then
            MsgBox "Product Protection Include/Exclude is Not Defined", vbOKOnly + vbApplicationModal, "Reschedule"
            mRescheduleOperation = False
            Exit Function
        End If
    End If
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbHourglass
    ilRet = mBuildSpots("R")
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbDefault
    mRescheduleOperation = True
End Function

Private Function mCheckCntr(slAction As String) As Integer
    Dim ilChf As Integer
    Dim ilPP As Integer
    Dim ilRet As Integer
    Dim llChfCode As Long
    
    mCheckCntr = False
    llChfCode = tmSdf.lChfCode
    If (imAdvtCntrAll) Or (imPPAll) Then
        mCheckCntr = True
    Else
        If UBound(imPPCode) = 0 Then
            'Check Contract
            For ilChf = 0 To UBound(lmChfCode) - 1 Step 1
                If llChfCode = lmChfCode(ilChf) Then
                    mCheckCntr = True
                    Exit For
                End If
            Next ilChf
        Else
            'Check Product protection
            tmChfSrchKey0.lCode = llChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                For ilPP = 0 To UBound(imPPCode) - 1 Step 1
                    If ((slAction = "R") And (rbcPPInclExcl(0).Value)) Or (slAction <> "R") Then
                        'Test for Include
                        If (tmChf.iMnfComp(0) = imPPCode(ilPP)) Or (tmChf.iMnfComp(1) = imPPCode(ilPP)) Then
                            mCheckCntr = True
                            Exit Function
                        End If
                    Else    'Test for exclude
                        If (tmChf.iMnfComp(0) = imPPCode(ilPP)) Or (tmChf.iMnfComp(1) = imPPCode(ilPP)) Then
                            Exit Function
                        End If
                    End If
                Next ilPP
                If (slAction = "R") And (rbcPPInclExcl(1).Value) Then
                    mCheckCntr = True
                End If
            End If
        End If
    End If
End Function

Private Sub mBuildPPCode()
    Dim ilPP As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    
    imPPAll = False
    If ckcProductProtection.Value = vbUnchecked Then
        For ilPP = 0 To lbcProductProtection.ListCount - 1 Step 1
            If lbcProductProtection.Selected(ilPP) Then
                slNameCode = tmProductProtectionCode(ilPP).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                imPPCode(UBound(imPPCode)) = Val(slCode)
                ReDim Preserve imPPCode(0 To UBound(imPPCode) + 1) As Integer
            End If
        Next ilPP
    Else
        imPPAll = True
    End If

End Sub
Private Sub mBuildChfCode()
    Dim llRow As Long
    
    imAdvtCntrAll = False
    If grdAdvtCntr.TextMatrix(grdAdvtCntr.FixedRows, ADVTINDEX) <> "" Then
        If ckcAdvtCntr.Value = vbUnchecked Then
            For llRow = grdAdvtCntr.FixedRows To grdAdvtCntr.Rows - 1 Step 1
                If grdAdvtCntr.TextMatrix(llRow, ADVTINDEX) <> "" Then
                    If grdAdvtCntr.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                        lmChfCode(UBound(lmChfCode)) = Val(grdAdvtCntr.TextMatrix(llRow, CHFCODEINDEX))
                        ReDim Preserve lmChfCode(0 To UBound(lmChfCode) + 1) As Long
                    End If
                End If
            Next llRow
        Else
            imAdvtCntrAll = True
        End If
    End If
End Sub

Private Function mBuildSpots(slAction As String) As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slSDate As String
    Dim slEDate As String
    Dim slSTime As String
    Dim slETime As String
    Dim llSTime As Long
    Dim llETime As Long
    Dim llTime As Long
    Dim ilVef As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilChf As Integer
    Dim ilVefCode As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slDays As String
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim llTSDate As Long
    Dim llTEDate As Long
    Dim llDate As Long
    Dim slTSTime As String
    Dim slTETime As String
    Dim llTSTime As Long
    Dim llTETime As Long
    Dim ilMaxFrontLoad As Integer
    Dim ilFLMax As Integer
    Dim llLockRecCode As Long
    Dim llLastLockValue As Long
    Dim ilLockFailed As Integer
    Dim slMoDate As String
    Dim llSDate As Long
    Dim llMoDate As Long
    Dim ilDateOk As Integer
    Dim ilDay As Integer
    
    slSDate = Trim$(edcFDDropDown(0).Text)
    llSDate = gDateValue(slSDate)
    slMoDate = gObtainPrevMonday(slSDate)
    llMoDate = gDateValue(slMoDate)
    slEDate = Trim$(edcFDDropDown(1).Text)
    slSTime = edcFTDropDown(0).Text
    slETime = edcFTDropDown(1).Text
    llSTime = gTimeToLong(slSTime, False)
    llETime = gTimeToLong(slETime, True)
    slDays = ""
    llLastLockValue = 0
    llLockRecCode = 0
    ilLockFailed = False
    'ReDim lgReschSdfCode(1 To 1) As Long
    ReDim lgReschSdfCode(0 To 0) As Long
    ReDim tmSdfAction(0 To 0) As SDF
    If slAction = "A" Then   'Advertiser/Contract
        ReDim tmSpotActionInfo(0 To 0) As SPOTACTIONINFO
    End If
    If (slAction = "R") Or ((slAction = "P") And (ckcPIReSchd.Value = vbChecked)) Or (slAction = "F") Then   'Reschedule
        'ReDim lgReschSdfCode(1 To 1) As Long
        ReDim lgReschSdfCode(0 To 0) As Long
        gGetSchParameters
        sgMovePass = "N"
        sgCompPass = "N"
        '1/11/11: Jim- Allow preempt mode for reschedule
        '6/22/11: Jim-Allow on all modes
        'If (slAction <> "R") Then
        '    sgPreemptPass = "N"
        'End If
        igUsePreferred = False
        'gObtainMissedReasonCode
        Randomize   'Remove this if same results are to be obtained
        If Not gOpenSchFiles() Then
            mBuildSpots = False
            Exit Function
        End If
        If (slAction = "P") Then
            slDays = "NNNNNNN"
            For llDate = gDateValue(slSDate) To gDateValue(slEDate) Step 1
                Mid$(slDays, gWeekDayLong(llDate) + 1, 1) = "Y"
            Next llDate
        End If
    End If
    If slAction = "F" Then
        llTSDate = gDateValue(Trim$(edcTDDropDown(0).Text))
        llTEDate = gDateValue(Trim$(edcTDDropDown(1).Text))
        slDays = "NNNNNNN"
        For llDate = llTSDate To llTEDate Step 1
            Mid$(slDays, gWeekDayLong(llDate) + 1, 1) = "Y"
        Next llDate
        slTSTime = edcTTDropDown(0).Text
        slTETime = edcTTDropDown(1).Text
        llTSTime = gTimeToLong(slTSTime, False)
        llTETime = gTimeToLong(slTETime, True)
        ilMaxFrontLoad = Val(edcMaxFrontLoad.Text)
    End If
    For ilVef = 0 To lbcVehicles.ListCount - 1 Step 1
        If lbcVehicles.Selected(ilVef) Then
            If (slAction = "R") Or ((slAction = "P") And (ckcPIReSchd.Value = vbChecked)) Then   'Reschedule
                'ReDim lgReschSdfCode(1 To 1) As Long
                ReDim lgReschSdfCode(0 To 0) As Long
            End If
            If (slAction = "P") Then
                ReDim tmSdfAction(0 To 0) As SDF
            End If
            If slAction = "E" Then
                ReDim tmSdfAction(0 To 0) As SDF
            End If
            If slAction = "C" Then
                ReDim tmSdfAction(0 To 0) As SDF
            End If
            If slAction = "H" Then
                ReDim tmSdfAction(0 To 0) As SDF
            End If
            If slAction = "F" Then
                ReDim tmSdfAction(0 To 0) As SDF
            End If
            slNameCode = tmUserVehicle(ilVef).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            If bmFirstCallToVpfFind Then
                imVpfIndex = gVpfFind(SpotAction, ilVefCode)
                bmFirstCallToVpfFind = False
            Else
                imVpfIndex = gVpfFindIndex(ilVefCode)
            End If
            If imVpfIndex <> -1 Then
                gUnpackDateLong tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), lmLastLogDate
                btrExtClear hmSdf   'Clear any previous extend operation
                ilExtLen = Len(tmSdf)
                llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
                btrExtClear hmSdf   'Clear any previous extend operation
                tmSdfSrchKey1.iVefCode = ilVefCode
                gPackDate slSDate, tmSdfSrchKey1.iDate(0), tmSdfSrchKey1.iDate(1)
                tmSdfSrchKey1.iTime(0) = 0
                tmSdfSrchKey1.iTime(1) = 0
                tmSdfSrchKey1.sSchStatus = ""
                ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_END_OF_FILE Then
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Function
                    End If
                    Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)
                    tlIntTypeBuff.iType = ilVefCode
                    ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
                    ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                    'If slSDate <> "" Then
                    '    gPackDate slSDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    If slMoDate <> "" Then
                        gPackDate slMoDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                        ilOffSet = gFieldOffset("Sdf", "SdfDate")
                        If slEDate <> "" Then
                            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
                        Else
                            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                        End If
                    End If
                    If slEDate <> "" Then
                        gPackDate slEDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                        ilOffSet = gFieldOffset("Sdf", "SdfDate")
                        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                    End If
                    ilRet = btrExtAddField(hmSdf, 0, ilExtLen)  'Extract Name
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Function
                    End If
                    ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                            Exit Function
                        End If
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                        Loop
                        Do While ilRet = BTRV_ERR_NONE
                            If (slAction = "A") Or ((slAction <> "A") And (tmSdf.sBill <> "Y")) Then
                                ilDateOk = True
                                If tmSdf.sSchStatus = "M" Then
                                    If (slAction = "A") Or (slAction = "R") Or ((slAction = "C") And (rbcCCancel(1).Value)) Or ((slAction = "H") And (rbcRemnant(0).Value)) Then
                                        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                                        If (llMoDate < llSDate) And (llDate < llSDate) Then
                                            ilDateOk = False
                                            tmClfSrchKey.lChfCode = tmSdf.lChfCode
                                            tmClfSrchKey.iLine = tmSdf.iLineNo
                                            tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                                            tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                                            ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                                            Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                                                ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                            Loop
                                            If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
                                                ilRet = gGetSpotFlight(tmSdf, tmClf, hmCff, hmSmf, tmCff)
                                                If ilRet Then
                                                    If tmCff.sDyWk = "W" Then            'weekly
                                                        For ilDay = 0 To 6 Step 1
                                                            If llMoDate + ilDay >= llSDate Then
                                                                If tmCff.iDay(ilDay) > 0 Or tmCff.sXDay(ilDay) = "1" Then
                                                                    ilDateOk = True
                                                                    Exit For
                                                                End If
                                                            End If
                                                        Next ilDay
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If llMoDate < llSDate Then
                                            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                                            If llDate < llSDate Then
                                                ilDateOk = False
                                            End If
                                        End If
                                    End If
                                Else
                                    If llMoDate < llSDate Then
                                        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                                        If llDate < llSDate Then
                                            ilDateOk = False
                                        End If
                                    End If
                                End If
                                gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                                If (llTime >= llSTime) And (llTime <= llETime) And (tmSdf.lChfCode > 0) And (ilDateOk) Then
                                    Select Case slAction
                                        Case "A"    'Advertiser/Contract
                                            ilIndex = -1
                                            For ilLoop = 0 To UBound(tmSpotActionInfo) - 1 Step 1
                                                If tmSdf.lChfCode = tmSpotActionInfo(ilLoop).lChfCode Then
                                                    ilIndex = ilLoop
                                                    Exit For
                                                End If
                                            Next ilLoop
                                            If ilIndex = -1 Then
                                                ilIndex = UBound(tmSpotActionInfo)
                                                tmSpotActionInfo(ilIndex).lChfCode = tmSdf.lChfCode
                                                tmSpotActionInfo(ilIndex).iSchdSpotCount = 0
                                                tmSpotActionInfo(ilIndex).iMissedSpotCount = 0
                                                tmSpotActionInfo(ilIndex).iCancelSpotCount = 0
                                                tmSpotActionInfo(ilIndex).iHiddenSpotCount = 0
                                                ReDim Preserve tmSpotActionInfo(0 To ilIndex + 1) As SPOTACTIONINFO
                                            End If
                                            If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                                                If tmSdf.sSpotType = "X" Then
                                                    tmSpotActionInfo(ilIndex).iSchdFillCount = tmSpotActionInfo(ilIndex).iSchdFillCount + 1
                                                Else
                                                    tmSpotActionInfo(ilIndex).iSchdSpotCount = tmSpotActionInfo(ilIndex).iSchdSpotCount + 1
                                                End If
                                            ElseIf tmSdf.sSchStatus = "C" Then
                                                tmSpotActionInfo(ilIndex).iCancelSpotCount = tmSpotActionInfo(ilIndex).iCancelSpotCount + 1
                                            ElseIf tmSdf.sSchStatus = "H" Then
                                                tmSpotActionInfo(ilIndex).iHiddenSpotCount = tmSpotActionInfo(ilIndex).iHiddenSpotCount + 1
                                            Else
                                                tmSpotActionInfo(ilIndex).iMissedSpotCount = tmSpotActionInfo(ilIndex).iMissedSpotCount + 1
                                            End If
                                        Case "P"  'Pre-empt
                                            '8/27/15: Retain MG and Outside unless specifed to pre-empt
                                            'If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                                            If (tmSdf.sSchStatus = "S") Or ((tmSdf.sSchStatus = "G") And (ckcPIPreemptMG.Value = vbChecked)) Or ((tmSdf.sSchStatus = "O") And (ckcPIPreemptMG.Value = vbChecked)) Then
                                                If mCheckSpot() Then
                                                    If mCheckCntr(slAction) Then
                                                        If mCheckLock(tmSdf) Then
                                                            tmSdfAction(UBound(tmSdfAction)) = tmSdf
                                                            ReDim Preserve tmSdfAction(0 To UBound(tmSdfAction) + 1) As SDF
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Case "R"  'Reschedule
                                            If (tmSdf.sSchStatus = "M") Then
                                                If mCheckCntr(slAction) Then
                                                    lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                                                    'ReDim Preserve lgReschSdfCode(1 To UBound(lgReschSdfCode) + 1) As Long
                                                    ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
                                                End If
                                            End If
                                        Case "C"  'Cancel
                                            If rbcCCancel(0).Value Then
                                                If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                                                    If mCheckSpot() Then
                                                        If mCheckCntr(slAction) Then
                                                            If mCheckLock(tmSdf) Then
                                                                tmSdfAction(UBound(tmSdfAction)) = tmSdf
                                                                ReDim Preserve tmSdfAction(0 To UBound(tmSdfAction) + 1) As SDF
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            If rbcCCancel(1).Value Then
                                                If (tmSdf.sSchStatus = "M") Then
                                                    If mCheckCntr(slAction) Then
                                                        tmSdfAction(UBound(tmSdfAction)) = tmSdf
                                                        ReDim Preserve tmSdfAction(0 To UBound(tmSdfAction) + 1) As SDF
                                                    End If
                                                End If
                                            End If
                                            If rbcCCancel(2).Value Then
                                                If (tmSdf.sSchStatus = "C") Then
                                                    If mCheckCntr(slAction) Then
                                                        tmSdfAction(UBound(tmSdfAction)) = tmSdf
                                                        ReDim Preserve tmSdfAction(0 To UBound(tmSdfAction) + 1) As SDF
                                                    End If
                                                End If
                                            End If
                                        Case "H"  'Hide Remnant
                                            If rbcRemnant(0).Value Then
                                                If (tmSdf.sSchStatus = "M") Then
                                                    If mCheckCntr(slAction) Then
                                                        tmSdfAction(UBound(tmSdfAction)) = tmSdf
                                                        ReDim Preserve tmSdfAction(0 To UBound(tmSdfAction) + 1) As SDF
                                                    End If
                                                End If
                                            End If
                                            If rbcRemnant(1).Value Then
                                                If (tmSdf.sSchStatus = "H") Then
                                                    If mCheckCntr(slAction) Then
                                                        tmSdfAction(UBound(tmSdfAction)) = tmSdf
                                                        ReDim Preserve tmSdfAction(0 To UBound(tmSdfAction) + 1) As SDF
                                                    End If
                                                End If
                                            End If
                                        Case "E"  'Erase Fill
                                            If (tmSdf.sSpotType = "X") Then
                                                If mCheckSpot() Then
                                                    If mCheckCntr(slAction) Then
                                                        If mCheckLock(tmSdf) Then
                                                            tmSdfAction(UBound(tmSdfAction)) = tmSdf
                                                            ReDim Preserve tmSdfAction(0 To UBound(tmSdfAction) + 1) As SDF
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Case "F"  'Front Load
                                            If (tmSdf.sSchStatus = "S") And (tmSdf.sSpotType <> "X") And (tmSdf.iGameNo <= 0) Then
                                                If mCheckSpot() Then
                                                    If mCheckCntr(slAction) Then
                                                        If mCheckLock(tmSdf) Then
                                                            tmSdfAction(UBound(tmSdfAction)) = tmSdf
                                                            ReDim Preserve tmSdfAction(0 To UBound(tmSdfAction) + 1) As SDF
                                                        End If
                                                    End If
                                                End If
                                            End If
                                    End Select
                                End If
                            End If
                            ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                            Do While ilRet = BTRV_ERR_REJECT_COUNT
                                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                            Loop
                        Loop
                        If (slAction = "P") Or ((slAction = "C") And (rbcCCancel(0).Value)) Or (slAction = "E") Or (slAction = "F") Then
                            If ckcRemoveLock.Value = vbChecked Then
                                ilRet = mUnlockOperation(1)
                            End If
                        End If
                        If slAction <> "A" Then
                            If slAction = "P" Then
                                For ilLoop = 0 To UBound(tmSdfAction) - 1 Step 1
                                    tmSdf = tmSdfAction(ilLoop)
                                    If mLockDay(ilVefCode, llLastLockValue, llLockRecCode) Then
                                        lmSsfDate = 0
                                        lmSsfRecPos = 0
                                        'If tmSdf.sSpotType <> "X" Then
                                        If ((tmSdf.sSpotType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmSdf.sSpotType = "M") And (tgSpf.sSchdPromo <> "Y")) Or ((tmSdf.sSpotType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmSdf.sSpotType = "Q") Or (tmSdf.sSpotType = "R") Or (tmSdf.sSpotType = "X") Then
                                            ilRet = gChgSchSpot("D", hmSdf, tmSdf, hmSmf, imGameNo, tmSmf, hmSsf, tmSsf, lmSsfDate, lmSsfRecPos, hmSxf, hmGsf, hmGhf)
                                        Else
                                            ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, imGameNo, tmSmf, hmSsf, tmSsf, lmSsfDate, lmSsfRecPos, hmSxf, hmGsf, hmGhf)
                                            If ilRet Then
                                                If ckcPIReSchd.Value = vbChecked Then
                                                    lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                                                    'ReDim Preserve lgReschSdfCode(1 To UBound(lgReschSdfCode) + 1) As Long
                                                    ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
                                                End If
                                            End If
                                        End If
                                    Else
                                        ilLockFailed = True
                                    End If
                                Next ilLoop
                                If llLockRecCode <> 0 Then
                                    ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                    llLastLockValue = 0
                                    llLockRecCode = 0
                                End If
                                If ilLockFailed Then
                                    ilRet = MsgBox("Not all spots Pre-empted as someone else is working on day(s)", vbOKOnly + vbInformation, "Pre-empt")
                                End If
                            End If
                            If (slAction = "R") Or (slAction = "P") Then
                                'If ((slAction = "R") And (UBound(lgReschSdfCode) > 1)) Or ((slAction = "P") And (ckcPIReSchd.Value = vbChecked) And (UBound(lgReschSdfCode) > 1)) Then  'Reschedule
                                If ((slAction = "R") And (UBound(lgReschSdfCode) > LBound(lgReschSdfCode))) Or ((slAction = "P") And (ckcPIReSchd.Value = vbChecked) And (UBound(lgReschSdfCode) > LBound(lgReschSdfCode))) Then  'Reschedule
                                    igReschNoPasses = 2
                                    If tgVpf(imVpfIndex).sMoveLLD <> "Y" Then
                                        igSetEarliestDateAsToday = 1
                                    Else
                                        igSetEarliestDateAsToday = 2
                                    End If
                                    If slAction = "R" Then
                                        ilRet = gReSchSpots(False, 0, "YYYYYYY", 0, 86400)
                                    End If
                                    If slAction = "P" Then
                                        ilRet = gReSchSpots(False, 2, slDays, llSTime, llETime)
                                    End If
                                End If
                            End If
                            If slAction = "E" Then
                                For ilLoop = 0 To UBound(tmSdfAction) - 1 Step 1
                                    tmSdf = tmSdfAction(ilLoop)
                                    If mLockDay(ilVefCode, llLastLockValue, llLockRecCode) Then
                                        imGameNo = tmSdf.iGameNo
                                        lmSsfDate = 0
                                        lmSsfRecPos = 0
                                        ilRet = gChgSchSpot("D", hmSdf, tmSdf, hmSmf, imGameNo, tmSmf, hmSsf, tmSsf, lmSsfDate, lmSsfRecPos, hmSxf, hmGsf, hmGhf)
                                    Else
                                        ilLockFailed = True
                                    End If
                                Next ilLoop
                                If llLockRecCode <> 0 Then
                                    ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                    llLastLockValue = 0
                                    llLockRecCode = 0
                                    If ilLockFailed Then
                                        ilRet = MsgBox("Not all spots Erased as someone else is working on day(s)", vbOKOnly + vbInformation, "Erase")
                                    End If
                                End If
                            End If
                            If slAction = "C" Then
                                For ilLoop = 0 To UBound(tmSdfAction) - 1 Step 1
                                    tmSdf = tmSdfAction(ilLoop)
                                    imGameNo = tmSdf.iGameNo
                                    lmSsfDate = 0
                                    lmSsfRecPos = 0
                                    If rbcCCancel(0).Value Then
                                        If mLockDay(ilVefCode, llLastLockValue, llLockRecCode) Then
                                            'If tmSdf.sSpotType <> "X" Then
                                            If ((tmSdf.sSpotType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmSdf.sSpotType = "M") And (tgSpf.sSchdPromo <> "Y")) Or ((tmSdf.sSpotType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmSdf.sSpotType = "Q") Or (tmSdf.sSpotType = "R") Or (tmSdf.sSpotType = "X") Then
                                                ilRet = gChgSchSpot("D", hmSdf, tmSdf, hmSmf, imGameNo, tmSmf, hmSsf, tmSsf, lmSsfDate, lmSsfRecPos, hmSxf, hmGsf, hmGhf)
                                            Else
                                                ilRet = gChgSchSpot("C", hmSdf, tmSdf, hmSmf, imGameNo, tmSmf, hmSsf, tmSsf, lmSsfDate, lmSsfRecPos, hmSxf, hmGsf, hmGhf)
                                            End If
                                        Else
                                            ilLockFailed = True
                                        End If
                                    Else
                                        tmSdfSrchKey3.lCode = tmSdf.lCode
                                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet = BTRV_ERR_NONE Then
                                            If ckcPIRetainReason.Value = vbChecked Then
                                                igMnfMissed = tmSdf.iMnfMissed
                                            End If
                                            If (rbcCCancel(1).Value) And (tmSdf.sSchStatus = "M") Then
                                                tmSdf.sSchStatus = "C"
                                                tmSdf.iMnfMissed = igMnfMissed
                                                ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                                            End If
                                            If (rbcCCancel(2).Value) And (tmSdf.sSchStatus = "C") Then
                                                tmSdf.sSchStatus = "M"
                                                ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                                            End If
                                        End If
                                    End If
                                Next ilLoop
                                If rbcCCancel(0).Value Then
                                    If llLockRecCode <> 0 Then
                                        ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                        llLastLockValue = 0
                                        llLockRecCode = 0
                                    End If
                                    If ilLockFailed Then
                                        ilRet = MsgBox("Not all spots Canceled as someone else is working on day(s)", vbOKOnly + vbInformation, "Cancel")
                                    End If
                                End If
                            End If
                            If slAction = "H" Then
                                For ilLoop = 0 To UBound(tmSdfAction) - 1 Step 1
                                    tmSdf = tmSdfAction(ilLoop)
                                    imGameNo = tmSdf.iGameNo
                                    lmSsfDate = 0
                                    lmSsfRecPos = 0
                                    tmSdfSrchKey3.lCode = tmSdf.lCode
                                    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet = BTRV_ERR_NONE Then
                                        igMnfMissed = tmSdf.iMnfMissed
                                        If (rbcRemnant(0).Value) And (tmSdf.sSchStatus = "M") Then
                                            tmSdf.sSchStatus = "H"
                                            tmSdf.iMnfMissed = igMnfMissed
                                            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                                        End If
                                        If (rbcRemnant(1).Value) And (tmSdf.sSchStatus = "H") Then
                                            tmSdf.sSchStatus = "M"
                                            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                                        End If
                                    End If
                                Next ilLoop
                            End If
                            If slAction = "F" Then
                                If UBound(tmSdfAction) > 0 Then  'Front Load
                                    For ilFLMax = 1 To ilMaxFrontLoad Step 1
                                        igReschNoPasses = 1
                                        If tgVpf(imVpfIndex).sMoveLLD <> "Y" Then
                                            igSetEarliestDateAsToday = 1
                                        Else
                                            igSetEarliestDateAsToday = 2
                                        End If
                                        ReDim lmFLChfCode(0 To 0) As Long
                                        For ilLoop = 0 To UBound(tmSdfAction) - 1 Step 1
                                            ilFound = False
                                            For ilChf = 0 To UBound(lmFLChfCode) - 1 Step 1
                                                If lmFLChfCode(ilChf) = tmSdfAction(ilLoop).lChfCode Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Next ilChf
                                            If Not ilFound Then
                                                If tmSdfAction(ilLoop).lCode > 0 Then
                                                    'ReDim lgReschSdfCode(1 To 2) As Long
                                                    ReDim lgReschSdfCode(0 To 1) As Long
                                                    'lgReschSdfCode(1) = tmSdfAction(ilLoop).lCode
                                                    lgReschSdfCode(0) = tmSdfAction(ilLoop).lCode
                                                    ilRet = gReSchSpots(False, 1, slDays, llTSTime, llTETime)
                                                    If ilRet Then
                                                        tmSdfSrchKey3.lCode = tmSdfAction(ilLoop).lCode
                                                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        If ilRet = BTRV_ERR_NONE Then
                                                            If (tmSdfAction(ilLoop).iDate(0) <> tmSdf.iDate(0)) Or (tmSdfAction(ilLoop).iDate(1) <> tmSdf.iDate(1)) Then
                                                                lmFLChfCode(UBound(lmFLChfCode)) = tmSdf.lChfCode
                                                                ReDim Preserve lmFLChfCode(0 To UBound(lmFLChfCode) + 1) As Long
                                                            End If
                                                        End If
                                                    End If
                                                    tmSdfAction(ilLoop).lCode = -1
                                                End If
                                            End If
                                        Next ilLoop
                                        If UBound(lmFLChfCode) <= 0 Then
                                            Exit For
                                        End If
                                    Next ilFLMax
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next ilVef
    If (slAction = "R") Or ((slAction = "P") And (ckcPIReSchd.Value = vbChecked)) Then   'Reschedule
        igSetEarliestDateAsToday = 1
        igReschNoPasses = 1
        gCloseSchFiles
    End If
    mBuildSpots = True
End Function

Private Function mPreemptOperation() As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    
    If Not mCheckFromDateTime() Then
        mPreemptOperation = False
        Exit Function
    End If
    'If (rbcRetainLock(0).Value = False) And (rbcRetainLock(1).Value = False) Then
    '    MsgBox "Clear or Bypass Lock Information is Not Defined", vbOkOnly + vbApplicationModal, "Cancel"
    '    mPreemptOperation = False
    '    Exit Function
    'End If
    ReDim lmChfCode(0 To 0) As Long
    ReDim imPPCode(0 To 0) As Integer
    imAdvtCntrAll = False
    imPPAll = False
    If lbcProductProtection.SelCount <= 0 Then
        mBuildChfCode
    Else
        mBuildPPCode
    End If
    If (UBound(lmChfCode) = 0) And (UBound(imPPCode) = 0) And (Not imAdvtCntrAll) And (Not imPPAll) Then
        MsgBox "Either Advertiser/Contract or Product Protection must been selected", vbOKOnly + vbApplicationModal, "Preempt"
        mPreemptOperation = False
        Exit Function
    End If
    If lbcVehicles.SelCount <= 0 Then
        MsgBox "No Vehicle(s) has been selected", vbOKOnly + vbApplicationModal, "Preempt"
        mPreemptOperation = False
        Exit Function
    End If
    If lbcMissedCancelReason.SelCount <= 0 Then
        MsgBox "Missed Reason needs to be selected", vbOKOnly + vbApplicationModal, "Preempt"
        mPreemptOperation = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbHourglass
    slNameCode = tmMissedCancelCode(lbcMissedCancelReason.ListIndex).sKey
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    igMnfMissed = Val(slCode)
    ilRet = mBuildSpots("P")
    If ilRet Then
        If ckcPILock.Value = vbChecked Then
            ilRet = mLockOperation()
        End If
    End If
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbDefault
    mPreemptOperation = True
End Function

Private Function mEraseOperation() As Integer
    Dim ilRet As Integer
    
    If Not mCheckFromDateTime() Then
        mEraseOperation = False
        Exit Function
    End If
    'If (rbcRetainLock(0).Value = False) And (rbcRetainLock(1).Value = False) Then
    '    MsgBox "Clear or Bypass Lock Information is Not Defined", vbOkOnly + vbApplicationModal, "Cancel"
    '    mEraseOperation = False
    '    Exit Function
    'End If
    ReDim lmChfCode(0 To 0) As Long
    ReDim imPPCode(0 To 0) As Integer
    imAdvtCntrAll = False
    imPPAll = False
    mBuildChfCode
    If (UBound(lmChfCode) = 0) And (Not imAdvtCntrAll) Then
        MsgBox "Advertiser/Contract must been selected", vbOKOnly + vbApplicationModal, "Preempt"
        mEraseOperation = False
        Exit Function
    End If
    If lbcVehicles.SelCount <= 0 Then
        MsgBox "No Vehicle(s) has been selected", vbOKOnly + vbApplicationModal, "Erase"
        mEraseOperation = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbHourglass
    ilRet = mBuildSpots("E")
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbDefault
    mEraseOperation = True
End Function

Private Function mCancelOperation() As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    
    If Not mCheckFromDateTime() Then
        mCancelOperation = False
        Exit Function
    End If
    If (rbcCCancel(0).Value = False) And (rbcCCancel(1).Value = False) And (rbcCCancel(2).Value = False) Then
        MsgBox "Cancel Information is Not Defined", vbOKOnly + vbApplicationModal, "Cancel"
        mCancelOperation = False
        Exit Function
    End If
    If (rbcCCancel(0).Value) Then
        'If (rbcRetainLock(0).Value = False) And (rbcRetainLock(1).Value = False) Then
        '    MsgBox "Clear or Bypass Lock Information is Not Defined", vbOkOnly + vbApplicationModal, "Cancel"
        '    mCancelOperation = False
        '    Exit Function
        'End If
    End If
    ReDim lmChfCode(0 To 0) As Long
    ReDim imPPCode(0 To 0) As Integer
    imAdvtCntrAll = False
    imPPAll = False
    If lbcProductProtection.SelCount <= 0 Then
        mBuildChfCode
    Else
        mBuildPPCode
    End If
    If (UBound(lmChfCode) = 0) And (UBound(imPPCode) = 0) And (Not imAdvtCntrAll) And (Not imPPAll) Then
        MsgBox "Either Advertiser/Contract or Product Protection must been selected", vbOKOnly + vbApplicationModal, "Cancel"
        mCancelOperation = False
        Exit Function
    End If
    If lbcVehicles.SelCount <= 0 Then
        MsgBox "No Vehicle(s) has been selected", vbOKOnly + vbApplicationModal, "Cancel"
        mCancelOperation = False
        Exit Function
    End If
    If ckcPIRetainReason.Value = vbUnchecked Then
        If lbcMissedCancelReason.SelCount <= 0 Then
            MsgBox "Canceled Reason needs to be selected", vbOKOnly + vbApplicationModal, "Cancel"
            mCancelOperation = False
            Exit Function
        End If
    End If
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbHourglass
    slNameCode = tmMissedCancelCode(lbcMissedCancelReason.ListIndex).sKey
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    igMnfMissed = Val(slCode)
    ilRet = mBuildSpots("C")
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbDefault
    mCancelOperation = True
End Function

Private Function mFrontLoadOperation() As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slFSDate As String
    Dim slFEDate As String
    Dim slTSDate As String
    Dim slTEDate As String
    
    If Not mCheckFromDateTime() Then
        mFrontLoadOperation = False
        Exit Function
    End If
    If Not mCheckToDateTime() Then
        mFrontLoadOperation = False
        Exit Function
    End If
    'If (rbcRetainLock(0).Value = False) And (rbcRetainLock(1).Value = False) Then
    '    MsgBox "Clear or Bypass Lock Information is Not Defined", vbOkOnly + vbApplicationModal, "Cancel"
    '    mFrontLoadOperation = False
    '    Exit Function
    'End If
    slFSDate = Trim$(edcFDDropDown(0).Text)
    slFEDate = Trim$(edcFDDropDown(1).Text)
    slTSDate = Trim$(edcTDDropDown(0).Text)
    slTEDate = Trim$(edcTDDropDown(1).Text)
    If gDateValue(slTEDate) >= gDateValue(slFSDate) Then
        MsgBox "To Dates overlap From Dates which is not allowed", vbOKOnly + vbApplicationModal, "Front Load"
        mFrontLoadOperation = False
        Exit Function
    End If
    If gObtainPrevMonday(slTSDate) <> gObtainPrevMonday(slFSDate) Then
        MsgBox "From and To Dates must be within same Week", vbOKOnly + vbApplicationModal, "Front Load"
        mFrontLoadOperation = False
        Exit Function
    End If
    If gObtainPrevMonday(slFEDate) <> gObtainPrevMonday(slFSDate) Then
        MsgBox "From Dates must be within same Week", vbOKOnly + vbApplicationModal, "Front Load"
        mFrontLoadOperation = False
        Exit Function
    End If
    If Trim$(edcMaxFrontLoad.Text) = "" Then
        MsgBox "Max Number of Spots/Advertiser/Vehicle must be defined", vbOKOnly + vbApplicationModal, "Front Load"
        mFrontLoadOperation = False
        Exit Function
    End If
    ReDim lmChfCode(0 To 0) As Long
    ReDim imPPCode(0 To 0) As Integer
    imAdvtCntrAll = False
    imPPAll = False
    If lbcProductProtection.SelCount <= 0 Then
        mBuildChfCode
    Else
        mBuildPPCode
    End If
    If (UBound(lmChfCode) = 0) And (UBound(imPPCode) = 0) And (Not imAdvtCntrAll) And (Not imPPAll) Then
        MsgBox "Either Advertiser/Contract or Product Protection must been selected", vbOKOnly + vbApplicationModal, "Front Load"
        mFrontLoadOperation = False
        Exit Function
    End If
    If lbcVehicles.SelCount <= 0 Then
        MsgBox "No Vehicle(s) has been selected", vbOKOnly + vbApplicationModal, "Front Load"
        mFrontLoadOperation = False
        Exit Function
    End If
    ilRet = MsgBox("This will Distort the Contract Horizontal Balance, Continue?", vbApplicationModal + vbInformation + vbYesNo, "Question")
    If ilRet = vbNo Then
        mFrontLoadOperation = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbHourglass
    igMnfMissed = 0
    ilRet = mBuildSpots("F")
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbDefault
    mFrontLoadOperation = True
End Function


Private Function mCheckLock(tlSdf As SDF) As Integer
    If ckcRemoveSpot.Value = vbChecked Then
        mCheckLock = True
    Else
        If gIsLocked(tlSdf) Then
            mCheckLock = False
        Else
            mCheckLock = True
        End If
    End If
End Function

Private Sub mFillAdvtCntrGrid(slAction As String)
    Dim llCol As Long
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilAdf As Integer
    Dim ilRet As Integer
    Dim ilShowAdvtCntr As Integer
    
    grdAdvtCntr.Redraw = False
    grdAdvtCntr.Row = 0
    For llCol = ADVTINDEX To MISSINDEX Step 1
        grdAdvtCntr.Col = llCol
        grdAdvtCntr.CellBackColor = LIGHTBLUE
    Next llCol
    grdAdvtCntr.rowHeight(0) = fgBoxGridH + 15
    If UBound(tmSpotActionInfo) <= 0 Then
        grdAdvtCntr.Redraw = True
        Exit Sub
    End If
    llRow = grdAdvtCntr.FixedRows
    For ilLoop = 0 To UBound(tmSpotActionInfo) - 1 Step 1
        tmChfSrchKey0.lCode = tmSpotActionInfo(ilLoop).lChfCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilAdf = gBinarySearchAdf(tmChf.iAdfCode)
            If ilAdf <> -1 Then
                ilShowAdvtCntr = False
                If slAction = "P" Then
                    If (tmSpotActionInfo(ilLoop).iSchdSpotCount > 0) Or (tmSpotActionInfo(ilLoop).iSchdFillCount > 0) Then
                        ilShowAdvtCntr = True
                    End If
                ElseIf slAction = "R" Then
                    If tmSpotActionInfo(ilLoop).iMissedSpotCount > 0 Then
                        ilShowAdvtCntr = True
                    End If
                ElseIf slAction = "C" Then
                    If rbcCCancel(0).Value Then
                        If (tmSpotActionInfo(ilLoop).iSchdSpotCount > 0) Or (tmSpotActionInfo(ilLoop).iSchdFillCount > 0) Then
                            ilShowAdvtCntr = True
                        End If
                    ElseIf rbcCCancel(1).Value Then
                        If tmSpotActionInfo(ilLoop).iMissedSpotCount > 0 Then
                            ilShowAdvtCntr = True
                        End If
                    ElseIf rbcCCancel(2).Value Then
                        If tmSpotActionInfo(ilLoop).iCancelSpotCount > 0 Then
                            ilShowAdvtCntr = True
                        End If
                    End If
                ElseIf slAction = "H" Then
                    If rbcRemnant(0).Value Then
                        If tmSpotActionInfo(ilLoop).iMissedSpotCount > 0 Then
                            ilShowAdvtCntr = True
                        End If
                    ElseIf rbcRemnant(1).Value Then
                        If tmSpotActionInfo(ilLoop).iHiddenSpotCount > 0 Then
                            ilShowAdvtCntr = True
                        End If
                    End If
                ElseIf slAction = "E" Then
                    If tmSpotActionInfo(ilLoop).iSchdFillCount > 0 Then
                        ilShowAdvtCntr = True
                    End If
                ElseIf slAction = "F" Then
                    If tmSpotActionInfo(ilLoop).iSchdSpotCount > 0 Then
                        ilShowAdvtCntr = True
                    End If
                End If
                
                If ilShowAdvtCntr Then
                    If llRow >= grdAdvtCntr.Rows Then
                        grdAdvtCntr.AddItem ""
                    End If
                    grdAdvtCntr.rowHeight(llRow) = fgBoxGridH + 15
                    For llCol = ADVTINDEX To MISSINDEX Step 1
                        grdAdvtCntr.Row = llRow
                        grdAdvtCntr.Col = llCol
                        grdAdvtCntr.CellBackColor = vbWhite
                        grdAdvtCntr.CellForeColor = vbBlue
                    Next llCol
                    If Trim$(tmChf.sProduct) <> "" Then
                        grdAdvtCntr.TextMatrix(llRow, ADVTINDEX) = Trim$(tgCommAdf(ilAdf).sName) & "/" & Trim$(tmChf.sProduct)
                    Else
                        grdAdvtCntr.TextMatrix(llRow, ADVTINDEX) = Trim$(tgCommAdf(ilAdf).sName)
                    End If
                    grdAdvtCntr.TextMatrix(llRow, CNTRINDEX) = Trim$(str(tmChf.lCntrNo))
                    If slAction = "P" Then
                        grdAdvtCntr.TextMatrix(llRow, SCHDINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iSchdSpotCount + tmSpotActionInfo(ilLoop).iSchdFillCount))
                        grdAdvtCntr.TextMatrix(llRow, MISSINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iMissedSpotCount))
                    ElseIf slAction = "R" Then
                        grdAdvtCntr.TextMatrix(llRow, SCHDINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iMissedSpotCount))
                        grdAdvtCntr.TextMatrix(llRow, MISSINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iSchdSpotCount))
                    ElseIf slAction = "C" Then
                        If rbcCCancel(0).Value Then
                            grdAdvtCntr.TextMatrix(llRow, SCHDINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iSchdSpotCount + tmSpotActionInfo(ilLoop).iSchdFillCount))
                            grdAdvtCntr.TextMatrix(llRow, MISSINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iCancelSpotCount))
                        ElseIf rbcCCancel(1).Value Then
                            grdAdvtCntr.TextMatrix(llRow, SCHDINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iMissedSpotCount))
                            grdAdvtCntr.TextMatrix(llRow, MISSINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iCancelSpotCount))
                        ElseIf rbcCCancel(2).Value Then
                            grdAdvtCntr.TextMatrix(llRow, SCHDINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iCancelSpotCount))
                            grdAdvtCntr.TextMatrix(llRow, MISSINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iMissedSpotCount))
                        End If
                    ElseIf slAction = "H" Then
                        If rbcRemnant(0).Value Then
                            grdAdvtCntr.TextMatrix(llRow, SCHDINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iMissedSpotCount))
                            grdAdvtCntr.TextMatrix(llRow, MISSINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iHiddenSpotCount))
                        ElseIf rbcRemnant(1).Value Then
                            grdAdvtCntr.TextMatrix(llRow, SCHDINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iHiddenSpotCount))
                            grdAdvtCntr.TextMatrix(llRow, MISSINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iMissedSpotCount))
                        End If
                    ElseIf slAction = "E" Then
                        grdAdvtCntr.TextMatrix(llRow, SCHDINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iSchdFillCount))
                    ElseIf slAction = "F" Then
                        grdAdvtCntr.TextMatrix(llRow, SCHDINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iSchdSpotCount))
                        grdAdvtCntr.TextMatrix(llRow, MISSINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).iMissedSpotCount))
                    End If
                    grdAdvtCntr.TextMatrix(llRow, CHFCODEINDEX) = Trim$(str(tmSpotActionInfo(ilLoop).lChfCode))
                    llRow = llRow + 1
                End If
            End If
        End If
    Next ilLoop
    mAdvtCntrSortCol ADVTINDEX
    grdAdvtCntr.Row = 0
    grdAdvtCntr.Col = CHFCODEINDEX
    grdAdvtCntr.Redraw = True

End Sub

Private Function mLockDay(ilVefCode As Integer, llLastLockValue As Long, llLockRecCode As Long) As Integer
    Dim slUserName As String
    Dim llLockValue As Long
    Dim ilRet As Integer
    Dim llSdfDate As Long
    
    mLockDay = True
    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSdfDate
    imGameNo = tmSdf.iGameNo
    If imGameNo = 0 Then
        llLockValue = 65536 * ilVefCode + llSdfDate
    Else
        llLockValue = 65536 * ilVefCode + imGameNo
    End If
    If (llLastLockValue <> llLockValue) Then
        If llLockRecCode <> 0 Then
            ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
            llLastLockValue = 0
        End If
        llLockRecCode = gCreateLockRec(hmRlf, "S", "S", llLockValue, True, slUserName)
        If llLockRecCode > 0 Then
            llLastLockValue = llLockValue
        Else
            mLockDay = False
        End If
    End If
End Function

Private Function mCheckSpot() As Integer
    Dim llDate As Long
    
    mCheckSpot = True
    If (tgVpf(imVpfIndex).sMoveLLD <> "Y") Then
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
        If (llDate <= lmLastLogDate) Then
            mCheckSpot = False
        End If
    End If
End Function

Private Function mRemnantOperation() As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    
    If Not mCheckFromDateTime("H") Then
        mRemnantOperation = False
        Exit Function
    End If
    If (rbcRemnant(0).Value = False) And (rbcRemnant(1).Value = False) Then
        MsgBox "Remnant Information is Not Defined", vbOKOnly + vbApplicationModal, "Cancel"
        mRemnantOperation = False
        Exit Function
    End If
    ReDim lmChfCode(0 To 0) As Long
    ReDim imPPCode(0 To 0) As Integer
    imAdvtCntrAll = False
    imPPAll = False
    If lbcProductProtection.SelCount <= 0 Then
        mBuildChfCode
    Else
        mBuildPPCode
    End If
    If (UBound(lmChfCode) = 0) And (UBound(imPPCode) = 0) And (Not imAdvtCntrAll) And (Not imPPAll) Then
        MsgBox "Either Advertiser/Contract or Product Protection must been selected", vbOKOnly + vbApplicationModal, "Cancel"
        mRemnantOperation = False
        Exit Function
    End If
    If lbcVehicles.SelCount <= 0 Then
        MsgBox "No Vehicle(s) has been selected", vbOKOnly + vbApplicationModal, "Cancel"
        mRemnantOperation = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbHourglass
    ilRet = mBuildSpots("H")
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAdvtCntr, grdAdvtCntr, vbDefault
    mRemnantOperation = True
End Function
