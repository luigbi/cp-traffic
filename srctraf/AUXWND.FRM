VERSION 5.00
Begin VB.Form AuxWnd 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9000
   ClientLeft      =   135
   ClientTop       =   135
   ClientWidth     =   13500
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9000
   ScaleWidth      =   13500
   Visible         =   0   'False
   Begin VB.PictureBox plcWindow 
      Appearance      =   0  'Flat
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
      Picture         =   "AUXWND.frx":0000
      ScaleHeight     =   8921.397
      ScaleMode       =   0  'User
      ScaleWidth      =   13515
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13520
      Begin VB.PictureBox pbcHor 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   7245
         Left            =   13560
         Picture         =   "AUXWND.frx":18B862
         ScaleHeight     =   7185
         ScaleWidth      =   10170
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   3000
         Visible         =   0   'False
         Width           =   10230
         Begin VB.Label lacGeneral 
            BackStyle       =   0  'Transparent
            Caption         =   "Advertisers"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   720
            TabIndex        =   81
            Top             =   1920
            Width           =   1170
         End
         Begin VB.Label lacSales 
            BackStyle       =   0  'Transparent
            Caption         =   "Ad Server Items"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   11
            Left            =   5040
            TabIndex        =   79
            Top             =   3015
            Width           =   1455
         End
         Begin VB.Label lacSales 
            BackStyle       =   0  'Transparent
            Caption         =   "Split Net"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   10
            Left            =   8085
            TabIndex        =   77
            Top             =   3360
            Width           =   870
         End
         Begin VB.Label lacAccount 
            BackStyle       =   0  'Transparent
            Caption         =   "Tax Table"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   5280
            TabIndex        =   75
            Top             =   5910
            Width           =   975
         End
         Begin VB.Label lacGeneral 
            BackStyle       =   0  'Transparent
            Caption         =   "Feed Names"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   8
            Left            =   5415
            TabIndex        =   72
            Top             =   1875
            Width           =   1155
         End
         Begin VB.Label laLists 
            BackStyle       =   0  'Transparent
            Caption         =   "Lists"
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
            Left            =   10365
            TabIndex        =   71
            Top             =   45
            Width           =   570
         End
         Begin VB.Label lacAccount 
            BackStyle       =   0  'Transparent
            Caption         =   "Tran Type"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   5
            Left            =   6480
            TabIndex        =   66
            Top             =   5910
            Width           =   1020
         End
         Begin VB.Label lacAccount 
            BackStyle       =   0  'Transparent
            Caption         =   "EDI"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   4
            Left            =   645
            TabIndex        =   56
            Top             =   5910
            Width           =   360
         End
         Begin VB.Label lacAccount 
            BackStyle       =   0  'Transparent
            Caption         =   "Lock Boxes"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   2655
            TabIndex        =   60
            Top             =   5910
            Width           =   1110
         End
         Begin VB.Label lacAccount 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Sorts"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   1200
            TabIndex        =   58
            Top             =   5910
            Width           =   1245
         End
         Begin VB.Label lacAccount 
            BackStyle       =   0  'Transparent
            Caption         =   "NTR Types"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   4020
            TabIndex        =   64
            Top             =   5910
            Width           =   1095
         End
         Begin VB.Label lacGeneral 
            BackStyle       =   0  'Transparent
            Caption         =   "Missed/Cancel Reason"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   9
            Left            =   6900
            TabIndex        =   62
            Top             =   1875
            Width           =   2085
         End
         Begin VB.Label lacGeneric 
            BackStyle       =   0  'Transparent
            Caption         =   "Generic"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   7980
            TabIndex        =   54
            Top             =   4995
            Width           =   900
         End
         Begin VB.Label lacPref 
            BackStyle       =   0  'Transparent
            Caption         =   "User"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   6390
            TabIndex        =   52
            Top             =   5010
            Width           =   435
         End
         Begin VB.Label lacPref 
            BackStyle       =   0  'Transparent
            Caption         =   "Site"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   5760
            TabIndex        =   50
            Top             =   5010
            Width           =   375
         End
         Begin VB.Label lacCopy 
            BackStyle       =   0  'Transparent
            Caption         =   "Announcer Names"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   645
            TabIndex        =   46
            Top             =   5010
            Width           =   1770
         End
         Begin VB.Label lacCopy 
            BackStyle       =   0  'Transparent
            Caption         =   "Media Definitions"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   2610
            TabIndex        =   48
            Top             =   5010
            Width           =   1695
         End
         Begin VB.Label lacProg 
            BackStyle       =   0  'Transparent
            Caption         =   "Genre Names"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   4
            Left            =   6300
            TabIndex        =   44
            Top             =   4170
            Width           =   1365
         End
         Begin VB.Label lacProg 
            BackStyle       =   0  'Transparent
            Caption         =   "Feed Type"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   4905
            TabIndex        =   42
            Top             =   4170
            Width           =   1230
         End
         Begin VB.Label lacProg 
            BackStyle       =   0  'Transparent
            Caption         =   "Event Types"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   3510
            TabIndex        =   40
            Top             =   4170
            Width           =   1170
         End
         Begin VB.Label lacProg 
            BackStyle       =   0  'Transparent
            Caption         =   "Event Names"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   2055
            TabIndex        =   38
            Top             =   4170
            Width           =   1260
         End
         Begin VB.Label lacProg 
            BackStyle       =   0  'Transparent
            Caption         =   "Avail Names"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   645
            TabIndex        =   36
            Top             =   4170
            Width           =   1185
         End
         Begin VB.Label lacSales 
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Teams"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   6675
            TabIndex        =   34
            Top             =   3360
            Width           =   1275
         End
         Begin VB.Label lacSales 
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Sources"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   5100
            TabIndex        =   32
            Top             =   3360
            Width           =   1395
         End
         Begin VB.Label lacSales 
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Regions"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3510
            TabIndex        =   30
            Top             =   3360
            Width           =   1395
         End
         Begin VB.Label lacSales 
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Offices"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   2040
            TabIndex        =   28
            Top             =   3360
            Width           =   1290
         End
         Begin VB.Label lacSales 
            BackStyle       =   0  'Transparent
            Caption         =   "Salespeople"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   4
            Left            =   630
            TabIndex        =   26
            Top             =   3360
            Width           =   1200
         End
         Begin VB.Label lacSales 
            BackStyle       =   0  'Transparent
            Caption         =   "Revenue Sets"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   5
            Left            =   7905
            TabIndex        =   24
            Top             =   3015
            Width           =   1350
         End
         Begin VB.Label lacSales 
            BackStyle       =   0  'Transparent
            Caption         =   "Research"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   7
            Left            =   6735
            TabIndex        =   22
            Top             =   3015
            Width           =   930
         End
         Begin VB.Label lacSales 
            BackStyle       =   0  'Transparent
            Caption         =   "Custom Demos"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   8
            Left            =   3330
            TabIndex        =   20
            Top             =   3015
            Width           =   1470
         End
         Begin VB.Label lacSales 
            BackStyle       =   0  'Transparent
            Caption         =   "Competitors"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   9
            Left            =   1920
            TabIndex        =   18
            Top             =   3015
            Width           =   1170
         End
         Begin VB.Label lacSales 
            BackStyle       =   0  'Transparent
            Caption         =   "Boilerplate"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   6
            Left            =   630
            TabIndex        =   16
            Top             =   3015
            Width           =   1050
         End
         Begin VB.Label lacGeneral 
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Groups"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   5
            Left            =   7455
            TabIndex        =   14
            Top             =   2220
            Width           =   1485
         End
         Begin VB.Label lacGeneral 
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicles"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   6375
            TabIndex        =   12
            Top             =   2220
            Width           =   885
         End
         Begin VB.Label lacGeneral 
            BackStyle       =   0  'Transparent
            Caption         =   "Program Exclusions"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   4245
            TabIndex        =   10
            Top             =   2220
            Width           =   1935
         End
         Begin VB.Label lacGeneral 
            BackStyle       =   0  'Transparent
            Caption         =   "Product Protection"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   2250
            TabIndex        =   8
            Top             =   2220
            Width           =   1800
         End
         Begin VB.Label lacGeneral 
            BackStyle       =   0  'Transparent
            Caption         =   "Potential Codes"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   7
            Left            =   645
            TabIndex        =   6
            Top             =   2220
            Width           =   1515
         End
         Begin VB.Label lacGeneral 
            BackStyle       =   0  'Transparent
            Caption         =   "Business Categories"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   6
            Left            =   3135
            TabIndex        =   4
            Top             =   1875
            Width           =   1980
         End
         Begin VB.Label lacGeneral 
            BackStyle       =   0  'Transparent
            Caption         =   "Agencies"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   1995
            TabIndex        =   2
            Top             =   1875
            Width           =   885
         End
      End
      Begin VB.CommandButton cmcAccount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tax Table"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   8250
         TabIndex        =   74
         TabStop         =   0   'False
         Tag             =   "List of reasons spots can be missed, for inclusion on invoices"
         Top             =   6960
         Width           =   2205
      End
      Begin VB.CommandButton cmcGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vehicles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   6335
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "List of all stations, networks, channels, or other airing entities"
         Top             =   2430
         Width           =   1830
      End
      Begin VB.CommandButton cmcGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vehicles Groups"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   5
         Left            =   8250
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "List of all stations, networks, channels, or other airing entities"
         Top             =   2430
         Width           =   2205
      End
      Begin VB.CommandButton cmcGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Agencies"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   2325
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "List of Advertising Agencies"
         Top             =   1845
         Width           =   1830
      End
      Begin VB.CommandButton cmcGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Advertisers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   645
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "List of Advertisers"
         Top             =   1845
         Width           =   1590
      End
      Begin VB.CommandButton cmcGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Product Protection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   2325
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "List of Product protection codes"
         Top             =   2430
         Width           =   1830
      End
      Begin VB.CommandButton cmcGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Business Categories"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   6
         Left            =   4245
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "List of Product protection codes"
         Top             =   1845
         Width           =   1995
      End
      Begin VB.CommandButton cmcGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Potential Codes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   7
         Left            =   645
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "List of Product protection codes"
         Top             =   2430
         Width           =   1590
      End
      Begin VB.CommandButton cmcGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Program Exclusions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   4
         Left            =   4245
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "List of Product protection codes"
         Top             =   2430
         Width           =   1995
      End
      Begin VB.CommandButton cmcSales 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sales Sources"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   6335
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   4290
         Width           =   1830
      End
      Begin VB.CommandButton cmcSales 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sales Regions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   4245
         TabIndex        =   31
         TabStop         =   0   'False
         Tag             =   "List of Market locations of Sales Reps"
         Top             =   4290
         Width           =   1995
      End
      Begin VB.CommandButton cmcSales 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sales Offices"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   2325
         TabIndex        =   29
         TabStop         =   0   'False
         Tag             =   "List of teams used to sub-group salespeople"
         Top             =   4290
         Width           =   1830
      End
      Begin VB.CommandButton cmcSales 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sales Teams"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   8250
         TabIndex        =   35
         TabStop         =   0   'False
         Tag             =   "List of Local and Rep salespeople"
         Top             =   4290
         Width           =   2205
      End
      Begin VB.CommandButton cmcSales 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Salespeople"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   4
         Left            =   645
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "List of special types of revenue totals not ordinarily covered"
         Top             =   4290
         Width           =   1590
      End
      Begin VB.CommandButton cmcSales 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Research"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   7
         Left            =   8250
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3705
         Width           =   2205
      End
      Begin VB.CommandButton cmcSales 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Custom Demos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   8
         Left            =   4245
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3705
         Width           =   1995
      End
      Begin VB.CommandButton cmcSales 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Competitors"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   9
         Left            =   2325
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3705
         Width           =   1830
      End
      Begin VB.CommandButton cmcSales 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Revenue Sets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   5
         Left            =   10560
         TabIndex        =   25
         TabStop         =   0   'False
         Tag             =   "List of special types of revenue totals not ordinarily covered"
         Top             =   3705
         Width           =   1830
      End
      Begin VB.CommandButton cmcSales 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Boilerplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   6
         Left            =   645
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "List of special types of revenue totals not ordinarily covered"
         Top             =   3705
         Width           =   1590
      End
      Begin VB.CommandButton cmcSales 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ad Server Items"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   6335
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   3705
         Width           =   1830
      End
      Begin VB.CommandButton cmcProg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Event Types"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   4245
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   5580
         Width           =   1995
      End
      Begin VB.CommandButton cmcProg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Event Names"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   2325
         MaskColor       =   &H00000942&
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "Update list of valid program names"
         Top             =   5580
         Width           =   1830
      End
      Begin VB.CommandButton cmcProg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Avail Names"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   645
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   5580
         Width           =   1590
      End
      Begin VB.CommandButton cmcProg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Feed Names"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   6335
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   5580
         Width           =   1830
      End
      Begin VB.CommandButton cmcProg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Genre Names"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   4
         Left            =   8250
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   5580
         Width           =   2205
      End
      Begin VB.CommandButton cmcGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Missed/Cancel Reasons"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   9
         Left            =   8250
         TabIndex        =   63
         TabStop         =   0   'False
         Tag             =   "List of reasons spots can be missed, for inclusion on invoices"
         Top             =   1845
         Width           =   2205
      End
      Begin VB.CommandButton cmcAccount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "NTR Types"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   6335
         TabIndex        =   65
         TabStop         =   0   'False
         Tag             =   "List of types of Item Billing allowed"
         Top             =   6960
         Width           =   1830
      End
      Begin VB.CommandButton cmcAccount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Invoice Sorts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   2325
         MaskColor       =   &H00000942&
         TabIndex        =   59
         TabStop         =   0   'False
         Tag             =   "List of addresses for payment of invoices"
         Top             =   6960
         Width           =   1830
      End
      Begin VB.CommandButton cmcAccount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lock Boxes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   4245
         TabIndex        =   61
         TabStop         =   0   'False
         Tag             =   "List of names and addresses for Electronic invoices"
         Top             =   6960
         Width           =   1995
      End
      Begin VB.CommandButton cmcAccount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "EDI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   4
         Left            =   645
         TabIndex        =   57
         TabStop         =   0   'False
         Tag             =   "List of addresses for payment of invoices"
         Top             =   6960
         Width           =   1590
      End
      Begin VB.CommandButton cmcAccount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tran Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   5
         Left            =   10545
         TabIndex        =   67
         TabStop         =   0   'False
         Tag             =   "List of addresses for payment of invoices"
         Top             =   6960
         Width           =   1830
      End
      Begin VB.CommandButton cmcCopy 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Media Definitions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   2655
         TabIndex        =   49
         TabStop         =   0   'False
         Tag             =   "Create and modify rules governing types of material aired"
         Top             =   8340
         Width           =   1860
      End
      Begin VB.CommandButton cmcCopy 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Announcer Names"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   645
         TabIndex        =   47
         TabStop         =   0   'False
         Tag             =   "Update list of announcers on commercials"
         Top             =   8340
         Width           =   1935
      End
      Begin VB.CommandButton cmcPref 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Site"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   7485
         TabIndex        =   51
         TabStop         =   0   'False
         Tag             =   "Client rules and specifications"
         Top             =   8340
         Width           =   1110
      End
      Begin VB.CommandButton cmcPref 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   8700
         TabIndex        =   53
         TabStop         =   0   'False
         Tag             =   "Names of valid users and what parts of the system they can use"
         Top             =   8340
         Width           =   1110
      End
      Begin VB.CommandButton cmcGeneric 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generic Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   9945
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   8340
         Width           =   1830
      End
      Begin VB.TextBox edcLinkDestDoneMsg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   4065
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.PictureBox pbcClickFocus 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Left            =   -315
         ScaleHeight     =   120
         ScaleWidth      =   60
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   4530
         Width           =   60
      End
      Begin VB.CommandButton cmcGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Feed Names"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   8
         Left            =   6335
         TabIndex        =   73
         TabStop         =   0   'False
         Tag             =   "List of Product protection codes"
         Top             =   1845
         Width           =   1830
      End
      Begin VB.CommandButton cmcSales 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Split Net"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   10
         Left            =   10545
         TabIndex        =   76
         TabStop         =   0   'False
         Tag             =   "List of Local and Rep salespeople"
         Top             =   4290
         Width           =   1830
      End
      Begin VB.Label lacOptionsGroup 
         BackStyle       =   0  'Transparent
         Caption         =   "OPTIONS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   7485
         TabIndex        =   87
         Top             =   7800
         Width           =   3855
      End
      Begin VB.Label lacCopyGroup 
         BackStyle       =   0  'Transparent
         Caption         =   "COPY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   645
         TabIndex        =   86
         Top             =   7800
         Width           =   3855
      End
      Begin VB.Label lacAccountingGroup 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNTING"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   645
         TabIndex        =   85
         Top             =   6420
         Width           =   3855
      End
      Begin VB.Label lacProgrammingGroup 
         BackStyle       =   0  'Transparent
         Caption         =   "PROGRAMMING"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   645
         TabIndex        =   84
         Top             =   5040
         Width           =   3855
      End
      Begin VB.Label lacSalesGroup 
         BackStyle       =   0  'Transparent
         Caption         =   "SALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   645
         TabIndex        =   83
         Top             =   3165
         Width           =   3375
      End
      Begin VB.Label lacGeneralGroup 
         BackStyle       =   0  'Transparent
         Caption         =   "GENERAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   645
         TabIndex        =   82
         Top             =   1305
         Width           =   3375
      End
      Begin VB.Image cmcCSLogo 
         Height          =   945
         Left            =   120
         Picture         =   "AUXWND.frx":19D091
         Stretch         =   -1  'True
         Top             =   121
         Width           =   7650
      End
      Begin VB.Image cmcStartJob 
         Height          =   600
         Left            =   12600
         Picture         =   "AUXWND.frx":59154F
         Stretch         =   -1  'True
         ToolTipText     =   "Back to Jobs Screen"
         Top             =   242
         Width           =   600
      End
      Begin VB.Label imcOutline 
         BackColor       =   &H0080FF80&
         Enabled         =   0   'False
         Height          =   435
         Left            =   600
         TabIndex        =   80
         Top             =   1800
         Visible         =   0   'False
         Width           =   1575
      End
   End
End
Attribute VB_Name = "AuxWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of AUXWND.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: AuxWnd.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Traffic input screen code
Option Explicit
Option Compare Text

Dim imPrevListIndex As Integer
Dim imPrevListGroup As Integer

Private Sub cmcAccount_Click(Index As Integer)
    Dim slStr As String

    gShowHelpMess tgListHelp(), -1
    If igGGFlag = 0 Then
        Exit Sub
    End If
    If Not igSiteSet Then
        MsgBox "Please Define Site Options", vbOKOnly + vbInformation, "Counterpoint"
        Exit Sub
    End If
    Select Case Index
        Case 0  'Tax Table
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            sgCommandStr = slStr
            On Error Resume Next
            On Error Resume Next
            TaxTable.Show vbModal
            slStr = sgDoneMsg
        Case 1  'Item Billing
            'If Not gWinRoom(igNoLJWinRes(ITEMBILLINGTYPESLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "I"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 2  'Invoice sort
            'If Not gWinRoom(igNoLJWinRes(INVOICESORTLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "V"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 3  'Lock Boxes
            'If Not gWinRoom(igNoLJWinRes(LOCKBOXESLIST)) Then
            '    Exit Sub
            'End If
            sgArfCallType = "L"
            'Screen.MousePointer = vbHourglass  'Wait
            igArfCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "NmAddr.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            NmAddr.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 4  'EDI Service
            'If Not gWinRoom(igNoLJWinRes(EDISERVICESLIST)) Then
            '    Exit Sub
            'End If
            sgArfCallType = "A"
            'Screen.MousePointer = vbHourglass  'Wait
            igArfCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "NmAddr.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            NmAddr.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 5  'Transaction type
            'If Not gWinRoom(igNoLJWinRes(INVOICESORTLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "Y"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
    End Select
    Form_Activate
    Exit Sub

    'Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Private Sub cmcAccount_GotFocus(Index As Integer)
    gCtrlGotFocus cmcAccount(Index)
    mListHelpMess Index + MISSEDREASONSLIST
End Sub

Private Sub cmcAccount_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move cmcAccount(Index).Left - 90, cmcAccount(Index).Top - 90, cmcAccount(Index).Width, cmcAccount(Index).height
    imcOutline.Visible = True
    If (imPrevListGroup <> 0) Or (imPrevListIndex <> Index - MISSEDREASONSLIST) Then
        mResetColor
    End If
    If igWinStatus(Index + MISSEDREASONSLIST) = 1 Then
        cmcAccount(Index).ToolTipText = "View Only"
        'lacAccount(Index).ForeColor = vbYellow     'vbGold; vbYellow
        imPrevListGroup = 0
        imPrevListIndex = Index
        imcOutline.BackColor = vbYellow
    Else
        cmcAccount(Index).ToolTipText = "Update Allowed"
        lacAccount(Index).ForeColor = vbBlack
        imcOutline.BackColor = &H80FF80
    End If
End Sub

Private Sub cmcCopy_Click(Index As Integer)
    Dim slStr As String
    If igGGFlag = 0 Then
        Exit Sub
    End If
    gShowHelpMess tgListHelp(), -1
    If Not igSiteSet Then
        MsgBox "Please Define Site Options", vbOKOnly + vbInformation, "Counterpoint"
        Exit Sub
    End If
    Select Case Index
        Case 0  'Media Definitions
            'If Not gWinRoom(igNoLJWinRes(MEDIADEFINITIONSLIST)) Then
            '    Exit Sub
            'End If
            'Screen.MousePointer = vbHourglass  'Wait
            igMcdCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igMcdCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igMcdCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "Media.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            Media.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 1  'Announcer Names
            'If Not gWinRoom(igNoLJWinRes(ANNOUNCERNAMESLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "A"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
    End Select
    Form_Activate
    Exit Sub

    'Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Private Sub cmcCopy_GotFocus(Index As Integer)
    gCtrlGotFocus cmcCopy(Index)
    mListHelpMess Index + MEDIADEFINITIONSLIST
End Sub

Private Sub cmcCopy_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move cmcCopy(Index).Left - 90, cmcCopy(Index).Top - 90, cmcCopy(Index).Width, cmcCopy(Index).height
    imcOutline.Visible = True
    If (imPrevListGroup <> 1) Or (imPrevListIndex <> Index - MEDIADEFINITIONSLIST) Then
        mResetColor
    End If
    If igWinStatus(Index + MEDIADEFINITIONSLIST) = 1 Then
        cmcCopy(Index).ToolTipText = "View Only"
        'lacCopy(Index).ForeColor = vbYellow     'vbGold; vbYellow
        imPrevListGroup = 1
        imPrevListIndex = Index
        imcOutline.BackColor = vbYellow
    Else
        cmcCopy(Index).ToolTipText = "Update Allowed"
        lacCopy(Index).ForeColor = vbBlack
        imcOutline.BackColor = &H80FF80
    End If
End Sub

Private Sub cmcCSLogo_Click()
    Dim ilAns As Integer

    ilAns = MsgBox(sgCSVersion, 0, "Counterpoint Software")
End Sub
Private Sub cmcGeneral_Click(Index As Integer)
    Dim slStr As String
    mResetColor
    If igGGFlag = 0 Then
        Exit Sub
    End If
    gShowHelpMess tgListHelp(), -1
    If Not igSiteSet Then
        MsgBox "Please Define Site Options", vbOKOnly + vbInformation, "Counterpoint"
        Exit Sub
    End If
    Select Case Index
        Case 0  'Vehicles
            'If Not gWinRoom(igNoLJWinRes(VEHICLESLIST)) Then
            '    Exit Sub
            'End If
            'Screen.MousePointer = vbHourglass  'Wait
            igVehCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            'If igTestSystem Then
                '10071 not used
'                If igDemo Then
'                    slStr = "Traffic^Test^HELP^DEMO\" & sgUserName & "\" & Trim$(str$(igVehCallSource))
'                Else
'                    slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igVehCallSource))
'                End If
                
            'Else
            '    If igDemo Then
            '        slStr = "Traffic^Prod^HELP^DEMO\" & sgUserName & "\" & Trim$(Str$(igVehCallSource))
            '    Else
            '        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(igVehCallSource))
            '    End If
            'End If
            'lgShellRet = Shell(sgExePath & "Vehicle.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            '10071
'            sgCommandStr = slStr
'            On Error Resume Next
'            Vehicle.Show vbModal
'            slStr = sgDoneMsg
            
            On Error Resume Next
            gCallVehicleProject Me
            
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbHourglass
            'Done in vehicles when erased and saved because of repopulation
'            sgMVefStamp = "~"   'Force Read
'            ilRet = gObtainVef()
'            sgVpfStamp = "~"    'Force read
'            ilRet = gVpfRead(AuxWnd)
'            gObtainVirtVehList
'            gAnyClustersDef
'            gAnyRepDef
            '10071 not needed
            'Screen.MousePointer = vbDefault    'Default
        Case 1  'Agencies
            'If Not gWinRoom(igNoLJWinRes(AGENCIESLIST)) Then
            '    Exit Sub
            'End If
            'Screen.MousePointer = vbHourglass  'Wait
            igAgyCallSource = CALLNONE
            igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igAgyCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igAgyCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "Agency.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            Agency.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 2  'Advertisers
            'If Not gWinRoom(igNoLJWinRes(ADVERTISERSLIST)) Then
            '    Exit Sub
            'End If
            'Screen.MousePointer = vbHourglass  'Wait
            igAdvtCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igAdvtCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igAdvtCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "Advt.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            Advt.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 3  'Product Protection
            'If Not gWinRoom(igNoLJWinRes(COMPETITIVESLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "C"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 4  'Exclusions
            'If Not gWinRoom(igNoLJWinRes(EXCLUSIONSLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "X"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 5  'Vehicle Group
            'If Not gWinRoom(igNoLJWinRes(VEHICLEGROUPSLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "H"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 6  'Business Categories
            'If Not gWinRoom(igNoLJWinRes(BUSCATEGORIESLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "B"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 7  'Potential Code
            'If Not gWinRoom(igNoLJWinRes(POTENTIALCODESLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "P"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 8  'Feed Name
            igFdNmCallSource = CALLNONE
            igChildDone = False
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igFdNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igFdNmCallSource))
            End If
            sgCommandStr = slStr
            On Error Resume Next
            FeedName.Show vbModal
            slStr = sgDoneMsg
        Case 9  'Missed Reason
            'If Not gWinRoom(igNoLJWinRes(MISSEDREASONSLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "M"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
    End Select
    imcOutline.Visible = False
    Form_Activate
    Exit Sub

    'Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Private Sub cmcGeneral_GotFocus(Index As Integer)
    gCtrlGotFocus cmcGeneral(Index)
    mListHelpMess Index + VEHICLESLIST
End Sub

Private Sub cmcGeneral_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move cmcGeneral(Index).Left - 90, cmcGeneral(Index).Top - 90, cmcGeneral(Index).Width, cmcGeneral(Index).height
    imcOutline.Visible = True
    If (imPrevListGroup <> 2) Or (imPrevListIndex <> Index - VEHICLESLIST) Then
        mResetColor
    End If
    If igWinStatus(Index + VEHICLESLIST) = 1 Then
        cmcGeneral(Index).ToolTipText = "View Only"
        'lacGeneral(Index).ForeColor = vbYellow     'vbGold; vbYellow
        imPrevListGroup = 2
        imPrevListIndex = Index
        imcOutline.BackColor = vbYellow
    Else
        cmcGeneral(Index).ToolTipText = "Update Allowed"
        lacGeneral(Index).ForeColor = vbBlack
        imcOutline.BackColor = &H80FF80
    End If
End Sub

Private Sub cmcGeneric_Click()
    Dim slStr As String        'General string
    If igGGFlag = 0 Then
        Exit Sub
    End If
    igRptCallType = GENERICBUTTON
    If igTestSystem Then
        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType))
    Else
        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType))
    End If
    sgCommandStr = slStr
    On Error Resume Next
    RptSelCreditStatus.Show vbModal         '7-6-15 changed to remove modules from traffic due to out of memory error
    slStr = sgDoneMsg
    imcOutline.Visible = False
End Sub

Private Sub cmcGeneric_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move cmcGeneric.Left - 90, cmcGeneric.Top - 90, cmcGeneric.Width, cmcGeneric.height
    imcOutline.Visible = True
    imcOutline.BackColor = &H80FF80
End Sub

Private Sub cmcPref_Click(Index As Integer)
    Dim ilRet As Integer
    Dim slStr As String
    
    Dim ilShell As Integer
    Dim dlShellRet As Double
    Dim slCommandStr As String
    Dim ilPos As Integer
    Dim slDate As String
    Dim ilLoop As Integer

    mResetColor
    gShowHelpMess tgListHelp(), -1
    If Index <> 0 Then
        If Not igSiteSet Then
            MsgBox "Please Define Site Options", vbOKOnly + vbInformation, "Counterpoint"
            Exit Sub
        End If
        'Dan M 1/11/10  need to allow guide to get into users even if no vehicle defined.
        'If Not igVehicleDefined Then
        If Not igVehicleDefined And Index <> 1 Then
            MsgBox "Please Define Vehicle", vbOKOnly + vbInformation, "Counterpoint"
            Exit Sub
        End If
    End If
    Select Case Index
        Case 0  'Site
            'If Not gWinRoom(igNoLJWinRes(SITELIST)) Then
            '    Exit Sub
            'End If
            'Screen.MousePointer = vbHourglass  'Wait
            igSiteCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igSiteCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igSiteCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "SiteOpt.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            'SiteOpt.Show vbModal
            mLoadSiteOpt
            
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
            Screen.MousePointer = vbHourglass
            sgSpfStamp = "~"    'Force read
            gSpfRead
            Screen.MousePointer = vbDefault
            'If tgSpf.sExport = "N" Then
            '    Traffic!mnuFileExport.Enabled = False
            'Else
                Traffic!mnuFileExport.Enabled = True
            'End If
            'If tgSpf.sImport = "N" Then
            '    Traffic!mnuFileImport.Enabled = False
            'Else
                Traffic!mnuFileImport.Enabled = True
            'End If
            'If tgSpf.sUsingMatrix = "N" Then           '4-24-03
            If ((Asc(tgSpf.sUsingFeatures) And MATRIXEXPORT) <> MATRIXEXPORT) And ((Asc(tgSaf(0).sFeatures1) And MATRIXCAL) <> MATRIXCAL) Then
                Traffic!mnuExportMatrix.Enabled = False
            Else
                Traffic!mnuExportMatrix.Enabled = True
            End If
            
            If ((Asc(tgSaf(0).sFeatures2) And TABLEAUEXPORT) <> TABLEAUEXPORT) And ((Asc(tgSaf(0).sFeatures2) And TABLEAUCAL) <> TABLEAUCAL) Then
                Traffic!mnuExportTableau.Enabled = False
            Else
                Traffic!mnuExportTableau.Enabled = True
            End If

            If tgSpf.sSystemType <> "R" Then
                If tgSpf.sGUseAffFeed = "Y" Then
                    Traffic!mnuExportStnFeed.Enabled = True
                Else
                    Traffic!mnuExportStnFeed.Enabled = False
                End If
            Else
                Traffic!mnuExportStnFeed.Enabled = False
            End If
            
            If (Asc(tgSpf.sUsingFeatures) And REVENUEEXPORT) <> REVENUEEXPORT Then
                Traffic!mnuExportProjections.Enabled = False
            Else
                Traffic!mnuExportProjections.Enabled = True
            End If

            If ((Asc(tgSpf.sUsingFeatures6) And GETPAIDEXPORT) <> GETPAIDEXPORT) Then
                Traffic!mnuGetPaid.Enabled = False
            Else
                Traffic!mnuGetPaid.Enabled = True
            End If

            If tgSpf.sDemoEstAllowed <> "Y" Then
                Traffic!mnuImportSatellite.Enabled = False
                Traffic!mnuImportMark.Enabled = True
                Traffic!mnuImportIMS.Enabled = True
                Traffic!mnuImportRadar.Enabled = True
            Else
                Traffic!mnuImportSatellite.Enabled = True
                Traffic!mnuImportMark.Enabled = False
                Traffic!mnuImportIMS.Enabled = False
                Traffic!mnuImportRadar.Enabled = False
            End If
            
            If (Asc(tgSaf(0).sFeatures6) And RABCALENDAR) = RABCALENDAR Then        '1-29-20
                Traffic!mnuExportRAB.Enabled = True
            Else
                Traffic!mnuExportRAB.Enabled = False
            End If
            
            'TTP 10205 - 6/21/21 - JW - WO Invoice Export - Export Menu
            'If (tgSpfx.iInvExpFeature And 1) > 0 Then
            If (tgSpfx.iInvExpFeature And INVEXP_AUDACYWO) > 0 Then
                Traffic!mnuExportWOInvoice.Enabled = True
            Else
                Traffic!mnuExportWOInvoice.Enabled = False
            End If
            'TTP 10233
            'If (tgSpfx.iInvExpFeature And 2) > 0 Then
            If (tgSpfx.iInvExpFeature And INVEXP_AUDACYLINE) > 0 Then
                Traffic!mnuExpCntrLine.Enabled = True
            Else
                Traffic!mnuExpCntrLine.Enabled = False
            End If
            
            '6/16/09:  Removed menu item Import Contracts because out of memory
            If tgSpf.sImptCntr = "P" Then           '11-9-04
            '    Traffic!mnuImportOrder.Enabled = False
            Else
            '    Traffic!mnuImportOrder.Enabled = True
            End If

'            If tgSpf.sSystemType <> "R" Then
'                cmcGeneral(8).Enabled = False
'                lacGeneral(8).Visible = False
'                cmcGeneral(7).Left = 4860
'                lacGeneral(7).Left = cmcGeneral(7).Left - 30
'            Else
'                cmcGeneral(8).Enabled = True
'                lacGeneral(8).Visible = True
'                cmcGeneral(8).Left = cmcGeneral(7).Left
'                lacGeneral(8).Left = cmcGeneral(8).Left - 30
'                cmcGeneral(7).Left = cmcGeneral(8).Left + cmcGeneral(8).Width + 30
'                lacGeneral(7).Left = cmcGeneral(7).Left - 30
'            End If
            mSetControls
            ilRet = gAlertForceCheck()
'        Case 1  'Vehicle
'            igVehOptCallSource = CALLNONE
'            sgVehNameToVehOpt = ""
'            Screen.MousePointer = vbHourGlass  'Wait
'            VehOpt.Show vbModal
'            Screen.MousePointer = vbDefault    'Default
'        Case 2  'User
        Case 1  'User
''            If Not gWinRoom(igNoLJWinRes(USERLIST)) Then
''                Exit Sub
''            End If
''            'Screen.MousePointer = vbHourGlass  'Wait
'''            Traffic!cmcAuxWnd.Value = False
''            UserOpt.Show vbModal
'''            Traffic!cmcAuxWnd.Value = True
''            'Screen.MousePointer = vbDefault    'Default
'            'If Not gWinRoom(igNoLJWinRes(USERLIST)) Then
'            '    Exit Sub
'            'End If
'            'Screen.MousePointer = vbHourglass  'Wait
'            igUrfCallSource = CALLNONE
'            igChildDone = False
'            'Traffic.edcLinkSrceHelpMsg.Text = ""
'            If igTestSystem Then
'                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igUrfCallSource))
'            Else
'                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igUrfCallSource))
'            End If
'            'lgShellRet = Shell(sgExePath & "UserOpt.Exe " & slStr, 1)
'            'Traffic.Enabled = False
'            'Screen.MousePointer = vbDefault  'Wait
'            'Do While Not igChildDone
'            '    DoEvents
'            'Loop
'            sgCommandStr = slStr
'            On Error Resume Next
'            UserOpt.Show vbModal
'            slStr = sgDoneMsg
'            'Traffic.Enabled = True
'            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
'            'For ilLoop = 0 To 10
'            '    DoEvents
'            'Next ilLoop
'            Screen.MousePointer = vbHourglass
        
            'dlShellRet = Shell(sgExePath & "ExportProj.Exe /IniLoc:" & CurDir$ & " /UserInput", vbMinimizedNoFocus)
            If igTestSystem Then
                slCommandStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
            Else
                slCommandStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
            End If
            If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
                ilPos = InStr(1, slCommandStr, "Guide", vbTextCompare)
                If ilPos > 0 Then
                    slCommandStr = Left(slCommandStr, ilPos - 1) & "CSI" & Mid(slCommandStr, ilPos + 5)
                End If
            End If
            'Dan M 9/20/10 problems in v57 reports.exe running GetCsiName
            'slDate = Trim$(gGetCSIName("SYSDate"))
            slDate = gCSIGetName()
            If slDate <> "" Then
                'use slDate when writing to file later
                slDate = " /D:" & slDate
                'slCommandStr = slCommandStr & " /D:" & slDate
                slCommandStr = slCommandStr & slDate
            End If
            slCommandStr = slCommandStr & " /ULF:" & lgUlfCode
            slCommandStr = slCommandStr & "/IniLoc:" & CurDir$ & " /UserInput"
            Screen.MousePointer = vbDefault
            Traffic.WindowState = vbMinimized
            gShellAndWait Traffic, sgExePath & "UserOpt.exe " & slCommandStr, vbNormalFocus, False    'vbFalse
            Traffic.WindowState = vbMaximized
            sgUrfStamp = "~"    'Force read
            'gUrfRead AuxWnd, sgUserName, True, tgUrf(), False  'Obtain user records
            gUrfRead AuxWnd, sgUserName, False, tgUrf(), False  'Obtain user records
            gInitSuperUser tgUrf(0)
            Screen.MousePointer = vbDefault
    End Select
    imcOutline.Visible = False
    Form_Activate
    Exit Sub

    'Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Private Sub cmcPref_GotFocus(Index As Integer)
    gCtrlGotFocus cmcPref(Index)
    mListHelpMess Index + SITELIST
End Sub

Private Sub cmcPref_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move cmcPref(Index).Left - 90, cmcPref(Index).Top - 90, cmcPref(Index).Width, cmcPref(Index).height
    imcOutline.Visible = True
    If (imPrevListGroup <> 3) Or (imPrevListIndex <> Index - SITELIST) Then
        mResetColor
    End If
    If igWinStatus(Index + SITELIST) = 1 Then
        cmcPref(Index).ToolTipText = "View Only"
        'lacPref(Index).ForeColor = vbYellow     'vbGold; vbYellow
        imPrevListGroup = 3
        imPrevListIndex = Index
        imcOutline.BackColor = vbYellow
    Else
        cmcPref(Index).ToolTipText = "Update Allowed"
        lacPref(Index).ForeColor = vbBlack
        imcOutline.BackColor = &H80FF80
    End If
End Sub

Private Sub cmcProg_Click(Index As Integer)
    Dim hlEtf As Integer        'Etf handle
    Dim llNoRec As Long
    Dim ilRet As Integer
    Dim slStr As String
    mResetColor
    If igGGFlag = 0 Then
        Exit Sub
    End If
    gShowHelpMess tgListHelp(), -1
    If Not igSiteSet Then
        MsgBox "Please Define Site Options", vbOKOnly + vbInformation, "Counterpoint"
        Exit Sub
    End If
    Select Case Index
        Case 0  'Event Types
            'If Not gWinRoom(igNoLJWinRes(EVENTTYPESLIST)) Then
            '    Exit Sub
            'End If
            'Screen.MousePointer = vbHourglass  'Wait
            igETypeCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igETypeCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igETypeCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "EType.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            EType.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 1  'Event Names
            If Not igVehicleDefined Then
                MsgBox "Please Define Vehicle", vbOKOnly + vbInformation, "Counterpoint"
                Exit Sub
            End If
            If Not igEtfRecsCreated Then    'Check if event records created
                hlEtf = CBtrvTable(ONEHANDLE)
                ilRet = btrOpen(hlEtf, "", sgDBPath & "Etf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                llNoRec = btrRecords(hlEtf) 'Obtain number of records
                ilRet = btrClose(hlEtf)
                btrDestroy hlEtf
                If llNoRec > 0 Then
                    igEtfRecsCreated = True
                Else
                    MsgBox "Please Define Event Types", vbOKOnly + vbInformation, "Counterpoint"
                    Exit Sub
                End If
            End If
            'If Not gWinRoom(igNoLJWinRes(EVENTNAMESLIST)) Then
            '    Exit Sub
            'End If
            'Screen.MousePointer = vbHourglass  'Wait
            igENameCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igENameCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igENameCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "EName.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            EName.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 2  'Avail Names
            'If Not gWinRoom(igNoLJWinRes(AVAILNAMESLIST)) Then
            '    Exit Sub
            'End If
            'Screen.MousePointer = vbHourglass  'Wait
            igANmCallSource = CALLNONE
            igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igANmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igANmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "AName.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            AName.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 3  'Feed Types
            'If Not gWinRoom(igNoLJWinRes(FEEDTYPESLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "N" 'Feed
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 4  'Genre Names
            'If Not gWinRoom(igNoLJWinRes(GENRESLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "E"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
    End Select
    imcOutline.Visible = False
    Form_Activate
    Exit Sub

    'Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Private Sub cmcProg_GotFocus(Index As Integer)
    gCtrlGotFocus cmcProg(Index)
    mListHelpMess Index + EVENTTYPESLIST
End Sub

Private Sub cmcProg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move cmcProg(Index).Left - 90, cmcProg(Index).Top - 90, cmcProg(Index).Width, cmcProg(Index).height
    imcOutline.Visible = True
    If (imPrevListGroup <> 4) Or (imPrevListIndex <> Index - EVENTTYPESLIST) Then
        mResetColor
    End If
    If igWinStatus(Index + EVENTTYPESLIST) = 1 Then
        cmcProg(Index).ToolTipText = "View Only"
        'lacProg(Index).ForeColor = vbYellow     'vbGold; vbYellow
        imPrevListGroup = 4
        imPrevListIndex = Index
        imcOutline.BackColor = vbYellow
    Else
        cmcProg(Index).ToolTipText = "Update Allowed"
        lacProg(Index).ForeColor = vbBlack
        imcOutline.BackColor = &H80FF80
    End If
End Sub

Private Sub cmcSales_Click(Index As Integer)
    Dim slStr As String

    mResetColor
    If igGGFlag = 0 Then
        Exit Sub
    End If
    gShowHelpMess tgListHelp(), -1
    If Not igSiteSet Then
        MsgBox "Please Define Site Options", vbOKOnly + vbInformation, "Counterpoint"
        Exit Sub
    End If
    Select Case Index
        Case 0  'Sales Source
            'If Not gWinRoom(igNoLJWinRes(SALESSOURCESLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "S"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 1  'Sales Region
            'If Not gWinRoom(igNoLJWinRes(SALESREGIONSLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "G"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 2  'Sales Offices
            'If Not gWinRoom(igNoLJWinRes(SALESOFFICESLIST)) Then
            '    Exit Sub
            'End If
            'Screen.MousePointer = vbHourglass  'Wait
            igSofCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igSofCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igSofCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "SOffice.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            SOffice.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 3  'Sales Teams
            'If Not gWinRoom(igNoLJWinRes(SALESTEAMSLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "T"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 4  'Salespeople
            'If Not gWinRoom(igNoLJWinRes(SALESPEOPLELIST)) Then
            '    Exit Sub
            'End If
            'Screen.MousePointer = vbHourglass  'Wait
            igSlfCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igSlfCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igSlfCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "SPerson.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            SPerson.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 5  'Revenue Sets
            'If Not gWinRoom(igNoLJWinRes(REVENUESETSLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "R"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 6  'Boilerplate
            'If Not gWinRoom(igNoLJWinRes(BOILERPLATESLIST)) Then
            '    Exit Sub
            'End If
            'Screen.MousePointer = vbHourglass  'Wait
            igCmmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igCmmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igCmmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "BPlate.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            BPlate.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 7  'Research
            'If Not gWinRoom(igNoLJWinRes(RESEARCHLIST)) Then
            '    Exit Sub
            'End If
            'Screen.MousePointer = vbHourglass  'Wait
            If tgSpf.sGUsePropSys <> "Y" Then
                Exit Sub
            End If
            igAgyCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            '10066 not used
'            If igTestSystem Then
'                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igAgyCallSource))
'            Else
'                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igAgyCallSource))
'            End If
            
            'lgShellRet = Shell(sgExePath & "Research.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            On Error Resume Next
            '10066
            mCallResearchProject
'            If (Asc(tgSpf.sMOFCopyAssign) And RSCHCUSTDEMO) = RSCHCUSTDEMO Then
'                RschByCust.Show vbModal
'            Else
'                Research.Show vbModal
'            End If
            'slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
        Case 8  'Custom Demos
            'If Not gWinRoom(igNoLJWinRes(DEMOSLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "D"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 9  'Competitors
            'If Not gWinRoom(igNoLJWinRes(COMPETITORSLIST)) Then
            '    Exit Sub
            'End If
            sgMnfCallType = "O"
            'Screen.MousePointer = vbHourglass  'Wait
            igMNmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            MultiNm.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
        Case 10 'Split Net
            sgSplitNetworkName = ""
            igAdfCode = 0
            igSplitType = 1
            igSplitChgd = False
            ReDim igSplitLnVefCode(0 To 0) As Integer
            ReDim lgNewRafCntr(0 To 0) As Long
            NetworkSplit.Show vbModal
        Case 11  'Geo Items
            'If Not gWinRoom(igNoLJWinRes(BOILERPLATESLIST)) Then
            '    Exit Sub
            'End If
            'Screen.MousePointer = vbHourglass  'Wait
            igCmmCallSource = CALLNONE
            'igChildDone = False
            'Traffic.edcLinkSrceHelpMsg.Text = ""
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igCmmCallSource))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igCmmCallSource))
            End If
            'lgShellRet = Shell(sgExePath & "BPlate.Exe " & slStr, 1)
            'Traffic.Enabled = False
            'Screen.MousePointer = vbDefault  'Wait
            'Do While Not igChildDone
            '    DoEvents
            'Loop
            sgCommandStr = slStr
            On Error Resume Next
            PodItem.Show vbModal
            slStr = sgDoneMsg
            'Traffic.Enabled = True
            'Traffic.edcLinkSrceHelpMsg.Text = "Ok"
            'For ilLoop = 0 To 10
            '    DoEvents
            'Next ilLoop
            'Screen.MousePointer = vbDefault    'Default
    End Select
    imcOutline.Visible = False
    Form_Activate
    Exit Sub

    'Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Private Sub cmcSales_GotFocus(Index As Integer)
    gCtrlGotFocus cmcSales(Index)
    mListHelpMess Index + SALESSOURCESLIST
End Sub

Private Sub cmcSales_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move cmcSales(Index).Left - 90, cmcSales(Index).Top - 90, cmcSales(Index).Width, cmcSales(Index).height
    imcOutline.Visible = True
    If (imPrevListGroup <> 5) Or (imPrevListIndex <> Index - SALESSOURCESLIST) Then
        mResetColor
    End If
    If igWinStatus(Index + SALESSOURCESLIST) = 1 Then
        cmcSales(Index).ToolTipText = "View Only"
        'lacSales(Index).ForeColor = vbYellow     'vbGold; vbYellow
        imPrevListGroup = 5
        imPrevListIndex = Index
        imcOutline.BackColor = vbYellow
    Else
        cmcSales(Index).ToolTipText = "Update Allowed"
        lacSales(Index).ForeColor = vbBlack
        imcOutline.BackColor = &H80FF80
    End If
End Sub

Private Sub cmcStartJob_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift Then
        Unload AuxWnd
    Else
        If igBasic10Showing Then
            Basic10.SetFocus
        Else
            Basic10.Show vbModeless
        End If
    End If
End Sub

Private Sub Form_Activate()
    If igGGFlag = 0 Then
        plcWindow.Enabled = False
    End If
    gShowPartBranner
    mSetControls
    igAuxWndShowing = True
    If igGGFlag = 0 Then
        'frcGeneral.Enabled = False
        'frcSales.Enabled = False
        'frcProg.Enabled = False
        'frcCopy.Enabled = False
        'frcAccount.Enabled = False
        'frcPref.Enabled = False
        cmcGeneric.Enabled = False
        lacGeneric.Enabled = False
        lacGeneric.FontBold = False
        'Note: plcWindow.Enabled = False, causes whole window not to show
    End If
    Me.KeyPreview = True
    Me.ZOrder 0 'Send to front
    pbcHor.Visible = False
    DoEvents
    pbcHor.Visible = True
End Sub
Private Sub Form_Click()
    '6/3/15: Ignore click event if screen disabled
    If (igGGFlag = 0) Or (pbcClickFocus.Enabled = False) Then
        Exit Sub
    End If
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    'ElseIf KeyCode = KEYF4 Then
    '    Basic10.Show vbModeless
    End If
End Sub

Private Sub Form_Load()
    
    lacGeneralGroup.ForeColor = RGB(131, 60, 11)
    lacSalesGroup.ForeColor = RGB(131, 60, 11)
    lacProgrammingGroup.ForeColor = RGB(131, 60, 11)
    lacAccountingGroup.ForeColor = RGB(131, 60, 11)
    lacCopyGroup.ForeColor = RGB(131, 60, 11)
    lacOptionsGroup.ForeColor = RGB(131, 60, 11)
    
    If igBasicOrAux = 1 Then
        Unload Basic10
    End If
    igAuxWndShowing = True
    igBasicOrAux = 2    'Indicate that aux loaded
    imPrevListGroup = -1
    imPrevListIndex = -1
    'AuxWnd.Caption = "Lists"
    'cmcCSLogo.Picture = IconTraf!imcCSLogo.Picture
    plcWindow.Left = 0
    plcWindow.Top = 0
    AuxWnd.height = plcWindow.height + fgPanelAdj
    AuxWnd.Width = plcWindow.Width + fgPanelAdj
    If lgCurrBPP <= 8 Then
        pbcHor.Picture = Traffic!pbc256List
    End If
    'cmcCSLogo.Picture = IconTraf!imcCSLogo.Picture
    gCenterForm AuxWnd
'    If tgSpf.sSystemType <> "R" Then
'        cmcGeneral(8).Enabled = False
'        lacGeneral(8).Visible = False
'        cmcGeneral(7).Left = 4860
'        lacGeneral(7).Left = cmcGeneral(7).Left - 30
'    Else
'        cmcGeneral(8).Enabled = True
'        lacGeneral(8).Visible = True
'        cmcGeneral(8).Left = cmcGeneral(7).Left
'        lacGeneral(8).Left = cmcGeneral(8).Left - 30
'        cmcGeneral(7).Left = cmcGeneral(8).Left + cmcGeneral(8).Width + 30
'        lacGeneral(7).Left = cmcGeneral(7).Left - 30
'    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
        AuxWnd!plcWindow.Enabled = False
    igAuxWndShowing = False
End Sub
Private Sub mListHelpMess(ilListNo As Integer)
    Select Case ilListNo
        Case VEHICLESLIST  'Vehicle
            gShowHelpMess tgListHelp(), VEHICLESLIST - VEHICLESLIST + 1
        Case AGENCIESLIST  'Agency
            gShowHelpMess tgListHelp(), AGENCIESLIST - VEHICLESLIST + 1
        Case ADVERTISERSLIST  'Advertiser
            gShowHelpMess tgListHelp(), ADVERTISERSLIST - VEHICLESLIST + 1
        Case COMPETITIVESLIST  'Competitive
            gShowHelpMess tgListHelp(), COMPETITIVESLIST - VEHICLESLIST + 1
        Case SALESSOURCESLIST  'Sales source
            gShowHelpMess tgListHelp(), SALESSOURCESLIST - VEHICLESLIST + 1
        Case SALESOFFICESLIST  'Sales Office
            gShowHelpMess tgListHelp(), SALESOFFICESLIST - VEHICLESLIST + 1
        Case SALESTEAMSLIST  'Sales Teams
            gShowHelpMess tgListHelp(), SALESTEAMSLIST - VEHICLESLIST + 1
        Case SALESPEOPLELIST  'Salespersons
            gShowHelpMess tgListHelp(), SALESPEOPLELIST - VEHICLESLIST + 1
        Case REVENUESETSLIST  'Revenue Sets
            gShowHelpMess tgListHelp(), REVENUESETSLIST - VEHICLESLIST + 1
        Case MISSEDREASONSLIST  'Missed reason
            gShowHelpMess tgListHelp(), MISSEDREASONSLIST - VEHICLESLIST + 1
        Case ITEMBILLINGTYPESLIST  'Item billing type
            gShowHelpMess tgListHelp(), ITEMBILLINGTYPESLIST - VEHICLESLIST + 1
        Case INVOICESORTLIST  'Invoice sort type
            gShowHelpMess tgListHelp(), INVOICESORTLIST - VEHICLESLIST + 1
        Case LOCKBOXESLIST 'Lock boxes
            gShowHelpMess tgListHelp(), LOCKBOXESLIST - VEHICLESLIST + 1
        Case EDISERVICESLIST  'Agency DP Service
            gShowHelpMess tgListHelp(), EDISERVICESLIST - VEHICLESLIST + 1
        Case EVENTTYPESLIST  'Event type
            gShowHelpMess tgListHelp(), EVENTTYPESLIST - VEHICLESLIST + 1
        Case EVENTNAMESLIST  'Event Name
            gShowHelpMess tgListHelp(), EVENTNAMESLIST - VEHICLESLIST + 1
        Case AVAILNAMESLIST  'Avail Names
            gShowHelpMess tgListHelp(), AVAILNAMESLIST - VEHICLESLIST + 1
        Case FEEDTYPESLIST  'Feed Types
            gShowHelpMess tgListHelp(), FEEDTYPESLIST - VEHICLESLIST + 1
        Case GENRESLIST  'Sports names
            gShowHelpMess tgListHelp(), GENRESLIST - VEHICLESLIST + 1
        Case MEDIADEFINITIONSLIST  'Media definitions
            gShowHelpMess tgListHelp(), MEDIADEFINITIONSLIST - VEHICLESLIST + 1
        Case ANNOUNCERNAMESLIST  'Annoucer Names
            gShowHelpMess tgListHelp(), ANNOUNCERNAMESLIST - VEHICLESLIST + 1
        Case SITELIST  'Site options
            gShowHelpMess tgListHelp(), SITELIST - VEHICLESLIST + 1
        Case USERLIST 'User option
            gShowHelpMess tgListHelp(), USERLIST - VEHICLESLIST + 1
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetControls                    *
'*                                                     *
'*             Created:10/27/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable/Disable list functions  *
'*                                                     *
'*******************************************************
Private Sub mSetControls()
    Dim ilLoop As Integer
    Dim ilIndex As Integer

    gSpfRead    'Make sure latest image is in memory
    'gSetUsingTraffic
    If (tgSpf.sAExcl = "N") Or (igGGFlag = 0) Then
        igWinStatus(EXCLUSIONSLIST) = 0
    Else
        If tgUrf(0).sWin(EXCLUSIONSLIST) = "I" Then
            igWinStatus(EXCLUSIONSLIST) = 2
        ElseIf tgUrf(0).sWin(EXCLUSIONSLIST) = "V" Then
            igWinStatus(EXCLUSIONSLIST) = 1
        Else
            igWinStatus(EXCLUSIONSLIST) = 0
        End If
    End If
    If tgSpf.sSystemType = "R" Then
        cmcGeneral(8).Visible = True
        lacGeneral(8).Visible = True
        'cmcGeneral(8).Left = cmcGeneral(6).Left + cmcGeneral(6).Width + 30
        'lacGeneral(8).Left = cmcGeneral(8).Left - 30
        'cmcGeneral(9).Left = cmcGeneral(8).Left + cmcGeneral(8).Width + 30
        'lacGeneral(9).Left = cmcGeneral(9).Left - 30
        cmcProg(3).Visible = False
        lacProg(3).Visible = False
        'cmcProg(4).Left = cmcProg(0).Left + cmcProg(0).Width + 30
        'lacProg(4).Left = cmcProg(4).Left - 30
        If tgUrf(0).sWin(FEEDNAMELIST) = "I" Then
            igWinStatus(FEEDNAMELIST) = 2
        ElseIf tgUrf(0).sWin(FEEDNAMELIST) = "V" Then
            igWinStatus(FEEDNAMELIST) = 1
        Else
            igWinStatus(FEEDNAMELIST) = 0
        End If
    Else
        cmcGeneral(8).Enabled = False
        lacGeneral(8).Visible = False
        'cmcGeneral(9).Left = cmcGeneral(6).Left + cmcGeneral(6).Width + 30       '4860
        'lacGeneral(9).Left = cmcGeneral(9).Left - 30
        cmcProg(3).Visible = True
        lacProg(3).Visible = True
        'cmcProg(3).Left = cmcProg(0).Left + cmcProg(0).Width + 30
        'lacProg(3).Left = cmcProg(3).Left - 30
        'cmcProg(4).Left = cmcProg(3).Left + cmcProg(3).Width + 30
        'lacProg(4).Left = cmcProg(4).Left - 30
        igWinStatus(FEEDNAMELIST) = 0
    End If
    If ((tgSpf.sAEDIC = "N") And (tgSpf.sAEDII = "N")) Or (igGGFlag = 0) Then
        igWinStatus(EDISERVICESLIST) = 0
    Else
        If tgUrf(0).sWin(EDISERVICESLIST) = "I" Then
            igWinStatus(EDISERVICESLIST) = 2
        ElseIf tgUrf(0).sWin(EDISERVICESLIST) = "V" Then
            igWinStatus(EDISERVICESLIST) = 1
        Else
            igWinStatus(EDISERVICESLIST) = 0
        End If
    End If
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) <> TAXONAIRTIME) And ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) <> TAXONNTR) Then
        igWinStatus(TAXTABLELIST) = 0
    Else
        If tgUrf(0).sWin(TAXTABLELIST) = "I" Then
            igWinStatus(TAXTABLELIST) = 2
        ElseIf tgUrf(0).sWin(TAXTABLELIST) = "V" Then
            igWinStatus(TAXTABLELIST) = 1
        Else
            igWinStatus(TAXTABLELIST) = 0
        End If
    End If
    If (Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) <> SPLITNETWORKS Then  'Allocation
        igWinStatus(SPLITNETSLIST) = 0
    Else
        If tgUrf(0).sWin(SPLITNETSLIST) = "I" Then
            igWinStatus(SPLITNETSLIST) = 2
        ElseIf tgUrf(0).sWin(SPLITNETSLIST) = "V" Then
            igWinStatus(SPLITNETSLIST) = 1
        Else
            igWinStatus(SPLITNETSLIST) = 0
        End If
    End If
    If (tgSpf.sUseCartNo = "N") Or (igGGFlag = 0) Then
        igWinStatus(MEDIADEFINITIONSLIST) = 0
    Else
        If tgUrf(0).sWin(MEDIADEFINITIONSLIST) = "I" Then
            igWinStatus(MEDIADEFINITIONSLIST) = 2
        ElseIf tgUrf(0).sWin(MEDIADEFINITIONSLIST) = "V" Then
            igWinStatus(MEDIADEFINITIONSLIST) = 1
        Else
            igWinStatus(MEDIADEFINITIONSLIST) = 0
        End If
    End If
    'General
    For ilLoop = VEHICLESLIST To POTENTIALCODESLIST Step 1
        If (igWinStatus(ilLoop) = 0) Or (igGGFlag = 0) Then
            cmcGeneral(ilLoop - VEHICLESLIST).Enabled = False
            lacGeneral(ilLoop - VEHICLESLIST).Enabled = False
            lacGeneral(ilLoop - VEHICLESLIST).ForeColor = vbBlack
            lacGeneral(ilLoop - VEHICLESLIST).FontBold = False
        Else
            cmcGeneral(ilLoop - VEHICLESLIST).Enabled = True
            lacGeneral(ilLoop - VEHICLESLIST).Enabled = True
            If igWinStatus(ilLoop) = 1 Then
                'lacGeneral(ilLoop - VEHICLESLIST).ForeColor = vbBlack     'vbGold; vbYellow
                lacGeneral(ilLoop - VEHICLESLIST).FontBold = False
            Else
                lacGeneral(ilLoop - VEHICLESLIST).ForeColor = vbBlack
            End If
        End If
    Next ilLoop
    If (igWinStatus(FEEDNAMELIST) = 0) Or (igGGFlag = 0) Then
        cmcGeneral(8).Enabled = False
        lacGeneral(8).Enabled = False
        lacGeneral(8).ForeColor = vbBlack
        lacGeneral(8).FontBold = False
    Else
        cmcGeneral(8).Enabled = True
        lacGeneral(8).Enabled = True
        If igWinStatus(FEEDNAMELIST) = 1 Then
            lacGeneral(8).FontBold = False
        Else
            lacGeneral(8).ForeColor = vbBlack
        End If
    End If
    If (igWinStatus(MISSEDREASONSLIST) = 0) Or (igGGFlag = 0) Then
        cmcGeneral(9).Enabled = False
        lacGeneral(9).Enabled = False
        lacGeneral(9).ForeColor = vbBlack
        lacGeneral(9).FontBold = False
    Else
        cmcGeneral(9).Enabled = True
        lacGeneral(9).Enabled = True
        If igWinStatus(MISSEDREASONSLIST) = 1 Then
            lacGeneral(9).FontBold = False
        Else
            lacGeneral(9).ForeColor = vbBlack
        End If
    End If
    'Sales
    For ilLoop = SALESSOURCESLIST To COMPETITORSLIST Step 1
        If ilLoop = RESEARCHLIST Then
            If tgSpf.sGUsePropSys = "Y" Then
                If (igWinStatus(ilLoop) = 0) Or (igGGFlag = 0) Then
                    cmcSales(ilLoop - SALESSOURCESLIST).Enabled = False
                    lacSales(ilLoop - SALESSOURCESLIST).Enabled = False
                    lacSales(ilLoop - SALESSOURCESLIST).ForeColor = vbBlack
                    lacSales(ilLoop - SALESSOURCESLIST).FontBold = False
                Else
                    cmcSales(ilLoop - SALESSOURCESLIST).Enabled = True
                    lacSales(ilLoop - SALESSOURCESLIST).Enabled = True
                    If igWinStatus(ilLoop) = 1 Then
                        'lacSales(ilLoop - SALESSOURCESLIST).ForeColor = vbBlack     'vbGold; vbYellow
                        lacSales(ilLoop - SALESSOURCESLIST).FontBold = False
                    Else
                        lacSales(ilLoop - SALESSOURCESLIST).ForeColor = vbBlack
                    End If
                End If
            Else
                cmcSales(ilLoop - SALESSOURCESLIST).Enabled = False
                lacSales(ilLoop - SALESSOURCESLIST).Enabled = False
                lacSales(ilLoop - SALESSOURCESLIST).ForeColor = vbBlack
                lacSales(ilLoop - SALESSOURCESLIST).FontBold = False
            End If
        Else
            If (igWinStatus(ilLoop) = 0) Or (igGGFlag = 0) Then
                cmcSales(ilLoop - SALESSOURCESLIST).Enabled = False
                lacSales(ilLoop - SALESSOURCESLIST).Enabled = False
                lacSales(ilLoop - SALESSOURCESLIST).ForeColor = vbBlack
                lacSales(ilLoop - SALESSOURCESLIST).FontBold = False
            Else
                cmcSales(ilLoop - SALESSOURCESLIST).Enabled = True
                lacSales(ilLoop - SALESSOURCESLIST).Enabled = True
                If igWinStatus(ilLoop) = 1 Then
                    'lacSales(ilLoop - SALESSOURCESLIST).ForeColor = vbBlack     'vbGold; vbYellow
                    lacSales(ilLoop - SALESSOURCESLIST).FontBold = False
                Else
                    lacSales(ilLoop - SALESSOURCESLIST).ForeColor = vbBlack
                End If
            End If
        End If
    Next ilLoop
    'Accounting
    For ilLoop = ITEMBILLINGTYPESLIST To TRANSACTIONSLIST Step 1
        If (igWinStatus(ilLoop) = 0) Or (igGGFlag = 0) Then
            cmcAccount(ilLoop - ITEMBILLINGTYPESLIST + 1).Enabled = False
            lacAccount(ilLoop - ITEMBILLINGTYPESLIST + 1).Enabled = False
            lacAccount(ilLoop - ITEMBILLINGTYPESLIST + 1).ForeColor = vbBlack
            lacAccount(ilLoop - ITEMBILLINGTYPESLIST + 1).FontBold = False
        Else
            cmcAccount(ilLoop - ITEMBILLINGTYPESLIST + 1).Enabled = True
            lacAccount(ilLoop - ITEMBILLINGTYPESLIST + 1).Enabled = True
            If igWinStatus(ilLoop) = 1 Then
                'lacAccount(ilLoop - ITEMBILLINGTYPESLIST + 1).ForeColor = vbBlack     'vbGold; vbYellow
                lacAccount(ilLoop - ITEMBILLINGTYPESLIST + 1).FontBold = False
            Else
                lacAccount(ilLoop - ITEMBILLINGTYPESLIST + 1).ForeColor = vbBlack
            End If
        End If
    Next ilLoop
    If (igWinStatus(TAXTABLELIST) = 0) Or (igGGFlag = 0) Then
        cmcAccount(0).Enabled = False
        lacAccount(0).Enabled = False
        lacAccount(0).ForeColor = vbBlack
        lacAccount(0).FontBold = False
    Else
        cmcAccount(0).Enabled = True
        lacAccount(0).Enabled = True
        If igWinStatus(TAXTABLELIST) = 1 Then
            lacAccount(0).FontBold = False
        Else
            lacAccount(0).ForeColor = vbBlack
        End If
    End If
    If (igWinStatus(SPLITNETSLIST) = 0) Or (igGGFlag = 0) Then
        cmcSales(10).Enabled = False
        lacSales(10).Enabled = False
        lacSales(10).ForeColor = vbBlack
        lacSales(10).FontBold = False
    Else
        cmcSales(10).Enabled = True
        lacSales(10).Enabled = True
        If igWinStatus(SPLITNETSLIST) = 1 Then
            lacSales(10).FontBold = False
        Else
            lacSales(10).ForeColor = vbBlack
        End If
    End If
    'L.Bianchi
    If (igWinStatus(PODITEMSLIST) = 0) Or (igGGFlag = 0) Then
        cmcSales(11).Enabled = False
        lacSales(11).Enabled = False
        lacSales(11).ForeColor = vbBlack
        lacSales(11).FontBold = False
    Else
        cmcSales(11).Enabled = True
        lacSales(11).Enabled = True
        If igWinStatus(PODITEMSLIST) = 1 Then
            lacSales(11).FontBold = False
        Else
            lacSales(11).ForeColor = vbBlack
        End If
    End If
    'Prog
    For ilLoop = EVENTTYPESLIST To GENRESLIST Step 1
        If (igWinStatus(ilLoop) = 0) Or (igGGFlag = 0) Then
            cmcProg(ilLoop - EVENTTYPESLIST).Enabled = False
            lacProg(ilLoop - EVENTTYPESLIST).Enabled = False
            lacProg(ilLoop - EVENTTYPESLIST).ForeColor = vbBlack
            lacProg(ilLoop - EVENTTYPESLIST).FontBold = False
        Else
            cmcProg(ilLoop - EVENTTYPESLIST).Enabled = True
            lacProg(ilLoop - EVENTTYPESLIST).Enabled = True
            If igWinStatus(ilLoop) = 1 Then
                'lacProg(ilLoop - EVENTTYPESLIST).ForeColor = vbBlack     'vbGold; vbYellow
                lacProg(ilLoop - EVENTTYPESLIST).FontBold = False
            Else
                lacProg(ilLoop - EVENTTYPESLIST).ForeColor = vbBlack
            End If
        End If
    Next ilLoop
    'Copy
    For ilLoop = MEDIADEFINITIONSLIST To ANNOUNCERNAMESLIST Step 1
        If (igWinStatus(ilLoop) = 0) Or (igGGFlag = 0) Then
            cmcCopy(ilLoop - MEDIADEFINITIONSLIST).Enabled = False
            lacCopy(ilLoop - MEDIADEFINITIONSLIST).Enabled = False
            lacCopy(ilLoop - MEDIADEFINITIONSLIST).ForeColor = vbBlack
            lacCopy(ilLoop - MEDIADEFINITIONSLIST).FontBold = False
        Else
            cmcCopy(ilLoop - MEDIADEFINITIONSLIST).Enabled = True
            lacCopy(ilLoop - MEDIADEFINITIONSLIST).Enabled = True
            If igWinStatus(ilLoop) = 1 Then
                'lacCopy(ilLoop - MEDIADEFINITIONSLIST).ForeColor = vbBlack     'vbGold; vbYellow
                lacCopy(ilLoop - MEDIADEFINITIONSLIST).FontBold = False
            Else
                lacCopy(ilLoop - MEDIADEFINITIONSLIST).ForeColor = vbBlack
            End If
        End If
    Next ilLoop
    'Preferences
    For ilLoop = SITELIST To USERLIST Step 1
        If (igWinStatus(ilLoop) = 0) Or (igGGFlag = 0) Then
            cmcPref(ilLoop - SITELIST).Enabled = False
            lacPref(ilLoop - SITELIST).Enabled = False
            lacPref(ilLoop - SITELIST).ForeColor = vbBlack
            lacPref(ilLoop - SITELIST).FontBold = False
        Else
            cmcPref(ilLoop - SITELIST).Enabled = True
            lacPref(ilLoop - SITELIST).Enabled = True
            If igWinStatus(ilLoop) = 1 Then
                'lacPref(ilLoop - SITELIST).ForeColor = vbBlack     'vbGold; vbYellow
                lacPref(ilLoop - SITELIST).FontBold = False
            Else
                lacPref(ilLoop - SITELIST).ForeColor = vbBlack
            End If
        End If
    Next ilLoop
    lacGeneric.ForeColor = vbBlack
    'Job task
    '7/22/11: igWinStatus now set in gSetWinStatusFromSPF for Using Traffic, Using Proposal and Using Affilaite
    'If tgSpf.sUsingTraffic <> "N" Then
        For ilIndex = LBound(tgUrf) To UBound(tgUrf) Step 1
            For ilLoop = RATECARDSJOB To FEEDJOB Step 1
                If tgUrf(ilIndex).sWin(ilLoop) = "I" Then
                    igWinStatus(ilLoop) = 2
                ElseIf tgUrf(ilIndex).sWin(ilLoop) = "V" Then
                    If igWinStatus(ilLoop) = 0 Then
                        igWinStatus(ilLoop) = 1
                    End If
                End If
            Next ilLoop
        Next ilIndex
    'Else
    '    For ilIndex = LBound(tgUrf) To UBound(tgUrf) Step 1
    '        For ilLoop = RATECARDSJOB To FEEDJOB Step 1
    '            If (((ilLoop = RATECARDSJOB) Or (ilLoop = PROPOSALSJOB) Or (ilLoop = BUDGETSJOB) Or (ilLoop > STATIONFEEDJOB)) And (tgSpf.sUsingTraffic = "N")) Or (tgSpf.sUsingTraffic <> "N") Then
    '                If tgUrf(ilIndex).sWin(ilLoop) = "I" Then
    '                    igWinStatus(ilLoop) = 2
    '                ElseIf tgUrf(ilIndex).sWin(ilLoop) = "V" Then
    '                    If igWinStatus(ilLoop) = 0 Then
    '                        igWinStatus(ilLoop) = 1
    '                    End If
    '                End If
    '            Else
    '                igWinStatus(ilLoop) = 0
    '            End If
    '        Next ilLoop
    '    Next ilIndex
    'End If
    If igGGFlag = 0 Then
        'For ilLoop = RATECARDSJOB To STATIONFEEDJOB Step 1
        For ilLoop = RATECARDSJOB To SLSPCOMMSJOB Step 1
            Traffic!cmcTask(ilLoop).Enabled = False
            Traffic!cmcTask(ilLoop).FontBold = False
        Next ilLoop
    Else
        'For ilLoop = RATECARDSJOB To STATIONFEEDJOB Step 1
        For ilLoop = RATECARDSJOB To SLSPCOMMSJOB Step 1
            Traffic!cmcTask(ilLoop).Visible = True
            Traffic!cmcTask(ilLoop).FontBold = True
            Traffic!cmcTask(ilLoop).Enabled = True
        Next ilLoop
        gSetWinStatusFromSPF igWinStatus()
        'If tgSpf.sGUsePropSys = "Y" Then
            If igWinStatus(BUDGETSJOB) = 0 Then
                Traffic!cmcTask(0).Enabled = False
                Traffic!cmcTask(0).FontBold = False
            Else
                If igWinStatus(BUDGETSJOB) = 1 Then
                    Traffic!cmcTask(0).FontBold = False
                Else
                End If
            End If
        'Else
        '    Traffic!cmcTask(0).Enabled = False
        'End If
        If igWinStatus(RATECARDSJOB) = 0 Then
            Traffic!cmcTask(1).Enabled = False
            Traffic!cmcTask(1).FontBold = False
        Else
            If igWinStatus(RATECARDSJOB) = 1 Then
                Traffic!cmcTask(1).FontBold = False
            Else
            End If
        End If
        'If tgSpf.sGUsePropSys = "Y" Then
            If igWinStatus(PROPOSALSJOB) = 0 Then
                Traffic!cmcTask(2).Enabled = False
                Traffic!cmcTask(2).FontBold = False
            Else
                If igWinStatus(PROPOSALSJOB) = 1 Then
                    Traffic!cmcTask(2).FontBold = False
                Else
                End If
            End If
        'Else
        '    Traffic!cmcTask(2).Enabled = False
        'End If
        If igWinStatus(CONTRACTSJOB) = 0 Then
            Traffic!cmcTask(3).Enabled = False
            Traffic!cmcTask(3).FontBold = False
        Else
            If igWinStatus(CONTRACTSJOB) = 1 Then
                Traffic!cmcTask(3).FontBold = False
            Else
            End If
        End If
        If igWinStatus(PROGRAMMINGJOB) = 0 Then
            Traffic!cmcTask(4).Enabled = False
            Traffic!cmcTask(4).FontBold = False
        Else
            If igWinStatus(PROGRAMMINGJOB) = 1 Then
                Traffic!cmcTask(4).FontBold = False
            Else
            End If
        End If
        If igWinStatus(SPOTSJOB) = 0 Then
            Traffic!cmcTask(5).Enabled = False
            Traffic!cmcTask(5).FontBold = False
        Else
            If igWinStatus(SPOTSJOB) = 1 Then
                Traffic!cmcTask(5).FontBold = False
            Else
            End If
        End If
        If igWinStatus(COPYJOB) = 0 Then
            Traffic!cmcTask(6).Enabled = False
        Else
            If igWinStatus(COPYJOB) = 1 Then
                Traffic!cmcTask(6).FontBold = False
                Traffic!cmcTask(6).FontBold = False
            Else
            End If
        End If
        If igWinStatus(LOGSJOB) = 0 Then
            Traffic!cmcTask(7).Enabled = False
            Traffic!cmcTask(7).FontBold = False
        Else
            If igWinStatus(LOGSJOB) = 1 Then
                Traffic!cmcTask(7).FontBold = False
            Else
            End If
        End If
        'Traffic!cmcTask(8).Left = Traffic!cmcTask(7).Left + Traffic!cmcTask(7).Width
        'Traffic!cmcTask(9).Left = Traffic!cmcTask(8).Left + Traffic!cmcTask(8).Width
        'Traffic!cmcTask(10).Left = Traffic!cmcTask(9).Left + Traffic!cmcTask(9).Width
        'Traffic!cmcTask(11).Left = Traffic!cmcTask(10).Left + Traffic!cmcTask(10).Width
        'Traffic!cmcTask(12).Left = Traffic!cmcTask(11).Left + Traffic!cmcTask(11).Width
        'Traffic!cmcReport.Left = Traffic!cmcTask(12).Left + Traffic!cmcTask(12).Width
        Traffic!cmcReport.Width = lgCurrHRes * 15 - Traffic!cmcReport.Left
        If tgSpf.sSystemType <> "R" Then
'            If tgSpf.sGUseAffFeed = "Y" Then
'                If igWinStatus(STATIONFEEDJOB) = 0 Then
'                    Traffic!cmcTask(8).Enabled = False
'                Else
'                    If igWinStatus(STATIONFEEDJOB) = 1 Then
'                        Traffic!cmcTask(8).FontBold = False
'                    Else
'                    End If
'                End If
'            Else
'                Traffic!cmcTask(8).Enabled = False
'            End If
            Traffic!cmcTask(8).Enabled = False
            Traffic!cmcTask(8).Visible = False
            'Traffic!cmcTask(9).Left = Traffic!cmcTask(9).Left - Traffic!cmcTask(8).Width
            'Traffic!cmcTask(10).Left = Traffic!cmcTask(10).Left - Traffic!cmcTask(8).Width
            'Traffic!cmcTask(11).Left = Traffic!cmcTask(11).Left - Traffic!cmcTask(8).Width
            'Traffic!cmcTask(12).Left = Traffic!cmcTask(12).Left - Traffic!cmcTask(8).Width
            'Traffic!cmcReport.Left = Traffic!cmcReport.Left - Traffic!cmcTask(8).Width
            Traffic!cmcReport.Width = Traffic!cmcReport.Width + Traffic!cmcTask(8).Width
        Else
            Traffic!cmcTask(8).Enabled = False
            Traffic!cmcTask(8).Visible = False
            'Traffic!cmcTask(9).Left = Traffic!cmcTask(9).Left - Traffic!cmcTask(8).Width
            'Traffic!cmcTask(10).Left = Traffic!cmcTask(10).Left - Traffic!cmcTask(8).Width
            'Traffic!cmcTask(11).Left = Traffic!cmcTask(11).Left - Traffic!cmcTask(8).Width
            'Traffic!cmcTask(12).Left = Traffic!cmcTask(12).Left - Traffic!cmcTask(8).Width
            'Traffic!cmcReport.Left = Traffic!cmcReport.Left - Traffic!cmcTask(8).Width
            Traffic!cmcReport.Width = Traffic!cmcReport.Width + Traffic!cmcTask(8).Width
        End If
        If igWinStatus(POSTLOGSJOB) = 0 Then
            Traffic!cmcTask(9).Enabled = False
            Traffic!cmcTask(9).FontBold = False
        Else
            If igWinStatus(POSTLOGSJOB) = 1 Then
                Traffic!cmcTask(9).FontBold = False
            Else
            End If
        End If
        If igWinStatus(INVOICESJOB) = 0 Then
            Traffic!cmcTask(10).Enabled = False
            Traffic!cmcTask(10).FontBold = False
        Else
            If igWinStatus(INVOICESJOB) = 1 Then
                Traffic!cmcTask(10).FontBold = False
            Else
            End If
        End If
        If igWinStatus(COLLECTIONSJOB) = 0 Then
            Traffic!cmcTask(11).Enabled = False
            Traffic!cmcTask(11).FontBold = False
        Else
            If igWinStatus(COLLECTIONSJOB) = 1 Then
                Traffic!cmcTask(11).FontBold = False
            Else
            End If
        End If
        If igWinStatus(SLSPCOMMSJOB) = 0 Then
            Traffic!cmcTask(12).Enabled = False
            Traffic!cmcTask(12).FontBold = False
        Else
            If igWinStatus(SLSPCOMMSJOB) = 1 Then
                Traffic!cmcTask(12).FontBold = False
            Else
            End If
        End If
        'If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
            'Traffic!cmcFind.Left = Traffic!cmcReport.Left
            'Traffic!cmcFind.Top = Traffic!cmcReport.Top
            'Traffic!cmcReport.Move Traffic!cmcFind.Left + Traffic!cmcFind.Width + 15, Traffic!cmcReport.Top, Traffic!pbcMsgArea.Width - Traffic!cmcFind.Left - Traffic!cmcFind.Width + 15
            Traffic!cmcFind.Visible = True
        'Else
        '    Traffic!cmcFind.Visible = False
        '4/10/12: Allow superuser access to Generic, only disallow guide from using Generic
        If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
        Else
            If sgUserName = sgSUName Then
                cmcGeneric.Visible = False      'Dan M 71709: guide no longer can access generic reports
                lacGeneric.Enabled = False
                lacGeneric.FontBold = False
            End If
        End If
        'End If
    End If
End Sub

Private Sub pbcClickFocus_GotFocus()
    imcOutline.Visible = False
    mResetColor
End Sub

Private Sub plcWindow_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Visible = False
    mResetColor
    mDisallowToolTip X, Y
End Sub

Private Sub plcWindow_Paint()
    'plcWindow.CurrentX = cmcCSLogo.Left + (3 * cmcCSLogo.Width) / 2
    'plcWindow.CurrentY = cmcCSLogo.Top + cmcCSLogo.Height / 3
    'plcWindow.Print "Counterpoint Software ®"
End Sub

Private Sub mResetColor()
    If imPrevListGroup = -1 Then
        Exit Sub
    End If
    Select Case imPrevListGroup
        Case 0
            'lacAccount(imPrevListIndex).ForeColor = vbBlack
        Case 1
            'lacCopy(imPrevListIndex).ForeColor = vbBlack
        Case 2
            'lacGeneral(imPrevListIndex).ForeColor = vbBlack
        Case 3
            'lacPref(imPrevListIndex).ForeColor = vbBlack
        Case 4
            'lacProg(imPrevListIndex).ForeColor = vbBlack
        Case 5
            'lacSales(imPrevListIndex).ForeColor = vbBlack
    End Select
    imPrevListGroup = -1
End Sub

Private Sub mDisallowToolTip(X As Single, Y As Single)
    Dim Ctrl As Control
    Dim ilLoop As Integer
    
    For Each Ctrl In AuxWnd.Controls
        If Ctrl.Enabled = False Then
            If TypeOf Ctrl Is CommandButton Then
                If (X >= Ctrl.Left) And (X <= Ctrl.Left + Ctrl.Width) Then
                    If (Y >= Ctrl.Top) And (Y <= Ctrl.Top + Ctrl.height) Then
                        plcWindow.ToolTipText = "Disallowed"
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next Ctrl
    plcWindow.ToolTipText = ""

End Sub
Private Sub mCallResearchProject()
    '10066
    Dim slCommandStr
    Dim ilPos As Integer
    Dim slDate As String
    Dim ilRet As Integer
    Dim dlShellRet As Double
    Dim slFormChoice As String
    '10365 we pass test/prod, but we don't use it.
    If igTestSystem Then
        slCommandStr = "Traffic^Test\" & sgUserName & "\"
    Else
        slCommandStr = "Traffic^Prod\" & sgUserName & "\"
    End If
    If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
        ilPos = InStr(1, slCommandStr, "Guide", vbTextCompare)
        If ilPos > 0 Then
            slCommandStr = Left(slCommandStr, ilPos - 1) & "CSI" & Mid(slCommandStr, ilPos + 5)
        End If
    End If
    If (Asc(tgSpf.sMOFCopyAssign) And RSCHCUSTDEMO) = RSCHCUSTDEMO Then
        slFormChoice = "RschByCust"
    Else
        slFormChoice = "Research"
    End If
    slCommandStr = slCommandStr & slFormChoice
    slDate = gCSIGetName()
    If slDate <> "" Then
        slDate = " /D:" & slDate
        slCommandStr = slCommandStr & slDate
    End If
    slCommandStr = slCommandStr & " /ULF:" & lgUlfCode
    On Error GoTo LoadErr
    ilRet = 0
    '10228
   ' AppActivate "Research"
    AppActivate "Research.exe"
    If ilRet = 1 Then
        dlShellRet = Shell(sgExePath & "Research.Exe " & slCommandStr, vbNormalFocus)
    End If
    Exit Sub
LoadErr:
    ilRet = 1
    Resume Next
End Sub

Sub mLoadSiteOpt()
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilShell As Integer
    Dim dlShellRet As Double
    Dim slCommandStr As String
    Dim ilPos As Integer
    Dim slDate As String
    Dim ilLoop As Integer
    
    If igTestSystem Then
        slCommandStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    Else
        slCommandStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    End If
    If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
        ilPos = InStr(1, slCommandStr, "Guide", vbTextCompare)
        If ilPos > 0 Then
            slCommandStr = Left(slCommandStr, ilPos - 1) & "CSI" & Mid(slCommandStr, ilPos + 5)
        End If
    End If
    'Dan M 9/20/10 problems in v57 reports.exe running GetCsiName
    'slDate = Trim$(gGetCSIName("SYSDate"))
    slDate = gCSIGetName()
    If slDate <> "" Then
        'use slDate when writing to file later
        slDate = " /D:" & slDate
        'slCommandStr = slCommandStr & " /D:" & slDate
        slCommandStr = slCommandStr & slDate
    End If
    slCommandStr = slCommandStr & " /ULF:" & lgUlfCode
    slCommandStr = slCommandStr & "/IniLoc:" & CurDir$ & " /UserInput"
    Screen.MousePointer = vbDefault
    Traffic.WindowState = vbMinimized
    gShellAndWait Traffic, sgExePath & "SiteOpt.exe " & slCommandStr, vbNormalFocus, False    'vbFalse
    Traffic.WindowState = vbMaximized
    sgUrfStamp = "~"    'Force read
    'gUrfRead AuxWnd, sgUserName, True, tgUrf(), False  'Obtain user records
    'gUrfRead AuxWnd, sgUserName, False, tgUrf(), False  'Obtain user records
    gInitSuperUser tgUrf(0)
    Screen.MousePointer = vbDefault
End Sub
