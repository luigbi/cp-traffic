VERSION 5.00
Object = "{0E9D0E41-7AB8-11D1-9400-00A0248F2EF0}#1.0#0"; "dzactx.dll"
Object = "{0FB90DC1-97D1-11D1-87C0-444553540000}#1.0#0"; "duzactx.dll"
Begin VB.Form CSISetup 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5775
   ClientLeft      =   2190
   ClientTop       =   2490
   ClientWidth     =   8775
   ClipControls    =   0   'False
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
   Icon            =   "CSISETUP.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5775
   ScaleWidth      =   8775
   Begin VB.DriveListBox cbcDrives 
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
      Height          =   315
      Left            =   1755
      TabIndex        =   73
      Top             =   5130
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.CommandButton cmcBack 
      Appearance      =   0  'Flat
      Caption         =   "< Back"
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
      Height          =   270
      Left            =   1575
      TabIndex        =   1
      Top             =   4635
      Width           =   1170
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
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
      Left            =   5985
      TabIndex        =   3
      Top             =   4635
      Width           =   1170
   End
   Begin VB.CommandButton cmcNext 
      Appearance      =   0  'Flat
      Caption         =   "Next >"
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
      Left            =   3810
      TabIndex        =   2
      Top             =   4635
      Width           =   1170
   End
   Begin VB.PictureBox plcCSISetup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   3960
      Left            =   480
      ScaleHeight     =   3930
      ScaleWidth      =   7755
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   7785
      Begin VB.FileListBox lbcZeroBTR 
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
         Height          =   225
         Left            =   300
         MultiSelect     =   2  'Extended
         Pattern         =   "*.btr"
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   3450
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.PictureBox plcDatabases 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   2640
         Left            =   4650
         ScaleHeight     =   2610
         ScaleWidth      =   7005
         TabIndex        =   81
         Top             =   3645
         Visible         =   0   'False
         Width           =   7035
         Begin VB.ListBox lbcDatabases 
            Height          =   1260
            ItemData        =   "CSISETUP.frx":08CA
            Left            =   1245
            List            =   "CSISETUP.frx":08CC
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   83
            Top             =   570
            Width           =   4710
         End
         Begin VB.Label lacDatabases 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "To Terminate, press Cancel"
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
            Height          =   285
            Index           =   2
            Left            =   0
            TabIndex        =   86
            Top             =   2355
            Visible         =   0   'False
            Width           =   7110
         End
         Begin VB.Label lacDatabases 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "press Next to begin Database conversion(s)"
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
            Height          =   255
            Index           =   1
            Left            =   75
            TabIndex        =   85
            Top             =   1995
            Width           =   6945
         End
         Begin VB.Label lacDatabases 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Databases Requiring Conversion"
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
            Height          =   315
            Index           =   0
            Left            =   15
            TabIndex        =   82
            Top             =   60
            Width           =   6975
         End
      End
      Begin VB.PictureBox pbcLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   540
         Picture         =   "CSISETUP.frx":08CE
         ScaleHeight     =   645
         ScaleWidth      =   735
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   30
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox plcWName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   1665
         Picture         =   "CSISETUP.frx":1805
         ScaleHeight     =   570
         ScaleWidth      =   4515
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   75
         Width           =   4515
      End
      Begin VB.PictureBox plcCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   2430
         Left            =   6540
         ScaleHeight     =   2430
         ScaleWidth      =   6930
         TabIndex        =   54
         Top             =   3690
         Visible         =   0   'False
         Width           =   6930
         Begin VB.Label lacCCorner 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cancel"
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
            Left            =   120
            TabIndex        =   59
            Top             =   60
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label lacCancel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "To terminate the install, press Cancel"
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
            Height          =   285
            Index           =   3
            Left            =   195
            TabIndex        =   58
            Top             =   1950
            Width           =   6465
         End
         Begin VB.Label lacCancel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "To continue with the install, press Next"
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
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   57
            Top             =   1515
            Width           =   6465
         End
         Begin VB.Label lacCancel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "The program has been cancelled before completion of the install"
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
            Height          =   660
            Index           =   1
            Left            =   1020
            TabIndex        =   56
            Top             =   585
            Width           =   4860
         End
         Begin VB.Label lacCancel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Cancelled"
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
            Height          =   315
            Index           =   0
            Left            =   285
            TabIndex        =   55
            Top             =   135
            Visible         =   0   'False
            Width           =   6360
         End
      End
      Begin VB.PictureBox plcBackup 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   2655
         Left            =   420
         ScaleHeight     =   2625
         ScaleWidth      =   6945
         TabIndex        =   62
         Top             =   1035
         Visible         =   0   'False
         Width           =   6975
         Begin VB.Label lacBackup 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "3.  All users off the system?"
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
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   67
            Top             =   1290
            Width           =   6600
         End
         Begin VB.Label lacBackup 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "2  Have you stopped the CSI Server application?"
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
            Height          =   285
            Index           =   6
            Left            =   150
            TabIndex        =   66
            Top             =   1020
            Width           =   6600
         End
         Begin VB.Label lacBackup 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "1.  Have you made a complete backup of the database and Counterpoint application?"
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
            Height          =   390
            Index           =   5
            Left            =   150
            TabIndex        =   65
            Top             =   600
            Width           =   6600
         End
         Begin VB.Label lacBCorner 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Backup"
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
            Height          =   255
            Left            =   90
            TabIndex        =   70
            Top             =   45
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label lacBackup 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Press Cancel, to terminate the Install"
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
            Height          =   285
            Index           =   4
            Left            =   150
            TabIndex        =   69
            Top             =   1935
            Width           =   6600
         End
         Begin VB.Label lacBackup 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Press Next, to continue with install if all three task have been comfirmed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   68
            Top             =   1680
            Width           =   6600
         End
         Begin VB.Label lacBackup 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Setup has detected that the system has been previously installed."
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
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   64
            Top             =   360
            Width           =   6600
         End
         Begin VB.Label lacBackup 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Check List"
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
            Index           =   1
            Left            =   150
            TabIndex        =   63
            Top             =   30
            Width           =   6600
         End
      End
      Begin VB.PictureBox plcFinish 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   2445
         Left            =   5250
         ScaleHeight     =   2415
         ScaleWidth      =   7470
         TabIndex        =   61
         Top             =   3060
         Visible         =   0   'False
         Width           =   7500
         Begin VB.Label lacFinish 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Run DDFReorg"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   4
            Left            =   60
            TabIndex        =   75
            Top             =   1335
            Width           =   7350
         End
         Begin VB.Label lacFCorner 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Finish"
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
            Left            =   105
            TabIndex        =   53
            Top             =   15
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label lacFinish 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "See"
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
            Height          =   285
            Index           =   3
            Left            =   30
            TabIndex        =   51
            Top             =   915
            Width           =   7410
         End
         Begin VB.Label lacFinish 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "To Exit, press Finish"
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
            Height          =   255
            Index           =   2
            Left            =   15
            TabIndex        =   52
            Top             =   1845
            Width           =   7425
         End
         Begin VB.Label lacFinish 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "The program has successfully installed V4.8"
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
            Index           =   1
            Left            =   30
            TabIndex        =   50
            Top             =   525
            Width           =   7395
         End
         Begin VB.Label lacFinish 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Finished"
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
            Height          =   315
            Index           =   0
            Left            =   225
            TabIndex        =   49
            Top             =   105
            Width           =   7050
         End
      End
      Begin VB.PictureBox plcProgress 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   2460
         Left            =   3660
         ScaleHeight     =   2430
         ScaleWidth      =   7005
         TabIndex        =   42
         Top             =   3810
         Visible         =   0   'False
         Width           =   7035
         Begin VB.DirListBox lbcDDFCheckDir 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   2295
            TabIndex        =   80
            Top             =   2160
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.PictureBox pbcProgress 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   375
            Left            =   1155
            ScaleHeight     =   345
            ScaleWidth      =   4635
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1395
            Width           =   4665
            Begin VB.Line lncProgress 
               BorderWidth     =   5
               X1              =   105
               X2              =   4485
               Y1              =   180
               Y2              =   180
            End
         End
         Begin VB.Label lacProgress 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            Index           =   3
            Left            =   1155
            TabIndex        =   72
            Top             =   1830
            Width           =   4680
         End
         Begin VB.Label lacProgreFileListBox 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            Left            =   2400
            TabIndex        =   47
            Top             =   930
            Width           =   4065
         End
         Begin VB.Label lacProgress 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "File: "
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
            Index           =   2
            Left            =   1155
            TabIndex        =   46
            Top             =   930
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label lacProgressFolder 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            Left            =   2385
            TabIndex        =   45
            Top             =   585
            Width           =   4125
         End
         Begin VB.Label lacProgress 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Installing into: "
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
            Index           =   1
            Left            =   1125
            TabIndex        =   44
            Top             =   585
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.Label lacProgress 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Progress Status"
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
            Height          =   315
            Index           =   0
            Left            =   15
            TabIndex        =   43
            Top             =   105
            Width           =   6975
         End
      End
      Begin VB.PictureBox plcSource 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   2625
         Left            =   5670
         ScaleHeight     =   2595
         ScaleWidth      =   6810
         TabIndex        =   26
         Top             =   2820
         Visible         =   0   'False
         Width           =   6840
         Begin VB.DirListBox lbcCheckSource 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   4770
            TabIndex        =   84
            Top             =   1455
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.FileListBox lbcSourceFiles 
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
            Height          =   225
            Left            =   5760
            MultiSelect     =   2  'Extended
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   225
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox edcSource 
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
            Left            =   1080
            TabIndex        =   29
            Top             =   405
            Width           =   5325
         End
         Begin VB.DriveListBox cbcSource 
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
            Height          =   315
            Left            =   1080
            TabIndex        =   31
            Top             =   930
            Width           =   3570
         End
         Begin VB.DirListBox lbcSource 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   1080
            TabIndex        =   33
            Top             =   1455
            Width           =   3540
         End
         Begin VB.Label lacSource 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Path:"
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
            Height          =   315
            Index           =   3
            Left            =   240
            TabIndex        =   28
            Top             =   435
            Width           =   540
         End
         Begin VB.Label lacSource 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Drive:"
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
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   30
            Top             =   930
            Width           =   660
         End
         Begin VB.Label lacSource 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Folder:"
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
            Height          =   315
            Index           =   2
            Left            =   210
            TabIndex        =   32
            Top             =   1455
            Width           =   660
         End
         Begin VB.Label lacSource 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Please choose the Source Folder"
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
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   27
            Top             =   45
            Width           =   6675
         End
      End
      Begin VB.PictureBox plcStart 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   3000
         Left            =   5985
         ScaleHeight     =   2970
         ScaleWidth      =   6780
         TabIndex        =   35
         Top             =   2655
         Visible         =   0   'False
         Width           =   6810
         Begin VB.ListBox lbcStart 
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
            Height          =   1590
            Left            =   750
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   690
            Width           =   5310
         End
         Begin VB.Label lacStart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "If you want to review or change any settings, press Back"
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
            Height          =   240
            Index           =   3
            Left            =   90
            TabIndex        =   40
            Top             =   2715
            Width           =   6645
         End
         Begin VB.Label lacStart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "If you are satisfied with the setting, press Next to begin copying files"
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
            Left            =   90
            TabIndex        =   39
            Top             =   2415
            Width           =   6645
         End
         Begin VB.Label lacSCorner 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Start"
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
            Left            =   120
            TabIndex        =   41
            Top             =   45
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label lacStart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Setup has enough information to start copying the program files"
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
            Left            =   180
            TabIndex        =   37
            Top             =   360
            Width           =   6510
         End
         Begin VB.Label lacStart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Start Copying Files"
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
            Index           =   0
            Left            =   60
            TabIndex        =   36
            Top             =   45
            Width           =   6645
         End
      End
      Begin VB.PictureBox plcBrowse 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   3015
         Left            =   6300
         ScaleHeight     =   2985
         ScaleWidth      =   7470
         TabIndex        =   16
         Top             =   2445
         Visible         =   0   'False
         Width           =   7500
         Begin VB.TextBox edcBrowse 
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
            Left            =   735
            TabIndex        =   19
            Top             =   555
            Width           =   6375
         End
         Begin VB.CommandButton cmcBrowseOk 
            Appearance      =   0  'Flat
            Caption         =   "Ok"
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
            Left            =   1950
            TabIndex        =   24
            Top             =   2580
            Width           =   1170
         End
         Begin VB.CommandButton cmcBrowseCancel 
            Appearance      =   0  'Flat
            Caption         =   "Cancel"
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
            Left            =   4125
            TabIndex        =   25
            Top             =   2580
            Width           =   1170
         End
         Begin VB.DirListBox lbcBrowse 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   4170
            TabIndex        =   23
            Top             =   1095
            Width           =   3210
         End
         Begin VB.DriveListBox cbcBrowse 
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
            Height          =   315
            Left            =   600
            TabIndex        =   21
            Top             =   1095
            Width           =   2340
         End
         Begin VB.Label lacBrowse 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Directories:"
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
            Height          =   315
            Index           =   4
            Left            =   3120
            TabIndex        =   22
            Top             =   1095
            Width           =   945
         End
         Begin VB.Label lacBrowse 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Drive:"
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
            Height          =   315
            Index           =   3
            Left            =   30
            TabIndex        =   20
            Top             =   1095
            Width           =   660
         End
         Begin VB.Label lacBrowse 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Please choose the Destination"
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
            Left            =   240
            TabIndex        =   17
            Top             =   90
            Width           =   6975
         End
         Begin VB.Label lacBrowse 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Path:"
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
            Height          =   315
            Index           =   1
            Left            =   195
            TabIndex        =   18
            Top             =   585
            Width           =   540
         End
      End
      Begin VB.PictureBox plcDestination 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   2115
         Left            =   6645
         ScaleHeight     =   2085
         ScaleWidth      =   7290
         TabIndex        =   9
         Top             =   2370
         Visible         =   0   'False
         Width           =   7320
         Begin VB.TextBox edcDestination 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   255
            TabIndex        =   79
            Top             =   1815
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.DriveListBox cbcDestinationDrive 
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
            Height          =   315
            Left            =   1800
            TabIndex        =   76
            Top             =   1020
            Width           =   3915
         End
         Begin VB.OptionButton rbcProdTest 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Production"
            Height          =   210
            Index           =   0
            Left            =   1110
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   600
            Width           =   180
         End
         Begin VB.OptionButton rbcProdTest 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Test"
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
            Left            =   2490
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   600
            Width           =   180
         End
         Begin VB.Label lacDestination 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "  Test"
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
            Index           =   4
            Left            =   2625
            TabIndex        =   78
            Top             =   600
            Width           =   630
         End
         Begin VB.Label lacDestination 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "  Production"
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
            Index           =   2
            Left            =   1260
            TabIndex        =   77
            Top             =   600
            Width           =   1080
         End
         Begin VB.Label lacBrowse 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Install into"
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
            Left            =   105
            TabIndex        =   13
            Top             =   600
            Width           =   990
         End
         Begin VB.Label lacDestination 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "To install to this Drive, press Next"
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
            Height          =   315
            Index           =   3
            Left            =   105
            TabIndex        =   12
            Top             =   1560
            Width           =   6975
         End
         Begin VB.Label lacDestination 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Destination Drive:"
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
            Height          =   315
            Index           =   1
            Left            =   105
            TabIndex        =   11
            Top             =   1065
            Width           =   1770
         End
         Begin VB.Label lacDestination 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "This setup program will install V4.8 in the following Destination"
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
            Index           =   0
            Left            =   135
            TabIndex        =   10
            Top             =   165
            Width           =   6975
         End
      End
      Begin VB.PictureBox plcWelcome 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   3225
         Left            =   -15
         Picture         =   "CSISETUP.frx":5026
         ScaleHeight     =   3195
         ScaleWidth      =   7755
         TabIndex        =   4
         Top             =   855
         Width           =   7785
         Begin VB.Label lacWMessage 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"CSISETUP.frx":98D5
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   480
            Index           =   5
            Left            =   1455
            TabIndex        =   90
            Top             =   1020
            Width           =   5565
         End
         Begin VB.Label lacWMessage 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "This software should be installed from a Counterpoint client PC."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   4
            Left            =   1455
            TabIndex        =   89
            Top             =   810
            Width           =   5430
         End
         Begin VB.Label lacWMessage 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "NOTE:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   3
            Left            =   705
            TabIndex        =   88
            Top             =   810
            Width           =   735
         End
         Begin VB.Label lacWCorner 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Welcome"
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
            Left            =   90
            TabIndex        =   8
            Top             =   30
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label lacWMessage 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Welcome to the install program for Counterpoint System V4.8"
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
            Height          =   300
            Index           =   0
            Left            =   255
            TabIndex        =   5
            Top             =   390
            Width           =   7185
         End
         Begin VB.Label lacWMessage 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "To Continue, press Next"
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
            Height          =   255
            Index           =   1
            Left            =   375
            TabIndex        =   6
            Top             =   1905
            Width           =   6945
         End
         Begin VB.Label lacWMessage 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "To Terminate, press Cancel"
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
            Height          =   285
            Index           =   2
            Left            =   315
            TabIndex        =   7
            Top             =   2340
            Width           =   7110
         End
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   0
         X2              =   7755
         Y1              =   705
         Y2              =   705
      End
   End
   Begin VB.FileListBox lbcAffReports 
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
      Height          =   225
      Left            =   60
      MultiSelect     =   2  'Extended
      Pattern         =   "*.Exe"
      TabIndex        =   60
      Top             =   5535
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7935
      Top             =   5145
   End
   Begin DUZACTXLibCtl.duzactxctrl zpcDUnZip 
      Left            =   7155
      OleObjectBlob   =   "CSISETUP.frx":995D
      Top             =   5190
   End
   Begin DZACTXLibCtl.dzactxctrl zpcDZip 
      Left            =   8430
      OleObjectBlob   =   "CSISETUP.frx":9A59
      Top             =   5160
   End
End
Attribute VB_Name = "CSISetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy

'**********************************************************
'                Btrieve Match MODULE DEFINITIONS
'
'   Created : 4/25/94       By : D. LeVine
'   Modified :              By :
'
'**********************************************************
Option Explicit
Option Compare Text
Dim hmMsg As Integer   'From file hanle
Dim hmFile As Integer           'file handle
Dim hmTo As Integer
Dim smVersion As String
Dim imScreenNo As Integer
Dim smCurDrive As String
Dim smSystemType As String   'T=Traffic; E=Engineering
Dim imDestByVersion As Integer  'True=Destination is version #; False=Destination is Prod or ProdEngr
Dim imCSIExist As Integer
Dim smDProd As String   'Production Prod folder name
Dim smDCsi As String    'Production CSI folder name
Dim smSProd As String   'Source Prod folder name (Prod_v48)
Dim imSourceExist As Integer
Dim imBackupAsked As Integer
Dim imCancelAns As Integer
Dim smSourceFolders() As String
Dim tmUnZipName() As UNZIPNAME
Dim tmZippedNames() As ZIPPEDNAMES
Dim smKillNames() As String
Dim smPattern As String
Dim smPassword As String
Dim smExport As String
Dim smImport As String
Dim smSave As String
'Module Status Flags
Dim imTerminate As Integer      'True = terminating task, False= OK
Dim imChgMode As Integer        'Change mode status (so change not entered when in change)
Dim imBSMode As Integer         'Backspace flag
Dim smNowDate As String   'Todays date
Dim imFirstTime As Integer
Dim imCancelled As Integer
Dim imProcessing As Integer
Dim smModFilePath As String
Dim smResultFilePath As String
Dim smAffReports As String
Dim smSaveFilePath As String
Dim smTodaysDate As String
Dim imDatabaseFound As Integer   'True = Database found; False = unable to find database
Dim smZeroNames(1 To 2) As String
Dim imDDFReorgOnly As Integer   'Convert databases only
Dim hmDDFFile As Integer
Dim tmDDFFile As DDFFILE
Dim lmActualRecLen As Long
Dim lmBaseRecLen As Long
Dim tmRec As BASEREC
Dim smFieldValue As String






Private Sub cbcBrowse_Change()
    lbcBrowse.Path = cbcBrowse.Drive
    edcBrowse.Text = lbcBrowse.Path
End Sub

Private Sub cbcDestinationDrive_Change()
    edcDestination.Text = left$(cbcDestinationDrive.Drive, 2) & "\"
End Sub

Private Sub cbcSource_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim slVersion As String
    
    smCurDrive = cbcSource.Drive
    If right(smCurDrive, 1) = "\" Then
        slStr = smCurDrive & "csi"
    ElseIf right(smCurDrive, 1) = ":" Then
        slStr = smCurDrive & "\csi"
    ElseIf Len(smCurDrive) > 1 Then
        slStr = smCurDrive & "\csi"
    Else
        slStr = smCurDrive & ":\csi"
    End If
    ilRet = 0
    '7449
On Error GoTo ERRNODRIVE
    lbcSource.Path = slStr
On Error GoTo 0
    If ilRet = 0 Then
        smVersion = ""
        For ilLoop = 0 To lbcSource.ListCount - 1 Step 1
            slVersion = lbcSource.List(ilLoop)
            ilPos = InStr(1, slVersion, "\Prod_", vbTextCompare)
            If ilPos > 0 Then
                smVersion = slVersion
                smVersion = Mid$(smVersion, ilPos + 6)
                ilPos = InStr(1, smVersion, "\", vbTextCompare)
                If ilPos > 0 Then
                    smVersion = left$(smVersion, ilPos - 1)
                End If
                smSystemType = "T"
                Exit For
            End If
        Next ilLoop
        For ilLoop = 0 To lbcSource.ListCount - 1 Step 1
            slVersion = lbcSource.List(ilLoop)
            ilPos = InStr(1, slVersion, "\ProdEngr_", vbTextCompare)
            If ilPos > 0 Then
                smVersion = slVersion
                smVersion = Mid$(smVersion, ilPos + 10)
                ilPos = InStr(1, smVersion, "\", vbTextCompare)
                If ilPos > 0 Then
                    smVersion = left$(smVersion, ilPos - 1)
                End If
                smSystemType = "E"
                Exit For
            End If
        Next ilLoop
        For ilLoop = 0 To lbcSource.ListCount - 1 Step 1
            slVersion = lbcSource.List(ilLoop)
            ilPos = InStr(1, slVersion, "\V", vbTextCompare)
            If ilPos > 0 Then
                slVersion = Mid$(slVersion, ilPos + 1)
                ilPos = InStr(1, slVersion, "\", vbTextCompare)
                If ilPos > 0 Then
                    slVersion = left$(slVersion, ilPos - 1)
                End If
                If Len(slVersion) = 3 Then
                    If (Mid$(slVersion, 2, 1) >= "0") And (Mid$(slVersion, 2, 1) <= "9") And (Mid$(slVersion, 3, 1) >= "0") And (Mid$(slVersion, 3, 1) <= "9") Then
                        smVersion = slVersion
                        smSystemType = "V"
                        Exit For
                    End If
                End If
            End If
        Next ilLoop
        If smVersion <> "" Then
            'If smSystemType = "T" Then
            '    smSProd = "Prod_" & smVersion
            'ElseIf smSystemType = "E" Then
            '    smSProd = "ProdEngr_" & smVersion
            'Else
                smSProd = smVersion
            'End If
            lbcSource.Path = smCurDrive & "\"
            edcSource.Text = lbcSource.Path
            'lbcSourceFiles.Path = lbcSource.Path
            imSourceExist = True
            If Len(smVersion) = 3 Then
                slVersion = left$(smVersion, 2) & "." & right$(smVersion, 1)
            Else
                slVersion = smVersion
            End If
            ilPos = InStr(1, lacFinish(1).Caption, "V4.8", 1)
            If ilPos > 0 Then
                lacFinish(1).Caption = left$(lacFinish(1).Caption, ilPos - 1) & slVersion & Mid$(lacFinish(1).Caption, ilPos + 4)
            End If
            ilPos = InStr(1, lacDestination(0).Caption, "V4.8", 1)
            If ilPos > 0 Then
                lacDestination(0).Caption = left$(lacDestination(0).Caption, ilPos - 1) & slVersion & Mid$(lacDestination(0).Caption, ilPos + 4)
            End If
            ilPos = InStr(1, lacWMessage(0).Caption, "V4.8", 1)
            If ilPos > 0 Then
                lacWMessage(0).Caption = left$(lacWMessage(0).Caption, ilPos - 1) & slVersion & Mid$(lacWMessage(0).Caption, ilPos + 4)
            End If
            lacSource(2).Visible = False
            lbcSource.Visible = False
        Else
            smSProd = "Prod"
            lbcSource.Path = smCurDrive & "\"
            edcSource.Text = lbcSource.Path
            lacSource(2).Visible = True
            lbcSource.Visible = True
        End If
    Else
        smSProd = "Prod"
        lbcSource.Path = smCurDrive
        edcSource.Text = lbcSource.Path
        lacSource(2).Visible = True
        lbcSource.Visible = True
        'lbcSourceFiles.Path = lbcSource.Path
    End If
'    lbcSource.Path = cbcSource.Drive
'    edcSource.Text = lbcSource.Path
    Exit Sub
ERRNODRIVE:
    ilRet = -1
    Resume Next
End Sub
Private Sub cmcBack_Click()
    If imDDFReorgOnly Then
        If imScreenNo = 1 Then
        ElseIf imScreenNo = 7 Then
            plcWelcome.Visible = True
            pbcLogo.Visible = False
            plcBrowse.Visible = False
            cmcBack.Enabled = False
            imScreenNo = 1
        ElseIf imScreenNo = 11 Then
            plcDatabases.Visible = False
            cmcBrowseOk.Visible = False
            cmcBrowseCancel.Visible = False
            plcBrowse.Visible = True
            cmcBack.Enabled = True
            cmcCancel.Enabled = True
            imScreenNo = 7
        ElseIf imScreenNo = 6 Then
            plcFinish.Visible = False
            cmcBrowseOk.Visible = False
            cmcBrowseCancel.Visible = False
            plcBrowse.Visible = True
            cmcBack.Enabled = True
            cmcCancel.Enabled = True
            cmcNext.Caption = "Next >"
            cmcNext.Enabled = True
            imScreenNo = 7
        End If
        Exit Sub
    End If
    If imSourceExist Then
        If imScreenNo = 1 Then
        ElseIf imScreenNo = 2 Then
            plcWelcome.Visible = True
            pbcLogo.Visible = False
            plcDestination.Visible = False
            cmcBack.Enabled = False
            imScreenNo = 1
        ElseIf (imScreenNo = 3) Or (imScreenNo = 10) Then
            plcDestination.Visible = True
            plcStart.Visible = False
            plcBackup.Visible = False
            imScreenNo = 2
        End If
    Else
        If imScreenNo = 1 Then
        ElseIf imScreenNo = 2 Then
            plcWelcome.Visible = True
            pbcLogo.Visible = False
            plcSource.Visible = False
            cmcBack.Enabled = False
            imScreenNo = 1
        ElseIf imScreenNo = 3 Then
            plcSource.Visible = True
            plcDestination.Visible = False
            imScreenNo = 2
        ElseIf (imScreenNo = 4) Or (imScreenNo = 10) Then
            plcDestination.Visible = True
            plcStart.Visible = False
            plcBackup.Visible = False
            imScreenNo = 3
        End If
    End If
End Sub
Private Sub cmcBrowseCancel_Click()
    cmcBack.Enabled = True
    cmcNext.Enabled = True
    cmcCancel.Enabled = True
    plcBrowse.Visible = False
End Sub
Private Sub cmcBrowseOk_Click()
    cmcBack.Enabled = True
    cmcNext.Enabled = True
    cmcCancel.Enabled = True
    plcBrowse.Visible = False
    edcDestination.Text = edcBrowse.Text
End Sub
Private Sub cmcCancel_Click()
    If imProcessing Then
        If imCancelled Then
            imCancelAns = 2
            Exit Sub
        End If
        imCancelled = True
        Exit Sub
    End If
    'Remove zero_btr files retained into temporary folder
    If imScreenNo = 11 Then
        mRemoveTempBTR smResultFilePath & "Updates"
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    'gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcNext_Click()
    Dim ilRet As Integer
    Dim slSource As String
    Dim slDestination As String
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim slVersion As String
    Dim slMsgFile As String
    
    If imCancelled Then
        imCancelAns = 1
        Exit Sub
    End If
    ilRet = 0
    If imDDFReorgOnly Then
        If imScreenNo = 1 Then
            'plcWelcome.Visible = False
            'pbcLogo.Visible = True
            cmcBrowseOk.Visible = False
            cmcBrowseCancel.Visible = False
            plcBrowse.Visible = True
            cmcBack.Enabled = True
            imScreenNo = 7
        ElseIf imScreenNo = 7 Then
            plcBrowse.Visible = False
            smResultFilePath = lbcBrowse.Path
            If right$(smResultFilePath, 1) <> "\" Then
                smResultFilePath = smResultFilePath & "\"
            End If
            mCheckDDF
            If lbcDatabases.ListCount > 0 Then
                plcDatabases.Visible = True
                cmcBack.Enabled = True
                cmcCancel.Enabled = True
                cmcNext.Caption = "Next >"
                cmcNext.Enabled = True
                imScreenNo = 11
                lbcDatabases.SetFocus
                For ilLoop = 0 To lbcDatabases.ListCount - 1 Step 1
                    lbcDatabases.ListIndex = ilLoop
                    SendKeys " ", True
                Next ilLoop
                lbcDatabases.ListIndex = -1
            Else
                lacFinish(1).Caption = "No databases found to be Converted"
                lacFinish(3).Caption = ""
                lacFinish(4).Caption = ""
                plcFinish.Visible = True
                cmcBack.Enabled = True
                cmcCancel.Enabled = False
                cmcNext.Caption = "Finish"
                cmcNext.Enabled = True
                imScreenNo = 6
            End If
        ElseIf imScreenNo = 5 Then
            mTerminate
        ElseIf imScreenNo = 6 Then
            mTerminate
        ElseIf imScreenNo = 11 Then
            'Convert Databases
            plcDatabases.Visible = False
            lacProgress(1).Caption = "Converting:"
            lacProgress(2).Caption = "Running:"
            lacProgress(1).Visible = True
            lacProgress(2).Visible = True
            plcProgress.Visible = True
            If Not mOpenMsgFile(slMsgFile) Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            lacFinish(3).Caption = "Results of MultiDDFReorg stored into " & slMsgFile
            Print #hmMsg, "CSI MultiDDFReorg " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
            Print #hmMsg, ""
            cmcBack.Enabled = False
            cmcNext.Enabled = False
            imScreenNo = 12
            tmcStart.Enabled = True
        ElseIf imScreenNo = 12 Then
            plcProgress.Visible = False
            plcFinish.Visible = True
            cmcBack.Enabled = False
            cmcCancel.Enabled = False
            cmcNext.Caption = "Finish"
            cmcNext.Enabled = True
            imScreenNo = 5
        End If
        Exit Sub
    End If
    If plcDestination.Visible = True Then
        If rbcProdTest(0).Value Or rbcProdTest(1).Value Then
            On Error GoTo mNextPath:
            cbcBrowse.Drive = left$(edcDestination.Text, 1)
            If ilRet = 0 Then
                slDestination = edcDestination.Text
                If right$(slDestination, 1) <> "\" Then
                    slDestination = slDestination & "\"
                End If
                If (rbcProdTest(0).Value) Or (imDestByVersion) Then
                    lbcBrowse.Path = slDestination & smDCsi & smDProd
                Else
                    lbcBrowse.Path = slDestination & smDCsi & "test" & Mid$(smDProd, 5)
                End If
                If ilRet = 0 Then
                    imCSIExist = True
                    If rbcProdTest(1).Value Then
                        smDProd = "test" & Mid$(smDProd, 5)
                    End If
                Else
                    ilRet = 0
                    If smSystemType <> "E" Then
                        If (Not imDestByVersion) Then
                            If rbcProdTest(0).Value Then
                                lbcBrowse.Path = slDestination & smDCsi & "prod\exe"
                            Else
                                lbcBrowse.Path = slDestination & smDCsi & "test\exe"
                            End If
                        Else
                            lbcBrowse.Path = slDestination & smDCsi & smDProd & "exe"
                        End If
                    Else
                        If Not imDestByVersion Then
                            If (rbcProdTest(0).Value) Then
                                lbcBrowse.Path = slDestination & smDCsi & "prodengr\exe"
                            Else
                                lbcBrowse.Path = slDestination & smDCsi & "testengr\exe"
                            End If
                        Else
                            lbcBrowse.Path = slDestination & smDCsi & smDProd & "exe"
                        End If
                    End If
                    If ilRet = 0 Then
                        If smSystemType <> "E" Then
                            If (Not imDestByVersion) Then
                                If rbcProdTest(0).Value Then
                                    smDProd = "prod\"
                                Else
                                    smDProd = "test\"
                                End If
                            End If
                        Else
                            If Not imDestByVersion Then
                                If rbcProdTest(0).Value Then
                                    smDProd = "prodengr\"
                                Else
                                    smDProd = "testengr\"
                                End If
                            End If
                        End If
                        imCSIExist = True
                    Else
                        imCSIExist = False
                    End If
                End If
            Else
                imCSIExist = False
            End If
        Else
            MsgBox "You must specify 'Install into Production or Test'"
            Exit Sub
        End If
    End If
    If imSourceExist Then
        If imScreenNo = 1 Then
            'plcWelcome.Visible = False
            'pbcLogo.Visible = True
            plcDestination.Visible = True
            cmcBack.Enabled = True
            imScreenNo = 2
        ElseIf (imScreenNo = 2) Or (imScreenNo = 10) Then
            If plcSource.Visible = True Then
                plcSource.Visible = False
                plcDestination.Visible = True
                cmcBack.Enabled = True
                imScreenNo = 2
            Else
                plcSource.Visible = False
                plcDestination.Visible = False
                plcBackup.Visible = False
                If imCSIExist And (Not imBackupAsked) Then
                    plcBackup.Visible = True
                    imBackupAsked = True
                    imScreenNo = 10
                Else
                    plcStart.Visible = True
                    lbcStart.Clear
                    lbcStart.AddItem "Source", 0
                    slSource = edcSource.Text
                    If right(slSource, 1) <> "\" Then
                        slSource = slSource & "\"
                    End If
                    slDestination = edcDestination.Text
                    If right$(slDestination, 1) <> "\" Then
                        slDestination = slDestination & "\"
                    End If
                    lbcStart.AddItem "     " & edcSource.Text & "     " & "(from " & slSource & "csi\" & LCase$(smSProd) & ")", 1
                    lbcStart.AddItem "Destination", 2
                    If smSystemType <> "E" Then
                        If Not imCSIExist Then
                            If rbcProdTest(0).Value Then
                                lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & "prod" & ")", 3
                            Else
                                lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & "test" & ")", 3
                            End If
                        ElseIf smDProd <> "" Then
                            If rbcProdTest(0).Value Then
                                lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & LCase$(left$(smDProd, Len(smDProd) - 1)) & ")", 3
                            Else
                                lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & LCase$(left$(smDProd, Len(smDProd) - 1)) & ")", 3
                            End If
                        Else
                            If rbcProdTest(0).Value Then
                                lbcStart.AddItem "     " & edcDestination.Text & "     " & "no \Prod subfolder", 3
                            Else
                                lbcStart.AddItem "     " & edcDestination.Text & "     " & "no \Test subfolder", 3
                            End If
                        End If
                    Else
                        If Not imCSIExist Then
                            If rbcProdTest(0).Value Then
                                lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & "prodengr" & ")", 3
                            Else
                                lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & "testengr" & ")", 3
                            End If
                        ElseIf smDProd <> "" Then
                            If rbcProdTest(0).Value Then
                                lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & LCase$(left$(smDProd, Len(smDProd) - 1)) & ")", 3
                            Else
                                lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & LCase$(left$(smDProd, Len(smDProd) - 1)) & ")", 3
                            End If
                        Else
                            If rbcProdTest(0).Value Then
                                lbcStart.AddItem "     " & edcDestination.Text & "     " & "no \ProdEngr subfolder", 3
                            Else
                                lbcStart.AddItem "     " & edcDestination.Text & "     " & "no \TestEngr subfolder", 3
                            End If
                        End If
                    End If
                    lbcStart.AddItem "Updating Previous Version", 4
                    If imCSIExist Then
                        lbcStart.AddItem "     " & "Yes", 5
                        lbcStart.AddItem "Backup Completed and CSI Server has been stopped", 6
                        lbcStart.AddItem "     " & "Yes", 7
                    Else
                        lbcStart.AddItem "     " & "No", 5
                    End If
                    imScreenNo = 3
                End If
            End If
        ElseIf imScreenNo = 3 Then
            plcStart.Visible = False
            plcProgress.Visible = True
            cmcBack.Enabled = False
            cmcNext.Enabled = False
            imScreenNo = 4
            tmcStart.Enabled = True
        ElseIf imScreenNo = 4 Then
            plcProgress.Visible = False
            If lbcDatabases.ListCount > 0 Then
                plcDatabases.Visible = True
                cmcBack.Enabled = False
                cmcCancel.Enabled = False   'True
                cmcNext.Caption = "Next >"
                cmcNext.Enabled = True
                imScreenNo = 11
                lbcDatabases.SetFocus
                For ilLoop = 0 To lbcDatabases.ListCount - 1 Step 1
                    lbcDatabases.ListIndex = ilLoop
                    SendKeys " ", True
                Next ilLoop
                lbcDatabases.ListIndex = -1
            Else
                plcFinish.Visible = True
                cmcBack.Enabled = False
                cmcCancel.Enabled = False
                cmcNext.Caption = "Finish"
                cmcNext.Enabled = True
                imScreenNo = 5
            End If
        ElseIf imScreenNo = 5 Then
            mTerminate
        ElseIf imScreenNo = 11 Then
            'Convert Databases
            plcDatabases.Visible = False
            lacProgress(1).Caption = "Converting:"
            lacProgress(2).Caption = "Running:"
            plcProgress.Visible = True
            cmcBack.Enabled = False
            cmcNext.Enabled = False
            imScreenNo = 12
            tmcStart.Enabled = True
        ElseIf imScreenNo = 12 Then
            plcProgress.Visible = False
            plcFinish.Visible = True
            cmcBack.Enabled = False
            cmcCancel.Enabled = False
            cmcNext.Caption = "Finish"
            cmcNext.Enabled = True
            imScreenNo = 5
        End If
    Else
        If imScreenNo = 1 Then
            'plcWelcome.Visible = False
            'pbcLogo.Visible = True
            plcSource.Visible = True
            cmcBack.Enabled = True
            imScreenNo = 2
        ElseIf imScreenNo = 2 Then
            smVersion = ""
            slStr = lbcSource.Path
            If right$(slStr, 1) <> "\" Then
                ilRet = 0
                On Error GoTo cmcNextErr:
                lbcCheckSource.Path = slStr & "\csi"
            Else
                lbcCheckSource.Path = slStr & "csi"
            End If
                If ilRet <> 0 Then
                    Exit Sub
                End If
                'For ilLoop = 0 To lbcCheckSource.ListCount - 1 Step 1
                '    slVersion = lbcCheckSource.List(ilLoop)
                '    ilPos = InStr(1, slVersion, "\Prod_", vbTextCompare)
                '    If ilPos > 0 Then
                '        smVersion = slVersion
                '        smVersion = Mid$(smVersion, ilPos + 6)
                '        ilPos = InStr(1, smVersion, "\", vbTextCompare)
                '        If ilPos > 0 Then
                '            smVersion = left$(smVersion, ilPos - 1)
                '        End If
                '        smSystemType = "T"
                '        Exit For
                '    End If
                'Next ilLoop
                'For ilLoop = 0 To lbcCheckSource.ListCount - 1 Step 1
                '    slVersion = lbcCheckSource.List(ilLoop)
                '    ilPos = InStr(1, slVersion, "\ProdEngr_", vbTextCompare)
                '    If ilPos > 0 Then
                '        smVersion = slVersion
                '        smVersion = Mid$(smVersion, ilPos + 10)
                '        ilPos = InStr(1, smVersion, "\", vbTextCompare)
                '        If ilPos > 0 Then
                '            smVersion = left$(smVersion, ilPos - 1)
                '        End If
                '        smSystemType = "E"
                '        Exit For
                '    End If
                'Next ilLoop
                '6998
                For ilLoop = 0 To lbcCheckSource.ListCount - 1 Step 1
                    slVersion = lbcCheckSource.List(ilLoop)
                    '6998 remove .
                    slVersion = Replace(slVersion, ".", "")
                    ilPos = InStr(1, slVersion, "\V", vbTextCompare)
                    If ilPos > 0 Then
                        slVersion = Mid$(slVersion, ilPos + 1)
                        ilPos = InStr(1, slVersion, "\", vbTextCompare)
                        If ilPos > 0 Then
                            slVersion = left$(slVersion, ilPos - 1)
                        End If
                        '6998: just use the first 3 letters
                        slVersion = Mid(slVersion, 1, 3)
                        If Len(slVersion) = 3 Then
                            If (Mid$(slVersion, 2, 1) >= "0") And (Mid$(slVersion, 2, 1) <= "9") And (Mid$(slVersion, 3, 1) >= "0") And (Mid$(slVersion, 3, 1) <= "9") Then
                                smVersion = slVersion
                                'smSystemType = "V"
                                Exit For
                            End If
                        End If
                    End If
                Next ilLoop
'                For ilLoop = 0 To lbcCheckSource.ListCount - 1 Step 1
'                    slVersion = lbcCheckSource.List(ilLoop)
'                    ilPos = InStr(1, slVersion, "\V", vbTextCompare)
'                    If ilPos > 0 Then
'                        slVersion = Mid$(slVersion, ilPos + 1)
'                        ilPos = InStr(1, slVersion, "\", vbTextCompare)
'                        If ilPos > 0 Then
'                            slVersion = left$(slVersion, ilPos - 1)
'                        End If
'                        If Len(slVersion) = 3 Then
'                            If (Mid$(slVersion, 2, 1) >= "0") And (Mid$(slVersion, 2, 1) <= "9") And (Mid$(slVersion, 3, 1) >= "0") And (Mid$(slVersion, 3, 1) <= "9") Then
'                                smVersion = slVersion
'                                'smSystemType = "V"
'                                Exit For
'                            End If
'                        End If
'                    End If
'                Next ilLoop
            'End If
            If smVersion <> "" Then
                'If smSystemType = "T" Then
                '    smSProd = "Prod_" & smVersion
                'ElseIf smSystemType = "E" Then
                '    smSProd = "ProdEngr_" & smVersion
                'Else
                    smSProd = smVersion
                'End If
                imSourceExist = True
                If Len(smVersion) = 3 Then
                    slVersion = left$(smVersion, 2) & "." & right$(smVersion, 1)
                Else
                    slVersion = smVersion
                End If
                ilPos = InStr(1, lacFinish(1).Caption, "V4.8", 1)
                If ilPos > 0 Then
                    lacFinish(1).Caption = left$(lacFinish(1).Caption, ilPos - 1) & slVersion & Mid$(lacFinish(1).Caption, ilPos + 4)
                End If
                ilPos = InStr(1, lacDestination(0).Caption, "V4.8", 1)
                If ilPos > 0 Then
                    lacDestination(0).Caption = left$(lacDestination(0).Caption, ilPos - 1) & slVersion & Mid$(lacDestination(0).Caption, ilPos + 4)
                End If
                ilPos = InStr(1, lacWMessage(0).Caption, "V4.8", 1)
                If ilPos > 0 Then
                    lacWMessage(0).Caption = left$(lacWMessage(0).Caption, ilPos - 1) & slVersion & Mid$(lacWMessage(0).Caption, ilPos + 4)
                End If
                plcSource.Visible = False
                If (imDestByVersion) And (smVersion <> "") Then
                    lacDestination(2).Caption = smVersion
                    rbcProdTest(0).Value = True
                    rbcProdTest(1).Visible = False
                    lacDestination(4).Visible = False
                End If
                plcDestination.Visible = True
                imScreenNo = 2
            Else
                Exit Sub
            End If
        ElseIf (imScreenNo = 3) Or (imScreenNo = 10) Then
            plcDestination.Visible = False
            plcBackup.Visible = False
            If imCSIExist And (Not imBackupAsked) Then
                plcBackup.Visible = True
                imBackupAsked = True
                imScreenNo = 10
            Else
                plcStart.Visible = True
                lbcStart.Clear
                lbcStart.AddItem "Source", 0
                slSource = edcSource.Text
                If right(slSource, 1) <> "\" Then
                    slSource = slSource & "\"
                End If
                slDestination = edcDestination.Text
                If right$(slDestination, 1) <> "\" Then
                    slDestination = slDestination & "\"
                End If
                lbcStart.AddItem "     " & edcSource.Text & "     " & "(from " & slSource & "csi\" & LCase$(smSProd) & ")", 1
                'If rbcProdTest(0).Value Then
                '    lbcStart.AddItem "Destination 'Production' in Folder", 2
                'Else
                '    lbcStart.AddItem "Destination 'Test' in Folder", 2
                'End If
                'If Not imCSIExist Then
                '    lbcStart.AddItem "     " & edcDestination.Text & "     " & "with subfolder \Prod", 3
                'ElseIf smDProd <> "" Then
                '    lbcStart.AddItem "     " & edcDestination.Text & "     " & "with subfolder \" & LCase$(left$(smDProd, Len(smDProd) - 1)), 3
                'Else
                '    lbcStart.AddItem "     " & edcDestination.Text & "     " & "no \Prod subfolder", 3
                'End If
                lbcStart.AddItem "Destination", 2
                If smSystemType <> "E" Then
                    If Not imCSIExist Then
                        If rbcProdTest(0).Value Then
                            lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & "prod" & ")", 3
                        Else
                            lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & "test" & ")", 3
                        End If
                    ElseIf smDProd <> "" Then
                        If rbcProdTest(0).Value Then
                            lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & LCase$(left$(smDProd, Len(smDProd) - 1)) & ")", 3
                        Else
                            lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & "test" & ")", 3
                        End If
                    Else
                        If rbcProdTest(0).Value Then
                            lbcStart.AddItem "     " & edcDestination.Text & "     " & "no \Prod subfolder", 3
                        Else
                            lbcStart.AddItem "     " & edcDestination.Text & "     " & "no \Test subfolder", 3
                        End If
                    End If
                Else
                    If Not imCSIExist Then
                        If rbcProdTest(0).Value Then
                            lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & "prodengr" & ")", 3
                        Else
                            lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & "testengr" & ")", 3
                        End If
                    ElseIf smDProd <> "" Then
                        If rbcProdTest(0).Value Then
                            lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & LCase$(left$(smDProd, Len(smDProd) - 1)) & ")", 3
                        Else
                            lbcStart.AddItem "     " & edcDestination.Text & "     " & "(to " & slDestination & smDCsi & "test" & ")", 3
                        End If
                    Else
                        If rbcProdTest(0).Value Then
                            lbcStart.AddItem "     " & edcDestination.Text & "     " & "no \ProdEngr subfolder", 3
                        Else
                            lbcStart.AddItem "     " & edcDestination.Text & "     " & "no \TestEngr subfolder", 3
                        End If
                    End If
                End If
                lbcStart.AddItem "Updating Previous Version", 4
                If imCSIExist Then
                    lbcStart.AddItem "     " & "Yes", 5
                    lbcStart.AddItem "Backup Completed and CSI Server has been stopped", 6
                    lbcStart.AddItem "     " & "Yes", 7
                Else
                    lbcStart.AddItem "     " & "No", 5
                End If
                imScreenNo = 4
            End If
        ElseIf imScreenNo = 4 Then
            plcStart.Visible = False
            plcProgress.Visible = True
            cmcBack.Enabled = False
            cmcNext.Enabled = False
            imScreenNo = 5
            tmcStart.Enabled = True
        ElseIf imScreenNo = 5 Then
            plcProgress.Visible = False
            If lbcDatabases.ListCount > 0 Then
                plcDatabases.Visible = True
                cmcBack.Enabled = False
                cmcCancel.Enabled = False   'True
                cmcNext.Caption = "Next >"
                cmcNext.Enabled = True
                imScreenNo = 11
                lbcDatabases.SetFocus
                For ilLoop = 0 To lbcDatabases.ListCount - 1 Step 1
                    lbcDatabases.ListIndex = ilLoop
                    SendKeys " ", True
                Next ilLoop
                lbcDatabases.ListIndex = -1
            Else
                plcFinish.Visible = True
                cmcBack.Enabled = False
                cmcCancel.Enabled = False
                cmcNext.Caption = "Finish"
                cmcNext.Enabled = True
                imScreenNo = 6
            End If
        ElseIf imScreenNo = 6 Then
            mTerminate
        ElseIf imScreenNo = 11 Then
            'Convert Databases
            plcDatabases.Visible = False
            lacProgress(1).Caption = "Converting:"
            lacProgress(2).Caption = "Running:"
            plcProgress.Visible = True
            cmcBack.Enabled = False
            cmcNext.Enabled = False
            imScreenNo = 12
            tmcStart.Enabled = True
        ElseIf imScreenNo = 12 Then
            plcProgress.Visible = False
            plcFinish.Visible = True
            cmcBack.Enabled = False
            cmcCancel.Enabled = False
            cmcNext.Caption = "Finish"
            cmcNext.Enabled = True
            imScreenNo = 5
        End If
    End If
    Exit Sub
mNextPath:
    ilRet = 1
    Resume Next
cmcNextErr:
    ilRet = Err
    Resume Next
End Sub

Private Sub Form_Activate()
    plcWelcome.BorderStyle = 0
    plcSource.BorderStyle = 0
    plcDestination.BorderStyle = 0
    plcBackup.BorderStyle = 0
    plcStart.BorderStyle = 0
    plcProgress.BorderStyle = 0
    plcDatabases.BorderStyle = 0
    plcFinish.BorderStyle = 0
    plcWelcome.PaintPicture plcWelcome.Picture, 0, 0
    plcSource.PaintPicture plcWelcome.Image, 0, 0, plcSource.Width, plcSource.Height, plcSource.left, plcSource.top - plcWelcome.top, plcSource.Width, plcSource.Height
    plcDestination.PaintPicture plcWelcome.Image, 0, 0, plcDestination.Width, plcDestination.Height, plcDestination.left, plcDestination.top - plcWelcome.top, plcDestination.Width, plcDestination.Height
    plcBackup.PaintPicture plcWelcome.Image, 0, 0, plcBackup.Width, plcBackup.Height, plcBackup.left, plcBackup.top - plcWelcome.top, plcBackup.Width, plcBackup.Height
    plcStart.PaintPicture plcWelcome.Image, 0, 0, plcStart.Width, plcStart.Height, plcStart.left, plcStart.top - plcWelcome.top, plcStart.Width, plcStart.Height
    plcProgress.PaintPicture plcWelcome.Image, 0, 0, plcProgress.Width, plcProgress.Height, plcProgress.left, plcProgress.top - plcWelcome.top, plcProgress.Width, plcProgress.Height
    plcDatabases.PaintPicture plcWelcome.Image, 0, 0, plcDatabases.Width, plcDatabases.Height, plcDatabases.left, plcDatabases.top - plcWelcome.top, plcDatabases.Width, plcDatabases.Height
    plcFinish.PaintPicture plcWelcome.Image, 0, 0, plcFinish.Width, plcFinish.Height, plcFinish.left, plcFinish.top - plcWelcome.top, plcFinish.Width, plcFinish.Height
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
        End
    Else
        mInit
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    'Reset used instead of Close to cause # Clients on network to be decrement
    'ilRet = btrReset(hgHlf)
    'btrDestroy hgHlf
    End
End Sub

Private Sub lacDestination_Click(Index As Integer)
    If (Index = 2) Then
        rbcProdTest(0).Value = True
    ElseIf Index = 4 Then
        rbcProdTest(1).Value = True
    End If
End Sub

Private Sub lbcBrowse_Change()
    edcBrowse.Text = lbcBrowse.Path
End Sub
Private Sub lbcSource_Change()
    edcSource.Text = lbcSource.Path
    lbcSourceFiles.Path = lbcSource.Path
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyFile                     *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Unzip files                     *
'*                                                     *
'*******************************************************
Private Function mCopyFile(slFromFile As String) As Integer
    Dim slDateTime As String
    Dim ilRet As Integer
    Dim slFDate As String
    Dim slFTime As String
    Dim llFSize As Long
    Dim slTDate As String
    Dim slTTime As String
    Dim llTSize As Long
    Dim ilAddFile As Integer
    Dim ilFileExist As Integer
    Dim ilOvMsg As Integer
    Dim llFTime As Long
    Dim llTTime As Long
    Dim ilSameTime As Integer
    ilRet = 0
    On Error GoTo mCopyFileErr:
    slDateTime = FileDateTime(smModFilePath & slFromFile)
    If ilRet = 0 Then
        slFDate = gAdjYear(Format$(slDateTime, "m/d/yy"))
        slFTime = Format$(slDateTime, "h:mm:ssAM/PM")
        slFTime = Format$(slFTime, "hh:mm:ss")
        llFSize = FileLen(smModFilePath & slFromFile)
        ilAddFile = False
        ilFileExist = False
        ilOvMsg = False
        ilRet = 0
        On Error GoTo mCopyFileErr:
        slDateTime = FileDateTime(smResultFilePath & slFromFile)
        If ilRet = 0 Then
            ilFileExist = True
            slTDate = gAdjYear(Format$(slDateTime, "m/d/yy"))
            slTTime = Format$(slDateTime, "h:mm:ssAM/PM")
            llTSize = FileLen(smResultFilePath & slFromFile)
            slTTime = Format$(slTTime, "hh:mm:ss")
            'ignore seconds as zip does not return the seconds
            llFTime = 3600 * Hour(slFTime) + 60 * Minute(slFTime) '+ Second(slFTime)
            llTTime = 3600 * Hour(slTTime) + 60 * Minute(slTTime) '+ Second(slTTime)
            If DateValue(slFDate) > DateValue(slTDate) + 1 Then
                ilAddFile = True
            ElseIf DateValue(slFDate) < DateValue(slTDate) - 1 Then
                ilOvMsg = True
            Else
                ilSameTime = False
                If (llFTime = llTTime) And (DateValue(slFDate) = DateValue(slTDate)) Then
                    ilSameTime = True
                Else
                    If Minute(slFTime) = Minute(slTTime) Then
                        If Hour(slFTime) = 0 Then
                            If (Hour(slTTime) >= 21) And (DateValue(slFDate) = DateValue(slTDate) + 1) Then
                                ilSameTime = True
                            End If
                            If (Hour(slTTime) <= 3) And (DateValue(slFDate) = DateValue(slTDate)) Then
                                ilSameTime = True
                            End If
                        ElseIf Hour(slFTime) = 1 Then
                            If (Hour(slTTime) >= 22) And (DateValue(slFDate) = DateValue(slTDate) + 1) Then
                                ilSameTime = True
                            End If
                            If (Hour(slTTime) <= 4) And (DateValue(slFDate) = DateValue(slTDate)) Then
                                ilSameTime = True
                            End If
                        ElseIf Hour(slFTime) = 2 Then
                            If (Hour(slTTime) >= 23) And (DateValue(slFDate) = DateValue(slTDate) + 1) Then
                                ilSameTime = True
                            End If
                            If (Hour(slTTime) <= 5) And (DateValue(slFDate) = DateValue(slTDate)) Then
                                ilSameTime = True
                            End If
                        ElseIf Hour(slFTime) = 23 Then
                            If (Hour(slTTime) >= 20) And (DateValue(slFDate) = DateValue(slTDate)) Then
                                ilSameTime = True
                            End If
                            If (Hour(slTTime) <= 2) And (DateValue(slFDate) = DateValue(slTDate) - 1) Then
                                ilSameTime = True
                            End If
                        ElseIf Hour(slFTime) = 22 Then
                            If (Hour(slTTime) >= 19) And (DateValue(slFDate) = DateValue(slTDate)) Then
                                ilSameTime = True
                            End If
                            If (Hour(slTTime) <= 1) And (DateValue(slFDate) = DateValue(slTDate) - 1) Then
                                ilSameTime = True
                            End If
                        ElseIf Hour(slFTime) = 21 Then
                            If (Hour(slTTime) >= 18) And (DateValue(slFDate) = DateValue(slTDate)) Then
                                ilSameTime = True
                            End If
                            If (Hour(slTTime) <= 0) And (DateValue(slFDate) = DateValue(slTDate) - 1) Then
                                ilSameTime = True
                            End If
                        Else
                            If (Hour(slTTime) >= Hour(slFTime) - 3) And (Hour(slTTime) <= Hour(slFTime) + 3) And (DateValue(slFDate) = DateValue(slTDate)) Then
                                ilSameTime = True
                            End If
                        End If
                    End If
                End If
                If ilSameTime And (llFSize = llTSize) Then
                    ilAddFile = True
                Else
                    If DateValue(slFDate) > DateValue(slTDate) And (Not ilSameTime) Then
                        ilAddFile = True
                    Else
                        If (DateValue(slFDate) = DateValue(slTDate)) And (llFTime > llTTime) And (Not ilSameTime) Then
                            ilAddFile = True
                        Else
                            ilOvMsg = True
                        End If
                    End If
                End If
            End If
        Else
            'ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
            ilAddFile = True
        End If
        'If ilOvMsg Then
        '    ilRet = MsgBox("Overwrite: " & Trim$(tmUnZipName(ilUpper).sName) & Chr(10) & Chr(13) & "     " & slDate & " " & slTime & Str$(llSize) & " with " & Chr(10) & Chr(13) & "     " & Format$(tmUnZipName(ilUpper).sDate, "m/d/yy") & " " & Format$(tmUnZipName(ilUpper).sTime, "h:mm:ssAM/PM") & Str$(tmUnZipName(ilUpper).lSize), 4, "Overwrite")
        '    If ilRet = 6 Then
        '        'ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
        '        ilAddFile = True
        '    End If
        'End If
        If ilAddFile Then
            If ilFileExist Then
                ilRet = GetAttr(smResultFilePath & slFromFile) And 1
                If ilRet <> 0 Then
                    SetAttr smResultFilePath & slFromFile, 0
                End If
            End If
            ilRet = 0
            On Error GoTo mCopyFileErr:
            FileCopy smModFilePath & slFromFile, smResultFilePath & slFromFile
            If ilRet = 0 Then
                If Not ilFileExist Then
                    Print #hmMsg, "     " & Trim$(slFromFile) & " " & Trim$(slFDate) & " " & Trim$(slFTime) & " Size" & Str$(llFSize) & " Copied into " & smResultFilePath & slFromFile
                Else
                    Print #hmMsg, "     " & Trim$(slFromFile) & " " & Trim$(slFDate) & " " & Trim$(slFTime) & " Size" & Str$(llFSize) & " Copied into " & smResultFilePath & slFromFile & " Replaced: " & Trim$(slTDate) & " " & Trim$(slTTime) & " Size" & Str$(llTSize)
                End If
                SetAttr smResultFilePath & slFromFile, 0
            Else
                Print #hmMsg, "Copying " & Trim$(smModFilePath & slFromFile) & " To " & Trim$(smResultFilePath & slFromFile) & " Error #" & Str$(ilRet)
            End If
        End If
    Else
        Print #hmMsg, "Unable to find " & Trim$(smModFilePath & slFromFile) & " to Copy To " & Trim$(smResultFilePath & slFromFile) & " Error #" & Str$(ilRet)
    End If
                
    Exit Function
mCopyFileErr:
    ilRet = Err
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:4/17/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
    
    Dim ilRet As Integer   'Return from btrieve calls
    Dim slDate As String
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilPass As Integer
    Dim ilPos As Integer
    Dim slVersion As String
    Dim slTest As String
    Dim slStr As String
    Dim ilCheckDest As Integer
    
    'ttp 5603
    frmSetupPassword.Show vbModal
    If Not bgSetupPassword Then
        imTerminate = True
        mTerminate
        Exit Sub
    End If
    imTerminate = False
    imCancelled = False
    imProcessing = False
    Screen.MousePointer = vbHourglass
    mParseCmmdLine
    plcCSISetup.Move 495, 510
    plcWelcome.Move 0, 735
    plcDestination.Move 225, 990
    plcBackup.Move 480, 1080
    imBackupAsked = False
    plcBrowse.Move 135, 825
    plcStart.Move 525, 825
    plcSource.Move 135, 1035
    plcProgress.Move 390, 960
    plcDatabases.Move 345, 990
    plcCancel.Move 315, 960
    plcFinish.Move 120, 960
    lncProgress.Visible = False
    '6776
    lacWMessage(3).Caption = ""
    lacWMessage(4).Caption = ""
    lacWMessage(5).Caption = ""
    
    imScreenNo = 1
    smCurDrive = CurDir$
    imSourceExist = False
    smSystemType = "T"
    ilRet = 0
    On Error GoTo mInitPath:
    cbcSource.Drive = left$(smCurDrive, 1)
    If ilRet = 0 Then
        On Error GoTo mInitPath:
        'lbcSource.Path = left$(cbcSource.List(ilLoop), 1) & ":\csi\prod"
        If right(smCurDrive, 1) = "\" Then
            slStr = smCurDrive & "csi"
        ElseIf right(smCurDrive, 1) = ":" Then
            slStr = smCurDrive & "\csi"
        ElseIf Len(smCurDrive) > 1 Then
            slStr = smCurDrive & "\csi"
        Else
            slStr = smCurDrive & ":\csi"
        End If
        ilRet = 0
        lbcSource.Path = slStr
        If ilRet = 0 Then
            smVersion = ""
            For ilLoop = 0 To lbcSource.ListCount - 1 Step 1
                slVersion = lbcSource.List(ilLoop)
                'ilPos = InStr(1, slVersion, "\Prod_", vbTextCompare)
                'If ilPos > 0 Then
                '    smVersion = Mid$(slVersion, ilPos + 6)
                '    ilPos = InStr(1, smVersion, "\", vbTextCompare)
                '    If ilPos > 0 Then
                '        smVersion = left$(smVersion, ilPos - 1)
                '    End If
                '    Exit For
                'End If
                'ilPos = InStr(1, slVersion, "\ProdEngr_", vbTextCompare)
                'If ilPos > 0 Then
                '    smVersion = Mid$(slVersion, ilPos + 10)
                '    ilPos = InStr(1, smVersion, "\", vbTextCompare)
                '    If ilPos > 0 Then
                '        smVersion = left$(smVersion, ilPos - 1)
                '    End If
                '    smSystemType = "E"
                '    Exit For
                'End If
                ilPos = InStr(1, slVersion, "\V", vbTextCompare)
                If ilPos > 0 Then
                    slVersion = Mid$(slVersion, ilPos + 1)
                    ilPos = InStr(1, slVersion, "\", vbTextCompare)
                    If ilPos > 0 Then
                        slVersion = left$(slVersion, ilPos - 1)
                    End If
                    If Len(slVersion) = 3 Then
                        If (Mid$(slVersion, 2, 1) >= "0") And (Mid$(slVersion, 2, 1) <= "9") And (Mid$(slVersion, 3, 1) >= "0") And (Mid$(slVersion, 3, 1) <= "9") Then
                            smVersion = slVersion
                            slVersion = Mid(smVersion, 2)
                            If Val(slVersion) > 50 Then
                                smSystemType = "T"
                            Else
                                smSystemType = "E"
                            End If
                            Exit For
                        End If
                    End If
                End If
            Next ilLoop
            If smVersion <> "" Then
                'If smSystemType = "T" Then
                '    smSProd = "Prod_" & smVersion
                'ElseIf smSystemType = "E" Then
                '    smSProd = "ProdEngr_" & smVersion
                'Else
                    smSProd = smVersion
                'End If
                ilPos = InStr(1, smCurDrive, "csi", vbTextCompare)
                If ilPos > 0 Then
                    lbcSource.Path = left$(smCurDrive, ilPos)
                Else
                    If right$(smCurDrive, 1) = "\" Then
                        lbcSource.Path = smCurDrive
                    Else
                        lbcSource.Path = smCurDrive & "\"
                    End If
                End If
                edcSource.Text = lbcSource.Path
                'lbcSourceFiles.Path = lbcSource.Path
                imSourceExist = True
            Else
                smSProd = "Prod"
                lbcSource.Path = smCurDrive
                edcSource.Text = lbcSource.Path
            End If
        Else
            smSProd = "Prod"
            lbcSource.Path = smCurDrive
            edcSource.Text = lbcSource.Path
            'lbcSourceFiles.Path = lbcSource.Path
        End If
    End If
    
    If smVersion <> "" Then
        If Len(smVersion) = 3 Then
            slVersion = left$(smVersion, 2) & "." & right$(smVersion, 1)
        Else
            slVersion = smVersion
        End If
        ilPos = InStr(1, lacFinish(1).Caption, "V4.8", 1)
        If ilPos > 0 Then
            lacFinish(1).Caption = left$(lacFinish(1).Caption, ilPos - 1) & slVersion & Mid$(lacFinish(1).Caption, ilPos + 4)
        End If
        ilPos = InStr(1, lacDestination(0).Caption, "V4.8", 1)
        If ilPos > 0 Then
            lacDestination(0).Caption = left$(lacDestination(0).Caption, ilPos - 1) & slVersion & Mid$(lacDestination(0).Caption, ilPos + 4)
        End If
        ilPos = InStr(1, lacWMessage(0).Caption, "V4.8", 1)
        If ilPos > 0 Then
            lacWMessage(0).Caption = left$(lacWMessage(0).Caption, ilPos - 1) & slVersion & Mid$(lacWMessage(0).Caption, ilPos + 4)
        End If
        ilPos = InStr(1, smVersion, ".", 1)
        If ilPos > 0 Then
            slVersion = left$(smVersion, ilPos - 1) & Mid$(smVersion, ilPos + 1)
        Else
            slVersion = smVersion
        End If
'        If smSystemType = "E" Then
'            ilPos = InStr(1, lacFinish(1).Caption, "Traffic", 1)
'            If ilPos > 0 Then
'                lacFinish(1).Caption = left$(lacFinish(1).Caption, ilPos - 1) & "Engineering" & Mid$(lacFinish(1).Caption, ilPos + 7)
'            End If
'            ilPos = InStr(1, lacDestination(0).Caption, "Traffic", 1)
'            If ilPos > 0 Then
'                lacDestination(0).Caption = left$(lacDestination(0).Caption, ilPos - 1) & "Engineering" & Mid$(lacDestination(0).Caption, ilPos + 7)
'            End If
'            ilPos = InStr(1, lacWMessage(0).Caption, "Traffic", 1)
'            If ilPos > 0 Then
'                lacWMessage(0).Caption = left$(lacWMessage(0).Caption, ilPos - 1) & "Engineering" & Mid$(lacWMessage(0).Caption, ilPos + 7)
'            End If
'        End If
    Else
        ilPos = InStr(1, lacWMessage(0).Caption, "V4.8", 1)
        If ilPos > 0 Then
            lacWMessage(0).Caption = left$(lacWMessage(0).Caption, ilPos - 1) & smVersion & Mid$(lacWMessage(0).Caption, ilPos + 4)
        End If
        slVersion = "v48"
    End If

    imCSIExist = False
    imDestByVersion = False
    For ilPass = 0 To 4 Step 1
        For ilLoop = 0 To cbcDrives.ListCount - 1 Step 1
            ilRet = 0
            On Error GoTo mInitPath:
            ilCheckDest = False
            If imDDFReorgOnly Then
                If (UCase$(left$(smCurDrive, 1)) = UCase$(left$(cbcDrives.List(ilLoop), 1))) Or (UCase$(left$(edcSource.Text, 1)) = UCase$(left$(cbcDrives.List(ilLoop), 1))) Then
                    ilCheckDest = True
                End If
            Else
                If (UCase$(left$(smCurDrive, 1)) <> UCase$(left$(cbcDrives.List(ilLoop), 1))) And (UCase$(left$(edcSource.Text, 1)) <> UCase$(left$(cbcDrives.List(ilLoop), 1))) Then
                    ilCheckDest = True
                End If
            End If
            If ilCheckDest Then
                cbcBrowse.Drive = left$(cbcDrives.List(ilLoop), 1)
                If ilRet = 0 Then
                    On Error GoTo mInitPath:
                    If ilPass = 0 Then
                        If smSystemType <> "E" Then
                            smDProd = "prod_" & slVersion
                        Else
                            smDProd = "prodengr_" & slVersion
                        End If
                        smDCsi = "csi"
                    ElseIf ilPass = 1 Then
                        If smSystemType <> "E" Then
                            smDProd = "prod"
                        Else
                            smDProd = "prodengr"
                        End If
                        smDCsi = "csi"
                    ElseIf ilPass = 2 Then
                        If smSystemType <> "E" Then
                            smDProd = "prod_"
                        Else
                            smDProd = "prodengr_"
                        End If
                        smDCsi = ""
                    ElseIf ilPass = 3 Then
                        If smSystemType <> "E" Then
                            smDProd = "prod"
                        Else
                            smDProd = "prodengr"
                        End If
                        smDCsi = ""
                    ElseIf ilPass = 4 Then
                        smDProd = smVersion
                        smDCsi = "csi"
                    End If
                    smDProd = smDProd & "\"
                    If smDCsi <> "" Then
                        smDCsi = smDCsi & "\"
                    End If
                    lbcBrowse.Path = left$(cbcBrowse.List(ilLoop), 1) & ":\" & smDCsi & left$(smDProd, Len(smDProd) - 1)
                    If ilRet = 0 Then
                        If smDCsi <> "" Then
                            lbcBrowse.Path = left$(cbcBrowse.List(ilLoop), 1) & ":\" & left$(smDCsi, Len(smDCsi) - 1)
                            edcDestination.Text = left$(lbcBrowse.Path, 3)
                        Else
                            lbcBrowse.Path = left$(cbcBrowse.List(ilLoop), 1) & ":"
                            edcDestination.Text = left$(cbcBrowse.List(ilLoop), 1) & ":\"
                        End If
                        imCSIExist = True
                        If ilPass = 4 Then
                            imDestByVersion = True
                        End If
                        Exit For
                    End If
                End If
            End If
        Next ilLoop
        If imCSIExist Then
            Exit For
        End If
    Next ilPass
    If Not imCSIExist Then
        smDCsi = "csi\"
        If smSystemType <> "E" Then
            smDProd = "prod\"
        Else
            smDProd = "prodengr\"
        End If
        edcDestination.Text = "c:\"
        cbcBrowse.Drive = "c"
    End If
    cbcDestinationDrive.Drive = left$(edcDestination.Text, 1)
    CSISetup.Height = cmcCancel.top + 5 * cmcCancel.Height / 3
    'gCenterStdAlone BtrMatch
    CSISetup.Move (Screen.Width - CSISetup.Width) \ 2, (Screen.Height - CSISetup.Height) \ 2 + 175 '+ Screen.Height \ 10
    CSISetup.Show
    
    If (imDestByVersion) And (smVersion <> "") Then
        lacDestination(2).Caption = smVersion
        rbcProdTest(0).Value = True
        rbcProdTest(1).Visible = False
        lacDestination(4).Visible = False
    End If
    imChgMode = False
    imBSMode = False
    imFirstTime = True
    smNowDate = Format$(Now, "m/d/yy")
    If smSystemType <> "E" Then
        smZeroNames(1) = "rnf.btr"
        smZeroNames(2) = "rtf.btr"
    End If
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    mTerminate
    Exit Sub
mInitPath:
    ilRet = 1
    Resume Next
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInstall                        *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create folders, Copy Images,   *
'*                      unzip images                   *
'*                                                     *
'*******************************************************
Private Function mInstall() As Integer
    Dim slSource As String
    Dim slDestination As String
    Dim ilRet As Integer
    Dim ilSourceCount As Integer
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilBase As Integer
    Dim ilPercent As Integer
    Dim slDrivePath As String
    Dim slMsgFile As String
    Dim ilError As Integer
    Dim slName As String
    Dim slToName As String
    Dim slDateTime As String
    Dim llPercent As Long
    Dim ilCount As Integer
    ReDim ilPos(0 To 1) As Integer
    
    imProcessing = True
    Screen.MousePointer = vbHourglass
    lacFinish(4).Caption = ""
    lacProgress(1).Visible = False
    lacProgress(2).Visible = False
    lbcDatabases.Clear
    lncProgress.X2 = lncProgress.X1
    DoEvents
    'Determine number of modules to be copied and unzipped
    slSource = edcSource.Text
    If right(slSource, 1) = "\" Then
        slSource = slSource & "csi" '\prod"
    ElseIf right(slSource, 1) = ":" Then
        slSource = slSource & "\csi"    '\prod"
    ElseIf Len(slSource) > 1 Then
        slSource = slSource & "\csi"    '\prod"
    Else
        slSource = slSource & ":\csi"   '\prod"
    End If
    If right$(slSource, 1) <> "\" Then
        slSource = slSource & "\"
    End If
    If Not imCSIExist Then
        If rbcProdTest(0).Value Then
            If smSystemType <> "E" Then
                smDProd = "prod\"
            Else
                smDProd = "prodengr\"
            End If
        Else
            If smSystemType <> "E" Then
                smDProd = "test\"
            Else
                smDProd = "testengr\"
            End If
        End If
    End If
'    If rbcProdTest(1).Value Then
'        smDProd = "Test\"
'    End If
    slDestination = edcDestination.Text
    If right$(slDestination, 1) <> "\" Then
        slDestination = slDestination & "\"
    End If
    slDestination = slDestination & smDCsi
    ilRet = 0
    On Error GoTo mMakeFolderErr:
    lbcSource.Path = slSource & smSProd  'left$(slSource, Len(slSource) - 5)
    If lbcSource.ListCount <= 0 Then
        lacFinish(3).Caption = ""
        imProcessing = False
        Screen.MousePointer = vbDefault
        MsgBox "Unable to find " & slSource & smSProd & " Error # " & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Create Error"
        mInstall = False
        Exit Function
    End If
    lbcSource.Path = left$(slSource, Len(slSource) - 1)
    ReDim smsourcefolder(0 To lbcSource.ListCount) As String
    
    For ilLoop = 0 To lbcSource.ListCount - 1 Step 1
        smsourcefolder(ilLoop) = lbcSource.List(ilLoop)
    Next ilLoop
    ilLoop = 0
    Do
        lbcSource.Path = smsourcefolder(ilLoop)
        If lbcSource.ListCount > 0 Then
            ilBase = UBound(smsourcefolder)
            ReDim Preserve smsourcefolder(0 To ilBase + lbcSource.ListCount) As String
            For ilIndex = 0 To lbcSource.ListCount - 1 Step 1
                smsourcefolder(ilIndex + ilBase) = lbcSource.List(ilIndex)
            Next ilIndex
        End If
        ilLoop = ilLoop + 1
    Loop While ilLoop < UBound(smsourcefolder)
    
    For ilLoop = 0 To UBound(smsourcefolder) - 1 Step 1
        lbcSource.Path = smsourcefolder(ilLoop)
        lbcSourceFiles.Path = lbcSource.Path
        ilSourceCount = ilSourceCount + lbcSourceFiles.ListCount
        'If lbcSource.ListCount > 0 Then
        '    ilBase = UBound(smSourceFolder)
        '    ReDim Preserve smSourceFolder(0 To ilBase + lbcSource.ListCount) As String
        '    For ilIndex = 0 To lbcSource.ListCount - 1 Step 1
        '        smSourceFolder(ilIndex + ilBase) = lbcSource.List(ilIndex)
        '    Next ilIndex
        'End If
    Next ilLoop
    'Determine if base path previously created. If not, then create
    slDrivePath = left$(slDestination, Len(slDestination) - 1)
    ilPos(0) = InStr(1, slDrivePath, "\", 1)
    If ilPos(0) > 0 Then
        Do
            ilPos(1) = InStr(ilPos(0) + 1, slDrivePath, "\", 1)
            If ilPos(1) > 0 Then
                On Error GoTo mMakeFolderErr:
                ilRet = 0
                MkDir left$(slDrivePath, ilPos(1) - 1)
                If (ilRet <> 0) And (ilRet <> 75) Then
                    imProcessing = False
                    Screen.MousePointer = vbDefault
                    MsgBox "Unable to create " & slDrivePath & " Error # " & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Create Error"
                    smResultFilePath = slDestination & smDProd
                    mInstall = False
                    Exit Function
                End If
                ilPos(0) = ilPos(1)
            Else
                If Len(slDrivePath) > ilPos(0) + 1 Then
                    On Error GoTo mMakeFolderErr:
                    ilRet = 0
                    MkDir slDrivePath
                    If (ilRet <> 0) And (ilRet <> 75) Then
                        imProcessing = False
                        Screen.MousePointer = vbDefault
                        MsgBox "Unable to create " & slDrivePath & " Error # " & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Create Error"
                        smResultFilePath = slDestination & smDProd
                        mInstall = False
                        Exit Function
                    End If
                End If
                Exit Do
            End If
        Loop While ilPos(1) > 0
    End If
    smResultFilePath = slDestination
    If Not mOpenMsgFile(slMsgFile) Then
        imProcessing = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    Print #hmMsg, "CSI Install " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmMsg, ""
    lacFinish(3).Caption = "Results of Install stored into " & slMsgFile
    'Test if traffic/Affiliat are being used, copy images to safe folder (Updates), then try to copy it back into
    'exe folder.  If unable to copy it back in, then it is in use.  First copying it to Updates saves the image
    'in case it is not part of the release.
    If imCSIExist Then
        If smSystemType <> "E" Then
            lacProgress(3).Caption = "Checking if Traffic System in Use"
        Else
            lacProgress(3).Caption = "Checking if Engineering System in Use"
        End If
        DoEvents
        ilRet = 0
        On Error GoTo mMakeFolderErr:
        If smSystemType <> "E" Then
            slDateTime = FileDateTime(smResultFilePath & smDProd & "Exe\Traffic.exe")
        Else
            slDateTime = FileDateTime(smResultFilePath & smDProd & "Exe\Engineer.exe")
        End If
        If ilRet = 0 Then
            ilRet = 0
            On Error GoTo mMakeFolderErr:
            If smSystemType <> "E" Then
                FileCopy smResultFilePath & smDProd & "Exe\Traffic.exe", smResultFilePath & smDProd & "Updates\Traffic.Exe"
            Else
                FileCopy smResultFilePath & smDProd & "Exe\Engineer.exe", smResultFilePath & smDProd & "Updates\Engineer.Exe"
            End If
            If ilRet <> 0 Then
                imProcessing = False
                Screen.MousePointer = vbDefault
                If smSystemType <> "E" Then
                    Print #hmMsg, "Unable to Copy " & smResultFilePath & smDProd & "Exe\Traffic.exe" & " To " & smResultFilePath & smDProd & "Updates\Traffic.Exe" & " Error #" & Str$(ilRet)
                Else
                    Print #hmMsg, "Unable to Copy " & smResultFilePath & smDProd & "Exe\Engineer.exe" & " To " & smResultFilePath & smDProd & "Updates\Engineer.Exe" & " Error #" & Str$(ilRet)
                End If
                MsgBox "Unable to copy traffic.exe to updates folder.  See " & slMsgFile & " for Error messages", vbOKOnly + vbCritical + vbApplicationModal, "Create Error"
                smResultFilePath = slDestination & smDProd
                Close #hmMsg
                mInstall = False
                Exit Function
            End If
            ilRet = 0
            If smSystemType <> "E" Then
                FileCopy smResultFilePath & smDProd & "Updates\Traffic.exe", smResultFilePath & smDProd & "Exe\Traffic.Exe"
            Else
                FileCopy smResultFilePath & smDProd & "Updates\Engineer.exe", smResultFilePath & smDProd & "Exe\Engineer.Exe"
            End If
            If ilRet <> 0 Then
                On Error Resume Next
                If smSystemType <> "E" Then
                    Kill smResultFilePath & smDProd & "Updates\Traffic.exe"
                Else
                    Kill smResultFilePath & smDProd & "Updates\Engineer.exe"
                End If
                imProcessing = False
                Screen.MousePointer = vbDefault
                If smSystemType <> "E" Then
                    Print #hmMsg, "Traffic in Use, Unable to Copy " & smResultFilePath & smDProd & "Updates\Traffic.exe" & " To " & smResultFilePath & smDProd & "Exe\Traffic.Exe" & " Error #" & Str$(ilRet)
                    MsgBox "Traffic in use.  See " & slMsgFile & " for Error messages", vbOKOnly + vbCritical + vbApplicationModal, "Create Error"
                Else
                    Print #hmMsg, "Engineering in Use, Unable to Copy " & smResultFilePath & smDProd & "Updates\Engineer.exe" & " To " & smResultFilePath & smDProd & "Exe\Engineer.Exe" & " Error #" & Str$(ilRet)
                    MsgBox "Engineering in use.  See " & slMsgFile & " for Error messages", vbOKOnly + vbCritical + vbApplicationModal, "Create Error"
                End If
                smResultFilePath = slDestination & smDProd
                Close #hmMsg
                mInstall = False
                Exit Function
            End If
            On Error Resume Next
            If smSystemType <> "E" Then
                Kill smResultFilePath & smDProd & "Updates\Traffic.exe"
            Else
                Kill smResultFilePath & smDProd & "Updates\Engineer.exe"
            End If
        End If
        DoEvents
        If smSystemType <> "E" Then
            ilRet = 0
            On Error GoTo mMakeFolderErr:
            slDateTime = FileDateTime(smResultFilePath & smDProd & "Exe\Traffic2.exe")
            If ilRet = 0 Then
                ilRet = 0
                On Error GoTo mMakeFolderErr:
                FileCopy smResultFilePath & smDProd & "Exe\Traffic2.exe", smResultFilePath & smDProd & "Updates\Traffic2.Exe"
                If ilRet <> 0 Then
                    imProcessing = False
                    Screen.MousePointer = vbDefault
                    Print #hmMsg, "Unable to Copy " & smResultFilePath & smDProd & "Exe\Traffic2.exe" & " To " & smResultFilePath & smDProd & "Updates\Traffic2.Exe" & " Error #" & Str$(ilRet)
                    MsgBox "Unable to copy traffic2.exe to updates folder.  See " & slMsgFile & " for Error messages", vbOKOnly + vbCritical + vbApplicationModal, "Create Error"
                    smResultFilePath = slDestination & smDProd
                    Close #hmMsg
                    mInstall = False
                    Exit Function
                End If
                ilRet = 0
                FileCopy smResultFilePath & smDProd & "Updates\Traffic2.exe", smResultFilePath & smDProd & "Exe\Traffic2.Exe"
                If ilRet <> 0 Then
                    On Error Resume Next
                    Kill smResultFilePath & smDProd & "Updates\Traffic2.exe"
                    imProcessing = False
                    Screen.MousePointer = vbDefault
                    Print #hmMsg, "Traffic2 in Use, Unable to Copy " & smResultFilePath & smDProd & "Updates\Traffic2.exe" & " To " & smResultFilePath & smDProd & "Exe\Traffic2.Exe" & " Error #" & Str$(ilRet)
                    MsgBox "Traffic in use.  See " & slMsgFile & " for Error messages", vbOKOnly + vbCritical + vbApplicationModal, "Create Error"
                    smResultFilePath = slDestination & smDProd
                    Close #hmMsg
                    mInstall = False
                    Exit Function
                End If
                On Error Resume Next
                Kill smResultFilePath & smDProd & "Updates\Traffic2.exe"
            End If
            lacProgress(3).Caption = "Checking if Affiliate System in Use"
            DoEvents
            ilRet = 0
            On Error GoTo mMakeFolderErr:
            slDateTime = FileDateTime(smResultFilePath & smDProd & "Exe\Affiliat.exe")
            If ilRet = 0 Then
                ilRet = 0
                On Error GoTo mMakeFolderErr:
                FileCopy smResultFilePath & smDProd & "Exe\Affiliat.exe", smResultFilePath & smDProd & "Updates\Affiliat.Exe"
                If ilRet <> 0 Then
                    imProcessing = False
                    Screen.MousePointer = vbDefault
                    Print #hmMsg, "Unable to Copy " & smResultFilePath & smDProd & "Exe\Affiliat.exe" & " To " & smResultFilePath & smDProd & "Updates\Affiliat.Exe" & " Error #" & Str$(ilRet)
                    MsgBox "Unable to copy Affiliat.exe to updates folder.  See " & slMsgFile & " for Error messages", vbOKOnly + vbCritical + vbApplicationModal, "Create Error"
                    smResultFilePath = slDestination & smDProd
                    Close #hmMsg
                    mInstall = False
                    Exit Function
                End If
                ilRet = 0
                FileCopy smResultFilePath & smDProd & "Updates\Affiliat.exe", smResultFilePath & smDProd & "Exe\Affiliat.Exe"
                If ilRet <> 0 Then
                    On Error Resume Next
                    Kill smResultFilePath & smDProd & "Updates\Affiliat.exe"
                    imProcessing = False
                    Screen.MousePointer = vbDefault
                    Print #hmMsg, "Affiliate in Use, Unable to Copy " & smResultFilePath & smDProd & "Updates\Affiliat.exe" & " To " & smResultFilePath & smDProd & "Exe\Affiliat.Exe" & " Error #" & Str$(ilRet)
                    MsgBox "Affiliate in use.  See " & slMsgFile & " for Error messages", vbOKOnly + vbCritical + vbApplicationModal, "Create Error"
                    smResultFilePath = slDestination & smDProd
                    Close #hmMsg
                    mInstall = False
                    Exit Function
                End If
                On Error Resume Next
                Kill smResultFilePath & smDProd & "Updates\Affiliat.exe"
            End If
            DoEvents
            ilRet = 0
            On Error GoTo mMakeFolderErr:
            slDateTime = FileDateTime(smResultFilePath & smDProd & "Exe\Affiliat2.exe")
            If ilRet = 0 Then
                ilRet = 0
                On Error GoTo mMakeFolderErr:
                FileCopy smResultFilePath & smDProd & "Exe\Affiliat2.exe", smResultFilePath & smDProd & "Updates\Affiliat2.Exe"
                If ilRet <> 0 Then
                    imProcessing = False
                    Screen.MousePointer = vbDefault
                    Print #hmMsg, "Unable to Copy " & smResultFilePath & smDProd & "Exe\Affiliat2.exe" & " To " & smResultFilePath & smDProd & "Updates\Affiliat2.Exe" & " Error #" & Str$(ilRet)
                    MsgBox "Unable to copy affiliat2.exe to updates folder.  See " & slMsgFile & " for Error messages", vbOKOnly + vbCritical + vbApplicationModal, "Create Error"
                    smResultFilePath = slDestination & smDProd
                    Close #hmMsg
                    mInstall = False
                    Exit Function
                End If
                ilRet = 0
                FileCopy smResultFilePath & smDProd & "Updates\Affiliat2.exe", smResultFilePath & smDProd & "Exe\Affiliat2.Exe"
                If ilRet <> 0 Then
                    On Error Resume Next
                    Kill smResultFilePath & smDProd & "Updates\Affiliat2.exe"
                    imProcessing = False
                    Screen.MousePointer = vbDefault
                    Print #hmMsg, "Affiliat2 in Use, Unable to Copy " & smResultFilePath & smDProd & "Updates\Affiliat2.exe" & " To " & smResultFilePath & smDProd & "Exe\Affiliat2.Exe" & " Error #" & Str$(ilRet)
                    MsgBox "Affiliate in use.  See " & slMsgFile & " for Error messages", vbOKOnly + vbCritical + vbApplicationModal, "Create Error"
                    smResultFilePath = slDestination & smDProd
                    Close #hmMsg
                    mInstall = False
                    Exit Function
                End If
                On Error Resume Next
                Kill smResultFilePath & smDProd & "Updates\Affiliat2.exe"
            End If
        End If
    End If
    'Move Zero BTR if updating system to temporary folder so it can be used to place into databases prior to running DDFReorg
    If (imCSIExist) And (smSystemType <> "E") Then
        lacProgress(3).Caption = "Copying Zero_Btr to temporary Folder"
        DoEvents
        mRemoveTempBTR slDestination & smDProd & "Updates"
        ilRet = mMoveZeroBTR(slDestination & smDProd & "Zero_BTR", slDestination & smDProd & "Updates")
        If Not ilRet Then
            imProcessing = False
            Screen.MousePointer = vbDefault
            MsgBox "Unable to copy zero_btr files to updates folder.  See " & slMsgFile & " for Error messages", vbOKOnly + vbCritical + vbApplicationModal, "Create Error"
            smResultFilePath = slDestination & smDProd
            Close #hmMsg
            mInstall = False
            Exit Function
        End If
    End If
    'Create All Folders
    lacProgress(3).Caption = "Creating Folders and Subfolders"
    DoEvents
    ilRet = 0
    ilError = False
    For ilLoop = 0 To UBound(smsourcefolder) - 1 Step 1
        slToName = Mid$(smsourcefolder(ilLoop), Len(slSource) + 1)
        ''If (InStr(1, slToName, "Zip", vbTextCompare) <= 0) Then
        'If (StrComp(slToName, "Zip", vbTextCompare) = 0) Then
        If (InStr(1, slToName, "Zip", vbTextCompare) <= 0) Or (InStr(1, slToName, "WinZip", vbTextCompare) > 0) Then
            'If InStr(1, slToName, "Prod", vbTextCompare) > 0 Then
            If InStr(1, slToName, smVersion, vbTextCompare) > 0 Then
               If StrComp(slToName, smSProd, vbTextCompare) = 0 Then
                    slToName = slDestination & left$(smDProd, Len(smDProd) - 1)
                Else
                    slToName = slDestination & smDProd & Mid$(slToName, Len(smSProd) + 2)
                End If
            Else
                slToName = slDestination & slToName
            End If
            
            'slDrivePath = slDestination & Mid$(smSourceFolder(ilLoop), Len(slSource) + 1)
            slDrivePath = slToName
            On Error GoTo mMakeFolderErr:
            ilRet = 0
            MkDir slDrivePath
            If ilRet = 0 Then
                Print #hmMsg, slDrivePath & " Created"
            ElseIf ilRet = 75 Then
                Print #hmMsg, slDrivePath & " Previously Existed"
            Else
                Print #hmMsg, "Unable to create " & slDrivePath & " Error #" & Str$(ilRet)
                ilError = True
            End If
        End If
    Next ilLoop
    If ilError Then
        imProcessing = False
        Screen.MousePointer = vbDefault
        MsgBox "Unable to create required Folders.  See " & slMsgFile & " for Error messages", vbOKOnly + vbCritical + vbApplicationModal, "Create Error"
        smResultFilePath = slDestination & smDProd
        Close #hmMsg
        mInstall = False
        Exit Function
    End If
    lacProgress(1).Caption = "Installing into:"
    lacProgress(2).Caption = "File:"
    lacProgress(1).Visible = True
    lacProgress(2).Visible = True
    lacProgress(3).Caption = ""
    DoEvents
    'Dan M 6/3/10 clean out setup folder. Removed from next loop.
    For ilLoop = 0 To UBound(smsourcefolder) - 1 Step 1
        smModFilePath = smsourcefolder(ilLoop) & "\"
        If mFolderToClean(smModFilePath) And imCSIExist Then
            mEraseOldFiles smModFilePath, slSource, slDestination
        End If
    Next ilLoop
    'Copy or unzip files into folders
    ilRet = 0
    ilCount = 0
    For ilLoop = 0 To UBound(smsourcefolder) - 1 Step 1
        smModFilePath = smsourcefolder(ilLoop) & "\"
        'Dan added 5/07/09 to clean out setup folder
        'lbcSource.Path = smSourceFolder(ilLoop)
        'If mFolderToClean(smModFilePath) And imCSIExist Then
         '   mEraseOldFiles smModFilePath, slSource, slDestination
        'End If
        lbcSourceFiles.Path = smsourcefolder(ilLoop)
        For ilIndex = 0 To lbcSourceFiles.ListCount - 1 Step 1
            slName = lbcSourceFiles.List(ilIndex)
            lacProgreFileListBox.Caption = slName
            slToName = Mid$(smsourcefolder(ilLoop), Len(slSource) + 1)
            ''If (InStr(1, slToName, "Zip", vbTextCompare) > 0) Then
            'If (StrComp(slToName, "Zip", vbTextCompare) <> 0) Then
            If (InStr(1, slToName, "Zip", vbTextCompare) > 0) And (InStr(1, slToName, "WinZip", vbTextCompare) <= 0) Then
                smResultFilePath = slDestination & smDProd
                ilRet = mUnZipFiles(slName)
            Else
                If InStr(1, slToName, smSProd, vbTextCompare) > 0 Then
                    slToName = slDestination & smDProd & Mid$(slToName, Len(smSProd) + 2) & "\"
                Else
                    slToName = slDestination & slToName & "\"
                End If
                'smResultFilePath = slToName & Mid$(smSourceFolder(ilLoop), Len(slSource) + 1) & "\"
                smResultFilePath = slToName
                lacProgressFolder.Caption = left$(smResultFilePath, Len(smResultFilePath) - 1)
                ilRet = mCopyFile(slName)
            End If
            If imCancelled Then
                Screen.MousePointer = vbDefault
                plcProgress.Visible = False
                plcCancel.Visible = True
                cmcNext.Enabled = True
                imCancelAns = 0
                Do
                    DoEvents
                Loop While imCancelAns = 0
                plcProgress.Visible = True
                plcCancel.Visible = False
                cmcNext.Enabled = False
                If imCancelAns = 2 Then
                    imProcessing = False
                    Screen.MousePointer = vbDefault
                    Print #hmMsg, ""
                    Print #hmMsg, "Cancelled " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                    Close #hmMsg
                    smResultFilePath = slDestination & smDProd
                    Close #hmMsg
                    mInstall = False
                    Exit Function
                Else
                    Screen.MousePointer = vbHourglass
                    imCancelled = False
                End If
            End If
            ilCount = ilCount + 1
            lacProgress(3).Caption = Str$(ilCount) & " of" & Str$(ilSourceCount)
            llPercent = (CLng(ilCount) * 100) \ ilSourceCount
            If llPercent >= 100 Then
                llPercent = 100
            End If
            lncProgress.X2 = ((pbcProgress.Width - 2 * lncProgress.X1) * llPercent) / 100
            If llPercent > 0 Then
                lncProgress.Visible = True
            End If
            DoEvents
        Next ilIndex
    Next ilLoop
    smResultFilePath = slDestination & smDProd
    mCheckDDF
    If lbcDatabases.ListCount <= 0 Then
        mRemoveTempBTR smResultFilePath & "Updates"
        Print #hmMsg, ""
        Print #hmMsg, "CSI File Images Install Successfully " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        Print #hmMsg, ""
        Close #hmMsg
        If imDatabaseFound Then
            lacFinish(4).Caption = "Installed" & Str$(ilCount) & " of" & Str$(ilSourceCount) & " files"
        Else
            lacFinish(4).Caption = "Installed" & Str$(ilCount) & " of" & Str$(ilSourceCount) & " files.  Unable to Find Databases, DDFReorg Not Run"
        End If
    Else
        Print #hmMsg, ""
        Print #hmMsg, "CSI File Images Install Successfully " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        Print #hmMsg, ""
        lacFinish(4).Caption = "Installed" & Str$(ilCount) & " of" & Str$(ilSourceCount) & " files"
    End If
    imProcessing = False
    mInstall = True
    Exit Function
mMakeFolderErr:
    ilRet = Err
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Function mOpenMsgFile(slMsgFile As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    ilRet = 0
    On Error GoTo mOpenMsgFileErr:
    'slToFile = sgExportPath & "ExpStnFd.Txt"
    If imDDFReorgOnly Then
        slToFile = smResultFilePath & "MultiDDFReorg.Txt"
    Else
        '4962
'        slToFile = smResultFilePath & "CSISetup.Txt"
        slToFile = smResultFilePath & "CSISetup_" & Format$(Now(), "mm-dd-yy") & ".Txt"
    End If
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
'Remove appending to file 11/4/03-Jim
'        slFileDate = Format$(slDateTime, "m/d/yy")
'        If DateValue(slFileDate) = DateValue(smTodaysDate) Then  'Append
'            On Error GoTo 0
'            ilRet = 0
'            On Error GoTo mOpenMsgFileErr:
'            hmMsg = FreeFile
'            Open slToFile For Append As hmMsg
'            If ilRet <> 0 Then
'                Screen.MousePointer = vbDefault
'                MsgBox "Open " & slToFile & " Error #" & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
'                mOpenMsgFile = False
'                Exit Function
'            End If
'        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            On Error GoTo mOpenMsgFileErr:
            hmMsg = FreeFile
            Open slToFile For Output As hmMsg
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & " Error #" & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
'        End If
    Else
        On Error GoTo 0
        ilRet = 0
        On Error GoTo mOpenMsgFileErr:
        hmMsg = FreeFile
        Open slToFile For Output As hmMsg
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile & " Error #" & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, ""
    slMsgFile = slToFile
    mOpenMsgFile = True
    Exit Function
mOpenMsgFileErr:
    ilRet = 1
    Resume Next
End Function
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
    Dim slCommand As String
    Dim slStr As String
    Dim ilPos As Integer
    Dim slChar As String
    Dim slName As String
    slCommand = Command$
    imDDFReorgOnly = False
    slName = App.EXEName
    ilPos = InStr(1, slName, ",", vbTextCompare)
    If ilPos > 0 Then
        slName = left$(slName, ilPos - 1)
    End If
    ilPos = InStr(1, slCommand, "/RunDDFReorgOnly", vbTextCompare)
    If (ilPos > 0) Or (StrComp(slName, "MultiDDFReorg", vbTextCompare) = 0) Then
        imDDFReorgOnly = True
        lacWMessage(0).Caption = "Welcome to Multi-DDFReorg Program for Counterpoint System"
        lacBrowse(0).Caption = "Open Folder that the Databases are Subfolder to"
        lacFinish(1).Caption = "The Running of DDFReorg was Successful"
        lacFinish(3).Caption = ""
        lacFinish(4).Caption = ""
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:4/17/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: terminate Links                *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
    Dim ilRet As Integer
    
    Erase smKillNames
    Erase tmUnZipName
    Erase tmZippedNames
    Screen.MousePointer = vbDefault
    Unload CSISetup
    Set CSISetup = Nothing
    End
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mUnZipFiles                     *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Unzip files                     *
'*                                                     *
'*******************************************************
Private Function mUnZipFiles(slFromFile As String) As Integer
    Dim slDateTime As String
    Dim slSubfolder As String
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim slChar As String
    Dim ilCount As Integer
    Dim ilLoop As Integer
    Dim slFileName As String
    Dim slDate As String
    Dim slTime As String
    Dim ilUpper As Integer
    Dim slFTime As String
    Dim llFTime As Long
    Dim slUTime As String
    Dim llUTime As Long
    Dim slFile As String
    Dim ilAddFile As Integer
    Dim ilFileExist As Integer
    Dim llSize As Long
    Dim ilOvMsg As Integer
    Dim ilSameTime As Integer
    Dim slVersion As String
    Dim slSaveOldSubfolder As String
    
    ReDim tmZippedNames(0 To 0) As ZIPPEDNAMES
    ilPos = InStr(1, smVersion, ".", 1)
    If ilPos > 0 Then
        slVersion = UCase$(left$(smVersion, ilPos - 1) & Mid$(smVersion, ilPos + 1))
    Else
        slVersion = UCase$(smVersion)
    End If
    slSaveOldSubfolder = "Updates\old"
    'Build array of files to be unzipped
    ReDim tmUnZipName(1 To 1) As UNZIPNAME
    ilUpper = 0
    ilPos = InStrRev(slFromFile, ".")
    If ilPos > 0 Then
        slChar = Mid$(slFromFile, ilPos + 1, 1)
        Select Case slChar
            Case "E"
                slSubfolder = "exe"
            Case "R", "J"
                slSubfolder = "reports"
            Case "A"
                slSubfolder = "affexe"
            Case "B"
                ilRet = 0
                On Error GoTo mUnZipFilesErr:
                slSubfolder = "affrepor"
                lbcAffReports.Path = smResultFilePath & slSubfolder
                If ilRet <> 0 Then
                    slSubfolder = "affreports"
                End If
            Case "N"
                slSubfolder = "NewDDF"
            Case "Z"
                slSubfolder = "zero_btr"
            Case "G"
                slSubfolder = "reports\generic"
            Case "D"
                slSubfolder = "data"
            Case "S"
                slSubfolder = "Setup\ClientInstall\System32"
            Case "F"
                slSubfolder = "Setup\Fonts"
        End Select
        lacProgressFolder.Caption = smResultFilePath & slSubfolder
        initUNZIPCmdStruct zpcDUnZip
        zpcDUnZip.ZIPFile = smModFilePath & slFromFile
        zpcDUnZip.QuietFlag = True
        zpcDUnZip.Filespec = "*.*"
        zpcDUnZip.NoDirectoryNamesFlag = True
        zpcDUnZip.ActionDZ = UNZIP_COUNTNAMEDZIPMEMBERS
        If zpcDUnZip.ErrorCode <> UE_OK Then
            'lbcResult.AddItem "Not Valid Zip File " & Trim$(smModFilePath & slFromFile) & " Error #" & Str$(zpcDUnZip.ErrorCode)
            Print #hmMsg, "     " & "Not Valid Zip File " & Trim$(smModFilePath & slFromFile) & " Error #" & Str$(zpcDUnZip.ErrorCode)
        Else
            ilCount = zpcDUnZip.ReturnCount
            For ilLoop = 1 To ilCount Step 1
                zpcDUnZip.NoDirectoryNamesFlag = True
                zpcDUnZip.ActionDZ = UNZIP_GETNEXTNAMEDZIPINFO
                ilUpper = UBound(tmUnZipName)
                slFileName = zpcDUnZip.zi_FileName
                ilPos = Len(slFileName)
                slChar = Mid$(slFileName, ilPos, 1)
                slFile = ""
                Do While slChar <> "\"
                    slFile = slChar & slFile
                    ilPos = ilPos - 1
                    If ilPos <= 0 Then
                        Exit Do
                    End If
                    slChar = Mid$(slFileName, ilPos, 1)
                Loop
                tmUnZipName(ilUpper).sName = slFile
                tmUnZipName(ilUpper).sZipName = gGetTrueFileExt(slFile)
                tmUnZipName(ilUpper).sDate = gAdjYear(Format$(zpcDUnZip.zi_DateTime, "m/d/yy"))
                tmUnZipName(ilUpper).sTime = Format$(zpcDUnZip.zi_DateTime, "h:mm:ssAM/PM")
                tmUnZipName(ilUpper).lSize = zpcDUnZip.zi_oSize
                tmUnZipName(ilUpper).sOldDate = ""
                tmUnZipName(ilUpper).sOldTime = ""
                tmUnZipName(ilUpper).lOldSize = 0
                ilAddFile = False
                ilFileExist = False
                ilOvMsg = False
                ilRet = 0
                On Error GoTo mUnZipFilesErr:
                slDateTime = FileDateTime(smResultFilePath & slSubfolder & "\" & Trim$(tmUnZipName(ilUpper).sZipName))
                If ilRet = 0 Then
                    ilFileExist = True
                    slDate = gAdjYear(Format$(slDateTime, "m/d/yy"))
                    slTime = Format$(slDateTime, "h:mm:ssAM/PM")
                    llSize = FileLen(smResultFilePath & slSubfolder & "\" & Trim$(tmUnZipName(ilUpper).sZipName))
                    tmUnZipName(ilUpper).sOldDate = slDate
                    tmUnZipName(ilUpper).sOldTime = slTime
                    tmUnZipName(ilUpper).lOldSize = llSize
                    slFTime = Format$(slTime, "hh:mm:ss")
                    'ignore seconds as zip does not return the seconds
                    llFTime = 3600 * Hour(slFTime) + 60 * Minute(slFTime) '+ Second(slFTime)
                    slUTime = Format$(tmUnZipName(ilUpper).sTime, "hh:mm:ss")
                    llUTime = 3600 * Hour(slUTime) + 60 * Minute(slUTime) '+ Second(slUTime)
                    'If DateValue(tmUnZipName(ilUpper).sDate) > DateValue(slDate) Then
                    '    'ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
                    '    ilAddFile = True
                    'ElseIf DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate) Then
                    '    If llUTime >= llFTime Then
                    '        'ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
                    '        ilAddFile = True
                    '    Else
                    '        'ilRet = MsgBox("Overwrite: " & Trim$(tmUnZipName(ilUpper).sName) & Chr(10) & Chr(13) & "     " & slDate & " " & slTime & Str$(llSize) & " with " & Chr(10) & Chr(13) & "     " & Format$(tmUnZipName(ilUpper).sDate, "m/d/yy") & " " & Format$(tmUnZipName(ilUpper).sTime, "h:mm:ssAM/PM") & Str$(tmUnZipName(ilUpper).lSize), 4, "Overwrite")
                    '        'If ilRet = 6 Then
                    '        '    'ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
                    '        '    ilAddFile = True
                    '        'End If
                    '        If llSize <> tmUnZipName(ilUpper).lSize Then
                    '            ilOvMsg = True
                    '        Else
                    '            ilOvMsg = True
                    '        End If
                    '    End If
                    'Else
                    '    'ilRet = MsgBox("Overwrite: " & Trim$(tmUnZipName(ilUpper).sName) & Chr(10) & Chr(13) & "     " & slDate & " " & slTime & Str$(llSize) & " with " & Chr(10) & Chr(13) & "     " & Format$(tmUnZipName(ilUpper).sDate, "m/d/yy") & " " & Format$(tmUnZipName(ilUpper).sTime, "h:mm:ssAM/PM") & Str$(tmUnZipName(ilUpper).lSize), 4, "Overwrite")
                    '    'If ilRet = 6 Then
                    '    '    'ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
                    '    '    ilAddFile = True
                    '    'End If
                    '    ilOvMsg = True
                    'End If
                    If DateValue(tmUnZipName(ilUpper).sDate) > DateValue(slDate) + 1 Then
                        ilAddFile = True
                    ElseIf DateValue(tmUnZipName(ilUpper).sDate) < DateValue(slDate) - 1 Then
                        ilOvMsg = True
                    Else
                        ilSameTime = False
                        If (llUTime = llFTime) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                            ilSameTime = True
                        Else
                            If Minute(slUTime) = Minute(slFTime) Then
                                If Hour(slUTime) = 0 Then
                                    If (Hour(slFTime) >= 21) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate) + 1) Then
                                        ilSameTime = True
                                    End If
                                    If (Hour(slFTime) <= 3) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                                        ilSameTime = True
                                    End If
                                ElseIf Hour(slUTime) = 1 Then
                                    If (Hour(slFTime) >= 22) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate) + 1) Then
                                        ilSameTime = True
                                    End If
                                    If (Hour(slFTime) <= 4) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                                        ilSameTime = True
                                    End If
                                ElseIf Hour(slUTime) = 2 Then
                                    If (Hour(slFTime) >= 23) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate) + 1) Then
                                        ilSameTime = True
                                    End If
                                    If (Hour(slFTime) <= 5) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                                        ilSameTime = True
                                    End If
                                ElseIf Hour(slUTime) = 23 Then
                                    If (Hour(slFTime) >= 20) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                                        ilSameTime = True
                                    End If
                                    If (Hour(slFTime) <= 2) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate) - 1) Then
                                        ilSameTime = True
                                    End If
                                ElseIf Hour(slUTime) = 22 Then
                                    If (Hour(slFTime) >= 19) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                                        ilSameTime = True
                                    End If
                                    If (Hour(slFTime) <= 1) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate) - 1) Then
                                        ilSameTime = True
                                    End If
                                ElseIf Hour(slUTime) = 21 Then
                                    If (Hour(slFTime) >= 18) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                                        ilSameTime = True
                                    End If
                                    If (Hour(slFTime) <= 0) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate) - 1) Then
                                        ilSameTime = True
                                    End If
                                Else
                                    If (Hour(slFTime) >= Hour(slUTime) - 3) And (Hour(slFTime) <= Hour(slUTime) + 3) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                                        ilSameTime = True
                                    End If
                                End If
                            End If
                        End If
                        If ilSameTime And (llSize = tmUnZipName(ilUpper).lSize) Then
                            ilAddFile = True
                        Else
                            If DateValue(tmUnZipName(ilUpper).sDate) > DateValue(slDate) And (Not ilSameTime) Then
                                ilAddFile = True
                            Else
                                If (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) And (llUTime > llFTime) And (Not ilSameTime) Then
                                    ilAddFile = True
                                Else
                                    ilOvMsg = True
                                End If
                            End If
                        End If
                    End If
                Else
                    'ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
                    ilAddFile = True
                End If
                'If ilOvMsg Then
                '    ilRet = MsgBox("Overwrite: " & Trim$(tmUnZipName(ilUpper).sZipName) & Chr(10) & Chr(13) & "     " & slDate & " " & slTime & Str$(llSize) & " with " & Chr(10) & Chr(13) & "     " & Format$(tmUnZipName(ilUpper).sDate, "m/d/yy") & " " & Format$(tmUnZipName(ilUpper).sTime, "h:mm:ssAM/PM") & Str$(tmUnZipName(ilUpper).lSize), 4, "Overwrite")
                '    If ilRet = 6 Then
                '        'ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
                '        ilAddFile = True
                '    End If
                'End If
                '5/19/05:  Always replace images as problem with replacing images from test or dev into prod as prod
                'might have had a new date or same date but newer time (modified test then prod).
                ilAddFile = True
                If ilAddFile Then
                    If ilFileExist Then
                        ilRet = GetAttr(smResultFilePath & slSubfolder & "\" & Trim$(tmUnZipName(ilUpper).sZipName)) And 1
                        If ilRet <> 0 Then
                            SetAttr smResultFilePath & slSubfolder & "\" & Trim$(tmUnZipName(ilUpper).sZipName), 0
                        End If
                    End If
                    ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
                End If
            Next ilLoop
            
    
            For ilLoop = 1 To UBound(tmUnZipName) - 1 Step 1
                'If (rbcSave(0).Value) And (tmUnZipName(ilLoop).lOldSize > 0) Then
                '    ilRet = mZipFile(slFromFile, smResultFilePath & slSubFolder & "\" & tmUnZipName(ilLoop).sName, tmUnZipName(ilLoop).sOldDate, tmUnZipName(ilLoop).sOldTime, tmUnZipName(ilLoop).lOldSize)
                'End If
                If (tmUnZipName(ilLoop).lOldSize > 0) Then
                    smSaveFilePath = smResultFilePath & slSaveOldSubfolder & "\"
                    ilRet = mZipFile(slFromFile, smResultFilePath & slSubfolder & "\" & tmUnZipName(ilLoop).sZipName, tmUnZipName(ilLoop).sOldDate, tmUnZipName(ilLoop).sOldTime, tmUnZipName(ilLoop).lOldSize)
                End If
                initUNZIPCmdStruct zpcDUnZip
                zpcDUnZip.ZIPFile = smModFilePath & slFromFile
                zpcDUnZip.Destination = smResultFilePath & slSubfolder
                zpcDUnZip.NoDirectoryNamesFlag = True
                zpcDUnZip.Filespec = Trim$(tmUnZipName(ilLoop).sName)
                zpcDUnZip.OverwriteFlag = True
                zpcDUnZip.DiagnosticFlag = True
                zpcDUnZip.DecryptFlag = True
                zpcDUnZip.DecryptCode = slVersion
                zpcDUnZip.ActionDZ = UNZIP_EXTRACT
                ilRet = zpcDUnZip.ErrorCode
                If ilRet = 0 Then
                    'lbcResult.AddItem Trim$(tmUnZipName(ilLoop).sZipName) & " Unzipped into " & smResultFilePath & slSubFolder
                    If Trim$(tmUnZipName(ilLoop).sOldDate) = "" Then
                        Print #hmMsg, "     " & Trim$(tmUnZipName(ilLoop).sZipName) & " " & Trim$(tmUnZipName(ilLoop).sDate) & " " & Trim$(tmUnZipName(ilLoop).sTime) & " Size" & Str$(tmUnZipName(ilLoop).lSize) & " Unzipped into " & smResultFilePath & slSubfolder
                    Else
                        Print #hmMsg, "     " & Trim$(tmUnZipName(ilLoop).sZipName) & " " & Trim$(tmUnZipName(ilLoop).sDate) & " " & Trim$(tmUnZipName(ilLoop).sTime) & " Size" & Str$(tmUnZipName(ilLoop).lSize) & " Unzipped into " & smResultFilePath & slSubfolder & " Replaced: " & Trim$(tmUnZipName(ilLoop).sOldDate) & " " & Trim$(tmUnZipName(ilLoop).sOldTime) & " Size" & Str$(tmUnZipName(ilLoop).lOldSize)
                    End If
                Else
                    'lbcResult.AddItem Trim$(tmUnZipName(ilLoop).sName) & " Error #" & Str$(ilRet)
                    Print #hmMsg, "     " & Trim$(tmUnZipName(ilLoop).sName) & " Error #" & Str$(ilRet)
                End If
            Next ilLoop
        End If
    End If
    Exit Function
mUnZipFilesErr:
    ilRet = Err
    Resume Next
End Function
Private Function mZipFile(slZipName As String, slInName As String, slInDate As String, slInTime As String, llSize As Long) As Integer
    Dim slDateTime As String
    Dim ilRet As Integer
    Dim slName As String
    Dim slZipPathName As String
    Dim ilPos As Integer
    Dim ilNameFd As Integer
    Dim ilLoop As Integer
    Dim slVersion As String
    
    ilPos = InStr(1, smVersion, ".", 1)
    If ilPos > 0 Then
        slVersion = UCase$(left$(smVersion, ilPos - 1) & Mid$(smVersion, ilPos + 1))
    Else
        slVersion = UCase$(smVersion)
    End If
    initZIPCmdStruct zpcDZip
    On Error GoTo mZipFilesErr:
    slZipPathName = smSaveFilePath & Trim$(slZipName)
    ilNameFd = False
    For ilLoop = LBound(tmZippedNames) To UBound(tmZippedNames) - 1 Step 1
        If StrComp(Trim$(slZipName), Trim$(tmZippedNames(ilLoop).sName), 1) = 0 Then
            ilNameFd = True
            Exit For
        End If
    Next ilLoop
    If Not ilNameFd Then
        ilRet = 0
        slDateTime = FileDateTime(slZipPathName)
        If ilRet = 0 Then
            Kill slZipPathName
        End If
        tmZippedNames(UBound(tmZippedNames)).sName = Trim$(slZipName)
        ReDim Preserve tmZippedNames(0 To UBound(tmZippedNames) + 1) As ZIPPEDNAMES
    End If
    On Error GoTo 0
    zpcDZip.ZIPFile = Trim$(slZipPathName)
    zpcDZip.ItemList = Trim$(slInName)
    zpcDZip.OldAsLatestFlag = True
    zpcDZip.NoDirectoryNamesFlag = True
    zpcDZip.EncryptFlag = True
    zpcDZip.EncryptCode = slVersion
    If Not ilNameFd Then
        zpcDZip.ActionDZ = ZIP_ADD
    Else
        zpcDZip.ActionDZ = ZIP_UPDATE
    End If
    ilRet = zpcDZip.ErrorCode
    If ilRet = 0 Then
        Print #hmMsg, "     " & Trim$(slInName) & " " & Trim$(slInDate) & " " & Trim$(slInTime) & " Size" & Str$(llSize) & " saved into " & Trim$(slZipPathName)
        'lbcResult.AddItem Trim$(slInName) & " zipped into " & slName
    Else
        Print #hmMsg, "     " & Trim$(slInName) & " saved into " & Trim$(slZipPathName) & " Failed: " & Str$(ilRet)
        'lbcResult.AddItem Trim$(slInName) & " not zipped"
    End If
    Exit Function
mZipFilesErr:
    ilRet = 1
    Resume Next
End Function

Private Sub tmcStart_Timer()
    Dim ilRet As Integer
    tmcStart.Enabled = False
    Screen.MousePointer = vbHourglass
    If imScreenNo = 12 Then
        ilRet = mRunDDFReorg()
        mRemoveTempBTR smResultFilePath & "Updates"
        Screen.MousePointer = vbDefault
        If ilRet Then
            cmcNext_Click
        Else
            If imDDFReorgOnly Then
                lacFinish(3).Caption = ""
                lacFinish(4).Caption = ""
                If imCancelled Then
                    lacFinish(0).Caption = "Cancelled"
                    lacFinish(1).Caption = "The Running of DDFReorg was terminated"
                Else
                    lacFinish(0).Caption = "Error when Running DDFReorg"
                    lacFinish(1).Caption = "The Running of DDFReorg was stopped because of errors"
                End If
            Else
                If imCancelled Then
                    lacFinish(0).Caption = "Cancelled"
                    lacFinish(1).Caption = "The Install Program was terminated"
                    If imCSIExist Then
                        lacFinish(3).Caption = "Restore the previous version from the Backup"
                    End If
                Else
                    lacFinish(0).Caption = "Install Error"
                    lacFinish(1).Caption = "The Install Program was stopped because of errors"
                End If
            End If
            imCancelled = False
            cmcNext_Click
        End If
    Else
        ilRet = mInstall()
        Screen.MousePointer = vbDefault
        If ilRet Then
            cmcNext_Click
        Else
            mRemoveTempBTR smResultFilePath & "Updates"
            If imCancelled Then
                lacFinish(0).Caption = "Cancelled"
                lacFinish(1).Caption = "The Install Program was terminated"
                If imCSIExist Then
                    lacFinish(3).Caption = "Restore the previous version from the Backup"
                End If
            Else
                lacFinish(0).Caption = "Install Error"
                lacFinish(1).Caption = "The Install Program was stopped because of errors"
            End If
            imCancelled = False
            'cmcNext_Click
            mTerminate
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub zpcDUnZip_UnZipRenameCallback(ItemName As String, iDate As Long, iTime As Long, lAttrib As Long, OrigItemName As String, rc As Long)
            
    Dim ilPos As Integer
    Dim slTempStr As String
    
    ilPos = InStrRev(ItemName, ".", Len(ItemName) - 1)
    slTempStr = UCase(Mid$(ItemName, ilPos + 1))
    
    If slTempStr = "FWE" Then
        ItemName = left$(ItemName, ilPos) & "exe"
    End If
    
    If slTempStr = "FWL" Then
        ItemName = left$(ItemName, ilPos) & "dll"
    End If

    If slTempStr = "FWA" Then
        ItemName = left$(ItemName, ilPos) & "oca"
    End If

    If slTempStr = "FWX" Then
        ItemName = left$(ItemName, ilPos) & "ocx"
    End If

End Sub


Private Sub mCheckDDF()
    Dim ilRet As Integer
    Dim slNewDDFDateTime As String
    Dim slNewDDFDate As String
    Dim slDataDateTime As String
    Dim slDataDate As String
    Dim ilLoop As Integer
    Dim slNewDDFTVIValue As String
    Dim slDBTVIValue As String
    
    imDatabaseFound = False
    ilRet = 0
    lacProgressFolder.Caption = ""
    lacProgreFileListBox.Caption = ""
    lacProgress(3).Caption = "Checking DDF Status"
    DoEvents
    lbcDatabases.Clear
    On Error GoTo mCheckDDFErr:
    slNewDDFTVIValue = mGetTVIDate(smResultFilePath & "NewDDF\")
    slNewDDFDateTime = FileDateTime(smResultFilePath & "NewDDF\Field.ddf")
    slNewDDFDate = gAdjYear(Format$(slNewDDFDateTime, "m/d/yy"))
    If ilRet = 0 Then
        lbcDDFCheckDir.Path = left$(smResultFilePath, Len(smResultFilePath) - 1)
        For ilLoop = 0 To lbcDDFCheckDir.ListCount - 1 Step 1
            ilRet = 0
            On Error GoTo mCheckDDFErr:
            slDataDateTime = FileDateTime(lbcDDFCheckDir.List(ilLoop) & "\" & "Field.ddf")
            If ilRet = 0 Then
                imDatabaseFound = True
                slDBTVIValue = mGetTVIDate(lbcDDFCheckDir.List(ilLoop) & "\")
                If (slNewDDFTVIValue <> "") Or (slDBTVIValue <> "") Then
                    If StrComp(slNewDDFTVIValue, slDBTVIValue, vbTextCompare) <> 0 Then
                        lbcDatabases.AddItem lbcDDFCheckDir.List(ilLoop)
                    End If
                Else
                    slDataDate = gAdjYear(Format$(slDataDateTime, "m/d/yy"))
                    If DateValue(slNewDDFDate) <> DateValue(slDataDate) Then
                        lbcDatabases.AddItem lbcDDFCheckDir.List(ilLoop)
                    End If
                End If
            End If
        Next ilLoop
        If lbcDatabases.ListCount <= 0 Then
            If imDatabaseFound Then
                Print #hmMsg, "No Database found that require Conversion by DDFReorg"
            Else
                Print #hmMsg, "Unable to Find Database, therefore DDFReorg not run"
            End If
        End If
    Else
        Print #hmMsg, "Unable to Find " & smResultFilePath & "NewDDF\Field.ddf" & " therefore DDFReorg not run" & " Error #" & Str$(ilRet)
    End If
    Exit Sub
mCheckDDFErr:
    ilRet = Err
    Resume Next

End Sub

Private Function mRunDDFReorg() As Integer
    Dim ilLoop As Integer
    Dim slExePath As String
    Dim ilPos As Integer
    Dim ilCount As Integer
    Dim ilCountTotal As Integer
    Dim llPercent As Long
    Dim ilRet As Integer
    Dim slName As String
    Dim llLen As Long
    Dim ilRunDDFOffst As Integer
    Dim slDBName As String
    
    imProcessing = True
    Screen.MousePointer = vbHourglass
    lncProgress.X2 = lncProgress.X1
    lncProgress.Visible = False
    DoEvents
    ilCountTotal = 0
    For ilLoop = 0 To lbcDatabases.ListCount - 1 Step 1
        If lbcDatabases.Selected(ilLoop) Then
            ilCountTotal = ilCountTotal + 1
        End If
    Next ilLoop
    'Determine if DDFOffst should be run
    ilRet = 0
    On Error GoTo mRunDDFReorgErr:
    llLen = FileLen(smResultFilePath & "Exe\" & "csi_io32.dll")
    If ilRet <> 0 Then
        Print #hmMsg, "Unable to find " & smResultFilePath & "Exe\" & "csi_io32.dll" & " Error # " & Str$(ilRet)
        llLen = 800000
    End If
    'Two different csi_io32 routines.  One that uses Classic cbtrv432 and the other that
    'Jeff wrote and does not use cbtrv432.  Jeff does not require the DDFOffst file and
    'is about 900000.  The other is about 90000.
    If llLen > 200000 Then
        ilRunDDFOffst = False
    Else
        ilRunDDFOffst = True
    End If

    ilCount = 0
    For ilLoop = 0 To lbcDatabases.ListCount - 1 Step 1
        slDBName = lbcDatabases.List(ilLoop)
        If lbcDatabases.Selected(ilLoop) Then
            lacProgressFolder.Caption = slDBName
            lacProgreFileListBox.Caption = "DDFReorg"
            Print #hmMsg, "User Running DDFReorg on " & slDBName
            DoEvents
            On Error GoTo mRunDDFReorgErr:
            ilRet = 0
            slName = lbcDatabases.List(ilLoop) & "\Messages"
            MkDir slName
            If (ilRet <> 0) And (ilRet <> 75) Then
                lacProgressFolder.Caption = ""
                lacProgreFileListBox.Caption = ""
                imProcessing = False
                Screen.MousePointer = vbDefault
                Print #hmMsg, "Unable to create " & slName & " Error # " & Str$(ilRet)
                MsgBox "Unable to create " & slName & " Error # " & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Create Error"
                mRunDDFReorg = False
                Exit Function
            End If
            ilRet = mMoveZeroBTR(smResultFilePath & "Updates", lbcDatabases.List(ilLoop))
            slExePath = smResultFilePath & "Exe\DDFReorg /N:" & smResultFilePath & "NewDDF" & " /O:" & lbcDatabases.List(ilLoop) & " /Exit /Run"
            gShellAndWait CSISetup, slExePath, vbMinimizedFocus 'vbNormalFocus    'vbMinimizedFocus 'vbHide
            'Move the zero_btr files into database folder
            If imCancelled Then
                Screen.MousePointer = vbDefault
                plcProgress.Visible = False
                plcCancel.Visible = True
                cmcNext.Enabled = True
                imCancelAns = 0
                Do
                    DoEvents
                Loop While imCancelAns = 0
                plcProgress.Visible = True
                plcCancel.Visible = False
                cmcNext.Enabled = False
                If imCancelAns = 2 Then
                    lacProgressFolder.Caption = ""
                    lacProgreFileListBox.Caption = ""
                    imProcessing = False
                    Screen.MousePointer = vbDefault
                    Print #hmMsg, ""
                    Print #hmMsg, "Cancelled " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                    Close #hmMsg
                    mRunDDFReorg = False
                    Exit Function
                Else
                    Screen.MousePointer = vbHourglass
                    imCancelled = False
                End If
            Else
                Print #hmMsg, "DDFReorg Completed for " & slDBName
            End If
            ilRet = mMoveZeroBTR(smResultFilePath & "Zero_BTR", lbcDatabases.List(ilLoop))
            If ilRunDDFOffst Then
                lacProgreFileListBox.Caption = "DDFOffst"
                Print #hmMsg, "Running DDFOffst on " & slDBName
                DoEvents
                slExePath = smResultFilePath & "Exe\DDFOffst " & " /F:" & lbcDatabases.List(ilLoop) & " /Exit /Run"
                gShellAndWait CSISetup, slExePath, vbMinimizedFocus 'vbNormalFocus    'vbMinimizedFocus 'vbHide
                'Move the zero_btr files into database folder
                If imCancelled Then
                    Screen.MousePointer = vbDefault
                    plcProgress.Visible = False
                    plcCancel.Visible = True
                    cmcNext.Enabled = True
                    imCancelAns = 0
                    Do
                        DoEvents
                    Loop While imCancelAns = 0
                    plcProgress.Visible = True
                    plcCancel.Visible = False
                    cmcNext.Enabled = False
                    If imCancelAns = 2 Then
                        lacProgressFolder.Caption = ""
                        lacProgreFileListBox.Caption = ""
                        imProcessing = False
                        Screen.MousePointer = vbDefault
                        Print #hmMsg, ""
                        Print #hmMsg, "Cancelled " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                        Close #hmMsg
                        mRunDDFReorg = False
                        Exit Function
                    Else
                        Screen.MousePointer = vbHourglass
                        imCancelled = False
                    End If
                Else
                    Print #hmMsg, "DDFOffst Completed for " & slDBName
                End If
            End If
            ilCount = ilCount + 1
            lacProgress(3).Caption = Str$(ilCount) & " of" & Str$(ilCountTotal)
            llPercent = (CLng(ilCount) * 100) \ ilCountTotal
            If llPercent >= 100 Then
                llPercent = 100
            End If
            lncProgress.X2 = ((pbcProgress.Width - 2 * lncProgress.X1) * llPercent) / 100
            If llPercent > 0 Then
                lncProgress.Visible = True
            End If
            DoEvents
        Else
            Print #hmMsg, "User Skipped Running DDFReorg on " & slDBName
        End If
    Next ilLoop
    lacProgressFolder.Caption = ""
    lacProgreFileListBox.Caption = ""
    imProcessing = False
    Print #hmMsg, "CSI Databases Converted Successfully " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmMsg, ""
    Close #hmMsg
    If imDDFReorgOnly Then
        lacFinish(4).Caption = "Converted" & Str$(ilCount) & " of" & Str$(ilCountTotal) & " databases"
    Else
        lacFinish(4).Caption = lacFinish(4).Caption & ", Converted" & Str$(ilCount) & " of" & Str$(ilCountTotal) & " databases"
    End If
    mRunDDFReorg = True
    Exit Function
mRunDDFReorgErr:
    ilRet = Err
    Resume Next
End Function

Public Function mMoveZeroBTR(slFromSubfolder As String, slToSubfolder As String) As Integer
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilRet As Integer
    Dim ilOk As Integer
    Dim ilTest As Integer
    
    mMoveZeroBTR = True
    If smSystemType = "E" Then
        Exit Function
    End If
    lbcZeroBTR.Path = slFromSubfolder
    For ilLoop = 0 To lbcZeroBTR.ListCount - 1 Step 1
        slName = lbcZeroBTR.List(ilLoop)
        ilOk = False
        'If InStr(1, slName, "r.btr", vbTextCompare) = 3 Then
        '    ilOk = True
        'Else
            For ilTest = 1 To UBound(smZeroNames) Step 1
                If InStr(1, slName, smZeroNames(ilTest), vbTextCompare) = 1 Then
                    ilOk = True
                    Exit For
                End If
            Next ilTest
        'End If
        If ilOk Then
            slName = slFromSubfolder & "\" & lbcZeroBTR.List(ilLoop)
            ilRet = 0
            On Error GoTo mMoveZeroBTRErr:
            FileCopy slName, slToSubfolder & "\" & lbcZeroBTR.List(ilLoop)
            If ilRet <> 0 Then
                Print #hmMsg, "Copying " & Trim$(slName) & " To " & Trim$(slToSubfolder) & " Error #" & Str$(ilRet)
                mMoveZeroBTR = False
            End If
        End If
    Next ilLoop
    Exit Function
mMoveZeroBTRErr:
    ilRet = Err
    Resume Next
    
End Function

Private Sub mRemoveTempBTR(slSubfolder As String)
    Dim ilLoop As Integer
    Dim ilRet As Integer
    ReDim smKillNames(0 To 0) As String
    If smSystemType = "E" Then
        Exit Sub
    End If
    lbcZeroBTR.Path = slSubfolder
    For ilLoop = 0 To lbcZeroBTR.ListCount - 1 Step 1
        smKillNames(UBound(smKillNames)) = slSubfolder & "\" & lbcZeroBTR.List(ilLoop)
        ReDim Preserve smKillNames(0 To UBound(smKillNames) + 1) As String
    Next ilLoop
    For ilLoop = 0 To UBound(smKillNames) - 1 Step 1
        ilRet = 0
        On Error GoTo mRemoveTempBTRErr:
        Kill smKillNames(ilLoop)
        If ilRet <> 0 Then
            Print #hmMsg, "Zero_Btr Names not removed from " & Trim$(slSubfolder) & " Error #" & Str$(ilRet)
        End If
    Next ilLoop
    Exit Sub
mRemoveTempBTRErr:
    ilRet = Err
    Resume Next

End Sub

Private Function mGetTVIDate(slDBPath As String) As String
    Dim ilRet As Integer
    mGetTVIDate = ""
    hmDDFFile = CBtrvTable()    'CBtrvObj()
    ilRet = btrOpen(hmDDFFile, "", slDBPath & "File.ddf", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmDDFFile)
        btrDestroy hmDDFFile
        Exit Function
    End If
    lmActualRecLen = Len(tmDDFFile) 'Len(tmRec)
    lmBaseRecLen = lmActualRecLen
    ilRet = btrGetFirst(hmDDFFile, tmRec, lmActualRecLen, INDEXKEY0, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmDDFFile)
        btrDestroy hmDDFFile
        Exit Function
    End If
    Do While ilRet = BTRV_ERR_NONE
        mGetFileFields
        If StrComp(left$(tmDDFFile.sName, 3), "TVI", vbTextCompare) = 0 Then
            mGetTVIDate = Trim$(tmDDFFile.sName)
            ilRet = btrClose(hmDDFFile)
            btrDestroy hmDDFFile
            Exit Function
        End If
        lmActualRecLen = lmBaseRecLen
        ilRet = btrGetNext(hmDDFFile, tmRec, lmActualRecLen, BTRV_LOCK_NONE)
    Loop
    ilRet = btrClose(hmDDFFile)
    btrDestroy hmDDFFile

End Function

Public Sub mGetFileFields()
    Dim tlDDFField As DDFFIELD
    
    tlDDFField.iOffset = 0
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmDDFFile.iFileID = Val(smFieldValue)
    tlDDFField.iOffset = 2
    tlDDFField.iSize = 20
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmDDFFile.sName = smFieldValue
    tlDDFField.iOffset = 22
    tlDDFField.iSize = 64
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmDDFFile.sLocation = smFieldValue
    tlDDFField.iOffset = 86
    tlDDFField.iSize = 1
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmDDFFile.sFlags = smFieldValue
    'tmGetDDFFile.sLocation = smFieldValue
    tlDDFField.iOffset = 87
    tlDDFField.iSize = 10
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmDDFFile.sReserved = smFieldValue
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetField                       *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get field from record           *
'*                                                     *
'*******************************************************
Private Sub mGetField(tlDDFField As DDFFIELD)
'
'   ilRecType: 0= tmRec
'
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilSize As Integer
    Dim ilValue As Integer
    Dim llValue As Long
    Dim ilDyMn As Integer
    Dim ilYear As Integer
    Dim ilHsSec As Integer
    Dim ilMinHr As Integer
    Dim slMoney As String
    'Dim slHex As String
    'Dim slInt2 As INT2
    'Dim slInt4 As INT4
    'Used with ePDNToStr
    'Dim slResult As String * 20
    'Dim ilPos As Integer
    ilIndex = tlDDFField.iOffset + 1
    ilSize = tlDDFField.iSize
    smFieldValue = "Error"
    Select Case Asc(tlDDFField.sDataType)
        Case 0  'String
            smFieldValue = ""
            For ilLoop = 1 To ilSize Step 1
                'If Asc(tmRec.sChar(ilIndex)) = 0 Then
                If (tmRec.sChar(ilIndex)) = 0 Then
                    smFieldValue = smFieldValue & " "
                Else
                    'If Asc(tmRec.sChar(ilIndex)) > 127 Then
                    '    smFieldValue = smFieldValue & Chr(Asc(tmRec.sChar(ilIndex)) - 128)
                    'Else
                        smFieldValue = smFieldValue & Chr(tmRec.sChar(ilIndex))
                    'End If
                End If
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 1  'Integer
            If ilSize = 1 Then
                ilValue = tmRec.sChar(ilIndex) 'Asc(tmRec.sChar(ilIndex))
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 2 Then
                ''If ilRecType = 0 Then
                ''    slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
                ''Else
                ''    slInt2.sInt2 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1)
                ''End If
                ''HMemCpy ilValue, slInt2, 2
                '''ilValue = (256 * CLng(Asc(tmRec.sChar(ilIndex + 1))) + Asc(tmRec.sChar(ilIndex))) And &HFFFF&
                'ilValue = 256 * (tmRec.sChar(ilIndex + 1)) + (tmRec.sChar(ilIndex))
                HMemCpy ilValue, tmRec.sChar(ilIndex), 2
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 4 Then
                ''If ilRecType = 0 Then
                ''    slInt4.sInt4 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1) & tmRec.sChar(ilIndex + 2) & tmRec.sChar(ilIndex + 3)
                ''Else
                ''    slInt4.sInt4 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1) & tgKeyRec.sChar(ilIndex + 2) & tgKeyRec.sChar(ilIndex + 3)
                ''End If
                ''HMemCpy llValue, slInt4, 4
                '''llValue = (256 * CLng(Asc(tmRec.sChar(ilIndex + 3))) + Asc(tmRec.sChar(ilIndex + 2))) And &HFFFF&
                '''llValue = 32768 * llValue + (256 * CLng(Asc(tmRec.sChar(ilIndex + 1))) + Asc(tmRec.sChar(ilIndex))) And &HFFFF&
                'llValue = 256 * CLng(tmRec.sChar(ilIndex + 3)) + tmRec.sChar(ilIndex + 2)
                'llValue = 65536 * llValue + 256 * CLng(tmRec.sChar(ilIndex + 1)) + tmRec.sChar(ilIndex)
                HMemCpy llValue, tmRec.sChar(ilIndex), 4
                smFieldValue = Trim$(Str$(llValue))
            End If
        Case 2  'IEEE Float
        Case 3  'Btrieve Date
            'ilDyMn = 256 * Asc(tmRec.sChar(ilIndex + 1)) + Asc(tmRec.sChar(ilIndex))
            'ilYear = 256 * Asc(tmRec.sChar(ilIndex + 3)) + Asc(tmRec.sChar(ilIndex + 2))
            ilDyMn = 256 * (tmRec.sChar(ilIndex + 1)) + (tmRec.sChar(ilIndex))
            ilYear = 256 * (tmRec.sChar(ilIndex + 3)) + (tmRec.sChar(ilIndex + 2))
            gUnpackDate ilDyMn, ilYear, smFieldValue
        Case 4  'Btrieve Time
            'ilHsSec = 256 * Asc(tmRec.sChar(ilIndex + 1)) + Asc(tmRec.sChar(ilIndex))
            'ilMinHr = 256 * Asc(tmRec.sChar(ilIndex + 3)) + Asc(tmRec.sChar(ilIndex + 2))
            ilHsSec = 256 * (tmRec.sChar(ilIndex + 1)) + (tmRec.sChar(ilIndex))
            ilMinHr = 256 * (tmRec.sChar(ilIndex + 3)) + (tmRec.sChar(ilIndex + 2))
            gUnpackTime ilHsSec, ilMinHr, "A", "1", smFieldValue
        Case 5  'COBOL Decimal COMP-3
            slMoney = ""
            For ilLoop = 1 To ilSize Step 1
                'slMoney = slMoney & tmRec.sChar(ilIndex)
                slMoney = slMoney & Chr(tmRec.sChar(ilIndex))
                ilIndex = ilIndex + 1
            Next ilLoop
            'ePDNToStr slMoney, ilSize, Asc(tlDDFField.sDec), slResult
            'ilPos = InStr(slResult, Chr(0))
            'If ilPos > 0 Then
            '    smFieldValue = Left$(slResult, ilPos - 1)
            'Else
            '    smFieldValue = Trim$(slResult)
            'End If
            gPDNToStr slMoney, Asc(tlDDFField.sDec), smFieldValue
        Case 6  'COBOL Money
            slMoney = ""
            For ilLoop = 1 To ilSize Step 1
                'slMoney = slMoney & tmRec.sChar(ilIndex)
                slMoney = slMoney & Chr(tmRec.sChar(ilIndex))
                ilIndex = ilIndex + 1
            Next ilLoop
            'ePDNToStr slMoney, ilSize, Asc(tlDDFField.sDec), slResult
            'ilPos = InStr(slResult, Chr(0))
            'If ilPos > 0 Then
            '    smFieldValue = Left$(slResult, ilPos - 1)
            'Else
            '    smFieldValue = Trim$(slResult)
            'End If
            gPDNToStr slMoney, Asc(tlDDFField.sDec), smFieldValue
        Case 7  'Logical
            If ilSize = 1 Then
                'ilValue = Asc(tmRec.sChar(ilIndex))
                ilValue = (tmRec.sChar(ilIndex))
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 2 Then
                ''If ilRecType = 0 Then
                ''    slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
                ''Else
                ''    slInt2.sInt2 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1)
                ''End If
                ''HMemCpy ilValue, slInt2, 2
                '''ilValue = (256 * CLng(Asc(tmRec.sChar(ilIndex + 1))) + Asc(tmRec.sChar(ilIndex))) And &HFFFF&
                'ilValue = (256 * (tmRec.sChar(ilIndex + 1))) + (tmRec.sChar(ilIndex))
                HMemCpy ilValue, tmRec.sChar(ilIndex), 2
                smFieldValue = Trim$(Str$(ilValue))
            End If
        Case 8  'COBOL Numeric
        Case 9  'BASIC Float
        Case 10 'Pascal LString
            'ilValue = Asc(tmRec.sChar(ilIndex))
            ilValue = tmRec.sChar(ilIndex)
            smFieldValue = ""
            ilIndex = ilIndex + 1
            For ilLoop = 1 To ilValue Step 1
                'If Asc(tmRec.sChar(ilIndex)) = 0 Then
                If (tmRec.sChar(ilIndex)) = 0 Then
                    smFieldValue = smFieldValue & " "
                Else
                    'smFieldValue = smFieldValue & tmRec.sChar(ilIndex)
                    smFieldValue = smFieldValue & Chr(tmRec.sChar(ilIndex))
                End If
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 11 'C ZString
            smFieldValue = ""
            'Do While Asc(tmRec.sChar(ilIndex)) <> 0
            '    smFieldValue = smFieldValue & tmRec.sChar(ilIndex)
            '    ilIndex = ilIndex + 1
            'Loop
            Do While (tmRec.sChar(ilIndex)) <> 0
                smFieldValue = smFieldValue & Chr(tmRec.sChar(ilIndex))
                ilIndex = ilIndex + 1
            Loop
        Case 12 'Variable Length Note
        Case 13 'LVar (length, Characters, 0, 0); Length does not include zeros
            ''If ilRecType = 0 Then
            ''    slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
            ''Else
            ''    slInt2.sInt2 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1)
            ''End If
            ''HMemCpy ilValue, slInt2, 2
            'ilValue = (256 * (tmRec.sChar(ilIndex + 1))) + (tmRec.sChar(ilIndex))
            HMemCpy ilValue, tmRec.sChar(ilIndex), 2
            smFieldValue = ""
            ilIndex = ilIndex + 2
            For ilLoop = 1 To ilValue Step 1
                'If Asc(tmRec.sChar(ilIndex)) = 0 Then
                '    smFieldValue = smFieldValue & " "
                'Else
                '    smFieldValue = smFieldValue & tmRec.sChar(ilIndex)
                'End If
                If (tmRec.sChar(ilIndex)) = 0 Then
                    smFieldValue = smFieldValue & " "
                Else
                    smFieldValue = smFieldValue & Chr(tmRec.sChar(ilIndex))
                End If
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 14 'Unsigned Binary
            ''If ilRecType = 0 Then
            ''    slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
            ''Else
            ''    slInt2.sInt2 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1)
            ''End If
            ''HMemCpy ilValue, slInt2, 2
            'ilValue = (256 * (tmRec.sChar(ilIndex + 1))) + (tmRec.sChar(ilIndex))
            HMemCpy ilValue, tmRec.sChar(ilIndex), 2
            llValue = gUnsignIntToLong(ilValue)
            smFieldValue = Trim$(Str$(llValue))
        Case 15 'AutoIncrement
            If ilSize = 2 Then
                ''If ilRecType = 0 Then
                ''    slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
                ''Else
                ''    slInt2.sInt2 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1)
                ''End If
                ''HMemCpy ilValue, slInt2, 2
                '''ilValue = (256 * CLng(Asc(tmRec.sChar(ilIndex + 1))) + Asc(tmRec.sChar(ilIndex))) And &HFFFF&
                'ilValue = (256 * (tmRec.sChar(ilIndex + 1))) + (tmRec.sChar(ilIndex))
                HMemCpy ilValue, tmRec.sChar(ilIndex), 2
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 4 Then
                ''If ilRecType = 0 Then
                ''    slInt4.sInt4 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1) & tmRec.sChar(ilIndex + 2) & tmRec.sChar(ilIndex + 3)
                ''Else
                ''    slInt4.sInt4 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1) & tgKeyRec.sChar(ilIndex + 2) & tgKeyRec.sChar(ilIndex + 3)
                ''End If
                ''HMemCpy llValue, slInt4, 4
                '''llValue = (256 * CLng(Asc(tmRec.sChar(ilIndex + 3))) + Asc(tmRec.sChar(ilIndex + 2))) And &HFFFF&
                '''llValue = 32768 * llValue + (256 * CLng(Asc(tmRec.sChar(ilIndex + 1))) + Asc(tmRec.sChar(ilIndex))) And &HFFFF&
                'llValue = 256 * CLng(tmRec.sChar(ilIndex + 3)) + tmRec.sChar(ilIndex + 2)
                'llValue = 65536 * llValue + 256 * CLng(tmRec.sChar(ilIndex + 1)) + tmRec.sChar(ilIndex)
                HMemCpy llValue, tmRec.sChar(ilIndex), 4
                smFieldValue = Trim$(Str$(llValue))
            Else
                smFieldValue = "Error"
            End If
        Case 16 'Bit
        Case 17 'COBOL Numeric STS
    End Select
End Sub
Private Sub mEraseOldFiles(slSourcePath As String, slSource As String, slDestination As String)
    Dim slFile As String
    Dim slPath As String
    Dim slToName As String

    slToName = Mid$(slSourcePath, Len(slSource) + 1)
    slPath = slDestination & smDProd & Mid$(slToName, Len(smSProd) + 2) 'write test/production
    On Error Resume Next
    slFile = Dir(slPath & "*.*")
    Do While slFile <> ""
        Kill slPath & slFile
        Print #hmMsg, "     Deleted old file " & slPath & slFile
        slFile = Dir
    Loop

End Sub
Private Function mFolderToClean(smModFilePath As String) As Boolean
    Dim slFolders(0 To 0) As String
    Dim c As Integer
    Dim ilTotal As Integer
    slFolders(0) = "Setup"
    ilTotal = UBound(slFolders)
    For c = 0 To ilTotal
        If InStr(1, smModFilePath, slFolders(c), vbTextCompare) > 0 Then
            mFolderToClean = True
            Exit Function
        End If
    Next c
    mFolderToClean = False
End Function

