VERSION 5.00
Begin VB.Form EngrList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3450
   ClientLeft      =   135
   ClientTop       =   135
   ClientWidth     =   7965
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
   ScaleHeight     =   3450
   ScaleWidth      =   7965
   Visible         =   0   'False
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   7350
      Top             =   960
   End
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
      Height          =   3435
      Left            =   0
      ScaleHeight     =   3435
      ScaleWidth      =   7965
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7965
      Begin VB.PictureBox pbcHor 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   3255
         Left            =   90
         Picture         =   "EngrList.frx":0000
         ScaleHeight     =   3195
         ScaleWidth      =   7740
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   90
         Width           =   7800
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Bus Controls"
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
            Index           =   17
            Left            =   3255
            TabIndex        =   39
            Top             =   1065
            Width           =   1515
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Audio Controls"
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
            Index           =   16
            Left            =   105
            TabIndex        =   38
            Top             =   555
            Width           =   1515
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Bus Groups"
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
            Index           =   15
            Left            =   4830
            TabIndex        =   11
            Top             =   1065
            Width           =   1515
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Audio Names"
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
            Index           =   14
            Left            =   1680
            TabIndex        =   1
            Top             =   555
            Width           =   1515
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Audio Types"
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
            Index           =   13
            Left            =   4830
            TabIndex        =   5
            Top             =   555
            Width           =   1515
         End
         Begin VB.Label lacPressMsg 
            BackStyle       =   0  'Transparent
            Caption         =   "Press F2 or Click Logo for Jobs"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   4050
            TabIndex        =   37
            Top             =   2940
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Material Types"
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
            Left            =   4830
            TabIndex        =   19
            Top             =   1575
            Width           =   1515
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
            Left            =   7110
            TabIndex        =   36
            Top             =   30
            Width           =   570
         End
         Begin VB.Label imcOutline 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   435
            Left            =   6675
            TabIndex        =   35
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "User Options"
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
            Index           =   12
            Left            =   3255
            TabIndex        =   31
            Top             =   2595
            Width           =   1515
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Site Options"
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
            Index           =   11
            Left            =   1680
            TabIndex        =   29
            Top             =   2595
            Width           =   1515
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Silences"
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
            Left            =   3255
            TabIndex        =   25
            Top             =   2085
            Width           =   1515
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Comments"
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
            Left            =   105
            TabIndex        =   13
            Top             =   1575
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Netcues"
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
            Left            =   105
            TabIndex        =   21
            Top             =   2085
            Width           =   1515
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Follows"
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
            Left            =   3255
            TabIndex        =   17
            Top             =   1575
            Width           =   1515
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Relays"
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
            Left            =   1680
            TabIndex        =   23
            Top             =   2085
            Width           =   1515
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Audio Sources"
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
            Left            =   3255
            TabIndex        =   3
            Top             =   555
            Width           =   1515
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Time Types"
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
            Left            =   4830
            TabIndex        =   27
            Top             =   2085
            Width           =   1515
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
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
            Index           =   2
            Left            =   1680
            TabIndex        =   15
            Top             =   1575
            Width           =   1515
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Buses"
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
            Left            =   1680
            TabIndex        =   9
            Top             =   1065
            Width           =   1515
         End
         Begin VB.Label lacTask 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Automation"
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
            Left            =   120
            TabIndex        =   7
            Top             =   1065
            Width           =   1515
         End
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Bus Groups"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   15
         Left            =   5040
         TabIndex        =   12
         Tag             =   "List of Advertising Agencies"
         Top             =   1185
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Audio Names"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   14
         Left            =   1890
         TabIndex        =   2
         Tag             =   "List of Product protection codes"
         Top             =   675
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Audio Types"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   13
         Left            =   5040
         TabIndex        =   6
         Tag             =   "List of Product protection codes"
         Top             =   675
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Follows"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   7
         Left            =   3465
         TabIndex        =   18
         Tag             =   "List of all stations, networks, channels, or other airing entities"
         Top             =   1695
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Netcues"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   315
         TabIndex        =   22
         Tag             =   "List of all stations, networks, channels, or other airing entities"
         Top             =   2190
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Buses"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1890
         TabIndex        =   10
         Tag             =   "List of Advertising Agencies"
         Top             =   1185
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Automation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   315
         TabIndex        =   8
         Tag             =   "List of Advertisers"
         Top             =   1185
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Audio Sources"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   3465
         TabIndex        =   4
         Tag             =   "List of Product protection codes"
         Top             =   675
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
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
         Height          =   345
         Index           =   2
         Left            =   1890
         TabIndex        =   16
         Tag             =   "List of Product protection codes"
         Top             =   1695
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Time Types"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   5040
         TabIndex        =   28
         Tag             =   "List of Product protection codes"
         Top             =   2205
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Relays"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   1890
         TabIndex        =   24
         Tag             =   "List of Product protection codes"
         Top             =   2190
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "User Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   12
         Left            =   3465
         TabIndex        =   32
         Top             =   2715
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Site Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   11
         Left            =   1890
         TabIndex        =   30
         Top             =   2715
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Silences"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   10
         Left            =   3465
         TabIndex        =   26
         Top             =   2205
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   9
         Left            =   315
         TabIndex        =   14
         Tag             =   "List of special types of revenue totals not ordinarily covered"
         Top             =   1695
         Width           =   1365
      End
      Begin VB.PictureBox pbcClickFocus 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Left            =   165
         ScaleHeight     =   120
         ScaleWidth      =   60
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   4530
         Width           =   60
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Material Types"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   5040
         TabIndex        =   20
         Tag             =   "List of Product protection codes"
         Top             =   1695
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Buses"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   17
         Left            =   3465
         TabIndex        =   41
         Tag             =   "List of Advertising Agencies"
         Top             =   1185
         Width           =   1365
      End
      Begin VB.CommandButton cmcTask 
         Appearance      =   0  'Flat
         Caption         =   "Audio Controls"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   16
         Left            =   315
         TabIndex        =   40
         Tag             =   "List of Product protection codes"
         Top             =   675
         Width           =   1365
      End
      Begin VB.Image cmcStartJob 
         Height          =   1335
         Left            =   6270
         Top             =   1965
         Width           =   1560
      End
      Begin VB.Image cmcCSLogo 
         Height          =   735
         Left            =   195
         Top             =   90
         Width           =   4305
      End
   End
End
Attribute VB_Name = "EngrList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: EngrList.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Traffic input screen code
Option Explicit
Option Compare Text

Dim imTaskIndex As Integer
Dim imPrevTaskIndex As Integer

'Control Index
Const AUTOINDEX = 0
Const BUSINDEX = 1
Const EVENTTYPEINDEX = 2
Const TIMETYPEINDEX = 3
Const MATERIALTYPEINDEX = 4
Const AUDIOINDEX = 5
Const RELAYINDEX = 6
Const FOLLOWINDEX = 7
Const NETCUEINDEX = 8
Const COMMENTINDEX = 9
Const SILENCEINDEX = 10
Const SITEINDEX = 11
Const USERINDEX = 12
Const AUDIOTYPEINDEX = 13
Const AUDIONAMEINDEX = 14
Const BUSGROUPINDEX = 15
Const AUDIOCONTROLINDEX = 16
Const BUSCONTROLINDEX = 17





Private Sub cmcCSLogo_Click()
    Dim ilAns As Integer
    
    ilAns = MsgBox(sgEngrVersion, 0, "Counterpoint Software")
End Sub


Private Sub cmcStartJob_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift Then
        Unload EngrList
    Else
        If igJobVisible Then
            EngrJob.SetFocus
        Else
            EngrJob.Show vbModeless
        End If
    End If
End Sub

Private Sub cmcTask_Click(Index As Integer)
    Dim ilRet As Integer
    If imPrevTaskIndex <> -1 Then
        imPrevTaskIndex = -1
    End If
    If tmcClick.Enabled = False Then
        imTaskIndex = Index
        tmcClick.Enabled = True
    End If
End Sub

Private Sub cmcTask_GotFocus(Index As Integer)
    gCtrlGotFocus cmcTask(Index)
End Sub

Private Sub cmcTask_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    imcOutline.Move cmcTask(Index).Left - 150, cmcTask(Index).Top - 120, cmcTask(Index).Width + 90, cmcTask(Index).Height
    cmcTask(Index).ToolTipText = ""
    If imPrevTaskIndex <> -1 Then
        'lacTask(imPrevTaskIndex).ForeColor = vbBlack
        imPrevTaskIndex = -1
    End If
    Select Case Index
        Case AUTOINDEX
            mSetCtrlProperties Index, AUTOLIST
        Case BUSINDEX
            mSetCtrlProperties Index, BUSLIST
        Case EVENTTYPEINDEX
            mSetCtrlProperties Index, EVENTTYPELIST
        Case TIMETYPEINDEX
            mSetCtrlProperties Index, TIMETYPELIST
        Case MATERIALTYPEINDEX
            mSetCtrlProperties Index, MATERIALTYPELIST
        Case AUDIOINDEX
            mSetCtrlProperties Index, AUDIOLIST
        Case RELAYINDEX
            mSetCtrlProperties Index, RELAYLIST
        Case FOLLOWINDEX
            mSetCtrlProperties Index, FOLLOWLIST
        Case NETCUEINDEX
            mSetCtrlProperties Index, NETCUELIST
        Case COMMENTINDEX
            mSetCtrlProperties Index, COMMENTLIST
        Case SILENCEINDEX
            mSetCtrlProperties Index, SILENCELIST
        Case SITEINDEX
            mSetCtrlProperties Index, SITELIST
        Case USERINDEX
            mSetCtrlProperties Index, USERLIST
        Case AUDIOTYPEINDEX
            mSetCtrlProperties Index, AUDIOTYPELIST
        Case AUDIONAMEINDEX
            mSetCtrlProperties Index, AUDIONAMELIST
        Case BUSGROUPINDEX
            mSetCtrlProperties Index, BUSGROUPLIST
        Case AUDIOCONTROLINDEX
            mSetCtrlProperties Index, AUDIOCONTROLLIST
        Case BUSCONTROLINDEX
            mSetCtrlProperties Index, BUSCONTROLLIST
    End Select
    imcOutline.Visible = True
End Sub

Private Sub Form_Activate()
'    If igGGFlag = 0 Then
'        plcWindow.Enabled = False
'    End If
    mSetControls
    imTaskIndex = 100
    igListVisible = True
'    If igGGFlag = 0 Then
'        'frcGeneral.Enabled = False
'        'frcSales.Enabled = False
'        'frcProg.Enabled = False
'        'frcCopy.Enabled = False
'        'frcAccount.Enabled = False
'        'frcPref.Enabled = False
'        cmcGeneric.Enabled = False
'        lacGeneric.Enabled = False
'        'Note: plcWindow.Enabled = False, causes whole window not to show
'    End If
    Me.ZOrder 0 'Send to front
    pbcHor.Visible = False
    DoEvents
    pbcHor.Visible = True
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Load()
    Dim ilLoop As Integer
    If igJobVisible Then
        Unload EngrJob
    End If
    igListVisible = True
    imTaskIndex = -1
    imPrevTaskIndex = -1
    plcWindow.Left = 0
    plcWindow.Top = 0
    EngrList.Height = plcWindow.Height + fgPanelAdj
    EngrList.Width = plcWindow.Width + fgPanelAdj
    If lgCurrBPP <= 8 Then
        pbcHor.Picture = EngrMain!pbc256List
    End If
    gCenterForm EngrList
End Sub
Private Sub Form_Unload(Cancel As Integer)
    igListVisible = False
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
    
    mSetCtrlEnable AUTOINDEX, AUTOLIST
    mSetCtrlEnable BUSINDEX, BUSLIST
    mSetCtrlEnable EVENTTYPEINDEX, EVENTTYPELIST
    mSetCtrlEnable TIMETYPEINDEX, TIMETYPELIST
    mSetCtrlEnable MATERIALTYPEINDEX, MATERIALTYPELIST
    mSetCtrlEnable AUDIOINDEX, AUDIOLIST
    mSetCtrlEnable RELAYINDEX, RELAYLIST
    mSetCtrlEnable FOLLOWINDEX, FOLLOWLIST
    mSetCtrlEnable NETCUEINDEX, NETCUELIST
    mSetCtrlEnable COMMENTINDEX, COMMENTLIST
    mSetCtrlEnable SILENCEINDEX, SILENCELIST
    mSetCtrlEnable SITEINDEX, SITELIST
    mSetCtrlEnable USERINDEX, USERLIST
    mSetCtrlEnable AUDIOTYPEINDEX, AUDIOTYPELIST
    mSetCtrlEnable AUDIONAMEINDEX, AUDIONAMELIST
    mSetCtrlEnable BUSGROUPINDEX, BUSGROUPLIST
    mSetCtrlEnable AUDIOCONTROLINDEX, AUDIOCONTROLLIST
    mSetCtrlEnable BUSCONTROLINDEX, BUSCONTROLLIST
End Sub

Private Sub pbcClickFocus_GotFocus()
    imcOutline.Visible = False
End Sub

Private Sub plcWindow_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    imcOutline.Visible = False
End Sub

Private Sub tmcClick_Timer()
    Dim ilCountActive As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    tmcClick.Enabled = False
    imcOutline.Visible = False
    igInitCallInfo = 0
    On Error Resume Next
    Select Case imTaskIndex
        Case AUTOINDEX
            EngrAutomation.Show vbModal
            ilRet = gGetTypeOfRecs_AEE_AutoEquip("C", sgCurrAEEStamp, "EngrAutomation-mPopulate", tgCurrAEE())
            ilRet = gGetTypeOfRecs_ACE_AutoContact("C", sgCurrACEStamp, "EngrEventType-mPopulate", tgCurrACE())
            ilRet = gGetTypeOfRecs_ADE_AutoDataFlags("C", sgCurrADEStamp, "EngrEventType-mPopulate", tgCurrADE())
            ilRet = gGetTypeOfRecs_AFE_AutoFormat("C", sgCurrAFEStamp, "EngrEventType-mPopulate", tgCurrAFE())
            ilRet = gGetTypeOfRecs_APE_AutoPath("C", sgCurrAPEStamp, "EngrEventType-mPopulate", tgCurrAPE())
            ilCountActive = 0
            For ilLoop = 0 To UBound(tgCurrAEE) - 1 Step 1
                If tgCurrAEE(ilLoop).sState = "A" Then
                    ilCountActive = ilCountActive + 1
                End If
            Next ilLoop
            If ilCountActive = 1 Then
                LSet tgAEE = tgCurrAEE(0)
                LSet tgACE = tgCurrACE(0)
                LSet tgADE = tgCurrADE(0)
                If tgCurrAFE(0).sSubType = "S" Then
                    LSet tgStartColAFE = tgCurrAFE(0)
                ElseIf tgCurrAFE(1).sSubType = "S" Then
                    LSet tgStartColAFE = tgCurrAFE(1)
                End If
                If tgCurrAFE(0).sSubType = "N" Then
                    LSet tgNoCharAFE = tgCurrAFE(0)
                ElseIf tgCurrAFE(1).sSubType = "N" Then
                    LSet tgNoCharAFE = tgCurrAFE(1)
                End If
                LSet tgAPE = tgCurrAPE(0)
            Else
                gInitMaxAFE
            End If
        Case BUSINDEX
            EngrBus.Show vbModal
        Case EVENTTYPEINDEX
            EngrEventType.Show vbModal
            gGetEPEUsedSummary
            gGetEPEManSummary
        Case TIMETYPEINDEX
            EngrTimeType.Show vbModal
        Case MATERIALTYPEINDEX
            EngrMatType.Show vbModal
        Case AUDIOINDEX
            EngrAudio.Show vbModal
        Case RELAYINDEX
            EngrRelay.Show vbModal
        Case FOLLOWINDEX
            EngrFollow.Show vbModal
        Case NETCUEINDEX
            EngrNetcue.Show vbModal
        'Case COMMENTINDEX
        '    EngrComment.Show vbModal
        Case SILENCEINDEX
            EngrSilence.Show vbModal
        Case SITEINDEX
            EngrSiteOption.Show vbModal
        Case USERINDEX
            EngrUserOption.Show vbModal
        Case AUDIOTYPEINDEX
            EngrAudioType.Show vbModal
        Case AUDIONAMEINDEX
            EngrAudioName.Show vbModal
        Case BUSGROUPINDEX
            EngrBusGroup.Show vbModal
        Case AUDIOCONTROLINDEX
            igInitCallInfo = 1
            EngrControlChar.Show vbModal
        Case BUSCONTROLINDEX
            igInitCallInfo = 2
            EngrControlChar.Show vbModal
    End Select
    imTaskIndex = -1
End Sub

Public Sub mSetCtrlProperties(ilCtrlIndex As Integer, ilTaskNo As Integer)
    If mCtrlUsed(ilTaskNo) Then
        If igListStatus(ilTaskNo) = 0 Then
            cmcTask(ilCtrlIndex).ToolTipText = "Disallowed"
        Else
            If igListStatus(ilTaskNo) = 1 Then
                cmcTask(ilCtrlIndex).ToolTipText = "View Only"
                imPrevTaskIndex = ilCtrlIndex
            Else
                cmcTask(ilCtrlIndex).ToolTipText = "Updates Allowed"
            End If
        End If
    Else
        cmcTask(ilCtrlIndex).ToolTipText = "Disallowed"
    End If
End Sub

Private Sub mSetCtrlEnable(ilCtrlIndex As Integer, ilTaskNo As Integer)
    If mCtrlUsed(ilTaskNo) Then
        If igListStatus(ilTaskNo) = 0 Then
            cmcTask(ilCtrlIndex).Enabled = False
            lacTask(ilCtrlIndex).Enabled = False
        Else
            cmcTask(ilCtrlIndex).Enabled = True
        End If
    Else
        cmcTask(ilCtrlIndex).Enabled = False
        lacTask(ilCtrlIndex).Enabled = False
    End If
End Sub

Private Function mCtrlUsed(ilTaskNo As Integer) As Integer
    Dim ilCtrlUsed As Integer
    ilCtrlUsed = True
    Select Case ilTaskNo
        Case BUSLIST
            If tgUsedSumEPE.sBus <> "Y" Then
                ilCtrlUsed = False
            End If
        Case BUSCONTROLLIST
            If tgUsedSumEPE.sBusControl <> "Y" Then
                ilCtrlUsed = False
            End If
        Case TIMETYPELIST
            If (tgUsedSumEPE.sStartType <> "Y") And (tgUsedSumEPE.sEndType <> "Y") Then
                ilCtrlUsed = False
            End If
        Case MATERIALTYPELIST
            If tgUsedSumEPE.sMaterialType <> "Y" Then
                ilCtrlUsed = False
            End If
        Case RELAYLIST
            If (tgUsedSumEPE.sRelay1 <> "Y") And (tgUsedSumEPE.sRelay2 <> "Y") Then
                ilCtrlUsed = False
            End If
        Case FOLLOWLIST
            If tgUsedSumEPE.sFollow <> "Y" Then
                ilCtrlUsed = False
            End If
        Case NETCUELIST
            If (tgUsedSumEPE.sStartNetcue <> "Y") And (tgUsedSumEPE.sStopNetcue <> "Y") Then
                ilCtrlUsed = False
            End If
        Case SILENCELIST
            If (tgUsedSumEPE.sSilence1 <> "Y") And (tgUsedSumEPE.sSilence2 <> "Y") And (tgUsedSumEPE.sSilence3 <> "Y") And (tgUsedSumEPE.sSilence4 <> "Y") Then
                ilCtrlUsed = False
            End If
        Case AUDIOCONTROLLIST
            If (tgUsedSumEPE.sAudioControl <> "Y") And (tgUsedSumEPE.sBkupAudioControl <> "Y") And (tgUsedSumEPE.sProtAudioControl <> "Y") Then
                ilCtrlUsed = False
            End If
        Case Else
            ilCtrlUsed = True
    End Select
    mCtrlUsed = ilCtrlUsed
End Function
