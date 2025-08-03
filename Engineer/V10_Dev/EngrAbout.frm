VERSION 5.00
Begin VB.Form EngrAbout 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5085
   ClientLeft      =   360
   ClientTop       =   3330
   ClientWidth     =   6855
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
   Icon            =   "EngrAbout.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5085
   ScaleWidth      =   6855
   Visible         =   0   'False
   Begin VB.PictureBox plcAbout 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5070
      Left            =   15
      Picture         =   "EngrAbout.frx":030A
      ScaleHeight     =   5070
      ScaleWidth      =   6735
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   -15
      Width           =   6735
      Begin VB.ListBox lbcInfo 
         Appearance      =   0  'Flat
         Height          =   870
         Index           =   1
         Left            =   300
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   3360
         Width           =   6045
      End
      Begin VB.ListBox lbcInfo 
         Appearance      =   0  'Flat
         Height          =   870
         Index           =   0
         Left            =   300
         TabIndex        =   0
         Top             =   2235
         Width           =   6045
      End
      Begin VB.CommandButton cmcOk 
         Appearance      =   0  'Flat
         Caption         =   "&Ok"
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
         Left            =   3000
         TabIndex        =   1
         Top             =   4545
         Width           =   1125
      End
      Begin VB.PictureBox plcName 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   855
         ScaleHeight     =   675
         ScaleWidth      =   4860
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   690
         Width           =   4920
         Begin VB.Label lacMsg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Engineering System Version xx.xxx"
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
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   4815
         End
         Begin VB.Label lacMsg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Copyright© 1993-2002 Counterpoint Software ®"
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
            TabIndex        =   10
            Top             =   225
            Width           =   4800
         End
         Begin VB.Label lacMsg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "All Rights Reserved"
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
            Left            =   0
            TabIndex        =   9
            Top             =   435
            Width           =   4770
         End
      End
      Begin VB.PictureBox plcStation 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   630
         Left            =   870
         ScaleHeight     =   570
         ScaleWidth      =   4860
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1470
         Width           =   4920
         Begin VB.Label lacMsg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "This product is Licensed to:"
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
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   4815
         End
         Begin VB.Label lacMsg 
            Alignment       =   2  'Center
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
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   4
            Left            =   15
            TabIndex        =   7
            Top             =   210
            Width           =   4800
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
         Left            =   -45
         ScaleHeight     =   165
         ScaleWidth      =   105
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2610
         Width           =   105
      End
   End
End
Attribute VB_Name = "EngrAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: EngrAbout.Frm
'
' Release: 1.0
'
' Description:
'   This file contains About code
Option Explicit
Option Compare Text
'Program library dates Field Areas
'General Files
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imShowingCSI As Integer
'Dates
Dim smNowDate As String
Dim lmNowDate As Long
Dim smLastInvDate As String 'Last Standard Month invoice date
Dim lmLastInvDate As Long
Dim imFirstActivate As Integer

Private Sub cmcOk_Click()
    mTerminate
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Exit Sub
    End If
    imFirstActivate = False
    EngrAbout.Refresh
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
        End
    End If
    mInit
    If imTerminate Then
        mTerminate
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
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim ilPos As Integer
    Dim slWinDir As String
    Dim llWinLen As Long
    Dim slWindowsDir As String * MAX_PATH
    
    imTerminate = False
    imFirstActivate = True
    
    EngrAbout.Height = plcAbout.Height + fgPanelAdj
    EngrAbout.Width = plcAbout.Width + fgPanelAdj
    gCenterForm EngrAbout
    Screen.MousePointer = vbHourglass
    lacMsg(0).Caption = "Engineering System " & sgEngrVersion '& ", O.S. # " & fgWinVersion
    imFirstFocus = True
    llWinLen = GetWindowsDirectory(slWindowsDir, MAX_PATH)
    slWinDir = Left$(slWindowsDir, llWinLen)
    lbcInfo(1).AddItem "Base Windows Folder: " & slWinDir, lbcInfo(1).ListCount
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Dim ilRet As Integer
    Screen.MousePointer = vbDefault
    igManUnload = vbYes
    Unload EngrAbout
    Set EngrAbout = Nothing   'Remove data segment
    igManUnload = vbNo
End Sub
Private Sub plcAbout_Paint()
    plcAbout.CurrentX = 0
    plcAbout.CurrentY = 0
    plcAbout.Print "Counterpoint Software ®"
End Sub
