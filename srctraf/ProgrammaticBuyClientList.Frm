VERSION 5.00
Begin VB.Form ProgrammaticBuyClientList 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Client List"
   ClientHeight    =   5355
   ClientLeft      =   465
   ClientTop       =   2415
   ClientWidth     =   4830
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5355
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmcExit 
      Appearance      =   0  'Flat
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2415
      TabIndex        =   4
      Top             =   4860
      Width           =   1050
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
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   1770
      Width           =   75
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      ScaleHeight     =   240
      ScaleWidth      =   5010
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -30
      Width           =   5010
   End
   Begin VB.CommandButton cmcSignOn 
      Appearance      =   0  'Flat
      Caption         =   "&Sign On"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   705
      TabIndex        =   3
      Top             =   4860
      Width           =   1305
   End
   Begin VB.PictureBox plcProgrammaticBuyClientList 
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
      Height          =   4530
      Left            =   105
      ScaleHeight     =   4470
      ScaleWidth      =   4440
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   225
      Width           =   4500
      Begin VB.ListBox lbcClientList 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4305
         ItemData        =   "ProgrammaticBuyClientList.frx":0000
         Left            =   105
         List            =   "ProgrammaticBuyClientList.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   90
         Width           =   4215
      End
   End
End
Attribute VB_Name = "ProgrammaticBuyClientList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of ProgrammaticBuyClientList.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ProgrammaticBuyClientList.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim smPassCommands As String
Dim smScreenCaption As String

Dim hmSaf As Integer
Dim tmSaf As SAF            'Schedule Attributes record image
Dim imSafRecLen As Integer

Dim imIgnoreChg As Integer
Dim smCommandKeys As String
Dim bmInModalModule As Boolean
Dim smImportName As String
Dim bmAutoMode As Boolean







Private Sub cmcExit_Click()
    sgSpecialPassword = ""
    igExitTraffic = True
    Unload ProgrammaticBuyClientList
    Set ProgrammaticBuyClientList = Nothing   'Remove data segment
End Sub
'
'
'               3-10-00 Create duplicate of rptselct , named rptselcb to separate the "bridge" reports
'                       Send appropirate command line string
'
Private Sub cmcSignOn_Click()
    Dim ilShell As Integer
    Dim ilRet As Integer
    Dim slStr As String
    
    On Error Resume Next
    If lbcClientList.ListIndex < 0 Then
        Exit Sub
    End If
    sgClientName = lbcClientList.List(lbcClientList.ListIndex)
    igExitTraffic = False
    Unload ProgrammaticBuyClientList
    ProgNetComm.Show vbModal
    
    
    Set ProgrammaticBuyClientList = Nothing   'Remove data segment
End Sub


Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    'Me.KeyPreview = False
End Sub

Private Sub Form_GotFocus()
    Me.WindowState = vbNormal
End Sub



Private Sub Form_Load()
    If App.PrevInstance Then
        End
    Else
        mInit
    End If
End Sub



Private Sub lbcClientList_DblClick()
    cmcSignOn_Click
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
    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    pbcClickFocus.Left = -pbcClickFocus.Width - 30
    'ProgrammaticBuyClientList.Height = cmcSignOn.Top + 3 * cmcSignOn.Height '/ 3
    gCenterStdAlone ProgrammaticBuyClientList
    Screen.MousePointer = vbHourglass
    smScreenCaption = "Client List Selection"
    'imcHelp.Picture = Traffic!imcHelp.Picture
    mPopulate
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate tree structure list   *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
    lbcClientList.Clear
    lbcClientList.AddItem "AdLarge Media, LLC"
    lbcClientList.AddItem "AgriNet"
    lbcClientList.AddItem "Aire Radio Networks"
    lbcClientList.AddItem "Anthem Broadcasting"
    lbcClientList.AddItem "AURN"
    lbcClientList.AddItem "Clear Channel Networks - AL"
    lbcClientList.AddItem "Clear Channel Networks - FL"
    lbcClientList.AddItem "Compass Media Marketing"
    lbcClientList.AddItem "Crystal Media"
    lbcClientList.AddItem "Curtis Networks Group"
    lbcClientList.AddItem "Dgital Media"
    lbcClientList.AddItem "Entravision Solutions"
    lbcClientList.AddItem "ESPN, Inc."
    lbcClientList.AddItem "FarmJournal Broadcasting"
    lbcClientList.AddItem "Focus 360 - National Focus"
    lbcClientList.AddItem "GoMedia"
    lbcClientList.AddItem "GOW Media"
    lbcClientList.AddItem "Hispanic Radio Network"
    lbcClientList.AddItem "Hoosier Ag Today"
    lbcClientList.AddItem "iHeart Networks"
    lbcClientList.AddItem "Illinois Radio Network"
    lbcClientList.AddItem "KLOVE/Air1"
    lbcClientList.AddItem "Learfield Communications"
    lbcClientList.AddItem "MetroNews Radio"
    lbcClientList.AddItem "MN Twins/Twins Ballpark"
    lbcClientList.AddItem "Network Indiana"
    lbcClientList.AddItem "Ohio Ag Net"
    lbcClientList.AddItem "Performance Racing"
    lbcClientList.AddItem "PodcastOne Sales"
    lbcClientList.AddItem "Radio America"
    lbcClientList.AddItem "Radio Oklahoma Network"
    lbcClientList.AddItem "Radio PA"
    lbcClientList.AddItem "Reach Media, Inc."
    lbcClientList.AddItem "RFD Radio Network"
    lbcClientList.AddItem "Salem Radio Network"
    lbcClientList.AddItem "Sports Byline USA"
    lbcClientList.AddItem "Sun Broadcast Group, Inc."
    lbcClientList.AddItem "Tennesse Radio Network"
    lbcClientList.AddItem "Three O 'Clock in the Morning"
    lbcClientList.AddItem "United Stations Radio Net"
    lbcClientList.AddItem "Univision"
    lbcClientList.AddItem "Westwood One"

    
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


Private Sub pbcClickFocus_GotFocus()
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smScreenCaption
End Sub



