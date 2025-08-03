VERSION 5.00
Begin VB.Form ProgrammaticBuyHelp 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4845
   ClientLeft      =   2640
   ClientTop       =   2610
   ClientWidth     =   8745
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4845
   ScaleWidth      =   8745
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
      Height          =   105
      Left            =   45
      ScaleHeight     =   105
      ScaleWidth      =   15
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1965
      Width           =   15
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Exit Help"
      Height          =   285
      Left            =   3810
      TabIndex        =   2
      Top             =   4350
      Width           =   1050
   End
   Begin VB.PictureBox plcHelp 
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
      Height          =   3945
      Left            =   105
      ScaleHeight     =   3885
      ScaleWidth      =   8430
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   8490
      Begin VB.TextBox edcMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3705
         Left            =   45
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   90
         Width           =   8325
      End
   End
End
Attribute VB_Name = "ProgrammaticBuyHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ProgrammaticBuyHelp.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract number screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim smMsg As String


Private Sub cmcDone_Click()
    igTerminateReturn = 1
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         slStr                                                   *
'******************************************************************************************


    gCtrlGotFocus cmcDone
End Sub


Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    'If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
    '    gFunctionKeyBranch KeyCode
    'End If
End Sub

Private Sub Form_Load()
    mInit
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    ProgrammaticBuyHelp.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    'ProgrammaticBuyHelp.Left = LiveLog.Left + LiveLog.Width / 2 - ProgrammaticBuyHelp.Width / 2
    'ProgrammaticBuyHelp.Top = LiveLog.Top + LiveLog.Height / 2 - ProgrammaticBuyHelp.Height / 2
    gCenterStdAlone ProgrammaticBuyHelp
    edcMsg.Move plcHelp.Left + 60, plcHelp.Top + 60
    If igHelpScreen = 1 Then
        smMsg = "Double-click the Live Log icon, and the selection window will appear." & sgCR & sgLF
        smMsg = smMsg & "The selection window will show every Log scheduled to be aired today by a live operator." & sgCR & sgLF
        smMsg = smMsg & "Any logs that are marked in red are currently in progress, being monitored by another operator." & sgCR & sgLF
        smMsg = smMsg & "Any logs in green are already aired and posted." & sgCR & sgLF
        smMsg = smMsg & "Select a Log that is shown in black" & sgCR & sgLF
        smMsg = smMsg & "Select the Log and/or Event you want, and press the Continue button to go to the Main screen."
    Else
        smMsg = "Overview" & sgCR & sgLF
        smMsg = smMsg & "1. Sign On.  Enter your name, and optionally, your location." & sgCR & sgLF
        smMsg = smMsg & "2. Enter the time the Event or program starts." & sgCR & sgLF
        smMsg = smMsg & "3. Post the times that each spot airs, and flag spots that don't air." & sgCR & sgLF
        smMsg = smMsg & "4. Post the copy that airs, if not the original copy." & sgCR & sgLF
        smMsg = smMsg & "5. Post any spots that were not scheduled to air."
        smMsg = smMsg & "6. Enter the time the Event or program ends (CRITICAL!)" & sgCR & sgLF
        smMsg = smMsg & "7. Sign Off." & sgCR & sgLF
        smMsg = smMsg & sgCR & sgLF
        smMsg = smMsg & "There are two scrolling windows on the screen.  The top window is everything waiting to be aired.  The bottom window is everything you mark as aired or not aired, and includes the sign on/sign off times, and program or Event start and end times." & sgCR & sgLF
        smMsg = smMsg & sgCR & sgLF
        smMsg = smMsg & "To post any event, click next to it in the green box.  The time right now will appear.  If that is not the time the event started, overtype it.  When you press tab, or click on another event, the event will be removed from the top window, and posted to the bottom window." & sgCR & sgLF
        smMsg = smMsg & sgCR & sgLF
        smMsg = smMsg & "If a spot airs, but the wrong copy runs, change the copy to what aired." & sgCR & sgLF
        smMsg = smMsg & sgCR & sgLF
        smMsg = smMsg & "If a spot does not air, click next to it in the red box.  Then select a reason missed in the box that pops up.  If you don't see the correct reason, type in a new reason.  It will appear in the list the next time you bottom window." & sgCR & sgLF
        smMsg = smMsg & sgCR & sgLF
        smMsg = smMsg & "If you realize that a spot in the bottom area has the wrong time, click on it and overtype the correct time." & sgCR & sgLF
        smMsg = smMsg & sgCR & sgLF
        smMsg = smMsg & "If you post the wrong spot, click on its time in the bottom area, and do one of two things:" & sgCR & sgLF
        smMsg = smMsg & "1. Press the Undo button," & sgCR & sgLF
        smMsg = smMsg & "2. Backspace to clear the time, and press tab." & sgCR & sgLF
        smMsg = smMsg & "Either way, the spot will be sent back to the top area, as though it were never posted.  Use either of these same two techniques to remove the time for a spot in the top window that you clicked on accidentally." & sgCR & sgLF
        smMsg = smMsg & sgCR & sgLF
        smMsg = smMsg & "Live spots in the top window have a blue background, recorded spots a yellow background." & sgCR & sgLF
        smMsg = smMsg & sgCR & sgLF
        smMsg = smMsg & "Aired spots in the bottom window are green, missed spots in the bottom window are red." & sgCR & sgLF
        smMsg = smMsg & sgCR & sgLF
        smMsg = smMsg & "If a spot airs that was not in the top window, press the Add Bonus Spot button to create it." & sgCR & sgLF
        smMsg = smMsg & "Enter its advertiser name, time aired, and copy." & sgCR & sgLF
    End If
    edcMsg.Text = smMsg
    Screen.MousePointer = vbDefault
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
'   mTerminate
'   Where:
'
    Screen.MousePointer = vbDefault
    Unload ProgrammaticBuyHelp
    Set ProgrammaticBuyHelp = Nothing   'Remove data segment
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = HELP_INDEX
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
