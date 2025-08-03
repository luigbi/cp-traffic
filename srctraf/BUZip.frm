VERSION 5.00
Begin VB.Form BUZip 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Counterpoint Backup"
   ClientHeight    =   3165
   ClientLeft      =   2745
   ClientTop       =   3435
   ClientWidth     =   6285
   ControlBox      =   0   'False
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox edcLastBackupDateTime 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton cmcStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox edcBackupStatus 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Left            =   1680
      Top             =   2400
   End
   Begin VB.CommandButton cmcClose 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblBURunningMsg 
      Alignment       =   2  'Center
      Caption         =   "Backup is now running on the server. Close this window to continue working or perform a PC shutdown."
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Current backup state :"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Last backup occurred on :"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "BUZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of BUZip.frm on Wed 6/17/09 @ 12:56 PM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit

'*********************************************************************************
'
'*********************************************************************************
Private Sub cmcStart_Click()
    Dim ilRet As Integer
    Dim slTrafficINIPathName As String

    On Error GoTo ErrHand
    cmcStart.Enabled = False
    ' Check to see if someone started a backup while we were on this screen.
    If gIsBackupRunning() Then
        edcBackupStatus.Text = "Someone else started a backup and it is currently in progress."
        Exit Sub
    End If

    slTrafficINIPathName = sgDBPath & "Traffic.ini"
    ilRet = csiStartBackup(sgDBPath, slTrafficINIPathName, 3)
    If ilRet <> 0 Then
        If ilRet = 8001 Then
            MsgBox "CSI_Server is not running. Backup is not available."
            Exit Sub
        End If
        MsgBox "Backup is not available. Error code = " & ilRet
        Exit Sub
    End If
    lblBURunningMsg.Visible = True
    ' Set up a timer to monitor the status every 3 seconds
    Timer1.Interval = 1000
    Timer1.Enabled = True
    Exit Sub

ErrHand:
    MsgBox "A general error has occured in cmcStart_Click"
End Sub

'*********************************************************************************
'
'*********************************************************************************
Private Sub Form_Load()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

    Dim slLastBackupDateTime As String

    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    lblBURunningMsg.Visible = False
    slLastBackupDateTime = gGetLastBackupDateTime()
    edcLastBackupDateTime.Text = slLastBackupDateTime & " EST"

    If gIsBackupRunning() Then
        edcBackupStatus.Text = "A backup is currently in progress."
        cmcStart.Enabled = False
    Else
        edcBackupStatus.Text = "Ok to start a backup now."
    End If
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHand:
    MsgBox "A general error has occured in Form_Load."
End Sub

'*********************************************************************************
'
'*********************************************************************************
Private Sub Timer1_Timer()
    Dim slLastBackupDateTime As String
    Dim smCSIServerINIFile As String
    Dim slStatus As String

    smCSIServerINIFile = sgExePath & "\CSI_Server.ini"
    Call gLoadINIValue(smCSIServerINIFile, "MainSettings", "LastBackupStatus", slStatus)
    edcBackupStatus.Text = slStatus
    If Not gIsBackupRunning() Then
        Timer1.Enabled = False
        slLastBackupDateTime = gGetLastBackupDateTime()
        edcLastBackupDateTime.Text = slLastBackupDateTime & " EST"
        Call gLoadINIValue(smCSIServerINIFile, "MainSettings", "LastBackupStatus", slStatus)
        edcBackupStatus.Text = slStatus
        cmcClose.Caption = "Done"
        lblBURunningMsg.Visible = False
    End If
End Sub

'*********************************************************************************
'
'*********************************************************************************
Private Sub cmcClose_Click()
    Unload Me
End Sub

