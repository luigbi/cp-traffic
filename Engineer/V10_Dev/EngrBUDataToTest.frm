VERSION 5.00
Begin VB.Form EngrBUDataToTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Counterpoint Backup Production Data to Test Data"
   ClientHeight    =   3045
   ClientLeft      =   2745
   ClientTop       =   435
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox edcLastCopyDateTime 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   360
      Width           =   3495
   End
   Begin VB.TextBox edcCopyStatus 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton cmcStart 
      Caption         =   "Start"
      Height          =   435
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   2280
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   4800
      TabIndex        =   0
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Label lblCopyRunningMsg 
      Alignment       =   2  'Center
      Caption         =   "Copy is now running on the server. Close this window to continue working or perform a PC shutdown."
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Current copy state :"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Last copy occurred on :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "EngrBUDataToTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of EngrBUDataToTest.frm on Wed 6/17/09 @ 12:
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  smCopyDataToTest                                                                      *
'******************************************************************************************

Option Explicit
Dim smTestDataPath As String

'*********************************************************************************
'
'*********************************************************************************
Private Sub cmcStart_Click()
    Dim slDateTime As String
    Dim ilRet As Integer
    Dim slINIPathName As String

    On Error GoTo ErrHand

    igBUerror = False
    igZipCancel = False
    If Not mGetTestIniInfo Then
        gOpenBUMsgFile "Database is not properly defined under [TestLocations] in the Engineer.ini file."
        igBUerror = True
        Exit Sub
    End If

    If Not gContDBPathResolved Then
        gOpenBUMsgFile "ServerDatabase is not defined in Engineer.ini file." & Chr(10) & Chr(13) _
                       & "The database path must be exactly as if sitting at the server." _
                       & Chr(10) & Chr(13) & "Example: ServerDatabase = C:\Csi\Prod\Data"
        igBUerror = True
        Exit Sub
    End If

    cmcStart.Enabled = False
    ' Check to see if someone started a backup while we were on this screen.
    If gIsBackupRunning() Then
        edcCopyStatus.text = "Someone else started a copy and it is currently in progress."
        Exit Sub
    End If

    slINIPathName = sgIniPathFileName   'sgDBPath & "Engineer.ini"
    ilRet = csiStartCopyDataToTest(sgDBPath, slINIPathName, 0)
    If ilRet <> 0 Then
        MsgBox "Backup is not available. Error code = " & ilRet
        Exit Sub
    End If
    lblCopyRunningMsg.Visible = True
    ' Set up a timer to monitor the status every 3 seconds
    Timer1.Interval = 1000
    Timer1.Enabled = True
    Exit Sub

    slDateTime = " " & Format$(gNow(), "ddd, m/d/yy h:mm AM/PM")
    EngrMain.mnuDate.Caption = slDateTime & "                                                   "

    Exit Sub

ErrHand:
    MsgBox "A general error has occured in cmcStart_Click"
End Sub

'*********************************************************************************
'
'*********************************************************************************
Private Function mGetTestIniInfo() As Integer

    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim slFileName As String
    Dim slReturn As String * 130
    Dim slLocation As String

    slLocation = "TestLocations"
    On Error GoTo gObtainIniValuesErr
    'If igDirectCall = -1 Then
    '    slFileName = sgIniPath & "Engineer.Ini"
    'Else
    '    slFileName = CurDir$ & "\Engineer.Ini"
    'End If
    slFileName = sgIniPathFileName
    ilRet = GetPrivateProfileString("TestLocations", "DBPath", "Not Found", slReturn, 128, slFileName)
    If Left$(slReturn, ilRet) = "Not Found" Then
        'Don't let them backup live data to testdata
        MsgBox "Database is not properly defined under [TestLocations] in the Engineer.ini file."
        mGetTestIniInfo = False
        Exit Function
    End If
    smTestDataPath = Left$(slReturn, ilRet)
    smTestDataPath = smTestDataPath & "\"
    mGetTestIniInfo = True
    Exit Function

gObtainIniValuesErr:
    ilFound = False
    Resume Next

End Function

'*********************************************************************************
'
'*********************************************************************************
Private Sub cmdCancel_Click()
    Unload Me
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


    On Error GoTo ErrHand
    lblCopyRunningMsg.Visible = False
    edcCopyStatus.text = ""
    edcLastCopyDateTime.text = gGetLastCopyDateTime()

    If gIsBackupRunning() Then
        edcCopyStatus.text = "A copy is currently in progress."
        cmcStart.Enabled = False
    Else
        edcCopyStatus.text = "Ok to start a copy now."
    End If
    Exit Sub

ErrHand:
    MsgBox "A general error has occured in Form_Load."
End Sub

'*********************************************************************************
'
'*********************************************************************************
Private Sub Timer1_Timer()
    Dim slLastCopyDateTime As String
    Dim smCSIServerINIFile As String
    Dim slStatus As String

    smCSIServerINIFile = sgExeDirectory & "CSI_Server.ini"
    Call gLoadINIValue(smCSIServerINIFile, "MainSettings", "LastBackupStatus", slStatus)
    edcCopyStatus.text = slStatus
    If Not gIsBackupRunning() Then
        Timer1.Enabled = False
        slLastCopyDateTime = gGetLastCopyDateTime()
        edcLastCopyDateTime.text = slLastCopyDateTime & " EST"
        cmcStart.Enabled = False
        Call gLoadINIValue(smCSIServerINIFile, "MainSettings", "LastBackupStatus", slStatus)
        edcCopyStatus.text = slStatus
        cmdCancel.Caption = "Done"
        lblCopyRunningMsg.Visible = False
    End If
End Sub
