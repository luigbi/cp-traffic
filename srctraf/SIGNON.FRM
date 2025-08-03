VERSION 5.00
Begin VB.Form Signon 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3045
   ClientLeft      =   2310
   ClientTop       =   2055
   ClientWidth     =   4425
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3045
   ScaleWidth      =   4425
   Begin VB.Frame frcSignon 
      Height          =   3135
      Left            =   10
      TabIndex        =   0
      Top             =   -80
      Width           =   4380
      Begin VB.PictureBox plcSignon 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2960
         Left            =   30
         ScaleHeight     =   2895
         ScaleWidth      =   4275
         TabIndex        =   1
         Top             =   120
         Width           =   4330
         Begin VB.PictureBox pbcClickFocus 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   -300
            ScaleHeight     =   165
            ScaleWidth      =   105
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   825
            Width           =   105
         End
         Begin VB.TextBox edcPassword 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1605
            PasswordChar    =   "*"
            TabIndex        =   10
            Top             =   1200
            Width           =   2325
         End
         Begin VB.TextBox edcName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1605
            TabIndex        =   9
            Top             =   810
            Width           =   2325
         End
         Begin VB.ListBox lbcUser 
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
            Height          =   240
            Left            =   225
            TabIndex        =   8
            Top             =   1620
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.OptionButton rbcSystemType 
            BackColor       =   &H80000009&
            Caption         =   "Production System"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   7
            Top             =   2820
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.OptionButton rbcSystemType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Test System"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   3090
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.CheckBox ckcRemoteUser 
            Caption         =   "Suppress Background Picture"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   105
            TabIndex        =   5
            Top             =   3420
            Visible         =   0   'False
            Width           =   2865
         End
         Begin VB.TextBox lacCode 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1605
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   1170
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.CommandButton cmcExit 
            Appearance      =   0  'Flat
            Caption         =   "Exit Traffic System"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   960
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   2355
            Width           =   2565
         End
         Begin VB.CommandButton cmcStart 
            Appearance      =   0  'Flat
            Caption         =   "Start Traffic System"
            Default         =   -1  'True
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   960
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   1875
            Width           =   2565
         End
         Begin VB.PictureBox plcName 
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
            Left            =   1605
            ScaleHeight     =   285
            ScaleWidth      =   2265
            TabIndex        =   12
            Top             =   780
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.PictureBox plcPassword 
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
            Height          =   345
            Left            =   1605
            ScaleHeight     =   285
            ScaleWidth      =   2265
            TabIndex        =   11
            Top             =   1170
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.Image cmcCSLogo 
            Height          =   510
            Left            =   45
            Picture         =   "SIGNON.frx":0000
            Stretch         =   -1  'True
            Top             =   45
            Width           =   4035
         End
         Begin VB.Label lacUserName 
            BackStyle       =   0  'Transparent
            Caption         =   "User Name"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   300
            TabIndex        =   20
            Top             =   820
            Width           =   1080
         End
         Begin VB.Label lacPassword 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   300
            TabIndex        =   19
            Top             =   1200
            Width           =   1080
         End
         Begin VB.Label lacProd 
            BackStyle       =   0  'Transparent
            Caption         =   "Production"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   345
            TabIndex        =   18
            Top             =   2775
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lacTest 
            BackStyle       =   0  'Transparent
            Caption         =   "Test"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   345
            TabIndex        =   17
            Top             =   3060
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label lacStart 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   16
            Top             =   1860
            Visible         =   0   'False
            Width           =   2565
         End
         Begin VB.Label lacExit 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   15
            Top             =   2340
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label imcOutline 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   300
            Left            =   225
            TabIndex        =   14
            Top             =   2040
            Visible         =   0   'False
            Width           =   330
         End
      End
   End
End
Attribute VB_Name = "Signon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SIGNON.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software ® All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SignOn.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Sign On input screen code
Option Explicit
Option Compare Text
'Dim tmLPopRec As LPOPREC
Dim imDBSelected As Integer  '0=Production; 1=Test
Dim smSpecial As String
Dim smGGBGSpecial As String




Private Sub cmcCSLogo_Click()
    Dim iAns As Integer
    Dim ilErr As Integer
    ''sgCSVersion set within InitSubs.Bas-gInitGlobalVar
    ''iAns = MsgBox(sgCSVersion, 0, "Counterpoint Software")
    'llShell = 0
    'ilErr = False
    'On Error GoTo mnuAboutHelpErr
    'llShell = Shell(sgExePath & "CSIAbout.Exe CS", 1)
    'If ilErr Then
    '    Exit Sub
    'End If
    ''While GetModuleUsage(ilShell) > 0
    ''    ilRet = DoEvents()
    ''Wend
    imcOutline.Visible = False
    CSIAbout.Show vbModal
    Exit Sub

    ilErr = True
    iAns = MsgBox(sgCSVersion, 0, "Counterpoint Software")
    Resume Next
End Sub



Private Sub cmcExit_Click()
    sgSpecialPassword = ""
    igExitTraffic = True
    Unload Signon
    Set Signon = Nothing   'Remove data segment
End Sub

Private Sub cmcExit_GotFocus()
    imcOutline.Move 1560, 2280, 1365, 360
    'imcOutline.Visible = True
End Sub

Private Sub cmcExit_LostFocus()
    imcOutline.Visible = False
End Sub

Private Sub cmcExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move 1560, 2280, 1365, 360 '465, 2190, 3405, 405
    'imcOutline.Visible = True
End Sub

Private Sub cmcStart_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  imSalesperson                 ilLoop                        ilPos                     *
'*                                                                                        *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilPWCompare As Integer
    Dim blStationUser As Boolean

    sgSpecialPassword = ""
    If edcName.Text = "" Then
        Beep
        edcName.SetFocus
        Exit Sub
    End If
    
    sgUserNameToPassToAffiliate = Trim$(edcName.Text)
    sgUserPasswordToPassToAffiliate = Trim$(edcPassword.Text)

    igTestSystem = rbcSystemType(1).Value
    If ckcRemoteUser.Value = vbChecked Then
        igRemoteUser = True
    Else
        igRemoteUser = False
    End If
    'Reset used instead of Close to cause # Clients on network to be decrement
    'Orlando meeting- ignore help messages
    'ilRet = btrClose(hgHlf)
    'btrDestroy hgHlf
    If imDBSelected = 0 Then
        If rbcSystemType(1).Value Then
            btrStopAppl
            gInitGlobalVar   'Initialize global variables
        Else
            gSetBrannerName
        End If
    Else
        If rbcSystemType(0).Value Then
            btrStopAppl
            gInitGlobalVar   'Initialize global variables
        Else
            gSetBrannerName
        End If
    End If
'    If Not gCheckDDFDates() Then
'        igExitTraffic = True
'        Unload Signon
'        Set Signon = Nothing   'Remove data segment
'        Exit Sub
'    End If
    blStationUser = False
    sgUrfStamp = "~" 'Clear time stamp incase same name
    sgUserName = Trim$(edcName.Text)
    '6/20/09:  Jim requested that the Guide sign in be changed to CSI for internal Guide only
    If StrComp(sgUserName, "CSI", vbTextCompare) = 0 Then
        sgUserName = "Guide"
    End If
    gUrfRead Signon, sgUserName, True, tgUrf(), False  'Obtain user records
    If Trim$(tgUrf(0).sName) = "" Then    'No matching name found
        If Not mStationUser() Then
            Beep
            edcName.SetFocus
            Exit Sub
        End If
        blStationUser = True
    End If
    '7/20/06-Allow blank for new users during sign on, then ask for password back in mCheckPW (Traffic.frm)
    'If Trim$(tgUrf(0).sPassword) = "" Then
    '    Beep
    '    edcName.SetFocus
    '    Exit Sub
    'End If
    igGGBDSignOn = False
    ilPWCompare = False
    If ((Asc(tgSpf.sUsingFeatures2) And STRONGPASSWORD) = STRONGPASSWORD) And (tgUrf(0).iCode > 2) Then
        If StrComp(Trim$(edcPassword.Text), Trim$(tgUrf(0).sPassword), vbBinaryCompare) = 0 Then
            ilPWCompare = True
        End If
    Else
        If tgUrf(0).iCode = 1 Then
            If StrComp(Trim$(edcPassword.Text), Trim$("JD#41"), vbTextCompare) = 0 Then
                ilPWCompare = True
            End If
        Else
            If StrComp(Trim$(edcPassword.Text), Trim$(tgUrf(0).sPassword), vbTextCompare) = 0 Then
                ilPWCompare = True
            End If
        End If
    End If
    If Not ilPWCompare Then
        If (tgUrf(0).iCode = 2) And ((StrComp(Trim$(edcPassword.Text), smSpecial, vbTextCompare) = 0) Or (Trim$(Left$(edcPassword.Text, 5)) = "Sport") Or (Trim$(Left$(edcPassword.Text, 5)) = "LiveL")) And (Len(Trim$(edcPassword.Text)) > 5) Then
            If Len(edcPassword.Text) = 9 Then
                '6/20/09:  Jim requested that the Guide sign in be changed to CSI for internal Guide only
                If StrComp(Trim$(edcName.Text), "Guide", vbTextCompare) = 0 Then
                    Beep
                    edcName.SetFocus
                    Exit Sub
                End If
                sgSpecialPassword = Mid$(edcPassword.Text, 6)
                'Dan M 4/08/09 'guide' and currently set user needs expanded rights
                gExpandGuideAsUser tgUrf(0)
            Else    ' Dan M  password 'Sports' sneaking through
                Beep
                edcName.SetFocus
                Exit Sub
            End If
        ElseIf (tgUrf(0).iCode = 2) And (StrComp(Left$(Trim$(edcPassword.Text), 7), smGGBGSpecial, vbTextCompare) = 0) And ((Len(Trim$(edcPassword.Text)) = Len(smGGBGSpecial) + 1) Or (Len(Trim$(edcPassword.Text)) = Len(smGGBGSpecial) + 6)) Then
            sgSpecialPassword = Trim$(edcPassword.Text)
            igGGBDSignOn = True
            If Len(sgSpecialPassword) = 8 Then
                igGGDBSignOnDays = Val(right(Trim$(sgSpecialPassword), 1))
                sgSpecialPassword = Left$(sgSpecialPassword, 7)
            ElseIf Len(sgSpecialPassword) = 13 Then
                igGGDBSignOnDays = gDateValue(Mid$(Trim$(sgSpecialPassword), 8, 2) & "/" & Mid$(Trim$(sgSpecialPassword), 10, 2) & "/" & Mid$(Trim$(sgSpecialPassword), 12, 2)) - gDateValue(Format(Now, "m/d/yy"))
                sgSpecialPassword = Left$(sgSpecialPassword, 7)
            Else
                igGGDBSignOnDays = 0
            End If
        Else
            Beep
            edcName.SetFocus
            Exit Sub
        End If
    End If

    igSportsSystem = 0
    If (Trim$(Left$(edcPassword.Text, 5)) = "Sport") And (sgSpecialPassword <> "") Or (tgUrf(0).sSportPropOnly = "Y") Then
        igSportsSystem = 1
    End If
    If (Trim$(Left$(edcPassword.Text, 5)) = "LiveL") And (sgSpecialPassword <> "") Or (tgUrf(0).sLiveLogPostOnly = "Y") Then
        igSportsSystem = 2
    End If
    If blStationUser Then
        igSportsSystem = 3
        sgStationCallLetters = UCase$(Trim$(edcName.Text))
        sgStationPW = UCase$(Trim$(edcPassword.Text))
    End If
    sgUserName = Trim$(tgUrf(0).sName)
    ilRet = csiSetStamp("RPTLIST", "")
    'ilRet = csiSetAlloc("RptList", 0, 300)
    'For ilLoop = 0 To 300 Step 1
    '    ilRet = csiSetRec("RPTLIST", ilLoop, VarPtr(tmLPopRec), 2406)
    'Next ilLoop
    'gObtainVehicleName tgUrf(0).iDefVeh, sgUserDefVehicleName, slVehType
    'gObtainSalespersonName tgUrf(0).iSlfCode, sgUserSalespersonName
    'ilRet = gVpfRead(Traffic)    'Obtain preference table
    'gObtainVehComboList
    'If tgUrf(0).sClnMoYr = "M" Then
    '    igCalByYear = False
    'Else
    '    igCalByYear = True
    'End If
    'Select Case tgUrf(0).sClnType
    '    Case "S"
    '        igCalMonthType = 0
    '    Case "R"
    '        igCalMonthType = 1
    '    Case "U"
    '        igCalMonthType = 2
    '    Case "D"
    '        igCalMonthType = 3
    '    Case "C"
    '        igCalMonthType = 4
    'End Select
    'If tgUrf(0).sClnLayout = "A" Then
    '    igCalAcross = True
    'Else
    '    igCalAcross = False
    'End If
    'igFromSignon = True
    'igShowPicture = True    'remove this line to get Jog screen up first instead of picture
    'fgCalLeft = tgUrf(0).iClnLeft
    'fgCalTop = tgUrf(0).iClnTop
    'fgCalcLeft = tgUrf(0).iClcLeft
    'fgCalcTop = tgUrf(0).iClcTop
    Traffic!edcPassword.Text = edcPassword.Text
    Unload Signon
    Set Signon = Nothing   'Remove data segment
End Sub

Private Sub cmcStart_GotFocus()
    'imcOutline.Move 1560, 1770, 1365, 360 '465, 1770, 3405, 405
    'imcOutline.Visible = True
End Sub

Private Sub cmcStart_LostFocus()
    'imcOutline.Visible = False
End Sub

Private Sub cmcStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'imcOutline.Move 1560, 1770, 1365, 360 '465, 1770, 3405, 405
    'imcOutline.Visible = True
End Sub

Private Sub edcName_GotFocus()
    'imcOutline.Visible = False
    gCtrlGotFocus edcName
End Sub

Private Sub edcName_LostFocus()
    If Trim(edcName) <> "" Then
        sgUserName = Trim(edcName)
        If StrComp(sgUserName, "CSI", vbTextCompare) = 0 Then
            sgUserName = "Guide"
        End If
        mCheckForBlankPassword
    End If
End Sub
Private Sub mCheckForBlankPassword()
    gUrfRead Signon, sgUserName, False, tgUrf(), False
    If Trim$(tgUrf(0).sName) = "" Then    'No matching name found
        Exit Sub
    End If
    'testing code so guide can login: work around until fix guide ability to change blank password
    If (StrComp(Trim(tgUrf(0).sPassword), "", vbTextCompare)) = 0 And tgUrf(0).iCode > 2 Then
        CSINewPW.lacCurrentPassword.Visible = False
        CSINewPW.edcCurrentPassword.Visible = False
        cmcStart_Click
    End If
End Sub
Private Sub edcPassword_GotFocus()
'    If edcName.Text = "" Then
'        Beep
'        edcName.SetFocus
'        Exit Sub
'    End If
'    sgUserName = Trim$(edcName.Text)
'    gUrfRead Traffic, sgUserName, True, tgUrf()  'Obtain user records
'    If tgUrf(0).sName = "" Then    'No matching name found
'        Beep
'        edcName.SetFocus
'        Exit Sub
'    End If
    'imcOutline.Visible = False
    gCtrlGotFocus edcPassword
End Sub
Private Sub edcPassword_KeyPress(KeyAscii As Integer)
    If Len(edcPassword) = 0 Then
        If KeyAscii = KEYASTERISK Then
            igShowPicture = True
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Form_Activate()
    '4/15/20
    Dim slName As String
    Dim slPassword As String
    Dim ilRet As Integer
    Dim slCommand As String
    'Traffic.Caption = sgBrannerMsg
    mInit
    slCommand = Command$
    If InStr(1, slCommand, UCase("FROMAFFILIATE"), vbTextCompare) > 0 Then
        ilRet = gParseItem(slCommand, 2, "\", slName)
        ilRet = gParseItem(slCommand, 3, "\", slPassword)
        edcName.Text = slName
        edcPassword.Text = slPassword
        cmcStart_Click
        Exit Sub
    End If
    edcName.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyC And Shift = 2 Then
        If edcPassword.Visible = True Then
            edcName.Text = "csi"
            lacCode.Visible = True
            lacCode.SetFocus
            edcPassword.Visible = False
            plcPassword.Visible = False
            plcName.Visible = False
            edcName.Visible = False
        Else
            edcName.Text = ""
            edcPassword.Visible = True
            lacCode.Visible = False
            plcPassword.Visible = True
            edcName.Visible = True
            plcName.Visible = True
            edcName.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    
    'slCommand = sgCommand
    'ilPos = InStr(1, slCommand, "/S:Test", 1)
    'If ilPos > 0 Then
    If igTestSystem Then
        imDBSelected = 1
    Else
        imDBSelected = 0
    End If
    If igTestSystem Then
        'igTestSystem = True
        rbcSystemType(1).Value = True
        rbcSystemType(0).Enabled = True
        rbcSystemType(1).Enabled = True
    End If
    'ilPos = InStr(1, slCommand, "/S:Prod", 1)
    'If ilPos > 0 Then
    If Not igTestSystem Then
        igTestSystem = False
        rbcSystemType(0).Value = True
        rbcSystemType(0).Enabled = True
        rbcSystemType(1).Enabled = True
    End If
    'ilPos = InStr(1, sgCommand, "/RemoteUser", 1)
    'If ilPos > 0 Then
    If igRemoteUser Then
        ckcRemoteUser.Value = vbChecked
    Else
        ckcRemoteUser.Value = vbUnchecked
    End If
    'cmcCSLogo.Picture = Traffic!imcCSLogo.Picture
    igShowPicture = False
    'SignOn.Move 0, 0, Traffic!pbcMsgArea.Width, Traffic!pbcMsgArea.Top
    'plcSignOn.Move (Traffic!pbcMsgArea.Width - plcSignOn.Width) \ 2, (Traffic!pbcMsgArea.Top - plcSignOn.Height) \ 2 + 175'+ Screen.Height \ 10
    If lgCurrBPP <= 8 Then
        plcSignon.Picture = Traffic!pbc256SignOn
    End If
    'Signon.Move  0, 0, Traffic.Width - 120, Traffic.height - 120
    'plcSignon.Move (Signon.Width - plcSignon.Width) \ 2, (Signon.height - plcSignon.height) \ 2 '+ Screen.Height \ 10
    'cmcStart.Move plcSignon.Left, plcSignon.Top
    'cmcExit.Move plcSignon.Left, plcSignon.Top
    
    Signon.Move (Signon.Width - plcSignon.Width) \ 2, (Signon.height - plcSignon.height) \ 2 '+ Screen.Height \ 10
    Signon.Left = (Traffic.Width - Signon.Width) / 2
    Signon.Top = (Traffic.height - Signon.height) / 2
    
    'plcSignOn.Left = 0
    'plcSignOn.Top = 0
    'SignOn.Height = plcSignon.Height + fgPanelAdj
    'SignOn.Width = plcSignon.Width + fgPanelAdj
    'Traffic!lbcVehicle.Clear
    'Traffic!lbcVehicle.Tag = ""
    'Traffic!lbcVehCombo.Clear
    'Traffic!lbcVehCombo.Tag = ""
    'Traffic!lbcUserVehicle.Clear
    'Traffic!lbcUserVehicle.Tag = ""
    'lgDateCreditShown = 0   'Reshow credit alert
    igShowCreditShown = False
'    ilBypassUpdate = False
'    On Error GoTo ByPassUpDateErr
'    ilRet = Len(tgUrf(0))
'    If Not ilBypassUpdate Then
        'gUrfUpdateCal Traffic, tgUrf()  'Update position of calendar/Calculator
'    End If
    'cmcCSLogo.Picture = IconTraf!imcCSLogo.Picture
'Remove call to load tgPopURF, it will be done when Start press by call to gUrfRead
'12/7/03
'    On Error GoTo SignOnErr
'    ilRet = gPopUserListBox(Signon, lbcUser, False)


    On Error GoTo 0
    'gCenterForm SignOn
    Screen.MousePointer = vbDefault
    
    Exit Sub
'ByPassUpDateErr:
'    ilBypassUpdate = True
'    Resume Next

    cmcExit_Click
    Exit Sub
End Sub

Private Sub lacCode_Change()
    edcPassword.Text = "login" & lacCode.Text
End Sub

Private Sub lacExit_Click()
    cmcExit_Click
End Sub

Private Sub lacExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'imcOutline.Move 1560, 2280, 1365, 360 '465, 2190, 3405, 405
    'imcOutline.Visible = True
End Sub

Private Sub lacStart_Click()
    cmcStart_Click
End Sub

Private Sub lacStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'imcOutline.Move 1560, 1770, 1365, 360 '465, 1770, 3405, 405
    'imcOutline.Visible = True
End Sub

Private Sub pbcClickFocus_GotFocus()
    'imcOutline.Visible = False
End Sub

Private Sub plcSignon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ((X >= 1560) And (X <= 1560 + 1365)) And ((Y >= 1770) And (Y <= 1770 + 360)) Then
        'imcOutline.Move 1560, 1770, 1365, 360 '465, 1770, 3405, 405
        'imcOutline.Visible = True
    ElseIf ((X >= 1560) And (X <= 1560 + 1365)) And ((Y >= 2280) And (Y <= 2280 + 360)) Then
        'imcOutline.Move 1560, 2280, 1365, 360 '465, 2190, 3405, 405
        'imcOutline.Visible = True
    Else
        'imcOutline.Visible = False
    End If
End Sub

Private Sub plcSignon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ((X >= 1560) And (X <= 1560 + 1365)) And ((Y >= 1770) And (Y <= 1770 + 360)) Then
        'imcOutline.Move 1560, 1770, 1365, 360 '465, 1770, 3405, 405
        'imcOutline.Visible = True
        cmcStart_Click
    ElseIf ((X >= 1560) And (X <= 1560 + 1365)) And ((Y >= 2280) And (Y <= 2280 + 360)) Then
        'imcOutline.Move 1560, 2280, 1365, 360 '465, 2190, 3405, 405
        'imcOutline.Visible = True
        cmcExit_Click
    End If
End Sub

Private Sub plcSignon_Paint()
    'plcSignon.CurrentX = cmcCSLogo.Left + (3 * cmcCSLogo.Width) / 2
    'plcSignon.CurrentY = cmcCSLogo.Top + cmcCSLogo.Height / 3
    'plcSignon.Print "Counterpoint Software ®"
End Sub
Private Sub plcName_Paint()
    'plcName.CurrentX = 0
    'plcName.CurrentY = 0
    'plcName.Print "Panel3D1"
End Sub

Private Sub rbcSystemType_GotFocus(Index As Integer)
    'imcOutline.Visible = False
End Sub

Private Sub mInit()
    Dim slDate As String
    Dim slMonth As String
    Dim slYear As String
    Dim llValue As Long
    Dim ilValue As Integer
    Dim slStr As String

    bgDevEnv = IsDevEnv()
    slDate = Format$(Now(), "m/d/yy")
    slMonth = month(slDate)
    slYear = year(slDate)
    llValue = Val(slMonth) * Val(slYear)
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    llValue = ilValue
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    slStr = Trim$(str$(ilValue))
    Do While Len(slStr) < 4
        slStr = "0" & slStr
    Loop
    smSpecial = "Login" & slStr
    smGGBGSpecial = "8373191"

End Sub

Private Function mStationUser() As Boolean
    Dim hlVef As Integer
    Dim tlVef As VEF
    Dim ilVefRecLen As Integer
    Dim tlVefSrchKey3 As VEFKEY3 'VEF key record image
    Dim hlVff As Integer
    Dim tlVff As VFF
    Dim ilVffRecLen As Integer
    Dim tlVffSrchKey1 As INTKEY0
    Dim blFound As Boolean
    Dim ilUrf As Integer
    Dim ilRet As Integer
    
    mStationUser = False
    hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlVef)
        btrDestroy hlVef
        Exit Function
    End If
    ilVefRecLen = Len(tlVef)
    tlVefSrchKey3.sName = UCase(Trim$(sgUserName))
    ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlVef)
        btrDestroy hlVef
        Exit Function
    End If
    'Check Password
    hlVff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlVff, "", sgDBPath & "Vff.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlVef)
        btrDestroy hlVef
        ilRet = btrClose(hlVff)
        btrDestroy hlVff
        Exit Function
    End If
    ilVffRecLen = Len(tlVff)
    tlVffSrchKey1.iCode = tlVef.iCode
    ilRet = btrGetEqual(hlVff, tlVff, ilVffRecLen, tlVffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlVef)
        btrDestroy hlVef
        ilRet = btrClose(hlVff)
        btrDestroy hlVff
        Exit Function
    End If
    ilRet = btrClose(hlVef)
    btrDestroy hlVef
    ilRet = btrClose(hlVff)
    btrDestroy hlVff
    If UCase$(Trim$(edcPassword.Text)) <> UCase$(Trim$(tlVff.sStationPassword)) Then
        Exit Function
    End If
    mStationUser = True
    edcPassword.Text = UCase$(edcPassword.Text)
    'Create User if not defined.
    blFound = False
    'For ilUrf = 1 To UBound(tgPopUrf) - 1 Step 1
    For ilUrf = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
        If StrComp(UCase$(Trim$(tgPopUrf(ilUrf).sName)), UCase$("Station"), vbTextCompare) = 0 Then
            tgUrf(0) = tgPopUrf(ilUrf)
            tgUrf(0).sPassword = UCase$(Trim$(tlVff.sStationPassword))
            Exit Function
        End If
    Next ilUrf
    'Create urf
    mAddStationUrf
    tgUrf(0).sPassword = UCase$(Trim$(tlVff.sStationPassword))
End Function

Private Sub mAddStationUrf()
    Dim hlUrf As Integer        'User Option file handle
    Dim tlUrf As URF
    Dim ilUrfRecLen As Integer
    Dim ilRet As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilLoop As Integer
    
    hlUrf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilUrfRecLen = Len(tlUrf)  'btrRecordLength(hlUrf)  'Get and save record length
    
    tlUrf.sName = "Station"
    tlUrf.iVefCode = 0
    tlUrf.sPassword = "Station"
    tlUrf.sRept = "Station"
    tlUrf.sPhoneNo = ""
    tlUrf.sCity = ""
    tlUrf.lEMailCefCode = 0
    tlUrf.iSlfCode = 0
    tlUrf.sClnMoYr = "M"    'Calendar by month
    tlUrf.sClnType = "S"    'Calendar type-standard
    tlUrf.sClnLayout = "A"  'Calendar- Across
    tlUrf.iClnLeft = 100    'Calendar left
    tlUrf.iClnTop = 200 'Calendar top
    tlUrf.iClcLeft = 2775   'Traffic.Width - 2775 - 100   '2625 width of calc form
    tlUrf.iClcTop = 200
    For ilLoop = 0 To 69 Step 1
        tlUrf.sWin(ilLoop) = "H"
    Next ilLoop
    tlUrf.iDefVeh = 0
    tlUrf.sGrid = "H"
    tlUrf.sPrice = "H"
    tlUrf.sCredit = "H"
    tlUrf.sPayRate = "H"
    tlUrf.sMerge = "H"
    tlUrf.sDelete = "N"
    tlUrf.sHideSpots = "H"
    tlUrf.sChgBilled = "H"
    tlUrf.sChgCntr = "H"
    tlUrf.sChgCrRt = "H"
    tlUrf.sBouChk = "H"
    tlUrf.sReprintLogAlert = "N"
    tlUrf.sIncompAlert = "N"
    tlUrf.sCompAlert = "N"
    tlUrf.sSchAlert = "N"
    tlUrf.sHoldAlert = "N"
    tlUrf.sRateCardAlert = "N"
    tlUrf.sResearchAlert = "N"
    tlUrf.sAvailAlert = "N"
    tlUrf.sCrdChkAlert = "N"
    tlUrf.sDeniedAlert = "N"
    tlUrf.sCrdLimitAlert = "N"
    tlUrf.sMoveAlert = "N"
    tlUrf.sPrgmmaticAlert = "H"
    tlUrf.sAdvanceAvails = "H"
    tlUrf.sShowNRMsg = "N"
    tlUrf.sWorkToDead = "N"
    tlUrf.sWorkToComp = "N"
    tlUrf.sWorkToHold = "N"
    tlUrf.sWorkToOrder = "N"
    tlUrf.sCompToIncomp = "N"
    tlUrf.sCompToDead = "N"
    tlUrf.sCompToHold = "N"
    tlUrf.sCompToOrder = "N"
    tlUrf.sIncompToDead = "N"
    tlUrf.sIncompToComp = "N"
    tlUrf.sIncompToHold = "N"
    tlUrf.sIncompToOrder = "N"
    tlUrf.sDeadToWork = "N"
    tlUrf.sHoldToOrder = "N"
    tlUrf.sReviseCntr = "N"
    tlUrf.sResvType = "H"
    tlUrf.sRemType = "H"
    tlUrf.sDRType = "H"
    tlUrf.sPIType = "H"
    tlUrf.sPSAType = "H"
    tlUrf.sPromoType = "H"
    tlUrf.sRefResvType = "H"
    tlUrf.sUseComputeCMC = "H"
    tlUrf.iGroupNo = 0
    tlUrf.sBlockRU = "N"
    tlUrf.sRCView = "R"
    tlUrf.sRegionCopy = "H"
    tlUrf.sChgPrices = "H"
    tlUrf.sChgLnBillPrice = "H"
    tlUrf.sActFlightButton = "H"
    tlUrf.sOldPassword(0) = ""
    tlUrf.sOldPassword(1) = ""
    tlUrf.sOldPassword(2) = ""
    tlUrf.sLiveLogPostOnly = "N"
    tlUrf.sSportPropOnly = "N"
    tlUrf.iRemoteUserID = 0
    tlUrf.iMnfHubCode = 0
    tlUrf.sShowNRMsg = "N"
            ' Dan M new selected fields 4/10/09
    tlUrf.sChangeCSIDate = "H"
    tlUrf.sAllowInvDisplay = "H"
    tlUrf.sActivityLog = "V"

    'Add new value into gInitSuperUser (initsubs.bas)
    tlUrf.iCode = 0  'Autoincrement
    tlUrf.iRemoteID = 0
    tlUrf.iAutoCode = tlUrf.iCode
    gUrfEncrypt tlUrf
    ilRet = btrInsert(hlUrf, tlUrf, ilUrfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlUrf)
        btrDestroy hlUrf
        Exit Sub
    End If
    gUrfDecrypt tlUrf
    Do
        tlUrf.iRemoteID = 0
        tlUrf.iAutoCode = tlUrf.iCode
        gGetSyncDateTime slSyncDate, slSyncTime
        gPackDate slSyncDate, tlUrf.iSyncDate(0), tlUrf.iSyncDate(1)
        gPackTime slSyncTime, tlUrf.iSyncTime(0), tlUrf.iSyncTime(1)
        gUrfEncrypt tlUrf
        ilRet = btrUpdate(hlUrf, tlUrf, ilUrfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    sgPopUrfTag = ""
    ilRet = gObtainUrf()
    gUrfDecrypt tlUrf
    tgUrf(0) = tlUrf
End Sub
