VERSION 5.00
Begin VB.Form ProgrammaticBuySignOn 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3600
   ClientLeft      =   2310
   ClientTop       =   2055
   ClientWidth     =   4380
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
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3600
   ScaleWidth      =   4380
   Begin VB.PictureBox plcProgrammaticBuySignOn 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   -15
      Picture         =   "ProgrammaticBuySignOn.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   -15
      Width           =   4395
      Begin VB.CheckBox ckcRemoteUser 
         Caption         =   "Suppress Background Picture"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   105
         TabIndex        =   9
         Top             =   3420
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.OptionButton rbcSystemType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Test System"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1155
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3090
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.OptionButton rbcSystemType 
         BackColor       =   &H80000009&
         Caption         =   "Production System"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1155
         TabIndex        =   7
         Top             =   2820
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   180
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
         TabIndex        =   12
         Top             =   1500
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox edcName 
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
         Height          =   270
         Left            =   1575
         TabIndex        =   1
         Top             =   690
         Width           =   2130
      End
      Begin VB.TextBox edcPassword 
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
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1575
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2130
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
         Left            =   1545
         ScaleHeight     =   285
         ScaleWidth      =   2130
         TabIndex        =   11
         Top             =   1050
         Width           =   2190
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
         Left            =   1545
         ScaleHeight     =   285
         ScaleWidth      =   2130
         TabIndex        =   10
         Top             =   660
         Width           =   2190
      End
      Begin VB.PictureBox pbcClickFocus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   -60
         ScaleHeight     =   165
         ScaleWidth      =   105
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   825
         Width           =   105
      End
      Begin VB.Label imcOutline 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   420
         Left            =   3105
         TabIndex        =   18
         Top             =   1920
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lacExit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   300
         Left            =   1560
         TabIndex        =   5
         Top             =   2295
         Width           =   1335
      End
      Begin VB.Label lacStart 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   300
         Left            =   1560
         TabIndex        =   3
         Top             =   1785
         Width           =   1350
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
         Left            =   1545
         TabIndex        =   17
         Top             =   3060
         Visible         =   0   'False
         Width           =   480
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
         Left            =   1545
         TabIndex        =   16
         Top             =   2775
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lacPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   435
         TabIndex        =   15
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label lacUserName 
         BackStyle       =   0  'Transparent
         Caption         =   "Buyer Name"
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
         Left            =   195
         TabIndex        =   14
         Top             =   690
         Width           =   1170
      End
      Begin VB.Image cmcCSLogo 
         Height          =   510
         Left            =   60
         Top             =   60
         Width           =   3210
      End
   End
   Begin VB.CommandButton cmcExit 
      Appearance      =   0  'Flat
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   6
      Top             =   2280
      Width           =   1365
   End
   Begin VB.CommandButton cmcStart 
      Appearance      =   0  'Flat
      Caption         =   "&START System"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   4
      Top             =   1770
      Width           =   1365
   End
End
Attribute VB_Name = "ProgrammaticBuySignOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim imDBSelected As Integer  '0=Production; 1=Test
Dim smSpecial As String
Dim smGGBGSpecial As String
Private Sub cmcCSLogo_Click()
         On Error GoTo CP1
gAddCallToStack "cmcCSLogo_Click"
         Dim iAns As Integer
         Dim ilErr As Integer
0001     imcOutline.Visible = False
0002     CSIAbout.Show vbModal
         Goto XP1
0003     ilErr = True
0004     iAns = MsgBox(sgCSVersion, 0, "Counterpoint Software")
         Resume Next
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, cmcCSLogo_Click")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub cmcExit_Click()
         On Error GoTo CP1
gAddCallToStack "cmcExit_Click"
0001     sgSpecialPassword = ""
0002     igExitTraffic = True
0003     Unload ProgrammaticBuySignOn
0004     Set ProgrammaticBuySignOn = Nothing   'Remove data segment
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, cmcExit_Click")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub cmcExit_GotFocus()
         On Error GoTo CP1
gAddCallToStack "cmcExit_GotFocus"
0001     imcOutline.Move 1560, 2280, 1365, 360
0002     imcOutline.Visible = True
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, cmcExit_GotFocus")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub cmcExit_LostFocus()
         On Error GoTo CP1
gAddCallToStack "cmcExit_LostFocus"
0001     imcOutline.Visible = False
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, cmcExit_LostFocus")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub cmcExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
         On Error GoTo CP1
gAddCallToStack "cmcExit_MouseMove"
0001     imcOutline.Move 1560, 2280, 1365, 360 '465, 2190, 3405, 405
0002     imcOutline.Visible = True
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, cmcExit_MouseMove")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub cmcStart_Click()
         On Error GoTo CP1
gAddCallToStack "cmcStart_Click"
         Dim ilRet As Integer
         Dim ilPWCompare As Integer
0001     sgSpecialPassword = ""
0002     If edcName.Text = "" Then
             Beep
0003         edcName.SetFocus
             Goto XP1
         End If
0004     igTestSystem = rbcSystemType(1).Value
0005     If ckcRemoteUser.Value = vbChecked Then
0006         igRemoteUser = True
         Else
0007         igRemoteUser = False
         End If
0008     sgUrfStamp = "~" 'Clear time stamp incase same name
0009     sgUserName = Trim$(edcName.Text)
0010     If StrComp(sgUserName, "CSI", vbTextCompare) = 0 Then
0011         sgUserName = "Guide"
         End If
0012     gUrfRead ProgrammaticBuySignOn, sgUserName, True, tgUrf(), False  'Obtain user records
0013     If Trim$(tgUrf(0).sName) = "" Then    'No matching name found
             Beep
0014         edcName.SetFocus
             Goto XP1
         End If
0015     igGGBDSignOn = False
0016     ilPWCompare = False
0017     If ((Asc(tgSpf.sUsingFeatures2) And STRONGPASSWORD) = STRONGPASSWORD) And (tgUrf(0).iCode > 2) Then
0018         If StrComp(Trim$(edcPassword.Text), Trim$(tgUrf(0).sPassword), vbBinaryCompare) = 0 Then
0019             ilPWCompare = True
             End If
         Else
0020         If tgUrf(0).iCode = 1 Then
0021             If StrComp(Trim$(edcPassword.Text), Trim$("JD#41"), vbTextCompare) = 0 Then
0022                 ilPWCompare = True
                 End If
             Else
0023             If StrComp(Trim$(edcPassword.Text), Trim$(tgUrf(0).sPassword), vbTextCompare) = 0 Then
0024                 ilPWCompare = True
                 End If
             End If
         End If
0025     If Not ilPWCompare Then
0026         If (tgUrf(0).iCode = 2) And ((StrComp(Trim$(edcPassword.Text), smSpecial, vbTextCompare) = 0) Or (Trim$(Left$(edcPassword.Text, 5)) = "Sport") Or (Trim$(Left$(edcPassword.Text, 5)) = "LiveL")) And (Len(Trim$(edcPassword.Text)) > 5) Then
0027             If Len(edcPassword.Text) = 9 Then
0028                 If StrComp(Trim$(edcName.Text), "Guide", vbTextCompare) = 0 Then
                         Beep
0029                     edcName.SetFocus
                         Goto XP1
                     End If
0030                 sgSpecialPassword = Mid$(edcPassword.Text, 6)
0031                 gExpandGuideAsUser tgUrf(0)
                 Else    ' Dan M  password 'Sports' sneaking through
                     Beep
0032                 edcName.SetFocus
                     Goto XP1
                 End If
0033         ElseIf (tgUrf(0).iCode = 2) And (StrComp(Left$(Trim$(edcPassword.Text), 7), smGGBGSpecial, vbTextCompare) = 0) And ((Len(Trim$(edcPassword.Text)) = Len(smGGBGSpecial) + 1) Or (Len(Trim$(edcPassword.Text)) = Len(smGGBGSpecial))) Then
0034             sgSpecialPassword = Trim$(edcPassword.Text)
0035             igGGBDSignOn = True
0036             If Len(sgSpecialPassword) = 8 Then
0037                 igGGDBSignOnDays = Val(right(Trim$(sgSpecialPassword), 1))
0038                 sgSpecialPassword = Left$(sgSpecialPassword, 7)
                 Else
0039                 igGGDBSignOnDays = 0
                 End If
             Else
                 Beep
0040             edcName.SetFocus
                 Goto XP1
             End If
         End If
0041     igSportsSystem = 0
0042     If (Trim$(Left$(edcPassword.Text, 5)) = "Sport") And (sgSpecialPassword <> "") Or (tgUrf(0).sSportPropOnly = "Y") Then
0043         igSportsSystem = 1
         End If
0044     If (Trim$(Left$(edcPassword.Text, 5)) = "LiveL") And (sgSpecialPassword <> "") Or (tgUrf(0).sLiveLogPostOnly = "Y") Then
0045         igSportsSystem = 2
         End If
0046     sgUserName = Trim$(tgUrf(0).sName)
0047     ilRet = csiSetStamp("RPTLIST", "")
0048     Unload ProgrammaticBuySignOn
0049     Set ProgrammaticBuySignOn = Nothing   'Remove data segment
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, cmcStart_Click")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub cmcStart_GotFocus()
         On Error GoTo CP1
gAddCallToStack "cmcStart_GotFocus"
0001     imcOutline.Move 1560, 1770, 1365, 360 '465, 1770, 3405, 405
0002     imcOutline.Visible = True
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, cmcStart_GotFocus")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub cmcStart_LostFocus()
         On Error GoTo CP1
gAddCallToStack "cmcStart_LostFocus"
0001     imcOutline.Visible = False
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, cmcStart_LostFocus")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub cmcStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
         On Error GoTo CP1
gAddCallToStack "cmcStart_MouseMove"
0001     imcOutline.Move 1560, 1770, 1365, 360 '465, 1770, 3405, 405
0002     imcOutline.Visible = True
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, cmcStart_MouseMove")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub edcName_GotFocus()
         On Error GoTo CP1
gAddCallToStack "edcName_GotFocus"
0001     imcOutline.Visible = False
0002     gCtrlGotFocus edcName
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, edcName_GotFocus")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub edcName_LostFocus()
         On Error GoTo CP1
gAddCallToStack "edcName_LostFocus"
0001     If Trim(edcName) <> "" Then
0002         sgUserName = Trim(edcName)
0003         If StrComp(sgUserName, "CSI", vbTextCompare) = 0 Then
0004             sgUserName = "Guide"
             End If
0005         mCheckForBlankPassword
         End If
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, edcName_LostFocus")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub mCheckForBlankPassword()
         On Error GoTo CP1
gAddCallToStack "mCheckForBlankPassword"
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, mCheckForBlankPassword")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub edcPassword_GotFocus()
         On Error GoTo CP1
gAddCallToStack "edcPassword_GotFocus"
0001     imcOutline.Visible = False
0002     gCtrlGotFocus edcPassword
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, edcPassword_GotFocus")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub edcPassword_KeyPress(KeyAscii As Integer)
         On Error GoTo CP1
gAddCallToStack "edcPassword_KeyPress"
0001     If Len(edcPassword) = 0 Then
0002         If KeyAscii = KEYASTERISK Then
0003             igShowPicture = True
0004             KeyAscii = 0
             End If
         End If
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, edcPassword_KeyPress")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub Form_Activate()
         On Error GoTo CP1
gAddCallToStack "Form_Activate"
0001     mInit
0002     edcName.SetFocus
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, Form_Activate")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub Form_Load()
         On Error GoTo CP1
gUserActivityLog "L", "ProgrammaticBuySignOn.FRM"
gAddCallToStack "Form_Load"
0001     If igTestSystem Then
0002         imDBSelected = 1
         Else
0003         imDBSelected = 0
         End If
0004     If igTestSystem Then
0005         rbcSystemType(1).Value = True
0006         rbcSystemType(0).Enabled = True
0007         rbcSystemType(1).Enabled = True
         End If
0008     If Not igTestSystem Then
0009         igTestSystem = False
0010         rbcSystemType(0).Value = True
0011         rbcSystemType(0).Enabled = True
0012         rbcSystemType(1).Enabled = True
         End If
0013     If igRemoteUser Then
0014         ckcRemoteUser.Value = vbChecked
         Else
0015         ckcRemoteUser.Value = vbUnchecked
         End If
0016     igShowPicture = False
0017     If lgCurrBPP <= 8 Then
0018         plcProgrammaticBuySignOn.Picture = Traffic!pbc256SignOn
         End If
0019     gCenterStdAlone ProgrammaticBuySignOn
0020     cmcStart.Move plcProgrammaticBuySignOn.Left, plcProgrammaticBuySignOn.Top
0021     cmcExit.Move plcProgrammaticBuySignOn.Left, plcProgrammaticBuySignOn.Top
0022     igShowCreditShown = False
0023     On Error GoTo CP1
0024     Screen.MousePointer = vbDefault
         Goto XP1
0025     cmcExit_Click
         Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, Form_Load")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub lacExit_Click()
         On Error GoTo CP1
gAddCallToStack "lacExit_Click"
0001     cmcExit_Click
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, lacExit_Click")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub lacExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
         On Error GoTo CP1
gAddCallToStack "lacExit_MouseMove"
0001     imcOutline.Move 1560, 2280, 1365, 360 '465, 2190, 3405, 405
0002     imcOutline.Visible = True
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, lacExit_MouseMove")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub lacStart_Click()
         On Error GoTo CP1
gAddCallToStack "lacStart_Click"
0001     cmcStart_Click
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, lacStart_Click")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub lacStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
         On Error GoTo CP1
gAddCallToStack "lacStart_MouseMove"
0001     imcOutline.Move 1560, 1770, 1365, 360 '465, 1770, 3405, 405
0002     imcOutline.Visible = True
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, lacStart_MouseMove")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub pbcClickFocus_GotFocus()
         On Error GoTo CP1
gAddCallToStack "pbcClickFocus_GotFocus"
0001     imcOutline.Visible = False
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, pbcClickFocus_GotFocus")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub plcProgrammaticBuySignOn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
         On Error GoTo CP1
gAddCallToStack "plcProgrammaticBuySignOn_MouseMove"
0001     If ((X >= 1560) And (X <= 1560 + 1365)) And ((Y >= 1770) And (Y <= 1770 + 360)) Then
0002         imcOutline.Move 1560, 1770, 1365, 360 '465, 1770, 3405, 405
0003         imcOutline.Visible = True
0004     ElseIf ((X >= 1560) And (X <= 1560 + 1365)) And ((Y >= 2280) And (Y <= 2280 + 360)) Then
0005         imcOutline.Move 1560, 2280, 1365, 360 '465, 2190, 3405, 405
0006         imcOutline.Visible = True
         Else
0007         imcOutline.Visible = False
         End If
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, plcProgrammaticBuySignOn_MouseMove")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub plcProgrammaticBuySignOn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
         On Error GoTo CP1
gAddCallToStack "plcProgrammaticBuySignOn_MouseUp"
0001     If ((X >= 1560) And (X <= 1560 + 1365)) And ((Y >= 1770) And (Y <= 1770 + 360)) Then
0002         imcOutline.Move 1560, 1770, 1365, 360 '465, 1770, 3405, 405
0003         imcOutline.Visible = True
0004         cmcStart_Click
0005     ElseIf ((X >= 1560) And (X <= 1560 + 1365)) And ((Y >= 2280) And (Y <= 2280 + 360)) Then
0006         imcOutline.Move 1560, 2280, 1365, 360 '465, 2190, 3405, 405
0007         imcOutline.Visible = True
0008         cmcExit_Click
         End If
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, plcProgrammaticBuySignOn_MouseUp")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub plcProgrammaticBuySignOn_Paint()
         On Error GoTo CP1
gAddCallToStack "plcProgrammaticBuySignOn_Paint"
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, plcProgrammaticBuySignOn_Paint")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub plcName_Paint()
         On Error GoTo CP1
gAddCallToStack "plcName_Paint"
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, plcName_Paint")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub rbcSystemType_GotFocus(Index As Integer)
         On Error GoTo CP1
gAddCallToStack "rbcSystemType_GotFocus"
0001     imcOutline.Visible = False
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, rbcSystemType_GotFocus")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub mInit()
         On Error GoTo CP1
gAddCallToStack "mInit"
         Dim slDate As String
         Dim slMonth As String
         Dim slYear As String
         Dim llValue As Long
         Dim ilValue As Integer
         Dim slStr As String
0001     pbcClickFocus.Left = -pbcClickFocus.Width - 30
0002     slDate = Format$(Now(), "ddddd")
0003     slMonth = Month(slDate)
0004     slYear = Year(slDate)
0005     llValue = Val(slMonth) * Val(slYear)
0006     ilValue = Int(10000 * Rnd(-llValue) + 1)
0007     llValue = ilValue
0008     ilValue = Int(10000 * Rnd(-llValue) + 1)
0009     slStr = Trim$(Str$(ilValue))
0010     Do While Len(slStr) < 4
0011         slStr = "0" & slStr
         Loop
0012     smSpecial = "Login" & slStr
0013     smGGBGSpecial = "8373191"
Goto XP1
CP1:
    gDbg_HandleError("ProgrammaticBuySignOn.FRM, mInit")
XP1:
Call gRemoveCallFromStack
End Sub
Private Sub Form_Unload(Cancel As Integer)
gUserActivityLog "U", "ProgrammaticBuySignOn.FRM"
End Sub
