VERSION 5.00
Object = "{0E9D0E41-7AB8-11D1-9400-00A0248F2EF0}#1.0#0"; "dzactx.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmUserLogEMail 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Log E-Mail"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5910
   ScaleWidth      =   9510
   Begin VB.TextBox edcCountdown 
      Height          =   315
      Left            =   3600
      TabIndex        =   7
      Top             =   930
      Width           =   645
   End
   Begin VB.CommandButton cmcSend 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2835
      TabIndex        =   15
      Top             =   4515
      Width           =   2010
   End
   Begin VB.CommandButton cmcAttachment 
      Caption         =   "Browse..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8130
      TabIndex        =   14
      Top             =   4065
      Width           =   1215
   End
   Begin VB.TextBox edcMessage 
      Height          =   1800
      Left            =   1185
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2160
      Width           =   8130
   End
   Begin VB.TextBox edcSubject 
      Height          =   315
      Left            =   1185
      TabIndex        =   9
      Top             =   1740
      Width           =   8130
   End
   Begin VB.TextBox edcAttachment 
      Height          =   315
      Left            =   1440
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4065
      Width           =   6510
   End
   Begin VB.CommandButton cmcExit 
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
      Height          =   450
      Left            =   5175
      TabIndex        =   16
      Top             =   4515
      Width           =   2010
   End
   Begin VB.ListBox lbcResults 
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   645
      ItemData        =   "AffUserLogEMail.frx":0000
      Left            =   885
      List            =   "AffUserLogEMail.frx":0002
      TabIndex        =   18
      Top             =   5070
      Visible         =   0   'False
      Width           =   8445
   End
   Begin VB.CheckBox ckcBlockSystem 
      Caption         =   "Traffic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   1860
      TabIndex        =   4
      Top             =   540
      Width           =   1050
   End
   Begin VB.TextBox edcFromEmail 
      Height          =   315
      Left            =   1200
      TabIndex        =   20
      Top             =   1350
      Width           =   3375
   End
   Begin VB.CheckBox ckcBlockSystem 
      Caption         =   "Affiliate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2940
      TabIndex        =   5
      Top             =   540
      Width           =   1050
   End
   Begin VB.CheckBox ckcMessageType 
      Caption         =   "Screen Alert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3015
      TabIndex        =   1
      Top             =   150
      Width           =   1485
   End
   Begin VB.CheckBox ckcMessageType 
      Caption         =   "E-Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   150
      Width           =   1485
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8535
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblFromEmail 
      Caption         =   "From Email:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   195
      TabIndex        =   19
      Top             =   1380
      Width           =   1125
   End
   Begin VB.Label lacSubject 
      Caption         =   "Subject:"
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
      Left            =   165
      TabIndex        =   8
      Top             =   1740
      Width           =   855
   End
   Begin VB.Label lacCountdown 
      Caption         =   "Countdown in Minutes to Shutdown:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   165
      TabIndex        =   6
      Top             =   975
      Width           =   3300
   End
   Begin VB.Label lacAttachment 
      Caption         =   "Attachment:"
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
      Left            =   165
      TabIndex        =   12
      Top             =   4080
      Width           =   1185
   End
   Begin VB.Label lacResults 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Results:"
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
      Left            =   90
      TabIndex        =   17
      Top             =   5100
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lacMessage 
      Caption         =   "Message:"
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
      Left            =   165
      TabIndex        =   10
      Top             =   2130
      Width           =   855
   End
   Begin VB.Label lacBlockSystem 
      Caption         =   "Block Users on:"
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
      Left            =   165
      TabIndex        =   3
      Top             =   525
      Width           =   1470
   End
   Begin VB.Label lacMessageType 
      Caption         =   "Message Type to be Sent as:"
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
      Left            =   165
      TabIndex        =   0
      Top             =   150
      Width           =   2625
   End
   Begin DZACTXLibCtl.dzactxctrl zpcDZip 
      Left            =   9135
      OleObjectBlob   =   "AffUserLogEMail.frx":0004
      Top             =   4440
   End
End
Attribute VB_Name = "frmUserLogEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Email sent
Private smUserName As String
Private rst_Ulf As ADODB.Recordset
Private rst_User As ADODB.Recordset
'Private Mail As MailSender
'Private smZipPathName As String
'Private bmZipFileCreated As Boolean
Private smToAddress As String
  'Dan M 11/04/09 update Cef email address if currently empty
Dim bmAddFromAddress As Boolean
Public bmEnableEmail As Boolean


Private Sub ckcMessageType_Click(Index As Integer)
    mSetCommands
End Sub

Private Sub edcCountdown_Change()
    mSignOffMsg
End Sub

Private Sub edcCountdown_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'if had no email address and now one is there, then
    If bmAddFromAddress And LenB(Trim(edcFromEmail)) > 0 Then
        'gInsertNewFromAddress edcFromEmail
        If Not ogEmailer Is Nothing Then
            ogEmailer.SetEmailAddressThisUser (Trim$(edcFromEmail.Text))
        End If
    End If
    Dim ilRet As Integer
    
    'Not required here as frmUserLog_close will handle this function
    'ilRet = mCloseCEFFile()
    rst_Ulf.Close
    rst_User.Close
    Set ogEmailer = Nothing
    Set frmUserLogEMail = Nothing
    
End Sub

Private Sub zpcDZip_ZipMajorStatus(ItemName As String, Percent As Long, Cancel As Long)

    lbcResults.List(0) = "Zipping... " & ItemName & "  " & Percent & " % "
    
End Sub
Private Sub zpcDZip_ZipMinorStatus(ItemName As String, Percent As Long, Cancel As Long)

    lbcResults.List(0) = "Zipping... " & ItemName & "  " & Percent & " % "
    
End Sub

Private Sub cmcAttachment_Click()

    Dim ilLoop As Integer
    Dim slCurDir As String
    
    slCurDir = CurDir
    CommonDialog1.Filter = "All Files|*.*"    'Setup the CommonDialog
    CommonDialog1.ShowOpen 'Show the Open Dialog

    If CommonDialog1.fileName <> "" Then
        edcAttachment.Text = Trim$(CommonDialog1.fileName)
    Else
        edcAttachment.Text = ""
    End If

    gChDrDir

End Sub
Private Sub cmcExit_Click()
    Unload frmUserLogEMail
End Sub
Private Sub cmcSend_Click()

    Dim ilRet As Integer
    Dim llEMailCefCode As Long
    Dim slSignOffDate As String
    Dim slStr As String
    Dim ilUlf As Integer
    Dim slEMail As String
    Dim slHost As String
    Dim ilPortNumber As Integer
    Dim slAccountName As String
    Dim ilSendEmail As Integer
    Dim slUserName As String
    'Dim Mail As MailSender
    'Dim tlEmailInfo As EmailInformation 'Dan M. to send to global procedure
    Dim slTempString As String
    lbcResults.Clear
    lbcResults.ForeColor = vbBlack
    
    Screen.MousePointer = vbHourglass
    'Dan M 10/10/11 move to new email class
    'Set Mail = New ASPEMAILLib.MailSender
    ogEmailer.Clear False, True
    If ckcMessageType(0).Value = vbChecked Then
        If igUserLogButton = 2 Then
            'Allow zero to be used to turn it off
            If Trim$(edcCountdown.Text) = "" Then
                Screen.MousePointer = vbDefault
                MsgBox "Countdown must be defined! ", vbCritical
                Exit Sub
            End If
            If Val(edcCountdown.Text) < 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Countdown must a positive number or zero! ", vbCritical
                Exit Sub
            End If
            If (ckcBlockSystem(0).Value = vbUnchecked) And (ckcBlockSystem(1).Value = vbUnchecked) Then
                Screen.MousePointer = vbDefault
                MsgBox "Block User on Traffic and/or Affiliate must be defined! ", vbCritical
                Exit Sub
            End If
        End If
    End If
    If ckcMessageType(1).Value = vbChecked Then
        If Trim$(edcSubject.Text) = "" Then
            Screen.MousePointer = vbDefault
            MsgBox "Subject must be defined! ", vbCritical
            Exit Sub
        End If
    End If
 'dan m 9/19/09 move to if below--only write cef comments for alerts
'    slStr = Trim$(edcMessage.text)
'    If slStr <> "" Then
'        llEMailCefCode = mPutCefComment(0, slStr)
'    Else
'        llEMailCefCode = 0
'    End If
    If ckcMessageType(0).Value = vbChecked Then
        slStr = Trim$(edcMessage.Text)
        If slStr <> "" Then
            llEMailCefCode = mPutCefComment(0, slStr)
        Else
            llEMailCefCode = 0
        End If
        If igUserLogButton = 2 Then
            'Send block
            slStr = Trim$(Str$(igUstCode)) & "|" & Trim$(Str$(llEMailCefCode)) & "|" & Trim$(edcCountdown.Text)
            If (ckcBlockSystem(0).Value = vbChecked) And (ckcBlockSystem(1).Value = vbChecked) Then
                ilRet = gAlertAdd("B", "B", 0, slStr)
            ElseIf (ckcBlockSystem(0).Value = vbChecked) Then
                ilRet = gAlertAdd("B", "T", 0, slStr)
            Else
                ilRet = gAlertAdd("B", "A", 0, slStr)
            End If
        Else
            'Send alert notice
            For ilUlf = 0 To UBound(lgUserLogUlfCode) - 1 Step 1
                SQLQuery = "SELECT * FROM ULF_User_Log WHERE ulfCode = " & lgUserLogUlfCode(ilUlf)
                Set rst_Ulf = gSQLSelectCall(SQLQuery)
                If Not rst_Ulf.EOF Then
                    slSignOffDate = Format$(Trim$(rst_Ulf!ulfSignOffDate), sgShowDateForm)
                    If DateValue(gAdjYear(slSignOffDate)) = DateValue("12/31/2069") Then
                        slStr = Trim$(Str$(rst_Ulf!ulfCode)) & "|" & Trim$(Str$(llEMailCefCode))
                        ilRet = gAlertAdd("N", "", 0, slStr)
                    End If
                End If
            Next ilUlf
        End If
    End If
    If ckcMessageType(1).Value = vbChecked Then
     ' Dan M 9/11/09 removed zipping ability (not implemented anyway) and moved to global
        mShowResults True
'
'        If Len(edcAttachment.text) > 0 Then
'            mShowResults True
'            mZipAllFiles ' - Zip up the files
'            bmZipFileCreated = True
'        Else
'            mShowResults True
'            bmZipFileCreated = False
'        End If
 'Dan M 9/14/09 this currently doing anything, commented out procedure
 '        mWriteEmailFile
    
        ilSendEmail = False
        For ilUlf = 0 To UBound(lgUserLogUlfCode) - 1 Step 1
            llEMailCefCode = 0
            SQLQuery = "SELECT * FROM ULF_User_Log WHERE ulfCode = " & lgUserLogUlfCode(ilUlf)
            Set rst_Ulf = gSQLSelectCall(SQLQuery)
            If Not rst_Ulf.EOF Then
                If rst_Ulf!ulfSystemType = "T" Then
                    SQLQuery = "SELECT * FROM URF_User_Options Where urfCode = " & rst_Ulf!ulfUrfCode
                    Set rst_User = gSQLSelectCall(SQLQuery)
                    If Not rst_User.EOF Then
                        If Trim$(gDecryptField(rst_User!urfRept)) <> "" Then
                            slUserName = Trim$(gDecryptField(rst_User!urfRept))
                        Else
                            slUserName = Trim$(gDecryptField(rst_User!urfName))
                        End If
                        llEMailCefCode = rst_User!urfEMailCefCode
                    End If
                Else
                    SQLQuery = "SELECT * FROM Ust Where ustCode = " & rst_Ulf!ulfUstCode
                    Set rst_User = gSQLSelectCall(SQLQuery)
                    If Not rst_User.EOF Then
                        If Trim$(rst_User!ustReportName) <> "" Then
                            slUserName = Trim$(rst_User!ustReportName)
                        Else
                            slUserName = Trim$(rst_User!ustname)
                        End If
                        
                        llEMailCefCode = rst_User!ustEmailcefcode
                    End If
                End If
            End If
            If llEMailCefCode > 0 Then
                ilRet = mGetCefComment(llEMailCefCode, slEMail)
                If Trim$(slEMail) <> "" Then
                    ogEmailer.AddTOAddress slEMail, slUserName
                    'Mail.AddAddress slEMail, slUserName
                    'Dan M 11/06/09 count number of emails sending; display to user.
                   ' ilSendEmail = True
                    ilSendEmail = ilSendEmail + 1
                End If
            End If
        Next ilUlf
        'Dan made msendemail global, added else statements 9/09/09
        lbcResults.ForeColor = vbRed
        'Dan M 11/06/09 display # of emails sent
        'If ilSendEmail Then
        If ilSendEmail > 0 Then
            With ogEmailer
                .FromAddress = Trim(edcFromEmail.Text)
                .FromName = smUserName
                .Subject = edcSubject.Text
                .Message = edcMessage.Text
                .Attachment = edcAttachment.Text
            End With
'            tlEmailInfo.sFromAddress = Trim(edcFromEmail.Text)
'            'tlEmailInfo.sFromAddress = sgEMail
'            tlEmailInfo.sFromName = smUserName
'            tlEmailInfo.sSubject = edcSubject.Text
'            tlEmailInfo.sMessage = edcMessage.Text
'            tlEmailInfo.sAttachment = edcAttachment.Text
            '10/11/11 Dan M added zip box so affiliate like traffic
            If ogEmailer.Send(lbcResults, zpcDZip) And UBound(lgUserLogUlfCode) > ilSendEmail Then
            'If gSendEmail(tlEmailInfo, lbcResults, , Mail) And UBound(lgUserLogUlfCode) > ilSendEmail Then
                If ilSendEmail = 1 Then
                    slTempString = " 1 email was sent: "
                Else
                    slTempString = ilSendEmail & " emails were sent: "
                End If
                If UBound(lgUserLogUlfCode) - ilSendEmail = 1 Then
                    lbcResults.AddItem slTempString & " a selected user does not have the capacity to receive email."
                Else
                    lbcResults.AddItem slTempString & UBound(lgUserLogUlfCode) - ilSendEmail & " selected users do not have the capacity to receive email."
                End If
            End If
            'to zip, send zip control
             'gSendEmail tlEmailInfo, lbcResults, zpcDZip, Mail
        'Dan M 11/06/09 now that 'logUser' makes sure selected, don't need this
'        ElseIf UBound(lgUserLogUlfCode) = 0 Then
'            lbcResults.AddItem "Please select users to receive email"
        Else
            lbcResults.AddItem "Email was not sent."
        End If

    End If  'email
    On Error Resume Next
    rst_User.Close
    rst_Ulf.Close
    If igUserLogButton = 3 Then
        ilRet = gAlertClear("A", "B", "", 0, "")
    End If
    cmcExit.Caption = "Done"
    'Set Mail = Nothing
    Screen.MousePointer = vbDefault

End Sub
Sub mShowResults(flag As Boolean)
    lbcResults.Visible = flag
    lacResults.Visible = flag
End Sub

Private Sub Form_Load()
    gCenterForm frmUserLogEMail
    mInit
End Sub


Private Sub mInit()
    
    Dim ilRet As Integer
                
    Set ogEmailer = New CEmail
    'Dan M moved to cmcSend 90909
   ' Set Mail = CreateObject("Persits.MailSender")
    If sgFileAttachment <> "" Then
        edcAttachment.Text = sgFileAttachmentName
    End If

    ilRet = mOpenCEFFile()
    If Trim$(sgReportName) <> "" Then
        smUserName = sgReportName
    Else
        smUserName = sgUserName
    End If
    'Dan M 11/06/09
    If bmEnableEmail Then
        ckcMessageType(1).Enabled = True
        'edcFromEmail.Text = gGetEMailAddress(bmAddFromAddress)
        edcFromEmail.Text = ogEmailer.UserEmailAddress
        If Len(edcFromEmail.Text) > 0 Then
            bmAddFromAddress = False
        Else
            bmAddFromAddress = True
        End If
        mSetSubject
    Else
        ckcMessageType(1).Enabled = False
    End If
    'mSetCommands
    'Dan M 11/06/09 code already here: if email enabled, default value is checked msetcommands moved to below and mSetSubject now set above
    If igUserLogButton = 1 Then 'Message
        If ckcMessageType(1).Enabled Then
            ckcMessageType(1).Value = vbChecked
           ' mSetSubject
        Else
            ckcMessageType(0).Value = vbChecked
        End If
        mMessageMsg
    ElseIf igUserLogButton = 2 Then 'Sign Off
        ckcMessageType(0).Value = vbChecked
        If ckcMessageType(1).Enabled Then
            ckcMessageType(1).Value = vbChecked
           ' mSetSubject
        End If
        edcCountdown.Text = 10
        ckcBlockSystem(0).Value = vbChecked
        ckcBlockSystem(1).Value = vbChecked
        mSignOffMsg
    ElseIf igUserLogButton = 3 Then 'Sign on
        ckcMessageType(0).Enabled = False
        'mSetSubject
        mSignOnMsg
    End If
    mSetCommands
End Sub
Private Sub mMessageMsg()
    Dim slName As String
    
    If sgReportName <> "" Then
        slName = sgReportName
    Else
        slName = sgUserName
    End If
    edcMessage.Text = "Message from " & slName & " on " & Now
End Sub
'Private Sub mWriteEmailFile()
'
'    Dim tlTxtStream As TextStream
'    Dim fs As New FileSystemObject
'    Dim slLocation As String
'    slLocation = sgDBPath & "Messages\CSIEmail" & CStr(igUstCode) & ".txt"
'    If fs.FileExists(slLocation) Then
'        fs.DeleteFile slLocation
'        DoEvents
'    End If
'
'    fs.CreateTextFile slLocation, False
'    DoEvents
'    Set tlTxtStream = fs.OpenTextFile(slLocation, ForWriting, False)
'    DoEvents
'    'Fix
'    tlTxtStream.WriteLine ("")  '(txtFromEmail.text)
'    DoEvents
'    tlTxtStream.Close
'    DoEvents
'End Sub


Private Sub mSetCommands()
'Dan M 11/06/09 site doesn't hold email address anymore
'    Dim rstSite As ADODB.Recordset
'    Dim slSiteFrom As String
    
'    SQLQuery = "SELECT * FROM Site Where siteCode = 1"
'    Set rstSite = gSQLSelectCall(SQLQuery)
'    If Not rstSite.EOF Then
'        slSiteFrom = Trim$(rstSite!siteEmailFromAddress)
'    End If
    
    If (sgAllowedToBlock <> "Y") Or (igUserLogButton <> 2) Then
        edcCountdown.Enabled = False
        lacCountdown.Enabled = False
        ckcBlockSystem(0).Enabled = False
        ckcBlockSystem(1).Enabled = False
        lacBlockSystem.Enabled = False
    End If
'Dan M 11/06/09 site doesn't hold email address anymore; always allow email, but toggle if selected
    mEmailOption ckcMessageType(1).Value
   ' If sgEMail = "" And slSiteFrom = "" Then
'        ckcMessageType(1).Enabled = False
'        edcSubject.Enabled = False
'        lacSubject.Enabled = False
'        edcAttachment.Enabled = False
'        lacAttachment.Enabled = False
'        cmcAttachment.Enabled = False
'    Else
'        If ckcMessageType(1).Value = vbChecked Then
'            edcSubject.Enabled = True
'            edcAttachment.Enabled = True
'            cmcAttachment.Enabled = True
'        Else
'            edcSubject.Enabled = False
'            edcAttachment.Enabled = False
'            cmcAttachment.Enabled = False
'        End If
'    End If
    If igUserLogButton = 3 Then 'Sign on
        ckcMessageType(0).Enabled = False
    End If
  '  rstSite.Close
End Sub
Private Sub mEmailOption(blAllow As Boolean)
    edcSubject.Enabled = blAllow
    lacSubject.Enabled = blAllow
    edcAttachment.Enabled = blAllow
    lacAttachment.Enabled = blAllow
    cmcAttachment.Enabled = blAllow
    'dan M 11/04/09 added from name and from email because site options doesn't contain this info anymore
    edcFromEmail.Enabled = blAllow
    lblFromEmail.Enabled = blAllow
'    If blAllow Then
'         edcFromEmail.text = sgEMail
'    Else
'        edcFromEmail.text = ""
'    End If
End Sub

Private Sub mSignOffMsg()
    Dim slNamePlusPhone As String
    Dim slZone As String
    Dim ilPos As Integer
    Dim slLen As String
    Dim ilLen As Integer

    If sgReportName <> "" Then
        slNamePlusPhone = sgReportName
    Else
        slNamePlusPhone = sgUserName
    End If
    If sgPhoneNo <> "" Then
        slNamePlusPhone = slNamePlusPhone & " at " & sgPhoneNo
    End If
    slZone = gGetLocalTZName()
    ilPos = InStr(1, slZone, " ", vbTextCompare)
    If ilPos > 0 Then
        slZone = Left$(slZone, ilPos - 1) & " Time"
    End If
    slLen = Trim$(edcCountdown.Text)
    ilLen = Val(slLen)
    edcMessage.Text = "System will be shutdown on " & Format$(Now, "m/d/yy") & " at " & Format$(gLongToTime(gTimeToLong(Now, False) + 60 * ilLen), "hh:mm am/pm") & " " & slZone & " " & "(" & slLen & " minutes from now)" & sgCRLF & "If you have any problems with this, call " & slNamePlusPhone
End Sub

Private Sub mSignOnMsg()
    Dim slNamePlusPhone As String
    Dim slZone As String
    
    If sgReportName <> "" Then
        slNamePlusPhone = sgReportName
    Else
        slNamePlusPhone = sgUserName
    End If
    If sgPhoneNo <> "" Then
        slNamePlusPhone = slNamePlusPhone & " at " & sgPhoneNo
    End If
    edcMessage.Text = "System is back up!!" & sgCRLF & "Thank you for your patience" & sgCRLF & "If you have any questions call " & slNamePlusPhone
End Sub

Private Sub mSetSubject()
    If igUserLogButton = 1 Then 'Message
        If sgReportName <> "" Then
            edcSubject.Text = "Message from " & sgReportName
        Else
            edcSubject.Text = "Message from " & sgUserName
        End If
        'Dan M commented out line below: blocking if statement above. 9/11/09
       ' edcSubject.text = "Message from "
    ElseIf igUserLogButton = 2 Then 'Sign Off
        edcSubject.Text = "Sign Off Message"
    ElseIf igUserLogButton = 3 Then 'Sign on
        edcSubject.Text = "Sign On Message"
    End If
End Sub

'Private Function mSendEmail() As Integer
'
'    Dim rstSite As ADODB.Recordset
'
'    SQLQuery = "SELECT * FROM Site Where siteCode = 1"
'    Set rstSite = gSQLSelectCall(SQLQuery)
'    If Not rstSite.EOF Then
'        If IsNull(rstSite!siteEmailHost) Then
'            Mail.Host = ""
'        Else
'            Mail.Host = Trim$(rstSite!siteEmailHost)
'        End If
'
'        If IsNull(rstSite!siteEmailPort) Then
'            Mail.Port = ""
'        Else
'            Mail.Port = Trim$(rstSite!siteEmailPort)
'        End If
'        If IsNull(rstSite!siteEmailAcctName) Then
'            Mail.Username = ""
'        Else
'            Mail.Username = Trim$(rstSite!siteEmailAcctName)
'        End If
'   'password and username are upgrades
'        If IsNull(rstSite!siteEmailPassword) Then
'            Mail.Password = ""
'        Else
'            Mail.Password = Trim$(rstSite!siteEmailPassword)
'        End If
'
'        'Use the site option if it exists otherwise use the current users name
'        If IsNull(rstSite!siteEmailFromName) Then
'            Mail.FromName = Trim$(smUserName)
'        Else
'            Mail.FromName = Trim$(rstSite!siteEmailFromName)
'        End If
'
'        'Use the site option if it exists otherwise use the current users email address
'        If IsNull(rstSite!siteEmailFromAddress) Or Trim$(rstSite!siteEmailFromAddress) = "" Then
'            Mail.From = Trim$(sgEMail)
'        Else
'            Mail.From = Trim$(rstSite!siteEmailFromAddress)
'        End If
'    End If
'
'     'Subject
'     If Trim(edcSubject.text) <> "" Then
'         Mail.Subject = Trim(edcSubject.text)
'     Else
'        MsgBox "Please enter a subject."
'        edcSubject.SetFocus
'        Exit Function
'     End If
'
'    'body
'     If Trim(edcMessage.text) <> "" Then
'         Mail.Body = Trim(edcMessage.text)
'     Else
'         MsgBox "Please enter the email body."
'         edcMessage.SetFocus
'         Exit Function
'     End If
'
'    'uncomment this if you want attachments zipped.
'    'Also search for the comment: "uncomment this line to zip the attachment" and uncomment as well
'    ' If smZipPathName <> "" Then
'    '     Mail.AddAttachment smZipPathName
'    ' End If
'
'    'attachments will not be zipped using this code.  Comment out if your going to zip them
'    If Trim$(edcAttachment.text) <> "" Then
'        Mail.AddAttachment edcAttachment.text
'    End If
'    'Dan M 9/09/09 so error will be read
'   On Error Resume Next
'    Mail.Send ' send message
'    If Err <> 0 Then ' error occurred
'        lbcResults.ForeColor = vbRed
'        'Dan M 9/05/09
'        lbcResults.AddItem Err.Description
'       ' lbcResults.text = Err.Description
'    Else
'        lbcResults.Clear
'        lbcResults.ForeColor = vbGreen
'      ' lbcResults.text = "Mail sent successfully"
'        'Dan M 9/05/09
'        lbcResults.AddItem "Mail sent successfully"
'    End If
'    lbcResults.Refresh
'    rstSite.Close
'    Exit Function
'End Function
'Dan M 9/14/09 commented out previous code and stored here
'Option Explicit
''Email sent
'Private smUserName As String
'Private rst_Ulf As ADODB.Recordset
'Private rst_User As ADODB.Recordset
'Private Mail As MailSender
'Private smZipPathName As String
'Private bmZipFileCreated As Boolean
'Private smToAddress As String
'
'' **************************************************************************************
''
''  Procedure:  initZIPCmdStruct()
''
''  Purpose:  Set the ZIP control values
''
'' **************************************************************************************
'Sub minitZIPCmdStruct()
'  zpcDZip.ActionDZ = 0 'NO_ACTION
'  zpcDZip.AddCommentFlag = False
'  zpcDZip.AfterDateFlag = False
'  zpcDZip.BackgroundProcessFlag = False
'  zpcDZip.Comment = ""
'  zpcDZip.CompressionFactor = 5
'  zpcDZip.ConvertLFtoCRLFFlag = False
'  zpcDZip.Date = ""
'  zpcDZip.DeleteOriginalFlag = False
'  zpcDZip.DiagnosticFlag = False
'  zpcDZip.DontCompressTheseSuffixesFlag = False
'  zpcDZip.DosifyFlag = False
'  zpcDZip.EncryptFlag = False
'  zpcDZip.FixFlag = False
'  zpcDZip.FixHarderFlag = False
'  zpcDZip.GrowExistingFlag = False
'  zpcDZip.IncludeFollowing = ""
'  zpcDZip.IncludeOnlyFollowingFlag = False
'  zpcDZip.IncludeSysandHiddenFlag = False
'  zpcDZip.IncludeVolumeFlag = False
'  zpcDZip.ItemList = ""
'  zpcDZip.MajorStatusFlag = True
'  zpcDZip.MessageCallbackFlag = True
'  zpcDZip.MinorStatusFlag = True
'  zpcDZip.MultiVolumeControl = 0
'  zpcDZip.NoDirectoryEntriesFlag = True
'  zpcDZip.NoDirectoryNamesFlag = True
'
'  zpcDZip.OldAsLatestFlag = False
'  zpcDZip.PathForTempFlag = False
'  zpcDZip.QuietFlag = False
'  zpcDZip.RecurseFlag = False
'  zpcDZip.StoreSuffixes = ""
'  zpcDZip.TempPath = ""
'  zpcDZip.ZIPFile = ""
'
'  'Write out a log file in the windows sub directory
'  zpcDZip.ZipSubOptions = 256
'
'  ' added for rev 3.00
'  zpcDZip.RenameCallbackFlag = False
'  zpcDZip.ExtProgTitle = ""
'  zpcDZip.ZIPString = ""
'
'End Sub
'
'Private Sub ckcMessageType_Click(Index As Integer)
'    mSetCommands
'End Sub
'
'Private Sub edcCountdown_Change()
'    mSignOffMsg
'End Sub
'
'Private Sub edcCountdown_KeyPress(KeyAscii As Integer)
'    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
'        Beep
'        KeyAscii = 0
'        Exit Sub
'    End If
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    Dim ilRet As Integer
'
'    ilRet = mCloseCEFFile()
'    Set frmUserLogEMail = Nothing
'
'End Sub
'
'Private Sub zpcDZip_ZipMajorStatus(ItemName As String, Percent As Long, Cancel As Long)
'
'    lbcResults.List(0) = "Zipping... " & ItemName & "  " & Percent & " % "
'
'End Sub
'Private Sub zpcDZip_ZipMinorStatus(ItemName As String, Percent As Long, Cancel As Long)
'
'    lbcResults.List(0) = "Zipping... " & ItemName & "  " & Percent & " % "
'
'End Sub
'
'Function mAddFileToZip(szZip As String, szFile As String) As Integer
'
'    'Init the Zip control structure
'    Call minitZIPCmdStruct
'
'    zpcDZip.ZIPFile = szZip    'The ZIP file name
'    zpcDZip.ItemList = szFile  'The file list to be added
'    zpcDZip.BackgroundProcessFlag = True
'    zpcDZip.ActionDZ = ZIP_ADD   'ADD files to the ZIP file
'    'Returns the error code.  This code can be translated by the sub mTranslateErrors.
'    'It is not currently being used to log to a file.
'    mAddFileToZip = zpcDZip.ErrorCode
'
'End Function
'
'Private Function mTranslateDynaErrors(iError As Integer) As String
'
'    Dim slErrMsg As String
'
'    Select Case iError
'        Case 0
'            slErrMsg = "Backup was Successful."
'        Case 1
'            slErrMsg = "Busy, can't re-enter now."
'        Case 2
'            slErrMsg = "Unexpected end of Zip file."
'        Case 3
'            slErrMsg = "Zip file structure invalid."
'        Case 4
'            slErrMsg = "Out of memory."
'        Case 5
'            slErrMsg = "Internal logic error."
'        Case 6
'            slErrMsg = "Entry too big to split."
'        Case 7
'            slErrMsg = "Invalid comment format."
'        Case 8
'            slErrMsg = "Zip file invalid or insufficient memory."
'        Case 9
'            slErrMsg = "Operation interrupted by application."
'        Case 10
'            slErrMsg = "Temporary file failure."
'        Case 11
'            slErrMsg = "Input file read failure."
'        Case 12
'            slErrMsg = "Nothing to do!"
'        Case 13
'            slErrMsg = "Missing or empty Zip file."
'        Case 14
'            slErrMsg = "Output file write failure, possible disk full."
'        Case 15
'            slErrMsg = "Could not create output file."
'        Case 16
'            slErrMsg = "Invalid control parameters."
'        Case 17
'            slErrMsg = "Could not complete operation."
'        Case 18
'            slErrMsg = "File not found or no read permission."
'        Case 19
'            slErrMsg = "Media Error Encountered."
'        Case 20
'            slErrMsg = "Invalid Multi-Volume control parameters."
'        Case 21
'            slErrMsg = "Improper use of Multi-Volume Zip file."
'    End Select
'
'    mTranslateDynaErrors = slErrMsg
'
'End Function
'
'
'Sub mZipAllFiles()
'
'    Dim ilRet As Integer
'    Dim ilPos As Integer
'    Dim slData As String
'    Dim slDateTime As String
'    Dim slStr As String
'    Dim ilLoop As Integer
'    Dim slCurDir As String
'
'    slCurDir = CurDir
'
'    smZipPathName = ""
'    DoEvents
'    slDateTime = " " & Format$(gNow(), "mmddyy")
'    smZipPathName = sgDBPath & Trim$(sgClientName) & slDateTime & ".zip"  'BuildZipName
'
'    'uncomment this line to zip the attachment
'    'ilRet = mAddFileToZip(smZipPathName, Trim$(edcAttachment.text))
'    gChDrDir
'    DoEvents
'
'End Sub
'
'Private Sub cmcAttachment_Click()
'
'    Dim ilLoop As Integer
'    Dim slCurDir As String
'
'    slCurDir = CurDir
'    CommonDialog1.Filter = "All Files|*.*"    'Setup the CommonDialog
'    CommonDialog1.ShowOpen 'Show the Open Dialog
'
'    If CommonDialog1.fileName <> "" Then
'        edcAttachment.text = Trim$(CommonDialog1.fileName)
'    Else
'        edcAttachment.text = ""
'    End If
'
'    gChDrDir
'
'End Sub
'Private Sub cmcExit_Click()
'    Unload frmUserLogEMail
'End Sub
'Private Sub cmcSend_Click()
'
'    Dim ilRet As Integer
'    Dim llEMailCefCode As Long
'    Dim slSignOffDate As String
'    Dim slStr As String
'    Dim ilUlf As Integer
'    Dim slEMail As String
'    Dim slHost As String
'    Dim ilPortNumber As Integer
'    Dim slAccountName As String
'    Dim ilSendEmail As Integer
'    Dim slUserName As String
'    lbcResults.text = ""
'    lbcResults.ForeColor = vbGreen
'
'    Screen.MousePointer = vbHourglass
'    Set Mail = New ASPEMAILLib.MailSender
'    If ckcMessageType(0).Value = vbChecked Then
'        If igUserLogButton = 2 Then
'            'Allow zero to be used to turn it off
'            If Trim$(edcCountdown.text) = "" Then
'                Screen.MousePointer = vbDefault
'                MsgBox "Countdown must be defined! ", vbCritical
'                Exit Sub
'            End If
'            If Val(edcCountdown.text) < 0 Then
'                Screen.MousePointer = vbDefault
'                MsgBox "Countdown must a positive number or zero! ", vbCritical
'                Exit Sub
'            End If
'            If (ckcBlockSystem(0).Value = vbUnchecked) And (ckcBlockSystem(1).Value = vbUnchecked) Then
'                Screen.MousePointer = vbDefault
'                MsgBox "Block User on Traffic and/or Affiliate must be defined! ", vbCritical
'                Exit Sub
'            End If
'        End If
'    End If
'    If ckcMessageType(1).Value = vbChecked Then
'        If Trim$(edcSubject.text) = "" Then
'            Screen.MousePointer = vbDefault
'            MsgBox "Subject must be defined! ", vbCritical
'            Exit Sub
'        End If
'    End If
'
'    slStr = Trim$(edcMessage.text)
'    If slStr <> "" Then
'        llEMailCefCode = mPutCefComment(0, slStr)
'    Else
'        llEMailCefCode = 0
'    End If
'    If ckcMessageType(0).Value = vbChecked Then
'        If igUserLogButton = 2 Then
'            'Send block
'            slStr = Trim$(Str$(igUstCode)) & "|" & Trim$(Str$(llEMailCefCode)) & "|" & Trim$(edcCountdown.text)
'            If (ckcBlockSystem(0).Value = vbChecked) And (ckcBlockSystem(1).Value = vbChecked) Then
'                ilRet = gAlertAdd("B", "B", 0, slStr)
'            ElseIf (ckcBlockSystem(0).Value = vbChecked) Then
'                ilRet = gAlertAdd("B", "T", 0, slStr)
'            Else
'                ilRet = gAlertAdd("B", "A", 0, slStr)
'            End If
'        Else
'            'Send alert notice
'            For ilUlf = 0 To UBound(lgUserLogUlfCode) - 1 Step 1
'                SQLQuery = "SELECT * FROM ULF_User_Log WHERE ulfCode = " & lgUserLogUlfCode(ilUlf)
'                Set rst_Ulf = gSQLSelectCall(SQLQuery)
'                If Not rst_Ulf.EOF Then
'                    slSignOffDate = Format$(Trim$(rst_Ulf!ulfSignOffDate), sgShowDateForm)
'                    If DateValue(gAdjYear(slSignOffDate)) = DateValue("12/31/2069") Then
'                        slStr = Trim$(Str$(rst_Ulf!ulfCode)) & "|" & Trim$(Str$(llEMailCefCode))
'                        ilRet = gAlertAdd("N", "", 0, slStr)
'                    End If
'                End If
'            Next ilUlf
'        End If
'    End If
'    If ckcMessageType(1).Value = vbChecked Then
'        If Len(edcAttachment.text) > 0 Then
'            mShowResults True
'            mZipAllFiles ' - Zip up the files
'            bmZipFileCreated = True
'        Else
'            mShowResults True
'            bmZipFileCreated = False
'        End If
'        mWriteEmailFile
'
'        ilSendEmail = False
'        For ilUlf = 0 To UBound(lgUserLogUlfCode) - 1 Step 1
'            llEMailCefCode = 0
'            SQLQuery = "SELECT * FROM ULF_User_Log WHERE ulfCode = " & lgUserLogUlfCode(ilUlf)
'            Set rst_Ulf = gSQLSelectCall(SQLQuery)
'            If Not rst_Ulf.EOF Then
'                If rst_Ulf!ulfSystemType = "T" Then
'                    SQLQuery = "SELECT * FROM URF_User_Options Where urfCode = " & rst_Ulf!ulfUrfCode
'                    Set rst_User = gSQLSelectCall(SQLQuery)
'                    If Not rst_User.EOF Then
'                        If Trim$(gDecryptField(rst_User!urfRept)) <> "" Then
'                            slUserName = Trim$(gDecryptField(rst_User!urfRept))
'                        Else
'                            slUserName = Trim$(gDecryptField(rst_User!urfName))
'                        End If
'                        llEMailCefCode = rst_User!urfEMailCefCode
'                    End If
'                Else
'                    SQLQuery = "SELECT * FROM Ust Where ustCode = " & rst_Ulf!ulfUstCode
'                    Set rst_User = gSQLSelectCall(SQLQuery)
'                    If Not rst_User.EOF Then
'                        If Trim$(rst_User!ustReportName) <> "" Then
'                            slUserName = Trim$(rst_User!ustReportName)
'                        Else
'                            slUserName = Trim$(rst_User!ustname)
'                        End If
'
'                        llEMailCefCode = rst_User!ustEMailCefCode
'                    End If
'                End If
'            End If
'            If llEMailCefCode > 0 Then
'                ilRet = mGetCefComment(llEMailCefCode, slEMail)
'                If Trim$(slEMail) <> "" Then
'                    Mail.AddAddress slEMail, slUserName
'                    ilSendEmail = True
'                End If
'            End If
'        Next ilUlf
'        'Dan added else statement 9/09/09
'        lbcResults.ForeColor = vbRed
'        If ilSendEmail Then
'            ilRet = mSendEmail()
'        ElseIf UBound(lgUserLogUlfCode) = 0 Then
'            lbcResults.AddItem "Please select users to receive email"
'        Else
'            lbcResults.AddItem "Email was not sent."
'        End If
'    End If  'email
'    On Error Resume Next
'    rst_User.Close
'    rst_Ulf.Close
'    Set Mail = Nothing
'    If igUserLogButton = 3 Then
'        ilRet = gAlertClear("A", "B", "", 0, "")
'    End If
'    cmcExit.Caption = "Done"
'    Screen.MousePointer = vbDefault
'
'End Sub
'Sub mShowResults(flag As Boolean)
'    lbcResults.Visible = flag
'    lacResults.Visible = flag
'End Sub
'
'Private Sub Form_Load()
'    gCenterForm frmUserLogEMail
'    mInit
'End Sub
'
'
'Private Sub mInit()
'
'    Dim ilRet As Integer
'    'Dan M moved to cmcSend 90909
'   ' Set Mail = CreateObject("Persits.MailSender")
'    If sgFileAttachment <> "" Then
'        edcAttachment.text = sgFileAttachmentName
'    End If
'
'    ilRet = mOpenCEFFile()
'    If Trim$(sgReportName) <> "" Then
'        smUserName = sgReportName
'    Else
'        smUserName = sgUserName
'    End If
'    mSetCommands
'    If igUserLogButton = 1 Then 'Message
'        If ckcMessageType(1).Enabled Then
'            ckcMessageType(1).Value = vbChecked
'            mSetSubject
'        Else
'            ckcMessageType(0).Value = vbChecked
'        End If
'    ElseIf igUserLogButton = 2 Then 'Sign Off
'        ckcMessageType(0).Value = vbChecked
'        If ckcMessageType(1).Enabled Then
'            ckcMessageType(1).Value = vbChecked
'            mSetSubject
'        End If
'        edcCountdown.text = 10
'        ckcBlockSystem(0).Value = vbChecked
'        ckcBlockSystem(1).Value = vbChecked
'        mSignOffMsg
'    ElseIf igUserLogButton = 3 Then 'Sign on
'        ckcMessageType(0).Enabled = False
'        mSetSubject
'        mSignOnMsg
'    End If
'
'End Sub
'
'Private Sub mWriteEmailFile()
'
'    Dim tlTxtStream As TextStream
'    Dim fs As New FileSystemObject
'    Dim slLocation As String
'    slLocation = sgDBPath & "Messages\CSIEmail" & CStr(igUstCode) & ".txt"
'    If fs.FileExists(slLocation) Then
'        fs.DeleteFile slLocation
'        DoEvents
'    End If
'
'    fs.CreateTextFile slLocation, False
'    DoEvents
'    Set tlTxtStream = fs.OpenTextFile(slLocation, ForWriting, False)
'    DoEvents
'    'Fix
'    tlTxtStream.WriteLine ("")  '(txtFromEmail.text)
'    DoEvents
'    tlTxtStream.Close
'    DoEvents
'End Sub
'
'
'Private Sub mSetCommands()
'
'    Dim rstSite As ADODB.Recordset
'    Dim slSiteFrom As String
'
'    SQLQuery = "SELECT * FROM Site Where siteCode = 1"
'    Set rstSite = gSQLSelectCall(SQLQuery)
'    If Not rstSite.EOF Then
'        slSiteFrom = Trim$(rstSite!siteEmailFromAddress)
'    End If
'
'    If (sgAllowedToBlock <> "Y") Or (igUserLogButton <> 2) Then
'        edcCountdown.Enabled = False
'        lacCountdown.Enabled = False
'        ckcBlockSystem(0).Enabled = False
'        ckcBlockSystem(1).Enabled = False
'        lacBlockSystem.Enabled = False
'    End If
'    If sgEMail = "" And slSiteFrom = "" Then
'        ckcMessageType(1).Enabled = False
'        edcSubject.Enabled = False
'        lacSubject.Enabled = False
'        edcAttachment.Enabled = False
'        lacAttachment.Enabled = False
'        cmcAttachment.Enabled = False
'    Else
'        If ckcMessageType(1).Value = vbChecked Then
'            edcSubject.Enabled = True
'            edcAttachment.Enabled = True
'            cmcAttachment.Enabled = True
'        Else
'            edcSubject.Enabled = False
'            edcAttachment.Enabled = False
'            cmcAttachment.Enabled = False
'        End If
'    End If
'    If igUserLogButton = 3 Then 'Sign on
'        ckcMessageType(0).Enabled = False
'    End If
'    rstSite.Close
'End Sub
'
'Private Sub mSignOffMsg()
'    Dim slNamePlusPhone As String
'    Dim slZone As String
'    Dim ilPos As Integer
'    Dim slLen As String
'    Dim ilLen As Integer
'
'    If sgReportName <> "" Then
'        slNamePlusPhone = sgReportName
'    Else
'        slNamePlusPhone = sgUserName
'    End If
'    If sgPhoneNo <> "" Then
'        slNamePlusPhone = slNamePlusPhone & " at " & sgPhoneNo
'    End If
'    slZone = gGetLocalTZName()
'    ilPos = InStr(1, slZone, " ", vbTextCompare)
'    If ilPos > 0 Then
'        slZone = Left$(slZone, ilPos - 1) & " Time"
'    End If
'    slLen = Trim$(edcCountdown.text)
'    ilLen = Val(slLen)
'    edcMessage.text = "System will be shutdown on " & Format$(Now, "m/d/yy") & " at " & Format$(gLongToTime(gTimeToLong(Now, False) + 60 * ilLen), "hh:mm am/pm") & " " & slZone & " " & "(" & slLen & " minutes from now)" & sgCRLF & "If you have any problems with this, call " & slNamePlusPhone
'End Sub
'
'Private Sub mSignOnMsg()
'    Dim slNamePlusPhone As String
'    Dim slZone As String
'
'    If sgReportName <> "" Then
'        slNamePlusPhone = sgReportName
'    Else
'        slNamePlusPhone = sgUserName
'    End If
'    If sgPhoneNo <> "" Then
'        slNamePlusPhone = slNamePlusPhone & " at " & sgPhoneNo
'    End If
'    edcMessage.text = "System is back up!!" & sgCRLF & "Thank you for your patience" & sgCRLF & "If you haver any questions call " & slNamePlusPhone
'End Sub
'
'Private Sub mSetSubject()
'    If igUserLogButton = 1 Then 'Message
'        If sgReportName <> "" Then
'            edcSubject.text = "Message from " & sgReportName
'        Else
'            edcSubject.text = "Message from " & sgUserName
'        End If
'        'Dan M commented out line below: blocking if statement above. 9/11/09
'       ' edcSubject.text = "Message from "
'    ElseIf igUserLogButton = 2 Then 'Sign Off
'        edcSubject.text = "Sign Off Message"
'    ElseIf igUserLogButton = 3 Then 'Sign on
'        edcSubject.text = "Sign On Message"
'    End If
'End Sub
'
'Private Function mSendEmail() As Integer
'
'    Dim rstSite As ADODB.Recordset
'
'    SQLQuery = "SELECT * FROM Site Where siteCode = 1"
'    Set rstSite = gSQLSelectCall(SQLQuery)
'    If Not rstSite.EOF Then
'        If IsNull(rstSite!siteEmailHost) Then
'            Mail.Host = ""
'        Else
'            Mail.Host = Trim$(rstSite!siteEmailHost)
'        End If
'
'        If IsNull(rstSite!siteEmailPort) Then
'            Mail.Port = ""
'        Else
'            Mail.Port = Trim$(rstSite!siteEmailPort)
'        End If
'        If IsNull(rstSite!siteEmailAcctName) Then
'            Mail.Username = ""
'        Else
'            Mail.Username = Trim$(rstSite!siteEmailAcctName)
'        End If
'   'password and username are upgrades
'        If IsNull(rstSite!siteEmailPassword) Then
'            Mail.Password = ""
'        Else
'            Mail.Password = Trim$(rstSite!siteEmailPassword)
'        End If
'
'        'Use the site option if it exists otherwise use the current users name
'        If IsNull(rstSite!siteEmailFromName) Then
'            Mail.FromName = Trim$(smUserName)
'        Else
'            Mail.FromName = Trim$(rstSite!siteEmailFromName)
'        End If
'
'        'Use the site option if it exists otherwise use the current users email address
'        If IsNull(rstSite!siteEmailFromAddress) Or Trim$(rstSite!siteEmailFromAddress) = "" Then
'            Mail.From = Trim$(sgEMail)
'        Else
'            Mail.From = Trim$(rstSite!siteEmailFromAddress)
'        End If
'    End If
'
'     'Subject
'     If Trim(edcSubject.text) <> "" Then
'         Mail.Subject = Trim(edcSubject.text)
'     Else
'        MsgBox "Please enter a subject."
'        edcSubject.SetFocus
'        Exit Function
'     End If
'
'    'body
'     If Trim(edcMessage.text) <> "" Then
'         Mail.Body = Trim(edcMessage.text)
'     Else
'         MsgBox "Please enter the email body."
'         edcMessage.SetFocus
'         Exit Function
'     End If
'
'    'uncomment this if you want attachments zipped.
'    'Also search for the comment: "uncomment this line to zip the attachment" and uncomment as well
'    ' If smZipPathName <> "" Then
'    '     Mail.AddAttachment smZipPathName
'    ' End If
'
'    'attachments will not be zipped using this code.  Comment out if your going to zip them
'    If Trim$(edcAttachment.text) <> "" Then
'        Mail.AddAttachment edcAttachment.text
'    End If
'    'Dan M 9/09/09 so error will be read
'   On Error Resume Next
'    Mail.Send ' send message
'    If Err <> 0 Then ' error occurred
'        lbcResults.ForeColor = vbRed
'        'Dan M 9/05/09
'        lbcResults.AddItem Err.Description
'       ' lbcResults.text = Err.Description
'    Else
'        lbcResults.Clear
'        lbcResults.ForeColor = vbGreen
'      ' lbcResults.text = "Mail sent successfully"
'        'Dan M 9/05/09
'        lbcResults.AddItem "Mail sent successfully"
'    End If
'    lbcResults.Refresh
'    rstSite.Close
'    Exit Function
'End Function
