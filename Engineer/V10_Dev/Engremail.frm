VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3AC17E6A-5D9D-4304-B1C6-E501BE604D2F}#4.8#0"; "PBEMail1.ocx"
Begin VB.Form EngrEmail 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6060
   ScaleWidth      =   9510
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8280
      Top             =   4920
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6060
      FormDesignWidth =   9510
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   5040
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.ComboBox cbcToName 
      Height          =   315
      ItemData        =   "Engremail.frx":0000
      Left            =   1170
      List            =   "Engremail.frx":0016
      TabIndex        =   3
      Top             =   945
      Width           =   3390
   End
   Begin VB.CommandButton cmdClrAttach 
      Caption         =   "Clear Attachments"
      Height          =   330
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   1860
   End
   Begin VB.CommandButton cmdExit 
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
      TabIndex        =   10
      Top             =   5520
      Width           =   2010
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7650
      Top             =   5490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtAttachments 
      Height          =   315
      Left            =   1740
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4515
      Width           =   7605
   End
   Begin VB.TextBox txtSubject 
      Height          =   315
      Left            =   1185
      TabIndex        =   5
      Top             =   1515
      Width           =   8130
   End
   Begin VB.TextBox txtMsgBox 
      Height          =   2280
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2025
      Width           =   9195
   End
   Begin VB.CommandButton cmdAttach 
      Caption         =   "Add Attachments"
      Height          =   330
      Left            =   150
      TabIndex        =   7
      Top             =   5160
      Width           =   1860
   End
   Begin VB.CommandButton cmdSend 
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
      TabIndex        =   9
      Top             =   5520
      Width           =   2010
   End
   Begin VB.TextBox txtFromEmail 
      Height          =   315
      Left            =   5910
      TabIndex        =   2
      Top             =   525
      Width           =   3375
   End
   Begin VB.TextBox txtToEmail 
      Height          =   315
      Left            =   5925
      TabIndex        =   4
      Top             =   975
      Width           =   3375
   End
   Begin VB.TextBox txtFromName 
      Height          =   315
      Left            =   1185
      TabIndex        =   1
      Top             =   510
      Width           =   3390
   End
   Begin VB.Label lblHeader 
      Caption         =   "Email"
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
      Left            =   90
      TabIndex        =   17
      Top             =   105
      Width           =   2205
   End
   Begin VB.Label lblAttachments 
      Caption         =   "Attachments:"
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
      Left            =   405
      TabIndex        =   16
      Top             =   4515
      Width           =   1185
   End
   Begin VB.Label lblToName 
      Caption         =   "To Name:"
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
      Left            =   225
      TabIndex        =   15
      Top             =   975
      Width           =   900
   End
   Begin VB.Label lblToEmail 
      Caption         =   "To Email:"
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
      Left            =   5025
      TabIndex        =   14
      Top             =   990
      Width           =   870
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
      Left            =   4845
      TabIndex        =   13
      Top             =   540
      Width           =   1125
   End
   Begin VB.Label lblFrom 
      Caption         =   "From Name:"
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
      Left            =   90
      TabIndex        =   12
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label lblSubject 
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
      Left            =   315
      TabIndex        =   11
      Top             =   1515
      Width           =   855
   End
   Begin PBEmail1.PBEmail PBEmail1 
      Left            =   8325
      Top             =   5445
      _ExtentX        =   1879
      _ExtentY        =   900
      From1           =   ""
      FromName1       =   ""
      To1             =   ""
      ToName1         =   ""
      Subject1        =   ""
      Body1           =   ""
      HtmlMsg         =   0   'False
      SMTPServer1     =   ""
      ActAsSMTP1      =   -1  'True
      REGUserName1    =   "Doug Smith"
      REGKeyCode1     =   "GIWDCDGJFDGIVTHSUCIDVHTCUTCJJHFWVUDGWSDBGTJVHUEFCVWICCTFRIDJVDTJBCBFWIFGFFFVFGDF"
      ReplyTo1        =   ""
      Organization1   =   ""
      Priority1       =   3
      XMailer1        =   "PABLOB.NET - PBEmail 1.x"
      CharSet1        =   2
      AskForReceipt1  =   0   'False
   End
End
Attribute VB_Name = "EngrEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbcToName_Click()

    Dim ilIndex As Integer
    Dim ilRet As Integer

    ilIndex = cbcToName.ListIndex

    Select Case ilIndex
        Case -1
            ilRet = ilRet
        Case 0
            txtToEmail.text = "NoraBrowne@Counterpoint.net"
            cbcToName.text = "Nora Browne"
        Case 1
            txtToEmail.text = "AnnaDow@Counterpoint.net"
            cbcToName.text = "Anna Dow"

        Case 2
            txtToEmail.text = "MaryNelson@Counterpoint.net"
            cbcToName.text = "Mary Nelson"

        Case 3
            txtToEmail.text = "SamanthaCarroll@Counterpoint.net"
            cbcToName.text = "Samantha Carroll"

        Case 4
            txtToEmail.text = "EricGarst@Counterpoint.net"
            cbcToName.text = "Eric Garst"

        Case 5
            txtToEmail.text = "Service@Counterpoint.net"
            cbcToName.text = "Service Department"
    End Select


End Sub


Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    gSetFonts EngrEmail
    gCenterForm EngrEmail
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set EngrEmail = Nothing
End Sub

Private Sub PBEmail1_ProgressChanged()
    'Update the ProgressBar in Sending form
    ProgressBar1.Value = PBEmail1.SendProgress
    DoEvents
End Sub

Private Sub PBEmail1_SendCompleted(Result As PBEmail1.PBErrorConstants)
    'Update the Sending Dialogs
'    Sending.cmdClose.Enabled = True
'    If Result = PBMessageSentOK Then
'        Sending.lblSending.Caption = "Message Sent Ok."
'    Else
'        Sending.lblSending.Caption = PBEmail1.GetLastError
'    End If
    ProgressBar1.Value = 100
    cmdExit.Caption = "Done"
    cmdClrAttach_Click
    mWriteEmailFile
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAttach_Click()

    Dim ilLoop As Integer

    'Setup the CommonDialog
    CommonDialog1.Filter = "All Files|*.*"

    'Show the Open Dialog
    CommonDialog1.ShowOpen

    If CommonDialog1.fileName <> "" Then
        If PBEmail1.Attachments.Add(CommonDialog1.fileName) = False Then
            MsgBox "Unable to Add Attachment!", vbInformation, "Error"
        End If
    End If

    'Update Attachment Text
    txtAttachments.text = ""
    For ilLoop = 1 To PBEmail1.Attachments.Count
        txtAttachments.text = txtAttachments.text & PBEmail1.Attachments.ItemName(ilLoop) & "; "
    Next ilLoop

End Sub

Private Sub cmdClrAttach_Click()

    Dim ilLoop As Integer

    'Clear the object
    PBEmail1.Attachments.Clear
    'Clear the Attachment Text
    txtAttachments.text = ""
    For ilLoop = 1 To PBEmail1.Attachments.Count
        txtAttachments.text = txtAttachments.text & PBEmail1.Attachments.ItemName(ilLoop) & "; "
    Next ilLoop

End Sub

Private Sub cmdExit_Click()
    Unload EngrEmail

End Sub

Private Sub cmdSend_Click()

    Dim ilResponse As Integer

    Screen.MousePointer = vbHourglass
    'Validate Data
    If txtFromEmail.text = "" Or InStr(1, txtFromEmail.text, "@") = 0 Then
        MsgBox "You must enter Your Email Address", vbInformation, "Error"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If txtToEmail.text = "" Or InStr(1, txtToEmail.text, "@") = 0 Then
        MsgBox "You must enter the Recipient's Email Address", vbInformation, "Error"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    PBEmail1.FromEmail = txtFromEmail.text
    PBEmail1.FromName = txtFromName.text

    'Reply-To
    PBEmail1.ReplyTo = "" 'txtReplyTo.Text

    'Organization
    PBEmail1.Organization = "" 'txtOrganization.Text

    'Recipient Information
    PBEmail1.ToEmail = txtToEmail.text
    PBEmail1.ToName = cbcToName.text

    'Priority
    PBEmail1.Priority = PBNormal
'    Select Case cmbPriority.ListIndex
'        Case 0:
'            'High Priority
'            PBEmail1.Priority = PBHigh
'        Case 1:
'            'Normal Priority
'            PBEmail1.Priority = PBNormal
'        Case 2:
'            'Low Priority
'            PBEmail1.Priority = PBLow
'    End Select

    'Content Type
    PBEmail1.ContentType = PB_TEXT
'    Select Case cboType.ListIndex
'        Case 0:
'            'TEXT Format
'            PBEmail1.ContentType = PB_TEXT
'        Case 1:
'            'HTML Format
'            PBEmail1.ContentType = PB_HTML
'    End Select

    'Message and Subject
    If txtMsgBox.text = "" Then
        txtMsgBox.text = " ** No Message **"
    End If
    PBEmail1.Body = txtMsgBox.text
    PBEmail1.Subject = txtSubject.text

    'Shows the Sending From, that Calls the SendEmail
    'Me.Enabled = False
    'CenterStdAlone Sending
    'Sending.Show vbModal

    'Send the Email
    ilResponse = PBEmail1.SendEmail

    'If ilResponse = PBMessageSending, then the Message
    'is on its way, just wait for the SendCompleted Event

    If ilResponse <> PBMessageSending Then
        Screen.MousePointer = vbDefault
        'There was an Error Starting the Send of the Email
        If ilResponse = PBDNSServerError Then
            'There Was an DNS Server Error, may not be connected?
            'Sending.lblSending.Caption = "DNS Error, Maybe not Connected!"
            MsgBox "DNS Error, May not be Connected!", vbOKOnly
        Else
            'Server Error, get The Error String
            'Sending.lblSending.Caption = PBEmail1.GetLastError
            MsgBox PBEmail1.GetLastError, vbOKOnly
        End If
        'Enable the Exit command
        'Sending.cmdClose.Enabled = True
        'Moved here at the one at end moved to complete event
        cmdClrAttach_Click
        mWriteEmailFile
        Screen.MousePointer = vbDefault
    Else
'Moved to PBEMail1_SendCompleted 3/30/04
'        'MsgBox "Message Was Sent Successfully!", vbOKOnly
'        ProgressBar1.Value = 100
'        cmdExit.Caption = "Done"
    End If
    cmdExit.SetFocus
    'If you Send Again, then Clear Attachments Now
'Moved to PBEMail1_SendCompleted 3/30/04
'    PBEmail1.Attachments.Clear
'    mWriteEmailFile
'    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
     mInit
End Sub


Private Sub mInit()
    If Trim$(tgCurrUIE(0).sShowName) <> "" Then
        txtFromName.text = Trim$(tgCurrUIE(0).sShowName)
    Else
        txtFromName.text = Trim$(tgCurrUIE(0).sSignOnName)
    End If

    mReadEmailFile
    'cbcToName.Text = "Counterpoint Service"

    If sgFileAttachment <> "" Then
        txtAttachments.text = sgFileAttachmentName
        txtSubject.text = Trim$(tgSOE.sClientName) & " - " & Trim$(sgFileAttachmentName)
        If PBEmail1.Attachments.Add(sgFileAttachment) = False Then
            MsgBox "Unable to Add Attachment!", vbInformation, "Error"
        End If
    Else
        txtSubject.text = Trim$(tgSOE.sClientName)
    End If

End Sub

Private Sub mWriteEmailFile()

    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject
    Dim slLocation As String

    slLocation = sgMsgDirectory & "CSIEmail" & CStr(tgCurrUIE(0).iCode) & ".txt"

    If fs.FileExists(slLocation) Then
        fs.DeleteFile slLocation
        DoEvents
    End If

    fs.CreateTextFile slLocation, False
    DoEvents
    Set tlTxtStream = fs.OpenTextFile(slLocation, ForWriting, False)
    DoEvents
    tlTxtStream.WriteLine (txtFromEmail.text)
    DoEvents
    tlTxtStream.Close
    DoEvents

End Sub

Private Sub mReadEmailFile()

    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject
    Dim slLocation As String

    slLocation = sgMsgDirectory & "CSIEmail" & CStr(tgCurrUIE(0).iCode) & ".txt"

    If fs.FileExists(slLocation) Then
        Set tlTxtStream = fs.OpenTextFile(slLocation, ForReading, False)
        'right now there is only one line, but we are ready for more
        Do While tlTxtStream.AtEndOfStream <> True
            txtFromEmail.text = tlTxtStream.ReadLine
        Loop
        tlTxtStream.Close
    Else
        txtFromEmail.text = ""
    End If
End Sub
