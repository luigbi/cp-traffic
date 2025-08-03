VERSION 5.00
Object = "{0E9D0E41-7AB8-11D1-9400-00A0248F2EF0}#1.0#0"; "dzactx.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EmailGeneric 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5715
   ScaleWidth      =   9510
   Begin VB.TextBox txtToName 
      Height          =   345
      Left            =   1155
      TabIndex        =   3
      Top             =   540
      Width           =   3375
   End
   Begin VB.ListBox lbcSendings 
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   885
      TabIndex        =   18
      Top             =   4785
      Visible         =   0   'False
      Width           =   8445
   End
   Begin VB.ComboBox cbcToName 
      Height          =   315
      Left            =   1170
      TabIndex        =   11
      Top             =   570
      Width           =   3390
   End
   Begin VB.CommandButton cmdClrAttach 
      Caption         =   "Clear Attachments"
      Height          =   330
      Left            =   150
      TabIndex        =   8
      Top             =   4350
      Width           =   1500
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
      Top             =   4230
      Width           =   2010
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7455
      Top             =   4155
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtAttachments 
      Height          =   315
      Left            =   1740
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3570
      Width           =   7605
   End
   Begin VB.TextBox txtSubject 
      Height          =   315
      Left            =   1185
      TabIndex        =   5
      Top             =   1140
      Width           =   8130
   End
   Begin VB.TextBox txtMsgBox 
      Height          =   1800
      Left            =   1185
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1650
      Width           =   8175
   End
   Begin VB.CommandButton cmdAttach 
      Caption         =   "Add Attachments"
      Height          =   300
      Left            =   150
      TabIndex        =   7
      Top             =   3930
      Width           =   1500
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
      Top             =   4230
      Width           =   2010
   End
   Begin VB.TextBox txtFromEmail 
      Height          =   315
      Left            =   5910
      TabIndex        =   2
      Top             =   150
      Width           =   3375
   End
   Begin VB.TextBox txtToEmail 
      Height          =   315
      Left            =   5925
      TabIndex        =   4
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox txtFromName 
      Height          =   315
      Left            =   1185
      TabIndex        =   1
      Top             =   135
      Width           =   3390
   End
   Begin VB.Image imcSpellCheck 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   690
      Picture         =   "emailGeneric.frx":0000
      ToolTipText     =   "Check Spelling"
      Top             =   3090
      Width           =   360
   End
   Begin DZACTXLibCtl.dzactxctrl zpcDZip 
      Left            =   8205
      OleObjectBlob   =   "emailGeneric.frx":0672
      Top             =   4095
   End
   Begin VB.Label lblResults 
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
      TabIndex        =   19
      Top             =   4815
      Visible         =   0   'False
      Width           =   720
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
      TabIndex        =   17
      Top             =   3570
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
      TabIndex        =   16
      Top             =   600
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
      TabIndex        =   15
      Top             =   615
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
      TabIndex        =   14
      Top             =   165
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
      TabIndex        =   13
      Top             =   165
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
      TabIndex        =   12
      Top             =   1140
      Width           =   855
   End
End
Attribute VB_Name = "EmailGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim omEmailer As CEmail
'7677
Private Const EMPLOYEES = 1
Dim tmEmpEmails(EMPLOYEES - 1) As EMPLOYEEEMAILS
Dim bmAddFromAddress As Boolean
Dim bmIsZip As Boolean
Dim bmIsService As Boolean
'create a listbox of 'to' that is counterpoint service people? default is false
Public Property Let isCounterpointService(ByVal blValue As Boolean)
    bmIsService = blValue
End Property
'zip attachments? default is false
Public Property Let isZipAttachment(ByVal blValue As Boolean)
    bmIsZip = blValue
End Property





Private Sub Form_Unload(Cancel As Integer)
    'no email address for user?
    If bmAddFromAddress And LenB(Trim(txtFromEmail)) > 0 Then 'And bmCefOpened Then
       omEmailer.SetEmailAddressThisUser (txtFromEmail.Text)
    End If
    Erase tmEmpEmails
    'this will return values to ogEmailer
    With omEmailer
        .ToAddress = Trim(txtToEmail.Text)
        .ToName = Trim(txtToName.Text)
        .FromAddress = Trim(txtFromEmail.Text)
        .FromName = Trim(txtFromName.Text)
    End With
    'if ogEmailer is set, the values above will remain!
    Set omEmailer = Nothing
    Set EmailGeneric = Nothing
   ' Unload EmailGeneric

End Sub

Private Sub imcSpellCheck_Click()
    omEmailer.SpellCheckUsingMSWord txtMsgBox
End Sub

Private Sub zpcDZip_ZipMajorStatus(ItemName As String, Percent As Long, Cancel As Long)
    lbcSendings.List(0) = "Zipping... " & ItemName & "  " & Percent & " % "

End Sub
Private Sub zpcDZip_ZipMinorStatus(ItemName As String, Percent As Long, Cancel As Long)

    lbcSendings.List(0) = "Zipping... " & ItemName & "  " & Percent & " % "

End Sub

Sub mShowResults(flag As Boolean)

    lbcSendings.Visible = flag
    lblResults.Visible = flag

End Sub
Sub mLoadCbcToName()
    Dim llCount As Long

'    tmEmpEmails(0).Name = "Nora Browne"
'    tmEmpEmails(1).Name = "Anna Dow"
'    tmEmpEmails(2).Name = "Eric Garst"
'    tmEmpEmails(3).Name = "Mary Nelson"
'    tmEmpEmails(4).Name = "Melinda Lachance Pomeroy"
'    tmEmpEmails(5).Name = "Sara Schones"
'    tmEmpEmails(6).Name = "Brooke Townsend"
'    tmEmpEmails(7).Name = "Support"
'
'    tmEmpEmails(0).Email = "norabrowne@counterpoint.net"
'    tmEmpEmails(1).Email = "annadow@counterpoint.net"
'    tmEmpEmails(2).Email = "ericgarst@counterpoint.net"
'    tmEmpEmails(3).Email = "marynelson@counterpoint.net"
'    tmEmpEmails(4).Email = "melindalachance@counterpoint.net"
'    tmEmpEmails(5).Email = "saraschones@counterpoint.net"
'    tmEmpEmails(6).Email = "brooketownsend@counterpoint.net"
'    tmEmpEmails(7).Email = "support@counterpoint.net"
    '7677
    tmEmpEmails(0).Name = "Counterpoint Support"
    tmEmpEmails(0).Email = "support@counterpoint.net"
    For llCount = 0 To EMPLOYEES - 1
      cbcToName.AddItem tmEmpEmails(llCount).Name
    Next

End Sub
Private Sub cbcToName_Click()
    txtToEmail.Text = tmEmpEmails(cbcToName.ListIndex).Email
    txtToName.Text = cbcToName.Text
End Sub

Private Sub cmdAttach_Click()
    Dim ilLoop As Integer
    
    CommonDialog1.Filter = "All Files|*.*"    'Setup the CommonDialog
    CommonDialog1.ShowOpen 'Show the Open Dialog
    If CommonDialog1.fileName <> "" Then
        If LenB(Trim(txtAttachments)) > 0 Then
            txtAttachments.Text = txtAttachments.Text & ";" & CommonDialog1.fileName
        Else
            txtAttachments.Text = CommonDialog1.fileName
        End If
    End If
    omEmailer.ChDrDir
End Sub

Private Sub cmdClrAttach_Click()
    txtAttachments.Text = ""    'Clear the Attachment Text

End Sub

Private Sub cmdExit_Click()
    Unload EmailGeneric
End Sub
Private Sub cmdSend_Click()
    
    Screen.MousePointer = vbHourglass
    lbcSendings.Clear
    lbcSendings.ForeColor = vbBlack
    DoEvents
    If txtFromEmail.Text = "" Then
        MsgBox "You must enter Your Email Address! ", vbCritical
        txtFromEmail.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If txtToEmail.Text = "" Then
        MsgBox "You must enter the Recipient's Email Address! ", vbCritical
        txtToEmail.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    mShowResults True
    With omEmailer
        .FromAddress = txtFromEmail
        .FromName = txtFromName
'        .ToAddress = txtToEmail
'        .ToName = txtToName.Text
        .AddTOAddress txtToEmail.Text, txtToName.Text
        .Subject = txtSubject
        .Message = txtMsgBox
        .Attachment = txtAttachments
        If bmIsZip Then
            .Success = .Send(lbcSendings, zpcDZip)
        Else
            .Success = .Send(lbcSendings)
        End If
    End With
    If omEmailer.Success Then
        lbcSendings.AddItem omEmailer.AdditionalMessageIfSuccess
    Else
        lbcSendings.AddItem omEmailer.AdditionalMessageIfFail
    End If
    '8245 from false, false
    omEmailer.Clear False, True
    cmdExit.Caption = "Done"
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
    mLoadCbcToName
    mInit
    EmailGeneric.Move (Screen.Width - EmailGeneric.Width) \ 2, (Screen.Height - EmailGeneric.Height) \ 2 + 115
End Sub


Private Sub mInit()
    Dim ilRet As Integer
    
    If ogEmailer Is Nothing Then
        Set omEmailer = New CEmail
    Else
        Set omEmailer = ogEmailer
    End If
    txtFromName.Text = omEmailer.UserName
    txtFromEmail.Text = omEmailer.UserEmailAddress
    If Len(txtFromEmail.Text) = 0 Then
        bmAddFromAddress = True
    Else
        bmAddFromAddress = False
    End If
    If bmIsService Then
        cbcToName.Visible = True
        txtToName.Visible = False
        cbcToName.TabIndex = 3
    Else
        cbcToName.Visible = False
        txtToName.Visible = True
        txtToName.Text = omEmailer.ToName
        txtToEmail.Text = omEmailer.ToAddress
    End If
    txtAttachments.Text = omEmailer.Attachment
    txtSubject.Text = omEmailer.Subject
    txtMsgBox.Text = omEmailer.Message
    Screen.MousePointer = vbDefault
End Sub
