VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EmailProposal 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6585
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
   ScaleHeight     =   6585
   ScaleWidth      =   9510
   Begin VB.CheckBox ckcSupress 
      Caption         =   "Suppress Rates"
      Height          =   255
      Left            =   6600
      TabIndex        =   22
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdAttach 
      Caption         =   "Browse"
      Height          =   300
      Left            =   8640
      TabIndex        =   20
      Top             =   4080
      Width           =   780
   End
   Begin VB.TextBox txtAttachments 
      Height          =   315
      Left            =   2070
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4080
      Width           =   6525
   End
   Begin VB.OptionButton OptOutput 
      Caption         =   "Xml"
      Height          =   315
      Index           =   2
      Left            =   4785
      TabIndex        =   18
      Top             =   3555
      Width           =   1350
   End
   Begin VB.OptionButton OptOutput 
      Caption         =   "OMD"
      Height          =   315
      Index           =   1
      Left            =   3390
      TabIndex        =   17
      Top             =   3555
      Width           =   1350
   End
   Begin VB.OptionButton OptOutput 
      Caption         =   "Csi"
      Height          =   315
      Index           =   0
      Left            =   1995
      TabIndex        =   16
      Top             =   3555
      Width           =   1350
   End
   Begin VB.TextBox txtToName 
      Height          =   345
      Left            =   1200
      TabIndex        =   2
      Top             =   525
      Width           =   3375
   End
   Begin VB.ListBox lbcSendings 
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   1005
      TabIndex        =   14
      Top             =   5625
      Visible         =   0   'False
      Width           =   8445
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
      Left            =   5295
      TabIndex        =   7
      Top             =   5070
      Width           =   2010
   End
   Begin VB.TextBox txtSubject 
      Height          =   315
      Left            =   1185
      TabIndex        =   4
      Top             =   1140
      Width           =   8130
   End
   Begin VB.TextBox txtMsgBox 
      Height          =   1800
      Left            =   1185
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1650
      Width           =   8175
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
      Left            =   2955
      TabIndex        =   6
      Top             =   5070
      Width           =   2010
   End
   Begin VB.TextBox txtFromEmail 
      Height          =   315
      Left            =   5910
      TabIndex        =   1
      Top             =   150
      Width           =   3375
   End
   Begin VB.TextBox txtToEmail 
      Height          =   315
      Left            =   5925
      TabIndex        =   3
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox txtFromName 
      Height          =   315
      Left            =   1185
      TabIndex        =   0
      Top             =   135
      Width           =   3390
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7920
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LabAttachMore 
      Caption         =   "Other Attachments:"
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
      Left            =   375
      TabIndex        =   21
      Top             =   4080
      Width           =   1665
   End
   Begin VB.Image imcSpellCheck 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   645
      Picture         =   "emailProposal.frx":0000
      ToolTipText     =   "Check Spelling"
      Top             =   2910
      Width           =   360
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
      Left            =   210
      TabIndex        =   15
      Top             =   5655
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblAttachments 
      Caption         =   "Proposal:"
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   1140
      Width           =   855
   End
End
Attribute VB_Name = "EmailProposal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim omEmailer As CEmail
Dim bmAddFromAddress As Boolean
Dim smAttachment As String
Dim bmSkipChangingAgency As Boolean
'7010
Dim bmAllowChangingAgency As Boolean
'pass in this value
Public AgencyCode As Integer
'pass in this value..then use and change if default changed
Public AgencyOutput As String
'pass this back to contract...file to be created.
Dim smCreateThisOutput As String
Private Const XML = 2
Private Const OMD = 1
Private Const CSI = 0

Private Sub cmdAttach_Click()
    Dim ilLoop As Integer
    '8225
    CommonDialog1.Filter = "All Files|*.*"    'Setup the CommonDialog
    CommonDialog1.InitDir = sgExportPath
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

Private Sub Form_Unload(Cancel As Integer)
    If Not omEmailer Is Nothing Then
        'no email address for user?
        If bmAddFromAddress And LenB(Trim(txtFromEmail)) > 0 Then 'And bmCefOpened Then
            omEmailer.SetEmailAddressThisUser (txtFromEmail.Text)
        End If
        'this will return values to ogEmailer
        With omEmailer
            .ToAddress = Trim(txtToEmail.Text)
            .ToName = Trim(txtToName.Text)
            .FromAddress = Trim(txtFromEmail.Text)
            .FromName = Trim(txtFromName.Text)
        End With
    End If
    'if ogEmailer is set, the values above will remain!
    Set omEmailer = Nothing
    Unload EmailProposal

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

Private Sub cmdExit_Click()
    Unload EmailProposal
End Sub
Private Sub cmdSend_Click()
    Dim slAttachment As String
    '5720
    Dim blUnload As Boolean
    '8940
    Dim blSupressRates As Boolean
    
    blSupressRates = False
    blUnload = False
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
    'now create the attachment! pass the possibly new output type.
    '8364
    'slAttachment = Contract.mExportProposal(smCreateThisOutput)
    '8940
    If ckcSupress.Enabled And ckcSupress.Value = vbChecked Then
        blSupressRates = True
    End If
    slAttachment = mExportProposal(smCreateThisOutput, False, False, blSupressRates)
    If Len(slAttachment) > 0 Then
        '8225
        If Len(txtAttachments.Text) > 0 Then
            slAttachment = slAttachment & ";" & txtAttachments.Text
        End If
        With omEmailer
            .FromAddress = txtFromEmail
            .FromName = txtFromName
            .AddTOAddress txtToEmail.Text, txtToName.Text
            .Subject = txtSubject
            .Message = txtMsgBox
            .Attachment = slAttachment
            .Success = .Send(lbcSendings)
        End With
        If Not omEmailer.Success Then
            lbcSendings.ForeColor = vbRed
            lbcSendings.AddItem "Please manually email Export File: " & slAttachment
        Else
            '5720
            blUnload = True
        End If
    Else
        lbcSendings.ForeColor = vbRed
        lbcSendings.AddItem "Couldn't create the export file.  Email not sent"
    End If
    omEmailer.Clear False, False
    cmdExit.Caption = "Done"
    Screen.MousePointer = vbDefault
    If blUnload Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    gCenterStdAlone EmailProposal
    mInit

End Sub


Private Sub mInit()
    Dim c As Integer
    
    '8940
    ckcSupress.Enabled = False
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
    txtToName.Text = omEmailer.ToName
    txtToEmail.Text = omEmailer.ToAddress
    smAttachment = omEmailer.Attachment
    txtSubject.Text = omEmailer.Subject
    txtMsgBox.Text = omEmailer.Message
    'agency output is only O, X, or blank(ignore c-could be anything other than o or x)
'    If AgencyOutput <> "O" And AgencyOutput <> "X" Then
    '6236 adv comes in as Z, change to X
    If AgencyOutput <> "O" And AgencyOutput <> "X" And AgencyOutput <> "Z" Then
        AgencyOutput = ""
    End If
    mSetOutputOptions
        ' 7010
    If (igWinStatus(AGENCIESLIST) <> 2) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        bmAllowChangingAgency = False
    Else
        bmAllowChangingAgency = True
    End If
        '4997
'    If (igWinStatus(AGENCIESLIST) <> 2) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
'        For c = OptOutput.LBound To OptOutput.UBound
'            OptOutput(c).Enabled = False
'        Next c
'    End If
    '6236 advertiser?
    If AgencyOutput = "Z" Then
        AgencyOutput = "X"
        For c = OptOutput.LBound To OptOutput.UBound
            OptOutput(c).Enabled = False
        Next c
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub mSetOutputOptions()
    'default to xml
    bmSkipChangingAgency = True
    If AgencyCode > 0 Then
        If AgencyOutput = "X" Then
            OptOutput(XML).Value = True
        ElseIf AgencyOutput = "O" Then
            OptOutput(OMD).Value = True
        Else
            OptOutput(CSI).Value = True
        End If
    Else
        OptOutput(XML).Value = True
    End If
    bmSkipChangingAgency = False
    ' file to create same as default
    If Len(AgencyOutput) > 0 Then
        smCreateThisOutput = AgencyOutput
    Else
        smCreateThisOutput = "C"
    End If
End Sub

Private Sub OptOutput_Click(Index As Integer)
    Dim blRet As Boolean
    Dim slOutputChoice As String
    Dim slCompare As String
    
   '7010
    Select Case Index
        Case 1
            slOutputChoice = "OMD"
            slCompare = "O"
        Case 2
            slOutputChoice = "Xml"
            slCompare = "X"
        Case Else
            slOutputChoice = "Csi"
            slCompare = ""
    End Select
    'change default agency output? 7010 added bmAllowChangingAgency
    If Not bmSkipChangingAgency And AgencyCode > 0 And bmAllowChangingAgency Then
        If slCompare <> AgencyOutput Then
            If MsgBox("Do you wish to permanently change the agency's output to " & slOutputChoice & "?", vbYesNo, "Change Agency's Default Output?") = vbYes Then
                If mChangeAgency(slCompare) Then
                    'new default. Use to compare if come back here later
                    AgencyOutput = slCompare
                Else
                    MsgBox "Couldn't change Agency to " & slCompare, vbOKOnly, "Error"
                End If
            End If
        End If
    End If
    '8940
    If slCompare = "X" Then
        ckcSupress.Enabled = True
    Else
        ckcSupress.Enabled = False
        ckcSupress.Value = vbUnchecked
    End If
    'output to create
    smCreateThisOutput = slCompare

'    'change default agency output?
'    If Not bmSkipChangingAgency And AgencyCode > 0 Then
'        Select Case Index
'            Case 1
'                slOutputChoice = "OMD"
'                slCompare = "O"
'            Case 2
'                slOutputChoice = "Xml"
'                slCompare = "X"
'            Case Else
'                slOutputChoice = "Csi"
'                slCompare = ""
'        End Select
'        If slCompare <> AgencyOutput Then
'            If MsgBox("Do you wish to permanently change the agency's output to " & slOutputChoice & "?", vbYesNo, "Change Agency's Default Output?") = vbYes Then
'                If mChangeAgency(slCompare) Then
'                    'new default. Use to compare if come back here later
'                    AgencyOutput = slCompare
'                Else
'                    MsgBox "Couldn't change Agency to " & slCompare, vbOKOnly, "Error"
'                End If
'            End If
'            'output to create
'            smCreateThisOutput = slCompare
'        End If
'    End If
End Sub
Private Function mChangeAgency(ByVal slNewOutput As String) As Boolean
    Dim hlAgf As Integer
    Dim ilRet As Integer
    Dim ilAgfRecLen As Integer
    Dim tlAgfSrchKey As INTKEY0
    Dim tlagf As AGF
    Dim blRet As Boolean
    
    If Len(slNewOutput) = 0 Then
        slNewOutput = "C"
    End If
    hlAgf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mBtrErr
    gBtrvErrorMsg ilRet, "mChangeAgency (btrOpen: Agf.Btr)", EmailProposal
    On Error GoTo 0
    ilAgfRecLen = Len(tlagf)
    tlAgfSrchKey.iCode = tgChfCntr.iAgfCode
    ilRet = btrGetEqual(hlAgf, tlagf, ilAgfRecLen, tlAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        tlagf.sCntrExptForm = slNewOutput
        ilRet = btrUpdate(hlAgf, tlagf, ilAgfRecLen)
        On Error GoTo mBtrErr
        gBtrvErrorMsg ilRet, "mChangeAgency (btrUpdate: Agf.Btr)", EmailProposal
        On Error GoTo 0
        blRet = True
    End If
    ilRet = btrClose(hlAgf)
    btrDestroy hlAgf
    mChangeAgency = blRet
    Exit Function
mBtrErr:
    mChangeAgency = False
End Function
