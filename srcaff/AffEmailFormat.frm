VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmAffEmailFormat 
   Caption         =   "Check Email Format"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7350
   Begin VB.TextBox txtResult 
      Height          =   3150
      Left            =   420
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   270
      Width           =   6570
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4035
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   4680
      FormDesignWidth =   7350
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3975
      TabIndex        =   1
      Top             =   4065
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Check"
      Height          =   375
      Left            =   2025
      TabIndex        =   0
      Top             =   4065
      Width           =   1335
   End
   Begin VB.Label labCount 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   315
      Width           =   3240
   End
End
Attribute VB_Name = "frmAffEmailFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const LOGFILE As String = ""
Private Const FORMNAME As String = "frmAffEMailFormat"

Private Sub cmdCancel_Click()
    Unload frmAffEmailFormat
End Sub

Private Sub cmdOk_Click()
    Screen.MousePointer = vbHourglass
    txtResult.Text = ""
    mTestFormat
    cmdOK.Enabled = False
    cmdCancel.Caption = "Done"
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / 3
    Me.Height = (Screen.Height) / 3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Screen.MousePointer = vbDefault
    txtResult.Text = "Check station email addresses to make sure they are in a valid format.  That is, the addresses fit the requirements to be an email address." _
    & "  This does not test to see if the email addresses are actually being used."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAffEmailFormat = Nothing
End Sub
Private Sub mTestFormat()
    '9938
   ' Dim myEMailer As CsiNetUtilities.CsiEmailer
    Dim myEmailer As CEmail

    Const INVALIDFILE As String = "EmailFormatImproper.txt"

  '  Dim llRet As Long
    Dim slAddress As String
    Dim rst As ADODB.Recordset
    Dim SQLQuery As String
  '  Dim slPreSql As String
  '  Dim ilCounter As Integer
    Dim llCountTotal As Long
    Dim llCountBad As Long
  '  Dim slPossibleRight As String
  '  Dim blDelete As Boolean
    Dim slResult As String
    '9938
    Dim slErrorMessage As String
    Dim slName As String
    Dim slMissingEmailMessage As String
    
    Screen.MousePointer = vbHourglass
    '9938
    slMissingEmailMessage = ""
    Set myEmailer = New CEmail
'    On Error GoTo errbox
'       'Set myEmailer = New CsiNetUtilities.CsiEmailer
'    On Error GoTo 0
    txtResult = "invalid"
    llCountTotal = 0
    llCountBad = 0
  '  slPreSql = "select count(*) as amount from artt where arttEmail <> ''"
    '9938 now check blank emails  also added arttType = 'P'
    'SQLQuery = "SELECT shttcallletters as letters , arttemail as email, arttLastName as last, arttFirstName as first, arttCode as code FROM artt inner join shtt on arttshttcode = shttcode order by shttCallLetters"
    SQLQuery = "SELECT shttcallletters as letters, arttemail as email, arttLastName as last, arttFirstName as first, arttCode as code, arttISCI2Contact as Isci,arttWebEmail as web FROM artt inner join shtt on arttshttcode = shttcode where arttType = 'P' order by shttCallLetters "
    On Error GoTo ERRBADSQL
   ' Set rst = gSQLSelectCall(slPreSql)
    'slPossibleRight = rst!amount
    'rst.Close
    Set rst = gSQLSelectCall(SQLQuery)
    On Error GoTo 0
    Do While Not rst.EOF
        slErrorMessage = ""
        slAddress = Trim(rst!eMail)
        If Len(slAddress) > 0 Then
            '9938
            llCountTotal = llCountTotal + 1
            If Not myEmailer.TestAddress(slAddress) Then
                slErrorMessage = myEmailer.ErrorMessage
                If Len(slErrorMessage) = 0 Then
                    slErrorMessage = " Issues not defined."
                Else
                    slErrorMessage = " Issue(s): " & slErrorMessage
                End If
                slName = Trim(rst!Last)
                If Len(slName) > 0 Then
                    If Len(Trim(rst!First)) > 0 Then
                        slName = slName & "," & Trim(rst!First)
                    End If
                Else
                    slName = Trim(rst!First)
                End If
                txtResult.Text = txtResult.Text & vbCrLf & Trim(rst!letters) & ": " & slName & " Address: " & slAddress & slErrorMessage
                llCountBad = llCountBad + 1
            End If
'            With myEmailer
'                llRet = .ValidateAddress(slAddress)
'             End With
'            If llRet = 1 Then
'                txtResult.Text = txtResult.Text & vbCrLf & Trim(rst!letters) & ": " & Trim(rst!Last) & "," & Trim(rst!First) & " " & slAddress
'                llCountBad = llCountBad + 1
'            Else
'                '8345  get rid of space at end!
'                If mTestForScrewyChars(slAddress) Then
'                    txtResult.Text = txtResult.Text & vbCrLf & Trim(rst!letters) & ": " & Trim(rst!Last) & "," & Trim(rst!First) & " " & slAddress & " Space at end is really an unwritable character.  Please remove"
'                    llCountBad = llCountBad + 1
'                Else
'                    llCountTotal = llCountTotal + 1
'                End If
'            End If
        Else
            '9938
            If rst!Web = "Y" Then
                slName = Trim(rst!Last)
                If Len(slName) > 0 Then
                    If Len(Trim(rst!First)) > 0 Then
                        slName = slName & "," & Trim(rst!First)
                    End If
                Else
                    slName = Trim(rst!First)
                End If
                slMissingEmailMessage = slMissingEmailMessage & vbCrLf & Trim(rst!letters) & ": " & slName & " defined as web contact but missing email. "
            End If
            If rst!ISCI = 1 Then
                 slName = Trim(rst!Last)
                If Len(slName) > 0 Then
                    If Len(Trim(rst!First)) > 0 Then
                        slName = slName & "," & Trim(rst!First)
                    End If
                Else
                    slName = Trim(rst!First)
                End If
                slMissingEmailMessage = slMissingEmailMessage & vbCrLf & Trim(rst!letters) & ": " & slName & " defined as ISCI contact but missing email. "
           
            End If
        End If
        rst.MoveNext
    Loop
    rst.Close
    txtResult.Text = txtResult & vbCrLf & llCountBad & " format issues of " & llCountTotal & " email addresses."
    If Len(slMissingEmailMessage) > 0 Then
        txtResult.Text = txtResult.Text & vbCrLf & "Undefined" & slMissingEmailMessage
    End If
   ' slResult = txtResult & vbCrLf & llCountBad & " format issues of " & slPossibleRight & " email addresses." & vbCrLf
    slResult = txtResult.Text
    'write to file
    gLogMsg vbCrLf & slResult, INVALIDFILE, True
    txtResult = slResult & vbCrLf & "File saved as " & INVALIDFILE & " in messages folder." & vbCrLf
Cleanup:
    Screen.MousePointer = vbDefault
    Set myEmailer = Nothing
    If Not rst Is Nothing Then
        If (rst.State And adStateOpen) <> 0 Then
            rst.Close
            Set rst = Nothing
        End If
    End If
    Exit Sub
'ERRDSN:
'    gLogMsg "That dsn name did not allow a successful connection", "affErrorLog.txt", False
'    txtResult.Text = "Error. That dsn name did not allow a successful connection"
'    GoTo Cleanup
'errbox:
'    gLogMsg "Couldn't find csiNetUtilities", "affErrorLog.txt", False
'    txtResult.Text = "Error. Couldn't find csiNetUtilities"
'    GoTo Cleanup
ERRBADSQL:
    gHandleError LOGFILE, FORMNAME & "-mtestformat"
    txtResult.Text = "problem testing email formats."
    GoTo Cleanup
End Sub
'9938
'Private Function mTestForScrewyChars(slString As String) As Boolean
'    Dim blRet As Boolean
'    Dim c As Integer
'    Dim ilAsc As Integer
'
'    blRet = False
'    For c = Len(slString) To 1 Step -1
'        ilAsc = Asc(Mid(slString, c, 1))
'        If ilAsc > 126 Then
'            blRet = True
'            Exit For
'        End If
'    Next c
'    mTestForScrewyChars = blRet
'End Function
