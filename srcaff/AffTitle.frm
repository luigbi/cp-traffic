VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmTitle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Title"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   ControlBox      =   0   'False
   Icon            =   "AffTitle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin V81Affiliate.CSI_ComboBoxList CSI_ComboBoxList1 
      Height          =   345
      Left            =   225
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   609
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   435
      Left            =   3945
      TabIndex        =   6
      Top             =   195
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   435
      Left            =   2377
      TabIndex        =   5
      Top             =   1995
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   4080
      TabIndex        =   4
      Top             =   1995
      Width           =   1320
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   435
      Left            =   675
      TabIndex        =   3
      Top             =   1995
      Width           =   1320
   End
   Begin VB.TextBox txtTitle 
      Height          =   330
      Left            =   900
      TabIndex        =   1
      Top             =   1245
      Width           =   3210
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5205
      Top             =   1290
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2625
      FormDesignWidth =   5760
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   225
      Left            =   225
      TabIndex        =   2
      Top             =   1305
      Width           =   630
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private imtntCode As Integer
Private bmIsDirty As Boolean
Private bmIgnoreChange As Boolean

Private Sub cmdDelete_Click()
    Dim ilResponse As Integer
    Dim rst As ADODB.Recordset
    Dim ilCode As Integer
    
    If Trim(txtTitle.Text) = "" Then
        txtTitle.SetFocus
        Exit Sub
    End If
    If Left(txtTitle.Text, 5) = "[New]" Then
        txtTitle.SetFocus
        Exit Sub
    End If
    If CSI_ComboBoxList1.ListIndex > 0 Then
        ilCode = CSI_ComboBoxList1.GetItemData(CSI_ComboBoxList1.ListIndex)
        SQLQuery = "SELECT * FROM artt WHERE"
        SQLQuery = SQLQuery & " arttTntCode = " & ilCode
        Set rst = gSQLSelectCall(SQLQuery)
        '3/23/16: correct test
        'If rst.EOF Then
        If Not rst.EOF Then
            ilResponse = gMsgBox("Unable to delete " & txtTitle.Text & " as it in use" & vbCrLf & vbCrLf, vbOK)
            Exit Sub
        End If
    End If
    ilResponse = gMsgBox("Are you sure you want to delete " & txtTitle.Text & vbCrLf & vbCrLf, vbYesNo)

    On Error Resume Next
    bmIgnoreChange = True
    If ilResponse = vbYes Then
        SQLQuery = "Delete From Tnt Where tntTitle = '" & Trim(txtTitle.Text) & "'"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "Title-cmdDelete_Click"
            Exit Sub
        End If
        If txtTitle.Text = sgTitle Then
            ' The title we started with got deleted so clear this item too.
            sgTitle = ""
            txtTitle.Text = ""
        End If
        mGetTitles
    End If
    CSI_ComboBoxList1.SelText ("[New]")
    bmIgnoreChange = False
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmTitle-cmcDelete"
    Exit Sub
End Sub

Private Sub cmdDone_Click()
    If Trim(txtTitle.Text) = "" Then
        gMsgBox "The title cannot be blank"
        txtTitle.SetFocus
        Exit Sub
    End If
    If Left(txtTitle.Text, 5) = "[New]" Then
        gMsgBox "The title cannot start with the word [New]"
        txtTitle.SetFocus
        Exit Sub
    End If
    sgTitle = Trim(txtTitle.Text)
    mSave
    Unload frmTitle
End Sub

Private Sub cmdSave_Click()
    If Trim(txtTitle.Text) = "" Then
        gMsgBox "The title cannot be blank"
        txtTitle.SetFocus
        Exit Sub
    End If
    If Left(txtTitle.Text, 5) = "[New]" Then
        gMsgBox "The title cannot start with the word [New]"
        txtTitle.SetFocus
        Exit Sub
    End If
    sgTitle = gFixQuote(Trim(txtTitle.Text))
    mSave
End Sub

Private Sub Form_Activate()
    bgFrmTitleCanceled = False
    txtTitle.SetFocus
    'Call CSI_ComboBoxList1.SetFont("Arial", 16)
    Call CSI_ComboBoxList1.SetDropDownNumRows(6)
    Call CSI_ComboBoxList1.SetDropDownCharWidth(20)
    mGetTitles
    If Trim(sgTitle) = "" Then
        'txtTitle.Text = "[New]"
        'CSI_ComboBoxList1.SetFocus
        CSI_ComboBoxList1.SelText ("[New]")
    End If
    bmIgnoreChange = False
    cmdSave.Enabled = False
End Sub

Private Sub Form_Load()
    Me.Width = (Screen.Width) / 3
    Me.Height = (Screen.Height) / 4
    Me.Top = (Screen.Height - Me.Height) / 3
    Me.Left = (Screen.Width - Me.Width) / 3
    gSetFonts frmTitle
    gCenterForm frmTitle
    bmIsDirty = False
    CSI_ComboBoxList1.SetFont txtTitle.FontName, txtTitle.FontSize
End Sub

Private Sub mGetTitles()
    Dim rstTitles As ADODB.Recordset

    bmIgnoreChange = True
    txtTitle.Text = ""
    CSI_ComboBoxList1.Clear
    CSI_ComboBoxList1.AddItem ("[New]")
    CSI_ComboBoxList1.SetItemData = -1  ' Indicates a new title

    SQLQuery = "Select tntCode, tntTitle From Tnt Order By tntTitle"
    Set rstTitles = gSQLSelectCall(SQLQuery)
    While Not rstTitles.EOF
        CSI_ComboBoxList1.AddItem (Trim(rstTitles!tntTitle))
        CSI_ComboBoxList1.SetItemData = rstTitles!tntCode
        rstTitles.MoveNext
    Wend
    txtTitle.Text = Trim(sgTitle)
    CSI_ComboBoxList1.SelText (Trim(sgTitle))
    bmIgnoreChange = False
End Sub

Private Sub cmdCancel_Click()
    ' sgTitle = ""
    bgFrmTitleCanceled = True
    Unload frmTitle
End Sub

Private Sub CSI_ComboBoxList1_OnChange()
    Dim ilIdx As Integer
    
    If bmIgnoreChange Then
        ' Another section may have already said to ignore changes and if so exit now.
        Exit Sub
    End If
    bmIgnoreChange = True
    txtTitle.Text = Trim(CSI_ComboBoxList1.Text)
    sgTitle = txtTitle.Text
    ilIdx = CSI_ComboBoxList1.ListIndex
    If ilIdx < 0 Then
        bmIgnoreChange = False
        Exit Sub
    End If
    imtntCode = CSI_ComboBoxList1.GetItemData(ilIdx)
    If txtTitle.Text = "[New]" Then
        txtTitle.Text = ""
        'SendKeys "{TAB}", False
        'txtTitle.SelStart = 0
        'txtTitle.SelLength = Len(txtTitle.Text)
    End If
    bmIgnoreChange = False
End Sub

Private Function mSave()
    Dim ilIdx As Integer
    Dim slTemp As String

    If Not bmIsDirty Then
        Exit Function
    End If
    On Error GoTo ErrHand
    mSave = False
    ilIdx = CSI_ComboBoxList1.ListIndex
    If ilIdx <> -1 Then
        imtntCode = CSI_ComboBoxList1.GetItemData(ilIdx)
    End If
    'D.S 6/2/14 TTP 5870
    slTemp = sgTitle
    sgTitle = Replace(slTemp, "&", " ")
    If imtntCode = -1 Then
        SQLQuery = "Insert Into Tnt (tntTitle, tntUsfCode) Values ('" & gFixQuote(sgTitle) & "', 0)"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "Title-mSave"
            mSave = False
            Exit Function
        End If
    Else
        SQLQuery = "Update Tnt Set tntTitle = '" & gFixQuote(sgTitle) & "', tntUsfCode = 0 Where tntCode = " & imtntCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "Title-mSave"
            mSave = False
            Exit Function
        End If
    End If
    mGetTitles
    mSave = True
    bmIsDirty = False
    cmdSave.Enabled = False
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmTitle-mSave"
    Exit Function
End Function

Private Sub txtTitle_Change()
    If bmIgnoreChange Then
        Exit Sub
    End If
    bmIsDirty = True
    cmdSave.Enabled = True
End Sub
