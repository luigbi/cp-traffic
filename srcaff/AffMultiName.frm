VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmMultiName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Title"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   ControlBox      =   0   'False
   Icon            =   "AffMultiName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin V81Affiliate.CSI_ComboBoxList cbcSelect 
      Height          =   345
      Left            =   2115
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   609
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   435
      Left            =   2205
      TabIndex        =   5
      Top             =   1995
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   3915
      TabIndex        =   4
      Top             =   1995
      Width           =   1320
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   435
      Left            =   510
      TabIndex        =   3
      Top             =   1995
      Width           =   1320
   End
   Begin VB.TextBox txtName 
      Height          =   330
      Left            =   1275
      TabIndex        =   1
      Top             =   1245
      Width           =   3960
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5460
      Top             =   2100
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
   Begin VB.Label lacName 
      Caption         =   "Name"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   1305
      Width           =   780
   End
End
Attribute VB_Name = "frmMultiName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lmMntCode As Long
Private bmIsDirty As Boolean
Private bmIgnoreChange As Boolean
Private rst_mnt As ADODB.Recordset


Private Sub cmdDone_Click()
    Dim ilRet As Integer
    
    sgMultiNameName = Trim(txtName.Text)
    If Not mSave() Then
        txtName.SetFocus
        Exit Sub
    End If
    igMultiNameReturn = True
    Unload frmMultiName
End Sub

Private Sub cmdSave_Click()
    sgMultiNameName = Trim(txtName.Text)
    If Not mSave() Then
        txtName.SetFocus
        Exit Sub
    End If
    bmIgnoreChange = True
    mGetMultiNames
    bmIgnoreChange = False
End Sub

Private Sub Form_Activate()
    igMultiNameReturn = False
    txtName.SetFocus
    Call cbcSelect.SetDropDownNumRows(6)
    If sgMultiNameType = "T" Then
        Call cbcSelect.SetDropDownCharWidth(20)
    Else
        Call cbcSelect.SetDropDownCharWidth(40)
    End If
    If sgMultiNameType = "T" Then
        frmMultiName.Caption = "Territory"
    ElseIf sgMultiNameType = "C" Then
        frmMultiName.Caption = "City"
    ElseIf sgMultiNameType = "Y" Then
        frmMultiName.Caption = "County"
    ElseIf sgMultiNameType = "M" Then
        frmMultiName.Caption = "Moniker"
    ElseIf sgMultiNameType = "A" Then
        frmMultiName.Caption = "Area"
    ElseIf sgMultiNameType = "O" Then
        frmMultiName.Caption = "Operator"
    Else
        frmMultiName.Caption = ""
    End If
    mGetMultiNames
    If Trim(sgMultiNameName) = "" Then
        cbcSelect.SelText ("[New]")
        txtName.SetFocus
    End If
    bmIgnoreChange = False
    cmdSave.Enabled = False
End Sub

Private Sub Form_Load()
    Me.Width = (Screen.Width) / 3
    Me.Height = (Screen.Height) / 4
    Me.Top = (Screen.Height - Me.Height) / 3
    Me.Left = (Screen.Width - Me.Width) / 3
    gSetFonts frmMultiName
    gCenterForm frmMultiName
    bmIsDirty = False
    cbcSelect.SetFont txtName.FontName, txtName.FontSize
End Sub

Private Sub mGetMultiNames()

    bmIgnoreChange = True
    txtName.Text = ""
    cbcSelect.Clear
    cbcSelect.AddItem ("[New]")
    cbcSelect.SetItemData = -1  ' Indicates a new title

    SQLQuery = "SELECT mntCode, mntName FROM mnt WHERE mntType = '" & sgMultiNameType & "' Order By mntName"
    Set rst_mnt = gSQLSelectCall(SQLQuery)
    While Not rst_mnt.EOF
        cbcSelect.AddItem (Trim(rst_mnt!mntName))
        cbcSelect.SetItemData = rst_mnt!mntCode
        rst_mnt.MoveNext
    Wend
    txtName.Text = Trim(sgMultiNameName)
    cbcSelect.SelText (Trim(sgMultiNameName))
    If cbcSelect.ListIndex > 0 Then
        lgMultiNameCode = cbcSelect.GetItemData(cbcSelect.ListIndex)
    Else
        lgMultiNameCode = 0
    End If
    bmIgnoreChange = False
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmMultiName-mGetMultiNames"
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' sgTitle = ""
    igMultiNameReturn = False
    Unload frmMultiName
End Sub

Private Sub cbcSelect_OnChange()
    Dim ilIdx As Integer
    
    If bmIgnoreChange Then
        ' Another section may have already said to ignore changes and if so exit now.
        Exit Sub
    End If
    bmIgnoreChange = True
    txtName.Text = Trim(cbcSelect.Text)
    sgMultiNameName = txtName.Text
    ilIdx = cbcSelect.ListIndex
    If ilIdx < 0 Then
        bmIgnoreChange = False
        Exit Sub
    End If
    lgMultiNameCode = cbcSelect.GetItemData(ilIdx)
    If txtName.Text = "[New]" Then
        txtName.Text = ""
        'SendKeys "{TAB}", False
        'txtName.SelStart = 0
        'txtName.SelLength = Len(txtName.Text)
    End If
    bmIgnoreChange = False
End Sub

Private Function mSave() As Integer
    Dim ilIdx As Integer
    Dim llCode As Long
    Dim ilRet As Integer

    If Not bmIsDirty Then
        mSave = True
        Exit Function
    End If
    mSave = False
    If Trim(txtName.Text) = "" Then
        gMsgBox "The Name cannot be blank"
        Exit Function
    End If
    If Left(txtName.Text, 5) = "[New]" Then
        gMsgBox "The Name cannot start with the word [New]"
        Exit Function
    End If
    On Error GoTo ErrHand
    ilIdx = cbcSelect.ListIndex
    If ilIdx <> -1 Then
        lmMntCode = cbcSelect.GetItemData(ilIdx)
    End If
    SQLQuery = "SELECT mntCode FROM mnt WHERE mntType = '" & sgMultiNameType & "' AND mntName = '" & sgMultiNameName & "'"
    Set rst_mnt = gSQLSelectCall(SQLQuery)
    If Not rst_mnt.EOF Then
        If lmMntCode <> rst_mnt!mntCode Then
            MsgBox sgMultiNameName & " previously used, Enter a different name"
            Exit Function
        End If
    End If
    If lmMntCode = -1 Then
        Do
            SQLQuery = "SELECT MAX(mntCode) from mnt"
            Set rst_mnt = gSQLSelectCall(SQLQuery)
            If IsNull(rst_mnt(0).Value) Then
                llCode = 1
            Else
                If Not rst_mnt.EOF Then
                    llCode = rst_mnt(0).Value + 1
                Else
                    llCode = 1
                End If
            End If
            ilRet = 0
            SQLQuery = "Insert Into mnt ( "
            SQLQuery = SQLQuery & "mntCode, "
            SQLQuery = SQLQuery & "mntType, "
            SQLQuery = SQLQuery & "mntName, "
            SQLQuery = SQLQuery & "mntState, "
            SQLQuery = SQLQuery & "mntUnused "
            SQLQuery = SQLQuery & ") "
            SQLQuery = SQLQuery & "Values ( "
            SQLQuery = SQLQuery & llCode & ", "
            SQLQuery = SQLQuery & "'" & sgMultiNameType & "', "
            SQLQuery = SQLQuery & "'" & gFixQuote(sgMultiNameName) & "', "
            SQLQuery = SQLQuery & "'" & "A" & "', "
            SQLQuery = SQLQuery & "'" & "" & "' "
            SQLQuery = SQLQuery & ") "
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand1:
                Screen.MousePointer = vbDefault
                If Not gHandleError4994("AffErrorLog.txt", "MultiName-mSave") Then
                    mSave = False
                    Exit Function
                End If
                ilRet = 1
            End If
        Loop While ilRet <> 0
        lgMultiNameCode = llCode
    Else
        SQLQuery = "Update Mnt Set mntName = '" & gFixQuote(sgMultiNameName) & "' Where mntCode = " & lmMntCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand1:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "MultiName-mSave"
            mSave = False
            Exit Function
        End If
        lgMultiNameCode = lmMntCode
    End If
    mSave = True
    bmIsDirty = False
    cmdSave.Enabled = False
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmMultiName-mSave"
ErrHand1:
    Screen.MousePointer = vbDefault
    gHandleError4994 "AffErorLog.txt", "frmMultiName-mSave"
    Exit Function
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst_mnt.Close
    Set frmMultiName = Nothing
End Sub

Private Sub txtName_Change()
    If bmIgnoreChange Then
        Exit Sub
    End If
    bmIsDirty = True
    cmdSave.Enabled = True
End Sub
