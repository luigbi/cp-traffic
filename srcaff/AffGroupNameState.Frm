VERSION 5.00
Begin VB.Form frmGroupNameState 
   Caption         =   "Group Name- State"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   Icon            =   "AffGroupNameState.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPostalName 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   2295
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1680
      Width           =   450
   End
   Begin VB.TextBox txtGroupName 
      Height          =   315
      Left            =   2295
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2250
      Width           =   2655
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   495
      TabIndex        =   7
      Top             =   2835
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2625
      TabIndex        =   8
      Top             =   2835
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4695
      TabIndex        =   9
      Top             =   2835
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   2295
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.ComboBox cboState 
      Height          =   315
      Left            =   2295
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lacPostalName 
      Caption         =   "Postal Name:"
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1875
   End
   Begin VB.Label lacGroupName 
      Caption         =   "Group Name:"
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   2250
      Width           =   2055
   End
   Begin VB.Label lacName 
      Caption         =   "Name:"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1875
   End
End
Attribute VB_Name = "frmGroupNameState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmGroupNameState
'*
'*  Created October,2005 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc. 2005
'*
'******************************************************
Option Explicit
Option Compare Text

Private smName As String
Private smGroupName As String
Private smPostalName As String

Private imInChg As Integer
Private imBSMode As Integer
Private imSave As Integer
Private imSntCode As Integer
Private imNameChange As Integer
Private imIgnoreChg As Integer
Private snt_rst As ADODB.Recordset


Private Sub cboState_Change()
    
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilLen As Integer
    Dim ilSel As Integer
    Dim llRow As Long

    If imInChg Then
        Exit Sub
    End If
    imInChg = True

    Screen.MousePointer = vbHourglass
    slName = LTrim$(cboState.Text)
    ilLen = Len(slName)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(cboState.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        On Error GoTo ErrHand
        cboState.ListIndex = llRow
        cboState.SelStart = ilLen
        cboState.SelLength = Len(cboState.Text)
        If cboState.ListIndex < 0 Then
            imSntCode = -1
        ElseIf (cboState.ListIndex = 0) And (sgStateCall <> "M") Then
            imSntCode = 0
        Else
            imSntCode = CInt(cboState.ItemData(cboState.ListIndex))
        End If
        If imSntCode < 0 Then
            mClearControls
        ElseIf (imSntCode = 0) And (sgStateCall <> "M") Then
            mClearControls
        Else                                                                'Load existing station data
            mBindControls
        End If
    End If
    Screen.MousePointer = vbDefault
    imInChg = False
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmGroupNameState-cboState_Change"
    imInChg = False
End Sub

Private Sub cboState_Click()
    If imInChg Then
        Exit Sub
    End If
    cboState_Change
End Sub

Private Sub cboState_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cboState_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboState_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboState.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    igStateReturnCode = 0
    igIndex = cboState.ListIndex
    Unload frmGroupNameState
End Sub

Private Sub cmdDone_Click()
    
    Dim ilRet As Integer

    ilRet = mSntChkForChange
    If ilRet Then
        ilRet = gMsgBox("Would You Like To Save the Changes", vbYesNo)
        If ilRet = vbYes Then
            ilRet = mSave()
        End If
        If Not ilRet Then
            Exit Sub
        End If
    End If
    igIndex = cboState.ListIndex
    igStateReturnCode = 0
    If igIndex > 0 Then
        igStateReturnCode = cboState.ItemData(igIndex)
    End If
    Unload frmGroupNameState
End Sub

Private Sub cmdSave_Click()
    Call mSave
End Sub

Private Sub Form_Activate()
    If sgStateCall <> "M" Then
        If cboState.Index >= 0 Then
            txtName.SetFocus
        End If
    End If
End Sub

Private Sub Form_Initialize()
    gSetFonts frmGroupNameState
    gCenterForm frmGroupNameState
End Sub

Private Sub Form_Load()
    
    Form_Initialize
    If (Not gWegenerExport) And (Not gOLAExport) Then
        txtGroupName.Enabled = False
        txtGroupName.Text = ""
    End If
    imSave = False
    Call mFillSntInfo("")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    ilRet = gPopStates()
    snt_rst.Close
    Set frmGroupNameState = Nothing

End Sub

Private Function mFillSntInfo(slFromSaveName As String)
    
    Dim snt_rst As ADODB.Recordset
    Dim slName As String
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    
    If imSave Then
        slName = Trim$(slFromSaveName)
    End If
    If cboState.Text = "[New]" And (Not imSave) Then
        cboState.ListIndex = 0
        mClearControls
        Exit Function
    End If
    cboState.Clear
    SQLQuery = "SELECT * FROM SNT"
    Set snt_rst = gSQLSelectCall(SQLQuery)
    While Not snt_rst.EOF
        cboState.AddItem Trim$(snt_rst!sntName)
        cboState.ItemData(cboState.NewIndex) = snt_rst!SntCode
        snt_rst.MoveNext
    Wend
    If sgStateCall <> "M" Then
        cboState.AddItem "[New]", 0
        cboState.ItemData(cboState.NewIndex) = 0
    End If

    If (sgStateCall <> "M") Or (imSave) Then
        If (sgStateCall = "N") And (Not imSave) Then
            cboState.ListIndex = 0
        Else
            If Not imSave Then
                slName = sgStateCall
            End If
            ilRow = SendMessageByString(cboState.hwnd, CB_FINDSTRING, -1, slName)
            If (ilRow > 0) Or ((ilRow = 0) And (sgStateCall = "M")) Then
                cboState.ListIndex = ilRow
                mBindControls
            Else
                mClearControls
            End If
            DoEvents
        End If
    End If
    Screen.MousePointer = vbDefault
    imInChg = False
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmGroupNameState-mFillSntInfo"
    imInChg = False
End Function

Private Function mSave()
    
    Dim ilRet As Integer
    Dim ilRow As Integer
    Dim slName As String
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    mSave = False
    ilRet = mSntChkForChange
    'slName = LTrim$(cboState.Text)
    slName = Trim$(txtName.Text)
    If Trim$(slName) = "" Then
        gMsgBox "State Name must be Defined.", vbOKOnly
        Exit Function
    End If
    'If cboState.text = "[New]" Or imNameChange Then
    '    SQLQuery = "select SntCode from SNT where SntStateAbbrName = '" & txtName.text & "'"
    '    Set rst = gSQLSelectCall(SQLQuery)
    '    If Not rst.EOF Then
    '        gMsgBox "The Format " & Trim$(txtName.text) & " name already exists"
    '        mSave = True
    '        Exit Function
    '    End If
    'End If
    
    ilRow = SendMessageByString(cboState.hwnd, CB_FINDSTRING, -1, slName)
    If ilRow >= 0 Then
        imSntCode = cboState.ItemData(ilRow)
    Else
        imSntCode = 0
    End If
    If Trim$(txtGroupName.Text) <> "" Then
        SQLQuery = "SELECT SntCode FROM snt WHERE sntGroupName = '" & Trim$(txtGroupName.Text) & "'"
        Set snt_rst = gSQLSelectCall(SQLQuery)
        If Not snt_rst.EOF Then
            If imSntCode <> snt_rst!SntCode Then
                gMsgBox "State Group Name previously defined with another state.", vbOKOnly
                Exit Function
            End If
        End If
    End If
    mSave = False
    mTrimAndFixQuotes
    If cboState.Text = "[New]" Then
        SQLQuery = "Insert into SNT "
        SQLQuery = SQLQuery & "(sntName, sntPostalName, sntGroupName, sntUstCode, sntUnused) "
        SQLQuery = SQLQuery & " VALUES ('" & txtName.Text & "','" & txtPostalName.Text & "','" & txtGroupName.Text & "'," & igUstCode & "," & "''" & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "GroupNameState-mSave"
            imInChg = False
            mSave = False
            Exit Function
        End If
        mClearControls
    Else
        SQLQuery = "UPDATE SNT"
        SQLQuery = SQLQuery & " SET sntGroupName = '" & Trim$(txtGroupName.Text) & "',"
        SQLQuery = SQLQuery & "sntUstCode = " & igUstCode
        SQLQuery = SQLQuery & " WHERE SntCode = " & imSntCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "GroupNameState-mSave"
            imInChg = False
            mSave = False
            Exit Function
        End If
    End If
    'rst.Close
    imSave = True
    ilRet = mFillSntInfo(slName)
    imSave = False
    imNameChange = False
    mSave = True
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmGroupNameState-mSave"
    imInChg = False
End Function

Private Sub mClearControls()

    imIgnoreChg = True
    txtName.Text = ""
    txtPostalName.Text = ""
    txtGroupName.Text = ""
    imIgnoreChg = False
End Sub

Private Function mSntChkForChange()
    
    On Error GoTo ErrHand
    
    mSntChkForChange = False

    If StrComp(txtName.Text, smName, vbTextCompare) <> 0 Then
        mSntChkForChange = True
        Exit Function
    End If
    If StrComp(txtGroupName.Text, smGroupName, vbTextCompare) <> 0 Then
        mSntChkForChange = True
        Exit Function
    End If
    If StrComp(txtPostalName.Text, smPostalName, vbTextCompare) <> 0 Then
        mSntChkForChange = True
        Exit Function
    End If
    
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in GroupNameState-mSntChkForChange: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    imInChg = False

End Function


Private Sub mTrimAndFixQuotes()

    On Error GoTo ErrHand
    
    txtName.Text = gFixQuote(Trim$(txtName.Text))
    imNameChange = False
Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in GroupNameState-mTrimAndFixQuotes: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    imInChg = False
End Sub



Private Sub txtGroupName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtGroupName_KeyPress(KeyAscii As Integer)
    Dim slChar As String
    
    If (KeyAscii <> KEYBACKSPACE) Then
        slChar = Chr(KeyAscii)
        If (Asc(slChar) < Asc("A")) Or (Asc(slChar) > Asc("Z")) Then
            'If (Asc(slChar) < Asc("a")) Or (Asc(slChar) > Asc("z")) Then
                If (Asc(slChar) < Asc("0")) Or (Asc(slChar) > Asc("9")) Then
                    If (Asc(slChar) <> Asc("-")) And (Asc(slChar) <> Asc("_")) Then
                        KeyAscii = 0    ' Ignore this key.
                    End If
                End If
            'End If
        End If
    End If
End Sub

Private Sub txtName_Change()
    If StrComp(txtName.Text, smName, vbTextCompare) <> 0 Then
        If Not imIgnoreChg Then
            imNameChange = True
        End If
    Else
        imNameChange = False
    End If
End Sub

Private Sub txtName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub



Private Sub mBindControls()

    Dim snt_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If cboState.Text = "[New]" Then
        Exit Sub
    End If
    
    imIgnoreChg = True
    SQLQuery = "SELECT * "
    SQLQuery = SQLQuery + " FROM SNT"
    SQLQuery = SQLQuery + " WHERE (SntCode = " & imSntCode & ")"
    Set snt_rst = gSQLSelectCall(SQLQuery)
    If snt_rst.EOF Then
        gMsgBox "No matching records were found", vbOKOnly
        mClearControls
    Else
        txtName.Text = Trim$(snt_rst!sntName)
        smName = txtName.Text
        txtPostalName.Text = Trim$(snt_rst!sntPostalName)
        smPostalName = txtPostalName.Text
        If (gWegenerExport) Or (gOLAExport) Then
            txtGroupName.Text = Trim$(snt_rst!sntGroupName)
            smGroupName = txtGroupName.Text
        Else
            txtGroupName.Text = ""
            smGroupName = txtGroupName.Text
        End If
    End If
    imIgnoreChg = False
Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmGroupNameState-mBindControls"
    imInChg = False
End Sub

Private Sub txtPostalName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
