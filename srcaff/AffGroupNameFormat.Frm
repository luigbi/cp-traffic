VERSION 5.00
Begin VB.Form frmGroupNameFormat 
   Caption         =   "Group Name- Format"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   Icon            =   "AffGroupNameFormat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGroupName 
      Height          =   315
      Left            =   2265
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1785
      Width           =   2655
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   405
      TabIndex        =   5
      Top             =   2835
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2535
      TabIndex        =   6
      Top             =   2835
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4605
      TabIndex        =   7
      Top             =   2835
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   2265
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.ComboBox cboFormat 
      Height          =   315
      Left            =   2265
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lacGroupName 
      Caption         =   "Group Name:"
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1785
      Width           =   2055
   End
   Begin VB.Label lblNAme 
      Caption         =   "Format Name:"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1875
   End
End
Attribute VB_Name = "frmGroupNameFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmGroupNameFormat
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

Private imInChg As Integer
Private imBSMode As Integer
Private imSave As Integer
Private imFmtCode As Integer
Private imNameChange As Integer
Private imIgnoreChg As Integer
Private fmt_rst As ADODB.Recordset


Private Sub cboFormat_Change()
    
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
    slName = LTrim$(cboFormat.Text)
    ilLen = Len(slName)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(cboFormat.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        On Error GoTo ErrHand
        cboFormat.ListIndex = llRow
        cboFormat.SelStart = ilLen
        cboFormat.SelLength = Len(cboFormat.Text)
        If cboFormat.ListIndex < 0 Then
            imFmtCode = -1
        ElseIf (cboFormat.ListIndex = 0) And (sgFormatCall <> "M") Then
            imFmtCode = 0
        Else
            imFmtCode = CInt(cboFormat.ItemData(cboFormat.ListIndex))
        End If
        If imFmtCode < 0 Then
            mClearControls
        ElseIf (imFmtCode = 0) And (sgFormatCall <> "M") Then
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
    gHandleError "AffErrorLog.txt", "GroupNameFormat-cboFormat_Change"
    imInChg = False
End Sub

Private Sub cboFormat_Click()
    If imInChg Then
        Exit Sub
    End If
    cboFormat_Change
End Sub

Private Sub cboFormat_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cboFormat_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboFormat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboFormat.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    igFormatReturn = False
    igFormatReturnCode = 0
    igIndex = cboFormat.ListIndex
    Unload frmGroupNameFormat
End Sub

Private Sub cmdDone_Click()
    
    Dim ilRet As Integer

    ilRet = mFmtChkForChange
    If ilRet Then
        ilRet = gMsgBox("Would You Like To Save the Changes", vbYesNo)
        If ilRet = vbYes Then
            ilRet = mSave()
        End If
        If Not ilRet Then
            Exit Sub
        End If
    End If
    igIndex = cboFormat.ListIndex
    igFormatReturnCode = 0
    If igIndex > 0 Then
        igFormatReturnCode = cboFormat.ItemData(igIndex)
    End If
    igFormatReturn = True
    Unload frmGroupNameFormat
End Sub

Private Sub cmdSave_Click()
    Call mSave
End Sub

Private Sub Form_Activate()
    If sgFormatCall <> "M" Then
        If cboFormat.ListIndex >= 0 Then
            txtName.SetFocus
        Else
            cboFormat.ListIndex = 0
            txtName.SetFocus
        End If
    End If
End Sub

Private Sub Form_Initialize()
    gSetFonts frmGroupNameFormat
    gCenterForm frmGroupNameFormat
End Sub

Private Sub Form_Load()
    
    Form_Initialize
    If (Not gWegenerExport) And (Not gOLAExport) Then
        txtGroupName.Enabled = False
        txtGroupName.Text = ""
        lacGroupName.Visible = False
        txtGroupName.Visible = False
    End If
    imSave = False
    If sgFormatCall <> "M" Then
        txtName.Locked = False
        txtName.BackColor = vbWhite
    End If
    Call mFillFmtInfo("")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    ilRet = gPopFormats()
    fmt_rst.Close
    Set frmGroupNameFormat = Nothing

End Sub

Private Function mFillFmtInfo(slFromSaveName As String)
    
    Dim fmt_rst As ADODB.Recordset
    Dim slName As String
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    
    If imSave Then
        slName = Trim$(slFromSaveName)
    End If
    If cboFormat.Text = "[New]" And (Not imSave) Then
        cboFormat.ListIndex = 0
        mClearControls
        Exit Function
    End If
    cboFormat.Clear
    SQLQuery = "SELECT * FROM Fmt_Station_Format"
    Set fmt_rst = gSQLSelectCall(SQLQuery)
    While Not fmt_rst.EOF
        cboFormat.AddItem Trim$(fmt_rst!FmtName)
        cboFormat.ItemData(cboFormat.NewIndex) = fmt_rst!FmtCode
        fmt_rst.MoveNext
    Wend
    If sgFormatCall <> "M" Then
        cboFormat.AddItem "[New]", 0
        cboFormat.ItemData(cboFormat.NewIndex) = 0
    End If

    If (sgFormatCall <> "M") Or (imSave) Then
        If (sgFormatCall = "N") And (Not imSave) Then
            cboFormat.ListIndex = 0
        Else
            If Not imSave Then
                slName = sgFormatName
            End If
            ilRow = SendMessageByString(cboFormat.hwnd, CB_FINDSTRING, -1, slName)
            If (ilRow > 0) Or ((ilRow = 0) And (sgFormatCall = "M")) Then
                cboFormat.ListIndex = ilRow
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
    gHandleError "AffErrorLog.txt", "GroupNameFormat-mFillFmtInfo"
    imInChg = False
End Function

Private Function mSave()
    
    Dim ilRet As Integer
    Dim ilRow As Integer
    Dim slName As String
    Dim ilFmtCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    mSave = False
    ilRet = mFmtChkForChange
    'slName = LTrim$(cboFormat.Text)
    slName = Trim$(txtName.Text)
    If Trim$(slName) = "" Then
        gMsgBox "Format Name must be Defined.", vbOKOnly
        Exit Function
    End If
    'If cboFormat.text = "[New]" Or imNameChange Then
    '    SQLQuery = "select fmtCode from FMT_Station_Format where fmtName = '" & txtName.text & "'"
    '    Set rst = gSQLSelectCall(SQLQuery)
    '    If Not rst.EOF Then
    '        gMsgBox "The Format " & Trim$(txtName.text) & " name already exists"
    '        mSave = True
    '        Exit Function
    '    End If
    'End If
    
    ilRow = SendMessageByString(cboFormat.hwnd, CB_FINDSTRING, -1, slName)
    If ilRow >= 0 Then
        ilFmtCode = cboFormat.ItemData(ilRow)
        If imFmtCode <> ilFmtCode Then
            gMsgBox "Format Name Previously Defined.", vbOKOnly
            Exit Function
        End If
    End If
    If Trim$(txtGroupName.Text) <> "" Then
        SQLQuery = "SELECT FmtCode FROM Fmt_Station_Format WHERE fmtGroupName = '" & Trim$(txtGroupName.Text) & "'"
        Set fmt_rst = gSQLSelectCall(SQLQuery)
        If Not fmt_rst.EOF Then
            If imFmtCode <> fmt_rst!FmtCode Then
                gMsgBox "Format Group Name previously defined with another format.", vbOKOnly
                Exit Function
            End If
        End If
    End If
    mSave = False
    mTrimAndFixQuotes
    If cboFormat.Text = "[New]" Then
        SQLQuery = "Insert into Fmt_Station_Format "
        SQLQuery = SQLQuery & "(fmtName, fmtUstCode, fmtGroupName, fmtDftCode, fmtUnused) "
        SQLQuery = SQLQuery & " VALUES ('" & txtName.Text & "'," & igUstCode & ",'" & txtGroupName.Text & "'," & 0 & ",''" & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "GroupNameFormat-mSave"
            imInChg = False
            mSave = False
            Exit Function
        End If
        mClearControls
    Else
        SQLQuery = "UPDATE Fmt_Station_Format"
        SQLQuery = SQLQuery & " SET fmtUstCode = " & igUstCode & ","
        If sgFormatCall <> "M" Then
            SQLQuery = SQLQuery & "fmtName = '" & Trim$(txtName.Text) & "'" & ","
        End If
        SQLQuery = SQLQuery & "fmtGroupName = '" & Trim$(txtGroupName.Text) & "'"
        SQLQuery = SQLQuery & " WHERE FmtCode = " & imFmtCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "GroupNameFormat-mSave"
            imInChg = False
            mSave = False
            Exit Function
        End If
    End If
    'rst.Close
    imSave = True
    ilRet = mFillFmtInfo(slName)
    imSave = False
    imNameChange = False
    mSave = True
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "GroupNameFormat-mSave"
    imInChg = False
End Function

Private Sub mClearControls()

    imIgnoreChg = True
    txtName.Text = ""
    txtGroupName.Text = ""
    imIgnoreChg = False
End Sub

Private Function mFmtChkForChange()
    
    On Error GoTo ErrHand
    
    mFmtChkForChange = False

    If StrComp(txtName.Text, smName, vbTextCompare) <> 0 Then
        mFmtChkForChange = True
        Exit Function
    End If
    If (gWegenerExport) Or (gOLAExport) Then
        If StrComp(txtGroupName.Text, smGroupName, vbTextCompare) <> 0 Then
            mFmtChkForChange = True
            Exit Function
        End If
    End If
    
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in GroupNameFormat-mFmtChkForChange: "
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
        gMsg = "A general error has occured in GroupNameFormat-mTrimAndFixQuotes: "
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

    Dim fmt_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If cboFormat.Text = "[New]" Then
        Exit Sub
    End If
    
    imIgnoreChg = True
    SQLQuery = "SELECT * "
    SQLQuery = SQLQuery + " FROM Fmt_Station_Format"
    SQLQuery = SQLQuery + " WHERE (fmtCode = " & imFmtCode & ")"
    Set fmt_rst = gSQLSelectCall(SQLQuery)
    If fmt_rst.EOF Then
        gMsgBox "No matching records were found", vbOKOnly
        mClearControls
    Else
        txtName.Text = Trim$(fmt_rst!FmtName)
        smName = txtName.Text
        If (gWegenerExport) Or (gOLAExport) Then
            txtGroupName.Text = Trim$(fmt_rst!fmtGroupName)
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
    gHandleError "AffErrorLog.txt", "GroupNameFormat-mBindControls"
    imInChg = False
End Sub
