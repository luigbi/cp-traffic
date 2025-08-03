VERSION 5.00
Begin VB.Form frmGroupNameMarket 
   Caption         =   "Group Name-DMA Market"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   Icon            =   "AffGroupNameMarket.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGroupName 
      Height          =   315
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1785
      Width           =   2655
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   450
      TabIndex        =   5
      Top             =   2835
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2580
      TabIndex        =   6
      Top             =   2835
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4650
      TabIndex        =   7
      Top             =   2835
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.ComboBox cboMarket 
      Height          =   315
      Left            =   2280
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
      Caption         =   "DMA Market Name:"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1875
   End
End
Attribute VB_Name = "frmGroupNameMarket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmGroupNameMarket
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
Private imMktCode As Integer
Private imNameChange As Integer
Private imIgnoreChg As Integer
Private mkt_rst As ADODB.Recordset


Private Sub cboMarket_Change()
    
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
    slName = LTrim$(cboMarket.Text)
    ilLen = Len(slName)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(cboMarket.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        On Error GoTo ErrHand
        cboMarket.ListIndex = llRow
        cboMarket.SelStart = ilLen
        cboMarket.SelLength = Len(cboMarket.Text)
        If cboMarket.ListIndex < 0 Then
            imMktCode = -1
        ElseIf (cboMarket.ListIndex = 0) And (sgGNMarketCall <> "M") Then
            imMktCode = 0
        Else
            imMktCode = CInt(cboMarket.ItemData(cboMarket.ListIndex))
        End If
        If imMktCode < 0 Then
            mClearControls
        ElseIf (imMktCode = 0) And (sgGNMarketCall <> "M") Then
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
    gHandleError "AffErrorLog.txt", "frmGroupNameMarket-cboMarket_Change"
    imInChg = False
End Sub

Private Sub cboMarket_Click()
    If imInChg Then
        Exit Sub
    End If
    cboMarket_Change
End Sub

Private Sub cboMarket_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cboMarket_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboMarket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboMarket.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    igGNMarketReturnCode = 0
    igIndex = cboMarket.ListIndex
    Unload frmGroupNameMarket
End Sub

Private Sub cmdDone_Click()
    
    Dim ilRet As Integer

    ilRet = mMktChkForChange
    If ilRet Then
        ilRet = gMsgBox("Would You Like To Save the Changes", vbYesNo)
        If ilRet = vbYes Then
            ilRet = mSave()
        End If
        If Not ilRet Then
            Exit Sub
        End If
    End If
    igIndex = cboMarket.ListIndex
    igGNMarketReturnCode = 0
    If igIndex > 0 Then
        igGNMarketReturnCode = cboMarket.ItemData(igIndex)
    End If
    Unload frmGroupNameMarket
End Sub

Private Sub cmdSave_Click()
    Call mSave
End Sub

Private Sub Form_Activate()
    If sgGNMarketCall <> "M" Then
        If cboMarket.Index >= 0 Then
            txtName.SetFocus
        End If
    End If
End Sub

Private Sub Form_Initialize()
    gSetFonts frmGroupNameMarket
    gCenterForm frmGroupNameMarket
End Sub

Private Sub Form_Load()
    
    Form_Initialize
    If (Not gWegenerExport) And (Not gOLAExport) Then
        txtGroupName.Enabled = False
        txtGroupName.Text = ""
    End If
    imSave = False
    Call mFileMktInfo("")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    ilRet = gPopMarkets()
    mkt_rst.Close
    Set frmGroupNameMarket = Nothing

End Sub

Private Function mFileMktInfo(slFromSaveName As String)
    
    Dim mkt_rst As ADODB.Recordset
    Dim slName As String
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    
    If imSave Then
        slName = Trim$(slFromSaveName)
    End If
    If cboMarket.Text = "[New]" And (Not imSave) Then
        cboMarket.ListIndex = 0
        mClearControls
        Exit Function
    End If
    cboMarket.Clear
    SQLQuery = "SELECT * FROM mkt"
    Set mkt_rst = gSQLSelectCall(SQLQuery)
    While Not mkt_rst.EOF
        cboMarket.AddItem Trim$(mkt_rst!mktName)
        cboMarket.ItemData(cboMarket.NewIndex) = mkt_rst!mktCode
        mkt_rst.MoveNext
    Wend
    If sgGNMarketCall <> "M" Then
        cboMarket.AddItem "[New]", 0
        cboMarket.ItemData(cboMarket.NewIndex) = 0
    End If

    If (sgGNMarketCall <> "M") Or (imSave) Then
        If (sgGNMarketCall = "N") And (Not imSave) Then
            cboMarket.ListIndex = 0
        Else
            If Not imSave Then
                slName = sgGNMarketCall
            End If
            ilRow = SendMessageByString(cboMarket.hwnd, CB_FINDSTRING, -1, slName)
            If (ilRow > 0) Or ((ilRow = 0) And (sgGNMarketCall = "M")) Then
                cboMarket.ListIndex = ilRow
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
    gHandleError "AffErrorLog.txt", "frmGroupNameMarket-mFileMktInfo"
End Function

Private Function mSave()
    
    Dim ilRet As Integer
    Dim ilRow As Integer
    Dim slName As String
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    mSave = False
    ilRet = mMktChkForChange
    'slName = LTrim$(cboMarket.Text)
    slName = Trim$(txtName.Text)
    If Trim$(slName) = "" Then
        gMsgBox "DMA Market Name must be Defined.", vbOKOnly
        Exit Function
    End If
    'If cboMarket.text = "[New]" Or imNameChange Then
    '    SQLQuery = "select mktCode from mkt where mktName = '" & txtName.text & "'"
    '    Set rst = gSQLSelectCall(SQLQuery)
    '    If Not rst.EOF Then
    '        gMsgBox "The Market " & Trim$(txtName.text) & " name already exists"
    '        mSave = True
    '        Exit Function
    '    End If
    'End If
    
    ilRow = SendMessageByString(cboMarket.hwnd, CB_FINDSTRING, -1, slName)
    If ilRow >= 0 Then
        imMktCode = cboMarket.ItemData(ilRow)
    Else
        imMktCode = 0
    End If
    If Trim$(txtGroupName.Text) <> "" Then
        SQLQuery = "SELECT mktCode FROM mkt WHERE mktGroupName = '" & Trim$(txtGroupName.Text) & "'"
        Set mkt_rst = gSQLSelectCall(SQLQuery)
        If Not mkt_rst.EOF Then
            If imMktCode <> mkt_rst!mktCode Then
                gMsgBox "DMA Market Group Name previously defined with another market.", vbOKOnly
                Exit Function
            End If
        End If
    End If
    mSave = False
    mTrimAndFixQuotes
    If cboMarket.Text = "[New]" Then
        SQLQuery = "Insert into mkt "
        SQLQuery = SQLQuery & "(mktName, mktUSFCode, mktGroupName, mktUnused) "
        SQLQuery = SQLQuery & " VALUES ('" & txtName.Text & "'," & igUstCode & ",'" & txtGroupName.Text & "'," & "''" & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "GroupNameMarket-mSave"
            imInChg = False
            mSave = False
            Exit Function
        End If
        mClearControls
    Else
        SQLQuery = "UPDATE mkt"
        SQLQuery = SQLQuery & " SET mktUsfCode = " & igUstCode & ","
        SQLQuery = SQLQuery & "mktGroupName = '" & Trim$(txtGroupName.Text) & "'"
        SQLQuery = SQLQuery & " WHERE mktCode = " & imMktCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "GroupNameMarket-mSave"
            imInChg = False
            mSave = False
            Exit Function
        End If
    End If
    'rst.Close
    imSave = True
    ilRet = mFileMktInfo(slName)
    imSave = False
    imNameChange = False
    mSave = True
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmGroupNameMarket-mSave"
    imInChg = False
End Function

Private Sub mClearControls()

    imIgnoreChg = True
    txtName.Text = ""
    txtGroupName.Text = ""
    imIgnoreChg = False
End Sub

Private Function mMktChkForChange()
    
    On Error GoTo ErrHand
    
    mMktChkForChange = False

    If StrComp(txtName.Text, smName, vbTextCompare) <> 0 Then
        mMktChkForChange = True
        Exit Function
    End If
    If StrComp(txtGroupName.Text, smGroupName, vbTextCompare) <> 0 Then
        mMktChkForChange = True
        Exit Function
    End If
    
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in GroupNameMarket-mmktChkForChange: "
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
        gMsg = "A general error has occured in GroupNameMarket-mTrimAndFixQuotes: "
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

    Dim mkt_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If cboMarket.Text = "[New]" Then
        Exit Sub
    End If
    
    imIgnoreChg = True
    SQLQuery = "SELECT * "
    SQLQuery = SQLQuery + " FROM mkt"
    SQLQuery = SQLQuery + " WHERE (mktCode = " & imMktCode & ")"
    Set mkt_rst = gSQLSelectCall(SQLQuery)
    If mkt_rst.EOF Then
        gMsgBox "No matching records were found", vbOKOnly
        mClearControls
    Else
        txtName.Text = Trim$(mkt_rst!mktName)
        smName = txtName.Text
        If (gWegenerExport) Or (gOLAExport) Then
            txtGroupName.Text = Trim$(mkt_rst!mktGroupName)
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
    gHandleError "AffErrorLog.txt", "frmGroupNameMarket-mBindControls"
    imInChg = False
End Sub
