VERSION 5.00
Begin VB.Form frmStationOwnerInfo 
   Caption         =   "Owner Information"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "AffStationOwnerInfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEmail 
      Height          =   375
      Left            =   2040
      MaxLength       =   70
      TabIndex        =   19
      Top             =   5160
      Width           =   6015
   End
   Begin VB.TextBox txtFax 
      Height          =   375
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   17
      Top             =   4680
      Width           =   2895
   End
   Begin VB.TextBox txtPhone 
      Height          =   375
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   15
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox txtZip 
      Height          =   375
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   13
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox txtCountry 
      Height          =   375
      Left            =   2040
      MaxLength       =   40
      TabIndex        =   11
      Top             =   3240
      Width           =   4815
   End
   Begin VB.TextBox txtState 
      Height          =   375
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox txtCity 
      Height          =   375
      Left            =   2040
      MaxLength       =   40
      TabIndex        =   7
      Top             =   2280
      Width           =   4815
   End
   Begin VB.TextBox txtAddress2 
      Height          =   375
      Left            =   2040
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1800
      Width           =   4815
   End
   Begin VB.TextBox txtAddress1 
      Height          =   375
      Left            =   2040
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   2040
      MaxLength       =   40
      TabIndex        =   2
      Top             =   840
      Width           =   4815
   End
   Begin VB.ComboBox cboOwnerInfo 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5760
      TabIndex        =   22
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   21
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   1080
      TabIndex        =   20
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label lblEmail 
      Caption         =   "Email:"
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label lblFax 
      Caption         =   "Fax #:"
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblPhone 
      Caption         =   "Phone:"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label lblZip 
      Caption         =   "Zip:"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Country:"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblState 
      Caption         =   "State:"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblCity 
      Caption         =   "City:"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblAdress2 
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblAddress1 
      Caption         =   "Address:"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblNAme 
      Caption         =   "Name:"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmStationOwnerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmStationOwnerInfo
'*
'*  Created October,2005 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc. 2005
'*
'******************************************************

Option Explicit
Option Compare Text

Private smName As String
Private smAddr1 As String
Private smAddr2 As String
Private smCity As String
Private smState As String
Private smCountry As String
Private smZip As String
Private smPhone As String
Private smFax As String
Private smEmail As String

Private imInChg As Integer
Private imBSMode As Integer
Private lmArttCode As Long
Private imIsFirstTime As Integer
Private imSave As Integer
Private imNameChange As Integer
Private imIgnoreChg As Integer


Private Sub cboOwnerInfo_Change()

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
    slName = LTrim$(cboOwnerInfo.Text)
    ilLen = Len(slName)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(cboOwnerInfo.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        On Error GoTo ErrHand
        cboOwnerInfo.ListIndex = llRow
        cboOwnerInfo.SelStart = ilLen
        cboOwnerInfo.SelLength = Len(cboOwnerInfo.Text)
        If cboOwnerInfo.ListIndex <= 0 Then
            lmArttCode = 0
        Else
            lmArttCode = CLng(cboOwnerInfo.ItemData(cboOwnerInfo.ListIndex))
        End If
        If lmArttCode <= 0 Then
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
    gHandleError "AffErrorLog.txt", "frmStationOwnerInfo-cboOwnerInfo"
End Sub

Private Sub cboOwnerInfo_Click()
    imInChg = False
    cboOwnerInfo_Change
End Sub

Private Sub cboOwnerInfo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cboOwnerInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboOwnerInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboOwnerInfo.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    igIndex = cboOwnerInfo.ListIndex
    lgOwnerReturnCode = -1
    Unload frmStationOwnerInfo
End Sub

Private Sub cmdDone_Click()

    Dim ilRet As Integer

    ilRet = mOwnerChkForChange
    If ilRet Then
        ilRet = gMsgBox("Would You Like To Save the Changes", vbYesNo)
        If ilRet = vbYes Then
            ilRet = mSave()
        End If
        If Not ilRet Then
            Exit Sub
        End If
    End If
    igIndex = cboOwnerInfo.ListIndex
    lgOwnerReturnCode = 0
    If igIndex > 0 Then
        lgOwnerReturnCode = cboOwnerInfo.ItemData(igIndex)
    End If
    Unload frmStationOwnerInfo
End Sub

Private Sub cmdSave_Click()
   Call mSave
End Sub

Private Sub Form_Activate()
    txtName.SetFocus
End Sub

Private Sub Form_Initialize()
    gSetFonts frmStationOwnerInfo
    gCenterForm frmStationOwnerInfo
End Sub

Private Sub Form_Load()
    imBSMode = False
    imInChg = False
    imIsFirstTime = True
    Call mFillOwnerInfo("")
    imIsFirstTime = False
    imIgnoreChg = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmStationOwnerInfo = Nothing
End Sub

Private Function mFillOwnerInfo(slFromSaveName As String) As Integer

    Dim artt_rst As ADODB.Recordset
    Dim slName As String
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    
    slName = Trim$(cboOwnerInfo.Text)
    If imSave Then
        slName = Trim$(slFromSaveName)
    End If
    If (cboOwnerInfo.Text = "[New]") And (Not imSave) Then
        cboOwnerInfo.ListIndex = 0
        mClearControls
        Exit Function
    End If
    cboOwnerInfo.Clear
    SQLQuery = "SELECT * FROM artt Where arttType = " & "'O'"
    Set artt_rst = gSQLSelectCall(SQLQuery)
    While Not artt_rst.EOF
        cboOwnerInfo.AddItem Trim$(artt_rst!arttLastName)
        cboOwnerInfo.ItemData(cboOwnerInfo.NewIndex) = artt_rst!arttCode
        artt_rst.MoveNext
    Wend
    cboOwnerInfo.AddItem "[New]", 0
    cboOwnerInfo.ItemData(cboOwnerInfo.NewIndex) = 0
    
    If (frmStation!cbcOwner.ListIndex <= 0) And (imIsFirstTime) Then
        cboOwnerInfo.ListIndex = 0
    Else
        If imIsFirstTime Then
            slName = Trim$(frmStation!cbcOwner.GetName(frmStation!cbcOwner.ListIndex))
        End If
        'If imSave Then
        '    slName = Trim$(txtName.Text)
        'End If
        imIgnoreChg = True
        ilRow = SendMessageByString(cboOwnerInfo.hwnd, CB_FINDSTRING, -1, slName)
        If ilRow > 0 Then
            cboOwnerInfo.ListIndex = ilRow
            mBindControls
        Else
            mClearControls
        End If
'        SQLQuery = "SELECT * FROM artt where arttCode = " & lmArttCode 'cboOwnerInfo.ItemData(cboOwnerInfo.ListIndex)
'        Set artt_rst = gSQLSelectCall(SQLQuery)
'        txtName.Text = Trim$(artt_rst!arttLastName)
'        smName = txtName.Text
'        txtAddress1.Text = Trim$(artt_rst!arttAddress1)
'        smAddr1 = txtAddress1.Text
'        txtAddress2.Text = Trim$(artt_rst!arttAddress2)
'        smAddr2 = txtAddress2.Text
'        txtCity.Text = Trim$(artt_rst!arttCity)
'        smCity = txtCity.Text
'        txtState.Text = Trim$(artt_rst!arttAddressState)
'        smState = txtState.Text
'        txtCountry.Text = Trim$(artt_rst!arttCountry)
'        smCountry = txtCountry.Text
'        txtZip.Text = Trim$(artt_rst!arttZip)
'        smZip = txtZip.Text
'        txtPhone.Text = Trim$(artt_rst!arttPhone)
'        smPhone = txtPhone.Text
'        txtFax.Text = Trim$(artt_rst!arttFax)
'        smFax = txtFax.Text
'        txtPhone.Text = Trim$(artt_rst!arttPhone)
'        smPhone = txtPhone.Text
'        txtEmail.Text = Trim$(artt_rst!arttEmail)
'        smEmail = txtEmail.Text
        DoEvents
    End If
    imIsFirstTime = False
    Screen.MousePointer = vbDefault
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationOwnerInfo-mFillOwnerInfo"
End Function

Public Function mOwnerChkForChange() As Integer

    On Error GoTo ErrHand
    
    mOwnerChkForChange = False

    If txtName.Text = "" Then
        mOwnerChkForChange = False
        Exit Function
    End If
    
    If StrComp(txtName.Text, smName, 1) <> 0 Then
        mOwnerChkForChange = True
        Exit Function
    End If
    If StrComp(txtAddress1.Text, smAddr1, 1) <> 0 Then
        mOwnerChkForChange = True
        Exit Function
    End If
    If StrComp(txtAddress2.Text, smAddr2, 1) <> 0 Then
        mOwnerChkForChange = True
        Exit Function
    End If
    If StrComp(txtCity.Text, smCity, 1) <> 0 Then
        mOwnerChkForChange = True
        Exit Function
    End If
    If StrComp(txtState.Text, smState, 1) <> 0 Then
        mOwnerChkForChange = True
        Exit Function
    End If
    If StrComp(txtCountry.Text, smCountry, 1) <> 0 Then
        mOwnerChkForChange = True
        Exit Function
    End If
    If StrComp(txtZip.Text, smZip, 1) <> 0 Then
        mOwnerChkForChange = True
        Exit Function
    End If
    If StrComp(txtFax.Text, smFax, 1) <> 0 Then
        mOwnerChkForChange = True
        Exit Function
    End If
    If StrComp(txtPhone.Text, smPhone, 1) <> 0 Then
        mOwnerChkForChange = True
        Exit Function
    End If
    If StrComp(txtEmail.Text, smEmail, 1) <> 0 Then
        mOwnerChkForChange = True
        Exit Function
    End If
    
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in StationOwnerInfo-mOwnerChkForChange: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
End Function

Private Function mSave() As Integer

    Dim ilRet As Integer
    Dim ilRow As Integer
    Dim slName As String
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    mSave = False
    'slName = LTrim$(cboOwnerInfo.Text)
    slName = Trim$(txtName.Text)
    If Trim$(slName) = "" Then
        gMsgBox "Owner Name must be Defined.", vbOKOnly
        Exit Function
    End If
    ''If cboOwnerInfo.text = "[New]" And Not imNameChange Then
    ''    SQLQuery = "select arttCode from artt where arttLastName = '" & txtName.text & "'"
    ''    Set rst = gSQLSelectCall(SQLQuery)
    ''    If Not rst.EOF Then
    ''        gMsgBox "The owner " & Trim$(txtName.text) & " already exists"
    ''        mSave = True
    ''        Exit Function
    ''    End If
    ''End If
    
    ''ilRow = SendMessageByString(cboOwnerInfo.hwnd, CB_FINDSTRING, -1, slName)
    ''If ilRow >= 0 Then
    ''    lmArttCode = cboOwnerInfo.ItemData(ilRow)
    ''Else
    ''    lmArttCode = 0
    ''End If
    'ilRow = SendMessageByString(cboOwnerInfo.hwnd, CB_FINDSTRING, -1, slName)
    'If ilRow >= 0 Then
    '    If rst!arttCode <> lmArttCode Then
    '        gMsgBox "The owner " & Trim$(txtName.Text) & " already exists"
    '        Exit Function
    '    End If
    'End If
    'Check
    SQLQuery = "SELECT arttCode FROM artt WHERE UCase(arttLastName) = '" & UCase(gFixQuote(txtName.Text)) & "'"
    SQLQuery = SQLQuery & " AND arttType = '" & "O" & "'"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If rst!arttCode <> lmArttCode Then
            gMsgBox "The owner " & Trim$(txtName.Text) & " already exists"
            Exit Function
        End If
    End If
    mSave = False
    mTrimAndFixQuotes
    If (cboOwnerInfo.Text = "[New]") Or (Trim$(cboOwnerInfo.Text) = "") Then
        SQLQuery = "Insert into Artt "
        SQLQuery = SQLQuery & "(arttType, arttLastName, arttAddress1, arttAddress2, arttCity, "
        SQLQuery = SQLQuery & "arttAddressState, arttCountry, arttZip, arttPhone, "
        SQLQuery = SQLQuery & "arttFax, arttEmail, arttEMailRights)"
        SQLQuery = SQLQuery & " VALUES ('O', '" & gFixQuote(txtName.Text) & "', '" & gFixQuote(txtAddress1.Text) & "', '" & gFixQuote(txtAddress2.Text) & "', '" & txtCity.Text & "', "
        SQLQuery = SQLQuery & "'" & txtState.Text & "', '" & txtCountry.Text & "', '" & txtZip.Text & "', '" & txtPhone.Text & "', "
        SQLQuery = SQLQuery & "'" & txtFax.Text & "', '" & txtEmail.Text & "', '" & "N" & "'" & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "StationOwnerInfo-mSave"
            mSave = False
            Exit Function
        End If
        mClearControls
    Else
        SQLQuery = "UPDATE artt"
        SQLQuery = SQLQuery & " SET arttLastName = '" & gFixQuote(txtName.Text) & "',"
        SQLQuery = SQLQuery & "arttAddress1 = '" & gFixQuote(txtAddress1.Text) & "',"
        SQLQuery = SQLQuery & "arttAddress2 = '" & gFixQuote(txtAddress2.Text) & "',"
        SQLQuery = SQLQuery & "arttCity = '" & txtCity.Text & "',"
        SQLQuery = SQLQuery & "arttAddressState = '" & txtState.Text & "',"
        SQLQuery = SQLQuery & "arttCountry = '" & txtCountry.Text & "',"
        SQLQuery = SQLQuery & "arttZip = '" & txtZip.Text & "',"
        SQLQuery = SQLQuery & "arttFax = '" & txtFax.Text & "',"
        SQLQuery = SQLQuery & "arttPhone = '" & txtPhone.Text & "',"
        SQLQuery = SQLQuery & "arttEmail = '" & txtEmail.Text & "',"
        SQLQuery = SQLQuery & "arttType = 'O'"
        SQLQuery = SQLQuery & " WHERE arttCode = " & lmArttCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "StationOwnerInfo-mSave"
            mSave = False
            Exit Function
        End If
    End If
    imSave = True
    ilRet = mFillOwnerInfo(slName)
    imSave = False
    imNameChange = False
    mSave = True
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationOwnerInfo-mSave"
End Function

Private Sub mClearControls()

    
    imIgnoreChg = True
    txtName.Text = ""
    txtAddress1.Text = ""
    txtAddress2.Text = ""
    txtCity.Text = ""
    txtState.Text = ""
    txtCountry.Text = ""
    txtZip.Text = ""
    txtFax.Text = ""
    txtEmail.Text = ""
    txtPhone.Text = ""
    If Not imIsFirstTime Then
        txtName.SetFocus
        cboOwnerInfo.ListIndex = 0
    End If

    If cboOwnerInfo.Text = "[New]" And Not imIsFirstTime Then
        cboOwnerInfo.ListIndex = 0
        txtName.SetFocus
    End If
    
Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in StationOwnerInfo-mClearControls: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    
End Sub

Private Sub mTrimAndFixQuotes()

    On Error GoTo ErrHand
    
    txtName.Text = gFixQuote(Trim$(txtName.Text))
    txtAddress1.Text = gFixQuote(Trim$(txtAddress1.Text))
    txtAddress2.Text = gFixQuote(Trim$(txtAddress2.Text))
    txtCity.Text = gFixQuote(Trim$(txtCity.Text))
    txtState.Text = gFixQuote(Trim$(txtState.Text))
    txtCountry.Text = gFixQuote(Trim$(txtCountry.Text))
    txtZip.Text = gFixQuote(Trim$(txtZip.Text))
    txtPhone.Text = gFixQuote(Trim$(txtPhone.Text))
    txtFax.Text = gFixQuote(Trim$(txtFax.Text))
    txtEmail.Text = gFixQuote(Trim$(txtEmail.Text))
    imNameChange = False
Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in StationOwnerInfo-mTrimAndFixQuotes: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If

End Sub

Private Sub txtAddress1_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtAddress2_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtCity_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub txtCountry_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub txtEmail_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub txtFax_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtName_Change()
    If mOwnerChkForChange Then
        If Not imIgnoreChg Then
            imNameChange = True
        End If
    End If
End Sub

Private Sub txtName_GotFocus()
    'Call mFillOwnerInfo
    gCtrlGotFocus ActiveControl
End Sub


Private Sub txtPhone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub txtState_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub txtZip_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub mBindControls()
    
    Dim artt_rst As ADODB.Recordset
    
    If cboOwnerInfo.Text = "[New]" Then
        Exit Sub
    End If
    
    SQLQuery = "SELECT * "
    SQLQuery = SQLQuery + " FROM artt"
    SQLQuery = SQLQuery + " WHERE (arttCode = " & lmArttCode & ")"
    Set artt_rst = gSQLSelectCall(SQLQuery)
    If artt_rst.EOF Then
        gMsgBox "No matching records were found", vbOKOnly
        mClearControls
    Else
        txtName.Text = Trim$(artt_rst!arttLastName)
        smName = txtName.Text
        txtAddress1.Text = Trim$(artt_rst!arttAddress1)
        smAddr1 = txtAddress1.Text
        txtAddress2.Text = Trim$(artt_rst!arttAddress2)
        smAddr2 = txtAddress2.Text
        txtCity.Text = Trim$(artt_rst!arttCity)
        smCity = txtCity.Text
        txtState.Text = Trim$(artt_rst!arttAddressState)
        smState = txtState.Text
        txtCountry.Text = Trim$(artt_rst!arttCountry)
        smCountry = txtCountry.Text
        txtZip.Text = Trim$(artt_rst!arttZip)
        smZip = txtZip.Text
        txtPhone.Text = Trim$(artt_rst!arttPhone)
        smPhone = txtPhone.Text
        txtFax.Text = Trim$(artt_rst!arttFax)
        smFax = txtFax.Text
        txtPhone.Text = Trim$(artt_rst!arttPhone)
        smPhone = txtPhone.Text
        txtEmail.Text = Trim$(artt_rst!arttEmail)
        smEmail = txtEmail.Text
    End If
Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationOwnerInfo-mBindControls"
End Sub
