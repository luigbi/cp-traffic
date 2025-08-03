VERSION 5.00
Begin VB.Form frmStationMktInfo 
   Caption         =   "DMA Market/Cluster Information"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   Icon            =   "AffStationMktInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGroupName 
      Height          =   375
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2340
      Width           =   2655
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtRank 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   5655
   End
   Begin VB.ComboBox cboMarket 
      Height          =   315
      Left            =   2640
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lacGroupName 
      Caption         =   "Group Name:"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   2460
      Width           =   2055
   End
   Begin VB.Label lblRank 
      Caption         =   "Rank:"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblNAme 
      Caption         =   "DMA Name:"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "frmStationMktInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmStationMktInfo
'*
'*  Created October,2005 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc. 2005
'*
'******************************************************
Option Explicit
Option Compare Text

Private smName As String
Private smRank As String
Private smGroupName As String

Private imInChg As Integer
Private imBSMode As Integer
Private imIsFirstTime As Integer
Private imSave As Integer
Private imMktCode As Integer
Private imNameChange As Integer
Private imIgnoreChg As Integer

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
        If cboMarket.ListIndex <= 0 Then
            imMktCode = 0
        Else
            imMktCode = CInt(cboMarket.ItemData(cboMarket.ListIndex))
        End If
        If imMktCode <= 0 Then
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
    gHandleError "AffErrorLog.txt", "frmStationMktInfo-cboMarket_Change"
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
    igMarketReturnCode = -1
    igIndex = cboMarket.ListIndex
    Unload frmStationMktInfo
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
    igMarketReturnCode = 0
    If igIndex > 0 Then
        igMarketReturnCode = cboMarket.ItemData(igIndex)
    End If
    Unload frmStationMktInfo
End Sub

Private Sub cmdSave_Click()
    Call mSave
End Sub

Private Sub Form_Activate()
    txtName.SetFocus
End Sub

Private Sub Form_Initialize()
    gSetFonts frmStationMktInfo
    gCenterForm frmStationMktInfo
End Sub

Private Sub Form_Load()
    
    Form_Initialize
    imIsFirstTime = True
    If (Not gWegenerExport) And (Not gOLAExport) Then
        txtGroupName.Enabled = False
        txtGroupName.Text = ""
        txtGroupName.Visible = False
        lacGroupName.Visible = False
    End If
    Call mFillMktInfo("")
    imIsFirstTime = False
    imIgnoreChg = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    ilRet = gPopMarkets()
    Set frmStationMktInfo = Nothing

End Sub

Private Function mFillMktInfo(slFromSaveName As String)
    
    Dim mkt_rst As ADODB.Recordset
    Dim slName As String
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    
    slName = Trim$(cboMarket.Text)
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
    cboMarket.AddItem "[New]", 0
    cboMarket.ItemData(cboMarket.NewIndex) = 0

    If (frmStation!cbcDMAMarket.ListIndex <= 0) And (imIsFirstTime) Then
        cboMarket.ListIndex = 0
    Else
        If imIsFirstTime Then
            slName = Trim$(frmStation!cbcDMAMarket.GetName(frmStation!cbcDMAMarket.ListIndex))
        End If
        'If imSave Then
        '    slName = Trim$(txtName.Text)
        'End If
        ilRow = SendMessageByString(cboMarket.hwnd, CB_FINDSTRING, -1, slName)
        If ilRow > 0 Then
            cboMarket.ListIndex = ilRow
            mBindControls
        Else
            mClearControls
        End If
        DoEvents
    End If
    Screen.MousePointer = vbDefault
    imInChg = False
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationMktInfo-mFillMktInfo"
    imInChg = False
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
        gMsgBox "Market Name must be Defined.", vbOKOnly
        Exit Function
    End If
    'If cboMarket.text = "[New]" And Not imNameChange Then
    '    SQLQuery = "select arttCode from artt where arttLastName = '" & txtName.text & "'"
    '    Set rst = gSQLSelectCall(SQLQuery)
    '    If Not rst.EOF Then
    '        gMsgBox "The market " & Trim$(txtName.text) & " already exists"
    '        mSave = True
    '        Exit Function
    '    End If
    'End If
    
    'ilRow = SendMessageByString(cboMarket.hwnd, CB_FINDSTRING, -1, slName)
    'If ilRow >= 0 Then
    '    imMktCode = cboMarket.ItemData(ilRow)
    'Else
    '    imMktCode = 0
    'End If
    ilRow = SendMessageByString(cboMarket.hwnd, CB_FINDSTRING, -1, slName)
    If ilRow >= 0 Then
        If cboMarket.ItemData(ilRow) <> imMktCode Then
            gMsgBox "The market " & Trim$(txtName.Text) & " already exists"
            mSave = True
            Exit Function
        End If
    End If
    mSave = False
    mTrimAndFixQuotes
    If (cboMarket.Text = "[New]") Or (Trim$(cboMarket.Text) = "") Then
        SQLQuery = "Insert into Mkt "
        SQLQuery = SQLQuery & "(mktName, mktRank,  mktUSFCode, mktGroupName) "
        SQLQuery = SQLQuery & " VALUES ('" & txtName.Text & "', '" & txtRank.Text & "', " & igUstCode & ", '" & txtGroupName.Text & "')"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "StationMktInfo-mSave"
            imInChg = False
            mSave = False
            Exit Function
        End If
        mClearControls
    Else
        SQLQuery = "UPDATE mkt"
        SQLQuery = SQLQuery & " SET mktName = '" & txtName.Text & "',"
        SQLQuery = SQLQuery & "mktRank = '" & Trim$(txtRank.Text) & "',"
        SQLQuery = SQLQuery & "mktGroupName = '" & Trim$(txtGroupName.Text) & "'"
        SQLQuery = SQLQuery & " WHERE mktCode = " & imMktCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "StationMktInfo-mSave"
            imInChg = False
            mSave = False
            Exit Function
        End If
    End If
    imSave = True
    ilRet = mFillMktInfo(slName)
    imSave = False
    imNameChange = False
    mSave = True
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationMktInfo-mSave"
    imInChg = False
End Function

Private Sub mClearControls()

    txtName.Text = ""
    txtRank.Text = ""
    txtGroupName.Text = ""
    
End Sub

Private Function mMktChkForChange()
    
    On Error GoTo ErrHand
    
    mMktChkForChange = False

    If StrComp(txtName.Text, smName, 1) <> 0 Then
        mMktChkForChange = True
        Exit Function
    End If
    If StrComp(txtRank.Text, smRank, 1) <> 0 Then
        mMktChkForChange = True
        Exit Function
    End If
    
    If StrComp(txtGroupName.Text, smGroupName, 1) <> 0 Then
        mMktChkForChange = True
        Exit Function
    End If
    
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in StationMktInfo-mMktChkForChange: "
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
        gMsg = "A general error has occured in StationMktInfo-mTrimAndFixQuotes: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    imInChg = False
End Sub


Private Sub txtGroupName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtName_Change()
    If mMktChkForChange Then
        If Not imIgnoreChg Then
            imNameChange = True
        End If
    End If
End Sub

Private Sub txtName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtRank_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub mBindControls()

    Dim mkt_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If cboMarket.Text = "[New]" Then
        Exit Sub
    End If
    
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
        txtRank.Text = Trim$(mkt_rst!mktRank)
        smRank = txtRank.Text
        If (gWegenerExport) Or (gOLAExport) Then
            txtGroupName.Text = Trim$(mkt_rst!mktGroupName)
            smGroupName = txtGroupName.Text
        Else
            txtGroupName.Text = ""
            smGroupName = txtGroupName.Text
        End If
    End If
Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationMktInfo-mBindControls"
    imInChg = False
End Sub
