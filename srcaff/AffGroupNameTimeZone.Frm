VERSION 5.00
Begin VB.Form frmGroupNameTimeZone 
   Caption         =   "Group Name- Time Zone"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   Icon            =   "AffGroupNameTimeZone.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton rbcUseZone 
      Caption         =   "PST"
      Height          =   240
      Index           =   3
      Left            =   5445
      TabIndex        =   12
      Top             =   2400
      Width           =   1110
   End
   Begin VB.OptionButton rbcUseZone 
      Caption         =   "MST"
      Height          =   240
      Index           =   2
      Left            =   4410
      TabIndex        =   11
      Top             =   2400
      Width           =   1110
   End
   Begin VB.OptionButton rbcUseZone 
      Caption         =   "CST"
      Height          =   240
      Index           =   1
      Left            =   3345
      TabIndex        =   10
      Top             =   2400
      Width           =   1110
   End
   Begin VB.OptionButton rbcUseZone 
      Caption         =   "EST"
      Height          =   240
      Index           =   0
      Left            =   2250
      TabIndex        =   9
      Top             =   2400
      Width           =   1110
   End
   Begin VB.TextBox txtGroupName 
      Height          =   315
      Left            =   2295
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1755
      Width           =   2655
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   510
      TabIndex        =   5
      Top             =   2835
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   2835
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4710
      TabIndex        =   7
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
   Begin VB.ComboBox cboTimeZone 
      Height          =   315
      Left            =   2295
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lacUseZone 
      Caption         =   "Use Zone:"
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   2385
      Width           =   1620
   End
   Begin VB.Label lacGroupName 
      Caption         =   "Group Name:"
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1755
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
Attribute VB_Name = "frmGroupNameTimeZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmGroupNameTimeZone
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
Private imTztCode As Integer
Private imNameChange As Integer
Private imIgnoreChg As Integer
Private tzt_rst As ADODB.Recordset


Private Sub cboTimeZone_Change()
    
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
    slName = LTrim$(cboTimeZone.Text)
    ilLen = Len(slName)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(cboTimeZone.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        On Error GoTo ErrHand
        cboTimeZone.ListIndex = llRow
        cboTimeZone.SelStart = ilLen
        cboTimeZone.SelLength = Len(cboTimeZone.Text)
        If cboTimeZone.ListIndex < 0 Then
            imTztCode = -1
        ElseIf (cboTimeZone.ListIndex = 0) And (sgTimeZoneCall <> "M") Then
            imTztCode = 0
        Else
            imTztCode = CInt(cboTimeZone.ItemData(cboTimeZone.ListIndex))
        End If
        If imTztCode < 0 Then
            mClearControls
        ElseIf (imTztCode = 0) And (sgTimeZoneCall <> "M") Then
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
    gHandleError "AffErrorLog.txt", "frmGroupNameTimeZone-cboTimeZone_Change"
    imInChg = False
End Sub

Private Sub cboTimeZone_Click()
    If imInChg Then
        Exit Sub
    End If
    cboTimeZone_Change
End Sub

Private Sub cboTimeZone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cboTimeZone_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboTimeZone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboTimeZone.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    igTimeZoneReturnCode = 0
    igIndex = cboTimeZone.ListIndex
    Unload frmGroupNameTimeZone
End Sub

Private Sub cmdDone_Click()
    
    Dim ilRet As Integer

    ilRet = mTztChkForChange
    If ilRet Then
        ilRet = gMsgBox("Would You Like To Save the Changes", vbYesNo)
        If ilRet = vbYes Then
            ilRet = mSave()
        End If
        If Not ilRet Then
            Exit Sub
        End If
    End If
    igIndex = cboTimeZone.ListIndex
    igTimeZoneReturnCode = 0
    If igIndex > 0 Then
        igTimeZoneReturnCode = cboTimeZone.ItemData(igIndex)
    End If
    Unload frmGroupNameTimeZone
End Sub

Private Sub cmdSave_Click()
    Call mSave
End Sub

Private Sub Form_Activate()
    If sgTimeZoneCall <> "M" Then
        If cboTimeZone.Index >= 0 Then
            txtName.SetFocus
        End If
    End If
End Sub

Private Sub Form_Initialize()
    gSetFonts frmGroupNameTimeZone
    gCenterForm frmGroupNameTimeZone
End Sub

Private Sub Form_Load()
    
    Form_Initialize
    If (Not gWegenerExport) And (Not gOLAExport) Then
        txtGroupName.Enabled = False
        txtGroupName.Text = ""
    End If
    imSave = False
    Call mFillTztInfo("")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    ilRet = gPopTimeZones()
    tzt_rst.Close
    Set frmGroupNameTimeZone = Nothing

End Sub

Private Function mFillTztInfo(slFromSaveName As String)
    
    Dim tzt_rst As ADODB.Recordset
    Dim slName As String
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    
    If imSave Then
        slName = Trim$(slFromSaveName)
    End If
    If cboTimeZone.Text = "[New]" And (Not imSave) Then
        cboTimeZone.ListIndex = 0
        mClearControls
        Exit Function
    End If
    cboTimeZone.Clear
    SQLQuery = "SELECT * FROM TZT"
    Set tzt_rst = gSQLSelectCall(SQLQuery)
    While Not tzt_rst.EOF
        cboTimeZone.AddItem Trim$(tzt_rst!tztName)
        cboTimeZone.ItemData(cboTimeZone.NewIndex) = tzt_rst!tztCode
        tzt_rst.MoveNext
    Wend
    If sgTimeZoneCall <> "M" Then
        cboTimeZone.AddItem "[New]", 0
        cboTimeZone.ItemData(cboTimeZone.NewIndex) = 0
    End If

    If (sgTimeZoneCall <> "M") Or (imSave) Then
        If (sgTimeZoneCall = "N") And (Not imSave) Then
            cboTimeZone.ListIndex = 0
        Else
            If Not imSave Then
                slName = sgTimeZoneCall
            End If
            ilRow = SendMessageByString(cboTimeZone.hwnd, CB_FINDSTRING, -1, slName)
            If (ilRow > 0) Or ((ilRow = 0) And (sgTimeZoneCall = "M")) Then
                cboTimeZone.ListIndex = ilRow
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
    gHandleError "AffErrorLog.txt", "frmGroupNameTimeZone-mFillTztInfo"
    imInChg = False
End Function

Private Function mSave()
    
    Dim ilRet As Integer
    Dim ilRow As Integer
    Dim slName As String
    Dim slZone As String
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    mSave = False
    ilRet = mTztChkForChange
    'slName = LTrim$(cboTimeZone.Text)
    slName = Trim$(txtName.Text)
    If Trim$(slName) = "" Then
        gMsgBox "Time Zone Name must be Defined.", vbOKOnly
        Exit Function
    End If
    If (rbcUseZone(0).Value = False) And (rbcUseZone(1).Value = False) And (rbcUseZone(2).Value = False) And (rbcUseZone(3).Value = False) Then
        gMsgBox "Use Zone must be Defined.", vbOKOnly
        Exit Function
    End If
    'If cboTimeZone.text = "[New]" Or imNameChange Then
    '    SQLQuery = "select TztCode from TZT where TztName = '" & txtName.text & "'"
    '    Set rst = gSQLSelectCall(SQLQuery)
    '    If Not rst.EOF Then
    '        gMsgBox "The Time zone " & Trim$(txtName.text) & " name already exists"
    '        mSave = True
    '        Exit Function
    '    End If
    'End If
    If rbcUseZone(0).Value Then
        slZone = "EST"
    End If
    If rbcUseZone(1).Value Then
        slZone = "CST"
    End If
    If rbcUseZone(2).Value Then
        slZone = "MST"
    End If
    If rbcUseZone(3).Value Then
        slZone = "PST"
    End If
    ilRow = SendMessageByString(cboTimeZone.hwnd, CB_FINDSTRING, -1, slName)
    If ilRow >= 0 Then
        imTztCode = cboTimeZone.ItemData(ilRow)
    Else
        imTztCode = 0
    End If
    If Trim$(txtGroupName.Text) <> "" Then
        SQLQuery = "SELECT tztCode FROM tzt WHERE tztGroupName = '" & Trim$(txtGroupName.Text) & "'"
        Set tzt_rst = gSQLSelectCall(SQLQuery)
        If Not tzt_rst.EOF Then
            If imTztCode <> tzt_rst!tztCode Then
                gMsgBox "Time Zone Group Name previously defined with another time zone.", vbOKOnly
                Exit Function
            End If
        End If
    End If
    mSave = False
    mTrimAndFixQuotes
    If cboTimeZone.Text = "[New]" Then
        SQLQuery = "Insert into TZT "
        SQLQuery = SQLQuery & "(tztName, tztGroupName, tztCSIName, tztUstCode, TztUnused) "
        SQLQuery = SQLQuery & " VALUES ('" & txtName.Text & "','" & txtGroupName.Text & "','" & slZone & "'," & igUstCode & "," & "''" & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "GroupNameTimeZone-mSave"
            imInChg = False
            mSave = False
            Exit Function
        End If
        mClearControls
    Else
        SQLQuery = "UPDATE TZT"
        SQLQuery = SQLQuery & " SET TztGroupName = '" & Trim$(txtGroupName.Text) & "',"
        SQLQuery = SQLQuery & "tztCSIName = '" & Trim$(slZone) & "',"
        SQLQuery = SQLQuery & "tztUstCode =" & igUstCode
        SQLQuery = SQLQuery & " WHERE TztCode = " & imTztCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "GroupNameTimeZone-mSave"
            imInChg = False
            mSave = False
            Exit Function
        End If
    End If
    'rst.Close
    imSave = True
    ilRet = mFillTztInfo(slName)
    imSave = False
    imNameChange = False
    mSave = True
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmGroupNameTimeZone-mSave"
    imInChg = False
End Function

Private Sub mClearControls()

    imIgnoreChg = True
    txtName.Text = ""
    txtGroupName.Text = ""
    rbcUseZone(0).Value = False
    rbcUseZone(1).Value = False
    rbcUseZone(2).Value = False
    rbcUseZone(3).Value = False
    imIgnoreChg = False
End Sub

Private Function mTztChkForChange()
    
    On Error GoTo ErrHand
    
    mTztChkForChange = False

    If StrComp(txtName.Text, smName, vbTextCompare) <> 0 Then
        mTztChkForChange = True
        Exit Function
    End If
    If StrComp(txtGroupName.Text, smGroupName, vbTextCompare) <> 0 Then
        mTztChkForChange = True
        Exit Function
    End If
    
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in GroupNameTimeZone-mTztChkForChange: "
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
        gMsg = "A general error has occured in GroupNameTimeZone-mTrimAndFixQuotes: "
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

    Dim tzt_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If cboTimeZone.Text = "[New]" Then
        Exit Sub
    End If
    
    imIgnoreChg = True
    SQLQuery = "SELECT * "
    SQLQuery = SQLQuery + " FROM TZT"
    SQLQuery = SQLQuery + " WHERE (TztCode = " & imTztCode & ")"
    Set tzt_rst = gSQLSelectCall(SQLQuery)
    If tzt_rst.EOF Then
        gMsgBox "No matching records were found", vbOKOnly
        mClearControls
    Else
        txtName.Text = Trim$(tzt_rst!tztName)
        smName = txtName.Text
        If (gWegenerExport) Or (gOLAExport) Then
            txtGroupName.Text = Trim$(tzt_rst!tztGroupName)
            smGroupName = txtGroupName.Text
            Select Case tzt_rst!tztCSIName
                Case "EST"
                    rbcUseZone(0).Value = True
                Case "CST"
                    rbcUseZone(1).Value = True
                Case "MST"
                    rbcUseZone(2).Value = True
                Case "PST"
                    rbcUseZone(3).Value = True
                Case Else
                    rbcUseZone(0).Value = False
                    rbcUseZone(1).Value = False
                    rbcUseZone(2).Value = False
                    rbcUseZone(3).Value = False
            End Select
        Else
            txtGroupName.Text = ""
            smGroupName = txtGroupName.Text
            rbcUseZone(0).Value = False
            rbcUseZone(1).Value = False
            rbcUseZone(2).Value = False
            rbcUseZone(3).Value = False
        End If
    End If
    imIgnoreChg = False
Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmGroupNameTimeZone-mBindControls"
    imInChg = False
End Sub
