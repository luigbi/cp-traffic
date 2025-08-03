VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDepartment 
   Caption         =   "Department"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   Icon            =   "AffDepartment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton rbcType 
      Caption         =   "Other"
      Height          =   315
      Index           =   2
      Left            =   4980
      TabIndex        =   8
      Top             =   2370
      Width           =   825
   End
   Begin VB.OptionButton rbcType 
      Caption         =   "Service Rep"
      Height          =   315
      Index           =   1
      Left            =   3435
      TabIndex        =   7
      Top             =   2370
      Width           =   1965
   End
   Begin VB.OptionButton rbcType 
      Caption         =   "Market Rep"
      Height          =   315
      Index           =   0
      Left            =   1860
      TabIndex        =   6
      Top             =   2370
      Width           =   1905
   End
   Begin VB.CommandButton cmcColor 
      Caption         =   "Select Color"
      Height          =   375
      Left            =   225
      TabIndex        =   3
      Top             =   1740
      Width           =   2475
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   405
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2535
      TabIndex        =   10
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmcSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4605
      TabIndex        =   11
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox edcName 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   1875
      MaxLength       =   30
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.ComboBox cbcDepartment 
      Height          =   315
      Left            =   1875
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog cdcColor 
      Left            =   5790
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lacType 
      Caption         =   "Department Type"
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   1725
   End
   Begin VB.Label lacSample 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sample"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3120
      TabIndex        =   4
      Top             =   1770
      Width           =   1920
   End
   Begin VB.Label lacName 
      Caption         =   "Department Name:"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1875
   End
End
Attribute VB_Name = "frmDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmDepartment
'*
'*  Created October,2005 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc. 2005
'*
'******************************************************
Option Explicit
Option Compare Text

Private imInChg As Integer
Private imBSMode As Integer
Private imSave As Integer
Private imDntCode As Integer
Private smSvName As String
Private lmSvColor As Long
Private lmColor As Long
Private smSvType As String
Private imIgnoreChg As Integer
Private dnt_rst As ADODB.Recordset


Private Sub cbcDepartment_Change()
    
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
    slName = LTrim$(cbcDepartment.Text)
    ilLen = Len(slName)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(cbcDepartment.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        On Error GoTo ErrHand
        cbcDepartment.ListIndex = llRow
        cbcDepartment.SelStart = ilLen
        cbcDepartment.SelLength = Len(cbcDepartment.Text)
        If cbcDepartment.ListIndex < 0 Then
            imDntCode = -1
        ElseIf (cbcDepartment.ListIndex = 0) Then
            imDntCode = 0
        Else
            imDntCode = CInt(cbcDepartment.ItemData(cbcDepartment.ListIndex))
        End If
        If imDntCode < 0 Then
            mClearControls
        ElseIf (imDntCode = 0) Then
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
    gHandleError "AffErrorLog.txt", "Department-cbcDepartment_Change"
    imInChg = False
End Sub

Private Sub cbcDepartment_Click()
    If imInChg Then
        Exit Sub
    End If
    cbcDepartment_Change
End Sub

Private Sub cbcDepartment_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cbcDepartment_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cbcDepartment_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cbcDepartment.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cmcCancel_Click()
    igGNMarketReturnCode = 0
    igIndex = cbcDepartment.ListIndex
    Unload frmDepartment
End Sub

Private Sub cmcColor_Click()
    cdcColor.CancelError = True
    On Error GoTo ErrHandle
    cdcColor.Flags = cdlCCRGBInit
    cdcColor.ShowColor
    lmColor = cdcColor.Color
    lacSample.BackColor = lmColor
    Exit Sub
ErrHandle:
    'User pressed cancel
    Exit Sub
End Sub

Private Sub cmcDone_Click()
    
    Dim ilRet As Integer
    Dim ilIndex As Integer

    ilRet = mDntChkForChange()
    If ilRet Then
        ilRet = gMsgBox("Would You Like To Save the Changes", vbYesNo)
        If ilRet = vbYes Then
            ilRet = mSave()
            If Not ilRet Then
                Exit Sub
            End If
        End If
    End If
    ilIndex = cbcDepartment.ListIndex
    igDepartmentReturn = True
    igDepartmentReturnCode = 0
    sgDepartmentName = ""
    If ilIndex > 0 Then
        igDepartmentReturnCode = cbcDepartment.ItemData(ilIndex)
        sgDepartmentName = cbcDepartment.List(ilIndex)
    End If
    Unload frmDepartment
End Sub

Private Sub cmcSave_Click()
    Call mSave
End Sub

Private Sub Form_Activate()
    If cbcDepartment.ListIndex >= 0 Then
        edcName.SetFocus
    End If
End Sub

Private Sub Form_Initialize()
    'gSetFonts frmDepartment
    gCenterForm frmDepartment
End Sub

Private Sub Form_Load()
    
    Form_Initialize
    imSave = False
    Call mFileDntInfo(sgDepartmentName)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    dnt_rst.Close
    
    Set frmDepartment = Nothing

End Sub

Private Function mFileDntInfo(slFromSaveName As String)
    
    Dim slName As String
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    
    If imSave Then
        slName = Trim$(slFromSaveName)
    End If
    If cbcDepartment.Text = "[New]" And (Not imSave) Then
        cbcDepartment.ListIndex = 0
        mClearControls
        Exit Function
    End If
    cbcDepartment.Clear
    SQLQuery = "SELECT * FROM dnt ORDER BY dntName"
    Set dnt_rst = gSQLSelectCall(SQLQuery)
    While Not dnt_rst.EOF
        cbcDepartment.AddItem Trim$(dnt_rst!dntName)
        cbcDepartment.ItemData(cbcDepartment.NewIndex) = dnt_rst!dntCode
        dnt_rst.MoveNext
    Wend
    cbcDepartment.AddItem "[New]", 0
    cbcDepartment.ItemData(cbcDepartment.NewIndex) = 0

    If Not imSave Then
        slName = sgDepartmentName
    End If
    ilRow = SendMessageByString(cbcDepartment.hwnd, CB_FINDSTRING, -1, slName)
    If (ilRow > 0) Then
        cbcDepartment.ListIndex = ilRow
        mBindControls
    Else
        cbcDepartment.ListIndex = 0
        mClearControls
        If slName <> "" Then
            edcName = slName
        End If
    End If
    DoEvents
    Screen.MousePointer = vbDefault
    imInChg = False
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Department-mFileDntInfo"
    imInChg = False
End Function

Private Function mSave()
    
    Dim ilRet As Integer
    Dim ilRow As Integer
    Dim slName As String
    Dim slType As String
    
    On Error GoTo ErrHand
    
    mSave = False
    ilRet = mDntChkForChange()
    slName = Trim$(edcName.Text)
    If Trim$(slName) = "" Then
        gMsgBox "Department Name must be Defined.", vbOKOnly
        Exit Function
    End If
    If (rbcType(0).Value = False) And (rbcType(1).Value = False) And (rbcType(2).Value = False) Then
        gMsgBox "Department Type must be Defined.", vbOKOnly
        Exit Function
    End If
    If rbcType(0).Value Then
        slType = "M"
    ElseIf rbcType(1).Value Then
        slType = "S"
    ElseIf rbcType(2).Value Then
        slType = "O"
    End If
    If cbcDepartment.Text = "[New]" Or StrComp(smSvName, slName, vbBinaryCompare) > 0 Then
        SQLQuery = "select dntCode from dnt where UCase(dntName) = '" & UCase(slName) & "'"
        Set dnt_rst = gSQLSelectCall(SQLQuery)
        If Not dnt_rst.EOF Then
            gMsgBox "The Department " & slName & " name already exists"
            mSave = False
            Exit Function
        End If
    End If
    
    'ilRow = SendMessageByString(cbcDepartment.hwnd, CB_FINDSTRING, -1, slName)
    ilRow = cbcDepartment.ListIndex
    If ilRow >= 0 Then
        imDntCode = cbcDepartment.ItemData(ilRow)
    Else
        imDntCode = 0
    End If
    mSave = False
    If cbcDepartment.Text = "[New]" Then
        SQLQuery = "Insert into dnt "
        SQLQuery = SQLQuery & "(dntName, dntColor, dntType, dntUnused) "
        SQLQuery = SQLQuery & " VALUES ('" & edcName.Text & "'," & lmColor & ",'" & slType & "'," & "''" & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "Department-mSave"
            mSave = False
            imInChg = False
            Exit Function
        End If
        mClearControls
    Else
        SQLQuery = "UPDATE dnt"
        SQLQuery = SQLQuery & " SET dntName = '" & edcName.Text & "',"
        SQLQuery = SQLQuery & "dntColor = " & lmColor & ","
        SQLQuery = SQLQuery & "dntType = '" & slType & "'"
        SQLQuery = SQLQuery & " WHERE dntCode = " & imDntCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "Department-mSave"
            mSave = False
            imInChg = False
            Exit Function
        End If
    End If
    'rst.Close
    imSave = True
    ilRet = mFileDntInfo(slName)
    imSave = False
    mSave = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmDepartment-mSave"
    imInChg = False
End Function

Private Sub mClearControls()

    imIgnoreChg = True
    edcName.Text = ""
    smSvName = ""
    lacSample.BackColor = vbWhite
    lmSvColor = 0
    rbcType(0).Value = False
    rbcType(1).Value = False
    rbcType(2).Value = False
    smSvType = ""
    imIgnoreChg = False
End Sub

Private Function mDntChkForChange()
    Dim slType As String
    
    On Error GoTo ErrHand
    
    mDntChkForChange = False

    If StrComp(edcName.Text, smSvName, vbTextCompare) <> 0 Then
        mDntChkForChange = True
        Exit Function
    End If
    If lmSvColor <> lmColor Then
        mDntChkForChange = True
        Exit Function
    End If
    slType = ""
    If rbcType(0).Value Then
        slType = "M"
    ElseIf rbcType(1).Value Then
        slType = "S"
    ElseIf rbcType(2).Value Then
        slType = "O"
    End If
    If smSvType <> slType Then
        mDntChkForChange = True
        Exit Function
    End If
    
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Department-mDntChkForChange: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    imInChg = False

End Function

Private Sub edcName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub



Private Sub mBindControls()

    On Error GoTo ErrHand
    
    If cbcDepartment.Text = "[New]" Then
        Exit Sub
    End If
    
    imIgnoreChg = True
    SQLQuery = "SELECT * "
    SQLQuery = SQLQuery + " FROM dnt"
    SQLQuery = SQLQuery + " WHERE (dntCode = " & imDntCode & ")"
    Set dnt_rst = gSQLSelectCall(SQLQuery)
    If dnt_rst.EOF Then
        gMsgBox "No matching records were found", vbOKOnly
        mClearControls
    Else
        edcName.Text = Trim$(dnt_rst!dntName)
        smSvName = edcName.Text
        lmColor = dnt_rst!dntColor
        lmSvColor = lmColor
        lacSample.BackColor = lmColor
        If dnt_rst!dntType = "M" Then
            rbcType(0).Value = True
        ElseIf dnt_rst!dntType = "S" Then
            rbcType(1).Value = True
        ElseIf dnt_rst!dntType = "O" Then
            rbcType(2).Value = True
        Else
            rbcType(0).Value = False
            rbcType(1).Value = False
            rbcType(2).Value = False
        End If
        smSvType = dnt_rst!dntType
    End If
    imIgnoreChg = False
Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmDepartment-mBindControls"
    imInChg = False
End Sub

Private Sub rbcType_Click(Index As Integer)
    If (Index = 0) And (rbcType(Index).Value) Then
        If ((Asc(sgSpfUsingFeatures9) And AFFILIATECRM) <> AFFILIATECRM) Then
            rbcType(0).Value = False
            MsgBox "This is a paid feature which has not been activated, Call Counterpoint to allow Affiliate Salespeople to use the Affiliate Management Screen"
        End If
    End If
End Sub
