VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmAffRep 
   Caption         =   "Affiliate A/E"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   Icon            =   "AffRep.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   5160
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3435
      TabIndex        =   16
      Top             =   3075
      Width           =   1335
   End
   Begin VB.ComboBox cboAffAE 
      Height          =   315
      ItemData        =   "AffRep.frx":08CA
      Left            =   1410
      List            =   "AffRep.frx":08CC
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3435
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   45
      Top             =   2685
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   3645
      FormDesignWidth =   5160
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1860
      TabIndex        =   14
      Top             =   3075
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Done"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3075
      Width           =   1335
   End
   Begin VB.OptionButton optState 
      Caption         =   "Dormant"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   12
      Top             =   2595
      Width           =   1095
   End
   Begin VB.OptionButton optState 
      Caption         =   "Active"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   11
      Top             =   2595
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtEMail 
      Height          =   285
      Left            =   1680
      MaxLength       =   70
      TabIndex        =   10
      Top             =   2115
      Width           =   3135
   End
   Begin VB.TextBox txtFax 
      Height          =   285
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   8
      Top             =   1755
      Width           =   2295
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1395
      Width           =   2295
   End
   Begin VB.TextBox txtLName 
      Height          =   285
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1035
      Width           =   3135
   End
   Begin VB.TextBox txtFName 
      Height          =   285
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   2
      Top             =   675
      Width           =   3135
   End
   Begin VB.Label lblIndex 
      Height          =   255
      Left            =   105
      TabIndex        =   15
      Top             =   2715
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "E-mail Address:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2115
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Fax:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1755
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Telephone:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1395
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1035
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "First Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   660
      Width           =   1095
   End
End
Attribute VB_Name = "frmAffRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  frmAffRep - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFieldChgd As Integer
Private imState As Integer
Private imIsRepDirty As Integer
Private imInChg As Integer
Private imBSMode As Integer
Private lmArttCode As Long



Private Sub mClearControls()
    lblIndex.Caption = 0
    txtFName.Text = ""
    txtLName.Text = ""
    txtPhone.Text = ""
    txtFax.Text = ""
    txtEmail.Text = ""
    imState = 0
    imFieldChgd = False
End Sub
Private Sub mBindControls()
    lblIndex.Caption = rst(0).Value
    txtFName.Text = Trim$(rst(1).Value)
    txtLName.Text = Trim$(rst(2).Value)
    txtPhone.Text = Trim$(rst(3).Value)
    txtFax.Text = Trim$(rst(4).Value)
    txtEmail.Text = Trim$(rst(5).Value)
    imState = rst(6).Value
    imFieldChgd = False
End Sub
Private Sub cmdCancel_Click()
    Unload frmAffRep
End Sub

Private Sub cmdOk_Click()
    Dim iIndex As Integer
    Dim i As Integer
    Dim sLastName As String
    Dim sFirstName As String
    Dim sEmail As String
    
    On Error GoTo ErrHand
    If imFieldChgd = False Then
        Unload frmAffRep
        Exit Sub
    End If
    If Not mSaveRec(True) Then
        Exit Sub
    End If
    Unload frmAffRep
    Set frmAffRep = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffRep-cmdOk"
End Sub

Private Sub cmdSave_Click()
    Dim ilRet As Integer
     
    ilRet = mSaveRec(False)
    If ilRet Then
        mPopAffAE
        cboAffAE.ListIndex = 0
        cboAffAE.Text = cboAffAE.List(0)
    End If
End Sub

Private Sub Form_Load()
    Dim sName As String
    Dim sAffRepFN As String
    Dim sAffRepLN As String
    Dim i As Integer
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    Me.Width = (Screen.Width) / 1.55
    Me.Height = (Screen.Height) / 2.4
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    imInChg = False
    mPopAffAE
    cboAffAE.ListIndex = 0
    cboAffAE.Text = cboAffAE.List(0)
    If sgUstWin(13) <> "I" Then
        cmdSave.Enabled = False
        cmdOK.Enabled = False
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffRep-Form Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAffRep = Nothing
End Sub

Private Sub optState_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub txtEMail_Change()
    imFieldChgd = True
End Sub

Private Sub txtEmail_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtFax_Change()
    imFieldChgd = True
End Sub

Private Sub txtFax_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtFName_Change()
    imFieldChgd = True
End Sub

Private Sub txtFName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtLName_Change()
    imFieldChgd = True
End Sub

Private Sub txtLName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtPhone_Change()
    imFieldChgd = True
End Sub

Private Sub txtPhone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cboAffAE_Change()
    Dim iLoop As Integer
    Dim sName As String
    Dim iLen As Integer
    Dim iSel As Integer
    Dim lRow As Long

    If imInChg Then
        Exit Sub
    End If
    imInChg = True

    Screen.MousePointer = vbHourglass
    sName = LTrim$(cboAffAE.Text)
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    lRow = SendMessageByString(cboAffAE.hwnd, CB_FINDSTRING, -1, sName)
    If lRow >= 0 Then
        On Error GoTo ErrHand
        cboAffAE.ListIndex = lRow
        cboAffAE.SelStart = iLen
        cboAffAE.SelLength = Len(cboAffAE.Text)
        If cboAffAE.ListIndex <= 0 Then
            lmArttCode = 0
        Else
            lmArttCode = CLng(cboAffAE.ItemData(cboAffAE.ListIndex))
        End If
        If lmArttCode <= 0 Then
            mClearControls
            imIsRepDirty = False
        Else                                                                'Load existing station data
            SQLQuery = "SELECT * "
            SQLQuery = SQLQuery + " FROM artt"
            SQLQuery = SQLQuery + " WHERE (arttCode = " & lmArttCode & ")"
            
            Set rst = gSQLSelectCall(SQLQuery)
            If rst.EOF Then
                gMsgBox "No matching records were found", vbOKOnly
                mClearControls
            Else
                mBindControls
            End If
            imIsRepDirty = True
        End If
    Else
        lmArttCode = 0
        imIsRepDirty = False
        mClearControls
    End If
    Screen.MousePointer = vbDefault
    imInChg = False
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffRep-cboAffAE"
    imInChg = False
End Sub

Private Sub cboAffAE_Click()
    cboAffAE_Change
End Sub


Private Sub cboAffAE_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboAffAE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboAffAE.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Function mSaveRec(ilAsk As Integer) As Integer
    Dim iIndex As Integer
    Dim i As Integer
    Dim sLastName As String
    Dim sFirstName As String
    Dim sEmail As String
    Dim ilRet As Integer
    
    If sgUstWin(13) <> "I" Then
        mSaveRec = False
        gMsgBox "Not Allowed to Save.", vbOKOnly
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    On Error GoTo ErrHand:
    
    'Determine state of rep (active or dormant)
    imState = -1
    For i = 0 To 1
        If optState(i).Value Then
            imState = i
            Exit For
        End If
    Next i
    sLastName = Trim$(txtLName.Text)
    sLastName = gFixQuote(sLastName)
    sFirstName = Trim$(txtFName.Text)
    sFirstName = gFixQuote(sFirstName)
    sEmail = Trim$(txtEmail.Text)
    sEmail = gFixQuote(sEmail)
    'Add new rep
    If imIsRepDirty = False Then
        SQLQuery = "INSERT INTO artt(arttFirstName,arttLastName,arttPhone,"
        SQLQuery = SQLQuery & "arttFax,arttEmail,arttEMailRights, arttState,arttUsfCode,"
        SQLQuery = SQLQuery & "arttAddress1,arttAddress2,arttCity,arttAddressState,"
        SQLQuery = SQLQuery & "arttZip, arttCountry, arttType)"
        SQLQuery = SQLQuery & "VALUES ('" & gFixQuote(sFirstName) & "',"
        SQLQuery = SQLQuery & "'" & gFixQuote(sLastName) & "', '" & txtPhone.Text & "',"
        SQLQuery = SQLQuery & "'" & txtFax.Text & "', '" & gFixQuote(sEmail) & "',"
        SQLQuery = SQLQuery & "'" & "N" & "',"
        SQLQuery = SQLQuery & "'" & imState & "'," & igUstCode & ","
        SQLQuery = SQLQuery & "'" & "" & "',"
        SQLQuery = SQLQuery & "'" & "" & "',"
        SQLQuery = SQLQuery & "'" & "" & "',"
        SQLQuery = SQLQuery & "'" & "" & "',"
        SQLQuery = SQLQuery & "'" & "" & "',"
        SQLQuery = SQLQuery & "'" & "" & "',"
        SQLQuery = SQLQuery & "'R'" & ")"
        If ilAsk Then
            ilRet = gMsgBox("Save all changes?", vbYesNo)
        Else
            ilRet = vbYes
        End If
        If ilRet = vbYes Then
            cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "AffRep-mSaveRec"
                cnn.RollbackTrans
                mSaveRec = False
                Exit Function
            End If
            cnn.CommitTrans
        End If
           
    Else
        'UPDATE existing rep
        SQLQuery = "UPDATE artt"
        SQLQuery = SQLQuery & " SET arttFirstName = '" & sFirstName & "',"
        SQLQuery = SQLQuery & "arttLastName = '" & sLastName & "',"
        SQLQuery = SQLQuery & "arttPhone = '" & Trim$(txtPhone.Text) & "',"
        SQLQuery = SQLQuery & "arttFax = '" & Trim$(txtFax.Text) & "',"
        SQLQuery = SQLQuery & "arttEmail = '" & sEmail & "',"
        SQLQuery = SQLQuery & "arttState = " & imState & ","
        SQLQuery = SQLQuery & "arttType = '" & "R" & "',"
        SQLQuery = SQLQuery & "arttUsfCode = " & igUstCode
        SQLQuery = SQLQuery & " WHERE arttCode = " & lmArttCode
        
        If ilAsk Then
            ilRet = gMsgBox("Save all changes?", vbYesNo)
        Else
            ilRet = vbYes
        End If
        If ilRet = vbYes Then
            cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "AffRep-mSaveRec"
                cnn.RollbackTrans
                mSaveRec = False
                Exit Function
            End If
            cnn.CommitTrans
        End If
    End If
    mSaveRec = True
    Screen.MousePointer = vbDefault
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffRep-mSaveRec"
    mSaveRec = False
    Exit Function
End Function

Private Sub mPopAffAE()

    On Error GoTo ErrHand:
    
    cboAffAE.Clear
    SQLQuery = "SELECT arttFirstName, arttLastName, arttCode FROM artt Where arttType = 'R' ORDER BY arttLastName"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        cboAffAE.AddItem Trim$(rst!arttFirstName) & " " & Trim$(rst!arttLastName)
        cboAffAE.ItemData(cboAffAE.NewIndex) = rst!arttCode
        rst.MoveNext
    Wend
    cboAffAE.AddItem "[New]", 0
    cboAffAE.ItemData(cboAffAE.NewIndex) = 0
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffRep-mPopAffAE"
    Exit Sub
End Sub
