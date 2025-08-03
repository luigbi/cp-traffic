VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmCmmtSrc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Title"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   ControlBox      =   0   'False
   Icon            =   "AffCmmtSrc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frcDefault 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1875
      Width           =   3090
      Begin VB.OptionButton rbcDefault 
         Caption         =   "No"
         Height          =   210
         Index           =   1
         Left            =   1890
         TabIndex        =   8
         Top             =   0
         Width           =   705
      End
      Begin VB.OptionButton rbcDefault 
         Caption         =   "Yes"
         Height          =   210
         Index           =   0
         Left            =   1125
         TabIndex        =   7
         Top             =   0
         Width           =   840
      End
      Begin VB.Label lacDefault 
         Caption         =   "Default"
         Height          =   195
         Left            =   0
         TabIndex        =   6
         Top             =   15
         Width           =   1035
      End
   End
   Begin VB.TextBox edcSortCode 
      Height          =   330
      Left            =   1275
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1365
      Width           =   660
   End
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
      TabIndex        =   10
      Top             =   2385
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   3915
      TabIndex        =   11
      Top             =   2385
      Width           =   1320
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   435
      Left            =   510
      TabIndex        =   9
      Top             =   2385
      Width           =   1320
   End
   Begin VB.TextBox edcName 
      Height          =   330
      Left            =   1275
      MaxLength       =   40
      TabIndex        =   2
      Top             =   870
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
      FormDesignHeight=   2955
      FormDesignWidth =   5760
   End
   Begin VB.Label lacSourceCode 
      Caption         =   "Sort Code"
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   1425
      Width           =   960
   End
   Begin VB.Label lacName 
      Caption         =   "Name"
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   930
      Width           =   780
   End
End
Attribute VB_Name = "frmCmmtSrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private imCstCode As Integer
Private bmIsDirty As Boolean
Private bmIgnoreChange As Boolean
Private rst_cst As ADODB.Recordset


Private Sub cmdDone_Click()
    Dim ilRet As Integer
    
    sgCmmtSrcName = Trim(edcName.Text)
    If Not mSave() Then
        edcName.SetFocus
        Exit Sub
    End If
    igCmmtSrcReturn = True
    Unload frmCmmtSrc
End Sub

Private Sub cmdSave_Click()
    sgCmmtSrcName = Trim(edcName.Text)
    If Not mSave() Then
        edcName.SetFocus
        Exit Sub
    End If
    'bmIgnoreChange = True
    mGetCmmtSrcs
    bmIgnoreChange = False
    bmIsDirty = False
End Sub

Private Sub edcSortCode_Change()
    If bmIgnoreChange Then
        Exit Sub
    End If
    bmIsDirty = True
    cmdSave.Enabled = True
End Sub

Private Sub edcSortCode_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Activate()
    igCmmtSrcReturn = False
    edcName.SetFocus
    Call cbcSelect.SetDropDownNumRows(6)
    Call cbcSelect.SetDropDownCharWidth(20)
    frmCmmtSrc.Caption = "Comment Source"
    mGetCmmtSrcs
    If Trim(sgCmmtSrcName) = "" Then
        cbcSelect.SelText ("[New]")
        edcName.SetFocus
    End If
    bmIsDirty = False
    bmIgnoreChange = False
    cmdSave.Enabled = False
End Sub

Private Sub Form_Load()
    Me.Width = (Screen.Width) / 2.5
    Me.Height = (Screen.Height) / 4
    Me.Top = (Screen.Height - Me.Height) / 3
    Me.Left = (Screen.Width - Me.Width) / 3
    gSetFonts frmCmmtSrc
    gCenterForm frmCmmtSrc
    bmIsDirty = False
    cbcSelect.SetFont edcName.FontName, edcName.FontSize
End Sub

Private Sub mGetCmmtSrcs()
    Dim ilLoop As Integer
    Dim slName As String
    
    On Error GoTo ErrHand:
    slName = sgCmmtSrcName
    ''bmIgnoreChange = True
    'edcName.Text = ""
    cbcSelect.Clear
    cbcSelect.AddItem ("[New]")
    cbcSelect.SetItemData = -1  ' Indicates a new title
    DoEvents
    SQLQuery = "SELECT * FROM cst Order By cstName"
    Set rst_cst = gSQLSelectCall(SQLQuery)
    While Not rst_cst.EOF
        cbcSelect.AddItem (Trim(rst_cst!cstName))
        cbcSelect.SetItemData = rst_cst!cstCode
        rst_cst.MoveNext
    Wend
    ''edcName.Text = Trim(sgCmmtSrcName)
    'cbcSelect.SelText ""
    'cbcSelect.SelText (Trim(sgCmmtSrcName))
    ''cbcSelect.Text = (Trim(sgCmmtSrcName))
    For ilLoop = 0 To cbcSelect.ListCount - 1 Step 1
        If StrComp(UCase$(Trim$(cbcSelect.GetName(ilLoop))), UCase$(Trim$(slName)), vbTextCompare) = 0 Then
            cbcSelect.SetListIndex = ilLoop
            Exit For
        End If
    Next ilLoop
    If cbcSelect.ListIndex > 0 Then
        igCmmtSrcCode = cbcSelect.GetItemData(cbcSelect.ListIndex)
    Else
        igCmmtSrcCode = 0
    End If
    bmIgnoreChange = False
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CmmtSrc-mGetCmmtSrcs"
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' sgTitle = ""
    igCmmtSrcReturn = False
    Unload frmCmmtSrc
End Sub

Private Sub cbcSelect_OnChange()
    Dim ilIdx As Integer
    
    On Error GoTo ErrHand:
    If bmIgnoreChange Then
        ' Another section may have already said to ignore changes and if so exit now.
        Exit Sub
    End If
    bmIgnoreChange = True
    edcName.Text = Trim(cbcSelect.Text)
    sgCmmtSrcName = edcName.Text
    ilIdx = cbcSelect.ListIndex
    If ilIdx < 0 Then
        bmIgnoreChange = False
        Exit Sub
    End If
    igCmmtSrcCode = cbcSelect.GetItemData(ilIdx)
    If edcName.Text = "[New]" Then
        edcName.Text = ""
        edcSortCode.Text = ""
        rbcDefault(1).Value = True
    Else
        SQLQuery = "SELECT * FROM cst WHERE cstCode = " & igCmmtSrcCode
        Set rst_cst = gSQLSelectCall(SQLQuery)
        If Not rst_cst.EOF Then
            edcName.Text = Trim$(rst_cst!cstName)
            edcSortCode.Text = Trim$(Str$(rst_cst!cstSortCode))
            If rst_cst!cstDefault = "Y" Then
                rbcDefault(0).Value = True
            Else
                rbcDefault(1).Value = True
            End If
        Else
            edcName.Text = ""
            edcSortCode.Text = ""
            rbcDefault(1).Value = True
        End If
    End If
    bmIgnoreChange = False
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CmmtSrc-cbcSelect"
End Sub

Private Function mSave() As Integer
    Dim ilIdx As Integer
    Dim ilCode As Integer
    Dim ilRet As Integer
    Dim ilSortCode As Integer
    Dim slDefault As String

    If Not bmIsDirty Then
        mSave = True
        Exit Function
    End If
    mSave = False
    If Trim(edcName.Text) = "" Then
        gMsgBox "The Name cannot be blank"
        Exit Function
    End If
    If Left(edcName.Text, 5) = "[New]" Then
        gMsgBox "The Name cannot start with the word [New]"
        Exit Function
    End If
    On Error GoTo ErrHand
    ilSortCode = Val(edcSortCode.Text)
    If rbcDefault(0).Value Then
        slDefault = "Y"
    Else
        slDefault = "N"
    End If
    ilIdx = cbcSelect.ListIndex
    If ilIdx <> -1 Then
        imCstCode = cbcSelect.GetItemData(ilIdx)
    End If
    SQLQuery = "SELECT cstCode FROM cst WHERE cstName = '" & sgCmmtSrcName & "'"
    Set rst_cst = gSQLSelectCall(SQLQuery)
    If Not rst_cst.EOF Then
        If imCstCode <> rst_cst!cstCode Then
            MsgBox sgCmmtSrcName & " previously used, Enter a different name"
            Exit Function
        End If
    End If
    If slDefault = "Y" Then
        SQLQuery = "Update cst Set "
        SQLQuery = SQLQuery & "cstDefault = '" & "N" & "'"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand1:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "CmmtSrc-mSave"
            mSave = False
            Exit Function
        End If
    End If
    If imCstCode = -1 Then
        Do
            SQLQuery = "SELECT MAX(cstCode) from cst"
            Set rst_cst = gSQLSelectCall(SQLQuery)
            If IsNull(rst_cst(0).Value) Then
                ilCode = 1
            Else
                If Not rst_cst.EOF Then
                    ilCode = rst_cst(0).Value + 1
                Else
                    ilCode = 1
                End If
            End If
            ilRet = 0
            SQLQuery = "Insert Into cst ( "
            SQLQuery = SQLQuery & "cstCode, "
            SQLQuery = SQLQuery & "cstName, "
            SQLQuery = SQLQuery & "cstDefault, "
            SQLQuery = SQLQuery & "cstSortCode, "
            SQLQuery = SQLQuery & "cstUnused "
            SQLQuery = SQLQuery & ") "
            SQLQuery = SQLQuery & "Values ( "
            SQLQuery = SQLQuery & ilCode & ", "
            SQLQuery = SQLQuery & "'" & gFixQuote(sgCmmtSrcName) & "', "
            SQLQuery = SQLQuery & "'" & slDefault & "', "
            SQLQuery = SQLQuery & ilSortCode & ", "
            SQLQuery = SQLQuery & "'" & "" & "' "
            SQLQuery = SQLQuery & ") "
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand1:
                Screen.MousePointer = vbDefault
                If Not gHandleError4994("AffErrorLog.txt", "CmmtSrc-mSave") Then
                    mSave = False
                    Exit Function
                End If
                ilRet = 1
            End If
        Loop While ilRet <> 0
        igCmmtSrcCode = ilCode
    Else
        SQLQuery = "Update cst Set "
        SQLQuery = SQLQuery & "cstName = '" & gFixQuote(sgCmmtSrcName) & "', "
        SQLQuery = SQLQuery & "cstDefault = '" & slDefault & "', "
        SQLQuery = SQLQuery & "cstSortCode = " & ilSortCode & ", "
        SQLQuery = SQLQuery & "cstUnused = '" & "" & "' "
        SQLQuery = SQLQuery & "Where cstCode = " & imCstCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand1:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "CmmtSrc-mSave"
            mSave = False
            Exit Function
        End If
        igCmmtSrcCode = imCstCode
    End If
    mSave = True
    bmIsDirty = False
    cmdSave.Enabled = False
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CmmtSrc-mSave"
    Exit Function
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmCmmtSrc = Nothing
End Sub

Private Sub edcName_Change()
    If bmIgnoreChange Then
        Exit Sub
    End If
    bmIsDirty = True
    cmdSave.Enabled = True
End Sub

Private Sub rbcDefault_Click(Index As Integer)
    If bmIgnoreChange Then
        Exit Sub
    End If
    bmIsDirty = True
    cmdSave.Enabled = True
End Sub
