VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form EngrFollow 
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   8490
   ControlBox      =   0   'False
   Icon            =   "EngrFollow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   8490
   Begin VB.CommandButton cmcErase 
      Caption         =   "&Erase"
      Height          =   375
      Left            =   5235
      TabIndex        =   12
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame frcDefine 
      Caption         =   "Step 2: Define Follow Properties"
      Height          =   1920
      Left            =   330
      TabIndex        =   2
      Top             =   885
      Width           =   7620
      Begin VB.Frame frcState 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   1305
         Width           =   2220
         Begin VB.OptionButton rbcState 
            Caption         =   "Active"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton rbcState 
            Caption         =   "Dormant"
            Height          =   255
            Index           =   1
            Left            =   975
            TabIndex        =   9
            Top             =   0
            Width           =   990
         End
      End
      Begin VB.TextBox edcName 
         Height          =   285
         Left            =   1665
         MaxLength       =   19
         TabIndex        =   4
         Top             =   405
         Width           =   2145
      End
      Begin VB.TextBox edcDescription 
         Height          =   285
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   6
         Top             =   825
         Width           =   5835
      End
      Begin VB.Label lacName 
         Caption         =   "Automation Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   405
         Width           =   1545
      End
      Begin VB.Label lacDescription 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   825
         Width           =   1380
      End
   End
   Begin VB.Frame frcSelect 
      Caption         =   "Step 1: Select Follow"
      Height          =   660
      Left            =   330
      TabIndex        =   0
      Top             =   120
      Width           =   3555
      Begin VB.ComboBox cbcSelect 
         BackColor       =   &H00FFFF80&
         Height          =   315
         ItemData        =   "EngrFollow.frx":030A
         Left            =   150
         List            =   "EngrFollow.frx":030C
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   3180
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   345
      Top             =   2925
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   3750
      FormDesignWidth =   8490
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3495
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmmDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   1710
      TabIndex        =   10
      Top             =   3135
      Width           =   1335
   End
End
Attribute VB_Name = "EngrFollow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrFollow - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFieldChgd As Integer
Private imState As Integer


Private Sub ClearControls()
    edcName.Text = ""
    edcDescription.Text = ""
    rbcState(0).Value = False
    rbcState(1).Value = False
    imState = 0
    imFieldChgd = False
End Sub
Private Sub BindControls()
End Sub
Private Sub cmcCancel_Click()
    Unload EngrFollow
End Sub

Private Sub cmcDone_Click()
    Dim iIndex As Integer
    Dim i As Integer
    Dim sLastName As String
    Dim sFirstName As String
    Dim sEMail As String
    
    On Error GoTo ErrHand
    If imFieldChgd = False Then
        Unload EngrFollow
        Exit Sub
    End If
'    Screen.MousePointer = vbHourglass
'
'    'Determine state of rep (active or dormant)
'    imState = -1
'    For i = 0 To 1
'        If optState(i).Value Then
'            imState = i
'            Exit For
'        End If
'    Next i
'    sLastName = Trim$(txtLName.Text)
'    sLastName = gFixQuote(sLastName)
'    sFirstName = Trim$(txtFName.Text)
'    sFirstName = gFixQuote(sFirstName)
'    sEMail = Trim$(txtEMail.Text)
'    sEMail = gFixQuote(sEMail)
'    'Add new rep
'    If IsRepDirty = False Then
'        SQLQuery = "INSERT INTO artt(arttFirstName,arttLastName,arttPhone,"
'        SQLQuery = SQLQuery & "arttFax,arttEmail,arttState,arttUsfCode,"
'        SQLQuery = SQLQuery & "arttAddress1,arttAddress2,arttCity,arttAddressState,"
'        SQLQuery = SQLQuery & "arttZip, arttCountry, arttType)"
'        SQLQuery = SQLQuery & "VALUES ('" & sFirstName & "',"
'        SQLQuery = SQLQuery & "'" & sLastName & "', '" & txtPhone.Text & "',"
'        SQLQuery = SQLQuery & "'" & txtFax.Text & "', '" & sEMail & "',"
'        SQLQuery = SQLQuery & "'" & imState & "'," & igUstCode & ","
'        SQLQuery = SQLQuery & "'" & "" & "',"
'        SQLQuery = SQLQuery & "'" & "" & "',"
'        SQLQuery = SQLQuery & "'" & "" & "',"
'        SQLQuery = SQLQuery & "'" & "" & "',"
'        SQLQuery = SQLQuery & "'" & "" & "',"
'        SQLQuery = SQLQuery & "'" & "" & "',"
'        SQLQuery = SQLQuery & "'R'" & ")"
'
'        If MsgBox("Save all changes?", vbYesNo) = vbYes Then
'            env.BeginTrans
'            cnn.Execute SQLQuery, rdExecDirect
'            env.CommitTrans
'            SQLQuery = "Select MAX(arttCode) from artt"
'            Set rst = cnn.OpenResultset(SQLQuery)
'            iARIndex = rst(0).Value
'        Else
'            iARIndex = -1
'        End If
'
'    Else
'        'UPDATE existing rep
'        SQLQuery = "UPDATE artt"
'        SQLQuery = SQLQuery & " SET arttFirstName = '" & sFirstName & "',"
'        SQLQuery = SQLQuery & "arttLastName = '" & sLastName & "',"
'        SQLQuery = SQLQuery & "arttPhone = '" & Trim$(txtPhone.Text) & "',"
'        SQLQuery = SQLQuery & "arttFax = '" & Trim$(txtFax.Text) & "',"
'        SQLQuery = SQLQuery & "arttEmail = '" & sEMail & "',"
'        SQLQuery = SQLQuery & "arttState = " & imState & ","
'        SQLQuery = SQLQuery & "arttType = '" & "R" & "',"
'        SQLQuery = SQLQuery & "arttUsfCode = " & igUstCode
'        SQLQuery = SQLQuery & " WHERE arttCode = '" & iARIndex & "'"
'
'        If MsgBox("Save all changes?", vbYesNo) = vbYes Then
'            env.BeginTrans
'            cnn.Execute SQLQuery, rdExecDirect
'            env.CommitTrans
'        End If
'    End If

    Unload EngrFollow
    Set EngrFollow = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In rdoErrors
        If gErrSQL.Number <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Bus Definition-cmcDone: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number, vbCritical
        End If
    Next gErrSQL
    env.RollbackTrans
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Bus Definition-cmcDone: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrFollow
    gCenterFormModal EngrFollow
End Sub

Private Sub Form_Load()
    Dim sName As String
    Dim sAffRepFN As String
    Dim sAffRepLN As String
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In rdoErrors
        If gErrSQL.Number <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Bus Definition-Form Load: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Bus Definition-Form Load: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set EngrFollow = Nothing
End Sub

Private Sub rbcState_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub edcName_Change()
    imFieldChgd = True
End Sub

Private Sub edcName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcCtrl_Change()
    imFieldChgd = True
End Sub

Private Sub edcCtrl_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDescription_Change()
    imFieldChgd = True
End Sub

Private Sub edcDescription_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcChannel_Change()
    imFieldChgd = True
End Sub

Private Sub edcChannel_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

