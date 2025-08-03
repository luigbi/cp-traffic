VERSION 5.00
Begin VB.Form frmWebVendors 
   Caption         =   "Vendor Setup"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ckcAutoVending 
      Caption         =   "Allow vendor to update delivery status"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   840
      Width           =   3255
   End
   Begin VB.CheckBox ckcAutoPosting 
      Caption         =   "Allow Auto-Posting"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.OptionButton rbcStatus 
      Caption         =   "Dormant"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.OptionButton rbcStatus 
      Caption         =   "Active"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox cboDelivery 
      Height          =   315
      Left            =   120
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   6240
      Width           =   2895
   End
   Begin VB.ComboBox cboExport 
      Height          =   315
      Left            =   2160
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   6120
      Width           =   2895
   End
   Begin VB.ComboBox cboImport 
      Height          =   315
      Left            =   4440
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   6120
      Width           =   2895
   End
   Begin VB.ComboBox cboHierarchy 
      Height          =   315
      Left            =   7440
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   6240
      Width           =   2895
   End
   Begin VB.CheckBox ckcUpdates 
      Caption         =   "Send only updates on change"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CheckBox ckcOverride 
      Caption         =   "Allow stations to override vendor times"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Frame frcStationConnection 
      Caption         =   "Station Connection Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5160
      TabIndex        =   19
      Top             =   3360
      Width           =   4815
      Begin VB.TextBox txtStationPassword 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtStationUserName 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lsc2 
         Caption         =   "Password"
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   735
         Width           =   1335
      End
      Begin VB.Label lsc1 
         Caption         =   "Sign-On Name"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   440
         Width           =   1455
      End
   End
   Begin VB.Frame frcVendorConnection 
      Caption         =   "Vendor Connection Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5160
      TabIndex        =   14
      Top             =   1680
      Width           =   4815
      Begin VB.TextBox txtAddress 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtVendorPassword 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtVendorUsername 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label le3 
         Caption         =   "Sign-On Name"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   640
         Width           =   1455
      End
      Begin VB.Label le2 
         Caption         =   "Address"
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   350
         Width           =   1575
      End
      Begin VB.Label le4 
         Caption         =   "Password"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   975
         Width           =   1455
      End
   End
   Begin VB.ComboBox cboFormat 
      Height          =   315
      ItemData        =   "AffVendors.frx":0000
      Left            =   6000
      List            =   "AffVendors.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4815
      TabIndex        =   13
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   6645
      TabIndex        =   15
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Vendor Type"
      Height          =   255
      Left            =   600
      TabIndex        =   37
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lbcType 
      Caption         =   "Type"
      Height          =   255
      Left            =   2040
      TabIndex        =   36
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lbcHierarchy 
      Caption         =   "Hierarchy"
      Height          =   255
      Left            =   2040
      TabIndex        =   35
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lbcVendorToWeb 
      Caption         =   "VendorToWeb"
      Height          =   255
      Left            =   2040
      TabIndex        =   34
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lbcWebToVendor 
      Caption         =   "WebtoVendor"
      Height          =   255
      Left            =   2040
      TabIndex        =   33
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lbcDelivery 
      Caption         =   "Delivery"
      Height          =   255
      Left            =   2040
      TabIndex        =   32
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Status"
      Height          =   255
      Left            =   600
      TabIndex        =   31
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lbcCode 
      Caption         =   "lbcCode"
      Height          =   255
      Left            =   9960
      TabIndex        =   30
      Top             =   360
      Width           =   615
   End
   Begin VB.Label ls6 
      Caption         =   "Delivery Method"
      Height          =   255
      Left            =   600
      TabIndex        =   29
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label ls3 
      Caption         =   "Web-to-Vendor"
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label ls4 
      Caption         =   "Vendor-to-Web"
      Height          =   255
      Left            =   600
      TabIndex        =   26
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label ls5 
      Caption         =   "Hierarchy"
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "frmWebVendors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmWebVendors
'*
'*  Created October,2015 by Dan Michaelson
'*  copied frmGroupName
'*  Copyright Counterpoint Software, Inc. 2015
'*
'******************************************************
Option Explicit
Option Compare Text
Const FORMNAME As String = "frmWebVendors"
Private Const ACTIVE As Integer = 0
Private Const DORMANT As Integer = 1
Private bmIsEditable As Boolean
Private imInChg As Integer
Private imBSMode As Integer
Private tmMyVendor As VendorInfo
Dim tmMyAvailableVendors() As VendorInfo
Private myRs As ADODB.Recordset




Private Sub Form_Initialize()
   ' gSetFonts frmWebVendors
    gCenterForm frmWebVendors
End Sub

Private Sub Form_Load()
    Form_Initialize
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not myRs Is Nothing Then
        If (myRs.State And adStateOpen) <> 0 Then
            myRs.Close
        End If
        Set myRs = Nothing
    End If
    Erase tmMyAvailableVendors
End Sub
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
    'this is my last find, which may be set in mBindControls.  Clear for next one
    If Not myRs Is Nothing Then
        If (myRs.State And adStateOpen) <> 0 Then
            myRs.Close
        End If
    End If
    If llRow >= 0 Then
On Error GoTo ErrHand
        cboFormat.ListIndex = llRow
        cboFormat.SelStart = ilLen
        cboFormat.SelLength = Len(cboFormat.Text)
        If cboFormat.ListIndex < 0 Then
            mClearControls
        Else
             mBindControls
        End If
    Else
        mClearControls
    End If
Cleanup:
    Screen.MousePointer = vbDefault
    imInChg = False
    Exit Sub
ErrHand:
    gHandleError "", FORMNAME & "-cboFormat_Change"
    GoTo Cleanup
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
    Unload Me
End Sub

Private Sub cmdDone_Click()
    
    If mIsChange() Then
        If gMsgBox("Would You Like To Save the Changes", vbYesNo) = vbYes Then
            If Not mSave() Then
                Exit Sub
            End If
        Else
            'dan 6/28.  they said no, so let them leave
            'Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If mIsChange() Then
        If rbcStatus(ACTIVE).Value = True Then
            mSave
        Else
            mDeactivate
        End If
    End If
    cboFormat.SetFocus
End Sub
'methods

Private Sub mInit()
    imInChg = True
    If igPasswordOk Then
        bmIsEditable = True
    Else
        mLockDown
        bmIsEditable = False
        cmdSave.Enabled = False
    End If
    'not currently implemented
    ckcUpdates.Visible = False
    lbcCode.Visible = False
    cboDelivery.Visible = False
    cboImport.Visible = False
    cboExport.Visible = False
    cboHierarchy.Visible = False
    If (StrComp(sgUserName, "Guide", 1) = 0) Then
        If Not bgLimitedGuide Then
            lbcCode.Visible = True
            lbcCode.ForeColor = vbBlue
            lbcCode.Caption = "0"
        End If
    End If
    mFillVendors
    lbcDelivery.BackColor = vbYellow
    lbcHierarchy.BackColor = vbYellow
    lbcWebToVendor.BackColor = vbYellow
    lbcVendorToWeb.BackColor = vbYellow
    lbcType.BackColor = vbYellow
    With cboExport
        .AddItem "None"
        .ItemData(.NewIndex) = MethodType.None
        .AddItem "csiStationService"
        .ItemData(.NewIndex) = MethodType.csiStationService
        .AddItem "csiVendorService"
        .ItemData(.NewIndex) = MethodType.csiVendorService
        .AddItem "Both Services"
        .ItemData(.NewIndex) = MethodType.BothServices
        '.Text = "None"
    End With
    With cboImport
        .AddItem "None"
        .ItemData(.NewIndex) = MethodType.None
        .AddItem "csiStationService"
        .ItemData(.NewIndex) = MethodType.csiStationService
        .AddItem "csiVendorService"
        .ItemData(.NewIndex) = MethodType.csiVendorService
        .AddItem "Both Services"
        .ItemData(.NewIndex) = MethodType.BothServices
        '.Text = "None"
    End With
    With cboHierarchy
        .AddItem "None"
        .ItemData(.NewIndex) = 0
        .AddItem "Pre-Automation"
        .ItemData(.NewIndex) = 1
        .AddItem "Automation"
        .ItemData(.NewIndex) = 2
        .AddItem "Heard"
        .ItemData(.NewIndex) = 3
        '.Text = "None"
    End With
    With cboDelivery
         .AddItem "None"
         '9204 from 44 to 2
         .ItemData(.NewIndex) = "2"
        .AddItem "Log"
        .ItemData(.NewIndex) = "1"
        .AddItem "Audio"
        .ItemData(.NewIndex) = "0"
       ' .Text = "None"
    End With
    imInChg = False
    'imInChg set in here
    mClearControls
    mEnableControls
    cboFormat.ListIndex = 0
End Sub
Private Sub mLockDown()
    rbcStatus(0).Enabled = False
    rbcStatus(1).Enabled = False
'    cboExport.Locked = True
'    cboExport.BackColor = vbYellow
'    cboImport.Locked = True
'    cboImport.BackColor = vbYellow
'    cboHierarchy.Locked = True
'    cboHierarchy.BackColor = vbYellow
'    cboDelivery.Locked = True
'    cboDelivery.BackColor = vbYellow
End Sub

Private Sub mSetType()
    With tmMyVendor
        'I'll test this later to see everything ok!
        .iIdCode = 0
        '8949
        If IsNumeric(lbcCode.Caption) Then
            .iIdCode = CInt(lbcCode.Caption)
        End If
        '8949 added 'or'.  note that they can be out of sync when leaving, but idcode points to current
        If cboFormat.ListIndex > -1 Or .iIdCode > 0 Then
          '  .iIdCode = cboFormat.ItemData(cboFormat.ListIndex)
            .sName = cboFormat.Text
            .sApprovalPassword = mApprovalPasswordForWebVendor(.iIdCode)
        End If
        .iHierarchy = 0
        If cboHierarchy.ListIndex > -1 Then
            .iHierarchy = cboHierarchy.ItemData(cboHierarchy.ListIndex)
        End If
        .iExportMethod = 0
        If cboExport.ListIndex > -1 Then
            .iExportMethod = cboExport.ItemData(cboExport.ListIndex)
        End If
        .iImportMethod = 0
        If cboImport.ListIndex > -1 Then
            .iImportMethod = cboImport.ItemData(cboImport.ListIndex)
        End If
        '8949
 '       .iIdCode = 0
'        If IsNumeric(lbcCode.Caption) Then
'            .iIdCode = CInt(lbcCode.Caption)
'        End If
        .sDeliveryType = "L"
        If cboDelivery.ListIndex > -1 Then
            If cboDelivery.ItemData(cboDelivery.ListIndex) = DeliveryType.Audio Then
                .sDeliveryType = "A"
            End If
        End If
        .sAddress = txtAddress.Text
        .sVendorPassword = txtVendorPassword.Text
        .sVendorUserName = txtVendorUsername.Text
        .sStationPassword = txtStationPassword.Text
        .sStationUserName = txtStationUserName.Text
        .sIsOverridable = ""
        If ckcOverride.Value = vbChecked Then
            .sIsOverridable = "Y"
        End If
        .sSendUpdatesOnly = ""
        If ckcUpdates.Value = vbChecked Then
            .sSendUpdatesOnly = "Y"
        End If
        '8457
        If ckcAutoPosting.Value = vbChecked Then
            .sAllowAutoPosting = "Y"
        Else
            .sAllowAutoPosting = "N"
        End If
        '8862
        .sAutoVendor = ""
        If ckcAutoVending.Value = vbChecked Then
            .sAutoVendor = "Y"
        End If
        .bIsActive = rbcStatus(ACTIVE).Value
    End With
End Sub
Private Sub mEnableControls()

'    cboDelivery.Enabled = False
'    cboExport.Enabled = False
'    cboImport.Enabled = False
'    cboHierarchy.Enabled = False
'    ckcOverride.Enabled = False
'    ckcUpdates.Enabled = False
'    ls3.Enabled = False
'    ls4.Enabled = False
'    ls5.Enabled = False
'    ls6.Enabled = False
    lsc1.Enabled = False
    lsc2.Enabled = False
    le2.Enabled = False
    le3.Enabled = False
    le4.Enabled = False
    mEnableConnections True, False
    mEnableConnections False, False
    '8457
    mEnableChoicesForBothWebVendors False
End Sub
Private Sub mEnableConnections(blIsVendor As Boolean, blValue As Boolean)
    If Not bmIsEditable Then
        blValue = False
    End If
    If blIsVendor Then
        frcVendorConnection.Enabled = blValue
        txtVendorUsername.Enabled = blValue
        txtAddress.Enabled = blValue
        txtVendorPassword.Enabled = blValue
        le2.Enabled = blValue
        le3.Enabled = blValue
        le4.Enabled = blValue
    Else
        frcStationConnection.Enabled = blValue
        txtStationUserName.Enabled = blValue
        txtStationPassword.Enabled = blValue
        lsc1.Enabled = blValue
        lsc2.Enabled = blValue
    End If
End Sub
Private Sub mEnableChoicesForBothWebVendors(blValue As Boolean)
    '8457
    ckcAutoPosting.Enabled = blValue
    ckcOverride.Enabled = blValue
    ckcUpdates.Enabled = blValue
    '8950 9009 removed
  '  ckcAutoVending.Enabled = blValue
End Sub
Private Sub mFillVendors()
 'I don't mess with imChg here: set before coming in!
    Dim ilCount As Integer
    
    On Error GoTo ErrHand
    'set bmIsjelli to true or false, etc.
    tmMyAvailableVendors = gGetAvailableVendors()
    cboFormat.Clear
    For ilCount = 0 To UBound(tmMyAvailableVendors) - 1
        cboFormat.AddItem tmMyAvailableVendors(ilCount).sName
        cboFormat.ItemData(cboFormat.NewIndex) = tmMyAvailableVendors(ilCount).iIdCode
    Next ilCount
Exit Sub
ErrHand:
    gHandleError "", FORMNAME & "-mFillVendors"
End Sub

Private Function mSave() As Boolean
    
    Dim blRet As Boolean
    Dim slIssue As String
    Dim blIsNew As Boolean
    Dim ilVendor As Vendors
    
    On Error GoTo ErrHand
    blIsNew = False
    blRet = True
    If Not mValidate(slIssue) Then
        blRet = False
        gMsgBox slIssue, vbOKOnly
        GoTo Cleanup
    '8355
    Else
        If Not mUserToContinueAfterWebReadyWarnings() Then
            blRet = False
            GoTo Cleanup
        End If
    End If
    If Not myRs Is Nothing Then
        'insert
        If myRs.EOF Then
            blIsNew = True
            SQLQuery = mInsertOrUpdateSql(False)
        Else
            SQLQuery = mInsertOrUpdateSql(True)
        End If
    End If
    If Len(SQLQuery) = 0 Then
        gMsgBox "Problem building sql in mSave.  Save failed.", vbOKOnly
        blRet = False
        GoTo Cleanup
    End If
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", FORMNAME & "-mSave"
        mBindControls
        mSave = False
        Exit Function
    End If
    rbcStatus(ACTIVE).Value = True
    If blIsNew Then
        '8156
        ilVendor = tmMyVendor.iIdCode
        gAdjustAllowedExportsImports ilVendor
        If Not mConvert() Then
            gMsgBox "Problem converting in mConvert.  Conversion failed.", vbOKOnly
            blRet = False
            GoTo Cleanup
        End If
    '8322
    ElseIf tmMyVendor.iIdCode = Vendors.NetworkConnect Then
        ilVendor = tmMyVendor.iIdCode
        gAdjustAllowedExportsImports ilVendor
    End If
Cleanup:
    '10467
    If Not bgVendorToWebAllowed Then
        gVendorToWebAllowed
    End If
    mBindControls
    mSave = blRet
    Exit Function
ErrHand:
    blRet = False
    gHandleError "", FORMNAME & "-mSave"
    GoTo Cleanup
End Function
Private Function mUserToContinueAfterWebReadyWarnings() As Boolean
    Dim blRet As Boolean
    Dim blStation As Boolean
    Dim blVendor As Boolean
    Dim slWarning As String
    Dim ilResponse As Integer
    
    blRet = True
    If rbcStatus(ACTIVE).Value = True Then
         With cboImport
            If .ListIndex > -1 Then
                If .ItemData(.ListIndex) = MethodType.csiStationService Or .ItemData(.ListIndex) = MethodType.BothServices Then
                    blStation = True
                End If
                If .ItemData(.ListIndex) = MethodType.csiVendorService Or .ItemData(.ListIndex) = MethodType.BothServices Then
                    blVendor = True
                End If
            End If
        End With
        With cboExport
            If .ListIndex > -1 Then
                If .ItemData(.ListIndex) = MethodType.csiStationService Or .ItemData(.ListIndex) = MethodType.BothServices Then
                    blStation = True
                End If
                If .ItemData(.ListIndex) = MethodType.csiVendorService Or .ItemData(.ListIndex) = MethodType.BothServices Then
                    blVendor = True
                End If
            End If
        End With
        If blVendor Or blStation Then
            If blVendor Then
                blVendor = False
                If Len(txtAddress) = 0 Or Len(txtVendorUsername) = 0 Or Len(txtVendorPassword) = 0 Then
                    blVendor = True
                End If
            End If
            If blStation Then
                blStation = False
                If Len(txtStationUserName) = 0 Or Len(txtStationPassword) = 0 Then
                    blStation = True
                End If
            End If
            'if true, we need to warn!
            If blStation Or blVendor Then
                slWarning = "This Web Vendor is active, but is not ready to export/import from web until "
                If blStation And blVendor Then
                    slWarning = slWarning & "'Vendor Connection Info' and 'Station Connection Info' values are filled out.  Continue with save anyway?"
                ElseIf blStation Then
                    slWarning = slWarning & "'Station Connection Info' values are filled out.  Continue with save anyway?"
                ElseIf blVendor Then
                    slWarning = slWarning & "'Vendor Connection Info' values are filled out.  Continue with save anyway?"
                End If
                ilResponse = MsgBox(slWarning, vbOKCancel, "Not Ready to send from web")
                If ilResponse = vbCancel Then
                    blRet = False
                End If
            End If
        End If
    End If 'is active (don't test if dormant!)
    mUserToContinueAfterWebReadyWarnings = blRet
End Function
Private Function mInsertOrUpdateSql(blIsUpdate As Boolean) As String
    Dim slRet As String
On Error GoTo errbox
    With tmMyVendor
        '8457 8862
        If blIsUpdate Then
            slRet = "Update WVT_Vendor_Table SET wvtVendorUserName = '" & Trim$(gFixQuote(.sVendorUserName)) & "', wvtVendorPassword = '" & Trim$(gFixQuote(.sVendorPassword)) & "', wvtAddress = '" & Trim$(gFixQuote(.sAddress)) & "', "
            slRet = slRet & " wvtStationUserName = '" & Trim$(gFixQuote(.sStationUserName)) & "', wvtStationPassword = '" & Trim$(gFixQuote(.sStationPassword)) & "', wvtSendUpdatesOnly = '" & .sSendUpdatesOnly & "', wvtIsOverridable = '" & .sIsOverridable & "',wvtAllowAutoPost = '" & .sAllowAutoPosting & "', wvtVendorUpdateVAT = '" & .sAutoVendor & "'"
            slRet = slRet & " WHERE wvtVendorID = " & .iIdCode
        Else
            slRet = "Insert into WVT_Vendor_Table (wvtVendorID, wvtName,wvtApprovalPassword,wvtVendorUserName, wvtVendorPassword, wvtExportMethod, "
            slRet = slRet & "wvtAddress, wvtStationUserName,wvtStationPassword, wvtSendUpdatesOnly, wvtImportMethod, wvtHierarchy, wvtDeliveryType, wvtIsOverridable,wvtAllowAutoPost,wvtVendorUpdateVAT) "
            slRet = slRet & " VALUES ( " & .iIdCode & ", '" & Trim$(gFixQuote(.sName)) & "','" & Trim$(gFixQuote(.sApprovalPassword)) & "','" & Trim$(gFixQuote(.sVendorUserName)) & "','" & Trim$(gFixQuote(.sVendorPassword)) & "', "
            slRet = slRet & .iExportMethod & ",'" & Trim$(gFixQuote(.sAddress)) & "','" & Trim$(gFixQuote(.sStationUserName)) & "','" & Trim$(gFixQuote(.sStationPassword)) & "','" & .sSendUpdatesOnly & "'," & .iImportMethod & " , " & .iHierarchy & ", '" & .sDeliveryType & "','" & .sIsOverridable & "','" & .sAllowAutoPosting & "','" & .sAutoVendor & "')"
        End If
    End With
    mInsertOrUpdateSql = slRet
    Exit Function
errbox:
    mInsertOrUpdateSql = ""
End Function
Private Function mValidate(slIssue As String) As Boolean
    Dim blRet As Boolean
    
    slIssue = ""
    blRet = True
    'vendor code can't be blank, or not a number, or 0
    With tmMyVendor
        If .iIdCode = 0 Then
            slIssue = "Internal issue: Vendor code cannot be 0"
            blRet = False
            GoTo Cleanup
        End If
        If Len(.sName) = 0 Then
            slIssue = "Internal issue: Vendor name cannot be blank"
            blRet = False
            GoTo Cleanup
        End If
    End With
Cleanup:
    mValidate = blRet
End Function
Private Function mApprovalPasswordForWebVendor(ilVendorCode As Integer) As String
    Dim slReal As String
    Dim c As Integer

    slReal = ""
    For c = 0 To UBound(tmMyAvailableVendors) - 1
        If ilVendorCode = tmMyAvailableVendors(c).iIdCode Then
            slReal = tmMyAvailableVendors(c).sApprovalPassword
            Exit For
        End If
    Next c
    mApprovalPasswordForWebVendor = slReal
End Function
Private Sub mClearControls()
    imInChg = True
    lbcCode.Caption = ""
    cboDelivery.Text = "None"
    gSetByItemData cboExport, 0
    gSetByItemData cboImport, 0
    cboHierarchy.Text = "None"
    txtAddress.Text = ""
    txtVendorPassword.Text = ""
    txtVendorUsername.Text = ""
    ckcOverride.Value = vbUnchecked
    '8457
    ckcAutoPosting.Value = vbUnchecked
'    If bmIsEditable Then
'        ckcOverride.Enabled = True
'    Else
'        ckcOverride.Enabled = False
'    End If
    rbcStatus(DORMANT).Value = True
    mCopyFromComboToLabels
    imInChg = False
End Sub
Private Sub mCopyFromComboToLabels()
    lbcDelivery.Caption = cboDelivery.Text
    lbcHierarchy.Caption = cboHierarchy.Text
    lbcVendorToWeb.Caption = cboImport.Text
    lbcWebToVendor.Caption = cboExport.Text
End Sub
Private Function mIsChange() As Boolean
    'TRUE is either change, or new!
    Dim blRet As Boolean
    
On Error GoTo ErrHand
    blRet = False
    If bmIsEditable Then
        mSetType
        'try to save 'blank' when first entering?
        If Not myRs Is Nothing Then
            'try to save blank after having chosen a vendor?
            If (myRs.State And adStateOpen) <> 0 Then
                ' a vendor not in table
                If Not myRs.EOF And Not myRs.BOF Then
                    With tmMyVendor
                        If Trim$(myRs!wvtVendorUserName) <> Trim$(.sVendorUserName) Then
                            blRet = True
                            GoTo Cleanup
                        End If
                        If Trim$(myRs!wvtVendorPassword) <> Trim$(.sVendorPassword) Then
                            blRet = True
                            GoTo Cleanup
                        End If
                        If Trim$(myRs!wvtAddress) <> Trim$(.sAddress) Then
                            blRet = True
                            GoTo Cleanup
                        End If
                        If Trim$(myRs!wvtStationUserName) <> Trim$(.sStationUserName) Then
                            blRet = True
                            GoTo Cleanup
                        End If
                        If Trim$(myRs!wvtStationPassword) <> Trim$(.sStationPassword) Then
                            blRet = True
                            GoTo Cleanup
                        End If
                        If Trim$(myRs!wvtIsOverridable) <> Trim$(.sIsOverridable) Then
                            blRet = True
                            GoTo Cleanup
                        End If
                        If Trim$(myRs!wvtSendUpdatesOnly) <> Trim$(.sSendUpdatesOnly) Then
                            blRet = True
                            GoTo Cleanup
                        End If
                        If Not .bIsActive Then
                            blRet = True
                            GoTo Cleanup
                        End If
                        '8457 can be blank, N or Y.  If one is Y and other not, then we have a difference
                        If myRs!wvtAllowAutoPost <> .sAllowAutoPosting And (.sAllowAutoPosting = "Y" Or myRs!wvtAllowAutoPost = "Y") Then
                            blRet = True
                            GoTo Cleanup
                        End If
                        '8862
                        If myRs!wvtVendorUpdateVAT <> .sAutoVendor And (.sAutoVendor = "Y" Or myRs!wvtVendorUpdateVAT = "Y") Then
                            blRet = True
                            GoTo Cleanup
                        End If
                    End With
                'no recordset, so this came from our list
                Else
                    If tmMyVendor.bIsActive Then
                        blRet = True
                    End If
                End If ' recordset not empty?
            End If 'not open
        End If ' recordset not nothing?
    End If  'is editable?
Cleanup:
    mIsChange = blRet
Exit Function

ErrHand:
    blRet = False
    gHandleError "", FORMNAME & "-mIsChange"
    GoTo Cleanup
End Function

Private Sub mBindControls()
    Dim llExport As Long
    Dim llImport As Long
    Dim slApproval As String
    Dim slDelivery As String
    Dim slAddress As String
    Dim slVendorUserName As String
    Dim slVendorPassword As String
    Dim slOverride As String
    Dim llHierarchy As Long
    Dim ilVendorCode As Integer
    Dim llDelivery As Long
    Dim slStationPassword As String
    Dim slStationUserName As String
    Dim slUpdate As String
    
    Dim blVendor As Boolean
    Dim blStation As Boolean
    Dim c As Integer
    '8457
    Dim slAutoPost As String
    '8862
    Dim blAutoVendShow As Boolean
    Dim slAutoVend As String
On Error GoTo ErrHand
    
    imInChg = True
    blVendor = False
    blStation = False
    With cboFormat
        '8949 problem!
        If .ListIndex > -1 Then
            ilVendorCode = .ItemData(.ListIndex)
        ElseIf IsNumeric(lbcCode.Caption) Then
            ilVendorCode = lbcCode.Caption
        End If
    End With
    SQLQuery = "SELECT * FROM WVT_Vendor_Table WHERE wvtVendorID = " & ilVendorCode
    Set myRs = gSQLSelectCall(SQLQuery)
    If myRs.EOF Then
        rbcStatus(DORMANT).Value = True
        For c = 0 To UBound(tmMyAvailableVendors) - 1
            If ilVendorCode = tmMyAvailableVendors(c).iIdCode Then
                With tmMyAvailableVendors(c)
                    llExport = .iExportMethod
                    llImport = .iImportMethod
                    slApproval = .sApprovalPassword
                    slDelivery = .sDeliveryType
                    llHierarchy = .iHierarchy
                End With
                Exit For
            End If
        Next c
    Else
        rbcStatus(ACTIVE).Value = True
        ilVendorCode = myRs!wvtVendorId
        slApproval = myRs!wvtApprovalPassword
        slVendorUserName = myRs!wvtVendorUserName
        slVendorPassword = myRs!wvtVendorPassword
        slAddress = myRs!wvtAddress
        slStationUserName = myRs!wvtStationUserName
        slStationPassword = myRs!wvtStationPassword
        slDelivery = myRs!wvtDeliveryType
        slOverride = myRs!wvtIsOverridable
        slUpdate = myRs!wvtSendUpdatesOnly
        llHierarchy = myRs!wvthierarchy
        llExport = myRs!wvtExportMethod
        llImport = myRs!wvtImportMethod
        slAutoPost = myRs!wvtAllowAutoPost
        '8862
        slAutoVend = myRs!wvtVendorUpdateVAT
    End If
    '8862  show auto vend?
    For c = 0 To UBound(tmMyAvailableVendors) - 1
        If ilVendorCode = tmMyAvailableVendors(c).iIdCode Then
            blAutoVendShow = tmMyAvailableVendors(c).bAutoVendingCSIAllowed
            Exit For
        End If
    Next c
    'now bind
     lbcCode.Caption = ilVendorCode
     '9204
     If slDelivery = "I" Then
        lbcType.Caption = "Informational"
     Else
        lbcType.Caption = "Delivery"
     End If
     If slDelivery = "A" Then
         llDelivery = DeliveryType.Audio
     ElseIf slDelivery = "L" Then
         llDelivery = DeliveryType.Logs
     Else
        llDelivery = DeliveryType.None
     End If
     gSetByItemData cboDelivery, llDelivery
     gSetByItemData cboExport, llExport
     gSetByItemData cboImport, llImport
     gSetByItemData cboHierarchy, llHierarchy
     txtAddress.Text = Trim$(slAddress)
     txtVendorPassword.Text = Trim$(slVendorPassword)
     txtVendorUsername.Text = Trim$(slVendorUserName)
     txtStationPassword.Text = Trim$(slStationPassword)
     txtStationUserName.Text = Trim$(slStationUserName)
     If slOverride = "Y" Then
         ckcOverride.Value = vbChecked
     Else
         ckcOverride.Value = vbUnchecked
     End If
     '8457
     If slAutoPost = "Y" Then
        ckcAutoPosting.Value = vbChecked
     Else
        ckcAutoPosting.Value = vbUnchecked
     End If
     '8862
    ckcAutoVending.Value = vbUnchecked
    ckcAutoVending.Enabled = False
    If blAutoVendShow Then
        '9009, fix 8950
        If bmIsEditable Then
            ckcAutoVending.Enabled = True
        End If
        If slAutoVend = "Y" Then
            ckcAutoVending.Value = vbChecked
        End If
    End If
    With cboImport
        If .ListIndex > -1 Then
            If .ItemData(.ListIndex) = MethodType.csiStationService Or .ItemData(.ListIndex) = MethodType.BothServices Then
                blStation = True
            End If
            If .ItemData(.ListIndex) = MethodType.csiVendorService Or .ItemData(.ListIndex) = MethodType.BothServices Then
                blVendor = True
            End If
        End If
    End With
    With cboExport
        If .ListIndex > -1 Then
            If .ItemData(.ListIndex) = MethodType.csiStationService Or .ItemData(.ListIndex) = MethodType.BothServices Then
                blStation = True
            End If
            If .ItemData(.ListIndex) = MethodType.csiVendorService Or .ItemData(.ListIndex) = MethodType.BothServices Then
                blVendor = True
            End If
        End If
    End With
    If bmIsEditable Then
        ckcOverride.Enabled = True
    Else
        ckcOverride.Enabled = False
    End If
    mEnableConnections True, blVendor
    mEnableConnections False, blStation
    '8457 9010 llImport to stop boxes from being enabled for AnalyticOwl(import = none).
    If (blVendor Or blStation) And llImport > 0 Then
        mEnableChoicesForBothWebVendors True
    Else
        mEnableChoicesForBothWebVendors False
    End If
    mCopyFromComboToLabels
Cleanup:
    imInChg = False
Exit Sub
ErrHand:
    gHandleError "", FORMNAME & "-mBindControls"
    GoTo Cleanup
End Sub

Private Sub mDeactivate()
    Dim ilCode As Integer
    Dim ilVendor As Vendors
    
    ilCode = 0
On Error GoTo ErrHand
    If Not myRs Is Nothing Then
        If (myRs.State And adStateOpen) <> 0 Then
            If Not myRs.EOF And Not myRs.BOF Then
                ilCode = myRs!wvtVendorId
            End If
        End If
    End If
    SQLQuery = "delete from WVT_Vendor_Table where wvtVendorID = " & ilCode
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", FORMNAME & "-mDeactivate"
        mBindControls
        Exit Sub
    ElseIf ilCode > 0 Then
        ilVendor = ilCode
        gAdjustAllowedExportsImports ilVendor
    End If
Cleanup:
    mBindControls
Exit Sub
ErrHand:
    gHandleError "", FORMNAME & "-mDeactivate"
    GoTo Cleanup
End Sub

Private Function mConvert() As Boolean
    Dim blRet As Boolean
    Dim ilVendor As Integer
    Dim rsConvert As ADODB.Recordset
    Dim slSql As String
    Dim blConvert As Boolean
    
    ilVendor = 0
    blConvert = False
    blRet = True
    '8949 problem!
    If IsNumeric(lbcCode.Caption) Then
        ilVendor = lbcCode.Caption
    End If
'    If cboFormat.ListIndex > -1 Then
'        ilVendor = cboFormat.ItemData(cboFormat.ListIndex)
'    End If
    If ilVendor > 0 Then
        slSql = "select count(*) as amount FROM VAT_Vendor_Agreement WHERE  vatWvtVendorID = " & ilVendor
        Set rsConvert = gSQLSelectCall(slSql)
        If Not rsConvert.EOF Then
            If rsConvert!amount = 0 Then
                blConvert = True
            End If
        End If
        If blConvert Then
            If gConvertToVat(ilVendor) = -1 Then
                blRet = False
            End If
        End If
    Else
        blRet = False
    End If
Cleanup:
    mConvert = blRet
    If Not rsConvert Is Nothing Then
        If (rsConvert.State And adStateOpen) <> 0 Then
            rsConvert.Close
        End If
    End If
    Exit Function
ErrHand:
    gHandleError "", FORMNAME & "-mConvert"
    blRet = False
    GoTo Cleanup
End Function







'no longer needed from here down
'
'Private Sub txtGroupName_GotFocus()
'    gCtrlGotFocus ActiveControl
'End Sub
'
'Private Sub txtGroupName_KeyPress(KeyAscii As Integer)
'    Dim slChar As String
'
'    If (KeyAscii <> KEYBACKSPACE) Then
'        slChar = Chr(KeyAscii)
'        If (Asc(slChar) < Asc("A")) Or (Asc(slChar) > Asc("Z")) Then
'            'If (Asc(slChar) < Asc("a")) Or (Asc(slChar) > Asc("z")) Then
'                If (Asc(slChar) < Asc("0")) Or (Asc(slChar) > Asc("9")) Then
'                    If (Asc(slChar) <> Asc("-")) And (Asc(slChar) <> Asc("_")) Then
'                        KeyAscii = 0    ' Ignore this key.
'                    End If
'                End If
'            'End If
'        End If
'    End If
'End Sub
'
'Private Sub txtName_Change()
'    If StrComp(txtName.Text, smName, vbTextCompare) <> 0 Then
'        If Not imInChg Then
'            imNameChange = True
'        End If
'    Else
'        imNameChange = False
'    End If
'End Sub
'
'Private Sub txtName_GotFocus()
'    gCtrlGotFocus ActiveControl
'End Sub
