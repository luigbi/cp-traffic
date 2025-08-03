VERSION 5.00
Begin VB.Form frmGroupNameVehicle 
   Caption         =   "Group Name- Vehicle"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   Icon            =   "AffGroupNameVehicle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOLAExportID 
      Height          =   315
      Left            =   3075
      MaxLength       =   8
      TabIndex        =   6
      Top             =   2190
      Width           =   2655
   End
   Begin VB.TextBox txtWegenerExportID 
      Height          =   315
      Left            =   3075
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1635
      Width           =   2655
   End
   Begin VB.TextBox txtGroupName 
      Height          =   315
      Left            =   3075
      MaxLength       =   10
      TabIndex        =   8
      Top             =   2790
      Width           =   2655
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   855
      TabIndex        =   9
      Top             =   3570
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2985
      TabIndex        =   10
      Top             =   3570
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5055
      TabIndex        =   11
      Top             =   3570
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   3075
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.ComboBox cboVehicle 
      Height          =   315
      Left            =   3075
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lacOLAExportID 
      Caption         =   "OLA Export ID:"
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   2190
      Width           =   2055
   End
   Begin VB.Label lacWegenerExportID 
      Caption         =   "Wegener Export ID:"
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1635
      Width           =   2415
   End
   Begin VB.Label lacGroupName 
      Caption         =   "Group Name:"
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   2790
      Width           =   2055
   End
   Begin VB.Label lblNAme 
      Caption         =   "Vehicle Name:"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1875
   End
End
Attribute VB_Name = "frmGroupNameVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmGroupNameVehicle
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
Private smWegenerExportID As String
Private smOLAExportID As String

Private imInChg As Integer
Private imBSMode As Integer
Private imSave As Integer
Private imVefCode As Integer
Private imVffCode As Integer
Private imNameChange As Integer
Private imIgnoreChg As Integer
Private vff_rst As ADODB.Recordset
Private vef_rst As ADODB.Recordset


Private Sub cboVehicle_Change()
    
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
    slName = LTrim$(cboVehicle.Text)
    ilLen = Len(slName)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(cboVehicle.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        On Error GoTo ErrHand
        cboVehicle.ListIndex = llRow
        cboVehicle.SelStart = ilLen
        cboVehicle.SelLength = Len(cboVehicle.Text)
        If cboVehicle.ListIndex < 0 Then
            imVefCode = -1
            imVffCode = -1
        ElseIf (cboVehicle.ListIndex = 0) And (sgVehicleCall <> "M") Then
            imVefCode = 0
            imVffCode = 0
        Else
            imVefCode = CInt(cboVehicle.ItemData(cboVehicle.ListIndex))
        End If
        If imVefCode < 0 Then
            mClearControls
        ElseIf (imVefCode = 0) And (sgVehicleCall <> "M") Then
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
    gHandleError "AffErrorLog.txt", "frmGroupNameVehicle-cboVehicle_Change"
    imInChg = False
End Sub

Private Sub cboVehicle_Click()
    If imInChg Then
        Exit Sub
    End If
    cboVehicle_Change
End Sub

Private Sub cboVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cboVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboVehicle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboVehicle.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    igVehicleReturnCode = 0
    igIndex = cboVehicle.ListIndex
    Unload frmGroupNameVehicle
End Sub

Private Sub cmdDone_Click()
    
    Dim ilRet As Integer

    ilRet = mVefChkForChange
    If ilRet Then
        ilRet = gMsgBox("Would You Like To Save the Changes", vbYesNo)
        If ilRet = vbYes Then
            ilRet = mSave()
        End If
        If Not ilRet Then
            Exit Sub
        End If
    End If
    igIndex = cboVehicle.ListIndex
    igVehicleReturnCode = 0
    If igIndex > 0 Then
        igVehicleReturnCode = cboVehicle.ItemData(igIndex)
    End If
    Unload frmGroupNameVehicle
End Sub

Private Sub cmdSave_Click()
    Call mSave
End Sub

Private Sub Form_Activate()
    If sgVehicleCall <> "M" Then
        If cboVehicle.ListIndex >= 0 Then
            txtName.SetFocus
        End If
    End If
End Sub

Private Sub Form_Initialize()
    gSetFonts frmGroupNameVehicle
    gCenterForm frmGroupNameVehicle
End Sub

Private Sub Form_Load()
    
    Form_Initialize
    If (Not gWegenerExport) And (Not gOLAExport) Then
        txtGroupName.Enabled = False
        txtGroupName.Text = ""
    End If
    imSave = False
    Call mFillVefInfo("")

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    vef_rst.Close
    vff_rst.Close
    
    Set frmGroupNameVehicle = Nothing

End Sub

Private Function mFillVefInfo(slFromSaveName As String)
    
    Dim slName As String
    Dim ilRow As Integer
    Dim llLoop As Long
    Dim llVpf As Long
    
    On Error GoTo ErrHand
    
    If imSave Then
        slName = Trim$(slFromSaveName)
    End If
    If cboVehicle.Text = "[New]" And (Not imSave) Then
        cboVehicle.ListIndex = 0
        mClearControls
        Exit Function
    End If
    cboVehicle.Clear
    For llLoop = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
        llVpf = gBinarySearchVpf(CLng(tgVehicleInfo(llLoop).iCode))
        If llVpf <> -1 Then
            If (tgVpfOptions(llVpf).sWegenerExport = "Y") Or (tgVpfOptions(llVpf).sOLAExport = "Y") Then
                cboVehicle.AddItem Trim$(tgVehicleInfo(llLoop).sVehicleName)
                cboVehicle.ItemData(cboVehicle.NewIndex) = tgVehicleInfo(llLoop).iCode
            End If
        End If
    Next llLoop
    If sgVehicleCall <> "M" Then
        cboVehicle.AddItem "[New]", 0
        cboVehicle.ItemData(cboVehicle.NewIndex) = 0
    End If

    If (sgVehicleCall <> "M") Or (imSave) Then
        If (sgVehicleCall = "N") And (Not imSave) Then
            cboVehicle.ListIndex = 0
        Else
            If Not imSave Then
                slName = sgVehicleCall
            End If
            ilRow = SendMessageByString(cboVehicle.hwnd, CB_FINDSTRING, -1, slName)
            If (ilRow > 0) Or ((ilRow = 0) And (sgVehicleCall = "M")) Then
                cboVehicle.ListIndex = ilRow
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
    gHandleError "AffErrorLog.txt", "frmGroupNameVehicle-mFillVefInfo"
    imInChg = False
End Function

Private Function mSave()
    
    Dim ilRet As Integer
    Dim ilRow As Integer
    Dim slName As String
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    mSave = False
    ilRet = mVefChkForChange
    'slName = LTrim$(cboVehicle.Text)
    slName = Trim$(txtName.Text)
    If Trim$(slName) = "" Then
        gMsgBox "Vehicle Name must be Defined.", vbOKOnly
        Exit Function
    End If
    'If cboVehicle.text = "[New]" Or imNameChange Then
    '    SQLQuery = "select VefCode from VFF_Vehicle_Features where VefName = '" & txtName.text & "'"
    '    Set rst = gSQLSelectCall(SQLQuery)
    '    If Not rst.EOF Then
    '        gMsgBox "The Vehicle " & Trim$(txtName.text) & " name already exists"
    '        mSave = True
    '        Exit Function
    '    End If
    'End If
    
    'ilRow = SendMessageByString(cboVehicle.hwnd, CB_FINDSTRING, -1, slName)
    'If ilRow >= 0 Then
    '    imVefCode = cboVehicle.ItemData(ilRow)
    'Else
    '    imVefCode = 0
    'End If
    If Trim$(txtGroupName.Text) <> "" Then
        SQLQuery = "SELECT vffCode, vffVefCode FROM VFF_Vehicle_Features WHERE vffGroupName = '" & Trim$(txtGroupName.Text) & "'"
        Set vff_rst = gSQLSelectCall(SQLQuery)
        If Not vff_rst.EOF Then
            If imVffCode <> vff_rst!vffCode Then
                '4/7/15: Allow duplicated Vehicle Group names
                'gMsgBox "Vehicle Group Name previously defined with another vehicle.", vbOKOnly
                ilRet = MsgBox("Vehicle Group Name previously defined with another vehicle, Continue with Save.", vbYesNo, "Duplicated Names")
                If ilRet = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End If
    '3/29/18: Disallow duplicate Wegeber Export ID as this causes problems when importing.  The wrong vehicle might be matched with the imported spots as the first matched ID is used to find the vehicle.
    If Trim$(txtWegenerExportID.Text) <> "" Then
        SQLQuery = "SELECT vffCode, vffVefCode FROM VFF_Vehicle_Features WHERE vffWegenerExportID = '" & Trim$(txtWegenerExportID.Text) & "'"
        Set vff_rst = cnn.Execute(SQLQuery)
        If Not vff_rst.EOF Then
            If imVffCode <> vff_rst!vffCode Then
                gMsgBox "Wegener Export ID previously defined with another vehicle.", vbOKOnly, "Dulpicated Name"
                Exit Function
            End If
        End If
    End If
    mSave = False
    mTrimAndFixQuotes
    If imVffCode = 0 Then
        'SQLQuery = "Insert into VFF_Vehicle_Features "
        'SQLQuery = SQLQuery & "(vffVefCode, vffGroupName, vffWegenerExportID, vffOLAExportID, vffUstCode, vffXDXMLForm , vffXDISCIPrefix, vffExportAirWave) "
        'SQLQuery = SQLQuery & " VALUES (" & imVefCode & ",'" & txtGroupName.Text & "','" & txtWegenerExportID.Text & "','" & txtOLAExportID.Text & "'," & igUstCode & ",'','',''" & ")"
        SQLQuery = "Insert Into VFF_Vehicle_Features ( "
        SQLQuery = SQLQuery & "vffVefCode, "
        SQLQuery = SQLQuery & "vffGroupName, "
        SQLQuery = SQLQuery & "vffWegenerExportID, "
        SQLQuery = SQLQuery & "vffOLAExportID, "
        SQLQuery = SQLQuery & "vffLiveCompliantAdj, "
        SQLQuery = SQLQuery & "vffUstCode, "
        SQLQuery = SQLQuery & "vffUrfCode, "
        SQLQuery = SQLQuery & "vffXDXMLForm, "
        SQLQuery = SQLQuery & "vffXDISCIPrefix, "
        SQLQuery = SQLQuery & "vffXDSaveCF, "
        SQLQuery = SQLQuery & "vffXDSaveHDD, "
        SQLQuery = SQLQuery & "vffXDSaveNAS, "
        SQLQuery = SQLQuery & "vffXDProgCodeID, "
        SQLQuery = SQLQuery & "vffCwfCode, "
        SQLQuery = SQLQuery & "vffAirWavePrgID, "
        SQLQuery = SQLQuery & "vffExportAirWave, "
        SQLQuery = SQLQuery & "vffExportNYESPN, "
        SQLQuery = SQLQuery & "vffPledgeVsAir, "
        SQLQuery = SQLQuery & "vffFedDeliveryZ1, "
        SQLQuery = SQLQuery & "vffFedDeliveryZ2, "
        SQLQuery = SQLQuery & "vffFedDeliveryZ3, "
        SQLQuery = SQLQuery & "vffFedDeliveryZ4, "
        SQLQuery = SQLQuery & "vffFedDeliveryZ5, "
        SQLQuery = SQLQuery & "vffLastAffExptDate, "
        SQLQuery = SQLQuery & "vffMoveSportToNon, "
        SQLQuery = SQLQuery & "vffMoveNonToSport, "
        SQLQuery = SQLQuery & "vffMoveSportToSport, "
        SQLQuery = SQLQuery & "vffMergeTraffic, "
        SQLQuery = SQLQuery & "vffMergeAffiliate, "
        SQLQuery = SQLQuery & "vffMergeWeb, "
        SQLQuery = SQLQuery & "vffPledgeByEvent, "
        SQLQuery = SQLQuery & "vffPledgeHdVtfCode, "
        SQLQuery = SQLQuery & "vffPledgeFtVtfCode, "
        SQLQuery = SQLQuery & "vffPledgeClearance, "
        SQLQuery = SQLQuery & "vffExportEncoESPN, "
        SQLQuery = SQLQuery & "vffWebName, "
        SQLQuery = SQLQuery & "vffSeasonGhfCode, "
        SQLQuery = SQLQuery & "vffMcfCode, "
        SQLQuery = SQLQuery & "vffExportAudio, "
        SQLQuery = SQLQuery & "vffExportMP2, "
        SQLQuery = SQLQuery & "vffExportCncSpot, "
        SQLQuery = SQLQuery & "vffExportEnco, "
        SQLQuery = SQLQuery & "vffExportCnCNetInv, "
        SQLQuery = SQLQuery & "vffIPumpEventTypeOV, "
        SQLQuery = SQLQuery & "vffExportIPump, "
        SQLQuery = SQLQuery & "vffAddr4, "
        SQLQuery = SQLQuery & "vffBBOpenCefCode, "
        SQLQuery = SQLQuery & "vffBBCloseCefCode, "
        SQLQuery = SQLQuery & "vffBBBothCefCode, "
        SQLQuery = SQLQuery & "vffMGsOnWeb, "
        SQLQuery = SQLQuery & "vffReplacementOnWeb, "
        SQLQuery = SQLQuery & "vffExportMatrix, "
        SQLQuery = SQLQuery & "vffSentToXDSStatus, "
        SQLQuery = SQLQuery & "vffStationComp, "
        SQLQuery = SQLQuery & "vffExportSalesForce, "
        SQLQuery = SQLQuery & "vffExportEfficio, "
        SQLQuery = SQLQuery & "vffExportJelli, "
        SQLQuery = SQLQuery & "vffOnXMLInsertion, "
        SQLQuery = SQLQuery & "vffOnInsertions, "
        SQLQuery = SQLQuery & "vffPostLogSource, "
        SQLQuery = SQLQuery & "vffExportTablrau, "
        SQLQuery = SQLQuery & "vffStationPassword, "
        SQLQuery = SQLQuery & "vffHonorZeroUnits, "
        SQLQuery = SQLQuery & "vffHideCommOnLog, "
        SQLQuery = SQLQuery & "vffHideCommOnWeb, "
        SQLQuery = SQLQuery & "vffConflictWinLen, "
        SQLQuery = SQLQuery & "vffACT1LineupCode, "
        SQLQuery = SQLQuery & "vffPrgmmaticAllow, "
        SQLQuery = SQLQuery & "vffSalesBrochure, "
        SQLQuery = SQLQuery & "vffCartOnWeb, "
        SQLQuery = SQLQuery & "vffDefaultAudioType, "
        'SQLQuery = SQLQuery & "vffUnused "
        SQLQuery = SQLQuery & "vffASICallLetters, "
        SQLQuery = SQLQuery & "vffASIBand "
        SQLQuery = SQLQuery & ") "
        SQLQuery = SQLQuery & "Values ( "
        SQLQuery = SQLQuery & imVefCode & ", "
        SQLQuery = SQLQuery & "'" & gFixQuote(txtGroupName.Text) & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(txtWegenerExportID.Text) & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(txtOLAExportID.Text) & "', "
        SQLQuery = SQLQuery & 5 & ","
        SQLQuery = SQLQuery & igUstCode & ", "
        SQLQuery = SQLQuery & 0 & ", "
        SQLQuery = SQLQuery & "'" & "" & "', "
        SQLQuery = SQLQuery & "'" & "" & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote("Y") & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
        SQLQuery = SQLQuery & "'" & "" & "', "
        SQLQuery = SQLQuery & 0 & ", "
        SQLQuery = SQLQuery & "'" & "" & "', "
        SQLQuery = SQLQuery & "'" & "" & "', "
        SQLQuery = SQLQuery & "'" & "" & "', "
        SQLQuery = SQLQuery & "'" & "N" & "', "
        SQLQuery = SQLQuery & "'" & "" & "', "
        SQLQuery = SQLQuery & "'" & "" & "', "
        SQLQuery = SQLQuery & "'" & "" & "', "
        SQLQuery = SQLQuery & "'" & "" & "', "
        SQLQuery = SQLQuery & "'" & "" & "', "
        SQLQuery = SQLQuery & "'" & Format$("1/1/1990", sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote("S") & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote("S") & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote("S") & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
        SQLQuery = SQLQuery & 0 & ", "
        SQLQuery = SQLQuery & 0 & ", "
        SQLQuery = SQLQuery & 0 & ", "
        SQLQuery = SQLQuery & "'" & "N" & "', "
        SQLQuery = SQLQuery & "'" & "" & "' "
        SQLQuery = SQLQuery & 0 & ", "
        SQLQuery = SQLQuery & 0 & ", "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
        SQLQuery = SQLQuery & "'" & "" & "', "          'vffIPumpEventTypeOV
        SQLQuery = SQLQuery & "'" & "N" & "', "         'vffExportIPump
        SQLQuery = SQLQuery & "'" & "" & "', "          'vffAddr4
        SQLQuery = SQLQuery & 0 & ", "                  'vffBBOpenCefCode
        SQLQuery = SQLQuery & 0 & ", "                  'vffBBCloseCefCode
        SQLQuery = SQLQuery & 0 & ", "                  'vffBBBothCefCode
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffMGsOrWeb
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffReplacementOrWeb
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffExportMatrix
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffSentToXDSStatus
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffStationComp
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffExportSalesForce, "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffExportEfficio, "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffExportJelli, "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffOnXMLInsertions, "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffOnInsertions, "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffPostLogSource, "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffExportTablrau "
        SQLQuery = SQLQuery & "'" & gFixQuote("") & "', "  'vffStationPassword "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffHonorZeroUnits "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffHideCommOnLog "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffHideCommOnWeb "
        SQLQuery = SQLQuery & "0" & ", "  'vffConflict Window Length (0=60 seconds) "
        SQLQuery = SQLQuery & "'" & gFixQuote("") & "', "  'vffACT1LineupCode "
        SQLQuery = SQLQuery & "'" & gFixQuote("") & "', "  'vffPrgmmaticAllow "
        SQLQuery = SQLQuery & "'" & gFixQuote("") & "', "  'vffSalesBrochure "
        SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'vffCartOnWeb "
        SQLQuery = SQLQuery & "'" & gFixQuote("R") & "', "  'vffDefaultAudioType "
        'SQLQuery = SQLQuery & "'" & gFixQuote("") & "'"  'vffUnused
        SQLQuery = SQLQuery & "'" & gFixQuote("") & "', "  'vffASICallLetters "
        SQLQuery = SQLQuery & "'" & gFixQuote("") & "'"  'vffASIBand
        SQLQuery = SQLQuery & ") "
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "GroupNameVehicle-mSave"
            imInChg = False
            mSave = False
            Exit Function
        End If
        mClearControls
    Else
        SQLQuery = "UPDATE VFF_Vehicle_Features"
        SQLQuery = SQLQuery & " SET vffGroupName = '" & Trim$(txtGroupName.Text) & "',"
        SQLQuery = SQLQuery & "vffWegenerExportID = '" & Trim$(txtWegenerExportID.Text) & "',"
        SQLQuery = SQLQuery & "vffOLAExportID = '" & Trim$(txtOLAExportID.Text) & "',"
        SQLQuery = SQLQuery & "vffUstCode =" & igUstCode
        SQLQuery = SQLQuery & " WHERE vffCode = " & imVffCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "GroupNameVehicle-mSave"
            imInChg = False
            mSave = False
            Exit Function
        End If
    End If
    imSave = True
    ilRet = mFillVefInfo(slName)
    imSave = False
    imNameChange = False
    mSave = True
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmGroupNameVehicle-mSave"
    imInChg = False
End Function

Private Sub mClearControls()

    imIgnoreChg = True
    txtName.Text = ""
    txtGroupName.Text = ""
    imIgnoreChg = False
End Sub

Private Function mVefChkForChange()
    
    On Error GoTo ErrHand
    
    mVefChkForChange = False

    If StrComp(txtName.Text, smName, vbTextCompare) <> 0 Then
        mVefChkForChange = True
        Exit Function
    End If
    If StrComp(txtGroupName.Text, smGroupName, vbTextCompare) <> 0 Then
        mVefChkForChange = True
        Exit Function
    End If
    If StrComp(txtWegenerExportID.Text, smWegenerExportID, vbTextCompare) <> 0 Then
        mVefChkForChange = True
        Exit Function
    End If
    If StrComp(txtOLAExportID.Text, smOLAExportID, vbTextCompare) <> 0 Then
        mVefChkForChange = True
        Exit Function
    End If
    
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in GroupNameVehicle-mVefChkForChange: "
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
        gMsg = "A general error has occured in GroupNameVehicle-mTrimAndFixQuotes: "
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

    
    On Error GoTo ErrHand
    
    If cboVehicle.Text = "[New]" Then
        Exit Sub
    End If
    
    imIgnoreChg = True
    SQLQuery = "SELECT * "
    SQLQuery = SQLQuery + " FROM VEF_Vehicles"
    SQLQuery = SQLQuery + " WHERE (VefCode = " & imVefCode & ")"
    Set vef_rst = gSQLSelectCall(SQLQuery)
    If vef_rst.EOF Then
        gMsgBox "No matching records were found", vbOKOnly
        mClearControls
    Else
        txtName.Text = Trim$(vef_rst!vefName)
        smName = txtName.Text
        If (gWegenerExport) Or (gOLAExport) Then
            SQLQuery = "SELECT * "
            SQLQuery = SQLQuery + " FROM VFF_Vehicle_Features"
            SQLQuery = SQLQuery + " WHERE (vffVefCode = " & imVefCode & ")"
            Set vff_rst = gSQLSelectCall(SQLQuery)
            If Not vff_rst.EOF Then
                imVffCode = vff_rst!vffCode
                txtGroupName.Text = Trim$(vff_rst!VffGroupName)
                smGroupName = txtGroupName.Text
                txtWegenerExportID.Text = Trim$(vff_rst!VffWegenerExportID)
                smWegenerExportID = txtWegenerExportID.Text
                txtOLAExportID.Text = Trim$(vff_rst!VffOLAExportID)
                smOLAExportID = txtOLAExportID.Text
            Else
                imVffCode = 0
                txtGroupName.Text = ""
                smGroupName = txtGroupName.Text
                txtWegenerExportID.Text = ""
                smWegenerExportID = txtWegenerExportID.Text
                txtOLAExportID.Text = ""
                smOLAExportID = txtOLAExportID.Text
            End If
        Else
            txtGroupName.Text = ""
            smGroupName = txtGroupName.Text
            txtWegenerExportID.Text = ""
            smWegenerExportID = txtWegenerExportID.Text
            txtOLAExportID.Text = ""
            smOLAExportID = txtOLAExportID.Text
        End If
    End If
    imIgnoreChg = False
Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmGroupNameVehicle-mBindControls"
    imInChg = False
End Sub

Private Sub txtOLAExportID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtWegenerExportID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
