VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFastEnd 
   Caption         =   "Fast End"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   Icon            =   "AffFastEnd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboStation 
      Height          =   315
      Left            =   3885
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   525
      Visible         =   0   'False
      Width           =   5385
   End
   Begin VB.PictureBox pbcType 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1080
      ScaleHeight     =   195
      ScaleWidth      =   6540
      TabIndex        =   0
      Top             =   120
      Width           =   6540
      Begin VB.OptionButton rbcType 
         Caption         =   "Affiliate"
         Height          =   240
         Index           =   0
         Left            =   2820
         TabIndex        =   2
         Top             =   0
         Width           =   1170
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Vehicle"
         Height          =   240
         Index           =   1
         Left            =   4455
         TabIndex        =   3
         Top             =   0
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Terminate Agreements by"
         Height          =   225
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2550
      End
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "Browse..."
      Enabled         =   0   'False
      Height          =   300
      Left            =   8160
      TabIndex        =   11
      Top             =   1530
      Width           =   1065
   End
   Begin VB.TextBox txtBrowse 
      Height          =   315
      Left            =   3885
      TabIndex        =   10
      Top             =   1560
      Width           =   3870
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "All"
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   2850
      Width           =   615
   End
   Begin VB.TextBox txtEndDate 
      Height          =   315
      Left            =   3885
      TabIndex        =   8
      Top             =   1035
      Width           =   1575
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9000
      Top             =   5565
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6330
      FormDesignWidth =   9900
   End
   Begin VB.CommandButton cmdMoveLeft 
      Caption         =   "<"
      Height          =   375
      Left            =   4815
      TabIndex        =   16
      Top             =   4305
      Width           =   615
   End
   Begin VB.CommandButton cmdMoveRight 
      Caption         =   ">"
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   3570
      Width           =   615
   End
   Begin VB.ComboBox cboVehicle 
      Height          =   315
      Left            =   3885
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   525
      Width           =   5385
   End
   Begin VB.ListBox lbcStations 
      Height          =   2400
      Index           =   1
      ItemData        =   "AffFastEnd.frx":08CA
      Left            =   5880
      List            =   "AffFastEnd.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   2490
      Width           =   3375
   End
   Begin VB.ListBox lbcStations 
      Height          =   2400
      Index           =   0
      ItemData        =   "AffFastEnd.frx":08CE
      Left            =   1080
      List            =   "AffFastEnd.frx":08D5
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   2490
      Width           =   3375
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   2040
      TabIndex        =   20
      Top             =   5730
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   6240
      TabIndex        =   21
      Top             =   5730
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9435
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblList 
      Caption         =   "List of Affiliates to Terminate"
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   1605
      Width           =   2340
   End
   Begin VB.Label lblWarning 
      Caption         =   "Warning:  All spots occurring after the entered End Date, posted or not, will be permanently removed."
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   5250
      Width           =   7695
   End
   Begin VB.Label lblEndDate 
      Caption         =   "End Date"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1035
      Width           =   1215
   End
   Begin VB.Label lblDontTerminate 
      Caption         =   "Do Not Terminate these Affiliates"
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label lblTerminate 
      Caption         =   "Terminate these Affiliates"
      Height          =   375
      Left            =   5880
      TabIndex        =   12
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label lblVehicles 
      Caption         =   "Terminate Agreements for"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   525
      Width           =   2295
   End
End
Attribute VB_Name = "frmFastEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmFastAdd - Allow users to terminate multiple agreements based
'*               from a given vehicle
'*
'*  Created March, 2004 by Doug Smith
'*
'*  Copyright Counterpoint Software,  Inc. 2004
'******************************************************

Option Explicit
Option Compare Text

Private imAgreeType As Integer
Private imVehInChg As Integer
Private imVehBSMode As Integer
Private imVefCode As Integer
Private lmAttCode As Long
Private imShttCode As Integer
Private smCurDate As String
Private smSvOnAirDate As String
Private smSvOffAirDate As String
Private smSvDropDate As String
Private smStationName As String
Private smVehicleName As String
Private imUnivisionType As Integer
Private imWebType As Integer
Private smCurDir As String
Private smCurDrive As String
Private imShttInChg As Integer
Private imShttBSMode As Integer
Private bmShowDates As Boolean

Private Sub mGenOK()
    
    On Error GoTo ErrHand
    
    If rbcType(0).Value Then
        If cboStation.Text <> "" And txtEndDate.Text <> "" And lbcStations(1).ListCount > 0 Then
            cmdUpdate.Enabled = True
        Else
            cmdUpdate.Enabled = False
        End If
    Else
        If cboVehicle.Text <> "" And txtEndDate.Text <> "" And lbcStations(1).ListCount > 0 Then
            cmdUpdate.Enabled = True
        Else
            cmdUpdate.Enabled = False
        End If
    End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastEnd-mGenOK: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastEndSummary.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub cboStation_Change()
    Dim llRow As Long
    Dim lTemp As Long
    Dim slName As String
    Dim ilLen As Integer
    
    On Error GoTo ErrHand
        
    If imShttInChg Then
        Exit Sub
    End If
    imShttInChg = True
    Screen.MousePointer = vbHourglass
    slName = LTrim$(cboStation.Text)
    ilLen = Len(slName)
    If imShttBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imShttBSMode = False
    End If
    
    llRow = SendMessageByString(cboStation.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        cboStation.ListIndex = llRow
        cboStation.SelStart = ilLen
        cboStation.SelLength = Len(cboStation.Text)
        lTemp = cboStation.ItemData(cboStation.ListIndex)
        cmcBrowse.Enabled = True
    Else
        cmcBrowse.Enabled = False
    End If
    Screen.MousePointer = vbDefault
    imShttInChg = False
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastEnd-cboStation_Change: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastEndSummary.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub cboStation_Click()
    
    On Error GoTo ErrHand
    
    lbcStations(0).Clear
    lbcStations(1).Clear
    If cboStation.Text <> "" Then
        mGetStaMark
    End If
    mGenOK
    If cboStation.ListIndex >= 0 Then
        cmcBrowse.Enabled = True
    Else
        cmcBrowse.Enabled = False
    End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastEnd-cboStation_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastEndSummary.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub cboStation_KeyDown(KeyCode As Integer, Shift As Integer)
    imShttBSMode = False
End Sub

Private Sub cboStation_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then
        If cboStation.SelLength <> 0 Then
            imShttBSMode = True
        End If
    End If
End Sub

Private Sub cboVehicle_Change()

    Dim llRow As Long
    Dim lTemp As Long
    Dim slName As String
    Dim ilLen As Integer
    
    On Error GoTo ErrHand
        
    If imVehInChg Then
        Exit Sub
    End If
    imVehInChg = True
    Screen.MousePointer = vbHourglass
    slName = LTrim$(cboVehicle.Text)
    ilLen = Len(slName)
    If imVehBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imVehBSMode = False
    End If
    
    llRow = SendMessageByString(cboVehicle.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        cboVehicle.ListIndex = llRow
        cboVehicle.SelStart = ilLen
        cboVehicle.SelLength = Len(cboVehicle.Text)
        lTemp = cboVehicle.ItemData(cboVehicle.ListIndex)
        cmcBrowse.Enabled = True
    Else
        cmcBrowse.Enabled = False
    End If
    Screen.MousePointer = vbDefault
    imVehInChg = False
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastEnd-cboVehicle_Change: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastEndSummary.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub

Private Sub cboVehicle_Click()
    
    On Error GoTo ErrHand
    
    lbcStations(0).Clear
    lbcStations(1).Clear
    If cboVehicle.Text <> "" Then
        mGetStaMark
    End If
    mGenOK
    If cboVehicle.ListIndex >= 0 Then
        cmcBrowse.Enabled = True
    Else
        cmcBrowse.Enabled = False
    End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastEnd-cboVehicle_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastEndSummary.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub

Private Sub cboVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
    imVehBSMode = False
End Sub

Private Sub cboVehicle_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then
        If cboVehicle.SelLength <> 0 Then
            imVehBSMode = True
        End If
    End If

End Sub

Private Sub cmcBrowse_Click()
    smCurDrive = Left$(CurDir$, 1)
    smCurDir = CurDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    
    txtBrowse.Text = Trim$(CommonDialog1.fileName)
    ChDrive smCurDir
    ChDir smCurDir
    mExternalFile
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub cmdAll_Click()
    
    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim ilExclCount As Integer
    Dim ilInclCount As Integer
    
    On Error GoTo ErrHand
    For ilLoop = 0 To lbcStations(0).ListCount - 1 Step 1
        lbcStations(1).AddItem lbcStations(0).List(ilLoop)
        lbcStations(1).ItemData(lbcStations(1).NewIndex) = lbcStations(0).ItemData(ilLoop)
    Next ilLoop
    
    lbcStations(0).Clear
    lbcStations(0).ListIndex = -1
    cmdAll.Visible = False
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastAdd-cmdAll_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload frmFastEnd
End Sub

Private Function mInit()

    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    mInit = True
    
    'Load the list of vehicles for the Create area
    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        '11/4/09-  Show Log and Conventional vehicle.  Let client pick which they want agreements to be used for
        'Temporarily include only for Special user until testing is complete
        'If tgVehicleInfo(ilLoop).sVehType = "C" Or tgVehicleInfo(ilLoop).sVehType = "A" Or tgVehicleInfo(ilLoop).sVehType = "G" Or tgVehicleInfo(ilLoop).sVehType = "I" Then
        If tgVehicleInfo(ilLoop).sVehType = "C" Or tgVehicleInfo(ilLoop).sVehType = "A" Or tgVehicleInfo(ilLoop).sVehType = "G" Or tgVehicleInfo(ilLoop).sVehType = "I" Or (tgVehicleInfo(ilLoop).sVehType = "L") Then
            'If (tgVehicleInfo(ilLoop).sOLAExport <> "Y") Then
                cboVehicle.AddItem Trim$(tgVehicleInfo(ilLoop).sVehicle)
                cboVehicle.ItemData(cboVehicle.NewIndex) = tgVehicleInfo(ilLoop).iCode
            'End If
        End If
    Next ilLoop
    mPopStations
    ilRet = gPopAttInfo()
    bmShowDates = False
    SQLQuery = "SELECT siteShowContrDate From Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If rst!siteShowContrDate = "Y" Then
            bmShowDates = True
        End If
    End If
    rst.Close
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastEnd - mInit: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastEndSummary.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    mInit = False
End Function

Private Sub mGetStaMark()
    
    Dim shtt_rst As ADODB.Recordset
    Dim att_rst As ADODB.Recordset
    Dim slDate As String
    Dim slEndDate As String
    Dim slRange As String
    Dim slTemp As String
    Dim llIdx As Long
    Dim llVef As Long
    
    
    On Error GoTo ErrHand
    slDate = Format(gNow(), "yyyy-mm-dd")
    If rbcType(0).Value Then
        imShttCode = CInt(cboStation.ItemData(cboStation.ListIndex))
        SQLQuery = "SELECT attCode, attShfCode, attDropDate, attOffAir, AttOnAir, attVefCode"
        SQLQuery = SQLQuery + " FROM att "
        SQLQuery = SQLQuery + " WHERE (attShfCode = '" & imShttCode & "')"
        SQLQuery = SQLQuery + " AND (attOffAir  >= '" & slDate & "')"
        SQLQuery = SQLQuery + " AND (attDropDate >= '" & slDate & "')"
        Set att_rst = gSQLSelectCall(SQLQuery)
        
        If att_rst.EOF Then
            lbcStations(0).Clear
            lbcStations(0).ForeColor = vbRed
            lbcStations(0).AddItem "No vehicles are pledged for: " & cboStation.Text
            Exit Sub
        End If
        
        lbcStations(0).ForeColor = vbBlack
        cmdAll.Visible = True
        lbcStations(0).Clear
        While Not att_rst.EOF
            llVef = gBinarySearchVef(CLng(att_rst!attvefCode))
            If llVef <> -1 Then
                If DateValue(gAdjYear(att_rst!attDropDate)) < DateValue(gAdjYear(att_rst!attOffAir)) Then
                    slEndDate = Format$(att_rst!attDropDate, sgShowDateForm)
                Else
                    slEndDate = Format$(att_rst!attOffAir, sgShowDateForm)
                End If
                If (DateValue(gAdjYear(att_rst!attOnAir)) = DateValue("1/1/1970")) Then 'Or (att_rst!attOnAir = "1/1/70") Then    'Placeholder value to prevent using Nulls/outer joins
                    slRange = ""
                Else
                    slRange = Format$(Trim$(att_rst!attOnAir), sgShowDateForm)
                End If
                If (DateValue(gAdjYear(slEndDate)) = DateValue("12/31/2069") Or DateValue(gAdjYear(slEndDate)) = DateValue("12/31/69")) Then  'Or (att_rst!attOffAir = "12/31/69") Then
                    If slRange <> "" Then
                        slRange = slRange & "-TFN"
                    End If
                Else
                    If slRange <> "" Then
                        slRange = slRange & "-" & slEndDate    'att_rst!attOffAir
                    Else
                        slRange = "Thru " & slEndDate 'att_rst!attOffAir
                    End If
                End If
                lbcStations(0).AddItem Trim$(tgVehicleInfo(llVef).sVehicle) & " " & slRange
                lbcStations(0).ItemData(lbcStations(0).NewIndex) = att_rst!attCode  'attShfCode
            End If
            att_rst.MoveNext
        Wend
        
    Else
        imVefCode = CInt(cboVehicle.ItemData(cboVehicle.ListIndex))
        
        'SQLQuery = "SELECT attCode, attShfCode, attDropDate, attOffAir, AttOnAir, shttCallLetters, shttMarket"
        'SQLQuery = SQLQuery + " FROM att, shtt"
        
        'SQLQuery = "SELECT attCode, attShfCode, attDropDate, attOffAir, AttOnAir, shttCallLetters, mktName"
        'SQLQuery = SQLQuery + " FROM att, shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode"
        'SQLQuery = SQLQuery + " WHERE (attVefCode = '" & imVefCode & "')"
        'SQLQuery = SQLQuery + " AND shttCode = attShfCode"
        'SQLQuery = SQLQuery + " AND (attOffAir  >= '" & slDate & "')"
        'SQLQuery = SQLQuery + " AND (attDropDate >= '" & slDate & "')"
        
        SQLQuery = "SELECT attCode, attShfCode, attDropDate, attOffAir, AttOnAir, shttCallLetters, shttMktCode"
        SQLQuery = SQLQuery + " FROM att, shtt "
        SQLQuery = SQLQuery + " WHERE (attVefCode = '" & imVefCode & "')"
        SQLQuery = SQLQuery + " AND shttCode = attShfCode"
        SQLQuery = SQLQuery + " AND (attOffAir  >= '" & slDate & "')"
        SQLQuery = SQLQuery + " AND (attDropDate >= '" & slDate & "')"
        Set att_rst = gSQLSelectCall(SQLQuery)
        
        If att_rst.EOF Then
            lbcStations(0).Clear
            lbcStations(0).ForeColor = vbRed
            lbcStations(0).AddItem "No stations are pledged for: " & cboVehicle.Text
            Exit Sub
        End If
        
        lbcStations(0).ForeColor = vbBlack
        If Not att_rst.EOF Then
            cmdAll.Visible = True
            lbcStations(0).Clear
            While Not att_rst.EOF
                If DateValue(gAdjYear(att_rst!attDropDate)) < DateValue(gAdjYear(att_rst!attOffAir)) Then
                    slEndDate = Format$(att_rst!attDropDate, sgShowDateForm)
                Else
                    slEndDate = Format$(att_rst!attOffAir, sgShowDateForm)
                End If
                If (DateValue(gAdjYear(att_rst!attOnAir)) = DateValue("1/1/1970")) Then 'Or (att_rst!attOnAir = "1/1/70") Then    'Placeholder value to prevent using Nulls/outer joins
                    slRange = ""
                Else
                    slRange = Format$(Trim$(att_rst!attOnAir), sgShowDateForm)
                End If
                If (DateValue(gAdjYear(slEndDate)) = DateValue("12/31/2069") Or DateValue(gAdjYear(slEndDate)) = DateValue("12/31/69")) Then  'Or (att_rst!attOffAir = "12/31/69") Then
                    If slRange <> "" Then
                        slRange = slRange & "-TFN"
                    End If
                Else
                    If slRange <> "" Then
                        slRange = slRange & "-" & slEndDate    'att_rst!attOffAir
                    Else
                        slRange = "Thru " & slEndDate 'att_rst!attOffAir
                    End If
                End If
                ''SQLQuery = "SELECT shttCallLetters, shttMarket"
                ''SQLQuery = SQLQuery + " FROM shtt"
                ''SQLQuery = SQLQuery + " WHERE (shttCode = " & att_rst!attShfCode & ")"
                ''Set shtt_rst = gSQLSelectCall(SQLQuery)
                'lbcStations(0).AddItem Trim$(att_rst!shttCallLetters) & " , " & Trim$(att_rst!shttMarket) & " " & slRange
                
                For llIdx = 0 To UBound(tgMarketInfo) - 1
                    If tgMarketInfo(llIdx).lCode = att_rst!shttMktCode Then
                        slTemp = tgMarketInfo(llIdx).sName
                        Exit For
                    End If
                Next llIdx
                
                
                'If IsNull(att_rst!mktName) = True Then
                '    lbcStations(0).AddItem Trim$(att_rst!shttCallLetters) & " " & slRange
                'Else
                '    lbcStations(0).AddItem Trim$(att_rst!shttCallLetters) & " , " & Trim$(att_rst!mktName) & " " & slRange
                'End If
                'lbcStations(0).ItemData(lbcStations(0).NewIndex) = att_rst!attCode  'attShfCode
                
                If IsNull(slTemp) = True Then
                    lbcStations(0).AddItem Trim$(att_rst!shttCallLetters) & " " & slRange
                Else
                    lbcStations(0).AddItem Trim$(att_rst!shttCallLetters) & " , " & Trim$(slTemp) & " " & slRange
                End If
                lbcStations(0).ItemData(lbcStations(0).NewIndex) = att_rst!attCode  'attShfCode
                
                att_rst.MoveNext
            Wend
        End If
    End If
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFastAdd-mGetStaMark"
End Sub


Private Sub cmdMoveLeft_Click()
    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim ilExclCount As Integer
    Dim ilInclCount As Integer
    
    On Error GoTo ErrHand
    For ilLoop = 0 To lbcStations(1).ListCount - 1 Step 1
        If lbcStations(1).Selected(ilLoop) Then
            lbcStations(0).AddItem lbcStations(1).List(ilLoop)
            lbcStations(0).ItemData(lbcStations(0).NewIndex) = lbcStations(1).ItemData(ilLoop)
        End If
    Next ilLoop
    
    ilInclCount = lbcStations(0).ListCount
    ilExclCount = lbcStations(1).ListCount
    For ilLoop = 0 To ilInclCount - 1 Step 1
        For ilIdx = 0 To ilExclCount - 1 Step 1
            If lbcStations(0).List(ilLoop) = lbcStations(1).List(ilIdx) Then
                lbcStations(1).RemoveItem (ilIdx)
                ilExclCount = ilExclCount - 1
                mGenOK
                Exit For
            End If
        Next ilIdx
    Next ilLoop
    
    lbcStations(1).ListIndex = -1
    lbcStations(0).ListIndex = -1
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastEnd-cmdMoveLeft_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastEndSummary.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If


End Sub

Private Sub cmdMoveRight_Click()

    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim ilExclCount As Integer
    Dim ilInclCount As Integer
    
    On Error GoTo ErrHand
    
    For ilLoop = 0 To lbcStations(0).ListCount - 1 Step 1
        If lbcStations(0).Selected(ilLoop) Then
            lbcStations(1).AddItem lbcStations(0).List(ilLoop)
            lbcStations(1).ItemData(lbcStations(1).NewIndex) = lbcStations(0).ItemData(ilLoop)
        End If
    Next ilLoop
    
    ilInclCount = lbcStations(1).ListCount
    ilExclCount = lbcStations(0).ListCount
    For ilLoop = 0 To ilInclCount - 1 Step 1
        For ilIdx = 0 To ilExclCount - 1 Step 1
            If lbcStations(1).List(ilLoop) = lbcStations(0).List(ilIdx) Then
                lbcStations(0).RemoveItem (ilIdx)
                ilExclCount = ilExclCount - 1
                mGenOK
                Exit For
            End If
        Next ilIdx
    Next ilLoop
    
    lbcStations(1).ListIndex = -1
    lbcStations(0).ListIndex = -1
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastEnd-cmdMoveRight_Click: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastEndSummary.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If


End Sub

Private Sub cmdUpdate_Click()
    mTerminateAgreements
    cmdCancel.Caption = "Done"
End Sub

Private Sub Form_Load()
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    cmdUpdate.Enabled = False
    imVehBSMode = False
    imVehInChg = False

    cmdCancel.Caption = "Cancel"
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 1.4
    Me.Left = (Screen.Width - Me.Width) / 2
    smCurDate = Format(gNow(), sgShowDateForm)
    frmFastEnd.Caption = "Affiliate Fast End - " & sgClientName
    gLogMsg "", "FastEndSummary.Txt", False
    gLogMsg "   *** Starting Fast End Program   ***", "FastEndSummary.Txt", False
    gLogMsg "", "FastEndSummary.Txt", False
    ilRet = mInit
    Screen.MousePointer = vbDefault
    If Not ilRet Then
        Exit Sub
    End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastEnd-Form_Load: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastEndSummary.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    gLogMsg "", "FastEndSummary.Txt", False
    gLogMsg "   *** Ending Fast End Program   ***", "FastEndSummary.Txt", False
    Unload frmFastEnd

End Sub

Private Sub lbcStations_DblClick(Index As Integer)

    On Error GoTo ErrHand
    
    If InStr(1, lbcStations(0).List(0), "No Stations", vbTextCompare) Then
        Exit Sub
    End If
    
    If lbcStations(0).ListIndex >= 0 Then
        lbcStations(1).AddItem lbcStations(0).Text
        lbcStations(1).ItemData(lbcStations(1).NewIndex) = lbcStations(0).ItemData(lbcStations(0).ListIndex)
        lbcStations(0).RemoveItem (lbcStations(0).ListIndex)
        lbcStations(1).ListIndex = -1
        lbcStations(0).ListIndex = -1
    End If

    If lbcStations(1).ListIndex >= 0 Then
        lbcStations(0).AddItem lbcStations(1).Text
        lbcStations(0).ItemData(lbcStations(0).NewIndex) = lbcStations(1).ItemData(lbcStations(1).ListIndex)
        lbcStations(1).RemoveItem (lbcStations(1).ListIndex)
        lbcStations(0).ListIndex = -1
        lbcStations(1).ListIndex = -1
    End If
    mGenOK
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastEnd-lbcStations_DblClick: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastEndSummary.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub

Private Sub mTerminateAgreements()

    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVefCode As Integer
    Dim ilRow As Integer
    Dim slOnAir As String
    Dim slOffAir As String
    Dim slDropDate As String
    Dim ilWebType As Integer
    Dim ilUnivisionType As Integer
    Dim att_updt As ADODB.Recordset
    Dim att_rst As ADODB.Recordset
    Dim att2_rst As ADODB.Recordset
    Dim shtt_rst As ADODB.Recordset
    Dim llGroupID As Long
    Dim llLastDate1 As Long
    Dim slMulticast2 As String
    Dim llLastDate2 As Long
    Dim slMulticastMsg As String
    Dim llVef As Long
    Dim ilShtt As Integer
    Dim llAtt As Long
    Dim llAttCode2 As Long
    
    On Error GoTo ErrHand
    'Verify the date is valid
    '11/15/14: Remove backslash if last character
    If right(txtEndDate.Text, 1) = "/" Then
        txtEndDate.Text = Left(txtEndDate.Text, Len(txtEndDate.Text) - 1)
    End If
    If Not gIsDate(txtEndDate.Text) Then
        gMsgBox "The End Date is not a valid date.  Please enter a new date."
        txtEndDate.Text = ""
        txtEndDate.SetFocus
        mGenOK
        Exit Sub
    End If
    
    'Force a Sunday day of the week.
    If Weekday(txtEndDate.Text) <> vbSunday Then
        gMsgBox "Date Must be a Sunday.  Please enter a Sunday date.", vbOKOnly
        Exit Sub
    End If
    
    'ilRow = SendMessageByString(cboVehicle.hwnd, CB_FINDSTRING, -1, cboVehicle.Text)
    'imVefCode = CInt(cboVehicle.ItemData(ilRow))
    slMulticastMsg = ""
    If rbcType(0).Value Then
        ilRow = SendMessageByString(cboStation.hwnd, CB_FINDSTRING, -1, cboStation.Text)
        imShttCode = CInt(cboStation.ItemData(ilRow))   'CInt(cboVehicle.ItemData(ilRow))
        smStationName = Trim$(cboStation.Text)
    Else
        ilRow = SendMessageByString(cboVehicle.hwnd, CB_FINDSTRING, -1, cboVehicle.Text)
        imVefCode = CInt(cboVehicle.ItemData(ilRow))
        smVehicleName = Trim$(cboVehicle.Text)
        
        '6/14/14: Verify that all multicast station are matched-up in the Include list.
        gAlignMulticastStations imVefCode, "A", lbcStations(1), lbcStations(0)
        
    End If

    Screen.MousePointer = vbHourglass
    For ilLoop = 0 To lbcStations(1).ListCount - 1 Step 1
        lmAttCode = lbcStations(1).ItemData(ilLoop)
        'smStationName = lbcStations(1).List(ilLoop)
        If rbcType(0).Value Then
            smVehicleName = Trim$(lbcStations(1).List(ilLoop))
        Else
            smStationName = Trim$(lbcStations(1).List(ilLoop))
        End If
        
        SQLQuery = "Select * from att"
        'SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode & ""
        'SQLQuery = SQLQuery + " AND attCode = " & lmAttCode & ")"
        SQLQuery = SQLQuery + " WHERE (attCode = " & lmAttCode & ")"
        Set att_rst = gSQLSelectCall(SQLQuery)
        
        ''gLogMsg "Terminating " & smStationName & " Agreement running on " & Trim$(cboVehicle.Text) & " as of " & Trim$(txtEndDate.Text), "FastEndSummary.Txt", False
        'gLogMsg "Terminating " & smStationName & " Agreement running on " & smVehicleName & " as of " & Trim$(txtEndDate.Text), "FastEndSummary.Txt", False
        
        '6/14/14
        imShttCode = att_rst!attshfCode
        imVefCode = att_rst!attvefCode
        ReDim llAttCode(0 To 1) As Long
        llAttCode(0) = lmAttCode
        If rbcType(0).Value Then
            If gIsMulticast(imShttCode) Then
                If att_rst!attMulticast = "Y" Then
                    llGroupID = gGetStaMulticastGroupID(imShttCode)
                    llLastDate1 = 0
                    If gDateValue(att_rst!attOffAir) <= gDateValue(att_rst!attDropDate) Then
                        If gDateValue(att_rst!attOffAir) > llLastDate1 Then
                            llLastDate1 = gDateValue(att_rst!attOffAir)
                        End If
                    Else
                        If gDateValue(att_rst!attDropDate) > llLastDate1 Then
                            llLastDate1 = gDateValue(att_rst!attDropDate)
                        End If
                    End If
                    SQLQuery = "Select shttCode FROM shtt where shttMultiCastGroupID = " & llGroupID & " And shttCode <> " & imShttCode
                    Set shtt_rst = gSQLSelectCall(SQLQuery)
                    While Not shtt_rst.EOF
                        llLastDate2 = 0
                        slMulticast2 = ""
                        llAttCode2 = 0
                        SQLQuery = "SELECT attCode, attOnAir, attOffAir, attDropDate, attMulticast FROM att"
                        SQLQuery = SQLQuery + " WHERE (attShfCode = " & shtt_rst!shttCode & " AND attVefCode = " & att_rst!attvefCode & ")"
                        Set att2_rst = gSQLSelectCall(SQLQuery)
                        While Not att2_rst.EOF
                            If gDateValue(att2_rst!attOffAir) <= gDateValue(att2_rst!attDropDate) Then
                                If gDateValue(att2_rst!attOffAir) > llLastDate2 Then
                                    llLastDate2 = gDateValue(att2_rst!attOffAir)
                                    slMulticast2 = att2_rst!attMulticast
                                    llAttCode2 = att2_rst!attCode
                                End If
                            Else
                                If gDateValue(att2_rst!attDropDate) > llLastDate2 Then
                                    llLastDate2 = gDateValue(att2_rst!attDropDate)
                                    slMulticast2 = att2_rst!attMulticast
                                    llAttCode2 = att2_rst!attCode
                                End If
                            End If
                            att2_rst.MoveNext
                        Wend
                        If (slMulticast2 = "Y") And (llLastDate1 = llLastDate2) Then
                            llAttCode(UBound(llAttCode)) = llAttCode2
                            ReDim Preserve llAttCode(0 To UBound(llAttCode) + 1) As Long
                            llVef = gBinarySearchVef(CLng(att_rst!attvefCode))
                            If llVef <> -1 Then
                                If slMulticastMsg = "" Then
                                    slMulticastMsg = Trim$(tgVehicleInfo(llVef).sVehicle)
                                Else
                                    slMulticastMsg = slMulticastMsg & "; " & Trim$(tgVehicleInfo(llVef).sVehicle)
                                End If
                                ilShtt = gBinarySearchStationInfoByCode(shtt_rst!shttCode)
                                If ilShtt <> -1 Then
                                    slMulticastMsg = slMulticastMsg & " " & Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
                                End If
                            Else
                                ilShtt = gBinarySearchStationInfoByCode(shtt_rst!shttCode)
                                If ilShtt <> -1 Then
                                    If slMulticastMsg = "" Then
                                        slMulticastMsg = Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
                                    Else
                                        slMulticastMsg = slMulticastMsg & "; " & Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
                                    End If
                                End If
                            End If
                        End If
                        shtt_rst.MoveNext
                    Wend
                End If
            End If
        End If
        
        If (rbcType(0).Value) And (slMulticastMsg <> "") Then
            gLogMsg "The following additional Multicast Agreements will be terminated: " & slMulticastMsg, "FastEndSummary.Txt", False
        End If
        
        For llAtt = 0 To UBound(llAttCode) - 1 Step 1
        
            
            lmAttCode = llAttCode(llAtt)
            
            SQLQuery = "Select * from att"
            SQLQuery = SQLQuery + " WHERE (attCode = " & lmAttCode & ")"
            Set att_rst = gSQLSelectCall(SQLQuery)
            
            imVefCode = att_rst!attvefCode
            imShttCode = att_rst!attshfCode
            
            llVef = gBinarySearchVef(CLng(imVefCode))
            If llVef <> -1 Then
                smVehicleName = Trim$(tgVehicleInfo(llVef).sVehicle)
            Else
                smVehicleName = "Vehicle Missing: " & imVefCode
            End If
            ilShtt = gBinarySearchStationInfoByCode(imShttCode)
            If ilShtt <> -1 Then
                smStationName = Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
            Else
                smStationName = "Station Missing: " & imShttCode
            End If
            'gLogMsg "Terminating " & smStationName & " Agreement running on " & Trim$(cboVehicle.Text) & " as of " & Trim$(txtEndDate.Text), "FastEndSummary.Txt", False
            gLogMsg "Terminating " & smStationName & " Agreement running on " & smVehicleName & " as of " & Trim$(txtEndDate.Text), "FastEndSummary.Txt", False
            
            imWebType = False
            imUnivisionType = False
            If att_rst!attExportType = 1 Then
                imWebType = True
            End If
            If att_rst!attExportType = 2 Then
                imUnivisionType = True
            End If
            
            
            'D.S. 10/25/04
            imAgreeType = att_rst!attExportType
            slOffAir = att_rst!attOffAir
            smSvOffAirDate = att_rst!attOffAir
            slOnAir = att_rst!attOnAir
            smSvOnAirDate = att_rst!attOnAir
            slDropDate = Trim$(txtEndDate.Text)
            smSvDropDate = Format$(att_rst!attDropDate, "m/d/yyyy")
            
            lmAttCode = att_rst!attCode
                        
            ilRet = mDeleteCPTT(False, slOnAir, slOffAir, slDropDate)
            'If we can't delete the cptt records don't update the att
            If ilRet Then
                'Is this a cancel before start delete?
                If DateValue(gAdjYear(slDropDate)) > DateValue(gAdjYear(slOnAir)) Then
                    SQLQuery = "UPDATE att SET "
                    If bmShowDates Then
                        SQLQuery = SQLQuery & "attDropDate = '" & Format$(txtEndDate.Text, sgSQLDateForm) & "',"
                    Else
                        SQLQuery = SQLQuery & "attOffAir = '" & Format$(txtEndDate.Text, sgSQLDateForm) & "',"
                    End If
                    SQLQuery = SQLQuery & "attEnterDate = '" & Format(gNow(), sgSQLDateForm) & "',"
                    SQLQuery = SQLQuery & "attEnterTime = '" & Format(gNow(), sgSQLTimeForm) & "',"
                    SQLQuery = SQLQuery & "attSentToXDSStatus = '" & "M" & "'"
                    SQLQuery = SQLQuery + " WHERE attCode = " & lmAttCode
                    SQLQuery = SQLQuery & " AND attOffAir > '" & Format$(txtEndDate.Text, sgSQLDateForm) & "'"
                    SQLQuery = SQLQuery & " AND attDropDate > '" & Format$(txtEndDate.Text, sgSQLDateForm) & "'"
                    Set att_updt = gSQLSelectCall(SQLQuery)
                Else
                    ' JD 12-18-2006 Added new function to properly remove an agreement.
                        'Yes, cancel before start - drop date is less than on air date
                    If Not gDeleteAgreement(lmAttCode, "FastEndSummary.Txt") Then
                        gLogMsg "FAIL: mTerminateAgreements - Unable to delete att code " & lmAttCode, "FastEndSummary.Txt", False
                    End If
                End If
            End If
        Next llAtt
    Next ilLoop
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFastEnd-mTerminateAgreements"
End Sub
Private Function mDeleteCPTT(iNewRec As Integer, sOnAir As String, sOffAir As String, sDropDate As String) As Integer
    
    'This module is for deleting CPTT records
    Dim sLLD As String 'Last Log date
    Dim lSvCPTTStart As Long
    Dim lSvCPTTEnd As Long
    Dim lCPTTStart As Long
    Dim lCPTTEnd As Long
    Dim lSDate As Long
    Dim lEDate As Long
    Dim slSDate As String
    Dim slEDate As String
    Dim lDate As Long
    Dim iCycle As Integer
    Dim sTime As String
    Dim iWkDay As Integer
    Dim sMsg As String
    Dim ilExported As Integer
    Dim rst_TestWk As ADODB.Recordset
    Dim ilRet As Integer
    Dim llTtlRows As Long
    Dim slCallLetters As String
    Dim slVehicleName As String
    
    On Error GoTo ErrHand
    
    sTime = Format("12:00AM", "hh:mm:ss")
        
    If DateValue(gAdjYear(sOnAir)) = DateValue("1/1/1970") Then
        mDeleteCPTT = True
        Exit Function
    End If
    
    If (DateValue(gAdjYear(sOnAir)) = DateValue(gAdjYear(smSvOnAirDate))) And (DateValue(gAdjYear(sOffAir)) = DateValue(gAdjYear(smSvOffAirDate))) And (DateValue(gAdjYear(sDropDate)) = DateValue(gAdjYear(smSvDropDate))) Then  'Append
        mDeleteCPTT = True
        Exit Function
    End If
    'Get the last log date
    SQLQuery = "SELECT vpfLLD, vpfLNoDaysCycle"
    SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
    SQLQuery = SQLQuery + " WHERE (vpfvefKCode =" & imVefCode & ")"
    
    Set rst = gSQLSelectCall(SQLQuery)
    If rst.EOF Then
        mDeleteCPTT = True
        Exit Function
    End If
    If IsNull(rst!vpfLLD) Then
        mDeleteCPTT = True
        Exit Function
    End If
    If Not gIsDate(rst!vpfLLD) Then
        'sLLD = "1/1/1970"
        'iWkDay = vbMonday
        mDeleteCPTT = True
        Exit Function
    Else
        'set sLLD to last log date
        sLLD = Format$(rst!vpfLLD, sgShowDateForm)
        '1= Sun, 2= Mon, 3= Tues, 4= Wed, 5= Th, 6= Fri, 7= Sat
        iWkDay = Weekday(gAdjYear(Format$(DateValue(sLLD) + 1, "m/d/yyyy")))
    End If
    iCycle = rst!vpfLNoDaysCycle
    lSvCPTTStart = 0
    lSvCPTTEnd = 0
    lCPTTStart = 0
    lCPTTEnd = 0
    
    If DateValue(gAdjYear(smSvOnAirDate)) <= DateValue(gAdjYear(sLLD)) Then
        If DateValue(gAdjYear(smSvDropDate)) <= DateValue(gAdjYear(sLLD)) Then
            lSvCPTTStart = DateValue(gAdjYear(smSvOnAirDate))
            If DateValue(gAdjYear(smSvDropDate)) < DateValue(gAdjYear(smSvOffAirDate)) Then
                lSvCPTTEnd = DateValue(gAdjYear(smSvDropDate)) '- iCycle
            Else
                lSvCPTTEnd = DateValue(gAdjYear(smSvOffAirDate))
            End If
        Else
            lSvCPTTStart = DateValue(gAdjYear(smSvOnAirDate))
            If DateValue(gAdjYear(sLLD)) < DateValue(gAdjYear(smSvOffAirDate)) Then
                lSvCPTTEnd = DateValue(gAdjYear(sLLD))
            Else
                lSvCPTTEnd = DateValue(gAdjYear(smSvOffAirDate))
            End If
        End If
    End If
    If DateValue(gAdjYear(sOnAir)) <= DateValue(gAdjYear(sLLD)) Then
        If DateValue(gAdjYear(sDropDate)) <= DateValue(gAdjYear(sLLD)) Then
            lCPTTStart = DateValue(gAdjYear(sOnAir))
            If DateValue(gAdjYear(sDropDate)) < DateValue(gAdjYear(sOffAir)) Then
                lCPTTEnd = DateValue(gAdjYear(sDropDate)) '- iCycle
            Else
                lCPTTEnd = DateValue(gAdjYear(sOffAir))
            End If
        Else
            lCPTTStart = DateValue(gAdjYear(sOnAir))
            If DateValue(gAdjYear(sLLD)) < DateValue(gAdjYear(sOffAir)) Then
                lCPTTEnd = DateValue(gAdjYear(sLLD))
            Else
                lCPTTEnd = DateValue(gAdjYear(sOffAir))
            End If
        End If
    End If
    
    'Check to see if start date Advanced.  If so delete weeks
    If ((lSvCPTTStart < lCPTTStart) And (lSvCPTTStart > 0)) Or ((lSvCPTTStart > 0) And (lCPTTStart = 0)) Then
        'Remove
        lSDate = lSvCPTTStart
        If iCycle Mod 7 = 0 Then
            Do While Weekday(Format$(lSDate, "m/d/yyyy")) <> iWkDay
                lSDate = lSDate - 1 '+ 1
            Loop
        End If
        If lCPTTStart > 0 Then
            lEDate = lCPTTStart - iCycle
        Else
            lEDate = lSvCPTTEnd
        End If
        If lSDate <= lEDate Then
            sMsg = "Deleted weeks: " & Format$(lSDate, sgShowDateForm) & "-" & Format$(lEDate, sgShowDateForm)
            gLogMsg smStationName & " " & sMsg, "FastEndSummary.Txt", False
            Do
                SQLQuery = "DELETE FROM cptt WHERE (cpttAtfCode = " & lmAttCode '& " And cpttShfCode =" & imShttCode & " And cpttVefCode =" & imVefCode
                SQLQuery = SQLQuery & " AND ((cpttStatus = 2) Or ((cpttStatus = 0) AND (cpttPostingStatus = 0)))"
                SQLQuery = SQLQuery & " AND cpttStartDate >= '" & Format$(lSDate, sgSQLDateForm) & "' And cpttStartDate <= '" & Format$(lSDate + 6, sgSQLDateForm) & "')"
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "FastEndSummary.Txt", "FastEnd-mDeleteCPTT"
                    mDeleteCPTT = False
                    Exit Function
                End If
                slCallLetters = gGetCallLettersByAttCode(lmAttCode)
                slVehicleName = gGetVehNameByVefCode(gGetVehCodeFromAttCode(CStr(lmAttCode)))
                gLogMsg "Deleting CPTT1: " & slCallLetters & " running on: " & slVehicleName & " for the period: " & Format$(lSDate, "mm/dd/yyyy") & " - " & Format$(lSDate + 6, "mm/dd/yyyy"), "FastEndSummary.Txt", False
                lSDate = lSDate + 7
            Loop While lSDate <= lEDate
            gFileChgdUpdate "cptt.mkd", True
        End If
    End If
    
    'Check to see if end date reduced.  If so, remove weeks
    If (lCPTTEnd < lSvCPTTEnd) And (lCPTTEnd > 0) Then
        'Remove
        lSDate = lCPTTEnd
        'Advance to next week as dates are for last week to air
        If iCycle Mod 7 = 0 Then
            lSDate = lSDate + 1
            Do While Weekday(Format$(lSDate, "m/d/yyyy")) <> iWkDay
                lSDate = lSDate + 1
            Loop
        Else
            lSDate = lSDate + 1
        End If
        lEDate = lSvCPTTEnd
        If lSDate <= lEDate Then
            sMsg = "Deleted weeks: " & Format$(lSDate, "m/d/yyyy") & "-" & Format$(lEDate, "m/d/yyyy")
            gLogMsg smStationName & " " & sMsg, "FastEndSummary.Txt", False
            Do
                'ilExported = False
                'If gCheckIfSpotsHaveBeenExported(imVefCode, Format$(lSDate, sgSQLDateForm), imAgreeType) Then
                '    ilExported = True
                'End If
                
                'D.S. 10/25/04
                'igChangedNewErased values  1 = changed, 2 = new, 3 = erased
                'If they are changing an agreement and it's already been exported then don't delete the CPTTs
                'If ilExported = False Then
                    SQLQuery = "DELETE FROM cptt WHERE (cpttAtfCode = " & lmAttCode '& " And cpttShfCode =" & imShttCode & " And cpttVefCode =" & imVefCode
                    SQLQuery = SQLQuery & " And cpttStartDate >='" & Format$(lSDate, sgSQLDateForm) & "' And cpttStartDate <= '" & Format$(lSDate + 6, sgSQLDateForm) & "')"
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "FastEndSummary.Txt", "FastEnd-mDeleteCPTT"
                        mDeleteCPTT = False
                        Exit Function
                    End If
                'End If
                slCallLetters = gGetCallLettersByAttCode(lmAttCode)
                slVehicleName = gGetVehNameByVefCode(gGetVehCodeFromAttCode(CStr(lmAttCode)))
                gLogMsg "Deleting CPTT2: " & slCallLetters & " running on: " & slVehicleName & " for the period: " & Format$(lSDate, "mm/dd/yyyy") & " - " & Format$(lSDate + 6, "mm/dd/yyyy"), "FastEndSummary.Txt", False
                
                'D.S. 10/25/04
                'If ilExported = False Then
                    SQLQuery = "DELETE FROM ast WHERE (astAtfCode = " & lmAttCode '& " And astShfCode =" & imShttCode & " And astVefCode =" & imVefCode
                    SQLQuery = SQLQuery & " And astAirDate >= '" & Format$(lSDate, sgSQLDateForm) & "' And astAirDate <= '" & Format$(lSDate + 6, sgSQLDateForm) & "')"
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "FastEndSummary.Txt", "FastEnd-mDeleteCPTT"
                        mDeleteCPTT = False
                        Exit Function
                    End If
                    gLogMsg "Deleting AST1: " & slCallLetters & " running on: " & slVehicleName & " for the period: " & Format$(lSDate, "mm/dd/yyyy") & " - " & Format$(lSDate + 6, "mm/dd/yyyy"), "FastEndSummary.Txt", False
                'End If
                
                'If ilExported Then
                    If (imWebType = True) Or (imUnivisionType = True) Then
                        ilRet = gAlertAdd("R", "S", imVefCode, Format$(lSDate, sgSQLDateForm))
                    End If
                'End If
                
                'Delete Spots on the web
                    If imWebType Then
                        slSDate = Format$(lSDate, sgSQLDateForm)
                        slEDate = DateAdd("d", 6, slSDate)
                        '9/17/12: SQL Server is only deleting spot from the second reference to pledgeStartDate
                        'SQLQuery = "Delete From Spots Where attCode = " & CLng(lmAttCode) & " And PledgeStartDate >= '" & Format$(gAdjYear(Format$(slSDate, "m/d/yy")), sgSQLDateForm) & "'" & " And PledgeStartDate >= '" & Format$(gAdjYear(Format$(slEDate, "m/d/yy")), sgSQLDateForm) & "'"
                        SQLQuery = "Delete From Spots Where attCode = " & CLng(lmAttCode) & " And PledgeStartDate >= '" & Format$(gAdjYear(Format$(slSDate, "m/d/yy")), sgSQLDateForm) & "'" & " And PledgeEndDate <= '" & Format$(gAdjYear(Format$(slEDate, "m/d/yy")), sgSQLDateForm) & "'"
                        llTtlRows = gExecWebSQLWithRowsEffected(SQLQuery)
                        If llTtlRows = -1 Then
                            gLogMsg "Error: gExecWebSQLWithRowsEffected Failed", "FastEndSummary.Txt", False
                            gLogMsg "    " & SQLQuery, "FastEndSummary.Txt", False
                            gMsgBox "The web failed to delete the specified week(s).", vbCritical
                        Else
                            gLogMsg "Deleting Web Spots1: " & slCallLetters & " running on: " & slVehicleName & " for the period: " & Format$(slSDate, "mm/dd/yyyy") & " - " & Format$(slEDate, "mm/dd/yyyy"), "FastEndSummary.Txt", False
                        End If
                        SQLQuery = "Delete From SpotRevisions Where attCode = " & CLng(lmAttCode) & " And PledgeStartDate >= '" & Format$(gAdjYear(Format$(slSDate, "m/d/yy")), sgSQLDateForm) & "'" & " And PledgeEndDate <= '" & Format$(gAdjYear(Format$(slEDate, "m/d/yy")), sgSQLDateForm) & "'"
                        llTtlRows = gExecWebSQLWithRowsEffected(SQLQuery)
                        If llTtlRows = -1 Then
                            gLogMsg "Error: gExecWebSQLWithRowsEffected Failed", "FastEndSummary.Txt", False
                            gLogMsg "    " & SQLQuery, "FastEndSummary.Txt", False
                            gMsgBox "The web failed to delete SpotRevisions the specified week(s).", vbCritical
                        Else
                            gLogMsg "Deleting Web SpotRevisions: " & slCallLetters & " running on: " & slVehicleName & " for the period: " & Format$(slSDate, "mm/dd/yyyy") & " - " & Format$(slEDate, "mm/dd/yyyy"), "FastEndSummary.Txt", False
                        End If
                    End If
                
                'advance the dates by one week until we reach the end date
                lSDate = lSDate + 7
                '9/17/12: advance to next week to handle pledge setup as After
            Loop While lSDate <= lEDate + 6
            
        End If
    End If

    mDeleteCPTT = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFastEnd-mDeleteCPTT"
    mDeleteCPTT = False
End Function

Private Sub rbcType_Click(Index As Integer)
    If rbcType(Index).Value Then
        lbcStations(0).Clear
        lbcStations(1).Clear
        If Index = 0 Then   'Affiliate
            'Populate cboVehicle with Affiliates
            cboStation.Visible = True
            cboVehicle.Visible = False
            lblList.Caption = "List of Vehicles to Terminate"
            lblDontTerminate.Caption = "Do Not Terminate these Vehicles"
            lblTerminate.Caption = "Terminate these Vehicles"
        Else                'Vehicle
            cboVehicle.Visible = True
            cboStation.Visible = False
            lblList.Caption = "List of Affiliates to Terminate"
            lblDontTerminate.Caption = "Do Not Terminate these Affiliates"
            lblTerminate.Caption = "Terminate these Affiliates"
        End If
    End If
End Sub

Private Sub txtBrowse_LostFocus()
    If txtBrowse.Text <> "" Then
        mExternalFile
    End If
End Sub

Private Sub txtEndDate_Change()
    mGenOK
End Sub

Private Sub txtEndDate_Click()
    mGenOK
End Sub
Private Sub mExternalFile()

    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject
    Dim slRetString As String
    Dim slLocation As String
    Dim slTemp As String
    Dim llRet As Long
    Dim ilLoop As Integer
    Dim llAtt As Long
    Dim ilStation As Integer
    Dim slCallLetters() As String
    Dim ilvehicle As Integer
    Dim ilVef As Integer
    Dim slVehicles() As String
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    slLocation = Trim$(txtBrowse.Text)
    If fs.FILEEXISTS(slLocation) Then
        Set tlTxtStream = fs.OpenTextFile(slLocation, ForReading, False)
    Else
        Screen.MousePointer = vbDefault
        gMsgBox "** No Data Available **"
        Exit Sub
    End If
        
    Do While tlTxtStream.AtEndOfStream <> True
        slRetString = tlTxtStream.ReadLine
        If rbcType(0).Value Then
            slVehicles = Split(slRetString, ",")
            If IsArray(slVehicles) Then
                For ilvehicle = 0 To UBound(slVehicles) Step 1
                    For ilVef = 0 To UBound(tgVehicleInfo) - 1 Step 1
                        If UCase(slVehicles(ilvehicle)) = UCase(Trim$(tgVehicleInfo(ilVef).sVehicle)) Then
                            ilLoop = 0
                            Do While ilLoop <= lbcStations(0).ListCount - 1
                                llAtt = gBinarySearchAtt(CLng(lbcStations(0).ItemData(ilLoop)))
                                If llAtt <> -1 Then
                                    If tgVehicleInfo(ilVef).iCode = tgAttInfo1(llAtt).attvefCode Then
                                        lbcStations(1).AddItem lbcStations(0).List(ilLoop)
                                        lbcStations(1).ItemData(lbcStations(1).NewIndex) = lbcStations(0).ItemData(ilLoop)
                                        lbcStations(0).RemoveItem (ilLoop)
                                        Exit Do
                                    End If
                                End If
                                ilLoop = ilLoop + 1
                            Loop
                        End If
                    Next ilVef
                Next ilvehicle
            End If
        Else
            slCallLetters = Split(slRetString, ",")
            If IsArray(slCallLetters) Then
                For ilStation = 0 To UBound(slCallLetters) Step 1
                    llRet = gBinarySearchStation(Trim$(slCallLetters(ilStation)))
                    If llRet <> -1 Then
                        ilLoop = 0
                        Do While ilLoop <= lbcStations(0).ListCount - 1
                            llAtt = gBinarySearchAtt(CLng(lbcStations(0).ItemData(ilLoop)))
                            If llAtt <> -1 Then
                                If tgStationInfo(llRet).iCode = tgAttInfo1(llAtt).attShttCode Then
                                    lbcStations(1).AddItem lbcStations(0).List(ilLoop)
                                    lbcStations(1).ItemData(lbcStations(1).NewIndex) = lbcStations(0).ItemData(ilLoop)
                                    lbcStations(0).RemoveItem (ilLoop)
                                End If
                            End If
                            ilLoop = ilLoop + 1
                        Loop
                    End If
                Next ilStation
            End If
        End If
    Loop
    tlTxtStream.Close
    mGenOK
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmFastEnd-mExternalFile: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "FastEndSummary.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    
End Sub


Private Sub mPopStations()
    Dim llMkt As Long
    Dim slMarket As String
    Dim shtt_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "SELECT distinct shttCallLetters, shttMktCode, shttCode"
    SQLQuery = SQLQuery + " FROM att, shtt "
    SQLQuery = SQLQuery + " WHERE "
    SQLQuery = SQLQuery + " shttCode = attShfCode"
    SQLQuery = SQLQuery + " AND (attOffAir  >= '" & Format(gNow(), sgSQLDateForm) & "')"
    SQLQuery = SQLQuery + " AND (attDropDate >= '" & Format(gNow(), sgSQLDateForm) & "')"
    Set shtt_rst = gSQLSelectCall(SQLQuery)
    Do While Not shtt_rst.EOF
        slMarket = ""
        For llMkt = 0 To UBound(tgMarketInfo) - 1
            If tgMarketInfo(llMkt).lCode = shtt_rst!shttMktCode Then
                slMarket = Trim$(tgMarketInfo(llMkt).sName)
                Exit For
            End If
        Next llMkt
        If slMarket <> "" Then
            cboStation.AddItem Trim$(shtt_rst!shttCallLetters) & ", " & slMarket
        Else
            cboStation.AddItem Trim$(shtt_rst!shttCallLetters)
        End If
        cboStation.ItemData(cboStation.NewIndex) = shtt_rst!shttCode
        shtt_rst.MoveNext
    Loop
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Fast End-mPopStations"
End Sub
