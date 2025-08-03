VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmVehicleSpec 
   Caption         =   "Specification Parameters"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   Icon            =   "AffVehicleSpec.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   9105
   Begin VB.TextBox edcTitle2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Vehicles to be Exported"
      Top             =   1890
      Width           =   3150
   End
   Begin VB.TextBox edcTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   510
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Vehicles not to be Exported"
      Top             =   1890
      Width           =   3150
   End
   Begin VB.CommandButton cmcReset 
      Caption         =   "Restore"
      Height          =   375
      Left            =   4770
      TabIndex        =   6
      Top             =   5370
      Width           =   1110
   End
   Begin VB.ListBox lbcExportVehicles 
      Height          =   2985
      ItemData        =   "AffVehicleSpec.frx":08CA
      Left            =   5010
      List            =   "AffVehicleSpec.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   3855
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   2985
      ItemData        =   "AffVehicleSpec.frx":08CE
      Left            =   150
      List            =   "AffVehicleSpec.frx":08D0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2175
      Width           =   3855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   6255
      TabIndex        =   2
      Top             =   5370
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3180
      TabIndex        =   1
      Top             =   5370
      Width           =   1110
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4725
      Width           =   45
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   540
      Top             =   5115
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6000
      FormDesignWidth =   9105
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   1710
      TabIndex        =   0
      Top             =   5370
      Width           =   1110
   End
   Begin V81Affiliate.AffExportCriteria udcCriteria 
      Height          =   1710
      Left            =   165
      TabIndex        =   7
      Top             =   135
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3016
   End
End
Attribute VB_Name = "frmVehicleSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmVehicleSpec - displays missed spots to be changed to Makegoods
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit


Private imFirstTime As Integer
Private imBSMode As Integer
Private imFieldChgd As Integer

'ISCI
Private imEmbeddedAllowed As Integer

Private rst_Eht As ADODB.Recordset
Private rst_Evt As ADODB.Recordset

Private Sub cmcReset_Click()
    lbcExportVehicles.Clear
    mPopVehicle True
    mSetCtrls
    'If restoring old setting, reset imFieldChg
    If (tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt <> -1) Or (lbcExportVehicles.ListCount <= 0) Then
        imFieldChgd = False
    Else
        imFieldChgd = True
    End If
    mSetCommands
End Sub

Private Sub cmdCancel_Click()
    Unload frmVehicleSpec
End Sub

Private Sub cmdClear_Click()
    lbcExportVehicles.Clear
    mPopVehicle False
    'mSetCtrls
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub cmdDone_Click()
    Dim ilRet As Integer
    
    If imFieldChgd Then
    '    If gMsgBox("Save all changes?", vbYesNo) = vbYes Then
            mMousePointer vbHourglass
            ilRet = mSave()
            mMousePointer vbDefault
            If Not ilRet Then
                Exit Sub
            End If
            igVehicleSpecChgFlag = True
    '    End If
    End If
    Unload frmVehicleSpec
    Exit Sub
   
End Sub


Private Sub Form_Activate()
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    If imFirstTime Then
        mMousePointer vbHourglass
        Me.Caption = Me.Caption & ": " & sgExportName
        mPopVehicle True
        'force redraw of controls within user control
        udcCriteria.Height = (3 * edcTitle1.Top) / 2 - udcCriteria.Top - 60
        'udcCriteria.Action 1
        udcCriteria.Action 2
        mSetCtrls
        imFirstTime = False
        mSetCommands
        mMousePointer vbDefault
    End If
    
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.25
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 1.4
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmVehicleSpec
    gCenterStdAlone Me
    
End Sub

Private Sub Form_Load()
    
    mMousePointer vbHourglass
    
    mInit
    mMousePointer vbDefault
    Exit Sub
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    rst_Eht.Close
    rst_Evt.Close
    Set frmVehicleSpec = Nothing
End Sub


Private Sub mInit()
    Dim ilRet As Integer
    Dim llVeh As Long
    
    
    imFirstTime = True
    imBSMode = False
    If (sgExptSpec <> "Y") Or (sgUstWin(14) = "V") Then
        lbcVehicles.Enabled = False
        lbcExportVehicles.Enabled = False
        udcCriteria.Enabled = False
        cmdClear.Enabled = False
        cmcReset.Enabled = False
        cmdDone.Enabled = False
    End If
        
    pbcClickFocus.Left = -100

End Sub
Private Sub mPopVehicle(blPopExportVehicles As Boolean)
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilEmbeddedAllowed As Integer
    Dim llNext As Long
    Dim slNowDate As String
    '7341
    Dim ilVff As Integer
    Dim ilVpf As Integer
    
    On Error GoTo ErrHand
    ilEmbeddedAllowed = False
    ilRet = gPopVff()
    lbcVehicles.Clear
    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        If (mVehicleAddTest(ilLoop) And (sgExportTypeChar = "A")) Or (sgExportTypeChar <> "A") Then
            lbcVehicles.AddItem Trim$(tgVehicleInfo(ilLoop).sVehicle)
            lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(ilLoop).iCode
            If tgVehicleInfo(ilLoop).iProducerArfCode > 0 Then
                ilEmbeddedAllowed = True
            End If
        End If
    Next ilLoop
    udcCriteria.Embedded = ilEmbeddedAllowed
    If blPopExportVehicles Then
        'SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & lgExportEhtCode
        'Set rst_Evt = gSQLSelectCall(SQLQuery)
        'Do While Not rst_Evt.EOF
        llNext = tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt
        If llNext <> -1 Then
            Do While llNext <> -1
                For ilLoop = 0 To lbcVehicles.ListCount - 1 Step 1
                    'If rst_Evt!evtVefCode = Val(lbcVehicles.ItemData(ilLoop)) Then
                    If tgEvtInfo(llNext).iVefCode = Val(lbcVehicles.ItemData(ilLoop)) Then
                        lbcVehicles.Selected(ilLoop) = True
                        Exit For
                    End If
                Next ilLoop
            '    rst_Evt.MoveNext
                llNext = tgEvtInfo(llNext).lNextEvt
            Loop
            imFieldChgd = False
        Else
            'Load based on setting in Traffic and/or Affiliate
            slNowDate = Format(gNow(), sgSQLDateForm)
            Select Case sgExportTypeChar
                Case "A"    'Aff Log
                    SQLQuery = "SELECT DISTINCT attVefCode FROM att WHERE attDropDate > '" & slNowDate & "' AND attOffAir > '" & slNowDate & "' AND attExportType <> 0"
                    Set rst = gSQLSelectCall(SQLQuery)
                    Do While Not rst.EOF
                        For ilLoop = 0 To lbcVehicles.ListCount - 1 Step 1
                            If rst!attvefCode = Val(lbcVehicles.ItemData(ilLoop)) Then
                                lbcVehicles.Selected(ilLoop) = True
                                Exit For
                            End If
                        Next ilLoop
                        rst.MoveNext
                    Loop
                Case "D"    'IDC
                    SQLQuery = "SELECT DISTINCT attVefCode FROM att WHERE attDropDate > '" & slNowDate & "' AND attOffAir > '" & slNowDate & "' AND RTrim(attIDCReceiverID) <> ''"
                    Set rst = gSQLSelectCall(SQLQuery)
                    Do While Not rst.EOF
                        For ilLoop = 0 To lbcVehicles.ListCount - 1 Step 1
                            If rst!attvefCode = Val(lbcVehicles.ItemData(ilLoop)) Then
                                lbcVehicles.Selected(ilLoop) = True
                                Exit For
                            End If
                        Next ilLoop
                        rst.MoveNext
                    Loop
                Case "X"    'X-Digital
                '7341 note the loop is reversed; otherwise, get errors
                    For ilLoop = lbcVehicles.ListCount - 1 To 0 Step -1
                           If igExportSource = 2 Then DoEvents
                            ilVff = gBinarySearchVff(lbcVehicles.ItemData(ilLoop))
                            ilVpf = gBinarySearchVpf(CLng(lbcVehicles.ItemData(ilLoop)))
                            If (ilVff <> -1) And (ilVpf <> -1) Then
                                If igExportSource = 2 Then DoEvents
                                If (Trim$(tgVffInfo(ilVff).sXDProgCodeID) <> "") Or (tgVpfOptions(ilVpf).iInterfaceID > 0) Then
                                    If (Trim$(UCase(tgVffInfo(ilVff).sXDProgCodeID)) <> "MERGE") Then
                                        lbcVehicles.Selected(ilLoop) = True
                                    End If
                                End If
                            End If
                    Next ilLoop
                
'                    SQLQuery = "SELECT DISTINCT vffVefCode FROM vff_Vehicle_Features WHERE RTrim(vffXDProgCodeID) <> ''"
'                    Set rst = gSQLSelectCall(SQLQuery)
'                    Do While Not rst.EOF
'                        For ilLoop = 0 To lbcVehicles.ListCount - 1 Step 1
'                            If rst!vffvefCode = Val(lbcVehicles.ItemData(ilLoop)) Then
'                                lbcVehicles.Selected(ilLoop) = True
'                                Exit For
'                            End If
'                        Next ilLoop
'                        rst.MoveNext
'                    Loop
                Case "W"    'Wegener
                    SQLQuery = "SELECT DISTINCT vpfVefKCode FROM vpf_Vehicle_Options WHERE vpfWegenerExport = 'Y'"
                    Set rst = gSQLSelectCall(SQLQuery)
                    Do While Not rst.EOF
                        For ilLoop = 0 To lbcVehicles.ListCount - 1 Step 1
                            If rst!vpfvefKCode = Val(lbcVehicles.ItemData(ilLoop)) Then
                                lbcVehicles.Selected(ilLoop) = True
                                Exit For
                            End If
                        Next ilLoop
                        rst.MoveNext
                    Loop
                '6079 'IPump
                Case "P"
                    SQLQuery = "SELECT DISTINCT vffVefCode FROM vff_Vehicle_Features WHERE RTrim(vffExportIPump) = 'Y' and length(vffIpumpEventTypeOv) = 0"
                    Set rst = gSQLSelectCall(SQLQuery)
                    Do While Not rst.EOF
                        For ilLoop = 0 To lbcVehicles.ListCount - 1 Step 1
                            If rst!vffvefCode = Val(lbcVehicles.ItemData(ilLoop)) Then
                                lbcVehicles.Selected(ilLoop) = True
                                Exit For
                            End If
                        Next ilLoop
                        rst.MoveNext
                    Loop
               
            End Select
        End If
        'cmdMoveRight_Click
    End If
    mSetCommands
    On Error Resume Next
    rst_Evt.Close
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmVehicleSpec-mPopulateVehicle"
End Sub


Private Sub mMousePointer(ilMousepointer As Integer)
    Screen.MousePointer = ilMousepointer
End Sub



Private Function mSave() As Integer
    
    Dim ilRet As Integer
    Dim llEvtInfo As Long
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim llNext As Long
    Dim llSvNext As Long
    Dim llCheck As Long
    
    On Error GoTo ErrHand
    
    If Not udcCriteria.Verify Then
        MsgBox "Not All Required Answers Specified", vbInformation + vbOKOnly, "Not Complete"
        mSave = False
        Exit Function
    End If
    If lbcExportVehicles.ListCount <= 0 Then
        MsgBox "No Vehicle to Export Specified", vbInformation + vbOKOnly, "Not Complete"
        mSave = False
        Exit Function
    End If
    udcCriteria.Action 5
    
    llNext = tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt
    Do While llNext <> -1
        llSvNext = tgEvtInfo(llNext).lNextEvt
        tgEvtInfo(llNext).lNextEvt = -9999
        llNext = llSvNext
    Loop
    tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt = -1
    For ilLoop = 0 To lbcExportVehicles.ListCount - 1 Step 1
        If tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt = -1 Then
            llNext = -1
        Else
            llNext = tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt
        End If
        llEvtInfo = UBound(tgEvtInfo)
        For llCheck = 0 To UBound(tgEvtInfo) - 1 Step 1
            If tgEvtInfo(llCheck).lNextEvt = -9999 Then
                llEvtInfo = llCheck
                Exit For
            End If
        Next llCheck
        tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt = llEvtInfo
        tgEvtInfo(llEvtInfo).iVefCode = Val(lbcExportVehicles.ItemData(ilLoop))
        If tgEvtInfo(llEvtInfo).lNextEvt = -9999 Then
            tgEvtInfo(llEvtInfo).lNextEvt = llNext
        Else
            tgEvtInfo(llEvtInfo).lNextEvt = llNext
            ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
        End If
    Next ilLoop
    imFieldChgd = False
    Screen.MousePointer = vbDefault
    mSave = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmVehicleSpec-mSave"
End Function

Private Sub mSetCommands()
End Sub



Private Sub mSetCtrls()
    Dim ilValue7 As Integer
    Dim ilValue8 As Integer
    Dim ilRet As Integer
    udcCriteria.Action 3
On Error Resume Next
    'Dan 11/25/14 7180 'reexport all' doesn't show here.
     If Trim$(sgExportTypeChar) = "X" Then
        udcCriteria.XReExportVisible = False
     End If
     '7459
     If Trim$(sgExportTypeChar) = "R" Then
        'dan m made global
        ilRet = gSiteISCIAndOrBreak()
'        SQLQuery = "Select spfUsingFeatures7,spfUsingFeatures8 From SPF_Site_Options"
'        Set rst = gSQLSelectCall(SQLQuery)
'        If Not rst.EOF Then
'            ilValue7 = Asc(rst!spfusingfeatures7)
'            If (ilValue7 And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT Then
'                ilRet = 1
'            End If
'            ilValue8 = Asc(rst!spfUsingFeatures8)
'            If (ilValue8 And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT Then
'                ilRet = ilRet + 2
'            End If
'        End If
        If ilRet = 3 Then
            udcCriteria.RPrefixVisible = True
        Else
            udcCriteria.RPrefixVisible = False
        End If
    End If
End Sub

Private Sub lbcExportVehicles_Click()
    Dim ilLoop As Integer
    For ilLoop = lbcExportVehicles.ListCount - 1 To 0 Step -1
        If lbcExportVehicles.Selected(ilLoop) Then
            lbcVehicles.AddItem lbcExportVehicles.List(ilLoop)
            lbcVehicles.ItemData(lbcVehicles.NewIndex) = lbcExportVehicles.ItemData(ilLoop)
            lbcExportVehicles.RemoveItem ilLoop
        End If
    Next ilLoop
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub lbcVehicles_Click()
    Dim ilLoop As Integer
    For ilLoop = lbcVehicles.ListCount - 1 To 0 Step -1
        If lbcVehicles.Selected(ilLoop) Then
            lbcExportVehicles.AddItem lbcVehicles.List(ilLoop)
            lbcExportVehicles.ItemData(lbcExportVehicles.NewIndex) = lbcVehicles.ItemData(ilLoop)
            lbcVehicles.RemoveItem ilLoop
        End If
    Next ilLoop
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub udcCriteria_SetChgFlag()
    imFieldChgd = True
End Sub

Private Function mVehicleAddTest(ilVefIndex As Integer) As Boolean
    Dim ilVff As Integer
    Dim ilVef As Integer
    Dim ilSetValue As Integer
    Dim rstATT As ADODB.Recordset
    
    On Error GoTo ErrHand:
    ilSetValue = True
    If tgVehicleInfo(ilVefIndex).sVehType = "L" Then
        'Check to see if any vehicle which belong to the Log vehicle is to be Merged
        ilSetValue = False
        For ilVef = 0 To UBound(tgVehicleInfo) - 1 Step 1
            If tgVehicleInfo(ilVefIndex).iCode = tgVehicleInfo(ilVef).iVefCode Then
                For ilVff = 0 To UBound(tgVffInfo) - 1 Step 1
                    If tgVehicleInfo(ilVef).iCode = tgVffInfo(ilVff).iVefCode Then
                        If tgVffInfo(ilVff).sMergeWeb <> "S" Then
                            ilSetValue = True
                            Exit For
                        End If
                    End If
                Next ilVff
            End If
        Next ilVef
        If Not ilSetValue Then
            'Test if agreement exist for Log vehicle
            SQLQuery = "Select MAX(attVefCode) from att where attVefCode =" & Str$(tgVehicleInfo(ilVefIndex).iCode)
            Set rstATT = gSQLSelectCall(SQLQuery)
            If rstATT(0).Value = tgVehicleInfo(ilVefIndex).iCode Then
                ilSetValue = True
            End If
        End If
    ElseIf ((tgVehicleInfo(ilVefIndex).sVehType = "C") Or (tgVehicleInfo(ilVefIndex).sVehType = "G") Or (tgVehicleInfo(ilVefIndex).sVehType = "A")) And (tgVehicleInfo(ilVefIndex).iVefCode > 0) Then
        'Check to see if the vehicle that references a Log vehicle is to have a separte agreement from the log vehicle
        ilSetValue = False
        For ilVff = 0 To UBound(tgVffInfo) - 1 Step 1
            If tgVehicleInfo(ilVefIndex).iCode = tgVffInfo(ilVff).iVefCode Then
                If tgVffInfo(ilVff).sMergeWeb = "S" Then
                    ilSetValue = True
                    Exit For
                End If
            End If
        Next ilVff
        If Not ilSetValue Then
            'Test if agreement exist for vehicle that references a Log vehicle
            SQLQuery = "Select MAX(attVefCode) from att where attVefCode =" & Str$(tgVehicleInfo(ilVefIndex).iCode)
            Set rstATT = gSQLSelectCall(SQLQuery)
            If rstATT(0).Value = tgVehicleInfo(ilVefIndex).iCode Then
                ilSetValue = True
            End If
        End If
    End If
    mVehicleAddTest = ilSetValue
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Vehicle Specification-mVehicleAddTest"
    mVehicleAddTest = False
End Function
