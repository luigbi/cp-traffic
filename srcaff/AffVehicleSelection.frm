VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmVehicleSelection 
   Caption         =   "Vehicle Selection"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   Icon            =   "AffVehicleSelection.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   9105
   Begin VB.CommandButton cmcPostpone 
      Caption         =   "Postpone All"
      Height          =   375
      Left            =   6195
      TabIndex        =   9
      Top             =   5370
      Width           =   1110
   End
   Begin VB.TextBox edcTitle2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Vehicles: Export Now"
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
      TabIndex        =   7
      Text            =   "Vehicles: Exported Later"
      Top             =   1890
      Width           =   3150
   End
   Begin VB.CommandButton cmcReset 
      Caption         =   "Restore"
      Height          =   375
      Left            =   4710
      TabIndex        =   5
      Top             =   5370
      Width           =   1110
   End
   Begin VB.ListBox lbcExportVehicles 
      Height          =   2985
      ItemData        =   "AffVehicleSelection.frx":08CA
      Left            =   5010
      List            =   "AffVehicleSelection.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2160
      Width           =   3855
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   2985
      ItemData        =   "AffVehicleSelection.frx":08CE
      Left            =   150
      List            =   "AffVehicleSelection.frx":08D0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2175
      Width           =   3855
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
      TabIndex        =   2
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
      TabIndex        =   6
      Top             =   135
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3016
   End
End
Attribute VB_Name = "frmVehicleSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmVehicleSelection - displays missed spots to be changed to Makegoods
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

Private Sub cmcPostpone_Click()
    Dim ilLoop As Integer
    For ilLoop = lbcExportVehicles.ListCount - 1 To 0 Step -1
        lbcVehicles.AddItem lbcExportVehicles.List(ilLoop)
        lbcVehicles.ItemData(lbcVehicles.NewIndex) = lbcExportVehicles.ItemData(ilLoop)
        lbcExportVehicles.RemoveItem ilLoop
    Next ilLoop
End Sub

Private Sub cmcReset_Click()
    lbcExportVehicles.Clear
    mPopVehicle
    mSetCtrls
    imFieldChgd = False
    mSetCommands
End Sub

Private Sub cmdCancel_Click()
    igExportReturn = 0
    Unload frmVehicleSelection
End Sub

Private Sub cmdDone_Click()
    Dim ilRet As Integer
    
    igExportReturn = 0
    If imFieldChgd Then
    '    If gMsgBox("Save all changes?", vbYesNo) = vbYes Then
            mMousePointer vbHourglass
            ilRet = mSave()
            mMousePointer vbDefault
            If Not ilRet Then
                Exit Sub
            End If
            igVehicleSpecChgFlag = True
            igExportReturn = 1
    '    End If
    End If
    Unload frmVehicleSelection
    Exit Sub
   
End Sub


Private Sub Form_Activate()
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    If imFirstTime Then
        mMousePointer vbHourglass
        Me.Caption = Me.Caption & ": " & sgExportName
        mPopVehicle
        udcCriteria.Enabled = False
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
    gSetFonts frmVehicleSelection
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
    Set frmVehicleSelection = Nothing
End Sub


Private Sub mInit()
    Dim ilRet As Integer
    Dim llVeh As Long
    
    
    imFirstTime = True
    imBSMode = False
        
    pbcClickFocus.Left = -100

End Sub
Private Sub mPopVehicle()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilEmbeddedAllowed As Integer
    Dim llVef As Long
    Dim llNext As Long
    
    On Error GoTo ErrHand
    ilEmbeddedAllowed = False
    ilRet = gPopVff()
    lbcVehicles.Clear
    lbcExportVehicles.Clear
    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        If tgVehicleInfo(ilLoop).iProducerArfCode > 0 Then
            ilEmbeddedAllowed = True
        End If
    Next ilLoop
    udcCriteria.Embedded = ilEmbeddedAllowed
    
    llNext = tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt
    Do While llNext <> -1
        llVef = gBinarySearchVef(CLng(tgEvtInfo(llNext).iVefCode))
        If llVef <> -1 Then
            lbcExportVehicles.AddItem Trim$(tgVehicleInfo(llVef).sVehicle)
            lbcExportVehicles.ItemData(lbcExportVehicles.NewIndex) = tgVehicleInfo(llVef).iCode
        End If
        llNext = tgEvtInfo(llNext).lNextEvt
    Loop
    If sgUstWin(14) = "V" Then
        lbcVehicles.Enabled = False
        lbcExportVehicles.Enabled = False
        cmcReset.Enabled = False
        cmdDone.Enabled = False
    End If
    mSetCommands
    On Error Resume Next
    rst_Evt.Close
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffExportLog.txt", "Vehicle Selection-mPopVehicle"
End Sub


Private Sub mMousePointer(ilMousepointer As Integer)
    Screen.MousePointer = ilMousepointer
End Sub



Private Function mSave()
    
    Dim ilRet As Integer
    Dim llEvtInfo As Long
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim llNext As Long
    Dim llSvNext As Long
    Dim llCheck As Long
    
    On Error GoTo ErrHand
    
    
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
    gHandleError "AffErrorLog.txt", ""
End Function

Private Sub mSetCommands()
End Sub



Private Sub mSetCtrls()
    udcCriteria.Action 3
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
