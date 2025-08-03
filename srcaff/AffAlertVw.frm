VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAlertVw 
   Caption         =   "View Alerts"
   ClientHeight    =   4125
   ClientLeft      =   900
   ClientTop       =   2820
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "AffAlertVw.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4125
   ScaleWidth      =   9720
   Begin VB.CommandButton cmcClear 
      Appearance      =   0  'Flat
      Caption         =   "&Clear"
      Height          =   285
      Left            =   5010
      TabIndex        =   12
      Top             =   3720
      Width           =   945
   End
   Begin ComctlLib.ListView lbcView 
      Height          =   2205
      Index           =   0
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   3889
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Creation Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Creation Time"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Vehicle Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Export Date"
         Object.Width           =   1341
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Reason"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame frcView 
      Caption         =   "View"
      Height          =   570
      Left            =   120
      TabIndex        =   7
      Top             =   345
      Width           =   9390
      Begin VB.OptionButton rbcView 
         Caption         =   "Info Alerts"
         Height          =   210
         Index           =   5
         Left            =   6405
         TabIndex        =   15
         Top             =   255
         Width           =   1335
      End
      Begin VB.OptionButton rbcView 
         Caption         =   "Web Vendors"
         Height          =   210
         Index           =   4
         Left            =   4815
         TabIndex        =   13
         Top             =   255
         Width           =   1425
      End
      Begin VB.OptionButton rbcView 
         Caption         =   "Agreement Alerts"
         Height          =   210
         Index           =   3
         Left            =   2865
         TabIndex        =   10
         Top             =   255
         Width           =   1905
      End
      Begin VB.OptionButton rbcView 
         Caption         =   "Traffic Alerts"
         Height          =   210
         Index           =   2
         Left            =   7800
         TabIndex        =   11
         Top             =   255
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.OptionButton rbcView 
         Caption         =   "Export ISCI"
         Height          =   210
         Index           =   1
         Left            =   1575
         TabIndex        =   9
         Top             =   255
         Width           =   1320
      End
      Begin VB.OptionButton rbcView 
         Caption         =   "Export Spots"
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   255
         Width           =   1455
      End
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   1770
      Width           =   75
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   3255
      TabIndex        =   0
      Top             =   3720
      Width           =   945
   End
   Begin ComctlLib.ListView lbcView 
      Height          =   2205
      Index           =   1
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1365
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   3889
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Creation Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Creation Time"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Vehicle Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Export Date"
         Object.Width           =   1341
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Reason"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ListView lbcView 
      Height          =   2205
      Index           =   2
      Left            =   135
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   3889
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Creation Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Creation Time"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Vehicle"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Log Date"
         Object.Width           =   1341
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Reason"
         Object.Width           =   2540
      EndProperty
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8955
      Top             =   3555
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   4125
      FormDesignWidth =   9720
   End
   Begin ComctlLib.ListView lbcView 
      Height          =   2205
      Index           =   3
      Left            =   135
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1335
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   3889
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Creation Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Creation Time"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Vehicle Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Changed Date"
         Object.Width           =   1341
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Reason"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "AufCode"
         Object.Width           =   0
      EndProperty
   End
   Begin ComctlLib.ListView lbcView 
      Height          =   2205
      Index           =   4
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1440
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   3889
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Creation Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Creation Time"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Station\Vehicle"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Monday Date"
         Object.Width           =   1341
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Vendor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Reason"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ListView lbcView 
      Height          =   2205
      Index           =   5
      Left            =   150
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1335
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   3889
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Creation Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Creation Time"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Source"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "File Name"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label lbcScreen 
      Caption         =   "View Alerts"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   45
      Width           =   1965
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   75
      Top             =   3630
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "frmAlertVw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: frmAlertVw.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim tmAufView() As AUFVIEW
Dim imRowSelected As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imShowHelpMsg As Integer    'True=Show Help messages; False=Ignore help message system



Private Sub cmcClear_Click()
    Dim ilRet As Integer
    Dim ilRow As Integer
    
    If rbcView(3).Value = False Then
        Exit Sub
    End If
    If (lbcView(3).ListItems.Count <= 0) Or (imRowSelected <= 0) Then
        Exit Sub
    End If
    ilRet = 0
    On Error GoTo cmcClearErr
    ilRow = lbcView(3).SelectedItem.Index
    If ilRet <> 0 Then
        Exit Sub
    End If
    On Error GoTo 0
    If ilRow >= 1 Then
        ilRet = MsgBox("Clear Selected row", vbYesNo + vbQuestion, "Clear Row")
        If ilRet = vbYes Then
            SQLAlertQuery = "UPDATE AUF_ALERT_USER SET "
            SQLAlertQuery = SQLAlertQuery & "aufStatus = 'C'" & ", "
            SQLAlertQuery = SQLAlertQuery & "aufClearDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
            SQLAlertQuery = SQLAlertQuery & "aufClearTime = '" & Format$(gNow(), sgSQLTimeForm) & "', "
            SQLAlertQuery = SQLAlertQuery & "aufClearUstCode = " & igUstCode & " "
            SQLAlertQuery = SQLAlertQuery & "WHERE aufCode = " & lbcView(3).ListItems(ilRow).SubItems(5)
            'cnn.Execute SQLAlertQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLAlertQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                gHandleError "AffErrorLog.txt", "AlertView-cmcClear_Click"
                Exit Sub
            End If
            lbcView(3).ListItems.Remove ilRow
            cmcClear.Enabled = False
            imRowSelected = -1
            ilRet = gAlertForceCheck()
        End If
    End If
    Exit Sub
cmcClearErr:
    ilRet = 1
    Resume Next
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "AlerVw-cmcClear"
    Exit Sub
End Sub

Private Sub cmcDone_Click()
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
    Dim ilLoop As Integer
    Dim ilCol As Integer
    If imFirstActivate Then
        imFirstActivate = False
        lbcView(0).ColumnHeaders.Item(1).Width = lbcView(0).Width / 8
        lbcView(0).ColumnHeaders.Item(2).Width = lbcView(0).Width / 7.75
        lbcView(0).ColumnHeaders.Item(4).Width = lbcView(0).Width / 10
        lbcView(0).ColumnHeaders.Item(5).Width = lbcView(0).Width / 6.25
        lbcView(0).ColumnHeaders.Item(3).Width = lbcView(0).Width - lbcView(0).ColumnHeaders.Item(1).Width - lbcView(0).ColumnHeaders.Item(2).Width - lbcView(0).ColumnHeaders.Item(4).Width - lbcView(0).ColumnHeaders.Item(5).Width - 7 * GRIDSCROLLWIDTH
        For ilLoop = 1 To 3 Step 1
            For ilCol = 1 To 5 Step 1
                lbcView(ilLoop).ColumnHeaders.Item(ilCol).Width = lbcView(0).ColumnHeaders.Item(ilCol).Width
            Next ilCol
        Next ilLoop
        lbcView(3).ColumnHeaders.Item(6).Width = 0
        lbcView(3).ColumnHeaders.Item(4).Width = lbcView(0).Width / 8
        lbcView(3).ColumnHeaders.Item(3).Width = lbcView(3).Width - lbcView(3).ColumnHeaders.Item(1).Width - lbcView(3).ColumnHeaders.Item(2).Width - lbcView(3).ColumnHeaders.Item(4).Width - lbcView(3).ColumnHeaders.Item(5).Width - 7 * GRIDSCROLLWIDTH
        '7967
        lbcView(4).ColumnHeaders.Item(1).Width = lbcView(4).Width / 8
        lbcView(4).ColumnHeaders.Item(2).Width = lbcView(4).Width / 7.75
        lbcView(4).ColumnHeaders.Item(4).Width = lbcView(4).Width / 14
        lbcView(4).ColumnHeaders.Item(5).Width = lbcView(4).Width / 14
        lbcView(4).ColumnHeaders.Item(6).Width = lbcView(4).Width / 18
        lbcView(4).ColumnHeaders.Item(7).Width = lbcView(4).Width / 5
        lbcView(4).ColumnHeaders.Item(3).Width = lbcView(4).Width - lbcView(4).ColumnHeaders.Item(1).Width - lbcView(4).ColumnHeaders.Item(2).Width - lbcView(4).ColumnHeaders.Item(4).Width - lbcView(4).ColumnHeaders.Item(5).Width - lbcView(4).ColumnHeaders.Item(6).Width - lbcView(4).ColumnHeaders.Item(7).Width - 7 * GRIDSCROLLWIDTH
    
        lbcView(5).ColumnHeaders.Item(1).Width = lbcView(0).Width / 8
        lbcView(5).ColumnHeaders.Item(2).Width = lbcView(0).Width / 7.75
        lbcView(5).ColumnHeaders.Item(3).Width = lbcView(0).Width / 8
        lbcView(5).ColumnHeaders.Item(4).Width = lbcView(5).Width - lbcView(5).ColumnHeaders.Item(1).Width - lbcView(5).ColumnHeaders.Item(2).Width - lbcView(5).ColumnHeaders.Item(3).Width - 6 * GRIDSCROLLWIDTH
    
        'moved from mInit
        rbcView(4).Visible = gAllowVendorAlerts(True)
        'rbcView(5).Visible = gPoolExist()
        If rbcView(4).Visible = False Then
            rbcView(5).Left = rbcView(4).Left
        End If
        If rbcView(5).Visible = True Then
            mRemoveInfo
        End If

    End If
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.25
    Me.Height = Screen.Height / 1.35
    Me.Top = (Screen.Height - Me.Height) / 1.5
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        mTerminate
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim llRet As Long
    imFirstActivate = True
    imTerminate = False
    Screen.MousePointer = vbHourglass
    imFirstFocus = True
    imRowSelected = -1
    'rbcView(4).Visible = gAllowVendorAlerts(True)
    'rbcView(5).Visible = gPoolExist()
    'If rbcView(4).Visible = False Then
    '    rbcView(5).Left = rbcView(4).Left
    'End If
    'If rbcView(5).Visible = True Then
    '    mRemoveInfo
    'End If
    mPopulate
    lbcView(1).Left = lbcView(0).Left
    lbcView(1).Top = lbcView(0).Top
    lbcView(1).Width = lbcView(0).Width
    lbcView(1).Height = lbcView(0).Height
    lbcView(2).Left = lbcView(0).Left
    lbcView(2).Top = lbcView(0).Top
    lbcView(2).Width = lbcView(0).Width
    lbcView(2).Height = lbcView(0).Height
    lbcView(3).Left = lbcView(0).Left
    lbcView(3).Top = lbcView(0).Top
    lbcView(3).Width = lbcView(0).Width
    lbcView(3).Height = lbcView(0).Height
    lbcView(4).Left = lbcView(0).Left
    lbcView(4).Top = lbcView(0).Top
    lbcView(4).Width = lbcView(0).Width
    lbcView(4).Height = lbcView(0).Height
    lbcView(5).Left = lbcView(0).Left
    lbcView(5).Top = lbcView(0).Top
    lbcView(5).Width = lbcView(0).Width
    lbcView(5).Height = lbcView(0).Height
    If lbcView(0).ListItems.Count > 0 Then
        rbcView(0).Value = True
    ElseIf lbcView(1).ListItems.Count > 0 Then
        rbcView(1).Value = True
    ElseIf lbcView(2).ListItems.Count > 0 Then
        rbcView(2).Value = True
    ElseIf lbcView(3).ListItems.Count > 0 Then
        rbcView(3).Value = True
    ElseIf lbcView(4).ListItems.Count > 0 Then
        rbcView(4).Value = True
        rbcView(4).Top = rbcView(3).Top
    ElseIf lbcView(5).ListItems.Count > 0 Then
        rbcView(5).Value = True
    Else
        rbcView(0).Value = True
    End If
'    '8129
'    rbcView(4).Visible = False
'    For ilRet = 0 To UBound(tgTaskInfo) Step 1
'        If Trim$(tgTaskInfo(ilRet).sTaskCode) = "WVM" Then
'            If tgTaskInfo(ilRet).iMenuIndex > 0 Then
'                rbcView(4).Visible = True
'            End If
'            Exit For
'        End If
'    Next ilRet
    '8273
    If imTerminate Then
        Exit Sub
    End If
'    gCenterModalForm frmAlertVw
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'
    Dim ilRet As Integer
    Screen.MousePointer = vbDefault
    Unload frmAlertVw
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmAufView
    Set frmAlertVw = Nothing   'Remove data segment
End Sub

Private Sub lbcView_Click(Index As Integer)
    Dim ilRet As Integer
    Dim ilRow As Integer
    
    If Index <> 3 Then
        Exit Sub
    End If
    If lbcView(Index).ListItems.Count <= 0 Then
        Exit Sub
    End If
    ilRet = 0
    On Error GoTo lbcViewErr
    ilRow = lbcView(Index).SelectedItem.Index
    If ilRet <> 0 Then
        Exit Sub
    End If
    On Error GoTo 0
    If ilRow >= 1 Then
        imRowSelected = ilRow
        cmcClear.Enabled = True
    Else
        imRowSelected = -1
        cmcClear.Enabled = False
    End If
    DoEvents
    Exit Sub
lbcViewErr:
    ilRet = 1
    Resume Next

End Sub

Private Sub lbcView_ColumnClick(Index As Integer, ByVal ColumnHeader As ComctlLib.ColumnHeader)
    lbcView(Index).SortKey = ColumnHeader.Index - 1
    lbcView(Index).Sorted = True
End Sub

Private Sub pbcClickFocus_GotFocus()
    Dim ilCode As Integer
    If imFirstFocus Then
        imFirstFocus = False
    End If
End Sub
Private Sub mPopulate()

    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim slTime As String
    Dim llTime As Long
    Dim ilVef As Integer
    Dim ilIndex As Integer
    Dim ilShowAuf As Integer
    Dim llDate As Long
    Dim slVehicleName As String
    Dim mItem As ListItem
    Dim tlAuf As AUF
    Dim llRet As Long
    '7967
    Dim slReason As String
    '8273
    Dim ilSkipVendor As Integer
    
    On Error GoTo ErrHand
    
    lbcView(0).ListItems.Clear
    lbcView(1).ListItems.Clear
    lbcView(2).ListItems.Clear
    lbcView(3).ListItems.Clear
    '7967
    lbcView(4).ListItems.Clear
    lbcView(5).ListItems.Clear
    llRet = SendMessageByNum(lbcView(3).hwnd, LV_SETEXTENDEDLISTVIEWSTYLE, 0, LV_FULLROWSSELECT)
    '8273 set to skip #4 if not visible.  66 is just a dummy number
    If gAllowVendorAlerts(True) Then
        ilSkipVendor = 66
    Else
        ilSkipVendor = 4
    End If
    '7/23/12: Bypass Traffic alerts
    For ilLoop = 0 To 5 Step 1
        'If rbcView(ilLoop).Visible Then
        If ilLoop <> 2 And ilLoop <> ilSkipVendor Then
            ReDim tmAufView(0 To 0) As AUFVIEW
            Select Case ilLoop
                Case 0              'Export Spots
                    SQLAlertQuery = "SELECT * FROM AUF_ALERT_USER WHERE (aufType = 'F' or aufType = 'R') AND aufSubType = 'S' AND aufStatus = 'R'"
                Case 1              'Expot ISCI
                    SQLAlertQuery = "SELECT * FROM AUF_ALERT_USER WHERE (aufType = 'F' or aufType = 'R') AND aufSubType = 'I' AND aufStatus = 'R'"
                Case 2              'Traffic Logs
                    SQLAlertQuery = "SELECT * FROM AUF_ALERT_USER WHERE aufType = 'L' AND (aufSubType = 'S' or aufSubType = 'C') AND aufStatus = 'R'"
                Case 3              'Agreement alter
                    SQLAlertQuery = "SELECT * FROM AUF_ALERT_USER WHERE aufType = 'P' AND aufSubType = 'A' AND aufStatus = 'R'"
                Case 4              '7967 web vendors
                    SQLAlertQuery = "SELECT * FROM AUF_ALERT_USER WHERE aufType = 'V' AND aufStatus = 'R'"  'AND aufSubType = 'E'
                Case 5              'Unfound Alert
                    SQLAlertQuery = "SELECT * FROM AUF_ALERT_USER WHERE aufType = 'U' AND aufStatus = 'R'"  'AND aufSubType = 'P'
            End Select
            Set rstAlert = gSQLSelectCall(SQLAlertQuery)
            Do While Not rstAlert.EOF
                tlAuf.sType = rstAlert!aufType
                tlAuf.sStatus = rstAlert!aufStatus
                tlAuf.sSubType = rstAlert!aufSubType
                tlAuf.iVefCode = rstAlert!aufVefCode
                tlAuf.lCode = rstAlert!aufCode
                '8133
                If ilLoop = 4 Then
                    tlAuf.iCountdown = rstAlert!aufcountdown
                    tlAuf.lUlfCode = rstAlert!aufulfcode
                End If
                If IsNull(rstAlert!aufMoWeekDate) Then
                    tlAuf.lMoWeekDate = 0
                ElseIf Not gIsDate(rstAlert!aufMoWeekDate) Then
                    tlAuf.lMoWeekDate = 0
                Else
                    tlAuf.lMoWeekDate = DateValue(gAdjYear(Format$(rstAlert!aufMoWeekDate, sgShowDateForm)))
                End If
                tlAuf.lEnteredDate = DateValue(gAdjYear(Format$(rstAlert!aufEnteredDate, sgShowDateForm)))
                tlAuf.lEnteredTime = gTimeToLong(Format$(rstAlert!aufEnteredTime, sgShowTimeWOSecForm), False)
                slDate = Trim$(Str$(tlAuf.lEnteredDate))
                Do While Len(slDate) < 5
                    slDate = "0" & slDate
                Loop
                slTime = Trim$(Str$(tlAuf.lEnteredTime))
                Do While Len(slTime) < 6
                    slTime = "0" & slTime
                Loop
                '7967 '8129
                If ilLoop = 4 Then
                    tlAuf.lChfCode = rstAlert!aufChfCode
                    tlAuf.lCefCode = rstAlert!aufcefcode
                End If
                tmAufView(UBound(tmAufView)).sKey = slDate & slTime
                tmAufView(UBound(tmAufView)).tAuf = tlAuf
                ReDim Preserve tmAufView(0 To UBound(tmAufView) + 1) As AUFVIEW
                rstAlert.MoveNext
            Loop
            
            If UBound(tmAufView) - 1 > 0 Then
                ArraySortTyp fnAV(tmAufView(), 0), UBound(tmAufView), 1, LenB(tmAufView(0)), 0, LenB(tmAufView(0).sKey), 0
            End If
            
            For ilIndex = 0 To UBound(tmAufView) - 1 Step 1
                'Test if status "R" is still valid.
                tlAuf = tmAufView(ilIndex).tAuf
                If tlAuf.sType = "F" Then
                    'If tlAuf.lMoWeekDate + 6 < DateValue(Format$(gNow(), "m/d/yy")) Then
                    If tlAuf.lMoWeekDate + 6 < DateValue(gAdjYear(Format$(gNow(), "m/d/yy"))) Then
                        ilShowAuf = False
                    Else
                        ilShowAuf = True
                    End If
                ElseIf tlAuf.sType = "R" Then
                    'If tlAuf.lMoWeekDate + 6 < DateValue(Format$(gNow(), "m/d/yy")) Then
                    If tlAuf.lMoWeekDate + 6 < DateValue(gAdjYear(Format$(gNow(), "m/d/yy"))) Then
                        ilShowAuf = False
                    Else
                        ilShowAuf = True
                    End If
                ElseIf tlAuf.sType = "L" Then
                    ilShowAuf = True
                ElseIf tlAuf.sType = "P" Then
                    ilShowAuf = True
                '7967
                ElseIf tlAuf.sType = "V" Then
                    If tlAuf.lMoWeekDate + 6 < DateValue(gAdjYear(Format$(gNow(), "m/d/yy"))) Then
                        ilShowAuf = False
                    Else
                        ilShowAuf = True
                    End If
                ElseIf tlAuf.sType = "U" Then
                    If tlAuf.lMoWeekDate + 6 < DateValue(gAdjYear(Format$(Now, "m/d/yy"))) Then
                        ilShowAuf = False
                    Else
                        ilShowAuf = True
                    End If
                Else
                    ilShowAuf = False
                End If
                If ilShowAuf Then
                    Select Case ilLoop
                        Case 0
                            Set mItem = lbcView(0).ListItems.Add()
                        Case 1
                            Set mItem = lbcView(1).ListItems.Add()
                        Case 2
                            Set mItem = lbcView(2).ListItems.Add()
                        Case 3
                            Set mItem = lbcView(3).ListItems.Add()
                        Case 4
                            Set mItem = lbcView(4).ListItems.Add()
                        Case 5
                            Set mItem = lbcView(5).ListItems.Add()
                    End Select
                    mItem.Text = Format$(tlAuf.lEnteredDate, sgShowDateForm)
                    mItem.SubItems(1) = Format$(gLongToTime(tlAuf.lEnteredTime), sgShowTimeWOSecForm)
                    For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
                        If tgVehicleInfo(ilVef).iCode = tlAuf.iVefCode Then
                            slVehicleName = tgVehicleInfo(ilVef).sVehicle
                            Exit For
                        End If
                    Next ilVef
                    If tlAuf.sType = "F" Then    'Export
                        mItem.SubItems(2) = slVehicleName
                        mItem.SubItems(3) = Format$(tlAuf.lMoWeekDate, sgShowDateForm)
                        If (ilLoop = 0) And (tlAuf.sSubType = "S") Then
                            mItem.SubItems(4) = "Final Log"
                        ElseIf (ilLoop = 1) And (tlAuf.sSubType = "I") Then
                            mItem.SubItems(4) = "Final Log"
                        End If
                    ElseIf tlAuf.sType = "R" Then    'Export
                        mItem.SubItems(2) = slVehicleName
                        mItem.SubItems(3) = Format$(tlAuf.lMoWeekDate, sgShowDateForm)
                        If (ilLoop = 0) And (tlAuf.sSubType = "S") Then
                            mItem.SubItems(4) = "Reprint Log"
                        ElseIf (ilLoop = 1) And (tlAuf.sSubType = "I") Then
                            mItem.SubItems(4) = "Reprint Log"
                        End If
                    ElseIf tlAuf.sType = "L" Then    'Log
                        mItem.SubItems(2) = slVehicleName
                        mItem.SubItems(3) = Format$(tlAuf.lMoWeekDate, sgShowDateForm)
                        If tlAuf.sSubType = "C" Then
                            mItem.SubItems(4) = "Copy Changed"
                        ElseIf tlAuf.sSubType = "S" Then
                            mItem.SubItems(4) = "Spot Changed"
                        ElseIf tlAuf.sSubType = "M" Then
                            mItem.SubItems(4) = "Missed Exist"
                        Else
                            mItem.SubItems(4) = ""
                        End If
                    ElseIf tlAuf.sType = "P" Then    'Agreement
                        mItem.SubItems(2) = slVehicleName
                        mItem.SubItems(3) = Format$(tlAuf.lMoWeekDate, sgShowDateForm)
                        If (tlAuf.sSubType = "A") Then
                            mItem.SubItems(4) = "Program Changed"
                        End If
                        mItem.SubItems(5) = Trim$(Str$(tlAuf.lCode))
                    '7967
                    ElseIf tlAuf.sType = "V" Then
                        '8133
                        '2 agreement 3 monday 4 vendor 5 export or import 6 reason
'                        If tlAuf.sSubType = "E" Then
'                            mItem.SubItems(5) = "Export"
'                            mItem.SubItems(4) = mVendorInfo(True, tlAuf.lChfCode, tlAuf.lCefCode, slVehicleName, slReason)
'                        Else
'                            mItem.SubItems(5) = "Import"
'                            mItem.SubItems(4) = mVendorInfo(False, tlAuf.lChfCode, tlAuf.lCefCode, slVehicleName, slReason)
'                        End If
'                        mItem.SubItems(2) = slVehicleName
'                        mItem.SubItems(6) = slReason
'                        mItem.SubItems(3) = Format$(tlAuf.lMoWeekDate, sgShowDateForm)
                        If tlAuf.sSubType = "E" Then
                            mItem.SubItems(5) = "Export"
                        Else
                            mItem.SubItems(5) = "Import"
                           ' mItem.SubItems(4) = mVendorInfo(False, tlAuf.lChfCode, tlAuf.lCefCode, slVehicleName, slReason)
                        End If
                        mItem.SubItems(4) = gVendorInitials(tlAuf.iVefCode) '  mVendorInfo(True, tlAuf.lChfCode, tlAuf.lCefCode)
                        mItem.SubItems(2) = mVendorStationVehicle(tlAuf.lUlfCode)
                        If tlAuf.lChfCode > 0 Then
                            mItem.SubItems(6) = gVendorIssue(True, tlAuf.iCountdown)
                        Else
                            mItem.SubItems(6) = gVendorWvmIssue(tlAuf.lCefCode)
                        End If
                        mItem.SubItems(3) = Format$(tlAuf.lMoWeekDate, sgShowDateForm)
                    ElseIf tlAuf.sType = "U" Then    'Info
                        If tlAuf.sSubType = "P" Then
                            mItem.SubItems(2) = "Pool"
                            mItem.SubItems(3) = "See PoolUnassignedLog_" & Format(tlAuf.lMoWeekDate, "mm-dd-yy") & ".txt" & " in Messages Subfolder"
                        End If
                        If tlAuf.sSubType = "C" Then
                            mItem.SubItems(2) = "Compel"
                            mItem.SubItems(3) = "See WegenerImportResult__" & Format(tlAuf.lMoWeekDate, "mm-dd-yy") & ".txt" & " in Messages Subfolder"
                        End If
                    End If
                Else
                    If (tlAuf.sType = "F") Or (tlAuf.sType = "R") Or (tlAuf.sType = "U") Then
                        SQLAlertQuery = "UPDATE AUF_ALERT_USER SET "
                        SQLAlertQuery = SQLAlertQuery & "aufStatus = 'C'" & ", "
                        SQLAlertQuery = SQLAlertQuery & "aufClearDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                        SQLAlertQuery = SQLAlertQuery & "aufClearTime = '" & Format$(gNow(), sgSQLTimeForm) & "', "
                        SQLAlertQuery = SQLAlertQuery & "aufClearUstCode = " & igUstCode & " "
                        SQLAlertQuery = SQLAlertQuery & "WHERE aufCode = " & tlAuf.lCode
                        'cnn.Execute SQLAlertQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLAlertQuery, False) <> 0 Then
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand:
                            gHandleError "AffErrorLog.txt", "AlertView-mPopulate"
                            Exit Sub
                        End If
                    End If
                End If
            Next ilIndex
        End If
    Next ilLoop
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "AlerVw-mPopulate"
End Sub

Private Sub rbcView_Click(Index As Integer)
    Dim ilRet As Integer
    Dim ilRow As Integer
    
    If rbcView(Index).Value Then
        lbcView(0).Visible = False
        lbcView(1).Visible = False
        lbcView(2).Visible = False
        lbcView(3).Visible = False
        lbcView(4).Visible = False
        lbcView(5).Visible = False
        lbcView(Index).Visible = True
        cmcClear.Enabled = False
        If Index = 3 Then
            If (lbcView(3).ListItems.Count > 0) And (imRowSelected > 0) Then
                cmcClear.Enabled = True
            End If
        End If
    End If
    Exit Sub
lbcViewErr:
    ilRet = 1
    Resume Next
End Sub
Private Function mVendorStationVehicle(llAttCode As Long) As String
    Dim slAgreementInfo As String
    Dim slSql As String
    
    slAgreementInfo = ""
    If llAttCode = 0 Then
        slAgreementInfo = ""
    Else
        slSql = "Select shttcallletters, vefName from att inner join VEF_Vehicles on attvefcode = vefcode inner join shtt on attshfcode = shttcode where attcode = " & llAttCode
        Set rst = gSQLSelectCall(slSql)
        If Not rst.EOF Then
            slAgreementInfo = Trim$(rst!shttCallLetters) & "\" & Trim$(rst!vefName)
        End If
    End If
    mVendorStationVehicle = slAgreementInfo
End Function

Private Sub mRemoveInfo()
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand
    gRemoveFiles sgMsgDirectory, "PoolUnassignedLog_", 30
    gRemoveFiles sgMsgDirectory, "WegenerImportResult_", 30
    slSQLQuery = "DELETE FROM AUF_ALERT_USER WHERE aufType = 'U' AND aufMoWeekDate <= '" & Format$(DateAdd("d", -30, Now), sgSQLDateForm) & "'"
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        gHandleError "AffErrorLog.txt", "frmAlertVw-mRemoveInfo"
        Exit Sub
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "AlerVw-mRemoveInfo"
End Sub
