VERSION 5.00
Begin VB.Form VehSel 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5055
   ClientLeft      =   2310
   ClientTop       =   1980
   ClientWidth     =   4740
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5055
   ScaleWidth      =   4740
   Begin VB.CheckBox ckcAll 
      Caption         =   "All"
      Height          =   210
      Left            =   165
      TabIndex        =   6
      Top             =   4095
      Width           =   1275
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2475
      TabIndex        =   5
      Top             =   4545
      Width           =   1260
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
      TabIndex        =   4
      Top             =   1770
      Width           =   75
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      ScaleHeight     =   240
      ScaleWidth      =   3990
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -30
      Width           =   3990
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Left            =   930
      TabIndex        =   3
      Top             =   4545
      Width           =   1260
   End
   Begin VB.PictureBox plcModel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3735
      Left            =   150
      ScaleHeight     =   3675
      ScaleWidth      =   4260
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   270
      Width           =   4320
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         Height          =   3600
         ItemData        =   "VehSel.frx":0000
         Left            =   30
         List            =   "VehSel.frx":0002
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   30
         Width           =   4200
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   90
      Top             =   4545
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "VehSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of VehSel.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: VehSel.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)

Public drf_rst As ADODB.Recordset

Private Sub ckcAll_Click()
    Dim Value As Integer
    Dim ilIndex As Integer
    Dim ilValue As Integer
    Dim llRg As Long
    Dim llRet As Long
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcVehicle.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcVehicle.hWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
End Sub

Private Sub cmcCancel_Click()
    igVehReturn = 0
    mTerminate
End Sub
Private Sub cmcDone_Click()
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim blRet As Boolean
    
    blRet = mSave()
    igVehReturn = 1
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    Me.KeyPreview = True
    VehSel.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    drf_rst.Close
    
    Set VehSel = Nothing   'Remove data segment

End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
    imTerminate = False
    imFirstActivate = True
    imAllClicked = False
    imSetAll = True


    Screen.MousePointer = vbHourglass
    'VehSel.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    If igVehSelType = 1 Then
        cmcDone.Caption = "Set Default"
    End If
    gCenterStdAlone VehSel
    'VehSel.Show
    Screen.MousePointer = vbHourglass
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slSQLQuery As String
    If igVehSelType = 1 Then
        slSQLQuery = "Select distinct vefName, vefCode from DRF_Demo_Rsrch_Data "
        slSQLQuery = slSQLQuery & " Left Outer Join vef_Vehicles On drfVefCode = vefCode"
        slSQLQuery = slSQLQuery & " where drfDnfCode = " & igVehSelCode
        slSQLQuery = slSQLQuery & " And drfVefCode <> 0"
        slSQLQuery = slSQLQuery & " And drfDemoDataType <> 'P'" & " Order by vefName"
        Set drf_rst = gSQLSelectCall(slSQLQuery)
        Do While Not drf_rst.EOF
            lbcVehicle.AddItem Trim$(drf_rst!VEFNAME)
            lbcVehicle.ItemData(lbcVehicle.NewIndex) = drf_rst!vefCode
            drf_rst.MoveNext
        Loop
        ckcAll.Value = vbChecked
    End If
    Exit Sub
mPopulateErr:
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
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload VehSel
    igManUnload = NO
End Sub

Private Sub lbcVehicle_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked  'False
        imSetAll = True
    End If
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    If igVehSelType = 1 Then
        plcScreen.Print "Set Vehicle Default"
    End If
End Sub

Private Function mSave() As Boolean
    Dim ilVef As Integer
    Dim llRet As Long
    Dim slSQLQuery As String
    Dim ilRet As Integer
    
    If igVehSelType = 1 Then
        ilRet = MsgBox("This will set the default Research Book for each selected vehicle, Ok to Continue", vbYesNo + vbQuestion, "Save")
        If ilRet = vbNo Then
            mSave = False
            Exit Function
        End If
        For ilVef = 0 To lbcVehicle.ListCount - 1 Step 1
            If lbcVehicle.Selected(ilVef) Then
                slSQLQuery = "UPDATE vef_Vehicles SET "
                slSQLQuery = slSQLQuery & "vefDnfCode = " & igVehSelCode
                slSQLQuery = slSQLQuery & " WHERE vefCode = " & lbcVehicle.ItemData(ilVef)
                llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            End If
        Next ilVef
    End If
    mSave = True
End Function
