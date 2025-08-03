VERSION 5.00
Begin VB.Form BudAdd 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4410
   ClientLeft      =   360
   ClientTop       =   2280
   ClientWidth     =   3930
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
   ScaleHeight     =   4410
   ScaleWidth      =   3930
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   4020
      Width           =   945
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
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   60
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5000
      Width           =   60
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   765
      TabIndex        =   3
      Top             =   4020
      Width           =   945
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   75
      ScaleHeight     =   240
      ScaleWidth      =   3030
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   3030
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
      Height          =   3510
      Left            =   150
      ScaleHeight     =   3450
      ScaleWidth      =   3450
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   3510
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         Height          =   3390
         Left            =   30
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   30
         Width           =   3390
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   255
      Top             =   4005
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "BudAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Budadd.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: BudAdd.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Trend input screen code
Option Explicit
Option Compare Text
'Vehicle
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer  'VEF record length
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer


Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
    Dim ilLoop As Integer
    Dim ilSaleOffice As Integer
    Dim ilVeh As Integer
    Dim ilUpper As Integer
    Dim slNameCode As String
    Dim slVehName As String
    Dim ilRet As Integer
    Dim slOffName As String
    Dim slOffState As String
    Dim slMktRank As String
    Dim slCode As String
    Dim ilOffCode As Integer
    Dim ilVefCode As Integer
    Dim slVehSort As String
    Dim slStr As String
    Dim ilFound As Integer
    igBDReturn = 1
    'Add Vehicles
    Screen.MousePointer = vbHourglass
    ilUpper = UBound(tgBvfRec)
    For ilVeh = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilVeh) Then
            slStr = lbcVehicle.List(ilVeh)
            ilFound = False
            For ilLoop = LBound(tgBudUserVehicle) To UBound(tgBudUserVehicle) - 1 Step 1
                slNameCode = tgBudUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                ilRet = gParseItem(slNameCode, 1, "\", slVehName)
                ilRet = gParseItem(slVehName, 3, "|", slVehName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                If slVehName = slStr Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If ilFound Then
                For ilSaleOffice = LBound(tgSalesOfficeCode) To UBound(tgSalesOfficeCode) - 1 Step 1
                    slNameCode = tgSalesOfficeCode(ilSaleOffice).sKey    'lbcSalesOfficeCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 1, "\", slStr)   'tmSaleOffice(ilLoop + 1).sName)
                    ilRet = gParseItem(slStr, 2, "|", slOffName)
                    ilRet = gParseItem(slStr, 3, "|", slOffState)
                    ilRet = gParseItem(slStr, 1, "|", slMktRank)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilOffCode = Val(slCode)
                    If slOffState <> "D" Then
                        tgBvfRec(ilUpper).tBvf.iSofCode = ilOffCode
                        tgBvfRec(ilUpper).tBvf.iVefCode = ilVefCode
                        
                        
'                        tgBvfRec(ilUpper).tBvf.iYear = tgBvfRec(LBound(tgBvfRec)).tBvf.iYear    'tgBvfRec(LBound(tgBvfRec)).tBvf.iYear
'                        tgBvfRec(ilUpper).tBvf.iSeqNo = tgBvfRec(LBound(tgBvfRec)).tBvf.iSeqNo
'                        tgBvfRec(ilUpper).tBvf.iStartDate(0) = tgBvfRec(LBound(tgBvfRec)).tBvf.iStartDate(0)
'                        tgBvfRec(ilUpper).tBvf.iStartDate(1) = tgBvfRec(LBound(tgBvfRec)).tBvf.iStartDate(1)
'                        tgBvfRec(ilUpper).tBvf.iMnfBudget = tgBvfRec(LBound(tgBvfRec)).tBvf.iMnfBudget    'tgBvfRec(LBound(tgBvfRec)).tBvf.iYear
                        tgBvfRec(ilUpper).tBvf.iYear = tgBvfRec(LBound(tgBvfRec) + igLBBvfRec).tBvf.iYear   'tgBvfRec(LBound(tgBvfRec)).tBvf.iYear
                        tgBvfRec(ilUpper).tBvf.iSeqNo = tgBvfRec(LBound(tgBvfRec) + igLBBvfRec).tBvf.iSeqNo
                        tgBvfRec(ilUpper).tBvf.iStartDate(0) = tgBvfRec(LBound(tgBvfRec) + igLBBvfRec).tBvf.iStartDate(0)
                        tgBvfRec(ilUpper).tBvf.iStartDate(1) = tgBvfRec(LBound(tgBvfRec) + igLBBvfRec).tBvf.iStartDate(1)
                        tgBvfRec(ilUpper).tBvf.iMnfBudget = tgBvfRec(LBound(tgBvfRec) + igLBBvfRec).tBvf.iMnfBudget   'tgBvfRec(LBound(tgBvfRec)).tBvf.iYear
                        
                        'tgBvfRec(ilUpper).tBvf.sSplit = tgBvfRec(LBound(tgBvfRec)).tBvf.sSplit
                        tgBvfRec(ilUpper).tBvf.sSplit = tgBvfRec(LBound(tgBvfRec) + igLBBvfRec).tBvf.sSplit
                        
                        For ilLoop = LBound(tgBvfRec(ilUpper).tBvf.lGross) To UBound(tgBvfRec(ilUpper).tBvf.lGross) Step 1
                            tgBvfRec(ilUpper).tBvf.lGross(ilLoop) = 0
                        Next ilLoop
                        
                        
                        If tmVef.iCode <> ilVefCode Then
                            tmVefSrchKey.iCode = ilVefCode
                            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            slVehSort = ""
                            slVehSort = Trim$(Str$(tmVef.iSort))
                            Do While Len(slVehSort) < 5
                                slVehSort = "0" & slVehSort
                            Loop
                        End If
                        tgBvfRec(ilUpper).sVehSort = slVehSort
                        tgBvfRec(ilUpper).sVehicle = slVehName
                        tgBvfRec(ilUpper).SOffice = slOffName
                        tgBvfRec(ilUpper).sMktRank = slMktRank
                        If Budget!rbcSort(1).Value Or Budget!rbcSort(2).Value Then    'Vehicle within office
                            tgBvfRec(ilUpper).sKey = tgBvfRec(ilUpper).sMktRank & tgBvfRec(ilUpper).SOffice & tgBvfRec(ilUpper).sVehSort & tgBvfRec(ilUpper).sVehicle
                        ElseIf Budget!rbcSort(0).Value Or Budget!rbcSort(3).Value Then
                            tgBvfRec(ilUpper).sKey = tgBvfRec(ilUpper).sVehSort & tgBvfRec(ilUpper).sVehicle & tgBvfRec(ilUpper).sMktRank & tgBvfRec(ilUpper).SOffice
                        End If
                        tgBvfRec(ilUpper).iStatus = 0
                        ilUpper = ilUpper + 1
                        'ReDim Preserve tgBvfRec(1 To ilUpper) As BVFREC
                        ReDim Preserve tgBvfRec(0 To ilUpper) As BVFREC
                    End If
                Next ilSaleOffice
            End If
        End If
    Next ilVeh
    Screen.MousePointer = vbDefault
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus cmcDone
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        'gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    'If (igWinStatus(RATECARDSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
    '    lbcBudget.Enabled = False
    'Else
    '    lbcBudget.Enabled = True
    'End If
    If (igWinStatus(BUDGETSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    'gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    BudAdd.Refresh
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    btrDestroy hmVef
    Set BudAdd = Nothing   'Remove data segment

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
    Dim ilRet As Integer
    imTerminate = False
    imFirstActivate = True
    igBDReturn = 0

    Screen.MousePointer = vbHourglass
    igLBBvfRec = 1
    BudAdd.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone BudAdd
    hmVef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", BudAdd
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    imTabDirection = 0  'Left to right movement
    'BudAdd.Show
    Screen.MousePointer = vbHourglass
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    'plcScreen.Caption = "Add Vehicle to " & sgBAName
'    gCenterModalForm BudAdd
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
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
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilVeh As Integer
    Dim ilLoop As Integer
    Dim ilCode As Integer

    For ilVeh = LBound(tgBudUserVehicle) To UBound(tgBudUserVehicle) - 1 Step 1
        slNameCode = tgBudUserVehicle(ilVeh).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 3, "|", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilCode = Val(slCode)
        ilFound = False
        'For ilLoop = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
        For ilLoop = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
            If (tgBvfRec(ilLoop).tBvf.iVefCode = ilCode) Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If (Not ilFound) Then
            'Test if vehicle has LCF within year
            tmVefSrchKey.iCode = ilCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                If tmVef.sState <> "D" Then
                    lbcVehicle.AddItem Trim$(tmVef.sName)
                End If
            End If
        End If
    Next ilVeh
    Exit Sub

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
    Unload BudAdd
    igManUnload = NO
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Add Vehicle to " & sgBAName
End Sub
