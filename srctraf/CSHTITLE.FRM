VERSION 5.00
Begin VB.Form CShTitle 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3585
   ClientLeft      =   735
   ClientTop       =   1620
   ClientWidth     =   6000
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3585
   ScaleWidth      =   6000
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   180
      Left            =   105
      Picture         =   "CSHTITLE.frx":0000
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.CommandButton cmcSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   3825
      TabIndex        =   11
      Top             =   3150
      Width           =   1050
   End
   Begin VB.ListBox lbcShortTitle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2895
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1755
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      Left            =   615
      MaxLength       =   20
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2310
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmcDropDown 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1785
      Picture         =   "CSHTITLE.frx":030A
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2310
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   5445
      Top             =   2805
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
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4560
      Width           =   75
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3525
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   135
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.PictureBox pbcTab 
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
      Height          =   75
      Left            =   30
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   2880
      Width           =   75
   End
   Begin VB.PictureBox pbcSTab 
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
      Height          =   135
      Left            =   195
      ScaleHeight     =   135
      ScaleWidth      =   90
      TabIndex        =   2
      Top             =   405
      Width           =   90
   End
   Begin VB.VScrollBar vbcShortTitle 
      Height          =   2160
      LargeChange     =   9
      Left            =   5205
      Min             =   1
      TabIndex        =   8
      Top             =   510
      Value           =   1
      Width           =   240
   End
   Begin VB.PictureBox pbcShortTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   495
      Picture         =   "CSHTITLE.frx":0404
      ScaleHeight     =   2175
      ScaleWidth      =   4710
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   495
      Width           =   4710
   End
   Begin VB.PictureBox plcShortTitle 
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
      Height          =   2295
      Left            =   420
      ScaleHeight     =   2235
      ScaleWidth      =   5040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   450
      Width           =   5100
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2340
      TabIndex        =   10
      Top             =   3150
      Width           =   1050
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   15
      ScaleHeight     =   270
      ScaleWidth      =   1260
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1260
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   855
      TabIndex        =   9
      Top             =   3150
      Width           =   1050
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5505
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2265
      Visible         =   0   'False
      Width           =   435
   End
End
Attribute VB_Name = "CShTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CSHTITLE.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CShTitle.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract item/package billing input screen code
Option Explicit
Option Compare Text
'Short Title
Dim tmCtrls(0 To 2) As FIELDAREA
Dim imLBCtrls As Integer
Dim imLB1Or2 As Integer
Dim tmShtTitle() As CSHTTITLE
Dim hmSif As Integer
Dim tmSif As SIF            'SIF record image
Dim imSifRecLen As Integer     'SIF record length
Dim hmMnf As Integer
Dim tmMnf As MNF            'MNF record image
Dim tmMnfSrchKey As INTKEY0 'MNF key record image
Dim imMnfRecLen As Integer     'MNF record length
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBoxNo As Integer
Dim imRowNo As Integer
Dim smShow() As String    'Show values
Dim imChg As Integer
Dim imSettingValue As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imFirstActivate As Integer
Dim imBypassFocus As Integer
Dim imUpdateAllowed As Integer
Dim smShtTitle As String     'Short Title, saved to determine if changed
Dim tmShtTitleCode() As SORTCODE
Dim smShtTitleCodeTag As String

'Short Title
Const VEHICLEINDEX = 1   'Vehicle control/field
Const SHORTTITLEINDEX = 2 'Short Title control/field
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    pbcArrow.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_Click()
    If mSaveRecChg(True) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    pbcArrow.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case SHORTTITLEINDEX
            lbcShortTitle.Visible = Not lbcShortTitle.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcSave_Click()
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    pbcShortTitle.Cls
    mVehPop
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    mShortTitlePop igAdfCode
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    mInitShow
    pbcShortTitle_Paint
    mSetCommands
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmcSave_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    pbcArrow.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imBoxNo
        Case SHORTTITLEINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcShortTitle, imBSMode, slStr)
            If ilRet = 1 Then
                lbcShortTitle.ListIndex = 0
            End If
            smShtTitle = edcDropDown.Text
    End Select
    imLbcArrowSetting = False
End Sub
Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case SHORTTITLEINDEX
            If lbcShortTitle.ListCount = 1 Then
                lbcShortTitle.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcIBSTab.SetFocus
                'Else
                '    pbcIBTab.SetFocus
                'End If
                'Exit Sub
            End If
    End Select
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo
            Case SHORTTITLEINDEX
                gProcessArrowKey Shift, KeyCode, lbcShortTitle, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub
Private Sub edcDropDown_LostFocus()
    Select Case imHdBoxNo
        Case VEHICLEINDEX
        Case SHORTTITLEINDEX
            smShtTitle = edcDropDown.Text
    End Select
End Sub
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imBoxNo
            Case SHORTTITLEINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
        End Select
        imDoubleClickName = False
    End If
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    Me.KeyPreview = True
    If sgCntrScreen = "Orders" Then
        If (igWinStatus(CONTRACTSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
            pbcShortTitle.Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            imUpdateAllowed = False
        Else
            pbcShortTitle.Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            imUpdateAllowed = True
        End If
    Else
        If (igWinStatus(PROPOSALSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
            pbcShortTitle.Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            imUpdateAllowed = False
        Else
            pbcShortTitle.Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            imUpdateAllowed = True
        End If
    End If
    gShowBranner imUpdateAllowed
    CShTitle.Refresh
    If imUpdateAllowed Then
        pbcSTab.SetFocus
    Else
        cmcDone.SetFocus
    End If
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
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    If Not igManUnload Then
        mSetShow imBoxNo
        imBoxNo = -1
        pbcArrow.Visible = False
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            If imBoxNo <> -1 Then
                mEnableBox imBoxNo
            End If
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    'Deallocate memory
    Erase tmShtTitle
    Erase smShow
    Erase tmShtTitleCode
    
    Set CShTitle = Nothing

End Sub
Private Sub lbcShortTitle_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcShortTitle, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcShortTitle_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcShortTitle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcShortTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcShortTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcShortTitle, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mIBEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    If (imRowNo < vbcShortTitle.Value) Or (imRowNo >= vbcShortTitle.Value + vbcShortTitle.LargeChange + 1) Then
        mSetShow ilBoxNo
        pbcArrow.Visible = False
        Exit Sub
    End If
    pbcArrow.Move plcShortTitle.Left - pbcArrow.Width - 15, plcShortTitle.Top + tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcShortTitle.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case VEHICLEINDEX 'Vehicle
        Case SHORTTITLEINDEX 'Short Title
            lbcShortTitle.Height = gListBoxHeight(lbcShortTitle.ListCount, 4)
            edcDropDown.Width = tmCtrls(SHORTTITLEINDEX).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 15
            gMoveTableCtrl pbcShortTitle, edcDropDown, tmCtrls(SHORTTITLEINDEX).fBoxX, tmCtrls(SHORTTITLEINDEX).fBoxY + (imRowNo - vbcShortTitle.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            smShtTitle = Trim$(tmShtTitle(imRowNo - 1).sShortTitle)
            gFindMatch smShtTitle, 0, lbcShortTitle
            imChgMode = True
            If gLastFound(lbcShortTitle) > 0 Then
                lbcShortTitle.ListIndex = gLastFound(lbcShortTitle)
            Else
                lbcShortTitle.ListIndex = 0
            End If
            If lbcShortTitle.ListIndex < 0 Then
                lbcShortTitle.ListIndex = 0
                edcDropDown.Text = smShtTitle
            Else
                edcDropDown.Text = lbcShortTitle.List(lbcShortTitle.ListIndex)
            End If
            imChgMode = False
            If imRowNo - vbcShortTitle.Value <= vbcShortTitle.LargeChange \ 2 Then
                lbcShortTitle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcShortTitle.Move edcDropDown.Left, edcDropDown.Top - lbcShortTitle.Height
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
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
    Dim ilRet As Integer    'Return Status
    ReDim smShow(0 To 2, 0 To 1) As String
    Screen.MousePointer = vbHourglass
    imLB1Or2 = 1
    imLBCtrls = 1
    imFirstActivate = True
    imTerminate = False
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    hmSif = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sif.Btr)", CShTitle
    On Error GoTo 0
    imSifRecLen = Len(tmSif) 'btrRecordLength(hmSif)    'Get Cff size
    hmMnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf.Btr)", CShTitle
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf) 'btrRecordLength(hmMnf)    'Get Cff size
    imSettingValue = True
    vbcShortTitle.Min = imLB1Or2    'LBound(smShow, 2)
    imSettingValue = True
    vbcShortTitle.Max = imLB1Or2    'LBound(smShow, 2)
    imSettingValue = True
    vbcShortTitle.Value = vbcShortTitle.Min
    imSettingValue = False
    imFirstActivate = True
    imBypassFocus = False
    imBoxNo = -1 'Initialize current Box to N/A
    imRowNo = -1
    imChg = False
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imSettingValue = False
    imChgMode = False
    imBSMode = False
    'gPDNToStr tgChfCntr.sInputGross, 2, smGross
    mVehPop
    If imTerminate Then
        Exit Sub
    End If
    mShortTitlePop igAdfCode
    If imTerminate Then
        Exit Sub
    End If
    CShTitle.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterModalForm CShTitle
    'Traffic!plcHelp.Cls
    mInitBox
    mInitShow
    pbcShortTitle_Paint
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    Dim flTextHeight As Single  'Standard text height
    flTextHeight = pbcShortTitle.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcShortTitle.Move 420, 450, pbcShortTitle.Width + vbcShortTitle.Width + fgPanelAdj - 15, pbcShortTitle.Height + fgPanelAdj
    pbcShortTitle.Move plcShortTitle.Left + fgBevelX, plcShortTitle.Top + fgBevelY
    vbcShortTitle.Move pbcShortTitle.Left + pbcShortTitle.Width, pbcShortTitle.Top + 15
    gSetCtrl tmCtrls(VEHICLEINDEX), 30, 225, 2325, fgBoxGridH
    'Short Title
    gSetCtrl tmCtrls(SHORTTITLEINDEX), 2370, tmCtrls(VEHICLEINDEX).fBoxY, 2325, fgBoxGridH
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitShow                       *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize show values         *
'*                                                     *
'*******************************************************
Private Sub mInitShow()
    Dim ilRowNo As Integer
    Dim ilBoxNo As Integer
    Dim slStr As String
    Dim ilSvRowNo As Integer
    Dim ilVef As Integer
    Dim ilLoop As Integer
    ilSvRowNo = imRowNo
    'ReDim smShow(1 To 2, 1 To UBound(tmShtTitle) + 1) As String
    ReDim smShow(0 To 2, 0 To UBound(tmShtTitle) + 1) As String
    For ilLoop = LBound(tmShtTitle) To UBound(tmShtTitle) - 1 Step 1
        ilRowNo = ilLoop + 1
        For ilBoxNo = VEHICLEINDEX To SHORTTITLEINDEX Step 1
            Select Case ilBoxNo 'Branch on box type (control)
                Case VEHICLEINDEX 'Vehicle
                    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If tmShtTitle(ilLoop).iVefCode = tgMVef(ilVef).iCode Then
                        ilVef = gBinarySearchVef(tmShtTitle(ilLoop).iVefCode)
                        If ilVef <> -1 Then
                            slStr = Trim$(tgMVef(ilVef).sName)
                            gSetShow pbcShortTitle, slStr, tmCtrls(VEHICLEINDEX)
                            smShow(VEHICLEINDEX, ilRowNo) = tmCtrls(VEHICLEINDEX).sShow
                    '        Exit For
                        End If
                    'Next ilVef
                Case SHORTTITLEINDEX 'Short Title
                    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If tmShtTitle(ilLoop).iVefCode = tgMVef(ilVef).iCode Then
                        ilVef = gBinarySearchVef(tmShtTitle(ilLoop).iVefCode)
                        If ilVef <> -1 Then
                            slStr = Trim$(tmShtTitle(ilLoop).sShortTitle)
                            gSetShow pbcShortTitle, slStr, tmCtrls(SHORTTITLEINDEX)
                            smShow(SHORTTITLEINDEX, ilRowNo) = tmCtrls(SHORTTITLEINDEX).sShow
                    '        Exit For
                        End If
                    'Next ilVef
            End Select
        Next ilBoxNo
    Next ilLoop
    imRowNo = ilSvRowNo
    imSettingValue = True
    vbcShortTitle.Min = imLB1Or2    'LBound(smShow, 2)
    imSettingValue = True
    If UBound(smShow, 2) - 2 <= vbcShortTitle.LargeChange Then
        vbcShortTitle.Max = imLB1Or2    'LBound(smShow, 2)
    Else
        vbcShortTitle.Max = UBound(smShow, 2) - 1 - vbcShortTitle.LargeChange
    End If
    imSettingValue = True
    vbcShortTitle.Value = vbcShortTitle.Min
    imSettingValue = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update or added record          *
'*                                                     *
'*******************************************************
Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilLoop As Integer   'For loop control
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slSyncDate As String
    Dim slSyncTime As String
    Screen.MousePointer = vbHourglass
    mSetShow imBoxNo
    gGetSyncDateTime slSyncDate, slSyncTime
    For ilLoop = LBound(tmShtTitle) To UBound(tmShtTitle) - 1 Step 1
        If Trim$(tmShtTitle(ilLoop).sShortTitle) <> "" Then
            If tmShtTitle(ilLoop).lSifCode = 0 Then
                imSifRecLen = Len(tmSif)
                Do  'Loop until record updated or added
                    tmSif.lCode = 0
                    tmSif.iAdfCode = igAdfCode
                    tmSif.sName = Trim$(tmShtTitle(ilLoop).sShortTitle)
                    tmSif.sState = "A"
                    tmSif.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                    tmSif.iRemoteID = tgUrf(0).iRemoteUserID
                    tmSif.lAutoCode = tmSif.lCode
                    ilRet = btrInsert(hmSif, tmSif, imSifRecLen, INDEXKEY0)
                    slMsg = "mSaveRec (btrInsert:Short Title)"
                Loop While ilRet = BTRV_ERR_CONFLICT
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, Contract
                On Error GoTo 0
                tmShtTitle(ilLoop).lSifCode = tmSif.lCode
                Do
                    tmSif.iRemoteID = tgUrf(0).iRemoteUserID
                    tmSif.lAutoCode = tmSif.lCode
                    gPackDate slSyncDate, tmSif.iSyncDate(0), tmSif.iSyncDate(1)
                    gPackTime slSyncTime, tmSif.iSyncTime(0), tmSif.iSyncTime(1)
                    ilRet = btrUpdate(hmSif, tmSif, imSifRecLen)
                    slMsg = "mSaveRec (btrUpdate:Short Title)"
                Loop While ilRet = BTRV_ERR_CONFLICT
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, Contract
                On Error GoTo 0
            End If
        Else
            tmShtTitle(ilLoop).lSifCode = 0
        End If
    Next ilLoop
    ReDim tgCShtTitle(0 To UBound(tmShtTitle)) As CSHTTITLE
    For ilLoop = LBound(tmShtTitle) To UBound(tmShtTitle) - 1 Step 1
        tgCShtTitle(ilLoop) = tmShtTitle(ilLoop)
    Next ilLoop
    imChg = False
    mSaveRec = True
    tmHdCtrls(10).iChg = True
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRecChg                     *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if record altered and*
'*                      requires updating              *
'*                                                     *
'*******************************************************
Private Function mSaveRecChg(ilAsk As Integer) As Integer
'
'   iAsk = True
'   iRet = mSaveRecChg(iAsk)
'   Where:
'       iAsk (I)- True = Ask if changed records should be updated;
'                 False= Update record if required without asking user
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRes As Integer
    Dim slMess As String
    If imChg Then
        If ilAsk Then
            slMess = "Save Changes"
            ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
            If ilRes = vbCancel Then
                mSaveRecChg = False
                Exit Function
            End If
            If ilRes = vbYes Then
                ilRes = mSaveRec()
                mSaveRecChg = ilRes
                Exit Function
            End If
            If ilRes = vbNo Then
            End If
        Else
            ilRes = mSaveRec()
            mSaveRecChg = ilRes
            Exit Function
        End If
    End If
    mSaveRecChg = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set command buttons (enable or *
'*                      disabled)                      *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
'
'   mSetCommands
'   Where:
'
    'Update button set if all mandatory fields have data and any field altered
    'Revert button set if any field changed
    If imChg Then
        cmcSave.Enabled = True
    Else
        cmcSave.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    pbcArrow.Visible = False
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case VEHICLEINDEX 'Vehicle
        Case SHORTTITLEINDEX 'Short Title index
            lbcShortTitle.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = smShtTitle
            If slStr = "[From Header]" Then
                slStr = ""
            End If
            gSetShow pbcShortTitle, slStr, tmCtrls(ilBoxNo)
            smShow(SHORTTITLEINDEX, imRowNo) = tmCtrls(ilBoxNo).sShow
            If StrComp(slStr, Trim$(tmShtTitle(imRowNo - 1).sShortTitle), 1) <> 0 Then
                Screen.MousePointer = vbHourglass
                slStr = smShtTitle
                If slStr = "[From Header]" Then
                    slStr = ""
                End If
                tmShtTitle(imRowNo - 1).sShortTitle = slStr
                imChg = True
                mShortTitlePop igAdfCode
                Screen.MousePointer = vbDefault
            End If
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mShortTitleBranch               *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      advertiser Short title and     *
'*                      process communication back from*
'*                      advertiser short title         *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mShortTitleBranch() As Integer
'
'   ilRet = mShortTitleBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
'    If imSelectedIndex = 0 Then 'New selected
'        imDoubleClickName = False
'        mShtTitleBranch = False
'        Exit Function
'    End If
    If edcDropDown.Text = "[From Header]" Then
        imDoubleClickName = False
        mShortTitleBranch = False
        Exit Function
    End If
    If (Not imDoubleClickName) And (edcDropDown.Text <> "[New]") Then
        imDoubleClickName = False
        mShortTitleBranch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoExeWinRes(ADVTPRODEXE)) Then
    '    imDoubleClickName = False
    '    mShortTitleBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourglass  'Wait
    igAdvtProdCallSource = CALLSOURCECONTRACT
    sgAdvtProdName = Contract.lbcAdvt.List(Contract.lbcAdvt.ListIndex)
    If edcDropDown.Text = "[New]" Then
        sgAdvtProdName = sgAdvtProdName & "\" & " "
    Else
        sgAdvtProdName = sgAdvtProdName & "\" & Trim$(edcDropDown.Text)
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'Traffic!edcLinkSrceHelpMsg.Text = ""
    If igTestSystem Then
        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    Else
        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    End If
    'lgShellRet = Shell(sgExePath & "ShtTitle.Exe " & slStr, 1)
    'Traffic.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    ShtTitle.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgAdvtProdName)
    igAdvtProdCallSource = Val(sgAdvtProdName)
    ilParse = gParseItem(slStr, 2, "\", sgAdvtProdName)
    'Traffic.Enabled = True
    'Traffic!edcLinkSrceHelpMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mShortTitleBranch = True
    imUpdateAllowed = ilUpdateAllowed
    gShowBranner imUpdateAllowed
    If igAdvtProdCallSource = CALLDONE Then  'Done
        igAdvtProdCallSource = CALLNONE
'        gSetMenuState True
        lbcShortTitle.Clear
        'lbcShtTitleCode.Tag = ""
        smShtTitleCodeTag = ""
        mShortTitlePop igAdfCode
        If imTerminate Then
            mShortTitleBranch = False
            Exit Function
        End If
        gFindMatch sgAdvtProdName, 1, lbcShortTitle
        If gLastFound(lbcShortTitle) > 0 Then
            imChgMode = True
            lbcShortTitle.ListIndex = gLastFound(lbcShortTitle)
            edcDropDown.Text = lbcShortTitle.List(lbcShortTitle.ListIndex)
            imChgMode = False
            mShortTitleBranch = False
        Else
            imChgMode = True
            lbcShortTitle.ListIndex = -1
            edcDropDown.Text = sgAdvtProdName
            imChgMode = False
            edcDropDown.SetFocus
            sgAdvtProdName = ""
            Exit Function
        End If
        sgAdvtProdName = ""
    End If
    If igAdvtProdCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igAdvtProdCallSource = CALLNONE
        sgAdvtProdName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igAdvtProdCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igAdvtProdCallSource = CALLNONE
        sgAdvtProdName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mShtTitlePop                    *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate advertiser short title*
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mShortTitlePop(ilAdvtCode As Integer)
'
'   mShtTitlePop
'   Where:
'       ilAdvtCode (I)- Adsvertiser code value
'
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim llLen As Long
    ilIndex = lbcShortTitle.ListIndex
    If ilIndex > 0 Then
        slName = lbcShortTitle.List(ilIndex)
    End If
    lbcShortTitle.Clear
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopShortTitleBox(Contract, ilAdvtCode, lbcShtTitle, lbcShtTitleCode)
    ilRet = gPopShortTitleBox(CShTitle, ilAdvtCode, lbcShortTitle, tmShtTitleCode(), smShtTitleCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mShtTitlePopErr
        gCPErrorMsg ilRet, "mShtTitlePop (gIMoveListBox)", CShTitle
        On Error GoTo 0
        imChgMode = True
        For ilLoop = LBound(tmShtTitle) To UBound(tmShtTitle) - 1 Step 1
            If Trim(tmShtTitle(ilLoop).sShortTitle) <> "" Then
                gFindMatch Trim(tmShtTitle(ilLoop).sShortTitle), 1, lbcShortTitle
                If gLastFound(lbcShortTitle) < 0 Then
                    tmShtTitle(ilLoop).lSifCode = 0
                    tmShtTitleCode(UBound(tmShtTitleCode)).sKey = Trim$(tmShtTitle(ilLoop).sShortTitle) & "\0"
                    ReDim Preserve tmShtTitleCode(0 To UBound(tmShtTitleCode) + 1) As SORTCODE
                Else
                    slNameCode = tmShtTitleCode(gLastFound(lbcShortTitle)).sKey    'lbcMster.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    tmShtTitle(ilLoop).lSifCode = Val(slCode)
                End If
            End If
        Next ilLoop
        lbcShortTitle.Clear
        If UBound(tmShtTitleCode) - 1 > 0 Then
            ArraySortTyp fnAV(tmShtTitleCode(), 0), UBound(tmShtTitleCode), 0, LenB(tmShtTitleCode(0)), 0, LenB(tmShtTitleCode(0).sKey), 0
        End If
        llLen = 0
        For ilLoop = 0 To UBound(tmShtTitleCode) - 1 Step 1
            slNameCode = tmShtTitleCode(ilLoop).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            slName = Trim$(slName)
            If Not gOkAddStrToListBox(slName, llLen, True) Then
                Exit For
            End If
            lbcShortTitle.AddItem slName  'Add ID to list box
        Next ilLoop
        lbcShortTitle.AddItem "[From Header]", 0  'Force as first item on list
'            lbcShtTitle.AddItem "[New]", 0  'Force as first item on list
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcShortTitle
            If gLastFound(lbcShortTitle) > 0 Then
                lbcShortTitle.ListIndex = gLastFound(lbcShortTitle)
            Else
                lbcShortTitle.ListIndex = -1
            End If
        Else
            lbcShortTitle.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mShtTitlePopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
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
    Unload CShTitle
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the Vehicle box      *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilClf As Integer
    Dim slGpSort As String
    Dim slVehName As String
    Dim slVehSort As String
    Dim ilVef As Integer
    Dim ilFound As Integer
    Dim slLineType As String

    ReDim tmShtTitle(0 To 0) As CSHTTITLE
    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
        slLineType = mGetLineType(ilClf + 1)
        If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
            If ((tgClfCntr(ilClf).iStatus = 0) Or (tgClfCntr(ilClf).iStatus = 1)) And (Not tgClfCntr(ilClf).iCancel) Then
                ilFound = False
                For ilLoop = LBound(tmShtTitle) To UBound(tmShtTitle) - 1 Step 1
                    If tmShtTitle(ilLoop).iVefCode = tgClfCntr(ilClf).ClfRec.iVefCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    tmShtTitle(UBound(tmShtTitle)).sKey = ""
                    tmShtTitle(UBound(tmShtTitle)).iVefCode = tgClfCntr(ilClf).ClfRec.iVefCode
                    tmShtTitle(UBound(tmShtTitle)).lSifCode = 0
                    tmShtTitle(UBound(tmShtTitle)).sShortTitle = ""
                    For ilLoop = LBound(tgCShtTitle) To UBound(tgCShtTitle) - 1 Step 1
                        If tgCShtTitle(ilLoop).iVefCode = tgClfCntr(ilClf).ClfRec.iVefCode Then
                            tmShtTitle(UBound(tmShtTitle)).sKey = tgCShtTitle(ilLoop).sKey
                            tmShtTitle(UBound(tmShtTitle)).lSifCode = tgCShtTitle(ilLoop).lSifCode
                            tmShtTitle(UBound(tmShtTitle)).sShortTitle = tgCShtTitle(ilLoop).sShortTitle
                            Exit For
                        End If
                    Next ilLoop
                    ReDim Preserve tmShtTitle(0 To UBound(tmShtTitle) + 1) As CSHTTITLE
                End If
            End If
        End If
    Next ilClf
    For ilLoop = LBound(tmShtTitle) To UBound(tmShtTitle) - 1 Step 1
        If Trim$(tmShtTitle(ilLoop).sKey) = "" Then
            slVehSort = "999"
            slVehName = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
            'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            '    If tmShtTitle(ilLoop).iVefCode = tgMVef(ilVef).iCode Then
                ilVef = gBinarySearchVef(tmShtTitle(ilLoop).iVefCode)
                If ilVef <> -1 Then
                    slVehName = tgMVef(ilVef).sName
                    slVehSort = Trim$(str$(tgMVef(ilVef).iSort))
                    tmMnfSrchKey.iCode = tgMVef(ilVef).iOwnerMnfCode
                    If tmMnf.iCode <> tmMnfSrchKey.iCode Then
                        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            tmMnf.iGroupNo = 999
                            tmMnf.iCode = 0
                        End If
                    End If
                    slGpSort = Trim$(str$(tmMnf.iGroupNo))
            '        Exit For
                End If
            'Next ilVef
            Do While Len(slGpSort) < 3
                slGpSort = "0" & slGpSort
            Loop
            Do While Len(slVehSort) < 3
                slVehSort = "0" & slVehSort
            Loop
            tmShtTitle(ilLoop).sKey = slGpSort & slVehSort & slVehName
        End If
    Next ilLoop
    If UBound(tmShtTitle) > 1 Then
        ArraySortTyp fnAV(tmShtTitle(), 0), UBound(tmShtTitle), 0, LenB(tmShtTitle(0)), 0, LenB(tmShtTitle(0).sKey), 0
    End If
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    pbcArrow.Visible = False
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcShortTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    ilCompRow = vbcShortTitle.LargeChange + 1
    If UBound(smShow, 2) - 1 > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(smShow, 2) - 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcShortTitle.Value - 1
                    mSetShow imBoxNo
                    If (ilBox = VEHICLEINDEX) Or (ilRowNo = UBound(smShow, 2)) Then
                        Beep
                        pbcClickFocus.SetFocus
                        Exit Sub
                    End If
                    imRowNo = ilRowNo
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcShortTitle_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer

    ilStartRow = vbcShortTitle.Value  'Top location
    ilEndRow = vbcShortTitle.Value + vbcShortTitle.LargeChange
    If ilEndRow > UBound(smShow, 2) Then
        ilEndRow = UBound(smShow, 2) 'include blank row as it might have data
    End If
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            pbcShortTitle.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcShortTitle.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            pbcShortTitle.Print smShow(ilBox, ilRow)
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    If imBoxNo = SHORTTITLEINDEX Then
        If mShortTitleBranch() Then
            Exit Sub
        End If
    End If
    imTabDirection = -1  'Set-right to left
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            imRowNo = 1
            imSettingValue = True
            vbcShortTitle.Value = vbcShortTitle.Min
            imSettingValue = False
            If UBound(smShow, 2) <= vbcShortTitle.LargeChange + 1 Then 'was <=
                vbcShortTitle.Max = imLB1Or2    'LBound(smShow, 2)
            Else
                vbcShortTitle.Max = UBound(smShow, 2) - vbcShortTitle.LargeChange ' - 1
            End If
            ilBox = SHORTTITLEINDEX
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case VEHICLEINDEX 'Name (first control within header)
            ilBox = SHORTTITLEINDEX
        Case SHORTTITLEINDEX
            mSetShow imBoxNo
            ilBox = SHORTTITLEINDEX
            If imRowNo <= 1 Then
                imBoxNo = -1
                cmcDone.SetFocus
                Exit Sub
            End If
            imRowNo = imRowNo - 1
            If imRowNo < vbcShortTitle.Value Then
                imSettingValue = True
                vbcShortTitle.Value = vbcShortTitle.Value - 1
                imSettingValue = False
            End If
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    If imBoxNo = SHORTTITLEINDEX Then
        If mShortTitleBranch() Then
            Exit Sub
        End If
    End If
    imTabDirection = 0  'Set-Left to right
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            imRowNo = UBound(smShow, 2)
            imSettingValue = True
            If imRowNo <= vbcShortTitle.LargeChange + 1 Then
                vbcShortTitle.Value = vbcShortTitle.Min
            Else
                vbcShortTitle.Value = imRowNo - vbcShortTitle.LargeChange
            End If
            imSettingValue = False
            ilBox = SHORTTITLEINDEX
        Case VEHICLEINDEX
        Case SHORTTITLEINDEX
            mSetShow imBoxNo
            If (imRowNo >= UBound(smShow, 2) - 1) Then
                pbcShortTitle_Paint
                If cmcDone.Enabled Then
                    cmcDone.SetFocus
                Else
                    cmcCancel.SetFocus
                End If
                Exit Sub
            End If
            ilBox = SHORTTITLEINDEX
            imRowNo = imRowNo + 1
            If imRowNo > vbcShortTitle.Value + vbcShortTitle.LargeChange Then
                imSettingValue = True
                vbcShortTitle.Value = vbcShortTitle.Value + 1
                imSettingValue = False
            End If
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case Else
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcShortTitle_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imBoxNo
        Case SHORTTITLEINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcShortTitle, edcDropDown, imChgMode, imLbcArrowSetting
    End Select
End Sub
Private Sub vbcShortTitle_Change()
    If imSettingValue Then
        pbcShortTitle.Cls
        pbcShortTitle_Paint
        imSettingValue = False
    Else
        mSetShow imBoxNo
        pbcShortTitle.Cls
        pbcShortTitle_Paint
        mEnableBox imBoxNo
    End If
End Sub
Private Sub vbcShortTitle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Short Titles"
End Sub
