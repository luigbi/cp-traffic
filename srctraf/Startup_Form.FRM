VERSION 5.00
Begin VB.Form Startup_Form 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4260
   ClientLeft      =   360
   ClientTop       =   3600
   ClientWidth     =   4065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   Icon            =   "Startup_Form.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4260
   ScaleWidth      =   4065
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   60
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1425
      Width           =   60
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1275
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   15
      TabIndex        =   6
      Top             =   2925
      Width           =   15
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   1170
      Width           =   45
   End
   Begin VB.PictureBox pbcFormName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   2460
      Left            =   525
      Picture         =   "Startup_Form.frx":08CA
      ScaleHeight     =   2460
      ScaleWidth      =   2850
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   945
      Width           =   2850
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   300
      TabIndex        =   7
      Top             =   3780
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   3765
      Width           =   1050
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2565
      TabIndex        =   9
      Top             =   3765
      Width           =   1050
   End
   Begin VB.PictureBox plcBackgroundName 
      ForeColor       =   &H00000000&
      Height          =   2580
      Left            =   480
      ScaleHeight     =   2520
      ScaleWidth      =   2925
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   900
      Width           =   2985
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   495
      TabIndex        =   1
      Top             =   420
      Width           =   3000
   End
   Begin VB.Label plcScreen 
      Caption         =   "Client Form Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2325
   End
End
Attribute VB_Name = "Startup_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software®, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Startup_Form.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Import Rate Card input screen code
Option Explicit
Option Compare Text
Dim tmCtrls(0 To 12)  As FIELDAREA

Dim smSave(0 To 1) As String
Dim imSave(0 To 8) As Integer

Dim imBoxNo As Integer   'Current Avail Name Box
Dim tmFileName As ANF        'ANF record image
Dim tmFileNameSrchKey As INTKEY0    'ANF key record image
Dim imFileNameRecLen As Integer        'ANF record length
Dim hmFileName As Integer 'File name file handle


Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim lmNowDate As Long
Dim imBypassFocus As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imUserInput As Integer
Dim imLBCtrls As Integer
Dim imBypassSetting As Integer
Dim imUpdateAllowed As Integer

Const NAMEINDEX = 1     'Name control/field
Const SUSTAININDEX = 2  'Sustain control/field
Const SPONSORINDEX = 3  'Sponsorship control/field
Const SORTCODEINDEX = 4
Const RPTDEFAULTINDEX = 5
Const TRAFTOAFFINDEX = 6
Const ISCIEXPORTINDEX = 7
Const AUDIOEXPORTINDEX = 8
Const AUTOEXPORTINDEX = 9
Const FILLEDBREAKINDEX = 10
Const SPOTTYPEINDEX = 11


Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub Form_Activate()
    If (igWinStatus(AVAILNAMESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcFormName.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcFormName.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    Startup_Form.Refresh

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
    btrStopAppl
    End
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
    Dim llNoRec As Long 'Number of records
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDate As String
    imTerminate = False
    mParseCmmdLine
    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    imFirstFocus = True
    imBypassFocus = False
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    gCenterStdAlone Startup_Form
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
'*             Created:6/30/93       By:D. LeVine      *
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
    'Position panel and picture areas with panel
    
    plcBackgroundName.Move 795, 735, pbcFormName.Width + fgPanelAdj, pbcFormName.Height + fgPanelAdj
    
    pbcFormName.Move plcBackgroundName.Left + fgBevelX, plcBackgroundName.Top + fgBevelY
    'Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
    'Sustaining
    gSetCtrl tmCtrls(SUSTAININDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slHelpSystem As String
    Dim ilPos As Integer
    Dim slChar As String
    slCommand = Command$
    igStdAloneMode = True 'False  'Switch from/to stand alone mode
    sgCallAppName = ""
    slStr = "Guide"
    ilTestSystem = False
    ilPos = InStr(1, slCommand, "/S:Test", 1)
    If ilPos > 0 Then
        ilTestSystem = True
    End If
    ilPos = InStr(1, slCommand, "/S:Prod", 1)
    If ilPos > 0 Then
        ilTestSystem = False
    End If


    gInitGlobalVar

    
    ilPos = InStr(1, slCommand, "/UserInput", 1)
    If ilPos > 0 Then
        imUserInput = True
        igBkgdProg = 0
    Else
        imUserInput = False
        igBkgdProg = 22
        Me.Visible = False
    End If

    'If password required
    If imUserInput Then
        Signon.Show vbModal
        If igExitTraffic Then
            imTerminate = True
            Exit Sub
        End If

    End If
    gInitStdAlone Startup_Form, slStr, ilTestSystem
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
    igParentRestarted = False
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Startup_Form
    Set Startup_Form = Nothing   'Remove data segment
    igManUnload = NO
End Sub

Private Sub mClearCtrlFields()
'
'   mClearCtrlFields
'   Where:
'
    Dim ilLoop As Integer
    mMoveCtrlToRec False
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
End Sub

Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            'mSendHelpMsg "Enter avail name"
            edcDropdown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropdown.MaxLength = 20
            gMoveFormCtrl pbcFormName, edcDropdown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcDropdown.Visible = True  'Set visibility
            edcDropdown.SetFocus
        Case SUSTAININDEX 'Allow sustaining
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
End Sub

Private Sub mMoveCtrlToRec(ilTestChg As Integer)
'
'   mMoveCtrlToRec iTest
'   Where:
'       iTest (I)- True = only move if field changed
'                  False = move regardless of change state
'
    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
        'tmFileName.sName = smSave(0)
    End If
    If Not ilTestChg Or tmCtrls(SUSTAININDEX).iChg Then
        Select Case imSave(0)
            Case 0  'Yes
                'tmFileName.sSustain = "Y"
            Case 1  'No
                'tmFileName.sSustain = "N"
            Case Else
                'tmFileName.sSustain = ""
        End Select
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()
'
'   mMoveRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    smSave(0) = Trim$(tmFileName.sName)
    'Select Case tmFileName.sSustain
    '    Case "Y"
    '        imSave(0) = 0
    '    Case "N"
    '        imSave(0) = 1
    '    Case Else
    '        imSave(0) = -1
    'End Select
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mOKName                         *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test that name is unique        *
'*                                                     *
'*******************************************************
Private Function mOKName()
    Dim slStr As String
    If edcDropdown.Text <> "" Then    'Test name
        slStr = Trim$(edcDropdown.Text)
        gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                If Trim$(edcDropdown.Text) = cbcSelect.List(gLastFound(cbcSelect)) Then
                    Beep
                    MsgBox "Name already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcDropdown.Text = Trim$(tmFileName.sName) 'Reset text
                    mSetShow imBoxNo
                    mSetChg imBoxNo
                    imBoxNo = 1
                    mEnableBox imBoxNo
                    mOKName = False
                    Exit Function
                End If
            End If
        End If
    End If
    mOKName = True
End Function

Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilLp As Integer
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer

    'imPopReqd = False
    'If igANmCallSource = CALLSOURCEFEED Then
    '    ilfilter(0) = CHARFILTER
    '    slFilter(0) = "F"
    '    ilOffSet(0) = gFieldOffset("Anf", "AnfBookLocalFeed") '2
    'Else
    '    ilfilter(0) = NOFILTER
    '    slFilter(0) = ""
    '    ilOffSet(0) = 0
    'End If
    'ilRet = gIMoveListBox(AName, cbcSelect, tgNameCode(), sgNameCodeTag, "Anf.btr", gFieldOffset("Anf", "AnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    'If ilRet <> CP_MSG_NOPOPREQ Then
    '    On Error GoTo mPopulateErr
    '    gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", AName
    '    On Error GoTo 0
    '    cbcSelect.AddItem "[New]", 0  'Force as first item on list
    '    imPopReqd = True
    'End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub mSetChg(ilBoxNo As Integer)
'
'   mSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'
    Dim slStr As String
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            'gSetChgFlag tmAnf.sName, edcDropdown, tmCtrls(ilBoxNo)
        Case SUSTAININDEX 'Sustain
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim ilAltered As Integer
    If imBypassSetting Then
        Exit Sub
    End If
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (ilAltered = YES) And (imUpdateAllowed) Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    If Not ilAltered Then
        cbcSelect.Enabled = True
    Else
        cbcSelect.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocusx                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcDropdown.SetFocus
        Case SUSTAININDEX 'Allow sustaining
            'pbcYN.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    '2/4/16: Add filter to handle the case where the name has illegal characters and it was pasted into the field
    If (ilBoxNo = NAMEINDEX) Then
        slStr = gReplaceIllegalCharacters(edcDropdown.Text)
        edcDropdown.Text = slStr
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcDropdown.Visible = False  'Set visibility
            slStr = edcDropdown.Text
            smSave(0) = slStr
            gSetShow pbcFormName, slStr, tmCtrls(ilBoxNo)
        Case SUSTAININDEX 'Sustaining
            'pbcYN.Visible = False  'Set visibility
            'If imSustain = 0 Then
            '    slStr = "Yes"
            'ElseIf imSustain = 1 Then
            '    slStr = "No"
            'Else
            '    slStr = ""
            'End If
            'gSetShow pbcFormName, slStr, tmCtrls(ilBoxNo)
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mTestFields(iTest, iState)
'   Where:
'       iTest (I)- Test all controls or control number specified
'       iState (I)- Test one of the following:
'                  ALLBLANK=All fields blank
'                  ALLMANBLANK=All mandatory
'                    field blank
'                  ALLMANDEFINED=All mandatory
'                    fields have data
'                  Plus
'                  NOMSG=No error message shown
'                  SHOWMSG=show error message
'       iRet (O)- True if all mandatory fields blank, False if not all blank
'
'
    Dim slStr As String
    If (ilCtrlNo = NAMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smSave(0), "", "Name must be specified", tmCtrls(NAMEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = NAMEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SUSTAININDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imSave(0) = 0 Then
            slStr = "Yes"
        ElseIf imSave(1) = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "A Yes Or No must be specified for Allow Sustaining", tmCtrls(SUSTAININDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SUSTAININDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    mTestFields = YES
End Function
Private Sub pbcFormName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus imBoxNo
End Sub
Private Sub pbcFormName_Paint()
    Dim ilBox As Integer
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        pbcFormName.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcFormName.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcFormName.Print tmCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    Select Case imBoxNo
        Case -1
            If (imSelectedIndex = 0) And (cbcSelect.Text = "[New]") Then
                ilBox = 1
                mSetCommands
            Else
                mSetChg 1
                ilBox = 2
            End If
        Case 1 'Name (first control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
                Exit Sub
            End If
            ilBox = 1
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
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    Select Case imBoxNo
        Case -1
            ilBox = UBound(tmCtrls)
        Case UBound(tmCtrls) 'Suppress (last control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            If (cmcUpdate.Enabled) And (igANmCallSource = CALLNONE) Then
                cmcUpdate.SetFocus
            Else
                cmcDone.SetFocus
            End If
            Exit Sub
        Case Else
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub

Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slStamp As String   'Date/Time stamp for file
    Dim ilQAsked As Integer

    mSetShow imBoxNo
    If mTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        mSaveRec = False
        Exit Function
    End If
    edcDropdown.Text = smSave(0)
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    Do  'Loop until record updated or added
        mMoveCtrlToRec True
        'If (tgSpf.sSystemType = "R") And (imSSFFixReq) And (imSelectedIndex > 0) And (Not ilQAsked) Then
        If imSelectedIndex = 0 Then 'New selected
            tmFileName.iCode = 0  'Autoincrement
            ilRet = btrInsert(hmFileName, tmFileName, imFileNameRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert)"
        Else 'Old record-Update
            ilRet = btrUpdate(hmFileName, tmFileName, imFileNameRecLen)
            slMsg = "mSaveRec (btr(Update)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, Startup_Form
    On Error GoTo 0
    mSaveRec = True
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function
