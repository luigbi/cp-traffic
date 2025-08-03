VERSION 5.00
Begin VB.Form PDFEMailPersonnel 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2190
   ClientLeft      =   1620
   ClientTop       =   3480
   ClientWidth     =   8175
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2190
   ScaleWidth      =   8175
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   6
      Left            =   7095
      Top             =   1575
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1770
      Width           =   75
   End
   Begin VB.TextBox edcDropdown 
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
      HelpContextID   =   8
      Left            =   375
      MaxLength       =   30
      TabIndex        =   4
      Top             =   585
      Visible         =   0   'False
      Width           =   1950
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
      Height          =   90
      Left            =   15
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   210
      Width           =   60
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
      Height          =   90
      Left            =   90
      ScaleHeight     =   90
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   1755
      Width           =   75
   End
   Begin VB.PictureBox pbcPDFEMailPersonnel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   1005
      Left            =   330
      Picture         =   "PDFEMailPersonnel.frx":0000
      ScaleHeight     =   1005
      ScaleWidth      =   7470
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   405
      Width           =   7470
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   7725
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -30
      Width           =   7725
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2775
      TabIndex        =   6
      Top             =   1740
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4545
      TabIndex        =   7
      Top             =   1740
      Width           =   945
   End
   Begin VB.PictureBox plcPDFEMailPersonnel 
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
      Height          =   1095
      Left            =   285
      ScaleHeight     =   1035
      ScaleWidth      =   7500
      TabIndex        =   2
      Top             =   360
      Width           =   7560
   End
End
Attribute VB_Name = "PDFEMailPersonnel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of PDFEMailPersonnel.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PDFEMailPersonnel.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Advertiser Short Title input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim tmCtrls(0 To 3)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer
Dim imRowNo As Integer
Dim smShow() As String    'Show values
Dim smSave() As String
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imUpdateAllowed As Integer    'User can update records
Dim smName As String

Const NAMEINDEX = 1     'Name control/field
Const PHONENUMBERINDEX = 2
Const EMAILINDEX = 3

Private Sub cmcCancel_Click()
    igArfCallSource = CALLCANCELLED
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus ActiveControl
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcDone_Click()
    Dim ilRet As Integer
    
    mMoveCtrlToRec
    igArfCallSource = CALLDONE
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    Dim ilLoop As Integer
    If imBoxNo = -1 Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    gCtrlGotFocus ActiveControl
End Sub


Private Sub edcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    'Dim ilKey As Integer
    'ilKey = KeyAscii
    'If Not gCheckKeyAscii(ilKey) Then
    '    KeyAscii = 0
    '    Exit Sub
    'End If
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(ADVERTISERSLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcPDFEMailPersonnel.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcPDFEMailPersonnel.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    Me.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        tmcTerminate.Enabled = True
        Exit Sub
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear each control on the      *
'*                      screen                         *
'*                                                     *
'*******************************************************
Private Sub mClearCtrlFields()
'
'   mClearCtrlFields
'   Where:
'
    Dim ilLoop As Integer

    edcDropdown.Text = ""
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcDropdown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropdown.MaxLength = 40
            edcDropdown.Text = smSave(ilBoxNo - 1, imRowNo)
            gMoveTableCtrl pbcPDFEMailPersonnel, edcDropdown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo) * (fgBoxGridH + 15)
            edcDropdown.Visible = True  'Set visibility
            edcDropdown.SetFocus
        Case PHONENUMBERINDEX 'Name
            edcDropdown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropdown.MaxLength = 20
            edcDropdown.Text = smSave(ilBoxNo - 1, imRowNo)
            gMoveTableCtrl pbcPDFEMailPersonnel, edcDropdown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo) * (fgBoxGridH + 15)
            edcDropdown.Visible = True  'Set visibility
            edcDropdown.SetFocus
        Case EMAILINDEX 'Name
            edcDropdown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropdown.MaxLength = 80
            edcDropdown.Text = smSave(ilBoxNo - 1, imRowNo)
            gMoveTableCtrl pbcPDFEMailPersonnel, edcDropdown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo) * (fgBoxGridH + 15)
            edcDropdown.Visible = True  'Set visibility
            edcDropdown.SetFocus
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
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
    Dim ilRow As Integer
    Dim slStr As String
    
    imLBCtrls = 1
    imTerminate = False
    imFirstActivate = True
    Screen.MousePointer = vbHourglass
    ReDim smShow(0 To 3, 0 To 3) As String
    ReDim smSave(0 To 3, 0 To 3) As String
    mInitBox
    gCenterStdAlone PDFEMailPersonnel
    Screen.MousePointer = vbHourglass
    imBoxNo = -1 'Initialize current Box to N/A
    imChgMode = False
    imBSMode = False
    ilRet = gParseItem(sgCommandStr, 1, "\", slStr)
    smName = sgCommandStr
    mMoveRecToCtrl
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
    flTextHeight = pbcPDFEMailPersonnel.TextHeight("1") - 35
    'Position panel and picture areas with panel
    'Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 225, 2160, fgBoxStH
    tmCtrls(NAMEINDEX).iReq = False
    gSetCtrl tmCtrls(PHONENUMBERINDEX), 2205, tmCtrls(NAMEINDEX).fBoxY, 1620, fgBoxStH
    tmCtrls(PHONENUMBERINDEX).iReq = False
    gSetCtrl tmCtrls(EMAILINDEX), 3840, tmCtrls(NAMEINDEX).fBoxY, 3600, fgBoxStH
    tmCtrls(EMAILINDEX).iReq = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:5/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec()
'
'   mMoveCtrlToRec iTest
'   Where:
'       iTest (I)- True = only move if field changed
'                  False = move regardless of change state
'
    Dim ilRow As Integer
    Dim ilAdd As Integer
    
    For ilRow = LBound(tgPdf) To UBound(tgPdf) Step 1
        tgPdf(ilRow).lCode = 0
        tgPdf(ilRow).sName = ""
        tgPdf(ilRow).sPhone = ""
        tgPdf(ilRow).sEMailAddress = ""
    Next ilRow
    ilAdd = 0
    For ilRow = 0 To 3 Step 1
        If Trim$(smSave(NAMEINDEX - 1, ilRow)) <> "" Then
            tgPdf(ilAdd).sName = smSave(NAMEINDEX - 1, ilRow)
            tgPdf(ilAdd).sPhone = smSave(PHONENUMBERINDEX - 1, ilRow)
            tgPdf(ilAdd).sEMailAddress = smSave(EMAILINDEX - 1, ilRow)
            ilAdd = ilAdd + 1
            If ilAdd >= UBound(tgPdf) Then
                ReDim Preserve tgPdf(0 To ilAdd) As PDF
            End If
        End If
    Next ilRow
    ReDim Preserve tgPdf(0 To ilAdd) As PDF
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
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
    Dim ilRow As Integer
    
    pbcPDFEMailPersonnel.Cls
    For ilRow = 0 To 3 Step 1
        smSave(NAMEINDEX, ilRow) = ""
        smSave(PHONENUMBERINDEX, ilRow) = ""
        smSave(EMAILINDEX, ilRow) = ""
        smShow(NAMEINDEX, ilRow) = ""
        smShow(PHONENUMBERINDEX, ilRow) = ""
        smShow(EMAILINDEX, ilRow) = ""
    Next ilRow
    edcDropdown.MaxLength = 0
    For ilRow = LBound(tgPdf) To UBound(tgPdf) - 1 Step 1
        imRowNo = ilRow
        edcDropdown.Text = Trim$(tgPdf(ilRow).sName)
        mSetShow NAMEINDEX
        edcDropdown.Text = Trim$(tgPdf(ilRow).sPhone)
        mSetShow PHONENUMBERINDEX
        edcDropdown.Text = Trim$(tgPdf(ilRow).sEMailAddress)
        mSetShow EMAILINDEX
    Next ilRow
    pbcPDFEMailPersonnel_Paint
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mSetChg                         *
'*                                                     *
'*             Created:5/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if value for a       *
'*                      control is different from the  *
'*                      record                         *
'*                                                     *
'*******************************************************
Private Sub mSetChg(ilBoxNo As Integer)
'
'   mSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            gSetChgFlag sgAVIndicatorID, edcDropdown, tmCtrls(ilBoxNo)
        Case PHONENUMBERINDEX 'Name
            gSetChgFlag sgXDSCue, edcDropdown, tmCtrls(ilBoxNo)
        Case EMAILINDEX 'Name
            gSetChgFlag sgXDSCue, edcDropdown, tmCtrls(ilBoxNo)
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:4/28/93       By:D. LeVine      *
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
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
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

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcDropdown.Visible = False  'Set visibility
            smSave(ilBoxNo - 1, imRowNo) = edcDropdown.Text
            gSetShow pbcPDFEMailPersonnel, smSave(ilBoxNo - 1, imRowNo), tmCtrls(ilBoxNo)
            smShow(ilBoxNo - 1, imRowNo) = tmCtrls(ilBoxNo).sShow
        Case PHONENUMBERINDEX 'Name
            edcDropdown.Visible = False  'Set visibility
            smSave(ilBoxNo - 1, imRowNo) = edcDropdown.Text
            gSetShow pbcPDFEMailPersonnel, smSave(ilBoxNo - 1, imRowNo), tmCtrls(ilBoxNo)
            smShow(ilBoxNo - 1, imRowNo) = tmCtrls(ilBoxNo).sShow
        Case EMAILINDEX 'Name
            edcDropdown.Visible = False  'Set visibility
            smSave(ilBoxNo - 1, imRowNo) = edcDropdown.Text
            gSetShow pbcPDFEMailPersonnel, smSave(ilBoxNo - 1, imRowNo), tmCtrls(ilBoxNo)
            smShow(ilBoxNo - 1, imRowNo) = tmCtrls(ilBoxNo).sShow
    End Select
    pbcPDFEMailPersonnel_Paint
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
    sgDoneMsg = Trim$(str$(igArfCallSource)) & "\" & sgArfName
    igManUnload = YES
    'Unload Traffic
    Unload PDFEMailPersonnel
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:4/21/93       By:D. LeVine      *
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
    Dim ilRow As Integer
    
    For ilRow = 0 To 3 Step 1
        If Trim$(smSave(NAMEINDEX, ilRow)) <> "" Then
            If (ilCtrlNo = NAMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
                If gFieldDefinedStr(smSave(NAMEINDEX, ilRow), "", "Name must be specified", tmCtrls(NAMEINDEX).iReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = NAMEINDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
            End If
            If (ilCtrlNo = PHONENUMBERINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
                If gFieldDefinedStr(smSave(PHONENUMBERINDEX, ilRow), "", "Phone Number must be specified", tmCtrls(PHONENUMBERINDEX).iReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = PHONENUMBERINDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
            End If
            If (ilCtrlNo = EMAILINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
                If gFieldDefinedStr(smSave(EMAILINDEX, ilRow), "", "E-Mail Address must be specified", tmCtrls(EMAILINDEX).iReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = EMAILINDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
            End If
        End If
    Next ilRow
    mTestFields = YES
End Function
Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcPDFEMailPersonnel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    

    For ilRow = 0 To 3 Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow
                    mSetShow imBoxNo
                    imRowNo = ilRowNo
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow


End Sub
Private Sub pbcPDFEMailPersonnel_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    For ilRow = 0 To 3 Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            pbcPDFEMailPersonnel.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcPDFEMailPersonnel.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            pbcPDFEMailPersonnel.Print Trim$(smShow(ilBox - 1, ilRow))
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilSvRowNo As Integer
    
    'Handle double names- if drop down selected the index is changed to the
    'first name without any events- forces back so change occurs
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    ilSvRowNo = imRowNo
    Select Case imBoxNo
        Case -1
            ilBox = NAMEINDEX
            imRowNo = 0
            mSetCommands
        Case NAMEINDEX 'Name
            mSetShow imBoxNo
            If ilSvRowNo = 0 Then
                pbcClickFocus.SetFocus
            Else
                imRowNo = ilSvRowNo - 1
                imBoxNo = EMAILINDEX
                mEnableBox imBoxNo
            End If
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
    Dim ilSvRowNo As Integer
    
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    ilSvRowNo = imRowNo
    Select Case imBoxNo
        Case -1
            ilBox = NAMEINDEX
        Case EMAILINDEX
            mSetShow imBoxNo
            If ilSvRowNo = 3 Then
                imBoxNo = -1
                cmcDone.SetFocus
            Else
                imRowNo = ilSvRowNo + 1
                imBoxNo = NAMEINDEX
                mEnableBox imBoxNo
            End If
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
Private Sub plcPDFEMailPersonnel_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "PDF E-Mail Invoice Personnel: " & smName
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload PDFEMailPersonnel
End Sub


