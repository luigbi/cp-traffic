VERSION 5.00
Begin VB.Form BudModel 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3765
   ClientLeft      =   1260
   ClientTop       =   1590
   ClientWidth     =   6420
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
   ScaleHeight     =   3765
   ScaleWidth      =   6420
   Begin VB.PictureBox plcSpec 
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
      Height          =   585
      Index           =   1
      Left            =   3525
      ScaleHeight     =   525
      ScaleWidth      =   2430
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   300
      Width           =   2490
      Begin VB.TextBox edcYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1305
         MaxLength       =   4
         TabIndex        =   13
         Top             =   135
         Width           =   795
      End
      Begin VB.Label lacYear 
         Appearance      =   0  'Flat
         Caption         =   "Budget Year"
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
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   1140
      End
   End
   Begin VB.PictureBox plcSpec 
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
      Height          =   1725
      Index           =   0
      Left            =   3240
      ScaleHeight     =   1665
      ScaleWidth      =   2955
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   300
      Width           =   3015
      Begin VB.TextBox edcYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1245
         MaxLength       =   4
         TabIndex        =   10
         Top             =   1170
         Width           =   795
      End
      Begin VB.PictureBox plcBT 
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   2760
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   810
         Width           =   2760
         Begin VB.OptionButton rbcBT 
            Caption         =   "Direct"
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
            Height          =   195
            Index           =   0
            Left            =   1125
            TabIndex        =   7
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton rbcBT 
            Caption         =   "Split"
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
            Height          =   195
            Index           =   1
            Left            =   1980
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.ComboBox cbcMnfBudget 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         TabIndex        =   5
         Top             =   330
         Width           =   2760
      End
      Begin VB.Label lacYear 
         Appearance      =   0  'Flat
         Caption         =   "Budget Year"
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
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1140
      End
      Begin VB.Label lacBudgetName 
         Appearance      =   0  'Flat
         Caption         =   "Budget Name"
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
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   90
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3495
      TabIndex        =   15
      Top             =   3360
      Width           =   945
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   1140
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   1140
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1995
      TabIndex        =   14
      Top             =   3360
      Width           =   945
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
      Height          =   2670
      Left            =   150
      ScaleHeight     =   2610
      ScaleWidth      =   2955
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   3015
      Begin VB.ListBox lbcBudget 
         Appearance      =   0  'Flat
         Height          =   2550
         Left            =   30
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   30
         Width           =   2895
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   840
      Top             =   3315
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "BudModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Budmodel.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: BudModel.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim tmMnfBudgetCode() As SORTCODE
Dim smMnfBudgetCodeTag As String

Dim tmBudgetCode() As SORTCODE
Dim smBudgetCodeTag As String

'Budget of Office
Dim hmBvf As Integer    'Budget Office file handle
Dim tmBvf As BVF
Dim tmBvfSrchKey As BVFKEY0    'Bvf key record image
Dim imBvfRecLen As Integer        'Bvf record length
'Budget of Salesperson
Dim hmBsf As Integer    'Budget by Salesperson file handle
Dim tmBsf As BSF
Dim tmBsfSrchKey As BSFKEY0    'Bsf key record image
Dim imBsfReclen As Integer        'Bsf record length
'Budget Names
Dim hmMnf As Integer        'Multi-Name file handle
Dim tmMnf As MNF            'MNF record image
Dim imMnfRecLen As Integer  'MNSF record length
'Program library dates Field Areas
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim smMnfName As String
Dim smCaption As String

Private Sub cbcMnfBudget_Change()
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered

    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    ilRet = gOptionLookAhead(cbcMnfBudget, imBSMode, slStr)
    If ilRet = 0 Then
        imSelectedIndex = cbcMnfBudget.ListIndex
        smMnfName = cbcMnfBudget.List(imSelectedIndex)
    Else
        imSelectedIndex = -1
        smMnfName = slStr
    End If
    mSetCommands
    Screen.MousePointer = vbDefault
    imChgMode = False
    imBypassSetting = False
    Exit Sub

    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    imChgMode = False
    Exit Sub
End Sub
Private Sub cbcMnfBudget_Click()
    cbcMnfBudget_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcMnfBudget_GotFocus()
    gCtrlGotFocus cbcMnfBudget
End Sub
Private Sub cbcMnfBudget_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcMnfBudget_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcMnfBudget.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    If ilKey = Asc("/") Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub cmcCancel_Click()
    igBDReturn = 0
    mTerminate
End Sub
Private Sub cmcDone_Click()
    Dim slNameCode As String
    Dim slNameYear As String
    Dim ilYear As Integer
    Dim slYear As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim ilMnfBudget As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilNoWks As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    If igBDView = 0 Then
        'Test budget name- If New then btrInsert, If Old test if year
        'and budget previously defined
        slYear = edcYear(0).Text
        ilYear = Val(slYear)
        If ilYear < 100 Then
            If ilYear >= 70 Then
                ilYear = 1900 + ilYear
            Else
                ilYear = 2000 + ilYear
            End If
        End If
        If tgSpf.sRUseCorpCal = "Y" Then
            ilFound = False
            For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
                If tgMCof(ilLoop).iYear = ilYear Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                MsgBox "Corporate Year Missing for" & Str$(ilYear), vbOKOnly + vbExclamation, "Budget"
                edcYear(0).SetFocus
                Exit Sub
            End If
        End If
        If imSelectedIndex >= 0 Then
            slNameCode = tmMnfBudgetCode(imSelectedIndex).sKey  'lbcMnfBudgetCode.List(imSelectedIndex)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilRet = gParseItem(slNameCode, 1, "\", sgBudgetName)
            ilMnfBudget = Val(slCode)
            tmBvfSrchKey.iYear = ilYear
            tmBvfSrchKey.iSeqNo = 1
            tmBvfSrchKey.iMnfBudget = ilMnfBudget
            ilRet = btrGetEqual(hmBvf, tmBvf, imBvfRecLen, tmBvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                MsgBox "Year previously defined for the selected Budget Name", vbExclamation, "Error"
                edcYear(0).SetFocus
                Exit Sub
            End If
            igNewYear = ilYear
            igNewMnfBudget = ilMnfBudget
        Else
            If Len(smMnfName) = 0 Then
                MsgBox "Budget Name must be defined", vbExclamation, "Error"
                cbcMnfBudget.SetFocus
                Exit Sub
            End If
            gFindMatch smMnfName, 0, cbcMnfBudget
            If gLastFound(cbcMnfBudget) >= 0 Then
                slNameCode = tmMnfBudgetCode(gLastFound(cbcMnfBudget)).sKey   'lbcMnfBudgetCode.List(gLastFound(cbcMnfBudget))
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilMnfBudget = Val(slCode)
                tmBvfSrchKey.iYear = ilYear
                tmBvfSrchKey.iSeqNo = 1
                tmBvfSrchKey.iMnfBudget = ilMnfBudget
                ilRet = btrGetEqual(hmBvf, tmBvf, imBvfRecLen, tmBvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    MsgBox "Year previously defined for the selected Budget Name", vbExclamation, "Error"
                    edcYear(0).SetFocus
                    Exit Sub
                End If
                igNewYear = ilYear
                igNewMnfBudget = ilMnfBudget
            Else
                tmMnf.iCode = 0
                tmMnf.sName = smMnfName
                tmMnf.sType = "U"
                tmMnf.sRPU = ""
                tmMnf.sCodeStn = ""
                tmMnf.iGroupNo = igBudgetType
                tmMnf.sUnitType = ""
                tmMnf.sSSComm = ""
                ilRet = btrInsert(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0)
                gGetSyncDateTime slSyncDate, slSyncTime
                Do
                    'tmMnfSrchKey.iCode = tmMnf.iCode
                    'ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
                    tmMnf.iAutoCode = tmMnf.iCode
                    gPackDate slSyncDate, tmMnf.iSyncDate(0), tmMnf.iSyncDate(1)
                    gPackTime slSyncTime, tmMnf.iSyncTime(0), tmMnf.iSyncTime(1)
                    ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
                igNewMnfBudget = tmMnf.iCode
                igNewYear = ilYear
            End If
            sgBudgetName = smMnfName
        End If
        If igBudgetType = 0 Then
            If lbcBudget.ListIndex >= 1 Then
                ilIndex = lbcBudget.ListIndex
                slNameCode = tmBudgetCode(ilIndex - 1).sKey 'lbcBudgetCode.List(ilIndex - 1)
                ilRet = gParseItem(slNameCode, 1, "\", slNameYear)
                ilRet = gParseItem(slNameYear, 1, "/", slYear)
                slYear = gSubStr("9999", slYear)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                igModelMnfBudget = Val(slCode)
                igModelYear = Val(slYear)
                'Test if # weeks differ between years
                If tgSpf.sRUseCorpCal = "Y" Then
                    slStartDate = gObtainYearStartDate(4, "1/15/" & Trim$(Str$(igModelYear)))
                    slEndDate = gObtainYearEndDate(4, "1/15/" & Trim$(Str$(igModelYear)))
                    ilNoWks = gDateValue(slEndDate) - gDateValue(slStartDate)
                    slStartDate = gObtainYearStartDate(4, "1/15/" & Trim$(Str$(igNewYear)))
                    slEndDate = gObtainYearEndDate(4, "1/15/" & Trim$(Str$(igNewYear)))
                Else
                    slStartDate = gObtainYearStartDate(0, "1/15/" & Trim$(Str$(igModelYear)))
                    slEndDate = gObtainYearEndDate(0, "1/15/" & Trim$(Str$(igModelYear)))
                    ilNoWks = gDateValue(slEndDate) - gDateValue(slStartDate)
                    slStartDate = gObtainYearStartDate(0, "1/15/" & Trim$(Str$(igNewYear)))
                    slEndDate = gObtainYearEndDate(0, "1/15/" & Trim$(Str$(igNewYear)))
                End If
                If (gDateValue(slEndDate) - gDateValue(slStartDate)) <> ilNoWks Then
                    MsgBox "Budget Amounts Altered because # Weeks Differs between Years", vbExclamation, "Information"
                End If
            Else
                igModelMnfBudget = 0
                igModelYear = 0
            End If
        Else
            igModelMnfBudget = 0
            igModelYear = 0
        End If
        If igBudgetType = 0 Then
            If rbcBT(0).Value Then
                igDirect = 0
            Else
                igDirect = 1
            End If
        Else
            igDirect = 0
        End If
    Else
        slYear = edcYear(1).Text
        ilYear = Val(slYear)
        If ilYear < 100 Then
            If ilYear >= 70 Then
                ilYear = 1900 + ilYear
            Else
                ilYear = 2000 + ilYear
            End If
        End If
        If tgSpf.sRUseCorpCal = "Y" Then
            ilFound = False
            For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
                If tgMCof(ilLoop).iYear = ilYear Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                MsgBox "Corporate Year Missing for" & Str$(ilYear), vbOKOnly + vbExclamation, "Budget"
                edcYear(1).SetFocus
                Exit Sub
            End If
        End If
        tmBsfSrchKey.iYear = ilYear
        tmBsfSrchKey.iSeqNo = 1
        tmBsfSrchKey.iSlfCode = 0
        ilRet = btrGetGreaterOrEqual(hmBsf, tmBsf, imBsfReclen, tmBsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        If (ilRet = BTRV_ERR_NONE) And (tmBsf.iYear = ilYear) Then
            MsgBox "Year previously defined for Salespeople", vbExclamation, "Error"
            edcYear(1).SetFocus
            Exit Sub
        End If
        igNewYear = ilYear

        If lbcBudget.ListIndex >= 1 Then
            ilIndex = lbcBudget.ListIndex
            slYear = lbcBudget.List(ilIndex)
            igModelYear = Val(slYear)
        Else
            igModelYear = 0
        End If
    End If
    igBDReturn = 1
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcYear_Change(Index As Integer)
    mSetCommands
End Sub
Private Sub edcYear_GotFocus(Index As Integer)
    gCtrlGotFocus edcYear(Index)
End Sub
Private Sub edcYear_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcYear_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim slStr As String
    Dim slComp As String
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcYear(Index).SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcYear(Index).Text
    slStr = Left$(slStr, edcYear(Index).SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcYear(Index).SelStart - edcYear(Index).SelLength)
    slComp = "2040"
    If gCompNumberStr(slStr, slComp) > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If igBudgetType = 0 Then
        smCaption = "New Budget"
        plcModel.Visible = True
        plcModel.Move 150, 270
        plcSpec(0).Move 3240, 270
        lacBudgetName.Caption = "Budget Name"
        lacYear(0).Caption = "Budget Year"
    Else
        smCaption = "New Actuals"
        plcModel.Visible = False
        plcSpec(0).Move 1635, 270
        lacBudgetName.Caption = "Actual Name"
        plcBT.Visible = False
        lacYear(0).Caption = "Actual Year"
    End If
    If (igWinStatus(RATECARDSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        lbcBudget.Enabled = False
    Else
        lbcBudget.Enabled = True
    End If
'    gShowBranner
    Me.KeyPreview = True
    BudModel.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
        If plcSpec(0).Visible Then
            plcSpec(0).Visible = False
            plcSpec(0).Visible = True
        ElseIf plcSpec(1).Visible Then
            plcSpec(1).Visible = False
            plcSpec(1).Visible = True
        End If
    End If
End Sub

Private Sub Form_Load()
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tmMnfBudgetCode
    Erase tmBudgetCode

    Screen.MousePointer = vbDefault
    btrExtClear hmMnf   'Clear any previous extend operation
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    btrExtClear hmBsf   'Clear any previous extend operation
    ilRet = btrClose(hmBsf)
    btrDestroy hmBsf
    btrExtClear hmBvf   'Clear any previous extend operation
    ilRet = btrClose(hmBvf)
    btrDestroy hmBvf
    
    Set BudModel = Nothing   'Remove data segment
 
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
    imFirstActivate = True
    imTerminate = False

    Screen.MousePointer = vbHourglass
    BudModel.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone BudModel
    hmBvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmBvf, "", sgDBPath & "Bvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Bvf.Btr)", BudModel
    On Error GoTo 0
    imBvfRecLen = Len(tmBvf)
    hmBsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmBsf, "", sgDBPath & "Bsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Bsf.Btr)", BudModel
    On Error GoTo 0
    imBsfReclen = Len(tmBsf)
    hmMnf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf.Btr)", BudModel
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    If igBDView = 0 Then
        plcSpec(0).Visible = True
        plcSpec(1).Visible = False
    Else
        plcSpec(0).Visible = False
        plcSpec(1).Visible = True
    End If
    'BudModel.Show
    Screen.MousePointer = vbHourglass
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    mMnfBudgetPop
    If imTerminate Then
        Exit Sub
    End If
    If lbcBudget.ListCount > 1 Then
        lbcBudget.ListIndex = 1
    Else
        lbcBudget.ListIndex = 0
    End If
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    gCenterModalForm BudModel
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMnfBudgetPop                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Mnf by Budget list    *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mMnfBudgetPop()
'
'   mMnfBudgetPop
'   Where:
'
    ReDim ilFilter(0 To 1) As Integer
    ReDim slFilter(0 To 1) As String
    ReDim ilOffSet(0 To 1) As Integer
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    ilFilter(0) = CHARFILTER
    slFilter(0) = "U"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilFilter(1) = INTEGERFILTER
    slFilter(1) = Trim$(Str$(igBudgetType))
    ilOffSet(1) = gFieldOffset("Mnf", "MnfGroupNo") '2
    'ilRet = gIMoveListBox(BudModel, cbcMnfBudget, lbcMnfBudgetCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    ilRet = gIMoveListBox(BudModel, cbcMnfBudget, tmMnfBudgetCode(), smMnfBudgetCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mMnfBudgetPopErr
        gCPErrorMsg ilRet, "mMnfBudgetPop (gIMoveListBox)", BudModel
        On Error GoTo 0
    End If
    Exit Sub
mMnfBudgetPopErr:
    On Error GoTo 0
    imTerminate = True
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

    If igBDView = 0 Then
        'Show Unique Budget names and year
        imPopReqd = False
        'ilRet = gPopVehBudgetBox(BudModel, 0, 1, lbcBudget, lbcBudgetCode)
        ilRet = gPopVehBudgetBox(BudModel, 0, 0, 1, lbcBudget, tmBudgetCode(), smBudgetCodeTag)
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mPopulateErr
            gCPErrorMsg ilRet, "mPopulate (gPopBudgetBox)", BudModel
            On Error GoTo 0
            lbcBudget.AddItem "[None]", 0  'Force as first item on list
            imPopReqd = True
        End If
    Else
        'Show unique years only
        imPopReqd = False
        'ilRet = gPopSlspBudgetBox(BudModel, lbcBudget, lbcBudgetCode)
        ilRet = gPopSlspBudgetBox(BudModel, lbcBudget, tmBudgetCode(), smBudgetCodeTag)
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mPopulateErr
            gCPErrorMsg ilRet, "mPopulate (gPopBudgetBox)", BudModel
            On Error GoTo 0
            lbcBudget.AddItem "[None]", 0  'Force as first item on list
            imPopReqd = True
        End If
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
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
    Dim slStr As String
    If igBDView = 0 Then
        If igBudgetType = 0 Then
            slStr = Trim$(smMnfName)
            If Len(slStr) > 0 Then
                If rbcBT(0).Value Or rbcBT(1).Value Then
                    slStr = Trim$(edcYear(0).Text)
                    If (Len(slStr) = 2) Or (Len(slStr) = 4) Then
                        cmcDone.Enabled = True
                    Else
                        cmcDone.Enabled = False
                    End If
                Else
                    cmcDone.Enabled = False
                End If
            Else
                cmcDone.Enabled = False
            End If
        Else
            slStr = Trim$(smMnfName)
            If Len(slStr) > 0 Then
                slStr = Trim$(edcYear(0).Text)
                If (Len(slStr) = 2) Or (Len(slStr) = 4) Then
                    cmcDone.Enabled = True
                Else
                    cmcDone.Enabled = False
                End If
            Else
                cmcDone.Enabled = False
            End If
        End If
    Else
        slStr = Trim$(edcYear(1).Text)
        If (Len(slStr) = 2) Or (Len(slStr) = 4) Then
            cmcDone.Enabled = True
        Else
            cmcDone.Enabled = False
        End If
    End If
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

    igManUnload = YES
    Unload BudModel
    igManUnload = NO
End Sub
Private Sub rbcBT_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcBT(Index).Value
    'End of coded added
    If Value Then
        mSetCommands
    End If
End Sub
Private Sub plcBT_Paint()
    plcBT.CurrentX = 0
    plcBT.CurrentY = 0
    plcBT.Print "Budget Type"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smCaption
End Sub
