VERSION 5.00
Begin VB.Form CPackage 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4935
   ClientLeft      =   1320
   ClientTop       =   2970
   ClientWidth     =   8160
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
   ScaleHeight     =   4935
   ScaleWidth      =   8160
   Begin VB.ListBox lbcCPMVefSort 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   5610
      Sorted          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4545
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.PictureBox pbcLbcVehicle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2775
      ScaleHeight     =   195
      ScaleWidth      =   840
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1275
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.ComboBox cbcSelect 
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
      Left            =   2910
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   285
      Width           =   5025
   End
   Begin VB.ListBox lbcVehGp3 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   5220
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1230
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Lines"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   150
      TabIndex        =   20
      Top             =   4425
      Width           =   1125
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1620
      Top             =   4470
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2025
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.ListBox lbcDaypart 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2145
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1725
      Visible         =   0   'False
      Width           =   2325
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
      Left            =   4950
      Picture         =   "CPACKAGE.frx":0000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   945
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pbcType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1005
      ScaleHeight     =   210
      ScaleWidth      =   1395
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   1395
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
      HelpContextID   =   8
      Left            =   2370
      MaxLength       =   20
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   945
      Visible         =   0   'False
      Width           =   2580
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
      Height          =   45
      Left            =   75
      ScaleHeight     =   45
      ScaleWidth      =   15
      TabIndex        =   12
      Top             =   1110
      Width           =   15
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
      Height          =   45
      Left            =   120
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   705
      Width           =   45
   End
   Begin VB.PictureBox pbcPackage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   195
      Picture         =   "CPACKAGE.frx":00FA
      ScaleHeight     =   375
      ScaleWidth      =   7680
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   780
      Width           =   7680
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
      ScaleWidth      =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4650
      Width           =   120
   End
   Begin VB.ComboBox cbcProgram 
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
      Left            =   2430
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   -60
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.ListBox lbcPrg 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1515
      Sorted          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   30
      ScaleHeight     =   270
      ScaleWidth      =   1485
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   1485
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2850
      TabIndex        =   15
      Top             =   4545
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4320
      TabIndex        =   16
      Top             =   4545
      Width           =   1050
   End
   Begin VB.PictureBox plcLines 
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
      Height          =   3030
      Left            =   120
      ScaleHeight     =   2970
      ScaleWidth      =   7815
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1350
      Width           =   7875
      Begin VB.PictureBox pbcLbcLines 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2730
         Left            =   90
         ScaleHeight     =   2730
         ScaleWidth      =   7620
         TabIndex        =   21
         Top             =   75
         Width           =   7620
      End
      Begin VB.ListBox lbcLines 
         Appearance      =   0  'Flat
         Height          =   2760
         Left            =   75
         MultiSelect     =   2  'Extended
         TabIndex        =   14
         Top             =   60
         Width           =   7650
      End
   End
   Begin VB.PictureBox pbcPackage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1620
      Picture         =   "CPACKAGE.frx":D1A4
      ScaleHeight     =   375
      ScaleWidth      =   4725
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   645
      Width           =   4725
   End
   Begin VB.PictureBox plcPackage 
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
      Height          =   480
      Left            =   165
      ScaleHeight     =   420
      ScaleWidth      =   7725
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   735
      Width           =   7785
   End
End
Attribute VB_Name = "CPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CPACKAGE.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CPackage.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract terminate screen code
Option Explicit
Option Compare Text
Dim tmCtrls(0 To 4)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer   'Current Avail Name Box
Dim smSave(0 To 3) As String    'Index 1: Vehicle Name; 2=Daypart Name; 3=Market Name
Dim imSave(0 To 1) As Integer   'Index 1: Invoice (0=Real; 1=Generate; 2= Equal)
Dim tmVehGp3Code() As SORTCODE
Dim smVehGp3CodeTag As String
Dim imVehGp3ChgMode As Integer
Dim smOrigVehGp3 As String
Dim tmTempVehGp3Code() As SORTCODE
Dim imMktAdj As Integer             '0 if market base or 2 if not market base
Dim smAlterName As String
Dim smOrigVehName As String
Dim smOrigDPName As String
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imChgMode As Integer
Dim imBSMode As Integer
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imSelectedIndex As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imFirstFocus As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imLbcArrowSetting As Integer
Dim imBypassFocus As Integer
Dim imTabDirection As Integer
Dim imDoubleClickName As Integer
Dim imRetBranch As Integer
Dim imUpdateAllowed As Integer
Dim imInPast As Integer
Dim imSetAll As Integer
Dim imAllClicked As Integer
'Dim imListField(1 To 6) As Integer
Dim imListField(0 To 6) As Integer

Dim imINPBCPaint As Integer
Dim imPbcIndex As Integer   '0=Dynamic package for Air Time; 1=Dynamic package for CPM
Dim imAdjY As Integer   '0 for imPBCIndex = 0; 15 for imPBCIndex = 1

Const INVOICEINDEX = 1
Const VEHICLEINDEX = 2
Const DAYPARTINDEX = 3
Const MKTINDEX = 4
Private Sub cbcSelect_Change()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim slName As String
    Dim ilClf As Integer
    Dim ilVef As Integer

    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    Screen.MousePointer = vbHourglass  'Wait
    If Contract.grdCPMBuy.Visible Then
        smAlterName = "Y"
        ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
        If ilRet = 0 Then
            imSelectedIndex = cbcSelect.ListIndex
            mLinePop imSelectedIndex
        Else
            cbcSelect.ListIndex = 0
            imSelectedIndex = 0
            mLinePop -1
            mClearCtrlFields
        End If
        pbcPackage_Paint imPbcIndex
        imChgMode = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    If ilRet = 0 Then
        slName = cbcSelect.List(cbcSelect.ListIndex)
        For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
            If (tgPackLn(ilLoop).sType = "O") Or (tgPackLn(ilLoop).sType = "A") Or (tgPackLn(ilLoop).sType = "E") Then
                slStr = Trim$(str$(tgPackLn(ilLoop).iLineNo)) & " "
                If slName = slStr & Trim$(tgPackLn(ilLoop).sVehName) & " " & Trim$(tgPackLn(ilLoop).sDPName) Then
                    For ilClf = 0 To UBound(tgClfCntr) Step 1
                        If tgPackLn(ilLoop).iLineNo = tgClfCntr(ilClf).ClfRec.iLine Then
                            imInPast = tgClfCntr(ilClf).iInPast
                            Exit For
                        End If
                    Next ilClf
                    smSave(1) = tgPackLn(ilLoop).sVehName
                    smOrigVehName = Trim$(smSave(1))
                    smSave(2) = tgPackLn(ilLoop).sDPName
                    smOrigDPName = Trim$(smSave(2))
                    smSave(3) = ""
                    smAlterName = tgPackLn(ilLoop).sAlterName
                    If tgPackLn(ilLoop).iMnfVehGp3Mkt > 0 Then
                        For ilVef = 0 To UBound(tmVehGp3Code) - 1 Step 1 'lbcVehGpCode.ListCount - 1 Step 1
                            slNameCode = tmVehGp3Code(ilVef).sKey   'lbcVehGpCode.List(ilVef)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            If Val(slCode) = tgPackLn(ilLoop).iMnfVehGp3Mkt Then
                                lbcVehGp3.ListIndex = ilVef + imMktAdj
                                smSave(3) = lbcVehGp3.List(ilVef + imMktAdj)
                                Exit For
                            End If
                        Next ilVef
                    End If
                    smSave(3) = Trim$(smSave(3))
                    smOrigVehGp3 = smSave(3)
                    If (tgPackLn(ilLoop).sType = "O") Then
                        imSave(1) = 1
                    ElseIf (tgPackLn(ilLoop).sType = "E") Then
                        imSave(1) = 2
                    Else
                        imSave(1) = 0
                    End If
                    ilRet = mLinePop(ilLoop)
                    Exit For
                End If
            End If
        Next ilLoop
    Else
        If ilRet = 1 Then
            cbcSelect.ListIndex = 0
        End If
        ilRet = mLinePop(-1)
        ilRet = 1   'Clear fields as no match name found
    End If
    pbcPackage(imPbcIndex).Cls
    pbcPackage(1).Cls
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            smSave(1) = slStr
        End If
    End If
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        Select Case ilLoop 'Branch on box type (control)
            Case INVOICEINDEX 'Name
                If imSave(1) = 0 Then
                    slStr = "Real"
                ElseIf imSave(1) = 1 Then
                    slStr = "Virtual"   '"Generate"
                ElseIf imSave(1) = 2 Then
                    slStr = "Equal"
                Else
                    slStr = ""
                End If
                gSetShow pbcPackage(imPbcIndex), slStr, tmCtrls(ilLoop)
            Case VEHICLEINDEX 'Vehicle
                gSetShow pbcPackage(imPbcIndex), smSave(1), tmCtrls(ilLoop)
            Case DAYPARTINDEX 'Daypart
                gSetShow pbcPackage(imPbcIndex), smSave(2), tmCtrls(ilLoop)
            Case MKTINDEX 'Market
                gSetShow pbcPackage(imPbcIndex), smSave(3), tmCtrls(ilLoop)
        End Select
    Next ilLoop
    pbcPackage_Paint imPbcIndex
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
Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcSelect_DropDown()
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
End Sub
Private Sub cbcSelect_GotFocus()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slName As String
    Dim slStr As String
    Dim ilLineNo As Integer
    Dim ilRet As Integer
    If imTerminate Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        If Contract.grdCPMBuy.Visible Then
            If igPackLnNo > 0 Then
                For ilIndex = 1 To cbcSelect.ListCount Step 1
                    If Contract.grdCPMBuy.Row = cbcSelect.ItemData(ilIndex) Then
                        cbcSelect.ListIndex = ilIndex
                        cmcDone.SetFocus
                        Exit Sub
                    End If
                Next ilIndex
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.ListIndex = 0
            End If
            Exit Sub
        End If
        If igPackLnNo > 0 Then 'If from sales office- set name and branch to control
            slName = ""
            For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                If (tgPackLn(ilLoop).sType = "O") Or (tgPackLn(ilLoop).sType = "A") Or (tgPackLn(ilLoop).sType = "E") Then
                    If tgPackLn(ilLoop).iLineNo = igPackLnNo Then
                        slStr = Trim$(str$(tgPackLn(ilLoop).iLineNo)) & " "
                        slName = slStr & Trim$(tgPackLn(ilLoop).sVehName) & " " & Trim$(tgPackLn(ilLoop).sDPName)
                        Exit For
                    Else
                        ilIndex = tgPackLn(ilLoop).iIndex
                        Do While ilIndex > 0
                            If tgPackLn(ilIndex).iLineNo = igPackLnNo Then
                                slStr = Trim$(str$(tgPackLn(ilLoop).iLineNo)) & " "
                                slName = slStr & Trim$(tgPackLn(ilLoop).sVehName) & " " & Trim$(tgPackLn(ilLoop).sDPName)
                                Exit For
                            Else
                                ilIndex = tgPackLn(ilIndex).iIndex
                            End If
                        Loop
                    End If
                End If
            Next ilLoop
            If slName <> "" Then
                mSetCommands
                gFindMatch slName, 1, cbcSelect
                If gLastFound(cbcSelect) > 0 Then
                    cbcSelect.ListIndex = gLastFound(cbcSelect)
                    cmcDone.SetFocus
                    Exit Sub
                End If
            Else
                cbcSelect.ListIndex = 0
            End If
            If pbcSTab.Enabled Then
                pbcSTab.SetFocus
            Else
                cmcCancel.SetFocus
            End If
            Exit Sub
        Else
            cbcSelect.ListIndex = 0
            If UBound(igLnMakingPack) > LBound(igLnMakingPack) Then
                For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
                    slName = lbcLines.List(ilLoop)
                    ilRet = gParseItem(slName, 1, "|", slStr)
                    'slStr = Trim$(Left$(slName, imListFieldChar(1) - 1))
                    ilLineNo = Val(slStr)
                    For ilIndex = LBound(igLnMakingPack) To UBound(igLnMakingPack) - 1 Step 1
                        If ilLineNo = igLnMakingPack(ilIndex) Then
                            lbcLines.Selected(ilLoop) = True
                            Exit For
                        End If
                    Next ilIndex
                Next ilLoop
            End If
            If pbcSTab.Enabled Then
                pbcSTab.SetFocus
            Else
                cmcCancel.SetFocus
            End If
            Exit Sub
        End If
    End If
    'slSvText = cbcSelect.Text
    'mPopulate
    'If imTerminate Then
    '    Exit Sub
    'End If
    'gCtrlGotFocus cbcSelect
    'If cbcSelect.ListCount <= 1 Then
    '    cbcSelect.ListIndex = 0
    '    mClearCtrlFields
    '    If pbcSTab.Enabled Then
    '        pbcSTab.SetFocus
    '    Else
    '        cmcCancel.SetFocus
    '    End If
    '    Exit Sub
    'End If
    'If (slSvText = "") Or (slSvText = "[New]") Then
    '    cbcSelect.ListIndex = 0
    '    cbcSelect_Change    'Call change so picture area repainted
    'Else
    '    gFindMatch slSvText, 1, cbcSelect
    '    If gLastFound(cbcSelect) > 0 Then
    '        If (slSvText <> cbcSelect.List(gLastFound(cbcSelect))) Or imPopReqd Then
    '            cbcSelect.ListIndex = gLastFound(cbcSelect)
    '            cbcSelect_Change    'Call change so picture area repainted
    '            imPopReqd = False
    '        End If
    '    Else
    '        cbcSelect.ListIndex = 0
    '        mClearCtrlFields
    '        cbcSelect_Change    'Call change so picture area repainted
    '    End If
    'End If
End Sub
Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim ilValue As Integer
    Dim llRet As Long
    Dim llRg As Long
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        If lbcLines.ListCount > 0 Then
            llRg = CLng(lbcLines.ListCount - 1) * &H10000 + 0
            'llRet = SendMessageByNum(lbcLines.hwnd, &H400 + 28, ilValue, llRg)
            llRet = SendMessageByNum(lbcLines.HWnd, LB_SELITEMRANGE, ilValue, llRg)
        End If
        imAllClicked = False
    End If
    DoEvents
    pbcLbcLines_Paint
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub
Private Sub cmcCancel_Click()
    igTerminateReturn = 0   'Cancel selected
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim ilPrevIndex As Integer
    Dim ilIndex As Integer
    Dim slName As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim ilClf As Integer
    Dim ilLine As Integer
    Dim ilTst As Integer
    Dim llDate As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim llAllowedChgWkStartDate As Long
    Dim slLnLLD As String
    Dim ilRsch As Integer
    
    If Contract.grdCPMBuy.Visible Then
        If mSaveCPM(imSelectedIndex) Then
            mTerminate
        End If
        Exit Sub
    End If
    ilFound = False
    For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
        If lbcLines.Selected(ilLoop) Then
            ilFound = True
            Exit For
        End If
    Next ilLoop
    If Not ilFound Then
        If imSelectedIndex > 0 Then
            'Remove all line from package
            slName = cbcSelect.List(imSelectedIndex)
            For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                If (tgPackLn(ilLoop).sType = "O") Or (tgPackLn(ilLoop).sType = "A") Or (tgPackLn(ilLoop).sType = "E") Then
                    slStr = Trim$(str$(tgPackLn(ilLoop).iLineNo)) & " "
                    If slName = slStr & Trim$(tgPackLn(ilLoop).sVehName) & " " & Trim$(tgPackLn(ilLoop).sDPName) Then
                        For ilClf = 0 To UBound(tgClfCntr) Step 1
                            If tgPackLn(ilLoop).iLineNo = tgClfCntr(ilClf).ClfRec.iLine Then
                                If (tgClfCntr(ilClf).iStatus = 0) Or ((tgClfCntr(ilClf).iStatus = 1) And (tgClfCntr(ilClf).ClfRec.sSchStatus <> "F") And (tgClfCntr(ilClf).ClfRec.sSchStatus <> "M")) Then 'new line
                                    igTerminateReturn = 3 'Remove package line
                                    igPackLnNo = tgPackLn(ilLoop).iLineNo
                                    mTerminate
                                    Exit Sub
                                Else    'Test dates-Later look at mLnDate
                                    ilFound = False
                                    For ilTst = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                                        If (tgClfCntr(ilTst).ClfRec.iPkLineNo = tgClfCntr(ilClf).ClfRec.iLine) Then
                                            If (tgClfCntr(ilTst).iStatus = 1) And ((tgClfCntr(ilTst).ClfRec.sSchStatus = "F") Or (tgClfCntr(ilTst).ClfRec.sSchStatus = "M")) Then
                                                llAllowedChgWkStartDate = mLnDate(ilTst + 1, slLnLLD, False)
                                                gUnpackDateLong tgClfCntr(ilTst).ClfRec.iStartDate(0), tgClfCntr(ilTst).ClfRec.iStartDate(1), llDate
                                                If llDate < llAllowedChgWkStartDate Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    Next ilTst
                                    If Not ilFound Then
                                        igTerminateReturn = 3 'Remove package line
                                        igPackLnNo = tgPackLn(ilLoop).iLineNo
                                        mTerminate
                                        Exit Sub
                                    End If
                                End If
                                igTerminateReturn = 0   'Cancel selected
                                mTerminate
                                Exit Sub
                            End If
                        Next ilClf
                    End If
                End If
            Next ilLoop
        End If
        igTerminateReturn = 0   'Cancel selected
        mTerminate
        Exit Sub
    End If
    If mTestSaveFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        Beep
        mEnableBox imBoxNo
        Exit Sub
    End If
    'Create new line which is a made up of lines selected
    If imSelectedIndex > 0 Then
        igTerminateReturn = 0
        'Test if lines removed- it type will be set in mSaveRec function to "S"
        slName = cbcSelect.List(imSelectedIndex)
        For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
            If (tgPackLn(ilLoop).sType = "O") Or (tgPackLn(ilLoop).sType = "A") Or (tgPackLn(ilLoop).sType = "E") Then
                slStr = Trim$(str$(tgPackLn(ilLoop).iLineNo)) & " "
                If slName = slStr & Trim$(tgPackLn(ilLoop).sVehName) & " " & Trim$(tgPackLn(ilLoop).sDPName) Then
                    If imSave(1) = 0 Then
                        If tgPackLn(ilLoop).sType <> "A" Then
                            igTerminateReturn = 2
                        End If
                        tgPackLn(ilLoop).sType = "A"
                    ElseIf imSave(1) = 2 Then
                        If tgPackLn(ilLoop).sType <> "E" Then
                            igTerminateReturn = 2
                        End If
                        tgPackLn(ilLoop).sType = "E"
                    Else
                        If tgPackLn(ilLoop).sType <> "O" Then
                            igTerminateReturn = 2
                        End If
                        tgPackLn(ilLoop).sType = "O"
                    End If
                    If Trim$(tgPackLn(ilLoop).sVehName) <> Trim$(smSave(1)) Then
                        igTerminateReturn = 2
                    End If
                    tgPackLn(ilLoop).sVehName = smSave(1)
                    If Trim$(tgPackLn(ilLoop).sDPName) <> Trim$(smSave(2)) Then
                        igTerminateReturn = 2
                    End If
                    tgPackLn(ilLoop).sDPName = smSave(2)
                    slStr = smSave(3)
                    gFindMatch slStr, imMktAdj, lbcVehGp3
                    If ((gLastFound(lbcVehGp3) > 1) And (imMktAdj = 2)) Or ((gLastFound(lbcVehGp3) >= 0) And (imMktAdj = 0)) Then
                        slNameCode = tmVehGp3Code(gLastFound(lbcVehGp3) - imMktAdj).sKey  'lbcVehGpCode.List(gLastFound(lbcVehGp) - 1)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If tgPackLn(ilLoop).iMnfVehGp3Mkt <> Val(slCode) Then
                            igTerminateReturn = 2
                        End If
                        tgPackLn(ilLoop).iMnfVehGp3Mkt = Val(slCode)
                    Else
                        If tgPackLn(ilLoop).iMnfVehGp3Mkt <> 0 Then
                            igTerminateReturn = 2
                        End If
                        tgPackLn(ilLoop).iMnfVehGp3Mkt = 0
                    End If
                    imVehGp3ChgMode = False
                    ilIndex = tgPackLn(ilLoop).iIndex
                    For ilClf = 0 To UBound(tgClfCntr) Step 1
                        If tgPackLn(ilLoop).iLineNo = tgClfCntr(ilClf).ClfRec.iLine Then
                            igTerminateReturn = 2 'Change mode
                            igPackLnNo = tgPackLn(ilLoop).iLineNo
                            ilPrevIndex = ilLoop
                            ilIndex = tgPackLn(ilLoop).iIndex
                            Do While ilIndex > 0
                                If ((Asc(tgSpf.sUsingFeatures) And HIDDENOVERRIDE) = HIDDENOVERRIDE) Then
                                    For ilRsch = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                                        If tgClfCntr(ilRsch).ClfRec.iLine = tgPackLn(ilIndex).iLineNo Then
                                            tgClfCntr(ilRsch).iResearch = False
                                            Exit For
                                        End If
                                    Next ilRsch
                                End If
                                tgPackLn(ilIndex).iLineNo = 0
                                tgPackLn(ilIndex).sType = ""
                                ilIndex = tgPackLn(ilIndex).iIndex
                            Loop
                            ilUpper = UBound(tgPackLn)
                            For ilLine = 0 To lbcLines.ListCount - 1 Step 1
                                If lbcLines.Selected(ilLine) Then
                                    tgPackLn(ilPrevIndex).iIndex = ilUpper
                                    tgPackLn(ilUpper).sType = "H"
                                    'tgPackLn(ilUpper).iAlter = True
                                    tgPackLn(ilUpper).sAlter = "Y"
                                    tgPackLn(ilUpper).sAlterName = "Y"
                                    slName = lbcLines.List(ilLine)
                                    ilRet = gParseItem(slName, 1, "|", slStr)
                                    'slStr = Trim$(Left$(slName, imListFieldChar(1) - 1))
                                    tgPackLn(ilUpper).iLineNo = Val(slStr)
                                    If ((Asc(tgSpf.sUsingFeatures) And HIDDENOVERRIDE) = HIDDENOVERRIDE) Then
                                        For ilRsch = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                                            If tgClfCntr(ilRsch).ClfRec.iLine = tgPackLn(ilUpper).iLineNo Then
                                                tgClfCntr(ilRsch).iResearch = False
                                                Exit For
                                            End If
                                        Next ilRsch
                                    End If
                                    tgPackLn(ilUpper).iIndex = -1
                                    tgPackLn(ilUpper).sVehName = ""
                                    tgPackLn(ilUpper).sDPName = ""
                                    tgPackLn(ilUpper).iMnfVehGp3Mkt = 0
                                    ReDim Preserve tgPackLn(0 To ilUpper + 1) As PACKLN
                                    ilPrevIndex = ilUpper
                                    ilUpper = ilUpper + 1
                                End If
                            Next ilLine
                            mTerminate
                            Exit Sub
                        End If
                    Next ilClf
                End If
            End If
        Next ilLoop
        mTerminate
        Exit Sub
    Else
        ilUpper = UBound(tgPackLn)
        If imSave(1) = 0 Then
            tgPackLn(ilUpper).sType = "A"
        ElseIf imSave(1) = 2 Then
            tgPackLn(ilUpper).sType = "E"
        Else
            tgPackLn(ilUpper).sType = "O"
        End If
        'tgPackLn(ilUpper).iAlter = True
        tgPackLn(ilUpper).sAlter = "Y"
        tgPackLn(ilUpper).sAlterName = "Y"
        tgPackLn(ilUpper).iLineNo = -1
        tgPackLn(ilUpper).iIndex = -1
        tgPackLn(ilUpper).sVehName = smSave(1)
        tgPackLn(ilUpper).sDPName = smSave(2)
        slStr = smSave(3)
        gFindMatch slStr, imMktAdj, lbcVehGp3
        If ((gLastFound(lbcVehGp3) > 1) And (imMktAdj = 2)) Or ((gLastFound(lbcVehGp3) >= 0) And (imMktAdj = 0)) Then
            slNameCode = tmVehGp3Code(gLastFound(lbcVehGp3) - imMktAdj).sKey  'lbcVehGpCode.List(gLastFound(lbcVehGp) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgPackLn(ilUpper).iMnfVehGp3Mkt = Val(slCode)
        Else
            tgPackLn(ilUpper).iMnfVehGp3Mkt = 0
        End If
        ReDim Preserve tgPackLn(0 To ilUpper + 1) As PACKLN
        ilPrevIndex = ilUpper
        ilUpper = ilUpper + 1
        For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
            If lbcLines.Selected(ilLoop) Then
                tgPackLn(ilPrevIndex).iIndex = ilUpper
                tgPackLn(ilUpper).sType = "H"
                'tgPackLn(ilUpper).iAlter = True
                tgPackLn(ilUpper).sAlter = "Y"
                tgPackLn(ilUpper).sAlterName = "Y"
                slName = lbcLines.List(ilLoop)
                ilRet = gParseItem(slName, 1, "|", slStr)
                'slStr = Trim$(Left$(slName, imListFieldChar(1) - 1))
                tgPackLn(ilUpper).iLineNo = Val(slStr)
                If ((Asc(tgSpf.sUsingFeatures) And HIDDENOVERRIDE) = HIDDENOVERRIDE) Then
                    For ilRsch = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                        If tgClfCntr(ilRsch).ClfRec.iLine = tgPackLn(ilUpper).iLineNo Then
                            tgClfCntr(ilRsch).iResearch = False
                            Exit For
                        End If
                    Next ilRsch
                End If
                tgPackLn(ilUpper).iIndex = -1
                tgPackLn(ilUpper).sVehName = ""
                tgPackLn(ilUpper).sDPName = ""
                tgPackLn(ilUpper).iMnfVehGp3Mkt = 0
                ReDim Preserve tgPackLn(0 To ilUpper + 1) As PACKLN
                ilPrevIndex = ilUpper
                ilUpper = ilUpper + 1
            End If
        Next ilLoop
    End If
    igTerminateReturn = 1   'Done selected
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case VEHICLEINDEX
            lbcVehicle.Visible = Not lbcVehicle.Visible
            imINPBCPaint = False
            If lbcVehicle.Visible Then
                pbcLbcVehicle.Visible = True
            Else
                pbcLbcVehicle.Visible = False
            End If
        Case DAYPARTINDEX
            lbcDaypart.Visible = Not lbcDaypart.Visible
        Case MKTINDEX
            lbcVehGp3.Visible = Not lbcVehGp3.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    If imChgMode Then
        Exit Sub
    End If
    imChgMode = True
    Select Case imBoxNo
        Case VEHICLEINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcVehicle, imBSMode, slStr)
            imLbcArrowSetting = False
        Case DAYPARTINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcDaypart, imBSMode, slStr)
            If ilRet = 1 Then
                If lbcDaypart.ListCount > 0 Then
                    lbcDaypart.ListIndex = 0
                End If
            End If
            imLbcArrowSetting = False
        Case MKTINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcVehGp3, imBSMode, slStr)
            If ilRet = 1 Then
                lbcVehGp3.ListIndex = 0
            End If
            gSetChgFlag smOrigVehGp3, lbcVehGp3, tmCtrls(imBoxNo)
            imLbcArrowSetting = False
    End Select
    mSetCommands
    imChgMode = False
End Sub
Private Sub edcDropDown_DblClick()
    If (imBoxNo = DAYPARTINDEX) Or (imBoxNo = MKTINDEX) Then
        imDoubleClickName = True    'Double click event is followed by a mouse up event
                                    'Process the double click event in the mouse up event
                                    'to avoid the mouse up event being in next form
    End If
End Sub
Private Sub edcDropDown_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer

    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If ActiveControl.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    Select Case imBoxNo
        Case VEHICLEINDEX
            If Not gCheckKeyAscii(ilKey) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case DAYPARTINDEX
            If Not gCheckKeyAscii(ilKey) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case MKTINDEX
            If Not gCheckKeyAscii(ilKey) Then
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        Select Case imBoxNo
            Case VEHICLEINDEX
                gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
            Case DAYPARTINDEX
                gProcessArrowKey Shift, KeyCode, lbcDaypart, imLbcArrowSetting
            Case MKTINDEX
                gProcessArrowKey Shift, KeyCode, lbcVehGp3, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
    End If
End Sub
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imBoxNo
            Case DAYPARTINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
            Case MKTINDEX
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
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    Me.KeyPreview = True
    'Me.Refresh
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        If (cbcSelect.Enabled) And (imBoxNo > 0) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        If ilReSet Then
            cbcSelect.Enabled = True
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
    On Error Resume Next
    
    Erase tmVehGp3Code
    Erase tmTempVehGp3Code
    
    Set CPackage = Nothing   'Remove data segment
    
End Sub

Private Sub lbcDaypart_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        If imLbcMouseDown Then
            imLbcArrowSetting = False
        Else
            imLbcArrowSetting = True
        End If
        gProcessLbcClick lbcDaypart, edcDropDown, imChgMode, imLbcArrowSetting
        imLbcMouseDown = False
    End If
End Sub
Private Sub lbcDaypart_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcDaypart_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcDaypart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcDaypart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcDaypart, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcLines_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
        pbcLbcLines_Paint
    End If
End Sub
Private Sub lbcLines_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    gCtrlGotFocus lbcLines
End Sub

Private Sub lbcLines_Scroll()
    pbcLbcLines_Paint
End Sub

Private Sub lbcVehGp3_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcVehGp3, edcDropDown, imVehGp3ChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcVehGp3_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcVehGp3_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcVehGp3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcVehGp3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcVehGp3, edcDropDown, imVehGp3ChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcVehicle_Click()
    If imLbcMouseDown Then
        imLbcArrowSetting = False
    Else
        imLbcArrowSetting = True
    End If
    gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
    If Not lbcVehicle.Visible Then
        pbcLbcVehicle.Visible = False
    Else
        pbcLbcVehicle_Paint
    End If
    imLbcMouseDown = False
End Sub
Private Sub lbcVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcVehicle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    smSave(1) = ""
    smSave(2) = ""
    smSave(3) = ""
    smOrigVehName = ""
    smOrigDPName = ""
    smOrigVehGp3 = ""
    smAlterName = "Y"
    imSave(1) = -1
    imInPast = False
    edcDropDown.Text = ""
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
        tmCtrls(ilLoop).sShow = ""
    Next ilLoop
    If imPbcIndex = 1 Then
        imSave(1) = 1
        smSave(3) = ""
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDPBranch                       *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to DP     *
'*                      names and process communication*
'*                      back from daypart names        *
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
Private Function mDPBranch()
'
'   ilRet = mDPBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcDaypart, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mDPBranch = False
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(AVAILNAMESLIST)) Then
    '    mDPBranch = True
    '    mEnableBox imRCBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    igDPNameCallSource = CALLSOURCERATECARD
    If edcDropDown.Text = "[New]" Then
        sgDPName = ""
    Else
        sgDPName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    igVefCode = -1
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    'If igTestSystem Then
    '    slStr = "Daypart^Test\" & sgUserName & "\" & Trim$(Str$(igANmCallSource)) & "\" & sgANmName
    'Else
    '    slStr = "Daypart^Prod\" & sgUserName & "\" & Trim$(Str$(igANmCallSource)) & "\" & sgANmName
    'End If
    'lgShellRet = Shell(sgExePath & "AName.Exe " & slStr, 1)
    imRetBranch = True
    'RateCard.Enabled = False
    Daypart.Show vbModal
    'Daypart.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'ilParse = gParseItem(slStr, 1, "\", sgANmName)
    'igANmCallSource = Val(sgANmName)
    'ilParse = gParseItem(slStr, 2, "\", sgANmName)
    'Daypart.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    imDoubleClickName = False
    mDPBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    'RateCard.Enabled = True
    gShowBranner imUpdateAllowed
    If igDPAltered Then
        Screen.MousePointer = vbHourglass
        sgMRdfStamp = ""
        ilRet = gObtainRcfRifRdf()
        mPkDPPop lbcDaypart
        lbcDaypart.AddItem "[New]", 0
        Screen.MousePointer = vbDefault
    End If
    If igDPNameCallSource = CALLDONE Then  'Done
        igDPNameCallSource = CALLNONE
        'mDPNameRowPop
        gFindMatch sgDPName, 1, lbcDaypart
        sgDPName = ""
        If gLastFound(lbcDaypart) > 0 Then
            imChgMode = True
            lbcDaypart.ListIndex = gLastFound(lbcDaypart)
            edcDropDown.Text = lbcDaypart.List(lbcDaypart.ListIndex)
            imChgMode = False
            mDPBranch = False
        Else
            imChgMode = True
            lbcDaypart.ListIndex = 0
            edcDropDown.Text = lbcDaypart.List(lbcDaypart.ListIndex)
            imChgMode = False
            edcDropDown.SetFocus
            pbcPackage_Paint imPbcIndex
            imRetBranch = False
            Exit Function
        End If
    End If
    pbcPackage(imPbcIndex).Cls
    pbcPackage_Paint imPbcIndex
    If igDPNameCallSource = CALLCANCELLED Then  'Cancelled
        igDPNameCallSource = CALLNONE
        sgDPName = ""
        mEnableBox imBoxNo
        pbcPackage_Paint imPbcIndex

        imRetBranch = False
        Exit Function
    End If
    If igDPNameCallSource = CALLTERMINATED Then
        igDPNameCallSource = CALLNONE
        sgDPName = ""
        mEnableBox imBoxNo
        pbcPackage_Paint imPbcIndex
        imRetBranch = False
        Exit Function
    End If
    imRetBranch = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilIndex As Integer
    Dim slStr As String
    Dim slName As String
    Dim ilRet As Integer
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case INVOICEINDEX 'Name
            If imSave(1) < 0 Then
                If tgSpf.sCPkAired = "Y" Then
                    imSave(1) = 0    'Aired
                ElseIf tgSpf.sCPkOrdered = "Y" Then
                    imSave(1) = 1    'Ordered
                Else
                    imSave(1) = 2
                End If
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcType.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcPackage(imPbcIndex), pbcType, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcType_Paint
            pbcType.Visible = True
            pbcType.SetFocus
        Case VEHICLEINDEX 'Vehicle
            lbcVehicle.height = gListBoxHeight(lbcVehicle.ListCount, 12)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            If tgSpf.iVehLen <= 40 Then
                edcDropDown.MaxLength = tgSpf.iVehLen
            Else
                edcDropDown.MaxLength = 20
            End If
            gMoveFormCtrl pbcPackage(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            imChgMode = True
            gFindMatch smSave(1), 0, lbcVehicle
            If gLastFound(lbcVehicle) >= 0 Then
                lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
            Else
                If (imSave(1) = 0) Then
                    ilIndex = -1
                    For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
                        If lbcLines.Selected(ilLoop) Then
                            ilIndex = ilLoop
                            Exit For
                        End If
                    Next ilLoop
                    If ilIndex >= 0 Then
                        slName = lbcLines.List(ilIndex)
                        ilRet = gParseItem(slName, 2, "|", slStr)
                        'slStr = Trim$(Mid$(slName, imListFieldChar(1), imListFieldChar(2) - imListFieldChar(1)))
                        gFindMatch slStr, 0, lbcVehicle
                        If gLastFound(lbcVehicle) >= 0 Then
                            lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                            edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                        Else
                            lbcVehicle.ListIndex = -1
                            edcDropDown.Text = smSave(1)
                        End If
                    Else
                        lbcVehicle.ListIndex = -1
                        edcDropDown.Text = smSave(1)
                    End If
                Else
                    lbcVehicle.ListIndex = -1
                    edcDropDown.Text = smSave(1)
                End If
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case DAYPARTINDEX 'Daypart
            lbcDaypart.height = gListBoxHeight(lbcDaypart.ListCount, 12)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcPackage(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcDaypart.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            imChgMode = True
            gFindMatch smSave(2), 0, lbcDaypart
            If gLastFound(lbcDaypart) >= 0 Then
                lbcDaypart.ListIndex = gLastFound(lbcDaypart)
                edcDropDown.Text = lbcDaypart.List(lbcDaypart.ListIndex)
            Else
                lbcDaypart.ListIndex = -1
                edcDropDown.Text = smSave(2)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case MKTINDEX
            mVehGp3Pop
            If imTerminate Then
                Exit Sub
            End If
            lbcVehGp3.height = gListBoxHeight(lbcVehGp3.ListCount, 10)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcPackage(imPbcIndex), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            imVehGp3ChgMode = True
            slStr = smSave(3)
            If imMktAdj = 0 Then
            gFindMatch slStr, 0, lbcVehGp3
            If gLastFound(lbcVehGp3) >= 0 Then
                lbcVehGp3.ListIndex = gLastFound(lbcVehGp3)
                edcDropDown.Text = lbcVehGp3.List(lbcVehGp3.ListIndex)
            Else
                If lbcVehGp3.ListCount > 0 Then
                    lbcVehGp3.ListIndex = 0
                    edcDropDown.Text = lbcVehGp3.List(0)
                Else
                    lbcVehGp3.ListIndex = 0
                    edcDropDown.Text = lbcVehGp3.List(0)
                End If
            End If
           Else
            gFindMatch slStr, 1, lbcVehGp3
            If gLastFound(lbcVehGp3) >= 1 Then
                lbcVehGp3.ListIndex = gLastFound(lbcVehGp3)
                edcDropDown.Text = lbcVehGp3.List(lbcVehGp3.ListIndex)
            Else
                If lbcVehGp3.ListCount > 1 Then
                    lbcVehGp3.ListIndex = 1
                    edcDropDown.Text = lbcVehGp3.List(1)
                Else
                    lbcVehGp3.ListIndex = 0
                    edcDropDown.Text = lbcVehGp3.List(0)
                End If
            End If
            End If
            imVehGp3ChgMode = False
            lbcVehGp3.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - lbcVehGp3.Width, edcDropDown.Top + edcDropDown.height
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
'*             Created:5/18/93       By:D. LeVine      *
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
    Screen.MousePointer = vbHourglass
    If Contract.grdCPMBuy.Visible Then
        imPbcIndex = 1
        imAdjY = -15
    Else
        imPbcIndex = 0
        imAdjY = 0
    End If
    imLBCtrls = 1
    imTerminate = False
    imFirstActivate = True
    imSelectedIndex = -1
    imBoxNo = -1 'Initialize current Box to N/A
    imChgMode = False
    imBSMode = False
    imSetAll = True
    imAllClicked = False
    imTabDirection = 0  'Left to right movement
    imDoubleClickName = False
    imRetBranch = False
    imBypassSetting = False
    imFirstFocus = True
    imLbcMouseDown = False
    imLbcArrowSetting = False
    imPopReqd = False
    imBypassFocus = False
    imInPast = False
    mInitBox
    imListField(1) = 15
    imListField(2) = 6 * igAlignCharWidth '     'Line #
    imListField(3) = 30 * igAlignCharWidth '
    imListField(4) = 55 * igAlignCharWidth '
    imListField(5) = 60 * igAlignCharWidth '
    imListField(6) = 120 * igAlignCharWidth '
    mPkDPPop lbcDaypart
    lbcDaypart.AddItem "[New]", 0
    If tgSpf.sMktBase = "Y" Then
        imMktAdj = 0
    Else
        imMktAdj = 2
    End If
    lbcVehGp3.Clear
    If Contract.grdCPMBuy.Visible Then
        mCPMVehPop
        mPopulate
    Else
        mVehGp3Pop
        mPkVehPop lbcVehicle, False, True
        mPopulate
    End If
    CPackage.height = cmcDone.Top + 5 * cmcDone.height / 3
    gCenterModalForm CPackage
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    flTextHeight = pbcPackage(imPbcIndex).TextHeight("1") - 35
    If imPbcIndex = 1 Then
        pbcPackage(1).Visible = True
        pbcPackage(0).Visible = False
    Else
        pbcPackage(0).Visible = True
        pbcPackage(1).Visible = False
    End If
    'plcSelect.Move 2880, 165
    cbcSelect.Move 2940, 180
    plcLines.Move 120, 1365
    'Position panel and picture areas with panel
    plcPackage.Move 165, 735, pbcPackage(imPbcIndex).Width + fgPanelAdj, pbcPackage(imPbcIndex).height + fgPanelAdj
    pbcPackage(imPbcIndex).Move plcPackage.Left + fgBevelX, plcPackage.Top + fgBevelY
    If imPbcIndex = 1 Then
        'Package Invoice
        gSetCtrl tmCtrls(INVOICEINDEX), 0, 0, 0, 0
        'Vehicle
        gSetCtrl tmCtrls(VEHICLEINDEX), 30, 30, 2340, fgBoxStH
        'Daypart
        gSetCtrl tmCtrls(DAYPARTINDEX), 2385, tmCtrls(VEHICLEINDEX).fBoxY, 2340, fgBoxStH
        'Market
        gSetCtrl tmCtrls(MKTINDEX), 0, 0, 0, 0
        tmCtrls(MKTINDEX).iReq = False
    Else
        'Package Invoice
        gSetCtrl tmCtrls(INVOICEINDEX), 30, 30, 1395, fgBoxStH
        'Vehicle
        gSetCtrl tmCtrls(VEHICLEINDEX), 1440, tmCtrls(INVOICEINDEX).fBoxY, 2340, fgBoxStH
        'Daypart
        gSetCtrl tmCtrls(DAYPARTINDEX), 3795, tmCtrls(INVOICEINDEX).fBoxY, 2340, fgBoxStH
        'Market
        gSetCtrl tmCtrls(MKTINDEX), 6150, tmCtrls(INVOICEINDEX).fBoxY, 1515, fgBoxStH
    End If
    If tgSpf.sMktBase <> "Y" Then
        tmCtrls(MKTINDEX).iReq = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLinePop                        *
'*                                                     *
'*             Created:8/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Population Line box            *
'*                                                     *
'*******************************************************
Private Function mLinePop(ilLnIndex As Integer) As Integer
    Dim ilClf As Integer
    Dim slStr As String
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilTest As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilCff As Integer
    Dim ilIndex As Integer
    Dim ilBypassLine As Integer
    Dim ilSelectIndex As Integer
    lbcLines.Clear
    If Contract.grdCPMBuy.Visible Then
        mLinePop = mCPMPop(ilLnIndex)
        Exit Function
    End If
    ilSelectIndex = -1
    
    If ilLnIndex >= LBound(tgPackLn) Then
        ilIndex = tgPackLn(ilLnIndex).iIndex
        Do While ilIndex > 0
            For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                If tgClfCntr(ilClf).ClfRec.iLine = tgPackLn(ilIndex).iLineNo Then
                    slStr = Trim$(str$(tgClfCntr(ilClf).ClfRec.iLine))
                    slRecCode = Trim$(str$(tgClfCntr(ilClf).ClfRec.iVefCode))
                    For ilTest = 0 To UBound(tmVehicleCode) - 1 Step 1  'Contract!lbcVehicle.ListCount - 1 Step 1
                        slNameCode = tmVehicleCode(ilTest).sKey    'Contract!lbcVehicle.List(ilTest)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If ilRet <> CP_MSG_NONE Then
                            mLinePop = False
                            Exit Function
                        End If
                        If slRecCode = slCode Then
                            ilRet = gParseItem(slNameCode, 1, "\", slName)
                            If ilRet <> CP_MSG_NONE Then
                                mLinePop = False
                                Exit Function
                            End If
                            ilRet = gParseItem(slName, 3, "|", slName)
                            If ilRet <> CP_MSG_NONE Then
                                mLinePop = False
                                Exit Function
                            End If
                            slStr = slStr & "|" & slName
                            Exit For
                        End If
                    Next ilTest
                    'mPrgPop (ilClf)
                    'If imTerminate Then
                    '    Exit Function
                    'End If
                    'For ilLoop = 0 To cbcProgram.ListCount - 1 Step 1
                    '    If lbcPrg(igTabMapIndex).ItemData(ilLoop) = tgClfCntr(ilClf).ClfRec.iRdfCode Then
                    '        slStr = slStr & "|" & cbcProgram.List(ilLoop)
                    '        Exit For
                    '    End If
                    'Next ilLoop
                    slStr = slStr & "|" & mMakeDPName(ilClf + 1)  '
                    slStr = slStr & "|" & Trim$(str$(tgClfCntr(ilClf).ClfRec.iLen))
                    If tgClfCntr(ilClf).iFirstCff <> -1 Then
                        slStartDate = ""
                        slEndDate = ""
                        ilCff = tgClfCntr(ilClf).iFirstCff
                        Do While ilCff <> -1
                            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                If slStartDate = "" Then
                                    gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), slStartDate
                                End If
                                gUnpackDate tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), slEndDate
                            End If
                            ilCff = tgCffCntr(ilCff).iNextCff
                        Loop
                        slStr = slStr & "|" & slStartDate & "-" & slEndDate
                    End If
                    'lbcLines.AddItem gAlignStringByPixel(slStr, "|", imListField(), imListFieldChar())
                    lbcLines.AddItem slStr
                    ilSelectIndex = lbcLines.ListCount - 1
                    Exit For
                End If
            Next ilClf
            ilIndex = tgPackLn(ilIndex).iIndex
        Loop
    End If
    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
        If tgClfCntr(ilClf).iCancel Then
            ilBypassLine = True
        Else
            ilBypassLine = False
            For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                If tgPackLn(ilLoop).iLineNo = tgClfCntr(ilClf).ClfRec.iLine Then
                    ilBypassLine = True
                    Exit For
                End If
            Next ilLoop
        End If
        If Not ilBypassLine Then
            slStr = Trim$(str$(tgClfCntr(ilClf).ClfRec.iLine))
            slRecCode = Trim$(str$(tgClfCntr(ilClf).ClfRec.iVefCode))
            For ilTest = 0 To UBound(tmVehicleCode) - 1 Step 1  'Contract!lbcVehicle.ListCount - 1 Step 1
                slNameCode = tmVehicleCode(ilTest).sKey    'Contract!lbcVehicle.List(ilTest)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If ilRet <> CP_MSG_NONE Then
                    mLinePop = False
                    Exit Function
                End If
                If slRecCode = slCode Then
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    If ilRet <> CP_MSG_NONE Then
                        mLinePop = False
                        Exit Function
                    End If
                    ilRet = gParseItem(slName, 3, "|", slName)
                    If ilRet <> CP_MSG_NONE Then
                        mLinePop = False
                        Exit Function
                    End If
                    slStr = slStr & "|" & slName
                    Exit For
                End If
            Next ilTest
            'mPrgPop (ilClf)
            'If imTerminate Then
            '    Exit Function
            'End If
            'For ilLoop = 0 To cbcProgram.ListCount - 1 Step 1
            '    If lbcPrg(igTabMapIndex).ItemData(ilLoop) = tgClfCntr(ilClf).ClfRec.iRdfCode Then
            '        slStr = slStr & "|" & cbcProgram.List(ilLoop)
            '        Exit For
            '    End If
            'Next ilLoop
            slStr = slStr & "|" & mMakeDPName(ilClf + 1)  '
            slStr = slStr & "|" & Trim$(str$(tgClfCntr(ilClf).ClfRec.iLen))
            If tgClfCntr(ilClf).iFirstCff <> -1 Then
                slStartDate = ""
                slEndDate = ""
                ilCff = tgClfCntr(ilClf).iFirstCff
                Do While ilCff <> -1
                    If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                        If slStartDate = "" Then
                            gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), slStartDate
                        End If
                        gUnpackDate tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), slEndDate
                    End If
                    ilCff = tgCffCntr(ilCff).iNextCff
                Loop
                slStr = slStr & "|" & slStartDate & "-" & slEndDate
            End If
            'lbcLines.AddItem gAlignStringByPixel(slStr, "|", imListField(), imListFieldChar())
            lbcLines.AddItem slStr
        End If
    Next ilClf
    For ilLoop = 0 To ilSelectIndex Step 1
        lbcLines.Selected(ilLoop) = True
    Next ilLoop
    pbcLbcLines_Paint
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llRow As Long
    If cbcSelect.ListCount > 0 Then
        Exit Sub
    End If
    imPopReqd = False
    If Contract.grdCPMBuy.Visible Then
        For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "P") Then
                slStr = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMID)) & " "
                cbcSelect.AddItem slStr & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE)) & " " & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART))
                cbcSelect.ItemData(cbcSelect.NewIndex) = llRow
            End If
        Next llRow
    Else
        For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
            If (tgPackLn(ilLoop).sType = "O") Or (tgPackLn(ilLoop).sType = "A") Or (tgPackLn(ilLoop).sType = "E") Then
                'If tgPackLn(ilLoop).iAlter Then
                If tgPackLn(ilLoop).sAlter = "Y" Then
                    slStr = Trim$(str$(tgPackLn(ilLoop).iLineNo)) & " "
                    cbcSelect.AddItem slStr & Trim$(tgPackLn(ilLoop).sVehName) & " " & Trim$(tgPackLn(ilLoop).sDPName)
                End If
            End If
        Next ilLoop
    End If
    cbcSelect.AddItem "[New]", 0  'Force as first item on list
    cbcSelect.ItemData(cbcSelect.NewIndex) = 0
    imPopReqd = True
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    If imSelectedIndex = 0 Then
        ilAltered = gAnyFieldChgd(tmCtrls(), VEHICLEINDEX)
        If Not ilAltered Then
            ilAltered = gAnyFieldChgd(tmCtrls(), DAYPARTINDEX)
        End If
        If Not ilAltered Then
            ilAltered = gAnyFieldChgd(tmCtrls(), MKTINDEX)
        End If
    Else
        ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    End If
    'Update button set if all mandatory fields have data and any field altered
    'If (mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (ilAltered = YES) Then
    '    cmcUpdate.Enabled = True
    'Else
    '    cmcUpdate.Enabled = False
    'End If
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
        Case INVOICEINDEX 'Name
            pbcType.Visible = True
            pbcType.SetFocus
        Case VEHICLEINDEX 'Allow sustaining
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case DAYPARTINDEX 'Suppress name on contract
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case MKTINDEX 'Suppress name on contract
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
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

    Select Case ilBoxNo 'Branch on box type (control)
        Case INVOICEINDEX 'Name
            pbcType.Visible = False  'Set visibility
            If imSave(1) = 0 Then
                slStr = "Real"
            ElseIf imSave(1) = 1 Then
                slStr = "Virtual"   '"Generate"
            ElseIf imSave(1) = 2 Then
                slStr = "Equal"
            Else
                slStr = ""
            End If
            gSetShow pbcPackage(imPbcIndex), slStr, tmCtrls(ilBoxNo)
        Case VEHICLEINDEX 'Vehicle
            pbcLbcVehicle.Visible = False
            edcDropDown.Visible = False  'Set visibility
            cmcDropDown.Visible = False
            lbcVehicle.Visible = False
            smSave(1) = edcDropDown.Text
            gSetShow pbcPackage(imPbcIndex), smSave(1), tmCtrls(ilBoxNo)
        Case DAYPARTINDEX 'Daypart
            edcDropDown.Visible = False  'Set visibility
            cmcDropDown.Visible = False
            lbcDaypart.Visible = False
            smSave(2) = edcDropDown.Text
            gSetShow pbcPackage(imPbcIndex), smSave(2), tmCtrls(ilBoxNo)
        Case MKTINDEX 'Market
            edcDropDown.Visible = False  'Set visibility
            cmcDropDown.Visible = False
            lbcVehGp3.Visible = False
            smSave(3) = edcDropDown.Text
            gSetShow pbcPackage(imPbcIndex), smSave(3), tmCtrls(ilBoxNo)
    End Select
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
    Unload CPackage
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
    If (ilCtrlNo = INVOICEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imSave(1) = 0 Then
            slStr = "Real"
        ElseIf imSave(1) = 1 Then
            slStr = "Virtual"   '"Generate"
        ElseIf imSave(1) = 2 Then
            slStr = "Equal"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Real Or Virtual must be specified for Package Invoice", tmCtrls(INVOICEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = INVOICEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = VEHICLEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcDropDown, "", "Vehicle must be specified", tmCtrls(VEHICLEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = VEHICLEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = DAYPARTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcDropDown, "", "Daypart must be specified", tmCtrls(DAYPARTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = DAYPARTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = DAYPARTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcDropDown, "", "Market must be specified", tmCtrls(MKTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = MKTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    mTestFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestSaveFields                 *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestSaveFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mTestSaveFields(iTest, iState)
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
    Dim ilVef As Integer
    Dim ilRet As Integer
    
    If Not Contract.grdCPMBuy.Visible Then
        If (ilCtrlNo = INVOICEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If imSave(1) = 0 Then
                slStr = "Real"
            ElseIf imSave(1) = 1 Then
                slStr = "Virtual"   '"Generate"
            ElseIf imSave(1) = 2 Then
                slStr = "Equal"
            Else
                slStr = ""
            End If
            If gFieldDefinedStr(slStr, "", "Real Or Virtual must be specified for Package Invoice", tmCtrls(INVOICEINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = INVOICEINDEX
                End If
                mTestSaveFields = NO
                Exit Function
            End If
        End If
    End If
    If (ilCtrlNo = VEHICLEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smSave(1), "", "Vehicle must be specified", tmCtrls(VEHICLEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = VEHICLEINDEX
            End If
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    '1/19/10:  Check that the name is not a package name
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If StrComp(Trim$(tgMVef(ilVef).sName), smSave(1), vbTextCompare) = 0 Then
            If tgMVef(ilVef).sType = "G" Then
                If (ilState And SHOWMSG) = SHOWMSG Then
                    ilRet = MsgBox("Enter a different Package Name as a Sport Vehicle Name is not allowed", vbOKOnly + vbExclamation, "Incomplete")
                End If
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = VEHICLEINDEX
                End If
                mTestSaveFields = NO
                Exit Function
            End If
        End If
    Next ilVef
    If (ilCtrlNo = DAYPARTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smSave(2), "", "Daypart must be specified", tmCtrls(DAYPARTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = DAYPARTINDEX
            End If
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If Not Contract.grdCPMBuy.Visible Then
        If (ilCtrlNo = MKTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedStr(smSave(3), "", "Market must be specified", tmCtrls(MKTINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = MKTINDEX
                End If
                mTestSaveFields = NO
                Exit Function
            End If
        End If
    End If
    mTestSaveFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGp3Branch                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      Vehicle group and process      *
'*                      communication back from        *
'*                      invoice sorting                *
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
Private Function mVehGp3Branch() As Integer
'
'   ilRet = mVehGp3Branch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcDropDown, lbcVehGp3, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mVehGp3Branch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(INVOICESORTLIST)) Then
    '    imDoubleClickName = False
    '    mVehGp3Branch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'ilEnable = cbcSelect.Enabled
    'cbcSelect.Enabled = False
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "H"
    igMNmCallSource = CALLSOURCEVEHICLE
    If edcDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'Traffic!edcLinkSrceHelpMsg.Text = ""
    If igTestSystem Then
        slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName & "\3"
    Else
        slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName & "\3"
    End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Traffic.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Traffic.Enabled = True
    'Traffic!edcLinkSrceHelpMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'cbcSelect.Enabled = ilEnable

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mVehGp3Branch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcVehGp3.Clear
        smVehGp3CodeTag = ""
        mVehGp3Pop
        If imTerminate Then
            mVehGp3Branch = False
            Exit Function
        End If
        gFindMatch sgMNmName, imMktAdj, lbcVehGp3
        sgMNmName = ""
        If gLastFound(lbcVehGp3) > 0 Then
            imVehGp3ChgMode = True
            lbcVehGp3.ListIndex = gLastFound(lbcVehGp3)
            edcDropDown.Text = lbcVehGp3.List(lbcVehGp3.ListIndex)
            imVehGp3ChgMode = False
            mVehGp3Branch = False
            'mSpecSetChg imBoxNo, True
        Else
            imVehGp3ChgMode = True
            If imMktAdj = 0 Then
                lbcVehGp3.ListIndex = 0
                edcDropDown.Text = lbcVehGp3.List(0)
            Else
                lbcVehGp3.ListIndex = 1
                edcDropDown.Text = lbcVehGp3.List(1)
            End If
            imVehGp3ChgMode = False
            'mSpecSetChg imBoxNo, True
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
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
'*      Procedure Name:mVehGp3Pop                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Vehicle Group Code    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mVehGp3Pop()
'
'   mVehGpPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String

    ilIndex = lbcVehGp3.ListIndex
    If ilIndex > 1 Then
        slName = lbcVehGp3.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gIMoveListBox(Vehicle, lbcVehGp, tmVehGpCode(), smVehGpCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gPopMnfPlusFieldsBox(CPackage, lbcVehGp3, tmVehGp3Code(), smVehGp3CodeTag, "H3")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehGp3PopErr
        gCPErrorMsg ilRet, "mVehGp3Pop (gPopMnfPlusFieldsBox)", CPackage
        On Error GoTo 0
        If tgSpf.sMktBase = "Y" Then
            ReDim tmTempVehGp3Code(LBound(tmVehGp3Code) To UBound(tmVehGp3Code)) As SORTCODE
            For ilLoop = LBound(tmVehGp3Code) To UBound(tmVehGp3Code) - 1 Step 1
                tmTempVehGp3Code(ilLoop) = tmVehGp3Code(ilLoop)
            Next ilLoop
            ReDim tmVehGp3Code(0 To 0) As SORTCODE
            For ilLoop = 0 To UBound(tmTempVehGp3Code) - 1 Step 1 'lbcVehGpCode.ListCount - 1 Step 1
                slNameCode = tmTempVehGp3Code(ilLoop).sKey   'lbcVehGpCode.List(ilVef)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                For ilIndex = 0 To UBound(igCntrMktCode) - 1 Step 1
                    If Val(slCode) = igCntrMktCode(ilIndex) Then
                        tmVehGp3Code(UBound(tmVehGp3Code)) = tmTempVehGp3Code(ilLoop)
                        ReDim Preserve tmVehGp3Code(0 To UBound(tmVehGp3Code) + 1) As SORTCODE
                        Exit For
                    End If
                Next ilIndex
            Next ilLoop
            lbcVehGp3.Clear
            For ilLoop = 0 To UBound(tmVehGp3Code) - 1 Step 1
                slNameCode = tmVehGp3Code(ilLoop).sKey    'lbcMster.List(ilLoop)
                ilRet = gParseItem(slNameCode, 1, "\", slNameCode)
                slNameCode = Trim$(slNameCode)
                lbcVehGp3.AddItem slNameCode  'Add ID to list box
            Next ilLoop
            imVehGp3ChgMode = True
            If ilIndex >= 0 Then
                gFindMatch slName, 0, lbcVehGp3
                If gLastFound(lbcVehGp3) >= 0 Then
                    lbcVehGp3.ListIndex = gLastFound(lbcVehGp3)
                Else
                    lbcVehGp3.ListIndex = -1
                End If
            Else
                lbcVehGp3.ListIndex = ilIndex
            End If

        Else
            lbcVehGp3.AddItem "[None]", 0  'Force as first item on list
            lbcVehGp3.AddItem "[New]", 0  'Force as first item on list
            imVehGp3ChgMode = True
            If ilIndex > 1 Then
                gFindMatch slName, 2, lbcVehGp3
                If gLastFound(lbcVehGp3) >= 2 Then
                    lbcVehGp3.ListIndex = gLastFound(lbcVehGp3)
                Else
                    lbcVehGp3.ListIndex = -1
                End If
            Else
                lbcVehGp3.ListIndex = ilIndex
            End If
            imVehGp3ChgMode = False
        End If
    End If
    Exit Sub
mVehGp3PopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

Private Sub lbcVehicle_Scroll()
    pbcLbcVehicle_Paint
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcLbcLines_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilLinesEnd As Integer
    Dim ilField As Integer
    Dim llWidth As Long
    Dim slFields(0 To 4) As String
    Dim llFgColor As Long
    Dim ilFieldIndex As Integer
    
    ilLinesEnd = lbcLines.TopIndex + lbcLines.height \ fgListHtArial825
    If ilLinesEnd > lbcLines.ListCount Then
        ilLinesEnd = lbcLines.ListCount
    End If
    If lbcLines.ListCount <= lbcLines.height \ fgListHtArial825 Then
        llWidth = lbcLines.Width - 30
    Else
        llWidth = lbcLines.Width - igScrollBarWidth - 30
    End If
    pbcLbcLines.Width = llWidth
    pbcLbcLines.Cls
    llFgColor = pbcLbcLines.ForeColor
    For ilLoop = lbcLines.TopIndex To ilLinesEnd - 1 Step 1
        pbcLbcLines.ForeColor = llFgColor
        If lbcLines.MultiSelect = 0 Then
            If lbcLines.ListIndex = ilLoop Then
                gPaintArea pbcLbcLines, CSng(0), CSng((ilLoop - lbcLines.TopIndex) * fgListHtArial825), CSng(pbcLbcLines.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcLines.ForeColor = vbWhite
            End If
        Else
            If lbcLines.Selected(ilLoop) Then
                gPaintArea pbcLbcLines, CSng(0), CSng((ilLoop - lbcLines.TopIndex) * fgListHtArial825), CSng(pbcLbcLines.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcLines.ForeColor = vbWhite
            End If
        End If
        slStr = lbcLines.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = imLBCtrls To UBound(imListField) - 1 Step 1
            pbcLbcLines.CurrentX = imListField(ilField)
            pbcLbcLines.CurrentY = (ilLoop - lbcLines.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            ilFieldIndex = ilFieldIndex + 1
            gAdjShowLen pbcLbcLines, slStr, imListField(ilField + 1) - imListField(ilField)
            pbcLbcLines.Print slStr
        Next ilField
        pbcLbcLines.ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub pbcLbcVehicle_Paint()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilField                       slFields                                                *
'******************************************************************************************

    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilLinesEnd As Integer
    Dim llWidth As Long
    Dim llFgColor As Long

    If imINPBCPaint Then
        Exit Sub
    End If
    imINPBCPaint = True
    pbcLbcVehicle.Move lbcVehicle.Left + 15, lbcVehicle.Top + 15, pbcLbcVehicle.Width - 30, lbcVehicle.height - 30 '2115, 1560
    ilLinesEnd = lbcVehicle.TopIndex + lbcVehicle.height \ fgListHtArial825
    If ilLinesEnd > lbcVehicle.ListCount Then
        ilLinesEnd = lbcVehicle.ListCount
    End If
    If lbcVehicle.ListCount <= lbcVehicle.height \ fgListHtArial825 Then
        llWidth = lbcVehicle.Width - 30
    Else
        llWidth = lbcVehicle.Width - igScrollBarWidth - 30
    End If
    pbcLbcVehicle.Width = llWidth
    pbcLbcVehicle.Cls
    llFgColor = pbcLbcVehicle.ForeColor
    For ilLoop = lbcVehicle.TopIndex To ilLinesEnd - 1 Step 1
        pbcLbcVehicle.ForeColor = llFgColor
        If lbcVehicle.Selected(ilLoop) Then
            gPaintArea pbcLbcVehicle, CSng(0), CSng((ilLoop - lbcVehicle.TopIndex) * fgListHtArial825), CSng(pbcLbcVehicle.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
            pbcLbcVehicle.ForeColor = vbWhite
        Else
            If Contract.grdCPMBuy.Visible Then
                slStr = Left(lbcCPMVefSort.List(ilLoop), 1)
                pbcLbcVehicle.FontItalic = True
            Else
                slStr = tgPackVeh(ilLoop).sSortChar
                pbcLbcVehicle.FontItalic = False
            End If
            If Left$(slStr, 1) = "A" Then
                pbcLbcVehicle.ForeColor = vbBlue
            Else
                pbcLbcVehicle.ForeColor = vbBlack
            End If
        End If
        pbcLbcVehicle.CurrentX = 15
        pbcLbcVehicle.CurrentY = (ilLoop - lbcVehicle.TopIndex) * fgListHtArial825 + 15
        slStr = lbcVehicle.List(ilLoop)
        gAdjShowLen pbcLbcVehicle, slStr, pbcLbcVehicle.Width
        pbcLbcVehicle.Print slStr
        pbcLbcVehicle.ForeColor = llFgColor
    Next ilLoop
    imINPBCPaint = False
End Sub

Private Sub pbcPackage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer

    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                If (imSelectedIndex > 0) And ((ilBox = VEHICLEINDEX) Or (ilBox = INVOICEINDEX)) Then
                    If imInPast Then
                        Beep
                        Exit Sub
                    End If
                    If (ilBox = VEHICLEINDEX) And (smAlterName = "N") Then
                        Beep
                        Exit Sub
                    End If
                End If
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus imBoxNo
End Sub
Private Sub pbcPackage_Paint(Index As Integer)
    Dim ilBox As Integer

    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        pbcPackage(imPbcIndex).CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcPackage(imPbcIndex).CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY + imAdjY
        pbcPackage(imPbcIndex).Print tmCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    If imRetBranch = True Then 'second gotfocus-ignore
        Exit Sub
    End If
    If GetFocus() <> pbcSTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to left
    If imBoxNo = DAYPARTINDEX Then
        If mDPBranch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo = MKTINDEX) Then
        If mVehGp3Branch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= UBound(tmCtrls)) Then
        If (imBoxNo <> INVOICEINDEX) Or (Not cbcSelect.Enabled) Then
            If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
    End If
    Select Case imBoxNo
        Case -1
            imTabDirection = 0  'Set-Left to right
            If (imSelectedIndex = 0) And (cbcSelect.Text = "[New]") Then
                If imPbcIndex = 1 Then
                    ilBox = VEHICLEINDEX
                Else
                    ilBox = 1
                End If
                mSetCommands
            Else
                'mSetChg 1
                tmCtrls(INVOICEINDEX).iChg = True
                If (imInPast) Or (smAlterName = "N") Then
                    ilBox = 3
                Else
                    ilBox = 2
                End If
            End If
        Case VEHICLEINDEX
            If imPbcIndex = 1 Then
                mSetShow imBoxNo
                imBoxNo = -1
                If cbcSelect.Enabled Then
                    cbcSelect.SetFocus
                    Exit Sub
                End If
                cmcCancel.SetFocus
                Exit Sub
            Else
                ilBox = imBoxNo - 1
            End If
        Case DAYPARTINDEX
            If imInPast Then
                mSetShow imBoxNo
                imBoxNo = -1
                If cbcSelect.Enabled Then
                    cbcSelect.SetFocus
                    Exit Sub
                End If
                cmcCancel.SetFocus
                Exit Sub
            Else
                If smAlterName = "N" Then
                    ilBox = imBoxNo - 2
                Else
                    ilBox = imBoxNo - 1
                End If
            End If
        Case 1 'Name (first control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
                Exit Sub
            End If
            cmcCancel.SetFocus
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
    If imRetBranch = True Then 'second gotfocus-ignore
        'imRetBranch = False
        Exit Sub
    End If
    If GetFocus() <> pbcTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = 0 'Set- Left to right
    If imBoxNo = DAYPARTINDEX Then
        If mDPBranch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo = MKTINDEX) Then
        If mVehGp3Branch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= UBound(tmCtrls)) Then
        If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    Select Case imBoxNo
        Case -1
            imTabDirection = -1  'Set-Right to left
            If imPbcIndex = 1 Then
                ilBox = DAYPARTINDEX
            Else
                ilBox = UBound(tmCtrls)
            End If
        Case UBound(tmCtrls) 'Suppress (last control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            'If (cmcUpdate.Enabled) And (igANmCallSource = CALLNONE) Then
            '    cmcUpdate.SetFocus
            'Else
                cmcDone.SetFocus
            'End If
            Exit Sub
        Case INVOICEINDEX
            If smAlterName = "N" Then
                ilBox = imBoxNo + 2
            Else
                ilBox = imBoxNo + 1
            End If
        Case DAYPARTINDEX
            If imPbcIndex = 1 Then
                mSetShow imBoxNo
                imBoxNo = -1
                'If (cmcUpdate.Enabled) And (igANmCallSource = CALLNONE) Then
                '    cmcUpdate.SetFocus
                'Else
                    cmcDone.SetFocus
                'End If
                Exit Sub
            Else
                ilBox = imBoxNo + 1
            End If
        Case Else
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcType_GotFocus()
    'pbcClickFocus.SetFocus
End Sub
Private Sub pbcType_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("R") Or (KeyAscii = Asc("r")) Then
        If tgSpf.sCPkAired = "Y" Then
            If imSave(1) <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSave(1) = 0
            pbcType_Paint
        End If
    ElseIf KeyAscii = Asc("V") Or (KeyAscii = Asc("v")) Then
        If tgSpf.sCPkOrdered = "Y" Then
            If imSave(1) <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSave(1) = 1
            pbcType_Paint
        End If
    ElseIf KeyAscii = Asc("E") Or (KeyAscii = Asc("e")) Then
        If tgSpf.sCPkEqual = "Y" Then
            If imSave(1) <> 2 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSave(1) = 2
            pbcType_Paint
        End If
    End If
    If KeyAscii = Asc(" ") Then
        If imSave(1) = 0 Then
            If tgSpf.sCPkOrdered = "Y" Then
                tmCtrls(imBoxNo).iChg = True
                imSave(1) = 1
                pbcType_Paint
            ElseIf tgSpf.sCPkEqual = "Y" Then
                tmCtrls(imBoxNo).iChg = True
                imSave(1) = 2
                pbcType_Paint
            End If
        ElseIf imSave(1) = 1 Then
            If tgSpf.sCPkEqual = "Y" Then
                tmCtrls(imBoxNo).iChg = True
                imSave(1) = 2
                pbcType_Paint
            ElseIf tgSpf.sCPkAired = "Y" Then
                tmCtrls(imBoxNo).iChg = True
                imSave(1) = 0
                pbcType_Paint
            End If
        ElseIf imSave(1) = 2 Then
            If tgSpf.sCPkAired = "Y" Then
                tmCtrls(imBoxNo).iChg = True
                imSave(1) = 0
                pbcType_Paint
            ElseIf tgSpf.sCPkOrdered = "Y" Then
                tmCtrls(imBoxNo).iChg = True
                imSave(1) = 1
                pbcType_Paint
            End If
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imSave(1) = 0 Then
        If tgSpf.sCPkOrdered = "Y" Then
            tmCtrls(imBoxNo).iChg = True
            imSave(1) = 1
        ElseIf tgSpf.sCPkEqual = "Y" Then
            tmCtrls(imBoxNo).iChg = True
            imSave(1) = 2
        End If
    ElseIf imSave(1) = 1 Then
        If tgSpf.sCPkEqual = "Y" Then
            tmCtrls(imBoxNo).iChg = True
            imSave(1) = 2
        ElseIf tgSpf.sCPkAired = "Y" Then
            tmCtrls(imBoxNo).iChg = True
            imSave(1) = 0
        End If
    ElseIf imSave(1) = 1 Then
        If tgSpf.sCPkAired = "Y" Then
            tmCtrls(imBoxNo).iChg = True
            imSave(1) = 0
        ElseIf tgSpf.sCPkOrdered = "Y" Then
            tmCtrls(imBoxNo).iChg = True
            imSave(1) = 1
        End If
    End If
    pbcType_Paint
    mSetCommands
End Sub
Private Sub pbcType_Paint()
    pbcType.Cls
    pbcType.CurrentX = fgBoxInsetX
    pbcType.CurrentY = 0 'fgBoxInsetY
    If imSave(1) = 0 Then
        pbcType.Print "Real"
    ElseIf imSave(1) = 1 Then
        pbcType.Print "Virtual" '"Generate"
    ElseIf imSave(1) = 2 Then
        pbcType.Print "Equal"
    Else
        pbcType.Print "   "
    End If
End Sub
Private Sub plcLines_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imBoxNo
        Case DAYPARTINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcDaypart, edcDropDown, imChgMode, imLbcArrowSetting
        Case MKTINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcVehGp3, edcDropDown, imVehGp3ChgMode, imLbcArrowSetting
    End Select
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Package Lines"
End Sub

Private Function mCPMPop(ilPkgIndex As Integer) As Integer
    Dim llRow As Long
    Dim slPkgID As String
    Dim slStr As String
    Dim ilSelectIndex As Integer
    Dim ilLoop As Integer
    
    ilSelectIndex = -1

    
    If ilPkgIndex > 0 Then
        llRow = cbcSelect.ItemData(ilPkgIndex)
        slPkgID = Contract.grdCPMBuy.TextMatrix(llRow, CPMID)
        For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
            If Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = slPkgID Then
                slStr = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMID))  'ID
                slStr = slStr & "|" & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE))   'Vehicle name
                slStr = slStr & "|" & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART))   'Daypart name
                slStr = slStr & "|" & ""    'Spot Length
                slStr = slStr & "|" & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) & "-" & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE))
                lbcLines.AddItem slStr
                lbcLines.ItemData(lbcLines.NewIndex) = llRow
                ilSelectIndex = lbcLines.ListCount - 1
            End If
        Next llRow
        llRow = cbcSelect.ItemData(ilPkgIndex)
        smSave(1) = Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE)
        smOrigVehName = Trim$(smSave(1))
        gSetShow pbcPackage(imPbcIndex), smSave(1), tmCtrls(VEHICLEINDEX)
        smSave(2) = Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART)
        smOrigDPName = Trim$(smSave(2))
        gSetShow pbcPackage(imPbcIndex), smSave(2), tmCtrls(DAYPARTINDEX)
        smSave(3) = ""
        smOrigVehGp3 = smSave(3)
        gSetShow pbcPackage(imPbcIndex), smSave(3), tmCtrls(MKTINDEX)
        imSave(1) = 1   'Order
        mSetShow INVOICEINDEX
    End If
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") And ((Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) <> "P") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) <> "H")) Then
            If Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID)) = "" Or Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID)) = "0" Then
                If Left$(Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMPRICETYPE)), 1) <> "F" Then
                    slStr = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMID))  'ID
                    slStr = slStr & "|" & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE))   'Vehicle name
                    slStr = slStr & "|" & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART))   'Daypart name
                    slStr = slStr & "|" & ""    'Spot Length
                    slStr = slStr & "|" & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) & "-" & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE))
                    lbcLines.AddItem slStr
                    lbcLines.ItemData(lbcLines.NewIndex) = llRow
                End If
            End If
        End If
    Next llRow
    For ilLoop = 0 To ilSelectIndex Step 1
        lbcLines.Selected(ilLoop) = True
    Next ilLoop
    pbcLbcLines_Paint
    mCPMPop = True
End Function

Private Function mSaveCPM(ilPkgIndex) As Boolean
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilFound As Integer
    Dim slDate As String
    Dim slLnLLD As String
    Dim llUnbilledDate As Long
    Dim llPkgRow As Long
    Dim slPkgID As String
    Dim ilVefCode As Integer
    Dim ilRdfCode As Integer
    Dim ilIDNo As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slImpressions As String
    Dim slCPM As String
    Dim slTotalCost As String
    Dim ilCol As Integer
    '10929
    Dim ilMatchingLength As Integer
    
    Contract.grdCPMBuy.Redraw = False
    mSaveCPM = True
    ilFound = False
    For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
        If lbcLines.Selected(ilLoop) Then
            ilFound = True
            Exit For
        End If
    Next ilLoop
    If Not ilFound Then
        If ilPkgIndex > 0 Then
            llPkgRow = cbcSelect.ItemData(ilPkgIndex)
            If Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMPCFCODE) = "" Then
                'Remove package and set hidden lines as standard lines
                mRemovePkg llPkgRow
            Else
                'Test if date in future
                slLnLLD = Format$(gNow(), "m/d/yy")
                '3/14/13: Use either unbilled date or todays date which is oldest
                If tgUrf(0).sChgCntr = "I" Then
                    If (tgChfCntr.sBillCycle = "C") Then    'Or (tgChfCntr.sBillCycle = "D") Then
                        gUnpackDate tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1), slDate
                        slDate = gObtainEndCal(slDate)
                    ElseIf (tgChfCntr.sBillCycle = "W") Then
                        gUnpackDate tgSaf(0).iBLastWeeklyDate(0), tgSaf(0).iBLastWeeklyDate(1), slDate
                        slDate = gObtainNextSunday(slDate)
                    Else
                        gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slDate
                        slDate = gObtainEndStd(slDate)
                    End If
                    llUnbilledDate = gDateValue(slDate) + 1
                    If (llUnbilledDate > 0) And (llUnbilledDate < gDateValue(slLnLLD)) Then
                        slLnLLD = Format$(llUnbilledDate, "m/d/yy")
                    End If
                    If gDateValue(Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMSTARTDATE)) > gDateValue(slLnLLD) Then
                        'Remove Package
                        mRemovePkg llPkgRow
                    End If
                End If
            End If
        End If
        imCPMChg = True
        mSetFGColorForRows
        Contract.grdCPMBuy.Redraw = False
        mComputeTotals
        Screen.MousePointer = vbDefault
        Contract.grdCPMBuy.Redraw = True
        mTerminate
        Exit Function
    End If
    imSave(1) = 1
    tmCtrls(INVOICEINDEX).iReq = False
    tmCtrls(MKTINDEX).iReq = False
    If mTestSaveFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        Beep
        mEnableBox imBoxNo
        mSaveCPM = False
        Exit Function
    End If
    If ilPkgIndex > 0 Then
        llPkgRow = cbcSelect.ItemData(ilPkgIndex)
        slPkgID = Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMID)
        'Change package plus add new lines
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMVEHICLE) = smSave(1)
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMDAYPART) = smSave(2)
        'De-package lines
        For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = slPkgID) And (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H") Then
                Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE))
                Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = ""
                Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "S"
                Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = True
                Contract.grdCPMBuy.rowHeight(llRow) = fgBoxGridH
            End If
        Next llRow
        
        'Package lines
        For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
            If lbcLines.Selected(ilLoop) Then
                llRow = lbcLines.ItemData(ilLoop)
                Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) = smHiddenSpace & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE))
                Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = slPkgID
                Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H"
                If Not bgHideHiddenLines Then
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "True"
                Else
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "False"
                End If
                If Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "True" Then
                    Contract.grdCPMBuy.rowHeight(llRow) = fgBoxGridH
                Else
                    Contract.grdCPMBuy.rowHeight(llRow) = 0
                End If
            End If
        Next ilLoop
    Else
        ilIDNo = 0
        For ilLoop = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
            '10929 short circuit
            If (Contract.grdCPMBuy.TextMatrix(ilLoop, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(ilLoop, CPMDAYPART) <> "") Then
                If Val(Contract.grdCPMBuy.TextMatrix(ilLoop, CPMID)) > ilIDNo Then
                    ilIDNo = Val(Contract.grdCPMBuy.TextMatrix(ilLoop, CPMID))
                End If
            Else
                Exit For
            End If
        Next ilLoop
        slPkgID = ilIDNo + 1
        '10929
        ilMatchingLength = -1
        For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
            If lbcLines.Selected(ilLoop) Then
                llRow = lbcLines.ItemData(ilLoop)
                Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) = smHiddenSpace & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE))
                Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = slPkgID
                Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H"
                If Not bgHideHiddenLines Then
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "True"
                Else
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "False"
                End If
                If Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "True" Then
                    Contract.grdCPMBuy.rowHeight(llRow) = fgBoxGridH
                Else
                    Contract.grdCPMBuy.rowHeight(llRow) = 0
                End If
                '10929 get length here
                If ilMatchingLength = -1 Then
                    ilMatchingLength = Contract.grdCPMBuy.TextMatrix(llRow, CPMLENGTH)
                ElseIf ilMatchingLength > 0 Then
                    If ilMatchingLength <> Contract.grdCPMBuy.TextMatrix(llRow, CPMLENGTH) Then
                        ilMatchingLength = 0
                    End If
                End If
            End If
        Next ilLoop

        'New package
        gFindMatch smSave(1), 0, lbcVehicle
        '10929 allow 0
        If gLastFound(lbcVehicle) >= 0 Then
            ilVefCode = lbcVehicle.ItemData(gLastFound(lbcVehicle))
        Else
            ilVefCode = 0
        End If
        gFindMatch smSave(2), 1, lbcDaypart
        If gLastFound(lbcDaypart) > 0 Then
            ilRdfCode = lbcDaypart.ItemData(gLastFound(lbcDaypart))
        Else
            ilRdfCode = 0
        End If
        'Find first blank line
        llRow = Contract.grdCPMBuy.rows - 1
        Do While Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) = ""
            llRow = llRow - 1
        Loop
        If llRow < Contract.grdCPMBuy.FixedRows Then
            Contract.grdCPMBuy.Row = Contract.grdCPMBuy.rows - 1
            Contract.grdCPMBuy.AddItem ""
            Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row + 1) = fgBoxGridH
            Contract.grdCPMBuy.Row = Contract.grdCPMBuy.Row + 1
            For ilCol = 0 To Contract.grdCPMBuy.cols - 1 Step 1
                Contract.grdCPMBuy.ColAlignment(ilCol) = flexAlignLeftCenter
            Next ilCol
        Else
            '10953 + 1 causes error
'            Contract.grdCPMBuy.Row = llRow + 1
            If Contract.grdCPMBuy.rows = llRow + 1 Then
               Contract.grdCPMBuy.rows = Contract.grdCPMBuy.rows + 1
            End If
            Contract.grdCPMBuy.Row = llRow + 1
        End If
        llPkgRow = Contract.grdCPMBuy.Row
        Contract.grdCPMBuy.Col = CPMEXPAND
        Contract.grdCPMBuy.CellFontName = "Monotype Sorts"
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMID) = slPkgID
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMVEHICLE) = smSave(1)
        Contract.grdCPMBuy.Col = CPMVEHICLE
        Contract.grdCPMBuy.CellFontItalic = True
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMDAYPART) = smSave(2)
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMLNTYPE) = "P"
        'Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMCOPYTYPE) = ""
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMPOSITION) = ""
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMPRIORITY) = ""
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMPRICETYPE) = "CPM"
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMPODTARGET) = ""
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMSTARTDATE) = ""  'Format$(slStartDate, sgShowDateForm)
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMENDDATE) = ""    'Format$(slEndDate, sgShowDateForm)
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMIMPRESSIONS) = ""    'slImpressions
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMIMPRESSIONCAP) = ""    'slImpressions
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMCPMVALUE) = ""   'slCPM
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMCOST) = ""   'slTotalCost
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMCOMMENT) = ""
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMVEFCODE) = ilVefCode
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMRDFCODE) = ilRdfCode
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMCHFCODE) = tgChfCntr.lCode
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMPACKAGEID) = "0"
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMPCFCODE) = "" 'New
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMSHOWROW) = "True"
        '10929 set the length if found all hidden the same, or set to 0
        If ilMatchingLength > -1 Then
            Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMLENGTH) = ilMatchingLength
        End If
        If Not bgHideHiddenLines Then
            Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMEXPAND) = ""
        Else
            If Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMSHOWROW) Then
                Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMEXPAND) = "s"
            Else
                Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMEXPAND) = "t"
            End If
        End If
        Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMSORT) = ""
    End If
    slImpressions = "0"
    slTotalCost = "0.00"
    slStartDate = ""
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = slPkgID) And (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H") Then
            slImpressions = gAddStr(slImpressions, Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS))
            slTotalCost = gAddStr(slTotalCost, Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST))
            If slStartDate = "" Then
                slStartDate = Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)
                slEndDate = Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)
            Else
                If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) < gDateValue(slStartDate) Then
                    slStartDate = Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)
                End If
                If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)) > gDateValue(slEndDate) Then
                    slEndDate = Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)
                End If
            End If
        End If
    Next llRow
    Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMIMPRESSIONS) = slImpressions
    Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMCOST) = slTotalCost
    Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMCPMVALUE) = gDivStr(gMulStr(slTotalCost, "1000"), slImpressions)
    Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMSTARTDATE) = Format$(slStartDate, sgShowDateForm)
    Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMENDDATE) = Format$(slEndDate, sgShowDateForm)
    
    'If Contract.grdCPMBuy.TextMatrix(0, CPMID) = "ID-Std" Then
    If imCPMSortMethod = 0 Then
        gCPMSort False
    Else
        gCPMSort True
    End If
    imCPMChg = True
    mSetFGColorForRows
    Contract.grdCPMBuy.Redraw = False
    mComputeTotals
    Screen.MousePointer = vbDefault
    Contract.grdCPMBuy.Redraw = True
    
End Function

Private Sub mRemovePkg(llPkgRow As Long)
    Dim llRow As Long
    Dim slPkgID As String
    
    slPkgID = Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMID)
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = slPkgID) And (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H") Then
            Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE))
            Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = ""
            Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "S"
        End If
    Next llRow
    Contract.grdCPMBuy.RemoveItem llPkgRow
End Sub


Private Sub mCPMVehPop()
    Dim ilVef As Integer
    Dim ilVpf As Integer
    Dim llRow As Long
    
    lbcCPMVefSort.Clear
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        ilVpf = gBinarySearchVpf(tgMVef(ilVef).iCode)
        If ilVpf <> -1 Then
            'If tgVpf(ilVpf).sGMedium = "M" Then
            If tgVpf(ilVpf).sGMedium = "P" Then
                If tgMVef(ilVef).sType = "P" Then
                    lbcCPMVefSort.AddItem "A" & Trim$(tgMVef(ilVef).sName)
                    lbcCPMVefSort.ItemData(lbcCPMVefSort.NewIndex) = tgMVef(ilVef).iCode
                Else
                    lbcCPMVefSort.AddItem "B" & Trim$(tgMVef(ilVef).sName)
                    lbcCPMVefSort.ItemData(lbcCPMVefSort.NewIndex) = tgMVef(ilVef).iCode
                End If
            End If
        End If
    Next ilVef
    'Add dynamic names
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "P") Then
                If (Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = "0") Or (Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE)) = "") Then
                    If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE) = "0") Or (Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE)) = "") Then
                        lbcCPMVefSort.AddItem "A" & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE))
                        lbcCPMVefSort.ItemData(lbcCPMVefSort.NewIndex) = 0
                    End If
                End If
            End If
        End If
    Next llRow
    
    lbcVehicle.Clear
    For ilVef = 0 To lbcCPMVefSort.ListCount - 1 Step 1
        lbcVehicle.AddItem Mid$(lbcCPMVefSort.List(ilVef), 2)
        lbcVehicle.ItemData(lbcVehicle.NewIndex) = lbcCPMVefSort.ItemData(ilVef)
    Next ilVef
End Sub
