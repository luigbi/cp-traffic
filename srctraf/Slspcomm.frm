VERSION 5.00
Begin VB.Form SlspComm 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5715
   ClientLeft      =   855
   ClientTop       =   1470
   ClientWidth     =   8265
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5715
   ScaleWidth      =   8265
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
      Left            =   3585
      TabIndex        =   1
      Top             =   150
      Width           =   3795
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7725
      Top             =   3420
   End
   Begin VB.PictureBox plcNum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   4575
      ScaleHeight     =   1140
      ScaleWidth      =   1095
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1755
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcNum 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1050
         Left            =   45
         Picture         =   "Slspcomm.frx":0000
         ScaleHeight     =   1050
         ScaleWidth      =   1020
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcNumOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   0
            Top             =   15
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imcNumInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   300
            Picture         =   "Slspcomm.frx":0B72
            Top             =   255
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox pbcArrow 
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
      Height          =   180
      Left            =   510
      Picture         =   "Slspcomm.frx":0E7C
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1095
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3915
      Visible         =   0   'False
      Width           =   2670
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
      Left            =   2040
      Picture         =   "Slspcomm.frx":1186
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   195
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
      Left            =   1020
      MaxLength       =   20
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox plcCalendar 
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
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   2295
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1755
      Visible         =   0   'False
      Width           =   1995
      Begin VB.CommandButton cmcCalUp 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1635
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalDn 
         Appearance      =   0  'Flat
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Slspcomm.frx":1280
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   240
            Left            =   510
            TabIndex        =   15
            Top             =   405
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
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
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   315
         TabIndex        =   12
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
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
      Left            =   7680
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4365
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
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
      Left            =   7725
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4110
      Visible         =   0   'False
      Width           =   525
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
      Left            =   7635
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2955
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcCreate 
      Appearance      =   0  'Flat
      Caption         =   "Cr&eate"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5130
      TabIndex        =   23
      Top             =   5250
      Width           =   1050
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
      Height          =   285
      Left            =   6225
      TabIndex        =   24
      Top             =   5250
      Width           =   1050
   End
   Begin VB.CommandButton cmcModel 
      Appearance      =   0  'Flat
      Caption         =   "&Model"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4065
      TabIndex        =   22
      Top             =   5250
      Width           =   1050
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
      Left            =   -15
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4560
      Width           =   75
   End
   Begin VB.CommandButton cmcSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   3075
      TabIndex        =   21
      Top             =   5250
      Width           =   945
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
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   17
      Top             =   4200
      Width           =   60
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
      Height          =   120
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   525
      Width           =   105
   End
   Begin VB.VScrollBar vbcComm 
      Height          =   4065
      LargeChange     =   18
      Left            =   7110
      Max             =   1
      Min             =   1
      TabIndex        =   18
      Top             =   720
      Value           =   1
      Width           =   240
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   1980
      TabIndex        =   20
      Top             =   5250
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
      Left            =   885
      TabIndex        =   19
      Top             =   5250
      Width           =   1050
   End
   Begin VB.PictureBox pbcComm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4080
      Left            =   810
      Picture         =   "Slspcomm.frx":409A
      ScaleHeight     =   4080
      ScaleWidth      =   6330
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   705
      Width           =   6330
      Begin VB.Label lacFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Height          =   240
         Left            =   15
         TabIndex        =   29
         Top             =   540
         Visible         =   0   'False
         Width           =   6300
      End
   End
   Begin VB.PictureBox plcComm 
      BackColor       =   &H00FFFFFF&
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
      Height          =   4215
      Left            =   750
      ScaleHeight     =   4155
      ScaleWidth      =   6600
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   630
      Width           =   6660
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   7665
      Picture         =   "Slspcomm.frx":79A34
      Top             =   5070
      Width           =   480
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   180
      Top             =   5145
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "SlspComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Slspcomm.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SlspComm.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Commission input screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim tmCtrls(0 To 5)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer   'Current event name Box
Dim imRowNo As Integer  'Current event row
Dim smSave() As String  '1=Vehicle Name; 2=Start Date; 3=End Date; 4=Under Goal %;
                        '5=Remnant Under Goal %
Dim smShow() As String  '1=Vehicle Name; 2=Start Date; 3=End Date; 4=Under Goal %;
                        '5=Remnant Under Goal %
Dim tmScf As SCF        'Scf record image
Dim tmScfSrchKey As LONGKEY0    'Scf key record image
Dim tmScfSrchKey1 As INTKEY0    'Scf key record image
Dim hmScf As Integer    'Sale Commission file handle
Dim imScfRecLen As Integer        'SCF record length
Dim imScfChg As Integer     'Indicates if field changed
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imChgMode As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imSlspSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imComboBoxIndex As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer
Dim imSettingValue As Integer
Dim smNowDate As String
'Calendar variables
Dim tmCDCtrls(0 To 7) As FIELDAREA  'Field area image
Dim imCalYear As Integer        'Month of displayed calendar
Dim imCalMonth As Integer       'Year of displayed calendar
Dim lmCalStartDate As Long      'Start date of displayed calendar
Dim lmCalEndDate As Long        'End date of displayed calendar
Dim imCalType As Integer        'Calendar type
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imDragShift As Integer

Dim tmUserVehicle() As SORTCODE
Dim smUserVehicleTag As String

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor

Const LBONE = 1

Const VEHINDEX = 1
Const SDATEINDEX = 2
Const EDATEINDEX = 3
Const GOALPVEHINDEX = 4
Const REMNANTPVEHINDEX = 5
Private Sub cbcSelect_Change()
    Dim ilRet As Integer
    If imChgMode = False Then
        imChgMode = True
        Screen.MousePointer = vbHourglass  'Wait
        If cbcSelect.Text <> "" Then
            gManLookAhead cbcSelect, imBSMode, imComboBoxIndex
        End If
        imSlspSelectedIndex = cbcSelect.ListIndex
        mClearCtrlFields
        ilRet = mReadRec()
        pbcComm.Cls
        mMoveRecToCtrl
        mInitShow
        mSetMinMax
        mSetCommands
        imChgMode = False
        Screen.MousePointer = vbDefault    'Default
    End If
End Sub
Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcSelect_GotFocus()

    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocus Then
        imFirstFocus = False
    End If
    imComboBoxIndex = imSlspSelectedIndex
    gCtrlGotFocus cbcSelect
    Exit Sub
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
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcCreate_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim slStr As String
    Dim ilVef As Integer
    If imSlspSelectedIndex < 0 Then
        ilRet = MsgBox("Salesperson must be selected", vbOKOnly + vbExclamation, "Commission")
        Exit Sub
    End If
    mSetShow imBoxNo
    SlspCrte.Show vbModal
    If igModReturn = 1 Then
        For ilLoop = LBound(tgScfAdd) To UBound(tgScfAdd) - 1 Step 1
            ilUpper = UBound(smShow, 2)
            ReDim Preserve smShow(0 To 5, 0 To ilUpper + 1) As String 'Values shown in program area
            ReDim Preserve smSave(0 To 5, 0 To ilUpper + 1) As String 'Values saved (program name) in program area
            ReDim Preserve tgScfRec(0 To UBound(tgScfRec) + 1) As SCFREC
            mInitNew ilUpper + 1
            'Get Vehicle Name
            'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1 'Traffic!lbcAdvertiser.ListCount - 1 Step 1
            '    If tgMVef(ilVef).iCode = tgScfAdd(ilLoop).iVefCode Then
                ilVef = gBinarySearchVef(tgScfAdd(ilLoop).iVefCode)
                If ilVef <> -1 Then
                    smSave(1, ilUpper) = Trim$(tgMVef(ilVef).sName)
                    slStr = Trim$(smSave(1, ilUpper))
                    gSetShow pbcComm, slStr, tmCtrls(VEHINDEX)
                    smShow(VEHINDEX, ilUpper) = tmCtrls(VEHINDEX).sShow
            '        Exit For
                End If
            'Next ilVef
            'Get Start Date
            gUnpackDate tgScfAdd(ilLoop).iStartDate(0), tgScfAdd(ilLoop).iStartDate(1), smSave(2, ilUpper)
            slStr = Trim$(smSave(2, ilUpper))
            gSetShow pbcComm, slStr, tmCtrls(SDATEINDEX)
            smShow(SDATEINDEX, ilUpper) = tmCtrls(SDATEINDEX).sShow
            'Get End Date
            gUnpackDate tgScfAdd(ilLoop).iEndDate(0), tgScfAdd(ilLoop).iEndDate(1), smSave(3, ilUpper)
            slStr = Trim$(smSave(3, ilUpper))
            If slStr = "" Then
                slStr = "TFN"
            End If
            gSetShow pbcComm, slStr, tmCtrls(EDATEINDEX)
            smShow(EDATEINDEX, ilUpper) = tmCtrls(EDATEINDEX).sShow
            'Get Under %
            smSave(4, ilUpper) = gIntToStrDec(tgScfAdd(ilLoop).iUnderComm, 2)
            slStr = Trim$(smSave(4, ilUpper))
            gSetShow pbcComm, slStr, tmCtrls(GOALPVEHINDEX)
            smShow(GOALPVEHINDEX, ilUpper) = tmCtrls(GOALPVEHINDEX).sShow
            'Get Remnant Under %
            smSave(5, ilUpper) = gIntToStrDec(tgScfAdd(ilLoop).iRemUnderComm, 2)
            slStr = Trim$(smSave(5, ilUpper))
            gSetShow pbcComm, slStr, tmCtrls(REMNANTPVEHINDEX)
            smShow(REMNANTPVEHINDEX, ilUpper) = tmCtrls(REMNANTPVEHINDEX).sShow
        Next ilLoop
        imRowNo = -1
        imBoxNo = -1
        imSettingValue = True
        vbcComm.Min = LBONE 'LBound(smShow, 2)
        imSettingValue = True
        If UBound(smShow, 2) - 1 <= vbcComm.LargeChange + 1 Then ' + 1 Then
            vbcComm.Max = LBONE 'LBound(smShow, 2)
        Else
            vbcComm.Max = UBound(smShow, 2) - vbcComm.LargeChange
        End If
        imSettingValue = True
        vbcComm.Value = vbcComm.Min
        imSettingValue = True
        pbcComm.Cls
        pbcComm_Paint
        imScfChg = True
        mSetCommands
    End If
End Sub
Private Sub cmcCreate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If mSaveRecChg(True) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case VEHINDEX
            lbcVehicle.Visible = Not lbcVehicle.Visible
        Case SDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case EDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case GOALPVEHINDEX
            plcNum.Visible = Not plcNum.Visible
        Case REMNANTPVEHINDEX
            plcNum.Visible = Not plcNum.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcModel_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim slStr As String
    Dim ilVef As Integer
    If imSlspSelectedIndex < 0 Then
        ilRet = MsgBox("Salesperson must be selected", vbOKOnly + vbExclamation, "Commission")
        Exit Sub
    End If
    mSetShow imBoxNo
    SlspMod.Show vbModal
    If igModReturn = 1 Then
        For ilLoop = LBound(tgScfAdd) To UBound(tgScfAdd) - 1 Step 1
            ilUpper = UBound(smShow, 2)
            ReDim Preserve smShow(0 To 5, 0 To ilUpper + 1) As String 'Values shown in program area
            ReDim Preserve smSave(0 To 5, 0 To ilUpper + 1) As String 'Values saved (program name) in program area
            ReDim Preserve tgScfRec(0 To UBound(tgScfRec) + 1) As SCFREC
            mInitNew ilUpper + 1
            'Get Vehicle Name
            'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1 'Traffic!lbcAdvertiser.ListCount - 1 Step 1
            '    If tgMVef(ilVef).iCode = tgScfAdd(ilLoop).iVefCode Then
                ilVef = gBinarySearchVef(tgScfAdd(ilLoop).iVefCode)
                If ilVef <> -1 Then
                    smSave(1, ilUpper) = Trim$(tgMVef(ilVef).sName)
                    slStr = Trim$(smSave(1, ilUpper))
                    gSetShow pbcComm, slStr, tmCtrls(VEHINDEX)
                    smShow(VEHINDEX, ilUpper) = tmCtrls(VEHINDEX).sShow
            '        Exit For
                End If
            'Next ilVef
            'Get Start Date
            gUnpackDate tgScfAdd(ilLoop).iStartDate(0), tgScfAdd(ilLoop).iStartDate(1), smSave(2, ilUpper)
            slStr = Trim$(smSave(2, ilUpper))
            gSetShow pbcComm, slStr, tmCtrls(SDATEINDEX)
            smShow(SDATEINDEX, ilUpper) = tmCtrls(SDATEINDEX).sShow
            'Get End Date
            gUnpackDate tgScfAdd(ilLoop).iEndDate(0), tgScfAdd(ilLoop).iEndDate(1), smSave(3, ilUpper)
            slStr = Trim$(smSave(3, ilUpper))
            If slStr = "" Then
                slStr = "TFN"
            End If
            gSetShow pbcComm, slStr, tmCtrls(EDATEINDEX)
            smShow(EDATEINDEX, ilUpper) = tmCtrls(EDATEINDEX).sShow
            'Get Under %
            smSave(4, ilUpper) = gIntToStrDec(tgScfAdd(ilLoop).iUnderComm, 2)
            slStr = Trim$(smSave(4, ilUpper))
            gSetShow pbcComm, slStr, tmCtrls(GOALPVEHINDEX)
            smShow(GOALPVEHINDEX, ilUpper) = tmCtrls(GOALPVEHINDEX).sShow
            'Get Remnant Under %
            smSave(5, ilUpper) = gIntToStrDec(tgScfAdd(ilLoop).iRemUnderComm, 2)
            slStr = Trim$(smSave(5, ilUpper))
            gSetShow pbcComm, slStr, tmCtrls(REMNANTPVEHINDEX)
            smShow(REMNANTPVEHINDEX, ilUpper) = tmCtrls(REMNANTPVEHINDEX).sShow
        Next ilLoop
        imRowNo = -1
        imBoxNo = -1
        imSettingValue = True
        vbcComm.Min = LBONE 'LBound(smShow, 2)
        imSettingValue = True
        If UBound(smShow, 2) - 1 <= vbcComm.LargeChange + 1 Then ' + 1 Then
            vbcComm.Max = LBONE 'LBound(smShow, 2)
        Else
            vbcComm.Max = UBound(smShow, 2) - vbcComm.LargeChange
        End If
        imSettingValue = True
        vbcComm.Value = vbcComm.Min
        imSettingValue = True
        pbcComm.Cls
        pbcComm_Paint
        imScfChg = True
        mSetCommands
    End If
End Sub
Private Sub cmcModel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcReport_Click()
    Dim slStr As String
    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = SLSPCOMMSJOB
    igRptType = 0
    ''Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "SlspComm^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType))
        Else
            slStr = "SlspComm^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "SlspComm^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    Else
    '        slStr = "SlspComm^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'SlspComm.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'SlspComm.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    sgCommandStr = slStr
    RptList.Show vbModal
    ''Screen.MousePointer = vbDefault    'Default
End Sub
Private Sub cmcReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcSave_Click()
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    ReDim tgScfDel(0 To 0) As SCFREC
    pbcComm.Cls
    pbcComm_Paint
    imBoxNo = -1    'Have to be after mSaveRecChg as it test imBoxNo = 1
    imScfChg = False
    mSetCommands
    pbcSTab.SetFocus
End Sub
Private Sub cmcSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Select Case imBoxNo
        Case VEHINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
        Case SDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case EDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case GOALPVEHINDEX
        Case REMNANTPVEHINDEX
    End Select
    imLbcArrowSetting = False
End Sub
Private Sub edcDropDown_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
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
    Dim ilKeyAscii As Integer
    ilKeyAscii = KeyAscii
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imBoxNo
        Case VEHINDEX
        Case SDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case EDATEINDEX
            'Disallow TFN for alternate
            If (Len(edcDropDown.Text) = edcDropDown.SelLength) Then
                If (KeyAscii = Asc("T")) Or (KeyAscii = Asc("t")) Then
                    edcDropDown.Text = "TFN"
                    edcDropDown.SelStart = 0
                    edcDropDown.SelLength = 3
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case GOALPVEHINDEX
            If Not mDropDownKeyPress(ilKeyAscii, False) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case REMNANTPVEHINDEX
            If Not mDropDownKeyPress(ilKeyAscii, False) Then
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo
            Case VEHINDEX
                If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
                    gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
                End If
            Case SDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
            Case EDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
            Case GOALPVEHINDEX
            Case REMNANTPVEHINDEX
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case imBoxNo
            Case VEHINDEX
            Case SDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
            Case EDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
            Case GOALPVEHINDEX
            Case REMNANTPVEHINDEX
        End Select
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True  'To get Alt J and Alt L keys
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(SLSPCOMMSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcComm.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcComm.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    Me.ZOrder 0 'Send to front
    SlspComm.Refresh
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
        plcCalendar.Visible = False
        plcNum.Visible = False
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

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        'Only expand first column
        'Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    End If
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    Erase tgScfRec
    Erase tgScfDel

    Erase tmUserVehicle
    smUserVehicleTag = ""

    btrExtClear hmScf   'Clear any previous extend operation
    ilRet = btrClose(hmScf)
    btrDestroy hmScf
    igJobShowing(SLSPCOMMSJOB) = False
    
    Set SlspComm = Nothing
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub imcHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub imcTrash_Click()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilUpperBound As Integer
    Dim ilRowNo As Integer
    Dim ilScf As Integer
    If (imRowNo < vbcComm.Value) Or (imRowNo > vbcComm.Value + vbcComm.LargeChange) Then
        Exit Sub
    End If
    ilRowNo = imRowNo
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
    ilUpperBound = UBound(smSave, 2)
    ilScf = ilRowNo
    If ilScf = ilUpperBound Then
        mInitNew ilScf
    Else
        If ilScf > 0 Then
            If tgScfRec(ilScf).iStatus = 1 Then
                tgScfDel(UBound(tgScfDel)).tScf = tgScfRec(ilScf).tScf
                tgScfDel(UBound(tgScfDel)).iStatus = tgScfRec(ilScf).iStatus
                tgScfDel(UBound(tgScfDel)).lRecPos = tgScfRec(ilScf).lRecPos
                ReDim Preserve tgScfDel(0 To UBound(tgScfDel) + 1) As SCFREC
            End If
            ilScf = ilRowNo
            'Remove record from tgRjf1Rec- Leave tgPjf2Rec
            For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
                tgScfRec(ilLoop) = tgScfRec(ilLoop + 1)
            Next ilLoop
            ReDim Preserve tgScfRec(0 To UBound(tgScfRec) - 1) As SCFREC
        End If
        For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
            For ilIndex = 1 To UBound(smSave, 1) Step 1
                smSave(ilIndex, ilLoop) = smSave(ilIndex, ilLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(smShow, 1) Step 1
                smShow(ilIndex, ilLoop) = smShow(ilIndex, ilLoop + 1)
            Next ilIndex
        Next ilLoop
        ilUpperBound = UBound(smSave, 2)
        ReDim Preserve smShow(0 To 5, 0 To ilUpperBound - 1) As String 'Values shown in program area
        ReDim Preserve smSave(0 To 5, 0 To ilUpperBound - 1) As String    'Values saved (program name) in program area
        imScfChg = True
    End If
    mSetCommands
    lacFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    imSettingValue = True
    vbcComm.Min = LBONE 'LBound(smShow, 2)
    imSettingValue = True
    If UBound(smShow, 2) - 1 <= vbcComm.LargeChange + 1 Then ' + 1 Then
        vbcComm.Max = LBONE 'LBound(smShow, 2)
    Else
        vbcComm.Max = UBound(smShow, 2) - vbcComm.LargeChange
    End If
    imSettingValue = True
    vbcComm.Value = vbcComm.Min
    imSettingValue = True
    pbcComm.Cls
    pbcComm_Paint
End Sub
Private Sub imcTrash_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash_Click
End Sub
Private Sub imcTrash_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    If State = vbEnter Then    'Enter drag over
        lacFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        imcTrash.Picture = IconTraf!imcTrashOpened.Picture
    ElseIf State = vbLeave Then
        lacFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    End If
End Sub
Private Sub imcTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub lbcVehicle_Click()
    gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Place box around calendar date *
'*                                                     *
'*******************************************************
Private Sub mBoxCalDate()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long

    slStr = edcDropDown.Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    lacDate.Caption = slDay
                    lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    lacDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            lacDate.Visible = False
        Else
            lacDate.Visible = False
        End If
    Else
        lacDate.Visible = False
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

    imScfChg = False
    lbcVehicle.ListIndex = -1
    ReDim tgScfRec(0 To 1) As SCFREC
    tgScfRec(0).iStatus = -1
    tgScfRec(0).lRecPos = 0
    tgScfRec(0).iDateChg = False
    tgScfRec(1).iStatus = -1
    tgScfRec(1).lRecPos = 0
    tgScfRec(1).iDateChg = False
    ReDim tgScfDel(0 To 0) As SCFREC
    tgScfDel(0).iStatus = -1
    tgScfDel(0).lRecPos = 0
    ReDim smShow(0 To 5, 0 To 1) As String 'Values shown in program area
    ReDim smSave(0 To 5, 0 To 1) As String 'Values saved (program name) in program area
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, 1) = ""
    Next ilLoop
    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, 1) = ""
    Next ilLoop
    vbcComm.Min = LBONE 'LBound(smShow, 2)
    imSettingValue = True
    vbcComm.Max = LBONE 'LBound(smShow, 2)
    imSettingValue = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDropDownKeyPress               *
'*                                                     *
'*             Created:5/11/94       By:D. Hannifan    *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                      in transaction section         *
'*******************************************************
Private Function mDropDownKeyPress(KeyAscii As Integer, ilNegAllowed As Integer) As Integer
    Dim ilPos As Integer
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flY As Single
    Dim flX As Single
    Dim slStr As String
    ilPos = InStr(edcDropDown.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                mDropDownKeyPress = False
                Exit Function
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If ilNegAllowed Then
        If (KeyAscii = KEYNEG) And ((Len(edcDropDown.Text) = 0) Or (Len(edcDropDown.Text) = edcDropDown.SelLength)) Then
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) And (KeyAscii <> KEYNEG) Then
                Beep
                mDropDownKeyPress = False
                Exit Function
            End If
        Else
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                Beep
                mDropDownKeyPress = False
                Exit Function
            End If
        End If
    Else
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
            Beep
            mDropDownKeyPress = False
            Exit Function
        End If
    End If
    slStr = edcDropDown.Text
    slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
    If gCompAbsNumberStr(slStr, "100.00") > 0 Then
        Beep
        mDropDownKeyPress = False
        Exit Function
    End If
    If KeyAscii <> KEYBACKSPACE Then
        Select Case Chr(KeyAscii)
            Case "7"
                ilRowNo = 1
                ilColNo = 1
            Case "8"
                ilRowNo = 1
                ilColNo = 2
            Case "9"
                ilRowNo = 1
                ilColNo = 3
            Case "4"
                ilRowNo = 2
                ilColNo = 1
            Case "5"
                ilRowNo = 2
                ilColNo = 2
            Case "6"
                ilRowNo = 2
                ilColNo = 3
            Case "1"
                ilRowNo = 3
                ilColNo = 1
            Case "2"
                ilRowNo = 3
                ilColNo = 2
            Case "3"
                ilRowNo = 3
                ilColNo = 3
            Case "0"
                ilRowNo = 4
                ilColNo = 1
            Case "00"   'Not possible
                ilRowNo = 4
                ilColNo = 2
            Case "."
                ilRowNo = 4
                ilColNo = 3
            Case "-"
                ilRowNo = 0
        End Select
        If ilRowNo > 0 Then
            flX = fgPadMinX + (ilColNo - 1) * fgPadDeltaX
            flY = fgPadMinY + (ilRowNo - 1) * fgPadDeltaY
            imcNumOutline.Move flX - 15, flY - 15
            imcNumOutline.Visible = True
        Else
            imcNumOutline.Visible = False
        End If
    Else
        imcNumOutline.Visible = False
    End If
    mDropDownKeyPress = True
    Exit Function
End Function
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
    Dim slStr As String
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
        Exit Sub
    End If
    If (imRowNo < vbcComm.Value) Or (imRowNo >= vbcComm.Value + vbcComm.LargeChange + 1) Then
        mSetShow ilBoxNo
        Exit Sub
    End If
    lacFrame.Move 0, tmCtrls(VEHINDEX).fBoxY + (imRowNo - vbcComm.Value) * (fgBoxGridH + 15) - 30
    lacFrame.Visible = True
    pbcArrow.Move pbcArrow.Left, plcComm.Top + tmCtrls(VEHINDEX).fBoxY + (imRowNo - vbcComm.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True

    Select Case ilBoxNo 'Branch on box type (control)
        Case VEHINDEX
            mVehPop
            If imTerminate Then
                Exit Sub
            End If
            lbcVehicle.Height = gListBoxHeight(lbcVehicle.ListCount, 10)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            If tgSpf.iVehLen <= 40 Then
                edcDropDown.MaxLength = tgSpf.iVehLen
            Else
                edcDropDown.MaxLength = 20
            End If
            gMoveTableCtrl pbcComm, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcComm.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If imRowNo - vbcComm.Value <= vbcComm.LargeChange \ 2 Then
                lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.Height
            End If
            imChgMode = True
            slStr = Trim$(smSave(1, imRowNo))
            If (slStr = "") And (imRowNo > 1) Then
                slStr = Trim$(smSave(1, imRowNo - 1))
            End If
            If slStr <> "" Then
                gFindMatch slStr, 0, lbcVehicle
                If gLastFound(lbcVehicle) > 0 Then
                    lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                Else
                    lbcVehicle.ListIndex = 0
                End If
            Else
                lbcVehicle.ListIndex = 0
            End If
            If lbcVehicle.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case SDATEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW + cmcDropDown.Width / 2
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcComm, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcComm.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcCalendar.Height < cmcDone.Top Then
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.Height
            End If
            slStr = Trim$(smSave(2, imRowNo))
            If (slStr = "") And (imRowNo > 1) Then
                If (smSave(1, imRowNo) = smSave(1, imRowNo - 1)) And (smSave(3, imRowNo - 1) <> "") Then
                    slStr = gObtainStartStd(Format$(gDateValue(Trim$(smSave(3, imRowNo - 1))) + 1, "m/d/yy"))
                End If
            End If
            If slStr = "" Then
                'Set to beginning of the year
                slStr = "1/15/" & Format$(gNow(), "yyyy")
                slStr = gObtainStartStd(slStr)
            End If
            edcDropDown.Text = slStr
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            If Trim$(smSave(2, imRowNo)) = "" Then
                pbcCalendar.Visible = True
            End If
            edcDropDown.SetFocus
        Case EDATEINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW + cmcDropDown.Width / 2
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcComm, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcComm.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcCalendar.Height < cmcDone.Top Then
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.Height
            End If
            slStr = Trim$(smSave(3, imRowNo))
            If slStr = "" Then
                slStr = "TFN"
            End If
            edcDropDown.Text = slStr
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            If Trim$(smSave(3, imRowNo)) = "" Then
                pbcCalendar.Visible = True
            End If
            edcDropDown.SetFocus
        Case GOALPVEHINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 6
            gMoveTableCtrl pbcComm, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcComm.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smSave(4, imRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case REMNANTPVEHINDEX
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 6
            gMoveTableCtrl pbcComm, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcComm.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smSave(5, imRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
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
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    imTerminate = False
    imFirstActivate = True
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165

    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    igJobShowing(SLSPCOMMSJOB) = True
    imFirstActivate = True
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    'SlspComm.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    mInitBox
    smNowDate = Format$(Now, "m/d/yy")
    gCenterForm SlspComm
    SlspComm.Show
    Screen.MousePointer = vbHourglass
    ReDim tgScfRec(0 To 1) As SCFREC
    tgScfRec(0).iStatus = -1
    tgScfRec(0).lRecPos = 0
    tgScfRec(1).iStatus = -1
    tgScfRec(1).lRecPos = 0
    ReDim tgScfDel(0 To 0) As SCFREC
    tgScfDel(0).iStatus = -1
    tgScfDel(0).lRecPos = 0
    ReDim smShow(0 To 5, 0 To 1) As String 'Values shown in program area
    ReDim smSave(0 To 5, 0 To 1) As String 'Values saved (program name) in program area
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, 1) = ""
    Next ilLoop
    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, 1) = ""
    Next ilLoop
'    mInitDDE
    'imcHelp.Picture = IconTraf!imcHelp.Picture
    imFirstFocus = True
    imDoubleClickName = False
    imLbcMouseDown = False
    imCalType = 0               'Standard type
    imBoxNo = -1                'Initialize current Box to N/A
    imRowNo = -1
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imChgMode = False
    imBSMode = False
    imBypassFocus = False
    imBypassSetting = False
    imSlspSelectedIndex = -1
    imScfChg = False
    hmScf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmScf, "", sgDBPath & "Scf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Scf.Btr)", SlspComm
    On Error GoTo 0
    imScfRecLen = Len(tmScf)
    cbcSelect.Clear 'Force list box to be populated
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    ilRet = gObtainVef()
    slDate = Format$(gNow(), "m/d/yy")
    gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
    If cbcSelect.ListCount > 0 Then
        cbcSelect.ListIndex = 0
    End If
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
    Dim ilLoop As Integer
    Dim llMax As Long
    Dim ilSpaceBetweenButtons As Integer
    Dim llAdjTop As Long

    flTextHeight = pbcComm.TextHeight("1") - 35
    'Position panel and picture areas with panel
    'plcSelect.Move 3555, 120
    cbcSelect.Move 3585, 150
    plcComm.Move 750, 630, pbcComm.Width + fgPanelAdj + vbcComm.Width, pbcComm.Height + fgPanelAdj
    pbcComm.Move plcComm.Left + fgBevelX, plcComm.Top + fgBevelY
    vbcComm.Move pbcComm.Left + pbcComm.Width - 15, pbcComm.Top
    'Vehicle
    gSetCtrl tmCtrls(VEHINDEX), 30, 375, 2655, fgBoxGridH
    'Start Date
    gSetCtrl tmCtrls(SDATEINDEX), 2700, tmCtrls(VEHINDEX).fBoxY, 990, fgBoxGridH
    'End Date
    gSetCtrl tmCtrls(EDATEINDEX), 3705, tmCtrls(VEHINDEX).fBoxY, 990, fgBoxGridH
    'Goal Percent
    gSetCtrl tmCtrls(GOALPVEHINDEX), 4710, tmCtrls(VEHINDEX).fBoxY, 780, fgBoxGridH
    'Remnant Percent
    gSetCtrl tmCtrls(REMNANTPVEHINDEX), 5505, tmCtrls(VEHINDEX).fBoxY, 780, fgBoxGridH

    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop



    llMax = 0
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        If ilLoop = VEHINDEX Then
            tmCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmCtrls(ilLoop).fBoxW)
            Do While (tmCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmCtrls(ilLoop).fBoxW = tmCtrls(ilLoop).fBoxW + 1
            Loop
        Else
            Do
                If tmCtrls(ilLoop).fBoxX < tmCtrls(ilLoop - 1).fBoxX + tmCtrls(ilLoop - 1).fBoxW + 15 Then
                    tmCtrls(ilLoop).fBoxX = tmCtrls(ilLoop).fBoxX + 15
                ElseIf tmCtrls(ilLoop).fBoxX > tmCtrls(ilLoop - 1).fBoxX + tmCtrls(ilLoop - 1).fBoxW + 15 Then
                    tmCtrls(ilLoop).fBoxX = tmCtrls(ilLoop).fBoxX - 15
                Else
                    Exit Do
                End If
            Loop
        End If
        If tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    pbcComm.Picture = LoadPicture("")
    pbcComm.Width = llMax
    plcComm.Width = llMax + vbcComm.Width + 2 * fgBevelX + 15
    lacFrame.Width = llMax - 15
    SlspComm.Width = plcComm.Width + 2 * plcComm.Left
    ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcDone.Left + cmcDone.Width))
    Do While ilSpaceBetweenButtons Mod 15 <> 0
        ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
    Loop
    cmcDone.Left = (SlspComm.Width - 6 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
    cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons
    cmcSave.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
    cmcModel.Left = cmcSave.Left + cmcSave.Width + ilSpaceBetweenButtons
    cmcCreate.Left = cmcModel.Left + cmcModel.Width + ilSpaceBetweenButtons
    cmcReport.Left = cmcCreate.Left + cmcCreate.Width + ilSpaceBetweenButtons
    cmcDone.Top = SlspComm.Height - (3 * cmcDone.Height) / 2 - 60
    cmcCancel.Top = cmcDone.Top
    cmcSave.Top = cmcDone.Top
    cmcModel.Top = cmcDone.Top
    cmcCreate.Top = cmcDone.Top
    cmcReport.Top = cmcDone.Top


    imcTrash.Top = cmcDone.Top + cmcDone.Height - imcTrash.Height
    imcTrash.Left = SlspComm.Width - (3 * imcTrash.Width) / 2

    llAdjTop = imcTrash.Top - plcComm.Top - 2 * fgBevelY - 120
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    llAdjTop = llAdjTop
    Do While plcComm.Top + llAdjTop + 2 * fgBevelY + 240 < imcTrash.Top
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    llAdjTop = llAdjTop
    plcComm.Height = llAdjTop + 2 * fgBevelY
    pbcComm.Left = plcComm.Left + fgBevelX
    pbcComm.Top = plcComm.Top + fgBevelY
    pbcComm.Height = plcComm.Height - 2 * fgBevelY

    vbcComm.Left = pbcComm.Left + pbcComm.Width + 15
    vbcComm.Top = pbcComm.Top
    vbcComm.Height = pbcComm.Height

    pbcTab.Top = SlspComm.Height
    pbcClickFocus.Top = SlspComm.Height
    cbcSelect.Left = plcComm.Left + plcComm.Width - cbcSelect.Width

End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNew                        *
'*                                                     *
'*             Created:9/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize values              *
'*                                                     *
'*******************************************************
Private Sub mInitNew(ilRowNo As Integer)
    Dim ilLoop As Integer

    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, ilRowNo) = ""
    Next ilLoop
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, ilRowNo) = ""
    Next ilLoop
    tgScfRec(ilRowNo).iStatus = 0
    tgScfRec(ilRowNo).lRecPos = 0
    tgScfRec(ilRowNo).iDateChg = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitShow                       *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mInitShow()
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim ilBoxNo As Integer
    For ilRowNo = 1 To UBound(smSave, 2) - 1 Step 1
        For ilBoxNo = VEHINDEX To REMNANTPVEHINDEX Step 1
            Select Case ilBoxNo
                Case VEHINDEX
                    slStr = Trim$(smSave(1, ilRowNo))
                    gSetShow pbcComm, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case SDATEINDEX
                    slStr = Trim$(smSave(2, ilRowNo))
                    gSetShow pbcComm, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case EDATEINDEX
                    slStr = Trim$(smSave(3, ilRowNo))
                    If slStr = "" Then
                        slStr = "TFN"
                    End If
                    gSetShow pbcComm, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case GOALPVEHINDEX
                    slStr = Trim$(smSave(4, ilRowNo))
                    gSetShow pbcComm, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
                Case REMNANTPVEHINDEX
                    slStr = Trim$(smSave(5, ilRowNo))
                    gSetShow pbcComm, slStr, tmCtrls(ilBoxNo)
                    smShow(ilBoxNo, ilRowNo) = tmCtrls(ilBoxNo).sShow
            End Select
        Next ilBoxNo
    Next ilRowNo
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec()
'
'   mMoveCtrlToRec
'   Where:
'
    Dim ilLoop As Integer
    Dim ilRowNo As Integer
    Dim slStr As String
    Dim llRowSDate As Long
    Dim llRowEDate As Long
    Dim llTstSDate As Long
    Dim llTstEDate As Long
    Dim ilPass As Integer
    Dim ilLp1 As Integer
    Dim ilRMonth As Integer
    Dim ilTMonth As Integer

    'Test and adjust overlapping dates
    'Pass one- test New, Pass two test changed
    For ilPass = 0 To 1 Step 1
        For ilRowNo = UBound(smSave, 2) - 1 To LBONE Step -1
            If ((tgScfRec(ilRowNo).iStatus = 0) And (ilPass = 0)) Or ((tgScfRec(ilRowNo).iStatus = 1) And (tgScfRec(ilRowNo).iDateChg = True) And (ilPass = 1)) Then
                llRowSDate = gDateValue(smSave(2, ilRowNo))
                If Trim$(smSave(3, ilRowNo)) <> "" Then
                    llRowEDate = gDateValue(smSave(3, ilRowNo))
                Else
                    llRowEDate = 999999999
                End If
                If smSave(3, ilRowNo) = "" Then
                    ilRMonth = 0
                Else
                    ilRMonth = Month(smSave(3, ilRowNo))
                End If
                For ilLoop = UBound(smSave, 2) - 1 To LBONE Step -1
                    If (StrComp(smSave(1, ilRowNo), smSave(1, ilLoop), 1) = 0) And (ilRowNo <> ilLoop) Then
                        llTstSDate = gDateValue(smSave(2, ilLoop))
                        If Trim$(smSave(3, ilLoop)) <> "" Then
                            llTstEDate = gDateValue(smSave(3, ilLoop))
                        Else
                            llTstEDate = 999999999
                        End If
                        If smSave(3, ilLoop) = "" Then
                            ilTMonth = 0
                        Else
                            ilTMonth = Month(smSave(3, ilLoop))
                        End If
                        If (llRowEDate >= llTstSDate) And (llTstEDate >= llRowSDate) Then
                            If (llRowSDate < llTstSDate) And (llRowEDate < llTstEDate) Then
                                'Truncate Row
                                smSave(2, ilLoop) = Format$(gDateValue(gObtainEndStd(smSave(3, ilRowNo))) + 1, "m/d/yy")
                                slStr = Trim$(smSave(2, ilLoop))
                                gSetShow pbcComm, slStr, tmCtrls(SDATEINDEX)
                                smShow(SDATEINDEX, ilLoop) = tmCtrls(SDATEINDEX).sShow
                            ElseIf (llTstSDate < llRowSDate) And (llTstEDate <= llRowEDate) Then
                                'Truncate test
                                smSave(3, ilLoop) = Format$(llRowSDate - 1, "m/d/yy")
                                slStr = Trim$(smSave(3, ilLoop))
                                gSetShow pbcComm, slStr, tmCtrls(EDATEINDEX)
                                smShow(EDATEINDEX, ilLoop) = tmCtrls(EDATEINDEX).sShow
                            ElseIf (llRowSDate <= llTstSDate) And (llRowEDate >= llTstEDate) Then
                                'Delete row
                                imRowNo = ilLoop
                                imcTrash_Click
                            ElseIf (llRowSDate = llTstSDate) And (llRowEDate < llTstEDate) Then
                                If ilRMonth = ilTMonth Then
                                    'Delete row
                                    imRowNo = ilLoop
                                    imcTrash_Click
                                Else
                                    'Truncate Row
                                    smSave(2, ilLoop) = Format$(gDateValue(gObtainEndStd(smSave(3, ilRowNo))) + 1, "m/d/yy")
                                    slStr = Trim$(smSave(2, ilLoop))
                                    gSetShow pbcComm, slStr, tmCtrls(SDATEINDEX)
                                    smShow(SDATEINDEX, ilLoop) = tmCtrls(SDATEINDEX).sShow
                                End If
                            Else
                                'Enclosed
                                If ilRMonth = ilTMonth Then
                                    'Truncate test
                                    smSave(3, ilLoop) = Format$(llRowSDate - 1, "m/d/yy")
                                    slStr = Trim$(smSave(3, ilLoop))
                                    gSetShow pbcComm, slStr, tmCtrls(EDATEINDEX)
                                    smShow(EDATEINDEX, ilLoop) = tmCtrls(EDATEINDEX).sShow
                                Else
                                    'Add row
                                    ReDim Preserve smShow(0 To 5, 0 To UBound(smSave, 1) + 1) As String 'Values shown in program area
                                    ReDim Preserve smSave(0 To 5, 0 To UBound(smSave, 1) + 1) As String 'Values saved (program name) in program area
                                    ReDim Preserve tgScfRec(0 To UBound(tgScfRec) + 1) As SCFREC
                                    mInitNew UBound(tgScfRec)
                                    For ilLp1 = LBONE To UBound(smSave, 1) Step 1
                                        smSave(ilLp1, UBound(tgScfRec) - 1) = smSave(ilLp1, ilLoop)
                                        smShow(ilLp1, UBound(tgScfRec) - 1) = smShow(ilLp1, ilLoop)
                                    Next ilLp1
                                    'Truncate Row
                                    smSave(2, UBound(tgScfRec) - 1) = Format$(gDateValue(gObtainEndStd(smSave(3, ilRowNo))) + 1, "m/d/yy")
                                    slStr = Trim$(smSave(2, UBound(tgScfRec) - 1))
                                    gSetShow pbcComm, slStr, tmCtrls(SDATEINDEX)
                                    smShow(SDATEINDEX, UBound(tgScfRec) - 1) = tmCtrls(SDATEINDEX).sShow
                                   'Truncate test
                                    smSave(3, ilLoop) = Format$(llRowSDate - 1, "m/d/yy")
                                    slStr = Trim$(smSave(3, ilLoop))
                                    gSetShow pbcComm, slStr, tmCtrls(EDATEINDEX)
                                    smShow(EDATEINDEX, ilLoop) = tmCtrls(EDATEINDEX).sShow
                                    If UBound(smSave, 2) <= vbcComm.LargeChange Then 'was <=
                                        vbcComm.Max = LBONE 'LBound(smSave, 2) '- 1
                                    Else
                                        vbcComm.Max = UBound(smSave, 2) - vbcComm.LargeChange '- 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next ilLoop
            End If
        Next ilRowNo
    Next ilPass
    For ilRowNo = LBONE To UBound(smSave, 2) - 1 Step 1
        'Set Vehicle
        For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If StrComp(Trim$(tgMVef(ilLoop).sName), Trim$(smSave(1, ilRowNo)), 1) = 0 Then
                tgScfRec(ilRowNo).tScf.iVefCode = tgMVef(ilLoop).iCode
                Exit For
            End If
        Next ilLoop
        'Start Date
        gPackDate Trim$(smSave(2, ilRowNo)), tgScfRec(ilRowNo).tScf.iStartDate(0), tgScfRec(ilRowNo).tScf.iStartDate(1)
        'End Date
        gPackDate Trim$(smSave(3, ilRowNo)), tgScfRec(ilRowNo).tScf.iEndDate(0), tgScfRec(ilRowNo).tScf.iEndDate(1)
        'Under %
        tgScfRec(ilRowNo).tScf.iUnderComm = gStrDecToInt(smSave(4, ilRowNo), 2)
        'Remnant Under %
        tgScfRec(ilRowNo).tScf.iRemUnderComm = gStrDecToInt(smSave(5, ilRowNo), 2)
    Next ilRowNo
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
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
    Dim ilRowNo As Integer
    Dim ilUpper As Integer
    ilUpper = UBound(tgScfRec)
    ReDim smShow(0 To 5, 0 To ilUpper) As String 'Values shown in program area
    ReDim smSave(0 To 5, 0 To ilUpper) As String 'Values saved (program name) in program area
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, ilUpper) = ""
    Next ilLoop
    For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
        smSave(ilLoop, ilUpper) = ""
    Next ilLoop
    'Init value in the case that no records are associated with the salesperson
    If ilUpper = LBONE Then
        ilRowNo = imRowNo
        imRowNo = 1
        mInitNew imRowNo
        imRowNo = ilRowNo
    End If
    For ilRowNo = LBONE To UBound(tgScfRec) - 1 Step 1
        'Get Vehicle Name
        'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1 'Traffic!lbcAdvertiser.ListCount - 1 Step 1
        '    If tgMVef(ilLoop).iCode = tgScfRec(ilRowNo).tScf.iVefCode Then
            ilLoop = gBinarySearchVef(tgScfRec(ilRowNo).tScf.iVefCode)
            If ilLoop <> -1 Then
                smSave(1, ilRowNo) = Trim$(tgMVef(ilLoop).sName)
        '        Exit For
            End If
        'Next ilLoop
        'Get Start Date
        gUnpackDate tgScfRec(ilRowNo).tScf.iStartDate(0), tgScfRec(ilRowNo).tScf.iStartDate(1), smSave(2, ilRowNo)
        'Get End Date
        gUnpackDate tgScfRec(ilRowNo).tScf.iEndDate(0), tgScfRec(ilRowNo).tScf.iEndDate(1), smSave(3, ilRowNo)
        'Get Under %
        smSave(4, ilRowNo) = gIntToStrDec(tgScfRec(ilRowNo).tScf.iUnderComm, 2)
        'Get Remnant Under %
        smSave(5, ilRowNo) = gIntToStrDec(tgScfRec(ilRowNo).tScf.iRemUnderComm, 2)
    Next ilRowNo
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Salesperson list box  *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilType As Integer
    ilIndex = cbcSelect.ListIndex
    If ilIndex > 1 Then
        slName = cbcSelect.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopCntrProjComboBox(CntrProj, cbcSelect, Traffic!lbcSaleCntrProj, Traffic!cbcSelectCombo, igSlfFirstNameFirst)
    ilType = 0  '0=All;1=Salesperson and Negotiator;4=Salespersons, Negotiator and Planner
    'ilRet = gPopSalespersonBox(CntrProj, ilType, True, True, cbcSelect, Traffic!lbcSalesperson, igSlfFirstNameFirst)
    ilRet = gPopSalespersonBox(SlspComm, ilType, True, False, cbcSelect, tmSlspCommSalesperson(), smSlspCommSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopluateErr
        gCPErrorMsg ilRet, "mPopluate (gIMoveListBox)", SlspComm
        On Error GoTo 0
        imChgMode = True
        If ilIndex >= 0 Then
            gFindMatch slName, 0, cbcSelect
            If gLastFound(cbcSelect) >= 0 Then
                cbcSelect.ListIndex = gLastFound(cbcSelect)
            Else
                cbcSelect.ListIndex = -1
            End If
        Else
            cbcSelect.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mPopluateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec() As Integer
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilSlfCode As Integer
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim llLoop As Long
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim slStr As String
    Dim slDate As String
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    ReDim tgScfRec(0 To 1) As SCFREC
    tgScfRec(0).iStatus = -1
    tgScfRec(0).lRecPos = 0
    tgScfRec(0).iDateChg = False
    tgScfRec(1).iStatus = -1
    tgScfRec(1).lRecPos = 0
    tgScfRec(1).iDateChg = False
    ReDim tgScfDel(0 To 0) As SCFREC
    tgScfDel(0).iStatus = -1
    tgScfDel(0).lRecPos = 0
    ilUpper = 1
    If imSlspSelectedIndex < 0 Then
        mReadRec = False
        Exit Function
    End If
    btrExtClear hmScf   'Clear any previous extend operation
    ilExtLen = Len(tgScfRec(1).tScf)  'Extract operation record size
    slNameCode = tmSlspCommSalesperson(imSlspSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilSlfCode = Val(slCode)
    tmScfSrchKey1.iCode = ilSlfCode
    ilRet = btrGetEqual(hmScf, tmScf, imScfRecLen, tmScfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
        ilRet = BTRV_ERR_END_OF_FILE
    Else
        If ilRet <> BTRV_ERR_NONE Then
            mReadRec = False
            Exit Function
        End If
    End If
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmScf, llNoRec, -1, "UC", "SCF", "") '"EG") 'Set extract limits (all records)
        ilOffSet = gFieldOffset("SCF", "SCFSLFCODE") 'GetOffSetForInt(tmScf, tmScf.iSlfCode)
        tlIntTypeBuff.iType = ilSlfCode
        ilRet = btrExtAddLogicConst(hmScf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        On Error GoTo mReadRecErr
        gBtrvErrorMsg ilRet, "mReadRec (btrExtAddLogicConst):" & "Scf.Btr", SlspComm
        On Error GoTo 0
        ilRet = btrExtAddField(hmScf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mReadRecErr
        gBtrvErrorMsg ilRet, "mReadRec (btrExtAddField):" & "Scf.Btr", SlspComm
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hmRpf)    'Extract record
        ilRet = btrExtGetNext(hmScf, tgScfRec(ilUpper).tScf, ilExtLen, tgScfRec(ilUpper).lRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mReadRecErr
            gBtrvErrorMsg ilRet, "mReadRec (btrExtGetNextExt):" & "Scf.Btr", SlspComm
            On Error GoTo 0
            'ilRet = btrExtGetFirst(hmRpf, tlRpfExt, ilExtLen, llRecPos)
            ilExtLen = Len(tgScfRec(1).tScf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmScf, tgScfRec(ilUpper).tScf, ilExtLen, tgScfRec(ilUpper).lRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                slStr = ""
                'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                '    If tgMVef(ilLoop).iCode = tgScfRec(ilUpper).tScf.iVefCode Then
                    ilLoop = gBinarySearchVef(tgScfRec(ilUpper).tScf.iVefCode)
                    If ilLoop <> -1 Then
                        slStr = tgMVef(ilLoop).sName
                '        Exit For
                    End If
                'Next ilLoop

                gUnpackDateForSort tgScfRec(ilUpper).tScf.iStartDate(0), tgScfRec(ilUpper).tScf.iStartDate(1), slDate
                tgScfRec(ilUpper).sKey = slStr & slDate
                tgScfRec(ilUpper).iStatus = 1
                tgScfRec(ilUpper).iDateChg = False
                ilUpper = ilUpper + 1
                ReDim Preserve tgScfRec(0 To ilUpper) As SCFREC
                tgScfRec(ilUpper).iStatus = -1
                tgScfRec(ilUpper).lRecPos = 0
                ilRet = btrExtGetNext(hmScf, tgScfRec(ilUpper).tScf, ilExtLen, tgScfRec(ilUpper).lRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmScf, tgScfRec(ilUpper).tScf, ilExtLen, tgScfRec(ilUpper).lRecPos)
                Loop
            Loop
        End If
    End If
    If ilUpper > 1 Then
        'ArraySortTyp fnAV(tgScfRec(), 1), UBound(tgScfRec) - 1, 0, LenB(tgScfRec(1)), 0, LenB(tgScfRec(1).sKey), 0
        For llLoop = LBound(tgScfRec) To UBound(tgScfRec) - 1 Step 1
            tgScfRec(llLoop) = tgScfRec(llLoop + 1)
        Next llLoop
        ReDim Preserve tgScfRec(0 To UBound(tgScfRec) - 1) As SCFREC
        ArraySortTyp fnAV(tgScfRec(), 0), UBound(tgScfRec), 0, LenB(tgScfRec(0)), 0, LenB(tgScfRec(0).sKey), 0
        ReDim Preserve tgScfRec(0 To UBound(tgScfRec) + 1) As SCFREC
        For llLoop = UBound(tgScfRec) - 1 To LBound(tgScfRec) Step -1
            tgScfRec(llLoop + 1) = tgScfRec(llLoop)
        Next llLoop
    End If
    mReadRec = True
    Exit Function
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:6/29/93       By:D. LeVine      *
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
    Dim ilCRet As Integer
    Dim slNameCode As String  'Name and Code number
    Dim slCode As String    'Code number
    Dim ilRowNo As Integer
    Dim slMsg As String
    Dim ilSlfCode As Integer
    Dim ilScf As Integer
    Dim tlScf As SCF
    Dim tlScf1 As MOVEREC
    Dim tlScf2 As MOVEREC
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    For ilRowNo = 1 To UBound(smSave, 2) - 1 Step 1
        If mTestSaveFields(ilRowNo) = NO Then
            mSaveRec = False
            imRowNo = ilRowNo
            Exit Function
        End If
    Next ilRowNo
    mMoveCtrlToRec
    Screen.MousePointer = vbHourglass  'Wait
    If imTerminate Then
        Screen.MousePointer = vbDefault    'Default
        mSaveRec = False
        Exit Function
    End If
    slNameCode = tmSlspCommSalesperson(imSlspSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilSlfCode = Val(slCode)
    ilRet = btrBeginTrans(hmScf, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 1", vbOKOnly + vbExclamation, "Commission")
        Exit Function
    End If
    ilLoop = 0
    For ilScf = LBONE To UBound(tgScfRec) - 1 Step 1
        Do  'Loop until record updated or added
            If (tgScfRec(ilScf).iStatus = 0) Then  'New selected
                'User
                gPackDate smNowDate, tgScfRec(ilScf).tScf.iDateEntrd(0), tgScfRec(ilScf).tScf.iDateEntrd(1)
                tgScfRec(ilScf).tScf.lCode = 0
                tgScfRec(ilScf).tScf.iSlfCode = ilSlfCode
                tgScfRec(ilScf).tScf.iUrfCode = tgUrf(0).iCode
                ilRet = btrInsert(hmScf, tgScfRec(ilScf).tScf, imScfRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmScf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 2", vbOKOnly + vbExclamation, "Commission")
                    Exit Function
                End If
                slMsg = "mSaveRec (btrInsert: Sales Commission)"
                ilRet = btrGetPosition(hmScf, tgScfRec(ilScf).lRecPos)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmScf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 3", vbOKOnly + vbExclamation, "Commission")
                    Exit Function
                End If
                tgScfRec(ilScf).iStatus = 1
            ElseIf (tgScfRec(ilScf).iStatus = 1) Then  'Old record-Update
                slMsg = "mSaveRec (btrGetEqual: Sales Commission)"
                tmScfSrchKey.lCode = tgScfRec(ilScf).tScf.lCode
                ilRet = btrGetEqual(hmScf, tlScf, imScfRecLen, tmScfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmScf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 4", vbOKOnly + vbExclamation, "Commission")
                    Exit Function
                End If
                LSet tlScf1 = tlScf
                LSet tlScf2 = tgScfRec(ilScf).tScf
                If StrComp(tlScf1.sChar, tlScf2.sChar, 0) <> 0 Then
                    tgScfRec(ilScf).tScf.iUrfCode = tgUrf(0).iCode
                    ilRet = btrUpdate(hmScf, tgScfRec(ilScf).tScf, imScfRecLen)
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                slMsg = "mSaveRec (btrUpdate: Sales Commission)"
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            ilCRet = btrAbortTrans(hmScf)
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 5", vbOKOnly + vbExclamation, "Commission")
            Exit Function
        End If
    Next ilScf
    For ilScf = LBound(tgScfDel) To UBound(tgScfDel) - 1 Step 1
        If tgScfDel(ilScf).iStatus = 1 Then
            Do
                slMsg = "mSaveRec (btrGetEqual: Sales Commission)"
                tmScfSrchKey.lCode = tgScfDel(ilScf).tScf.lCode
                ilRet = btrGetEqual(hmScf, tmScf, imScfRecLen, tmScfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmScf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 6", vbOKOnly + vbExclamation, "Commission")
                    Exit Function
                End If
                ilRet = btrDelete(hmScf)
                slMsg = "mSaveRec (btrDelete: Sales Commission)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                ilCRet = btrAbortTrans(hmScf)
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 7", vbOKOnly + vbExclamation, "Commission")
                Exit Function
            End If
        End If
    Next ilScf
    ilRet = btrEndTrans(hmScf)
    mSaveRec = True
    Screen.MousePointer = vbDefault    'Default
    Exit Function

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
'*             Created:9/24/93       By:D. LeVine      *
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
    Dim ilLoop As Integer
    Dim ilNew As Integer
    If imScfChg And (UBound(tgScfRec) > LBONE) Or (UBound(tgScfDel) > LBound(tgScfDel)) Then
        If ilAsk Then
            ilNew = True
            For ilLoop = LBONE To UBound(tgScfRec) - 1 Step 1
                If tgScfRec(ilLoop).iStatus <> 0 Then
                    ilNew = False
                    Exit For
                End If
            Next ilLoop
            For ilLoop = LBound(tgScfDel) To UBound(tgScfDel) - 1 Step 1
                If tgScfDel(ilLoop).iStatus <> 0 Then
                    ilNew = False
                    Exit For
                End If
            Next ilLoop
            If Not ilNew Then
                slMess = "Save Changes"
            Else
                slMess = "Add Changes"
            End If
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
    If (imBypassSetting) Or (Not imUpdateAllowed) Then
        Exit Sub
    End If
    ilAltered = imScfChg
    If (Not ilAltered) And (UBound(tgScfDel) > LBound(tgScfDel)) Then
        ilAltered = True
    End If
    If ilAltered Then
        pbcComm.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        cbcSelect.Enabled = False
    Else
        If imSlspSelectedIndex < 0 Then
            pbcComm.Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            cbcSelect.Enabled = True
        Else
            pbcComm.Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            cbcSelect.Enabled = True
        End If
    End If
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields() = YES) And (ilAltered) And (UBound(tgScfRec) > 1) And (imUpdateAllowed) Then
        cmcSave.Enabled = True
    Else
        cmcSave.Enabled = False
    End If
    If imSlspSelectedIndex < 0 Then
        cmcModel.Enabled = False
        cmcCreate.Enabled = False
    Else
        cmcModel.Enabled = True
        cmcCreate.Enabled = True
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
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
        Case VEHINDEX
            edcDropDown.SetFocus
        Case SDATEINDEX
            edcDropDown.SetFocus
        Case EDATEINDEX
            edcDropDown.SetFocus
        Case GOALPVEHINDEX
            edcDropDown.SetFocus
        Case REMNANTPVEHINDEX
            edcDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetMinMax                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set scroll bar min/max         *
'*                                                     *
'*******************************************************
Private Sub mSetMinMax()
    imSettingValue = True
    vbcComm.Min = LBONE 'LBound(smShow, 2)
    imSettingValue = True
    If UBound(smShow, 2) - 1 <= vbcComm.LargeChange + 1 Then ' + 1 Then
        vbcComm.Max = LBONE 'LBound(smShow, 2)
    Else
        vbcComm.Max = UBound(smShow, 2) - vbcComm.LargeChange
    End If
    imSettingValue = True
    If vbcComm.Value = vbcComm.Min Then
        vbcComm_Change
    Else
        vbcComm.Value = vbcComm.Min
    End If
    imSettingValue = False
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
    pbcArrow.Visible = False
    lacFrame.Visible = False
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case VEHINDEX
            lbcVehicle.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcComm, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(1, imRowNo)) <> slStr Then
                imScfChg = True
                smSave(1, imRowNo) = slStr
            End If
        Case SDATEINDEX
            plcCalendar.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            If gValidDate(slStr) Then
                gSetShow pbcComm, slStr, tmCtrls(ilBoxNo)
                smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
                If Trim$(smSave(2, imRowNo)) <> slStr Then
                    imScfChg = True
                    smSave(2, imRowNo) = slStr
                    tgScfRec(imRowNo).iDateChg = True
                End If
            End If
        Case EDATEINDEX
            plcCalendar.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            If StrComp(slStr, "TFN", 1) <> 0 Then
                If Trim$(smSave(3, imRowNo)) <> slStr Then
                    imScfChg = True
                    smSave(3, imRowNo) = slStr
                    tgScfRec(imRowNo).iDateChg = True
                End If
                If gValidDate(slStr) Then
                    smSave(3, imRowNo) = slStr
                End If
            Else
                If Trim$(smSave(3, imRowNo)) <> "" Then
                    imScfChg = True
                    tgScfRec(imRowNo).iDateChg = True
                End If
                smSave(3, imRowNo) = ""
                slStr = "TFN"
            End If
            gSetShow pbcComm, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
        Case GOALPVEHINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcComm, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(4, imRowNo)) <> slStr Then
                imScfChg = True
                smSave(4, imRowNo) = slStr
            End If
        Case REMNANTPVEHINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcComm, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            If Trim$(smSave(5, imRowNo)) <> slStr Then
                imScfChg = True
                smSave(5, imRowNo) = slStr
            End If
    End Select
    mSetCommands
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
    'Unload Traffic
    Unload SlspComm
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields() As Integer
'
'   iRet = mTestFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRowNo As Integer
    For ilRowNo = LBONE To UBound(smSave, 2) - 1 Step 1
        If Trim$(smSave(1, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smSave(2, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        'If Trim$(smSave(3, ilRowNo)) = "" Then
        '    mTestFields = NO
        '    Exit Function
        'End If
        If Trim$(smSave(4, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smSave(5, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
    Next ilRowNo
    mTestFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestSaveFields                 *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestSaveFields(ilRowNo As Integer) As Integer
'
'   iRet = mTestSaveFields(ilRowNo)
'   Where:
'       ilRowNo (I)- row number to be checked
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    If Trim$(smSave(1, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = VEHINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smSave(2, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("Start Date must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = SDATEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If gDateValue(smSave(2, ilRowNo)) <> gDateValue(gObtainStartStd(smSave(2, ilRowNo))) Then
        Beep
        ilRes = MsgBox("Start Date must be first day of the Month", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = SDATEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smSave(4, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("Under Goal Percent must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = GOALPVEHINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smSave(5, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("Under Remnant Percent must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = REMNANTPVEHINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    mTestSaveFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
    Dim ilRet As Integer
    ilRet = gPopUserVehicleBox(SlspComm, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHNTR + ACTIVEVEH + DORMANTVEH, lbcVehicle, tmUserVehicle(), smUserVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", SlspComm
        On Error GoTo 0
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0

    llDate = lmCalStartDate

    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                edcDropDown.Text = Format$(llDate, "m/d/yy")
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                imBypassFocus = True
                edcDropDown.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate

    edcDropDown.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    plcCalendar.Visible = False
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcClickFocus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcComm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    imDragShift = Shift
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub pbcComm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    If Button = 2 Then
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    Screen.MousePointer = vbDefault
    ilCompRow = vbcComm.LargeChange + 1
    If UBound(tgScfRec) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(tgScfRec) + 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcComm.Value - 1
                    If ilRowNo > UBound(smSave, 2) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    If (ilBox > VEHINDEX) And (Trim$(smSave(1, ilRowNo)) = "") Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    mSetShow imBoxNo
                    imRowNo = ilRow + vbcComm.Value - 1
                    If (imRowNo = UBound(smSave, 2)) And (Trim$(smSave(1, imRowNo)) = "") Then
                        mInitNew imRowNo
                    End If
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mSetFocus imBoxNo
End Sub
Private Sub pbcComm_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    mPaintCommTitle
    ilStartRow = vbcComm.Value '+ 1  'Top location
    ilEndRow = vbcComm.Value + vbcComm.LargeChange ' + 1
    If ilEndRow > UBound(smSave, 2) Then
        If Trim$(smShow(1, UBound(smShow, 2))) <> "" Then
            ilEndRow = UBound(smSave, 2) 'include blank row as it might have data
        Else
            ilEndRow = UBound(smSave, 2) - 1
        End If
    End If
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            'If ilBox <> TOTALINDEX Then
            '    gPaintArea pbcProj, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
            'Else
            '    gPaintArea pbcProj, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
            'End If
            pbcComm.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcComm.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            slStr = Trim$(smShow(ilBox, ilRow))
            pbcComm.Print slStr
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcNum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    imcNumInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 4 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            For ilColNo = 1 To 3 Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcNumInv.Move flX, flY
                    imcNumInv.Visible = True
                    imcNumOutline.Move flX - 15, flY - 15
                    imcNumOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcNum_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    imcNumInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 4 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            For ilColNo = 1 To 3 Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcNumInv.Move flX, flY
                    imcNumOutline.Move flX - 15, flY - 15
                    imcNumOutline.Visible = True
                    Select Case ilRowNo
                        Case 1
                            Select Case ilColNo
                                Case 1
                                    slKey = "7"
                                Case 2
                                    slKey = "8"
                                Case 3
                                    slKey = "9"
                            End Select
                        Case 2
                            Select Case ilColNo
                                Case 1
                                    slKey = "4"
                                Case 2
                                    slKey = "5"
                                Case 3
                                    slKey = "6"
                            End Select
                        Case 3
                            Select Case ilColNo
                                Case 1
                                    slKey = "1"
                                Case 2
                                    slKey = "2"
                                Case 3
                                    slKey = "3"
                            End Select
                        Case 4
                            Select Case ilColNo
                                Case 1
                                    slKey = "0"
                                Case 2
                                    slKey = "00"
                                Case 3
                                    slKey = "."
                            End Select
                    End Select
                    imBypassFocus = True    'Don't change select text
                    edcDropDown.SetFocus
                    'SendKeys slKey
                    gSendKeys edcDropDown, slKey
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to left
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            If (UBound(smSave, 2) = 1) Then
                imTabDirection = 0  'Set-Left to right
                imRowNo = 1
                mInitNew imRowNo
            Else
                If UBound(smSave, 2) <= vbcComm.LargeChange Then 'was <=
                    vbcComm.Max = LBONE 'LBound(smSave, 2)
                Else
                    vbcComm.Max = UBound(smSave, 2) - vbcComm.LargeChange '- 1
                End If
                imRowNo = 1
                If imRowNo >= UBound(smSave, 2) Then
                    mInitNew imRowNo
                End If
                imSettingValue = True
                vbcComm.Value = vbcComm.Min
                imSettingValue = False
            End If
            ilBox = VEHINDEX
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case VEHINDEX, 0
            mSetShow imBoxNo
            If (imBoxNo < 1) And (imRowNo < 1) Then 'Modelled from Proposal
                Exit Sub
            End If
            ilBox = REMNANTPVEHINDEX
            If imRowNo <= 1 Then
                imBoxNo = -1
                imRowNo = -1
                cmcDone.SetFocus
                Exit Sub
            End If
            imRowNo = imRowNo - 1
            If imRowNo < vbcComm.Value Then
                imSettingValue = True
                vbcComm.Value = vbcComm.Value - 1
                imSettingValue = False
            End If
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case SDATEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo - 1
        Case EDATEINDEX
            slStr = edcDropDown.Text
            If (slStr <> "") And (StrComp(slStr, "TFN", 1) <> 0) Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            End If
            ilBox = imBoxNo - 1
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcSTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilLoop As Integer
    Dim slStr As String

    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    imTabDirection = 0 'Set- Left to right
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            imRowNo = UBound(smSave, 2) - 1
            imSettingValue = True
            If imRowNo <= vbcComm.LargeChange + 1 Then
                vbcComm.Value = 1
            Else
                vbcComm.Value = imRowNo - vbcComm.LargeChange - 1
            End If
            imSettingValue = False
            ilBox = REMNANTPVEHINDEX
        Case REMNANTPVEHINDEX
            mSetShow imBoxNo
            If mTestSaveFields(imRowNo) = NO Then
                mEnableBox imBoxNo
                Exit Sub
            End If
            If imRowNo >= UBound(smSave, 2) Then
                imScfChg = True
                ReDim Preserve smShow(0 To 5, 0 To imRowNo + 1) As String 'Values shown in program area
                ReDim Preserve smSave(0 To 5, 0 To imRowNo + 1) As String 'Values saved (program name) in program area
                For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
                    smShow(ilLoop, imRowNo + 1) = ""
                Next ilLoop
                For ilLoop = LBound(smSave, 1) To UBound(smSave, 1) Step 1
                    smSave(ilLoop, imRowNo + 1) = ""
                Next ilLoop
                ReDim Preserve tgScfRec(0 To UBound(tgScfRec) + 1) As SCFREC
                tgScfRec(UBound(tgScfRec)).iStatus = 0
                tgScfRec(UBound(tgScfRec)).lRecPos = 0
            End If
            If imRowNo >= UBound(smSave, 2) - 1 Then
                imRowNo = imRowNo + 1
                mInitNew imRowNo
                If UBound(smSave, 2) <= vbcComm.LargeChange Then 'was <=
                    vbcComm.Max = LBONE 'LBound(smSave, 2) '- 1
                Else
                    vbcComm.Max = UBound(smSave, 2) - vbcComm.LargeChange '- 1
                End If
            Else
                imRowNo = imRowNo + 1
            End If
            If imRowNo > vbcComm.Value + vbcComm.LargeChange Then
                imSettingValue = True
                vbcComm.Value = vbcComm.Value + 1
                imSettingValue = False
            End If
            If imRowNo >= UBound(smSave, 2) Then
                imBoxNo = 0
                mSetCommands
                'lacFrame.Move 0, tmCtrls(PROPNOINDEX).fBoxY + (imRowNo - vbcProj.Value) * (fgBoxGridH + 15) - 30
                'lacFrame.Visible = True
                pbcArrow.Move pbcArrow.Left, plcComm.Top + tmCtrls(VEHINDEX).fBoxY + (imRowNo - vbcComm.Value) * (fgBoxGridH + 15) + 45
                pbcArrow.Visible = True
                pbcArrow.SetFocus
                Exit Sub
            Else
                ilBox = VEHINDEX
            End If
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case 0
            ilBox = VEHINDEX
        Case SDATEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo + 1
        Case EDATEINDEX
            slStr = edcDropDown.Text
            If (slStr <> "") And (StrComp(slStr, "TFN", 1) <> 0) Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            End If
            ilBox = imBoxNo + 1
        Case Else
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub tmcDrag_Timer()
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            ilCompRow = vbcComm.LargeChange + 1
            If UBound(smSave, 2) > ilCompRow Then
                ilMaxRow = ilCompRow
            Else
                ilMaxRow = UBound(smSave, 2)
            End If
            For ilRow = 1 To ilMaxRow Step 1
                If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(VEHINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(VEHINDEX).fBoxY + tmCtrls(VEHINDEX).fBoxH)) Then
                    mSetShow imBoxNo
                    imBoxNo = -1
                    imRowNo = -1
                    imRowNo = ilRow + vbcComm.Value - 1
                    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                    lacFrame.Move 0, tmCtrls(VEHINDEX).fBoxY + (imRowNo - vbcComm.Value) * (fgBoxGridH + 15) - 30
                    'If gInvertArea call then remove visible setting
                    lacFrame.Visible = True
                    pbcArrow.Move pbcArrow.Left, plcComm.Top + tmCtrls(VEHINDEX).fBoxY + (imRowNo - vbcComm.Value) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    imcTrash.Enabled = True
                    lacFrame.Drag vbBeginDrag
                    lacFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                    Exit Sub
                End If
            Next ilRow
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub
Private Sub vbcComm_Change()
    If imSettingValue Then
        pbcComm.Cls
        pbcComm_Paint
        imSettingValue = False
    Else
        mSetShow imBoxNo
        imBoxNo = -1
        imRowNo = -1
        pbcComm.Cls
        pbcComm_Paint
    End If
End Sub
Private Sub vbcComm_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    gCtrlGotFocus ActiveControl
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Commission"
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPaintLnTitle                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint Header Titles            *
'*                                                     *
'*******************************************************
Private Sub mPaintCommTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer
    Dim ilHalfY As Integer

    llColor = pbcComm.ForeColor
    slFontName = pbcComm.FontName
    flFontSize = pbcComm.FontSize
    ilFillStyle = pbcComm.FillStyle
    llFillColor = pbcComm.FillColor
    pbcComm.ForeColor = BLUE
    pbcComm.FontBold = False
    pbcComm.FontSize = 7
    pbcComm.FontName = "Arial"
    pbcComm.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    ilHalfY = tmCtrls(VEHINDEX).fBoxY / 2
    Do While ilHalfY Mod 15 <> 0
        ilHalfY = ilHalfY - 1
    Loop
    pbcComm.Line (tmCtrls(VEHINDEX).fBoxX - 15, 15)-Step(tmCtrls(VEHINDEX).fBoxW + 15, tmCtrls(VEHINDEX).fBoxY - 30), BLUE, B
    pbcComm.CurrentX = tmCtrls(VEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcComm.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcComm.Print "Vehicle"
    pbcComm.Line (tmCtrls(SDATEINDEX).fBoxX - 15, 15)-Step(tmCtrls(SDATEINDEX).fBoxW + 15, tmCtrls(SDATEINDEX).fBoxY - 30), BLUE, B
    pbcComm.CurrentX = tmCtrls(SDATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcComm.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcComm.Print "Start Date"
    pbcComm.Line (tmCtrls(EDATEINDEX).fBoxX - 15, 15)-Step(tmCtrls(EDATEINDEX).fBoxW + 15, tmCtrls(EDATEINDEX).fBoxY - 30), BLUE, B
    pbcComm.CurrentX = tmCtrls(EDATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcComm.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcComm.Print "End Date"
    pbcComm.Line (tmCtrls(GOALPVEHINDEX).fBoxX - 15, 15)-Step(tmCtrls(GOALPVEHINDEX).fBoxW + 15, tmCtrls(GOALPVEHINDEX).fBoxY - 30), BLUE, B
    pbcComm.CurrentX = tmCtrls(GOALPVEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcComm.CurrentY = 15
    pbcComm.Print "Under"
    pbcComm.CurrentX = tmCtrls(GOALPVEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcComm.CurrentY = ilHalfY + 15
    pbcComm.Print "Goal %"
    pbcComm.Line (tmCtrls(REMNANTPVEHINDEX).fBoxX - 15, 15)-Step(tmCtrls(REMNANTPVEHINDEX).fBoxW + 15, tmCtrls(REMNANTPVEHINDEX).fBoxY - 30), BLUE, B
    pbcComm.CurrentX = tmCtrls(REMNANTPVEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcComm.CurrentY = 15
    pbcComm.Print "Under"
    pbcComm.CurrentX = tmCtrls(REMNANTPVEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcComm.CurrentY = ilHalfY + 15
    pbcComm.Print "Remnant %"

    ilLineCount = 0
    llTop = tmCtrls(1).fBoxY
    Do
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            pbcComm.Line (tmCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmCtrls(1).fBoxH + 15
    Loop While llTop + tmCtrls(1).fBoxH < pbcComm.Height
    vbcComm.LargeChange = ilLineCount - 1
    pbcComm.FontSize = flFontSize
    pbcComm.FontName = slFontName
    pbcComm.FontSize = flFontSize
    pbcComm.ForeColor = llColor
    pbcComm.FontBold = True
End Sub

