VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form PostRepTimes 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4470
   ClientLeft      =   885
   ClientTop       =   2415
   ClientWidth     =   11205
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
   ScaleHeight     =   4470
   ScaleWidth      =   11205
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "PostRepTimes.frx":0000
      Left            =   1095
      List            =   "PostRepTimes.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3540
      Visible         =   0   'False
      Width           =   4425
   End
   Begin VB.PictureBox pbcPaid 
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
      Left            =   1575
      ScaleHeight     =   210
      ScaleWidth      =   840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.ListBox lbcCopyNm 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "PostRepTimes.frx":0004
      Left            =   1080
      List            =   "PostRepTimes.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3255
      Visible         =   0   'False
      Width           =   4425
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
      Left            =   2160
      Picture         =   "PostRepTimes.frx":0008
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   870
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   5685
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1275
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "PostRepTimes.frx":0102
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcTmeOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   330
            Top             =   270
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "PostRepTimes.frx":0DC0
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
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
      Left            =   3105
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   990
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
         Left            =   1620
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   30
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
         Left            =   30
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   30
         Picture         =   "PostRepTimes.frx":10CA
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
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
            TabIndex        =   18
            Top             =   390
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
         TabIndex        =   21
         Top             =   30
         Width           =   1305
      End
   End
   Begin VB.ComboBox cbcLine 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2760
      TabIndex        =   15
      Top             =   270
      Width           =   8145
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   3630
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   9195
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   30
      Width           =   45
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   30
      Picture         =   "PostRepTimes.frx":3EE4
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   675
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   10485
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   9
      Top             =   30
      Width           =   45
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   30
      TabIndex        =   2
      Top             =   345
      Width           =   30
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1245
      MaxLength       =   10
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   885
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6780
      TabIndex        =   12
      Top             =   4065
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   5055
      TabIndex        =   11
      Top             =   4065
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
      Left            =   11115
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   13
      Top             =   15
      Width           =   75
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   3285
      TabIndex        =   10
      Top             =   4065
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTimes 
      Height          =   2955
      Left            =   210
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   675
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5212
      _Version        =   393216
      Rows            =   10
      Cols            =   9
      FixedCols       =   0
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   10410
      Picture         =   "PostRepTimes.frx":41EE
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label lacSpotStatus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   10800
      TabIndex        =   25
      Top             =   3690
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label lacContract 
      Height          =   210
      Left            =   2880
      TabIndex        =   24
      Top             =   30
      Width           =   6240
   End
   Begin VB.Label lacScreen 
      Caption         =   "Post Rep Times"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   4140
   End
End
Attribute VB_Name = "PostRepTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of PostRepTimes.Frm on Wed 6/17/09 @ 12:
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmSbfSrchKey2                 smMissedCodeTag                                         *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  mMoveCtrlToRec                mInitCif                                                *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PostRepTimes.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim hmSdf As Integer    'Demo Book Name file handle
Dim tmSdf As SDF
Dim tmSdfSrchKey3 As LONGKEY0
Dim tmSdfSrchKey5 As LONGKEY0
Dim imSdfRecLen As Integer        'Sdf record length
Dim tmRepSdf() As SDF

Dim hmSmf As Integer
Dim tmSmf As SMF
Dim imSmfRecLen As Integer
Dim tmSmfSrchKey2 As LONGKEY0   'SdfCode

Dim hmCif As Integer
Dim tmCif As CIF        'Rvf record image
Dim tmCifSrchKey As LONGKEY0
Dim imCifRecLen As Integer        'RvF record length

Dim hmCpf As Integer        'Copy Product/Agency file handle
Dim tmCpf As CPF            'CPF record image
Dim tmCpfSrchKey As LONGKEY0 'CPF key record image
Dim imCpfRecLen As Integer     'CPF record length

Dim hmSbf As Integer        'Special billing
Dim tmSbf() As SBF    'SBF record image of billing items
Dim tmExistSbf() As SBF
Dim tmSbfSrchKey0 As SBFKEY0
Dim tmSbfSrchKey1 As LONGKEY0            'SBF record image
Dim imSbfRecLen As Integer        'SBF record length

Dim hmRdf As Integer            'Contract line flight file handle
Dim imRdfRecLen As Integer        'CFF record length
Dim tmRdf As RDF
Dim tmRdfSrchKey As INTKEY0            'RWF record image

Dim hmCff As Integer

Dim tmVehicle() As SORTCODE
Dim smVehicleTag As String

Dim smLnVehicle As String 'Line vehicle
Dim smBonusOnInv As String
Dim smAllowRepMG As String

Dim lmDelSdfCode() As Long

Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imLastSelectRow As Integer
Dim imCtrlKey As Integer
Dim imLastColSorted As Integer
Dim imLastSort As Integer
Dim imChg As Integer
Dim imIgnoreScroll As Integer
Dim imFromArrow As Integer
Dim imUpdateAllowed As Integer
Dim imChgMode As Integer
Dim imLbcMouseDown As Integer
Dim imLbcArrowSetting As Integer
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imBypassFocus As Integer
Dim smMissedTime As String
Dim imStatus As Integer     '0=Paid; 1=Bonus

Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imCtrlVisible As Integer
Dim lmTopRow As Long
Dim imInitNoRows As Integer

Dim tmCopyNmCode() As SORTCODE
Dim smCopyNmCodeTag As String
Dim tmMissedCode() As SORTCODE

Dim imOrderedSpots() As Integer

'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer

Dim imCalYear As Integer    'Month of displayed calendar
Dim imLBCalYear As Integer
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer

Const STATUSINDEX = 0
Const DATEINDEX = 1
Const TIMEINDEX = 2
Const ISCIINDEX = 3
Const MGAIRVEHICLEINDEX = 4
Const MGMISSEDDATEINDEX = 5
Const SDFCODEINDEX = 6
Const SORTINDEX = 7
Const CHGINDEX = 8


Private Sub cbcLine_Change()
    Dim ilRet As Integer
    Dim ilLineNo As Integer
    Dim ilClf As Integer
    Dim ilVeh As Integer

    mClearGrid
    ilClf = cbcLine.ItemData(cbcLine.ListIndex)
    ilLineNo = tgClfRep(ilClf).ClfRec.iLine
    If ((tgClfRep(ilClf).ClfRec.iStartTime(0) <> 1) Or (tgClfRep(ilClf).ClfRec.iStartTime(1) <> 0)) Then
        gUnpackTime tgClfRep(ilClf).ClfRec.iStartTime(0), tgClfRep(ilClf).ClfRec.iStartTime(1), "A", "1", smMissedTime
    Else
        tmRdfSrchKey.iCode = tgClfRep(ilClf).ClfRec.iRdfCode  ' Daypart File Code
        ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            'gUnpackTime tmRdf.iStartTime(0, 7), tmRdf.iStartTime(1, 7), "A", "1", smMissedTime
            gUnpackTime tmRdf.iStartTime(0, 6), tmRdf.iStartTime(1, 6), "A", "1", smMissedTime
        End If
    End If
    smLnVehicle = ""
    For ilVeh = 0 To lbcVehicle.ListCount - 1 Step 1
        If Val(lbcVehicle.ItemData(ilVeh)) = tgClfRep(ilClf).ClfRec.iVefCode Then
            smLnVehicle = Trim$(lbcVehicle.List(ilVeh))
            Exit For
        End If
    Next ilVeh

    ilRet = mReadRec(lgRepSpotChfCode, ilLineNo)
    If ilRet Then
        mMoveRecToCtrl
    End If
End Sub

Private Sub cbcLine_Click()
    cbcLine_Change
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
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer

    If imChg Then
        If MsgBox("Save all changes?", vbYesNo) = vbYes Then
            ilRet = mSaveRec()
            If Not ilRet Then
                Exit Sub
            End If
        End If
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcDropDown_Click()
    Select Case lmEnableCol
        Case STATUSINDEX
        Case DATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case TIMEINDEX
            plcTme.Visible = Not plcTme.Visible
        Case ISCIINDEX
            lbcCopyNm.Visible = Not lbcCopyNm.Visible
        Case MGAIRVEHICLEINDEX
            lbcVehicle.Visible = Not lbcVehicle.Visible
        Case MGMISSEDDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
    End Select
End Sub

Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcUpdate_Click()
    Dim ilRet As Integer
    Dim ilLineNo As Integer

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    ilRet = mSaveRec()
    If ilRet Then
        mClearGrid
        ilLineNo = tgClfRep(cbcLine.ItemData(cbcLine.ListIndex)).ClfRec.iLine
        ilRet = mReadRec(lgRepSpotChfCode, ilLineNo)
        If ilRet Then
            mMoveRecToCtrl
        End If
        mSetCommands
    End If
End Sub

Private Sub cmcUpdate_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer

    Select Case lmEnableCol
        Case STATUSINDEX
        Case DATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case TIMEINDEX
        Case ISCIINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcCopyNm, imBSMode, slStr)
        Case MGAIRVEHICLEINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
        Case MGMISSEDDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
    End Select
    grdTimes.CellForeColor = vbBlack

End Sub

Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub

Private Sub edcDropDown_GotFocus()
    Select Case lmEnableCol
        Case STATUSINDEX
        Case DATEINDEX
        Case TIMEINDEX
        Case ISCIINDEX
        Case MGAIRVEHICLEINDEX
        Case MGMISSEDDATEINDEX
    End Select
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPos                         slStr                                                   *
'******************************************************************************************

    Dim ilKey As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer

    ilKey = KeyAscii

    Select Case lmEnableCol
        Case STATUSINDEX
        Case DATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case TIMEINDEX
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                ilFound = False
                For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
                    If KeyAscii = igLegalTime(ilLoop) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        Case ISCIINDEX
        Case MGAIRVEHICLEINDEX
        Case MGMISSEDDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select

End Sub

Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String

    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        Select Case lmEnableCol
            Case STATUSINDEX
            Case DATEINDEX
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
            Case TIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case ISCIINDEX
                gProcessArrowKey Shift, KeyCode, lbcCopyNm, imLbcArrowSetting
            Case MGAIRVEHICLEINDEX
                gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
            Case MGMISSEDDATEINDEX
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
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case lmEnableCol
            Case STATUSINDEX
            Case DATEINDEX
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
            Case TIMEINDEX
            Case ISCIINDEX
            Case MGAIRVEHICLEINDEX
            Case MGMISSEDDATEINDEX
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
        End Select
    End If
End Sub

Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
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
'    gShowBranner
    PostRepTimes.Refresh
    Me.KeyPreview = True
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

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    tmcStart.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer

    On Error Resume Next
    Erase tmCopyNmCode
    Erase tmMissedCode
    Erase tmSbf
    Erase tmExistSbf
    Erase lmDelSdfCode

    btrExtClear hmSdf   'Clear any previous extend operation
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    btrExtClear hmCif   'Clear any previous extend operation
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    btrExtClear hmCpf   'Clear any previous extend operation
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    btrExtClear hmSbf   'Clear any previous extend operation
    ilRet = btrClose(hmSbf)
    btrDestroy hmSbf
    btrExtClear hmRdf   'Clear any previous extend operation
    ilRet = btrClose(hmRdf)
    btrDestroy hmRdf
    btrExtClear hmSmf   'Clear any previous extend operation
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    btrDestroy hmCff
    Set PostRepTimes = Nothing   'Remove data segment
End Sub


Private Sub grdTimes_EnterCell()
    If lmEnableCol = DATEINDEX Then
        If Trim$(grdTimes.TextMatrix(lmEnableRow, DATEINDEX)) = "" Then
            imcTrash_Click
            Exit Sub
        End If
    End If
    mSetShow
End Sub

Private Sub grdTimes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lmTopRow = grdTimes.TopRow
    grdTimes.Redraw = False
End Sub


Private Sub grdTimes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilCol As Integer
    Dim ilRow As Integer

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    imIgnoreScroll = False
    If Y < grdTimes.rowHeight(0) Then
        grdTimes.Col = grdTimes.MouseCol
        mSortCol grdTimes.Col
        Exit Sub
    End If
    pbcArrow.Visible = False
    If igRepSpotBilled Then
        grdTimes.Redraw = True
        Exit Sub
    End If
    ilCol = grdTimes.MouseCol
    ilRow = grdTimes.MouseRow
    If ilCol < grdTimes.FixedCols Then
        grdTimes.Redraw = True
        Exit Sub
    End If
    If ilRow < grdTimes.FixedRows Then
        grdTimes.Redraw = True
        Exit Sub
    End If
    If smAllowRepMG = "N" Then
        If (ilCol = MGAIRVEHICLEINDEX) Or (ilCol = MGMISSEDDATEINDEX) Then
            grdTimes.Redraw = True
            Exit Sub
        End If
    End If
    If grdTimes.TextMatrix(ilRow, STATUSINDEX) = "" Then
        grdTimes.Redraw = False
        Do
            ilRow = ilRow - 1
        Loop While grdTimes.TextMatrix(ilRow, STATUSINDEX) = ""
        grdTimes.Row = ilRow + 1
        grdTimes.Col = DATEINDEX
        grdTimes.Redraw = True
    Else
        grdTimes.Row = ilRow
        grdTimes.Col = ilCol
    End If
    grdTimes.Redraw = True
    lmTopRow = grdTimes.TopRow
    If Not mColOk() Then
        pbcArrow.Move grdTimes.Left - pbcArrow.Width - 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + (grdTimes.rowHeight(grdTimes.Row) - pbcArrow.height) / 2
        pbcArrow.Visible = True
        Exit Sub
    End If
    mEnableBox
End Sub

Private Sub grdTimes_Scroll()
    If imIgnoreScroll Then  'Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdTimes.Redraw = False Then
        grdTimes.Redraw = True
        If lmTopRow < grdTimes.FixedRows Then
            grdTimes.TopRow = grdTimes.FixedRows
        Else
            grdTimes.TopRow = lmTopRow
        End If
        grdTimes.Refresh
        grdTimes.Redraw = False
    End If
    If (imCtrlVisible) And (grdTimes.Row >= grdTimes.FixedRows) And (grdTimes.Col >= 0) And (grdTimes.Col < grdTimes.Cols - 1) Then
        If grdTimes.RowIsVisible(grdTimes.Row) Then
            pbcArrow.Move grdTimes.Left - pbcArrow.Width - 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + (grdTimes.rowHeight(grdTimes.Row) - pbcArrow.height) / 2
            pbcArrow.Visible = True
            mSetFocus
        Else
            pbcSetFocus.SetFocus
            edcDropDown.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        slNameCode                    slCode                    *
'*                                                                                        *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilAdf As Integer

    imFirstActivate = True
    imTerminate = False
    imIgnoreScroll = False
    imFromArrow = False
    imCtrlVisible = False
    imChgMode = False
    imCalType = 0   'Standard
    imDoubleClickName = False
    imLbcArrowSetting = False
    imLbcMouseDown = False
    imBypassFocus = False
    lmEnableRow = -1
    lmEnableCol = -1

    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    'mParseCmmdLine
    PostRepTimes.height = PostRep!pbcPostRep.height 'cmcDone.Top + 5 * cmcDone.Height / 3
    cmcDone.Top = PostRepTimes.height - imcTrash.height - lacSpotStatus.height - 0.5 * cmcDone.height
    cmcUpdate.Top = cmcDone.Top
    cmcCancel.Top = cmcDone.Top
    grdTimes.height = cmcDone.Top - grdTimes.Top - 120
    lacSpotStatus.Top = grdTimes.Top + grdTimes.height + 60
    imcTrash.Top = lacSpotStatus.Top + lacSpotStatus.height + 60
    gCenterStdAlone PostRepTimes
    'PostRepTimes.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imChg = False

    imFirstFocus = True
    imLastSelectRow = 0
    imCtrlKey = False
    hmSdf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", PostRepTimes
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)  'Get and save CHF record length
    hmCif = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", PostRepTimes
    On Error GoTo 0
    imCifRecLen = Len(tmCif)  'Get and save CHF record length
    hmCpf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cpf.Btr)", PostRepTimes
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)  'Get and save CHF record length
    hmSbf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sbf.Btr)", PostRepTimes
    On Error GoTo 0
    ReDim tmSbf(0 To 0) As SBF
    imSbfRecLen = Len(tmSbf(0))  'Get and save CHF record length
    hmRdf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rdf.Btr)", PostRepTimes
    On Error GoTo 0
    imRdfRecLen = Len(tmRdf)  'Get and save CHF record length
    hmSmf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf.Btr)", PostRepTimes
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)  'Get and save CHF record length

    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", PostRepTimes
    On Error GoTo 0

    ilAdf = gBinarySearchAdf(tgChfRep.iAdfCode)
    smBonusOnInv = "Y"
    smAllowRepMG = "Y"
    If ilAdf <> -1 Then
        If (tgCommAdf(ilAdf).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdf).sAddrID) <> "") Then
            slStr = slStr & " " & Trim$(tgCommAdf(ilAdf).sName) & ", " & Trim$(tgCommAdf(ilAdf).sAddrID)
        Else
            slStr = slStr & " " & Trim$(tgCommAdf(ilAdf).sName)
        End If
        smBonusOnInv = Trim$(tgCommAdf(ilAdf).sBonusOnInv)
        smAllowRepMG = tgCommAdf(ilAdf).sAllowRepMG
    End If
    lacContract.Caption = slStr

    mInitBox


    mCopyPop
    mVehPop
    mPopulate
    If imTerminate Then
        Exit Sub
    End If

    slStr = Trim$(str$(tgChfRep.lCntrNo))
    lacScreen.Caption = "Post Rep Times: " & Format$(lgRepSpotEndStd, "mmm, yyyy")
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

    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload PostRepTimes
    igManUnload = NO
End Sub




Private Sub imcTrash_Click()
    Dim llRow As Long
    Dim ilCol As Integer
    Dim llEnableRow As Long

    If (lmEnableRow >= grdTimes.FixedRows) And (lmEnableRow < grdTimes.Rows) Then
        llEnableRow = lmEnableRow
        lmDelSdfCode(UBound(lmDelSdfCode)) = Val(grdTimes.TextMatrix(llEnableRow, SDFCODEINDEX))
        If lmDelSdfCode(UBound(lmDelSdfCode)) > 0 Then
            imChg = True
            ReDim Preserve lmDelSdfCode(0 To UBound(lmDelSdfCode) + 1) As Long
        End If
        grdTimes.Redraw = False
        mSetShow
        For llRow = llEnableRow To grdTimes.Rows - 2 Step 1
            For ilCol = STATUSINDEX To CHGINDEX Step 1
                grdTimes.TextMatrix(llRow, ilCol) = grdTimes.TextMatrix(llRow + 1, ilCol)
                If (ilCol = SDFCODEINDEX) Then
                    grdTimes.TextMatrix(llRow + 1, ilCol) = -1
                Else
                    grdTimes.TextMatrix(llRow + 1, ilCol) = ""
                End If
            Next ilCol
        Next llRow
        pbcClickFocus.SetFocus
        grdTimes.Redraw = True
    End If
End Sub

Private Sub lacScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub lacSpotStatus_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub lbcCopyNm_Click()
    If imLbcMouseDown Then
        imLbcArrowSetting = False
    Else
        imLbcArrowSetting = True
    End If
    gProcessLbcClick lbcCopyNm, edcDropDown, imChgMode, imLbcArrowSetting
    imLbcMouseDown = False
End Sub

Private Sub lbcCopyNm_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcCopyNm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub


Private Sub lbcVehicle_Click()
    If imLbcMouseDown Then
        imLbcArrowSetting = False
    Else
        imLbcArrowSetting = True
    End If
    gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
    imLbcMouseDown = False
End Sub

Private Sub lbcVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcVehicle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    If imFirstFocus Then
        imFirstFocus = False
    End If
    If grdTimes.Visible Then
        mSetShow
        grdTimes.Row = 0
        grdTimes.Col = SDFCODEINDEX
        mSetCommands
    End If
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub mPopulate()

    Dim ilClf As Integer
    Dim ilFound As Integer
    Dim slVehicle As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slDaypartTimes As String
    Dim slDaypartDays As String
    Dim ilDay As Integer
    Dim ilLoop As Integer
    Dim ilVef As Integer
    Dim slSFlightDate As String
    Dim slEFlightDate As String
    Dim ilIncludeFlight As Integer
    Dim slStr As String
    Dim ilCff As Integer
    Dim ilRet As Integer
    Dim ilOrderedSpots As Integer
    Dim llDate As Long

    ReDim imOrderedSpots(0 To 0) As Integer
    cbcLine.Clear
    For ilClf = LBound(tgClfRep) To UBound(tgClfRep) - 1 Step 1
        ilFound = False
        For ilLoop = 0 To UBound(tgRepSpotLineInfo) - 1 Step 1
            If tgClfRep(ilClf).ClfRec.iLine = tgRepSpotLineInfo(ilLoop).iLineNo Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If ilFound Then
            ilVef = gBinarySearchVef(tgClfRep(ilClf).ClfRec.iVefCode)
            If ilVef <> -1 Then
                slStr = Trim$(tgMVef(ilVef).sName)
            End If
            slVehicle = slStr
            If ((tgClfRep(ilClf).ClfRec.iStartTime(0) <> 1) Or (tgClfRep(ilClf).ClfRec.iStartTime(1) <> 0)) Then
                gUnpackTime tgClfRep(ilClf).ClfRec.iStartTime(0), tgClfRep(ilClf).ClfRec.iStartTime(1), "A", "1", slStartTime
                gUnpackTime tgClfRep(ilClf).ClfRec.iEndTime(0), tgClfRep(ilClf).ClfRec.iEndTime(1), "A", "1", slEndTime
                slDaypartTimes = slStartTime & "-" & slEndTime
            Else
                tmRdfSrchKey.iCode = tgClfRep(ilClf).ClfRec.iRdfCode  ' Daypart File Code
                ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    'gUnpackTime tmRdf.iStartTime(0, 7), tmRdf.iStartTime(1, 7), "A", "1", slStartTime
                    gUnpackTime tmRdf.iStartTime(0, 6), tmRdf.iStartTime(1, 6), "A", "1", slStartTime
                    'gUnpackTime tmRdf.iEndTime(0, 7), tmRdf.iEndTime(1, 7), "A", "1", slEndTime
                    gUnpackTime tmRdf.iEndTime(0, 6), tmRdf.iEndTime(1, 6), "A", "1", slEndTime
                    slDaypartTimes = slStartTime & "-" & slEndTime
                End If
            End If
            ilOrderedSpots = 0
            slDaypartDays = String(7, " ")
            ilCff = tgClfRep(ilClf).iFirstCff
            Do While ilCff <> -1
                gUnpackDate tgCffRep(ilCff).CffRec.iStartDate(0), tgCffRep(ilCff).CffRec.iStartDate(1), slSFlightDate
                gUnpackDate tgCffRep(ilCff).CffRec.iEndDate(0), tgCffRep(ilCff).CffRec.iEndDate(1), slEFlightDate
                ilIncludeFlight = True
                If tgChfRep.sBillCycle = "C" Then
                    If (gDateValue(slSFlightDate) > lgRepSpotEndCal) Or (gDateValue(slEFlightDate) < lgRepSpotStartCal) Then
                        ilIncludeFlight = False
                    End If
                Else
                    If (gDateValue(slSFlightDate) > lgRepSpotEndStd) Or (gDateValue(slEFlightDate) < lgRepSpotStartStd) Then
                        ilIncludeFlight = False
                    End If
                End If
                'Test if CBS
                If gDateValue(slEFlightDate) < gDateValue(slSFlightDate) Then
                    ilIncludeFlight = False
                End If
                If ilIncludeFlight Then
                    For ilDay = 0 To 6 Step 1
                        If tgCffRep(ilCff).CffRec.iDay(ilDay) > 0 Then
                             If Trim$(Mid$(slDaypartDays, ilDay + 1, 1)) = "" Then
                                Mid$(slDaypartDays, ilDay + 1, 1) = "Y"
                            ElseIf Trim$(Mid$(slDaypartDays, ilDay + 1, 1)) = "N" Then
                                Mid$(slDaypartDays, ilDay + 1, 1) = "M"
                            End If
                       Else
                            If Trim$(Mid$(slDaypartDays, ilDay + 1, 1)) = "" Then
                                Mid$(slDaypartDays, ilDay + 1, 1) = "N"
                            ElseIf Trim$(Mid$(slDaypartDays, ilDay + 1, 1)) = "Y" Then
                                Mid$(slDaypartDays, ilDay + 1, 1) = "M"
                            End If
                        End If
                    Next ilDay
                    'If tgCffRep(ilCff).CffRec.sDyWk <> "D" Then
                    '    ilOrderedSpots = ilOrderedSpots + tgCffRep(ilCff).CffRec.iSpotsWk
                    'Else
                    '    For ilDay = 0 To 6 Step 1
                    '        ilOrderedSpots = ilOrderedSpots + tgCffRep(ilCff).CffRec.iDay(ilDay)
                    '    Next ilDay
                    'End If
                    If tgCffRep(ilCff).CffRec.sDyWk <> "D" Then
                        For llDate = gDateValue(slSFlightDate) To gDateValue(slEFlightDate) Step 7
                            ilIncludeFlight = True
                            If tgChfRep.sBillCycle = "C" Then
                                If (llDate > lgRepSpotEndCal) Or (llDate < lgRepSpotStartCal) Then
                                    ilIncludeFlight = False
                                End If
                            Else
                                If (llDate > lgRepSpotEndStd) Or (llDate < lgRepSpotStartStd) Then
                                    ilIncludeFlight = False
                                End If
                            End If
                            If ilIncludeFlight Then
                                ilOrderedSpots = ilOrderedSpots + tgCffRep(ilCff).CffRec.iSpotsWk
                            End If
                        Next llDate
                    Else
                        For llDate = gDateValue(slSFlightDate) To gDateValue(slEFlightDate) Step 1
                            ilIncludeFlight = True
                            If tgChfRep.sBillCycle = "C" Then
                                If (llDate > lgRepSpotEndCal) Or (llDate < lgRepSpotStartCal) Then
                                    ilIncludeFlight = False
                                End If
                            Else
                                If (llDate > lgRepSpotEndStd) Or (llDate < lgRepSpotStartStd) Then
                                    ilIncludeFlight = False
                                End If
                            End If
                            If ilIncludeFlight Then
                                ilDay = gWeekDayLong(llDate)
                                ilOrderedSpots = ilOrderedSpots + tgCffRep(ilCff).CffRec.iDay(ilDay)
                            End If
                        Next llDate
                    End If
                End If
                ilCff = tgCffRep(ilCff).iNextCff
            Loop
            slStr = Trim$(str$(tgClfRep(ilClf).ClfRec.iLine)) & ": " & slVehicle & ", " & slDaypartTimes & " " & slDaypartDays & ", " & Trim$(str$(tgClfRep(ilClf).ClfRec.iLen)) & "s" & " Ordered: " & ilOrderedSpots & " spots"
            cbcLine.AddItem slStr
            cbcLine.ItemData(cbcLine.NewIndex) = ilClf
            imOrderedSpots(UBound(imOrderedSpots)) = ilOrderedSpots
            ReDim Preserve imOrderedSpots(0 To UBound(imOrderedSpots) + 1) As Integer
        End If
    Next ilClf
    For ilLoop = 0 To cbcLine.ListCount - 1 Step 1
        If igRepSpotLineNo = tgClfRep(cbcLine.ItemData(ilLoop)).ClfRec.iLine Then
            cbcLine.ListIndex = ilLoop
        End If
    Next ilLoop
End Sub


Private Sub mSetCommands()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         ilRow                                                   *
'******************************************************************************************


    If imChg Then
        cmcUpdate.Enabled = True
        cbcLine.Enabled = False
    Else
        cmcUpdate.Enabled = False
        cbcLine.Enabled = True
    End If
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  flTextHeight                  ilRow                         ilCol                     *
'*                                                                                        *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
    Dim ilLoop As Integer
    'flTextHeight = pbcDates.TextHeight("1") - 35

    mGridLayout
    mGridColumnWidths
    mGridColumns
    grdTimes.Move 180, cbcLine.Top + cbcLine.height + 120, grdTimes.Width, cmcDone.Top - (lacScreen.Top + lacScreen.height) - 240
    'grdTimes.Height = grdTimes.RowPos(0) + 14 * grdTimes.RowHeight(0) + fgPanelAdj - 15
    imInitNoRows = (cmcDone.Top - 120 - grdTimes.Top) \ fgFlexGridRowH
    grdTimes.height = grdTimes.RowPos(0) + imInitNoRows * (fgFlexGridRowH) + fgPanelAdj - 15
    mClearGrid
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
End Sub

Private Sub mGridLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdTimes.Rows - 1 Step 1
        grdTimes.rowHeight(ilRow) = fgFlexGridRowH
    Next ilRow
    For ilCol = 0 To grdTimes.Cols - 1 Step 1
        grdTimes.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdTimes.Row = grdTimes.FixedRows - 1
    grdTimes.Col = STATUSINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Status"
    grdTimes.Col = DATEINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Aired Date"
    grdTimes.Col = TIMEINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Aired Time"
    grdTimes.Col = ISCIINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "ISCI"
    grdTimes.Col = SDFCODEINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Sdf Code"
    grdTimes.Col = MGAIRVEHICLEINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Outside/MG Vehicle"
    grdTimes.Col = MGMISSEDDATEINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Missed Date"
    grdTimes.Col = SORTINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Sort"
    grdTimes.Col = CHGINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Changed"

End Sub

Private Sub mGridColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdTimes.ColWidth(SDFCODEINDEX) = 0
    grdTimes.ColWidth(SORTINDEX) = 0
    grdTimes.ColWidth(CHGINDEX) = 0
    grdTimes.ColWidth(STATUSINDEX) = 0.1 * grdTimes.Width
    grdTimes.ColWidth(DATEINDEX) = 0.1 * grdTimes.Width
    grdTimes.ColWidth(TIMEINDEX) = 0.12 * grdTimes.Width
    grdTimes.ColWidth(ISCIINDEX) = 0.25 * grdTimes.Width
    grdTimes.ColWidth(MGAIRVEHICLEINDEX) = 0.25 * grdTimes.Width
    grdTimes.ColWidth(MGMISSEDDATEINDEX) = 0.1 * grdTimes.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdTimes.Width
    For ilCol = 0 To grdTimes.Cols - 1 Step 1
        llWidth = llWidth + grdTimes.ColWidth(ilCol)
        If (grdTimes.ColWidth(ilCol) > 15) And (grdTimes.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdTimes.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdTimes.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdTimes.Width
            For ilCol = 0 To grdTimes.Cols - 1 Step 1
                If (grdTimes.ColWidth(ilCol) > 15) And (grdTimes.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdTimes.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdTimes.FixedCols To grdTimes.Cols - 1 Step 1
                If grdTimes.ColWidth(ilCol) > 15 Then
                    ilColInc = grdTimes.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdTimes.ColWidth(ilCol) = grdTimes.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub


Private Sub mSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdTimes.FixedRows To grdTimes.Rows - 1 Step 1
        slStr = Trim$(grdTimes.TextMatrix(llRow, STATUSINDEX))
        If slStr <> "" Then
            If ilCol = STATUSINDEX Then
                slSort = grdTimes.TextMatrix(llRow, STATUSINDEX)
                Do While Len(slSort) < 10
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = DATEINDEX Then
                slSort = Trim$(str$(gDateValue(grdTimes.TextMatrix(llRow, DATEINDEX))))
                Do While Len(slSort) < 10
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = TIMEINDEX Then
                slSort = Trim$(str$(gTimeToLong(grdTimes.TextMatrix(llRow, TIMEINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = ISCIINDEX) Then
                slSort = grdTimes.TextMatrix(llRow, ISCIINDEX)
                Do While Len(slSort) < 20
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = MGAIRVEHICLEINDEX) Then
                slSort = grdTimes.TextMatrix(llRow, MGAIRVEHICLEINDEX)
                Do While Len(slSort) < 20
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = MGMISSEDDATEINDEX Then
                slSort = Trim$(str$(gDateValue(grdTimes.TextMatrix(llRow, MGMISSEDDATEINDEX))))
                Do While Len(slSort) < 10
                    slSort = "0" & slSort
                Loop
            End If
            slStr = grdTimes.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastColSorted) Or ((ilCol = imLastColSorted) And (imLastSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdTimes.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdTimes.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastColSorted Then
        imLastColSorted = SORTINDEX
    Else
        imLastColSorted = -1
        imLastSort = -1
    End If
    gGrid_SortByCol grdTimes, STATUSINDEX, SORTINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox()
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If (grdTimes.Row < grdTimes.FixedRows) Or (grdTimes.Row >= grdTimes.Rows) Or (grdTimes.Col < grdTimes.FixedCols) Or (grdTimes.Col >= grdTimes.Cols - 1) Then
        Exit Sub
    End If
    If igRepSpotBilled Then
        lmEnableRow = -1
        lmEnableCol = -1
        Exit Sub
    End If
    lmEnableRow = grdTimes.Row
    lmEnableCol = grdTimes.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdTimes.Left - pbcArrow.Width - 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + (grdTimes.rowHeight(grdTimes.Row) - pbcArrow.height) / 2
    pbcArrow.Visible = True
    imCtrlVisible = True
    slStr = grdTimes.TextMatrix(lmEnableRow, lmEnableCol)
    Select Case grdTimes.Col
        Case STATUSINDEX
            If Trim$(grdTimes.TextMatrix(grdTimes.Row, STATUSINDEX)) = "Bonus" Then
                imStatus = 1
            Else
                imStatus = 0
            End If
        Case DATEINDEX
            edcDropDown.MaxLength = 10
            slStr = grdTimes.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdTimes.Row > grdTimes.FixedRows Then
                    slStr = grdTimes.TextMatrix(grdTimes.Row - 1, grdTimes.Col)
                End If
            End If
            If slStr <> "" Then
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint
            Else
                slStr = Format$(lgRepSpotEndStd, "m/d/yy")
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                slStr = ""
            End If
            edcDropDown.Text = slStr
        Case TIMEINDEX
            edcDropDown.MaxLength = 11
            slStr = grdTimes.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            edcDropDown.Text = slStr
        Case ISCIINDEX
            lbcCopyNm.height = gListBoxHeight(lbcCopyNm.ListCount, 10)
            edcDropDown.MaxLength = 20
            imChgMode = True
            gFindMatch slStr, 0, lbcCopyNm
            If gLastFound(lbcCopyNm) >= 0 Then
                lbcCopyNm.ListIndex = gLastFound(lbcCopyNm)
                edcDropDown.Text = lbcCopyNm.List(lbcCopyNm.ListIndex)
            Else
                edcDropDown.Text = slStr
            End If
            imChgMode = False
        Case MGAIRVEHICLEINDEX
            lbcVehicle.height = gListBoxHeight(lbcVehicle.ListCount, 10)
            edcDropDown.MaxLength = 35
            imChgMode = True
            gFindMatch slStr, 1, lbcVehicle
            If gLastFound(lbcVehicle) >= 1 Then
                lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
            Else
                gFindMatch smLnVehicle, 1, lbcVehicle
                If gLastFound(lbcVehicle) >= 1 Then
                    lbcVehicle.ListIndex = gLastFound(lbcVehicle)
                    edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
                Else
                    edcDropDown.Text = slStr
                End If
            End If
            imChgMode = False
        Case MGMISSEDDATEINDEX
            edcDropDown.MaxLength = 10
            slStr = grdTimes.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdTimes.Row > grdTimes.FixedRows Then
                    slStr = grdTimes.TextMatrix(grdTimes.Row - 1, grdTimes.Col)
                End If
            End If
            If slStr <> "" Then
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint
            Else
                slStr = Format$(lgRepSpotEndStd, "m/d/yy")
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                slStr = ""
            End If
            edcDropDown.Text = slStr
    End Select
    mSetFocus
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer

    If (grdTimes.Row < grdTimes.FixedRows) Or (grdTimes.Row >= grdTimes.Rows) Or (grdTimes.Col < grdTimes.FixedCols) Or (grdTimes.Col >= grdTimes.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    pbcArrow.Visible = False
    pbcArrow.Move grdTimes.Left - pbcArrow.Width - 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + (grdTimes.rowHeight(grdTimes.Row) - pbcArrow.height) / 2
    pbcArrow.Visible = True
    llColPos = 0
    For ilCol = 0 To grdTimes.Col - 1 Step 1
        llColPos = llColPos + grdTimes.ColWidth(ilCol)
    Next ilCol
    Select Case grdTimes.Col
        Case STATUSINDEX
            pbcPaid.Move grdTimes.Left + llColPos + 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + 30, grdTimes.ColWidth(grdTimes.Col) - 30, grdTimes.rowHeight(grdTimes.Row) - 15
            pbcPaid.Visible = True
            pbcPaid.SetFocus
        Case DATEINDEX
            edcDropDown.Move grdTimes.Left + llColPos + 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + 30, grdTimes.ColWidth(grdTimes.Col) - 30, grdTimes.rowHeight(grdTimes.Row) - 15
            If edcDropDown.Top + edcDropDown.height + plcCalendar.height < cmcDone.Top Then
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            Else
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.height
            End If
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TIMEINDEX
            edcDropDown.Move grdTimes.Left + llColPos + 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + 30, grdTimes.ColWidth(grdTimes.Col) - 30, grdTimes.rowHeight(grdTimes.Row) - 15
            If edcDropDown.Top + edcDropDown.height + plcTme.height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            Else
                plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.height
            End If
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ISCIINDEX
            edcDropDown.Move grdTimes.Left + llColPos + 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + 30, grdTimes.ColWidth(grdTimes.Col) - 30, grdTimes.rowHeight(grdTimes.Row) - 15
            If edcDropDown.Top + edcDropDown.height + lbcCopyNm.height < cmcDone.Top Then
                lbcCopyNm.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            Else
                lbcCopyNm.Move edcDropDown.Left, edcDropDown.Top - lbcCopyNm.height
            End If
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case MGAIRVEHICLEINDEX
            edcDropDown.Move grdTimes.Left + llColPos + 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + 30, grdTimes.ColWidth(grdTimes.Col) - 30, grdTimes.rowHeight(grdTimes.Row) - 15
            If edcDropDown.Top + edcDropDown.height + lbcVehicle.height < cmcDone.Top Then
                lbcVehicle.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - lbcVehicle.Width, edcDropDown.Top + edcDropDown.height
            Else
                lbcVehicle.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - lbcVehicle.Width, edcDropDown.Top - lbcVehicle.height
            End If
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case MGMISSEDDATEINDEX
            edcDropDown.Move grdTimes.Left + llColPos + 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + 30, grdTimes.ColWidth(grdTimes.Col) - 30, grdTimes.rowHeight(grdTimes.Row) - 15
            If edcDropDown.Top + edcDropDown.height + plcCalendar.height < cmcDone.Top Then
                plcCalendar.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - plcCalendar.Width, edcDropDown.Top + edcDropDown.height
            Else
                plcCalendar.Move edcDropDown.Left + edcDropDown.Width + cmcDropDown.Width - plcCalendar.Width, edcDropDown.Top - plcCalendar.height
            End If
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slAudience                    slPopulation                  slRating                  *
'*                                                                                        *
'******************************************************************************************

'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String

    pbcArrow.Visible = False
    If (lmEnableRow >= grdTimes.FixedRows) And (lmEnableRow < grdTimes.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case STATUSINDEX
                pbcPaid.Visible = False  'Set visibility
                If imStatus = 1 Then
                    slStr = "Bonus"
                Else
                    slStr = "Paid"
                End If
                If StrComp(grdTimes.TextMatrix(lmEnableRow, lmEnableCol), slStr, vbTextCompare) <> 0 Then
                    imChg = True
                    grdTimes.TextMatrix(lmEnableRow, CHGINDEX) = "1"
                End If
                grdTimes.TextMatrix(lmEnableRow, lmEnableCol) = slStr
             Case DATEINDEX
                edcDropDown.Visible = False  'Set visibility
                plcCalendar.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                If gDateValue(grdTimes.TextMatrix(lmEnableRow, lmEnableCol)) <> gDateValue(slStr) Then
                    imChg = True
                    grdTimes.TextMatrix(lmEnableRow, CHGINDEX) = "1"
                End If
                grdTimes.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                If slStr <> "" Then
                    If grdTimes.TextMatrix(lmEnableRow, STATUSINDEX) = "" Then
                        grdTimes.TextMatrix(lmEnableRow, STATUSINDEX) = "Paid"
                    End If
                End If
            Case TIMEINDEX
                edcDropDown.Visible = False  'Set visibility
                plcTme.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                If gTimeToLong(grdTimes.TextMatrix(lmEnableRow, lmEnableCol), False) <> gTimeToLong(slStr, False) Then
                    imChg = True
                    grdTimes.TextMatrix(lmEnableRow, CHGINDEX) = "1"
                End If
                grdTimes.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                If slStr <> "" Then
                    If grdTimes.TextMatrix(lmEnableRow, STATUSINDEX) = "" Then
                        grdTimes.TextMatrix(lmEnableRow, STATUSINDEX) = "Paid"
                    End If
                End If
            Case ISCIINDEX
                edcDropDown.Visible = False  'Set visibility
                cmcDropDown.Visible = False
                lbcCopyNm.Visible = False
                slStr = Trim$(edcDropDown.Text)
                If StrComp(grdTimes.TextMatrix(lmEnableRow, lmEnableCol), slStr, vbTextCompare) <> 0 Then
                    imChg = True
                    grdTimes.TextMatrix(lmEnableRow, CHGINDEX) = "1"
                End If
                grdTimes.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                If slStr <> "" Then
                    gFindMatch Trim$(slStr), 0, lbcCopyNm
                    If gLastFound(lbcCopyNm) < 0 Then
                        lbcCopyNm.RemoveItem 0
                        lbcCopyNm.AddItem slStr
                        lbcCopyNm.ItemData(lbcCopyNm.NewIndex) = 0
                        lbcCopyNm.AddItem "[None]", 0
                    End If
                End If
                If slStr <> "" Then
                    If grdTimes.TextMatrix(lmEnableRow, STATUSINDEX) = "" Then
                        grdTimes.TextMatrix(lmEnableRow, STATUSINDEX) = "Paid"
                    End If
                End If
            Case MGAIRVEHICLEINDEX
                edcDropDown.Visible = False  'Set visibility
                cmcDropDown.Visible = False
                lbcVehicle.Visible = False
                slStr = Trim$(edcDropDown.Text)
                If StrComp(grdTimes.TextMatrix(lmEnableRow, lmEnableCol), slStr, vbTextCompare) <> 0 Then
                    imChg = True
                    grdTimes.TextMatrix(lmEnableRow, CHGINDEX) = "1"
                End If
                grdTimes.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case MGMISSEDDATEINDEX
                edcDropDown.Visible = False  'Set visibility
                plcCalendar.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                If gDateValue(grdTimes.TextMatrix(lmEnableRow, lmEnableCol)) <> gDateValue(slStr) Then
                    imChg = True
                    grdTimes.TextMatrix(lmEnableRow, CHGINDEX) = "1"
                End If
                grdTimes.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                If slStr <> "" Then
                    If grdTimes.TextMatrix(lmEnableRow, STATUSINDEX) = "" Then
                        grdTimes.TextMatrix(lmEnableRow, STATUSINDEX) = "Paid"
                    End If
                    If grdTimes.TextMatrix(lmEnableRow, MGAIRVEHICLEINDEX) = "" Then
                        grdTimes.TextMatrix(lmEnableRow, MGAIRVEHICLEINDEX) = smLnVehicle
                    End If
                End If
        End Select
    End If
    pbcArrow.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    mGetTotals
    mSetCommands
End Sub

Private Sub pbcPaid_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcPaid_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("P") Or (KeyAscii = Asc("p")) Then
        imStatus = 0
        pbcPaid_Paint
    ElseIf KeyAscii = Asc("B") Or (KeyAscii = Asc("b")) Then
        imStatus = 1
        pbcPaid_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imStatus = 0 Then
            imStatus = 1
            pbcPaid_Paint
        ElseIf imStatus = 1 Then
            imStatus = 0
            pbcPaid_Paint
        End If
    End If
End Sub

Private Sub pbcPaid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imStatus = 0 Then
        imStatus = 1
    ElseIf imStatus = 1 Then
        imStatus = 0
    End If
    pbcPaid_Paint
End Sub

Private Sub pbcPaid_Paint()
    pbcPaid.Cls
    pbcPaid.CurrentX = fgBoxInsetX
    pbcPaid.CurrentY = 0 'fgBoxInsetY
    If imStatus = 0 Then
        pbcPaid.Print "Paid"
    ElseIf imStatus = 1 Then
        pbcPaid.Print "Bonus"
    Else
        pbcPaid.Print "   "
    End If
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilPrev As Integer

    If GetFocus() <> pbcSTab.HWnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mEnableBox
        Exit Sub
    End If
    If imCtrlVisible Then
        If (lmEnableCol = DATEINDEX) And ((lmEnableRow >= grdTimes.FixedRows) And (lmEnableRow < grdTimes.Rows)) Then
            If Trim$(edcDropDown.Text) = "" Then
                imcTrash_Click
                Exit Sub
            End If
        End If
        mSetShow
        Do
            ilPrev = False
            If (grdTimes.Col = DATEINDEX) Or (grdTimes.Col = STATUSINDEX) Then
                If grdTimes.Row > grdTimes.FixedRows Then
                    lmTopRow = -1
                    grdTimes.Row = grdTimes.Row - 1
                    If Not grdTimes.RowIsVisible(grdTimes.Row) Then
                        grdTimes.TopRow = grdTimes.TopRow - 1
                    End If
                    grdTimes.Col = ISCIINDEX
                    mEnableBox
                Else

                    cmcCancel.SetFocus
                End If
            Else
                grdTimes.Col = grdTimes.Col - 1
                If mColOk() Then
                    mEnableBox
                Else
                    ilPrev = True
                End If
            End If
        Loop While ilPrev
    Else
        lmTopRow = -1
        grdTimes.TopRow = grdTimes.FixedRows
        grdTimes.Col = DATEINDEX
        grdTimes.Row = grdTimes.FixedRows
        If mColOk() Then
            mEnableBox
        Else
            cmcCancel.SetFocus
        End If
    End If
End Sub

Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    Dim ilCol As Integer
    Dim ilNext As Integer
    Dim llEnableRow As Long

    If GetFocus() <> pbcTab.HWnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        llEnableRow = lmEnableRow
        If (lmEnableCol = DATEINDEX) And ((lmEnableRow >= grdTimes.FixedRows) And (lmEnableRow < grdTimes.Rows)) Then
            If Trim$(edcDropDown.Text) = "" Then
                imcTrash_Click
                Exit Sub
            End If
        End If
        mSetShow
        Do
            ilNext = False
            If (grdTimes.Col = ISCIINDEX) Or ((grdTimes.Col = MGAIRVEHICLEINDEX) And ((edcDropDown.Text) = "")) Or (grdTimes.Col = MGMISSEDDATEINDEX) Then
                llRow = grdTimes.Rows
                Do
                    llRow = llRow - 1
                Loop While grdTimes.TextMatrix(llRow, STATUSINDEX) = ""
                llRow = llRow + 1
                If (grdTimes.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdTimes.Row = grdTimes.Row + 1
                    If Not grdTimes.RowIsVisible(grdTimes.Row) Or (grdTimes.Row - (grdTimes.TopRow - grdTimes.FixedRows) >= imInitNoRows) Then
                        imIgnoreScroll = True
                        grdTimes.TopRow = grdTimes.TopRow + 1
                    End If
                    grdTimes.Col = DATEINDEX
                    'grdTimes.TextMatrix(grdTimes.Row, CODEINDEX) = 0
                    If Trim$(grdTimes.TextMatrix(grdTimes.Row, STATUSINDEX)) <> "" Then
                        If mColOk() Then
                            mEnableBox
                        Else
                            cmcCancel.SetFocus
                        End If
                    Else
                        imFromArrow = True
                        pbcArrow.Move grdTimes.Left - pbcArrow.Width - 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + (grdTimes.rowHeight(grdTimes.Row) - pbcArrow.height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    End If
                Else
                    If Trim$(grdTimes.TextMatrix(llEnableRow, STATUSINDEX)) <> "" Then
                        lmTopRow = -1
                        If grdTimes.Row + 1 >= grdTimes.Rows Then
                            grdTimes.AddItem ""
                            grdTimes.rowHeight(grdTimes.Row + 1) = fgFlexGridRowH
                            grdTimes.TextMatrix(grdTimes.Row + 1, SDFCODEINDEX) = -1
                            grdTimes.TextMatrix(grdTimes.Row + 1, CHGINDEX) = "0"
                            If smAllowRepMG = "N" Then
                                For ilCol = MGAIRVEHICLEINDEX To MGMISSEDDATEINDEX Step 1
                                    grdTimes.Col = ilCol
                                    grdTimes.Row = grdTimes.Row + 1
                                    grdTimes.CellBackColor = LIGHTYELLOW
                                Next ilCol
                            End If
                        End If
                        grdTimes.Row = grdTimes.Row + 1
                        If (Not grdTimes.RowIsVisible(grdTimes.Row)) Or (grdTimes.Row - (grdTimes.TopRow - grdTimes.FixedRows) >= imInitNoRows) Then
                            imIgnoreScroll = True
                            grdTimes.TopRow = grdTimes.TopRow + 1
                        End If
                        grdTimes.Col = DATEINDEX
                        grdTimes.TextMatrix(grdTimes.Row, SDFCODEINDEX) = -1
                        grdTimes.TextMatrix(grdTimes.Row, CHGINDEX) = "0"
                        'mEnableBox
                        imFromArrow = True
                        pbcArrow.Move grdTimes.Left - pbcArrow.Width - 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + (grdTimes.rowHeight(grdTimes.Row) - pbcArrow.height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    Else
                        pbcClickFocus.SetFocus
                    End If
                End If
            Else
                grdTimes.Col = grdTimes.Col + 1
                If mColOk() Then
                    mEnableBox
                Else
                    ilNext = True
                End If
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdTimes.TopRow = grdTimes.FixedRows
        grdTimes.Col = ISCIINDEX
        grdTimes.Row = grdTimes.FixedRows
        If mColOk() Then
            mEnableBox
        Else
            cmcCancel.SetFocus
        End If
    End If
End Sub

Private Function mSaveRec() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slMsg                         llLoop                                                  *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mSaveRecErr                                                                           *
'******************************************************************************************

    Dim llRow As Long
    Dim ilRet As Integer
    Dim ilError As Integer
    Dim slStr As String
    Dim slAiredPrice As String
    Dim llAiredPrice As Long
    Dim tlSdf As SDF
    Dim tlClf As CLF
    Dim tlSbf As SBF
    'TTP 10855 - fix potential NTR Overflow errors
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim slDate As String
    Dim ilMissedVefCode As Integer
    Dim slMissedDate As String
    Dim slSchDate As String
    Dim slSchTime As String
    Dim slType As String
    Dim ilSdf As Integer

    If Not imUpdateAllowed Then
        mSaveRec = False
        Exit Function
    End If
    ilError = False
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdTimes, grdTimes, vbHourglass
    For llRow = grdTimes.FixedRows To grdTimes.Rows - 1 Step 1
        If mGridFieldsOk(llRow) = False Then
            ilError = True
        End If
    Next llRow
    If (ilError) Then
        gSetMousePointer grdTimes, grdTimes, vbDefault
        Screen.MousePointer = vbDefault
        Beep
        MsgBox "Check input fields marked in Red as in error", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
        mSaveRec = False
        Exit Function
    End If
    ReDim tmSbf(0 To 0) As SBF
    For llRow = grdTimes.FixedRows To grdTimes.Rows - 1 Step 1
        slStr = Trim$(grdTimes.TextMatrix(llRow, STATUSINDEX))
        If slStr <> "" Then
            slDate = Trim$(grdTimes.TextMatrix(llRow, DATEINDEX))
            If slDate <> "" Then
                mInitSdf llRow
                If tmSdf.lCode > 0 Then
                    tmSdfSrchKey3.lCode = tmSdf.lCode
                    ilRet = btrGetEqual(hmSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                    End If
                Else
                    tlSdf.lSmfCode = 0
                    ilRet = btrInsert(hmSdf, tmSdf, imSdfRecLen, INDEXKEY3)
                End If
                slStr = Trim$(grdTimes.TextMatrix(llRow, MGMISSEDDATEINDEX))
                If slStr <> "" Then
                    If tlSdf.lSmfCode > 0 Then
                        tmSmfSrchKey2.lCode = tlSdf.lCode
                        ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    Else
                        tmSmf.lChfCode = 0
                    End If
                    ilMissedVefCode = tgClfRep(cbcLine.ItemData(cbcLine.ListIndex)).ClfRec.iVefCode
                    slMissedDate = Trim$(grdTimes.TextMatrix(llRow, MGMISSEDDATEINDEX))
                    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slSchDate
                    gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slSchTime
                    slType = "O"
                    ilRet = gMakeSmf(hmSmf, tmSmf, slType, tmSdf, ilMissedVefCode, slMissedDate, smMissedTime, 0, slSchDate, slSchTime)
                    If ilRet Then
                        tmSdfSrchKey3.lCode = tmSdf.lCode
                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet = BTRV_ERR_NONE Then
                            tmSdf.lSmfCode = tmSmf.lCode
                            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                        End If
                    End If
                Else
                    If tlSdf.lSmfCode > 0 Then
                        tmSmfSrchKey2.lCode = tlSdf.lCode
                        ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            ilRet = btrDelete(hmSmf)
                        End If
                    End If
                End If
                If Left(Trim$(grdTimes.TextMatrix(llRow, STATUSINDEX)), 1) = "B" Then
                    llAiredPrice = 0
                Else
                    tlClf = tgClfRep(cbcLine.ItemData(cbcLine.ListIndex)).ClfRec
                    ilRet = gGetFlightPrice(tmSdf, tlClf, hmCff, hmSmf, slAiredPrice)
                    If InStr(1, slAiredPrice, ".", vbTextCompare) <= 0 Then
                        llAiredPrice = 0
                    Else
                        llAiredPrice = gStrDecToLong(slAiredPrice, 2)
                    End If
                End If
                If tmSdf.iMnfMissed <= 0 Then
                    mInitSbf llAiredPrice, slDate, Left(Trim$(grdTimes.TextMatrix(llRow, STATUSINDEX)), 1)
                End If
            End If
        End If
    Next llRow
    For ilSdf = 0 To UBound(lmDelSdfCode) - 1 Step 1
        tmSdfSrchKey3.lCode = lmDelSdfCode(ilSdf)
        ilRet = btrGetEqual(hmSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmSdf)
        End If
    Next ilSdf
    For llSbf = 0 To UBound(tmExistSbf) - 1 Step 1
        tmSbfSrchKey1.lCode = tmExistSbf(llSbf).lCode
        ilRet = btrGetEqual(hmSbf, tlSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmSbf)
        End If
    Next llSbf
    For llSbf = 0 To UBound(tmSbf) - 1 Step 1
        tmSbf(llSbf).lGross = tmSbf(llSbf).lSpotPrice * tmSbf(llSbf).iAirNoSpots
        ilRet = btrInsert(hmSbf, tmSbf(llSbf), imSbfRecLen, INDEXKEY1)
    Next llSbf
    imChg = False
    mSaveRec = True
    gSetMousePointer grdTimes, grdTimes, vbDefault
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr: 'VBC NR
    On Error GoTo 0
    ilRet = btrAbortTrans(hmCif)
    gSetMousePointer grdTimes, grdTimes, vbDefault
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function



'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer record values to      *
'*                      controls on the screen         *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llRet                         ilStatus                                                *
'******************************************************************************************

    Dim llRow As Long
    Dim ilCol As Integer
    Dim llLoop As Long
    Dim slStr As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim ilVeh As Integer
    Dim llColor As Long

    grdTimes.Redraw = False
    llRow = grdTimes.FixedRows
    llColor = grdTimes.ForeColor

    For llLoop = 0 To UBound(tmRepSdf) - 1 Step 1
        If llRow >= grdTimes.Rows Then
            grdTimes.AddItem ""
            grdTimes.rowHeight(llRow) = fgFlexGridRowH
        End If
        grdTimes.TextMatrix(llRow, STATUSINDEX) = ""
        grdTimes.TextMatrix(llRow, MGAIRVEHICLEINDEX) = ""
        grdTimes.TextMatrix(llRow, MGMISSEDDATEINDEX) = ""
        If tmRepSdf(llLoop).sSpotType = "X" Then
            grdTimes.TextMatrix(llRow, STATUSINDEX) = "Bonus"
        Else
            grdTimes.TextMatrix(llRow, STATUSINDEX) = "Paid"
        End If
        If (tmRepSdf(llLoop).sSchStatus = "G") Or (tmRepSdf(llLoop).sSchStatus = "O") Then
            tmSmfSrchKey2.lCode = tmRepSdf(llLoop).lCode
            ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slDate
            Else
                slDate = ""
            End If
            grdTimes.TextMatrix(llRow, MGMISSEDDATEINDEX) = slDate
            For ilVeh = 0 To lbcVehicle.ListCount - 1 Step 1
                If Val(lbcVehicle.ItemData(ilVeh)) = tmRepSdf(llLoop).iVefCode Then
                    grdTimes.TextMatrix(llRow, MGAIRVEHICLEINDEX) = Trim$(lbcVehicle.List(ilVeh))
                    Exit For
                End If
            Next ilVeh
        End If
        gUnpackDate tmRepSdf(llLoop).iDate(0), tmRepSdf(llLoop).iDate(1), slStr
        grdTimes.TextMatrix(llRow, DATEINDEX) = slStr
        gUnpackTime tmRepSdf(llLoop).iTime(0), tmRepSdf(llLoop).iTime(1), "A", "1", slStr
        grdTimes.TextMatrix(llRow, TIMEINDEX) = slStr
        grdTimes.TextMatrix(llRow, ISCIINDEX) = ""
        If tmRepSdf(llLoop).sPtType = "1" Then  '  Single Copy
            ' Read CIF using lCopyCode from SDF
            tmCifSrchKey.lCode = tmRepSdf(llLoop).lCopyCode
            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                If tmCif.lcpfCode > 0 Then
                    tmCpfSrchKey.lCode = tmCif.lcpfCode
                    ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        grdTimes.TextMatrix(llRow, ISCIINDEX) = Trim$(tmCpf.sISCI)
                    End If
                End If
            End If
        End If
        grdTimes.TextMatrix(llRow, SDFCODEINDEX) = Trim$(str$(tmRepSdf(llLoop).lCode))
        If smAllowRepMG = "N" Then
            For ilCol = MGAIRVEHICLEINDEX To MGMISSEDDATEINDEX Step 1
                grdTimes.Col = ilCol
                grdTimes.Row = llRow
                grdTimes.CellBackColor = LIGHTYELLOW
            Next ilCol
        End If
        If (tmRepSdf(llLoop).sBill = "Y") Or (igRepSpotBilled) Then
            For ilCol = STATUSINDEX To MGMISSEDDATEINDEX Step 1
                grdTimes.Col = ilCol
                grdTimes.Row = llRow
                grdTimes.CellForeColor = DARKGREEN
            Next ilCol
        End If
        grdTimes.ForeColor = llColor
        llRow = llRow + 1
    Next llLoop
    If llRow >= grdTimes.Rows Then
        grdTimes.AddItem ""
        grdTimes.rowHeight(llRow) = fgFlexGridRowH
        grdTimes.TextMatrix(llRow, SDFCODEINDEX) = -1
        If smAllowRepMG = "N" Then
            For ilCol = MGAIRVEHICLEINDEX To MGMISSEDDATEINDEX Step 1
                grdTimes.Col = ilCol
                grdTimes.Row = llRow
                grdTimes.CellBackColor = LIGHTYELLOW
            Next ilCol
        End If
    End If
    'Remove highlight
    imLastColSorted = -1
    mSortCol TIMEINDEX
    mSortCol DATEINDEX
    grdTimes.Row = 0
    grdTimes.Col = SDFCODEINDEX
    grdTimes.Redraw = True
    mGetTotals
    mSetCommands

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGridFieldsOk                   *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mGridFieldsOk(llRowNo As Long) As Integer
'
'   iRet = mGridFieldsOk()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim slStr As String
    Dim ilError As Integer

    ilError = False
    slStr = Trim$(grdTimes.TextMatrix(llRowNo, STATUSINDEX))
    If slStr <> "" Then
        slStr = Trim$(grdTimes.TextMatrix(llRowNo, DATEINDEX))
        If slStr <> "" Then
            If Not gValidDate(slStr) Then
                ilError = True
                grdTimes.Row = llRowNo
                grdTimes.Col = DATEINDEX
                grdTimes.CellForeColor = vbRed
            End If
            slStr = Trim$(grdTimes.TextMatrix(llRowNo, TIMEINDEX))
            If slStr = "" Then
                grdTimes.TextMatrix(llRowNo, TIMEINDEX) = "Missing"
                ilError = True
                grdTimes.Row = llRowNo
                grdTimes.Col = TIMEINDEX
                grdTimes.CellForeColor = vbRed
            Else
                If Not gValidTime(slStr) Then
                    ilError = True
                    grdTimes.Row = llRowNo
                    grdTimes.Col = TIMEINDEX
                    grdTimes.CellForeColor = vbRed
                End If
            End If
            If (Trim$(Trim$(grdTimes.TextMatrix(llRowNo, MGAIRVEHICLEINDEX))) <> "") Or (Trim$(Trim$(grdTimes.TextMatrix(llRowNo, MGMISSEDDATEINDEX))) <> "") Then
                slStr = Trim$(grdTimes.TextMatrix(llRowNo, MGMISSEDDATEINDEX))
                If slStr = "" Then
                    ilError = True
                    grdTimes.Row = llRowNo
                    grdTimes.Col = MGMISSEDDATEINDEX
                    grdTimes.CellForeColor = vbRed
                Else
                    If Not gValidDate(slStr) Then
                        ilError = True
                        grdTimes.Row = llRowNo
                        grdTimes.Col = MGMISSEDDATEINDEX
                        grdTimes.CellForeColor = vbRed
                    End If
                End If
                If Trim$(grdTimes.TextMatrix(llRowNo, MGAIRVEHICLEINDEX)) = "" Then
                    grdTimes.TextMatrix(llRowNo, MGAIRVEHICLEINDEX) = "Missing"
                    ilError = True
                    grdTimes.Row = llRowNo
                    grdTimes.Col = MGAIRVEHICLEINDEX
                    grdTimes.CellForeColor = vbRed
                End If
            End If
        End If
    End If
    If ilError Then
        mGridFieldsOk = False
    Else
        mGridFieldsOk = True
    End If
End Function


Private Function mColOk() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilPos                         ilValue                   *
'*                                                                                        *
'******************************************************************************************


    If Not imUpdateAllowed Then
        mColOk = False
    ElseIf igRepSpotBilled Then
        mColOk = False
    Else
        mColOk = True
        If grdTimes.CellBackColor = LIGHTYELLOW Then
            mColOk = False
            Exit Function
        End If
        If grdTimes.ForeColor = DARKGREEN Then
            mColOk = False
            Exit Function
        End If
    End If
End Function


Private Sub mInitSdf(llRow As Long)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCode                                                                                *
'******************************************************************************************

    Dim slStr As String

    tmSdf.lCode = 0
    If Val(grdTimes.TextMatrix(llRow, SDFCODEINDEX)) > 0 Then
        tmSdf.lCode = Val(grdTimes.TextMatrix(llRow, SDFCODEINDEX))
    End If
    tmSdf.iVefCode = tgClfRep(cbcLine.ItemData(cbcLine.ListIndex)).ClfRec.iVefCode
    tmSdf.lChfCode = tgChfRep.lCode
    tmSdf.iLineNo = tgClfRep(cbcLine.ItemData(cbcLine.ListIndex)).ClfRec.iLine
    tmSdf.iAdfCode = tgChfRep.iAdfCode
    slStr = grdTimes.TextMatrix(llRow, DATEINDEX)
    gPackDate slStr, tmSdf.iDate(0), tmSdf.iDate(1)
    slStr = grdTimes.TextMatrix(llRow, TIMEINDEX)
    gPackTime slStr, tmSdf.iTime(0), tmSdf.iTime(1)
    slStr = Trim$(grdTimes.TextMatrix(llRow, STATUSINDEX))
    tmSdf.sPriceType = "L"
    tmSdf.sSpotType = "A"
    If StrComp(slStr, "Bonus", vbTextCompare) = 0 Then
        If smBonusOnInv <> "N" Then
            tmSdf.sPriceType = "B"  'tmClf.sPriceType 'T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
            tmSdf.sSpotType = "X"   '"T"=Remnant; Q=per Inquire; S=PSA; M=Promo; X=Extra Spot
        Else
            tmSdf.sPriceType = "N"  'tmClf.sPriceType 'T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
            tmSdf.sSpotType = "X"   '"T"=Remnant; Q=per Inquire; S=PSA; M=Promo; X=Extra Spot
        End If
    End If
    tmSdf.iMnfMissed = 0
    If Trim$(grdTimes.TextMatrix(llRow, MGAIRVEHICLEINDEX)) <> "" Then
        slStr = Trim$(grdTimes.TextMatrix(llRow, MGAIRVEHICLEINDEX))
        gFindMatch slStr, 1, lbcVehicle
        If gLastFound(lbcVehicle) >= 1 Then
            tmSdf.sSchStatus = "O"
            tmSdf.iVefCode = lbcVehicle.ItemData(gLastFound(lbcVehicle))
        Else
            tmSdf.sSchStatus = "S"
        End If
    Else
        tmSdf.sSchStatus = "S"
    End If
    tmSdf.sTracer = ""
    tmSdf.sAffChg = ""
    tmSdf.sPtType = "0"
    tmSdf.lCopyCode = mAddCopy(llRow)
    If tmSdf.lCopyCode > 0 Then
        tmSdf.sPtType = "1"
    End If
    tmSdf.iLen = tgClfRep(cbcLine.ItemData(cbcLine.ListIndex)).ClfRec.iLen
    tmSdf.sBill = "N"
    tmSdf.iRotNo = 0
    tmSdf.lSmfCode = 0
    tmSdf.iUrfCode = tgUrf(0).iCode
    tmSdf.lFsfCode = 0
    tmSdf.iGameNo = 0
    tmSdf.sXCrossMidnight = "N"
    tmSdf.sWasMG = "N"
    tmSdf.sFromWorkArea = "N"
    tmSdf.sUnused = ""
End Sub






Private Sub pbcTme_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeInv.Visible = True
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

Private Sub pbcTme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
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
                            End Select
                        Case 5
                            Select Case ilColNo
                                Case 1
                                    slKey = ":"
                                Case 2
                                    slKey = "AM"
                                Case 3
                                    slKey = "PM"
                            End Select
                    End Select
                    imBypassFocus = True    'Don't change select text
                    edcDropDown.SetFocus
                    'SendKeys slKey
                    gSendKeys edcDropDown, slKey
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

Private Sub tmcStart_Timer()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVef                         slNameCode                    slCode                    *
'*  ilRet                                                                                 *
'******************************************************************************************


    tmcStart.Enabled = False
    If imTerminate Then
        mTerminate
    End If
    'Check if allowed to defined input
    'One vehicle only
    imUpdateAllowed = True
    If (igWinStatus(POSTLOGSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    End If
    If Not imUpdateAllowed Then
        cmcUpdate.Enabled = False
        grdTimes.Enabled = False
    End If
End Sub

Private Function mReadRec(llChfCode As Long, ilLineNo As Integer) As Integer
    Dim ilRet As Integer
    Dim llDate As Long
    Dim ilIncludeSpot As Integer
    Dim ilUpper As Integer

    ReDim tmRepSdf(0 To 0) As SDF
    ReDim lmDelSdfCode(0 To 0) As Long
    tmSdfSrchKey5.lCode = llChfCode
    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey5, INDEXKEY5, BTRV_LOCK_NONE, SETFORWRITE)  'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.lChfCode = llChfCode)
        If (tmSdf.iLineNo = ilLineNo) Then
            If (tmSdf.sSchStatus <> "G") And (tmSdf.sSchStatus <> "O") Then
                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
            Else
                tmSmfSrchKey2.lCode = tmSdf.lCode
                ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llDate
                Else
                    llDate = 0
                End If
            End If
            ilIncludeSpot = True
            If tgChfRep.sBillCycle = "C" Then
                If (llDate > lgRepSpotEndCal) Or (llDate < lgRepSpotStartCal) Then
                    ilIncludeSpot = False
                End If
            Else
                If (llDate > lgRepSpotEndStd) Or (llDate < lgRepSpotStartStd) Then
                    ilIncludeSpot = False
                End If
            End If
            If ilIncludeSpot Then
                tmRepSdf(UBound(tmRepSdf)) = tmSdf
                ReDim Preserve tmRepSdf(0 To UBound(tmRepSdf) + 1) As SDF
            End If
        End If
        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Loop
    ReDim tmExistSbf(0 To 0) As SBF
    ilUpper = 0
    tmSbfSrchKey0.lChfCode = llChfCode
    If tgChfRep.sBillCycle = "C" Then
        gPackDateLong lgRepSpotEndCal, tmSbfSrchKey0.iDate(0), tmSbfSrchKey0.iDate(1)
    Else
        gPackDateLong lgRepSpotEndStd, tmSbfSrchKey0.iDate(0), tmSbfSrchKey0.iDate(1)
    End If
    tmSbfSrchKey0.sTranType = "T"
    ilRet = btrGetEqual(hmSbf, tmExistSbf(ilUpper), imSbfRecLen, tmSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)  'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmExistSbf(ilUpper).lChfCode = llChfCode)
        gUnpackDateLong tmExistSbf(ilUpper).iDate(0), tmExistSbf(ilUpper).iDate(1), llDate
        If tgChfRep.sBillCycle = "C" Then
            If lgRepSpotEndCal <> llDate Then
                Exit Do
            End If
        Else
            If lgRepSpotEndStd <> llDate Then
                Exit Do
            End If
        End If
        If tmExistSbf(ilUpper).sTranType = "T" Then
            If tmExistSbf(ilUpper).iLineNo = ilLineNo Then
                ReDim Preserve tmExistSbf(0 To ilUpper + 1) As SBF
                ilUpper = ilUpper + 1
            End If
        End If
        ilRet = btrGetNext(hmSbf, tmExistSbf(ilUpper), imSbfRecLen, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Loop
    mReadRec = True
End Function

Private Sub mClearGrid()
    Dim llRow As Long
    Dim ilCol As Integer

    grdTimes.Redraw = False
    grdTimes.Rows = imInitNoRows
    For llRow = grdTimes.FixedRows To grdTimes.Rows - 1 Step 1
        grdTimes.rowHeight(llRow) = fgFlexGridRowH
        For ilCol = 0 To grdTimes.Cols - 1 Step 1
            If (smAllowRepMG = "N") And ((ilCol = MGAIRVEHICLEINDEX) Or (ilCol = MGMISSEDDATEINDEX)) Then
                grdTimes.Col = ilCol
                grdTimes.Row = llRow
                grdTimes.CellBackColor = LIGHTYELLOW
            End If
            If (ilCol = SDFCODEINDEX) Then
                grdTimes.TextMatrix(llRow, ilCol) = -1
            Else
                grdTimes.TextMatrix(llRow, ilCol) = ""
            End If
        Next ilCol
    Next llRow
    grdTimes.Redraw = True
End Sub

Private Sub mInitSbf(llPrice As Long, slDate As String, slPaidOrBonus As String)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilFound                                                                               *
'******************************************************************************************

    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim llDate As Long

    llDate = gDateValue(slDate)
    For ilLoop = 0 To UBound(tmSbf) - 1 Step 1
        If tmSbf(ilLoop).iLineNo = tgClfRep(cbcLine.ItemData(cbcLine.ListIndex)).ClfRec.iLine Then
            If tmSbf(ilLoop).lSpotPrice = llPrice Then
                If tgChfRep.sBillCycle = "C" Then
                    If slPaidOrBonus = "B" Then
                        tmSbf(ilLoop).iBonusNoSpots = tmSbf(ilLoop).iBonusNoSpots + 1
                    Else
                        tmSbf(ilLoop).iAirNoSpots = tmSbf(ilLoop).iAirNoSpots + 1
                    End If
                Else
                    If llDate > lgRepSpotEndStd Then
                        If slPaidOrBonus = "B" Then
                            tmSbf(ilLoop).iCalCarryBonus = tmSbf(ilLoop).iCalCarryBonus + 1
                        Else
                            tmSbf(ilLoop).iCalCarryOver = tmSbf(ilLoop).iCalCarryOver + 1
                        End If
                    Else
                        If slPaidOrBonus = "B" Then
                            tmSbf(ilLoop).iBonusNoSpots = tmSbf(ilLoop).iBonusNoSpots + 1
                        Else
                            tmSbf(ilLoop).iAirNoSpots = tmSbf(ilLoop).iAirNoSpots + 1
                        End If
                    End If
                End If
                Exit Sub
            End If
        End If
    Next ilLoop
    ilUpper = UBound(tmSbf)
    tmSbf(ilUpper).lCode = 0
    tmSbf(ilUpper).lChfCode = tgChfRep.lCode
    If tgChfRep.sBillCycle = "C" Then
        gPackDateLong lgRepSpotEndCal, tmSbf(ilUpper).iDate(0), tmSbf(ilUpper).iDate(1)
    Else
        gPackDateLong lgRepSpotEndStd, tmSbf(ilUpper).iDate(0), tmSbf(ilUpper).iDate(1)
    End If
    tmSbf(ilUpper).sTranType = "T"
    tmSbf(ilUpper).iBillVefCode = tgClfRep(cbcLine.ItemData(cbcLine.ListIndex)).ClfRec.iVefCode
    tmSbf(ilUpper).iMnfItem = 0
    tmSbf(ilUpper).iNoItems = 0
    tmSbf(ilUpper).lGross = 0
    tmSbf(ilUpper).sUnused1 = ""
    gPackDate "", tmSbf(ilUpper).iPrintInvDate(0), tmSbf(ilUpper).iPrintInvDate(1)
    tmSbf(ilUpper).sDescr = ""
    tmSbf(ilUpper).sAgyComm = "Y"
    tmSbf(ilUpper).sSlsComm = "Y"
    tmSbf(ilUpper).sUnused2 = ""
    tmSbf(ilUpper).sBilled = "N"
    If tgChfRep.iPctTrade = 0 Then
        tmSbf(ilUpper).sCashTrade = "C"
    ElseIf tgChfRep.iPctTrade = 100 Then
        tmSbf(ilUpper).sCashTrade = "T"
    Else
        tmSbf(ilUpper).sCashTrade = "B"
    End If
    tmSbf(ilUpper).iAirVefCode = tmSbf(ilUpper).iBillVefCode
    If tgChfRep.sBillCycle = "C" Then
        If slPaidOrBonus = "B" Then
            tmSbf(ilUpper).iBonusNoSpots = 1
        Else
            tmSbf(ilUpper).iAirNoSpots = 1
        End If
    Else
        If llDate > lgRepSpotEndStd Then
            If slPaidOrBonus = "B" Then
                tmSbf(ilUpper).iCalCarryBonus = 1
            Else
                tmSbf(ilUpper).iCalCarryOver = 1
            End If
        Else
            If slPaidOrBonus = "B" Then
                tmSbf(ilUpper).iBonusNoSpots = 1
            Else
                tmSbf(ilUpper).iAirNoSpots = 1
            End If
        End If
    End If
    tmSbf(ilUpper).iBonusNoSpots = 0
    tmSbf(ilUpper).iTrfCode = 0
    tmSbf(ilUpper).sUnused3 = ""
    gPackDate "", tmSbf(ilUpper).iImportDate(0), tmSbf(ilUpper).iImportDate(1)
    gPackDate "", tmSbf(ilUpper).iExportDate(0), tmSbf(ilUpper).iExportDate(1)
    tmSbf(ilUpper).lRefInvNo = 0
    tmSbf(ilUpper).iCombineID = 0
    tmSbf(ilUpper).iUrfCode = tgUrf(0).iCode
    tmSbf(ilUpper).lOGross = 0
    tmSbf(ilUpper).iCommPct = 0
    gPackDate "", tmSbf(ilUpper).iRecDate(0), tmSbf(ilUpper).iRecDate(1)
    gPackDate Format$(gNow(), "m/d/yy"), tmSbf(ilUpper).iPostDate(0), tmSbf(ilUpper).iPostDate(1)
    tmSbf(ilUpper).sPostStatus = "P"
    tmSbf(ilUpper).iCalCarryOver = 0
    tmSbf(ilUpper).lSpotPrice = llPrice
    tmSbf(ilUpper).iMissCarryOver = 0
    tmSbf(ilUpper).sInserted = "N"
    tmSbf(ilUpper).sBarterPaid = "N"
    tmSbf(ilUpper).lAcquisitionCost = 0
    tmSbf(ilUpper).iSpotLen = tgClfRep(cbcLine.ItemData(cbcLine.ListIndex)).ClfRec.iLen
    tmSbf(ilUpper).iIhfCode = 0
    tmSbf(ilUpper).iLineNo = tgClfRep(cbcLine.ItemData(cbcLine.ListIndex)).ClfRec.iLine
    ReDim Preserve tmSbf(0 To ilUpper + 1) As SBF
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyPop                        *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection Name    *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mCopyPop()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mCopyPopErr                                                                           *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilAdvtCode As Integer
    Dim ilISCIProd As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String

    ilISCIProd = 4
    ilAdvtCode = tgChfRep.iAdfCode
    ilRet = gPopCopyForAdvtBox(PostRepTimes, ilAdvtCode, ilISCIProd, 1, lbcCopyNm, tmCopyNmCode(), smCopyNmCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        lbcCopyNm.Clear
        For ilLoop = 0 To UBound(tmCopyNmCode) - 1 Step 1
            slNameCode = tmCopyNmCode(ilLoop).sKey
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            lbcCopyNm.AddItem Trim$(slName)
            lbcCopyNm.ItemData(lbcCopyNm.NewIndex) = slCode
        Next ilLoop
        lbcCopyNm.AddItem "[None]", 0   ' Add this at the top of the list
    End If
    Exit Sub
mCopyPopErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
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


Private Sub mGetTotals()
    Dim slOrdered As String
    Dim llRow As Long
    Dim ilPosted As Integer
    Dim slStr As String

    If cbcLine.ListIndex < 0 Then
        lacSpotStatus.Visible = False
        Exit Sub
    End If
    slOrdered = "Ordered: " & Trim$(str$(imOrderedSpots(cbcLine.ListIndex)))
    ilPosted = 0
    For llRow = grdTimes.FixedRows To grdTimes.Rows - 1 Step 1
        slStr = Trim$(grdTimes.TextMatrix(llRow, STATUSINDEX))
        If slStr <> "" Then
            slStr = Trim$(grdTimes.TextMatrix(llRow, DATEINDEX))
            If slStr <> "" Then
                ilPosted = ilPosted + 1
            End If
        End If
    Next llRow
    lacSpotStatus.Caption = slOrdered & " Posted: " & Trim$(str$(ilPosted))
    If imOrderedSpots(cbcLine.ListIndex) = ilPosted Then
        lacSpotStatus.BackColor = vbGreen
    Else
        lacSpotStatus.BackColor = vbRed
    End If
    lacSpotStatus.Visible = True
End Sub

Private Function mAddCopy(llRowNo As Long) As Long
    Dim slISCI As String
    Dim ilRet As Integer
    Dim slDate As String

    slISCI = Trim$(grdTimes.TextMatrix(llRowNo, ISCIINDEX))
    gFindMatch slISCI, 0, lbcCopyNm
    If gLastFound(lbcCopyNm) >= 0 Then
        If Val(lbcCopyNm.ItemData(gLastFound(lbcCopyNm))) > 0 Then
            mAddCopy = Val(lbcCopyNm.ItemData(gLastFound(lbcCopyNm)))
            Exit Function
        End If
    Else
        mAddCopy = 0
        Exit Function
    End If
    'Add Product/ISCI
    tmCpf.lCode = 0
    tmCpf.sName = ""
    tmCpf.sISCI = slISCI
    tmCpf.sCreative = ""
    tmCpf.iRotEndDate(0) = 0
    tmCpf.iRotEndDate(1) = 0
    tmCpf.lSifCode = 0
    ilRet = btrInsert(hmCpf, tmCpf, imCpfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        mAddCopy = 0
        Exit Function
    End If
    'Add copy inventory
    tmCif.lCode = 0  'Autoincrement
    tmCif.iMcfCode = 0
    tmCif.sName = ""
    tmCif.sCut = ""
    tmCif.sReel = ""
    tmCif.iLen = tgClfRep(cbcLine.ItemData(cbcLine.ListIndex)).ClfRec.iLen
    tmCif.iEtfCode = 0
    tmCif.iEnfCode = 0
    tmCif.iAdfCode = tgChfRep.iAdfCode
    tmCif.lcpfCode = tmCpf.lCode
    tmCif.iMnfComp(0) = 0
    tmCif.iMnfComp(1) = 0
    tmCif.iMnfAnn = 0
    tmCif.sHouse = "Y"
    tmCif.sCleared = "Y"
    tmCif.lCsfCode = 0
    tmCif.iNoTimesAir = 0
    tmCif.sCartDisp = "N"
    tmCif.sTapeDisp = "N"
    tmCif.sPurged = "A"
    tmCif.iPurgeDate(0) = 0
    tmCif.iPurgeDate(1) = 0
    slDate = Format$(gNow(), "m/d/yy")
    gPackDate slDate, tmCif.iDateEntrd(0), tmCif.iDateEntrd(1)
    tmCif.iUsedDate(0) = 0
    tmCif.iUsedDate(1) = 0
    tmCif.iRotStartDate(0) = 0
    tmCif.iRotStartDate(1) = 0
    tmCif.iRotEndDate(0) = 0
    tmCif.iRotEndDate(1) = 0
    tmCif.iUrfCode = tgUrf(0).iCode
    tmCif.sPrint = "N"
    tmCif.iLangMnfCode = 0
    tmCif.sUnused = ""
    ilRet = btrInsert(hmCif, tmCif, imCifRecLen, INDEXKEY0)
    If ilRet = BTRV_ERR_NONE Then
        mAddCopy = tmCif.lCode
    Else
        mAddCopy = 0
    End If
End Function

Private Function mVehPop() As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilLoop As Integer

    ilRet = gPopUserVehicleBox(PostRepTimes, VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + ACTIVEVEH + DORMANTVEH, lbcVehicle, tmVehicle(), smVehicleTag)

    lbcVehicle.Clear
    For ilLoop = 0 To UBound(tmVehicle) - 1 Step 1
        slNameCode = tmVehicle(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 3, "|", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        lbcVehicle.AddItem Trim$(slName)
        lbcVehicle.ItemData(lbcVehicle.NewIndex) = slCode
    Next ilLoop
    Erase tmVehicle
    lbcVehicle.AddItem "", 0
    lbcVehicle.ItemData(lbcVehicle.NewIndex) = 0
    mVehPop = True
End Function
