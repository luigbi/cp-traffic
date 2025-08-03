VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form AlertVw 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4125
   ClientLeft      =   885
   ClientTop       =   2415
   ClientWidth     =   9315
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
   ScaleHeight     =   4125
   ScaleWidth      =   9315
   Begin VB.CommandButton cmcSpots 
      Appearance      =   0  'Flat
      Caption         =   "S&pots"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7770
      TabIndex        =   14
      Top             =   3720
      Width           =   945
   End
   Begin VB.CommandButton cmcSchedule 
      Appearance      =   0  'Flat
      Caption         =   "&Schedule"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5970
      TabIndex        =   13
      Top             =   3720
      Width           =   1455
   End
   Begin VB.ListBox lbcMkt 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "AlertVw.frx":0000
      Left            =   7635
      List            =   "AlertVw.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   12
      Top             =   705
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.CommandButton cmcViewCntr 
      Appearance      =   0  'Flat
      Caption         =   "&View Proposal"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4140
      TabIndex        =   11
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmcChgCntr 
      Appearance      =   0  'Flat
      Caption         =   "&Change Proposal"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Frame frcView 
      Caption         =   "View"
      Height          =   570
      Left            =   1110
      TabIndex        =   6
      Top             =   -270
      Visible         =   0   'False
      Width           =   7590
      Begin VB.OptionButton rbcView 
         Caption         =   "Pool Alert"
         Height          =   210
         Index           =   4
         Left            =   4590
         TabIndex        =   19
         Top             =   255
         Width           =   1230
      End
      Begin VB.OptionButton rbcView 
         Caption         =   "Rep-Net Messages"
         Height          =   210
         Index           =   3
         Left            =   2565
         TabIndex        =   15
         Top             =   255
         Width           =   1890
      End
      Begin VB.OptionButton rbcView 
         Caption         =   "Affiliate Alert"
         Height          =   210
         Index           =   2
         Left            =   5910
         TabIndex        =   9
         Top             =   255
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.OptionButton rbcView 
         Caption         =   "Logs"
         Height          =   210
         Index           =   1
         Left            =   1545
         TabIndex        =   8
         Top             =   255
         Width           =   1005
      End
      Begin VB.OptionButton rbcView 
         Caption         =   "Contracts"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   255
         Width           =   1335
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
      Left            =   630
      TabIndex        =   0
      Top             =   3720
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCntr 
      Height          =   2415
      Left            =   480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1395
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      Rows            =   10
      Cols            =   10
      FixedCols       =   0
      BackColorSel    =   12632256
      ForeColorSel    =   16711680
      BackColorUnpopulated=   16777215
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
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
      _Band(0).Cols   =   10
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLog 
      Height          =   2415
      Left            =   690
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1230
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   16777215
      Rows            =   10
      Cols            =   7
      FixedCols       =   0
      BackColorSel    =   12632256
      BackColorUnpopulated=   16777215
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
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
      _Band(0).Cols   =   7
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAffExport 
      Height          =   2415
      Left            =   915
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1050
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   4260
      _Version        =   393216
      Rows            =   10
      Cols            =   7
      FixedCols       =   0
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
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
      _Band(0).Cols   =   7
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdRNMsg 
      Height          =   2415
      Left            =   1125
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   930
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      Rows            =   10
      Cols            =   11
      FixedCols       =   0
      BackColorSel    =   12632256
      ForeColorSel    =   16711680
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      GridColorUnpopulated=   16777215
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      SelectionMode   =   1
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
      _Band(0).Cols   =   11
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPrgmmatic 
      Height          =   2415
      Left            =   1500
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   765
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      Rows            =   10
      Cols            =   12
      FixedCols       =   0
      BackColorSel    =   12632256
      ForeColorSel    =   16711680
      BackColorUnpopulated=   16777215
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
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
      _Band(0).Cols   =   12
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPool 
      Height          =   2415
      Left            =   1680
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   600
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      Rows            =   10
      Cols            =   6
      FixedCols       =   0
      BackColorSel    =   12632256
      ForeColorSel    =   16711680
      BackColorUnpopulated=   16777215
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
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
      _Band(0).Cols   =   6
   End
   Begin ComctlLib.TabStrip tbcAlert 
      Height          =   3270
      Left            =   135
      TabIndex        =   17
      Top             =   315
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   5768
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   6
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Contracts"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Logs"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Programmatic Buys"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Affiliate Alert"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Rep-Net Messages"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pool Alert"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lbcScreen 
      Caption         =   "View Alerts"
      Height          =   180
      Left            =   120
      TabIndex        =   5
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
Attribute VB_Name = "AlertVw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of AlertVw.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: AlertVw.Frm
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
Dim tmAuf As AUF        'Rvf record image
Dim tmAufSrchKey0 As LONGKEY0
Dim tmAufSrchKey1 As AUFKEY1    'Rvf key record image
Dim imAufRecLen As Integer        'RvF record length
'Contract line
Dim hmCHF As Integer        'Contract line file handle
Dim tmChf As CHF            'CHF record image
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim imCHFRecLen As Integer     'CHF record length
'Sales Office
Dim tmSof As SOF            'SOF record image
Dim tmSofSrchKey As INTKEY0 'SOF key record image
Dim imSofRecLen As Integer  'SOF record length
Dim hmSof As Integer        'Selling Office file handle
'Comment
Dim hmCef As Integer        'Comment file handle
Dim tmCef As CEF            'CEF record image
Dim tmCefSrchKey As LONGKEY0 'CEF key record image
Dim imCefRecLen As Integer     'CEF record length
'Vehicle Option
Dim hmVLF As Integer        'Vehicle Option file handle
Dim hmVsf As Integer

Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim tmVehGp3Code() As SORTCODE
Dim smVehGp3CodeTag As String
Dim imScheduleExist As Integer
Dim imCompleteExist As Integer
Dim imPrgmmaticScheduleExist As Integer
Dim imLastSelectRow As Integer
Dim imCtrlKey As Integer
Dim imLastCntrColSorted As Integer
Dim imLastCntrSort As Integer
Dim imLastLogColSorted As Integer
Dim imLastLogSort As Integer
Dim imLastPrgmmaticColSorted As Integer
Dim imLastPrgmmaticSort As Integer
Dim imLastAffColSorted As Integer
Dim imLastAffSort As Integer
Dim imLastRNMsgColSorted As Integer
Dim imLastRNMsgSort As Integer
Dim lmCntrRowSelected As Long
Dim lmPrgmmaticRowSelected As Long
Dim lmRNMsgRowSelected As Long
Dim lmLogRowSelected As Long
Dim imTabSelected As Integer
Dim lmPoolRowSelected As Long
Dim imLastPoolColSorted As Integer
Dim imLastPoolSort As Integer

'Contract
Const C_CREATEDATEINDEX = 0
Const C_CREATETIMEINDEX = 1
Const C_CNTRNOINDEX = 2
Const C_ADVTNAMEINDEX = 3
Const C_PRODUCTINDEX = 4
Const C_CNTRSTARTDATEINDEX = 5
Const C_SALESOFFICEINDEX = 6
Const C_REASONINDEX = 7
Const C_AUFCODEINDEX = 8
Const C_SORTINDEX = 9

'Log
Const L_CREATEDATEINDEX = 0
Const L_CREATETIMEINDEX = 1
Const L_VEHICLEINDEX = 2
Const L_LOGDATEINDEX = 3
Const L_REASONINDEX = 4
Const L_AUFCODEINDEX = 5
Const L_SORTINDEX = 6

'Affiliate Export
Const A_CREATEDATEINDEX = 0
Const A_CREATETIMEINDEX = 1
Const A_VEHICLEINDEX = 2
Const A_LOGDATEINDEX = 3
Const A_REASONINDEX = 4
Const A_AUFCODEINDEX = 5
Const A_SORTINDEX = 6

'Rep-Net
Const RN_CREATEDATEINDEX = 0
Const RN_CREATETIMEINDEX = 1
Const RN_TYPEMSGINDEX = 2
Const RN_STATUSINDEX = 3
Const RN_FROMIDINDEX = 4
Const RN_CNTRNOINDEX = 5
Const RN_ADVTNAMEINDEX = 6
Const RN_VEHICLEINDEX = 7
Const RN_MESSAGEINDEX = 8
Const RN_AUFCODEINDEX = 9
Const RN_SORTINDEX = 10

'Programmatic Buys
Const P_CREATEDATEINDEX = 0
Const P_CREATETIMEINDEX = 1
Const P_AGENCYINDEX = 2
Const P_CNTRNOINDEX = 3
Const P_ADVTNAMEINDEX = 4
Const P_PRODUCTINDEX = 5
Const P_CNTRSTARTDATEINDEX = 6
Const P_SALESOFFICEINDEX = 7
Const P_REASONINDEX = 8
Const P_AUFCODEINDEX = 9
Const P_CNTRSTATUSINDEX = 10
Const P_SORTINDEX = 11

Const UP_CREATEDATEINDEX = 0
Const UP_CREATETIMEINDEX = 1
Const UP_FILENAMEINDEX = 2
Const UP_FILEDATEINDEX = 3
Const UP_AUFCODEINDEX = 4
Const UP_SORTINDEX = 5


Private Sub cmcChgCntr_Click()
    Dim ilRet As Integer
    Dim slDate As String
    Dim slTime As String
    Dim slStr As String

    'If rbcView(0).Value Then
    If imTabSelected = 0 Then
        If igJobShowing(CONTRACTSJOB) Then
            '4/5/10:  Enable is always False as Image is not active
            'If (Contract!cmcUpdate.Enabled = False) Then
            If igOKToCallCntr Then
                igTerminateAndUnload = True
                Unload Contract
                DoEvents
                Sleep 50
            Else
                ilRet = MsgBox("Unable to Change Contract until Current Altered Contract has been Saved", vbInformation + vbOKOnly, "Information")
                Exit Sub
            End If
        End If
        igAlertCntrStatus = 0
        If lmCntrRowSelected >= grdCntr.FixedRows Then
            If Trim$(grdCntr.TextMatrix(lmCntrRowSelected, C_CREATEDATEINDEX)) = "" Then
                Exit Sub
            End If
            lgAlertCntrNo = grdCntr.TextMatrix(lmCntrRowSelected, C_CNTRNOINDEX)
            If Trim$(grdCntr.TextMatrix(lmCntrRowSelected, C_REASONINDEX)) = "Complete" Then
                If tgSpf.sGUsePropSys = "Y" Then
                    If igWinStatus(PROPOSALSJOB) = 0 Then
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
                igAlertCntrStatus = 1
            '8/15/18
            ElseIf Trim$(grdCntr.TextMatrix(lmCntrRowSelected, C_REASONINDEX)) = "Unapproved" Then
                If tgSpf.sGUsePropSys = "Y" Then
                    If igWinStatus(PROPOSALSJOB) = 0 Then
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
                igAlertCntrStatus = 1
            ElseIf Trim$(grdCntr.TextMatrix(lmCntrRowSelected, C_REASONINDEX)) = "Not Scheduled" Then
                If igWinStatus(CONTRACTSJOB) = 0 Then
                    Exit Sub
                End If
                igAlertCntrStatus = 4
            Else
                If tgSpf.sGUsePropSys = "Y" Then
                    If igWinStatus(PROPOSALSJOB) = 0 Then
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
                igAlertCntrStatus = 2
            End If
        End If
        mTerminate
    ElseIf imTabSelected = 2 Then
        If igJobShowing(CONTRACTSJOB) Then
            '4/5/10:  Enable is always False as Image is not active
            'If (Contract!cmcUpdate.Enabled = False) Then
            If igOKToCallCntr Then
                igTerminateAndUnload = True
                Unload Contract
                DoEvents
                Sleep 50
            Else
                ilRet = MsgBox("Unable to Change Contract until Current Altered Contract has been Saved", vbInformation + vbOKOnly, "Information")
                Exit Sub
            End If
        End If
        '0= not from alert screen, 1= Change mode: Complete; 2=Change mode: Rev Complete; 3=View mode
        igAlertCntrStatus = 0
        If lmPrgmmaticRowSelected >= grdPrgmmatic.FixedRows Then
            If Trim$(grdPrgmmatic.TextMatrix(lmPrgmmaticRowSelected, P_CREATEDATEINDEX)) = "" Then
                Exit Sub
            End If
            lgAlertCntrNo = grdPrgmmatic.TextMatrix(lmPrgmmaticRowSelected, P_CNTRNOINDEX)
            slStr = Trim$(grdPrgmmatic.TextMatrix(lmPrgmmaticRowSelected, P_CNTRSTATUSINDEX))
            If (slStr = "W") Or (slStr = "C") Then
                If tgSpf.sGUsePropSys = "Y" Then
                    If igWinStatus(PROPOSALSJOB) = 0 Then
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
                igAlertCntrStatus = 1
            ElseIf (slStr = "N") Or (slStr = "O") Then
                If igWinStatus(CONTRACTSJOB) = 0 Then
                    Exit Sub
                End If
                igAlertCntrStatus = 4
            Else
                If tgSpf.sGUsePropSys = "Y" Then
                    If igWinStatus(PROPOSALSJOB) = 0 Then
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
                igAlertCntrStatus = 2
            End If
        End If
        mTerminate    'ElseIf rbcView(3).Value Then
    ElseIf imTabSelected = 4 Then
        If lmRNMsgRowSelected >= grdRNMsg.FixedRows Then
            tmAufSrchKey0.lCode = grdRNMsg.TextMatrix(lmRNMsgRowSelected, RN_AUFCODEINDEX)
            ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                tmAuf.sStatus = "C"
                tmAuf.sClearMethod = "A"
                slDate = Format$(gNow(), "m/d/yy")
                slTime = Format$(gNow(), "h:mm:ssAM/PM")
                gPackDate slDate, tmAuf.iClearDate(0), tmAuf.iClearDate(1)
                gPackTime slTime, tmAuf.iClearTime(0), tmAuf.iClearTime(1)
                tmAuf.iClearUrfCode = tgUrf(0).iCode
                ilRet = btrUpdate(hgAuf, tmAuf, imAufRecLen)
                grdRNMsg.Redraw = False
                grdRNMsg.RemoveItem lmRNMsgRowSelected
                grdRNMsg.Row = grdRNMsg.Rows - 1
                grdRNMsg.AddItem ""
                grdRNMsg.RowHeight(grdRNMsg.Rows - 1) = fgBoxGridH + 15
                grdRNMsg.Row = 0
                grdRNMsg.Col = RN_AUFCODEINDEX
                lmRNMsgRowSelected = -1
                grdRNMsg.Redraw = True
            End If
        End If
    End If
    Exit Sub
cmcChgCntrErr: 'VBC NR
    ilRet = 1
    Resume Next
End Sub

Private Sub cmcDone_Click()
    igAlertCntrStatus = 0
    igAlertSpotStatus = 0
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcSchedule_Click()
    mSchedule
End Sub

Private Sub cmcSpots_Click()
    mSpots
End Sub

Private Sub cmcViewCntr_Click()

    Dim ilRet As Integer
    Dim slStr As String

    If igJobShowing(CONTRACTSJOB) Then
        '4/5/10:  Enable is always False as Image is not active
        'If (Contract!cmcUpdate.Enabled = False) Then
        If igOKToCallCntr Then
            igTerminateAndUnload = True
            Unload Contract
            DoEvents
            Sleep 50
        Else
            ilRet = MsgBox("Unable to view Contract until Current Altered Contract has been Saved", vbInformation + vbOKOnly, "Information")
            Exit Sub
        End If
    End If
    igAlertCntrStatus = 0
    If imTabSelected = 0 Then
        If lmCntrRowSelected >= grdCntr.FixedRows Then
            If Trim$(grdCntr.TextMatrix(lmCntrRowSelected, C_CREATEDATEINDEX)) = "" Then
                Exit Sub
            End If
            lgAlertCntrNo = grdCntr.TextMatrix(lmCntrRowSelected, C_CNTRNOINDEX)
            If Trim$(grdCntr.TextMatrix(lmCntrRowSelected, C_REASONINDEX)) = "Complete" Then
                If tgSpf.sGUsePropSys = "Y" Then
                    If igWinStatus(PROPOSALSJOB) = 0 Then
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
                igAlertCntrStatus = 3
            ElseIf Trim$(grdCntr.TextMatrix(lmCntrRowSelected, C_REASONINDEX)) = "Unapproved" Then
                If tgSpf.sGUsePropSys = "Y" Then
                    If igWinStatus(PROPOSALSJOB) = 0 Then
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
                igAlertCntrStatus = 3
            Else
                If igWinStatus(CONTRACTSJOB) = 0 Then
                    Exit Sub
                End If
                igAlertCntrStatus = 5
            End If
        End If
    ElseIf imTabSelected = 2 Then
        If lmPrgmmaticRowSelected >= grdPrgmmatic.FixedRows Then
            If Trim$(grdPrgmmatic.TextMatrix(lmPrgmmaticRowSelected, P_CREATEDATEINDEX)) = "" Then
                Exit Sub
            End If
            lgAlertCntrNo = grdPrgmmatic.TextMatrix(lmPrgmmaticRowSelected, P_CNTRNOINDEX)
            slStr = Trim$(grdPrgmmatic.TextMatrix(lmPrgmmaticRowSelected, P_REASONINDEX))
            If (slStr = "W") Or (slStr = "C") Then
                If tgSpf.sGUsePropSys = "Y" Then
                    If igWinStatus(PROPOSALSJOB) = 0 Then
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
                igAlertCntrStatus = 3
            Else
                If igWinStatus(CONTRACTSJOB) = 0 Then
                    Exit Sub
                End If
                igAlertCntrStatus = 5
            End If
        End If
    End If
    mTerminate
    Exit Sub
cmcViewCntrErr: 'VBC NR
    ilRet = 1
    Resume Next
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
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
    AlertVw.Refresh
    Me.KeyPreview = True
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_Initialize()
    'Me.Width = (CLng(90) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    'Me.Height = (CLng(90) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    Me.Width = (CLng(75) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    Me.Height = (CLng(90) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    gCenterStdAlone AlertVw
    DoEvents
    mSetControls
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
    If imTerminate Then
        mTerminate
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tmAufView
    Erase tmVehGp3Code
    
    btrExtClear hmVsf   'Clear any previous extend operation
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    btrExtClear hmVLF   'Clear any previous extend operation
    ilRet = btrClose(hmVLF)
    btrDestroy hmVLF
    btrExtClear hmCef   'Clear any previous extend operation
    ilRet = btrClose(hmCef)
    btrDestroy hmCef
    btrExtClear hmSof   'Clear any previous extend operation
    ilRet = btrClose(hmSof)
    btrExtClear hmCHF   'Clear any previous extend operation
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    
    Set AlertVw = Nothing   'Remove data segment

End Sub

Private Sub grdAffExport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < grdAffExport.RowHeight(0) Then
        grdAffExport.Col = grdAffExport.MouseCol
        mAffSortCol grdAffExport.Col
        Exit Sub
    End If
End Sub

Private Sub grdCntr_Click()


    If grdCntr.Row >= grdCntr.FixedRows Then
        If grdCntr.TextMatrix(grdCntr.Row, C_CREATEDATEINDEX) <> "" Then
            If (lmCntrRowSelected = grdCntr.Row) Then
                If imCtrlKey Then
                    lmCntrRowSelected = -1
                    grdCntr.Row = 0
                    grdCntr.Col = C_AUFCODEINDEX
                End If
            Else
                lmCntrRowSelected = grdCntr.Row
            End If
        Else
            lmCntrRowSelected = -1
            grdCntr.Row = 0
            grdCntr.Col = C_AUFCODEINDEX
        End If
    End If
    mSetCommands
End Sub

Private Sub grdCntr_DblClick()
    If cmcSchedule.Enabled Then
        mSchedule
    ElseIf cmcChgCntr.Enabled Then
        cmcChgCntr_Click
    ElseIf cmcViewCntr.Enabled Then
        cmcViewCntr_Click
    End If
End Sub

Private Sub grdCntr_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If
End Sub

Private Sub grdCntr_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
End Sub

Private Sub grdCntr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llAufCode As Long
    Dim llRow As Long

    If Y < grdCntr.RowHeight(0) Then
        Screen.MousePointer = vbHourglass
        llAufCode = -1
        If lmCntrRowSelected >= grdCntr.FixedRows Then
            If Trim$(grdCntr.TextMatrix(lmCntrRowSelected, C_CREATEDATEINDEX)) <> "" Then
                llAufCode = grdCntr.TextMatrix(lmCntrRowSelected, C_AUFCODEINDEX)
            End If
        End If
        grdCntr.Col = grdCntr.MouseCol
        mCntrSortCol grdCntr.Col
        grdCntr.Row = 0
        grdCntr.Col = C_AUFCODEINDEX
        lmCntrRowSelected = -1
        If llAufCode <> -1 Then
            For llRow = grdCntr.FixedRows To grdCntr.Rows - 1 Step 1
                If llAufCode = grdCntr.TextMatrix(llRow, C_AUFCODEINDEX) Then
                    grdCntr.Row = llRow
                    grdCntr.RowSel = llRow
                    grdCntr.Col = C_CREATEDATEINDEX
                    grdCntr.ColSel = C_REASONINDEX
                    lmCntrRowSelected = llRow
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            Next llRow
        End If
        mSetCommands
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

End Sub

Private Sub grdLog_Click()
    Dim slStr As String
    
    If igWinStatus(SPOTSJOB) = 0 Then
        Exit Sub
    End If
    If grdLog.Row >= grdLog.FixedRows Then
        If grdLog.TextMatrix(grdLog.Row, L_CREATEDATEINDEX) <> "" Then
            If (lmLogRowSelected = grdLog.Row) Then
                If imCtrlKey Then
                    mPaintLog lmLogRowSelected, False
                    lmLogRowSelected = -1
                    grdLog.Row = 0
                    grdLog.Col = L_AUFCODEINDEX
                End If
            Else
                slStr = grdLog.TextMatrix(grdLog.Row, L_REASONINDEX)
                If InStr(1, slStr, "Missed", vbTextCompare) = 1 Then
                    mPaintLog lmLogRowSelected, False
                    lmLogRowSelected = grdLog.Row
                    mPaintLog lmLogRowSelected, True
                Else
                    mPaintLog lmLogRowSelected, False
                    lmLogRowSelected = -1
                    grdLog.Row = 0
                    grdLog.Col = L_AUFCODEINDEX
                End If
            End If
        Else
            mPaintLog lmLogRowSelected, False
            lmLogRowSelected = -1
            grdLog.Row = 0
            grdLog.Col = L_AUFCODEINDEX
        End If
    End If
    mSetCommands

End Sub

Private Sub grdLog_DblClick()
    If cmcSpots.Enabled Then
        mSpots
    End If

End Sub

Private Sub grdLog_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If
End Sub

Private Sub grdLog_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
End Sub

Private Sub grdLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llAufCode As Long
    Dim llRow As Long

    If Y < grdLog.RowHeight(0) Then
        Screen.MousePointer = vbHourglass
        llAufCode = -1
        If lmLogRowSelected >= grdLog.FixedRows Then
            If Trim$(grdLog.TextMatrix(lmLogRowSelected, L_CREATEDATEINDEX)) <> "" Then
                llAufCode = grdLog.TextMatrix(lmLogRowSelected, L_AUFCODEINDEX)
            End If
        End If
        grdLog.Col = grdLog.MouseCol
        mLogSortCol grdLog.Col
        grdLog.Row = 0
        grdLog.Col = L_AUFCODEINDEX
        lmLogRowSelected = -1
        If llAufCode <> -1 Then
            For llRow = grdLog.FixedRows To grdLog.Rows - 1 Step 1
                If llAufCode = grdLog.TextMatrix(llRow, L_AUFCODEINDEX) Then
                    grdLog.Row = llRow
                    grdLog.RowSel = llRow
                    grdLog.Col = L_CREATEDATEINDEX
                    grdLog.ColSel = L_REASONINDEX
                    lmLogRowSelected = llRow
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            Next llRow
        End If
        mSetCommands
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


End Sub

Private Sub grdPool_Click()
    lmPoolRowSelected = -1
    grdPool.Row = 0
    grdPool.Col = UP_AUFCODEINDEX
End Sub

Private Sub grdPool_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < grdPool.RowHeight(0) Then
        Screen.MousePointer = vbHourglass
        grdPool.Col = grdPool.MouseCol
        mPoolSortCol grdPool.Col
        grdPool.Row = 0
        grdPool.Col = UP_AUFCODEINDEX
        lmPoolRowSelected = -1
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


End Sub

Private Sub grdPrgmmatic_Click()

    If grdPrgmmatic.Row >= grdPrgmmatic.FixedRows Then
        If grdPrgmmatic.TextMatrix(grdPrgmmatic.Row, P_CREATEDATEINDEX) <> "" Then
            If (lmPrgmmaticRowSelected = grdPrgmmatic.Row) Then
                If imCtrlKey Then
                    lmPrgmmaticRowSelected = -1
                    grdPrgmmatic.Row = 0
                    grdPrgmmatic.Col = P_AUFCODEINDEX
                End If
            Else
                lmPrgmmaticRowSelected = grdPrgmmatic.Row
            End If
        Else
            lmPrgmmaticRowSelected = -1
            grdPrgmmatic.Row = 0
            grdPrgmmatic.Col = P_AUFCODEINDEX
        End If
    End If
    mSetCommands
End Sub

Private Sub grdPrgmmatic_DblClick()
    If cmcSchedule.Enabled Then
        mSchedule
    ElseIf cmcChgCntr.Enabled Then
        cmcChgCntr_Click
    ElseIf cmcViewCntr.Enabled Then
        cmcViewCntr_Click
    End If
End Sub

Private Sub grdPrgmmatic_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If
End Sub

Private Sub grdPrgmmatic_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
End Sub

Private Sub grdPrgmmatic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llAufCode As Long
    Dim llRow As Long

    If Y < grdPrgmmatic.RowHeight(0) Then
        Screen.MousePointer = vbHourglass
        llAufCode = -1
        If lmPrgmmaticRowSelected >= grdPrgmmatic.FixedRows Then
            If Trim$(grdPrgmmatic.TextMatrix(lmPrgmmaticRowSelected, P_CREATEDATEINDEX)) <> "" Then
                llAufCode = grdPrgmmatic.TextMatrix(lmPrgmmaticRowSelected, P_AUFCODEINDEX)
            End If
        End If
        grdPrgmmatic.Col = grdPrgmmatic.MouseCol
        mPrgmmaticSortCol grdPrgmmatic.Col
        grdPrgmmatic.Row = 0
        grdPrgmmatic.Col = P_AUFCODEINDEX
        lmPrgmmaticRowSelected = -1
        If llAufCode <> -1 Then
            For llRow = grdPrgmmatic.FixedRows To grdPrgmmatic.Rows - 1 Step 1
                If llAufCode = grdPrgmmatic.TextMatrix(llRow, P_AUFCODEINDEX) Then
                    grdPrgmmatic.Row = llRow
                    grdPrgmmatic.RowSel = llRow
                    grdPrgmmatic.Col = P_CREATEDATEINDEX
                    grdPrgmmatic.ColSel = P_REASONINDEX
                    lmPrgmmaticRowSelected = llRow
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            Next llRow
        End If
        mSetCommands
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

End Sub

Private Sub grdRNMsg_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llRow                                                                                 *
'******************************************************************************************

    If grdRNMsg.TextMatrix(grdRNMsg.Row, RN_CREATEDATEINDEX) <> "" Then
        If (lmRNMsgRowSelected = grdRNMsg.Row) Then
            If imCtrlKey Then
                lmRNMsgRowSelected = -1
                grdRNMsg.Row = 0
                grdRNMsg.Col = RN_AUFCODEINDEX
            End If
        Else
            lmRNMsgRowSelected = grdRNMsg.Row
        End If
    Else
        lmRNMsgRowSelected = -1
        grdRNMsg.Row = 0
        grdRNMsg.Col = RN_AUFCODEINDEX
    End If
    mSetCommands
End Sub

Private Sub grdRNMsg_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If
End Sub

Private Sub grdRNMsg_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
End Sub

Private Sub grdRNMsg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llAufCode As Long
    Dim llRow As Long

    If Y < grdRNMsg.RowHeight(0) Then
        Screen.MousePointer = vbHourglass
        llAufCode = -1
        If lmRNMsgRowSelected >= grdRNMsg.FixedRows Then
            If Trim$(grdRNMsg.TextMatrix(lmRNMsgRowSelected, RN_CREATEDATEINDEX)) <> "" Then
                llAufCode = grdRNMsg.TextMatrix(lmRNMsgRowSelected, RN_AUFCODEINDEX)
            End If
        End If
        grdRNMsg.Col = grdRNMsg.MouseCol
        mRNMsgSortCol grdRNMsg.Col
        grdRNMsg.Row = 0
        grdRNMsg.Col = RN_AUFCODEINDEX
        lmRNMsgRowSelected = -1
        If llAufCode <> -1 Then
            For llRow = grdRNMsg.FixedRows To grdRNMsg.Rows - 1 Step 1
                If llAufCode = grdRNMsg.TextMatrix(llRow, RN_AUFCODEINDEX) Then
                    grdRNMsg.Row = llRow
                    grdRNMsg.RowSel = llRow
                    grdRNMsg.Col = RN_CREATEDATEINDEX
                    grdRNMsg.ColSel = RN_MESSAGEINDEX
                    lmRNMsgRowSelected = llRow
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            Next llRow
        End If
        mSetCommands
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

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
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String

    imFirstActivate = True
    imTerminate = False

    Screen.MousePointer = vbHourglass
    'mParseCmmdLine
    'AlertVw.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    'gCenterStdAlone AlertVw
    'AlertVw.Show
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdCntr, grdAffExport, vbHourglass
    gSetMousePointer grdLog, grdRNMsg, vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    lmCntrRowSelected = -1
    lmRNMsgRowSelected = -1
    lmLogRowSelected = -1

    imFirstFocus = True
    imLastSelectRow = 0
    imCtrlKey = False
    imAufRecLen = Len(tmAuf)
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", AlertVw
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)  'Get and save CHF record length
    hmSof = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sof.Btr)", AlertVw
    On Error GoTo 0
    imSofRecLen = Len(tmSof)  'Get and save CHF record length
    hmCef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cef.Btr)", AlertVw
    On Error GoTo 0
    imCefRecLen = Len(tmCef)  'Get and save CHF record length
    hmVLF = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vlf.Btr)", AlertVw
    On Error GoTo 0
    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vsf.Btr)", AlertVw
    On Error GoTo 0
    'Convert Selling to Airing alerts as Logs are by Airing and Alerts entered by Selling
    gAlertVehicleReplace hmVLF

    mInitBox

    mPopulate
    'If ((Asc(tgSpf.sAutoType2) And RN_REP) <> RN_REP) And ((Asc(tgSpf.sAutoType2) And RN_NET) <> RN_NET) Then
    '    rbcView(3).Enabled = False
    'ElseIf (Trim$(grdRNMsg.TextMatrix(grdRNMsg.FixedRows, RN_CREATEDATEINDEX)) = "") Then
    '    rbcView(3).Enabled = False
    'End If
    If Not gPoolExist() Then
        tbcAlert.Tabs.Remove (6)
    Else
        mRemovePool
    End If
    If (((Asc(tgSpf.sAutoType2) And RN_REP) <> RN_REP) And ((Asc(tgSpf.sAutoType2) And RN_NET) <> RN_NET)) Then
        tbcAlert.Tabs.Remove (5)
    End If
    tbcAlert.Tabs.Remove (4)
    If (Asc(tgSaf(0).sFeatures5) And PROGRAMMATICALLOWED) <> PROGRAMMATICALLOWED Then
        tbcAlert.Tabs.Remove (3)
    ElseIf (tgUrf(0).sPrgmmaticAlert <> "I") And (tgUrf(0).sPrgmmaticAlert <> "V") Then
        tbcAlert.Tabs.Remove (3)
    End If
    If igAlertSpotStatus <> 1 Then
        If Trim$(grdCntr.TextMatrix(grdCntr.FixedRows, C_CREATEDATEINDEX)) <> "" Then
            'rbcView(0).Value = True
            tbcAlert.SelectedItem = tbcAlert.Tabs(1)
    
        ElseIf Trim$(grdLog.TextMatrix(grdLog.FixedRows, L_CREATEDATEINDEX)) <> "" Then
            'rbcView(1).Value = True
            tbcAlert.SelectedItem = tbcAlert.Tabs(2)
        'ElseIf Trim$(grdAffExport.TextMatrix(grdAffExport.FixedRows, A_CREATEDATEINDEX)) <> "" Then
        '    rbcView(2).Value = True
        ElseIf Asc(tgSaf(0).sFeatures5) And PROGRAMMATICALLOWED = PROGRAMMATICALLOWED Then
        
        'ElseIf (Trim$(grdRNMsg.TextMatrix(grdRNMsg.FixedRows, RN_CREATEDATEINDEX)) <> "") And (rbcView(3).Enabled) Then
        ElseIf (Trim$(grdRNMsg.TextMatrix(grdRNMsg.FixedRows, RN_CREATEDATEINDEX)) <> "") And (((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Or ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET)) Then
            'rbcView(3).Value = True
            If tbcAlert.Tabs(3).Caption = "Rep-Net Messages" Then
                tbcAlert.SelectedItem = tbcAlert.Tabs(3)
            ElseIf tbcAlert.Tabs(4).Caption = "Rep-Net Messages" Then
                tbcAlert.SelectedItem = tbcAlert.Tabs(4)
            ElseIf tbcAlert.Tabs(5).Caption = "Rep-Net Messages" Then
                tbcAlert.SelectedItem = tbcAlert.Tabs(5)
            End If
        Else
            'rbcView(0).Value = True
            tbcAlert.SelectedItem = tbcAlert.Tabs(1)
        End If
    Else
        'rbcView(1).Value = True
        tbcAlert.SelectedItem = tbcAlert.Tabs(2)
    End If
    igAlertCntrStatus = 0
    igAlertSpotStatus = 0
    sgAlertSpotVehicle = ""
    sgAlertSpotMoDate = ""
    If (tgSpf.sMktBase = "Y") Then
        ReDim igCntrMktCode(0 To 0) As Integer
        ilRet = gPopMnfPlusFieldsBox(AlertVw, lbcMkt, tmVehGp3Code(), smVehGp3CodeTag, "H3")
        For ilLoop = 0 To lbcMkt.ListCount - 1 Step 1
            slNameCode = tmVehGp3Code(ilLoop).sKey    'lbcVehCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            igCntrMktCode(UBound(igCntrMktCode)) = Val(slCode)
            ReDim Preserve igCntrMktCode(0 To UBound(igCntrMktCode) + 1) As Integer
        Next ilLoop
    End If
    If imTerminate Then
        Exit Sub
    End If
    cmcChgCntr.Enabled = False
    cmcViewCntr.Enabled = False
'    gCenterModalForm AlertVw
    Screen.MousePointer = vbDefault
    gSetMousePointer grdCntr, grdAffExport, vbDefault
    gSetMousePointer grdLog, grdRNMsg, vbDefault
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
    Unload AlertVw
    igManUnload = NO
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
    If grdCntr.Visible Then
        lmCntrRowSelected = -1
        grdCntr.Row = 0
        grdCntr.Col = C_AUFCODEINDEX
        mSetCommands
    End If
    If grdPrgmmatic.Visible Then
        lmPrgmmaticRowSelected = -1
        grdPrgmmatic.Row = 0
        grdPrgmmatic.Col = P_AUFCODEINDEX
        mSetCommands
    End If
    If grdRNMsg.Visible Then
        lmRNMsgRowSelected = -1
        grdRNMsg.Row = 0
        grdRNMsg.Col = RN_AUFCODEINDEX
        mSetCommands
    End If
    If grdLog.Visible Then
        lmLogRowSelected = -1
        grdLog.Row = 0
        grdLog.Col = L_AUFCODEINDEX
        mSetCommands
    End If
    If grdPool.Visible Then
        lmPoolRowSelected = -1
        grdPool.Row = 0
        grdPool.Col = UP_AUFCODEINDEX
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llRet                                                                                 *
'******************************************************************************************


    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim slTime As String
    Dim llTime As Long
    Dim ilAgf As Integer
    Dim ilAdf As Integer
    Dim ilVef As Integer
    Dim ilSlf As Integer
    Dim slType As String
    Dim ilIndex As Integer
    Dim ilShowAuf As Integer
    Dim ilUpdateAuf As Integer
    Dim llDate As Long
    Dim ilForceAlert As Integer
    Dim llCRow As Long
    Dim ilCol As Integer
    Dim llLRow As Long
    Dim llPRow As Long
    Dim llARow As Long
    Dim llRNRow As Long
    Dim slStartDate As String
    Dim llCntrDate As Long
    Dim slMessage As String
    Dim ilCntrFd As Integer
    Dim llURow As Long
    Dim llUPRow As Long
    Dim llRow As Long
    Dim slFields(0 To 7) As String

    imAufRecLen = Len(tmAuf)

    grdCntr.Redraw = False
    'grdCntr.Rows = 14
    For llCRow = grdCntr.FixedRows To grdCntr.Rows - 1 Step 1
        grdCntr.RowHeight(llCRow) = fgBoxGridH + 15
        For ilCol = 0 To grdCntr.Cols - 1 Step 1
            If ilCol = C_AUFCODEINDEX Then
                grdCntr.TextMatrix(llCRow, ilCol) = 0
            Else
                grdCntr.TextMatrix(llCRow, ilCol) = ""
            End If
        Next ilCol
    Next llCRow
    llCRow = grdCntr.FixedRows

    grdLog.Redraw = False
    'grdLog.Rows = 14
    For llLRow = grdLog.FixedRows To grdLog.Rows - 1 Step 1
        grdLog.RowHeight(llLRow) = fgBoxGridH + 15
        For ilCol = 0 To grdLog.Cols - 1 Step 1
            If ilCol = L_AUFCODEINDEX Then
                grdLog.TextMatrix(llLRow, ilCol) = 0
            Else
                grdLog.TextMatrix(llLRow, ilCol) = ""
            End If
        Next ilCol
    Next llLRow
    llLRow = grdLog.FixedRows

    grdPrgmmatic.Redraw = False
    'grdCntr.Rows = 14
    For llPRow = grdPrgmmatic.FixedRows To grdPrgmmatic.Rows - 1 Step 1
        grdPrgmmatic.RowHeight(llCRow) = fgBoxGridH + 15
        For ilCol = 0 To grdPrgmmatic.Cols - 1 Step 1
            If ilCol = P_AUFCODEINDEX Then
                grdPrgmmatic.TextMatrix(llCRow, ilCol) = 0
            Else
                grdPrgmmatic.TextMatrix(llCRow, ilCol) = ""
            End If
        Next ilCol
    Next llPRow
    llPRow = grdPrgmmatic.FixedRows

    grdAffExport.Redraw = False
    'grdAffExport.Rows = 14
    For llARow = grdAffExport.FixedRows To grdAffExport.Rows - 1 Step 1
        grdAffExport.RowHeight(llARow) = fgBoxGridH + 15
        For ilCol = 0 To grdAffExport.Cols - 1 Step 1
            If ilCol = A_AUFCODEINDEX Then
                grdAffExport.TextMatrix(llARow, ilCol) = 0
            Else
                grdAffExport.TextMatrix(llARow, ilCol) = ""
            End If
        Next ilCol
    Next llARow
    llARow = grdAffExport.FixedRows

    grdRNMsg.Redraw = False
    'grdRNMsg.Rows = 14
    For llRNRow = grdRNMsg.FixedRows To grdRNMsg.Rows - 1 Step 1
        grdRNMsg.RowHeight(llRNRow) = fgBoxGridH + 15
        For ilCol = 0 To grdRNMsg.Cols - 1 Step 1
            If ilCol = RN_AUFCODEINDEX Then
                grdRNMsg.TextMatrix(llRNRow, ilCol) = 0
            Else
                grdRNMsg.TextMatrix(llRNRow, ilCol) = ""
            End If
        Next ilCol
    Next llRNRow
    llRNRow = grdRNMsg.FixedRows

    grdPrgmmatic.Redraw = False
    'grdCntr.Rows = 14
    For llUPRow = grdPool.FixedRows To grdPool.Rows - 1 Step 1
        grdPool.RowHeight(llUPRow) = fgBoxGridH + 15
        For ilCol = 0 To grdPool.Cols - 1 Step 1
            If ilCol = UP_AUFCODEINDEX Then
                grdPool.TextMatrix(llUPRow, ilCol) = 0
            Else
                grdPool.TextMatrix(llUPRow, ilCol) = ""
            End If
        Next ilCol
    Next llUPRow
    llUPRow = grdPool.FixedRows
    
    ilForceAlert = False
    imScheduleExist = False
    imCompleteExist = False
    imPrgmmaticScheduleExist = False
    
    'Remove alerts for contracts erased
    'tmAufSrchKey1.sType = "C"
    'tmAufSrchKey1.sStatus = ""
    'ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    'Do While (ilRet = BTRV_ERR_NONE) And (tmAuf.sType = "C")
    '    tmChfSrchKey.lCode = tmAuf.lChfCode
    '    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '    If ilRet = BTRV_ERR_NONE Then
    '        ilRet = btrGetNext(hgAuf, tmAuf, imAufRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
    '    Else
    '        ilRet = gAlertContractErase(tmAuf.lChfCode)
    '        tmAufSrchKey1.sType = "C"
    '        tmAufSrchKey1.sStatus = ""
    '        ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    '        'ilRet = btrGetGreaterOrEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    '    End If
    'Loop

    ReDim llAufCode(0 To 0) As Long
    '5/24/11: Because of speed issues with this code, only going to remove those that cause an issue
    'tmAufSrchKey1.sType = "C"
    'tmAufSrchKey1.sStatus = ""
    'ilRet = btrGetGreaterOrEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    'Do While (ilRet = BTRV_ERR_NONE) And (tmAuf.sType = "C")
    '    tmChfSrchKey.lCode = tmAuf.lChfCode
    '    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '    If ilRet <> BTRV_ERR_NONE Then
    '        llAufCode(UBound(llAufCode)) = tmAuf.lCode
    '        ReDim Preserve llAufCode(0 To UBound(llAufCode) + 1) As Long
    '    End If
    '    ilRet = btrGetNext(hgAuf, tmAuf, imAufRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
    'Loop
    'For ilLoop = 0 To UBound(llAufCode) - 1 Step 1
    '    tmAufSrchKey0.lCode = llAufCode(ilLoop)
    '    ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    '    If ilRet = BTRV_ERR_NONE Then
    '        ilRet = btrDelete(hgAuf)
    '    End If
    'Next ilLoop

    For ilLoop = 0 To 6 Step 1
        '7/23/12: Only show Traffic alters
        If (ilLoop = 0) Or (ilLoop = 1) Or (ilLoop = 5) Or (ilLoop = 6) Then
            ReDim tmAufView(0 To 0) As AUFVIEW
            Select Case ilLoop
                Case 0
                    slType = "C"    'Contract
                Case 1
                    slType = "L"    'Log
                Case 2
                    slType = "F"    'Affiliate Export All
                Case 3
                    slType = "R"    'Affiliate Export Change
                Case 4
                    slType = "P"    'Program Change
                Case 5
                    slType = "M"    'Messages
                Case 6
                    slType = "U"    'Unassigned Pool
            End Select
            tmAufSrchKey1.sType = slType
            tmAufSrchKey1.sStatus = "R"
            ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            Do While (ilRet = BTRV_ERR_NONE) And (slType = tmAuf.sType) And (tmAuf.sStatus = "R")
                gUnpackDateForSort tmAuf.iEnteredDate(0), tmAuf.iEnteredDate(1), slDate
                gUnpackTimeLong tmAuf.iEnteredTime(0), tmAuf.iEnteredTime(1), False, llTime
                slTime = Trim$(str$(llTime))
                Do While Len(slTime) < 6
                    slTime = "0" & slTime
                Loop
                tmAufView(UBound(tmAufView)).sKey = slDate & slTime
                tmAufView(UBound(tmAufView)).tAuf = tmAuf
                ReDim Preserve tmAufView(0 To UBound(tmAufView) + 1) As AUFVIEW
                ilRet = btrGetNext(hgAuf, tmAuf, imAufRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
            Loop
            If UBound(tmAufView) - 1 > 0 Then
                ArraySortTyp fnAV(tmAufView(), 0), UBound(tmAufView), 1, LenB(tmAufView(0)), 0, LenB(tmAufView(0).sKey), 0
            End If
            For ilIndex = 0 To UBound(tmAufView) - 1 Step 1
                'Test if status "R" is still valid.
                'If Contract, check status
                'If Spot or Copy, check date
                tmAuf = tmAufView(ilIndex).tAuf
                If tmAuf.sType = "C" Then
                    tmChfSrchKey.lCode = tmAuf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If tmAuf.sSubType = "C" Then
                            If igWinStatus(PROPOSALSJOB) = 0 Or tmChf.sSource = "P" Then
                                ilShowAuf = False
                                ilUpdateAuf = False
                            Else
                                If tmChf.sStatus <> tmAuf.sSubType Then
                                    ilShowAuf = False
                                    ilUpdateAuf = True
                                Else
                                    ilShowAuf = True
                                    ilUpdateAuf = False
                                End If
                            End If
                        '8/15/18: Add Unapproved alerts
                        ElseIf tmAuf.sSubType = "I" Then    'Unapproved
                            If igWinStatus(PROPOSALSJOB) = 0 Or tmChf.sSource = "P" Then
                                ilShowAuf = False
                                ilUpdateAuf = False
                            Else
                                If tmChf.sStatus <> tmAuf.sSubType Then
                                    ilShowAuf = False
                                    ilUpdateAuf = True
                                Else
                                    ilShowAuf = True
                                    ilUpdateAuf = False
                                End If
                            End If
                        ElseIf tmAuf.sSubType = "S" Then
                            If igWinStatus(CONTRACTSJOB) = 0 Or tmChf.sSource = "P" Then
                                ilShowAuf = False
                                ilUpdateAuf = False
                            Else
                                If (tmChf.sSchStatus = "A") Or (tmChf.sSchStatus = "N") Or (tmChf.sSchStatus = "I") Then
                                    ilShowAuf = True
                                    ilUpdateAuf = False
                                Else
                                    ilShowAuf = False
                                    ilUpdateAuf = True
                                End If
                            End If
                        ElseIf tmAuf.sSubType = "P" Then    'Programmatic Buy
                            If tmChf.sSource = "P" Then
                                If (tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M") Or (tmChf.sSchStatus = "I") Then
                                    gUnpackDateLong tmChf.iOHDDate(0), tmChf.iOHDDate(1), llCntrDate
                                    If llCntrDate + 7 < gDateValue(gNow()) Then
                                        ilShowAuf = False
                                        ilUpdateAuf = True
                                    ElseIf igWinStatus(CONTRACTSJOB) = 0 Then
                                        ilShowAuf = False
                                        ilUpdateAuf = False
                                    Else
                                        ilShowAuf = True
                                        ilUpdateAuf = False
                                    End If
                                Else
                                    If igWinStatus(PROPOSALSJOB) = 0 Then
                                        ilShowAuf = False
                                        ilUpdateAuf = False
                                    Else
                                        ilShowAuf = True
                                        ilUpdateAuf = False
                                    End If
                                End If
                            Else
                                ilShowAuf = False
                                ilUpdateAuf = False
                            End If
                        End If
                    Else
                        ilShowAuf = False
                        '8/18/10:  Retain contract alert
                        ilUpdateAuf = False 'True
                        '5/24/11: Remove those in error
                        'If ilRet = 30002 Then
                        '    ilRet = csiHandleValue(0, 7)
                        'End If
                        'MsgBox "Contract Not Found, ChfCode " & tmAuf.lChfCode & ", Error " & ilRet & "Call Counterpoint", vbOKOnly + vbInformation, "Information"
                        'gLogMsg "View Alerts: Contract Not Found, ChfCode " & tmAuf.lChfCode & ", Error " & ilRet, "TrafficErrors.Txt", False
                        llAufCode(UBound(llAufCode)) = tmAuf.lCode
                        ReDim Preserve llAufCode(0 To UBound(llAufCode) + 1) As Long
                    End If
                    If (ilShowAuf) And (tmAuf.sSubType = "C") Then
                        If tgUrf(0).sCompAlert = "N" Then
                            ilShowAuf = False
                        End If
                    End If
                    If (ilShowAuf) And (tmAuf.sSubType = "I") Then
                        If tgUrf(0).sIncompAlert = "N" Then
                            ilShowAuf = False
                        End If
                    End If
                    If (ilShowAuf) And (tmAuf.sSubType = "S") Then
                        If tgUrf(0).sSchAlert = "N" Then
                            ilShowAuf = False
                        End If
                    End If
                    If (ilShowAuf) And (tmAuf.sSubType = "P") Then
                        If (tgUrf(0).sPrgmmaticAlert <> "I") And (tgUrf(0).sPrgmmaticAlert <> "V") Then
                            ilShowAuf = False
                        End If
                    End If
                    If ilShowAuf Then
                        ilShowAuf = gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode(), tmChf.sSource)
                    End If
                ElseIf tmAuf.sType = "L" Then
                    If (igWinStatus(LOGSJOB) = 0) And (igWinStatus(SPOTSJOB) = 0) Then
                        ilShowAuf = False
                        ilUpdateAuf = False
                    Else
                        gUnpackDateLong tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), llDate
                        If llDate + 6 < gDateValue(Format$(gNow(), "m/d/yy")) Then   '4-27-04
                            ilShowAuf = False
                            ilUpdateAuf = True
                        Else
                            ilShowAuf = True
                            ilUpdateAuf = False
                        End If
                    End If
                ElseIf tmAuf.sType = "M" Then
                    ilUpdateAuf = False
                    If tgUrf(0).sShowNRMsg = "N" Then
                        ilShowAuf = False
                    Else
                        ilShowAuf = True
                    End If
                ElseIf tmAuf.sType = "U" Then
                    gUnpackDateLong tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), llDate
                    If llDate + 6 < gDateValue(Format$(Now, "m/d/yy")) Then
                        ilShowAuf = False
                        ilUpdateAuf = True
                    Else
                        If tmAuf.sSubType = "P" Then
                            ilShowAuf = True
                            ilUpdateAuf = False
                        Else
                            ilShowAuf = False
                            ilUpdateAuf = False
                        End If
                    End If
                Else
                    ilShowAuf = True
                    ilUpdateAuf = False
                End If
                If ilShowAuf Then
                    Select Case ilLoop
                        Case 0
                            If tmAuf.sSubType <> "P" Then
                                If llCRow >= grdCntr.Rows Then
                                    grdCntr.AddItem ""
                                    grdCntr.RowHeight(llCRow) = fgBoxGridH + 15
                                End If
                            Else
                                If llPRow >= grdPrgmmatic.Rows Then
                                    grdPrgmmatic.AddItem ""
                                    grdPrgmmatic.RowHeight(llCRow) = fgBoxGridH + 15
                                End If
                            End If
                        Case 1
                            If llLRow >= grdLog.Rows Then
                                grdLog.AddItem ""
                                grdLog.RowHeight(llLRow) = fgBoxGridH + 15
                            End If
                        Case 2
                            If llARow >= grdAffExport.Rows Then
                                grdAffExport.AddItem ""
                                grdAffExport.RowHeight(llARow) = fgBoxGridH + 15
                            End If
                        Case 3
                            If llARow >= grdAffExport.Rows Then
                                grdAffExport.AddItem ""
                                grdAffExport.RowHeight(llARow) = fgBoxGridH + 15
                            End If
                        Case 4
                            If llARow >= grdAffExport.Rows Then
                                grdAffExport.AddItem ""
                                grdAffExport.RowHeight(llARow) = fgBoxGridH + 15
                            End If
                        Case 5
                            If llRNRow >= grdRNMsg.Rows Then
                                grdRNMsg.AddItem ""
                                grdRNMsg.RowHeight(llRNRow) = fgBoxGridH + 15
                            End If
                        Case 6
                            If llUPRow >= grdPool.Rows Then
                                grdPool.AddItem ""
                                grdPool.RowHeight(llUPRow) = fgBoxGridH + 15
                            End If
                    End Select
                    gUnpackDate tmAuf.iEnteredDate(0), tmAuf.iEnteredDate(1), slDate
                    gUnpackTime tmAuf.iEnteredTime(0), tmAuf.iEnteredTime(1), "A", "1", slTime
                    If slType = "C" Then    'Contract
                        If tmAuf.sSubType <> "P" Then
                            grdCntr.TextMatrix(llCRow, C_CREATEDATEINDEX) = slDate
                            grdCntr.TextMatrix(llCRow, C_CREATETIMEINDEX) = slTime
                            grdCntr.TextMatrix(llCRow, C_CNTRNOINDEX) = Trim$(str$(tmChf.lCntrNo))
        
                            ilAdf = gBinarySearchAdf(tmChf.iAdfCode)
                            If ilAdf <> -1 Then
                                If (tgCommAdf(ilAdf).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdf).sAddrID) <> "") Then
                                    grdCntr.TextMatrix(llCRow, C_ADVTNAMEINDEX) = Trim$(tgCommAdf(ilAdf).sName) & ", " & Trim$(tgCommAdf(ilAdf).sAddrID)
                                Else
                                    grdCntr.TextMatrix(llCRow, C_ADVTNAMEINDEX) = Trim$(tgCommAdf(ilAdf).sName)
                                End If
                            Else
                                grdCntr.TextMatrix(llCRow, C_ADVTNAMEINDEX) = ""
                            End If
                            gUnpackDate tmChf.iStartDate(0), tmChf.iStartDate(1), slStartDate
                            grdCntr.TextMatrix(llCRow, C_CNTRSTARTDATEINDEX) = slStartDate
                            grdCntr.TextMatrix(llCRow, C_PRODUCTINDEX) = Trim$(tmChf.sProduct)
                            grdCntr.TextMatrix(llCRow, C_SALESOFFICEINDEX) = ""
                            ilSlf = gBinarySearchSlf(tmChf.iSlfCode(0))
                            If ilSlf <> -1 Then
                                If tgMSlf(ilSlf).iSofCode <> 0 Then
                                    tmSofSrchKey.iCode = tgMSlf(ilSlf).iSofCode
                                    ilRet = btrGetEqual(hmSof, tmSof, imSofRecLen, tmSofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                    If ilRet = BTRV_ERR_NONE Then
                                        grdCntr.TextMatrix(llCRow, C_SALESOFFICEINDEX) = Trim$(tmSof.sName)
                                    End If
                                End If
                            End If
                            If tmAuf.sSubType = "C" Then
                                imCompleteExist = True
                                If tmChf.iCntRevNo > 0 Then
                                    grdCntr.TextMatrix(llCRow, C_REASONINDEX) = "Rev Complete"
                                Else
                                    grdCntr.TextMatrix(llCRow, C_REASONINDEX) = "Complete"
                                End If
                            ElseIf tmAuf.sSubType = "S" Then
                                imScheduleExist = True
                                If (tmChf.sSchStatus = "A") Or (tmChf.sSchStatus = "N") Then
                                    grdCntr.TextMatrix(llCRow, C_REASONINDEX) = "Not Scheduled"
                                ElseIf (tmChf.sSchStatus = "I") Then
                                    grdCntr.TextMatrix(llCRow, C_REASONINDEX) = "Scheduling"
                                End If
                            ElseIf tmAuf.sSubType = "I" Then
                                imCompleteExist = True
                                grdCntr.TextMatrix(llCRow, C_REASONINDEX) = "Unapproved"
                            Else
                                grdCntr.TextMatrix(llCRow, C_REASONINDEX) = ""
                            End If
                            grdCntr.TextMatrix(llCRow, C_AUFCODEINDEX) = tmAuf.lCode
                            llCRow = llCRow + 1
                        Else
                            ilCntrFd = 0    'Match not found, add row
                            llURow = llPRow
                            For llRow = grdPrgmmatic.FixedRows To grdPrgmmatic.Rows - 1 Step 1
                                If grdPrgmmatic.TextMatrix(llRow, P_CREATEDATEINDEX) <> "" Then
                                    If Val(grdPrgmmatic.TextMatrix(llRow, P_CNTRNOINDEX)) = tmChf.lCntrNo Then
                                        ilCntrFd = 1    'Match found, update
                                        llURow = llRow
                                        'Determine if update current displayed or leave current displayed
                                        'Test date and time created.  Not a perfect solution
                                        If Val(grdPrgmmatic.TextMatrix(llRow, P_AUFCODEINDEX)) > tmAuf.lCode Then
                                            ilCntrFd = 2    'Bypass this auf record
                                            tmAufSrchKey0.lCode = tmAuf.lCode
                                            ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                            If ilRet = BTRV_ERR_NONE Then
                                                ilForceAlert = True
                                                tmAuf.sStatus = "C"
                                                tmAuf.sClearMethod = "M"
                                                slDate = Format$(gNow(), "m/d/yy")
                                                slTime = Format$(gNow(), "h:mm:ssAM/PM")
                                                gPackDate slDate, tmAuf.iClearDate(0), tmAuf.iClearDate(1)
                                                gPackTime slTime, tmAuf.iClearTime(0), tmAuf.iClearTime(1)
                                                tmAuf.iClearUrfCode = tgUrf(0).iCode
                                                ilRet = btrUpdate(hgAuf, tmAuf, imAufRecLen)
                                            End If
                                        End If
                                        Exit For
                                    End If
                                End If
                            Next llRow
                            If ilCntrFd <= 1 Then
                                grdPrgmmatic.TextMatrix(llURow, P_CREATEDATEINDEX) = slDate
                                grdPrgmmatic.TextMatrix(llURow, P_CREATETIMEINDEX) = slTime
                                grdPrgmmatic.TextMatrix(llURow, P_CNTRNOINDEX) = Trim$(str$(tmChf.lCntrNo))
            
                                If tmChf.iAgfCode > 0 Then
                                    ilAgf = gBinarySearchAgf(tmChf.iAgfCode)
                                    If ilAgf <> -1 Then
                                        grdPrgmmatic.TextMatrix(llURow, P_AGENCYINDEX) = Trim$(tgCommAgf(ilAgf).sName) & ", " & Trim$(tgCommAgf(ilAgf).sCityID)
                                    Else
                                        grdPrgmmatic.TextMatrix(llURow, P_AGENCYINDEX) = ""
                                    End If
                                Else
                                    grdPrgmmatic.TextMatrix(llURow, P_AGENCYINDEX) = ""
                                End If
                                
                                ilAdf = gBinarySearchAdf(tmChf.iAdfCode)
                                If ilAdf <> -1 Then
                                    If (tgCommAdf(ilAdf).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdf).sAddrID) <> "") Then
                                        grdPrgmmatic.TextMatrix(llURow, P_ADVTNAMEINDEX) = Trim$(tgCommAdf(ilAdf).sName) & ", " & Trim$(tgCommAdf(ilAdf).sAddrID)
                                    Else
                                        grdPrgmmatic.TextMatrix(llURow, P_ADVTNAMEINDEX) = Trim$(tgCommAdf(ilAdf).sName)
                                    End If
                                Else
                                    grdPrgmmatic.TextMatrix(llURow, P_ADVTNAMEINDEX) = ""
                                End If
                                gUnpackDate tmChf.iStartDate(0), tmChf.iStartDate(1), slStartDate
                                grdPrgmmatic.TextMatrix(llURow, P_CNTRSTARTDATEINDEX) = slStartDate
                                grdPrgmmatic.TextMatrix(llURow, P_PRODUCTINDEX) = Trim$(tmChf.sProduct)
                                grdPrgmmatic.TextMatrix(llURow, P_SALESOFFICEINDEX) = ""
                                ilSlf = gBinarySearchSlf(tmChf.iSlfCode(0))
                                If ilSlf <> -1 Then
                                    If tgMSlf(ilSlf).iSofCode <> 0 Then
                                        tmSofSrchKey.iCode = tgMSlf(ilSlf).iSofCode
                                        ilRet = btrGetEqual(hmSof, tmSof, imSofRecLen, tmSofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                        If ilRet = BTRV_ERR_NONE Then
                                            grdPrgmmatic.TextMatrix(llURow, P_SALESOFFICEINDEX) = Trim$(tmSof.sName)
                                        End If
                                    End If
                                End If
                                'W=Working Proposal; D=Dead Proposal; C=Completed Proposal; I=Incomplete Proposal;
                                'H=Hold; O=Order
                                grdPrgmmatic.TextMatrix(llURow, P_CNTRSTATUSINDEX) = tmChf.sStatus
                                Select Case tmChf.sStatus
                                    Case "W"
                                        If tmChf.iCntRevNo > 0 Then
                                            grdPrgmmatic.TextMatrix(llURow, P_REASONINDEX) = "Rev Working"
                                            grdPrgmmatic.TextMatrix(llURow, P_CNTRSTATUSINDEX) = "R" & tmChf.sStatus
                                        Else
                                            grdPrgmmatic.TextMatrix(llURow, P_REASONINDEX) = "Working"
                                        End If
                                    Case "D"
                                        grdPrgmmatic.TextMatrix(llURow, P_REASONINDEX) = "Rejected"
                                    Case "C"
                                        If tmChf.iCntRevNo > 0 Then
                                            grdPrgmmatic.TextMatrix(llURow, P_REASONINDEX) = "Rev Complete"
                                            grdPrgmmatic.TextMatrix(llURow, P_CNTRSTATUSINDEX) = "R" & tmChf.sStatus
                                        Else
                                            grdPrgmmatic.TextMatrix(llURow, P_REASONINDEX) = "Complete"
                                        End If
                                    Case "I"
                                        If tmChf.iCntRevNo > 0 Then
                                            grdPrgmmatic.TextMatrix(llURow, P_REASONINDEX) = "Rev Unapproved"
                                            grdPrgmmatic.TextMatrix(llURow, P_CNTRSTATUSINDEX) = "R" & tmChf.sStatus
                                        Else
                                            grdPrgmmatic.TextMatrix(llURow, P_REASONINDEX) = "Unapproved"
                                        End If
                                    Case "G"
                                        imPrgmmaticScheduleExist = True
                                        If (tmChf.sSchStatus = "I") Then
                                            grdPrgmmatic.TextMatrix(llURow, P_REASONINDEX) = "Scheduling Hold"
                                        Else
                                            grdPrgmmatic.TextMatrix(llURow, P_REASONINDEX) = "Approved Hold"
                                        End If
                                    Case "N"
                                        imPrgmmaticScheduleExist = True
                                        If (tmChf.sSchStatus = "I") Then
                                            grdPrgmmatic.TextMatrix(llURow, P_REASONINDEX) = "Scheduling Order"
                                        Else
                                            grdPrgmmatic.TextMatrix(llURow, P_REASONINDEX) = "Approved Order"
                                        End If
                                    Case "H"
                                        grdPrgmmatic.TextMatrix(llURow, P_REASONINDEX) = "Hold"
                                    Case "O"
                                        grdPrgmmatic.TextMatrix(llURow, P_REASONINDEX) = "Order"
                                    Case Else
                                        grdPrgmmatic.TextMatrix(llURow, P_REASONINDEX) = ""
                                End Select
                                grdPrgmmatic.TextMatrix(llURow, P_AUFCODEINDEX) = tmAuf.lCode
                                If ilCntrFd = 0 Then
                                    llPRow = llPRow + 1
                                End If
                            End If
                        End If
                    ElseIf slType = "L" Then    'Log
                        grdLog.TextMatrix(llLRow, L_CREATEDATEINDEX) = slDate
                        grdLog.TextMatrix(llLRow, L_CREATETIMEINDEX) = slTime
    
                        ilVef = gBinarySearchVef(tmAuf.iVefCode)
                        If ilVef <> -1 Then
                            grdLog.TextMatrix(llLRow, L_VEHICLEINDEX) = Trim$(tgMVef(ilVef).sName)
                        Else
                            grdLog.TextMatrix(llLRow, L_VEHICLEINDEX) = ""
                        End If
                        gUnpackDate tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), slDate
                        grdLog.TextMatrix(llLRow, L_LOGDATEINDEX) = slDate
                        If tmAuf.sSubType = "C" Then
                            grdLog.TextMatrix(llLRow, L_REASONINDEX) = "Copy Changed"
                        ElseIf tmAuf.sSubType = "S" Then
                            grdLog.TextMatrix(llLRow, L_REASONINDEX) = "Spot Changed"
                        ElseIf tmAuf.sSubType = "M" Then
                            grdLog.TextMatrix(llLRow, L_REASONINDEX) = "Missed in Closed Week"
                        Else
                            grdLog.TextMatrix(llLRow, L_REASONINDEX) = ""
                        End If
                        grdLog.TextMatrix(llLRow, L_AUFCODEINDEX) = tmAuf.lCode
                        llLRow = llLRow + 1
                    ElseIf slType = "F" Then    'Affiliate Export-Final
                        grdAffExport.TextMatrix(llARow, A_CREATEDATEINDEX) = slDate
                        grdAffExport.TextMatrix(llARow, A_CREATETIMEINDEX) = slTime
    
                        ilVef = gBinarySearchVef(tmAuf.iVefCode)
                        If ilVef <> -1 Then
                            grdAffExport.TextMatrix(llARow, A_VEHICLEINDEX) = Trim$(tgMVef(ilVef).sName)
                        Else
                            grdAffExport.TextMatrix(llARow, A_VEHICLEINDEX) = ""
                        End If
                        gUnpackDate tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), slDate
                        grdAffExport.TextMatrix(llARow, A_LOGDATEINDEX) = slDate
                        If tmAuf.sSubType = "I" Then
                            grdAffExport.TextMatrix(llARow, A_REASONINDEX) = "New-ISCI"
                        ElseIf tmAuf.sSubType = "S" Then
                            grdAffExport.TextMatrix(llARow, A_REASONINDEX) = "New-Spot"
                        Else
                            grdAffExport.TextMatrix(llARow, A_REASONINDEX) = ""
                        End If
                        grdAffExport.TextMatrix(llARow, A_AUFCODEINDEX) = tmAuf.lCode
                        llARow = llARow + 1
                    ElseIf slType = "R" Then    'Affiliate Export-Reprint
                        grdAffExport.TextMatrix(llARow, A_CREATEDATEINDEX) = slDate
                        grdAffExport.TextMatrix(llARow, A_CREATETIMEINDEX) = slTime
    
                        ilVef = gBinarySearchVef(tmAuf.iVefCode)
                        If ilVef <> -1 Then
                            grdAffExport.TextMatrix(llARow, A_VEHICLEINDEX) = Trim$(tgMVef(ilVef).sName)
                        Else
                            grdAffExport.TextMatrix(llARow, A_VEHICLEINDEX) = ""
                        End If
                        gUnpackDate tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), slDate
                        grdAffExport.TextMatrix(llARow, A_LOGDATEINDEX) = slDate
                        If tmAuf.sSubType = "I" Then
                            grdAffExport.TextMatrix(llARow, A_REASONINDEX) = "Changed-ISCI"
                        ElseIf tmAuf.sSubType = "S" Then
                            grdAffExport.TextMatrix(llARow, A_REASONINDEX) = "Changed-Spot"
                        Else
                            grdAffExport.TextMatrix(llARow, A_REASONINDEX) = ""
                        End If
                        grdAffExport.TextMatrix(llARow, A_AUFCODEINDEX) = tmAuf.lCode
                        llARow = llARow + 1
                    ElseIf slType = "P" Then    'Affiliate Export-Reprint
                        grdAffExport.TextMatrix(llARow, A_CREATEDATEINDEX) = slDate
                        grdAffExport.TextMatrix(llARow, A_CREATETIMEINDEX) = slTime
    
                        ilVef = gBinarySearchVef(tmAuf.iVefCode)
                        If ilVef <> -1 Then
                            grdAffExport.TextMatrix(llARow, A_VEHICLEINDEX) = Trim$(tgMVef(ilVef).sName)
                        Else
                            grdAffExport.TextMatrix(llARow, A_VEHICLEINDEX) = ""
                        End If
                        gUnpackDate tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), slDate
                        grdAffExport.TextMatrix(llARow, A_LOGDATEINDEX) = slDate
                        If tmAuf.sSubType = "A" Then
                            grdAffExport.TextMatrix(llARow, A_REASONINDEX) = "Changed-Program"
                        Else
                            grdAffExport.TextMatrix(llARow, A_REASONINDEX) = ""
                        End If
                        grdAffExport.TextMatrix(llARow, A_AUFCODEINDEX) = tmAuf.lCode
                        llARow = llARow + 1
                    ElseIf slType = "M" Then    'Messages
                        tmCefSrchKey.lCode = tmAuf.lCefCode
                        tmCef.sComment = ""
                        imCefRecLen = Len(tmCef)    '1009
                        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            'If tmCef.iStrLen > 0 Then
                            '    slMessage = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                            'End If
                            slMessage = gStripChr0(tmCef.sComment)
                        End If
                        grdRNMsg.TextMatrix(llRNRow, RN_CREATEDATEINDEX) = slDate
                        grdRNMsg.TextMatrix(llRNRow, RN_CREATETIMEINDEX) = slTime
                        gParseItemFields slMessage, "|", slFields()
                        For ilCol = UBound(slFields) - 1 To LBound(slFields) Step -1
                            slFields(ilCol + 1) = slFields(ilCol)
                        Next ilCol
                        slFields(0) = ""
                        If slFields(1) = "C" Then
                            grdRNMsg.TextMatrix(llRNRow, RN_TYPEMSGINDEX) = "Contract"
                        ElseIf slFields(1) = "P" Then
                            grdRNMsg.TextMatrix(llRNRow, RN_TYPEMSGINDEX) = "Posted Spot"
                        ElseIf slFields(1) = "A" Then
                            grdRNMsg.TextMatrix(llRNRow, RN_TYPEMSGINDEX) = "Avail Summary"
                        ElseIf slFields(1) = "R" Then
                            grdRNMsg.TextMatrix(llRNRow, RN_TYPEMSGINDEX) = "Copy"
                        ElseIf slFields(1) = "F" Then
                            grdRNMsg.TextMatrix(llRNRow, RN_TYPEMSGINDEX) = "Find Field"
                        ElseIf slFields(1) = "O" Then
                            grdRNMsg.TextMatrix(llRNRow, RN_TYPEMSGINDEX) = "Open File"
                        ElseIf slFields(1) = "M" Then
                            grdRNMsg.TextMatrix(llRNRow, RN_TYPEMSGINDEX) = "Message"
                        End If
                        If slFields(2) = "E" Then
                            grdRNMsg.TextMatrix(llRNRow, RN_STATUSINDEX) = "Error"
                        ElseIf slFields(2) = "O" Then
                            grdRNMsg.TextMatrix(llRNRow, RN_STATUSINDEX) = "Ok"
                        End If
                        grdRNMsg.TextMatrix(llRNRow, RN_FROMIDINDEX) = slFields(3)
                        grdRNMsg.TextMatrix(llRNRow, RN_CNTRNOINDEX) = slFields(4)
                        grdRNMsg.TextMatrix(llRNRow, RN_ADVTNAMEINDEX) = slFields(5)
                        grdRNMsg.TextMatrix(llRNRow, RN_VEHICLEINDEX) = slFields(6)
                        grdRNMsg.TextMatrix(llRNRow, RN_MESSAGEINDEX) = slFields(7)
                        grdRNMsg.TextMatrix(llRNRow, RN_AUFCODEINDEX) = tmAuf.lCode
                        llRNRow = llRNRow + 1
                    ElseIf slType = "U" Then    'Unassigned Pool
                        grdPool.TextMatrix(llUPRow, UP_CREATEDATEINDEX) = slDate
                        grdPool.TextMatrix(llUPRow, UP_CREATETIMEINDEX) = slTime
                        gUnpackDate tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), slDate
                        grdPool.TextMatrix(llUPRow, UP_FILENAMEINDEX) = "See PoolUnassignedLog_" & Format(slDate, "mm-dd-yy") & ".txt" & " in Messages Subfolder"
                        grdPool.TextMatrix(llUPRow, UP_FILEDATEINDEX) = slDate
                        grdPool.TextMatrix(llUPRow, UP_AUFCODEINDEX) = tmAuf.lCode
                        llUPRow = llUPRow + 1
                    End If
                Else
                    If ilUpdateAuf Then
                        tmAufSrchKey0.lCode = tmAuf.lCode
                        ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet = BTRV_ERR_NONE Then
                            ilForceAlert = True
                            tmAuf.sStatus = "C"
                            tmAuf.sClearMethod = "M"
                            slDate = Format$(gNow(), "m/d/yy")
                            slTime = Format$(gNow(), "h:mm:ssAM/PM")
                            gPackDate slDate, tmAuf.iClearDate(0), tmAuf.iClearDate(1)
                            gPackTime slTime, tmAuf.iClearTime(0), tmAuf.iClearTime(1)
                            tmAuf.iClearUrfCode = tgUrf(0).iCode
                            ilRet = btrUpdate(hgAuf, tmAuf, imAufRecLen)
                        End If
                    End If
                End If
            Next ilIndex
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(llAufCode) - 1 Step 1
        tmAufSrchKey0.lCode = llAufCode(ilLoop)
        ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hgAuf)
        End If
    Next ilLoop
    If ilForceAlert Then
        ilRet = gAlertForceCheck()
    End If
    mSetCommands
    'Remove highlight
    mCntrSortCol C_CREATETIMEINDEX
    mCntrSortCol C_CREATEDATEINDEX
    grdCntr.Row = 0
    grdCntr.Col = C_AUFCODEINDEX
    lmCntrRowSelected = -1
    grdCntr.Redraw = True
    mLogSortCol L_CREATETIMEINDEX
    mLogSortCol L_CREATEDATEINDEX
    grdLog.Row = 0
    grdLog.Col = L_AUFCODEINDEX
    lmLogRowSelected = -1
    grdLog.Redraw = True
    mPrgmmaticSortCol P_CNTRNOINDEX
    mPrgmmaticSortCol P_ADVTNAMEINDEX
    mPrgmmaticSortCol P_AGENCYINDEX
    mPrgmmaticSortCol P_REASONINDEX
    grdPrgmmatic.Row = 0
    grdPrgmmatic.Col = P_AUFCODEINDEX
    lmPrgmmaticRowSelected = -1
    grdPrgmmatic.Redraw = True
    mAffSortCol A_CREATETIMEINDEX
    mAffSortCol A_CREATEDATEINDEX
    grdAffExport.Row = 0
    grdAffExport.Col = A_AUFCODEINDEX
    grdAffExport.Redraw = True
    mRNMsgSortCol RN_CREATETIMEINDEX
    mRNMsgSortCol RN_CREATEDATEINDEX
    mRNMsgSortCol RN_FROMIDINDEX
    mRNMsgSortCol RN_CNTRNOINDEX
    grdRNMsg.Row = 0
    grdRNMsg.Col = RN_AUFCODEINDEX
    lmRNMsgRowSelected = -1
    grdRNMsg.Redraw = True
    mPoolSortCol UP_CREATETIMEINDEX
    mPoolSortCol UP_CREATEDATEINDEX
    grdPool.Row = 0
    grdPool.Col = UP_AUFCODEINDEX
    lmPoolRowSelected = -1
    grdPool.Redraw = True
End Sub

Private Sub rbcView_Click(Index As Integer)
    If rbcView(Index).Value Then
        grdCntr.Visible = False
        grdLog.Visible = False
        grdAffExport.Visible = False
        grdRNMsg.Visible = False
        grdPool.Visible = False
        If Index = 0 Then
            grdCntr.Visible = True
        End If
        If Index = 1 Then
            grdLog.Visible = True
        End If
        If Index = 2 Then
            grdAffExport.Visible = True
        End If
        If Index = 3 Then
            grdRNMsg.Visible = True
        End If
        If Index = 4 Then
            grdPool.Visible = True
        End If
        mSetCommands
    End If
End Sub

Private Sub mSetCommands()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRow                                                                                 *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  cmcEraseErr                                                                           *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilOk As Integer

    'If rbcView(0).Value Then
    If imTabSelected = 0 Then
        cmcSpots.Enabled = False
        If imCompleteExist Then
            cmcChgCntr.Caption = "&Change Proposal"
            cmcViewCntr.Caption = "&View Proposal"
            If tgSpf.sGUsePropSys = "Y" Then
                If lmCntrRowSelected >= grdCntr.FixedRows Then
                    ilOk = False
                    If (Not igJobShowing(CONTRACTSJOB)) Then
                        ilOk = True
                    Else
                        If (Contract.cmcUpdate.Enabled = False) Then
                            ilOk = True
                        End If
                    End If
                    If ilOk Then
                        If igWinStatus(PROPOSALSJOB) = 1 Then
                            cmcChgCntr.Enabled = False
                            cmcViewCntr.Enabled = True
                        Else
                            If Trim$(grdCntr.TextMatrix(lmCntrRowSelected, C_REASONINDEX)) = "Rev Complete" Then
                                If tgUrf(0).sReviseCntr <> "N" Then
                                    cmcChgCntr.Enabled = True
                                Else
                                    cmcChgCntr.Enabled = False
                                End If
                            ElseIf Trim$(grdCntr.TextMatrix(lmCntrRowSelected, C_REASONINDEX)) = "Complete" Then
                                cmcChgCntr.Enabled = True
                            ElseIf Trim$(grdCntr.TextMatrix(lmCntrRowSelected, C_REASONINDEX)) = "Unapproved" Then
                                cmcChgCntr.Enabled = True
                            Else
                                cmcChgCntr.Enabled = False
                            End If
                            cmcViewCntr.Enabled = True
                        End If
                    Else
                        cmcChgCntr.Enabled = False
                        cmcViewCntr.Enabled = False
                    End If
                Else
                    cmcChgCntr.Enabled = False
                    cmcViewCntr.Enabled = False
                End If
            Else
                cmcChgCntr.Enabled = False
                cmcViewCntr.Enabled = False
            End If
        Else
            cmcChgCntr.Enabled = False
            cmcViewCntr.Enabled = False
        End If
        If imScheduleExist Then
            If Not imCompleteExist Then
                cmcChgCntr.Caption = "&Change Order"
                cmcViewCntr.Caption = "&View Order"
            End If
            If lmCntrRowSelected >= grdCntr.FixedRows Then
                If Trim$(grdCntr.TextMatrix(lmCntrRowSelected, C_REASONINDEX)) = "Not Scheduled" Then
                    cmcSchedule.Enabled = True
                    ilOk = False
                    If (Not igJobShowing(CONTRACTSJOB)) Then
                        ilOk = True
                    Else
                        If (Contract.cmcUpdate.Enabled = False) Then
                            ilOk = True
                        End If
                    End If
                    If ilOk Then
                        If igWinStatus(CONTRACTSJOB) = 2 Then
                            If tgUrf(0).sReviseCntr <> "N" Then
                                cmcChgCntr.Enabled = True
                                cmcChgCntr.Caption = "&Change Order"
                            End If
                            cmcViewCntr.Enabled = True
                            cmcViewCntr.Caption = "&View Order"
                        ElseIf igWinStatus(CONTRACTSJOB) = 1 Then
                            cmcViewCntr.Enabled = True
                            cmcViewCntr.Caption = "&View Order"
                        End If
                    End If
                Else
                    cmcSchedule.Enabled = False
                End If
            Else
                cmcSchedule.Enabled = True
            End If
        Else
            cmcSchedule.Enabled = False
        End If
    'ElseIf rbcView(1).Value Then
    ElseIf imTabSelected = 1 Then
        If lmLogRowSelected >= grdLog.FixedRows Then
            cmcSpots.Enabled = True
        Else
            cmcSpots.Enabled = False
        End If
        cmcChgCntr.Enabled = False
        cmcViewCntr.Enabled = False
        cmcSchedule.Enabled = False
    'ElseIf rbcView(3).Value Then
    ElseIf imTabSelected = 2 Then
        cmcSpots.Enabled = False
        cmcChgCntr.Caption = "&Change"
        cmcViewCntr.Caption = "&View"
        ilOk = False
        If (Not igJobShowing(CONTRACTSJOB)) Then
            ilOk = True
        Else
            If (Contract.cmcUpdate.Enabled = False) Then
                ilOk = True
            End If
        End If
        If lmPrgmmaticRowSelected >= grdPrgmmatic.FixedRows And ilOk Then
            If igWinStatus(PROPOSALSJOB) = 1 Then
                cmcChgCntr.Enabled = False
                cmcViewCntr.Enabled = True
                cmcSchedule.Enabled = False
            Else
                Select Case Trim$(grdPrgmmatic.TextMatrix(lmPrgmmaticRowSelected, P_CNTRSTATUSINDEX))
                    Case "W", "RW"
                        If igWinStatus(PROPOSALSJOB) = 1 Then
                            cmcChgCntr.Enabled = False
                            cmcViewCntr.Enabled = True
                        ElseIf igWinStatus(PROPOSALSJOB) = 2 Then
                            If tgUrf(0).sPrgmmaticAlert = "I" Then
                                cmcChgCntr.Enabled = True
                            Else
                                cmcChgCntr.Enabled = False
                            End If
                            cmcViewCntr.Enabled = True
                        End If
                        cmcSchedule.Enabled = False
                    Case "D"
                        cmcChgCntr.Enabled = False
                        cmcViewCntr.Enabled = True
                        cmcSchedule.Enabled = False
                    Case "C", "RC"
                        If igWinStatus(PROPOSALSJOB) = 1 Then
                            cmcChgCntr.Enabled = False
                            cmcViewCntr.Enabled = True
                            cmcSchedule.Enabled = False
                        ElseIf igWinStatus(PROPOSALSJOB) = 2 Then
                            If Trim$(grdPrgmmatic.TextMatrix(lmPrgmmaticRowSelected, P_REASONINDEX)) = "Rev Complete" Then
                                'If tgUrf(0).sReviseCntr <> "N" Then
                                If (tgUrf(0).sReviseCntr <> "N") And (tgUrf(0).sPrgmmaticAlert = "I") Then
                                    cmcChgCntr.Enabled = True
                                Else
                                    cmcChgCntr.Enabled = False
                                End If
                            Else
                                If tgUrf(0).sPrgmmaticAlert = "I" Then
                                    cmcChgCntr.Enabled = True
                                Else
                                    cmcChgCntr.Enabled = False
                                End If
                            End If
                            cmcViewCntr.Enabled = True
                            cmcSchedule.Enabled = False
                        End If
                    Case "I", "RI"
                        If igWinStatus(PROPOSALSJOB) = 1 Then
                            cmcChgCntr.Enabled = False
                            cmcViewCntr.Enabled = True
                        ElseIf igWinStatus(PROPOSALSJOB) = 2 Then
                            If tgUrf(0).sPrgmmaticAlert = "I" Then
                                cmcChgCntr.Enabled = True
                            Else
                                cmcChgCntr.Enabled = False
                            End If
                            cmcViewCntr.Enabled = True
                        End If
                        cmcSchedule.Enabled = False
                    Case "G"
                        If igWinStatus(CONTRACTSJOB) = 1 Then
                            cmcChgCntr.Enabled = False
                            cmcViewCntr.Enabled = True
                            cmcSchedule.Enabled = False
                        ElseIf igWinStatus(CONTRACTSJOB) = 2 Then
                            'If tgUrf(0).sReviseCntr <> "N" Then
                            If (tgUrf(0).sReviseCntr <> "N") And (tgUrf(0).sPrgmmaticAlert = "I") Then
                                cmcChgCntr.Enabled = True
                            Else
                                cmcChgCntr.Enabled = False
                            End If
                            cmcViewCntr.Enabled = True
                            cmcSchedule.Enabled = True
                        End If
                    Case "N"
                        If igWinStatus(CONTRACTSJOB) = 1 Then
                            cmcChgCntr.Enabled = False
                            cmcViewCntr.Enabled = True
                            cmcSchedule.Enabled = False
                        ElseIf igWinStatus(CONTRACTSJOB) = 2 Then
                            'If tgUrf(0).sReviseCntr <> "N" Then
                            If (tgUrf(0).sReviseCntr <> "N") And (tgUrf(0).sPrgmmaticAlert = "I") Then
                                cmcChgCntr.Enabled = True
                            Else
                                cmcChgCntr.Enabled = False
                            End If
                            cmcViewCntr.Enabled = True
                            cmcSchedule.Enabled = True
                        End If
                    Case "H"
                        If igWinStatus(CONTRACTSJOB) = 1 Then
                            cmcChgCntr.Enabled = False
                            cmcViewCntr.Enabled = True
                            cmcSchedule.Enabled = False
                        ElseIf igWinStatus(CONTRACTSJOB) = 2 Then
                            If tgUrf(0).sPrgmmaticAlert = "I" Then
                                cmcChgCntr.Enabled = True
                            Else
                                cmcChgCntr.Enabled = False
                            End If
                            cmcViewCntr.Enabled = True
                            cmcSchedule.Enabled = False
                        End If
                    Case "O"
                        If igWinStatus(CONTRACTSJOB) = 1 Then
                            cmcChgCntr.Enabled = False
                            cmcViewCntr.Enabled = True
                            cmcSchedule.Enabled = False
                        ElseIf igWinStatus(CONTRACTSJOB) = 2 Then
                            If tgUrf(0).sPrgmmaticAlert = "I" Then
                                cmcChgCntr.Enabled = True
                            Else
                                cmcChgCntr.Enabled = False
                            End If
                            cmcViewCntr.Enabled = True
                            cmcSchedule.Enabled = False
                        End If
                    Case Else
                        cmcChgCntr.Enabled = False
                        cmcViewCntr.Enabled = False
                        cmcSchedule.Enabled = False
                End Select
                If tgUrf(0).sPrgmmaticAlert <> "I" Then
                    cmcChgCntr.Enabled = False
                End If
            End If
        Else
            cmcChgCntr.Enabled = False
            cmcViewCntr.Enabled = False
            cmcSchedule.Enabled = False
        End If
    ElseIf imTabSelected = 4 Then
        cmcChgCntr.Caption = "&Clear"
        cmcChgCntr.Enabled = True
        cmcViewCntr.Enabled = False
        cmcSchedule.Enabled = False
        cmcSpots.Enabled = False
    ElseIf imTabSelected = 5 Then
        cmcChgCntr.Enabled = False
        cmcViewCntr.Enabled = False
        cmcSchedule.Enabled = False
        cmcSpots.Enabled = False
    Else
        cmcChgCntr.Enabled = False
        cmcViewCntr.Enabled = False
        cmcSchedule.Enabled = False
        cmcSpots.Enabled = False
    End If
    Exit Sub
cmcEraseErr: 'VBC NR
    ilRet = 1
    Resume Next
End Sub

Private Sub mSchedule()

    Dim ilSchSelCntr As Integer
    Dim ilRet As Integer
    Dim slStr As String

    ilSchSelCntr = False
    If imTabSelected = 0 Then
        If lmCntrRowSelected >= grdCntr.FixedRows Then
            If grdCntr.TextMatrix(lmCntrRowSelected, C_CREATEDATEINDEX) = "" Then
                Exit Sub
            End If
            tmAufSrchKey0.lCode = grdCntr.TextMatrix(lmCntrRowSelected, C_AUFCODEINDEX)
            ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                If tmAuf.sType = "C" Then
                    tmChfSrchKey.lCode = tmAuf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If (tmChf.sSchStatus = "A") Or (tmChf.sSchStatus = "N") Then
                            ilSchSelCntr = True
                        Else
                            mPopulate
                            Exit Sub
                        End If
                    Else
                        mPopulate
                        Exit Sub
                    End If
                Else
                    mPopulate
                    Exit Sub
                End If
            Else
                mPopulate
                Exit Sub
            End If
        End If
    ElseIf imTabSelected = 2 Then
        If lmPrgmmaticRowSelected >= grdPrgmmatic.FixedRows Then
            If grdPrgmmatic.TextMatrix(lmPrgmmaticRowSelected, P_CREATEDATEINDEX) = "" Then
                Exit Sub
            End If
            tmAufSrchKey0.lCode = grdPrgmmatic.TextMatrix(lmPrgmmaticRowSelected, P_AUFCODEINDEX)
            ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                If tmAuf.sType = "C" Then
                    tmChfSrchKey.lCode = tmAuf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If (tmChf.sSchStatus = "A") Or (tmChf.sSchStatus = "N") Then
                            ilSchSelCntr = True
                        Else
                            mPopulate
                            Exit Sub
                        End If
                    Else
                        mPopulate
                        Exit Sub
                    End If
                Else
                    mPopulate
                    Exit Sub
                End If
            Else
                mPopulate
                Exit Sub
            End If
        End If
    End If
    If ilSchSelCntr Then
        If igTestSystem Then
            slStr = "Traffic^Test\" & sgUserName & "\" & "#" & Trim$(str$(tmChf.lCode))
        Else
            slStr = "Traffic^Prod\" & sgUserName & "\" & "#" & Trim$(str$(tmChf.lCode))
        End If
    Else
        If igTestSystem Then
            slStr = "Traffic^Test\" & sgUserName & "\" & "Hold"
        Else
            slStr = "Traffic^Prod\" & sgUserName & "\" & "Hold"
        End If
    End If
    sgCommandStr = slStr
    CntrSch.Show vbModal
    slStr = sgDoneMsg
    mPopulate

    Exit Sub
mScheduleErr: 'VBC NR
    ilRet = 1
    Resume Next
End Sub

Private Sub mGridCntrLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdCntr.Rows - 1 Step 1
        grdCntr.RowHeight(ilRow) = fgBoxGridH + 15
    Next ilRow
    For ilCol = 0 To grdCntr.Cols - 1 Step 1
        grdCntr.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridCntrColumns()

    grdCntr.Row = grdCntr.FixedRows - 1
    grdCntr.Col = C_CREATEDATEINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    grdCntr.CellBackColor = LIGHTBLUE
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Create Date"
    grdCntr.Col = C_CREATETIMEINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    grdCntr.CellBackColor = LIGHTBLUE
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Create Time"
    grdCntr.Col = C_CNTRNOINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    grdCntr.CellBackColor = LIGHTBLUE
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Contract #"
    grdCntr.Col = C_ADVTNAMEINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    grdCntr.CellBackColor = LIGHTBLUE
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Advertiser"
    grdCntr.Col = C_PRODUCTINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    grdCntr.CellBackColor = LIGHTBLUE
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Product"
    grdCntr.Col = C_CNTRSTARTDATEINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    grdCntr.CellBackColor = LIGHTBLUE
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Start Date"
    grdCntr.Col = C_SALESOFFICEINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    grdCntr.CellBackColor = LIGHTBLUE
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Sales Office"
    grdCntr.Col = C_REASONINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    grdCntr.CellBackColor = LIGHTBLUE
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Status"
    grdCntr.Col = C_AUFCODEINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    grdCntr.CellBackColor = LIGHTBLUE
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Auf Code"
    grdCntr.Col = C_SORTINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    grdCntr.CellBackColor = LIGHTBLUE
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Sort"

End Sub

Private Sub mGridCntrColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdCntr.ColWidth(C_AUFCODEINDEX) = 0
    grdCntr.ColWidth(C_SORTINDEX) = 0
    grdCntr.ColWidth(C_CREATEDATEINDEX) = 0.092 * grdCntr.Width
    grdCntr.ColWidth(C_CREATETIMEINDEX) = 0.1 * grdCntr.Width
    grdCntr.ColWidth(C_ADVTNAMEINDEX) = 0.2 * grdCntr.Width
    grdCntr.ColWidth(C_PRODUCTINDEX) = 0.2 * grdCntr.Width
    grdCntr.ColWidth(C_CNTRNOINDEX) = 0.08 * grdCntr.Width
    grdCntr.ColWidth(C_CNTRSTARTDATEINDEX) = 0.08 * grdCntr.Width
    grdCntr.ColWidth(C_SALESOFFICEINDEX) = 0.1 * grdCntr.Width
    grdCntr.ColWidth(C_REASONINDEX) = 0.1 * grdCntr.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdCntr.Width
    For ilCol = 0 To grdCntr.Cols - 1 Step 1
        llWidth = llWidth + grdCntr.ColWidth(ilCol)
        If (grdCntr.ColWidth(ilCol) > 15) And (grdCntr.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdCntr.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdCntr.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdCntr.Width
            For ilCol = 0 To grdCntr.Cols - 1 Step 1
                If (grdCntr.ColWidth(ilCol) > 15) And (grdCntr.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdCntr.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdCntr.FixedCols To grdCntr.Cols - 1 Step 1
                If grdCntr.ColWidth(ilCol) > 15 Then
                    ilColInc = grdCntr.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdCntr.ColWidth(ilCol) = grdCntr.ColWidth(ilCol) + 15
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
'*  flTextHeight                  ilLoop                        ilRow                     *
'*  ilCol                                                                                 *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
    'flTextHeight = pbcDates.TextHeight("1") - 35

'    mGridCntrLayout
'    mGridCntrColumnWidths
'    mGridCntrColumns
'    grdCntr.Move 180, frcView.Top + frcView.Height + 120
'    grdCntr.Height = grdCntr.RowPos(0) + 14 * grdCntr.RowHeight(0) + fgPanelAdj - 15
'
'    mGridLogLayout
'    mGridLogColumnWidths
'    mGridLogColumns
'    grdLog.Move grdCntr.Left, grdCntr.Top, grdCntr.Width, grdCntr.Height
'
'    mGridAffExportLayout
'    mGridAffExportColumnWidths
'    mGridAffExportColumns
'    grdAffExport.Move grdCntr.Left, grdCntr.Top, grdCntr.Width, grdCntr.Height
'
'    mGridRNMsgLayout
'    mGridRNMsgColumnWidths
'    mGridRNMsgColumns
'    grdRNMsg.Move grdCntr.Left, grdCntr.Top, grdCntr.Width, grdCntr.Height

End Sub

Private Sub mGridLogLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdLog.Rows - 1 Step 1
        grdLog.RowHeight(ilRow) = fgBoxGridH + 15
    Next ilRow
    For ilCol = 0 To grdLog.Cols - 1 Step 1
        grdLog.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridLogColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdLog.Row = grdLog.FixedRows - 1
    grdLog.Col = L_CREATEDATEINDEX
    grdLog.CellFontBold = False
    grdLog.CellFontName = "Arial"
    grdLog.CellFontSize = 6.75
    grdLog.CellForeColor = vbBlue
    grdLog.CellBackColor = LIGHTBLUE
    grdLog.TextMatrix(grdLog.Row, grdLog.Col) = "Create Date"
    grdLog.Col = L_CREATETIMEINDEX
    grdLog.CellFontBold = False
    grdLog.CellFontName = "Arial"
    grdLog.CellFontSize = 6.75
    grdLog.CellForeColor = vbBlue
    grdLog.CellBackColor = LIGHTBLUE
    grdLog.TextMatrix(grdLog.Row, grdLog.Col) = "Create Time"
    grdLog.Col = L_VEHICLEINDEX
    grdLog.CellFontBold = False
    grdLog.CellFontName = "Arial"
    grdLog.CellFontSize = 6.75
    grdLog.CellForeColor = vbBlue
    grdLog.CellBackColor = LIGHTBLUE
    grdLog.TextMatrix(grdLog.Row, grdLog.Col) = "Vehicle Name"
    grdLog.Col = L_LOGDATEINDEX
    grdLog.CellFontBold = False
    grdLog.CellFontName = "Arial"
    grdLog.CellFontSize = 6.75
    grdLog.CellForeColor = vbBlue
    grdLog.CellBackColor = LIGHTBLUE
    grdLog.TextMatrix(grdLog.Row, grdLog.Col) = "Log Date"
    grdLog.Col = L_REASONINDEX
    grdLog.CellFontBold = False
    grdLog.CellFontName = "Arial"
    grdLog.CellFontSize = 6.75
    grdLog.CellForeColor = vbBlue
    grdLog.CellBackColor = LIGHTBLUE
    grdLog.TextMatrix(grdLog.Row, grdLog.Col) = "Reason"
    grdLog.Col = L_AUFCODEINDEX
    grdLog.CellFontBold = False
    grdLog.CellFontName = "Arial"
    grdLog.CellFontSize = 6.75
    grdLog.CellForeColor = vbBlue
    grdLog.CellBackColor = LIGHTBLUE
    grdLog.TextMatrix(grdLog.Row, grdLog.Col) = "Auf Code"
    grdLog.Col = L_SORTINDEX
    grdLog.CellFontBold = False
    grdLog.CellFontName = "Arial"
    grdLog.CellFontSize = 6.75
    grdLog.CellForeColor = vbBlue
    grdLog.CellBackColor = LIGHTBLUE
    grdLog.TextMatrix(grdLog.Row, grdLog.Col) = "Sort"

End Sub

Private Sub mGridLogColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdLog.ColWidth(L_AUFCODEINDEX) = 0
    grdLog.ColWidth(L_SORTINDEX) = 0
    grdLog.ColWidth(L_CREATEDATEINDEX) = 0.092 * grdLog.Width
    grdLog.ColWidth(L_CREATETIMEINDEX) = 0.1 * grdLog.Width
    grdLog.ColWidth(L_VEHICLEINDEX) = 0.4 * grdLog.Width
    grdLog.ColWidth(L_LOGDATEINDEX) = 0.092 * grdLog.Width
    grdLog.ColWidth(L_REASONINDEX) = 0.15 * grdLog.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdLog.Width
    For ilCol = 0 To grdLog.Cols - 1 Step 1
        llWidth = llWidth + grdLog.ColWidth(ilCol)
        If (grdLog.ColWidth(ilCol) > 15) And (grdLog.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdLog.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdLog.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdLog.Width
            For ilCol = 0 To grdLog.Cols - 1 Step 1
                If (grdLog.ColWidth(ilCol) > 15) And (grdLog.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdLog.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdLog.FixedCols To grdLog.Cols - 1 Step 1
                If grdLog.ColWidth(ilCol) > 15 Then
                    ilColInc = grdLog.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdLog.ColWidth(ilCol) = grdLog.ColWidth(ilCol) + 15
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


Private Sub mGridAffExportLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdAffExport.Rows - 1 Step 1
        grdAffExport.RowHeight(ilRow) = fgBoxGridH + 15
    Next ilRow
    For ilCol = 0 To grdAffExport.Cols - 1 Step 1
        grdAffExport.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridAffExportColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdAffExport.Row = grdAffExport.FixedRows - 1
    grdAffExport.Col = A_CREATEDATEINDEX
    grdAffExport.CellFontBold = False
    grdAffExport.CellFontName = "Arial"
    grdAffExport.CellFontSize = 6.75
    grdAffExport.CellForeColor = vbBlue
    grdAffExport.CellBackColor = LIGHTBLUE
    grdAffExport.TextMatrix(grdAffExport.Row, grdAffExport.Col) = "Create Date"
    grdAffExport.Col = A_CREATETIMEINDEX
    grdAffExport.CellFontBold = False
    grdAffExport.CellFontName = "Arial"
    grdAffExport.CellFontSize = 6.75
    grdAffExport.CellForeColor = vbBlue
    grdAffExport.CellBackColor = LIGHTBLUE
    grdAffExport.TextMatrix(grdAffExport.Row, grdAffExport.Col) = "Create Time"
    grdAffExport.Col = A_VEHICLEINDEX
    grdAffExport.CellFontBold = False
    grdAffExport.CellFontName = "Arial"
    grdAffExport.CellFontSize = 6.75
    grdAffExport.CellForeColor = vbBlue
    grdAffExport.CellBackColor = LIGHTBLUE
    grdAffExport.TextMatrix(grdAffExport.Row, grdAffExport.Col) = "Vehicle Name"
    grdAffExport.Col = A_LOGDATEINDEX
    grdAffExport.CellFontBold = False
    grdAffExport.CellFontName = "Arial"
    grdAffExport.CellFontSize = 6.75
    grdAffExport.CellForeColor = vbBlue
    grdAffExport.CellBackColor = LIGHTBLUE
    grdAffExport.TextMatrix(grdAffExport.Row, grdAffExport.Col) = "Log Date"
    grdAffExport.Col = A_REASONINDEX
    grdAffExport.CellFontBold = False
    grdAffExport.CellFontName = "Arial"
    grdAffExport.CellFontSize = 6.75
    grdAffExport.CellForeColor = vbBlue
    grdAffExport.CellBackColor = LIGHTBLUE
    grdAffExport.TextMatrix(grdAffExport.Row, grdAffExport.Col) = "Reason"
    grdAffExport.Col = A_AUFCODEINDEX
    grdAffExport.CellFontBold = False
    grdAffExport.CellFontName = "Arial"
    grdAffExport.CellFontSize = 6.75
    grdAffExport.CellForeColor = vbBlue
    grdAffExport.CellBackColor = LIGHTBLUE
    grdAffExport.TextMatrix(grdAffExport.Row, grdAffExport.Col) = "Auf Code"
    grdAffExport.Col = A_SORTINDEX
    grdAffExport.CellFontBold = False
    grdAffExport.CellFontName = "Arial"
    grdAffExport.CellFontSize = 6.75
    grdAffExport.CellForeColor = vbBlue
    grdAffExport.CellBackColor = LIGHTBLUE
    grdAffExport.TextMatrix(grdAffExport.Row, grdAffExport.Col) = "Sort"

End Sub

Private Sub mGridAffExportColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdAffExport.ColWidth(A_AUFCODEINDEX) = 0
    grdAffExport.ColWidth(A_SORTINDEX) = 0
    grdAffExport.ColWidth(A_CREATEDATEINDEX) = 0.092 * grdAffExport.Width
    grdAffExport.ColWidth(A_CREATETIMEINDEX) = 0.1 * grdAffExport.Width
    grdAffExport.ColWidth(A_VEHICLEINDEX) = 0.4 * grdAffExport.Width
    grdAffExport.ColWidth(A_LOGDATEINDEX) = 0.092 * grdAffExport.Width
    grdAffExport.ColWidth(A_REASONINDEX) = 0.15 * grdAffExport.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdAffExport.Width
    For ilCol = 0 To grdAffExport.Cols - 1 Step 1
        llWidth = llWidth + grdAffExport.ColWidth(ilCol)
        If (grdAffExport.ColWidth(ilCol) > 15) And (grdAffExport.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdAffExport.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdAffExport.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdAffExport.Width
            For ilCol = 0 To grdAffExport.Cols - 1 Step 1
                If (grdAffExport.ColWidth(ilCol) > 15) And (grdAffExport.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdAffExport.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdAffExport.FixedCols To grdAffExport.Cols - 1 Step 1
                If grdAffExport.ColWidth(ilCol) > 15 Then
                    ilColInc = grdAffExport.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdAffExport.ColWidth(ilCol) = grdAffExport.ColWidth(ilCol) + 15
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

Private Sub mCntrSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdCntr.FixedRows To grdCntr.Rows - 1 Step 1
        slStr = Trim$(grdCntr.TextMatrix(llRow, C_CREATEDATEINDEX))
        If slStr <> "" Then
            If ilCol = C_CREATEDATEINDEX Then
                slSort = Trim$(str$(gDateValue(grdCntr.TextMatrix(llRow, C_CREATEDATEINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = C_CREATETIMEINDEX) Then
                slSort = Trim$(str$(gTimeToLong(grdCntr.TextMatrix(llRow, C_CREATETIMEINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = C_CNTRNOINDEX) Then
                slSort = Trim$(grdCntr.TextMatrix(llRow, C_CNTRNOINDEX))
                Do While Len(slSort) < 8
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = C_CNTRSTARTDATEINDEX) Then
                slSort = Trim$(str$(gDateValue(grdCntr.TextMatrix(llRow, C_CNTRSTARTDATEINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdCntr.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdCntr.TextMatrix(llRow, C_SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastCntrColSorted) Or ((ilCol = imLastCntrColSorted) And (imLastCntrSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdCntr.TextMatrix(llRow, C_SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdCntr.TextMatrix(llRow, C_SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastCntrColSorted Then
        imLastCntrColSorted = C_SORTINDEX
    Else
        imLastCntrColSorted = -1
        imLastCntrSort = -1
    End If
    gGrid_SortByCol grdCntr, C_CREATEDATEINDEX, C_SORTINDEX, imLastCntrColSorted, imLastCntrSort
    imLastCntrColSorted = ilCol
End Sub

Private Sub mLogSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdLog.FixedRows To grdLog.Rows - 1 Step 1
        slStr = Trim$(grdLog.TextMatrix(llRow, L_CREATEDATEINDEX))
        If slStr <> "" Then
            If ilCol = L_CREATEDATEINDEX Then
                slSort = Trim$(str$(gDateValue(grdLog.TextMatrix(llRow, L_CREATEDATEINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = L_CREATETIMEINDEX) Then
                slSort = Trim$(str$(gTimeToLong(grdLog.TextMatrix(llRow, L_CREATETIMEINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = L_LOGDATEINDEX) Then
                slSort = Trim$(str$(gTimeToLong(grdLog.TextMatrix(llRow, L_LOGDATEINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdLog.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdLog.TextMatrix(llRow, L_SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastLogColSorted) Or ((ilCol = imLastLogColSorted) And (imLastLogSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdLog.TextMatrix(llRow, L_SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdLog.TextMatrix(llRow, L_SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastLogColSorted Then
        imLastLogColSorted = L_SORTINDEX
    Else
        imLastLogColSorted = -1
        imLastLogSort = -1
    End If
    gGrid_SortByCol grdLog, L_CREATEDATEINDEX, L_SORTINDEX, imLastLogColSorted, imLastLogSort
    imLastLogColSorted = ilCol
End Sub

Private Sub mPrgmmaticSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdPrgmmatic.FixedRows To grdPrgmmatic.Rows - 1 Step 1
        slStr = Trim$(grdPrgmmatic.TextMatrix(llRow, P_CREATEDATEINDEX))
        If slStr <> "" Then
            If ilCol = P_CREATEDATEINDEX Then
                slSort = Trim$(str$(gDateValue(grdPrgmmatic.TextMatrix(llRow, P_CREATEDATEINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = P_CREATETIMEINDEX) Then
                slSort = Trim$(str$(gTimeToLong(grdPrgmmatic.TextMatrix(llRow, P_CREATETIMEINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = P_CNTRNOINDEX) Then
                slSort = Trim$(grdPrgmmatic.TextMatrix(llRow, P_CNTRNOINDEX))
                Do While Len(slSort) < 8
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = P_CNTRSTARTDATEINDEX) Then
                slSort = Trim$(str$(gDateValue(grdPrgmmatic.TextMatrix(llRow, P_CNTRSTARTDATEINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdPrgmmatic.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdPrgmmatic.TextMatrix(llRow, P_SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastPrgmmaticColSorted) Or ((ilCol = imLastPrgmmaticColSorted) And (imLastPrgmmaticSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdPrgmmatic.TextMatrix(llRow, P_SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdPrgmmatic.TextMatrix(llRow, P_SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastPrgmmaticColSorted Then
        imLastPrgmmaticColSorted = P_SORTINDEX
    Else
        imLastPrgmmaticColSorted = -1
        imLastPrgmmaticSort = -1
    End If
    gGrid_SortByCol grdPrgmmatic, P_CREATEDATEINDEX, P_SORTINDEX, imLastPrgmmaticColSorted, imLastPrgmmaticSort
    imLastPrgmmaticColSorted = ilCol
End Sub
Private Sub mAffSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdAffExport.FixedRows To grdAffExport.Rows - 1 Step 1
        slStr = Trim$(grdAffExport.TextMatrix(llRow, A_CREATEDATEINDEX))
        If slStr <> "" Then
            If ilCol = A_CREATEDATEINDEX Then
                slSort = Trim$(str$(gDateValue(grdAffExport.TextMatrix(llRow, A_CREATEDATEINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = A_CREATETIMEINDEX) Then
                slSort = Trim$(str$(gTimeToLong(grdAffExport.TextMatrix(llRow, A_CREATETIMEINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = A_LOGDATEINDEX) Then
                slSort = Trim$(str$(gTimeToLong(grdAffExport.TextMatrix(llRow, A_LOGDATEINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdAffExport.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdAffExport.TextMatrix(llRow, A_SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastAffColSorted) Or ((ilCol = imLastAffColSorted) And (imLastAffSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdAffExport.TextMatrix(llRow, A_SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdAffExport.TextMatrix(llRow, A_SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastAffColSorted Then
        imLastAffColSorted = A_SORTINDEX
    Else
        imLastAffColSorted = -1
        imLastAffSort = -1
    End If
    gGrid_SortByCol grdAffExport, A_CREATEDATEINDEX, A_SORTINDEX, imLastAffColSorted, imLastAffSort
    imLastAffColSorted = ilCol
End Sub

Private Sub mGridRNMsgLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdRNMsg.Rows - 1 Step 1
        grdRNMsg.RowHeight(ilRow) = fgBoxGridH + 15
    Next ilRow
    For ilCol = 0 To grdRNMsg.Cols - 1 Step 1
        grdRNMsg.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridRNMsgColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdRNMsg.Row = grdRNMsg.FixedRows - 1
    grdRNMsg.Col = RN_CREATEDATEINDEX
    grdRNMsg.CellFontBold = False
    grdRNMsg.CellFontName = "Arial"
    grdRNMsg.CellFontSize = 6.75
    grdRNMsg.CellForeColor = vbBlue
    grdRNMsg.CellBackColor = LIGHTBLUE
    grdRNMsg.TextMatrix(grdRNMsg.Row, grdRNMsg.Col) = "Create Date"
    grdRNMsg.Col = RN_CREATETIMEINDEX
    grdRNMsg.CellFontBold = False
    grdRNMsg.CellFontName = "Arial"
    grdRNMsg.CellFontSize = 6.75
    grdRNMsg.CellForeColor = vbBlue
    grdRNMsg.CellBackColor = LIGHTBLUE
    grdRNMsg.TextMatrix(grdRNMsg.Row, grdRNMsg.Col) = "Create Time"
    grdRNMsg.Col = RN_TYPEMSGINDEX
    grdRNMsg.CellFontBold = False
    grdRNMsg.CellFontName = "Arial"
    grdRNMsg.CellFontSize = 6.75
    grdRNMsg.CellForeColor = vbBlue
    grdRNMsg.CellBackColor = LIGHTBLUE
    grdRNMsg.TextMatrix(grdRNMsg.Row, grdRNMsg.Col) = "Type"
    grdRNMsg.Col = RN_STATUSINDEX
    grdRNMsg.CellFontBold = False
    grdRNMsg.CellFontName = "Arial"
    grdRNMsg.CellFontSize = 6.75
    grdRNMsg.CellForeColor = vbBlue
    grdRNMsg.CellBackColor = LIGHTBLUE
    grdRNMsg.TextMatrix(grdRNMsg.Row, grdRNMsg.Col) = "Status"
    grdRNMsg.Col = RN_FROMIDINDEX
    grdRNMsg.CellFontBold = False
    grdRNMsg.CellFontName = "Arial"
    grdRNMsg.CellFontSize = 6.75
    grdRNMsg.CellForeColor = vbBlue
    grdRNMsg.CellBackColor = LIGHTBLUE
    grdRNMsg.TextMatrix(grdRNMsg.Row, grdRNMsg.Col) = "From"
    grdRNMsg.Col = RN_CNTRNOINDEX
    grdRNMsg.CellFontBold = False
    grdRNMsg.CellFontName = "Arial"
    grdRNMsg.CellFontSize = 6.75
    grdRNMsg.CellForeColor = vbBlue
    grdRNMsg.CellBackColor = LIGHTBLUE
    grdRNMsg.TextMatrix(grdRNMsg.Row, grdRNMsg.Col) = "Contract #"
    grdRNMsg.Col = RN_ADVTNAMEINDEX
    grdRNMsg.CellFontBold = False
    grdRNMsg.CellFontName = "Arial"
    grdRNMsg.CellFontSize = 6.75
    grdRNMsg.CellForeColor = vbBlue
    grdRNMsg.CellBackColor = LIGHTBLUE
    grdRNMsg.TextMatrix(grdRNMsg.Row, grdRNMsg.Col) = "Advertiser"
    grdRNMsg.Col = RN_VEHICLEINDEX
    grdRNMsg.CellFontBold = False
    grdRNMsg.CellFontName = "Arial"
    grdRNMsg.CellFontSize = 6.75
    grdRNMsg.CellForeColor = vbBlue
    grdRNMsg.CellBackColor = LIGHTBLUE
    grdRNMsg.TextMatrix(grdRNMsg.Row, grdRNMsg.Col) = "Vehicle"
    grdRNMsg.Col = RN_MESSAGEINDEX
    grdRNMsg.CellFontBold = False
    grdRNMsg.CellFontName = "Arial"
    grdRNMsg.CellFontSize = 6.75
    grdRNMsg.CellForeColor = vbBlue
    grdRNMsg.CellBackColor = LIGHTBLUE
    grdRNMsg.TextMatrix(grdRNMsg.Row, grdRNMsg.Col) = "Message"
    grdRNMsg.Col = RN_AUFCODEINDEX
    grdRNMsg.CellFontBold = False
    grdRNMsg.CellFontName = "Arial"
    grdRNMsg.CellFontSize = 6.75
    grdRNMsg.CellForeColor = vbBlue
    grdRNMsg.CellBackColor = LIGHTBLUE
    grdRNMsg.TextMatrix(grdRNMsg.Row, grdRNMsg.Col) = "Auf Code"
    grdRNMsg.Col = RN_SORTINDEX
    grdRNMsg.CellFontBold = False
    grdRNMsg.CellFontName = "Arial"
    grdRNMsg.CellFontSize = 6.75
    grdRNMsg.CellForeColor = vbBlue
    grdRNMsg.CellBackColor = LIGHTBLUE
    grdRNMsg.TextMatrix(grdRNMsg.Row, grdRNMsg.Col) = "Sort"

End Sub

Private Sub mGridRNMsgColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdRNMsg.ColWidth(RN_AUFCODEINDEX) = 0
    grdRNMsg.ColWidth(RN_SORTINDEX) = 0
    grdRNMsg.ColWidth(RN_CREATEDATEINDEX) = 0.092 * grdRNMsg.Width
    grdRNMsg.ColWidth(RN_CREATETIMEINDEX) = 0.1 * grdRNMsg.Width
    grdRNMsg.ColWidth(RN_ADVTNAMEINDEX) = 0.12 * grdRNMsg.Width
    grdRNMsg.ColWidth(RN_VEHICLEINDEX) = 0.12 * grdRNMsg.Width
    grdRNMsg.ColWidth(RN_FROMIDINDEX) = 0.05 * grdRNMsg.Width
    grdRNMsg.ColWidth(RN_CNTRNOINDEX) = 0.08 * grdRNMsg.Width
    grdRNMsg.ColWidth(RN_TYPEMSGINDEX) = 0.11 * grdRNMsg.Width
    grdRNMsg.ColWidth(RN_STATUSINDEX) = 0.05 * grdRNMsg.Width
    grdRNMsg.ColWidth(RN_MESSAGEINDEX) = 0.36 * grdRNMsg.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdRNMsg.Width
    For ilCol = 0 To grdRNMsg.Cols - 1 Step 1
        llWidth = llWidth + grdRNMsg.ColWidth(ilCol)
        If (grdRNMsg.ColWidth(ilCol) > 15) And (grdRNMsg.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdRNMsg.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdRNMsg.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdRNMsg.Width
            For ilCol = 0 To grdRNMsg.Cols - 1 Step 1
                If (grdRNMsg.ColWidth(ilCol) > 15) And (grdRNMsg.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdRNMsg.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdRNMsg.FixedCols To grdRNMsg.Cols - 1 Step 1
                If grdRNMsg.ColWidth(ilCol) > 15 Then
                    ilColInc = grdRNMsg.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdRNMsg.ColWidth(ilCol) = grdRNMsg.ColWidth(ilCol) + 15
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
    'grdRNMsg.WordWrap = True
    grdRNMsg.ColWordWrapOption(RN_MESSAGEINDEX) = True
End Sub

Private Sub mRNMsgSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdRNMsg.FixedRows To grdRNMsg.Rows - 1 Step 1
        slStr = Trim$(grdRNMsg.TextMatrix(llRow, RN_CREATEDATEINDEX))
        If slStr <> "" Then
            If ilCol = RN_CREATEDATEINDEX Then
                slSort = Trim$(str$(gDateValue(grdRNMsg.TextMatrix(llRow, RN_CREATEDATEINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = RN_CREATETIMEINDEX) Then
                slSort = Trim$(str$(gTimeToLong(grdRNMsg.TextMatrix(llRow, RN_CREATETIMEINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = RN_CNTRNOINDEX) Then
                slSort = Trim$(grdRNMsg.TextMatrix(llRow, RN_CNTRNOINDEX))
                Do While Len(slSort) < 8
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdRNMsg.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdRNMsg.TextMatrix(llRow, RN_SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastRNMsgColSorted) Or ((ilCol = imLastRNMsgColSorted) And (imLastRNMsgSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdRNMsg.TextMatrix(llRow, RN_SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdRNMsg.TextMatrix(llRow, RN_SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastRNMsgColSorted Then
        imLastRNMsgColSorted = RN_SORTINDEX
    Else
        imLastRNMsgColSorted = -1
        imLastRNMsgSort = -1
    End If
    gGrid_SortByCol grdRNMsg, RN_CREATEDATEINDEX, RN_SORTINDEX, imLastRNMsgColSorted, imLastRNMsgSort
    imLastRNMsgColSorted = ilCol
End Sub

Private Sub mSetControls()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRow                         ilCol                                                   *
'******************************************************************************************

    Dim ilGap As Integer

    ilGap = cmcChgCntr.Left - (cmcDone.Left + cmcDone.Width)
    cmcDone.Top = Me.Height - cmcDone.Height - 120
    cmcChgCntr.Top = cmcDone.Top
    cmcViewCntr.Top = cmcDone.Top
    cmcSchedule.Top = cmcDone.Top
    cmcSpots.Top = cmcDone.Top
    cmcDone.Left = AlertVw.Width / 2 - (4 * ilGap + cmcDone.Width + cmcChgCntr.Width + cmcViewCntr.Width + cmcSchedule.Width + cmcSpots.Width) / 2
    cmcChgCntr.Left = cmcDone.Left + cmcDone.Width + ilGap
    cmcViewCntr.Left = cmcChgCntr.Left + cmcChgCntr.Width + ilGap
    cmcSchedule.Left = cmcViewCntr.Left + cmcViewCntr.Width + ilGap
    cmcSpots.Left = cmcSchedule.Left + cmcSchedule.Width + ilGap
    frcView.Left = AlertVw.Width / 2 - frcView.Width / 2
    frcView.Top = 210
    tbcAlert.Move 120, lbcScreen.Top + lbcScreen.Height + 60, AlertVw.Width - 2 * 120, cmcDone.Top - lbcScreen.Top - lbcScreen.Height - 120
    grdCntr.Move tbcAlert.Left + 120, tbcAlert.Top + 355, tbcAlert.Width - 2 * 120, tbcAlert.Height - 355 - 2 * 120
    grdCntr.RowHeight(0) = fgBoxGridH + 15
    mGridCntrColumnWidths
    mGridCntrColumns
    gGrid_IntegralHeight grdCntr, fgBoxGridH + 15
    gGrid_FillWithRows grdCntr, fgBoxGridH + 15
    mGridCntrLayout
    grdCntr.Height = grdCntr.Height + 45
    grdCntr.Top = (tbcAlert.Height - 355) / 2 - grdCntr.Height / 2 + tbcAlert.Top + 355
    grdLog.Move grdCntr.Left, grdCntr.Top, grdCntr.Width, grdCntr.Height
    grdLog.RowHeight(0) = fgBoxGridH + 15
    mGridLogColumnWidths
    mGridLogColumns
    gGrid_IntegralHeight grdLog, fgBoxGridH + 15
    gGrid_FillWithRows grdLog, fgBoxGridH + 15
    mGridLogLayout
    grdLog.Height = grdLog.Height + 45

    grdAffExport.Move grdCntr.Left, grdCntr.Top, grdCntr.Width, grdCntr.Height
    grdAffExport.RowHeight(0) = fgBoxGridH + 15
    mGridAffExportColumnWidths
    mGridAffExportColumns
    gGrid_IntegralHeight grdAffExport, fgBoxGridH + 15
    gGrid_FillWithRows grdAffExport, fgBoxGridH + 15
    mGridAffExportLayout
    grdAffExport.Height = grdAffExport.Height + 45

    'grdRNMsg.Move grdCntr.Left, grdCntr.Top, grdCntr.Width, grdCntr.Height
    grdRNMsg.Move grdCntr.Left, grdCntr.Top, grdCntr.Width, grdCntr.Height - cmcDone.Height
    grdRNMsg.RowHeight(0) = fgBoxGridH + 15
    mGridRNMsgColumnWidths
    mGridRNMsgColumns
    gGrid_IntegralHeight grdRNMsg, fgBoxGridH + 15
    gGrid_FillWithRows grdRNMsg, fgBoxGridH + 15
    mGridRNMsgLayout
    grdRNMsg.Height = grdRNMsg.Height + 45

    grdPrgmmatic.Move grdCntr.Left, grdCntr.Top, grdCntr.Width, grdCntr.Height - cmcDone.Height
    grdPrgmmatic.RowHeight(0) = fgBoxGridH + 15
    mGridPrgmmaticColumnWidths
    mGridPrgmmaticColumns
    gGrid_IntegralHeight grdPrgmmatic, fgBoxGridH + 15
    gGrid_FillWithRows grdPrgmmatic, fgBoxGridH + 15
    mGridPrgmmaticLayout
    grdPrgmmatic.Height = grdPrgmmatic.Height + 45

    grdPool.Move grdCntr.Left, grdCntr.Top, grdCntr.Width, grdCntr.Height - cmcDone.Height
    grdPool.RowHeight(0) = fgBoxGridH + 15
    mGridPoolColumnWidths
    mGridPoolColumns
    gGrid_IntegralHeight grdPool, fgBoxGridH + 15
    gGrid_FillWithRows grdPool, fgBoxGridH + 15
    mGridPoolLayout
    grdPool.Height = grdPool.Height + 45

End Sub


Private Sub mSpots()
    Dim ilRet As Integer
    If lmLogRowSelected < grdLog.FixedRows Then
        Exit Sub
    End If
    If igWinStatus(SPOTSJOB) = 0 Then
        Exit Sub
    End If
    If igJobShowing(SPOTSJOB) Then
        ilRet = MsgBox("Unable to view Spots until Current Spot Screen Closed", vbInformation + vbOKOnly, "Information")
        Exit Sub
    End If
    If tgSpf.sSystemType = "R" Then
        sgSpotCallType = "T"
    Else
        sgSpotCallType = "A"
    End If
    igAlertSpotStatus = 1   '2
    sgAlertSpotVehicle = grdLog.TextMatrix(lmLogRowSelected, A_VEHICLEINDEX)
    sgAlertSpotMoDate = grdLog.TextMatrix(lmLogRowSelected, A_LOGDATEINDEX)
    mTerminate
End Sub

Private Sub mPaintLog(llRow, blSelect As Boolean)
    Dim ilCol As Integer
    
    If llRow < grdLog.FixedRows Then
        Exit Sub
    End If
    grdLog.Row = llRow
    For ilCol = 0 To grdLog.Cols - 1 Step 1
        grdLog.Col = ilCol
        If blSelect Then
            grdLog.CellBackColor = GRAY
        Else
            grdLog.CellBackColor = vbWhite
        End If
    Next ilCol
End Sub

Private Sub tbcAlert_Click()

    grdCntr.Visible = False
    grdLog.Visible = False
    grdPrgmmatic.Visible = False
    grdAffExport.Visible = False
    grdRNMsg.Visible = False
    'imTabSelected = tbcAlert.SelectedItem.Index - 1
    If tbcAlert.SelectedItem = "Contracts" Then
        imTabSelected = 0
        grdCntr.Visible = True
    End If
    If tbcAlert.SelectedItem = "Logs" Then
        imTabSelected = 1
        grdLog.Visible = True
    End If
    If tbcAlert.SelectedItem = "Programmatic Buys" Then
        imTabSelected = 2
        grdPrgmmatic.Visible = True
    End If
    'If Index = 2 Then
    If tbcAlert.SelectedItem = "Affiliate Alert" Then
        imTabSelected = 3
        grdAffExport.Visible = True
    End If
    If tbcAlert.SelectedItem = "Rep-Net Messages" Then
        imTabSelected = 4
        If ((Asc(tgSpf.sAutoType2) And RN_REP) <> RN_REP) And ((Asc(tgSpf.sAutoType2) And RN_NET) <> RN_NET) Then
            grdRNMsg.Visible = False
        ElseIf (Trim$(grdRNMsg.TextMatrix(grdRNMsg.FixedRows, RN_CREATEDATEINDEX)) = "") Then
            grdRNMsg.Visible = False
        Else
            grdRNMsg.Visible = True
        End If
    If tbcAlert.SelectedItem = "Pool Alert" Then
        imTabSelected = 5
        grdPool.Visible = True
    End If
    End If
    mSetCommands
End Sub

Private Sub mGridPrgmmaticLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdPrgmmatic.Rows - 1 Step 1
        grdPrgmmatic.RowHeight(ilRow) = fgBoxGridH + 15
    Next ilRow
    For ilCol = 0 To grdPrgmmatic.Cols - 1 Step 1
        grdPrgmmatic.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub
Private Sub mGridPrgmmaticColumns()

    grdPrgmmatic.Row = grdPrgmmatic.FixedRows - 1
    grdPrgmmatic.Col = P_CREATEDATEINDEX
    grdPrgmmatic.CellFontBold = False
    grdPrgmmatic.CellFontName = "Arial"
    grdPrgmmatic.CellFontSize = 6.75
    grdPrgmmatic.CellForeColor = vbBlue
    grdPrgmmatic.CellBackColor = LIGHTBLUE
    grdPrgmmatic.TextMatrix(grdPrgmmatic.Row, grdPrgmmatic.Col) = "Create Date"
    grdPrgmmatic.Col = P_CREATETIMEINDEX
    grdPrgmmatic.CellFontBold = False
    grdPrgmmatic.CellFontName = "Arial"
    grdPrgmmatic.CellFontSize = 6.75
    grdPrgmmatic.CellForeColor = vbBlue
    grdPrgmmatic.CellBackColor = LIGHTBLUE
    grdPrgmmatic.TextMatrix(grdPrgmmatic.Row, grdPrgmmatic.Col) = "Create Time"
    grdPrgmmatic.Col = P_AGENCYINDEX
    grdPrgmmatic.CellFontBold = False
    grdPrgmmatic.CellFontName = "Arial"
    grdPrgmmatic.CellFontSize = 6.75
    grdPrgmmatic.CellForeColor = vbBlue
    grdPrgmmatic.CellBackColor = LIGHTBLUE
    grdPrgmmatic.TextMatrix(grdPrgmmatic.Row, grdPrgmmatic.Col) = "Agency"
    grdPrgmmatic.Col = P_CNTRNOINDEX
    grdPrgmmatic.CellFontBold = False
    grdPrgmmatic.CellFontName = "Arial"
    grdPrgmmatic.CellFontSize = 6.75
    grdPrgmmatic.CellForeColor = vbBlue
    grdPrgmmatic.CellBackColor = LIGHTBLUE
    grdPrgmmatic.TextMatrix(grdPrgmmatic.Row, grdPrgmmatic.Col) = "Contract #"
    grdPrgmmatic.Col = P_ADVTNAMEINDEX
    grdPrgmmatic.CellFontBold = False
    grdPrgmmatic.CellFontName = "Arial"
    grdPrgmmatic.CellFontSize = 6.75
    grdPrgmmatic.CellForeColor = vbBlue
    grdPrgmmatic.CellBackColor = LIGHTBLUE
    grdPrgmmatic.TextMatrix(grdPrgmmatic.Row, grdPrgmmatic.Col) = "Advertiser"
    grdPrgmmatic.Col = P_PRODUCTINDEX
    grdPrgmmatic.CellFontBold = False
    grdPrgmmatic.CellFontName = "Arial"
    grdPrgmmatic.CellFontSize = 6.75
    grdPrgmmatic.CellForeColor = vbBlue
    grdPrgmmatic.CellBackColor = LIGHTBLUE
    grdPrgmmatic.TextMatrix(grdPrgmmatic.Row, grdPrgmmatic.Col) = "Product"
    grdPrgmmatic.Col = P_CNTRSTARTDATEINDEX
    grdPrgmmatic.CellFontBold = False
    grdPrgmmatic.CellFontName = "Arial"
    grdPrgmmatic.CellFontSize = 6.75
    grdPrgmmatic.CellForeColor = vbBlue
    grdPrgmmatic.CellBackColor = LIGHTBLUE
    grdPrgmmatic.TextMatrix(grdPrgmmatic.Row, grdPrgmmatic.Col) = "Start Date"
    grdPrgmmatic.Col = P_SALESOFFICEINDEX
    grdPrgmmatic.CellFontBold = False
    grdPrgmmatic.CellFontName = "Arial"
    grdPrgmmatic.CellFontSize = 6.75
    grdPrgmmatic.CellForeColor = vbBlue
    grdPrgmmatic.CellBackColor = LIGHTBLUE
    grdPrgmmatic.TextMatrix(grdPrgmmatic.Row, grdPrgmmatic.Col) = "Sales Office"
    grdPrgmmatic.Col = P_REASONINDEX
    grdPrgmmatic.CellFontBold = False
    grdPrgmmatic.CellFontName = "Arial"
    grdPrgmmatic.CellFontSize = 6.75
    grdPrgmmatic.CellForeColor = vbBlue
    grdPrgmmatic.CellBackColor = LIGHTBLUE
    grdPrgmmatic.TextMatrix(grdPrgmmatic.Row, grdPrgmmatic.Col) = "Status"
    grdPrgmmatic.Col = P_AUFCODEINDEX
    grdPrgmmatic.CellFontBold = False
    grdPrgmmatic.CellFontName = "Arial"
    grdPrgmmatic.CellFontSize = 6.75
    grdPrgmmatic.CellForeColor = vbBlue
    grdPrgmmatic.CellBackColor = LIGHTBLUE
    grdPrgmmatic.TextMatrix(grdPrgmmatic.Row, grdPrgmmatic.Col) = "Auf Code"
    grdPrgmmatic.Col = P_CNTRSTATUSINDEX
    grdPrgmmatic.CellFontBold = False
    grdPrgmmatic.CellFontName = "Arial"
    grdPrgmmatic.CellFontSize = 6.75
    grdPrgmmatic.CellForeColor = vbBlue
    grdPrgmmatic.CellBackColor = LIGHTBLUE
    grdPrgmmatic.TextMatrix(grdPrgmmatic.Row, grdPrgmmatic.Col) = "Status Char"
    grdPrgmmatic.Col = P_SORTINDEX
    grdPrgmmatic.CellFontBold = False
    grdPrgmmatic.CellFontName = "Arial"
    grdPrgmmatic.CellFontSize = 6.75
    grdPrgmmatic.CellForeColor = vbBlue
    grdPrgmmatic.CellBackColor = LIGHTBLUE
    grdPrgmmatic.TextMatrix(grdPrgmmatic.Row, grdPrgmmatic.Col) = "Sort"

End Sub
Private Sub mGridPrgmmaticColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdPrgmmatic.ColWidth(P_AUFCODEINDEX) = 0
    grdPrgmmatic.ColWidth(P_CNTRSTATUSINDEX) = 0
    grdPrgmmatic.ColWidth(P_SORTINDEX) = 0
    grdPrgmmatic.ColWidth(P_CREATEDATEINDEX) = 0.07 * grdPrgmmatic.Width
    grdPrgmmatic.ColWidth(P_CREATETIMEINDEX) = 0.08 * grdPrgmmatic.Width
    grdPrgmmatic.ColWidth(P_AGENCYINDEX) = 0.15 * grdPrgmmatic.Width
    grdPrgmmatic.ColWidth(P_ADVTNAMEINDEX) = 0.15 * grdPrgmmatic.Width
    grdPrgmmatic.ColWidth(P_PRODUCTINDEX) = 0.15 * grdPrgmmatic.Width
    grdPrgmmatic.ColWidth(P_CNTRNOINDEX) = 0.08 * grdPrgmmatic.Width
    grdPrgmmatic.ColWidth(P_CNTRSTARTDATEINDEX) = 0.07 * grdPrgmmatic.Width
    grdPrgmmatic.ColWidth(P_SALESOFFICEINDEX) = 0.1 * grdPrgmmatic.Width
    grdPrgmmatic.ColWidth(P_REASONINDEX) = 0.1 * grdPrgmmatic.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdPrgmmatic.Width
    For ilCol = 0 To grdPrgmmatic.Cols - 1 Step 1
        llWidth = llWidth + grdPrgmmatic.ColWidth(ilCol)
        If (grdPrgmmatic.ColWidth(ilCol) > 15) And (grdPrgmmatic.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdPrgmmatic.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdPrgmmatic.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdPrgmmatic.Width
            For ilCol = 0 To grdPrgmmatic.Cols - 1 Step 1
                If (grdPrgmmatic.ColWidth(ilCol) > 15) And (grdPrgmmatic.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdPrgmmatic.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdPrgmmatic.FixedCols To grdPrgmmatic.Cols - 1 Step 1
                If grdPrgmmatic.ColWidth(ilCol) > 15 Then
                    ilColInc = grdPrgmmatic.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdPrgmmatic.ColWidth(ilCol) = grdPrgmmatic.ColWidth(ilCol) + 15
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

Private Sub mGridPoolColumns()
    grdPool.Row = grdPool.FixedRows - 1
    grdPool.Col = UP_CREATEDATEINDEX
    grdPool.CellFontBold = False
    grdPool.CellFontName = "Arial"
    grdPool.CellFontSize = 6.75
    grdPool.CellForeColor = vbBlue
    grdPool.CellBackColor = LIGHTBLUE
    grdPool.TextMatrix(grdPool.Row, grdPool.Col) = "Create Date"
    grdPool.Col = UP_CREATETIMEINDEX
    grdPool.CellFontBold = False
    grdPool.CellFontName = "Arial"
    grdPool.CellFontSize = 6.75
    grdPool.CellForeColor = vbBlue
    grdPool.CellBackColor = LIGHTBLUE
    grdPool.TextMatrix(grdPool.Row, grdPool.Col) = "Create Time"
    grdPool.Col = UP_FILENAMEINDEX
    grdPool.CellFontBold = False
    grdPool.CellFontName = "Arial"
    grdPool.CellFontSize = 6.75
    grdPool.CellForeColor = vbBlue
    grdPool.CellBackColor = LIGHTBLUE
    grdPool.TextMatrix(grdPool.Row, grdPool.Col) = "File Name"

End Sub
Private Sub mGridPoolColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdPool.ColWidth(UP_AUFCODEINDEX) = 0
    grdPool.ColWidth(UP_SORTINDEX) = 0
    grdPool.ColWidth(UP_FILEDATEINDEX) = 0
    grdPool.ColWidth(UP_CREATEDATEINDEX) = 0.07 * grdPool.Width
    grdPool.ColWidth(UP_CREATETIMEINDEX) = 0.08 * grdPool.Width
    grdPool.ColWidth(UP_FILENAMEINDEX) = grdPool.Width - grdPool.ColWidth(UP_CREATEDATEINDEX) - grdPool.ColWidth(UP_CREATETIMEINDEX) - (GRIDSCROLLWIDTH + 45)
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdPool.Width
    For ilCol = 0 To grdPool.Cols - 1 Step 1
        llWidth = llWidth + grdPool.ColWidth(ilCol)
        If (grdPool.ColWidth(ilCol) > 15) And (grdPool.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdPool.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdPool.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdPool.Width
            For ilCol = 0 To grdPool.Cols - 1 Step 1
                If (grdPool.ColWidth(ilCol) > 15) And (grdPool.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdPool.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdPool.FixedCols To grdPool.Cols - 1 Step 1
                If grdPool.ColWidth(ilCol) > 15 Then
                    ilColInc = grdPool.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdPool.ColWidth(ilCol) = grdPool.ColWidth(ilCol) + 15
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
Private Sub mPoolSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdPool.FixedRows To grdPool.Rows - 1 Step 1
        slStr = Trim$(grdPool.TextMatrix(llRow, UP_CREATEDATEINDEX))
        If slStr <> "" Then
            If ilCol = UP_CREATEDATEINDEX Then
                slSort = Trim$(str$(gDateValue(grdPool.TextMatrix(llRow, UP_CREATEDATEINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = UP_CREATETIMEINDEX) Then
                slSort = Trim$(str$(gTimeToLong(grdPool.TextMatrix(llRow, UP_CREATETIMEINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = UP_FILENAMEINDEX) Then
                slSort = Trim$(str$(gDateValue(grdPool.TextMatrix(llRow, UP_FILEDATEINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdPool.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdPool.TextMatrix(llRow, UP_SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastPoolColSorted) Or ((ilCol = imLastPoolColSorted) And (imLastPoolSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdPool.TextMatrix(llRow, UP_SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdPool.TextMatrix(llRow, UP_SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastPoolColSorted Then
        imLastPoolColSorted = UP_SORTINDEX
    Else
        imLastPoolColSorted = -1
        imLastPoolSort = -1
    End If
    gGrid_SortByCol grdPool, UP_CREATEDATEINDEX, UP_SORTINDEX, imLastPoolColSorted, imLastPoolSort
    imLastPoolColSorted = ilCol
End Sub
Private Sub mGridPoolLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdPool.Rows - 1 Step 1
        grdPool.RowHeight(ilRow) = fgBoxGridH + 15
    Next ilRow
    For ilCol = 0 To grdPool.Cols - 1 Step 1
        grdPool.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub
Private Sub mRemovePool()
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand
    gRemoveFiles sgDBPath & "Messages\", "PoolUnassignedLog_", 30
    slSQLQuery = "DELETE FROM AUF_ALERT_USER WHERE aufType = 'U' AND aufSubType = 'P' AND aufMoWeekDate <= '" & Format$(DateAdd("d", -30, Now), sgSQLDateForm) & "'"
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        gHandleError "AffErrorLog.txt", "frmAlertVw-mRemovePool"
        Exit Sub
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "AlerVw-mPopulate"
End Sub
