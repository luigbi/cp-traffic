VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ImptCntr 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5790
   ClientLeft      =   450
   ClientTop       =   2355
   ClientWidth     =   8460
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5790
   ScaleWidth      =   8460
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Files"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4695
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1275
      Width           =   1410
   End
   Begin VB.FileListBox lbcBrowserFile 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   4695
      MultiSelect     =   2  'Extended
      Pattern         =   "H*.or"
      TabIndex        =   0
      Top             =   45
      Width           =   3630
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   225
      Left            =   2925
      TabIndex        =   29
      Top             =   2280
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.PictureBox plChkMove 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6765
      ScaleHeight     =   240
      ScaleWidth      =   7710
      TabIndex        =   13
      Top             =   1935
      Visible         =   0   'False
      Width           =   7710
      Begin VB.OptionButton rbcChkVeh 
         Caption         =   "No"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   4065
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   630
      End
      Begin VB.OptionButton rbcChkVeh 
         Caption         =   "Yes"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3390
         TabIndex        =   14
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.PictureBox pbcChkVeh 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1230
      Left            =   3435
      ScaleHeight     =   1200
      ScaleWidth      =   3825
      TabIndex        =   28
      Top             =   3945
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   1665
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   5
      Top             =   2505
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
         TabIndex        =   8
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Imptcntr.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   10
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   7
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.TextBox edcNoWks 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4785
      MaxLength       =   2
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1620
      Width           =   390
   End
   Begin VB.CommandButton cmcSDate 
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
      Left            =   2685
      Picture         =   "Imptcntr.frx":2E1A
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1620
      Width           =   195
   End
   Begin VB.TextBox edcSDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1770
      MaxLength       =   10
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1620
      Width           =   930
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6495
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5310
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5880
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5310
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6150
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5175
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcErrors 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2550
      Left            =   285
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2700
      Visible         =   0   'False
      Width           =   7905
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   5565
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.CommandButton cmcFileConv 
      Appearance      =   0  'Flat
      Caption         =   "Convert &Files"
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
      Left            =   2040
      TabIndex        =   17
      Top             =   5325
      Width           =   1830
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
      Left            =   4635
      TabIndex        =   16
      Top             =   5325
      Width           =   1050
   End
   Begin VB.Label lacMsg 
      Appearance      =   0  'Flat
      Caption         =   $"Imptcntr.frx":2F14
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   120
      TabIndex        =   30
      Top             =   135
      Width           =   4365
   End
   Begin VB.Label lacNoWks 
      Appearance      =   0  'Flat
      Caption         =   "Number of Weeks         (Blank -> closing not required)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3195
      TabIndex        =   11
      Top             =   1620
      Width           =   4875
   End
   Begin VB.Label lacSDate 
      Appearance      =   0  'Flat
      Caption         =   "Closing Start Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1620
      Width           =   1590
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   225
      Top             =   5190
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacErrors 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5730
      TabIndex        =   24
      Top             =   2295
      Width           =   1725
   End
   Begin VB.Label lacCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1050
      TabIndex        =   23
      Top             =   2295
      Width           =   1725
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   5865
      TabIndex        =   21
      Top             =   1965
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      Caption         =   "Line #"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   5100
      TabIndex        =   20
      Top             =   1965
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   3900
      TabIndex        =   19
      Top             =   1965
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      Caption         =   "Processing Contract #"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   1905
      TabIndex        =   18
      Top             =   1965
      Visible         =   0   'False
      Width           =   1965
   End
End
Attribute VB_Name = "ImptCntr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc.  All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ImptCntr.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the import contract conversion input screen code
Option Explicit
Option Compare Text
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls as Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim imFirstActivate As Integer
Dim imBypassFocus As Integer
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim hmCntr As Integer   'Contract header input file hanle
Dim hmLine As Integer   'Contract Line input file handle
Dim hmFlight As Integer 'Line flight input file handle
Dim hmMove As Integer   'Move input file hanle
Dim hmMsg As Integer    'Message file handle
Dim hmChf As Integer    'file handle
Dim imChfRecLen As Integer  'Record length
Dim tmChf As CHF
Dim tmChfSrchKey As CHFKEY1
Dim hmClf As Integer    'file handle
Dim imClfRecLen As Integer  'Record length
Dim tmClf As CLF
Dim tmClfSrchKey As CLFKEY0
Dim imPriceLevel As Integer
Dim hmCff As Integer    'file handle
Dim imCffRecLen As Integer  'Record length
Dim tmCff As CFF
Dim imSchCntr As Integer
Dim hmPrf As Integer    'file handle
Dim tmPrf As PRF
Dim imPrfRecLen As Integer  'Record length
Dim tmPrfSrchKey As PRFKEY1
Dim hmRcf As Integer    'file handle
Dim tmRcf As RCF
Dim imRcfRecLen As Integer  'Record length
Dim imRcfCode As Integer    'Latest RCF code
Dim hmRdf As Integer    'file handle
Dim tmRdf As RDF
Dim imRdfRecLen As Integer  'Record length
Dim imToRdfCode As Integer
Dim hmMnf As Integer    'file handle
Dim tmMnf As MNF        'Record structure
Dim imMnfRecLen As Integer  'Record length
Dim hmAdf As Integer    'file handle
Dim tmAdf As ADF        'Record structure
Dim imAdfRecLen As Integer  'Record length
Dim tmAdfSrchKey As INTKEY0
Dim smAdvtName As String
Dim smAdfCreditRestr As String
Dim hmAgf As Integer    'file handle
Dim tmAgf As AGF        'Record structure
Dim imAgfRecLen As Integer  'Record length
Dim tmAgfSrchKey As INTKEY0
Dim smAgyName As String
Dim smAgyCity As String
Dim smAgyAddr1 As String
Dim smAgyCreditRestr As String
Dim hmSlf As Integer    'file handle
Dim tmSlf As SLF        'Record structure
Dim imSlfRecLen As Integer  'Record length
Dim hmVef As Integer    'file handle
Dim tmVef As VEF        'Record structure
Dim imVefRecLen As Integer  'Record length
Dim hmVsf As Integer    'file handle
Dim tmVsf As VSF
Dim imVsfRecLen As Integer  'Record length
Dim hmVpf As Integer    'file handle
Dim tmVpf As VPF        'Record structure
Dim imVpfRecLen As Integer  'Record length
'Dim hmIcf As Integer    'file handle
'Dim imIcfRecLen As Integer  'Record length
'Dim tmIcf As ICF
Dim hmSif As Integer    'file handle
Dim tmSif As SIF
Dim imSifRecLen As Integer  'Record length
Dim tmSaf As SAF
Dim hmSaf As Integer
Dim imSafRecLen As Integer
Dim hmRlf As Integer

Dim smSeqNo As String
Dim imCurSeqNo As Integer
Dim lmLastMoveID As Long
Dim imProcMove As Integer

Dim imWarningCount As Integer   'Number of Warning messages printed
Dim imInfoCount As Integer      'Number of Information message printed

Dim imMS12MVefCode() As Integer

Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim tmSortCode() As SORTCODE

Dim lmMGRecNo As Long    'MG record number
Dim imShowMGErr As Integer
Dim imTerminate As Integer
Dim imConverting As Integer
Dim smFieldValues(1 To 50) As String    'largest number of fields (header)
Dim smMoveValues(1 To 24) As String     'Move values
Dim tmMoveRec As MGMOVEREC
Dim imFlightError As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim smRptDate As String
Dim smRptName As String
Dim smNowDate As String
Dim lmNowDate As Long
Dim imMatchLnUpper As Integer
Dim imNewAdvt As Integer
Dim imNewAgy As Integer
Dim imPorO As Integer   '0=Proposal; 1=Order
Dim imPkgRdfCode As Integer
Dim smNameMsg As String
Dim smSyncDate As String
Dim smSyncTime As String
Dim imBSMode As Integer
Dim imVpfIndex As Integer
Dim lmCompTime As Long
Dim lmSepLength As Long
Dim hmLcf As Integer    'file handle
Dim imLcfRecLen As Integer  'Record length
Dim tmLcf As LCF
Dim lmRecCount As Long
Dim lmErrorCount As Long
Dim imRet As Integer
Dim tmSpotMove(1 To 2) As SPOTMOVE
Dim tmVcf0() As VCF
Dim tmVcf6() As VCF
Dim tmVcf7() As VCF
'Spot record
Dim hmSdf As Integer    'file handle
Dim tmSdf As SDF
Dim imSdfRecLen As Integer
Dim tmSdfSrchKey0 As SDFKEY0
Dim tmSdfSrchKey1 As SDFKEY1
Dim tmSdfSrchKey3 As LONGKEY0
'Spot MG record
Dim hmSmf As Integer
Dim tmSmf As SMF
Dim imSmfRecLen As Integer
Dim tmSmfSrchKey2 As LONGKEY0   'SdfCode
Dim hmSsf As Integer
Dim tmSsf As SSF                'SSF record image
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image
Dim imSsfRecLen As Integer
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
'MG Track record
Dim hmMtf As Integer
Dim tmMtf As MTF
Dim imMtfRecLen As Integer
Dim tmMtfSrchKey As LONGKEY0
Dim tmMtfSrchKey1 As LONGKEY0
Dim tmMtfSrchKey2 As LONGKEY0
Dim lmPrevSdfCreated() As Long  'RefTrackID of spot previously Undone
Dim lmPrevMoveDeleted() As Long 'RefTrackID of Spot previously Undone move
Dim smPreemptPass As String
'Copy Rotation
Dim hmCrf As Integer
'Dim tmRdf As RDF
Dim tmBRdf As RDF
'Dim tmRcf As RCF
Dim tmErrMsg() As ERRMSG
Dim smBlankSpaces As String

Private Sub ckcAll_Click()
    If lbcBrowserFile.ListCount <= 0 Then
        Exit Sub
    End If
    'Code added because Value removed as parameter
    Dim Value As Integer
    
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of Coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcBrowserFile.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcBrowserFile.hwnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
End Sub

Private Sub ckcAll_GotFocus()
    plcCalendar.Visible = False
End Sub

' MsgBox parameters
'Const vbOkOnly = 0                 ' OK button only
'Const vbCritical = 16          ' Critical message
'Const vbApplicationModal = 0
'Const INDEXKEY0 = 0
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcSDate.SelStart = 0
    edcSDate.SelLength = Len(edcSDate.Text)
    edcSDate.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcSDate.SelStart = 0
    edcSDate.SelLength = Len(edcSDate.Text)
    edcSDate.SetFocus
End Sub
Private Sub cmcCancel_Click()
    If imConverting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcFileConv_Click()
    Dim ilRet As Integer
    Dim slStr As String
    Dim slDate As String
    Dim ilConvCHFRet As Integer
    
    If mUnschdCntr() Then
        MsgBox "Must schedule Contracts prior to Import", vbOkOnly + vbCritical + vbApplicationModal, "Unscheduled Contracts"
        cmcCancel.SetFocus
        Exit Sub
    End If
    slStr = Trim$(edcNoWks.Text)
    If slStr <> "" Then
        If Val(slStr) <> 0 Then
            slDate = Trim$(edcSDate.Text)
            If slDate <> "" Then
                If Not gValidDate(slDate) Then
                    MsgBox "Closing Date is not valid", vbOkOnly + vbCritical + vbApplicationModal, "Unscheduled Contracts"
                    cmcCancel.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
    ilRet = mSortNames()
    If Not ilRet Then
        MsgBox "Import File not selected or Sequence Numbers out of Order", vbOkOnly + vbCritical + vbApplicationModal, "Import Contracts"
        cmcCancel.SetFocus
        Exit Sub
    End If
    ilRet = mChkVeh()
    Screen.MousePointer = vbDefault
    If ilRet = 1 Then
        cmcCancel.SetFocus
        Exit Sub
    ElseIf ilRet = 2 Then
        'MsgBox "Vehicle(s) missing, see VehChk.Txt for list and/or Screen", vbOkOnly + vbCritical + vbApplicationModal, "Unscheduled Contracts"
        MsgBox "Missing Vehicle and/or Mo-Su 12M-12M Daypart and/or Avail Name, see VehChk.Txt for list and/or Screen", vbOkOnly + vbCritical + vbApplicationModal, "Import Contracts"
        cmcCancel.SetFocus
        Exit Sub
    End If
    lmErrorCount = 0
    ReDim lgReschSdfCode(1 To 1) As Long
    ReDim tmErrMsg(0 To 0) As ERRMSG
    gGetSchParameters
    gObtainMissedReasonCode
    Randomize   'Remove this if same results are to be obtained
    If gOpenSchFiles() Then
        imConverting = True
        imWarningCount = 0
        imInfoCount = 0
        ilConvCHFRet = mConvCHF()
        If ilConvCHFRet Then
            DoEvents
            cmcFileConv.Visible = False
            cmcCancel.Caption = "Done"
            'ilRet = mSetSeqNo(Val(smSeqNo))
            slStr = Trim$(edcNoWks.Text)
            If slStr <> "" Then
                If Val(slStr) <> 0 Then
                    slDate = Trim$(edcSDate.Text)
                    If slDate <> "" Then
                        If gValidDate(slDate) Then
                            ilRet = mSchdUnschd()
                        End If
                    End If
                End If
            End If
        End If
        gCloseSchFiles
        Screen.MousePointer = vbDefault
        'If lmErrorCount > 0 Then
        '    lacCntr(0).Caption = "Number of Warning Messages & Str$(imWarningCount)"
        '    lacCntr(1).Caption = "Number of Information Messages" & Str$(imInfoCount)
        '    lacCntr(0).Visible = True
        '    lacCntr(1).Visible = True
        'End If
        lbcErrors.AddItem "Number of Information Messages" & str$(imInfoCount), 0
        lbcErrors.AddItem "Number of Warning Messages" & str$(imWarningCount), 0
        imConverting = False
        cmcCancel.Caption = "&Done"
        If Not ilConvCHFRet Then
            MsgBox "Major Error, Import Stopped", vbOKOnly + vbCritical, "Major Error"
        ElseIf (imWarningCount > 0) Then
            MsgBox "Number of Warning Messages" & str$(imWarningCount) & ", Number of Information Messages" & str$(imInfoCount), vbOKOnly + vbInformation, "Error Counts"
        End If
        'cmcCancel.SetFocus
    End If
    Exit Sub
End Sub
Private Sub cmcFileConv_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcFrom_Click()
'    CMDialogBox.DialogTitle = "From File"
'    CMDialogBox.Filter = "Import|H*.OR|ASC|*.Asc|Text|*.Txt|Print|*.Prn|All|*.*"
'    CMDialogBox.InitDir = Left$(sgImportPath, Len(sgImportPath) - 1)
'    CMDialogBox.DefaultExt = ".OR"
'    CMDialogBox.Action = 1 'Open dialog
'    DoEvents
'    edcFrom.Text = CMDialogBox.fileName
'    If InStr(1, sgCurDir, ":") > 0 Then
'        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
'        ChDir sgCurDir
'    End If
'    DoEvents
End Sub
Private Sub cmcFrom_GotFocus()
'    plcCalendar.Visible = False
End Sub

Private Sub cmcSDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcSDate.SelStart = 0
    edcSDate.SelLength = Len(edcSDate.Text)
    edcSDate.SetFocus
End Sub
Private Sub edcFrom_GotFocus()
'    plcCalendar.Visible = False
'    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
'        imFirstFocus = False
'        'Show branner
'    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcNoWks_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub edcSDate_Change()
    Dim slStr As String
    slStr = edcSDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcSDate_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcSDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcSDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcSDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcSDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcSDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcSDate.Text = slDate
            End If
        End If
        edcSDate.SelStart = 0
        edcSDate.SelLength = Len(edcSDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcSDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcSDate.Text = slDate
            End If
        End If
        edcSDate.SelStart = 0
        edcSDate.SelLength = Len(edcSDate.Text)
    End If
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    
    lbcBrowserFile.fileName = Left$(sgImportPath, Len(sgImportPath) - 1)

    imFirstActivate = False
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Click()
    plcCalendar.Visible = False
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        gFunctionKeyBranch KeyCode
        'plcFrom.Visible = False
        'plcFrom.Visible = True
        plChkMove.Visible = False
        plChkMove.Visible = True
    End If
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
End Sub
Private Sub Form_Unload(Cancel As Integer)
    gGetSchParameters
    'ilRet = btrReset(hgHlf)
    'btrDestroy hgHlf
    'btrStopAppl
    'End
End Sub
Private Sub imcHelp_Click()
    plcCalendar.Visible = False
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lacCntr_Click(Index As Integer)
    plcCalendar.Visible = False
End Sub
Private Sub lacMsg_Click()
    plcCalendar.Visible = False
End Sub

Private Sub lbcBrowserFile_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked  '9-12-02 False
        imSetAll = True
    End If
End Sub

Private Sub lbcBrowserFile_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub lbcErrors_GotFocus()
    plcCalendar.Visible = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddAdvertiser                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Add advertiser                 *
'*                                                     *
'*******************************************************
Private Function mAddAdvertiser(slName As String, slDirect As String, slAddr1 As String, slAddr2 As String, slAddr3 As String, slBAddr1 As String, slBAddr2 As String, slBAddr3 As String) As Integer
'
'   ilCode = mAddAdvertiser (slName, slDirect)
'   Where:
'       slName(I)- Advertiser name
'       slDirect(I)- "A"=Agency, "D" = Direct
'       ilCode(O)- Advertiser code
'
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilUpper As Integer
    tmAdf.iCode = 0    'Internal code number for advertiser
    tmAdf.sName = slName  'Name
    tmAdf.sAbbr = Left$(slName, 7) 'Abbreviation
    tmAdf.sProduct = "" 'Product name
    tmAdf.iSlfCode = 0 'Salesperson code number
    tmAdf.iAgfCode = 0 'Agency code number
    tmAdf.sBuyer = ""   'Buyers name
    tmAdf.sCodeRep = "" 'Rep advertiser Code
    tmAdf.sCodeAgy = ""
    tmAdf.sCodeStn = "" 'Station advertiser Code
    tmAdf.iMnfComp(0) = 0 'Competitive code
    tmAdf.iMnfComp(1) = 0 'Competitive code
    tmAdf.iMnfExcl(0) = 0 'Program Exclusions code
    tmAdf.iMnfExcl(1) = 0 'Program Exclusions code
    tmAdf.sCppCpm = "N"    'P=CPP; M=CPM; N=N/A
    For ilLoop = 0 To 3 Step 1
        tmAdf.sDemo(ilLoop) = ""    'First-four Demo target
        'slStr = ""
        'gStrToPDN slStr, 2, 4, tmAdf.sTarget(ilLoop)
        tmAdf.iMnfDemo(ilLoop) = 0
        tmAdf.lTarget(ilLoop) = 0
    Next ilLoop
    tmAdf.sCreditRestr = "N"
    'tmAdf.sCreditLimit = "0"
    'slStr = ""
    'gStrToPDN slStr, 2, 5, tmAdf.sCreditLimit
    tmAdf.lCreditLimit = 0
    tmAdf.sPaymRating = "1"
    tmAdf.sShowISCI = "N"
    tmAdf.iMnfSort = 0
    If slDirect = "D" Then
        tmAdf.sBillAgyDir = "D"
        If Trim$(slAddr1) = "" Then
            tmAdf.sCntrAddr(0) = "*************************"
            For ilLoop = 1 To 2 Step 1
                tmAdf.sCntrAddr(ilLoop) = ""
            Next ilLoop
        Else
            tmAdf.sCntrAddr(0) = Trim$(slAddr1)
            tmAdf.sCntrAddr(1) = Trim$(slAddr2)
            tmAdf.sCntrAddr(2) = Trim$(slAddr3)
        End If
        tmAdf.sBillAddr(0) = Trim$(slBAddr1)
        tmAdf.sBillAddr(1) = Trim$(slBAddr2)
        tmAdf.sBillAddr(2) = Trim$(slBAddr3)
    Else
        tmAdf.sBillAgyDir = "A"
        For ilLoop = 0 To 2 Step 1
            tmAdf.sCntrAddr(ilLoop) = ""
        Next ilLoop
        For ilLoop = 0 To 2 Step 1
            tmAdf.sBillAddr(ilLoop) = ""
        Next ilLoop
    End If
    tmAdf.iArfLkCode = 0
    'Phone number (123) 456-789A Ext(BCDE)
    'Stored as 123456789ABCDE
    tmAdf.sPhone = "______________"
    tmAdf.sFax = "__________"
    tmAdf.iArfCntrCode = 0
    tmAdf.iArfInvCode = 0
    tmAdf.sCntrPrtSz = "N"
    '12/17/06-Change to tax by agency or vehicle
    'tmAdf.sSlsTax(0) = "N"
    'tmAdf.sSlsTax(1) = "N"
    tmAdf.iTrfCode = 0
    tmAdf.sCrdApp = "A" '"R" changed 6/30/00 via jim request
    tmAdf.sCrdRtg = ""
    tmAdf.iPnfBuyer = 0
    tmAdf.iPnfPay = 0
    'slStr = ""
    'gStrToPDN slStr, 0, 2, tmAdf.sPct90
    tmAdf.iPct90 = 0
    slStr = ""
    gStrToPDN slStr, 2, 6, tmAdf.sCurrAR
    slStr = ""
    gStrToPDN slStr, 2, 6, tmAdf.sUnbilled
    slStr = ""
    gStrToPDN slStr, 2, 6, tmAdf.sHiCredit
    slStr = ""
    gStrToPDN slStr, 2, 6, tmAdf.sTotalGross
    slDate = Format$(gNow(), "m/d/yy")
    gPackDate slDate, tmAdf.iDateEntrd(0), tmAdf.iDateEntrd(1)
    tmAdf.iNSFChks = 0
    tmAdf.iDateLstInv(0) = 0  'No date
    tmAdf.iDateLstInv(1) = 0
    tmAdf.iDateLstPaym(0) = 0  'No date
    tmAdf.iDateLstPaym(1) = 0
    tmAdf.iAvgToPay = 0
    tmAdf.iLstToPay = 0
    tmAdf.iNoInvPd = 0
    tmAdf.sNewBus = "N"
    tmAdf.iEndDate(0) = 0
    tmAdf.iEndDate(1) = 0
    tmAdf.iMerge = 0
    tmAdf.iurfCode = 2
    tmAdf.sState = "A"
    tmAdf.iCrdAppDate(0) = 0
    tmAdf.iCrdAppDate(1) = 0
    tmAdf.iCrdAppTime(0) = 0
    tmAdf.iCrdAppTime(1) = 0
    tmAdf.sPkInvShow = "T"
    tmAdf.sUnused2 = ""
    tmAdf.sRateOnInv = "Y"
    tmAdf.iMnfBus = 0
    tmAdf.lGuar = 0
    tmAdf.sAllowRepMG = "N"
    tmAdf.sBonusOnInv = "Y"
    tmAdf.sRepInvGen = "I"
    tmAdf.iMnfInvTerms = 0
    tmAdf.sPolitical = "N"
    tmAdf.sAddrID = ""
    tmAdf.iTrfCode = 0
    tmAdf.sUnused2 = ""
    ilRet = btrInsert(hmAdf, tmAdf, imAdfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        If ilRet >= 30000 Then
            ilRet = csiHandleValue(0, 7)
        End If
        slStr = "Advertiser " & Trim$(slName) & " not added for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " Error" & str$(ilRet)
        'lbcErrors.AddItem slStr
        mAddMsg slStr
        Print #hmMsg, slStr
        mAddAdvertiser = -2
        Exit Function
    End If

    ilUpper = UBound(tgCommAdf)
    tgCommAdf(ilUpper).iCode = tmAdf.iCode
    tgCommAdf(ilUpper).sName = Trim$(tmAdf.sName)
    tgCommAdf(ilUpper).sAbbr = Trim$(tmAdf.sAbbr)
    tgCommAdf(ilUpper).sBillAgyDir = tmAdf.sBillAgyDir
    tgCommAdf(ilUpper).sState = tmAdf.sState
    tgCommAdf(ilUpper).sAllowRepMG = tmAdf.sAllowRepMG
    tgCommAdf(ilUpper).sBonusOnInv = tmAdf.sBonusOnInv
    tgCommAdf(ilUpper).sRepInvGen = tmAdf.sRepInvGen
    tgCommAdf(ilUpper).sFirstByteCntrAddr = Left$(tmAdf.sCntrAddr(0), 1)
    tgCommAdf(ilUpper).sPolitical = tmAdf.sPolitical                '5-26-06
    tgCommAdf(ilUpper).sAddrID = Trim$(tmAdf.sAddrID)
    tgCommAdf(ilUpper).iTrfCode = tmAdf.iTrfCode
    ilUpper = ilUpper + 1
    ReDim Preserve tgCommAdf(1 To ilUpper) As ADFEXT
    mAddAdvertiser = tmAdf.iCode
    imInfoCount = imInfoCount + 1
    Print #hmMsg, smBlankSpaces & smBlankSpaces & "Advertiser: " & Trim$(tmAdf.sName) & " added"
'    'lbcErrors.AddItem smBlankSpaces & "Advertiser: " & Trim$(tmAdf.sName) & " added"
'    mAddMsg smBlankSpaces & "Advertiser: " & Trim$(tmAdf.sName) & " added"
    imNewAdvt = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddAgency                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Add Agency                     *
'*                                                     *
'*******************************************************
Private Function mAddAgency(slName As String, slCityID As String, slAddr1 As String, slAddr2 As String, slAddr3 As String, slBAddr1 As String, slBAddr2 As String, slBAddr3 As String, slCrdApp As String) As Integer
'
'   ilCode = mAddAgency( slName )
'   Where:
'       slName(I)- Agency name
'       ilCode(O)- Agency code
'
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilUpper As Integer
    Dim ilRet As Integer
    Dim slDate As String
    tmAgf.iCode = 0
    tmAgf.sName = slName
    tmAgf.sAbbr = Left$(tmAgf.sName, 5)
    If Trim$(slCityID) = "" Then
        tmAgf.sCityID = "*****"
    Else
        tmAgf.sCityID = slCityID
    End If
    slStr = "15.00"
    'gStrToPDN slStr, 2, 3, tmAgf.sComm
    tmAgf.iComm = gStrDecToInt(slStr, 2)
    tmAgf.iSlfCode = 0
    tmAgf.sBuyer = ""
    tmAgf.sCodeRep = ""
    tmAgf.sCodeStn = ""
    tmAgf.sCreditRestr = "N"
    'tmAgf.sCreditLimit = "0"
    'slStr = ""
    'gStrToPDN slStr, 2, 5, tmAgf.sCreditLimit
    tmAgf.lCreditLimit = 0
    tmAgf.sPaymRating = "1"
    tmAgf.sShowISCI = "N"
    tmAgf.iMnfSort = 0
    If Trim$(slAddr1) = "" Then
        tmAgf.sCntrAddr(0) = "*************************"
        For ilLoop = 1 To 2 Step 1
            tmAgf.sCntrAddr(ilLoop) = ""
        Next ilLoop
    Else
        tmAgf.sCntrAddr(0) = Trim$(slAddr1)
        tmAgf.sCntrAddr(1) = Trim$(slAddr2)
        tmAgf.sCntrAddr(2) = Trim$(slAddr3)
    End If
    tmAgf.sBillAddr(0) = Trim$(slBAddr1)
    tmAgf.sBillAddr(1) = Trim$(slBAddr2)
    tmAgf.sBillAddr(2) = Trim$(slBAddr3)
    tmAgf.iArfLkCode = 0
    'Phone number (123) 456-789A Ext(BCDE)
    'Stored as 123456789ABCDE
    tmAgf.sPhone = "______________"
    tmAgf.sFax = "__________"
    tmAgf.iArfCntrCode = 0
    tmAgf.iArfInvCode = 0
    tmAgf.sCntrPrtSz = "N"
    '12/17/06-Change to tax by agency or vehicle
    'tmAgf.sSlsTax(0) = "N"
    'tmAgf.sSlsTax(1) = "N"
    tmAgf.iTrfCode = 0
    tmAgf.sCrdApp = slCrdApp    '"R"
    tmAgf.sCrdRtg = ""
    tmAgf.iPnfBuyer = 0
    tmAgf.iPnfPay = 0
    'slStr = ""
    'gStrToPDN slStr, 0, 2, tmAgf.sPct90
    tmAgf.iPct90 = 0
    slStr = ""
    gStrToPDN slStr, 2, 6, tmAgf.sCurrAR
    slStr = ""
    gStrToPDN slStr, 2, 6, tmAgf.sUnbilled
    slStr = ""
    gStrToPDN slStr, 2, 6, tmAgf.sHiCredit
    slStr = ""
    gStrToPDN slStr, 2, 6, tmAgf.sTotalGross
    slDate = Format$(gNow(), "m/d/yy")
    gPackDate slDate, tmAgf.iDateEntrd(0), tmAgf.iDateEntrd(1)
    tmAgf.iNSFChks = 0
    tmAgf.iDateLstInv(0) = 0  'No date
    tmAgf.iDateLstInv(1) = 0
    tmAgf.iDateLstPaym(0) = 0  'No date
    tmAgf.iDateLstPaym(1) = 0
    tmAgf.iAvgToPay = 0
    tmAgf.iLstToPay = 0
    tmAgf.iNoInvPd = 0
    tmAgf.iMerge = 0
    tmAgf.iurfCode = 2
    tmAgf.sState = "A"
    tmAgf.iCrdAppDate(0) = 0
    tmAgf.iCrdAppDate(1) = 0
    tmAgf.iCrdAppTime(0) = 0
    tmAgf.iCrdAppTime(1) = 0
    tmAgf.sPkInvShow = "T"
    tmAgf.iTrfCode = 0
    tmAgf.iRemoteID = tgUrf(0).iRemoteUserID
    tmAgf.iAutoCode = 0
    ilRet = btrInsert(hmAgf, tmAgf, imAgfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        If ilRet >= 30000 Then
            ilRet = csiHandleValue(0, 7)
        End If
        slStr = "Agency " & Trim$(slName) & ", " & Trim$(slCityID) & " not added for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " Error" & str$(ilRet)
        'lbcErrors.AddItem slStr
        mAddMsg slStr
        Print #hmMsg, slStr
        mAddAgency = -2
        Exit Function
    End If
    tmAgf.iRemoteID = tgUrf(0).iRemoteUserID
    tmAgf.iAutoCode = tmAgf.iCode
    tmAgf.iSourceID = tgUrf(0).iRemoteUserID
    gPackDate smSyncDate, tmAgf.iSyncDate(0), tmAgf.iSyncDate(1)
    gPackTime smSyncTime, tmAgf.iSyncTime(0), tmAgf.iSyncTime(1)
    ilRet = btrUpdate(hmAgf, tmAgf, imAgfRecLen)
    ilUpper = UBound(tgCommAgf)
    tgCommAgf(ilUpper).iCode = tmAgf.iCode
    tgCommAgf(ilUpper).sName = Trim$(tmAgf.sName)
    tgCommAgf(ilUpper).sCityID = Trim$(tmAgf.sCityID)
    tgCommAgf(ilUpper).sCreditRestr = tmAgf.sCreditRestr
    tgCommAgf(ilUpper).iMnfSort = tmAgf.iMnfSort
    tgCommAgf(ilUpper).sState = tmAgf.sState
    tgCommAgf(ilUpper).iTrfCode = tmAgf.iTrfCode
    ilUpper = ilUpper + 1
    ReDim Preserve tgCommAgf(1 To ilUpper) As AGFEXT
    mAddAgency = tmAgf.iCode
    'MAI/Sirius: Added 10/4 Added imInfoCount, second smBlankSpaces and removed mAddMsg
    imInfoCount = imInfoCount + 1
    Print #hmMsg, smBlankSpaces & smBlankSpaces & "Agency: " & Trim$(tmAgf.sName) & ", " & Trim$(tmAgf.sCityID) & " added"
    'lbcErrors.AddItem smBlankSpaces & "Agency: " & Trim$(tmAgf.sName) & ", " & Trim$(tmAgf.sCityID) & " added"
'    mAddMsg smBlankSpaces & smBlankSpaces & "Agency: " & Trim$(tmAgf.sName) & ", " & Trim$(tmAgf.sCityID) & " added"
    imNewAgy = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddComp                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Add Comp                       *
'*                                                     *
'*******************************************************
Private Function mAddComp(slName As String) As Integer
'
'   ilCode = mAddComp( slName )
'   Where:
'       slName(I)- Competitive name
'       ilCode(O)- Competitive code
'
    Dim slStr As String
    Dim ilUpper As Integer
    Dim ilRet As Integer
    If Trim$(slName) = "" Then
        mAddComp = 0
        Exit Function
    End If
    tmMnf.iCode = 0    'Internal code number for multi-name
    tmMnf.sType = "C" '"I"=Item Billing; "A"=Announcer; "P"=Sport; "S"=Sales Source;
                        '"T"=Sales Team; "R"=Revenue Sets; "M"= Missed Reason; "C"=Competitive;
                        '"N"=Feed type; "V"=Invoice sorting; Sales regions
    tmMnf.sName = slName  'Name
    slStr = ""
    gStrToPDN slStr, 2, 5, tmMnf.sRPU
    tmMnf.sUnitType = Left$(slName, 5)
    slStr = ""
    gStrToPDN slStr, 4, 4, tmMnf.sSSComm
    tmMnf.iMerge = 0   'Merge code number
    tmMnf.iGroupNo = 0 'If "A": Group number or if "S": sales origin (1=Local; 2=Regional; 3=National)
    tmMnf.sCodeStn = ""  '"C" and "R" codes assigned by station
    tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
    tmMnf.iAutoCode = 0
    ilRet = btrInsert(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        If ilRet >= 30000 Then
            ilRet = csiHandleValue(0, 7)
        End If
        slStr = "Product Protection " & Trim$(slName) & " not added for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " Error" & str$(ilRet)
        'lbcErrors.AddItem slStr
        mAddMsg slStr
        Print #hmMsg, slStr
        mAddComp = -2
        Exit Function
    End If
    tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
    tmMnf.iAutoCode = tmMnf.iCode
    gPackDate smSyncDate, tmMnf.iSyncDate(0), tmMnf.iSyncDate(1)
    gPackTime smSyncTime, tmMnf.iSyncTime(0), tmMnf.iSyncTime(1)
    ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
    ilUpper = UBound(tgCompMnf)
    tgCompMnf(ilUpper).iCode = tmMnf.iCode
    tgCompMnf(ilUpper).sName = tmMnf.sName
    ilUpper = ilUpper + 1
    ReDim Preserve tgCompMnf(1 To ilUpper) As MNFCOMPEXT
    mAddComp = tmMnf.iCode
    imInfoCount = imInfoCount + 1
    Print #hmMsg, smBlankSpaces & smBlankSpaces & "Product Protection: " & Trim$(slName) & " added"
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddMsg                         *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add message to list box if room *
'*                                                     *
'*******************************************************
Private Sub mAddMsg(slStr As String)
    If lbcErrors.ListCount < 30000 Then
        lbcErrors.AddItem slStr
        If lbcErrors.ListCount = 30000 Then
            lbcErrors.AddItem "See Printout for other error messages"
        End If
    End If
    If Left$(slStr, 1) <> " " Then
        lmErrorCount = lmErrorCount + 1
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddSalesperson                 *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Add salesperson                *
'*                                                     *
'*******************************************************
Private Function mAddSalesperson(slLastName As String, slFirstName As String) As Integer
'
'   mAddAdvt slLastName
'   Where:
'       slLastName(I)- Last name
'
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilUpper As Integer
    tmSlf.iCode = 0        'Internal code number for salesperson
    tmSlf.sFirstName = slFirstName
    tmSlf.sLastName = slLastName    'Last name
    tmSlf.iSofCode = 0     'Sales office code
    'Phone number (123) 456-789A Ext(BCDE)
    'Stored as 123456789ABCDE
    tmSlf.sPhone = "______________"
    tmSlf.sFax = "__________"
    tmSlf.iMnfSlsTeam = 0  'Sales team code
    'tmSlf.sSS = ""       'Social security
    'tmSlf.sGLAcct = ""  'G/L Account #
    'slStr = ""
    'gStrToPDN slStr, 4, 4, tmSlf.sComm
    'tmSlf.lComm = 0
    'slStr = ""
    'gStrToPDN slStr, 2, 4, tmSlf.sDrawAmt
    tmSlf.iMerge = 0       'Merge code number
    tmSlf.sJobTitle = "S"       'Salesperson
    tmSlf.sState = "A"       'Active
    tmSlf.iurfCode = 2     'Last user code number who altered this file
    tmSlf.sCodeStn = ""  'Station salesperson code
    tmSlf.lSalesGoal = 0       'Merge code number
    tmSlf.iUnderComm = 0       'Merge code number
    tmSlf.iOverComm = 0       'Merge code number
    tmSlf.iRemUnderComm = 0       'Merge code number
    tmSlf.iRemOverComm = 0       'Merge code number
    tmSlf.lStartCommPaid = 0       'Merge code number
    tmSlf.lStartSales = 0       'Merge code number
    tmSlf.iStartCommDate(0) = 0
    tmSlf.iStartCommDate(1) = 0
    tmSlf.iRemoteID = tgUrf(0).iRemoteUserID
    tmSlf.iAutoCode = 0
    ilRet = btrInsert(hmSlf, tmSlf, imSlfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        If ilRet >= 30000 Then
            ilRet = csiHandleValue(0, 7)
        End If
        slStr = "Salesperson " & Trim$(slFirstName) & " " & Trim$(slLastName) & " not added for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " Error" & str$(ilRet)
        'lbcErrors.AddItem slStr
        mAddMsg slStr
        Print #hmMsg, slStr
        mAddSalesperson = -2
        Exit Function
    End If
    tmSlf.iRemoteID = tgUrf(0).iRemoteUserID
    tmSlf.iAutoCode = tmSlf.iCode
    gPackDate smSyncDate, tmSlf.iSyncDate(0), tmSlf.iSyncDate(1)
    gPackTime smSyncTime, tmSlf.iSyncTime(0), tmSlf.iSyncTime(1)
    ilRet = btrUpdate(hmSlf, tmSlf, imSlfRecLen)
    ilUpper = UBound(tgCSlf)
    tgCSlf(ilUpper).iCode = tmSlf.iCode
    tgCSlf(ilUpper).sFirstName = tmSlf.sFirstName
    tgCSlf(ilUpper).sLastName = tmSlf.sLastName
    ilUpper = ilUpper + 1
    ReDim Preserve tgCSlf(1 To ilUpper) As SLFEXT
    mAddSalesperson = tmSlf.iCode
    imInfoCount = imInfoCount + 1
    Print #hmMsg, smBlankSpaces & smBlankSpaces & "Salesperson: " & Trim$(tmSlf.sFirstName) & " " & Trim$(tmSlf.sLastName) & " added"
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddVehicle                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Add Vehicle                    *
'*                                                     *
'*******************************************************
Private Function mAddVehicle(slType As String, slName As String) As Integer
'
'   ilCode = mAddVehicle( slType, slName )
'   Where:
'       slName(I)- Competitive name
'       ilCode(O)- Competitive code
'
    Dim slStr As String
    Dim ilRet As Integer
    If Trim$(slName) = "" Then
        mAddVehicle = 0
        Exit Function
    End If
    If slType = "C" Then
        slStr = "Vehicle " & Trim$(slName) & " not found for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo))
        'lbcErrors.AddItem slStr
        mAddMsg slStr
        Print #hmMsg, slStr
        mAddVehicle = -2
        Exit Function
    End If
    tmVef.iCode = 0
    tmVef.sName = slName
    'tmVef.sMktName = ""
    tmVef.iMnfVehGp3Mkt = 0
    tmVef.iMnfVehGp5Rsch = 0
    tmVef.sAddr(0) = ""
    tmVef.sAddr(1) = ""
    tmVef.sAddr(2) = ""
    tmVef.sPhone = ""
    tmVef.sFax = ""
    tmVef.sDialPos = ""
    tmVef.lPvfCode = 0
    'tmVef.sFormat = ""
    tmVef.iMnfVehGp4Fmt = 0
    tmVef.iMnfVehGp2 = 0
    tmVef.sType = slType
    tmVef.sCodeStn = ""
    tmVef.iVefCode = 0
    tmVef.iCombineVefCode = 0
    tmVef.iOwnerMnfCode = 0
    tmVef.iProdPct(1) = 0
    tmVef.iProdPct(2) = 0
    tmVef.iProdPct(3) = 0
    tmVef.iProdPct(4) = 0
    tmVef.iProdPct(5) = 0
    tmVef.iProdPct(6) = 0
    tmVef.iProdPct(7) = 0
    tmVef.iProdPct(8) = 0
    tmVef.sState = "A"
    tmVef.iMnfGroup(1) = 0
    tmVef.iMnfGroup(2) = 0
    tmVef.iMnfGroup(3) = 0
    tmVef.iMnfGroup(4) = 0
    tmVef.iMnfGroup(5) = 0
    tmVef.iMnfGroup(6) = 0
    tmVef.iMnfGroup(7) = 0
    tmVef.iMnfGroup(8) = 0
    tmVef.iSort = 0
    tmVef.iDnfCode = 0
    tmVef.iReallDnfCode = 0
    tmVef.iMnfDemo = 0
    tmVef.iMnfSSCode(1) = 0
    tmVef.iMnfSSCode(2) = 0
    tmVef.iMnfSSCode(3) = 0
    tmVef.iMnfSSCode(4) = 0
    tmVef.iMnfSSCode(5) = 0
    tmVef.iMnfSSCode(6) = 0
    tmVef.iMnfSSCode(7) = 0
    tmVef.iMnfSSCode(8) = 0
    tmVef.sUpdateRVF(1) = "Y"
    tmVef.sUpdateRVF(2) = "Y"
    tmVef.sUpdateRVF(3) = "Y"
    tmVef.sUpdateRVF(4) = "Y"
    tmVef.sUpdateRVF(5) = "Y"
    tmVef.sUpdateRVF(6) = "Y"
    tmVef.sUpdateRVF(7) = "Y"
    tmVef.sUpdateRVF(8) = "Y"
    tmVef.sUnused3 = ""
    tmVef.lVsfCode = 0
    tmVef.lRateAud = 0
    tmVef.lCPPCPM = 0
    tmVef.lYearAvails = 0
    tmVef.iPctSellout = 0
    tmVef.iMnfVehGp6Sub = 0
    tmVef.iNrfCode = 0
    tmVef.iSSMnfCode = 0
    tmVef.sStdPrice = ""
    tmVef.sStdInvTime = ""
    tmVef.sStdAlter = ""
    tmVef.iStdIndex = 0
    tmVef.iRemoteID = tgUrf(0).iRemoteUserID
    tmVef.iAutoCode = tmVef.iCode
    ilRet = btrInsert(hmVef, tmVef, imVefRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        If ilRet >= 30000 Then
            ilRet = csiHandleValue(0, 7)
        End If
        slStr = "Vehicle " & Trim$(slName) & " not added for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " Error" & str$(ilRet)
        'lbcErrors.AddItem slStr
        mAddMsg slStr
        Print #hmMsg, slStr
        mAddVehicle = -2
        Exit Function
    End If
    tmVef.iRemoteID = tgUrf(0).iRemoteUserID
    tmVef.iAutoCode = tmVef.iCode
    'tmVef.iSourceID = tgUrf(0).iRemoteUserID
    'gPackDate smSyncDate, tmVef.iSyncDate(0), tmVef.iSyncDate(1)
    'gPackTime smSyncTime, tmVef.iSyncTime(0), tmVef.iSyncTime(1)
    ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
    tgMVef(UBound(tgMVef)) = tmVef
    ReDim Preserve tgMVef(1 To UBound(tgMVef) + 1) As VEF
    If UBound(tgMVef) > 2 Then
        ArraySortTyp fnAV(tgMVef(), 1), UBound(tgMVef) - 1, 0, LenB(tgMVef(1)), 0, -1, 0
    End If
    'Create Vpf
    ilRet = gVpfFindHd(ImptCntr, tmVef.iCode, hmVpf)
    'ilUpper = UBound(tgVef)
    'tgVef(ilUpper) = tmVef
    'ilUpper = ilUpper + 1
    'ReDim Preserve tgVef(1 To ilUpper) As VEF
    mAddVehicle = tmVef.iCode
    imInfoCount = imInfoCount + 1
    If slType = "P" Then
        Print #hmMsg, smBlankSpaces & smBlankSpaces & "Package vehicle: " & Trim$(tmVef.sName) & " added"
    Else
        Print #hmMsg, smBlankSpaces & smBlankSpaces & "Vehicle: " & Trim$(tmVef.sName) & " added"
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAnyConflicts                   *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test if Ok to book spot        *
'*                                                     *
'*******************************************************
Private Function mAnyConflicts(ilAvailIndex As Integer, ilAdfCode As Integer, ilMnfComp0 As Integer, ilMnfComp1 As Integer, slType As String, ilDay As Integer) As Integer
    Dim ilSpotIndex As Integer
    Dim ilAvHour As Integer
    Dim ilHour As Integer
    Dim ilEvt As Integer
    Dim ilPSACount As Integer
    Dim ilPromoCount As Integer
    Dim tlAvail As AVAILSS
    Dim llSsfRecPos As Long
   LSet tlAvail = tmSsf.tPas(ADJSSFPASBZ + ilAvailIndex)
    If (tlAvail.iRecType = 2) And (((slType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((slType = "M") And (tgSpf.sSchdPromo <> "Y"))) Then
        'Get count- and test if max exceded
        ilAvHour = tlAvail.iTime(1) \ 256  'Obtain month
        ilPSACount = 0
        ilPromoCount = 0
        'Get start of hour
        For ilEvt = 1 To tmSsf.iCount Step 1
           LSet tlAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
            If tlAvail.iRecType = 2 Then
                ilHour = tlAvail.iTime(1) \ 256  'Obtain month
                If ilHour > ilAvHour Then
                    Exit For
                End If
                If (ilAvHour = ilHour) Then
                    For ilSpotIndex = ilEvt + 1 To ilEvt + tlAvail.iNoSpotsThis Step 1
                       LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
                        If (tmSpot.iRank And RANKMASK) = 1050 Then 'Promo
                            ilPromoCount = ilPromoCount + 1
                        ElseIf (tmSpot.iRank And RANKMASK) = 1060 Then 'PSA
                            ilPSACount = ilPSACount + 1
                        End If
                    Next ilSpotIndex
                End If
            End If
        Next ilEvt
        If slType = "S" Then
            If ilDay <= 4 Then
                If ilPSACount >= tgVpf(imVpfIndex).iMMFPSA(ilAvHour + 1) Then
                    mAnyConflicts = True
                    Exit Function
                End If
            ElseIf ilDay = 5 Then
                If ilPSACount >= tgVpf(imVpfIndex).iMSAPSA(ilAvHour + 1) Then
                    mAnyConflicts = True
                    Exit Function
                End If
            ElseIf ilDay = 6 Then
                If ilPSACount >= tgVpf(imVpfIndex).iMSUPSA(ilAvHour + 1) Then
                    mAnyConflicts = True
                    Exit Function
                End If
            End If
        Else
            If ilDay <= 4 Then
                If ilPromoCount >= tgVpf(imVpfIndex).iMMFPromo(ilAvHour + 1) Then
                    mAnyConflicts = True
                    Exit Function
                End If
            ElseIf ilDay = 5 Then
                If ilPromoCount >= tgVpf(imVpfIndex).iMSAPromo(ilAvHour + 1) Then
                    mAnyConflicts = True
                    Exit Function
                End If
            ElseIf ilDay = 6 Then
                If ilPromoCount >= tgVpf(imVpfIndex).iMSUPromo(ilAvHour + 1) Then
                    mAnyConflicts = True
                    Exit Function
                End If
            End If
        End If
    End If
    'If rbcAC(2).Value Then
    '    mAnyConflicts = False
    '    Exit Function
    'ElseIf rbcAC(0).Value Then
        llSsfRecPos = 0 'Only used if preempting- Not preempting
        If Not gAdvtTest(tmSsf, llSsfRecPos, tmSpotMove(), imVpfIndex, lmSepLength, ilAvailIndex, tmChf.iAdfCode, tmChf.iMnfComp(0), tmChf.iMnfComp(1), 0, 0, "I", "N", imPriceLevel, True) Then
            mAnyConflicts = True
            Exit Function
        End If
        If Not gCompetitiveTest(lmCompTime, hmSsf, tmSsf, llSsfRecPos, tmSpotMove(), imVpfIndex, tmClf.iLen, tmChf.iMnfComp(0), tmChf.iMnfComp(1), ilAvailIndex, tmVcf0(), tmVcf6(), tmVcf7(), 0, 0, "I", "N", imPriceLevel, True) Then
            mAnyConflicts = True
            Exit Function
        End If
    'Else
    '   LSet tlAvail = tmSsf.tPas(ADJSSFPASBZ + ilAvailIndex)
    '    For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tlAvail.iNoSpotsThis Step 1
    '       LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
    '        ilMatchComp = False
    '        If (ilMnfComp0 = 0) And (ilMnfComp1 = 0) And (tmSpot.iMnfComp(0) = 0) And (tmSpot.iMnfComp(1) = 0) Then
    '            ilMatchComp = True
    '        Else
    '            If (ilMnfComp0 <> 0) And ((ilMnfComp0 = tmSpot.iMnfComp(0)) Or (ilMnfComp0 = tmSpot.iMnfComp(1))) Then
    '                ilMatchComp = True
    '            End If
    '            If (ilMnfComp1 <> 0) And ((ilMnfComp1 = tmSpot.iMnfComp(0)) Or (ilMnfComp1 = tmSpot.iMnfComp(1))) Then
    '                ilMatchComp = True
    '            End If
    '        End If
    '        'Advertiser conflict if same competitive only
    '        If (tmSpot.iAdfCode = ilAdfCode) And (ilMatchComp) Then
    '            mAnyConflicts = True
    '            Exit Function
    '        End If
    '        If (ilMnfComp0 <> 0) And ((ilMnfComp0 = tmSpot.iMnfComp(0)) Or (ilMnfComp0 = tmSpot.iMnfComp(1))) Then
    '            mAnyConflicts = True
    '            Exit Function
    '        ElseIf (ilMnfComp1 <> 0) And ((ilMnfComp1 = tmSpot.iMnfComp(0)) Or (ilMnfComp1 = tmSpot.iMnfComp(1))) Then
    '            mAnyConflicts = True
    '            Exit Function
    '        End If
    '    Next ilSpotIndex
    'End If
    mAnyConflicts = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mBonusSpot                       *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Book a bonus spot              *
'*                                                     *
'*******************************************************
Private Function mBonusSpot(llSDate As Long, llEDate As Long, llSTime As Long, llETime As Long, ilDays() As Integer, ilToVefCode As Integer, llTrackID As Long, llEnteredDate As Long, llRefTrackID As Long, llTransGpID As Long, ilAnfCode As Integer) As Integer
'
'
    Dim slDate As String
    Dim llSundayDate As Long
    Dim llLcfEarliestDate As Long
    Dim llLcfLatestDate As Long
    Dim llDate As Long
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilEvt As Integer
    Dim ilAvEvt As Integer
    Dim ilUnits As Integer
    Dim ilLen As Integer
    Dim ilPosition As Integer
    Dim llSsfRecPos As Long
    Dim llSdfRecPos As Long
    Dim llTime As Long
    Dim ilType As Integer
    Dim ilBkQH As Integer
    Dim slSchStatus As String
    Dim ilSpot As Integer
    Dim ilDay As Integer
    Dim llExtendDate As Long
    Dim ilNoDays As Integer
    Dim ilFound As Integer
    Dim ilDate As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilTest As Integer
    Dim llLockRecCode As Long
    Dim slUserName As String
    Dim ilPriceLevel As Integer
    Dim ilVefIndex As Integer
    
    'Make sure SSF Created
    slDate = Format$(llEDate, "m/d/yy")
    slDate = gObtainNextSunday(slDate)
    llSundayDate = gDateValue(slDate)
    llLcfEarliestDate = gGetEarliestLCFDate(hmLcf, "C", ilToVefCode)
    If (llLcfEarliestDate > 0) And (llSundayDate < llLcfEarliestDate) Then
        For llDate = llSundayDate - 6 To llLcfEarliestDate - 1 Step 1
            slDate = Format$(llDate, "m/d/yy")
            gPackDate slDate, ilLogDate0, ilLogDate1
            ilRet = gExtendTFN(hmLcf, hmSsf, hmSdf, hmSmf, "C", ilToVefCode, ilLogDate0, ilLogDate1, False)
            If Not ilRet Then
                igBtrError = 0
                sgErrLoc = "mBonusSpot, Extend Date"
                mBonusSpot = False
                Exit Function
            End If
        Next llDate
    End If
    llLcfLatestDate = gGetLatestLCFDate(hmLcf, "C", ilToVefCode)
    If (llLcfLatestDate > 0) Then
        llExtendDate = llLcfLatestDate + 1
    Else
        llExtendDate = lmNowDate
    End If
    For llDate = llExtendDate To llSundayDate Step 1
        slDate = Format$(llDate, "m/d/yy")
        gPackDate slDate, ilLogDate0, ilLogDate1
        ilRet = gExtendTFN(hmLcf, hmSsf, hmSdf, hmSmf, "C", ilToVefCode, ilLogDate0, ilLogDate1, False)
        If Not ilRet Then
            igBtrError = 0
            sgErrLoc = "mBonusSpot, Extend Date"
            mBonusSpot = False
            Exit Function
        End If
    Next llDate
    ilPosition = -1
    'ilBkQH = 1045   'ignore booking quarter hour since manual move
    '3/1/01- JIm make bonus priority one so that we max the number scheduled
    ilBkQH = 1
    imPriceLevel = 0
    ilType = 0
    'Mix days
    ilNoDays = llEDate - llSDate + 1
    ReDim llRndDates(1 To ilNoDays) As Long
    For ilLoop = 1 To ilNoDays Step 1
        llRndDates(ilLoop) = 0
    Next ilLoop
    For ilLoop = 1 To ilNoDays Step 1
        ilTest = Int(ilNoDays * Rnd + 1)
        Do
            ilFound = False
            For ilIndex = 1 To ilNoDays Step 1
                If llRndDates(ilIndex) = llSDate + ilTest - 1 Then
                    ilFound = True
                    Exit For
                End If
            Next ilIndex
            If Not ilFound Then
                Exit Do
            End If
            ilTest = Int(ilNoDays * Rnd + 1)
        Loop
        llRndDates(ilLoop) = llSDate + ilTest - 1
    Next ilLoop
    ilVefIndex = gBinarySearchVef(ilToVefCode)
    'For llDate = llSDate To llEDate Step 1
    For ilDate = 1 To ilNoDays Step 1
        llDate = llRndDates(ilDate)
        ilDay = gWeekDayLong(llDate)
        If ilDays(ilDay) <> 0 Then   'ckcDay(ilDay).Value Then
            slDate = Format$(llDate, "m/d/yy")
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            gPackDate slDate, ilLogDate0, ilLogDate1
            
            llLockRecCode = gCreateLockRec(hmRlf, "S", "S", 65536 * ilToVefCode + gDateValue(slDate), False, slUserName)
            If llLockRecCode > 0 Then
                If tgMVef(ilVefIndex).sType <> "G" Then
                    tmSsfSrchKey.iType = ilType
                    tmSsfSrchKey.iVefCode = ilToVefCode
                    tmSsfSrchKey.iDate(0) = ilLogDate0
                    tmSsfSrchKey.iDate(1) = ilLogDate1
                    tmSsfSrchKey.iStartTime(0) = 0
                    tmSsfSrchKey.iStartTime(1) = 0
                    ilRet = gSSFGetEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                Else
                    tmSsfSrchKey2.iVefCode = ilToVefCode
                    tmSsfSrchKey2.iDate(0) = ilLogDate0
                    tmSsfSrchKey2.iDate(1) = ilLogDate1
                    ilRet = gSSFGetEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                    ilType = tmSsf.iType
                End If
                Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilToVefCode) And (tmSsf.iDate(0) = ilLogDate0) And (tmSsf.iDate(1) = ilLogDate1)
                    ilRet = gSSFGetPosition(hmSsf, llSsfRecPos)
                    ilEvt = 1
                    Do While ilEvt <= tmSsf.iCount
                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                        'If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                        If (tmAvail.iRecType = 2) Or (tmAvail.iRecType = 8) Or (tmAvail.iRecType = 9) Then
                            gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                            If llTime >= llETime Then
                                Exit Do
                            End If
                            If llTime >= llSTime Then
                                ilAvEvt = ilEvt
                                'Test if within selected times
                                ilLen = tmAvail.iLen
                                ilUnits = tmAvail.iAvInfo And &H1F
                                For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                    ilEvt = ilEvt + 1
                                   LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                    If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                        ilUnits = ilUnits - 1
                                        ilLen = ilLen - (tmSpot.iPosLen And &HFFF)
                                    End If
                                Next ilSpot
                                ''If (ilUnits > 0) And (tmClf.iLen <= ilLen) Then
                                'If (ilUnits > 0) And (tmClf.iLen <= ilLen) And (((ilAnfCode = 0) And (tmAvail.ianfCode = 0)) Or (tmAvail.ianfCode = ilAnfCode)) Then
                                If (ilUnits > 0) And (tmClf.iLen <= ilLen) And ((ilAnfCode = 0) Or (tmAvail.ianfCode = ilAnfCode)) Then
                                    If (tmClf.iLen = 30) Or (tmClf.iLen = 60) Or (tmClf.iLen = tmAvail.iLen) Then
                                        If (tmAvail.iRecType = 2) Or ((tmAvail.iRecType = 8) And (tmChf.sType = "S")) Or ((tmAvail.iRecType = 9) And (tmChf.sType = "M")) Then
                                            If Not mAnyConflicts(ilAvEvt, tmChf.iAdfCode, tmChf.iMnfComp(0), tmChf.iMnfComp(1), tmChf.sType, ilDay) Then
                                                ilRet = btrBeginTrans(hmSdf, 1000)
                                                If ilRet <> BTRV_ERR_NONE Then
                                                    If ilRet >= 30000 Then
                                                        igBtrError = csiHandleValue(0, 7)
                                                    Else
                                                        igBtrError = ilRet
                                                    End If
                                                    sgErrLoc = "mBonusSpot"
                                                    ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                                    mBonusSpot = False
                                                    Exit Function
                                                End If
                                                ilRet = mMakeUnschSpot("B", tmChf.lCode, tmChf.iAdfCode, tmClf.iLine, llDate, ilToVefCode, tmClf.iLen)
                                                If ilRet <> BTRV_ERR_NONE Then
                                                    If ilRet >= 30000 Then
                                                        igBtrError = csiHandleValue(0, 7)
                                                    Else
                                                        igBtrError = ilRet
                                                    End If
                                                    sgErrLoc = "mMakeUnschSpot"
                                                    ilCRet = btrAbortTrans(hmSdf)
                                                    ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                                    mBonusSpot = False
                                                    Exit Function
                                                End If
                                                slSchStatus = "O"
                                                If tmChf.sType = "S" Then
                                                    ilBkQH = 1060
                                                ElseIf tmChf.sType = "M" Then
                                                    ilBkQH = 1050
                                                ElseIf tmChf.sType = "Q" Then
                                                    ilBkQH = 1030
                                                Else
                                                    ilBkQH = 1045
                                                End If
                                                '3/1/01- JIm make bonus priority one so that we max the number scheduled
                                                ilBkQH = 1
                                                ilPriceLevel = 0
                                                tmSdfSrchKey3.lCode = tmSdf.lCode
                                                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                If ilRet = BTRV_ERR_NONE Then
                                                    ilRet = btrGetPosition(hmSdf, llSdfRecPos)
                                                End If
                                                If ilRet = BTRV_ERR_NONE Then
                                                    'BookSpot Re-Read Ssf so handle is correct
                                                    ilRet = gBookSpot(slSchStatus, hmSdf, tmSdf, llSdfRecPos, ilBkQH, hmSsf, tmSsf, llSsfRecPos, ilAvEvt, ilPosition, tmChf, tmClf, tmRdf, imVpfIndex, hmSmf, tmSmf, hmClf, hmCrf, ilPriceLevel, False)
                                                    If ilRet Then
                                                        'tmSdfSrchKey3.lCode = tmSdf.lCode
                                                        'ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                        'If ilRet = BTRV_ERR_NONE Then
                                                        '    tmSmfSrchKey2.lCode = tmSdf.lCode
                                                        '    ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                                        '    If ilRet = BTRV_ERR_NONE Then
                                                        '        tmSmf.lTrackID = llTrackID
                                                        '        tmSmf.lRefTrackID = llRefTrackID
                                                        '        ilRet = btrUpdate(hmSmf, tmSmf, imSmfRecLen)
                                                        '    End If
                                                        'End If
                                                        ilRet = mUpdateSMF(llTrackID, llEnteredDate, llRefTrackID, 0, 0, llTransGpID, 0)
                                                        If ilRet <> BTRV_ERR_NONE Then
                                                            If ilRet >= 30000 Then
                                                                igBtrError = csiHandleValue(0, 7)
                                                            Else
                                                                igBtrError = ilRet
                                                            End If
                                                            sgErrLoc = "mBonusSpot"
                                                            ilCRet = btrAbortTrans(hmSdf)
                                                            ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                                            mBonusSpot = False
                                                            Exit Function
                                                        End If
                                                        ''mMakeTracer llSdfRecPos, "S"
                                                        'ilRet = gMakeTracer(hmSdf, tmSdf, llSdfRecPos, hmStf, lmLastLogDate, "S", "M", tmSdf.iRotNo)
                                                        'If ilRet Then
                                                        '
                                                            mBonusSpot = True
                                                            ilRet = btrEndTrans(hmSdf)
                                                            ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                                            Exit Function
                                                        'End If
                                                    End If
                                                End If
                                                mBonusSpot = False
                                                ilCRet = btrAbortTrans(hmSdf)
                                                ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                ilEvt = ilEvt + tmAvail.iNoSpotsThis
                            End If
                        End If
                        ilEvt = ilEvt + 1
                        DoEvents
                    Loop
                    imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
            End If
        End If
    Next ilDate
    igBtrError = 0
    sgErrLoc = "mBonusSpot, no avails available for bonus"
    mBonusSpot = False
    Exit Function
End Function
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
    slStr = edcSDate.Text
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
'*      Procedure Name:mChkVeh                         *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Check Line vehicles             *
'*                                                     *
'*                                                     *
'*******************************************************
Private Function mChkVeh() As Integer
    Dim slFromFile As String
    Dim slFromCntr As String
    Dim slFromLine As String
    Dim slFromMove As String
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim ilFound As Integer
    Dim ilVefCode As Integer
    Dim ilAnfFound As Integer
    Dim ilAnfCode As Integer
    Dim ilRdfFound As Integer
    Dim slLine As String
    Dim slMove As String
    Dim ilEof As Integer
    Dim ilTst As Integer
    Dim ilAdd As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilIndex As Integer
    Dim llLnStartTime As Long
    Dim llLnEndTime As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilDay As Integer
    Dim ilPass As Integer

    ReDim imMS12MVefCode(0 To 0) As Integer

    pbcChkVeh.Move (ImptCntr.Width - pbcChkVeh.Width) / 2, (ImptCntr.Height - pbcChkVeh.Height) / 2
'    slFromFile = edcFrom.Text
    For ilPass = 0 To UBound(tmSortCode) - 1 Step 1
        ilRet = 0
        On Error GoTo mChkVehErr:
        slStr = tmSortCode(ilPass).sKey
        ilRet = gParseItem(slStr, 2, "\", slFromFile)
        slFromFile = sgImportPath & slFromFile
        slFromCntr = slFromFile
        slFromLine = slFromFile
        ilPos = Len(slFromLine) - 1
        Do While Mid$(slFromLine, ilPos, 1) <> "\"
            ilPos = ilPos - 1
            If ilPos <= 0 Then
                Exit Do
            End If
        Loop
        smSeqNo = Mid$(slFromCntr, ilPos + 6, 3)
        If ilPass = 0 Then
            If (mGetSeqNo() + 1) <> Val(smSeqNo) Then
                MsgBox "Import out of Sequence, Last Import was Sequence Number" & str$(mGetSeqNo()), vbOkOnly + vbCritical + vbApplicationModal, "Sequence Error"
                mChkVeh = 1
                Exit Function
            End If
        End If
        ilPos = ilPos + 1
        Mid$(slFromLine, ilPos, 1) = "L"
        slFromMove = slFromFile
        ilPos = Len(slFromMove) - 1
        Do While Mid$(slFromMove, ilPos, 1) <> "\"
            ilPos = ilPos - 1
            If ilPos <= 0 Then
                Exit Do
            End If
        Loop
        ilPos = ilPos + 1
        Mid$(slFromMove, ilPos, 1) = "M"
        ilRet = 0
        hmLine = FreeFile
        Open slFromLine For Input Access Read As hmLine
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slFromLine & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mChkVeh = 1
            Exit Function
        End If
        ilRet = 0
        hmMove = FreeFile
        Open slFromMove For Input Access Read As hmMove
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slFromMove & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mChkVeh = 1
            Exit Function
        End If
        If ilPass = 0 Then
            ilRet = 0
            On Error GoTo mChkVehFileErr:
            slToFile = sgImportPath & "VehChk.Txt"
            slDateTime = FileDateTime(slToFile)
            If ilRet = 0 Then
                Kill slToFile
                On Error GoTo 0
                ilRet = 0
                On Error GoTo mChkVehFileErr:
                hmMsg = FreeFile
                Open slToFile For Output As hmMsg
                If ilRet <> 0 Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                    mChkVeh = 1
                    Exit Function
                End If
            Else
                On Error GoTo 0
                ilRet = 0
                On Error GoTo mChkVehFileErr:
                hmMsg = FreeFile
                Open slToFile For Output As hmMsg
                If ilRet <> 0 Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                    mChkVeh = 1
                    Exit Function
                End If
            End If
        End If
        Screen.MousePointer = vbHourglass
        pbcChkVeh.Visible = True
        DoEvents
        ilEof = False
        mChkVeh = 0
        If ilPass = 0 Then
            lbcErrors.Clear
            lbcErrors.Visible = True
            Print #hmMsg, "** Vehicle Test: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        End If
        Print #hmMsg, " "
        Print #hmMsg, "** Checking: " & slFromMove & " and " & slFromLine & " **"
        Do
            'Get Lines
            ilRet = 0
            On Error GoTo mChkVehErr:
            Line Input #hmLine, slLine
            On Error GoTo 0
            If ilRet <> 0 Then
                Exit Do
            End If
            If Trim$(slLine) <> "" Then
                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                    Exit Do
                End If
            End If
            DoEvents
            If Trim$(slLine) <> "" Then
                slLine = mFilter(slLine)
                gParseCDFields slLine, True, smFieldValues()    'Change case
                If Trim$(smFieldValues(4)) <> "P" Then
                    ilFound = False
                    For ilLoop = LBound(tgMVef) To UBound(tgMVef) Step 1
                        If (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S") Then
                            If StrComp(Trim$(tgMVef(ilLoop).sName), Trim$(smFieldValues(5)), 1) = 0 Then
                                ilVefCode = tgMVef(ilLoop).iCode
                                ilFound = True
                                ilAdd = True
                                For ilTst = LBound(imMS12MVefCode) To UBound(imMS12MVefCode) - 1 Step 1
                                    If tgMVef(ilLoop).iCode = imMS12MVefCode(ilTst) Then
                                        ilAdd = False
                                        Exit For
                                    End If
                                Next ilTst
                                If ilAdd Then
                                    imMS12MVefCode(UBound(imMS12MVefCode)) = tgMVef(ilLoop).iCode
                                    ReDim Preserve imMS12MVefCode(0 To UBound(imMS12MVefCode) + 1) As Integer
                                End If
                                Exit For
                            End If
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        slStr = "Vehicle " & Trim$(smFieldValues(5)) & " not found"
                        Print #hmMsg, slStr
                        lbcErrors.AddItem "Missing: " & Trim$(smFieldValues(5))
                        mChkVeh = 2
                    End If
                    ilAnfFound = False
                    If Trim$(smFieldValues(18)) <> "" Then
                        For ilLoop = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1 Step 1
                            If StrComp(Trim$(tgAvailAnf(ilLoop).sName), Trim$(smFieldValues(18)), 1) = 0 Then
                                ilAnfFound = True
                                ilAnfCode = tgAvailAnf(ilLoop).iCode
                                Exit For
                            End If
                        Next ilLoop
                        If Not ilAnfFound Then
                            slStr = "Avail Name " & Trim$(smFieldValues(18)) & " not found in " & Trim$(smFieldValues(5))
                            Print #hmMsg, slStr
                            lbcErrors.AddItem "Avail Name Missing: " & Trim$(smFieldValues(18))
                            mChkVeh = 2
                        End If
                    End If
                    If ilFound And ilAnfFound Then
                        llLnStartTime = gTimeToLong(smFieldValues(6), False)
                        llLnEndTime = gTimeToLong(smFieldValues(7), True)
                        ilRdfFound = False
                        For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                            If ilVefCode = tgMRif(llRif).iVefCode Then
                                'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                                '    If (tgMRif(ilRif).iRdfcode = tgMRdf(ilRdf).iCode) Then
                                    ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfcode)
                                    If ilRdf <> -1 Then
                                        For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                            If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                                gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), False, llStartTime
                                                gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilIndex), tgMRdf(ilRdf).iEndTime(1, ilIndex), True, llEndTime
                                                'Ignoring day test because allowed days are obtained from Flight records not Line records
                                                'ilOk = True
                                                'For ilDay = 1 To 7 Step 1
                                                '    If tgMRdf(ilRdf).sWkDays(ilIndex, ilDay) <> "Y" Then
                                                '        ilOk = False
                                                '    End If
                                                'Next ilDay
                                                'If (llStartTime = 0) And (llEndTime = 0) And (ilOk = True) And (tgMRdf(ilRdf).sInOut = "N") And (tgMRdf(ilRdf).sState = "A") Then
                                                '    Exit For
                                                'Else
                                                '    ilOk = False
                                                'End If
                                                If (tgMRdf(ilRdf).sInOut = "I") And (tgMRdf(ilRdf).sState = "A") And (tgMRdf(ilRdf).ianfCode = ilAnfCode) Then
                                                    If (llLnStartTime >= llStartTime) And (llLnEndTime <= llEndTime) Then
                                                        ilRdfFound = True
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        Next ilIndex
                                        If ilRdfFound Then
                                '            Exit For
                                        End If
                                    End If
                                'Next ilRdf
                                If ilRdfFound Then
                                    Exit For
                                End If
                            End If
                        Next llRif
                        If Not ilRdfFound Then
                            slStr = "Daypart with Avail Name " & Trim$(smFieldValues(18)) & " not found in " & Trim$(smFieldValues(5))
                            Print #hmMsg, slStr
                            lbcErrors.AddItem "Daypart with Avail Name Missing: " & Trim$(smFieldValues(18))
                            mChkVeh = 2
                        End If
                    End If
                End If
            End If
        Loop Until ilEof
        If rbcChkVeh(0).Value Then
            Do
                'Get Lines
                ilRet = 0
                On Error GoTo mChkVehErr:
                Line Input #hmMove, slMove
                On Error GoTo 0
                If ilRet <> 0 Then
                    Exit Do
                End If
                If Trim$(slMove) <> "" Then
                    If (Asc(slMove) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                        Exit Do
                    End If
                End If
                DoEvents
                If Trim$(slMove) <> "" Then
                    slMove = mFilter(slMove)
                    gParseCDFields slMove, True, smFieldValues()    'Change case
                    Select Case UCase$(smFieldValues(15))
                        Case "C", "S", "M", "U", "V", "D"
                            ilFound = False
                            For ilLoop = LBound(tgMVef) To UBound(tgMVef) Step 1
                                If (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S") Then
                                    If StrComp(Trim$(tgMVef(ilLoop).sName), Trim$(smFieldValues(7)), 1) = 0 Then
                                        ilFound = True
                                        ilAdd = True
                                        For ilTst = LBound(imMS12MVefCode) To UBound(imMS12MVefCode) - 1 Step 1
                                            If tgMVef(ilLoop).iCode = imMS12MVefCode(ilTst) Then
                                                ilAdd = False
                                                Exit For
                                            End If
                                        Next ilTst
                                        If ilAdd Then
                                            imMS12MVefCode(UBound(imMS12MVefCode)) = tgMVef(ilLoop).iCode
                                            ReDim Preserve imMS12MVefCode(0 To UBound(imMS12MVefCode) + 1) As Integer
                                        End If
                                        Exit For
                                    End If
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                slStr = "Vehicle " & Trim$(smFieldValues(7)) & " not found"
                                Print #hmMsg, slStr
                                lbcErrors.AddItem "Missing: " & Trim$(smFieldValues(7))
                                mChkVeh = 2
                            End If
                    End Select
                    Select Case UCase$(smFieldValues(15))
                        Case "C", "S", "M", "B", "U", "V"
                            ilFound = False
                            For ilLoop = LBound(tgMVef) To UBound(tgMVef) Step 1
                                If (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S") Then
                                    If StrComp(Trim$(tgMVef(ilLoop).sName), Trim$(smFieldValues(11)), 1) = 0 Then
                                        ilVefCode = tgMVef(ilLoop).iCode
                                        ilFound = True
                                        ilAdd = True
                                        For ilTst = LBound(imMS12MVefCode) To UBound(imMS12MVefCode) - 1 Step 1
                                            If tgMVef(ilLoop).iCode = imMS12MVefCode(ilTst) Then
                                                ilAdd = False
                                                Exit For
                                            End If
                                        Next ilTst
                                        If ilAdd Then
                                            imMS12MVefCode(UBound(imMS12MVefCode)) = tgMVef(ilLoop).iCode
                                            ReDim Preserve imMS12MVefCode(0 To UBound(imMS12MVefCode) + 1) As Integer
                                        End If
                                        Exit For
                                    End If
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                slStr = "Vehicle " & Trim$(smFieldValues(11)) & " not found"
                                Print #hmMsg, slStr
                                lbcErrors.AddItem "Missing: " & Trim$(smFieldValues(11))
                                mChkVeh = 2
                            End If
                    End Select
    
                    'From avail name is not used
                    'ilAnfFound = False
                    'If Trim$(smFieldValues(23)) <> "" Then
                    '    For ilLoop = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1 Step 1
                    '        If StrComp(Trim$(tgAvailAnf(ilLoop).sName), Trim$(smFieldValues(23)), 1) = 0 Then
                    '            ilAnfFound = True
                    '            Exit For
                    '        End If
                    '    Next ilLoop
                    '    If Not ilAnfFound Then
                    '        slStr = "Avail Name " & Trim$(smFieldValues(23)) & " not found"
                    '        Print #hmMsg, slStr
                    '        lbcErrors.AddItem "Missing: " & Trim$(smFieldValues(23))
                    '        mChkVeh = 2
                    '    End If
                    'End If
    
                    ilAnfFound = False
                    If Trim$(smFieldValues(24)) <> "" Then
                        For ilLoop = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1 Step 1
                            If StrComp(Trim$(tgAvailAnf(ilLoop).sName), Trim$(smFieldValues(24)), 1) = 0 Then
                                ilAnfCode = tgAvailAnf(ilLoop).iCode
                                ilAnfFound = True
                                Exit For
                            End If
                        Next ilLoop
                        If Not ilAnfFound Then
                            slStr = "Avail Name " & Trim$(smFieldValues(24)) & " not found"
                            Print #hmMsg, slStr
                            lbcErrors.AddItem "Missing: " & Trim$(smFieldValues(24))
                            mChkVeh = 2
                        End If
                    End If
                    If ilFound And ilAnfFound Then
                        Select Case UCase$(smFieldValues(15))
                            Case "C", "S", "M", "B", "U", "V"
                                If smFieldValues(13) <> "" Then
                                    llLnStartTime = gTimeToLong(smFieldValues(13), False)
                                Else
                                    llLnStartTime = -1
                                End If
                                If smFieldValues(14) <> "" Then
                                    llLnEndTime = gTimeToLong(smFieldValues(14), True)
                                Else
                                    llLnEndTime = -1
                                End If
                                If llLnStartTime = -1 Then
                                    llLnStartTime = 0
                                    llLnEndTime = 86400
                                Else
                                    llLnStartTime = llLnStartTime  'CLng(gTimeToCurrency(slToTime, False))
                                    llLnEndTime = llLnStartTime
                                End If
                                ilDay = gWeekDayLong(gDateValue(smFieldValues(12)))
                                ilRdfFound = False
                                For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                                    If ilVefCode = tgMRif(llRif).iVefCode Then
                                        'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                                        '    If (tgMRif(ilRif).iRdfcode = tgMRdf(ilRdf).iCode) Then
                                            ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfcode)
                                            If ilRdf <> -1 Then
                                                For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                                    If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                                        gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), False, llStartTime
                                                        gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilIndex), tgMRdf(ilRdf).iEndTime(1, ilIndex), True, llEndTime
                                                        'ilOk = True
                                                        'For ilDay = 1 To 7 Step 1
                                                        '    If tgMRdf(ilRdf).sWkDays(ilIndex, ilDay) <> "Y" Then
                                                        '        ilOk = False
                                                        '    End If
                                                        'Next ilDay
                                                        'If (llStartTime = 0) And (llEndTime = 0) And (ilOk = True) And (tgMRdf(ilRdf).sInOut = "N") And (tgMRdf(ilRdf).sState = "A") Then
                                                        '    Exit For
                                                        'Else
                                                        '    ilOk = False
                                                        'End If
                                                        If (tgMRdf(ilRdf).sWkDays(ilIndex, ilDay + 1) = "Y") And (tgMRdf(ilRdf).sInOut = "I") And (tgMRdf(ilRdf).sState = "A") And (tgMRdf(ilRdf).ianfCode = ilAnfCode) Then
                                                            If (llLnStartTime >= llStartTime) And (llLnEndTime <= llEndTime) Then
                                                                ilRdfFound = True
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                Next ilIndex
                                                If ilRdfFound Then
                                        '            Exit For
                                                End If
                                            End If
                                        'Next ilRdf
                                        If ilRdfFound Then
                                            Exit For
                                        End If
                                    End If
                                Next llRif
                                If Not ilRdfFound Then
                                    'Look for a base daypart that can be used- mGetRdf has two passes (the second pass looked for a base daypart)
                                    For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                                        'If ilVefCode = tgMRif(ilRif).iVefCode Then
                                            'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                                            '    If (tgMRif(ilRif).iRdfcode = tgMRdf(ilRdf).iCode) And (tgMRdf(ilRdf).sBase = "Y") Then
                                                ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfcode)
                                                If ilRdf <> -1 Then
                                                    If tgMRdf(ilRdf).sBase = "Y" Then
                                                        For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                                            If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                                                gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), False, llStartTime
                                                                gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilIndex), tgMRdf(ilRdf).iEndTime(1, ilIndex), True, llEndTime
                                                                'ilOk = True
                                                                'For ilDay = 1 To 7 Step 1
                                                                '    If tgMRdf(ilRdf).sWkDays(ilIndex, ilDay) <> "Y" Then
                                                                '        ilOk = False
                                                                '    End If
                                                                'Next ilDay
                                                                'If (llStartTime = 0) And (llEndTime = 0) And (ilOk = True) And (tgMRdf(ilRdf).sInOut = "N") And (tgMRdf(ilRdf).sState = "A") Then
                                                                '    Exit For
                                                                'Else
                                                                '    ilOk = False
                                                                'End If
                                                                If (tgMRdf(ilRdf).sWkDays(ilIndex, ilDay + 1) = "Y") And (tgMRdf(ilRdf).sInOut = "I") And (tgMRdf(ilRdf).sState = "A") And (tgMRdf(ilRdf).ianfCode = ilAnfCode) Then
                                                                    If (llLnStartTime >= llStartTime) And (llLnEndTime <= llEndTime) Then
                                                                        ilRdfFound = True
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            End If
                                                        Next ilIndex
                                                    End If
                                                    If ilRdfFound Then
                                            '            Exit For
                                                    End If
                                                End If
                                            'Next ilRdf
                                            If ilRdfFound Then
                                                Exit For
                                            End If
                                        'End If
                                    Next llRif
                                End If
                                If Not ilRdfFound Then
                                    slStr = "Daypart with Avail Name " & Trim$(smFieldValues(24)) & " not found in " & Trim$(smFieldValues(11))
                                    Print #hmMsg, slStr
                                    lbcErrors.AddItem "Daypart with Avail Name Missing: " & Trim$(smFieldValues(24))
                                    mChkVeh = 2
                                End If
                        End Select
                    End If
                End If
            Loop Until ilEof
        End If
        If Not mTestMSU12M12M() Then
            mChkVeh = 2
        End If
        Close hmLine
        Close hmMove
    Next ilPass
    Print #hmMsg, "** Completed Vehicle Test: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    Close #hmMsg
    pbcChkVeh.Visible = False
    Exit Function
mChkVehFileErr:
    ilRet = Err.Number
    Resume Next
mChkVehErr:
    ilRet = Err.Number
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mConvCHF                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Convert CHF                    *
'*                                                     *
'*******************************************************
Private Function mConvCHF() As Integer
    Dim slFromFile As String
    Dim slFromCntr As String
    Dim slFromLine As String
    Dim slFromFlight As String
    Dim slFromMove As String
    Dim ilRet As Integer
    Dim slHeader As String
    Dim slLine As String
    Dim slFlight As String
    Dim slMove As String
    Dim ilEof As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilFirstHd As Integer
    Dim llPercent As Long
    Dim llCount As Long
    Dim ilErrorCount As Integer
    Dim ilLnIndex As Integer
    Dim llCntrRevDate As Long
    Dim llCntrRevTime As Long
    Dim llPrevCntrRevDate As Long
    Dim ilFound As Integer
    Dim ilPass As Integer
'    slFromFile = edcFrom.Text
    For ilPass = 0 To UBound(tmSortCode) - 1 Step 1
        ilRet = 0
        On Error GoTo mConvCHFErr:
        slStr = tmSortCode(ilPass).sKey
        ilRet = gParseItem(slStr, 2, "\", slFromFile)
        slFromFile = sgImportPath & slFromFile
        slFromCntr = slFromFile
        slFromLine = slFromFile
        ilPos = Len(slFromLine) - 1
        Do While Mid$(slFromLine, ilPos, 1) <> "\"
            ilPos = ilPos - 1
            If ilPos <= 0 Then
                Exit Do
            End If
        Loop
        smSeqNo = Mid$(slFromCntr, ilPos + 6, 3)
        If (mGetSeqNo() + 1) <> Val(smSeqNo) Then
            MsgBox "Import out of Sequence, Last Import was Sequence Number" & str$(mGetSeqNo()), vbOkOnly + vbCritical + vbApplicationModal, "Sequence Error"
            'edcFrom.SetFocus
            cmcCancel.SetFocus
            mConvCHF = False
            Exit Function
        Else
            imProcMove = True
            If imCurSeqNo = Val(smSeqNo) Then
                imProcMove = False
            End If
        End If
        ilPos = ilPos + 1
        Mid$(slFromLine, ilPos, 1) = "L"
        slFromFlight = slFromFile
        ilPos = Len(slFromFlight) - 1
        Do While Mid$(slFromFlight, ilPos, 1) <> "\"
            ilPos = ilPos - 1
            If ilPos <= 0 Then
                Exit Do
            End If
        Loop
        ilPos = ilPos + 1
        Mid$(slFromFlight, ilPos, 1) = "F"
        slFromMove = slFromFile
        ilPos = Len(slFromMove) - 1
        Do While Mid$(slFromMove, ilPos, 1) <> "\"
            ilPos = ilPos - 1
            If ilPos <= 0 Then
                Exit Do
            End If
        Loop
        ilPos = ilPos + 1
        Mid$(slFromMove, ilPos, 1) = "M"
        hmCntr = FreeFile
        Open slFromCntr For Input Access Read As hmCntr
        If ilRet <> 0 Then
            MsgBox "Open " & slFromFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
    '        edcFrom.SetFocus
            cmcCancel.SetFocus
            mConvCHF = False
            Exit Function
        End If
        hmLine = FreeFile
        Open slFromLine For Input Access Read As hmLine
        If ilRet <> 0 Then
            Close hmCntr
            MsgBox "Open " & "L" & Mid(slFromFile, 2) & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
    '        edcFrom.SetFocus
            cmcCancel.SetFocus
            mConvCHF = False
            Exit Function
        End If
        hmFlight = FreeFile
        Open slFromFlight For Input Access Read As hmFlight
        If ilRet <> 0 Then
            Close hmCntr
            Close hmLine
            MsgBox "Open " & "F" & Mid(slFromFile, 2) & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
    '        edcFrom.SetFocus
            cmcCancel.SetFocus
            mConvCHF = False
            Exit Function
        End If
        hmMove = FreeFile
        Open slFromMove For Input Access Read As hmMove
        If ilRet <> 0 Then
            Close hmCntr
            Close hmLine
            Close hmFlight
            MsgBox "Open " & "M" & Mid(slFromFile, 2) & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
    '        edcFrom.SetFocus
            cmcCancel.SetFocus
            mConvCHF = False
            Exit Function
        End If
        smRptName = slFromFile
        ilPos = InStr(smRptName, ".")
        If ilPos > 0 Then
            smRptName = Left$(smRptName, ilPos - 1)
        End If
        ilPos = Len(smRptName) - 1
        Do While Mid$(smRptName, ilPos, 1) <> "\"
            ilPos = ilPos - 1
            If ilPos <= 0 Then
                Exit Do
            End If
        Loop
        smRptName = Mid$(smRptName, ilPos + 1) & ".txt"
        If Not mOpenMsgFile() Then
            Close hmCntr
            Close hmLine
            Close hmFlight
            Close hmMove
            cmcCancel.SetFocus
            mConvCHF = False
            Exit Function
        End If
        Print #hmMsg, "** Import Contracts from " & slFromCntr
        Print #hmMsg, "** Import MG's from " & slFromMove
        ReDim tgMoveRec(0 To 0) As MGMOVEREC
        lmRecCount = 0
        llCount = 0
        lmTotalNoBytes = LOF(hmCntr) \ 128  'The Loc returns current position \128
        lmTotalNoBytes = lmTotalNoBytes + LOF(hmMove) \ 128 'The Loc returns current position \128
        ilFirstHd = True
        ilEof = False
        lacCntr(0).Visible = True
        lacCntr(1).Visible = True
        lacCntr(2).Visible = False
        lacCntr(3).Visible = False
        DoEvents
        If imTerminate Then
            Print #hmMsg, "** Import Terminated " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Close hmMove
            Close hmFlight
            Close hmLine
            Close hmCntr
            mTerminate
            mConvCHF = False
            Exit Function
        End If
        lbcErrors.Clear
        lbcErrors.Visible = True
        plcGauge.Value = llPercent
        lacCount.Caption = ""
        lacErrors.Caption = ""
        imFlightError = False
        Screen.MousePointer = vbHourglass
        ReDim lmPrevSdfCreated(0 To 0) As Long
        ReDim lmPrevMoveDeleted(0 To 0) As Long
        'mClearIcf
        'tmIcf.sType = "0"
        'tmIcf.sAdvtName = smRptName
        'tmIcf.sStatus = "A"
        'tmIcf.sErrorMess = ""
        'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
        'mClearIcf
        slHeader = ""
        slLine = ""
        slFlight = ""
        slMove = ""
        lmMGRecNo = 1
        llPrevCntrRevDate = 0
        Do
            If Trim$(slHeader) = "" Then
                ilRet = 0
                On Error GoTo mConvCHFErr:
                Line Input #hmCntr, slHeader
                On Error GoTo 0
                If ilRet <> 0 Then
                    ilEof = True
                    Exit Do
                End If
            End If
            On Error GoTo 0
            If Trim$(slHeader) <> "" Then
                If (Asc(slHeader) = 26) Or (ilRet <> 0) Or (ilEof) Then    'Ctrl Z
                    ilEof = True
                    Exit Do
                Else
                    DoEvents
                    'Test if whole header line imported
                    'Last Field must be: BOOKED or CANCEL_H or CANCEL_O or HOLD or REV_HOLD or REV_ORDER
                    Do
                        ilFound = False
                        slStr = UCase$(right$(Trim$(slHeader), 4))
                        Select Case slStr
                            Case "OKED"
                                ilFound = True
                            Case "EL_H"
                                ilFound = True
                            Case "EL_O"
                                ilFound = True
                            Case "HOLD"
                                ilFound = True
                            Case "RDER"
                                ilFound = True
                        End Select
                        If Not ilFound Then
                            ilRet = 0
                            On Error GoTo mConvCHFErr:
                            Line Input #hmCntr, slStr
                            On Error GoTo 0
                            If ilRet <> 0 Then
                                ilEof = True
                                Exit Do
                            End If
                            slHeader = slHeader & slStr
                        Else
                            Exit Do
                        End If
                    Loop While Not ilFound
                    gGetSyncDateTime smSyncDate, smSyncTime
        '           ilRet = mParseItem(slInputStr, ilItemNo, slDelimiter, slOutStr)
                    'gParseCDFields slHeader, False, smFieldValues()
                    slHeader = mFilter(slHeader)
                    gParseCDFields slHeader, True, smFieldValues()    'Change case
                    llCntrRevDate = gDateValue(smFieldValues(3))
                    llCntrRevTime = gTimeToLong(smFieldValues(4), False)
                    If llPrevCntrRevDate <> llCntrRevDate Then
                        ilRet = mReadMoveRecord(llCntrRevDate)
                        llPrevCntrRevDate = llCntrRevDate
                    End If
                    ''Determine if Contract or Move should be processed
                    'Do
                    '    ilDoCntr = False
                    '    If Len(smMoveValues(1)) <= 0 Then
                    '        Exit Do
                    '    End If
                    '    If gDateValue(smFieldValues(3)) < tgMoveRec.lCreatedDate Then
                    '        Exit Do
                    '    ElseIf gDateValue(smFieldValues(3)) = tgMoveRec.lCreatedDate Then
                    '        If gTimeToLong(smFieldValues(4), False) <= gTimeToLong(smMoveValues(20), False) Then
                    '            Exit Do
                    '        Else
                    '            tmMoveRec = tgMoveRec(ilLoop)
                    '            ilRet = mImportMG()
                    '            lmRecCount = lmRecCount + 1
                    '        End If
                    '    Else
                    '        tmMoveRec = tgMoveRec(ilLoop)
                    '        ilRet = mImportMG()
                    '        lmRecCount = lmRecCount + 1
                    '    End If
                    'Loop While Not ilDoCntr
                    For ilLoop = LBound(tgMoveRec) To UBound(tgMoveRec) - 1 Step 1
                        'Test if Contract exist because
                        'contract entered as hold, spot moved, then contract changed to Order.
                        'The hold image is not retained (i.e. only receive the Order image after the move operation)
                        If (tgMoveRec(ilLoop).iStatus = 0) And (tgMoveRec(ilLoop).lCreatedDate <= llCntrRevDate) Then
                            tmMoveRec = tgMoveRec(ilLoop)
                            tmChfSrchKey.lCntrNo = tmMoveRec.lCntrNo
                            tmChfSrchKey.iCntRevNo = 32000
                            tmChfSrchKey.iPropVer = 32000
                            ilRet = btrGetGreaterOrEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)
                            If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmMoveRec.lCntrNo) Then   'Contract found
                                If (tmMoveRec.lCreatedDate < llCntrRevDate) Or ((tmMoveRec.lCreatedDate = llCntrRevDate) And (tmMoveRec.lCreatedTime <= llCntrRevTime)) Then
                                    'Process undo events
                                    If (tmMoveRec.sOper = "D") Or (tmMoveRec.sOper = "U") Or (tmMoveRec.sOper = "V") Or ((tmMoveRec.sOper = "M") And (tmMoveRec.sOpResult <> "R")) Then
                                        If tmMoveRec.lCreatedTime <> 0 Then
                                            imShowMGErr = True
                                            ilRet = mImportMG()
                                            If Not ilRet Then
                                                mConvCHF = False
                                                Exit Function
                                            End If
                                            tgMoveRec(ilLoop).iStatus = 1
                                            ilRet = mSetCurSeqNo(Val(smSeqNo), tgMoveRec(ilLoop).lMoveID)
                                        Else
                                            imShowMGErr = False
                                            ilRet = mImportMG()
                                            If ilRet Then
                                                tgMoveRec(ilLoop).iStatus = 1
                                                ilRet = mSetCurSeqNo(Val(smSeqNo), tgMoveRec(ilLoop).lMoveID)
                                            Else
                                                mConvCHF = False
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                                If tgMoveRec(ilLoop).iStatus = 0 Then
                                    If (tmMoveRec.lCreatedDate < llCntrRevDate) Or ((tmMoveRec.lCreatedDate = llCntrRevDate) And (tmMoveRec.lCreatedTime <= llCntrRevTime) And (tmMoveRec.lCreatedTime <> 0)) Then
                                        'Process undo events
                                        imShowMGErr = True
                                        ilRet = mImportMG()
                                        If Not ilRet Then
                                            mConvCHF = False
                                            Exit Function
                                        End If
                                        tgMoveRec(ilLoop).iStatus = 1
                                        ilRet = mSetCurSeqNo(Val(smSeqNo), tgMoveRec(ilLoop).lMoveID)
                                    End If
                                End If
                            End If
                        End If
                    Next ilLoop
                    imPorO = 1
                    smNameMsg = "Order"
                    lacCntr(0).Caption = "Processing Order #"
                    lacCntr(1).Caption = smFieldValues(1) & " Rev. " & smFieldValues(2)
                    lacCntr(2).Caption = "Line #"
                    lacCntr(3).Caption = ""
                    DoEvents
                    ReDim lgReschSdfCode(1 To 1) As Long
                    mProcessHeader
                    Do
                        'Get Lines
                        If Trim$(slLine) = "" Then
                            ilRet = 0
                            On Error GoTo mConvCHFErr:
                            Line Input #hmLine, slLine
                            On Error GoTo 0
                            If ilRet <> 0 Then
                                slHeader = ""
                                slLine = ""
                                Exit Do
                            End If
                            If Trim$(slLine) <> "" Then
                                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                                    slHeader = "" 'slLine
                                    slLine = ""
                                    'ilEof = True
                                    Exit Do
                                End If
                            End If
                        ElseIf ilEof Then
                            Exit Do
                        End If
                        DoEvents
                        If Trim$(slLine) <> "" Then
                            slLine = mFilter(slLine)
                            gParseCDFields slLine, True, smFieldValues()    'Change case
                            If Val(smFieldValues(1)) <> tgChfImpt.lCntrNo Then
                                slHeader = "" 'slLine
                                Exit Do
                            End If
                            If Val(smFieldValues(3)) <> tgChfImpt.iExtRevNo Then
                                slHeader = "" 'slLine
                                Exit Do
                            End If
                            lacCntr(2).Visible = True
                            lacCntr(3).Visible = True
                            lacCntr(3).Caption = smFieldValues(2)
                            mProcessLine
                            slLine = ""
                        End If
                    Loop Until ilEof
                    Do
                        'Get Flights
                        If Trim$(slFlight) = "" Then
                            ilRet = 0
                            On Error GoTo mConvCHFErr:
                            Line Input #hmFlight, slFlight
                            On Error GoTo 0
                            If ilRet <> 0 Then
                                slFlight = ""
                                Exit Do
                            End If
                            If Trim$(slFlight) <> "" Then
                                If (Asc(slFlight) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                                    slFlight = ""
                                    'ilEof = True
                                    Exit Do
                                End If
                            End If
                        ElseIf ilEof Then
                            Exit Do
                        End If
                        DoEvents
                        If Trim$(slFlight) <> "" Then
                            slFlight = mFilter(slFlight)
                            gParseCDFields slFlight, True, smFieldValues()    'Change case
                            If Val(smFieldValues(1)) <> tgChfImpt.lCntrNo Then
                                Exit Do
                            End If
                            If Val(smFieldValues(3)) <> tgChfImpt.iExtRevNo Then
                                Exit Do
                            End If
                            'If Val(smFieldValues(2)) <> tgClfImpt(UBound(tgClfImpt) - 1).ClfRec.iLine Then
                            '    slLine = ""'slFlight
                            '    Exit Do
                            'End If
                            ilLnIndex = -1
                            For ilLoop = LBound(tgClfImpt) To UBound(tgClfImpt) - 1 Step 1
                                If Val(smFieldValues(2)) = tgClfImpt(ilLoop).ClfRec.iLine Then
                                    ilLnIndex = ilLoop
                                    Exit For
                                End If
                            Next ilLoop
                            If ilLnIndex = -1 Then
                                slStr = "Line Missing for flight on " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & ", Flight Line " & Trim$(smFieldValues(2))
                                'lbcErrors.AddItem slStr
                                mAddMsg slStr
                                Print #hmMsg, slStr
                                'Exit Do
                                imTerminate = True
                                Exit Do
                            Else
                                mProcessFlight ilLnIndex
                            End If
                            slFlight = ""
                        End If
                    Loop
                    lacCntr(0).Caption = "Adding " & smNameMsg '& " #" & Str$(tgChfImpt.lCntrNo) & " Rev " & Trim$(Str$(tgChfImpt.iExtRevNo))
                    lacCntr(2).Visible = False
                    lacCntr(3).Visible = False
                    DoEvents
                    If imTerminate Then
                        Print #hmMsg, "** Import Terminated " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                        Close #hmMsg
                        Close hmMove
                        Close hmFlight
                        Close hmLine
                        Close hmCntr
                        Screen.MousePointer = vbDefault
                        mTerminate
                        mConvCHF = False
                        Exit Function
                    End If
                    ilErrorCount = lbcErrors.ListCount
                    imSchCntr = True
                    If Not mSaveRec() Then
                        Print #hmMsg, "** Import Failed " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                        Close #hmMsg
                        Close hmMove
                        Close hmFlight
                        Close hmLine
                        Close hmCntr
                        Screen.MousePointer = vbDefault
                        mTerminate
                        mConvCHF = False
                        Exit Function
                    End If
                    If imSchCntr Then
                        'Schedule the contract
                        slStr = smAdvtName & str$(tgChfImpt.lCntrNo)
                        lacCntr(0).Caption = "Scheduling"
                        DoEvents
                        slStr = smBlankSpaces & smBlankSpaces & "Contract " & str$(tgChfImpt.lCntrNo) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " being scheduled"
                        Print #hmMsg, slStr
                        ilRet = gSchCntr(tgChfImpt.lCode, lbcErrors, slStr)
                        If Not ilRet Then
                            Print #hmMsg, "** Scheduling Import Failed " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                            Close #hmMsg
                            Close hmMove
                            Close hmFlight
                            Close hmLine
                            Close hmCntr
                            Screen.MousePointer = vbDefault
                            mTerminate
                            mConvCHF = False
                            Exit Function
                        End If
                        ilRet = gReSchSpots(False, 0, "YYYYYYY", 0, 86400)
                        If Not ilRet Then
                            Print #hmMsg, "** Rescheduling Import Failed " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                            Close #hmMsg
                            Close hmMove
                            Close hmFlight
                            Close hmLine
                            Close hmCntr
                            Screen.MousePointer = vbDefault
                            mTerminate
                            mConvCHF = False
                            Exit Function
                        End If
                    End If
                    If ilErrorCount = lbcErrors.ListCount Then
                        llCount = llCount + 1
                        lacCount.Caption = Trim$(str$(llCount)) & " converted"
                    Else
                        ilErrorCount = lbcErrors.ListCount
                        lacErrors.Caption = Trim$(str$(ilErrorCount)) & " with errors"
                    End If
                    imFlightError = False
                    lmProcessedNoBytes = Loc(hmCntr) + Loc(hmMove)
                    llPercent = (lmProcessedNoBytes * CSng(100)) / lmTotalNoBytes
                    If llPercent >= 100 Then
                        If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                            llPercent = 99
                        Else
                            llPercent = 100
                        End If
                    End If
                    plcGauge.Value = llPercent
                End If
            End If
        Loop Until ilEof
        'Do remaining moves
        ilRet = mReadMoveRecord(999999999)
        For ilLoop = LBound(tgMoveRec) To UBound(tgMoveRec) - 1 Step 1
            tmMoveRec = tgMoveRec(ilLoop)
            imShowMGErr = True
            ilRet = mImportMG()
            If Not ilRet Then
                mConvCHF = False
                Exit Function
            End If
            ilRet = mSetCurSeqNo(Val(smSeqNo), tgMoveRec(ilLoop).lMoveID)
            'lmRecCount = lmRecCount + 1
        Next ilLoop
        ilRet = mSetSeqNo(Val(smSeqNo))
        Print #hmMsg, " "
        'Print #hmMsg, "Total # of major Errors: " & Trim$(Str$(lmErrorCount))
        Print #hmMsg, "Total # of Warnings: " & Trim$(str$(imWarningCount))
        Print #hmMsg, "Total # of Information messages: " & Trim$(str$(imInfoCount))
        Print #hmMsg, " "
        Print #hmMsg, "** Completed Import of Contracts: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Close hmMove
        Close hmFlight
        Close hmLine
        Close hmCntr
        On Error GoTo mConvCHFErr:
        Kill slFromCntr & "~"
        Name slFromCntr As slFromCntr & "~"
        On Error GoTo mConvCHFErr:
        Kill slFromLine & "~"
        Name slFromLine As slFromLine & "~"
        On Error GoTo mConvCHFErr:
        Kill slFromFlight & "~"
        Name slFromFlight As slFromFlight & "~"
        On Error GoTo mConvCHFErr:
        Kill slFromMove & "~"
        Name slFromMove As slFromMove & "~"
    Next ilPass
    If (lbcErrors.ListCount <= 0) And (Not imTerminate) Then
        'Delete file
        'Kill slFromFile
    End If

    mConvCHF = True
    If Err = 62 Then
        lacCntr(0).Visible = True
        lacCntr(1).Visible = True
        lacCntr(2).Visible = True
        lacCntr(3).Visible = True
    Else
        lacCntr(0).Caption = ""
        lacCntr(1).Caption = ""
        lacCntr(2).Caption = ""
        lacCntr(3).Caption = ""
    End If
'    On Error GoTo mConvCHFErr:
'    Kill slFromCntr & "~"
'    Name slFromCntr As slFromCntr & "~"
'    On Error GoTo mConvCHFErr:
'    Kill slFromLine & "~"
'    Name slFromLine As slFromLine & "~"
'    On Error GoTo mConvCHFErr:
'    Kill slFromFlight & "~"
'    Name slFromFlight As slFromFlight & "~"
'    On Error GoTo mConvCHFErr:
'    Kill slFromMove & "~"
'    Name slFromMove As slFromMove & "~"
    Screen.MousePointer = vbDefault
    Exit Function
mConvCHFErr:
    ilRet = Err.Number
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mFilter                         *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove illegal characters      *
'*                      like: C/R and LF               *                                                   *
'*******************************************************
Private Function mFilter(slInField As String) As String
    Dim slOutField As String
    Dim slChar As String
    Dim ilAscBlank As Integer
    Dim ilPos As Integer
    slOutField = ""
    ilAscBlank = Asc(" ")
    For ilPos = 1 To Len(slInField) Step 1
        slChar = Mid$(slInField, ilPos, 1)
        If Asc(slChar) >= ilAscBlank Then
            slOutField = slOutField & slChar
        End If
    Next ilPos
    mFilter = slOutField
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mFindBonusSpot                  *
'*                                                     *
'*             Created:1/30/00       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Find Spot                      *
'*                                                     *
'*******************************************************
Private Function mFindBonusSpot(llChfCode As Long, ilLineNo As Integer, ilVefCode As Integer, llSdfDate As Long, llRefTrackID As Long) As Integer
'
'   Return tmSdf and tmSmf (If MG)
'

    Dim slDate As String
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim llSunDate As Long
    Dim ilEndPass As Integer
    Dim ilPass As Integer
    Dim ilRet As Integer
    Dim llTstDate As Long
    lgSsfDate(0) = 0
    tmSdfSrchKey0.iVefCode = ilVefCode
    tmSdfSrchKey0.lChfCode = llChfCode
    tmSdfSrchKey0.iLineNo = ilLineNo
    tmSdfSrchKey0.lFsfCode = 0
    slDate = Format$(llSdfDate, "m/d/yy")
    slDate = gObtainPrevMonday(slDate)
    llSunDate = gDateValue(slDate) + 6
    gPackDate slDate, ilDate0, ilDate1
    tmSdfSrchKey0.iDate(0) = ilDate0
    tmSdfSrchKey0.iDate(1) = ilDate1
    tmSdfSrchKey0.sSchStatus = ""
    tmSdfSrchKey0.iTime(0) = 0
    tmSdfSrchKey0.iTime(1) = 0
    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.lChfCode = llChfCode) And (tmSdf.iLineNo = ilLineNo)
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llTstDate
        If llTstDate > llSunDate Then
            If ilPass = ilEndPass Then
                'mFindBonusSpot = False
                'Exit Function
                Exit Do
            Else
                Exit Do
            End If
        End If
        'If (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G")  Then
            'If (tmSdf.sPriceType = "B") And (tmSdf.sSpotType = "X") Then
            If (tmSdf.sSpotType = "X") Then
                If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                    tmSmfSrchKey2.lCode = tmSdf.lCode
                    ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        tmMtfSrchKey.lCode = tmSmf.lMtfCode
                        ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            If llRefTrackID = tmMtf.lTrackID Then
                                mFindBonusSpot = True
                                Exit Function
                            End If
                        End If
                    Else
                        mFindBonusSpot = False
                        Exit Function
                    End If
                Else
                    If tmSdf.sTracer = "*" Then
                        tmMtfSrchKey.lCode = tmSdf.lSmfCode
                        ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            If llRefTrackID = tmMtf.lTrackID Then
                                mFindBonusSpot = True
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        'End If
        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    tmMtfSrchKey1.lCode = llRefTrackID
    ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE Then
        tmSdf.lCode = 0 'spot previously removed
        tmSdf.sSchStatus = ""
        mFindBonusSpot = True
        Exit Function
    End If
    mFindBonusSpot = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mFindSpot                       *
'*                                                     *
'*             Created:1/30/00       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Find Spot                      *
'*                                                     *
'*******************************************************
Private Function mFindSpot(slOper As String, ilVefCode As Integer, llSdfDate As Long, llRefTrackID As Long, llTransGpID As Long) As Integer
'
'   Return tmSdf and tmSmf (If MG)
'

    Dim slDate As String
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim llMonDate As Long
    Dim llSunDate As Long
    Dim ilStartPass As Integer
    Dim ilEndPass As Integer
    Dim ilPass As Integer
    Dim ilRet As Integer
    Dim llTstDate As Long
    Dim llMtfCode As Long
    Dim tlClf As CLF
    If llRefTrackID <= 0 Then
        ilStartPass = 3 '0
        ilEndPass = 4   '1
    Else
        ilStartPass = 1
        ilEndPass = 4   '1
    End If
    llMtfCode = 0
    lgSsfDate(0) = 0
    For ilPass = ilStartPass To ilEndPass Step 1
        tmSdfSrchKey0.iVefCode = ilVefCode
        tmSdfSrchKey0.lChfCode = tmChf.lCode    'llChfCode
        tmSdfSrchKey0.iLineNo = tmClf.iLine
        tmSdfSrchKey0.lFsfCode = 0
        slDate = Format$(llSdfDate, "m/d/yy")
        slDate = gObtainPrevMonday(slDate)
        llMonDate = gDateValue(slDate)
        llSunDate = gDateValue(slDate) + 6
        gPackDate slDate, ilDate0, ilDate1
        tmSdfSrchKey0.iDate(0) = ilDate0
        tmSdfSrchKey0.iDate(1) = ilDate1
        tmSdfSrchKey0.sSchStatus = ""
        tmSdfSrchKey0.iTime(0) = 0
        tmSdfSrchKey0.iTime(1) = 0
        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        'Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.lChfCode = llChfCode) And (tmSdf.iLineNo = ilLineNo)
        Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.lChfCode = tmChf.lCode) 'And (tmSdf.iLineNo = tmClf.iLine)
            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llTstDate
            'If llTstDate > llSunDate Then
            '    If ilPass = ilEndPass Then
            '        mFindSpot = False
            '        Exit Function
            '    Else
            '        Exit Do
            '    End If
            'End If
            ''If llRefTrackID <= 0 Then 'Find missed spot
            ''    If ilPass = 1 Then
            ''        If (tmSdf.sSchStatus = "C") Or (tmSdf.sSchStatus = "H") Or (tmSdf.sSchStatus = "M") Then
            ''            mFindSpot = True
            ''            Exit Function
            ''        End If
            ''    Else
            ''        If (tmSdf.sSchStatus = "S") Then
            ''            mFindSpot = True
            ''            Exit Function
            ''        End If
            ''    End If
            ''Else
            ''    If (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
            ''        tmSmfSrchKey2.lCode = tmSdf.lCode
            ''        ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            ''        If ilRet = BTRV_ERR_NONE Then
            ''            If llRefTrackID = tmSmf.lRefTrackID Then
            ''                mFindSpot = True
            ''                Exit Function
            ''            End If
            ''        End If
            ''    End If
            ''End If
            'Bonus spots can be moved-treat like require spot
            'If (llTstDate >= llMonDate) And (llTstDate <= llSunDate) And (tmSdf.sSpotType <> "X") Then
            If ((llTstDate >= llMonDate) And (llTstDate <= llSunDate) And (tmSdf.sSpotType <> "X") And (slOper <> "M")) Or ((llTstDate >= llMonDate) And (llTstDate <= llSunDate) And (slOper = "M")) Then
                Select Case ilPass
                    Case 1
                        If (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
                            tmSmfSrchKey2.lCode = tmSdf.lCode
                            ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                            If ilRet = BTRV_ERR_NONE Then
                                tmMtfSrchKey.lCode = tmSmf.lMtfCode
                                ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                If ilRet = BTRV_ERR_NONE Then
                                    'If llRefTrackID = tmSmf.lRefTrackID Then
                                    If (llRefTrackID = tmMtf.lTrackID) Then
                                        'llRefTrackID = tmSmf.lRefTrackID
                                        mFindSpot = True
                                        Exit Function
                                    End If
                                    If (llRefTrackID = tmMtf.lRefTrackID) Then
                                        'llRefTrackID = tmSmf.lRefTrackID
                                        If slOper = "S" Then
                                            tmSdf.lCode = 0 'spot previously removed
                                            tmSdf.sSchStatus = ""
                                        End If
                                        mFindSpot = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    Case 2
                        If (tmSdf.sSchStatus = "C") Or (tmSdf.sSchStatus = "H") Or (tmSdf.sSchStatus = "M") Then
                            If (tmSdf.iLineNo = tmClf.iLine) Then
                                If tmSdf.sTracer = "*" Then
                                    tmMtfSrchKey.lCode = tmSdf.lSmfCode
                                    ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                    If ilRet = BTRV_ERR_NONE Then
                                        'If llRefTrackID = tmSmf.lRefTrackID Then
                                        If (llRefTrackID = tmMtf.lTrackID) Then
                                            'llRefTrackID = tmSmf.lRefTrackID
                                            mFindSpot = True
                                            Exit Function
                                        End If
                                        If (llRefTrackID = tmMtf.lRefTrackID) Then
                                            'llRefTrackID = tmSmf.lRefTrackID
                                            If slOper = "S" Then
                                                tmSdf.lCode = 0 'spot previously removed
                                                tmSdf.sSchStatus = ""
                                            End If
                                            mFindSpot = True
                                            Exit Function
                                        End If
                                    End If
                                End If
                            Else
                                'Test if line is a package, then spot ok
                                If (tmClf.sType = "O") Or (tmClf.sType = "A") Or (tmClf.sType = "E") Then
                                    'Read line to determine if part of package
                                    tmClfSrchKey.lChfCode = tmChf.lCode 'llChfCode
                                    tmClfSrchKey.iLine = tmSdf.iLineNo
                                    tmClfSrchKey.iCntRevNo = tmChf.iCntRevNo ' 0 show latest version
                                    tmClfSrchKey.iPropVer = tmChf.iPropVer ' 0 show latest version
                                    ilRet = btrGetGreaterOrEqual(hmClf, tlClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCode = tlClf.lChfCode) And (tmSdf.iLineNo = tlClf.iLine) Then
                                        If (tlClf.iPkLineNo = tmClf.iLine) Then
                                            If tmSdf.sTracer = "*" Then
                                                tmMtfSrchKey.lCode = tmSdf.lSmfCode
                                                ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                If ilRet = BTRV_ERR_NONE Then
                                                    'If llRefTrackID = tmSmf.lRefTrackID Then
                                                    If (llRefTrackID = tmMtf.lTrackID) Then
                                                        'llRefTrackID = tmSmf.lRefTrackID
                                                        mFindSpot = True
                                                        Exit Function
                                                    End If
                                                    If (llRefTrackID = tmMtf.lRefTrackID) Then
                                                        'llRefTrackID = tmSmf.lRefTrackID
                                                        If slOper = "S" Then
                                                            tmSdf.lCode = 0 'spot previously removed
                                                            tmSdf.sSchStatus = ""
                                                        End If
                                                        mFindSpot = True
                                                        Exit Function
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Case 3
                        If (tmSdf.sSchStatus = "C") Or (tmSdf.sSchStatus = "H") Or (tmSdf.sSchStatus = "M") Then
                            If (tmSdf.iLineNo = tmClf.iLine) Then
                                If tmSdf.sTracer <> "*" Then
                                    mFindSpot = True
                                    'llRefTrackID = 0
                                    tmMtf.lTransGpID = 0
                                    Exit Function
                                End If
                            Else
                                'Test if line is a package, then spot ok
                                If (tmClf.sType = "O") Or (tmClf.sType = "A") Or (tmClf.sType = "E") Then
                                    'Read line to determine if part of package
                                    tmClfSrchKey.lChfCode = tmChf.lCode 'llChfCode
                                    tmClfSrchKey.iLine = tmSdf.iLineNo
                                    tmClfSrchKey.iCntRevNo = tmChf.iCntRevNo ' 0 show latest version
                                    tmClfSrchKey.iPropVer = tmChf.iPropVer ' 0 show latest version
                                    ilRet = btrGetGreaterOrEqual(hmClf, tlClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCode = tlClf.lChfCode) And (tmSdf.iLineNo = tlClf.iLine) Then
                                        If (tlClf.iPkLineNo = tmClf.iLine) Then
                                            If tmSdf.sTracer <> "*" Then
                                                'llRefTrackID = 0
                                                mFindSpot = True
                                                tmMtf.lTransGpID = 0
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Case 4
                        If (tmSdf.sSchStatus = "S") Then
                            If (tmSdf.iLineNo = tmClf.iLine) Then
                                'llRefTrackID = 0
                                mFindSpot = True
                                tmMtf.lTransGpID = 0
                                Exit Function
                            Else
                                'Test if line is a package, then spot ok
                                If (tmClf.sType = "O") Or (tmClf.sType = "A") Or (tmClf.sType = "E") Then
                                    'Read line to determine if part of package
                                    tmClfSrchKey.lChfCode = tmChf.lCode
                                    tmClfSrchKey.iLine = tmSdf.iLineNo
                                    tmClfSrchKey.iCntRevNo = tmChf.iCntRevNo ' 0 show latest version
                                    tmClfSrchKey.iPropVer = tmChf.iPropVer ' 0 show latest version
                                    ilRet = btrGetGreaterOrEqual(hmClf, tlClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCode = tlClf.lChfCode) And (tmSdf.iLineNo = tlClf.iLine) Then
                                        If (tlClf.iPkLineNo = tmClf.iLine) Then
                                            'llRefTrackID = 0
                                            mFindSpot = True
                                            tmMtf.lTransGpID = 0
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
                            tmSmfSrchKey2.lCode = tmSdf.lCode
                            ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                            If (ilRet = BTRV_ERR_NONE) And (tmSmf.lMtfCode = 0) Then
                                If (tmSdf.iLineNo = tmClf.iLine) Then
                                    'llRefTrackID = 0
                                    mFindSpot = True
                                    tmMtf.lTransGpID = 0
                                    Exit Function
                                Else
                                    'Test if line is a package, then spot ok
                                    If (tmClf.sType = "O") Or (tmClf.sType = "A") Or (tmClf.sType = "E") Then
                                        'Read line to determine if part of package
                                        tmClfSrchKey.lChfCode = tmChf.lCode
                                        tmClfSrchKey.iLine = tmSdf.iLineNo
                                        tmClfSrchKey.iCntRevNo = tmChf.iCntRevNo ' 0 show latest version
                                        tmClfSrchKey.iPropVer = tmChf.iPropVer ' 0 show latest version
                                        ilRet = btrGetGreaterOrEqual(hmClf, tlClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                        If (ilRet = BTRV_ERR_NONE) And (tmChf.lCode = tlClf.lChfCode) And (tmSdf.iLineNo = tlClf.iLine) Then
                                            If (tlClf.iPkLineNo = tmClf.iLine) Then
                                                'llRefTrackID = 0
                                                mFindSpot = True
                                                tmMtf.lTransGpID = 0
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                End Select
            End If
            ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If (ilStartPass = 1) And (ilPass = 2) And (llMtfCode > 0) Then
            tmMtfSrchKey.lCode = llMtfCode
            ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                If slOper = "S" Then
                    tmSdf.lCode = 0 'spot previously removed
                    tmSdf.sSchStatus = ""
                End If
                mFindSpot = True
                Exit Function
            End If
        End If
        If ilPass = 1 Then
            tmMtfSrchKey1.lCode = llRefTrackID
            ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                'tmSdf.lCode = 0 'spot previously removed
                'tmSdf.sSchStatus = ""
                'mFindSpot = True
                'Exit Function
                llMtfCode = tmMtf.lCode
            End If
        End If
        If (ilPass = 1) And (tmMoveRec.sOper = "S") And (llMtfCode = 0) Then
            tmMtfSrchKey2.lCode = llRefTrackID
            ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmMtf.lRefTrackID = llRefTrackID)
                If (tmMtf.lTransGpID = llTransGpID) Then
                    'tmSdf.lCode = 0 'spot previously removed
                    'tmSdf.sSchStatus = ""
                    'mFindSpot = True
                    'Exit Function
                    llMtfCode = tmMtf.lCode
                    Exit Do
                End If
                ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
    Next ilPass
    mFindSpot = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mFindUndoSpot                   *
'*                                                     *
'*             Created:1/30/00       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Find Spot                      *
'*                                                     *
'*******************************************************
Private Function mFindUndoSpot(slTypeUndo As String, llChfCode As Long, ilLineNo As Integer, ilVefCode As Integer, llSdfDate As Long, llTrackID As Long, llRefTrackID As Long) As Integer
'
'   Return tmSdf and tmSmf (If MG)
'

    Dim slDate As String
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim llMonDate As Long
    Dim llSunDate As Long
    Dim ilRet As Integer
    Dim llTstDate As Long
    lgSsfDate(0) = 0
    tmSdfSrchKey0.iVefCode = ilVefCode
    tmSdfSrchKey0.lChfCode = llChfCode
    tmSdfSrchKey0.iLineNo = ilLineNo
    tmSdfSrchKey0.lFsfCode = 0
    slDate = Format$(llSdfDate, "m/d/yy")
    slDate = gObtainPrevMonday(slDate)
    llMonDate = gDateValue(slDate)
    llSunDate = gDateValue(slDate) + 6
    gPackDate slDate, ilDate0, ilDate1
    tmSdfSrchKey0.iDate(0) = ilDate0
    tmSdfSrchKey0.iDate(1) = ilDate1
    tmSdfSrchKey0.sSchStatus = ""
    tmSdfSrchKey0.iTime(0) = 0
    tmSdfSrchKey0.iTime(1) = 0
    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.lChfCode = llChfCode) And (tmSdf.iLineNo = ilLineNo)
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llTstDate
        If llTstDate > llSunDate Then
            Exit Do
        End If
        If (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
            'If tmSdf.sPriceType <> "B" Then
            If tmSdf.sSpotType <> "X" Then
                tmSmfSrchKey2.lCode = tmSdf.lCode
                ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    tmMtfSrchKey.lCode = tmSmf.lMtfCode
                    ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        If slTypeUndo = "U" Then
                            If (llRefTrackID = tmMtf.lTrackID) Then
                                mFindUndoSpot = True
                                Exit Function
                            End If
                        Else
                            If (llTrackID = tmMtf.lRefTrackID) And (llRefTrackID = tmMtf.lTrackID) Then
                                mFindUndoSpot = True
                                Exit Function
                            End If
                        End If
                    End If
                Else
                    mFindUndoSpot = False
                    Exit Function
                End If
            End If
        ElseIf (tmSdf.sSchStatus = "M") And (tmSdf.sTracer = "*") Then
            tmMtfSrchKey.lCode = tmSdf.lSmfCode
            ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                If slTypeUndo = "U" Then
                    If (llRefTrackID = tmMtf.lTrackID) Then
                        mFindUndoSpot = True
                        Exit Function
                    End If
                Else
                    If (llTrackID = tmMtf.lRefTrackID) And (llRefTrackID = tmMtf.lTrackID) Then
                        mFindUndoSpot = True
                        Exit Function
                    End If
                End If
            End If
        End If
        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    'Check if spot not scheduled
    tmMtfSrchKey1.lCode = llRefTrackID
    ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (llTrackID = tmMtf.lRefTrackID) And (llRefTrackID = tmMtf.lTrackID)
        gUnpackDateLong tmMtf.iSdfDate(0), tmMtf.iSdfDate(1), llTstDate
        If (llTstDate >= llMonDate) And (llTstDate <= llSunDate) Then
            mFindUndoSpot = True
            tmSdf.lCode = 0
            Exit Function
        End If
        ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    tmMtfSrchKey1.lCode = llRefTrackID
    ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If (ilRet = BTRV_ERR_NONE) And (llTrackID = tmMtf.lRefTrackID) Then
        mFindUndoSpot = True
        tmSdf.lCode = 0
        Exit Function
    End If
    mFindUndoSpot = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetAdvertiserCode              *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Add advertiser code #          *
'*                                                     *
'*******************************************************
Private Function mGetAdvertiserCode(slName As String, slDirect As String, slAddr1 As String, slAddr2 As String, slAddr3 As String, slBAddr1 As String, slBAddr2 As String, slBAddr3 As String) As Integer
'
'   ilCode = mGetAdvtCode (slName, slDirect)
'   Where:
'       slName(I)- Advertiser name
'       slDirect(I)- "A"=Agency, "D" = Direct  (required if need to add advertiser)
'       ilCode(O)- Advertiser code #
'
'
    Dim ilLoop As Integer
    imNewAdvt = False
    smAdvtName = Trim$(slName)
    If smAdvtName = "" Then
        smAdvtName = "No Advertiser"
        mGetAdvertiserCode = -1
        Exit Function
    End If
    For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) Step 1
        If StrComp(Trim$(tgCommAdf(ilLoop).sName), Trim$(slName), 1) = 0 Then
            mGetAdvertiserCode = tgCommAdf(ilLoop).iCode
            smAdfCreditRestr = tgCommAdf(ilLoop).sCreditRestr
            'If not matching case- then update
            'If StrComp(Trim$(tgCommAdf(ilLoop).sName), Trim$(slName), 0) <> 0 Then
            '    tmAdfSrchKey.iCode = tgCommAdf(ilLoop).iCode
            '    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            '    If ilRet = BTRV_ERR_NONE Then   'Contract found
            '        Do
            '            tmAdf.sName = slName
            '            ilRet = btrUpdate(hmAdf, tmAdf, imAdfRecLen)
            '            If ilRet = BTRV_ERR_CONFLICT Then
            '                tmAdfSrchKey.iCode = tgCommAdf(ilLoop).iCode
            '                ilCRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            '                If ilCRet <> BTRV_ERR_NONE Then
            '                    Exit Do
            '                End If
            '            End If
            '        Loop While ilRet = BTRV_ERR_CONFLICT
            '    End If
            'End If
            Exit Function
        End If
    Next ilLoop
    ''Set advertiser as missing -1, don't add
    'mGetAdvertiserCode = -1
    'Exit Function
    mGetAdvertiserCode = mAddAdvertiser(slName, slDirect, slAddr1, slAddr2, slAddr3, slBAddr1, slBAddr2, slBAddr3)
    smAdfCreditRestr = tmAdf.sCreditRestr
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetAgencyCode                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get agency code #              *
'*                                                     *
'*******************************************************
Private Function mGetAgencyCode(slName As String, slCityID As String, slAddr1 As String, slAddr2 As String, slAddr3 As String, slBAddr1 As String, slBAddr2 As String, slBAddr3 As String) As Integer
'
'   ilCode = mGetAgencyCode (slName)
'   Where:
'       slName(I)- Agency name
'       ilCode(O)- Agency code #
'
'
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim tlAgf As AGF

    imNewAgy = False
    smAgyName = Trim$(slName)
    smAgyCity = Trim$(slCityID)
    smAgyAddr1 = Trim$(slAddr1)
    If (smAgyName = "") Or (smAgyCity = "") Then
        If smAgyName = "" Then
            smAgyName = "No Agency"
        End If
        If smAgyCity = "" Then
            smAgyCity = "No City"
        End If
        mGetAgencyCode = -1
        Exit Function
    End If
    For ilLoop = LBound(tgCommAgf) To UBound(tgCommAgf) Step 1
        If (StrComp(Trim$(tgCommAgf(ilLoop).sName), Trim$(slName), 1) = 0) And (StrComp(Trim$(tgCommAgf(ilLoop).sCityID), Trim$(slCityID), 1) = 0) Then
            mGetAgencyCode = tgCommAgf(ilLoop).iCode
            smAgyCreditRestr = tgCommAgf(ilLoop).sCreditRestr
            'Update fields
        '    If (StrComp(Trim$(tgCommAgf(ilLoop).sName), Trim$(slName), 0) <> 0) Or (StrComp(Trim$(tgCommAgf(ilLoop).sCityID), Trim$(slCityID), 0) <> 0) Then
        '        tmAgfSrchKey.iCode = tgCommAgf(ilLoop).iCode
        '        ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        '        If ilRet = BTRV_ERR_NONE Then   'Contract found
        '            Do
        '                tmAgf.sName = slName
        '                ilRet = btrUpdate(hmAgf, tmAgf, imAgfRecLen)
        '                If ilRet = BTRV_ERR_CONFLICT Then
        '                    tmAgfSrchKey.iCode = tgCommAgf(ilLoop).iCode
        '                    ilCRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        '                    If ilCRet <> BTRV_ERR_NONE Then
        '                        Exit Do
        '                    End If
        '                End If
        '            Loop While ilRet = BTRV_ERR_CONFLICT
        '        End If
        '    End If
            Exit Function
        End If
    Next ilLoop
    'If name matches, then add agency with new city
    For ilLoop = LBound(tgCommAgf) To UBound(tgCommAgf) Step 1
        If (StrComp(Trim$(tgCommAgf(ilLoop).sName), Trim$(slName), 1) = 0) Then
            tmAgfSrchKey.iCode = tgCommAgf(ilLoop).iCode
            ilRet = btrGetEqual(hmAgf, tlAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then   'Contract found
                mGetAgencyCode = mAddAgency(slName, slCityID, slAddr1, slAddr2, slAddr3, slBAddr1, slBAddr2, slBAddr3, tlAgf.sCrdApp)
            Else
                mGetAgencyCode = mAddAgency(slName, slCityID, slAddr1, slAddr2, slAddr3, slBAddr1, slBAddr2, slBAddr3, "A") ' "R")  changed 6/30/00, Jim
            End If
            smAgyCreditRestr = tmAgf.sCreditRestr
            Exit Function
        End If
    Next ilLoop
    ''Set agency as missing -1, don't add
    'mGetAgencyCode = -1
    'Exit Function
    mGetAgencyCode = mAddAgency(slName, slCityID, slAddr1, slAddr2, slAddr3, slBAddr1, slBAddr2, slBAddr3, "A") '"R")
    smAgyCreditRestr = tmAgf.sCreditRestr
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetChfClfRdf                   *
'*                                                     *
'*             Created:1/30/00       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Chf; Clf and Rdf           *
'*                                                     *
'*******************************************************
Private Function mGetChfClfRdf(llCntrNo As Long, ilLineNo As Integer) As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim llSDate As Long
    Dim llEDate As Long
    Dim tlCffSrchKey As CFFKEY0            'CFF record image
    mGetChfClfRdf = True
    tmChfSrchKey.lCntrNo = llCntrNo
    tmChfSrchKey.iCntRevNo = 32000
    tmChfSrchKey.iPropVer = 32000
    ilRet = btrGetGreaterOrEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)
    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = llCntrNo) Then   'Contract found
        If tmChf.sSchStatus <> "F" Then
            slStr = "Contract" & str$(llCntrNo) & " must be scheduled prior to import"
            'lbcErrors.AddItem slStr
            ilFound = False
            For ilLoop = 0 To UBound(tmErrMsg) - 1 Step 1
                If (tmErrMsg(ilLoop).lCntrNo = llCntrNo) And (tmErrMsg(ilLoop).iLineNo = 0) And (tmErrMsg(ilLoop).iRdfcode = 0) Then
                    ilFound = True
                End If
            Next ilLoop
            If Not ilFound Then
                mAddMsg slStr
                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                tmErrMsg(UBound(tmErrMsg)).lCntrNo = llCntrNo
                tmErrMsg(UBound(tmErrMsg)).iLineNo = 0
                tmErrMsg(UBound(tmErrMsg)).iRdfcode = 0
                ReDim Preserve tmErrMsg(0 To UBound(tmErrMsg) + 1) As ERRMSG
            End If
            mGetChfClfRdf = False
            Exit Function
        End If
        If ilLineNo > 0 Then
            tmClfSrchKey.lChfCode = tmChf.lCode
            tmClfSrchKey.iLine = ilLineNo
            tmClfSrchKey.iCntRevNo = tmChf.iCntRevNo ' 0 show latest version
            tmClfSrchKey.iPropVer = tmChf.iPropVer ' 0 show latest version
            ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            If (tmClf.lChfCode <> tmChf.lCode) Or (tmClf.iLine <> ilLineNo) Then
                ilFound = False
                For ilLoop = 0 To UBound(tmErrMsg) - 1 Step 1
                    If (tmErrMsg(ilLoop).lCntrNo = llCntrNo) And (tmErrMsg(ilLoop).iLineNo = ilLineNo) And (tmErrMsg(ilLoop).iRdfcode = 0) Then
                        ilFound = True
                    End If
                Next ilLoop
                If Not ilFound Then
                    slStr = "Contract" & str$(llCntrNo) & " Line" & str$(ilLineNo) & " Line missing"
                    'lbcErrors.AddItem slStr
                    mAddMsg slStr
                    Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                    tmErrMsg(UBound(tmErrMsg)).lCntrNo = llCntrNo
                    tmErrMsg(UBound(tmErrMsg)).iLineNo = ilLineNo
                    tmErrMsg(UBound(tmErrMsg)).iRdfcode = 0
                    ReDim Preserve tmErrMsg(0 To UBound(tmErrMsg) + 1) As ERRMSG
                End If
                mGetChfClfRdf = False
                Exit Function
            End If
        Else
            ilFound = False
            tmMtfSrchKey1.lCode = tmMoveRec.lRefTrackID
            ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                tmClfSrchKey.lChfCode = tmChf.lCode
                tmClfSrchKey.iLine = tmMtf.iLineNo
                tmClfSrchKey.iCntRevNo = tmChf.iCntRevNo ' 0 show latest version
                tmClfSrchKey.iPropVer = tmChf.iPropVer ' 0 show latest version
                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                If (tmClf.lChfCode = tmChf.lCode) And (tmClf.iLine = tmMtf.iLineNo) Then
                    ilFound = True
                End If
            End If
            If Not ilFound Then
                'Find first non-package line which is not CBS
                tmClfSrchKey.lChfCode = tmChf.lCode
                tmClfSrchKey.iLine = 0
                tmClfSrchKey.iCntRevNo = tmChf.iCntRevNo ' 0 show latest version
                tmClfSrchKey.iPropVer = tmChf.iPropVer ' 0 show latest version
                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (tmClf.lChfCode = tmChf.lCode) And (ilRet = BTRV_ERR_NONE)
                    If (tmClf.sType <> "O") And (tmClf.sType <> "A") And (tmClf.sType <> "E") Then
                        If ((tmClf.iLen = tmMoveRec.iSpotLen) Or (tmMoveRec.iSpotLen = 0)) Then
                            tlCffSrchKey.lChfCode = tmClf.lChfCode
                            tlCffSrchKey.iClfLine = tmClf.iLine
                            tlCffSrchKey.iCntRevNo = tmClf.iCntRevNo
                            tlCffSrchKey.iPropVer = tmClf.iPropVer
                            tlCffSrchKey.iStartDate(0) = 0
                            tlCffSrchKey.iStartDate(1) = 0
                            ilRet = btrGetGreaterOrEqual(hmCff, tmCff, imCffRecLen, tlCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            If (tmCff.lChfCode = tmClf.lChfCode) And (tmCff.iClfLine = tmClf.iLine) Then
                                gUnpackDateLong tmCff.iStartDate(0), tmCff.iStartDate(1), llSDate
                                gUnpackDateLong tmCff.iEndDate(0), tmCff.iEndDate(1), llEDate
                                If llSDate <= llEDate Then
                                    ilFound = True
                                    Exit Do
                                End If
                            End If
                        End If
                    End If
                    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
            If Not ilFound Then
                'Find CBS line since no line exist that is not CBS
                tmClfSrchKey.lChfCode = tmChf.lCode
                tmClfSrchKey.iLine = 0
                tmClfSrchKey.iCntRevNo = tmChf.iCntRevNo ' 0 show latest version
                tmClfSrchKey.iPropVer = tmChf.iPropVer ' 0 show latest version
                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (tmClf.lChfCode = tmChf.lCode) And (ilRet = BTRV_ERR_NONE)
                    If (tmClf.sType <> "O") And (tmClf.sType <> "A") And (tmClf.sType <> "E") Then
                        If ((tmClf.iLen = tmMoveRec.iSpotLen) Or (tmMoveRec.iSpotLen = 0)) Then
                            ilFound = True
                            Exit Do
                        End If
                    End If
                    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
            If Not ilFound Then
                ilFound = False
                For ilLoop = 0 To UBound(tmErrMsg) - 1 Step 1
                    If (tmErrMsg(ilLoop).lCntrNo = llCntrNo) And (tmErrMsg(ilLoop).iLineNo = ilLineNo) And (tmErrMsg(ilLoop).iRdfcode = 0) Then
                        ilFound = True
                    End If
                Next ilLoop
                If Not ilFound Then
                    slStr = "Contract" & str$(llCntrNo) & " Line" & str$(ilLineNo) & " Bonus Line not found"
                    'lbcErrors.AddItem slStr
                    mAddMsg slStr
                    Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                    tmErrMsg(UBound(tmErrMsg)).lCntrNo = llCntrNo
                    tmErrMsg(UBound(tmErrMsg)).iLineNo = ilLineNo
                    tmErrMsg(UBound(tmErrMsg)).iRdfcode = 0
                    ReDim Preserve tmErrMsg(0 To UBound(tmErrMsg) + 1) As ERRMSG
                End If
                mGetChfClfRdf = False
                Exit Function
            End If
        End If
        tmRdf.iCode = -1
        'For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
        '    If tgMRdf(ilLoop).iCode = tmClf.iRdfcode Then
            ilLoop = gBinarySearchRdf(tmClf.iRdfcode)
            If ilLoop <> -1 Then
                tmRdf = tgMRdf(ilLoop)
        '        Exit For
            End If
        'Next ilLoop
        If tmRdf.iCode = -1 Then
            ilFound = False
            For ilLoop = 0 To UBound(tmErrMsg) - 1 Step 1
                If (tmErrMsg(ilLoop).lCntrNo = llCntrNo) And (tmErrMsg(ilLoop).iLineNo = ilLineNo) And (tmErrMsg(ilLoop).iRdfcode = tmClf.iRdfcode) Then
                    ilFound = True
                End If
            Next ilLoop
            If Not ilFound Then
                slStr = "Contract" & str$(llCntrNo) & " Line" & str$(ilLineNo) & " Daypart missing"
                'lbcErrors.AddItem slStr
                mAddMsg slStr
                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                tmErrMsg(UBound(tmErrMsg)).lCntrNo = llCntrNo
                tmErrMsg(UBound(tmErrMsg)).iLineNo = ilLineNo
                tmErrMsg(UBound(tmErrMsg)).iRdfcode = tmClf.iRdfcode
                ReDim Preserve tmErrMsg(0 To UBound(tmErrMsg) + 1) As ERRMSG
            End If
            mGetChfClfRdf = False
            Exit Function
        End If
    Else
        ilFound = False
        For ilLoop = 0 To UBound(tmErrMsg) - 1 Step 1
            If (tmErrMsg(ilLoop).lCntrNo = llCntrNo) And (tmErrMsg(ilLoop).iLineNo = 0) And (tmErrMsg(ilLoop).iRdfcode = 0) Then
                ilFound = True
            End If
        Next ilLoop
        If Not ilFound Then
            slStr = "Contract " & Trim$(str$(llCntrNo)) & " missing"
            'lbcErrors.AddItem slStr
            mAddMsg slStr
            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
            tmErrMsg(UBound(tmErrMsg)).lCntrNo = llCntrNo
            tmErrMsg(UBound(tmErrMsg)).iLineNo = 0
            tmErrMsg(UBound(tmErrMsg)).iRdfcode = 0
            ReDim Preserve tmErrMsg(0 To UBound(tmErrMsg) + 1) As ERRMSG
        End If
        mGetChfClfRdf = False
        Exit Function
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetCompCode                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get competitive code #         *
'*                                                     *
'*******************************************************
Private Function mGetCompCode(slName As String) As Integer
'
'   ilCode = mGetCompCode (slName)
'   Where:
'       slName(I)- Competitive name
'       ilCode(O)- Competitive code #
'
'
    Dim ilLoop As Integer

    For ilLoop = LBound(tgCompMnf) To UBound(tgCompMnf) Step 1
        If StrComp(Trim$(tgCompMnf(ilLoop).sName), Trim$(slName), 1) = 0 Then
            mGetCompCode = tgCompMnf(ilLoop).iCode
            'Remove case change in mChfConv if this coded added
            'If StrComp(Trim$(tgCompMnf(ilLoop).sName), Trim$(slName), 0) <> 0 Then
            '    tmMnfSrchKey.iCode = tgCompMnf(ilLoop).iCode
            '    ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            '    If ilRet = BTRV_ERR_NONE Then   'Contract found
            '        Do
            '            tmMnf.sName = slName
            '            ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
            '            If ilRet = BTRV_ERR_CONFLICT Then
            '                tmMnfSrchKey.iCode = tgCompMnf(ilLoop).iCode
            '                ilCRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            '                If ilCRet <> BTRV_ERR_NONE Then
            '                    Exit Do
            '                End If
            '            End If
            '        Loop While ilRet = BTRV_ERR_CONFLICT
            '    End If
            'End If
            Exit Function
        End If
    Next ilLoop
    mGetCompCode = mAddComp(slName)
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetLnRdf                       *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set line RdfCode                *
'*                                                     *
'*******************************************************
Private Function mGetRdf(ilToVefCode As Integer, llToDate As Long, llToTime As Long, ilAnfCode As Integer) As Integer
    Dim slStartTime As String
    Dim slEndTime As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    ReDim ilDay(0 To 6) As Integer 'Valid days
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilOk As Integer
    Dim ilDayIndex As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilRdfIndex As Integer   'Rdf with smallest subset of times and days OK
    Dim llSubSetExtraTime As Long
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilPass As Integer   'pass 1 for at matching vehicle, pass 2 ignore vehicle test
    Dim llRifStart As Long
    Dim llRifEnd As Long
    If llToTime = -1 Then
        llStartTime = 0
        llEndTime = 86400
    Else
        llStartTime = llToTime  'CLng(gTimeToCurrency(slToTime, False))
        llEndTime = llStartTime
    End If
    For ilLoop = 0 To 6 Step 1
        ilDay(ilLoop) = 0
    Next ilLoop
    ilDay(gWeekDayLong(llToDate)) = 1
    ilRdfIndex = -1
    llStartDate = llToDate  'gDateValue(slToDate)
    llEndDate = llStartDate
    'Find best time which does not valid days
    For ilPass = 1 To 2 Step 1
        If ilPass = 1 Then
            llRifStart = LBound(tgMRif)
            llRifEnd = UBound(tgMRif) - 1
        Else
            llRifStart = LBound(tgMRif)
            llRifEnd = LBound(tgMRif)
        End If
        For llRif = llRifStart To llRifEnd Step 1
            If ((ilToVefCode = tgMRif(llRif).iVefCode) And (ilPass = 1)) Or ((ilToVefCode <> tgMRif(llRif).iVefCode) And (ilPass = 2)) Then
                For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                    If ((tgMRif(llRif).iRdfcode = tgMRdf(ilRdf).iCode) And (ilPass = 1)) Or ((ilPass = 2) And (tgMRdf(ilRdf).sBase = "Y")) Then
                        slEndDate = "12/31/2069"
                        slStartDate = "1/1/1970"
                        If (llStartDate >= gDateValue(slStartDate)) And (llEndDate <= gDateValue(slEndDate)) Then
                            If (tgMRdf(ilRdf).iLtfCode(0) = 0) And (tgMRdf(ilRdf).iLtfCode(1) = 0) And (tgMRdf(ilRdf).iLtfCode(2) = 0) Then
                                For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                    If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                        gUnpackTime tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), "A", "1", slStartTime
                                        gUnpackTime tgMRdf(ilRdf).iEndTime(0, ilIndex), tgMRdf(ilRdf).iEndTime(1, ilIndex), "A", "1", slEndTime
                                        If (llStartTime = CLng(gTimeToCurrency(slStartTime, False))) And (llEndTime = CLng(gTimeToCurrency(slEndTime, True))) Then
                                            'Exact time match- check days
                                            ilOk = True
                                            For ilDayIndex = 0 To 6 Step 1
                                                If (ilDay(ilDayIndex) > 0) And (tgMRdf(ilRdf).sWkDays(ilIndex, ilDayIndex + 1) <> "Y") Then
                                                    ilOk = False
                                                    Exit For
                                                End If
                                            Next ilDayIndex
                                            If ilOk Then
                                                'If ((ilAnfCode = 0) And (tgMRdf(ilRdf).sInOut = "N")) Or ((tgMRdf(ilRdf).sInOut = "I") And (tgMRdf(ilRdf).iAnfCode = ilAnfCode)) Then
                                                If ((ilAnfCode = 0) And (tgMRdf(ilRdf).sInOut = "N")) Or ((tgMRdf(ilRdf).sInOut = "I") And (tgMRdf(ilRdf).ianfCode = ilAnfCode)) Then
                                             Else
                                                    ilOk = False
                                                End If
                                            End If
                                            If ilOk Then
                                                mGetRdf = tgMRdf(ilRdf).iCode
                                                Exit Function
                                            End If
                                        ElseIf (llStartTime >= CLng(gTimeToCurrency(slStartTime, False))) And (llEndTime <= CLng(gTimeToCurrency(slEndTime, True))) Then
                                            'Subset of times- check days
                                            ilOk = True
                                            For ilDayIndex = 0 To 6 Step 1
                                                If (ilDay(ilDayIndex) > 0) And (tgMRdf(ilRdf).sWkDays(ilIndex, ilDayIndex + 1) <> "Y") Then
                                                    ilOk = False
                                                    Exit For
                                                End If
                                            Next ilDayIndex
                                            If ilOk Then
                                                'If ((ilAnfCode = 0) And (tgMRdf(ilRdf).sInOut = "N")) Or ((tgMRdf(ilRdf).sInOut = "I") And (tgMRdf(ilRdf).iAnfCode = ilAnfCode)) Then
                                                If ((ilAnfCode = 0) And (tgMRdf(ilRdf).sInOut = "N")) Or ((tgMRdf(ilRdf).sInOut = "I") And (tgMRdf(ilRdf).ianfCode = ilAnfCode)) Then
                                                Else
                                                    ilOk = False
                                                End If
                                            End If
                                            If ilOk Then
                                                If ilRdfIndex = -1 Then
                                                    ilRdfIndex = ilRdf
                                                    llSubSetExtraTime = CLng(gTimeToCurrency(slEndTime, True)) - CLng(gTimeToCurrency(slStartTime, False)) - (llEndTime - llStartTime)
                                                Else
                                                    If CLng(gTimeToCurrency(slEndTime, True)) - CLng(gTimeToCurrency(slStartTime, False)) - (llEndTime - llStartTime) < llSubSetExtraTime Then
                                                        ilRdfIndex = ilRdf
                                                        llSubSetExtraTime = CLng(gTimeToCurrency(slEndTime, True)) - CLng(gTimeToCurrency(slStartTime, False)) - (llEndTime - llStartTime)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next ilIndex
                            End If
                        End If
                    End If
                Next ilRdf
            End If
        Next llRif
        If ilRdfIndex <> -1 Then
            Exit For
        End If
    Next ilPass
    If ilRdfIndex = -1 Then
        mGetRdf = -1
    Else
        mGetRdf = tgMRdf(ilRdfIndex).iCode
        tmBRdf = tgMRdf(ilRdfIndex)
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGetRecLength                   *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the record length from   *
'*                     the database                    *
'*                                                     *
'*******************************************************
Private Function mGetRecLength(slFileName As String) As Integer
'
'   ilRecLen = mGetRecLength(slName)
'   Where:
'       slName (I)- Name of the file
'       ilRecLen (O)- record length within the file
'
    Dim hlFile As Integer
    Dim ilRet As Integer
    hlFile = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mGetRecLength = -ilRet
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        Exit Function
    End If
    mGetRecLength = btrRecordLength(hlFile)  'Get and save record length
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetSalespersonCode             *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get salesperson code #         *
'*                                                     *
'*******************************************************
Private Function mGetSalespersonCode(slLastName As String, slFirstName As String) As Integer
'
'   ilCode = mGetSalespersonCode (slName)
'   Where:
'       slName(I)- Salesperson last name
'       ilCode(O)- Salesperson code #
'
'
    Dim ilLoop As Integer
    For ilLoop = LBound(tgCSlf) To UBound(tgCSlf) Step 1
        If (StrComp(Trim$(tgCSlf(ilLoop).sLastName), Trim$(slLastName), 1) = 0) And (StrComp(Trim$(tgCSlf(ilLoop).sFirstName), Trim$(slFirstName), 1) = 0) Then
            mGetSalespersonCode = tgCSlf(ilLoop).iCode
            'Remove case convertion in mChfConv when this code added
            'If (StrComp(Trim$(tgCSlf(ilLoop).sLastName), Trim$(slLastName), 0) <> 0) Or (StrComp(Trim$(tgCSlf(ilLoop).sFirstName), Trim$(slFirstName), 0) <> 0) Then
            '    tmSlfSrchKey.iCode = tgCSlf(ilLoop).iCode
            '    ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            '    If ilRet = BTRV_ERR_NONE Then   'Contract found
            '        Do
            '            tmSlf.sFirstName = slFirstName
            '            tmSlf.sLastName = slLastName    'Last name
            '            ilRet = btrUpdate(hmSlf, tmSlf, imSlfRecLen)
            '            If ilRet = BTRV_ERR_CONFLICT Then
            '                tmSlfSrchKey.iCode = tgCSlf(ilLoop).iCode
            '                ilCRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            '                If ilCRet <> BTRV_ERR_NONE Then
            '                    Exit Do
            '                End If
            '            End If
            '        Loop While ilRet = BTRV_ERR_CONFLICT
            '    End If
            'End If
            Exit Function
        End If
    Next ilLoop
    mGetSalespersonCode = mAddSalesperson(slLastName, slFirstName)
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetSeqNo                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get the Contract Import        *
'*                      sequence number                *
'*                                                     *
'*******************************************************
Private Function mGetSeqNo() As Integer
    Dim ilRet As Integer
    'hmSaf = CBtrvTable(ONEHANDLE)
    'ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'If ilRet = BTRV_ERR_NONE Then
        imSafRecLen = Len(tmSaf) 'btrRecordLength(hmSaf)  'Get and save record length
        ilRet = btrGetFirst(hmSaf, tmSaf, imSafRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            mGetSeqNo = -1
            imCurSeqNo = -1
            lmLastMoveID = 0
        Else
            If tmSaf.iImptSeqNo = 999 Then
                tmSaf.iImptSeqNo = -1   'this number is incremented to determine what the next number is.
                                        'The wrap around number is zero (0).
            End If
            mGetSeqNo = tmSaf.iImptSeqNo
            imCurSeqNo = tmSaf.iCurSeqNo
            lmLastMoveID = tmSaf.lLastMoveID
        End If
    'Else
    '    mGetSeqNo = -1
    'End If
    'ilRet = btrClose(hmSaf)
    'btrDestroy hmSaf
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetVehicleCode                 *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get vehicle code #             *
'*                                                     *
'*******************************************************
Private Function mGetVehicleCode(slType As String, slName As String) As Integer
'
'   ilCode = mGetVehicleCode slType, slName)
'   Where:
'       slType(I)- P= Package;H=Hidden;C=Conventional
'       slName(I)- Vehicle name
'       ilCode(O)- Vehicle code #
'
'
    Dim ilLoop As Integer

    For ilLoop = LBound(tgMVef) To UBound(tgMVef) Step 1
        If ((slType <> "P") And (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S")) Or ((slType = "P") And (tgMVef(ilLoop).sType = "P")) Then
            If StrComp(Trim$(tgMVef(ilLoop).sName), Trim$(slName), 1) = 0 Then
                mGetVehicleCode = tgMVef(ilLoop).iCode
                Exit Function
            End If
        End If
    Next ilLoop
    If slType = "P" Then
        mGetVehicleCode = mAddVehicle("P", slName)
    Else
        mGetVehicleCode = mAddVehicle("C", slName)
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mImportMG                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Import MG's                    *
'*                                                     *
'*******************************************************
Private Function mImportMG() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDate As String
    Dim slTime As String
    Dim llPercent As Long
    Dim ilFromVefCode As Integer
    Dim ilToVefCode As Integer
    Dim ilToRdfCode As Integer
    Dim ilVehicleOk As Integer
    Dim llMonDate As Long
    Dim llSunDate As Long
    Dim llRdfSTime As Long
    Dim llRdfETime As Long
    Dim ilDay As Integer
    Dim slLength As String
    Dim llRefTrackID As Long
    Dim llTrackID As Long
    Dim llTransGpID As Long
    Dim llPrevTransGpID As Long
    Dim llFromPrice As Long
    Dim llToPrice As Long
    Dim ilFirstMTF As Integer
    Dim ilFromAnfCode As Integer
    Dim ilToAnfCode As Integer
    Dim ilFound As Integer
    Dim llUndoRefTrackID As Long
    Dim slSchStatus As String
    Dim tlMtf As MTF
    ReDim ilDays(0 To 6) As Integer
    lacCntr(0).Caption = "Processing MG"
    smMoveValues(5) = Trim$(str$(tmMoveRec.lCntrNo))
    smMoveValues(6) = Trim$(str$(tmMoveRec.iLineNo))
    smMoveValues(7) = Trim$(tmMoveRec.sFromVehicle)
    smMoveValues(8) = Format$(tmMoveRec.lFromDate, "m/d/yy")
    smMoveValues(11) = Trim$(tmMoveRec.sToVehicle)
    smMoveValues(12) = Format$(tmMoveRec.lToDate, "m/d/yy")
    lacCntr(1).Caption = smMoveValues(5)
    lacCntr(2) = "Record #"
    lacCntr(3).Caption = Trim$(str$(tmMoveRec.lRecNo))
    lacCntr(2).Visible = True
    lacCntr(3).Visible = True
    DoEvents
    ilVehicleOk = True
    mImportMG = True
    If Not imProcMove Then
        If tmMoveRec.lMoveID = lmLastMoveID Then
            imProcMove = True
        End If
        imInfoCount = imInfoCount + 1
        slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " previously processed"
        'mAddMsg slStr
        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
        Exit Function
    End If
    Select Case tmMoveRec.sOper
        Case "C", "S", "M", "U", "V", "D"
            ilFromVefCode = tmMoveRec.iFromVefCode
            If ilFromVefCode = -1 Then
                tmMtfSrchKey1.lCode = tmMoveRec.lRefTrackID
                ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_NONE Then
                    imWarningCount = imWarningCount + 1
                    slStr = smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " missing"
                    'lbcErrors.AddItem slStr
                    If imShowMGErr Then
                        'mAddMsg slStr
                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                    End If
                    ilVehicleOk = False
                    'mImportMG = False
                End If
            End If
    End Select
    Select Case tmMoveRec.sOper
        Case "C", "S", "M", "B", "U", "V"
            ilToVefCode = tmMoveRec.iToVefCode
            If (ilToVefCode = -1) And (tmMoveRec.sOper = "B") Then
                imWarningCount = imWarningCount + 1
                slStr = smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " missing"
                'lbcErrors.AddItem slStr
                If imShowMGErr Then
                    'mAddMsg slStr
                    Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                End If
                ilVehicleOk = False
                'mImportMG = False
            End If
    End Select
    If ilVehicleOk Then
        llRefTrackID = tmMoveRec.lRefTrackID
        llTrackID = tmMoveRec.lTrackID
        llTransGpID = tmMoveRec.lTransGpID
        llFromPrice = tmMoveRec.lFromPrice  'gStrDecToLong(smMoveValues(16), 2)
        llToPrice = tmMoveRec.lToPrice  'gStrDecToLong(smMoveValues(17), 2)
        ilFromAnfCode = tmMoveRec.iFromAnfCode
        ilToAnfCode = tmMoveRec.iToAnfCode
        imToRdfCode = mGetRdf(ilToVefCode, tmMoveRec.lToDate, tmMoveRec.lToSTime, ilToAnfCode)
        ReDim lgReschSdfCode(1 To 1) As Long
        'Get line number as smMoveValues(6) could to package line number
        ilRet = mGetChfClfRdf(tmMoveRec.lCntrNo, tmMoveRec.iLineNo)
        If ilRet Then
            Select Case tmMoveRec.sOper
                Case "C"    'Combine
                    ilRet = mFindSpot("C", ilFromVefCode, tmMoveRec.lFromDate, llRefTrackID, llTransGpID)
                    If ilRet Then
                        llPrevTransGpID = tmMtf.lTransGpID
                        ilFirstMTF = True
                        tmMtfSrchKey1.lCode = llTrackID
                        ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            ilFirstMTF = False
                        End If
                        If tmSdf.lCode <> 0 Then
                            'If ilFirstMTF Then
                            '    If ilToVefCode <> -1 Then
                                    ilRet = mRemoveSpot(tmMoveRec.lCreatedDate, ilFromVefCode, llTrackID, llRefTrackID, llTransGpID, "C")
                            '    Else
                            '        ilRet = mRemoveSpot(tmMoveRec.lCreatedDate, ilFromVefCode, llTrackID, llRefTrackID, "")
                            '    End If
                            'Else
                            '    ilRet = mRemoveSpot(tmMoveRec.lCreatedDate, ilFromVefCode, llTrackID, llRefTrackID, "")
                            'End If
                        Else
                            'ilRet = True
                            ilRet = mMakeMTF(tmChf.lCode, tmClf.iLine, ilFromVefCode, llTrackID, tmMoveRec.lCreatedDate, llRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, llTransGpID, "", llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
                        End If
                        If ilRet And (ilToVefCode <> -1) Then
                            'Only create one scheduled spot
                            If ilFirstMTF Then
                                ilRet = mMakeSchdMgSpot(tmClf.iLine, ilToVefCode, llTrackID, llRefTrackID, llTransGpID, llPrevTransGpID, ilToAnfCode)
                                If Not ilRet Then
                                    mImportMG = False
                                    Exit Function
                                End If
                            Else
                                imInfoCount = imInfoCount + 1
                                slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " on " & smMoveValues(12) & " part of 1 for N"
                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                            End If
                        ElseIf ilRet And (ilToVefCode = -1) Then
                            If ilFirstMTF Then
                                imInfoCount = imInfoCount + 1
                                slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Combined into Vehicle " & smMoveValues(11) & " on " & smMoveValues(12)
                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                            Else
                                imInfoCount = imInfoCount + 1
                                slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " on " & smMoveValues(12) & " part of 1 for N"
                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                            End If
                        ElseIf Not ilRet Then
                            If igBtrError >= 30000 Then
                                ilRet = csiHandleValue(0, 7)
                            Else
                                ilRet = igBtrError
                            End If
                            If tmSdf.lCode > 0 Then
                                slStr = "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " on " & smMoveValues(8) & " Unable to remove spot"
                            Else
                                slStr = "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " on " & smMoveValues(8) & " Unable to Make MTF"
                            End If
                            mAddMsg slStr
                            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo) & ", Error = " & str$(ilRet)
                            mImportMG = False
                            Exit Function
                        End If
                    Else
                        imWarningCount = imWarningCount + 1
                        slStr = smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " on " & smMoveValues(8) & " Spot missing, might have been previously deleted"
                        'lbcErrors.AddItem slStr
                        'mAddMsg slStr
                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                    End If
                Case "S"    'Split (always M for 1, but if 1 is from combine then result is M for N)
                    ilRet = mFindSpot("S", ilFromVefCode, tmMoveRec.lFromDate, llRefTrackID, llTransGpID)
                    'If Not ilRet Then
                    '    'Test if MTF or SMF record matches, then second part of split (original from non mg spot)
                    '    tmMtfSrchKey2.lCode = llRefTrackID
                    '    ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    '    If ilRet = BTRV_ERR_NONE Then
                    '        tmSdf.lCode = 0 'spot previously removed
                    '        tmSdf.sSchStatus = ""
                    '        ilRet = True
                    '    End If
                    'End If
                    If ilRet Then
                        llPrevTransGpID = tmMtf.lTransGpID
                        If tmSdf.lCode <> 0 Then
                            If ilToVefCode <> -1 Then
                                ilRet = mRemoveSpot(tmMoveRec.lCreatedDate, ilFromVefCode, llTrackID, llRefTrackID, llTransGpID, "")
                            Else
                                ilRet = mRemoveSpot(tmMoveRec.lCreatedDate, ilFromVefCode, llTrackID, llRefTrackID, llTransGpID, "S")
                            End If
                            If Not ilRet Then
                                If igBtrError >= 30000 Then
                                    ilRet = csiHandleValue(0, 7)
                                Else
                                    ilRet = igBtrError
                                End If
                                slStr = "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Parent ID " & Trim$(str$(tmMoveRec.lRefTrackID)) & " remove spot failed in Split"
                                'lbcErrors.AddItem slStr
                                mAddMsg slStr
                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo) & ", Error = " & str$(ilRet)
                                mImportMG = False
                                Exit Function
                            End If
                        Else
                            'ilRet = True
                            If ilToVefCode = -1 Then
                                ilRet = mMakeMTF(tmChf.lCode, tmClf.iLine, ilFromVefCode, llTrackID, tmMoveRec.lCreatedDate, llRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, llTransGpID, "", llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
                            Else
                                ilRet = True
                            End If
                            If Not ilRet Then
                                If imShowMGErr Then
                                    imWarningCount = imWarningCount + 1
                                    slStr = smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Parent ID " & Trim$(str$(tmMoveRec.lRefTrackID)) & " make MTF failed in Split"
                                    'lbcErrors.AddItem slStr
                                    'mAddMsg slStr
                                    Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                End If
                            End If
                        End If
                        If ilRet And (ilToVefCode <> -1) Then
                            ilRet = mMakeSchdMgSpot(tmClf.iLine, ilToVefCode, llTrackID, llRefTrackID, llTransGpID, llPrevTransGpID, ilToAnfCode)
                            If Not ilRet Then
                                mImportMG = False
                                Exit Function
                            End If
                        ElseIf ilRet And (ilToVefCode = -1) Then
                            imInfoCount = imInfoCount + 1
                            slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Split into Vehicle " & smMoveValues(11) & " on " & smMoveValues(12)
                            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                        End If
                    Else
                        imWarningCount = imWarningCount + 1
                        slStr = smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Parent ID " & Trim$(str$(tmMoveRec.lRefTrackID)) & " Unable to find spot"
                        'lbcErrors.AddItem slStr
                        'mAddMsg slStr
                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo) & " Error #" & str$(ilRet)
                    End If
                Case "M"    'Move (1 for 1)
                    ilFound = False
                    If tmMoveRec.sOpResult = "D" Then
                        For ilLoop = 0 To UBound(lmPrevMoveDeleted) - 1 Step 1
                            If lmPrevMoveDeleted(ilLoop) = tmMoveRec.lTrackID Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If Not ilFound Then
                        ilRet = mFindSpot("M", ilFromVefCode, tmMoveRec.lFromDate, llRefTrackID, llTransGpID)
                        If ilRet Then
                            llPrevTransGpID = tmMtf.lTransGpID
                            llTransGpID = tmMtf.lTransGpID
                            If tmMoveRec.sOpResult <> "D" Then
                                'Check values from move of from smf
                                'ilRet = mMakeMTF(tmChf.lCode, tmClf.iLine, ilFromVefCode, llTrackID, tmMoveRec.lCreatedDate, llRefTrackID, smMoveValues(16))
                                If tmSdf.lCode <> 0 Then
                                    If tmSdf.sSpotType = "X" Then
                                        tmMoveRec.sOper = "B"
                                    End If
                                    If ilToVefCode <> -1 Then
                                        ilRet = mRemoveSpot(tmMoveRec.lCreatedDate, ilFromVefCode, llTrackID, llRefTrackID, llTransGpID, "")
                                    Else
                                        ilRet = mRemoveSpot(tmMoveRec.lCreatedDate, ilFromVefCode, llTrackID, llRefTrackID, llTransGpID, "M")
                                    End If
                                    If Not ilRet Then
                                        If igBtrError >= 30000 Then
                                            ilRet = csiHandleValue(0, 7)
                                        Else
                                            ilRet = igBtrError
                                        End If
                                        slStr = "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Parent ID " & Trim$(str$(tmMoveRec.lRefTrackID)) & " remove spot failed in Move"
                                        'lbcErrors.AddItem slStr
                                        mAddMsg slStr
                                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo) & ", Error = " & str$(ilRet)
                                        mImportMG = False
                                        Exit Function
                                    End If
                                Else
                                    'ilRet = True
                                    If ilToVefCode = -1 Then
                                        ilRet = mMakeMTF(tmChf.lCode, tmClf.iLine, ilFromVefCode, llTrackID, tmMoveRec.lCreatedDate, llRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, llTransGpID, "", llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
                                    Else
                                        ilRet = True
                                    End If
                                    If Not ilRet Then
                                        imWarningCount = imWarningCount + 1
                                        slStr = smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Parent ID " & Trim$(str$(tmMoveRec.lRefTrackID)) & " make MTF failed in Move"
                                        'lbcErrors.AddItem slStr
                                        'mAddMsg slStr
                                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                    End If
                                End If
                                'Check values from move of from smf
                                If ilRet And (ilToVefCode <> -1) Then
                                    ilRet = mMakeSchdMgSpot(tmClf.iLine, ilToVefCode, llTrackID, llRefTrackID, llTransGpID, llPrevTransGpID, ilToAnfCode)
                                    If Not ilRet Then
                                        mImportMG = False
                                        Exit Function
                                    End If
                                ElseIf ilRet And (ilToVefCode = -1) Then
                                    imInfoCount = imInfoCount + 1
                                    slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Moved to Vehicle " & smMoveValues(11) & " on " & smMoveValues(12)
                                    Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                End If
                            Else
                                If tmSdf.lCode <> 0 Then
                                    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                                    gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                                    slSchStatus = tmSdf.sSchStatus
                                    If Not gChgSchSpot("D", hmSdf, tmSdf, hmSmf, 0, tmSmf, hmSsf, tgSsf(0), lgSsfDate(0), lgSsfRecPos(0)) Then
                                        If igBtrError >= 30000 Then
                                            ilRet = csiHandleValue(0, 7)
                                        Else
                                            ilRet = igBtrError
                                        End If
                                        slStr = "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " on " & smMoveValues(8) & " Unable to delete Moved Spot"
                                        'lbcErrors.AddItem slStr & " " & sgErrLoc & Str$(ilRet)
                                        'mAddMsg slStr & " " & sgErrLoc & Str$(ilRet)
                                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo) & ", Error = " & str$(ilRet)
                                        mImportMG = False
                                        Exit Function
                                    Else
                                        lmPrevMoveDeleted(UBound(lmPrevMoveDeleted)) = tmMoveRec.lTrackID
                                        ReDim Preserve lmPrevMoveDeleted(0 To UBound(lmPrevMoveDeleted) + 1) As Long
                                        If (slSchStatus = "G") Or (slSchStatus = "O") Then
                                            tmMtfSrchKey.lCode = tmMtf.lCode
                                            ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                            If ilRet = BTRV_ERR_NONE Then
                                                ilRet = btrDelete(hmMtf)
                                            End If
                                            If (ilRet = BTRV_ERR_NONE) Then 'And (ilToVefCode <> -1) Then
                                                ilFound = False
                                                tlMtf = tmMtf
                                                'tmMtfSrchKey1.lCode = llRefTrackID
                                                'ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                                'Do While (ilRet = BTRV_ERR_NONE) And (tmMtf.lTrackID = llRefTrackID) And (tlMtf.lTransGpID = tmMtf.lTransGpID)
                                                '    If Trim$(tmMtf.sSdfSchStatus) <> "" Then
                                                '        If Not ilFound Then
                                                '            tlMtf = tmMtf
                                                '            ilFound = True
                                                '        Else
                                                '            If tmMtf.lCode < tlMtf.lCode Then
                                                '                tlMtf = tmMtf
                                                '            End If
                                                '        End If
                                                '    End If
                                                '    ilRet = btrDelete(hmMtf)
                                                '    ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                                'Loop
                                                tmMtfSrchKey1.lCode = llTrackID
                                                ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                                Do While (ilRet = BTRV_ERR_NONE) And (tmMtf.lTrackID = llTrackID) And (tlMtf.lRefTrackID = tmMtf.lRefTrackID)
                                                    If Trim$(tmMtf.sSdfSchStatus) <> "" Then
                                                        If Not ilFound Then
                                                            tlMtf = tmMtf
                                                            ilFound = True
                                                        Else
                                                            If tmMtf.lCode < tlMtf.lCode Then
                                                                tlMtf = tmMtf
                                                            End If
                                                        End If
                                                    End If
                                                    ilRet = btrDelete(hmMtf)
                                                    ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                                Loop
                                                If ilFound Then
                                                    tmMtf = tlMtf
                                                    If ilToVefCode <> -1 Then
                                                        'Was it from a MG
                                                        If (tmMtf.sSdfSchStatus = "G") Or (tmMtf.sSdfSchStatus = "O") Then
                                                            'reschedule spot
                                                            ilRet = mMakeSchdMgSpot(tmMtf.iLineNo, ilToVefCode, tmMtf.lTrackID, tmMtf.lRefTrackID, llTransGpID, llPrevTransGpID, ilToAnfCode)
                                                            If Not ilRet Then
                                                                mImportMG = False
                                                                Exit Function
                                                            End If
                                                        ElseIf (tmMtf.sSdfSchStatus = "S") Or (tmMtf.sSdfSchStatus = "M") Then
                                                            'Regular spot
                                                            imInfoCount = imInfoCount + 1
                                                            slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Move Spot Removed "
                                                            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                                        End If
                                                    Else
                                                        imInfoCount = imInfoCount + 1
                                                        slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Move Spot Removed"
                                                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                                    End If
                                                Else
                                                    'Regular spot
                                                    imInfoCount = imInfoCount + 1
                                                    slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Spot Removed Only "
                                                    Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                                End If
                                            ElseIf (ilRet <> BTRV_ERR_NONE) Then
                                                imWarningCount = imWarningCount + 1
                                                slStr = smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Move Spot Removed Failed on MTF "
                                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                            End If
                                        Else    'If (ilRet <> BTRV_ERR_NONE) Then
                                            imInfoCount = imInfoCount + 1
                                            slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Spot Removed"
                                            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                        End If
                                    End If
                                Else
                                    lmPrevMoveDeleted(UBound(lmPrevMoveDeleted)) = tmMoveRec.lTrackID
                                    ReDim Preserve lmPrevMoveDeleted(0 To UBound(lmPrevMoveDeleted) + 1) As Long
                                    'tmMtfSrchKey.lCode = tmMtf.lCode
                                    'ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                    'If ilRet = BTRV_ERR_NONE Then
                                    '    ilRet = btrDelete(hmMtf)
                                    'End If
                                    'If (ilRet = BTRV_ERR_NONE) And (ilToVefCode <> -1) Then
                                    '    lmPrevMoveDeleted(UBound(lmPrevMoveDeleted)) = tmMoveRec.lTrackID
                                    '    ReDim Preserve lmPrevMoveDeleted(0 To UBound(lmPrevMoveDeleted) + 1) As Long
                                    '    tlMtf = tmMtf
                                    '    tmMtfSrchKey1.lCode = tmMtf.lRefTrackID
                                    '    ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                    '    If (ilRet = BTRV_ERR_NONE) And (tlMtf.lTransGpID = tmMtf.lTransGpID) Then
                                    '        ilRet = btrDelete(hmMtf)
                                    '    Else
                                    '        tmMtf = tlMtf
                                    '    End If
                                    '    'Was it from a MG
                                    '    If (tmMtf.sSdfSchStatus = "G") Or (tmMtf.sSdfSchStatus = "O") Then
                                    '        'reschedule spot
                                    '        ilRet = mMakeSchdMgSpot(tmMtf.iLineNo, ilToVefCode, tmMtf.lTrackID, tmMtf.lRefTrackID, llTransGpID, llPrevTransGpID)
                                    '    ElseIf tmMtf.sSdfSchStatus = "S" Then
                                    '        'Regular spot
                                    '        slStr = "    Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Moved Spot Removed "
                                    '        Print #hmMsg, slStr & " Rec#" & Str$(tmMoveRec.lRecNo)
                                    '    End If
                                    'End If
                                    If (slSchStatus = "G") Or (slSchStatus = "O") Then
                                        tmMtfSrchKey.lCode = tmMtf.lCode
                                        ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                        If ilRet = BTRV_ERR_NONE Then
                                            ilRet = btrDelete(hmMtf)
                                        End If
                                        If (ilRet = BTRV_ERR_NONE) Then 'And (ilToVefCode <> -1) Then
                                            ilFound = False
                                            tlMtf = tmMtf
                                            'tmMtfSrchKey1.lCode = llRefTrackID
                                            'ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                            'Do While (ilRet = BTRV_ERR_NONE) And (tmMtf.lTrackID = llRefTrackID) And (tlMtf.lTransGpID = tmMtf.lTransGpID)
                                            '    If Trim$(tmMtf.sSdfSchStatus) <> "" Then
                                            '        If Not ilFound Then
                                            '            tlMtf = tmMtf
                                            '            ilFound = True
                                            '        Else
                                            '            If tmMtf.lCode < tlMtf.lCode Then
                                            '                tlMtf = tmMtf
                                            '            End If
                                            '        End If
                                            '    End If
                                            '    ilRet = btrDelete(hmMtf)
                                            '    ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                            'Loop
                                            tmMtfSrchKey1.lCode = llTrackID
                                            ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                            Do While (ilRet = BTRV_ERR_NONE) And (tmMtf.lTrackID = llTrackID) And (tlMtf.lRefTrackID = tmMtf.lRefTrackID)
                                                If Trim$(tmMtf.sSdfSchStatus) <> "" Then
                                                    If Not ilFound Then
                                                        tlMtf = tmMtf
                                                        ilFound = True
                                                    Else
                                                        If tmMtf.lCode < tlMtf.lCode Then
                                                            tlMtf = tmMtf
                                                        End If
                                                    End If
                                                End If
                                                ilRet = btrDelete(hmMtf)
                                                ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                            Loop
                                            If ilFound Then
                                                tmMtf = tlMtf
                                                If ilToVefCode <> -1 Then
                                                    'Was it from a MG
                                                    If (tmMtf.sSdfSchStatus = "G") Or (tmMtf.sSdfSchStatus = "O") Then
                                                        'reschedule spot
                                                        ilRet = mMakeSchdMgSpot(tmMtf.iLineNo, ilToVefCode, tmMtf.lTrackID, tmMtf.lRefTrackID, llTransGpID, llPrevTransGpID, ilToAnfCode)
                                                        If Not ilRet Then
                                                            mImportMG = False
                                                            Exit Function
                                                        End If
                                                    ElseIf (tmMtf.sSdfSchStatus = "S") Or (tmMtf.sSdfSchStatus = "M") Then
                                                        'Regular spot
                                                        imInfoCount = imInfoCount + 1
                                                        slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Move Spot Removed "
                                                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                                    End If
                                                Else
                                                    imInfoCount = imInfoCount + 1
                                                    slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Move Spot Removed"
                                                    Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                                End If
                                            Else
                                                'Regular spot
                                                imInfoCount = imInfoCount + 1
                                                slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Spot Removed Only "
                                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                            End If
                                        ElseIf (ilRet <> BTRV_ERR_NONE) Then
                                            imWarningCount = imWarningCount + 1
                                            slStr = smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Move Spot Removed Failed on MTF "
                                            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                        End If
                                    Else    'If (ilRet <> BTRV_ERR_NONE) Then
                                        imInfoCount = imInfoCount + 1
                                        slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Spot Removed"
                                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                    End If
                                End If
                            End If
                        Else
                            If tmMoveRec.sOpResult <> "D" Then
                                imWarningCount = imWarningCount + 1
                                slStr = smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " on " & smMoveValues(8) & " Spot missing, might have been previously deleted"
                                'lbcErrors.AddItem slStr
                                'mAddMsg slStr
                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                           Else
                                If imShowMGErr Then
                                    imWarningCount = imWarningCount + 1
                                    slStr = smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " on " & smMoveValues(8) & " Spot missing, might have been previously deleted"
                                    'lbcErrors.AddItem slStr
                                    'mAddMsg slStr
                                    Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                End If
                                'mImportMG = False
                            End If
                        End If
                    Else
                        imInfoCount = imInfoCount + 1
                        slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Moved Spot Previously Removed "
                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                    End If
                Case "B"    'Bonus
                    ilToRdfCode = imToRdfCode 'mGetRdf(ilToVefCode, tmMoveRec.lToDate, tmMoveRec.lToSTime)
                    If (ilToRdfCode < 0) And ((tmMoveRec.lToSTime = -1) Or (tmMoveRec.lToETime = -1)) Then
                        slStr = "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " on " & smMoveValues(12) & " can't find daypart"
                        'lbcErrors.AddItem slStr
                        mAddMsg slStr
                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                        mImportMG = False
                        Exit Function
                    Else
                        If ilToRdfCode >= 0 Then
                            For ilLoop = UBound(tmBRdf.iStartTime, 2) To LBound(tmBRdf.iStartTime, 2) Step -1
                                If (tmBRdf.iStartTime(0, ilLoop) <> 1) Or (tmBRdf.iStartTime(1, ilLoop) <> 0) Then
                                    gUnpackTimeLong tmBRdf.iStartTime(0, ilLoop), tmBRdf.iStartTime(1, ilLoop), False, llRdfSTime
                                    gUnpackTimeLong tmBRdf.iEndTime(0, ilLoop), tmBRdf.iEndTime(1, ilLoop), True, llRdfETime
                                    'For ilDay = 0 To 6 Step 1
                                    '    If tmBRdf.sWkDays(ilLoop, ilDay + 1) = "Y" Then
                                    '        ilDays(ilDay) = True
                                    '    Else
                                    '        ilDays(ilDay) = False
                                    '    End If
                                    'Next ilDay
                                    ''Find time is split 12m-1am then 8p-12m, ignore split
                                    ''Exit For
                                End If
                            Next ilLoop
                        Else
                            'For ilDay = 0 To 6 Step 1
                            '    ilDays(ilDay) = True
                            'Next ilDay
                        End If
                        For ilDay = 0 To 6 Step 1
                            ilDays(ilDay) = tmMoveRec.iDays(ilLoop)
                        Next ilDay
                        imVpfIndex = gVpfFindHd(ImptCntr, ilToVefCode, hmVpf)
                        If tgVpf(imVpfIndex).sSCompType = "T" Then
                            gUnpackLength tgVpf(imVpfIndex).iSCompLen(0), tgVpf(imVpfIndex).iSCompLen(1), "3", False, slLength
                            lmCompTime = CLng(gLengthToCurrency(slLength))
                        Else
                            lmCompTime = 0&
                        End If
                        lmSepLength = 1 'Within same avail only
                        slDate = gObtainPrevMonday(smMoveValues(12))
                        llMonDate = gDateValue(slDate)
                        llSunDate = gDateValue(slDate) + 6
                        'If Trim$(smMoveValues(13)) <> "" Then
                        '    llRdfSTime = gTimeToLong(Trim$(smMoveValues(13)), False)
                        'End If
                        If tgVpf(imVpfIndex).sGMedium <> "S" Then
                            If tmMoveRec.lToSTime <> -1 Then
                                llRdfSTime = tmMoveRec.lToSTime
                            End If
                            'If Trim$(smMoveValues(14)) <> "" Then
                            '    llRdfETime = gTimeToLong(Trim$(smMoveValues(14)), True)
                            'End If
                            If tmMoveRec.lToETime <> -1 Then
                                llRdfETime = tmMoveRec.lToETime
                            End If
                        End If
                        ilRet = mBonusSpot(llMonDate, llSunDate, llRdfSTime, llRdfETime, ilDays(), ilToVefCode, llTrackID, tmMoveRec.lCreatedDate, llRefTrackID, llTransGpID, ilToAnfCode)
                        If Not ilRet Then
                            ilRet = mMakeMTF(tmChf.lCode, tmClf.iLine, ilToVefCode, llTrackID, tmMoveRec.lCreatedDate, llRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, llTransGpID, "", llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
                            ilRet = mMakeUnschSpot("B", tmChf.lCode, tmChf.iAdfCode, tmClf.iLine, llMonDate, ilToVefCode, tmClf.iLen)
                            tmSdf.sTracer = "*"
                            tmSdf.lSmfCode = tmMtf.lCode
                            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                            imInfoCount = imInfoCount + 1
                            slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " on " & smMoveValues(12) & " can't schedule Bonus spot, Missed spot created"
                            'lbcErrors.AddItem slStr
                            'mAddMsg slStr
                            If igBtrError = 0 Then
                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo) '& " " & sgErrLoc
                            Else
                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo) & " " & str$(igBtrError) '& " " & sgErrLoc & Str$(igBtrError)
                            End If
                        End If
                    End If
                Case "D"    'Delete Bonus
                    ilFound = False
                    If tmMoveRec.sOpResult = "D" Then
                        For ilLoop = 0 To UBound(lmPrevMoveDeleted) - 1 Step 1
                            If tmMoveRec.lTrackID <> 0 Then
                                If lmPrevMoveDeleted(ilLoop) = tmMoveRec.lTrackID Then
                                    ilFound = True
                                    Exit For
                                End If
                            Else
                                If lmPrevMoveDeleted(ilLoop) = tmMoveRec.lRefTrackID Then
                                    ilFound = True
                                    Exit For
                                End If
                            End If
                        Next ilLoop
                    End If
                    If Not ilFound Then
                        ilRet = mFindBonusSpot(tmChf.lCode, tmClf.iLine, ilFromVefCode, tmMoveRec.lFromDate, llRefTrackID)
                        If ilRet Then
                            If tmSdf.lCode <> 0 Then
                                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                                gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                                If Not gChgSchSpot("D", hmSdf, tmSdf, hmSmf, 0, tmSmf, hmSsf, tgSsf(0), lgSsfDate(0), lgSsfRecPos(0)) Then
                                    If igBtrError >= 30000 Then
                                        ilRet = csiHandleValue(0, 7)
                                    Else
                                        ilRet = igBtrError
                                    End If
                                    slStr = "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " on " & smMoveValues(8) & " Unable to delete Bonus"
                                    'lbcErrors.AddItem slStr & " " & sgErrLoc & Str$(ilRet)
                                    mAddMsg slStr & " " & sgErrLoc & str$(ilRet)
                                    Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo) & " " & sgErrLoc & str$(ilRet)
                                    mImportMG = False
                                    Exit Function
                                Else
                                    imInfoCount = imInfoCount + 1
                                    slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " Bonus deleted on " & slDate & " " & slTime
                                    Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                End If
                            Else
                                imInfoCount = imInfoCount + 1
                                slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " Bonus previously deleted on " & slDate & " " & slTime
                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                            End If
                            tmMtfSrchKey.lCode = tmMtf.lCode
                            ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                            If ilRet = BTRV_ERR_NONE Then
                                ilRet = btrDelete(hmMtf)
                            End If
                        Else
                            If imShowMGErr Then
                                imWarningCount = imWarningCount + 1
                                slStr = smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " on " & smMoveValues(8) & " Bonus Spot missing"
                                'lbcErrors.AddItem slStr
                                'mAddMsg slStr
                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                            End If
                            'mImportMG = False
                        End If
                    Else
                        imInfoCount = imInfoCount + 1
                        slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Bonus Spot Previously Removed "
                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                    End If
                Case "U"    'Undo (Remove)
                    'tmMtfSrchKey1.lCode = llRefTrackID
                    'ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    'If ilRet = BTRV_ERR_NONE Then
                    '    'Determine if at root- if so, add missed spot
                    '    tmMtfSrchKey1.lCode = tmMtf.lRefTrackID
                    '    ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    '    If ilRet <> BTRV_ERR_NONE Then
                    '        'Restore spot
                    '        Do
                    '            tmMtfSrchKey1.lCode = llRefTrackID
                    '            ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                    '            If ilRet = BTRV_ERR_NONE Then
                    '                ilRet = btrDelete(hmMtf)
                    '            Else
                    '                Exit Do
                    '            End If
                    '        Loop While ilRet = BTRV_ERR_CONFLICT
                    '        gUnpackDate tmMtf.iSdfDate(0), tmMtf.iSdfDate(1), slDate
                    '        ilLen = tmClf.iLen
                    '        tmClfSrchKey.lChfCode = tmChf.lCode
                    '        tmClfSrchKey.iLine = tmMtf.iLineNo
                    '        tmClfSrchKey.iCntRevNo = tmChf.iCntRevNo ' 0 show latest version
                    '        tmClfSrchKey.iPropVer = tmChf.iPropVer ' 0 show latest version
                    '        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    '        If (tmClf.lChfCode = tmChf.lCode) And (tmClf.iLine = tmMtf.iLineNo) Then
                    '            ilLen = tmClf.iLen
                    '        End If
                    '        ilRet = mMakeUnschSpot("A", tmChf.lCode, tmChf.iAdfCode, tmMtf.iLineNo, slDate, tmMtf.iSdfVefCode, ilLen)
                    '    End If
                    'Else
                    '    'ilRet = mFindUndoSpot(tmChf.lCode, tmClf.iLine, ilFromVefCode, tmMoveRec.lFromDate, llTrackID)
                    '    If ilRet Then
                    '        'Reset spot as a Missed spot
                    '    Else
                    '        slStr = "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " on " & smMoveValues(8) & " Undo Spot missing"
                    '        lbcErrors.AddItem slStr
                    '        Print #hmMsg, slStr
                    '    End If
                    'End If
                    ilRet = mFindUndoSpot("U", tmChf.lCode, tmClf.iLine, ilFromVefCode, tmMoveRec.lFromDate, llTrackID, llRefTrackID)
                    If ilRet Then
                        If tmSdf.lCode <> 0 Then
                            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                            gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                            If Not gChgSchSpot("D", hmSdf, tmSdf, hmSmf, 0, tmSmf, hmSsf, tgSsf(0), lgSsfDate(0), lgSsfRecPos(0)) Then
                                If igBtrError >= 30000 Then
                                    ilRet = csiHandleValue(0, 7)
                                Else
                                    ilRet = igBtrError
                                End If
                                slStr = "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " on " & smMoveValues(8) & " Unable to delete Bonus"
                                'lbcErrors.AddItem slStr & " " & sgErrLoc & Str$(ilRet)
                                mAddMsg slStr & " " & sgErrLoc & str$(ilRet)
                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo) & " " & sgErrLoc & str$(ilRet)
                                mImportMG = False
                                Exit Function
                            Else
                                tmMtfSrchKey.lCode = tmMtf.lCode
                                ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                If ilRet = BTRV_ERR_NONE Then
                                    ilRet = btrDelete(hmMtf)
                                End If
                                ilFound = False
                                tlMtf = tmMtf
                                'Undo move of a move
                                'tmMtfSrchKey1.lCode = llRefTrackID
                                'ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                'Do While (ilRet = BTRV_ERR_NONE) And (tmMtf.lTrackID = llRefTrackID)
                                '    If tmMtf.lRefTrackID = llRefTrackID Then
                                '        ilRet = btrDelete(hmMtf)
                                '        Exit Do
                                '    End If
                                '    ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                'Loop
                                'tmMtfSrchKey1.lCode = llRefTrackID
                                'ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                'Do While (ilRet = BTRV_ERR_NONE) And (tmMtf.lTrackID = llRefTrackID)
                                '    If tmMtf.lRefTrackID = llTrackID Then
                                '        ilRet = btrDelete(hmMtf)
                                '        ilFound = True
                                '        Exit Do
                                '    End If
                                '    ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                'Loop
                                'If ilFound = False Then
                                '    'Test if from a MG that was MG, then Split, then undone
                                '    tmMtfSrchKey1.lCode = llRefTrackID
                                '    ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                '    Do While (ilRet = BTRV_ERR_NONE)
                                '        If tmMtf.lRefTrackID = llTrackID Then
                                '            ilRet = btrDelete(hmMtf)
                                '            ilFound = True
                                '            Exit Do
                                '        End If
                                '        ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                '    Loop
                                'End If
                                tmMtfSrchKey1.lCode = llRefTrackID
                                ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                Do While (ilRet = BTRV_ERR_NONE) And (tmMtf.lTrackID = llRefTrackID)
                                    If ((tmMtf.lRefTrackID = llTrackID) Or (tmMtf.lRefTrackID = tmMtf.lTrackID)) And (tlMtf.lTransGpID = tmMtf.lTransGpID) Then
                                        If Trim$(tmMtf.sSdfSchStatus) <> "" Then
                                            'tlMtf = tmMtf
                                            'ilFound = True
                                            If Not ilFound Then
                                               tlMtf = tmMtf
                                               ilFound = True
                                            Else
                                               If tmMtf.lCode < tlMtf.lCode Then
                                                  tlMtf = tmMtf
                                               End If
                                            End If
                                        End If
                                        ilRet = btrDelete(hmMtf)
                                        'Exit Do
                                    End If
                                    ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                Loop
                                tmMtfSrchKey1.lCode = llTrackID
                                ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                Do While (ilRet = BTRV_ERR_NONE)
                                    If (tmMtf.lTrackID = llTrackID) And (tlMtf.lTransGpID = tmMtf.lTransGpID) Then
                                        If Trim$(tmMtf.sSdfSchStatus) <> "" Then
                                            'tlMtf = tmMtf
                                            'ilFound = True
                                            If Not ilFound Then
                                               tlMtf = tmMtf
                                               ilFound = True
                                            Else
                                               If tmMtf.lCode < tlMtf.lCode Then
                                                  tlMtf = tmMtf
                                               End If
                                            End If
                                        End If
                                        ilRet = btrDelete(hmMtf)
                                        'Exit Do
                                    End If
                                    ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                Loop
                                'If ilFound And (ilToVefCode <> -1) Then
                                If ilFound Then
                                    'tlMtf = tmMtf
                                    'tmMtfSrchKey1.lCode = tlMtf.lRefTrackID
                                    'ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                    'If (ilRet = BTRV_ERR_NONE) And ((tlMtf.lTransGpID = tmMtf.lTransGpID) Or (Trim$(tmMtf.sSdfSchStatus) <> "")) Then
                                    '    ilRet = btrDelete(hmMtf)
                                    'Else
                                    '    tmMtf = tlMtf
                                    'End If
                                    tmMtf = tlMtf
                                    llPrevTransGpID = 0
                                    'Was it from a MG
                                    If (tmMtf.sSdfSchStatus = "G") Or (tmMtf.sSdfSchStatus = "O") Then
                                        If ilToVefCode = -1 Then
                                            ilRet = mMakeMTF(tmChf.lCode, tmMtf.iLineNo, ilToVefCode, tmMtf.lTrackID, tmMoveRec.lCreatedDate, tmMtf.lRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, tmMtf.lPrevTransGpID, "", llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
                                            imInfoCount = imInfoCount + 1
                                            slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Undo Split Spot "
                                            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                        Else
                                            'reschedule spot
                                            ilRet = mMakeSchdMgSpot(tmMtf.iLineNo, ilToVefCode, tmMtf.lTrackID, tmMtf.lRefTrackID, tmMtf.lPrevTransGpID, llPrevTransGpID, ilToAnfCode)
                                            If Not ilRet Then
                                                mImportMG = False
                                                Exit Function
                                            End If
                                        End If
                                    Else 'If (tmMtf.sSdfSchStatus = "S") Or (tmMtf.sSdfSchStatus = "M") Then
                                        'Regular spot
                                        If tmMoveRec.sOpResult = "R" Then
                                            'Reschedule spot
                                            ilRet = mMakeUnschSpot(tmMoveRec.sOper, tmChf.lCode, tmChf.iAdfCode, tmMtf.iLineNo, tmMoveRec.lToDate, ilToVefCode, tmClf.iLen)
                                            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                                            gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                                            imInfoCount = imInfoCount + 1
                                            slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Undo Combine made Missed spot at " & slDate & " " & slTime
                                            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                        Else
                                            imInfoCount = imInfoCount + 1
                                            slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " Undo Combine Spot removed "
                                            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            tmMtfSrchKey.lCode = tmMtf.lCode
                            ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                            If ilRet = BTRV_ERR_NONE Then
                                ilRet = btrDelete(hmMtf)
                            End If
                            'If (ilToVefCode <> -1) Then
                                tlMtf = tmMtf
                                tmMtfSrchKey2.lCode = llRefTrackID
                                ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                Do While (ilRet = BTRV_ERR_NONE) And (llRefTrackID = tmMtf.lRefTrackID)
                                    If (tlMtf.lTransGpID = tmMtf.lTransGpID) Then
                                        If Trim$(tmMtf.sSdfSchStatus) <> "" Then
                                            'tlMtf = tmMtf
                                            'ilFound = True
                                            If Not ilFound Then
                                               tlMtf = tmMtf
                                               ilFound = True
                                            Else
                                               If tmMtf.lCode < tlMtf.lCode Then
                                                  tlMtf = tmMtf
                                               End If
                                            End If
                                        End If
                                        ilRet = btrDelete(hmMtf)
                                    End If
                                    ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                Loop
                                tmMtfSrchKey1.lCode = llRefTrackID
                                ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                Do While (ilRet = BTRV_ERR_NONE) And (llRefTrackID = tmMtf.lTrackID)
                                    If (tlMtf.lTransGpID = tmMtf.lTransGpID) And (llTrackID = tmMtf.lRefTrackID) Then
                                        If Trim$(tmMtf.sSdfSchStatus) <> "" Then
                                            'tlMtf = tmMtf
                                            'ilFound = True
                                            If Not ilFound Then
                                               tlMtf = tmMtf
                                               ilFound = True
                                            Else
                                               If tmMtf.lCode < tlMtf.lCode Then
                                                  tlMtf = tmMtf
                                               End If
                                            End If
                                        End If
                                        ilRet = btrDelete(hmMtf)
                                    End If
                                    ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                Loop
                                tmMtf = tlMtf
                                llPrevTransGpID = 0
                                'Was it from a MG
                                If (tmMtf.sSdfSchStatus = "G") Or (tmMtf.sSdfSchStatus = "O") Then
                                    If ilToVefCode = -1 Then
                                        ilRet = mMakeMTF(tmChf.lCode, tmMtf.iLineNo, ilFromVefCode, tmMtf.lTrackID, tmMoveRec.lCreatedDate, tmMtf.lRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, tmMtf.lPrevTransGpID, "", llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
                                        imInfoCount = imInfoCount + 1
                                        slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Undo Combine Spot "
                                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                    Else
                                        'reschedule spot
                                        ilRet = mMakeSchdMgSpot(tmMtf.iLineNo, ilToVefCode, tmMtf.lTrackID, tmMtf.lRefTrackID, tmMtf.lPrevTransGpID, llPrevTransGpID, ilToAnfCode)
                                        If Not ilRet Then
                                            mImportMG = False
                                            Exit Function
                                        End If
                                    End If
                                Else 'If (tmMtf.sSdfSchStatus = "S") Or (tmMtf.sSdfSchStatus = "M") Then
                                    'Regular spot
                                    If tmMoveRec.sOpResult = "R" Then
                                        'Reschedule spot- This could cause a problem- sdf previously created
                                        ilRet = mMakeUnschSpot(tmMoveRec.sOper, tmChf.lCode, tmChf.iAdfCode, tmMtf.iLineNo, tmMoveRec.lToDate, ilToVefCode, tmClf.iLen)
                                        gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                                        gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                                        imInfoCount = imInfoCount + 1
                                        slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Undo Combine made Missed spot at " & slDate & " " & slTime
                                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                    Else
                                        imInfoCount = imInfoCount + 1
                                        slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Undo Combine Spot Removed "
                                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                    End If
                                End If
                            'End If
                        End If
                    Else
                        If imShowMGErr Then
                            imWarningCount = imWarningCount + 1
                            slStr = smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " on " & smMoveValues(8) & " Undo Combine Spot missing, might have been previously deleted"
                            'lbcErrors.AddItem slStr
                            'mAddMsg slStr
                            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                        End If
                        'mImportMG = False
                    End If
                Case "V"    'Undo Split
                    ilRet = mFindUndoSpot("V", tmChf.lCode, tmClf.iLine, ilFromVefCode, tmMoveRec.lFromDate, llTrackID, llRefTrackID)
                    If ilRet Then
                        If tmSdf.lCode <> 0 Then
                            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                            gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                            If Not gChgSchSpot("D", hmSdf, tmSdf, hmSmf, 0, tmSmf, hmSsf, tgSsf(0), lgSsfDate(0), lgSsfRecPos(0)) Then
                                If igBtrError >= 30000 Then
                                    ilRet = csiHandleValue(0, 7)
                                Else
                                    ilRet = igBtrError
                                End If
                                slStr = "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " on " & smMoveValues(8) & " Unable to delete Split Spot"
                                'lbcErrors.AddItem slStr & " " & sgErrLoc & Str$(ilRet)
                                mAddMsg slStr & " " & sgErrLoc & str$(ilRet)
                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo) & " " & sgErrLoc & str$(ilRet)
                                mImportMG = False
                                Exit Function
                            Else
                                tmMtfSrchKey.lCode = tmMtf.lCode
                                ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                If ilRet = BTRV_ERR_NONE Then
                                    ilRet = btrDelete(hmMtf)
                                End If
                                llUndoRefTrackID = tmMtf.lRefTrackID
                                tlMtf = tmMtf
                                'If (ilRet = BTRV_ERR_NONE) And (ilToVefCode <> -1) Then
                                '    'MG came from a MG
                                '    tmMtfSrchKey1.lCode = tlMtf.lRefTrackID
                                '    ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                '    Do While (ilRet = BTRV_ERR_NONE) And (tmMtf.lTrackID = tlMtf.lRefTrackID)
                                '        If tlMtf.lTransGpID = tmMtf.lTransGpID Then
                                '            Exit Do
                                '        End If
                                '        ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                '    Loop
                                '    If (ilRet = BTRV_ERR_NONE) And (tmMtf.lTrackID = tlMtf.lRefTrackID) And (tlMtf.lTransGpID = tmMtf.lTransGpID) Then
                                '        ilRet = btrDelete(hmMtf)
                                '    Else
                                '        'MG came from a Spot
                                '        tmMtfSrchKey1.lCode = tlMtf.lTrackID
                                '        ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                '        Do While (ilRet = BTRV_ERR_NONE) And (tmMtf.lTrackID = tlMtf.lTrackID)
                                '            If tlMtf.lTransGpID = tmMtf.lTransGpID Then
                                '                Exit Do
                                '            End If
                                '            ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                '        Loop
                                '        If (ilRet = BTRV_ERR_NONE) And (tmMtf.lTrackID = tlMtf.lTrackID) And (tlMtf.lTransGpID = tmMtf.lTransGpID) Then
                                '            ilRet = btrDelete(hmMtf)
                                '        Else
                                '            'MG that was delete came from a Undo Split, find MTF
                                '            tmMtfSrchKey1.lCode = tlMtf.lTrackID
                                '            ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                '            Do While (ilRet = BTRV_ERR_NONE) And (tmMtf.lTrackID = tlMtf.lTrackID)
                                '                If tlMtf.lRefTrackID = tmMtf.lRefTrackID Then
                                '                    Exit Do
                                '                End If
                                '                ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                '            Loop
                                '            If (ilRet = BTRV_ERR_NONE) And (tmMtf.lTrackID = tlMtf.lTrackID) And (tlMtf.lRefTrackID = tmMtf.lRefTrackID) Then
                                '                ilRet = btrDelete(hmMtf)
                                '            Else
                                '                tmMtf.sSdfSchStatus = ""
                                '            End If
                                '        End If
                                '    End If
                                    tmMtfSrchKey1.lCode = llRefTrackID
                                    ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                    Do While (ilRet = BTRV_ERR_NONE) And (tmMtf.lTrackID = llRefTrackID)
                                        'If (tmMtf.lRefTrackID = llTrackID) Or (tmMtf.lRefTrackID = tmMtf.lTrackID) Then
                                        If (tlMtf.lTransGpID = tmMtf.lTransGpID) Then
                                            If Trim$(tmMtf.sSdfSchStatus) <> "" Then
                                                'tlMtf = tmMtf
                                                'llUndoRefTrackID = tmMtf.lRefTrackID   'tmMtf.lRefTrackID
                                                'ilFound = True
                                                If Not ilFound Then
                                                   tlMtf = tmMtf
                                                   llUndoRefTrackID = tmMtf.lRefTrackID
                                                   ilFound = True
                                                Else
                                                   If tmMtf.lCode < tlMtf.lCode Then
                                                      tlMtf = tmMtf
                                                      llUndoRefTrackID = tmMtf.lRefTrackID
                                                   End If
                                                End If
                                            End If
                                            ilRet = btrDelete(hmMtf)
                                            'Exit Do
                                        End If
                                        ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                    Loop
                                    tmMtfSrchKey1.lCode = llTrackID
                                    ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                    Do While (ilRet = BTRV_ERR_NONE)
                                        If (tmMtf.lTrackID = llTrackID) And (tlMtf.lTransGpID = tmMtf.lTransGpID) Then
                                            If Trim$(tmMtf.sSdfSchStatus) <> "" Then
                                                'tlMtf = tmMtf
                                                'llUndoRefTrackID = tmMtf.lTrackID
                                                'ilFound = True
                                                If Not ilFound Then
                                                   tlMtf = tmMtf
                                                   llUndoRefTrackID = tmMtf.lTrackID
                                                   ilFound = True
                                                Else
                                                   If tmMtf.lCode < tlMtf.lCode Then
                                                      tlMtf = tmMtf
                                                      llUndoRefTrackID = tmMtf.lTrackID
                                                   End If
                                                End If
                                            End If
                                            ilRet = btrDelete(hmMtf)
                                            'Exit Do
                                        End If
                                        ilRet = btrGetNext(hmMtf, tmMtf, imMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                    Loop
                                    tmMtf = tlMtf
                                    ilFound = False
                                    For ilLoop = 0 To UBound(lmPrevSdfCreated) - 1 Step 1
                                        If lmPrevSdfCreated(ilLoop) = llUndoRefTrackID Then
                                            ilFound = True
                                            Exit For
                                        End If
                                    Next ilLoop
                                    If Not ilFound Then
                                        'Was it from a MG
                                        llPrevTransGpID = 0
                                        If (tmMtf.sSdfSchStatus = "G") Or (tmMtf.sSdfSchStatus = "O") Then
                                            If ilToVefCode = -1 Then
                                                ilRet = mMakeMTF(tmChf.lCode, tmMtf.iLineNo, ilToVefCode, tmMtf.lTrackID, tmMoveRec.lCreatedDate, tmMtf.lRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, tmMtf.lPrevTransGpID, "", llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
                                                imInfoCount = imInfoCount + 1
                                                slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Undo Split Spot "
                                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                            Else
                                                'reschedule spot
                                                ilRet = mMakeSchdMgSpot(tmMtf.iLineNo, ilToVefCode, tmMtf.lTrackID, tmMtf.lRefTrackID, tmMtf.lPrevTransGpID, llPrevTransGpID, ilToAnfCode)
                                                If Not ilRet Then
                                                    mImportMG = False
                                                    Exit Function
                                                End If
                                            End If
                                            lmPrevSdfCreated(UBound(lmPrevSdfCreated)) = llUndoRefTrackID
                                            ReDim Preserve lmPrevSdfCreated(0 To UBound(lmPrevSdfCreated) + 1) As Long
                                        Else 'If (tmMtf.sSdfSchStatus = "S") Or (tmMtf.sSdfSchStatus = "M") Then
                                            'Regular spot
                                            If tmMoveRec.sOpResult = "R" Then
                                                'Reschedule spot
                                                ilRet = mMakeUnschSpot(tmMoveRec.sOper, tmChf.lCode, tmChf.iAdfCode, tmMtf.iLineNo, tmMoveRec.lToDate, ilToVefCode, tmClf.iLen)
                                                lmPrevSdfCreated(UBound(lmPrevSdfCreated)) = llUndoRefTrackID
                                                ReDim Preserve lmPrevSdfCreated(0 To UBound(lmPrevSdfCreated) + 1) As Long
                                                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                                                gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                                                imInfoCount = imInfoCount + 1
                                                slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Undo Split Spot Missed spot at " & slDate & " " & slTime
                                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                            Else
                                                imInfoCount = imInfoCount + 1
                                                slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Undo Split Spot Removed "
                                                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                            End If
                                        End If
                                    Else
                                        imInfoCount = imInfoCount + 1
                                        slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Undo Split Spot, second part removed"
                                        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                    End If
                                'End If
                            End If
                        Else
                            tmMtfSrchKey.lCode = tmMtf.lCode
                            ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                            If ilRet = BTRV_ERR_NONE Then
                                ilRet = btrDelete(hmMtf)
                            End If
                            If (ilRet = BTRV_ERR_NONE) And (ilToVefCode <> -1) Then
                                tlMtf = tmMtf
                                tmMtfSrchKey1.lCode = tmMtf.lRefTrackID
                                ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                If (ilRet = BTRV_ERR_NONE) And (tlMtf.lTransGpID = tmMtf.lTransGpID) Then
                                    ilRet = btrDelete(hmMtf)
                                Else
                                    tmMtf = tlMtf
                                End If
                                ilFound = False
                                For ilLoop = 0 To UBound(lmPrevSdfCreated) - 1 Step 1
                                    If lmPrevSdfCreated(ilLoop) = tmMtf.lRefTrackID Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilLoop
                                If Not ilFound Then
                                    'Was it from a MG
                                    llPrevTransGpID = 0
                                    If (tmMtf.sSdfSchStatus = "G") Or (tmMtf.sSdfSchStatus = "O") Then
                                        If ilToVefCode = -1 Then
                                            ilRet = mMakeMTF(tmChf.lCode, tmMtf.iLineNo, ilToVefCode, tmMtf.lTrackID, tmMoveRec.lCreatedDate, tmMtf.lRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, tmMtf.lPrevTransGpID, "", llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
                                            imInfoCount = imInfoCount + 1
                                            slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Undo Split Spot "
                                            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                        Else
                                            'reschedule spot
                                            ilRet = mMakeSchdMgSpot(tmMtf.iLineNo, ilToVefCode, tmMtf.lTrackID, tmMtf.lRefTrackID, tmMtf.lPrevTransGpID, llPrevTransGpID, ilToAnfCode)
                                            If Not ilRet Then
                                                mImportMG = False
                                                Exit Function
                                            End If
                                        End If
                                        lmPrevSdfCreated(UBound(lmPrevSdfCreated)) = tmMtf.lRefTrackID
                                        ReDim Preserve lmPrevSdfCreated(0 To UBound(lmPrevSdfCreated) + 1) As Long
                                    Else 'If (tmMtf.sSdfSchStatus = "S") Or (tmMtf.sSdfSchStatus = "M") Then
                                        'Regular spot
                                        If tmMoveRec.sOpResult = "R" Then
                                            'Reschedule spot- This could cause a problem- sdf previously created
                                            ilRet = mMakeUnschSpot(tmMoveRec.sOper, tmChf.lCode, tmChf.iAdfCode, tmMtf.iLineNo, tmMoveRec.lToDate, ilToVefCode, tmClf.iLen)
                                            lmPrevSdfCreated(UBound(lmPrevSdfCreated)) = tmMtf.lRefTrackID
                                            ReDim Preserve lmPrevSdfCreated(0 To UBound(lmPrevSdfCreated) + 1) As Long
                                            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                                            gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                                            imInfoCount = imInfoCount + 1
                                            slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Undo Split Spot Missed spot at " & slDate & " " & slTime
                                            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                        Else
                                            imInfoCount = imInfoCount + 1
                                            slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Undo Split Spot Removed "
                                            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                        End If
                                    End If
                                Else
                                    imInfoCount = imInfoCount + 1
                                    slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " Undo Split Spot, second part removed"
                                    Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                                End If
                            End If
                        End If
                    Else
                        If imShowMGErr Then
                            imWarningCount = imWarningCount + 1
                            slStr = smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(7) & " on " & smMoveValues(8) & " Undo Split Spot missing, might have been previously deleted"
                            'lbcErrors.AddItem slStr
                            'mAddMsg slStr
                            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                        End If
                        'mImportMG = False
                    End If
            End Select
        Else
            mImportMG = False
            Exit Function
        End If
    End If
    'DoEvents
    'If imTerminate Then
    '    Print #hmMsg, "** Import Terminated " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    '    Close #hmMsg
    '    Close hmMove
    '    Screen.MousePointer = vbDefault
    '    mTerminate
    '    mImportMG = False
    '    Exit Function
    'End If
    'ilErrorCount = lbcErrors.ListCount
    'If ilErrorCount = lbcErrors.ListCount Then
    '    lmRecCount = lmRecCount + 1
    '    lacCount.Caption = Trim$(Str$(lmRecCount)) & " converted"
    'Else
    '    ilErrorCount = lbcErrors.ListCount
    '    lacErrors.Caption = Trim$(Str$(ilErrorCount)) & " with errors"
    'End If
    lmProcessedNoBytes = Loc(hmCntr) + Loc(hmMove)
    llPercent = (lmProcessedNoBytes * CSng(100)) / lmTotalNoBytes
    If llPercent >= 100 Then
        If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
            llPercent = 99
        Else
            llPercent = 100
        End If
    End If
    plcGauge.Value = llPercent
    Exit Function
End Function
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
    Dim slMnfType As String
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilDay As Integer
    Dim ilOk As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim slStr As String
    imTerminate = False
    imFirstActivate = True
    'mParseCmmdLine
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imConverting = False
    imFirstFocus = True
    lmTotalNoBytes = 0
    lmProcessedNoBytes = 0
    imBSMode = False
    hmMsg = -1
    imBypassFocus = False
    imAllClicked = False
    imSetAll = True
    mInitBox
    hmAdf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Advertiser Error:" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imAdfRecLen = Len(tmAdf)
    hmAgf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Agency Error:" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imAgfRecLen = Len(tmAgf)
    hmSlf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Salesperson Error:" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imSlfRecLen = Len(tmSlf)
    hmSaf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Schedule Error:" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imSafRecLen = Len(tmSaf)
    hmMnf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Salesperson Error:" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
    hmPrf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Advertiser Product Error:" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imPrfRecLen = Len(tmPrf)
    hmRcf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmRcf, "", sgDBPath & "Rcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Rate Card Error:" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imRcfRecLen = Len(tmRcf)
    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Rate Card Program/Time Error:" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imRdfRecLen = Len(tmRdf)
    hmCff = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Flight Error:" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    ReDim tgCffImpt(1 To 1) As CFFLIST      'CFF record image
    imCffRecLen = Len(tmCff)
    hmClf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Line Error" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    ReDim tgClfImpt(1 To 1) As CLFLIST      'CLF record image
    imClfRecLen = Len(tgClfImpt(1).ClfRec)
    hmChf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Header Error" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imChfRecLen = Len(tgChfImpt)
    hmVef = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Receivable Error" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmVsf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Combo Vehicle Error" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imVsfRecLen = Len(tmVsf)
    hmSif = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Short Title Error" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imSifRecLen = Len(tmSif)
    hmVpf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmVpf, "", sgDBPath & "Vpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Advertiser Error:" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imVpfRecLen = Len(tmVpf)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Spot (sdf) Error" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)
    hmSmf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open MG Spot (smf) Error" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imSmfRecLen = Len(tmSmf)
    'Copy Rotation
    hmCrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Copy Rotation Error" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    hmMtf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmMtf, "", sgDBPath & "Mtf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open MG Track (Mtf) Error" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imMtfRecLen = Len(tmMtf)
    hmSsf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Spot Summary (ssf) Error" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imSsfRecLen = Len(tmSsf)
    hmLcf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Spot (lcf) Error" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imLcfRecLen = Len(tmLcf)
    hmRlf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmRlf, "", sgDBPath & "Rlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Open Record Block (Rlf) Error" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    
    ReDim tmVcf0(1 To 1) As VCF
    ReDim tmVcf6(1 To 1) As VCF
    ReDim tmVcf7(1 To 1) As VCF
    'hmIcf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    'ilRet = btrOpen(hmIcf, "", sgDBPath & "Icf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'If ilRet <> BTRV_ERR_NONE Then
    '    Screen.MousePointer = vbDefault
    '    MsgBox "Open Import Error", vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
    '    mTerminate
    '    Exit Sub
    'End If
    'imIcfRecLen = Len(tmIcf)
    'Populate arrays to determine if records exist
    ilRet = gObtainAdvt()
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Advertiser Error", vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    ilRet = gObtainAgency()
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Agency Error", vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    ilRet = gObtainSalesperson()
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Salesperson Error", vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    ilRet = gObtainComp()
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Competitive Error", vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    ilRet = gObtainVef()
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Vehicle Error", vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    ilRet = gVpfRead()

    ilRet = gObtainAvail()
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Avail Error", vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    ilRet = mObtainSalesperson()
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Salesperson Error", vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If

    'ilRet = gObtainRcfRifRdf()  'mObtainRateCardPT(0)    'Vehicle zero is the all vehicle
    ReDim tgMRcf(1 To 2) As RCF
    ilRet = gObtainLatestRcf(tgMRcf(1))
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Rate Card Error", vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imRcfCode = tgMRcf(1).iCode
    gRCRead ImptCntr, tgMRcf(1).iCode, tmRcf, tgMRif(), tgMRdf()
    'Determine package rdf record
    sgMRdfStamp = ""
    ilRet = gObtainRdf(sgMRdfStamp, tgMRdf())
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Daypart Error", vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imPkgRdfCode = -1
    For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
        For ilIndex = LBound(tgMRdf(ilLoop).iStartTime, 2) To UBound(tgMRdf(ilLoop).iStartTime, 2) Step 1
            If (tgMRdf(ilLoop).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilLoop).iStartTime(1, ilIndex) <> 0) Then
                gUnpackTimeLong tgMRdf(ilLoop).iStartTime(0, ilIndex), tgMRdf(ilLoop).iStartTime(1, ilIndex), False, llStartTime
                gUnpackTimeLong tgMRdf(ilLoop).iEndTime(0, ilIndex), tgMRdf(ilLoop).iEndTime(1, ilIndex), False, llEndTime
                ilOk = True
                For ilDay = 1 To 7 Step 1
                    If tgMRdf(ilLoop).sWkDays(ilIndex, ilDay) <> "Y" Then
                        ilOk = False
                    End If
                Next ilDay
                If (llStartTime = 0) And (llEndTime = 0) And (ilOk = True) And (tgMRdf(ilLoop).sInOut = "N") And (tgMRdf(ilLoop).sState = "A") Then
                    imPkgRdfCode = tgMRdf(ilLoop).iCode
                    Exit For
                End If
            End If
        Next ilIndex
        If imPkgRdfCode <> -1 Then
            Exit For
        End If
    Next ilLoop
    ReDim tgDMnf(1 To 1) As MNF
    slMnfType = "D"
    ilRet = gObtainMnfForType(slMnfType, sgDMnfStamp, tgDMnf())
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Demo Error", vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    ReDim tgPotMnf(1 To 1) As MNF
    slMnfType = "P"
    ilRet = gObtainMnfForType(slMnfType, sgPotMnfStamp, tgPotMnf())
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Potential Error", vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
        mTerminate
        Exit Sub
    End If
    imBSMode = False
    imCalType = 0   'Standard
    smBlankSpaces = "        "
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    smRptDate = Format$(gNow(), "m/d/yy")
    'gPackDate smRptDate, tmIcf.iDate(0), tmIcf.iDate(1)
    'smRptTime = Format$(gNow(), "h:m:s AM/PM")
    'gPackTime smRptTime, tmIcf.iTime(0), tmIcf.iTime(1)
    'tmIcf.iSeqNo = 0
    'tmIcf.iUrfCode = 0 'tgUrf(0).iCode
    plcGauge.Move ImptCntr.Width / 2 - plcGauge.Width / 2
    'cmcFileConv.Move ImptCntr.Width / 2 - cmcFileConv.Width / 2
    'cmcCancel.Move ImptCntr.Width / 2 - cmcCancel.Width / 2 - cmcCancel.Width
    gCenterStdAlone ImptCntr
    If mTestRecLengths() Then
        Screen.MousePointer = vbDefault
        mTerminate
    End If
    slStr = Format$(gNow(), "m/d/yy")
    slStr = gObtainNextMonday(slStr)
    slStr = gIncOneWeek(slStr)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    edcSDate.Text = slStr
    lacMsg.Caption = lacMsg.Caption & "  Last Sequence Number Imported" & str$(mGetSeqNo())
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
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
    Dim ilLoop As Integer
    'flTextHeight = lacMsg.TextHeight("1") - 35
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    plcCalendar.Move edcSDate.Left, edcSDate.Top + edcSDate.Height + 15
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeMTF                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Make MG History record         *
'*                                                     *
'*******************************************************
Private Function mMakeMTF(llChfCode As Long, ilLineNo As Integer, ilVefCode As Integer, llTrackID As Long, llEnteredDate As Long, llRefTrackID As Long, llFromPrice As Long, llToPrice As Long, llTransGpID As Long, slSchStatus As String, llPrevTransGpID As Long, llToStartTime As Long, llToEndTime As Long, ilDays() As Integer)
    Dim ilRet As Integer
    Dim ilLoop As Integer
    tmMtf.lCode = 0
    tmMtf.lChfCode = llChfCode
    tmMtf.iLineNo = ilLineNo
    tmMtf.lTrackID = llTrackID
    tmMtf.lRefTrackID = llRefTrackID
    tmMtf.lTransGpID = llTransGpID
    gPackDateLong llEnteredDate, tmMtf.iEnteredDate(0), tmMtf.iEnteredDate(1)
    tmMtf.iSdfVefCode = ilVefCode
    tmMtf.iSdfDate(0) = tmSdf.iDate(0)
    tmMtf.iSdfDate(1) = tmSdf.iDate(1)
    tmMtf.iSdfTime(0) = tmSdf.iTime(0)
    tmMtf.iSdfTime(1) = tmSdf.iTime(1)
    tmMtf.sSdfSchStatus = slSchStatus
    tmMtf.lFromPrice = llFromPrice
    tmMtf.lToPrice = llToPrice
    tmMtf.lPrevTransGpID = llPrevTransGpID
    If llToStartTime = -1 Then
        tmMtf.iToStartTime(0) = 1
        tmMtf.iToStartTime(1) = 0
    Else
        gPackTimeLong llToStartTime, tmMtf.iToStartTime(0), tmMtf.iToStartTime(1)
    End If
    If llToEndTime = -1 Then
        tmMtf.iToEndTime(0) = 1
        tmMtf.iToEndTime(1) = 0
    Else
        gPackTimeLong llToEndTime, tmMtf.iToEndTime(0), tmMtf.iToEndTime(1)
    End If
    tmMtf.iToRdfCode = imToRdfCode
    For ilLoop = 0 To 6 Step 1
        tmMtf.iDays(ilLoop) = ilDays(ilLoop)
    Next ilLoop
    ilRet = btrInsert(hmMtf, tmMtf, imMtfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        mMakeMTF = False
        imRet = ilRet
        igBtrError = imRet
    Else
        mMakeMTF = True
        lgMtfNoRecs = btrRecords(hmMtf)
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeSchdMgSpot                 *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Create a Sdf records and        *
'*                     schedule                        *
'*                                                     *
'*******************************************************
Private Function mMakeSchdMgSpot(ilLineNo As Integer, ilToVefCode As Integer, llTrackID As Long, llRefTrackID As Long, llTransGpID As Long, llPrevTransGpID As Long, ilToAnfCode As Integer) As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilToRdfCode As Integer
    mMakeSchdMgSpot = True
    ilRet = mMakeUnschSpot(tmMoveRec.sOper, tmChf.lCode, tmChf.iAdfCode, ilLineNo, tmMoveRec.lToDate, ilToVefCode, tmClf.iLen)
    If ilRet <> BTRV_ERR_NONE Then
        If ilRet >= 30000 Then
            ilRet = csiHandleValue(0, 7)
        End If
        slStr = "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " on " & smMoveValues(12) & " Unable to create spot"
        'lbcErrors.AddItem slStr
        mAddMsg slStr
        Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo) & " Error #" & str$(ilRet)
        mMakeSchdMgSpot = False
        Exit Function
        'ilRet = mMakeMTF(tmSdf.lChfCode, tmSdf.iLineNo, tmSdf.iVefCode, llTrackID, tmMoveRec.lCreatedDate, llRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, llTransGpID, "", llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
    Else
        ilToRdfCode = mGetRdf(ilToVefCode, tmMoveRec.lToDate, tmMoveRec.lToSTime, ilToAnfCode)
        If ilToRdfCode < 0 Then
            imWarningCount = imWarningCount + 1
            slStr = smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " on " & smMoveValues(12) & " can't find daypart, Missed spot created"
            'lbcErrors.AddItem slStr
            'mAddMsg slStr
            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
            'mMakeSchdMgSpot = False
            tmSdfSrchKey3.lCode = tmSdf.lCode
            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                ilRet = mMakeMTF(tmSdf.lChfCode, tmSdf.iLineNo, tmSdf.iVefCode, llTrackID, tmMoveRec.lCreatedDate, llRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, llTransGpID, "", llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
                'ilRet = btrDelete(hmSdf)
                tmSdf.sTracer = "*"
                tmSdf.lSmfCode = tmMtf.lCode
                ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
            End If
        Else
            If tmMoveRec.sOper = "B" Then
                ilRet = gMGSchSpots(True, "O", tmSdf.lCode, ilToRdfCode, ilToVefCode, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
            Else
                ilRet = gMGSchSpots(True, "G", tmSdf.lCode, ilToRdfCode, ilToVefCode, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
            End If
            If Not ilRet Then
                imInfoCount = imInfoCount + 1
                slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " on " & smMoveValues(12) & " can't schedule MG spot, Missed spot created"
                'lbcErrors.AddItem slStr
                'mAddMsg slStr
                Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
                'mMakeSchdMgSpot = False
                tmSdfSrchKey3.lCode = tmSdf.lCode
                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = mMakeMTF(tmSdf.lChfCode, tmSdf.iLineNo, tmSdf.iVefCode, llTrackID, tmMoveRec.lCreatedDate, llRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, llTransGpID, "", llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
                    'ilRet = btrDelete(hmSdf)
                    tmSdf.sTracer = "*"
                    tmSdf.lSmfCode = tmMtf.lCode
                    ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                End If
            Else
                ilRet = mUpdateSMF(llTrackID, tmMoveRec.lCreatedDate, llRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, llTransGpID, llPrevTransGpID)
                ilRet = gReSchSpots(False, 0, "YYYYYYY", 0, 86400)
            End If
        End If
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeUnschSpot                  *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Create a Sdf records            *
'*                                                     *
'*                     Similar to code within          *
'*                     CntSchd.Bas                     *
'*                                                     *
'*******************************************************
Private Function mMakeUnschSpot(slType As String, llChfCode As Long, ilAdfCode As Integer, ilLineNo As Integer, llInDate As Long, ilVefCode As Integer, ilLen As Integer) As Integer
'
'   mMakeUnschSpot slType, llChfCode, ilAdfCode, ilLineNo, lDate, ilVefCode, ilLen
'   Where:
'       slType(I)- B=Bonus; Not B = Unschd spot
'       llChfCode(I)- Chf Code
'       ilLineNo(I)- Line number
'       slDate(I)- Date to create spot for
'       ilExtraSpot(I)- True=Extra Bonus Spot
'       llSdfRecPos(O)- Sdf Record position
'
    Dim llDate As Long
    Dim ilDay As Integer
    Dim ilRet As Integer

    tmSdf.lCode = 0
    tmSdf.iVefCode = ilVefCode      'Vehicle Code (combos not allowed)
    tmSdf.lChfCode = llChfCode    'Contract code
    tmSdf.iLineNo = ilLineNo    'Line number
    tmSdf.lFsfCode = 0
    tmSdf.iAdfCode = ilAdfCode 'Advertiser code number
    gPackDateLong llInDate, tmSdf.iDate(0), tmSdf.iDate(1)
    llDate = llInDate   'gDateValue(slDate)
    ilDay = gWeekDayLong(llDate)
    tmSdf.iTime(0) = 0
    tmSdf.iTime(1) = 0
    tmSdf.sSchStatus = "M"    'S=Scheduled, M=Missed,
                                'G=Makegood, A=on alternate log but not MG, B=on alternate Log and MG,
                                'C=Cancelled
    tmSdf.iMnfMissed = 0   'Missed reason
    tmSdf.sTracer = " "   'M=Mouse move, N=On demand & mouse moved, C=Created in post log,
                            'N=N/A, D=on Demand & created in post log
    tmSdf.sAffChg = " "   'T=Time change, C=Copy change, B=Time and copy changed, blank=no change
    tmSdf.sPtType = "0"
    tmSdf.lCopyCode = 0        'Copy inventory code
    tmSdf.iRotNo = 0
    tmSdf.iLen = ilLen         'Spot length
    'If rbcFillInv(0).Value Then
    '    tmSdf.sPriceType = "B"  'tmClf.sPriceType 'T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
    '    tmSdf.sSpotType = "X"   '"T"=Remnant; Q=per Inquire; S=PSA; M=Promo; X=Extra Spot
    'Else
    '    tmSdf.sPriceType = "N"  'tmClf.sPriceType 'T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
    '    tmSdf.sSpotType = "X"   '"T"=Remnant; Q=per Inquire; S=PSA; M=Promo; X=Extra Spot
    'End If
    If slType = "B" Then
        tmSdf.sPriceType = "B"  'tmClf.sPriceType 'T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
        tmSdf.sSpotType = "X"   '"T"=Remnant; Q=per Inquire; S=PSA; M=Promo; X=Extra Spot
    Else
        tmSdf.sPriceType = "L" 'Use line pricetmCClf.sPriceType 'T=True; N=No Charge; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
        tmSdf.sSpotType = "A" 'tmCClf.sBB       'Spot type: N=N/A (regular spot); O=Open BB; C=Close BB; F=Floater BB;
                                'M=Commercial Promo(create avails if not found); P=Pledge spot;
                                'D=Donut; B=Bookend; S=PSA; R=Promo;
    End If
    tmSdf.sBill = "N"
    tmSdf.lSmfCode = 0
    tmSdf.iurfCode = tgUrf(0).iCode      'Last user who modified spot
    tmSdf.iGameNo = 0
    tmSdf.sXCrossMidnight = "N"
    tmSdf.sUnused = ""
    ilRet = btrInsert(hmSdf, tmSdf, imSdfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        mMakeUnschSpot = ilRet
        Exit Function
    End If
    mMakeUnschSpot = ilRet
    Exit Function
End Function





'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainSalesperson              *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgCSlf for collection  *
'*                                                     *
'*******************************************************
Private Function mObtainSalesperson() As Integer
'
'   ilRet = mObtainSalesperson ()
'   Where:
'       tgCSlf() (I)- SLFEXT record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Slf date/time stamp
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffset As Integer
    Dim ilUpperBound As Integer

    slStamp = gFileDateTime(sgDBPath & "Slf.Btr")
    If sgCSlfStamp <> "" Then
        If StrComp(slStamp, sgCSlfStamp, 1) = 0 Then
            If UBound(tgCSlf) > 1 Then
                mObtainSalesperson = True
                Exit Function
            End If
        End If
    End If
    ReDim tgCSlf(1 To 1) As SLFEXT
    ilRecLen = Len(tmSlf) 'btrRecordLength(hmSlf)  'Get and save record length
    sgCSlfStamp = slStamp
    'llNoRec = btrRecords(hmSlf) 'Obtain number of records
    llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlFile) 'Obtain number of records
    btrExtClear hmSlf   'Clear any previous extend operation
    ilRet = btrGetFirst(hmSlf, tmSlf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        mObtainSalesperson = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            mObtainSalesperson = False
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hmSlf, llNoRec, 0, "UC", "SLFEXTPK", SLFEXTPK) 'Set extract limits (all records)
    ilOffset = gFieldOffset("Slf", "SlfCode")   '0
    ilRet = btrExtAddField(hmSlf, ilOffset, 2)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        mObtainSalesperson = False
        Exit Function
    End If
    ilOffset = gFieldOffset("Slf", "SlfFirstName")  '2
    ilRet = btrExtAddField(hmSlf, ilOffset, 20)  'Extract First Name field
    If ilRet <> BTRV_ERR_NONE Then
        mObtainSalesperson = False
        Exit Function
    End If
    ilOffset = gFieldOffset("Slf", "SlfLastName")   '22
    ilRet = btrExtAddField(hmSlf, ilOffset, 20) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        mObtainSalesperson = False
        Exit Function
    End If
'    ilRet = btrExtGetNextExt(hmSlf)    'Extract record
'    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
'        If ilRet <> BTRV_ERR_NONE Then
'            mObtainSalesperson = False
'            Exit Function
'        End If
'    End If
'    ilUpperBound = UBound(tgCSlf)
'    ilExtLen = Len(tgCSlf(ilUpperBound))  'Extract operation record size
'    ilRet = btrExtGetFirst(hmSlf, tgCSlf(ilUpperBound), ilExtLen, llRecPos)
'    Do While ilRet = BTRV_ERR_NONE
'        ilUpperBound = ilUpperBound + 1
'        ReDim Preserve tgCSlf(1 To ilUpperBound) As SLFEXT
'        ilRet = btrExtGetNext(hmSlf, tgCSlf(ilUpperBound), ilExtLen, llRecPos)
'    Loop
    
    ilRet = btrExtGetNext(hmSlf, tgCSlf(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            mObtainSalesperson = False
            Exit Function
        End If
        ilUpperBound = UBound(tgCSlf)
        ilExtLen = Len(tgCSlf(ilUpperBound))  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmSlf, tgCSlf(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilUpperBound = ilUpperBound + 1
            ReDim Preserve tgCSlf(1 To ilUpperBound) As SLFEXT
            ilRet = btrExtGetNext(hmSlf, tgCSlf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmSlf, tgCSlf(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If

    mObtainSalesperson = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Function mOpenMsgFile()
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    On Error GoTo mOpenMsgFileErr:
    slToFile = sgImportPath & smRptName
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        slFileDate = Format$(slDateTime, "m/d/yy")
        If gDateValue(slFileDate) = lmNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            On Error GoTo mOpenMsgFileErr:
            hmMsg = FreeFile
            Open slToFile For Append As hmMsg
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            On Error GoTo mOpenMsgFileErr:
            hmMsg = FreeFile
            Open slToFile For Output As hmMsg
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        On Error GoTo mOpenMsgFileErr:
        hmMsg = FreeFile
        Open slToFile For Output As hmMsg
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, "** Import Contracts: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    Print #hmMsg, ""
    mOpenMsgFile = True
    Exit Function
mOpenMsgFileErr:
    ilRet = Err.Number
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mProcessFlight                  *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Process flight type record      *
'*                                                     *
'*******************************************************
Private Sub mProcessFlight(ilLnIndex As Integer)
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim ilCff As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStr As String
    Dim ilAnyDays As Integer
    Dim slActPrice As String
    Dim ilNoSpots As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilReplace As Integer
    Dim ilLastCff As Integer
    ilNoSpots = 0
    If smFieldValues(6) = "D" Then
        For ilLoop = 0 To 6 Step 1
            ilNoSpots = ilNoSpots + Val(smFieldValues(ilLoop + 8))
        Next ilLoop
    Else
        ilNoSpots = Val(smFieldValues(7))
    End If
    ilUpper = UBound(tgCffImpt)
    ilReplace = False
    'If imPorO = 1 Then
        'If gDateValue(smFieldValues(4)) < lmNowDate Then
        '    slStr = "Flight date prior to today for " & Trim$(Str$(tgChfImpt.lCntrNo)) & ", Line " & Trim$(Str$(tgClfImpt(ilLnIndex).ClfRec.iLine))
        '    lbcErrors.AddItem slStr
        '    Print #hmMsg, slStr
        '    imFlightError = True
        '    'If Trim$(tmIcf.sErrorMess) = "" Then
        '    '    tmIcf.sErrorMess = "Flight date prior to today "
        '    'Else
        '    '    tmIcf.sErrorMess = tmIcf.sErrorMess & "; "
        '    'End If
        '    'tmIcf.sErrorMess = tmIcf.sErrorMess & Trim$(Str$(tgClfImpt(ilLnIndex).ClfRec.iLine))
        '    Exit Sub
        'End If
    'End If
    If tgClfImpt(ilLnIndex).iFirstCff = -1 Then
        tgClfImpt(ilLnIndex).iFirstCff = ilUpper
    Else
        If ilNoSpots = 0 Then
            Exit Sub
        End If
        'Test if flight part of another flight
        ilCff = tgClfImpt(ilLnIndex).iFirstCff
        ilLastCff = ilCff
        Do While ilCff <> -1
            If tgCffImpt(ilCff).iStatus >= 0 Then
                gUnpackDate tgCffImpt(ilCff).CffRec.iStartDate(0), tgCffImpt(ilCff).CffRec.iStartDate(1), slStartDate
                gUnpackDate tgCffImpt(ilCff).CffRec.iEndDate(0), tgCffImpt(ilCff).CffRec.iEndDate(1), slEndDate
                If gDateValue(slEndDate) < gDateValue(slStartDate) Then
                    ilReplace = True
                    ilUpper = ilCff
                    Exit Do
                Else
                    If (gDateValue(smFieldValues(4)) <= gDateValue(slEndDate)) And (gDateValue(smFieldValues(5)) >= gDateValue(slStartDate)) Then
                        slStr = "Conflicting flights for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & ", Line " & Trim$(str$(tgClfImpt(ilLnIndex).ClfRec.iLine))
                        'lbcErrors.AddItem slStr
                        mAddMsg slStr
                        Print #hmMsg, slStr
                        imFlightError = True
                        'If Trim$(tmIcf.sErrorMess) = "" Then
                        '    tmIcf.sErrorMess = "Flight Error for line "
                        'Else
                        '    tmIcf.sErrorMess = tmIcf.sErrorMess & "; "
                        'End If
                        'tmIcf.sErrorMess = tmIcf.sErrorMess & Trim$(Str$(tgClfImpt(ilLnIndex).ClfRec.iLine))
                        Exit Sub
                    End If
                End If
            End If
            ilLastCff = ilCff
            ilCff = tgCffImpt(ilCff).iNextCff
        Loop
        If Not ilReplace Then
            tgCffImpt(ilLastCff).iNextCff = ilUpper
        End If
    End If
    tgCffImpt(ilUpper).CffRec.lChfCode = tgChfImpt.lCode        'Contract code
    tgCffImpt(ilUpper).CffRec.iClfLine = Val(smFieldValues(2))        'Line number
    tgCffImpt(ilUpper).CffRec.iCntRevNo = Val(smFieldValues(3))     'Version number
    tgCffImpt(ilUpper).CffRec.iPropVer = 0     'Version number
    gPackDate smFieldValues(4), tgCffImpt(ilUpper).CffRec.iStartDate(0), tgCffImpt(ilUpper).CffRec.iStartDate(1)
    llStartDate = gDateValue(smFieldValues(4))
    llEndDate = gDateValue(smFieldValues(5))
    If ilNoSpots = 0 Then
        llEndDate = llStartDate - 1
    Else
        If llEndDate = llStartDate Then
            Do While gWeekDayLong(llEndDate) <> 6
                llEndDate = llEndDate + 1
            Loop
        End If
    End If
    gPackDateLong llEndDate, tgCffImpt(ilUpper).CffRec.iEndDate(0), tgCffImpt(ilUpper).CffRec.iEndDate(1)
    If smFieldValues(6) = "D" Then
        tgCffImpt(ilUpper).CffRec.sDyWk = "D"
        tgCffImpt(ilUpper).CffRec.iSpotsWk = 0
        For ilLoop = 0 To 6 Step 1
            tgCffImpt(ilUpper).CffRec.iDay(ilLoop) = Val(smFieldValues(ilLoop + 8))
        Next ilLoop
    Else
        ilAnyDays = False
        tgCffImpt(ilUpper).CffRec.sDyWk = "W"
        tgCffImpt(ilUpper).CffRec.iSpotsWk = Val(smFieldValues(7))
        For ilLoop = 0 To 6 Step 1
            If smFieldValues(ilLoop + 8) = "Y" Then
                ilAnyDays = True
                tgCffImpt(ilUpper).CffRec.iDay(ilLoop) = 1
            Else
                tgCffImpt(ilUpper).CffRec.iDay(ilLoop) = 0
            End If
        Next ilLoop
        tgCffImpt(ilUpper).CffRec.iXSpotsWk = 0
        For ilLoop = 0 To 6 Step 1
            tgCffImpt(ilUpper).CffRec.sXDay(ilLoop) = "0"
        Next ilLoop
        If Not ilAnyDays Then
            slStr = "No air days defined for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & ", Line " & Trim$(str$(tgClfImpt(ilLnIndex).ClfRec.iLine))
            'lbcErrors.AddItem slStr
            mAddMsg slStr
            Print #hmMsg, slStr
            imFlightError = True
            'If Trim$(tmIcf.sErrorMess) = "" Then
            '    tmIcf.sErrorMess = "No air days for line "
            'Else
            '    tmIcf.sErrorMess = tmIcf.sErrorMess & "; "
            'End If
            'tmIcf.sErrorMess = tmIcf.sErrorMess & Trim$(Str$(tgClfImpt(ilLnIndex).ClfRec.iLine))
            Exit Sub
        End If
    End If
    tgCffImpt(ilUpper).CffRec.sDelete = "N"       'Y=Yes(flight-line deleted); N=No
    'gPDNToStr tgClfImpt(ilLnIndex).ClfRec.sActPrice, 2, slActPrice
    slActPrice = smFieldValues(15)
    tgCffImpt(ilUpper).CffRec.lActPrice = gStrDecToLong(slActPrice, 2)
    tgCffImpt(ilUpper).CffRec.lPropPrice = 0
    tgCffImpt(ilUpper).CffRec.lBBPrice = 0
    tgCffImpt(ilUpper).CffRec.sPriceType = tgClfImpt(ilLnIndex).ClfRec.sPriceType
    tgCffImpt(ilUpper).lRecPos = 0
    tgCffImpt(ilUpper).iStatus = 0
    tgCffImpt(ilUpper).iNextCff = -1
    If Not ilReplace Then
        ilUpper = ilUpper + 1
        ReDim Preserve tgCffImpt(1 To ilUpper) As CFFLIST
        tgCffImpt(ilUpper).lRecPos = 0
        tgCffImpt(ilUpper).iStatus = -1
        tgCffImpt(ilUpper).iNextCff = -1
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mProcessHeader                  *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Process header type record      *
'*                                                     *
'*******************************************************
Private Sub mProcessHeader()
    'First determine if this is a new contract or a change
    'to a existing contract.
    'If a change, determine what have been altered and set header change
    'flag
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slLastName As String
    Dim slFirstName As String
    Dim ilFound As Integer
    Dim slDemo As String
    Dim ilDemo As Integer
    tgChfImpt.lCode = 0    'Internal code number for Contract
    tgChfImpt.lCntrNo = Val(smFieldValues(1))  'Contract number
    'tmIcf.sType = "1"
    'tmIcf.lCntrNo = Val(smFieldValues(2))  'Contract number
    tgChfImpt.sAgyEstNo = ""  'Agenct estimate number
    'tgChfImpt.lPurcOrderNo = Val(smFieldValues(27))  'Purchase Order number
    tgChfImpt.iExtRevNo = Val(smFieldValues(2))  'Master Contract number
    tgChfImpt.iPkageGenMeth = 0 'weekly dollar generation
    gPackDate smFieldValues(3), tgChfImpt.iOHDDate(0), tgChfImpt.iOHDDate(1)
    gPackTime smFieldValues(4), tgChfImpt.iOHDTime(0), tgChfImpt.iOHDTime(1)
    tgChfImpt.iCntRevNo = Val(smFieldValues(2))   'Revision number
    tgChfImpt.iPropVer = 0
    tgChfImpt.iPropDate(0) = 0
    tgChfImpt.iPropDate(1) = 0
    tgChfImpt.iPropTime(0) = 0
    tgChfImpt.iPropTime(1) = 0
    tgChfImpt.sType = smFieldValues(5) 'Type of contract: P=Proposal; H=Hold; V=Reservation; J=Rejection;
                        'E=Order; C=Contract; B=Allocation contract; D=Deferred Contract;
                        'T=Remnant; R=Direct Response; Q=Per inQuiry; S=PSA; M=Promo
    If StrComp(Trim$(smFieldValues(15)), "DIRECT ADVERTISER", 1) <> 0 Then
        tgChfImpt.iAgfCode = mGetAgencyCode(smFieldValues(15), smFieldValues(16), smFieldValues(17), smFieldValues(18), smFieldValues(19), smFieldValues(20), smFieldValues(21), smFieldValues(22))
        'Test if match exist in file
        tgChfImpt.iAdfCode = mGetAdvertiserCode(smFieldValues(23), "A", "", "", "", "", "", "")
    Else
        tgChfImpt.iAgfCode = 0
        tgChfImpt.iAdfCode = mGetAdvertiserCode(smFieldValues(23), "D", smFieldValues(24), smFieldValues(25), smFieldValues(26), smFieldValues(27), smFieldValues(28), smFieldValues(29))
        smAgyCreditRestr = ""
    End If
    'tmIcf.sAdvtName = smFieldValues(11)
    tgChfImpt.sProduct = smFieldValues(30)   'Product (default to advertiser product)
    'tmIcf.sProduct = tgChfImpt.sProduct
    tgChfImpt.lSifCode = 0
    For ilLoop = 0 To 9 Step 1
        tgChfImpt.iSlfCode(ilLoop) = 0
        tgChfImpt.lComm(ilLoop) = 0
        tgChfImpt.iMnfSubCmpy(ilLoop) = 0
    Next ilLoop
    slLastName = gChangeCase(smFieldValues(34), 2)
    slFirstName = gChangeCase(smFieldValues(33), 2)
    tgChfImpt.iSlfCode(0) = mGetSalespersonCode(slLastName, slFirstName) 'Salesperson code number (if direct: default to advt otherwisw default to agency)
    If tgChfImpt.iSlfCode(0) > 0 Then
        tgChfImpt.lComm(0) = 100 * Val(smFieldValues(39))
    End If
    slLastName = gChangeCase(smFieldValues(36), 2)
    slFirstName = gChangeCase(smFieldValues(35), 2)
    tgChfImpt.iSlfCode(1) = mGetSalespersonCode(slLastName, slFirstName) 'Salesperson code number (if direct: default to advt otherwisw default to agency)
    If tgChfImpt.iSlfCode(1) > 0 Then
        tgChfImpt.lComm(1) = 100 * Val(smFieldValues(40))
    End If
    slLastName = gChangeCase(smFieldValues(38), 2)
    slFirstName = gChangeCase(smFieldValues(37), 2)
    tgChfImpt.iSlfCode(2) = mGetSalespersonCode(slLastName, slFirstName) 'Salesperson code number (if direct: default to advt otherwisw default to agency)
    If tgChfImpt.iSlfCode(2) > 0 Then
        tgChfImpt.lComm(2) = 100 * Val(smFieldValues(41))
    End If
    If Trim$(smFieldValues(45)) <> "" Then
        tgChfImpt.iMnfComp(0) = mGetCompCode(smFieldValues(45))
    Else
        tgChfImpt.iMnfComp(0) = 0
    End If
    tgChfImpt.iMnfComp(1) = 0   'mGetCompCode(slName) 'Competitive code number
    tgChfImpt.iMnfExcl(0) = 0 'Program exclusions
    tgChfImpt.iMnfExcl(1) = 0 'Program exclusions
    tgChfImpt.sBuyer = smFieldValues(31)   'Buyers name (default name from advertiser)
    tgChfImpt.sPhone = smFieldValues(32) & "______________"   'Buyers Phone number and extension(if direct: default to advt otherwisw default to agency)
    tgChfImpt.sPriceCntr = "Y"  'Y=Show prices on contract; N=Don't show prices
    tgChfImpt.sPriceInv = "Y" 'Y=Show prices on invoices; N=Don't show prices
    'slStr = ""
    'gStrToPDN slStr, 0, 2, tgChfImpt.sPctBudget
    tgChfImpt.iPctBudget = 0
    For ilLoop = 0 To 4 Step 1
        tgChfImpt.iMnfRevSet(ilLoop) = 0 '
    Next ilLoop
    tgChfImpt.sCppCpm = "N"    'P=CPP; M=CPM; N=N/A
    tgChfImpt.iMerchPct = 0
    tgChfImpt.iPromoPct = 0
    For ilLoop = 0 To 9 Step 1
        tgChfImpt.iSlspCommPct(ilLoop) = tgChfImpt.lComm(ilLoop) \ 100
    Next ilLoop
    For ilLoop = 0 To 3 Step 1
        tgChfImpt.iMnfDemo(ilLoop) = 0    'First-four Demo target
        tgChfImpt.lTarget(ilLoop) = 0
    Next ilLoop
    For ilDemo = 0 To 2 Step 1
        slDemo = ""
        If Trim$(smFieldValues(42 + ilDemo)) <> "" Then
            For ilLoop = 1 To Len(smFieldValues(42 + ilDemo)) Step 1
                slStr = Mid$(smFieldValues(42 + ilDemo), ilLoop, 1)
                If slStr <> " " Then
                    'tgChfImpt.sDemo(0) = Trim$(tgChfImpt.sDemo(0)) & slStr
                    slDemo = slDemo & slStr
                End If
            Next ilLoop
            ilFound = False
            For ilLoop = LBound(tgDMnf) To UBound(tgDMnf) - 1 Step 1
                'If StrComp(Trim$(tgDMnf(ilLoop).sName), Trim$(tgChfImpt.sDemo(0)), 1) = 0 Then
                If StrComp(Trim$(tgDMnf(ilLoop).sName), Trim$(slDemo), 1) = 0 Then
                    tgChfImpt.iMnfDemo(ilDemo) = tgDMnf(ilLoop).iCode
                    tgChfImpt.sCppCpm = "P"
                    Exit For
                End If
            Next ilLoop
        End If
    Next ilDemo
    If smFieldValues(6) = "T" Then
        slStr = smFieldValues(7)    '"100"
    Else
        slStr = "0"
    End If
    'gStrToPDN slStr, 0, 2, tgChfImpt.sPctTrade
    tgChfImpt.iPctTrade = gStrDecToInt(slStr, 0)
    tgChfImpt.iRcfCode = imRcfCode 'Rate Card code
    slStr = smFieldValues(46)
    'gStrToPDN slStr, 2, 6, tgChfImpt.sInputGross
    tgChfImpt.lInputGross = gStrDecToLong(slStr, 2)
    'slStr = smFieldValues(26)
    'gStrToPDN slStr, 2, 6, tmIcf.sGross
    tgChfImpt.sBillCycle = "S" 'S=Standard; C=Calendar; W=Weekly
    tgChfImpt.sInvGp = "A"       'Invoicing grouping flag: A= all Spots; P= per Product; T= per Tag
    'slStr = ""
    'gStrToPDN slStr, 2, 3, tgChfImpt.sPctTag
    tgChfImpt.iPctTag = 0
    tgChfImpt.lCxfCode = 0        'Comment code number
    tgChfImpt.lCxfChgR = 0        'Comment code number
    tgChfImpt.lCxfInt = 0        'Comment code number
    tgChfImpt.lCxfMerch = 0        'Comment code number
    tgChfImpt.lCxfProm = 0        'Comment code number
    tgChfImpt.lCxfCanc = 0        'Comment code number
    tgChfImpt.iPropVer = 0
    If tgSpf.sInvAirOrder = "O" Then
        tgChfImpt.sMGMiss = "M"
    Else
        tgChfImpt.sMGMiss = "G"
    End If
    If imPorO = 0 Then
        tgChfImpt.sSchStatus = "P"
        tgChfImpt.sStatus = "W"
    Else
        Select Case tgChfImpt.sType
            Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
                tgChfImpt.sSchStatus = "M"
                tgChfImpt.sStatus = "O"
            Case Else   'Standard or Reservation
                tgChfImpt.sStatus = "N" 'N is type O but needs scheduling
                tgChfImpt.sSchStatus = "A"  '"N"
        End Select
    End If
    tgChfImpt.sTitle = ""
    tgChfImpt.iDtNeed(0) = 0
    tgChfImpt.iDtNeed(1) = 0
    tgChfImpt.iMnfBus = 0
    tgChfImpt.lGuar = 0
    tgChfImpt.sUnused3 = ""
    For ilLoop = 0 To 6 Step 1
        tgChfImpt.iMnfCmpy(ilLoop) = 0
        tgChfImpt.iCmpyPct(ilLoop) = 0
    Next ilLoop
    tgChfImpt.sResvNew = "N"
    tgChfImpt.lChfCode = 0
    tgChfImpt.iMnfPotnType = 0
    'If (Trim$(smFieldValues(30)) <> "") And (imPorO = 0) Then
    '    For ilLoop = LBound(tgPotMnf) To UBound(tgPotMnf) - 1 Step 1
    '        'If StrComp(Trim$(tgDMnf(ilLoop).sName), Trim$(tgChfImpt.sDemo(0)), 1) = 0 Then
    '        If StrComp(Trim$(tgPotMnf(ilLoop).sName), Trim$(smFieldValues(30)), 1) = 0 Then
    '            tgChfImpt.iMnfPotnType = tgPotMnf(ilLoop).iCode
    '            Exit For
    '        End If
    '    Next ilLoop
    'End If
    tgChfImpt.sPrint = "N"    'Contract print state: N=New and not printed; C=Changed and not printed; P=already printed (don't ask revision number until printed)
    tgChfImpt.sDiscrep = "N"
    tgChfImpt.sNewBus = "N"   'New business state: Y= advertiser new business flag is Y-> set it to N; N= not new business
    tgChfImpt.sAgyCTrade = "N"  'Trades Commissionable
    tgChfImpt.iStartDate(0) = 0
    tgChfImpt.iStartDate(1) = 0
    tgChfImpt.iEndDate(0) = 0
    tgChfImpt.iEndDate(1) = 0
    tgChfImpt.lVefCode = 0 'Vehicle Code (- => combo. note: Vehicle of all lines.  If one does not exist for all lines it will be created)
    tgChfImpt.iurfCode = 2     'Last user who modified contract
    'sSellNet As String * 1  'Y= selling
    'tgChfImpt.sSellNet = "N"
    If (StrComp(Trim$(smFieldValues(48)), "Cancel_O", 1) = 0) Or (StrComp(Trim$(smFieldValues(48)), "Cancel_H", 1) = 0) Then
        tgChfImpt.sCBSOrder = "C"   'remove bonus spots
    Else
        tgChfImpt.sCBSOrder = "N"
    End If
    tgChfImpt.iHdChg = 0   'Bit field indicating header changes
                        ' 0     Advertiser changed
                        ' 1     Competitive # 1 changed
                        ' 2     Competitive # 2 changed
                        ' 3     Exclusion # 1 changed
                        ' 4     Exclusion # 2 changed
                        ' 5     Status P (Prevent) changed
                        ' 6     Status M (Manual) changed
    tgChfImpt.iAlertDate(0) = 0
    tgChfImpt.iAlertDate(1) = 0
    tgChfImpt.iAlertTime(0) = 0
    tgChfImpt.iAlertTime(1) = 0
    tgChfImpt.sAlertFlag = "Y"
    tgChfImpt.iSlfLock = 0
    tgChfImpt.iLockDate(0) = 0
    tgChfImpt.iLockDate(1) = 0
    tgChfImpt.sDelete = "N"
    tgChfImpt.lAudMGAdj = 0

    ReDim tgCffImpt(1 To 1) As CFFLIST      'CFF record image
    tgCffImpt(1).lRecPos = 0
    tgCffImpt(1).iStatus = -1
    tgCffImpt(1).iNextCff = -1
    ReDim tgClfImpt(1 To 1) As CLFLIST      'CLF record image
    tgClfImpt(1).lRecPos = 0
    tgClfImpt(1).iStatus = -1
    tgClfImpt(1).iFirstCff = -1
    'If tgChfImpt.iAdfCode > 0 Then
    '    tgChfImpt.lSifCode = mGetSifCode(smFieldValues(21), tgChfImpt.iAdfCode)
    'Else
        tgChfImpt.lSifCode = 0
    'End If

    'ilRet = btrInsert(hmChf, tgChfImpt, imChfRecLen, INDEXKEY0)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mProcessLine                    *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Process line type record        *
'*                                                     *
'*            Note: the line version number is         *
'*                  increment if this is not a new line*
'*                                                     *
'*******************************************************
Private Sub mProcessLine()
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    ilUpper = UBound(tgClfImpt)
    tgClfImpt(ilUpper).ClfRec.lChfCode = tgChfImpt.lCode        'Contract code
    tgClfImpt(ilUpper).ClfRec.iLine = Val(smFieldValues(2))        'Line number
    tgClfImpt(ilUpper).ClfRec.iCntRevNo = Val(smFieldValues(3))     'Version number
    tgClfImpt(ilUpper).ClfRec.iPropVer = 0     'Version number
    tgClfImpt(ilUpper).ClfRec.iVefCode = mGetVehicleCode(smFieldValues(4), smFieldValues(5))
    'scan rate card program for matching vehicle and times
    tgClfImpt(ilUpper).ClfRec.sBB = "N"       'BB: N=N/A; O=Open BB; C=Close BB; B= Open/Close BB;
                            'F=Floating; 1=1 Open/week; 2=1 Close/week; 4=1 Open & Close/Week;
                            '3=1 Floating/Week; A=Any
    tgClfImpt(ilUpper).ClfRec.sExtra = "N"    'Extra Spot Types: N=N/a; C=Commercial Promo (create breaks if not found);
                            'D=Donuts(two spots within break- any position but one spot between);
                            'B=Bookends (two spots within break-first and last)
    tgClfImpt(ilUpper).ClfRec.sPgmTime = "T"  'If library buy retain spots with P=Program buy; T=Time buy when library moved
    tgClfImpt(ilUpper).ClfRec.iBreak = 0       'Break number if buying a specific break
    tgClfImpt(ilUpper).ClfRec.iPosition = 0    'Position number if buying a specific break and position
    'Override times
    gPackTime smFieldValues(6), tgClfImpt(ilUpper).ClfRec.iStartTime(0), tgClfImpt(ilUpper).ClfRec.iStartTime(1)
    gPackTime smFieldValues(7), tgClfImpt(ilUpper).ClfRec.iEndTime(0), tgClfImpt(ilUpper).ClfRec.iEndTime(1)
    tgClfImpt(ilUpper).ClfRec.iRdfcode = 0     'Rate Card Program/Time Code
    tgClfImpt(ilUpper).ClfRec.iNoGames = 0     'Number of games to distribute spot across
    tgClfImpt(ilUpper).ClfRec.iSpotsOrdered = 0    'Number of spots to be disttibuted across the number of games
    tgClfImpt(ilUpper).ClfRec.iSpotsBooked = 0 'Total spots booked for a game buy
    'tgClfImpt(ilUpper).ClfRec.iSpotsWrite = 0  'Number of spots written onto the log (for late weekly buys only)
    tgClfImpt(ilUpper).ClfRec.iSpotsWrite = Val(smFieldValues(2))        'Save Original Line number
    'Scan fligt
    tgClfImpt(ilUpper).ClfRec.sCntPct = "1"  'Entered # of spots/week and spots/day are by: 1= Daily by count; 2=weekly by count;3=Daily by %; 4=weekly by %
                            'The scheduler will convert 3 and 4 to 1 & 2. 3 and 4 are for
                            'reservation contract only
    tgClfImpt(ilUpper).ClfRec.iLen = Val(smFieldValues(15))         'Spot length
    tgClfImpt(ilUpper).ClfRec.sPreempt = "P"  'N=Non-preemptible (spot can't be moved when scheduling other spots);
                            'P=Preemptible
    'slStr = smFieldValues(9)
    'gStrToPDN slStr, 2, 5, tgClfImpt(ilUpper).ClfRec.sActPrice
    'slStr = ""
    'gStrToPDN slStr, 2, 5, tgClfImpt(ilUpper).ClfRec.sPropPrice
    'slStr = ""
    'gStrToPDN slStr, 2, 5, tgClfImpt(ilUpper).ClfRec.sHighPrice
    'slStr = ""
    'gStrToPDN slStr, 2, 5, tgClfImpt(ilUpper).ClfRec.sAvgPrice
    'slStr = ""
    'gStrToPDN slStr, 2, 5, tgClfImpt(ilUpper).ClfRec.sLowPrice
    'slStr = ""
    'gStrToPDN slStr, 2, 5, tgClfImpt(ilUpper).ClfRec.sBBPrice
    tgClfImpt(ilUpper).ClfRec.sPriceType = "T" 'T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
    tgClfImpt(ilUpper).ClfRec.iPriority = -1    'Priority flag (calculated)
    tgClfImpt(ilUpper).ClfRec.lCxfCode = 0        'Comment code
    Select Case tgChfImpt.sType
        Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
            tgClfImpt(ilUpper).ClfRec.sSchStatus = "M"    'Manual
        Case Else   'Contract; Reservation
            tgClfImpt(ilUpper).ClfRec.sSchStatus = "A"  '"N"    'New scheduling required
    End Select
    tgClfImpt(ilUpper).ClfRec.iurfCode = 2     'Last user who modified line
    tgClfImpt(ilUpper).ClfRec.sDelete = "N"   'Y=Yes(line deleted); N=No
    tgClfImpt(ilUpper).ClfRec.iEntryDate(0) = tgChfImpt.iOHDDate(0)
    tgClfImpt(ilUpper).ClfRec.iEntryDate(1) = tgChfImpt.iOHDDate(1)
    tgClfImpt(ilUpper).ClfRec.iEntryTime(0) = tgChfImpt.iOHDTime(0)
    tgClfImpt(ilUpper).ClfRec.iEntryTime(1) = tgChfImpt.iOHDTime(1)
    tgClfImpt(ilUpper).ClfRec.iMnfDemo = 0
    tgClfImpt(ilUpper).ClfRec.lCPM = 0
    tgClfImpt(ilUpper).ClfRec.lCPP = 0
    tgClfImpt(ilUpper).ClfRec.lGrImp = 0
    tgClfImpt(ilUpper).ClfRec.iDnfCode = 0
    tgClfImpt(ilUpper).ClfRec.iRdfcode = 0 'Set in SetLnRdf
    tgClfImpt(ilUpper).ClfRec.iStartDate(0) = 0
    tgClfImpt(ilUpper).ClfRec.iStartDate(1) = 0
    tgClfImpt(ilUpper).ClfRec.iEndDate(0) = 0
    tgClfImpt(ilUpper).ClfRec.iEndDate(1) = 0
    tgClfImpt(ilUpper).ClfRec.iMnfSocEco = 0
    tgClfImpt(ilUpper).ClfRec.iPkLineNo = 0
    If smFieldValues(4) = "P" Then
        tgClfImpt(ilUpper).ClfRec.sType = "O"   'Virtual
    ElseIf smFieldValues(4) = "H" Then
        tgClfImpt(ilUpper).ClfRec.sType = "H"   'Hidden
        tgClfImpt(ilUpper).ClfRec.iPkLineNo = 100 * (tgClfImpt(ilUpper).ClfRec.iLine \ 100)
    Else
        tgClfImpt(ilUpper).ClfRec.sType = "S"
    End If
    tgClfImpt(ilUpper).lRecPos = 0
    tgClfImpt(ilUpper).iStatus = 0
    tgClfImpt(ilUpper).iFirstCff = -1
    tgClfImpt(ilUpper).iCancel = False
    tgClfImpt(ilUpper).iOverride = True
    tgClfImpt(ilUpper).iGame = False
    tgClfImpt(ilUpper).iLibBuy = False
    tgClfImpt(ilUpper).iPriChgd = 0
    If Trim$(smFieldValues(18)) <> "" Then
        For ilLoop = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1 Step 1
            If StrComp(Trim$(tgAvailAnf(ilLoop).sName), Trim$(smFieldValues(18)), 1) = 0 Then
                tgClfImpt(ilUpper).iPriChgd = tgAvailAnf(ilLoop).iCode
                Exit For
            End If
        Next ilLoop
    End If
    ilUpper = ilUpper + 1
    ReDim Preserve tgClfImpt(1 To ilUpper) As CLFLIST
    tgClfImpt(ilUpper).lRecPos = 0
    tgClfImpt(ilUpper).iStatus = -1
    tgClfImpt(ilUpper).iFirstCff = -1
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadMoveRecord                 *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove Move Record             *
'*                                                     *
'*******************************************************
Private Function mReadMoveRecord(llToDate As Long) As Integer
    Dim ilRet As Integer
    Dim slMove As String
    Dim ilLoop As Integer
    Dim ilEof As Integer
    Dim ilUpper As Integer
    Dim slStr As String
    'Remove any processed records
    ilLoop = LBound(tgMoveRec)
    Do
        If ilLoop >= UBound(tgMoveRec) Then
            Exit Do
        End If
        If tgMoveRec(ilLoop).iStatus = 1 Then
            For ilUpper = ilLoop To UBound(tgMoveRec) - 1 Step 1
                tgMoveRec(ilUpper) = tgMoveRec(ilUpper + 1)
            Next ilUpper
            ReDim Preserve tgMoveRec(0 To UBound(tgMoveRec) - 1) As MGMOVEREC
        Else
            ilLoop = ilLoop + 1
        End If
    Loop
    Do
        ilRet = 0
        On Error GoTo mReadMoveRecordErr:
        Line Input #hmMove, slMove
        On Error GoTo 0
        If ilRet <> 0 Then
            mReadMoveRecord = True
            Exit Function
        End If
        If Len(slMove) > 0 Then
            If (Asc(slMove) = 26) Then     'Ctrl Z
                mReadMoveRecord = True
                Exit Function
            Else
                DoEvents
                slMove = mFilter(slMove)
                gParseCDFields slMove, True, smMoveValues()    'Change case
                ilUpper = UBound(tgMoveRec)
                tgMoveRec(ilUpper).iStatus = 0
                tgMoveRec(ilUpper).lCreatedDate = gDateValue(smMoveValues(1))
                tgMoveRec(ilUpper).lRefTrackID = Val(smMoveValues(2))
                tgMoveRec(ilUpper).lTrackID = Val(smMoveValues(3))
                tgMoveRec(ilUpper).lTransGpID = Val(smMoveValues(4))
                tgMoveRec(ilUpper).lCntrNo = Val(smMoveValues(5))
                tgMoveRec(ilUpper).iLineNo = Val(smMoveValues(6))
                tgMoveRec(ilUpper).sFromVehicle = smMoveValues(7)
                tgMoveRec(ilUpper).iFromVefCode = -1
                For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    If (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S") Then
                        If StrComp(smMoveValues(7), Trim$(tgMVef(ilLoop).sName), 1) = 0 Then
                            tgMoveRec(ilUpper).iFromVefCode = tgMVef(ilLoop).iCode
                            Exit For
                        End If
                    End If
                Next ilLoop
                tgMoveRec(ilUpper).lFromDate = gDateValue(smMoveValues(8))
                tgMoveRec(ilUpper).lFromSTime = CLng(gTimeToCurrency(smMoveValues(9), False))
                tgMoveRec(ilUpper).lFromETime = CLng(gTimeToCurrency(smMoveValues(10), True))
                tgMoveRec(ilUpper).sToVehicle = smMoveValues(11)
                tgMoveRec(ilUpper).iToVefCode = -1
                For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    If (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S") Then
                        If StrComp(smMoveValues(11), Trim$(tgMVef(ilLoop).sName), 1) = 0 Then
                            tgMoveRec(ilUpper).iToVefCode = tgMVef(ilLoop).iCode
                            Exit For
                        End If
                    End If
                Next ilLoop
                tgMoveRec(ilUpper).lToDate = gDateValue(smMoveValues(12))
                If smMoveValues(13) <> "" Then
                    tgMoveRec(ilUpper).lToSTime = gTimeToLong(smMoveValues(13), False)
                Else
                    tgMoveRec(ilUpper).lToSTime = -1
                End If
                If smMoveValues(14) <> "" Then
                    tgMoveRec(ilUpper).lToETime = gTimeToLong(smMoveValues(14), True)
                Else
                    tgMoveRec(ilUpper).lToETime = -1
                End If
                tgMoveRec(ilUpper).sOper = UCase$(smMoveValues(15))
                tgMoveRec(ilUpper).lFromPrice = gStrDecToLong(smMoveValues(16), 2)
                tgMoveRec(ilUpper).lToPrice = gStrDecToLong(smMoveValues(17), 2)
                tgMoveRec(ilUpper).sOpResult = UCase$(smMoveValues(18))
                tgMoveRec(ilUpper).lCreatedTime = gTimeToLong(smMoveValues(20), False)
                If Trim$(smMoveValues(21)) <> "" Then
                    tgMoveRec(ilUpper).iSpotLen = Val(smMoveValues(21))
                Else
                    tgMoveRec(ilUpper).iSpotLen = 0 'Undefined
                End If
                If Trim$(smMoveValues(22)) <> "" Then
                    For ilLoop = 0 To 6 Step 1
                        tgMoveRec(ilUpper).iDays(ilLoop) = 0
                    Next ilLoop
                    slStr = smMoveValues(22)
                    Do
                        tgMoveRec(ilUpper).iDays(Val(Left$(slStr, 1)) - 1) = 1
                        slStr = Mid$(slStr, 2)
                    Loop While Len(slStr) > 0
                Else
                    For ilLoop = 0 To 6 Step 1
                        tgMoveRec(ilUpper).iDays(ilLoop) = 1
                    Next ilLoop
                End If
                tgMoveRec(ilUpper).lMoveID = Val(smMoveValues(19))
                tgMoveRec(ilUpper).lRecNo = lmMGRecNo
                tgMoveRec(ilUpper).iFromAnfCode = 0
                If Trim$(smMoveValues(23)) <> "" Then
                    For ilLoop = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1 Step 1
                        If StrComp(Trim$(tgAvailAnf(ilLoop).sName), Trim$(smMoveValues(23)), 1) = 0 Then
                            tgMoveRec(ilUpper).iFromAnfCode = tgAvailAnf(ilLoop).iCode
                            Exit For
                        End If
                    Next ilLoop
                End If
                tgMoveRec(ilUpper).iToAnfCode = 0
                If Trim$(smMoveValues(24)) <> "" Then
                    For ilLoop = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1 Step 1
                        If StrComp(Trim$(tgAvailAnf(ilLoop).sName), Trim$(smMoveValues(24)), 1) = 0 Then
                            tgMoveRec(ilUpper).iToAnfCode = tgAvailAnf(ilLoop).iCode
                            Exit For
                        End If
                    Next ilLoop
                End If
                lmMGRecNo = lmMGRecNo + 1
                ReDim Preserve tgMoveRec(0 To ilUpper + 1) As MGMOVEREC
                If tgMoveRec(ilUpper).lCreatedDate > llToDate Then
                    Exit Do
                End If
            End If
        End If
    Loop Until ilEof
    mReadMoveRecord = True
    Exit Function
mReadMoveRecordErr:
    ilRet = Err.Number
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mRemoveSpot                     *
'*                                                     *
'*             Created:1/30/00       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove spot and create MTF     *
'*                                                     *
'*******************************************************
Private Function mRemoveSpot(llEnteredDate As Long, ilFromVefCode As Integer, llTrackID As Long, llRefTrackID As Long, llTransGpID As Long, slExtraMTFType As String) As Integer
'
'   Input tmSdf anf tmSmf
'
    Dim ilRet As Integer
    Dim llPrevTransGpID As Long
    Dim llDate As Long
    Dim ilLoop As Integer
    ReDim ilDays(0 To 6) As Integer
    If tmSdf.lCode <> 0 Then
        If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
            tmSmfSrchKey2.lCode = tmSdf.lCode
            ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                If tmSmf.lMtfCode <> 0 Then
                    tmMtfSrchKey.lCode = tmSmf.lMtfCode
                    ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        'slFPrice = gLongToStrDec(tmMtf.lFromPrice, 2)
                        'slTPrice = gLongToStrDec(tmMtf.lToPrice, 2)
                        If (slExtraMTFType = "M") Then
                            llTransGpID = tmMtf.lTransGpID  'Retain the previous Group ID
                        End If
                        llPrevTransGpID = tmMtf.lTransGpID
                        For ilLoop = 0 To 6 Step 1
                            ilDays(ilLoop) = tmMtf.iDays(ilLoop)
                        Next ilLoop
                        ilRet = mMakeMTF(tmSdf.lChfCode, tmSdf.iLineNo, tmSdf.iVefCode, tmMtf.lTrackID, llEnteredDate, tmMtf.lRefTrackID, tmMtf.lFromPrice, tmMtf.lToPrice, llTransGpID, tmSdf.sSchStatus, llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, ilDays())
                        'PrevTransGpID could be zero, but using the one from MTF
                        If slExtraMTFType = "C" Then
                            ilRet = mMakeMTF(tmChf.lCode, tmSdf.iLineNo, ilFromVefCode, llTrackID, tmMoveRec.lCreatedDate, llRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, llTransGpID, tmSdf.sSchStatus, llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
                        ElseIf (slExtraMTFType = "S") Then
                            ilRet = mMakeMTF(tmChf.lCode, tmSdf.iLineNo, ilFromVefCode, llTrackID, tmMoveRec.lCreatedDate, llRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, llTransGpID, tmSdf.sSchStatus, llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
                        ElseIf (slExtraMTFType = "M") Then
                            ilRet = mMakeMTF(tmChf.lCode, tmSdf.iLineNo, ilFromVefCode, llTrackID, tmMoveRec.lCreatedDate, llRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, llTransGpID, tmSdf.sSchStatus, llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
                        End If
                        tmMtfSrchKey.lCode = tmSmf.lMtfCode
                        ilRet = btrGetEqual(hmMtf, tmMtf, imMtfRecLen, tmMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            ilRet = btrDelete(hmMtf)
                        End If
                    Else
                        igBtrError = ilRet
                        mRemoveSpot = False
                        Exit Function
                    End If
                Else
                    For ilLoop = 0 To 6 Step 1
                        ilDays(ilLoop) = 0
                    Next ilLoop
                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                    ilDays(gWeekDayLong(llDate)) = 1
                    ilRet = mMakeMTF(tmChf.lCode, tmSdf.iLineNo, ilFromVefCode, llTrackID, tmMoveRec.lCreatedDate, llRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, llTransGpID, tmSdf.sSchStatus, 0, tmMoveRec.lToSTime, tmMoveRec.lToETime, ilDays())
                End If
            Else
                igBtrError = ilRet
                mRemoveSpot = False
                Exit Function
            End If
        Else
            'Days are not used as reschd is not a mg.
            For ilLoop = 0 To 6 Step 1
                ilDays(ilLoop) = 1
            Next ilLoop
            ilRet = mMakeMTF(tmChf.lCode, tmSdf.iLineNo, ilFromVefCode, llTrackID, tmMoveRec.lCreatedDate, llRefTrackID, tmMoveRec.lFromPrice, tmMoveRec.lToPrice, llTransGpID, tmSdf.sSchStatus, 0, tmMoveRec.lToSTime, tmMoveRec.lToETime, ilDays())
        End If
        lgSsfDate(0) = 0
        If Not gChgSchSpot("D", hmSdf, tmSdf, hmSmf, 0, tmSmf, hmSsf, tgSsf(0), lgSsfDate(0), lgSsfRecPos(0)) Then
            mRemoveSpot = False
            Exit Function
        End If
    End If
    mRemoveSpot = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Save contract (header, line,    *
'*                     and flights)                    *
'*                                                     *
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
    Dim slStr As String
    Dim ilClf As Integer
    Dim ilCff As Integer
    Dim ilFound As Integer
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llTestDate As Long
    Dim slDate As String
    Dim slAdfLimit As String
    Dim slAdfCurrAR As String
    Dim slAdfUnbilled As String
    Dim slAgfLimit As String
    Dim slAgfCurrAR As String
    Dim slAgfUnbilled As String
    Dim slNetAmount As String
    Dim slAdjNetAmount As String
    Dim slAvgNetAmount As String
    Dim slGrossAmount As String
    Dim slAgyRate As String
    Dim slPctTrade As String
    Dim ilAdfCreditOk As Integer
    Dim ilAgfCreditOk As Integer
    Dim slNoPeriods As String
    Dim llInputGross As Long

    'If imNewAdvt Then
    '    'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
    '    'tmIcf.iLineNo = 0
    '    'tmIcf.iFlightNo = 0
    '    'slStr = "Advertiser: " & smAdvtName & " for " & Trim$(Str$(tmIcf.lCntrNo)) & " added" '& Trim$(Str$(tgChfImpt.lCntrNo)) & " missing"
    '    slStr = "Advertiser: " & smAdvtName & " for " & Trim$(Str$(tgChfImpt.lCntrNo)) & " added" '& Trim$(Str$(tgChfImpt.lCntrNo)) & " missing"
    '    lbcErrors.AddItem slStr
    '    Print #hmMsg, slStr
    '    'tmIcf.sErrorMess = "New " & smAdvtName
    '    'tmIcf.sType = "1"
    '    'tmIcf.sStatus = "A"
    '    'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
    'End If
    'If imNewAgy Then
    '    'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
    '    'tmIcf.iLineNo = 0
    '    'tmIcf.iFlightNo = 0
    '    'slStr = "Agency: " & smAgyName & ", " & smAgyCity & " for " & Trim$(Str$(tmIcf.lCntrNo)) & " added"   ' & Trim$(Str$(tgChfImpt.lCntrNo)) & " missing"
    '    slStr = "Agency: " & smAgyName & ", " & smAgyCity & " for " & Trim$(Str$(tgChfImpt.lCntrNo)) & " added"   ' & Trim$(Str$(tgChfImpt.lCntrNo)) & " missing"
    '    lbcErrors.AddItem slStr
    '    Print #hmMsg, slStr
    '    'tmIcf.sErrorMess = "New " & smAgyName & ", " & smAgyCity & " at " & smAgyAddr1
    '    'tmIcf.sType = "1"
    '    'tmIcf.sStatus = "A"
    '    'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
    'End If
    If imFlightError Then
        mSaveRec = False
        imSchCntr = False
        'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
        'tmIcf.iLineNo = 0
        'tmIcf.iFlightNo = 0
        ''Set in
        ''tmIcf.sErrorMess = "Contract previously entered"
        'tmIcf.sType = "1"
        'tmIcf.sStatus = "T"
        'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
        'mClearIcf
        Exit Function
    End If
    If (tgChfImpt.iAdfCode = -1) Or (tgChfImpt.iAgfCode = -1) Then
        mSaveRec = False
        imSchCntr = False
        'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
        'tmIcf.iLineNo = 0
        'tmIcf.iFlightNo = 0
        'MAI/Sirius "F"
        If (tgChfImpt.iAdfCode = -1) And (tgChfImpt.iAgfCode = -1) Then
            'slStr = "Advertiser: " & smAdvtName & " Agency: " & smAgyName & ", " & smAgyCity & " for " & Trim$(Str$(tmIcf.lCntrNo)) & " missing"'& Trim$(Str$(tgChfImpt.lCntrNo)) & " missing"
            slStr = "Advertiser: " & smAdvtName & " Agency: " & smAgyName & ", " & smAgyCity & " for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " Advertiser and Agency name missing from import, Contract not added" '& Trim$(Str$(tgChfImpt.lCntrNo)) & " missing"
            'lbcErrors.AddItem slStr
            mAddMsg slStr
            Print #hmMsg, slStr
        '    tmIcf.sErrorMess = "Missing " & smAdvtName & ", " & smAgyName & ", " & smAgyCity
        ElseIf (tgChfImpt.iAdfCode = -1) Then
            'slStr = "Advertiser: " & smAdvtName & " for " & Trim$(Str$(tmIcf.lCntrNo)) & " missing" '& Trim$(Str$(tgChfImpt.lCntrNo)) & " missing"
            slStr = "Advertiser: " & smAdvtName & " for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " Advertiser name missing from Import, Contract not added"  '& Trim$(Str$(tgChfImpt.lCntrNo)) & " missing"
            'lbcErrors.AddItem slStr
            mAddMsg slStr
            Print #hmMsg, slStr
        '    tmIcf.sErrorMess = "Missing " & smAdvtName
        Else
            'slStr = "Agency: " & smAgyName & ", " & smAgyCity & " for " & Trim$(Str$(tmIcf.lCntrNo)) & " missing"   ' & Trim$(Str$(tgChfImpt.lCntrNo)) & " missing"
            slStr = "Agency: " & smAgyName & ", " & smAgyCity & " for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " Agency name missing from import, Contract not added"    ' & Trim$(Str$(tgChfImpt.lCntrNo))
            'lbcErrors.AddItem slStr
            mAddMsg slStr
            Print #hmMsg, slStr
        '    tmIcf.sErrorMess = "Missing " & smAgyName & ", " & smAgyCity & " at " & smAgyAddr1
        End If
        'tmIcf.sType = "1"
        'tmIcf.sStatus = "T"
        'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
        'mClearIcf
        Exit Function
    End If
    If (tgChfImpt.iAdfCode = -2) Or (tgChfImpt.iAgfCode = -2) Then
        mSaveRec = False
        imSchCntr = False
    End If
    If (tgChfImpt.iSlfCode(0) = -2) Or (tgChfImpt.iSlfCode(1) = -2) Or (tgChfImpt.iSlfCode(2) = -2) Then
        mSaveRec = False
        imSchCntr = False
    End If
    If tgChfImpt.iMnfComp(0) = -2 Then
        mSaveRec = False
        imSchCntr = False
    End If

'    If (smAdfCreditRestr = "P") Or (smAgyCreditRestr = "P") Then
'        mSaveRec = True
'        imSchCntr = False
'        'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
'        'tmIcf.iLineNo = 0
'        'tmIcf.iFlightNo = 0
'        If (smAdfCreditRestr = "P") And (smAgyCreditRestr = "P") Then
'            'slStr = "Advertiser: " & smAdvtName & " Agency: " & smAgyName & ", " & smAgyCity & " for " & Trim$(Str$(tmIcf.lCntrNo)) & " No New Orders"  ' & Trim$(Str$(tgChfImpt.lCntrNo)) & " No New Orders"
'            slStr = "Advertiser: " & smAdvtName & " Agency: " & smAgyName & ", " & smAgyCity & " for " & Trim$(Str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(Str$(tgChfImpt.iExtRevNo)) & " No New Orders, Contract not added"   ' & Trim$(Str$(tgChfImpt.lCntrNo)) & " No New Orders"
'            'lbcErrors.AddItem slStr
'            mAddMsg slStr
'            Print #hmMsg, slStr
'        '    tmIcf.sErrorMess = "No new orders " & smAdvtName & ", " & smAgyName & ", " & smAgyCity
'        ElseIf (smAdfCreditRestr = "P") Then
'            'slStr = "Advertiser: " & smAdvtName & " for " & Trim$(Str$(tmIcf.lCntrNo)) & " no new orders"   ' & Trim$(Str$(tgChfImpt.lCntrNo)) & " no new orders"
'            slStr = "Advertiser: " & smAdvtName & " for " & Trim$(Str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(Str$(tgChfImpt.iExtRevNo)) & " no new orders, Contract not added"    ' & Trim$(Str$(tgChfImpt.lCntrNo)) & " no new orders"
'            'lbcErrors.AddItem slStr
'            mAddMsg slStr
'            Print #hmMsg, slStr
'        '    tmIcf.sErrorMess = "No new orders " & smAdvtName
'        Else
'            'slStr = "Agency: " & smAgyName & ", " & smAgyCity & " for " & Trim$(Str$(tmIcf.lCntrNo)) & " no new order"  '
'            slStr = "Agency: " & smAgyName & ", " & smAgyCity & " for " & Trim$(Str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(Str$(tgChfImpt.iExtRevNo)) & " no new order, Contract not added"   '
'            'lbcErrors.AddItem slStr
'            mAddMsg slStr
'            Print #hmMsg, slStr
'        '    tmIcf.sErrorMess = "No new orders " & smAgyName & ", " & smAgyCity & " at " & smAgyAddr1
'        End If
'        'tmIcf.sType = "1"
'        'tmIcf.sStatus = "T"
'        'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
'        'mClearIcf
'        Exit Function
'    End If
    If UBound(tgClfImpt) <= 1 Then
        'slStr = "No Lines defined for " & Trim$(Str$(tmIcf.lCntrNo))
        slStr = "No Lines defined for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo))
        'lbcErrors.AddItem slStr
        mAddMsg slStr
        Print #hmMsg, slStr
        mSaveRec = False    'True
        imSchCntr = False
        'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
        'tmIcf.sErrorMess = "No lines defined"
        'tmIcf.sType = "1"
        'tmIcf.sStatus = "T"
        'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
        ''MsgBox slStr, vbOkOnly + vbCritical + vbApplicationModal, "Error"
        ''mSaveRec = False
        'mClearIcf
        Exit Function
    End If
    llInputGross = tgChfImpt.lInputGross
    imMatchLnUpper = UBound(tgClfImpt)
    'Test if contract exist- if so bypass (later- code as update to existing contract)
    tmChfSrchKey.lCntrNo = tgChfImpt.lCntrNo
    tmChfSrchKey.iCntRevNo = 32000
    tmChfSrchKey.iPropVer = 32000
    'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)
    ilRet = btrGetGreaterOrEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tgChfImpt.lCntrNo)
        If tgChfImpt.iCntRevNo = tmChf.iCntRevNo Then
            slStr = smBlankSpaces & smBlankSpaces & "Contract " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tmChf.iCntRevNo)) & " previously imported, Contract not added"
            'lbcErrors.AddItem slStr
            mAddMsg slStr
            Print #hmMsg, slStr
            mSaveRec = True
            imSchCntr = False
            'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
            'tmIcf.sErrorMess = smNameMsg & " type ,  " & tmChf.sType & " not matching " & tgChfImpt.sType
            'tmIcf.sType = "1"
            'tmIcf.sStatus = "T"
            'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
            'mClearIcf
            Exit Function
        End If
        ilRet = btrGetNext(hmChf, tmChf, imChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    'Test if contract exist- if so bypass (later- code as update to existing contract)
    tmChfSrchKey.lCntrNo = tgChfImpt.lCntrNo
    tmChfSrchKey.iCntRevNo = 32000
    tmChfSrchKey.iPropVer = 32000
    'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)
    ilRet = btrGetGreaterOrEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)
    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tgChfImpt.lCntrNo) Then   'Contract found
        If tmChf.sSchStatus <> "F" Then
            slStr = "Contract " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " must be scheduled prior to import, Contract not added"
            'lbcErrors.AddItem slStr
            mAddMsg slStr
            Print #hmMsg, slStr
            mSaveRec = False    'True
            imSchCntr = False
            'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
            'tmIcf.sErrorMess = smNameMsg & " type ,  " & tmChf.sType & " not matching " & tgChfImpt.sType
            'tmIcf.sType = "1"
            'tmIcf.sStatus = "T"
            'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
            'mClearIcf
            Exit Function
        End If
        ''mSaveRec = True
        ''tmIcf.iSeqNo = tmIcf.iSeqNo + 1
        ''tmIcf.sErrorMess = "Contract previously entered"
        ''tmIcf.sType = "1"
        ''tmIcf.sStatus = "T"
        ''ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
        ''mClearIcf
        ''Exit Function
        'Merge the two contracts
        If tmChf.sType <> tgChfImpt.sType Then
            slStr = "Type Not Matching " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " Contract not added"
            'lbcErrors.AddItem slStr
            mAddMsg slStr
            Print #hmMsg, slStr
            mSaveRec = False    'True
            imSchCntr = False
            'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
            'tmIcf.sErrorMess = smNameMsg & " type ,  " & tmChf.sType & " not matching " & tgChfImpt.sType
            'tmIcf.sType = "1"
            'tmIcf.sStatus = "T"
            'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
            'mClearIcf
            Exit Function
        End If
        If ((tgChfImpt.sStatus = "N") And (tmChf.sStatus <> "O") And (tmChf.sStatus <> "H")) Or ((tgChfImpt.sStatus = "O") And (tmChf.sStatus <> tgChfImpt.sStatus)) Then
            slStr = "Status Not Matching " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iCntRevNo)) & " Contract not added"
            'lbcErrors.AddItem slStr
            mAddMsg slStr
            Print #hmMsg, slStr
            mSaveRec = False    'True
            imSchCntr = False
            'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
            'tmIcf.sErrorMess = smNameMsg & " status ,  " & tmChf.sStatus & " not matching " & tgChfImpt.sStatus
            'tmIcf.sType = "1"
            'tmIcf.sStatus = "T"
            'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
            'mClearIcf
            Exit Function
        End If
        If tmChf.iAdfCode <> tgChfImpt.iAdfCode Then
            tmAdfSrchKey.iCode = tmChf.iAdfCode
            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
            'tmIcf.sErrorMess = smNameMsg & " previously entered,  " & smAdvtName & " not matching " & Trim$(tmAdf.sName)
            'tmIcf.sType = "1"
            'tmIcf.sStatus = "W"
            'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
            'tmIcf.sErrorMess = ""
            tgChfImpt.iHdChg = tgChfImpt.iHdChg Or HDADVTCHG
        End If
        If tmChf.iAgfCode <> tgChfImpt.iAgfCode Then
            tmAgfSrchKey.iCode = tmChf.iAgfCode
            ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
            'tmIcf.sErrorMess = smNameMsg & " previously entered, " & smAgyName & ", " & smAgyCity & " not matching" & Trim$(tmAgf.sName) & ", " & Trim$(tmAgf.sCityID)
            'tmIcf.sType = "1"
            'tmIcf.sStatus = "W"
            'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
            'tmIcf.sErrorMess = ""
        End If
        If tmChf.iMnfComp(0) <> tgChfImpt.iMnfComp(0) Then
            tgChfImpt.iHdChg = tgChfImpt.iHdChg Or HDCOMP1CHG
        End If
        If tmChf.iMnfComp(1) <> tgChfImpt.iMnfComp(1) Then
            tgChfImpt.iHdChg = tgChfImpt.iHdChg Or HDCOMP2CHG
        End If
        If tmChf.iMnfExcl(0) <> tgChfImpt.iMnfExcl(0) Then
            tgChfImpt.iHdChg = tgChfImpt.iHdChg Or HDEXCL1CHG
        End If
        If tmChf.iMnfExcl(1) <> tgChfImpt.iMnfExcl(1) Then
            tgChfImpt.iHdChg = tgChfImpt.iHdChg Or HDEXCL2CHG
        End If
    End If
    'Move Package line times into hidden line times 3/9/01- MAI
    'All hidden lines must be for the same daypart as the package
    For ilClf = LBound(tgClfImpt) To imMatchLnUpper - 1 Step 1
        If (tgClfImpt(ilClf).iStatus = 0) And (tgClfImpt(ilClf).ClfRec.sType = "O") Then
            For ilLoop = LBound(tgClfImpt) To imMatchLnUpper - 1 Step 1
                If (tgClfImpt(ilLoop).iStatus = 0) And (tgClfImpt(ilLoop).ClfRec.sType = "H") And (tgClfImpt(ilClf).ClfRec.iLine = tgClfImpt(ilLoop).ClfRec.iPkLineNo) Then
                    tgClfImpt(ilLoop).ClfRec.iStartTime(0) = tgClfImpt(ilClf).ClfRec.iStartTime(0)
                    tgClfImpt(ilLoop).ClfRec.iStartTime(1) = tgClfImpt(ilClf).ClfRec.iStartTime(1)
                    tgClfImpt(ilLoop).ClfRec.iEndTime(0) = tgClfImpt(ilClf).ClfRec.iEndTime(0)
                    tgClfImpt(ilLoop).ClfRec.iEndTime(1) = tgClfImpt(ilClf).ClfRec.iEndTime(1)
                End If
            Next ilLoop
        End If
    Next ilClf
    For ilClf = LBound(tgClfImpt) To imMatchLnUpper - 1 Step 1
        If tgClfImpt(ilClf).ClfRec.iVefCode = -2 Then
            mSaveRec = False
            imSchCntr = False
        End If
    Next ilClf
    'End of move times from package to hidden lines
    For ilLoop = LBound(tgClfImpt) To imMatchLnUpper - 1 Step 1
        If tgClfImpt(ilLoop).iFirstCff = -1 Then
            slStr = "No flights defined for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & ", Line # " & Trim$(str$(tgClfImpt(ilLoop).ClfRec.iLine)) & " Contract not added"
            'lbcErrors.AddItem slStr
            mAddMsg slStr
            Print #hmMsg, slStr
            mSaveRec = False    'True
            imSchCntr = False
            'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
            'tmIcf.sErrorMess = "No Flights defined for line " & Trim$(Str$(tgClfImpt(ilLoop).ClfRec.iLine))
            'tmIcf.sType = "1"
            'tmIcf.sStatus = "T"
            'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
            'mClearIcf
            ''MsgBox slStr, vbOkOnly + vbCritical + vbApplicationModal, "Error"
            ''mSaveRec = False
            Exit Function
        End If
        If tgClfImpt(ilLoop).ClfRec.iVefCode <= 0 Then
            slStr = "Vehicle(s) missing for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & ", Line # " & Trim$(str$(tgClfImpt(ilLoop).ClfRec.iLine)) & " Contract not added"
            'lbcErrors.AddItem slStr
            mAddMsg slStr
            Print #hmMsg, slStr
            mSaveRec = False    'True
            imSchCntr = False
            'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
            'tmIcf.sErrorMess = "No Flights defined for line " & Trim$(Str$(tgClfImpt(ilLoop).ClfRec.iLine))
            'tmIcf.sType = "1"
            'tmIcf.sStatus = "T"
            'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
            'mClearIcf
            ''MsgBox slStr, vbOkOnly + vbCritical + vbApplicationModal, "Error"
            ''mSaveRec = False
            Exit Function
        End If
        If Not mSetLnRdf(tgClfImpt(ilLoop)) Then
            slStr = "No Rate Card Daypart defined for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & ", Line # " & Trim$(str$(tgClfImpt(ilLoop).ClfRec.iLine)) & " Contract not added"
            'lbcErrors.AddItem slStr
            mAddMsg slStr
            Print #hmMsg, slStr
            mSaveRec = False    'True
            imSchCntr = False
            'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
            'tmIcf.sErrorMess = "No Rate Card Program/Time defined for line " & Trim$(Str$(tgClfImpt(ilLoop).ClfRec.iLine))
            'tmIcf.sType = "1"
            'tmIcf.sStatus = "T"
            'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
            'mClearIcf
            Exit Function
        End If
        gUnpackTime tgClfImpt(ilLoop).ClfRec.iStartTime(0), tgClfImpt(ilLoop).ClfRec.iStartTime(1), "A", "1", slStartTime
        llStartTime = gTimeToLong(slStartTime, False)
        gUnpackTime tgClfImpt(ilLoop).ClfRec.iEndTime(0), tgClfImpt(ilLoop).ClfRec.iEndTime(1), "A", "1", slEndTime
        llEndTime = gTimeToLong(slEndTime, True)
        'If llEndTime < llStartTime Then
        If (llEndTime < llStartTime) And (tgClfImpt(ilLoop).ClfRec.sType <> "O") Then
            slStr = "End Time prior to Start Time for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & ", Line # " & Trim$(str$(tgClfImpt(ilLoop).ClfRec.iLine)) & " Contract not added"
            'lbcErrors.AddItem slStr
            mAddMsg slStr
            Print #hmMsg, slStr
            mSaveRec = False    'True
            imSchCntr = False
            'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
            'tmIcf.sErrorMess = End Time prior to Start Time for line " & Trim$(Str$(tgClfImpt(ilLoop).ClfRec.iLine))
            'tmIcf.sType = "1"
            'tmIcf.sStatus = "T"
            'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
            'mClearIcf
            Exit Function
        End If
    Next ilLoop
    imSchCntr = True
    slStr = smBlankSpaces & smBlankSpaces & "Contract " & str$(tgChfImpt.lCntrNo) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " being added"
    Print #hmMsg, slStr
    mSetHdDate
    'gPDNToStr tgChfImpt.sPctTrade, 0, slPctTrade
    slPctTrade = gIntToStrDec(tgChfImpt.iPctTrade, 0)
    If Val(slPctTrade) <> 100 Then  'Ignore trades
        If (smAdfCreditRestr <> "N") Or ((smAgyCreditRestr <> "N") And (smAgyCreditRestr <> "")) Then
            slAgyRate = ""
            If smAdfCreditRestr <> "N" Then
                tmAdfSrchKey.iCode = tgChfImpt.iAdfCode
                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    'gPDNToStr tmAdf.sCreditLimit, 2, slAdfLimit
                    slAdfLimit = gLongToStrDec(tmAdf.lCreditLimit, 2)
                    gPDNToStr tmAdf.sCurrAR, 2, slAdfCurrAR
                    gPDNToStr tmAdf.sUnbilled, 2, slAdfUnbilled
                Else
                    slAdfLimit = ""
                    slAdfCurrAR = ""
                    slAdfUnbilled = "0"
                End If
            End If
            If (tgChfImpt.iAgfCode > 0) Then
                tmAgfSrchKey.iCode = tgChfImpt.iAgfCode
                ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    'gPDNToStr tmAgf.sCreditLimit, 2, slAgfLimit
                    slAgfLimit = gLongToStrDec(tmAgf.lCreditLimit, 2)
                    gPDNToStr tmAgf.sCurrAR, 2, slAgfCurrAR
                    gPDNToStr tmAgf.sUnbilled, 2, slAgfUnbilled
                    'gPDNToStr tmAgf.sComm, 2, slAgyRate
                    slAgyRate = gIntToStrDec(tmAgf.iComm, 2)
                Else
                    slAgfLimit = ""
                    slAgfCurrAR = ""
                    slAgfUnbilled = "0"
                End If
            Else
                slAgfLimit = ""
                slAgfCurrAR = ""
                slAgfUnbilled = "0"
            End If
            'Credit Limit
            'gPDNToStr tgChfImpt.sInputGross, 2, slGrossAmount
            slGrossAmount = gLongToStrDec(llInputGross, 2)
            If slAgyRate = "" Then
                slNetAmount = slGrossAmount
            Else
                slNetAmount = gDivStr(gMulStr(slGrossAmount, gSubStr("100.00", slAgyRate)), "100.00")
            End If
            If (smAdfCreditRestr = "L") And (tgSpf.iRNoWks > 0) Then
                'Must have a credit if any Gross Amount
                gUnpackDate tgChfImpt.iStartDate(0), tgChfImpt.iStartDate(1), slStartDate
                gUnpackDate tgChfImpt.iEndDate(0), tgChfImpt.iEndDate(1), slEndDate
                slEndDate = Trim$(str$(gDateValue(slEndDate)))
                slStartDate = Trim$(str$(gDateValue(slStartDate)))
                slNoPeriods = gDivStr(gSubStr(slEndDate, slStartDate), "7")
                If Val(slNoPeriods) <= 0 Then
                    slNoPeriods = "1"
                End If
                slAdjNetAmount = slNetAmount
                If Val(slEndDate) >= (lmNowDate + tgSpf.iRNoWks * 7) Then
                    slAvgNetAmount = gDivStr(slNetAmount, slNoPeriods)
                    If Val(slStartDate) > lmNowDate Then
                        slNoPeriods = gDivStr(gSubStr(slStartDate, Trim$(str$(lmNowDate))), "7")
                        If Val(slNoPeriods) > tgSpf.iRNoWks Then
                            slNoPeriods = "0"
                        Else
                            slNoPeriods = gSubStr(Trim$(str$(tgSpf.iRNoWks)), slNoPeriods)
                        End If
                    Else
                        slNoPeriods = Trim$(str$(tgSpf.iRNoWks))
                    End If
                    slAdjNetAmount = gMulStr(slAvgNetAmount, slNoPeriods)
                End If
                ilAdfCreditOk = False
                If gCompNumberStr(slAdjNetAmount, ".00") > 0 Then
                    If tgSpf.sRCurrAmt = "Y" Then
                        slStr = gAddStr(slAdfCurrAR, slAdfUnbilled)
                    Else
                        slStr = slAdfUnbilled
                    End If
                    slStr = gAddStr(slStr, slAdjNetAmount)
                    If InStr(slStr, "-") = 0 Then
                        If gCompNumberStr(slStr, slAdfLimit) <= 0 Then
                            ilAdfCreditOk = True
                        End If
                    Else
                        ilAdfCreditOk = True
                    End If
                Else
                    ilAdfCreditOk = True
                End If
'                If Not ilAdfCreditOk Then
'                    'mSaveRec = True
'                    'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
'                    'tmIcf.iLineNo = 0
'                    'tmIcf.iFlightNo = 0
'                    slStr = smBlankSpaces & "Advertiser: " & smAdvtName & " for " & Trim$(Str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(Str$(tgChfImpt.iExtRevNo)) & " Warning- Exceeded Credit Limit"
'                    'lbcErrors.AddItem slStr
'                    mAddMsg slStr
'                    Print #hmMsg, slStr
'                    'tmIcf.sErrorMess = "Warning- Exceeded Credit Limit " & smAdvtName
'                    'tmIcf.sType = "1"
'                    'tmIcf.sStatus = "W"
'                    'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
'                    'tmIcf.sErrorMess = ""
'                    ''mClearIcf
'                    ''Exit Function
'                End If
            End If
            If (smAdfCreditRestr = "W") Or (smAdfCreditRestr = "M") Or (smAdfCreditRestr = "T") Then
                If smAdfCreditRestr = "W" Then
                    gUnpackDate tgChfImpt.iStartDate(0), tgChfImpt.iStartDate(1), slStartDate
                    gUnpackDate tgChfImpt.iEndDate(0), tgChfImpt.iEndDate(1), slEndDate
                    slEndDate = Trim$(str$(gDateValue(slEndDate)))
                    slStartDate = Trim$(str$(gDateValue(slStartDate)))
                    slNoPeriods = gDivStr(gSubStr(slEndDate, slStartDate), "7")
                    If Val(slNoPeriods) <= 0 Then
                        slNoPeriods = "1"
                    End If
                    slAvgNetAmount = gDivStr(slNetAmount, slNoPeriods)
                ElseIf smAdfCreditRestr = "M" Then
                    gUnpackDate tgChfImpt.iStartDate(0), tgChfImpt.iStartDate(1), slStartDate
                    gUnpackDate tgChfImpt.iEndDate(0), tgChfImpt.iEndDate(1), slEndDate
                    slEndDate = Trim$(str$(gDateValue(slEndDate)))
                    slStartDate = Trim$(str$(gDateValue(slStartDate)))
                    slNoPeriods = gDivStr(gSubStr(slEndDate, slStartDate), "28")
                    If Val(slNoPeriods) <= 0 Then
                        slNoPeriods = "1"
                    End If
                    slAvgNetAmount = gDivStr(slNetAmount, slNoPeriods)
                Else
                    slAvgNetAmount = slNetAmount
                End If
                'Must have a credit if any Gross Amount
                ilAdfCreditOk = False
                If gCompNumberStr(slAvgNetAmount, ".00") > 0 Then
                    slStr = gAddStr(slAdfCurrAR, slAdfUnbilled)
                    If InStr(slStr, "-") > 0 Then
                        If gCompAbsNumberStr(slAvgNetAmount, slStr) <= 0 Then
                            ilAdfCreditOk = True
                        End If
                    End If
                Else
                    ilAdfCreditOk = True
                End If
'                If Not ilAdfCreditOk Then
'                    mSaveRec = True
'                    'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
'                    'tmIcf.iLineNo = 0
'                    'tmIcf.iFlightNo = 0
'                    slStr = "Advertiser: " & smAdvtName & " for " & Trim$(Str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(Str$(tgChfImpt.iExtRevNo)) & " Insufficient Cash, Contract not added"
'                    'lbcErrors.AddItem slStr
'                    mAddMsg slStr
'                    Print #hmMsg, slStr
'                    'tmIcf.sErrorMess = "Insufficient Cash " & smAdvtName
'                    'tmIcf.sType = "1"
'                    'tmIcf.sStatus = "T"
'                    'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
'                    'mClearIcf
'                    Exit Function
'                End If
            End If
            If tgChfImpt.iAgfCode > 0 Then
                If (smAgyCreditRestr = "L") And (tgSpf.iRNoWks > 0) Then
                    'Must have a credit if any Gross Amount
                    gUnpackDate tgChfImpt.iStartDate(0), tgChfImpt.iStartDate(1), slStartDate
                    gUnpackDate tgChfImpt.iEndDate(0), tgChfImpt.iEndDate(1), slEndDate
                    slEndDate = Trim$(str$(gDateValue(slEndDate)))
                    slStartDate = Trim$(str$(gDateValue(slStartDate)))
                    slNoPeriods = gDivStr(gSubStr(slEndDate, slStartDate), "7")
                    If Val(slNoPeriods) <= 0 Then
                        slNoPeriods = "1"
                    End If
                    slAdjNetAmount = slNetAmount
                    If Val(slEndDate) >= (lmNowDate + tgSpf.iRNoWks * 7) Then
                        slAvgNetAmount = gDivStr(slNetAmount, slNoPeriods)
                        If Val(slStartDate) > lmNowDate Then
                            slNoPeriods = gDivStr(gSubStr(slStartDate, Trim$(str$(lmNowDate))), "7")
                            If Val(slNoPeriods) > tgSpf.iRNoWks Then
                                slNoPeriods = "0"
                            Else
                                slNoPeriods = gSubStr(Trim$(str$(tgSpf.iRNoWks)), slNoPeriods)
                            End If
                        Else
                            slNoPeriods = Trim$(str$(tgSpf.iRNoWks))
                        End If
                        slAdjNetAmount = gMulStr(slAvgNetAmount, slNoPeriods)
                    End If
                    ilAgfCreditOk = False
                    If gCompNumberStr(slAdjNetAmount, ".00") > 0 Then
                        If tgSpf.sRCurrAmt = "Y" Then
                            slStr = gAddStr(slAgfCurrAR, slAgfUnbilled)
                        Else
                            slStr = slAgfUnbilled
                        End If
                        slStr = gAddStr(slStr, slAdjNetAmount)
                        If InStr(slStr, "-") = 0 Then
                            If gCompNumberStr(slStr, slAgfLimit) <= 0 Then
                                ilAgfCreditOk = True
                            End If
                        Else
                            ilAgfCreditOk = True
                        End If
                    Else
                        ilAgfCreditOk = True
                    End If
'                    If Not ilAgfCreditOk Then
'                        ''mSaveRec = True
'                        'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
'                        'tmIcf.iLineNo = 0
'                        'tmIcf.iFlightNo = 0
'                        slStr = smBlankSpaces & "Agency: " & smAgyName & ", " & smAgyCity & " for " & Trim$(Str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(Str$(tgChfImpt.iExtRevNo)) & " Warning- Exceeded Credit Limit"
'                        'lbcErrors.AddItem slStr
'                        mAddMsg slStr
'                        Print #hmMsg, slStr
'                        'tmIcf.sErrorMess = "Warning- Exceeded Credit Limit " & smAgyName & ", " & smAgyCity
'                        'tmIcf.sType = "1"
'                        'tmIcf.sStatus = "W"
'                        'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
'                        'tmIcf.sErrorMess = ""
'                        ''mClearIcf
'                        ''Exit Function
'                    End If
                End If
                If (smAgyCreditRestr = "W") Or (smAgyCreditRestr = "M") Or (smAgyCreditRestr = "T") Then
                    If smAgyCreditRestr = "W" Then
                        gUnpackDate tgChfImpt.iStartDate(0), tgChfImpt.iStartDate(1), slStartDate
                        gUnpackDate tgChfImpt.iEndDate(0), tgChfImpt.iEndDate(1), slEndDate
                        slEndDate = Trim$(str$(gDateValue(slEndDate)))
                        slStartDate = Trim$(str$(gDateValue(slStartDate)))
                        slNoPeriods = gDivStr(gSubStr(slEndDate, slStartDate), "7")
                        If Val(slNoPeriods) <= 0 Then
                            slNoPeriods = "1"
                        End If
                        slAvgNetAmount = gDivStr(slNetAmount, slNoPeriods)
                    ElseIf smAgyCreditRestr = "M" Then
                        gUnpackDate tgChfImpt.iStartDate(0), tgChfImpt.iStartDate(1), slStartDate
                        gUnpackDate tgChfImpt.iEndDate(0), tgChfImpt.iEndDate(1), slEndDate
                        slEndDate = Trim$(str$(gDateValue(slEndDate)))
                        slStartDate = Trim$(str$(gDateValue(slStartDate)))
                        slNoPeriods = gDivStr(gSubStr(slEndDate, slStartDate), "28")
                        If Val(slNoPeriods) <= 0 Then
                            slNoPeriods = "1"
                        End If
                        slAvgNetAmount = gDivStr(slNetAmount, slNoPeriods)
                    Else
                        slAvgNetAmount = slNetAmount
                    End If
                    'Must have a credit if any Gross Amount
                    ilAgfCreditOk = False
                    If gCompNumberStr(slAvgNetAmount, ".00") > 0 Then
                        slStr = gAddStr(slAgfCurrAR, slAgfUnbilled)
                        If InStr(slStr, "-") > 0 Then
                            If gCompAbsNumberStr(slAvgNetAmount, slStr) <= 0 Then
                                ilAgfCreditOk = True
                            End If
                        End If
                    Else
                        ilAgfCreditOk = True
                    End If
'                    If Not ilAgfCreditOk Then
'                        mSaveRec = True
'                        'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
'                        'tmIcf.iLineNo = 0
'                        'tmIcf.iFlightNo = 0
'                        slStr = "Agency: " & smAgyName & ", " & smAgyCity & " for " & Trim$(Str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(Str$(tgChfImpt.iExtRevNo)) & " Insufficient Cash, Contract not added"
'                        'lbcErrors.AddItem slStr
'                        mAddMsg slStr
'                        Print #hmMsg, slStr
'                        'tmIcf.sErrorMess = "Insufficient Cash " & smAgyName & ", " & smAgyCity
'                        'tmIcf.sType = "1"
'                        'tmIcf.sStatus = "T"
'                        'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
'                        'mClearIcf
'                        Exit Function
'                    End If
                End If
            End If
        End If
    End If
    'tmIcf.sErrorMess = ""
    'Moved prior to testing for across midnight 6/30/00
    ''Set rate card program/times
    'For ilLoop = LBound(tgClfImpt) To imMatchLnUpper - 1 Step 1
    '    If Not mSetLnRdf(tgClfImpt(ilLoop)) Then
    '        slStr = "No Rate Card Daypart defined for " & Trim$(Str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(Str$(tgChfImpt.iExtRevNo)) & ", Line # " & Trim$(Str$(tgClfImpt(ilLoop).ClfRec.iLine)) & " Contract not added"
    '        lbcErrors.AddItem slStr
    '        Print #hmMsg, slStr
    '        mSaveRec = True
    '        'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
    '        'tmIcf.sErrorMess = "No Rate Card Program/Time defined for line " & Trim$(Str$(tgClfImpt(ilLoop).ClfRec.iLine))
    '        'tmIcf.sType = "1"
    '        'tmIcf.sStatus = "T"
    '        'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
    '        'mClearIcf
    '        Exit Function
    '    End If
    'Next ilLoop
    If Not mSetHdVeh() Then
        mSaveRec = False    'True 'Ignore error
        imSchCntr = False
        'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
        'tmIcf.sErrorMess = "Can't set vehicles for " & smNameMsg & " header"
        'tmIcf.sType = "1"
        'tmIcf.sStatus = "T"
        'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
        'mClearIcf
        ''mSaveRec = False
        Exit Function
    End If
    'mSetHdDate
    If Trim$(tgChfImpt.sProduct) <> "" Then
        ilFound = False
        tmPrfSrchKey.iAdfCode = tgChfImpt.iAdfCode
        ilRet = btrGetGreaterOrEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
        Do While (ilRet = BTRV_ERR_NONE) And (tmPrf.iAdfCode = tgChfImpt.iAdfCode)
            If StrComp(Trim$(tgChfImpt.sProduct), Trim$(tmPrf.sName), 1) = 0 Then
                ilFound = True
                tgChfImpt.sProduct = tmPrf.sName    'Use name from file incase of differ case
                'If StrComp(Trim$(tgChfImpt.sProduct), Trim$(tmPrf.sName), 0) <> 0 Then
                '    ilRet = btrGetPosition(hmPrf, llPrfRecPos)
                '    If ilRet = BTRV_ERR_NONE Then   'Contract found
                '        Do
                '            tmPrf.sName = tgChfImpt.sProduct
                '            ilRet = btrUpdate(hmPrf, tmPrf, imPrfRecLen)
                '            If ilRet = BTRV_ERR_CONFLICT Then
                '                ilCRet = btrGetDirect(hmPrf, tmPrf, imPrfRecLen, llPrfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                '                If ilCRet <> BTRV_ERR_NONE Then
                '                    Exit Do
                '                End If
                '            End If
                '        Loop While ilRet = BTRV_ERR_CONFLICT
                '    End If
                'End If
                Exit Do
            End If
            ilRet = btrGetNext(hmPrf, tmPrf, imPrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If Not ilFound Then
            Do  'Loop until record updated or added
                tmPrf.lCode = 0
                tmPrf.iAdfCode = tgChfImpt.iAdfCode
                tmPrf.sName = Trim$(tgChfImpt.sProduct)
                tmPrf.iMnfComp(0) = tgChfImpt.iMnfComp(0)
                tmPrf.iMnfComp(1) = tgChfImpt.iMnfComp(1)
                tmPrf.iMnfExcl(0) = tgChfImpt.iMnfExcl(0)
                tmPrf.iMnfExcl(1) = tgChfImpt.iMnfExcl(1)
                tmAdfSrchKey.iCode = tgChfImpt.iAdfCode
                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    tmPrf.iPnfBuyer = tmAdf.iPnfBuyer
                Else
                    tmPrf.iPnfBuyer = 0
                End If
                tmPrf.sCppCpm = tgChfImpt.sCppCpm
                For ilLoop = 0 To 3
                    tmPrf.iMnfDemo(ilLoop) = tgChfImpt.iMnfDemo(ilLoop)
                    tmPrf.lTarget(ilLoop) = tgChfImpt.lTarget(ilLoop)
                    tmPrf.lLastCPP(ilLoop) = 0
                    tmPrf.lLastCPM(ilLoop) = 0
                Next ilLoop
                tmPrf.sState = "A"
                tmPrf.iurfCode = 2 'Use first record retained for user
                tmPrf.iRemoteID = tgUrf(0).iRemoteUserID
                tmPrf.lAutoCode = tmPrf.lCode
                ilRet = btrInsert(hmPrf, tmPrf, imPrfRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    slStr = "Product " & Trim$(tgChfImpt.sProduct) & " not added for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " Error" & str$(ilRet)
                    'lbcErrors.AddItem slStr
                    mAddMsg slStr
                    Print #hmMsg, slStr
                    'ilRet = BTRV_ERR_NONE
                    mSaveRec = False    'True 'Ignore error
                    imSchCntr = False
                    Exit Function
                Else
                    tmPrf.iRemoteID = tgUrf(0).iRemoteUserID
                    tmPrf.lAutoCode = tmPrf.lCode
                    tmPrf.iSourceID = tgUrf(0).iRemoteUserID
                    gPackDate smSyncDate, tmPrf.iSyncDate(0), tmPrf.iSyncDate(1)
                    gPackTime smSyncTime, tmPrf.iSyncTime(0), tmPrf.iSyncTime(1)
                    ilRet = btrUpdate(hmPrf, tmPrf, imPrfRecLen)
                End If
                slMsg = "mSaveRec (btrInsert: Product)" & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo))
            Loop While ilRet = BTRV_ERR_CONFLICT
            'If mBtrvErrorMsg(ilRet, slMsg, True) Then
            '    mSaveRec = False
            '    'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
            '    'tmIcf.sErrorMess = "Btrieve error inserting advertiser product " & Trim$(Str$(ilRet))
            '    'tmIcf.sType = "1"
            '    'tmIcf.sStatus = "T"
            '    'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
            '    'mClearIcf
            '    Exit Function
            'End If
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                slStr = "Product " & Trim$(tgChfImpt.sProduct) & " not updated for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " Error" & str$(ilRet)
                'lbcErrors.AddItem slStr
                mAddMsg slStr
                Print #hmMsg, slStr
                'ilRet = BTRV_ERR_NONE
                mSaveRec = False    'True 'Ignore error
                imSchCntr = False
                Exit Function
            End If
        End If
    End If
    tgChfImpt.iurfCode = tgUrf(0).iCode 'Use first record retained for user
    'tgChfImpt.iRemoteID = tgUrf(0).iRemoteUserID
    'tgChfImpt.lAutoCode = 0
    ilRet = btrInsert(hmChf, tgChfImpt, imChfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        If ilRet >= 30000 Then
            ilRet = csiHandleValue(0, 7)
        End If
        slStr = "Contract " & str$(tgChfImpt.lCntrNo) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " not added" & " Error" & str$(ilRet)
        'lbcErrors.AddItem slStr
        mAddMsg slStr
        Print #hmMsg, slStr
        mSaveRec = False    'True
        imSchCntr = False
        Exit Function
    End If
    tgChfImpt.lSpotChfCode = 0
    'tgChfImpt.iRemoteID = tgUrf(0).iRemoteUserID
    'tgChfImpt.lAutoCode = tgChfImpt.lCode
    'tgChfImpt.iSourceID = tgUrf(0).iRemoteUserID
    'gPackDate smSyncDate, tgChfImpt.iSyncDate(0), tgChfImpt.iSyncDate(1)
    'gPackTime smSyncTime, tgChfImpt.iSyncTime(0), tgChfImpt.iSyncTime(1)
    'tgChfImpt.iOrigRemoteID = tgChfImpt.iRemoteID
    'tgChfImpt.lOrigAutoCode = tgChfImpt.lCode
    tgChfImpt.sNoAssigned = "Y"
    ilRet = btrUpdate(hmChf, tgChfImpt, imChfRecLen)
    slMsg = "mSaveRec (btrInsert: " & smNameMsg & ")" & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo))
    'If mBtrvErrorMsg(ilRet, slMsg, True) Then
    '    mSaveRec = False
    '    'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
    '    'tmIcf.sErrorMess = "Btrieve " & smNameMsg & " header insert error " & Trim$(Str$(ilRet))
    '    'tmIcf.sType = "1"
    '    'tmIcf.sStatus = "T"
    '    'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
    '    'mClearIcf
    '    Exit Function
    'End If
    If ilRet <> BTRV_ERR_NONE Then
        If ilRet >= 30000 Then
            ilRet = csiHandleValue(0, 7)
        End If
        slStr = "Contract " & str$(tgChfImpt.lCntrNo) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " not updated" & " Error" & str$(ilRet)
        'lbcErrors.AddItem slStr
        mAddMsg slStr
        Print #hmMsg, slStr
        'mSaveRec = True
        'Exit Function
        'MAI/Sirius: Added 10/4 Chaned to False and annd imSchCntr
        mSaveRec = False    'True
        imSchCntr = False
        Exit Function
    End If
    'tmIcf.sErrorMess = ""
    'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
    'tmIcf.sType = "1"
    'tmIcf.sStatus = "A"
    'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
    For ilClf = LBound(tgClfImpt) To imMatchLnUpper - 1 Step 1 'UBound(tgClfImpt) - 1 Step 1
        If tgClfImpt(ilClf).iStatus = 0 Then
            'tmIcf.sType = "2"
            'tmIcf.iLineNo = tgClfImpt(ilClf).ClfRec.iLine
            'tmIcf.iFlightNo = 0
            'Determine date range of flights
            llStartDate = 0
            llEndDate = 0
            ilCff = tgClfImpt(ilClf).iFirstCff
            Do While ilCff <> -1
                If tgCffImpt(ilCff).iStatus >= 0 Then
                    If llStartDate = 0 Then
                        gUnpackDateLong tgCffImpt(ilCff).CffRec.iStartDate(0), tgCffImpt(ilCff).CffRec.iStartDate(1), llStartDate
                        gUnpackDateLong tgCffImpt(ilCff).CffRec.iEndDate(0), tgCffImpt(ilCff).CffRec.iEndDate(1), llEndDate
                    Else
                        gUnpackDateLong tgCffImpt(ilCff).CffRec.iStartDate(0), tgCffImpt(ilCff).CffRec.iStartDate(1), llTestDate
                        If llTestDate < llStartDate Then
                            llStartDate = llTestDate
                        End If
                        gUnpackDateLong tgCffImpt(ilCff).CffRec.iEndDate(0), tgCffImpt(ilCff).CffRec.iEndDate(1), llTestDate
                        If llTestDate > llEndDate Then
                            llEndDate = llTestDate
                        End If
                    End If
                End If
                ilCff = tgCffImpt(ilCff).iNextCff
            Loop
            slDate = Format$(llStartDate, "m/d/yy")
            gPackDate slDate, tgClfImpt(ilClf).ClfRec.iStartDate(0), tgClfImpt(ilClf).ClfRec.iStartDate(1)
            slDate = Format$(llEndDate, "m/d/yy")
            gPackDate slDate, tgClfImpt(ilClf).ClfRec.iEndDate(0), tgClfImpt(ilClf).ClfRec.iEndDate(1)
            Do  'Loop until record updated or added
                tgClfImpt(ilClf).ClfRec.lChfCode = tgChfImpt.lCode
                tgClfImpt(ilClf).ClfRec.iCntRevNo = tgChfImpt.iCntRevNo
                ilRet = btrInsert(hmClf, tgClfImpt(ilClf).ClfRec, imClfRecLen, INDEXKEY0)
                slMsg = "mSaveRec (btrInsert: " & smNameMsg & " Line)" & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & ", Line " & Trim$(str$(tgClfImpt(ilClf).ClfRec.iLine))
            Loop While ilRet = BTRV_ERR_CONFLICT
            'If mBtrvErrorMsg(ilRet, slMsg, True) Then
            '    mSaveRec = False
            '    'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
            '    'tmIcf.sErrorMess = "Btrieve line insert error " & Trim$(Str$(ilRet))
            '    'tmIcf.sStatus = "P"
            '    'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
            '    'mClearIcf
            '    Exit Function
            'End If
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                slStr = "Contract " & str$(tgChfImpt.lCntrNo) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " Line # " & Trim$(str$(tgClfImpt(ilClf).ClfRec.iLine)) & " not added" & " Error" & str$(ilRet)
                'lbcErrors.AddItem slStr
                mAddMsg slStr
                Print #hmMsg, slStr
                'mSaveRec = True
                'Exit Function
                mSaveRec = False    'True
                imSchCntr = False
                Exit Function
            Else
                'mSaveRec = False
                ilRet = btrGetPosition(hmClf, tgClfImpt(ilClf).lRecPos)
                ilCff = tgClfImpt(ilClf).iFirstCff
                Do While ilCff <> -1
                    If tgCffImpt(ilCff).iStatus >= 0 Then
                        Do  'Loop until record updated or added
                            tgCffImpt(ilCff).CffRec.lChfCode = tgChfImpt.lCode
                            tgCffImpt(ilCff).CffRec.iClfLine = tgClfImpt(ilClf).ClfRec.iLine
                            tgCffImpt(ilCff).CffRec.iCntRevNo = tgClfImpt(ilClf).ClfRec.iCntRevNo
                            tgCffImpt(ilCff).CffRec.iPropVer = tgClfImpt(ilClf).ClfRec.iPropVer
                            tgCffImpt(ilCff).CffRec.sDelete = "N"
                            ilRet = btrInsert(hmCff, tgCffImpt(ilCff), imCffRecLen, INDEXKEY0)
                            slMsg = "mSaveRec (btrInsert: " & smNameMsg & " Line Flight)" & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & ", Line " & Trim$(str$(tgClfImpt(ilClf).ClfRec.iLine))
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        'gUnpackDate tgCffImpt(ilCff).CffRec.iStartDate(0), tgCffImpt(ilCff).CffRec.iStartDate(1), slStartDate
                        'gUnpackDate tgCffImpt(ilCff).CffRec.iEndDate(0), tgCffImpt(ilCff).CffRec.iEndDate(1), slEndDate
                        'If mBtrvErrorMsg(ilRet, slMsg, True) Then
                        '    mSaveRec = False
                        '    'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
                        '    'tmIcf.iFlightNo = tmIcf.iFlightNo + 1
                        '    'tmIcf.sErrorMess = "Btrieve flight insert error " & Trim$(Str$(ilRet))
                        '    'tmIcf.sStatus = "P"
                        '    'tmIcf.sFlightDates = slStartDate & "-" & slEndDate
                        '    'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
                        '    'mClearIcf
                        '    Exit Function
                        'End If
                        If ilRet <> BTRV_ERR_NONE Then
                            gUnpackDate tgCffImpt(ilCff).CffRec.iStartDate(0), tgCffImpt(ilCff).CffRec.iStartDate(1), slStartDate
                            gUnpackDate tgCffImpt(ilCff).CffRec.iEndDate(0), tgCffImpt(ilCff).CffRec.iEndDate(1), slEndDate
                            If ilRet >= 30000 Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            slStr = "Contract " & str$(tgChfImpt.lCntrNo) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " Line # " & Trim$(str$(tgClfImpt(ilClf).ClfRec.iLine)) & " Flight " & slStartDate & "-" & slEndDate & " not added" & " Error" & str$(ilRet)
                            'lbcErrors.AddItem slStr
                            mAddMsg slStr
                            Print #hmMsg, slStr
                            'mSaveRec = True
                            'Exit Function
                            mSaveRec = False    'True
                            imSchCntr = False
                            Exit Function
                        End If
                        'tmIcf.iSeqNo = tmIcf.iSeqNo + 1
                        'tmIcf.iFlightNo = tmIcf.iFlightNo + 1
                        'tmIcf.sStatus = "A"
                        'tmIcf.sFlightDates = slStartDate & "-" & slEndDate
                        'ilRet = btrInsert(hmIcf, tmIcf, imIcfRecLen, INDEXKEY0)
                        ilRet = btrGetPosition(hmCff, tgCffImpt(ilCff).lRecPos)
                    End If
                    ilCff = tgCffImpt(ilCff).iNextCff
                Loop
            End If
        End If
    Next ilClf
    mSaveRec = True
    'mClearIcf
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSchdUnschd                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Schedule spots missed          *
'*                                                     *
'*******************************************************
Private Function mSchdUnschd() As Integer
    Dim ilVehCode As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slDate As String
    Dim llDate As Long
    Dim ilVpfIndex As Integer
    Dim ilRet As Integer
    Dim llPercent As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slTime As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilFound As Integer
    Dim llSpotTime As Long
    Dim ilVeh As Integer
    ilRet = gObtainVef()
    'Remove Move; Compact; Preempt passes
    sgMovePass = "N"
    sgCompPass = "N"
    'Allow preempt 11/18/98 (required for rank to work when altered)
    smPreemptPass = sgPreemptPass
    'sgPreemptPass = "N"
    llPercent = 0
    plcGauge.Value = llPercent
    For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If (tgMVef(ilVeh).sType = "C") Or (tgMVef(ilVeh).sType = "S") Then
            lacCntr(0).Left = 120
            lacCntr(0).Width = 8190
            lacCntr(0).Caption = "Import Done, Processing Closing for " & Trim$(tgMVef(ilVeh).sName)
            lacCntr(1).Visible = False
            lacCntr(2).Visible = False
            lacCntr(3).Visible = False
            DoEvents
            ilVehCode = tgMVef(ilVeh).iCode
            ilVpfIndex = gVpfFind(ImptCntr, ilVehCode)
            slStartDate = edcSDate.Text
            slEndDate = Format$(gDateValue(slStartDate) + 7 * Val(edcNoWks.Text) - 1, "m/d/yy")
            slStartTime = "12AM"
            slEndTime = "12AM"
            llStartDate = gDateValue(slStartDate)
            llEndDate = gDateValue(slEndDate)
            llStartTime = CLng(gTimeToCurrency(slStartTime, False))
            llEndTime = CLng(gTimeToCurrency(slEndTime, True)) - 1
            ReDim lgReschSdfCode(1 To 1) As Long
            tmSdfSrchKey1.iVefCode = ilVehCode
            slDate = Format$(llStartDate, "m/d/yy")
            gPackDate slDate, ilDate0, ilDate1
            tmSdfSrchKey1.iDate(0) = ilDate0
            tmSdfSrchKey1.iDate(1) = ilDate1
            tmSdfSrchKey1.iTime(0) = 0
            tmSdfSrchKey1.iTime(1) = 0
            tmSdfSrchKey1.sSchStatus = "M"
            ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVehCode)
                ilFound = False
                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                If (llDate > llEndDate) Then
                    Exit Do
                End If
                If (tmSdf.sSchStatus = "M") And (llDate >= llStartDate) Then
                    gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                    llSpotTime = CLng(gTimeToCurrency(slTime, False))
                    If (llSpotTime >= llStartTime) And (llSpotTime <= llEndTime) Then
                        ''ilRet = btrGetPosition(hmSdf, lgReschRecPos(UBound(lgReschRecPos)))
                        'Test handles in gReschSpots
                        'If tmSdf.sTracer <> "*" Then
                            lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                            ReDim Preserve lgReschSdfCode(1 To UBound(lgReschSdfCode) + 1) As Long
                        'Else
                        '    If tmSdf.sSpotType = "X" Then   'Bonus spot
                        '    Else
                        '    End If
                        'End If
                    End If
                End If
                ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Loop
            'sgPreemptPass = smPreemptPass
            ilRet = gReSchSpots(True, 0, "YYYYYYY", 0, 86400)
            'sgPreemptPass = "N"
        End If
        llPercent = ((CLng(ilVeh) + 1) * CSng(100)) / UBound(tgMVef)
        If llPercent >= 100 Then
            llPercent = 100
        End If
        plcGauge.Value = llPercent
    Next ilVeh
    lacCntr(0).Left = 1875
    lacCntr(0).Width = 1920
    lacCntr(0).Caption = "DONE"
    mSchdUnschd = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCurSeqNo                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set the Contract Import        *
'*                      sequence number                *
'*                                                     *
'*******************************************************
Private Function mSetCurSeqNo(ilSeqNo As Integer, llMoveID As Long) As Integer
    Dim ilRet As Integer
    mSetCurSeqNo = False
    'hmSaf = CBtrvTable(TWOHANDLES)
    'ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'If ilRet = BTRV_ERR_NONE Then
        imSafRecLen = Len(tmSaf) 'btrRecordLength(hmSaf)  'Get and save record length
        ilRet = btrGetFirst(hmSaf, tmSaf, imSafRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            tmSaf.iCurSeqNo = ilSeqNo
            tmSaf.lLastMoveID = llMoveID
            ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
            If ilRet = BTRV_ERR_NONE Then
                mSetCurSeqNo = True
            End If
        End If
    'End If
    'ilRet = btrClose(hmSaf)
    'btrDestroy hmSaf
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetHdDate                      *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set header date range           *
'*                                                     *
'*            Code taken from Contract.Frm             *
'*                                                     *
'*******************************************************
Private Sub mSetHdDate()
    Dim llStartDate As Long
    Dim slStartDate As String
    Dim llEndDate As Long
    Dim slEndDate As String
    Dim ilClf As Integer
    Dim ilCff As Integer
    llStartDate = 0
    llEndDate = 0
    For ilClf = LBound(tgClfImpt) To UBound(tgClfImpt) - 1 Step 1
        If ((tgClfImpt(ilClf).iStatus = 0) Or (tgClfImpt(ilClf).iStatus = 1)) And (Not tgClfImpt(ilClf).iCancel) Then
            ilCff = tgClfImpt(ilClf).iFirstCff
            Do While ilCff <> -1
                If (tgCffImpt(ilCff).iStatus = 0) Or (tgCffImpt(ilCff).iStatus = 1) Then
                    gUnpackDate tgCffImpt(ilCff).CffRec.iStartDate(0), tgCffImpt(ilCff).CffRec.iStartDate(1), slStartDate    'Week Start date
                    gUnpackDate tgCffImpt(ilCff).CffRec.iEndDate(0), tgCffImpt(ilCff).CffRec.iEndDate(1), slEndDate    'Week Start date
                    If llStartDate = 0 Then
                        llStartDate = gDateValue(slStartDate)
                        llEndDate = gDateValue(slEndDate)
                    Else
                        If gDateValue(slStartDate) < llStartDate Then
                            llStartDate = gDateValue(slStartDate)
                        End If
                        If gDateValue(slEndDate) > llEndDate Then
                            llEndDate = gDateValue(slEndDate)
                        End If
                    End If
                End If
                ilCff = tgCffImpt(ilCff).iNextCff
            Loop
        End If
    Next ilClf
    slStartDate = Format$(llStartDate, "m/d/yy")
    slEndDate = Format$(llEndDate, "m/d/yy")
    gPackDate slStartDate, tgChfImpt.iStartDate(0), tgChfImpt.iStartDate(1)
    gPackDate slEndDate, tgChfImpt.iEndDate(0), tgChfImpt.iEndDate(1)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetHdVeh                       *
'*                                                     *
'*             Created:8/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine Vehicle for header   *
'*                                                     *
'*          1/18/99 Fix links when more than 50 veh-   *
'*          previously not linking to next 50 properly *
'*******************************************************
Private Function mSetHdVeh()
    ReDim tlNewVsf(0 To 1) As VSF    'Combo Vehicle
    Dim ilNoFac As Integer
    Dim ilOldCount As Integer
    ReDim tlOldVsf(0 To 1) As VSF
    Dim ilRecLen As Integer     'Vsf record length
    Dim tlVsfSrchKey As LONGKEY0
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilFound As Integer
    Dim ilMatch As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilClf As Integer
    Dim ilVsf As Integer
    Dim ilOldVsf As Integer
    Dim ilCount As Integer
    Dim llLkVsfCode As Long
    ilNoFac = 0
    ilCount = 0
    ilIndex = LBound(tlNewVsf)
    ilRecLen = Len(tlOldVsf(0))  'btrRecordLength(hlVpf)  'Get and save record length
    For ilClf = LBound(tgClfImpt) To UBound(tgClfImpt) - 1 Step 1
        'Ignore package vehicles
        If (tgClfImpt(ilClf).iStatus >= 0) And (tgClfImpt(ilClf).ClfRec.sType <> "O") And (tgClfImpt(ilClf).ClfRec.sType <> "A") And (tgClfImpt(ilClf).ClfRec.sType <> "E") Then
            ilFound = False
            For ilVsf = LBound(tlNewVsf) To UBound(tlNewVsf) - 1 Step 1
                For ilLoop = LBound(tlNewVsf(0).iFSCode) To UBound(tlNewVsf(0).iFSCode) Step 1
                    If tlNewVsf(ilVsf).iFSCode(ilLoop) > 0 Then
                        If tlNewVsf(ilVsf).iFSCode(ilLoop) = tgClfImpt(ilClf).ClfRec.iVefCode Then
                            ilFound = True
                        End If
                    Else
                        Exit For
                    End If
                Next ilLoop
            Next ilVsf
            If Not ilFound Then
                If ilNoFac > UBound(tlNewVsf(0).iFSCode) Then
                    ilNoFac = 0
                    ilIndex = ilIndex + 1
                    ReDim Preserve tlNewVsf(0 To UBound(tlNewVsf) + 1) As VSF
                End If
                tlNewVsf(ilIndex).iFSCode(ilNoFac) = tgClfImpt(ilClf).ClfRec.iVefCode
                ilNoFac = ilNoFac + 1
                ilCount = ilCount + 1
            End If
        End If
    Next ilClf
    If ilCount = 0 Then
        If UBound(tgClfImpt) <= 1 Then
            slStr = "Line definition not complete or no lines defined " & smNameMsg & " " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo))
        Else
            slStr = "No Vehicle defined for the lines in " & smNameMsg & " " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " Contract not added"
        End If
        mAddMsg slStr
        Print #hmMsg, slStr
        mSetHdVeh = False
        Exit Function
    ElseIf ilCount = 1 Then
        tgChfImpt.lVefCode = tlNewVsf(0).iFSCode(0)
    Else
        'Short title not used in this import
        'For ilVsf = LBound(tlNewVsf) To UBound(tlNewVsf) - 1 Step 1
        '    For ilVef = LBound(tlNewVsf(0).iFSCode) To UBound(tlNewVsf(0).iFSCode) Step 1
        '        If tlNewVsf(ilVsf).iFSCode(ilVef) > 0 Then
        '            For ilLoop = LBound(tgShtTitle) To UBound(tgShtTitle) - 1 Step 1
        '                If tgShtTitle(ilLoop).iVefCode = tlNewVsf(ilVsf).iFSCode(ilVef) Then
        '                    tlNewVsf(ilVsf).lFSComm(ilVef) = tgShtTitle(ilLoop).lSifCode
        '                    Exit For
        '                End If
        '            Next ilLoop
        '        End If
        '    Next ilVef
        'Next ilVsf
        'If igUserByVeh Then
            'Test if any combo exist that uses the same Vehicles
            ilFound = False
            If tgChfImpt.lVefCode < 0 Then
                tlVsfSrchKey.lCode = -tgChfImpt.lVefCode
                ilRet = btrGetEqual(hmVsf, tlOldVsf(0), ilRecLen, tlVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    llLkVsfCode = tlOldVsf(0).lLkVsfCode
                    Do While llLkVsfCode > 0
                        tlVsfSrchKey.lCode = llLkVsfCode
                        ilRet = btrGetEqual(hmVsf, tlOldVsf(UBound(tlOldVsf)), ilRecLen, tlVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            Exit Do
                        End If
                        llLkVsfCode = tlOldVsf(UBound(tlOldVsf)).lLkVsfCode
                        ReDim Preserve tlOldVsf(0 To UBound(tlOldVsf) + 1) As VSF
                    Loop
                    ilFound = False
                    If tlOldVsf(0).sType = "F" Then
                        ilOldCount = 0
                        For ilVsf = LBound(tlOldVsf) To UBound(tlOldVsf) - 1 Step 1
                            For ilLoop = LBound(tlOldVsf(0).iFSCode) To UBound(tlOldVsf(0).iFSCode) Step 1
                                If tlOldVsf(ilVsf).iFSCode(ilLoop) > 0 Then
                                    ilOldCount = ilOldCount + 1
                                Else
                                    Exit For
                                End If
                            Next ilLoop
                        Next ilVsf
                        If ilOldCount = ilCount Then
                            For ilVsf = LBound(tlNewVsf) To UBound(tlNewVsf) - 1 Step 1
                                For ilLoop = LBound(tlNewVsf(0).iFSCode) To UBound(tlNewVsf(0).iFSCode) Step 1
                                    If tlNewVsf(ilVsf).iFSCode(ilLoop) > 0 Then
                                        ilMatch = False
                                        For ilOldVsf = LBound(tlOldVsf) To UBound(tlOldVsf) - 1 Step 1
                                            For ilIndex = LBound(tlOldVsf(0).iFSCode) To UBound(tlOldVsf(0).iFSCode) Step 1
                                                If (tlNewVsf(ilVsf).iFSCode(ilLoop) = tlOldVsf(ilOldVsf).iFSCode(ilIndex)) And (tlNewVsf(ilVsf).lFSComm(ilLoop) = tlOldVsf(ilOldVsf).lFSComm(ilIndex)) Then
                                                    ilMatch = True
                                                    Exit For
                                                End If
                                            Next ilIndex
                                            If ilMatch Then
                                                Exit For
                                            End If
                                        Next ilOldVsf
                                        If Not ilMatch Then
                                            Exit For
                                        End If
                                    Else
                                        ilMatch = True
                                    End If
                                Next ilLoop
                                If Not ilMatch Then
                                    Exit For
                                End If
                            Next ilVsf
                        Else
                            ilMatch = False
                        End If
                        If ilMatch Then
                            ilFound = True
                            'Exit Do
                        End If
                    End If
                End If
            End If
            If ilFound Then
                tgChfImpt.lVefCode = -tlOldVsf(0).lCode
            Else
                llLkVsfCode = 0
                For ilVsf = UBound(tlNewVsf) - 1 To 0 Step -1
                    tlNewVsf(ilVsf).lCode = 0
                    tlNewVsf(ilVsf).sType = "F"
                    tlNewVsf(ilVsf).sName = ""
                    'tlNewVsf.sMktName = ""
                    tlNewVsf(ilVsf).lLkVsfCode = llLkVsfCode
                    For ilLoop = LBound(tlNewVsf(0).lFSComm) To UBound(tlNewVsf(0).lFSComm) Step 1
                        tlNewVsf(ilVsf).iNoSpots(ilLoop) = 0
                    Next ilLoop
                    tlNewVsf(ilVsf).sSource = "S"
                    tlNewVsf(ilVsf).iMerge = 0
                    ilRet = btrInsert(hmVsf, tlNewVsf(ilVsf), ilRecLen, INDEXKEY0)
                    'If mBtrvErrorMsg(ilRet, "mSetHdVeh (btrGetEqual)" & Trim$(Str$(tgChfImpt.lCntrNo)), False) Then
                    '    mSetHdVeh = False
                    '    Exit Function
                    'End If
                    If ilRet <> BTRV_ERR_NONE Then
                        If ilRet >= 30000 Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        slStr = "mSetHdVeh btrInsert error for " & Trim$(str$(tgChfImpt.lCntrNo)) & " Rev " & Trim$(str$(tgChfImpt.iExtRevNo)) & " Error" & str$(ilRet) & " Contract not added"
                        'lbcErrors.AddItem slStr
                        mAddMsg slStr
                        Print #hmMsg, slStr
                        mSetHdVeh = False
                        Exit Function
                    End If
                    llLkVsfCode = tlNewVsf(ilVsf).lCode    '.lLkVsfCode chged to .lcode
                Next ilVsf
                tgChfImpt.lVefCode = -tlNewVsf(0).lCode
            End If
        'Else
        '    tgChfImpt.iVefCode = 0
        'End If
    End If
    mSetHdVeh = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetLnRdf                       *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set line RdfCode                *
'*                                                     *
'*******************************************************
Private Function mSetLnRdf(tlClf As CLFLIST)
    Dim slStartTime As String
    Dim slEndTime As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llMStartTime As Long
    Dim llMEndTime As Long
    ReDim ilDay(0 To 6) As Integer 'Valid days
    ReDim ilMDay(0 To 6) As Integer 'Valid days
    Dim ilCff As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilMIndex As Integer
    Dim ilOk As Integer
    Dim ilDayIndex As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilRdfIndex As Integer   'Rdf with smallest subset of times and days OK
    Dim llSubSetExtraTime As Long
    Dim llRif As Long
    Dim ilAnfCode As Integer

    gUnpackTime tlClf.ClfRec.iStartTime(0), tlClf.ClfRec.iStartTime(1), "A", "1", slStartTime
    llStartTime = gTimeToLong(slStartTime, False)
    gUnpackTime tlClf.ClfRec.iEndTime(0), tlClf.ClfRec.iEndTime(1), "A", "1", slEndTime
    llEndTime = gTimeToLong(slEndTime, True)
    If llEndTime < llStartTime Then
        llMEndTime = llEndTime
        llMStartTime = gTimeToLong("12M", False)
        llEndTime = gTimeToLong("12M", True)
    Else
        llMStartTime = 0
        llMEndTime = 0
    End If
    For ilLoop = 0 To 6 Step 1
        ilDay(ilLoop) = 0
        ilMDay(ilLoop) = 0
    Next ilLoop
    ilRdfIndex = -1
    ilAnfCode = tlClf.iPriChgd
    llStartDate = 1000000
    llEndDate = 0
    ilCff = tlClf.iFirstCff
    Do While ilCff <> -1
        gUnpackDate tgCffImpt(ilCff).CffRec.iStartDate(0), tgCffImpt(ilCff).CffRec.iStartDate(1), slStartDate
        If gDateValue(slStartDate) < llStartDate Then
            llStartDate = gDateValue(slStartDate)
        End If
        gUnpackDate tgCffImpt(ilCff).CffRec.iEndDate(0), tgCffImpt(ilCff).CffRec.iEndDate(1), slEndDate
        If gDateValue(slEndDate) > llEndDate Then
            llEndDate = gDateValue(slEndDate)
        End If
        For ilLoop = 0 To 6 Step 1
            If tgCffImpt(ilCff).CffRec.iDay(ilLoop) > 0 Then
                ilDay(ilLoop) = 1
            End If
        Next ilLoop
        ilCff = tgCffImpt(ilCff).iNextCff
    Loop
    If llMEndTime <> 0 Then
        For ilLoop = 0 To 6 Step 1
            If ilLoop < 6 Then
                ilMDay(ilLoop + 1) = ilDay(ilLoop)
            Else
                ilMDay(0) = ilDay(ilLoop)
            End If
        Next ilLoop
    End If
    'Find best time which does not valid days
    For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
        If tlClf.ClfRec.iVefCode = tgMRif(llRif).iVefCode Then
            'For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
            '    If tgMRif(ilRif).iRdfcode = tgMRdf(ilLoop).iCode Then
                ilLoop = gBinarySearchRdf(tgMRif(llRif).iRdfcode)
                If ilLoop <> -1 Then
                    'gUnpackDate tgMRif(ilLoop).iStartDate(0), tgMRif(ilLoop).iStartDate(1), slStartDate
                    'gUnpackDate tgMRif(ilLoop).iEndDate(0), tgMRif(ilLoop).iEndDate(1), slEndDate
                    'Ignore date test as the line could have started in the middle and gos
                    'pass the end.
                    'Also, when starting a system, the start date of the rate card will not
                    'be the true start date and line might be prior to the dates
        '            If slEndDate = "" Then
                        slEndDate = "12/31/2069"
        '            End If
                    slStartDate = "1/1/1970"
                    If (llStartDate >= gDateValue(slStartDate)) And (llEndDate <= gDateValue(slEndDate)) Then
                        If (tgMRdf(ilLoop).iLtfCode(0) = 0) And (tgMRdf(ilLoop).iLtfCode(1) = 0) And (tgMRdf(ilLoop).iLtfCode(2) = 0) Then
                            For ilIndex = LBound(tgMRdf(ilLoop).iStartTime, 2) To UBound(tgMRdf(ilLoop).iStartTime, 2) Step 1
                                If (tgMRdf(ilLoop).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilLoop).iStartTime(1, ilIndex) <> 0) Then
                                    gUnpackTime tgMRdf(ilLoop).iStartTime(0, ilIndex), tgMRdf(ilLoop).iStartTime(1, ilIndex), "A", "1", slStartTime
                                    gUnpackTime tgMRdf(ilLoop).iEndTime(0, ilIndex), tgMRdf(ilLoop).iEndTime(1, ilIndex), "A", "1", slEndTime
                                    If (llStartTime = gTimeToLong(slStartTime, False)) And (llEndTime = gTimeToLong(slEndTime, True)) Then
                                        'Exact time match- check days
                                        ilOk = True
                                        For ilDayIndex = 0 To 6 Step 1
                                            If (ilDay(ilDayIndex) > 0) And (tgMRdf(ilLoop).sWkDays(ilIndex, ilDayIndex + 1) <> "Y") Then
                                                ilOk = False
                                                Exit For
                                            End If
                                        Next ilDayIndex
                                        If (llMEndTime <> 0) And (ilOk) Then
                                            'Both segments must match times and days as we can not have overrides across midnight
                                            ilOk = False
                                            For ilMIndex = LBound(tgMRdf(ilLoop).iStartTime, 2) To UBound(tgMRdf(ilLoop).iStartTime, 2) Step 1
                                                If (tgMRdf(ilLoop).iStartTime(0, ilMIndex) <> 1) Or (tgMRdf(ilLoop).iStartTime(1, ilMIndex) <> 0) Then
                                                    gUnpackTime tgMRdf(ilLoop).iStartTime(0, ilMIndex), tgMRdf(ilLoop).iStartTime(1, ilMIndex), "A", "1", slStartTime
                                                    gUnpackTime tgMRdf(ilLoop).iEndTime(0, ilMIndex), tgMRdf(ilLoop).iEndTime(1, ilMIndex), "A", "1", slEndTime
                                                    If (llMStartTime = gTimeToLong(slStartTime, False)) And (llMEndTime = gTimeToLong(slEndTime, True)) Then
                                                        'Exact time match- check days
                                                        ilOk = True
                                                        For ilDayIndex = 0 To 6 Step 1
                                                            If (ilMDay(ilDayIndex) > 0) And (tgMRdf(ilLoop).sWkDays(ilMIndex, ilDayIndex + 1) <> "Y") Then
                                                                ilOk = False
                                                                Exit For
                                                            End If
                                                        Next ilDayIndex
                                                        If ilOk Then
                                                            If ((ilAnfCode = 0) And (tgMRdf(ilLoop).sInOut = "N")) Or ((tgMRdf(ilLoop).sInOut = "I") And (tgMRdf(ilLoop).ianfCode = ilAnfCode)) Then
                                                            Else
                                                                ilOk = False
                                                            End If
                                                        End If
                                                    'ElseIf (llMStartTime >= gTimeToLong(slStartTime, False)) And (llMEndTime <= gTimeToLong(slEndTime, True)) Then
                                                    '    'Subset of times- check days
                                                    '    ilOk = True
                                                    '    For ilDayIndex = 0 To 6 Step 1
                                                    '        If (ilDay(ilDayIndex) > 0) And (tgMRdf(ilLoop).sWkDays(ilIndex, ilDayIndex + 1) <> "Y") Then
                                                    '            ilOk = False
                                                    '            Exit For
                                                    '        End If
                                                    '    Next ilDayIndex
                                                    '    If ilOk Then
                                                    '        If ilRdfIndex = -1 Then
                                                    '            ilRdfIndex = ilLoop
                                                    '            llSubSetExtraTime = gTimeToLong(slEndTime, True) - gTimeToLong(slStartTime, False) - (llMEndTime - llMStartTime)
                                                    '        Else
                                                    '            If gTimeToLong(slEndTime, True) - gTimeToLong(slStartTime, False) - (llMEndTime - llMStartTime) < llSubSetExtraTime Then
                                                    '                ilRdfIndex = ilLoop
                                                    '                llSubSetExtraTime = gTimeToLong(slEndTime, True) - gTimeToLong(slStartTime, False) - (llMEndTime - llMStartTime)
                                                    '            End If
                                                    '        End If
                                                    '        ilOk = False
                                                    '    End If
                                                    End If
                                                End If
                                            Next ilMIndex
                                        ElseIf (llMEndTime = 0) And (ilOk) Then
                                            If ((ilAnfCode = 0) And (tgMRdf(ilLoop).sInOut = "N")) Or ((tgMRdf(ilLoop).sInOut = "I") And (tgMRdf(ilLoop).ianfCode = ilAnfCode)) Then
                                            Else
                                                ilOk = False
                                            End If
                                        End If
                                        If ilOk Then
                                            tlClf.ClfRec.iRdfcode = tgMRdf(ilLoop).iCode
                                            'Remove times as exact times found
                                            tlClf.ClfRec.iStartTime(0) = 1
                                            tlClf.ClfRec.iStartTime(1) = 0
                                            tlClf.ClfRec.iEndTime(0) = 1
                                            tlClf.ClfRec.iEndTime(1) = 0
                                            mSetLnRdf = True
                                            Exit Function
                                        End If
                                    ElseIf (llStartTime >= gTimeToLong(slStartTime, False)) And (llEndTime <= gTimeToLong(slEndTime, True)) And (llMEndTime = 0) Then
                                        'Subset of times- check days
                                        ilOk = True
                                        For ilDayIndex = 0 To 6 Step 1
                                            If (ilDay(ilDayIndex) > 0) And (tgMRdf(ilLoop).sWkDays(ilIndex, ilDayIndex + 1) <> "Y") Then
                                                ilOk = False
                                                Exit For
                                            End If
                                        Next ilDayIndex
                                        If ilOk Then
                                            If ((ilAnfCode = 0) And (tgMRdf(ilLoop).sInOut = "N")) Or ((tgMRdf(ilLoop).sInOut = "I") And (tgMRdf(ilLoop).ianfCode = ilAnfCode)) Then
                                            Else
                                                ilOk = False
                                            End If
                                        End If
                                        If ilOk Then
                                            If ilRdfIndex = -1 Then
                                                ilRdfIndex = ilLoop
                                                llSubSetExtraTime = gTimeToLong(slEndTime, True) - gTimeToLong(slStartTime, False) - (llEndTime - llStartTime)
                                            Else
                                                If gTimeToLong(slEndTime, True) - gTimeToLong(slStartTime, False) - (llEndTime - llStartTime) < llSubSetExtraTime Then
                                                    ilRdfIndex = ilLoop
                                                    llSubSetExtraTime = gTimeToLong(slEndTime, True) - gTimeToLong(slStartTime, False) - (llEndTime - llStartTime)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next ilIndex
                        End If
                    End If
                End If
            'Next ilLoop
        End If
    Next llRif
    If ilRdfIndex = -1 Then
        If (tlClf.ClfRec.sType = "O") Or (tlClf.ClfRec.sType = "A") Or (tlClf.ClfRec.sType = "E") Then
            If imPkgRdfCode <> -1 Then
                tlClf.ClfRec.iRdfcode = imPkgRdfCode
                mSetLnRdf = True
            Else
                mSetLnRdf = False
            End If
        Else
            mSetLnRdf = False
        End If
    Else
        tlClf.ClfRec.iRdfcode = tgMRdf(ilRdfIndex).iCode
        mSetLnRdf = True
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetSeqNo                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set the Contract Import        *
'*                      sequence number                *
'*                                                     *
'*******************************************************
Private Function mSetSeqNo(ilSeqNo As Integer) As Integer
    Dim ilRet As Integer
    mSetSeqNo = False
    'hmSaf = CBtrvTable(TWOHANDLES)
    'ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'If ilRet = BTRV_ERR_NONE Then
        imSafRecLen = Len(tmSaf) 'btrRecordLength(hmSaf)  'Get and save record length
        ilRet = btrGetFirst(hmSaf, tmSaf, imSafRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            tmSaf.iImptSeqNo = ilSeqNo
            ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
            If ilRet = BTRV_ERR_NONE Then
                mSetSeqNo = True
            End If
        End If
    'End If
    'ilRet = btrClose(hmSaf)
    'btrDestroy hmSaf
End Function
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
    Erase tmSortCode
    
    Erase tgDMnf
    'Erase tgCommAdf
    'Erase tgCommAgf
    Erase tgCompMnf
    'Erase tgCSlf
    Erase tgCffImpt
    Erase tgClfImpt
    'Erase tgVef
    Erase tgMRdf
    Erase tgMoveRec
    Erase tmErrMsg
    Erase lmPrevMoveDeleted
    Erase lmPrevSdfCreated
    Erase imMS12MVefCode

    ilRet = btrClose(hmRlf)
    btrDestroy hmRlf
    ilRet = btrClose(hmVpf)
    btrDestroy hmVpf
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    ilRet = btrClose(hmSaf)
    btrDestroy hmSaf
    ilRet = btrClose(hmSlf)
    btrDestroy hmSlf
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmChf)
    btrDestroy hmChf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmRcf)
    btrDestroy hmRcf
    ilRet = btrClose(hmRdf)
    btrDestroy hmRdf
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmMtf)
    btrDestroy hmMtf
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf

    ilRet = gObtainSalesperson()
    ilRet = gObtainComp()
    ilRet = gObtainAvail()
'    ilRet = gObtainAdvt()
'    ilRet = gObtainAgency()
    sgMRcfStamp = ""
    sgMRifStamp = ""
    sgMRdfStamp = ""
    ilRet = gObtainRcfRifRdf()
    'ilRet = btrClose(hmIcf)
    'btrDestroy hmIcf
    Screen.MousePointer = vbDefault
    'igParentRestarted = False
    'If Not igStdAloneMode Then
    '    If StrComp(sgCallAppName, "Traffic", 1) = 0 Then
    '        edcLinkDestHelpMsg.LinkExecute "@" & "Done"
    '    Else
    '        edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
    '        edcLinkDestHelpMsg.LinkTopic = sgCallAppName & "|DoneMsg"
    '        edcLinkDestHelpMsg.LinkItem = "edcLinkSrceDoneMsg"
    '        edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
    '        edcLinkDestHelpMsg.LinkExecute "Done"
    '    End If
    '    Do While Not igParentRestarted
    '        DoEvents
    '    Loop
    'End If
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ImptCntr
    Set ImptCntr = Nothing   'Remove data segment
    igManUnload = NO
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestMSU12M12M                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test that each vehicle         *
'*                      has a Mo-Su 12m-12m daypart    *
'*                                                     *
'*******************************************************
Private Function mTestMSU12M12M() As Integer
    Dim ilVeh As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilIndex As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilOk As Integer
    Dim ilDay As Integer
    Dim slStr As String
    Dim ilLoop As Integer

    mTestMSU12M12M = True
    For ilVeh = LBound(imMS12MVefCode) To UBound(imMS12MVefCode) - 1 Step 1
        ilOk = False
        For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
            If imMS12MVefCode(ilVeh) = tgMRif(llRif).iVefCode Then
                ilOk = False
                'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                '    If (tgMRif(ilRif).iRdfcode = tgMRdf(ilRdf).iCode) Then
                    ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfcode)
                    If ilRdf <> -1 Then
                        For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                            If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), False, llStartTime
                                gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilIndex), tgMRdf(ilRdf).iEndTime(1, ilIndex), False, llEndTime
                                ilOk = True
                                For ilDay = 1 To 7 Step 1
                                    If tgMRdf(ilRdf).sWkDays(ilIndex, ilDay) <> "Y" Then
                                        ilOk = False
                                    End If
                                Next ilDay
                                If (llStartTime = 0) And (llEndTime = 0) And (ilOk = True) And (tgMRdf(ilRdf).sInOut = "N") And (tgMRdf(ilRdf).sState = "A") Then
                                    Exit For
                                Else
                                    ilOk = False
                                End If
                            End If
                        Next ilIndex
                        If ilOk Then
                '            Exit For
                        End If
                    End If
                'Next ilRdf
                If ilOk Then
                    Exit For
                End If
            End If
        Next llRif
        If Not ilOk Then
            For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                If (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S") Then
                    If imMS12MVefCode(ilVeh) = tgMVef(ilLoop).iCode Then
                        slStr = "Vehicle " & Trim$(tgMVef(ilLoop).sName) & " missing Daypart Mo-Su 12m-12m"
                        Print #hmMsg, slStr
                        lbcErrors.AddItem slStr
                        mTestMSU12M12M = False
                        Exit For
                    End If
                End If
            Next ilLoop
        End If
    Next ilVeh
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestRecLengths                 *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if record lengths match    *
'*                                                     *
'*******************************************************
Private Function mTestRecLengths() As Integer
    Dim ilSizeError As Integer
    Dim ilSize As Integer
    ilSizeError = False
    ilSize = mGetRecLength("Adf.Btr")
    If ilSize <> Len(tmAdf) Then
        If ilSize > 0 Then
            MsgBox "Adf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmAdf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Adf error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Agf.Btr")
    If ilSize <> Len(tmAgf) Then
        If ilSize > 0 Then
            MsgBox "Agf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmAgf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Agf error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Slf.Btr")
    If ilSize <> Len(tmSlf) Then
        If ilSize > 0 Then
            MsgBox "Slf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmSlf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Slf error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Mnf.Btr")
    If ilSize <> Len(tmMnf) Then
        If ilSize > 0 Then
            MsgBox "Mnf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmMnf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Mnf error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Prf.Btr")
    If ilSize <> Len(tmPrf) Then
        If ilSize > 0 Then
            MsgBox "Prf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmPrf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Prf error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Cff.Btr")
    If ilSize <> Len(tgCffImpt(1).CffRec) Then
        If ilSize > 0 Then
            MsgBox "Cff size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tgCffImpt(1).CffRec)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Cff error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Clf.Btr")
    If ilSize <> Len(tgClfImpt(1).ClfRec) Then
        If ilSize > 0 Then
            MsgBox "Clf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tgClfImpt(1).ClfRec)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Clf error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Chf.Btr")
    If ilSize <> Len(tgChfImpt) Then
        If ilSize > 0 Then
            MsgBox "Chf size error: Btrieve Size" & str$(mGetRecLength("Chf.Btr")) & " Internal size" & str$(Len(tgChfImpt)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Chf error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Vef.Btr")
    If mGetRecLength("Vef.Btr") <> Len(tmVef) Then
        If ilSize > 0 Then
            MsgBox "Vef size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmVef)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Vef error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Vsf.Btr")
    If ilSize <> Len(tmVsf) Then
        If ilSize > 0 Then
            MsgBox "Vsf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmVsf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Vsf error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Mtf.Btr")
    If ilSize <> Len(tmMtf) Then
        If ilSize > 0 Then
            MsgBox "Mtf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmMtf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Mtf error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    mTestRecLengths = ilSizeError
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mUnschdCntr                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if any Unscheduled   *
'*                      contracts exist                *
'*                                                     *
'*******************************************************
Private Function mUnschdCntr() As Integer
    Dim llNoRec As Long         'Number of records in Sof
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffset As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    llNoRec = gExtNoRec(Len(tgChfImpt)) 'Obtain number of records
    btrExtClear hmChf   'Clear any previous extend operation
    ilRet = btrGetFirst(hmChf, tgChfImpt, imChfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        mUnschdCntr = False
        Exit Function
    End If
    Call btrExtSetBounds(hmChf, llNoRec, -1, "UC", "CHF", "") 'Set extract limits (all records)
    tlCharTypeBuff.sType = "A"  'Old contract which requires scheduling
    ilOffset = gFieldOffset("Chf", "ChfSchStatus")
    ilRet = btrExtAddLogicConst(hmChf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlCharTypeBuff, 1)
    tlCharTypeBuff.sType = "I"  'Old contract which requires scheduling
    ilOffset = gFieldOffset("Chf", "ChfSchStatus")
    ilRet = btrExtAddLogicConst(hmChf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlCharTypeBuff, 1)
    tlCharTypeBuff.sType = "N"  'New contract which requires scheduling
    ilOffset = gFieldOffset("Chf", "ChfSchStatus")
    ilRet = btrExtAddLogicConst(hmChf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
    ilRet = btrExtAddField(hmChf, 0, Len(tgChfImpt))  'Extract iCode field
    ilExtLen = Len(tgChfImpt)  'Extract operation record size
    ilRet = btrExtGetNext(hmChf, tgChfImpt, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        ilExtLen = Len(tgChfImpt)  'Extract operation record size
        'ilRet = btrExtGetFirst(hmChf, tgChfImpt, ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmChf, tgChfImpt, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            mUnschdCntr = True
            Exit Function
            ilRet = btrExtGetNext(hmChf, tgChfImpt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmChf, tgChfImpt, ilExtLen, llRecPos)
            Loop
        Loop
    End If
    mUnschdCntr = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mUpdateSMF                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Update SMF with Track ID       *
'*                                                     *
'*******************************************************
Private Function mUpdateSMF(llTrackID As Long, llEnteredDate As Long, llRefTrackID As Long, llFromPrice As Long, llToPrice As Long, llTransGpID As Long, llPrevTransGpID As Long) As Integer
    Dim ilRet As Integer
    Dim slDate As String
    Dim slTime As String
    Dim slStr As String
    tmSdfSrchKey3.lCode = tmSdf.lCode
    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE Then
        If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
            gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
            imInfoCount = imInfoCount + 1
            If tmSdf.sSpotType = "X" Then
                slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Bonus Spot created on Vehicle " & smMoveValues(11) & " on " & slDate & " at " & slTime
            Else
                If tmMoveRec.sOper = "C" Then
                    slStr = " Combine"
                ElseIf tmMoveRec.sOper = "S" Then
                    slStr = " Split"
                ElseIf tmMoveRec.sOper = "M" Then
                    slStr = " Move"
                ElseIf tmMoveRec.sOper = "U" Then
                    slStr = " Undo Combine"
                ElseIf tmMoveRec.sOper = "V" Then
                    slStr = " Undo Split"
                Else
                    slStr = "?" & tmMoveRec.sOper & "?"
                End If
                slStr = smBlankSpaces & smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " MG Spot on Vehicle " & smMoveValues(11) & " at " & slDate & " " & slTime & " operation" & slStr
            End If
            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
            tmSmfSrchKey2.lCode = tmSdf.lCode
            ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                ilRet = mMakeMTF(tmSdf.lChfCode, tmSdf.iLineNo, tmSdf.iVefCode, llTrackID, llEnteredDate, llRefTrackID, llFromPrice, llToPrice, llTransGpID, tmSdf.sSchStatus, llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
                If ilRet Then
                    tmSmf.lMtfCode = tmMtf.lCode
                    ilRet = btrUpdate(hmSmf, tmSmf, imSmfRecLen)
                Else
                    ilRet = imRet
                End If
            End If
        Else
            'Remove spot but leave MTF so that image of spot can be found
            If tmMoveRec.sOper = "C" Then
                slStr = " Combine"
            ElseIf tmMoveRec.sOper = "S" Then
                slStr = " Split"
            ElseIf tmMoveRec.sOper = "M" Then
                slStr = " Move"
            ElseIf tmMoveRec.sOper = "U" Then
                slStr = " Undo Combine"
            ElseIf tmMoveRec.sOper = "V" Then
                slStr = " Undo Split"
            Else
                slStr = "?" & tmMoveRec.sOper & "?"
            End If
            'MAI/Sirius: Added 10/4 Added smBlankSpaces and Count
            imWarningCount = imWarningCount + 1
            slStr = smBlankSpaces & "Contract " & smMoveValues(5) & " Line " & smMoveValues(6) & " Vehicle " & smMoveValues(11) & " on " & smMoveValues(12) & " can't schedule spot as MG" & " operation" & slStr
            'lbcErrors.AddItem slStr
            'mAddMsg slStr
            Print #hmMsg, slStr & " Rec#" & str$(tmMoveRec.lRecNo)
            ilRet = mMakeMTF(tmSdf.lChfCode, tmSdf.iLineNo, tmSdf.iVefCode, llTrackID, llEnteredDate, llRefTrackID, llFromPrice, llToPrice, llTransGpID, "", llPrevTransGpID, tmMoveRec.lToSTime, tmMoveRec.lToETime, tmMoveRec.iDays())
            ilRet = btrDelete(hmSdf)
        End If
    End If
    mUpdateSMF = ilRet
End Function
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
                edcSDate.Text = Format$(llDate, "m/d/yy")
                edcSDate.SelStart = 0
                edcSDate.SelLength = Len(edcSDate.Text)
                imBypassFocus = True
                edcSDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcSDate.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcChkVeh_Paint()
    pbcChkVeh.CurrentX = (pbcChkVeh.Width - pbcChkVeh.TextWidth("Checking Vehicles, Please Wait....")) / 2
    pbcChkVeh.CurrentY = (pbcChkVeh.Height - pbcChkVeh.TextHeight("Checking Vehicles, Please Wait....")) / 2 - 30
    pbcChkVeh.Print "Checking Vehicles, Please Wait...."
End Sub
Private Sub plChkMove_Paint()
    plChkMove.CurrentX = 0
    plChkMove.CurrentY = 0
    plChkMove.Print "Check Move Vehicles and Avail Names"
End Sub

Private Function mSortNames() As Integer
    Dim ilLoop As Integer
    Dim ilCount As Integer
    Dim slStr As String
    Dim slSeqNo As String
    Dim ilSeqNoWrapAround As Integer
    Dim ilRet As Integer
    Dim slLastSeqNo As String
    
    ReDim tmSortCode(0 To 0) As SORTCODE
    ilCount = 0
    ilSeqNoWrapAround = False
    For ilLoop = 0 To lbcBrowserFile.ListCount - 1 Step 1
        If lbcBrowserFile.Selected(ilLoop) Then
            ilCount = ilCount + 1
            slStr = lbcBrowserFile.List(ilLoop)
            slSeqNo = Mid$(slStr, 6, 3)
            If Val(slSeqNo) = 999 Then
                ilSeqNoWrapAround = True
            End If
        End If
    Next ilLoop
    If ilCount <= 0 Then
        mSortNames = False
        Exit Function
    End If
    For ilLoop = 0 To lbcBrowserFile.ListCount - 1 Step 1
        If lbcBrowserFile.Selected(ilLoop) Then
            slStr = lbcBrowserFile.List(ilLoop)
            slSeqNo = Mid$(slStr, 6, 3)
            If ilSeqNoWrapAround Then
                If Val(slSeqNo) <= 99 Then
                    slSeqNo = "1" & slSeqNo
                Else
                    slSeqNo = "0" & slSeqNo
                End If
            Else
                slSeqNo = "0" & slSeqNo
            End If
            tmSortCode(UBound(tmSortCode)).sKey = slSeqNo & "\" & slStr
            ReDim Preserve tmSortCode(0 To UBound(tmSortCode) + 1) As SORTCODE
        End If
    Next ilLoop
    If UBound(tmSortCode) > 0 Then
        ArraySortTyp fnAV(tmSortCode(), 0), UBound(tmSortCode), 0, LenB(tmSortCode(0)), 0, LenB(tmSortCode(0).sKey), 0 '100, 0
    End If
    'Check seq numbers
    For ilLoop = 0 To UBound(tmSortCode) - 1 Step 1
        slStr = tmSortCode(ilLoop).sKey
        ilRet = gParseItem(slStr, 1, "\", slSeqNo)
        slSeqNo = Mid$(slSeqNo, 2)
        If ilLoop = 0 Then
            If Val(slSeqNo) = 999 Then
                slSeqNo = "-1"
            End If
            slLastSeqNo = slSeqNo
        Else
            If Val(slLastSeqNo) + 1 <> Val(slSeqNo) Then
                mSortNames = False
                Exit Function
            End If
            If Val(slSeqNo) = 999 Then
                slSeqNo = "-1"
            End If
            slLastSeqNo = slSeqNo
        End If
    Next ilLoop
    mSortNames = True
End Function

