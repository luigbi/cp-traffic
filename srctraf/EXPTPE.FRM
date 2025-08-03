VERSION 5.00
Object = "{0E9D0E41-7AB8-11D1-9400-00A0248F2EF0}#1.0#0"; "dzactx.dll"
Begin VB.Form ExptPE 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5220
   ClientLeft      =   1485
   ClientTop       =   2100
   ClientWidth     =   7110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
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
   ScaleHeight     =   5220
   ScaleWidth      =   7110
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5580
      Top             =   2415
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   4065
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   25
      Top             =   1035
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
         TabIndex        =   20
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
         TabIndex        =   28
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
         Picture         =   "EXPTPE.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   26
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
            TabIndex        =   27
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
         TabIndex        =   21
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.TextBox edcSpotDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3315
      MaxLength       =   10
      TabIndex        =   5
      Top             =   720
      Width           =   1110
   End
   Begin VB.CommandButton cmcEndDate 
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
      Left            =   4425
      Picture         =   "EXPTPE.frx":2E1A
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   720
      Width           =   195
   End
   Begin VB.CommandButton cmcStartDate 
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
      Left            =   4425
      Picture         =   "EXPTPE.frx":2F14
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   405
      Width           =   195
   End
   Begin VB.TextBox edcCntrDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3315
      MaxLength       =   10
      TabIndex        =   2
      Top             =   405
      Width           =   1110
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   60
      ScaleHeight     =   270
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   30
      Width           =   5355
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4185
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4605
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "&Generate"
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
      Top             =   3915
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
      Left            =   3930
      TabIndex        =   18
      Top             =   3915
      Width           =   1050
   End
   Begin VB.ListBox lbcPart 
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
      Height          =   1290
      Left            =   1905
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   1335
      Width           =   4935
   End
   Begin VB.PictureBox plcSUseProd 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   5235
      TabIndex        =   7
      Top             =   1035
      Visible         =   0   'False
      Width           =   5235
      Begin VB.OptionButton rbcUsersOff 
         Caption         =   "Yes"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   2445
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   645
      End
      Begin VB.OptionButton rbcUsersOff 
         Caption         =   "No"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3090
         TabIndex        =   9
         Top             =   0
         Width           =   570
      End
   End
   Begin VB.Frame frcExtract 
      Caption         =   "Also Extract: "
      Height          =   660
      Left            =   120
      TabIndex        =   12
      Top             =   2700
      Width           =   6735
      Begin VB.CheckBox ckcContractHistory 
         Caption         =   "Contract Version History"
         Height          =   195
         Left            =   4035
         TabIndex        =   16
         Top             =   300
         Value           =   1  'Checked
         Width           =   2460
      End
      Begin VB.CheckBox ckcResearch 
         Caption         =   "Research"
         Height          =   195
         Left            =   2700
         TabIndex        =   15
         Top             =   300
         Width           =   1230
      End
      Begin VB.CheckBox ckcCollection 
         Caption         =   "Collection"
         Height          =   225
         Left            =   1185
         TabIndex        =   14
         Top             =   300
         Width           =   1200
      End
      Begin VB.CheckBox ckcCopy 
         Caption         =   "Copy"
         Height          =   210
         Left            =   150
         TabIndex        =   13
         Top             =   300
         Width           =   885
      End
   End
   Begin DZACTXLibCtl.dzactxctrl dzactxctrl1 
      Left            =   5790
      OleObjectBlob   =   "EXPTPE.frx":300E
      Top             =   4530
   End
   Begin VB.Label lacParticipant 
      Caption         =   "Extract Participant:"
      Height          =   270
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1710
   End
   Begin VB.Label lacToZip 
      Appearance      =   0  'Flat
      Caption         =   "Data To:"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   31
      Top             =   4755
      Width           =   6735
   End
   Begin VB.Label lacToDB 
      Appearance      =   0  'Flat
      Caption         =   "Data To:"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   19
      Top             =   4515
      Width           =   6810
   End
   Begin VB.Label lacFromDB 
      Appearance      =   0  'Flat
      Caption         =   "Data From:"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   30
      Top             =   4290
      Width           =   6825
   End
   Begin VB.Label lacEndDate 
      Appearance      =   0  'Flat
      Caption         =   "Extract Spots Airing On or After"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2925
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   3855
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacStartDate 
      Appearance      =   0  'Flat
      Caption         =   "Extract Contracts Airing On or After"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   405
      Width           =   3285
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   150
      TabIndex        =   29
      Top             =   3600
      Width           =   6690
   End
End
Attribute VB_Name = "ExptPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software®. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExptPE.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Database View input screen code
Option Explicit
Option Compare Text
Dim hmMsg As Integer   'From file hanle
Dim lmNowDate As Long   'Todays date
Dim lmPrevStdStartDate As Long '30 days prior to now date
Dim hmFrom As Integer
Dim hmTo As Integer
Dim tmRec As BASEREC
Dim smZipName As String
Dim imExit As Integer
Dim imRun As Integer
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imGenerating As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imRecLen As Integer
Dim imSvRecLen As Integer
Dim imShowHelpMsg As Integer    'True=Show help message; False=Ignore help message system
Dim imIncludeCollection As Integer
Dim imIncludeCopy As Integer
Dim imIncludeResearch As Integer
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim imDateBox As Integer    '1=Start Date; 2=Export Date
Dim tmAdf As ADF
Dim tmAgf As AGF
Dim tmSlf As SLF
Dim tmPrf As PRF
Dim tmSif As SIF
Dim tmSaf As SAF
Dim tmBsf As BSF
Dim tmBvf As BVF
Dim tmChf As CHF
Dim tmClf As CLF
Dim tmCff As CFF
Dim tmCgf As CGF
Dim tmMsf As MSF
Dim tmVsf As VSF
Dim tmCxf As CXF
Dim imCxfRecLen As Integer
Dim tmCdf As CDF
Dim tmSbf As SBF
Dim imSbfRecLen As Integer
Dim tmCif As CIF
Dim tmScf As SCF
Dim hmFChf As Integer
Dim hmFClf As Integer
Dim hmFCff As Integer
Dim hmFCgf As Integer
Dim hmFMsf As Integer
Dim hmFVsf As Integer
Dim hmFCxf As Integer
Dim hmFSbf As Integer
Dim hmFTzf As Integer
Dim hmFCnf As Integer
Dim hmFCsf As Integer
Dim hmFCpf As Integer
Dim hmFScf As Integer
Dim hmFStf As Integer
Dim hmFDpf As Integer
Dim hmTChf As Integer
Dim hmTClf As Integer
Dim hmTCff As Integer
Dim hmTCgf As Integer
Dim hmTMsf As Integer
Dim hmTVsf As Integer
Dim hmTCxf As Integer
Dim hmTSbf As Integer
Dim hmTTzf As Integer
Dim hmTCnf As Integer
Dim hmTCsf As Integer
Dim hmTCpf As Integer
Dim hmTScf As Integer
Dim hmTStf As Integer
Dim hmTDpf As Integer
Dim tmVcf As VCF
Dim tmVlf As VLF
Dim tmSdf As SDF
Dim tmSmf As SMF
Dim hmFSmf As Integer
Dim hmTSmf As Integer
Dim hmFSdf As Integer
Dim hmTSdf As Integer
Dim tmDlf As DLF
Dim tmCrf As CRF
Dim tmCnf As CNF
Dim tmCsf As CSF
Dim tmCpf As CPF
Dim tmVef As VEF
Dim tmRif As RIF
Dim tmLcf As LCF
Dim tmTzf As TZF
Dim tmDrf As DRF
Dim tmDpf As DPF
Dim tmMtf As MTF
Dim tmLtf As LTF
Dim tmLvf As LVF
Dim tmLef As LEF
Dim tmRvf As RVF
Dim tmPhf As RVF
Dim tmVof As VOF
Dim tmVpf As VPF
Dim tmPvf As PVF
Dim tmPjf As PJF
Dim tmStf As STF
Dim tmUrf As URF
Dim tmIhf As IHF
Dim tmIsf As ISF
Dim tmGhf As GHF
Dim tmGsf As GSF
Dim tmIntSrchKey As INTKEY0    'Key record image
Dim tmLongSrchKey As LONGKEY0    'Key record image
'Used by mSetCopyPurgeStatus
Dim hmTCif As Integer
Dim hmTMcf As Integer
Dim tmMcf As MCF
Dim imRemoveMcf() As Integer
Dim imIncludeTaxTable As Integer
' MsgBox parameters
'Const vbOkOnly = 0                 ' OK button only
'Const vbCritical = 16          ' Critical message
'Const vbApplicationModal = 0
Dim smToDBPath As String
Dim smZipPath As String
Dim smPartExtName As String 'Producer Extract Name
Dim imVBErrorNo As Integer
Dim lmCount As Long
Dim imCntrUpper0 As Integer
Dim imCntrUpper1 As Integer

Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    If imDateBox = 0 Then
        edcCntrDate.SelStart = 0
        edcCntrDate.SelLength = Len(edcCntrDate.Text)
        edcCntrDate.SetFocus
    Else
        edcSpotDate.SelStart = 0
        edcSpotDate.SelLength = Len(edcSpotDate.Text)
        edcSpotDate.SetFocus
    End If
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    If imDateBox = 0 Then
        edcCntrDate.SelStart = 0
        edcCntrDate.SelLength = Len(edcCntrDate.Text)
        edcCntrDate.SetFocus
    Else
        edcSpotDate.SelStart = 0
        edcSpotDate.SelLength = Len(edcSpotDate.Text)
        edcSpotDate.SetFocus
    End If
End Sub
Private Sub cmcCancel_Click()
    If imGenerating Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcEndDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcSpotDate.SelStart = 0
    edcSpotDate.SelLength = Len(edcSpotDate.Text)
    edcSpotDate.SetFocus
End Sub
Private Sub cmcEndDate_GotFocus()
    If imDateBox <> 1 Then
        plcCalendar.Visible = False
    End If
    imDateBox = 1
    plcCalendar.Move edcSpotDate.Left, edcSpotDate.Top + edcSpotDate.Height
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcGen_Click()
    Dim slToFile
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilPart As Integer
    Dim slCntrDate As String
    Dim slSpotDate As String
    Dim slStr As String
    If imGenerating Then
        Exit Sub
    End If
    If Not igPasswordOk Then
        Exit Sub
    End If
    'If Trim$(smPartExtName) = "" Then
    If lbcPart.ListIndex < 0 Then
        MsgBox "Participant Extraction Name not Selected"
        Beep
        cmcCancel.SetFocus
        Exit Sub
    Else
        smPartExtName = lbcPart.List(lbcPart.ListIndex)
    End If
    igPECode = -1
    For ilLoop = LBound(tgPEMnf) To UBound(tgPEMnf) - 1 Step 1
        If StrComp(Trim$(tgPEMnf(ilLoop).sName), smPartExtName, 1) = 0 Then
            igPECode = tgPEMnf(ilLoop).iCode
            Exit For
        End If
    Next ilLoop
    If igPECode = -1 Then
        MsgBox "Participant Extraction Name not Found"
        Beep
        cmcCancel.SetFocus
        Exit Sub
    End If
    slStr = edcCntrDate.Text
    If Not gValidDate(slStr) Then
        Beep
        edcCntrDate.SetFocus
        Exit Sub
    End If
    slCntrDate = slStr
    slStr = edcSpotDate.Text
    If Not gValidDate(slStr) Then
        Beep
        edcSpotDate.SetFocus
        Exit Sub
    End If
    slSpotDate = slStr
    If (rbcUsersOff(0).Value = False) And (rbcUsersOff(1).Value = False) Then
        MsgBox "All Users Off System Not Specified"
        Beep
        cmcCancel.SetFocus
        Exit Sub
    End If
    If gDateValue(slSpotDate) > lmPrevStdStartDate Then
        lmPrevStdStartDate = gDateValue(slSpotDate)
    End If
    If ckcCollection.Value = vbChecked Then
        imIncludeCollection = True
    Else
        imIncludeCollection = False
    End If
    If ckcCopy.Value = vbChecked Then
        imIncludeCopy = True
    Else
        imIncludeCopy = False
    End If
    If ckcResearch.Value = vbChecked Then
        imIncludeResearch = True
    Else
        imIncludeResearch = False
    End If
    imIncludeTaxTable = False
    If Not mOpenMsgFile() Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    Print #hmMsg, "     Contract Start Extract Date :" & slCntrDate
    Print #hmMsg, "     Spot Start Extract Date :" & slSpotDate
    ilRet = 0
    Screen.MousePointer = vbHourglass
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    imGenerating = True
    ilRet = gObtainVef()
    ReDim igPECSRNVefCode(0 To 0) As Integer
    ReDim igPEAVefCode(0 To 0) As Integer
    ReDim igPEPVefCode(0 To 0) As Integer
    ReDim igPELVefCode(0 To 0) As Integer
    ReDim igPELtfCode(0 To 0) As Integer
    ReDim lgPELvfCode(0 To 0) As Long
    ReDim igAddPkVefCode(0 To 0) As Integer
    ReDim igPECntrAdfCode(0 To 0) As Integer
    ReDim igPECntrAgfCode(0 To 0) As Integer
    ReDim igPECntrSlfCode(0 To 0) As Integer
    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S") Or (tgMVef(ilLoop).sType = "R") Or (tgMVef(ilLoop).sType = "N") Or (tgMVef(ilLoop).sType = "G") Then
            For ilPart = 1 To 8 Step 1
                If tgMVef(ilLoop).iMnfGroup(ilPart) = igPECode Then
                    igPECSRNVefCode(UBound(igPECSRNVefCode)) = tgMVef(ilLoop).iCode
                    ReDim Preserve igPECSRNVefCode(0 To UBound(igPECSRNVefCode) + 1) As Integer
                    If tgMVef(ilLoop).iTrfCode > 0 Then
                        imIncludeTaxTable = True
                    End If
                    Exit For
                End If
            Next ilPart
        ElseIf tgMVef(ilLoop).sType = "A" Then
            For ilPart = 1 To 8 Step 1
                If tgMVef(ilLoop).iMnfGroup(ilPart) = igPECode Then
                    igPEAVefCode(UBound(igPEAVefCode)) = tgMVef(ilLoop).iCode
                    ReDim Preserve igPEAVefCode(0 To UBound(igPEAVefCode) + 1) As Integer
                    If tgMVef(ilLoop).iTrfCode > 0 Then
                        imIncludeTaxTable = True
                    End If
                    Exit For
                End If
            Next ilPart
        ElseIf tgMVef(ilLoop).sType = "P" Then
            For ilPart = 1 To 8 Step 1
                If tgMVef(ilLoop).iMnfGroup(ilPart) = igPECode Then
                    igPEPVefCode(UBound(igPEPVefCode)) = tgMVef(ilLoop).iCode
                    ReDim Preserve igPEPVefCode(0 To UBound(igPEPVefCode) + 1) As Integer
                    If tgMVef(ilLoop).iTrfCode > 0 Then
                        imIncludeTaxTable = True
                    End If
                    Exit For
                End If
            Next ilPart
        ElseIf tgMVef(ilLoop).sType = "L" Then
            For ilPart = 1 To 8 Step 1
                If tgMVef(ilLoop).iMnfGroup(ilPart) = igPECode Then
                    igPELVefCode(UBound(igPELVefCode)) = tgMVef(ilLoop).iCode
                    ReDim Preserve igPELVefCode(0 To UBound(igPELVefCode) + 1) As Integer
                    If tgMVef(ilLoop).iTrfCode > 0 Then
                        imIncludeTaxTable = True
                    End If
                    Exit For
                End If
            Next ilPart
        End If
    Next ilLoop
    'Get Contract so that reduce advertiser, agencies and salespersons
    If Not mGetChfAdfAgfSlf(slCntrDate) Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Advertiser
    'Use fix so that tax flag set
    'If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Adf.Btr")
    'Else
    '    ilRet = mCopyFile("Adf.Btr")
    'End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Agency
    'Use fix so that tax flag set
    'If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Agf.Btr")
    'Else
    '    ilRet = mCopyFile("Agf.Btr")
    'End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Avail Name
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Anf.Btr")
    Else
        ilRet = mCopyFile("Anf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Lock Address
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Arf.Btr")
    Else
        ilRet = mCopyFile("Arf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'If Not mCopyFile("Bof.Btr") Then
    '    imGenerating = False
    '    Screen.MousePointer = vbDefault
    '    Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    '    Close #hmMsg
    '    Exit Sub
    'End If
    'If Not mCopyFile("Bsf.Btr") Then
    '    imGenerating = False
    '    Screen.MousePointer = vbDefault
    '    Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    '    Close #hmMsg
    '    Exit Sub
    'End If
    'If Not mCopyFile("Bvf.Btr") Then
    '    imGenerating = False
    '    Screen.MousePointer = vbDefault
    '    Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    '    Close #hmMsg
    '    Exit Sub
    'End If
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Bsf.Btr")
    Else
        ilRet = mCopyFile("Bsf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Bvf.Btr")
    Else
        ilRet = mCopyFile("Bvf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Ccf.Btr")
    Else
        ilRet = mCopyFile("Ccf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Comments-Dunning
    If imIncludeCollection Then
        'If rbcUsersOff(1).Value Then
        '    ilRet = mCopyFileVar("Cdf.Btr")
        'Else
        '    ilRet = mCopyFile("Cdf.Btr")
        'End If
        ilRet = mCopyCdfRecs()
        If Not ilRet Then
            imGenerating = False
            Screen.MousePointer = vbDefault
            Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Exit Sub
        End If
    End If
    'Comment Events
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileVar("Cef.Btr")
    Else
        ilRet = mCopyFile("Cef.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Contract(cff; chf; clf; vsf; cxf; sbf)
'    If Not mCopyCntrFiles(slCntrDate) Then
    If Not mCopyCntrFiles() Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    ''If Not mCopyFileFix("Cif.Btr") Then
    'If rbcUsersOff(1).Value Then
    '    ilRet = mCopyFileFix("Cif.Btr")
    'Else
    '    ilRet = mCopyFile("Cif.Btr")
    'End If
    'If Not ilRet Then
    '    imGenerating = False
    '    Screen.MousePointer = vbDefault
    '    Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    '    Close #hmMsg
    '    Exit Sub
    'End If
    'Comments-Boilerplates
    If Not mCopyFileVar("Cmf.Btr") Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    ''If Not mCopyFileFix("Cnf.Btr") Then
    'If rbcUsersOff(1).Value Then
    '    ilRet = mCopyFileFix("Cnf.Btr")
    'Else
    '    ilRet = mCopyFile("Cnf.Btr")
    'End If
    'If Not ilRet Then
    '    imGenerating = False
    '    Screen.MousePointer = vbDefault
    '    Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    '    Close #hmMsg
    '    Exit Sub
    'End If
    ''Corp Calendar
    'If Not mCopyFileFix("Cof.Btr") Then
    '    imGenerating = False
    '    Screen.MousePointer = vbDefault
    '    Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    '    Close #hmMsg
    '    Exit Sub
    'End If
    ''If Not mCopyFileFix("Cpf.Btr") Then
    'If rbcUsersOff(1).Value Then
    '    ilRet = mCopyFileFix("Cpf.Btr")
    'Else
    '    ilRet = mCopyFile("Cpf.Btr")
    'End If
    'If Not ilRet Then
    '    imGenerating = False
    '    Screen.MousePointer = vbDefault
    '    Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    '    Close #hmMsg
    '    Exit Sub
    'End If
    'Copy rotation
    'If Not mCopyCrfRecs(slCntrDate) Then
    If imIncludeCopy Then
        '4/30/07:  Get copy using contract date
        'If Not mCopyCrfRecs(Format$(lmPrevStdStartDate, "m/d/yy")) Then
        If Not mCopyCrfRecs(slCntrDate) Then
            imGenerating = False
            Screen.MousePointer = vbDefault
            Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Exit Sub
        End If
        If Not mCopyCifRecs() Then
            imGenerating = False
            Screen.MousePointer = vbDefault
            Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Exit Sub
        End If
    Else
        ReDim tgCopyUsed(0 To 0) As COPYUSED
    End If
    'Copy Script
    ''If Not mCopyFileVar("Csf.Btr") Then
    'If rbcUsersOff(1).Value Then
    '    ilRet = mCopyFileVar("Csf.Btr")
    'Else
    '    ilRet = mCopyFile("Csf.Btr")
    'End If
    'If Not ilRet Then
    '    imGenerating = False
    '    Screen.MousePointer = vbDefault
    '    Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    '    Close #hmMsg
    '    Exit Sub
    'End If
    'If Not mCopyFileFix("Cyf.Btr") Then
    '    imGenerating = False
    '    Screen.MousePointer = vbDefault
    '    Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    '    Close #hmMsg
    '    Exit Sub
    'End If
    'Delivery Links
    If Not mCopyFileFix("Dlf.Btr") Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Demo Data Files
    'If Not mCopyFileFix("Dnf.Btr") Then
    If imIncludeResearch Then
        If rbcUsersOff(1).Value Then
            ilRet = mCopyFileFix("Dnf.Btr")
        Else
            ilRet = mCopyFile("Dnf.Btr")
        End If
        If Not ilRet Then
            imGenerating = False
            Screen.MousePointer = vbDefault
            Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Exit Sub
        End If
        If rbcUsersOff(1).Value Then
            ilRet = mCopyFileFix("Def.Btr")
        Else
            ilRet = mCopyFile("Def.Btr")
        End If
        If Not ilRet Then
            imGenerating = False
            Screen.MousePointer = vbDefault
            Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Exit Sub
        End If
        'If Not mCopyFileFix("Drf.Btr") Then
        If Not mCopyDrfRecs() Then
            imGenerating = False
            Screen.MousePointer = vbDefault
            Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Exit Sub
        End If
    End If
    'Delete records
    'If Not mCopyFileFix("Dsf.Btr") Then
    '    imGenerating = False
    '    Screen.MousePointer = vbDefault
    '    Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    '    Close #hmMsg
    '    Exit Sub
    'End If
    'Engineering links
    If Not mCopyFileFix("Egf.Btr") Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Event Names
    'If Not mCopyFileFix("Enf.Btr") Then
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Enf.Btr")
    Else
        ilRet = mCopyFile("Enf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Event Types
    'If Not mCopyFileFix("Etf.Btr") Then
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Etf.Btr")
    Else
        ilRet = mCopyFile("Etf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Log Calendar
    If Not mCopyLcfRecs(slSpotDate) Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Library Title
    'If Not mCopyFileFix("Ltf.Btr") Then
    ilRet = mCopyFileFix("Ltf.Btr")
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Library Version
    'If Not mCopyFileFix("Lvf.Btr") Then
    ilRet = mCopyFileFix("Lvf.Btr")
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Library Events
    'If Not mCopyFileFix("Lef.Btr") Then
    ilRet = mCopyFileFix("Lef.Btr")
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Media Code
    'If Not mCopyFileFix("Mcf.Btr") Then
    If imIncludeCopy Then
        If rbcUsersOff(1).Value Then
            ilRet = mCopyFileFix("Mcf.Btr")
        Else
            ilRet = mCopyFile("Mcf.Btr")
        End If
        If Not ilRet Then
            imGenerating = False
            Screen.MousePointer = vbDefault
            Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Exit Sub
        End If
    End If
    'Multi-Name
    'If Not mCopyFileFix("Mnf.Btr") Then
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Mnf.Btr")
    Else
        ilRet = mCopyFile("Mnf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'MG Track (this requires lgChfCode)
    'If Not mCopyFileFix("Mtf.Btr", True) Then
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Mtf.Btr")
    Else
        ilRet = mCopyFile("Mtf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    ''Receivables Files
    'If Not mCopyFileFix("Phf.Btr") Then
    '    imGenerating = False
    '    Screen.MousePointer = vbDefault
    '    Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    '    Close #hmMsg
    '    Exit Sub
    'End If
    ''Projection
    If Not mCopyPjfRecs() Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Personnel
    'If Not mCopyFileFix("Pnf.Btr") Then
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Pnf.Btr")
    Else
        ilRet = mCopyFile("Pnf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Products
    'If Not mCopyFileFix("Prf.Btr") Then
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Prf.Btr")
    Else
        ilRet = mCopyFile("Prf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Package Spots
    If imIncludeCollection Then
        If Not mCopyPsfRecs(slCntrDate) Then
            imGenerating = False
            Screen.MousePointer = vbDefault
            Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Exit Sub
        End If
    End If
    'Packages
    If Not mCopyPvfRecs() Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Rate Card
    'If Not mCopyFileFix("Rcf.Btr") Then
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Rcf.Btr")
    Else
        ilRet = mCopyFile("Rcf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Dayparts
    'If Not mCopyFileFix("Rdf.Btr") Then
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Rdf.Btr")
    Else
        ilRet = mCopyFile("Rdf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Rate Card Items
    If Not mCopyFileFix("Rif.Btr") Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Report Names
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileVar("Rnf.Btr")
    Else
        ilRet = mCopyFile("Rnf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Report Names
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Rtf.Btr")
    Else
        ilRet = mCopyFile("Rtf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'If imIncludeCollection Then
        If Not mCopyFileFix("Phf.Btr") Then
            imGenerating = False
            Screen.MousePointer = vbDefault
            Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Exit Sub
        End If
        If Not mCopyFileFix("Rvf.Btr") Then
            imGenerating = False
            Screen.MousePointer = vbDefault
            Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Exit Sub
        End If
    'End If
    'Salesperson Commission
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Scf.Btr")
    Else
        ilRet = mCopyFile("Scf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Spots- This must be prior to SMF
    If Not mCopySdfRecs(slSpotDate) Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Spots
    If Not mCopySmfRecs(slSpotDate) Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Spot Summary
    If Not mCopySsfRecs(slSpotDate) Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    If imIncludeCopy Then
        mSetCopyPurgeStatus
    End If
    'End If
    'Short Titles
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Sif.Btr")
    Else
        ilRet = mCopyFile("Sif.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Salesperson
    'If Not mCopyFileFix("Slf.Btr") Then
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Slf.Btr")
    Else
        ilRet = mCopyFile("Slf.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Set Names
    If Not mCopyFileVar("Snf.Btr") Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Sale Office
    'If Not mCopyFileFix("Sof.Btr") Then
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Sof.Btr")
    Else
        ilRet = mCopyFile("Sof.Btr")
    End If
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Site Options
    If Not mCopyFileFix("Spf.Btr") Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Schedule attributes
    If rbcUsersOff(1).Value Then
        ilRet = mCopyFileFix("Saf.Btr")
    Else
        ilRet = mCopyFile("Saf.Btr")
    End If
    'Report set name
    If Not mCopyFileFix("Srf.Btr") Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Spot Tracking
    If Not mCopyFileFix("Stf.Btr") Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'User
    If Not mCopyUrfRecs() Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Time zone- part of sdf
    'If Not mCopyFileFix("Tzf.Btr", False) Then
    '    imGenerating = False
    '    Screen.MousePointer = vbDefault
    '    Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    '    Close #hmMsg
    '    Exit Sub
    'End If
    ''User
    'If Not mCopyFileFix("Urf.Btr") Then
    '    imGenerating = False
    '    Screen.MousePointer = vbDefault
    '    Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    '    Close #hmMsg
    '    Exit Sub
    'End If
    'Vehicle Conflicts
    If Not mCopyVcfRecs(slSpotDate) Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Vehicle
    If Not mCopyFileFix("Vef.Btr") Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Vehicle Links
    If Not mCopyVlfRecs(slSpotDate) Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Vehicle Log info
    ilRet = mCopyFileFix("Vof.Btr")
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    'Vehicle Options
    'If Not mCopyFileFix("Vpf.Btr") Then
    ilRet = mCopyFileFix("Vpf.Btr")
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    ''Vehicle Combos
    'If Not mCopyFileFix("Vsf.Btr") Then
    '    imGenerating = False
    '    Screen.MousePointer = vbDefault
    '    Print #hmMsg, "** Terminated Generation of Database: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    '    Close #hmMsg
    '    Exit Sub
    'End If
    'Clone Game definition stuff
    ilRet = mCopyFileFix("Ihf.Btr")
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    ilRet = mCopyFileFix("Isf.Btr")
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    ilRet = mCopyFileFix("IIf.Btr")
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    ilRet = mCopyFileFix("Itf.Btr")
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    ilRet = mCopyFileFix("Ghf.Btr")
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If
    ilRet = mCopyFileFix("Gsf.Btr")
    If Not ilRet Then
        imGenerating = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Exit Sub
    End If

    If imIncludeTaxTable Then
        ilRet = mCopyFileFix("Trf.Btr")
        If Not ilRet Then
            imGenerating = False
            Screen.MousePointer = vbDefault
            Print #hmMsg, "** Terminated Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Exit Sub
        End If
    End If
    
    lacProcessing.Caption = "Making Result File: " & smZipName
    DoEvents
    ilRet = mZipFiles()
    Print #hmMsg, "** Completed Generation of Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    Close #hmMsg
    On Error GoTo 0
    
    lacProcessing.Caption = "Extraction Completed"
    DoEvents
    Screen.MousePointer = vbDefault
    imGenerating = False
    If imExit Then
        mTerminate
    End If
    cmcCancel.Caption = "&Done"
    cmcCancel.SetFocus
    cmcGen.Enabled = False
    Exit Sub
cmcExportErr:
    ilRet = Err
    Resume Next
End Sub
Private Sub cmcGen_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcStartDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcCntrDate.SelStart = 0
    edcCntrDate.SelLength = Len(edcCntrDate.Text)
    edcCntrDate.SetFocus
End Sub
Private Sub cmcStartDate_GotFocus()
    If imDateBox <> 0 Then
        plcCalendar.Visible = False
    End If
    imDateBox = 0
    plcCalendar.Move edcCntrDate.Left, edcCntrDate.Top + edcCntrDate.Height
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcCntrDate_Change()
    Dim slStr As String
    slStr = edcCntrDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcCntrDate_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    If imDateBox <> 0 Then
        plcCalendar.Visible = False
    End If
    imDateBox = 0
    plcCalendar.Move edcCntrDate.Left, edcCntrDate.Top + edcCntrDate.Height
    gCtrlGotFocus edcCntrDate
End Sub
Private Sub edcCntrDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcCntrDate_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcCntrDate.SelLength <> 0 Then    'avoid deleting two characters
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
Private Sub edcCntrDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcCntrDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcCntrDate.Text = slDate
            End If
        End If
        edcCntrDate.SelStart = 0
        edcCntrDate.SelLength = Len(edcCntrDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcCntrDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcCntrDate.Text = slDate
            End If
        End If
        edcCntrDate.SelStart = 0
        edcCntrDate.SelLength = Len(edcCntrDate.Text)
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcSpotDate_Change()
    Dim slStr As String
    slStr = edcSpotDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcSpotDate_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    If imDateBox <> 1 Then
        plcCalendar.Visible = False
    End If
    imDateBox = 1
    plcCalendar.Move edcSpotDate.Left, edcSpotDate.Top + edcSpotDate.Height
    gCtrlGotFocus edcSpotDate
End Sub
Private Sub edcSpotDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcSpotDate_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcSpotDate.SelLength <> 0 Then    'avoid deleting two characters
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
Private Sub edcSpotDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcSpotDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcSpotDate.Text = slDate
            End If
        End If
        edcSpotDate.SelStart = 0
        edcSpotDate.SelLength = Len(edcSpotDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcSpotDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcSpotDate.Text = slDate
            End If
        End If
        edcSpotDate.SelStart = 0
        edcSpotDate.SelLength = Len(edcSpotDate.Text)
    End If
End Sub
Private Sub Form_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    CSIPass.Show vbModal
    If Not igPasswordOk Then
        cmcGen.Enabled = False
    End If
    mInit
    If imTerminate Then
        cmcCancel_Click
    Else
        If imRun Then
            tmcStart.Enabled = True
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    btrStopAppl
    End
End Sub
Private Sub imcHelp_Click()
    Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcPart_GotFocus()
    plcCalendar.Visible = False
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
    Dim ilCalType As Integer
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    If imDateBox = 0 Then
        slStr = edcCntrDate.Text
    Else
        slStr = edcSpotDate.Text
    End If
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(Str$(Day(llDate)))
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
'*      Procedure Name:mBtrErrorMsg                    *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Display error message           *
'*                                                     *
'*******************************************************
Private Function mBtrErrorMsg(ilRet As Integer, slMess As String) As Integer
    Dim slErrMsg As String
    If ilRet = BTRV_ERR_NONE Then
        mBtrErrorMsg = False
        Exit Function
    End If
    If ilRet >= 30000 Then
        ilRet = csiHandleValue(0, 7)
    End If
    slErrMsg = gGetBtrvErrorMsg(ilRet)
    slErrMsg = slErrMsg & " " & slMess
    Print #hmMsg, slErrMsg
    mBtrErrorMsg = True
End Function
Private Function mCloneCntr(llChfCode As Long) As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilInsert As Integer
    Dim ilTest As Integer
    Dim ilCff As Integer
    Dim ilCgf As Integer
    Dim ilMsf As Integer
    Dim ilSpots As Integer
    Dim llGross As Long
    Dim ilFound As Integer
    Dim tlCffSrchKey As CFFKEY0
    Dim tlClfSrchKey As CLFKEY0
    Dim tlSbfSrchKey As SBFKEY0    'SBF key record image
    Dim tlCgfSrchKey As LONGKEY0
    Dim tlMsfSrchKey As LONGKEY0
    
                
    imRecLen = Len(tgChf)
    ''ilRet = gObtainCntr(hmFChf, hmFClf, hmFCff, llChfCode, False, tgChf, tgClf(), tgCff())
    'If ckcContractHistory.Value = vbChecked Then
    '    ilRet = gObtainCntrPlusGame(hmFChf, hmFClf, hmFCff, hmFCgf, hmFMsf, llChfCode, True, tgChf, tgClf(), tgCff(), tgCgfCntr(), tgMsfCntr())
    'Else
        'ilRet = gObtainCntrPlusGame(hmFChf, hmFClf, hmFCff, hmFCgf, hmFMsf, llChfCode, False, tgChf, tgClf(), tgCff(), tgCgfCntr(), tgMsfCntr())
        ilRet = gObtainCntrPlusGame(hmFChf, hmFClf, hmFCff, hmFCgf, llChfCode, False, tgChf, tgClf(), tgCff(), tgCgfCntr())
    'End If
    tgChf.lInputGross = 0
    tgChf.lCxfMerch = 0
    tgChf.iMerchPct = 0
    tgChf.lCxfProm = 0
    tgChf.iPromoPct = 0
    If tgChf.lCxfCode > 0 Then
        imCxfRecLen = Len(tmCxf)
        tmLongSrchKey.lCode = tgChf.lCxfCode
        'ilRet = btrGetEqual(hmFCxf, tmCxf, imCxfRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        ilRet = gCXFGetEqual(hmFCxf, tmCxf, imCxfRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrInsert(hmTCxf, tmCxf, imCxfRecLen, INDEXKEY0)
        End If
        If ilRet <> BTRV_ERR_NONE Then
            tgChf.lCxfCode = 0
        End If
    End If
    If tgChf.lCxfChgR > 0 Then
        imCxfRecLen = Len(tmCxf)
        tmLongSrchKey.lCode = tgChf.lCxfChgR
        'ilRet = btrGetEqual(hmFCxf, tmCxf, imCxfRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        ilRet = gCXFGetEqual(hmFCxf, tmCxf, imCxfRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrInsert(hmTCxf, tmCxf, imCxfRecLen, INDEXKEY0)
        End If
        If ilRet <> BTRV_ERR_NONE Then
            tgChf.lCxfChgR = 0
        End If
    End If
    If tgChf.lCxfInt > 0 Then
        imCxfRecLen = Len(tmCxf)
        tmLongSrchKey.lCode = tgChf.lCxfInt
        'ilRet = btrGetEqual(hmFCxf, tmCxf, imCxfRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        ilRet = gCXFGetEqual(hmFCxf, tmCxf, imCxfRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrInsert(hmTCxf, tmCxf, imCxfRecLen, INDEXKEY0)
        End If
        If ilRet <> BTRV_ERR_NONE Then
            tgChf.lCxfInt = 0
        End If
    End If
    If tgChf.lCxfCanc > 0 Then
        imCxfRecLen = Len(tmCxf)
        tmLongSrchKey.lCode = tgChf.lCxfCanc
        'ilRet = btrGetEqual(hmFCxf, tmCxf, imCxfRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        ilRet = gCXFGetEqual(hmFCxf, tmCxf, imCxfRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrInsert(hmTCxf, tmCxf, imCxfRecLen, INDEXKEY0)
        End If
        If ilRet <> BTRV_ERR_NONE Then
            tgChf.lCxfCanc = 0
        End If
    End If
    ilRet = mUpdateVsf(tgChf)
    If Not ilRet Then
        mCloneCntr = False
        Exit Function
    End If
    ReDim tgPkageLnCheck(0 To 0) As PKAGELNCHECK
    For ilLoop = LBound(tgClf) To UBound(tgClf) - 1 Step 1
        tmClf = tgClf(ilLoop).ClfRec
        If (tmClf.sType = "O") Or (tmClf.sType = "A") Or (tmClf.sType = "E") Or (tmClf.sType = "H") Then
            ilInsert = False
            If (tmClf.sType = "H") Then
                For ilTest = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                    If (tmClf.iVefCode = igPECSRNVefCode(ilTest)) Then
                        ilInsert = True
                        Exit For
                    End If
                Next ilTest
            Else
                For ilTest = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                    If (tmClf.iVefCode = igPECSRNVefCode(ilTest)) Then
                        ilInsert = True
                        Exit For
                    End If
                Next ilTest
                If Not ilInsert Then
                    For ilTest = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        If (tmClf.iVefCode = tgMVef(ilTest).iCode) And (tgMVef(ilTest).sType = "P") Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilTest
                End If
            End If
            
            ilFound = False
            If tmClf.sType = "H" Then
                For ilTest = 0 To UBound(tgPkageLnCheck) - 1 Step 1
                    If tgPkageLnCheck(ilTest).iPkLineNo = tmClf.iPkLineNo Then
                        If Not ilInsert Then
                            tgPkageLnCheck(ilTest).iInclude = False
                        End If
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
            Else
                For ilTest = 0 To UBound(tgPkageLnCheck) - 1 Step 1
                    If tgPkageLnCheck(ilTest).iPkLineNo = tmClf.iLine Then
                        If Not ilInsert Then
                            tgPkageLnCheck(ilTest).iInclude = False
                        End If
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
            End If
            If Not ilFound Then
                If tmClf.sType = "H" Then
                    tgPkageLnCheck(UBound(tgPkageLnCheck)).iPkLineNo = tmClf.iPkLineNo
                Else
                    tgPkageLnCheck(UBound(tgPkageLnCheck)).iPkLineNo = tmClf.iLine
                End If
                If ilInsert Then
                    tgPkageLnCheck(UBound(tgPkageLnCheck)).iInclude = True
                Else
                    tgPkageLnCheck(UBound(tgPkageLnCheck)).iInclude = True
                End If
                ReDim Preserve tgPkageLnCheck(0 To UBound(tgPkageLnCheck) + 1) As PKAGELNCHECK
            End If
        End If
    Next ilLoop
    llGross = 0
    For ilLoop = LBound(tgClf) To UBound(tgClf) - 1 Step 1
        ilInsert = False
        For ilTest = 0 To UBound(igPECSRNVefCode) - 1 Step 1
            If (tgClf(ilLoop).ClfRec.iVefCode = igPECSRNVefCode(ilTest)) And ((tgClf(ilLoop).ClfRec.sType = "S") Or (tgClf(ilLoop).ClfRec.sType = "H")) Then
                ilInsert = True
                Exit For
            End If
        Next ilTest
        If Not ilInsert Then
            If (tgClf(ilLoop).ClfRec.sType = "O") Or (tgClf(ilLoop).ClfRec.sType = "A") Or (tgClf(ilLoop).ClfRec.sType = "E") Then
                For ilTest = 0 To UBound(tgPkageLnCheck) - 1 Step 1
                    If (tgClf(ilLoop).ClfRec.iLine = tgPkageLnCheck(ilTest).iPkLineNo) Then
                        If tgPkageLnCheck(ilTest).iInclude Then
                            ilInsert = True
                        End If
                        Exit For
                    End If
                Next ilTest
            End If
        End If
        If ilInsert Then
            tlClfSrchKey.lChfCode = tgClf(ilLoop).ClfRec.lChfCode
            tlClfSrchKey.iLine = tgClf(ilLoop).ClfRec.iLine
            tlClfSrchKey.iCntRevNo = tgClf(ilLoop).ClfRec.iCntRevNo
            tlClfSrchKey.iPropVer = tgClf(ilLoop).ClfRec.iPropVer
            ilRet = btrGetEqual(hmTClf, tmClf, Len(tmClf), tlClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilInsert = False
            End If
            If ilInsert Then
                If tgClf(ilLoop).ClfRec.lCxfCode > 0 Then
                    imCxfRecLen = Len(tmCxf)
                    tmLongSrchKey.lCode = tgClf(ilLoop).ClfRec.lCxfCode
                    'ilRet = btrGetEqual(hmFCxf, tmCxf, imCxfRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    ilRet = gCXFGetEqual(hmFCxf, tmCxf, imCxfRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        ilRet = btrInsert(hmTCxf, tmCxf, imCxfRecLen, INDEXKEY0)
                    End If
                    If ilRet <> BTRV_ERR_NONE Then
                        tgClf(ilLoop).ClfRec.lCxfCode = 0
                    End If
                End If
                If tgClf(ilLoop).ClfRec.sType = "H" Then
                    For ilTest = 0 To UBound(tgPkageLnCheck) - 1 Step 1
                        If (tgClf(ilLoop).ClfRec.iPkLineNo = tgPkageLnCheck(ilTest).iPkLineNo) Then
                            If Not tgPkageLnCheck(ilTest).iInclude Then
                                tgClf(ilLoop).ClfRec.sType = "S"    'Remove hidden lines
                            End If
                            Exit For
                        End If
                    Next ilTest
                End If
                If Not imIncludeResearch Then
                    tgClf(ilLoop).ClfRec.iDnfCode = 0
                End If
                ilRet = btrInsert(hmTClf, tgClf(ilLoop).ClfRec, Len(tgClf(ilLoop).ClfRec), INDEXKEY0)
                If mBtrErrorMsg(ilRet, "mCloneCntr" & Str$(tgChf.lCntrNo) & " Line" & Str$(tgClf(ilLoop).ClfRec.iLine)) Then
                    mCloneCntr = False
                    Exit Function
                End If
                If (tgClf(ilLoop).ClfRec.sType = "O") Or (tgClf(ilLoop).ClfRec.sType = "A") Or (tgClf(ilLoop).ClfRec.sType = "E") Then
                    ilFound = False
                    For ilTest = 0 To UBound(igAddPkVefCode) - 1 Step 1
                        If (tgClf(ilLoop).ClfRec.iVefCode = igAddPkVefCode(ilTest)) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilTest
                    If Not ilFound Then
                        igAddPkVefCode(UBound(igAddPkVefCode)) = tgClf(ilLoop).ClfRec.iVefCode
                        ReDim Preserve igAddPkVefCode(0 To UBound(igAddPkVefCode) + 1) As Integer
                    End If
                End If
                For ilCff = LBound(tgCff) To UBound(tgCff) - 1 Step 1
                    If tgClf(ilLoop).ClfRec.iLine = tgCff(ilCff).CffRec.iClfLine Then
                        If (tgClf(ilLoop).ClfRec.sType = "O") Or (tgClf(ilLoop).ClfRec.sType = "A") Or (tgClf(ilLoop).ClfRec.sType = "E") Then
                        Else
                            llGross = llGross + mComputePrice(tgCff(ilCff).CffRec)
                        End If
                        tlCffSrchKey.lChfCode = tgCff(ilCff).CffRec.lChfCode
                        tlCffSrchKey.iClfLine = tgCff(ilCff).CffRec.iClfLine
                        tlCffSrchKey.iCntRevNo = tgCff(ilCff).CffRec.iCntRevNo
                        tlCffSrchKey.iPropVer = tgCff(ilCff).CffRec.iPropVer
                        tlCffSrchKey.iStartDate(0) = tgCff(ilCff).CffRec.iStartDate(0)
                        tlCffSrchKey.iStartDate(1) = tgCff(ilCff).CffRec.iStartDate(1)
                        ilRet = btrGetEqual(hmTCff, tmCff, Len(tmCff), tlCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            ilRet = btrInsert(hmTCff, tgCff(ilCff).CffRec, Len(tgCff(ilCff).CffRec), INDEXKEY0)
                            If mBtrErrorMsg(ilRet, "mCloneCntr" & Str$(tgChf.lCntrNo) & " Line/Flight" & Str$(tgClf(ilLoop).ClfRec.iLine)) Then
                                mCloneCntr = False
                                Exit Function
                            End If
                        End If
                    End If
                Next ilCff
                For ilCgf = LBound(tgCgfCntr) To UBound(tgCgfCntr) - 1 Step 1
                    If tgClf(ilLoop).ClfRec.lCode = tgCgfCntr(ilCgf).CgfRec.lClfCode Then
                        tlCgfSrchKey.lCode = tgCgfCntr(ilCgf).CgfRec.lCode
                        ilRet = btrGetEqual(hmTCgf, tmCgf, Len(tmCgf), tlCgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            ilRet = btrInsert(hmTCgf, tgCgfCntr(ilCgf).CgfRec, Len(tgCgfCntr(ilCgf).CgfRec), INDEXKEY0)
                            If mBtrErrorMsg(ilRet, "mCloneCntr" & Str$(tgChf.lCntrNo) & " Game" & Str$(tgClf(ilLoop).ClfRec.iLine)) Then
                                mCloneCntr = False
                                Exit Function
                            End If
                        End If
                    End If
                Next ilCgf
            End If
        End If
    Next ilLoop

    imSbfRecLen = Len(tmSbf)
    tlSbfSrchKey.lChfCode = llChfCode
    tlSbfSrchKey.iDate(0) = 0
    tlSbfSrchKey.iDate(1) = 0
    tlSbfSrchKey.sTranType = " "
    ilRet = btrGetGreaterOrEqual(hmFSbf, tmSbf, imSbfRecLen, tlSbfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.lChfCode = llChfCode)
        ilInsert = False
        For ilTest = 0 To UBound(igPECSRNVefCode) - 1 Step 1
            If (tmSbf.iAirVefCode = igPECSRNVefCode(ilTest)) Then
                ilInsert = True
                Exit For
            End If
        Next ilTest
        If ilInsert Then
            ilRet = btrInsert(hmTSbf, tmSbf, imSbfRecLen, INDEXKEY0)
            If mBtrErrorMsg(ilRet, "mCloneCntr" & Str$(tgChf.lCntrNo) & " Special Bill") Then
                mCloneCntr = False
                Exit Function
            End If
        End If
        ilRet = btrGetNext(hmFSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    
'    For ilMsf = LBound(tgMsfCntr) To UBound(tgMsfCntr) - 1 Step 1
'        ilInsert = False
'        For ilTest = 0 To UBound(igPECSRNVefCode) - 1 Step 1
'            If (tgMsfCntr(ilMsf).MsfRec.iVefCode = igPECSRNVefCode(ilTest)) Then
'                ilInsert = True
'                Exit For
'            End If
'        Next ilTest
'        If ilInsert Then
'            tlMsfSrchKey.lCode = tgMsfCntr(ilMsf).MsfRec.lCode
'            ilRet = btrGetEqual(hmTMsf, tmMsf, Len(tmMsf), tlMsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'            If ilRet <> BTRV_ERR_NONE Then
'                ilRet = btrInsert(hmTMsf, tgMsfCntr(ilMsf).MsfRec, Len(tgMsfCntr(ilMsf).MsfRec), INDEXKEY0)
'                If mBtrErrorMsg(ilRet, "mCloneCntr" & Str$(tgChf.lCntrNo) & " Multi-Media" & Str$(tgClf(ilLoop).ClfRec.iLine)) Then
'                    mCloneCntr = False
'                    Exit Function
'                End If
'            End If
'        End If
'    Next ilMsf
    
    tgChf.lInputGross = llGross
    ilRet = btrInsert(hmTChf, tgChf, imRecLen, INDEXKEY0)
    If mBtrErrorMsg(ilRet, "mCloneCntr" & Str$(tgChf.lCntrNo)) Then
        mCloneCntr = False
        Exit Function
    End If
    mCloneCntr = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mComputePrice                   *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute Flight price           *
'*                                                     *
'*******************************************************
Private Function mComputePrice(tlCff As CFF) As Long
    Dim llStartDate As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llMoStartDate As Long
    Dim llEndDate As Long
    Dim ilSpots As Integer
    Dim ilDay As Integer
    Dim llLoopDate As Long
    Dim llPrice As Long
    
    llPrice = 0
    gUnpackDate tlCff.iStartDate(0), tlCff.iStartDate(1), slStartDate
    gUnpackDate tlCff.iEndDate(0), tlCff.iEndDate(1), slEndDate
    llStartDate = gDateValue(slStartDate)
    slStartDate = gObtainPrevMonday(slStartDate)
    llMoStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
    For llLoopDate = llMoStartDate To llEndDate Step 7
        If tlCff.sDyWk = "D" Then
            ilSpots = 0
            For ilDay = 0 To 6 Step 1
                If (llLoopDate + ilDay >= llStartDate) And (llLoopDate + ilDay <= llEndDate) Then
                    ilSpots = ilSpots + tlCff.iDay(ilDay)
                End If
            Next ilDay
        Else
            ilSpots = tlCff.iSpotsWk + tlCff.iXSpotsWk
        End If
        If ilSpots > 0 Then
            If tlCff.sPriceType = "T" Then
                llPrice = llPrice + ilSpots * tlCff.lActPrice
            End If
        End If
    Next llLoopDate
    mComputePrice = llPrice
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyFileFix                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Fixed Length Files        *
'*                                                     *
'*******************************************************
'Private Function mCopyCntrFiles(slCntrDate As String) As Integer
Private Function mCopyCntrFiles() As Integer
    Dim ilRet As Integer
'    Dim llNoRecords As Long
'    Dim ilExtLen As Integer
'    Dim ilOffset As Integer
'    Dim llNoRec As Integer
'    Dim ilLoop As Integer
'    Dim ilTest As Integer
'    Dim llRecPos As Long
'    Dim ilInsert As Integer
    Dim ilCff As Integer
'    Dim ilUpper0 As Integer
'    Dim ilUpper1 As Integer
'    Dim ilVsfReclen As Integer
'    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
'    ReDim lgChfCode(1 To 32000, 1 To 1) As Long
    Dim llLoop As Long
'    For ilLoop = 1 To 32000 Step 1
'        lgChfCode(ilLoop, 1) = 0
'    Next ilLoop
'    ilUpper0 = 1
'    ilUpper1 = 1
'    ilUpper0 = imCntrUpper0
'    ilUpper1 = imCntrUpper1
    lacProcessing.Caption = "Processing " & "Contracts"
    DoEvents
'    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
'    ilRet = btrOpen(hmFrom, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open From Chf") Then
'        mCopyCntrFiles = False
'        Exit Function
'    End If
'    imRecLen = btrRecordLength(hmFrom)  'Get and save record length
'    imSvRecLen = imRecLen
'    llNoRecords = btrRecords(hmFrom)
    hmFChf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open 'From' Chf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    hmFClf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open 'From' Clf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    hmFCff = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open 'From' Cff") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    hmFVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open 'From' Vsf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
'    ilVsfReclen = Len(tmVsf)
    hmFCxf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open From Cxf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    hmFSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open From Sbf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    hmFCgf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFCgf, "", sgDBPath & "Cgf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open From Cgf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    hmFMsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFMsf, "", sgDBPath & "Msf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open From Msf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    mRemoveFile "Chf.Btr"
    mRemoveFile "Clf.Btr"
    mRemoveFile "Cff.Btr"
    mRemoveFile "Vsf.Btr"
    mRemoveFile "Cxf.Btr"
    mRemoveFile "Sbf.Btr"
    mRemoveFile "Cgf.Btr"
    mRemoveFile "Msf.Btr"
    ilRet = btrClone(hmFChf, smToDBPath & "Chf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Clone Chf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    ilRet = btrClone(hmFClf, smToDBPath & "Clf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Clone Clf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    ilRet = btrClone(hmFCff, smToDBPath & "Cff.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Clone Cff") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    ilRet = btrClone(hmFVsf, smToDBPath & "Vsf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Clone Vsf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    ilRet = btrClone(hmFCxf, smToDBPath & "Cxf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Clone Cxf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    ilRet = btrClone(hmFSbf, smToDBPath & "Sbf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Clone Sbf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    ilRet = btrClone(hmFCgf, smToDBPath & "Cgf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Clone Cgf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    ilRet = btrClone(hmFMsf, smToDBPath & "Msf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Clone Msf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    
    hmTChf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTChf, "", smToDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open To Chf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    hmTClf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTClf, "", smToDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open 'To' Clf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    hmTCff = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTCff, "", smToDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open 'To' Cff") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    hmTVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTVsf, "", smToDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open 'To' Vsf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    hmTCxf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTCxf, "", smToDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open 'To' Cxf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    hmTSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTSbf, "", smToDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open 'To' Sbf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    hmTCgf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTCgf, "", smToDBPath & "Cgf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open 'To' Cgf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
    hmTMsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTMsf, "", smToDBPath & "Msf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCntrFiles: Open 'To' Msf") Then
        mCopyCntrFiles = False
        Exit Function
    End If
'    ilRet = btrGetFirst(hmFrom, tmChf, imRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'    If (ilRet = BTRV_ERR_NONE) Then
'        ilExtLen = imSvRecLen  'Extract operation record size
'        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlVef) 'Obtain number of records
'        btrExtClear hmFrom   'Clear any previous extend operation
'        Call btrExtSetBounds(hmFrom, llNoRec, -1, "UC", "CHF", "") 'Set extract limits (all records)
'        gPackDate slCntrDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
'        ilOffset = gFieldOffset("Chf", "ChfEndDate")
'        ilRet = btrExtAddLogicConst(hmFrom, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
'        ilOffset = 0
'        ilRet = btrExtAddField(hmFrom, ilOffset, ilExtLen)  'Extract First Name field
'        ilRet = btrExtGetNext(hmFrom, tmChf, ilExtLen, llRecPos)
'        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
'            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
'                If ilRet >= 30000 Then
'                    ilRet = csiHandleValue(0, 7)
'                End If
'                Print #hmMsg, gGetBtrvErrorMsg(ilRet) & "mCopyCntrFile: ExtGetNext"
'                btrDestroy hmFrom
'                btrDestroy hmFChf
'                btrDestroy hmFClf
'                btrDestroy hmFCff
'                btrDestroy hmFVsf
'                btrDestroy hmFCxf
'                btrDestroy hmTChf
'                btrDestroy hmTClf
'                btrDestroy hmTCff
'                btrDestroy hmTVsf
'                btrDestroy hmTCxf
'                mCopyCntrFiles = False
'                Exit Function
'            End If
'            ilExtLen = imSvRecLen
'            Do While ilRet = BTRV_ERR_REJECT_COUNT
'                ilRet = btrExtGetNext(hmFrom, tmChf, ilExtLen, llRecPos)
'            Loop
'            Do While ilRet = BTRV_ERR_NONE
'                DoEvents
'                If imTerminate Then
'                    btrDestroy hmFrom
'                    btrDestroy hmFChf
'                    btrDestroy hmFClf
'                    btrDestroy hmFCff
'                    btrDestroy hmFVsf
'                    btrDestroy hmFCxf
'                    btrDestroy hmTChf
'                    btrDestroy hmTClf
'                    btrDestroy hmTCff
'                    btrDestroy hmTVsf
'                    btrDestroy hmTCxf
'                    mCopyCntrFiles = False
'                    Exit Function
'                End If
'                'ilRet = gObtainCntr(hmFChf, hmFClf, hmFCff, tmChf.lCode, False, tgChf, tgClf(), tgCff())
'                'For ilLoop = LBound(tgClf) To UBound(tgClf) - 1 Step 1
'                '    ilInsert = False
'                '    For ilTest = 0 To UBound(igPECSRNVefCode) - 1 Step 1
'                '        If (tgClf(ilLoop).ClfRec.iVefCode = igPECSRNVefCode(ilTest)) And ((tgClf(ilLoop).ClfRec.sType = "S") Or (tgClf(ilLoop).ClfRec.sType = "H")) Then
'                '            ilInsert = True
'                '            Exit For
'                '        End If
'                '    Next ilTest
'                '    If ilInsert Then
'                '        lgChfCode(ilUpper0, ilUpper1) = tmChf.lCode
'                '        ilUpper0 = ilUpper0 + 1
'                '        If ilUpper0 > 32000 Then
'                '            ilUpper0 = 1
'                '            ilUpper1 = ilUpper1 + 1
'                '            ReDim Preserve lgChfCode(1 To 32000, 1 To ilUpper1) As Long
'                '        End If
'                '        Exit For
'                '    End If
'                'Next ilLoop
'                ilInsert = False
'                If tmChf.lVefCode > 0 Then
'                    For ilTest = 0 To UBound(igPECSRNVefCode) - 1 Step 1
'                        If (tmChf.lVefCode = CLng(igPECSRNVefCode(ilTest))) Then
'                            ilInsert = True
'                            Exit For
'                        End If
'                    Next ilTest
'                ElseIf tmChf.lVefCode < 0 Then
'                    tmLongSrchKey.lCode = -tmChf.lVefCode
'                    ilRet = btrGetEqual(hmFVsf, tmVsf, ilVsfReclen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                    Do While ilRet = BTRV_ERR_NONE
'                        For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
'                            If tmVsf.iFSCode(ilLoop) > 0 Then
'                                For ilTest = 0 To UBound(igPECSRNVefCode) - 1 Step 1
'                                    If (tmVsf.iFSCode(ilLoop) = igPECSRNVefCode(ilTest)) Then
'                                        ilInsert = True
'                                        Exit For
'                                    End If
'                                Next ilTest
'                            End If
'                            If ilInsert Then
'                                Exit Do
'                            End If
'                        Next ilLoop
'                        If tmVsf.lLkVsfCode <= 0 Then
'                            Exit Do
'                        End If
'                        tmLongSrchKey.lCode = tmVsf.lLkVsfCode
'                        ilRet = btrGetEqual(hmFVsf, tmVsf, ilVsfReclen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                    Loop
'                End If
'                If ilInsert Then
'                    lgChfCode(ilUpper0, ilUpper1) = tmChf.lCode
'                    ilUpper0 = ilUpper0 + 1
'                    If ilUpper0 > 32000 Then
'                        ilUpper0 = 1
'                        ilUpper1 = ilUpper1 + 1
'                        ReDim Preserve lgChfCode(1 To 32000, 1 To ilUpper1) As Long
'                        For ilLoop = 1 To 32000 Step 1
'                            lgChfCode(ilLoop, ilUpper1) = 0
'                        Next ilLoop
'                    End If
'                End If
'                ilExtLen = imSvRecLen
'                ilRet = btrExtGetNext(hmFrom, tmChf, ilExtLen, llRecPos)
'                Do While ilRet = BTRV_ERR_REJECT_COUNT
'                    ilRet = btrExtGetNext(hmFrom, tmChf, ilExtLen, llRecPos)
'                Loop
'            Loop
'        End If
'    End If
    lmCount = 0
    For llLoop = LBound(lgChfCode) To UBound(lgChfCode) - 1 Step 1
        ilRet = mCloneCntr(lgChfCode(llLoop))
        If Not ilRet Then
'                btrDestroy hmFrom
            btrDestroy hmFChf
            btrDestroy hmFClf
            btrDestroy hmFCff
            btrDestroy hmFVsf
            btrDestroy hmFCxf
            btrDestroy hmFSbf
            btrDestroy hmFCgf
            btrDestroy hmFMsf
            btrDestroy hmTChf
            btrDestroy hmTClf
            btrDestroy hmTCff
            btrDestroy hmTVsf
            btrDestroy hmTCxf
            btrDestroy hmTSbf
            btrDestroy hmTCgf
            btrDestroy hmTMsf
            mCopyCntrFiles = False
            Exit Function
        End If
        lmCount = lmCount + 1
    Next llLoop
    Print #hmMsg, "    " & Str$(lmCount) & " Contracts (chf.btr, clf.btr, cff.btr, cxf.btr, vsf.btr) records Transferred"
'    btrDestroy hmFrom
    btrDestroy hmFChf
    btrDestroy hmFClf
    btrDestroy hmFCff
    btrDestroy hmFVsf
    btrDestroy hmFCxf
    btrDestroy hmFSbf
    btrDestroy hmFCgf
    btrDestroy hmFMsf
    btrDestroy hmTChf
    btrDestroy hmTClf
    btrDestroy hmTCff
    btrDestroy hmTVsf
    btrDestroy hmTCxf
    btrDestroy hmTSbf
    btrDestroy hmTCgf
    btrDestroy hmTMsf
    mCopyCntrFiles = True
    Exit Function
mCopyCntrFilesErr:
    Screen.MousePointer = vbDefault
'    btrDestroy hmFrom
    btrDestroy hmFChf
    btrDestroy hmFClf
    btrDestroy hmFCff
    btrDestroy hmFVsf
    btrDestroy hmFCxf
    btrDestroy hmFSbf
    btrDestroy hmFCgf
    btrDestroy hmFMsf
    btrDestroy hmTChf
    btrDestroy hmTClf
    btrDestroy hmTCff
    btrDestroy hmTVsf
    btrDestroy hmTCxf
    btrDestroy hmTSbf
    btrDestroy hmTCgf
    btrDestroy hmTMsf
    On Error GoTo 0
    mCopyCntrFiles = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyCrfRecs                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Variable length files     *
'*                                                     *
'*******************************************************
Private Function mCopyCrfRecs(slActiveDate As String) As Integer
    Dim ilRet As Integer
    Dim llNoRecords As Long
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Integer
    Dim ilLoop As Integer
    Dim llRecPos As Long
    Dim ilInsert As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlCnfSrchKey As CNFKEY0
    Dim ilCnfRecLen As Integer
    Dim ilCsfRecLen As Integer
    
    lacProcessing.Caption = "Processing " & "Crf.Btr"
    DoEvents
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCrfRecs: Open Crf") Then
        mCopyCrfRecs = False
        Exit Function
    End If
    imRecLen = btrRecordLength(hmFrom)  'Get and save record length
    imSvRecLen = imRecLen
    llNoRecords = btrRecords(hmFrom)
    mRemoveFile "Crf.Btr"
    ilRet = btrClone(hmFrom, smToDBPath & "Crf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyCrfRecs: Clone Crf") Then
        mCopyCrfRecs = False
        Exit Function
    End If
    hmFCnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFCnf, "", sgDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCrfRecs: Open 'From' Crf") Then
        mCopyCrfRecs = False
        Exit Function
    End If
    mRemoveFile "Cnf.Btr"
    ilRet = btrClone(hmFCnf, smToDBPath & "Cnf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyCrfRecs: Clone Cnf") Then
        mCopyCrfRecs = False
        Exit Function
    End If
    ilCnfRecLen = Len(tmCnf)

    hmFCsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFCsf, "", sgDBPath & "Csf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCrfFiles: Open From Csf") Then
        mCopyCrfRecs = False
        Exit Function
    End If
    mRemoveFile "Csf.Btr"
    ilRet = btrClone(hmFCsf, smToDBPath & "Csf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyCrfRecs: Clone Csf") Then
        mCopyCrfRecs = False
        Exit Function
    End If
    ilCsfRecLen = Len(tmCsf)


    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCrfRecs: Open 'To' Crf") Then
        mCopyCrfRecs = False
        Exit Function
    End If
    hmTCnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTCnf, "", smToDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCrfRecs: Open 'To' Cnf") Then
        mCopyCrfRecs = False
        Exit Function
    End If
    hmTCsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTCsf, "", smToDBPath & "Csf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCrfRecs: Open 'To' Csf") Then
        mCopyCrfRecs = False
        Exit Function
    End If
    
    lmCount = 0
    ilRet = btrGetFirst(hmFrom, tmCrf, imRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If (ilRet = BTRV_ERR_NONE) Then
        ilExtLen = imSvRecLen  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlVef) 'Obtain number of records
        btrExtClear hmFrom   'Clear any previous extend operation
        Call btrExtSetBounds(hmFrom, llNoRec, -1, "UC", "CRF", "") 'Set extract limits (all records)
        gPackDate slActiveDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Crf", "CrfEndDate")
        ilRet = btrExtAddLogicConst(hmFrom, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        ilOffSet = 0
        ilRet = btrExtAddField(hmFrom, ilOffSet, ilExtLen)  'Extract First Name field
        ilRet = btrExtGetNext(hmFrom, tmCrf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmMsg, gGetBtrvErrorMsg(ilRet) & "mCopyCrfRecs: ExtGetNext"
                btrDestroy hmFrom
                btrDestroy hmTo
                btrDestroy hmFCnf
                btrDestroy hmTCnf
                btrDestroy hmFCsf
                btrDestroy hmTCsf
                mCopyCrfRecs = False
                Exit Function
            End If
            ilExtLen = imSvRecLen
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmFrom, tmCrf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                DoEvents
                If imTerminate Then
                    btrDestroy hmFrom
                    btrDestroy hmTo
                    btrDestroy hmFCnf
                    btrDestroy hmTCnf
                    btrDestroy hmFCsf
                    btrDestroy hmTCsf
                    mCopyCrfRecs = False
                    Exit Function
                End If
                'tmCrf = tmRec
                ilInsert = False
                For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                    If tmCrf.iVefCode = igPECSRNVefCode(ilLoop) Then
                        ilInsert = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilInsert Then
                    For ilLoop = 0 To UBound(igPEAVefCode) - 1 Step 1
                        If tmCrf.iVefCode = igPEAVefCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                End If
                If ilInsert Then
                    ilRet = btrInsert(hmTo, tmCrf, ilExtLen, INDEXKEY0)
                    If mBtrErrorMsg(ilRet, "mCopyCrfRecs: Insert Crf") Then
                        btrDestroy hmFrom
                        btrDestroy hmTo
                        btrDestroy hmFCnf
                        btrDestroy hmTCnf
                        btrDestroy hmFCsf
                        btrDestroy hmTCsf
                        mCopyCrfRecs = False
                        Exit Function
                    End If
                    lmCount = lmCount + 1
                    If tmCrf.lCsfCode > 0 Then
                        ilCsfRecLen = Len(tmCxf)
                        tmLongSrchKey.lCode = tmCrf.lCsfCode
                        ilRet = btrGetEqual(hmFCsf, tmCsf, ilCsfRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            ilRet = btrInsert(hmTCsf, tmCsf, ilCsfRecLen, INDEXKEY0)
                        End If
                        If ilRet <> BTRV_ERR_NONE Then
                            tmCrf.lCsfCode = 0
                        End If
                    End If
                    tlCnfSrchKey.lCrfCode = tmCrf.lCode
                    tlCnfSrchKey.iInstrNo = 0
                    ilRet = btrGetGreaterOrEqual(hmFCnf, tmCnf, ilCnfRecLen, tlCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
                        ilRet = btrInsert(hmTCnf, tmCnf, ilCnfRecLen, INDEXKEY0)
                        If mBtrErrorMsg(ilRet, "mCopyCrfRecs: Insert CNF") Then
                            btrDestroy hmFrom
                            btrDestroy hmTo
                            btrDestroy hmFCnf
                            btrDestroy hmTCnf
                            btrDestroy hmFCsf
                            btrDestroy hmTCsf
                            mCopyCrfRecs = False
                            Exit Function
                        End If
                        ilRet = btrGetNext(hmFCnf, tmCnf, ilCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
                ilExtLen = imSvRecLen
                ilRet = btrExtGetNext(hmFrom, tmCrf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmFrom, tmCrf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    Print #hmMsg, "    " & Str$(lmCount) & " Copy Rotations (crf.btr) records Transferred"
    btrDestroy hmFrom
    btrDestroy hmTo
    btrDestroy hmFCnf
    btrDestroy hmTCnf
    btrDestroy hmFCsf
    btrDestroy hmTCsf
    mCopyCrfRecs = True
    Exit Function
mCopyCrfRecsErr:
    Screen.MousePointer = vbDefault
    btrDestroy hmFrom
    btrDestroy hmTo
    btrDestroy hmFCnf
    btrDestroy hmTCnf
    btrDestroy hmFCsf
    btrDestroy hmTCsf
    On Error GoTo 0
    mCopyCrfRecs = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyFile                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Fixed File                *
'*                                                     *
'*******************************************************
Private Function mCopyFile(slFileName As String) As Integer
    Dim ilRet As Integer
    Dim slStr As String
    lacProcessing.Caption = "Copying " & slFileName
    DoEvents
    ilRet = 0
    On Error GoTo mCopyFileErr0:
    slStr = FileDateTime(smToDBPath & slFileName)
    If ilRet <> 0 Then
        ilRet = 0
        slStr = FileDateTime(sgDBPath & slFileName)
        If ilRet = 0 Then
            On Error GoTo mCopyFileErr1
            FileCopy sgDBPath & slFileName, smToDBPath & slFileName
            Print #hmMsg, "    " & slFileName & " copied"
        Else
            Print #hmMsg, "File Not Copied as Missing " & slFileName
        End If
    Else
        ilRet = 0
        On Error GoTo mCopyFileErr0:
        If FileDateTime(sgDBPath & slFileName) <> FileDateTime(smToDBPath & slFileName) Then
            If ilRet = 0 Then
                mRemoveFile slFileName
                On Error GoTo mCopyFileErr1
                FileCopy sgDBPath & slFileName, smToDBPath & slFileName
                Print #hmMsg, "    " & slFileName & " copied"
            End If
        Else
            Print #hmMsg, "    " & slFileName & " not copied as same date"
        End If
    End If
    On Error GoTo 0
    mCopyFile = True
    Exit Function
mCopyFileErr0:
    ilRet = Err
    Resume Next
mCopyFileErr1:
    imVBErrorNo = Err
    Print #hmMsg, "Copy File Error" & Str$(imVBErrorNo) & " on " & slFileName
    On Error GoTo 0
    mCopyFile = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyFileFix                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Fixed Length Files        *
'*                                                     *
'*******************************************************
Private Function mCopyFileFix(slFileName As String) As Integer
    Dim ilRet As Integer
    Dim llNoRecords As Long
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Integer
    Dim llRecPos As Long
    Dim ilLoop As Integer
    Dim ilInsert As Integer
    Dim ilTest As Integer
    Dim ilKeyNo As Integer
    Dim slNameOfFile As String
    Dim ilPos As Integer
    Dim llLoop As Long
    Dim llDate As Long
    
    lacProcessing.Caption = "Processing " & slFileName
    DoEvents
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyFileFix: Open " & slFileName) Then
        mCopyFileFix = False
        Exit Function
    End If
    imRecLen = btrRecordLength(hmFrom)  'Get and save record length
    imSvRecLen = imRecLen
    llNoRecords = btrRecords(hmFrom)
    mRemoveFile slFileName
    ilRet = btrClone(hmFrom, smToDBPath & slFileName, 0)
    If mBtrErrorMsg(ilRet, "mCopyFileFix: Clone " & slFileName) Then
        mCopyFileFix = False
        Exit Function
    End If
    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyFileFix: Open To " & slFileName) Then
        mCopyFileFix = False
        Exit Function
    End If
    lmCount = 0
    ilKeyNo = INDEXKEY0
    If (StrComp(slFileName, "Dlf.btr", 1) = 0) Or (StrComp(slFileName, "Egf.btr", 1) = 0) Then
        ilKeyNo = INDEXKEY1
    ElseIf (StrComp(slFileName, "Rif.btr", 1) = 0) Then
        ilKeyNo = INDEXKEY1
    ElseIf (StrComp(slFileName, "Rvf.btr", 1) = 0) Or (StrComp(slFileName, "Phf.btr", 1) = 0) Then
        ilKeyNo = INDEXKEY2
    ElseIf (StrComp(slFileName, "Stf.btr", 1) = 0) Then
        ilKeyNo = INDEXKEY1
    End If
    ilPos = InStr(1, slFileName, ".", vbTextCompare)
    If ilPos > 0 Then
        slNameOfFile = Left$(slFileName, ilPos - 1)
    Else
        slNameOfFile = slFileName
    End If
    ilRet = btrGetFirst(hmFrom, tmRec, imRecLen, ilKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)
    If (ilRet = BTRV_ERR_NONE) Then
        ilExtLen = imSvRecLen  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlVef) 'Obtain number of records
        btrExtClear hmFrom   'Clear any previous extend operation
        Call btrExtSetBounds(hmFrom, llNoRec, -1, "UC", slNameOfFile, "") 'Set extract limits (all records)
        ilOffSet = 0
        ilRet = btrExtAddField(hmFrom, ilOffSet, ilExtLen)  'Extract First Name field
        If (StrComp(slFileName, "Dlf.btr", 1) = 0) Or (StrComp(slFileName, "Egf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmDlf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Rif.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmRif, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Mtf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmMtf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Vef.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmVef, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Ltf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmLtf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Lvf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmLvf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Lef.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmLef, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Vof.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmVof, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Vpf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmVpf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Rvf.btr", 1) = 0) Or (StrComp(slFileName, "Phf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmRvf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Adf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmAdf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Agf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmAgf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Slf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmSlf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Prf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmPrf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Sif.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmSif, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Saf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmSaf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Bsf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmBsf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Bvf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmBvf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Scf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmScf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Stf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmStf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Ihf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmIhf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Isf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmIsf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Ghf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmGhf, ilExtLen, llRecPos)
        ElseIf (StrComp(slFileName, "Gsf.btr", 1) = 0) Then
            ilRet = btrExtGetNext(hmFrom, tmGsf, ilExtLen, llRecPos)
        Else
            ilRet = btrExtGetNext(hmFrom, tmRec, ilExtLen, llRecPos)
        End If
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmMsg, gGetBtrvErrorMsg(ilRet) & "mCopyFileFix: ExtGetNext"
                btrDestroy hmFrom
                btrDestroy hmTo
                mCopyFileFix = False
                Exit Function
            End If
            ilExtLen = imSvRecLen
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                'ilRet = btrExtGetNext(hmFrom, tmRec, ilExtLen, llRecPos)
                If (StrComp(slFileName, "Dlf.btr", 1) = 0) Or (StrComp(slFileName, "Egf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmDlf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Rif.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmRif, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Mtf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmMtf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Vef.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmVef, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Ltf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmLtf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Lvf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmLvf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Lef.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmLef, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Vof.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmVof, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Vpf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmVpf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Rvf.btr", 1) = 0) Or (StrComp(slFileName, "Phf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmRvf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Adf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmAdf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Agf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmAgf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Slf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmSlf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Prf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmPrf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Sif.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmSif, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Saf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmSaf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Bsf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmBsf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Bvf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmBvf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Scf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmScf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Stf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmStf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Ihf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmIhf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Isf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmIsf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Ghf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmGhf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Gsf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmGsf, ilExtLen, llRecPos)
                Else
                    ilRet = btrExtGetNext(hmFrom, tmRec, ilExtLen, llRecPos)
                End If
            Loop
            Do While ilRet = BTRV_ERR_NONE
                DoEvents
                If imTerminate Then
                    btrDestroy hmFrom
                    btrDestroy hmTo
                    mCopyFileFix = False
                    Exit Function
                End If
                ilInsert = False
                If (StrComp(slFileName, "Dlf.btr", 1) = 0) Or (StrComp(slFileName, "Egf.btr", 1) = 0) Then
                    'tmDlf = tmRec
                    For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If tmDlf.iVefCode = igPECSRNVefCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilInsert Then
                        For ilLoop = 0 To UBound(igPEAVefCode) - 1 Step 1
                            If tmDlf.iVefCode = igPEAVefCode(ilLoop) Then
                                ilInsert = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                ElseIf (StrComp(slFileName, "Rif.btr", 1) = 0) Then
                    'tmRif = tmRec
                    For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If tmRif.iVefCode = igPECSRNVefCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilInsert Then
                        For ilLoop = 0 To UBound(igAddPkVefCode) - 1 Step 1
                            If tmRif.iVefCode = igAddPkVefCode(ilLoop) Then
                                ilInsert = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                ElseIf (StrComp(slFileName, "Mtf.btr", 1) = 0) Then
                    'tmMtf = tmRec
                    For llLoop = LBound(lgChfCode) To UBound(lgChfCode) - 1 Step 1
                       If tmMtf.lChfCode = lgChfCode(llLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next llLoop
                ElseIf (StrComp(slFileName, "Vef.btr", 1) = 0) Then
                    'tmVef = tmRec
                    For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If tmVef.iCode = igPECSRNVefCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilInsert Then
                        For ilLoop = 0 To UBound(igPEAVefCode) - 1 Step 1
                            If tmVef.iCode = igPEAVefCode(ilLoop) Then
                                ilInsert = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If Not ilInsert Then
                        For ilLoop = 0 To UBound(igPELVefCode) - 1 Step 1
                            If tmVef.iCode = igPELVefCode(ilLoop) Then
                                ilInsert = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If Not ilInsert Then
                        For ilLoop = 0 To UBound(igAddPkVefCode) - 1 Step 1
                            If tmVef.iCode = igAddPkVefCode(ilLoop) Then
                                ilInsert = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                ElseIf (StrComp(slFileName, "Ltf.btr", 1) = 0) Then
                    'tmLtf = tmRec
                    For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If tmLtf.iVefCode = igPECSRNVefCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilInsert Then
                        For ilLoop = 0 To UBound(igPEAVefCode) - 1 Step 1
                            If tmLtf.iVefCode = igPEAVefCode(ilLoop) Then
                                ilInsert = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If Not ilInsert Then
                        For ilLoop = 0 To UBound(igPELVefCode) - 1 Step 1
                            If tmLtf.iVefCode = igPELVefCode(ilLoop) Then
                                ilInsert = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If ilInsert Then
                        igPELtfCode(UBound(igPELtfCode)) = tmLtf.iCode
                        ReDim Preserve igPELtfCode(0 To UBound(igPELtfCode) + 1) As Integer
                    End If
                ElseIf (StrComp(slFileName, "Lvf.btr", 1) = 0) Then
                    'tmLvf = tmRec
                    For ilLoop = 0 To UBound(igPELtfCode) - 1 Step 1
                        If tmLvf.iLtfCode = igPELtfCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                    If ilInsert Then
                        lgPELvfCode(UBound(lgPELvfCode)) = tmLvf.lCode
                        ReDim Preserve lgPELvfCode(0 To UBound(lgPELvfCode) + 1) As Long
                    End If
                ElseIf (StrComp(slFileName, "Lef.btr", 1) = 0) Then
                    'tmLef = tmRec
                    For ilLoop = 0 To UBound(lgPELvfCode) - 1 Step 1
                        If tmLef.lLvfCode = lgPELvfCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                ElseIf (StrComp(slFileName, "Vof.btr", 1) = 0) Then
                    'tmVof = tmRec
                    For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If tmVof.iVefCode = igPECSRNVefCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilInsert Then
                        For ilLoop = 0 To UBound(igPEAVefCode) - 1 Step 1
                            If tmVof.iVefCode = igPEAVefCode(ilLoop) Then
                                ilInsert = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If Not ilInsert Then
                        For ilLoop = 0 To UBound(igPELVefCode) - 1 Step 1
                            If tmVof.iVefCode = igPELVefCode(ilLoop) Then
                                ilInsert = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                ElseIf (StrComp(slFileName, "Vpf.btr", 1) = 0) Then
                    'tmVpf = tmRec
                    For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If tmVpf.iVefKCode = igPECSRNVefCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilInsert Then
                        For ilLoop = 0 To UBound(igPEAVefCode) - 1 Step 1
                            If tmVpf.iVefKCode = igPEAVefCode(ilLoop) Then
                                ilInsert = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If Not ilInsert Then
                        For ilLoop = 0 To UBound(igPELVefCode) - 1 Step 1
                            If tmVpf.iVefKCode = igPELVefCode(ilLoop) Then
                                ilInsert = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                ElseIf (StrComp(slFileName, "Rvf.btr", 1) = 0) Or (StrComp(slFileName, "Phf.btr", 1) = 0) Then
                    'tmRvf = tmRec
                    If tmRvf.iAirVefCode > 0 Then
                        For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                            If tmRvf.iAirVefCode = igPECSRNVefCode(ilLoop) Then
                                ilInsert = True
                                Exit For
                            End If
                        Next ilLoop
                        If Not ilInsert Then
                            For ilLoop = 0 To UBound(igPEAVefCode) - 1 Step 1
                                If tmRvf.iAirVefCode = igPEAVefCode(ilLoop) Then
                                    ilInsert = True
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                        If Not ilInsert Then
                            For ilLoop = 0 To UBound(igPELVefCode) - 1 Step 1
                                If tmRvf.iAirVefCode = igPELVefCode(ilLoop) Then
                                    ilInsert = True
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                    Else
                        ilRet = mBinarySearch(igPECntrAgfCode(), tmRvf.iAgfCode)
                        If ilRet <> -1 Then
                            ilInsert = True
                        Else
                            If tmRvf.iAdfCode > 0 Then
                                ilRet = mBinarySearch(igPECntrAdfCode(), tmAdf.iCode)
                                If ilRet <> -1 Then
                                    ilInsert = True
                                End If
                            End If
                        End If
                    End If
                    If ilInsert Then
                        If Not imIncludeCollection Then
                            If (tmRvf.sCashTrade <> "M") And (tmRvf.sCashTrade <> "P") Then
                                ilInsert = False
                            End If
                        End If
                    End If
                ElseIf (StrComp(slFileName, "Adf.btr", 1) = 0) Then
                    ilRet = mBinarySearch(igPECntrAdfCode(), tmAdf.iCode)
                    If ilRet <> -1 Then
                        ilInsert = True
                        If tmAdf.iTrfCode > 0 Then
                            imIncludeTaxTable = True
                        End If
                    End If
                ElseIf (StrComp(slFileName, "Agf.btr", 1) = 0) Then
                    ilRet = mBinarySearch(igPECntrAgfCode(), tmAgf.iCode)
                    If ilRet <> -1 Then
                        ilInsert = True
                        If tmAgf.iTrfCode > 0 Then
                            imIncludeTaxTable = True
                        End If
                    End If
                ElseIf (StrComp(slFileName, "Slf.btr", 1) = 0) Then
                    ilRet = mBinarySearch(igPECntrSlfCode(), tmSlf.iCode)
                    If ilRet <> -1 Then
                        ilInsert = True
                    End If
                ElseIf (StrComp(slFileName, "Prf.btr", 1) = 0) Then
                    ilRet = mBinarySearch(igPECntrAdfCode(), tmPrf.iAdfCode)
                    If ilRet <> -1 Then
                        ilInsert = True
                    End If
                ElseIf (StrComp(slFileName, "Sif.btr", 1) = 0) Then
                    ilRet = mBinarySearch(igPECntrAdfCode(), tmSif.iAdfCode)
                    If ilRet <> -1 Then
                        ilInsert = True
                    End If
                ElseIf (StrComp(slFileName, "Saf.btr", 1) = 0) Then
                    If tmSaf.iVefCode = 0 Then
                        ilInsert = True
                    Else
                        For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                            If tmSaf.iVefCode = igPECSRNVefCode(ilLoop) Then
                                ilInsert = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                ElseIf (StrComp(slFileName, "Bsf.btr", 1) = 0) Then
                    ilRet = mBinarySearch(igPECntrSlfCode(), tmBsf.iSlfCode)
                    If ilRet <> -1 Then
                        ilInsert = True
                    End If
                ElseIf (StrComp(slFileName, "Bvf.btr", 1) = 0) Then
                    For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If tmBvf.iVefCode = igPECSRNVefCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                ElseIf (StrComp(slFileName, "Scf.btr", 1) = 0) Then
                    'tmVef = tmRec
                    For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If tmScf.iVefCode = igPECSRNVefCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                ElseIf (StrComp(slFileName, "Stf.btr", 1) = 0) Then
                    'tmVef = tmRec
                    For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If tmStf.iVefCode = igPECSRNVefCode(ilLoop) Then
                            gUnpackDateLong tmStf.iLogDate(0), tmStf.iLogDate(1), llDate
                            If llDate >= lmPrevStdStartDate Then
                                ilInsert = True
                            End If
                            Exit For
                        End If
                    Next ilLoop
                ElseIf (StrComp(slFileName, "Ihf.btr", 1) = 0) Then
                    For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If tmIhf.iVefCode = igPECSRNVefCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                ElseIf (StrComp(slFileName, "Isf.btr", 1) = 0) Then
                    'tmVof = tmRec
                    For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If tmIsf.iVefCode = igPECSRNVefCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                ElseIf (StrComp(slFileName, "Ghf.btr", 1) = 0) Then
                    'tmVof = tmRec
                    For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If tmGhf.iVefCode = igPECSRNVefCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                ElseIf (StrComp(slFileName, "Gsf.btr", 1) = 0) Then
                    'tmVof = tmRec
                    For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If tmGsf.iVefCode = igPECSRNVefCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                Else
                    ilInsert = True
                End If
                If ilInsert Then
                    'ilRet = btrInsert(hmTo, tmRec, imRecLen, INDEXKEY0)
                    If (StrComp(slFileName, "Dlf.btr", 1) = 0) Or (StrComp(slFileName, "Egf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmDlf, imRecLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Rif.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmRif, imRecLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Mtf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmMtf, imRecLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Vef.btr", 1) = 0) Then
                        If Not imIncludeResearch Then
                            tmVef.iDnfCode = 0
                            tmVef.iReallDnfCode = 0
                        End If
                        ilRet = btrInsert(hmTo, tmVef, imRecLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Ltf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmLtf, imRecLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Lvf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmLvf, imRecLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Lef.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmLef, imRecLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Vof.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmVof, imRecLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Vpf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmVpf, imRecLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Rvf.btr", 1) = 0) Or (StrComp(slFileName, "Phf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmRvf, imRecLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Adf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmAdf, ilExtLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Agf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmAgf, ilExtLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Slf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmSlf, ilExtLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Prf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmPrf, ilExtLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Sif.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmSif, ilExtLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Saf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmSaf, ilExtLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Bsf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmBsf, ilExtLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Bvf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmBvf, ilExtLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Scf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmScf, ilExtLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Stf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmStf, ilExtLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Ihf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmIhf, imRecLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Isf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmIsf, imRecLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Ghf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmGhf, imRecLen, INDEXKEY0)
                    ElseIf (StrComp(slFileName, "Gsf.btr", 1) = 0) Then
                        ilRet = btrInsert(hmTo, tmGsf, imRecLen, INDEXKEY0)
                    Else
                        ilRet = btrInsert(hmTo, tmRec, imRecLen, INDEXKEY0)
                    End If
                    If mBtrErrorMsg(ilRet, "mCopyFileFix: Insert " & slFileName) Then
                        btrDestroy hmFrom
                        btrDestroy hmTo
                        mCopyFileFix = False
                        Exit Function
                    End If
                    lmCount = lmCount + 1
                End If
                ilExtLen = imSvRecLen
                'ilRet = btrExtGetNext(hmFrom, tmRec, ilExtLen, llRecPos)
                If (StrComp(slFileName, "Dlf.btr", 1) = 0) Or (StrComp(slFileName, "Egf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmDlf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Rif.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmRif, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Mtf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmMtf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Vef.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmVef, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Ltf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmLtf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Lvf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmLvf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Lef.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmLef, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Vof.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmVof, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Vpf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmVpf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Rvf.btr", 1) = 0) Or (StrComp(slFileName, "Phf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmRvf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Adf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmAdf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Agf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmAgf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Slf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmSlf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Prf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmPrf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Sif.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmSif, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Saf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmSaf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Bsf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmBsf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Bvf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmBvf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Scf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmScf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Stf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmStf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Ihf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmIhf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Isf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmIsf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Ghf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmGhf, ilExtLen, llRecPos)
                ElseIf (StrComp(slFileName, "Gsf.btr", 1) = 0) Then
                    ilRet = btrExtGetNext(hmFrom, tmGsf, ilExtLen, llRecPos)
                Else
                    ilRet = btrExtGetNext(hmFrom, tmRec, ilExtLen, llRecPos)
                End If
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    'ilRet = btrExtGetNext(hmFrom, tmRec, ilExtLen, llRecPos)
                    If (StrComp(slFileName, "Dlf.btr", 1) = 0) Or (StrComp(slFileName, "Egf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmDlf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Rif.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmRif, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Mtf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmMtf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Vef.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmVef, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Ltf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmLtf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Lvf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmLvf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Lef.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmLef, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Vof.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmVof, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Vpf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmVpf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Rvf.btr", 1) = 0) Or (StrComp(slFileName, "Phf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmRvf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Adf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmAdf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Agf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmAgf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Slf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmSlf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Prf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmPrf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Sif.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmSif, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Saf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmSaf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Bsf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmBsf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Bvf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmBvf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Scf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmScf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Stf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmStf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Ihf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmIhf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Isf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmIsf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Ghf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmGhf, ilExtLen, llRecPos)
                    ElseIf (StrComp(slFileName, "Gsf.btr", 1) = 0) Then
                        ilRet = btrExtGetNext(hmFrom, tmGsf, ilExtLen, llRecPos)
                    Else
                        ilRet = btrExtGetNext(hmFrom, tmRec, ilExtLen, llRecPos)
                    End If
                Loop
            Loop
        End If
    End If
    Print #hmMsg, "    " & Str$(lmCount) & " records from " & slFileName & " Transferred"
    btrDestroy hmFrom
    btrDestroy hmTo
    mCopyFileFix = True
    Exit Function
mCopyFileFixErr:
    Screen.MousePointer = vbDefault
    imVBErrorNo = Err
    Print #hmMsg, "Copy File Fix Error" & Str$(imVBErrorNo) & " on " & slFileName
    btrDestroy hmFrom
    btrDestroy hmTo
    On Error GoTo 0
    mCopyFileFix = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyFileVar                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Variable length files     *
'*                                                     *
'*******************************************************
Private Function mCopyFileVar(slFileName As String) As Integer
    Dim ilRet As Integer
    Dim llNoRecords As Long
    Dim llRecPos As Long
    lacProcessing.Caption = "Processing " & slFileName
    DoEvents
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyFileVar: Open " & slFileName) Then
        mCopyFileVar = False
        Exit Function
    End If
    imRecLen = 10000    'btrRecordLength(hmFrom)  'Get and save record length
    imSvRecLen = imRecLen
    llNoRecords = btrRecords(hmFrom)
    mRemoveFile slFileName
    ilRet = btrClone(hmFrom, smToDBPath & slFileName, 0)
    If mBtrErrorMsg(ilRet, "mCopyFileVar: Clone " & slFileName) Then
        mCopyFileVar = False
        Exit Function
    End If
    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyFileVar: Open To " & slFileName) Then
        mCopyFileVar = False
        Exit Function
    End If
    lmCount = 0
    ilRet = btrGetFirst(hmFrom, tmRec, imRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE)
        DoEvents
        If imTerminate Then
            btrDestroy hmFrom
            btrDestroy hmTo
            mCopyFileVar = False
            Exit Function
        End If
        ilRet = btrInsert(hmTo, tmRec, imRecLen, INDEXKEY0)
        If mBtrErrorMsg(ilRet, "mCopyFileVar: Insert " & slFileName) Then
            btrDestroy hmFrom
            btrDestroy hmTo
            mCopyFileVar = False
            Exit Function
        End If
        lmCount = lmCount + 1
        imRecLen = imSvRecLen
        ilRet = btrGetNext(hmFrom, tmRec, imRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Print #hmMsg, "    " & Str$(lmCount) & " records from " & slFileName & " Transferred"
    btrDestroy hmFrom
    btrDestroy hmTo
    mCopyFileVar = True
    Exit Function
mCopyFileVarErr:
    Screen.MousePointer = vbDefault
    btrDestroy hmFrom
    btrDestroy hmTo
    On Error GoTo 0
    mCopyFileVar = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyLcfRecs                   *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Fixed Length Files        *
'*                                                     *
'*******************************************************
Private Function mCopyLcfRecs(slSpotDate As String) As Integer
    Dim ilRet As Integer
    Dim llNoRecords As Long
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Integer
    Dim ilLoop As Integer
    Dim ilInsert As Integer
    Dim llInDate As Long
    Dim llTstDate As Long
    Dim llRecPos As Long
    'Dim slType As String
    Dim ilType As Integer
    Dim tlLcfSrchKey As LCFKEY0     'LCF key record image
    Dim ilLcfRecLen As Integer         'LCF record length
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    llInDate = gDateValue(slSpotDate)
    lacProcessing.Caption = "Processing " & "Lcf.Btr"
    DoEvents
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyLcfRecs: Open Lcf") Then
        mCopyLcfRecs = False
        Exit Function
    End If
    imRecLen = btrRecordLength(hmFrom)  'Get and save record length
    'Determine LCF start date for airing vehciles
    ilLcfRecLen = Len(tmLcf)
    'slType = "O"
    ilType = 0
    ReDim lgLCFAirStartDate(0 To UBound(igPEAVefCode)) As Long
    For ilLoop = 0 To UBound(igPEAVefCode) - 1 Step 1
        tlLcfSrchKey.iType = ilType
        tlLcfSrchKey.sStatus = "C"
        tlLcfSrchKey.iVefCode = igPEAVefCode(ilLoop)
        tlLcfSrchKey.iLogDate(0) = 257  'Year 1/1/2100
        tlLcfSrchKey.iLogDate(1) = 2100
        tlLcfSrchKey.iSeqNo = 1
        ilRet = btrGetLessOrEqual(hmFrom, tmLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If (ilRet = BTRV_ERR_NONE) And (tmLcf.sStatus = "C") And (tmLcf.iVefCode = igPEAVefCode(ilLoop)) And (tmLcf.iType = ilType) Then
            gUnpackDateLong tmLcf.iLogDate(0), tmLcf.iLogDate(1), llTstDate
            If llTstDate <= llInDate Then
                lgLCFAirStartDate(ilLoop) = llTstDate - 6
            Else
                lgLCFAirStartDate(ilLoop) = llInDate
            End If
        Else
            lgLCFAirStartDate(ilLoop) = llInDate
        End If
    Next ilLoop
    imSvRecLen = imRecLen
    llNoRecords = btrRecords(hmFrom)
    mRemoveFile "Lcf.Btr"
    ilRet = btrClone(hmFrom, smToDBPath & "Lcf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyLcfRecs: Clone Lcf") Then
        mCopyLcfRecs = False
        Exit Function
    End If
    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyLcfRecs: Open 'To' Lcf") Then
        mCopyLcfRecs = False
        Exit Function
    End If
    lmCount = 0
    ilRet = btrGetFirst(hmFrom, tmLcf, imRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If (ilRet = BTRV_ERR_NONE) Then
        ilExtLen = imSvRecLen  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlVef) 'Obtain number of records
        btrExtClear hmFrom   'Clear any previous extend operation
        Call btrExtSetBounds(hmFrom, llNoRec, -1, "UC", "LCF", "") 'Set extract limits (all records)
        'Do test below as airing dates could be prior to specified dates
        'gPackDate slSpotDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        'ilOffset = gFieldOffset("Lcf", "LcfLogDate")
        'ilRet = btrExtAddLogicConst(hmFrom, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        ilOffSet = 0
        ilRet = btrExtAddField(hmFrom, ilOffSet, ilExtLen)  'Extract First Name field
        ilRet = btrExtGetNext(hmFrom, tmLcf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmMsg, gGetBtrvErrorMsg(ilRet) & "mCopyLcfRecs: ExtGetNext"
                btrDestroy hmFrom
                btrDestroy hmTo
                mCopyLcfRecs = False
                Exit Function
            End If
            ilExtLen = imSvRecLen
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmFrom, tmLcf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                DoEvents
                If imTerminate Then
                    btrDestroy hmFrom
                    btrDestroy hmTo
                    mCopyLcfRecs = False
                    Exit Function
                End If
                ilInsert = False
                'tmLcf = tmRec
                For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                    If tmLcf.iVefCode = igPECSRNVefCode(ilLoop) Then
                        If (tmLcf.iLogDate(0) >= 1) And (tmLcf.iLogDate(0) <= 7) And (tmLcf.iLogDate(1) = 0) Then
                            ilInsert = True
                        Else
                            gUnpackDateLong tmLcf.iLogDate(0), tmLcf.iLogDate(1), llTstDate
                            If llTstDate >= llInDate Then
                                ilInsert = True
                            End If
                        End If
                        Exit For
                    End If
                Next ilLoop
                If Not ilInsert Then
                    For ilLoop = 0 To UBound(igPEAVefCode) - 1 Step 1
                        If tmLcf.iVefCode = igPEAVefCode(ilLoop) Then
                            If (tmLcf.iLogDate(0) >= 1) And (tmLcf.iLogDate(0) <= 7) And (tmLcf.iLogDate(1) = 0) Then
                                ilInsert = True
                            Else
                                gUnpackDateLong tmLcf.iLogDate(0), tmLcf.iLogDate(1), llTstDate
                                If llTstDate >= lgLCFAirStartDate(ilLoop) Then
                                    ilInsert = True
                                End If
                            End If
                            Exit For
                        End If
                    Next ilLoop
                End If
                If ilInsert Then
                    ilRet = btrInsert(hmTo, tmLcf, ilExtLen, INDEXKEY0)
                    If mBtrErrorMsg(ilRet, "mCopyLcfRecs: Insert") Then
                        btrDestroy hmFrom
                        btrDestroy hmTo
                        mCopyLcfRecs = False
                        Exit Function
                    End If
                    lmCount = lmCount + 1
                End If
                ilExtLen = imSvRecLen
                ilRet = btrExtGetNext(hmFrom, tmLcf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmFrom, tmLcf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    Print #hmMsg, "    " & Str$(lmCount) & " Library Calendar Structures (lcf.btr) records Transferred"
    btrDestroy hmFrom
    btrDestroy hmTo
    mCopyLcfRecs = True
    Exit Function
mCopyLcfRecsErr:
    Screen.MousePointer = vbDefault
    btrDestroy hmFrom
    btrDestroy hmTo
    On Error GoTo 0
    mCopyLcfRecs = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyPjfRecs                   *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Fixed Length Files        *
'*                                                     *
'*******************************************************
Private Function mCopyPjfRecs() As Integer
    Dim ilRet As Integer
    Dim llNoRecords As Long
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Integer
    Dim ilLoop As Integer
    Dim llRecPos As Long
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    lacProcessing.Caption = "Processing " & "Pjf.Btr"
    DoEvents
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & "Pjf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyPjfRecs: Open Pjf") Then
        mCopyPjfRecs = False
        Exit Function
    End If
    imRecLen = btrRecordLength(hmFrom)  'Get and save record length
    imSvRecLen = imRecLen
    llNoRecords = btrRecords(hmFrom)
    mRemoveFile "Pjf.Btr"
    ilRet = btrClone(hmFrom, smToDBPath & "Pjf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyPjfRecs: Clone") Then
        mCopyPjfRecs = False
        Exit Function
    End If
    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & "Pjf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyPjfRecs: Open 'To' Pjf") Then
        mCopyPjfRecs = False
        Exit Function
    End If
    lmCount = 0
    ilRet = btrGetFirst(hmFrom, tmPjf, imRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If (ilRet = BTRV_ERR_NONE) Then
        ilExtLen = imSvRecLen  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlVef) 'Obtain number of records
        btrExtClear hmFrom   'Clear any previous extend operation
        Call btrExtSetBounds(hmFrom, llNoRec, -1, "UC", "PJF", "") 'Set extract limits (all records)
        tlDateTypeBuff.iDate0 = 0
        tlDateTypeBuff.iDate1 = 0
        ilOffSet = gFieldOffset("Pjf", "PjfRolloverDate")
        ilRet = btrExtAddLogicConst(hmFrom, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        ilOffSet = 0
        ilRet = btrExtAddField(hmFrom, ilOffSet, ilExtLen)  'Extract First Name field
        ilRet = btrExtGetNext(hmFrom, tmPjf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmMsg, gGetBtrvErrorMsg(ilRet) & "mCopyPjfRecs: ExtGetNext"
                btrDestroy hmFrom
                btrDestroy hmTo
                mCopyPjfRecs = False
                Exit Function
            End If
            ilExtLen = imSvRecLen
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmFrom, tmPjf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                DoEvents
                If imTerminate Then
                    btrDestroy hmFrom
                    btrDestroy hmTo
                    mCopyPjfRecs = False
                    Exit Function
                End If
                If mBinarySearch(igPECntrSlfCode(), tmPjf.iSlfCode) <> -1 Then
                    ilRet = btrInsert(hmTo, tmPjf, ilExtLen, INDEXKEY0)
                    If mBtrErrorMsg(ilRet, "mCopyPjfRecs: Insert") Then
                        btrDestroy hmFrom
                        btrDestroy hmTo
                        mCopyPjfRecs = False
                        Exit Function
                    End If
                    lmCount = lmCount + 1
                End If
                ilExtLen = imSvRecLen
                ilRet = btrExtGetNext(hmFrom, tmPjf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmFrom, tmPjf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    Print #hmMsg, "    " & Str$(lmCount) & " Contract Projections (pjf.btr) records Transferred"
    btrDestroy hmFrom
    btrDestroy hmTo
    mCopyPjfRecs = True
    Exit Function
mCopyPjfRecsErr:
    Screen.MousePointer = vbDefault
    btrDestroy hmFrom
    btrDestroy hmTo
    On Error GoTo 0
    mCopyPjfRecs = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyPsfRecs                   *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Fixed Length Files        *
'*                                                     *
'*******************************************************
Private Function mCopyPsfRecs(slStartDate As String) As Integer
    Dim ilRet As Integer
    Dim llNoRecords As Long
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Integer
    Dim ilLoop As Integer
    Dim llRecPos As Long
    Dim ilInsert As Integer
    Dim llSpotDate As Long
    Dim llTstDate As Long
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    lacProcessing.Caption = "Processing " & "Psf.Btr"
    DoEvents
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & "Psf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyPsfRecs: Open") Then
        mCopyPsfRecs = False
        Exit Function
    End If
    imRecLen = btrRecordLength(hmFrom)  'Get and save record length
    imSvRecLen = imRecLen
    llNoRecords = btrRecords(hmFrom)
    mRemoveFile "Psf.Btr"
    ilRet = btrClone(hmFrom, smToDBPath & "Psf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyPsfRecs: Clone") Then
        mCopyPsfRecs = False
        Exit Function
    End If
    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & "Psf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyPsfRecs: Open 'To' Psf") Then
        mCopyPsfRecs = False
        Exit Function
    End If
    llSpotDate = gDateValue(slStartDate)
    lmCount = 0
    ilRet = btrGetFirst(hmFrom, tmSdf, imRecLen, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
    If (ilRet = BTRV_ERR_NONE) Then
        ilExtLen = imSvRecLen  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlVef) 'Obtain number of records
        btrExtClear hmFrom   'Clear any previous extend operation
        Call btrExtSetBounds(hmFrom, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)
        'gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        'ilOffset = gFieldOffset("Psf", "SdfDate")
        'ilRet = btrExtAddLogicConst(hmFrom, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        'gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        'ilOffset = gFieldOffset("Psf", "SdfDate")
        'ilRet = btrExtAddLogicConst(hmFrom, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        ilOffSet = 0
        ilRet = btrExtAddField(hmFrom, ilOffSet, ilExtLen)  'Extract First Name field
        ilRet = btrExtGetNext(hmFrom, tmSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmMsg, gGetBtrvErrorMsg(ilRet) & "mCopyPsfRecs: ExtGetNext"
                btrDestroy hmFrom
                btrDestroy hmTo
                mCopyPsfRecs = False
                Exit Function
            End If
            ilExtLen = imSvRecLen
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmFrom, tmSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                DoEvents
                If imTerminate Then
                    btrDestroy hmFrom
                    btrDestroy hmTo
                    mCopyPsfRecs = False
                    Exit Function
                End If
                ilInsert = False
                For ilLoop = 0 To UBound(igAddPkVefCode) - 1 Step 1
                    If tmSdf.iVefCode = igAddPkVefCode(ilLoop) Then
                        ilInsert = True
                        Exit For
                    End If
                Next ilLoop
                If ilInsert Then
                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llTstDate
                    If llTstDate < llSpotDate Then
                        ilInsert = False
                    End If
                End If
                If ilInsert Then
                    If Not mCopyOK(tmSdf.sPtType, tmSdf.lCopyCode) Then
                        tmSdf.iRotNo = 0
                        tmSdf.lCopyCode = 0
                        tmSdf.sPtType = ""
                    End If
                    If Not imIncludeCollection Then
                        tmSdf.sBill = "N"
                    End If
                    ilRet = btrInsert(hmTo, tmSdf, ilExtLen, INDEXKEY0)
                    If mBtrErrorMsg(ilRet, "mCopyPsfRecs: Insert") Then
                        btrDestroy hmFrom
                        btrDestroy hmTo
                        mCopyPsfRecs = False
                        Exit Function
                    End If
                    lmCount = lmCount + 1
                End If
                ilExtLen = imSvRecLen
                ilRet = btrExtGetNext(hmFrom, tmSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmFrom, tmSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    Print #hmMsg, "    " & Str$(lmCount) & " Package Spots (psf.btr) records Transferred"
    btrDestroy hmFrom
    btrDestroy hmTo
    mCopyPsfRecs = True
    Exit Function
mCopyPsfRecsErr:
    Screen.MousePointer = vbDefault
    btrDestroy hmFrom
    btrDestroy hmTo
    On Error GoTo 0
    mCopyPsfRecs = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyPvfRecs                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Variable length files     *
'*                                                     *
'*******************************************************
Private Function mCopyPvfRecs() As Integer
    Dim ilRet As Integer
    Dim llNoRecords As Long
    Dim llRecPos As Long
    Dim llDate As Long
    Dim llInDate As Long
    Dim llTDate As Long
    Dim ilLoop As Integer
    Dim ilInsert As Integer
    Dim ilVeh As Integer
    Dim tlPvfSrchKey As LONGKEY0
    lacProcessing.Caption = "Processing " & "Pvf.Btr"
    DoEvents
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & "Pvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyPvfRecs: Open") Then
        mCopyPvfRecs = False
        Exit Function
    End If
    imRecLen = btrRecordLength(hmFrom)  'Get and save record length
    imSvRecLen = imRecLen
    llNoRecords = btrRecords(hmFrom)
    mRemoveFile "Pvf.Btr"
    ilRet = btrClone(hmFrom, smToDBPath & "Pvf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyPvfRecs: Clone") Then
        mCopyPvfRecs = False
        Exit Function
    End If
    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & "Pvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyPvfRecs: Open 'To' Pvf") Then
        mCopyPvfRecs = False
        Exit Function
    End If
    lmCount = 0
    For ilLoop = 0 To UBound(igAddPkVefCode) - 1 Step 1
        For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If igAddPkVefCode(ilLoop) = tgMVef(ilVeh).iCode Then
                If tgMVef(ilVeh).lPvfCode > 0 Then
                    tlPvfSrchKey.lCode = tgMVef(ilVeh).lPvfCode
                    ilRet = btrGetEqual(hmFrom, tmPvf, imRecLen, tlPvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    Do While (ilRet = BTRV_ERR_NONE)
                        DoEvents
                        If imTerminate Then
                            btrDestroy hmFrom
                            btrDestroy hmTo
                            mCopyPvfRecs = False
                            Exit Function
                        End If
                        ilRet = btrInsert(hmTo, tmPvf, imRecLen, INDEXKEY0)
                        If mBtrErrorMsg(ilRet, "mCopyPvfRecs: Insert") Then
                            btrDestroy hmFrom
                            btrDestroy hmTo
                            mCopyPvfRecs = False
                            Exit Function
                        End If
                        lmCount = lmCount + 1
                        If tmPvf.lLkPvfCode <= 0 Then
                            Exit Do
                        End If
                        imRecLen = imSvRecLen
                        tlPvfSrchKey.lCode = tmPvf.lLkPvfCode
                        ilRet = btrGetEqual(hmFrom, tmPvf, imRecLen, tlPvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
                Exit For
            End If
        Next ilVeh
    Next ilLoop
    Print #hmMsg, "    " & Str$(lmCount) & " Standard Package (Pvf.Btr) records Transferred"
    btrDestroy hmFrom
    btrDestroy hmTo
    mCopyPvfRecs = True
    Exit Function
mCopyPvfRecsErr:
    Screen.MousePointer = vbDefault
    btrDestroy hmFrom
    btrDestroy hmTo
    On Error GoTo 0
    mCopyPvfRecs = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopySdfRecs                   *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Fixed Length Files        *
'*                                                     *
'*******************************************************
Private Function mCopySdfRecs(slSpotDate As String) As Integer
    Dim ilRet As Integer
    Dim llNoRecords As Long
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Integer
    Dim ilLoop As Integer
    Dim ilInsert As Integer
    Dim llRecPos As Long
    Dim llSpotDate As Long
    Dim llTstDate As Long
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    lacProcessing.Caption = "Processing " & "Sdf.Btr"
    DoEvents
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopySdfRecs: Open") Then
        mCopySdfRecs = False
        Exit Function
    End If
    imRecLen = btrRecordLength(hmFrom)  'Get and save record length
    imSvRecLen = imRecLen
    llNoRecords = btrRecords(hmFrom)
    mRemoveFile "Sdf.Btr"
    ilRet = btrClone(hmFrom, smToDBPath & "Sdf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopySdfRecs: Clone Sdf") Then
        mCopySdfRecs = False
        Exit Function
    End If
    hmFSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopySdfRecs: Open 'From' Smf") Then
        mCopySdfRecs = False
        Exit Function
    End If
    mRemoveFile "Smf.Btr"
    ilRet = btrClone(hmFSmf, smToDBPath & "Smf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopySdfRecs: Clone Smf") Then
        mCopySdfRecs = False
        Exit Function
    End If
    hmFTzf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopySdfRecs: Open 'From' Tzf") Then
        mCopySdfRecs = False
        Exit Function
    End If
    mRemoveFile "Tzf.Btr"
    ilRet = btrClone(hmFTzf, smToDBPath & "Tzf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopySdfRecs: Clone Tzf") Then
        mCopySdfRecs = False
        Exit Function
    End If
    
    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopySdfRecs: Open 'To' Sdf") Then
        mCopySdfRecs = False
        Exit Function
    End If
    hmTSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTSmf, "", smToDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopySdfRecs: Open 'To' Smf") Then
        mCopySdfRecs = False
        Exit Function
    End If
    hmTTzf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTTzf, "", smToDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopySdfRecs: Open 'To' Tzf") Then
        mCopySdfRecs = False
        Exit Function
    End If
    llSpotDate = gDateValue(slSpotDate)
    lmCount = 0
    ilRet = btrGetFirst(hmFrom, tmSdf, imRecLen, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
    If (ilRet = BTRV_ERR_NONE) Then
        ilExtLen = imSvRecLen  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlVef) 'Obtain number of records
        btrExtClear hmFrom   'Clear any previous extend operation
        Call btrExtSetBounds(hmFrom, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)
        'gPackDate slSpotDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        'ilOffset = gFieldOffset("Sdf", "SdfDate")
        'ilRet = btrExtAddLogicConst(hmFrom, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        ilOffSet = 0
        ilRet = btrExtAddField(hmFrom, ilOffSet, ilExtLen)  'Extract First Name field
        ilRet = btrExtGetNext(hmFrom, tmSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmMsg, gGetBtrvErrorMsg(ilRet) & "mCopySdfRecs: ExtGetNext"
                btrDestroy hmFrom
                btrDestroy hmTo
                btrDestroy hmFSmf
                btrDestroy hmTSmf
                btrDestroy hmFTzf
                btrDestroy hmTTzf
                mCopySdfRecs = False
                Exit Function
            End If
            ilExtLen = imSvRecLen
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmFrom, tmSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                DoEvents
                If imTerminate Then
                    btrDestroy hmFrom
                    btrDestroy hmTo
                    btrDestroy hmFSmf
                    btrDestroy hmTSmf
                    btrDestroy hmFTzf
                    btrDestroy hmTTzf
                    mCopySdfRecs = False
                    Exit Function
                End If
                ilInsert = False
                For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                    If tmSdf.iVefCode = igPECSRNVefCode(ilLoop) Then
                        ilInsert = True
                        Exit For
                    End If
                Next ilLoop
                If ilInsert Then
                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llTstDate
                    If llTstDate < llSpotDate Then
                        ilInsert = False
                    End If
                End If
                If ilInsert Then
                    ilRet = mCopyOK(tmSdf.sPtType, tmSdf.lCopyCode)
                    If Not ilRet Then
                        tmSdf.sPtType = ""
                        tmSdf.lCopyCode = 0
                        tmSdf.iRotNo = 0
                    End If
                    If Not imIncludeCollection Then
                        tmSdf.sBill = "N"
                    End If
                    ilRet = btrInsert(hmTo, tmSdf, ilExtLen, INDEXKEY0)
                    If mBtrErrorMsg(ilRet, "mCopySdfRecs: Insert Sdf") Then
                        btrDestroy hmFrom
                        btrDestroy hmTo
                        btrDestroy hmFSmf
                        btrDestroy hmTSmf
                        btrDestroy hmFTzf
                        btrDestroy hmTTzf
                        mCopySdfRecs = False
                        Exit Function
                    End If
                    lmCount = lmCount + 1
                    If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                        tmLongSrchKey.lCode = tmSdf.lCode
                        ilRet = btrGetEqual(hmFSmf, tmSmf, Len(tmSmf), tmLongSrchKey, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            ilRet = mCopyOK(tmSmf.sPtType, tmSmf.lCopyCode)
                            If Not ilRet Then
                                tmSmf.sPtType = ""
                                tmSmf.lCopyCode = 0
                                tmSmf.iRotNo = 0
                            End If
                            ilRet = btrInsert(hmTSmf, tmSmf, Len(tmSmf), INDEXKEY0)
                            If mBtrErrorMsg(ilRet, "mCopySdfRecs: Insert Smf") Then
                                btrDestroy hmFrom
                                btrDestroy hmTo
                                btrDestroy hmFSmf
                                btrDestroy hmTSmf
                                btrDestroy hmFTzf
                                btrDestroy hmTTzf
                                mCopySdfRecs = False
                                Exit Function
                            End If
                        End If
                    End If
                End If
                ilExtLen = imSvRecLen
                ilRet = btrExtGetNext(hmFrom, tmSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmFrom, tmSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    Print #hmMsg, "    " & Str$(lmCount) & " Spots (sdf.btr, smf.btr, tzf.btr) records Transferred"
    btrDestroy hmFrom
    btrDestroy hmTo
    btrDestroy hmFSmf
    btrDestroy hmTSmf
    btrDestroy hmFTzf
    btrDestroy hmTTzf
    mCopySdfRecs = True
    Exit Function
mCopySdfRecsErr:
    Screen.MousePointer = vbDefault
    btrDestroy hmFrom
    btrDestroy hmTo
    btrDestroy hmFSmf
    btrDestroy hmTSmf
    btrDestroy hmFTzf
    btrDestroy hmTTzf
    On Error GoTo 0
    mCopySdfRecs = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopySmfRecs                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Fixed Length Files        *
'*                                                     *
'*******************************************************
Private Function mCopySmfRecs(slInDate As String) As Integer
    Dim ilRet As Integer
    Dim llNoRecords As Long
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Integer
    Dim ilLoop As Integer
    Dim llRecPos As Long
    Dim slDate As String
    Dim ilInsert As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    lacProcessing.Caption = "Processing " & "Smf.Btr"
    DoEvents
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopySmfRecs: Open") Then
        mCopySmfRecs = False
        Exit Function
    End If
    hmFSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopySmfRecs: Open 'From' Sdf") Then
        mCopySmfRecs = False
        Exit Function
    End If
    hmFTzf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopySmfRecs: Open 'From' Tzf") Then
        mCopySmfRecs = False
        Exit Function
    End If
    imRecLen = btrRecordLength(hmFrom)  'Get and save record length
    imSvRecLen = imRecLen
    llNoRecords = btrRecords(hmFrom)
    'The remove and Clone are done in mCopySdfRecs
    'mRemoveFile "Smf.Btr"
    'ilRet = btrClone(hmFrom, smToDBPath & "Smf.Btr", 0)
    'If ilRet <> BTRV_ERR_NONE Then
    '    Screen.MousePointer = vbDefault
    '    MsgBox "Unable to clone file- Error code " & Trim$(Str$(ilRet))
    '    mCopySmfRecs = False
    '    Exit Function
    'End If
    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopySmfRecs: Open To Smf") Then
        mCopySmfRecs = False
        Exit Function
    End If
    hmTSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTSdf, "", smToDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopySmfRecs: Open 'To' Sdf") Then
        mCopySmfRecs = False
        Exit Function
    End If
    hmTTzf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTTzf, "", smToDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopySmfRecs: Open 'To' Tzf") Then
        mCopySmfRecs = False
        Exit Function
    End If
    lmCount = 0
    ilRet = btrGetFirst(hmFrom, tmSmf, imRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If (ilRet = BTRV_ERR_NONE) Then
        ilExtLen = imSvRecLen  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlVef) 'Obtain number of records
        btrExtClear hmFrom   'Clear any previous extend operation
        Call btrExtSetBounds(hmFrom, llNoRec, -1, "UC", "SMF", "") 'Set extract limits (all records)
        gPackDate slInDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Smf", "SmfMissedDate")
        ilRet = btrExtAddLogicConst(hmFrom, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        ilOffSet = 0
        ilRet = btrExtAddField(hmFrom, ilOffSet, ilExtLen)  'Extract First Name field
        ilRet = btrExtGetNext(hmFrom, tmSmf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmMsg, gGetBtrvErrorMsg(ilRet) & "mCopySmfRecs: ExtGetNext"
                btrDestroy hmFrom
                btrDestroy hmFSdf
                btrDestroy hmFTzf
                btrDestroy hmTo
                btrDestroy hmTSdf
                btrDestroy hmTTzf
                mCopySmfRecs = False
                Exit Function
            End If
            ilExtLen = imSvRecLen
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmFrom, tmSmf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                DoEvents
                If imTerminate Then
                    btrDestroy hmFrom
                    btrDestroy hmFSdf
                    btrDestroy hmFTzf
                    btrDestroy hmTo
                    btrDestroy hmTSdf
                    btrDestroy hmTTzf
                    mCopySmfRecs = False
                    Exit Function
                End If
                tmLongSrchKey.lCode = tmSmf.lSdfCode
                ilRet = btrGetEqual(hmFSdf, tmSdf, Len(tmSdf), tmLongSrchKey, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    ilInsert = False
                    For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If tmSdf.iVefCode = igPECSRNVefCode(ilLoop) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilLoop
                    If ilInsert Then
                        tmLongSrchKey.lCode = tmSmf.lCode
                        ilRet = btrGetEqual(hmTo, tmRec, ilExtLen, tmLongSrchKey, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            ilRet = mCopyOK(tmSmf.sPtType, tmSmf.lCopyCode)
                            If Not ilRet Then
                                tmSmf.sPtType = ""
                                tmSmf.lCopyCode = 0
                                tmSmf.iRotNo = 0
                            End If
                            ilRet = btrInsert(hmTo, tmSmf, ilExtLen, INDEXKEY0)
                            If mBtrErrorMsg(ilRet, "mCopySmfRecs: Insert Smf") Then
                                btrDestroy hmFrom
                                btrDestroy hmFSdf
                                btrDestroy hmFTzf
                                btrDestroy hmTo
                                btrDestroy hmTSdf
                                btrDestroy hmTTzf
                                mCopySmfRecs = False
                                Exit Function
                            End If
                            lmCount = lmCount + 1
                        End If
                        tmLongSrchKey.lCode = tmSdf.lCode
                        ilRet = btrGetEqual(hmTSdf, tmSdf, Len(tmSdf), tmLongSrchKey, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            ilRet = mCopyOK(tmSdf.sPtType, tmSdf.lCopyCode)
                            If Not ilRet Then
                                tmSdf.sPtType = ""
                                tmSdf.lCopyCode = 0
                                tmSdf.iRotNo = 0
                            End If
                            If Not imIncludeCollection Then
                                tmSdf.sBill = "N"
                            End If
                            ilRet = btrInsert(hmTSdf, tmSdf, Len(tmSdf), INDEXKEY0)
                            If mBtrErrorMsg(ilRet, "mCopySmfRecs: Insert Sdf") Then
                                btrDestroy hmFrom
                                btrDestroy hmFSdf
                                btrDestroy hmFTzf
                                btrDestroy hmTo
                                btrDestroy hmTSdf
                                btrDestroy hmTTzf
                                mCopySmfRecs = False
                                Exit Function
                            End If
                        End If
                    End If
                End If
                ilExtLen = imSvRecLen
                ilRet = btrExtGetNext(hmFrom, tmSmf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmFrom, tmSmf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    Print #hmMsg, "    " & Str$(lmCount) & " Extra Spots (smf.btr, sdf.btr, tzf.btr) records Transferred"
    btrDestroy hmFrom
    btrDestroy hmFSdf
    btrDestroy hmFTzf
    btrDestroy hmTo
    btrDestroy hmTSdf
    btrDestroy hmTTzf
    mCopySmfRecs = True
    Exit Function
mCopySmfRecsErr:
    Screen.MousePointer = vbDefault
    btrDestroy hmFrom
    btrDestroy hmFSdf
    btrDestroy hmFTzf
    btrDestroy hmTo
    btrDestroy hmTSdf
    btrDestroy hmTTzf
    On Error GoTo 0
    mCopySmfRecs = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopySsfRecs                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Variable length files     *
'*                                                     *
'*******************************************************
Private Function mCopySsfRecs(slSpotDate As String) As Integer
    Dim ilRet As Integer
    Dim llNoRecords As Long
    Dim llRecPos As Long
    Dim llDate As Long
    Dim llInDate As Long
    Dim ilType As Integer
    Dim llTstDate As Long
    Dim ilLoop As Integer
    Dim ilInsert As Integer
    Dim slErrDate As String
    Dim tlSsfSrchKey2 As SSFKEY2     'SSF key record image
    Dim ilSsfRecLen As Integer      'SSF record length
    lacProcessing.Caption = "Processing " & "Ssf.Btr"
    DoEvents
    llInDate = gDateValue(slSpotDate)
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopySsfRecs: Open") Then
        mCopySsfRecs = False
        Exit Function
    End If
    imRecLen = Len(tgSsf(0))    'btrRecordLength(hmFrom)  'Get and save record length
    ilType = 0
    ReDim lgSSFAirStartDate(0 To UBound(igPEAVefCode)) As Long
    For ilLoop = 0 To UBound(igPEAVefCode) - 1 Step 1
        tlSsfSrchKey2.iVefCode = igPEAVefCode(ilLoop)
        tlSsfSrchKey2.iDate(0) = 257
        tlSsfSrchKey2.iDate(1) = 2100
        ilSsfRecLen = Len(tgSsf(0))
        'ilRet = btrGetLessOrEqual(hmFrom, tgSsf(0), ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        ilRet = gSSFGetLessOrEqualKey2(hmFrom, tgSsf(0), ilSsfRecLen, tlSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
        If (ilRet = BTRV_ERR_NONE) And (tgSsf(0).iVefCode = igPEAVefCode(ilLoop)) Then
            gUnpackDateLong tgSsf(0).iDate(0), tgSsf(0).iDate(1), llTstDate
            If llTstDate <= llInDate Then
                Do Until gWeekDayLong(llTstDate) = 0
                    llTstDate = llTstDate - 1
                Loop
                lgSSFAirStartDate(ilLoop) = llTstDate
            Else
                lgSSFAirStartDate(ilLoop) = llInDate
            End If
        Else
            lgSSFAirStartDate(ilLoop) = llInDate
        End If
    Next ilLoop
    imSvRecLen = imRecLen
    llNoRecords = btrRecords(hmFrom)
    mRemoveFile "Ssf.Btr"
    ilRet = btrClone(hmFrom, smToDBPath & "Ssf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopySsfRecs: Clone") Then
        mCopySsfRecs = False
        Exit Function
    End If
    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopySsfRecs: Open 'To' Ssf") Then
        mCopySsfRecs = False
        Exit Function
    End If
    lmCount = 0
    'ilRet = btrGetFirst(hmFrom, tgSsf(0), imRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    ilRet = gSSFGetFirst(hmFrom, tgSsf(0), imRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE)
        DoEvents
        If imTerminate Then
            btrDestroy hmFrom
            btrDestroy hmTo
            mCopySsfRecs = False
            Exit Function
        End If
        ilInsert = False
        For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
            If tgSsf(0).iVefCode = igPECSRNVefCode(ilLoop) Then
                gUnpackDateLong tgSsf(0).iDate(0), tgSsf(0).iDate(1), llTstDate
                If llTstDate >= llInDate Then
                    ilInsert = True
                End If
                Exit For
            End If
        Next ilLoop
        If Not ilInsert Then
            For ilLoop = 0 To UBound(igPEAVefCode) - 1 Step 1
                If tgSsf(0).iVefCode = igPEAVefCode(ilLoop) Then
                    gUnpackDateLong tgSsf(0).iDate(0), tgSsf(0).iDate(1), llTstDate
                    If llTstDate >= lgSSFAirStartDate(ilLoop) Then
                        ilInsert = True
                    End If
                    Exit For
                End If
            Next ilLoop
        End If
        If ilInsert Then
            ilRet = btrInsert(hmTo, tgSsf(0), imRecLen, INDEXKEY0)
            If mBtrErrorMsg(ilRet, "mCopySsfRecs: Insert") Then
                btrDestroy hmFrom
                btrDestroy hmTo
                mCopySsfRecs = False
                Exit Function
            End If
            lmCount = lmCount + 1
        End If
        imRecLen = imSvRecLen
        'ilRet = btrGetNext(hmFrom, tgSsf(0), imRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilRet = gSSFGetNext(hmFrom, tgSsf(0), imRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Do While (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_NONE)
            gUnpackDate tgSsf(0).iDate(0), tgSsf(0).iDate(1), slErrDate
            For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                If tgSsf(0).iVefCode = tgMVef(ilLoop).iCode Then
                    Print #hmMsg, "    Error" & Str$(ilRet) & " in SSF " & Trim$(tgMVef(ilLoop).sName) & " " & slErrDate
                    Exit For
                End If
            Next ilLoop
            imRecLen = imSvRecLen
            'ilRet = btrGetNext(hmFrom, tgSsf(0), imRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            ilRet = gSSFGetNext(hmFrom, tgSsf(0), imRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Loop
    Print #hmMsg, "    " & Str$(lmCount) & " Spot Summary (ssf.btr) records Transferred"
    btrDestroy hmFrom
    btrDestroy hmTo
    mCopySsfRecs = True
    Exit Function
mCopySsfRecsErr:
    Screen.MousePointer = vbDefault
    btrDestroy hmFrom
    btrDestroy hmTo
    On Error GoTo 0
    mCopySsfRecs = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyVcfRecs                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Variable length files     *
'*                                                     *
'*******************************************************
Private Function mCopyVcfRecs(slSpotDate As String) As Integer
    Dim ilRet As Integer
    Dim llNoRecords As Long
    Dim llRecPos As Long
    Dim llDate As Long
    Dim llInDate As Long
    Dim llTDate As Long
    Dim ilLoop As Integer
    Dim ilInsert As Integer
    lacProcessing.Caption = "Processing " & "Vcf.Btr"
    DoEvents
    llInDate = gDateValue(slSpotDate)
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & "Vcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyVcfRecs: Open") Then
        mCopyVcfRecs = False
        Exit Function
    End If
    imRecLen = btrRecordLength(hmFrom)  'Get and save record length
    imSvRecLen = imRecLen
    llNoRecords = btrRecords(hmFrom)
    mRemoveFile "Vcf.Btr"
    ilRet = btrClone(hmFrom, smToDBPath & "Vcf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyVcfRecs: Clone") Then
        mCopyVcfRecs = False
        Exit Function
    End If
    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & "Vcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyVcfRecs: Open 'To' Vcf") Then
        mCopyVcfRecs = False
        Exit Function
    End If
    lmCount = 0
    ilRet = btrGetFirst(hmFrom, tmVcf, imRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE)
        DoEvents
        If imTerminate Then
            btrDestroy hmFrom
            btrDestroy hmTo
            mCopyVcfRecs = False
            Exit Function
        End If
        'gUnpackDateLong tmVcf.iEffDate(0), tmVcf.iEffDate(1), llDate
        'gUnpackDateLong tmVcf.iTermDate(0), tmVcf.iTermDate(1), llTDate
        'If ((llDate >= llSDate) And (llDate <= llInDate)) Or ((llTDate >= llSDate) And (llTDate <= llInDate)) Or ((llTDate = 0)) Then
        '    ilRet = btrInsert(hmTo, tmVcf, imRecLen, INDEXKEY0)
        'End If
        For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
            If tmVcf.iSellCode = igPECSRNVefCode(ilLoop) Then
                ilInsert = True
            ElseIf tmVcf.iCSV(1) = igPECSRNVefCode(ilLoop) Then
                ilInsert = True
            ElseIf tmVcf.iCSV(2) = igPECSRNVefCode(ilLoop) Then
                ilInsert = True
            ElseIf tmVcf.iCSV(3) = igPECSRNVefCode(ilLoop) Then
                ilInsert = True
            ElseIf tmVcf.iCSV(4) = igPECSRNVefCode(ilLoop) Then
                ilInsert = True
            ElseIf tmVcf.iCSV(5) = igPECSRNVefCode(ilLoop) Then
                ilInsert = True
            End If
            If ilInsert Then
                Exit For
            End If
        Next ilLoop
        imRecLen = imSvRecLen
        gUnpackDateLong tmVcf.iEffDate(0), tmVcf.iEffDate(1), llDate
        If ilInsert Then
            ilRet = btrInsert(hmTo, tmVcf, imRecLen, INDEXKEY0)
            If mBtrErrorMsg(ilRet, "mCopyVcfRecs: Insert") Then
                mCopyVcfRecs = False
                btrDestroy hmFrom
                btrDestroy hmTo
                Exit Function
            End If
            lmCount = lmCount + 1
        End If
        ilRet = btrGetNext(hmFrom, tmVcf, imRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Print #hmMsg, "    " & Str$(lmCount) & " Vehicle Conflict (vcf.btr) records Transferred"
    btrDestroy hmFrom
    btrDestroy hmTo
    mCopyVcfRecs = True
    Exit Function
mCopyVcfRecsErr:
    Screen.MousePointer = vbDefault
    btrDestroy hmFrom
    btrDestroy hmTo
    On Error GoTo 0
    mCopyVcfRecs = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyVlfRecs                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Variable length files     *
'*                                                     *
'*******************************************************
Private Function mCopyVlfRecs(slSpotDate As String) As Integer
    Dim ilRet As Integer
    Dim llNoRecords As Long
    Dim llRecPos As Long
    Dim llDate As Long
    Dim llInDate As Long
    Dim llTDate As Long
    Dim ilLoop As Integer
    Dim ilInsert As Integer
    lacProcessing.Caption = "Processing " & "Vlf.Btr"
    DoEvents
    llInDate = gDateValue(slSpotDate)
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyVlfRecs: Open") Then
        mCopyVlfRecs = False
        Exit Function
    End If
    imRecLen = btrRecordLength(hmFrom)  'Get and save record length
    imSvRecLen = imRecLen
    llNoRecords = btrRecords(hmFrom)
    mRemoveFile "Vlf.Btr"
    ilRet = btrClone(hmFrom, smToDBPath & "Vlf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyVlfRecs: Clone") Then
        mCopyVlfRecs = False
        Exit Function
    End If
    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyVlfRecs: Open 'To' Vlf") Then
        mCopyVlfRecs = False
        Exit Function
    End If
    lmCount = 0
    ilRet = btrGetFirst(hmFrom, tmVlf, imRecLen, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE)
        DoEvents
        If imTerminate Then
            btrDestroy hmFrom
            btrDestroy hmTo
            mCopyVlfRecs = False
            Exit Function
        End If
        'gUnpackDateLong tmVlf.iEffDate(0), tmVlf.iEffDate(1), llDate
        'gUnpackDateLong tmVlf.iTermDate(0), tmVlf.iTermDate(1), llTDate
        'If ((llDate >= llSDate) And (llDate <= llInDate)) Or ((llTDate >= llSDate) And (llTDate <= llInDate)) Or ((llTDate = 0)) Then
        '    ilRet = btrInsert(hmTo, tmVlf, imRecLen, INDEXKEY0)
        'End If
        ilInsert = False
        For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
            If tmVlf.iSellCode = igPECSRNVefCode(ilLoop) Then
                ilInsert = True
                Exit For
            End If
        Next ilLoop
        If Not ilInsert Then
            For ilLoop = 0 To UBound(igPEAVefCode) - 1 Step 1
                If tmVlf.iAirCode = igPEAVefCode(ilLoop) Then
                    ilInsert = True
                    Exit For
                End If
            Next ilLoop
        End If
        If ilInsert Then
            ilRet = btrInsert(hmTo, tmVlf, imRecLen, INDEXKEY0)
            If mBtrErrorMsg(ilRet, "mCopyVlfRecs: Insert") Then
                btrDestroy hmFrom
                btrDestroy hmTo
                mCopyVlfRecs = False
                Exit Function
            End If
            lmCount = lmCount + 1
        End If
        imRecLen = imSvRecLen
        ilRet = btrGetNext(hmFrom, tmVlf, imRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Print #hmMsg, "    " & Str$(lmCount) & " Vehicle Links (vlf.btr) records Transferred"
    btrDestroy hmFrom
    btrDestroy hmTo
    mCopyVlfRecs = True
    Exit Function
mCopyVlfRecsErr:
    Screen.MousePointer = vbDefault
    btrDestroy hmFrom
    btrDestroy hmTo
    On Error GoTo 0
    mCopyVlfRecs = False
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
    Dim llNoRec As Long 'Number of records
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilTest As Integer
    Dim ilVefCode As Integer
    Dim slMnfType As String
    Dim slMMnfStamp As String
    Dim slFYear As String
    Dim slFMonth As String
    Dim slFDay As String
    Dim slDate As String
    imTerminate = False
    imExit = False
    imRun = False
    mParseCmmdLine
    
    If Not gCheckDDFDates() Then
        imTerminate = True
        Exit Sub
    End If
    
    If Not mCheckRecLength() Then
        imTerminate = True
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imGenerating = False
    imFirstFocus = True
    imBypassFocus = False
    
    imBSMode = False
    imCalType = 0   'Standard
    imDateBox = 0
    mInitBox
    slStr = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(slStr)
    lmPrevStdStartDate = gDateValue(gObtainStartStd(Format$(gDateValue(gObtainStartStd(Format$(lmNowDate, "m/d/yy"))) - 1, "m/d/yy")))
    slStr = gObtainYearStartDate(0, slStr)
    edcCntrDate.Text = slStr
    slStr = Format$(gNow(), "m/d/yy")
    slStr = gObtainStartStd(slStr)
    edcSpotDate.Text = slStr
    lacFromDB.Caption = "Data From: " & Left$(sgDBPath, Len(sgDBPath) - 1)
    lacToDB.Caption = "Data To: " & Left$(smToDBPath, Len(smToDBPath) - 1)
    slDate = Format$(gNow(), "mm/dd/yyyy")
    gObtainYearMonthDayStr slDate, True, slFYear, slFMonth, slFDay
    smZipName = "PE" & right$(slFYear, 2) & slFMonth & slFDay & ".csi"
    lacToZip.Caption = "Result File: " & smToDBPath & "PE\" & smZipName
    'plcScreen.Caption = "Participant Extraction for " & smPartExtName
    slMnfType = "H1"    'Participant
    slMMnfStamp = ""
    ilRet = gObtainMnfForType(slMnfType, slMMnfStamp, tgPEMnf())
    For ilLoop = LBound(tgPEMnf) To UBound(tgPEMnf) - 1 Step 1
        lbcPart.AddItem Trim$(tgPEMnf(ilLoop).sName)
    Next ilLoop
    For ilLoop = 0 To lbcPart.ListCount - 1 Step 1
        If StrComp(Trim$(lbcPart.List(ilLoop)), smPartExtName, 1) = 0 Then
            lbcPart.ListIndex = ilLoop
            Exit For
        End If
    Next ilLoop
    rbcUsersOff(1).Value = True
    If imIncludeCollection Then
        ckcCollection.Value = vbChecked
    Else
        ckcCollection.Value = vbUnchecked
    End If
    If imIncludeCopy Then
        ckcCopy.Value = vbChecked
    Else
        ckcCopy.Value = vbUnchecked
    End If
    If imIncludeResearch Then
        ckcResearch.Value = vbChecked
    Else
        ckcResearch.Value = vbUnchecked
    End If
    'gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    'pbcCalendar_Paint   'mBoxCalDate called within paint
    'lacDate.Visible = False
    gCenterStdAlone ExptPE
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    Exit Sub
mInitErr:
    Screen.MousePointer = vbDefault
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
    Dim ilLoop As Integer
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    plcCalendar.Move edcCntrDate.Left, edcCntrDate.Top + edcCntrDate.Height
End Sub
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
    slToFile = sgExportPath & "ExptPE.Txt"
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
                MsgBox "Open " & slToFile, vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
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
                MsgBox "Open " & slToFile, vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
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
            MsgBox "Open " & slToFile, vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, "** Generating Producer Extraction Database: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    Print #hmMsg, ""
    mOpenMsgFile = True
    Exit Function
mOpenMsgFileErr:
    ilRet = 1
    Resume Next
End Function
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
    Dim ilPos As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slHelpSystem As String
    Dim slToDBName As String
    Dim slPEName As String
    Dim slChar As String
    slCommand = Command$
    slToDBName = "Extract"
    'imIncludeCollection = False
    imIncludeCollection = False
    imIncludeCopy = False
    imIncludeResearch = False
    ilPos = InStr(1, slCommand, "/CS", 1)
    If (StrComp(slCommand, "Debug", 1) = 0) Or (ilPos > 0) Then
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
        ilPos = InStr(1, slCommand, "/E:", 1)
        If ilPos > 0 Then
            slToDBName = ""
            ilPos = ilPos + 3
            slChar = Trim$(Mid$(slCommand, ilPos, 1))
            Do While slChar = ""
                ilPos = ilPos + 1
                slChar = Mid$(slCommand, ilPos, 1)
            Loop
            Do While slChar <> ""
                slToDBName = slToDBName & slChar
                ilPos = ilPos + 1
                If ilPos > Len(slCommand) Then
                    Exit Do
                End If
                slChar = Mid$(slCommand, ilPos, 1)
            Loop
            If right$(slToDBName, 1) = "\" Then
                slToDBName = Left$(slToDBName, Len(slToDBName) - 1)
            End If
        End If
        'Get participant Extraction Name
        ilPos = InStr(1, slCommand, "/PEN:", 1)
        If ilPos > 0 Then
            slPEName = ""
            ilPos = ilPos + 5
            slChar = Trim$(Mid$(slCommand, ilPos, 1))
            Do While slChar = ""
                ilPos = ilPos + 1
                slChar = Mid$(slCommand, ilPos, 1)
            Loop
            Do While slChar <> "/"
                slPEName = slPEName & slChar
                ilPos = ilPos + 1
                If ilPos > Len(slCommand) Then
                    Exit Do
                End If
                slChar = Mid$(slCommand, ilPos, 1)
            Loop
            smPartExtName = Trim$(slPEName)
        End If
        ilPos = InStr(1, slCommand, "/Collection", 1)
        If (ilPos > 0) Then
            imIncludeCollection = True
        End If
        ilPos = InStr(1, slCommand, "/Copy", 1)
        If (ilPos > 0) Then
            imIncludeCopy = True
        End If
        ilPos = InStr(1, slCommand, "/Research", 1)
        If (ilPos > 0) Then
            imIncludeResearch = True
        End If
        ilPos = InStr(1, slCommand, "/Exit", 1)
        If ilPos > 0 Then
            imExit = True
        End If
        ilPos = InStr(1, slCommand, "/Run", 1)
        If (ilPos > 0) And (Len(smPartExtName) <> 0) Then
            imRun = True
        End If
        slCommand = "Traffic\Guide"
    Else
        igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            End
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
        imShowHelpMsg = True
        ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
        If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
            imShowHelpMsg = False
        End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    End If
    gInitStdAlone ExptPE, slStr, ilTestSystem
    smToDBPath = sgDBPath & slToDBName & "\"
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mRemoveFile                     *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Sub mRemoveFile(slFileName As String)
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    ilRet = 0
    On Error GoTo mRemoveFileErr:
    slToFile = smToDBPath & slFileName
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        Kill slToFile
    End If
    On Error GoTo 0
    Exit Sub
mRemoveFileErr:
    ilRet = 1
    Resume Next
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
    Erase igPECSRNVefCode
    Erase igPEAVefCode
    Erase igPEPVefCode
    Erase igPELVefCode
    Erase tgPEMnf
    Erase lgChfCode
    Erase lgLCFAirStartDate
    Erase lgSSFAirStartDate
    Erase igPELtfCode
    Erase lgPELvfCode
    Erase tgPkageLnCheck
    Erase igAddPkVefCode
    Erase igPECntrAdfCode
    Erase igPECntrAgfCode
    Erase igPECntrSlfCode
    Screen.MousePointer = vbDefault
    igParentRestarted = False
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
    Unload ExptPE
    Set ExptPE = Nothing   'Remove data segment
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyOK                         *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Time zone copy                 *
'*                                                     *
'*******************************************************
Private Function mCopyOK(slPtType As String, llCopyCode As Long) As Integer
    Dim ilTzfRecLen As Integer
    Dim ilRet As Integer
    Dim ilTzf As Integer
    Dim llRet As Long
    Dim tlTzfSrchKey As LONGKEY0
    If Not imIncludeCopy Then
        mCopyOK = False
        Exit Function
    End If
    mCopyOK = True
    If slPtType = "1" Then  '  Single Copy
        llRet = mSetCopyUsed(llCopyCode)
    ElseIf slPtType = "2" Then  '  Combo Copy
    ElseIf slPtType = "3" Then  '  Time Zone Copy
        ' Read TZF using lCopyCode from SDF
        ilTzfRecLen = Len(tmTzf)
        tlTzfSrchKey.lCode = llCopyCode
        ilRet = btrGetEqual(hmFTzf, tmTzf, ilTzfRecLen, tlTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrInsert(hmTTzf, tmTzf, ilTzfRecLen, INDEXKEY0)
            If mBtrErrorMsg(ilRet, "mCopyOK: Insert") Then
                mCopyOK = False
                Exit Function
            End If
            For ilTzf = 1 To 6 Step 1
                If tmTzf.lCifZone(ilTzf) > 0 Then ' Process just the first positive Zone
                    llRet = mSetCopyUsed(tmTzf.lCifZone(ilTzf))
                End If
            Next ilTzf
        End If
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mUpdateVsf                      *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Process VSF record from a      *
'*                      Contract                       *
'*                                                     *
'*******************************************************
Private Function mUpdateVsf(tlChf As CHF) As Integer
    Dim ilVsfReclen As Integer
    Dim ilRetain As Integer
    Dim ilTest As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim tlVsf As VSF
    ilVsfReclen = Len(tmVsf)
    If tlChf.lVefCode > 0 Then
        mUpdateVsf = True
        Exit Function
    ElseIf tlChf.lVefCode < 0 Then
        tmLongSrchKey.lCode = -tlChf.lVefCode
        ilRet = btrGetEqual(hmFVsf, tmVsf, ilVsfReclen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Do While ilRet = BTRV_ERR_NONE
            tmLongSrchKey.lCode = tmVsf.lCode
            ilRet = btrGetEqual(hmTVsf, tlVsf, ilVsfReclen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                Exit Do     'Vsf aready moved to extracted database
            End If
            For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                If tmVsf.iFSCode(ilLoop) > 0 Then
                    ilRetain = False
                    For ilTest = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If (tmVsf.iFSCode(ilLoop) = igPECSRNVefCode(ilTest)) Then
                            ilRetain = True
                            Exit For
                        End If
                    Next ilTest
                    If Not ilRetain Then
                        tmVsf.iFSCode(ilLoop) = 0
                    End If
                End If
            Next ilLoop
            ilRet = btrInsert(hmTVsf, tmVsf, ilVsfReclen, INDEXKEY0)
            If mBtrErrorMsg(ilRet, "mUpdateVsf: Insert") Then
                mUpdateVsf = False
                Exit Function
            End If
            If tmVsf.lLkVsfCode <= 0 Then
                Exit Do
            End If
            tmLongSrchKey.lCode = tmVsf.lLkVsfCode
            ilRet = btrGetEqual(hmFVsf, tmVsf, ilVsfReclen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    mUpdateVsf = True
    Exit Function
End Function
Private Function mZipFiles() As Integer
    Dim slDateTime As String
    Dim ilRet As Integer
    initZIPCmdStruct dzactxctrl1
    On Error GoTo mZipFilesErr:
    ilRet = 0
    slDateTime = FileDateTime(smToDBPath & "PE\" & smZipName)
    If ilRet = 0 Then
        Kill smToDBPath & "PE\" & smZipName
    End If
    On Error GoTo 0
    dzactxctrl1.ZIPFile = smToDBPath & "PE\" & smZipName
    dzactxctrl1.ItemList = smToDBPath & "*.btr"
    dzactxctrl1.ActionDZ = ZIP_ADD
    ilRet = dzactxctrl1.ErrorCode
    Exit Function
mZipFilesErr:
    ilRet = 1
    Resume Next
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
        slDay = Trim$(Str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                If imDateBox = 0 Then
                    edcCntrDate.Text = Format$(llDate, "m/d/yy")
                    edcCntrDate.SelStart = 0
                    edcCntrDate.SelLength = Len(edcCntrDate.Text)
                    imBypassFocus = True
                    edcCntrDate.SetFocus
                Else
                    edcSpotDate.Text = Format$(llDate, "m/d/yy")
                    edcSpotDate.SelStart = 0
                    edcSpotDate.SelLength = Len(edcSpotDate.Text)
                    imBypassFocus = True
                    edcSpotDate.SetFocus
                End If
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    If imDateBox = 0 Then
        edcCntrDate.SetFocus
    Else
        edcSpotDate.SetFocus
    End If
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    slStr = Trim$(Str$(imCalMonth)) & "/15/" & Trim$(Str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub rbcUsersOff_GotFocus(Index As Integer)
    plcCalendar.Visible = False
End Sub
Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    cmcGen_Click
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Participant Extraction"
End Sub
Private Sub plcSUseProd_Paint()
    plcSUseProd.CurrentX = 0
    plcSUseProd.CurrentY = 0
    plcSUseProd.Print "Are all users off the system"
End Sub

Private Function mGetChfAdfAgfSlf(slCntrDate As String) As Integer
    Dim ilRet As Integer
    Dim llNoRecords As Long
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Integer
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim llLoop As Long
    Dim llRecPos As Long
    Dim ilInsert As Integer
    Dim ilVsfReclen As Integer
    Dim ilPass As Integer
    Dim ilSlf As Integer
    Dim llDate As Long
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    ReDim lgChfCode(0 To 0) As Long
    lacProcessing.Caption = "Retrieving " & "Contracts"
    DoEvents
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mGetChfAdfAgfSlf: Open From Chf") Then
        mGetChfAdfAgfSlf = False
        Exit Function
    End If
    imRecLen = btrRecordLength(hmFrom)  'Get and save record length
    imSvRecLen = imRecLen
    
    hmFVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mGetChfAdfAgfSlf: Open 'From' Vsf") Then
        mGetChfAdfAgfSlf = False
        Exit Function
    End If
    ilVsfReclen = Len(tmVsf)
    
    llNoRecords = btrRecords(hmFrom)
    ilRet = btrGetFirst(hmFrom, tmChf, imRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If (ilRet = BTRV_ERR_NONE) Then
        ilExtLen = imSvRecLen  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlVef) 'Obtain number of records
        btrExtClear hmFrom   'Clear any previous extend operation
        Call btrExtSetBounds(hmFrom, llNoRec, -1, "UC", "CHF", "") 'Set extract limits (all records)
        gPackDate slCntrDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Chf", "ChfEndDate")
        ilRet = btrExtAddLogicConst(hmFrom, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        ilOffSet = 0
        ilRet = btrExtAddField(hmFrom, ilOffSet, ilExtLen)  'Extract First Name field
        ilRet = btrExtGetNext(hmFrom, tmChf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmMsg, gGetBtrvErrorMsg(ilRet) & "mGetChfAdfAgfSlf: ExtGetNext"
                btrDestroy hmFrom
                btrDestroy hmFVsf
                mGetChfAdfAgfSlf = False
                Exit Function
            End If
            ilExtLen = imSvRecLen
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmFrom, tmChf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                DoEvents
                If imTerminate Then
                    btrDestroy hmFrom
                    btrDestroy hmFVsf
                    mGetChfAdfAgfSlf = False
                    Exit Function
                End If
                ilInsert = False
                If tmChf.lVefCode > 0 Then
                    For ilTest = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                        If (tmChf.lVefCode = CLng(igPECSRNVefCode(ilTest))) Then
                            ilInsert = True
                            Exit For
                        End If
                    Next ilTest
                ElseIf tmChf.lVefCode < 0 Then
                    tmLongSrchKey.lCode = -tmChf.lVefCode
                    ilRet = btrGetEqual(hmFVsf, tmVsf, ilVsfReclen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    Do While ilRet = BTRV_ERR_NONE
                        For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                            If tmVsf.iFSCode(ilLoop) > 0 Then
                                For ilTest = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                                    If (tmVsf.iFSCode(ilLoop) = igPECSRNVefCode(ilTest)) Then
                                        ilInsert = True
                                        Exit For
                                    End If
                                Next ilTest
                            End If
                            If ilInsert Then
                                Exit Do
                            End If
                        Next ilLoop
                        If tmVsf.lLkVsfCode <= 0 Then
                            Exit Do
                        End If
                        tmLongSrchKey.lCode = tmVsf.lLkVsfCode
                        ilRet = btrGetEqual(hmFVsf, tmVsf, ilVsfReclen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
                'Bypass old proposals
                If ilInsert Then
                    If (tmChf.sStatus <> "H") And (tmChf.sStatus <> "O") Then
                        gUnpackDateLong tmChf.iEndDate(0), tmChf.iEndDate(1), llDate
                        If llDate < lmPrevStdStartDate Then
                            ilInsert = False
                        End If
                    End If
                End If
                If (ilInsert) And (ckcContractHistory.Value = vbUnchecked) Then
                    If tmChf.sDelete = "Y" Then
                        ilInsert = False
                    End If
                End If
                If ilInsert Then
                    lgChfCode(UBound(lgChfCode)) = tmChf.lCode
                    ReDim Preserve lgChfCode(0 To UBound(lgChfCode) + 1) As Long
                    ilRet = mBinarySearch(igPECntrAdfCode(), tmChf.iAdfCode)
                    If ilRet = -1 Then
                        igPECntrAdfCode(UBound(igPECntrAdfCode)) = tmChf.iAdfCode
                        ReDim Preserve igPECntrAdfCode(0 To UBound(igPECntrAdfCode) + 1) As Integer
                        If UBound(igPECntrAdfCode) > 1 Then
                            ArraySortTyp fnAV(igPECntrAdfCode(), 0), UBound(igPECntrAdfCode), 0, LenB(igPECntrAdfCode(0)), 0, -1, 0
                        End If
                    End If
                    If tmChf.iAgfCode > 0 Then
                        ilRet = mBinarySearch(igPECntrAgfCode(), tmChf.iAgfCode)
                        If ilRet = -1 Then
                            igPECntrAgfCode(UBound(igPECntrAgfCode)) = tmChf.iAgfCode
                            ReDim Preserve igPECntrAgfCode(0 To UBound(igPECntrAgfCode) + 1) As Integer
                            If UBound(igPECntrAgfCode) > 1 Then
                                ArraySortTyp fnAV(igPECntrAgfCode(), 0), UBound(igPECntrAgfCode), 0, LenB(igPECntrAgfCode(0)), 0, -1, 0
                            End If
                        End If
                    End If
                    For ilSlf = LBound(tmChf.iSlfCode) To UBound(tmChf.iSlfCode) Step 1
                        If tmChf.iSlfCode(ilSlf) > 0 Then
                            ilRet = mBinarySearch(igPECntrSlfCode(), tmChf.iSlfCode(ilSlf))
                            If ilRet = -1 Then
                                igPECntrSlfCode(UBound(igPECntrSlfCode)) = tmChf.iSlfCode(ilSlf)
                                ReDim Preserve igPECntrSlfCode(0 To UBound(igPECntrSlfCode) + 1) As Integer
                                If UBound(igPECntrSlfCode) > 1 Then
                                    ArraySortTyp fnAV(igPECntrSlfCode(), 0), UBound(igPECntrSlfCode), 0, LenB(igPECntrSlfCode(0)), 0, -1, 0
                                End If
                            End If
                        End If
                    Next ilSlf
                End If
                ilExtLen = imSvRecLen
                ilRet = btrExtGetNext(hmFrom, tmChf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmFrom, tmChf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    btrDestroy hmFrom
    btrDestroy hmFVsf
    
    For ilPass = 0 To 1 Step 1
        hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
        If ilPass = 0 Then
            ilRet = btrOpen(hmFrom, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If mBtrErrorMsg(ilRet, "mGetChfAdfAgfSlf: Open From Rvf") Then
                mGetChfAdfAgfSlf = False
                Exit Function
            End If
        Else
            ilRet = btrOpen(hmFrom, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If mBtrErrorMsg(ilRet, "mGetChfAdfAgfSlf: Open From Phf") Then
                mGetChfAdfAgfSlf = False
                Exit Function
            End If
        End If
        imRecLen = Len(tmRvf)
        ilRet = btrGetFirst(hmFrom, tmRvf, imRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Do While (ilRet = BTRV_ERR_NONE)
            ilInsert = False
            For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
                If tmRvf.iAirVefCode = igPECSRNVefCode(ilLoop) Then
                    ilInsert = True
                    Exit For
                End If
            Next ilLoop
            If Not ilInsert Then
                For ilLoop = 0 To UBound(igPEAVefCode) - 1 Step 1
                    If tmRvf.iAirVefCode = igPEAVefCode(ilLoop) Then
                        ilInsert = True
                        Exit For
                    End If
                Next ilLoop
            End If
            If Not ilInsert Then
                For ilLoop = 0 To UBound(igPELVefCode) - 1 Step 1
                    If tmRvf.iAirVefCode = igPELVefCode(ilLoop) Then
                        ilInsert = True
                        Exit For
                    End If
                Next ilLoop
            End If
            If ilInsert Then
                ilRet = mBinarySearch(igPECntrAdfCode(), tmRvf.iAdfCode)
                If ilRet = -1 Then
                    igPECntrAdfCode(UBound(igPECntrAdfCode)) = tmRvf.iAdfCode
                    ReDim Preserve igPECntrAdfCode(0 To UBound(igPECntrAdfCode) + 1) As Integer
                    If UBound(igPECntrAdfCode) > 1 Then
                        ArraySortTyp fnAV(igPECntrAdfCode(), 0), UBound(igPECntrAdfCode), 0, LenB(igPECntrAdfCode(0)), 0, -1, 0
                    End If
                End If
                If tmRvf.iAgfCode > 0 Then
                    ilRet = mBinarySearch(igPECntrAgfCode(), tmRvf.iAgfCode)
                    If ilRet = -1 Then
                        igPECntrAgfCode(UBound(igPECntrAgfCode)) = tmRvf.iAgfCode
                        ReDim Preserve igPECntrAgfCode(0 To UBound(igPECntrAgfCode) + 1) As Integer
                        If UBound(igPECntrAgfCode) > 1 Then
                            ArraySortTyp fnAV(igPECntrAgfCode(), 0), UBound(igPECntrAgfCode), 0, LenB(igPECntrAgfCode(0)), 0, -1, 0
                        End If
                    End If
                End If
                If tmRvf.iSlfCode > 0 Then
                    ilRet = mBinarySearch(igPECntrSlfCode(), tmRvf.iSlfCode)
                    If ilRet = -1 Then
                        igPECntrSlfCode(UBound(igPECntrSlfCode)) = tmRvf.iSlfCode
                        ReDim Preserve igPECntrSlfCode(0 To UBound(igPECntrSlfCode) + 1) As Integer
                        If UBound(igPECntrSlfCode) > 1 Then
                            ArraySortTyp fnAV(igPECntrSlfCode(), 0), UBound(igPECntrSlfCode), 0, LenB(igPECntrSlfCode(0)), 0, -1, 0
                        End If
                    End If
                End If
            End If
            ilRet = btrGetNext(hmFrom, tmRvf, imRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        btrDestroy hmFrom
    Next ilPass
    
    mGetChfAdfAgfSlf = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchSlf                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Advt index into tgCommAdf*
'*                                                     *
'*******************************************************
Public Function mBinarySearch(ilArray() As Integer, ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(ilArray)
    ilMax = UBound(ilArray) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = ilArray(ilMiddle) Then
            'found the match
            mBinarySearch = ilMiddle
            Exit Function
        ElseIf ilCode < ilArray(ilMiddle) Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearch = -1
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyCifRecs                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Variable length files     *
'*                                                     *
'*******************************************************
Private Function mCopyCifRecs() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilInsert As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilCpfRecLen As Integer
    Dim ilCsfRecLen As Integer
    Dim ilCnfRecLen As Integer
    Dim llRet As Long
    
    lacProcessing.Caption = "Processing " & "Cif.Btr"
    DoEvents
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCifRecs: Open Cif") Then
        mCopyCifRecs = False
        Exit Function
    End If
    imRecLen = btrRecordLength(hmFrom)  'Get and save record length
    mRemoveFile "Cif.Btr"
    ilRet = btrClone(hmFrom, smToDBPath & "Cif.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyCifRecs: Clone Cif") Then
        mCopyCifRecs = False
        Exit Function
    End If
    hmFCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCifRecs: Open 'From' Cpf") Then
        mCopyCifRecs = False
        Exit Function
    End If
    mRemoveFile "Cpf.Btr"
    ilRet = btrClone(hmFCpf, smToDBPath & "Cpf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyCifRecs: Clone Cpf") Then
        mCopyCifRecs = False
        Exit Function
    End If
    ilCpfRecLen = Len(tmCpf)

    hmFCsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFCsf, "", sgDBPath & "Csf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCifFiles: Open From Csf") Then
        mCopyCifRecs = False
        Exit Function
    End If
    'Remove done within mCopyCrfRecs as it is prior to this routine
    'mRemoveFile "Csf.Btr"
    'ilRet = btrClone(hmFCsf, smToDBPath & "Csf.Btr", 0)
    'If mBtrErrorMsg(ilRet, "mCopyCifRecs: Clone Csf") Then
    '    mCopyCifRecs = False
    '    Exit Function
    'End If
    ilCsfRecLen = Len(tmCsf)


    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCifRecs: Open 'To' Cif") Then
        mCopyCifRecs = False
        Exit Function
    End If
    hmTCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTCpf, "", smToDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCifRecs: Open 'To' Cpf") Then
        mCopyCifRecs = False
        Exit Function
    End If
    hmTCsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTCsf, "", smToDBPath & "Csf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCifRecs: Open 'To' Csf") Then
        mCopyCifRecs = False
        Exit Function
    End If
    
    lmCount = 0
    ilRet = btrGetFirst(hmFrom, tmCif, imRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE)
        ilInsert = False
        ilRet = mBinarySearch(igPECntrAdfCode(), tmCif.iAdfCode)
        If ilRet <> -1 Then
            If tmCif.sPurged = "H" Then
                ilInsert = False
            Else
                ilInsert = True
            End If
        Else
            If tmCif.sPurged = "P" Then
                ilInsert = True
            ElseIf tmCif.sPurged = "A" Then
                ilInsert = True
                'tmCif.sPurged = "P"
                'gPackDateLong lmNowDate, tmCif.iPurgeDate(0), tmCif.iPurgeDate(1)
            End If
            'tmCif.iAdfCode = 0
            'tmCif.lCsfCode = 0
            'tmCif.lCpfCode = 0
            If ilInsert Then
                mPurgeCif
            End If
        End If
        If ilInsert Then
            ilRet = btrInsert(hmTo, tmCif, imRecLen, INDEXKEY0)
            If mBtrErrorMsg(ilRet, "mCopyCifRecs: Insert Cif") Then
                btrDestroy hmFrom
                btrDestroy hmTo
                btrDestroy hmFCpf
                btrDestroy hmTCpf
                btrDestroy hmFCsf
                btrDestroy hmTCsf
                mCopyCifRecs = False
                Exit Function
            End If
            lmCount = lmCount + 1
            If tmCif.lCsfCode > 0 Then
                ilCsfRecLen = Len(tmCxf)
                tmLongSrchKey.lCode = tmCif.lCsfCode
                ilRet = btrGetEqual(hmFCsf, tmCsf, ilCsfRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrInsert(hmTCsf, tmCsf, ilCsfRecLen, INDEXKEY0)
                End If
                If ilRet <> BTRV_ERR_NONE Then
                    tmCif.lCsfCode = 0
                End If
            End If
            If tmCif.lcpfCode > 0 Then
                tmLongSrchKey.lCode = tmCif.lcpfCode
                ilRet = btrGetEqual(hmFCpf, tmCpf, ilCpfRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If (ilRet = BTRV_ERR_NONE) Then
                    tmLongSrchKey.lCode = tmCif.lcpfCode
                    ilRet = btrGetEqual(hmTCpf, tmCpf, ilCpfRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If (ilRet <> BTRV_ERR_NONE) Then
                        ilRet = btrInsert(hmTCpf, tmCpf, ilCpfRecLen, INDEXKEY0)
                        If mBtrErrorMsg(ilRet, "mCopyCifRecs: Insert Cpf") Then
                            btrDestroy hmFrom
                            btrDestroy hmTo
                            btrDestroy hmFCpf
                            btrDestroy hmTCpf
                            btrDestroy hmFCsf
                            btrDestroy hmTCsf
                            mCopyCifRecs = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        ilRet = btrGetNext(hmFrom, tmCif, imRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Print #hmMsg, "    " & Str$(lmCount) & " Copy Rotations (Cif.btr) records Transferred"
    ReDim tgCopyUsed(0 To lmCount) As COPYUSED
    lmCount = 0
    ilRet = btrGetFirst(hmTo, tmCif, imRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE)
        If lmCount >= UBound(tgCopyUsed) Then
            ReDim Preserve tgCopyUsed(0 To lmCount + 1) As COPYUSED
        End If
        tgCopyUsed(lmCount).lCifCode = tmCif.lCode
        tgCopyUsed(lmCount).iMcfCode = tmCif.iMcfCode
        tgCopyUsed(lmCount).sUsed = "N"
        lmCount = lmCount + 1
        ilRet = btrGetNext(hmTo, tmCif, imRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If UBound(tgCopyUsed) > 1 Then
        ArraySortTyp fnAV(tgCopyUsed(), 0), UBound(tgCopyUsed), 0, LenB(tgCopyUsed(0)), 0, -2, 0
    End If
    hmTCnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTCnf, "", smToDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCifRecs: Open 'To' Cnf") Then
        mCopyCifRecs = False
        Exit Function
    End If
    ilCnfRecLen = Len(tmCnf)
    ilRet = btrGetFirst(hmTCnf, tmCnf, ilCnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE)
        llRet = mSetCopyUsed(tmCnf.lCifCode)
        ilRet = btrGetNext(hmTCnf, tmCnf, ilCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    btrDestroy hmTCnf
    btrDestroy hmFrom
    btrDestroy hmTo
    btrDestroy hmFCpf
    btrDestroy hmTCpf
    btrDestroy hmFCsf
    btrDestroy hmTCsf
    mCopyCifRecs = True
    Exit Function
mCopyCifRecsErr:
    Screen.MousePointer = vbDefault
    btrDestroy hmFrom
    btrDestroy hmTo
    btrDestroy hmFCsf
    btrDestroy hmTCsf
    On Error GoTo 0
    mCopyCifRecs = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyCdfRecs                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Variable length files     *
'*                                                     *
'*******************************************************
Private Function mCopyCdfRecs() As Integer
    Dim ilRet As Integer
    Dim ilInsert As Integer
    
    lacProcessing.Caption = "Processing " & "Cdf.Btr"
    DoEvents
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & "Cdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCdfRecs: Open " & "Cdf.Btr") Then
        mCopyCdfRecs = False
        Exit Function
    End If
    imRecLen = Len(tmCdf)    'btrRecordLength(hmFrom)  'Get and save record length
    imSvRecLen = imRecLen
    mRemoveFile "Cdf.Btr"
    ilRet = btrClone(hmFrom, smToDBPath & "Cdf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyCdfRecs: Clone " & "Cdf.Btr") Then
        mCopyCdfRecs = False
        Exit Function
    End If
    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & "Cdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyCdfRecs: Open To " & "Cdf.Btr") Then
        mCopyCdfRecs = False
        Exit Function
    End If
    lmCount = 0
    ilRet = btrGetFirst(hmFrom, tmCdf, imRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE)
        DoEvents
        If imTerminate Then
            btrDestroy hmFrom
            btrDestroy hmTo
            mCopyCdfRecs = False
            Exit Function
        End If
        ilInsert = False
        If tmCdf.iAdfCode > 0 Then
            ilRet = mBinarySearch(igPECntrAdfCode(), tmCdf.iAdfCode)
            If ilRet <> -1 Then
                ilInsert = True
            End If
        End If
        If tmCdf.iAgfCode > 0 Then
            ilRet = mBinarySearch(igPECntrAgfCode(), tmCdf.iAgfCode)
            If ilRet <> -1 Then
                ilInsert = True
            End If
        End If
        If ilInsert Then
            ilRet = btrInsert(hmTo, tmCdf, imRecLen, INDEXKEY0)
            If mBtrErrorMsg(ilRet, "mCopyCdfRecs: Insert " & "Cdf.Btr") Then
                btrDestroy hmFrom
                btrDestroy hmTo
                mCopyCdfRecs = False
                Exit Function
            End If
            lmCount = lmCount + 1
        End If
        imRecLen = imSvRecLen
        ilRet = btrGetNext(hmFrom, tmCdf, imRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Print #hmMsg, "    " & Str$(lmCount) & " records from " & "Cdf.Btr" & " Transferred"
    btrDestroy hmFrom
    btrDestroy hmTo
    mCopyCdfRecs = True
    Exit Function
mCopyCdfRecsErr:
    Screen.MousePointer = vbDefault
    btrDestroy hmFrom
    btrDestroy hmTo
    On Error GoTo 0
    mCopyCdfRecs = False
    Exit Function
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyUrfRecs                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy URF file                  *
'*                                                     *
'*******************************************************
Private Function mCopyUrfRecs() As Integer
    Dim ilRet As Integer
    Dim ilInsert As Integer
    
    lacProcessing.Caption = "Processing " & "Urf.Btr"
    DoEvents
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyUrfRecs: Open " & "Urf.Btr") Then
        mCopyUrfRecs = False
        Exit Function
    End If
    imRecLen = Len(tmUrf)    'btrRecordLength(hmFrom)  'Get and save record length
    imSvRecLen = imRecLen
    mRemoveFile "Urf.Btr"
    ilRet = btrClone(hmFrom, smToDBPath & "Urf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyUrfRecs: Clone " & "Urf.Btr") Then
        mCopyUrfRecs = False
        Exit Function
    End If
    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyUrfRecs: Open To " & "Urf.Btr") Then
        mCopyUrfRecs = False
        Exit Function
    End If
    lmCount = 0
    ilRet = btrGetFirst(hmFrom, tmUrf, imRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE)
        DoEvents
        If imTerminate Then
            btrDestroy hmFrom
            btrDestroy hmTo
            mCopyUrfRecs = False
            Exit Function
        End If
        ilInsert = False
        gUrfDecrypt tmUrf
        If (tmUrf.iCode <= 2) Then
            ilInsert = True
        Else
            If tmUrf.iSlfCode > 0 Then
                ilRet = mBinarySearch(igPECntrSlfCode(), tmUrf.iSlfCode)
                If ilRet <> -1 Then
                    ilInsert = True
                End If
            End If
        End If
        If ilInsert Then
            gUrfEncrypt tmUrf
            ilRet = btrInsert(hmTo, tmUrf, imRecLen, INDEXKEY0)
            If mBtrErrorMsg(ilRet, "mCopyUrfRecs: Insert " & "Urf.Btr") Then
                btrDestroy hmFrom
                btrDestroy hmTo
                mCopyUrfRecs = False
                Exit Function
            End If
            lmCount = lmCount + 1
        End If
        imRecLen = imSvRecLen
        ilRet = btrGetNext(hmFrom, tmUrf, imRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Print #hmMsg, "    " & Str$(lmCount) & " records from " & "Urf.Btr" & " Transferred"
    btrDestroy hmFrom
    btrDestroy hmTo
    mCopyUrfRecs = True
    Exit Function
mCopyUrfRecsErr:
    Screen.MousePointer = vbDefault
    btrDestroy hmFrom
    btrDestroy hmTo
    On Error GoTo 0
    mCopyUrfRecs = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyDrfRecs                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Copy Variable length files     *
'*                                                     *
'*******************************************************
Private Function mCopyDrfRecs() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilInsert As Integer
    Dim ilDpfRecLen As Integer
    Dim tlDpfSrchKey As DPFKEY1    'DPF key record image
    
    lacProcessing.Caption = "Processing " & "Drf.Btr"
    DoEvents
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFrom, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyDrfRecs: Open Drf") Then
        mCopyDrfRecs = False
        Exit Function
    End If
    imRecLen = btrRecordLength(hmFrom)  'Get and save record length
    mRemoveFile "Drf.Btr"
    ilRet = btrClone(hmFrom, smToDBPath & "Drf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyDrfRecs: Clone Drf") Then
        mCopyDrfRecs = False
        Exit Function
    End If
    hmFDpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmFDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyDrfRecs: Open 'From' Dpf") Then
        mCopyDrfRecs = False
        Exit Function
    End If
    mRemoveFile "Dpf.Btr"
    ilRet = btrClone(hmFDpf, smToDBPath & "Dpf.Btr", 0)
    If mBtrErrorMsg(ilRet, "mCopyDrfRecs: Clone Dpf") Then
        mCopyDrfRecs = False
        Exit Function
    End If
    ilDpfRecLen = Len(tmDpf)



    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTo, "", smToDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyDrfRecs: Open 'To' Drf") Then
        mCopyDrfRecs = False
        Exit Function
    End If
    hmTDpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTDpf, "", smToDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mCopyDrfRecs: Open 'To' Dpf") Then
        mCopyDrfRecs = False
        Exit Function
    End If
    
    lmCount = 0
    ilRet = btrGetFirst(hmFrom, tmDrf, imRecLen, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE)
        ilInsert = False
        For ilLoop = 0 To UBound(igPECSRNVefCode) - 1 Step 1
            If tmDrf.iVefCode = igPECSRNVefCode(ilLoop) Then
                ilInsert = True
                Exit For
            End If
        Next ilLoop
        If Not ilInsert Then
            If tmDrf.sDemoDataType = "P" Then
                ilInsert = True
            End If
        End If
        If ilInsert Then
            ilRet = btrInsert(hmTo, tmDrf, imRecLen, INDEXKEY0)
            If mBtrErrorMsg(ilRet, "mCopyDrfRecs: Insert Drf") Then
                btrDestroy hmFrom
                btrDestroy hmTo
                btrDestroy hmFDpf
                btrDestroy hmTDpf
                mCopyDrfRecs = False
                Exit Function
            End If
            lmCount = lmCount + 1
            
            tlDpfSrchKey.lDrfCode = tmDrf.lCode
            tlDpfSrchKey.iMnfDemo = 0
            ilRet = btrGetGreaterOrEqual(hmFDpf, tmDpf, ilDpfRecLen, tlDpfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmDpf.lDrfCode = tmDrf.lCode)
                ilRet = btrInsert(hmTDpf, tmDpf, ilDpfRecLen, INDEXKEY0)
                If mBtrErrorMsg(ilRet, "mCopyDrfRecs: Insert Dpf") Then
                    btrDestroy hmFrom
                    btrDestroy hmTo
                    btrDestroy hmFDpf
                    btrDestroy hmTDpf
                    mCopyDrfRecs = False
                    Exit Function
                End If
                ilRet = btrGetNext(hmFDpf, tmDpf, ilDpfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
        ilRet = btrGetNext(hmFrom, tmDrf, imRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Print #hmMsg, "    " & Str$(lmCount) & " Copy Rotations (Drf.btr) records Transferred"
    btrDestroy hmFrom
    btrDestroy hmTo
    btrDestroy hmFDpf
    btrDestroy hmTDpf
    mCopyDrfRecs = True
    Exit Function
mCopyDrfRecsErr:
    Screen.MousePointer = vbDefault
    btrDestroy hmFrom
    btrDestroy hmTo
    btrDestroy hmFDpf
    btrDestroy hmTDpf
    On Error GoTo 0
    mCopyDrfRecs = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCopyUsed                    *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set Copy Used flag              *
'*                                                     *
'*******************************************************
Public Function mSetCopyUsed(llCifCode As Long) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    llMin = LBound(tgCopyUsed)
    llMax = UBound(tgCopyUsed) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCifCode = tgCopyUsed(llMiddle).lCifCode Then
            'found the match
            tgCopyUsed(llMiddle).sUsed = "Y"
            mSetCopyUsed = llMiddle
            Exit Function
        ElseIf llCifCode < tgCopyUsed(llMiddle).lCifCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    mSetCopyUsed = -1
End Function

Private Sub mSetCopyPurgeStatus()
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim ilCifRecLen As Integer
    Dim ilMcfRecLen As Integer
    Dim ilAllPurged As Integer
    Dim ilMcfRemove As Integer
    
    hmTCif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTCif, "", smToDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mSetCopyPurgeStatus: Open 'To' Cif") Then
        Exit Sub
    End If
    ilCifRecLen = Len(tmCif)
    hmTMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTMcf, "", smToDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If mBtrErrorMsg(ilRet, "mSetCopyPurgeStatus: Open 'To' Mcf") Then
        Exit Sub
    End If
    ilMcfRecLen = Len(tmMcf)
    For llLoop = 0 To UBound(tgCopyUsed) - 1 Step 1
        tmLongSrchKey.lCode = tgCopyUsed(llLoop).lCifCode
        ilRet = btrGetEqual(hmTCif, tmCif, ilCifRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            tgCopyUsed(llLoop).sPurged = tmCif.sPurged
            If tgCopyUsed(llLoop).sUsed <> "Y" Then
                If tmCif.sPurged = "A" Then
                    'tmCif.sPurged = "P"
                    mPurgeCif
                    ilRet = btrUpdate(hmTCif, tmCif, ilCifRecLen)
                    tgCopyUsed(llLoop).sPurged = tmCif.sPurged
                End If
            End If
        End If
    Next llLoop
    ReDim imRemoveMcf(0 To 0) As Integer
    ilRet = btrGetFirst(hmTMcf, tmMcf, ilMcfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE)
        DoEvents
        If imTerminate Then
            btrDestroy hmTMcf
            btrDestroy hmTCif
            Exit Sub
        End If
        ilAllPurged = True
        For llLoop = 0 To UBound(tgCopyUsed) - 1 Step 1
            If tmMcf.iCode = tgCopyUsed(llLoop).iMcfCode Then
                If tgCopyUsed(llLoop).sPurged = "A" Then
                    ilAllPurged = False
                    Exit For
                End If
                If tgCopyUsed(llLoop).sUsed = "Y" Then
                    ilAllPurged = False
                    Exit For
                End If
            End If
        Next llLoop
        If ilAllPurged Then
            imRemoveMcf(UBound(imRemoveMcf)) = tmMcf.iCode
            ReDim Preserve imRemoveMcf(0 To UBound(imRemoveMcf) + 1) As Integer
        End If
        ilRet = btrGetNext(hmTMcf, tmMcf, ilMcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    For ilMcfRemove = 0 To UBound(imRemoveMcf) - 1 Step 1
        For llLoop = 0 To UBound(tgCopyUsed) - 1 Step 1
            If imRemoveMcf(ilMcfRemove) = tgCopyUsed(llLoop).iMcfCode Then
                tmLongSrchKey.lCode = tgCopyUsed(llLoop).lCifCode
                ilRet = btrGetEqual(hmTCif, tmCif, ilCifRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmTCif)
                End If
            End If
        Next llLoop
        tmIntSrchKey.iCode = imRemoveMcf(ilMcfRemove)
        ilRet = btrGetEqual(hmTMcf, tmMcf, ilMcfRecLen, tmIntSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmTMcf)
        End If
    Next ilMcfRemove
    btrDestroy hmTMcf
    btrDestroy hmTCif
End Sub
Private Sub mPurgeCif()
    tmCif.sCut = ""
    tmCif.sReel = ""
    tmCif.iLen = 0
    tmCif.iEtfCode = 0
    tmCif.iEnfCode = 0
    tmCif.iAdfCode = 0
    tmCif.lcpfCode = 0
    tmCif.iMnfComp(0) = 0
    tmCif.iMnfComp(1) = 0
    tmCif.iMnfAnn = 0
    tmCif.lCsfCode = 0
    tmCif.iNoTimesAir = 0
    'tmCif.sCartDisp = tmMcf.sCartDisp
    'tmCif.sTapeDisp = tmMcf.sTapeDisp
    tmCif.sPurged = "P"
    gPackDateLong lmNowDate, tmCif.iPurgeDate(0), tmCif.iPurgeDate(1)
    tmCif.iUsedDate(0) = 0
    tmCif.iUsedDate(1) = 0
    tmCif.iRotStartDate(0) = 0
    tmCif.iRotStartDate(1) = 0
    tmCif.iRotEndDate(0) = 0
    tmCif.iRotEndDate(1) = 0
    tmCif.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
    tmCif.sPrint = "N"
End Sub

Private Function mCheckRecLength() As Integer
    If Not gRecLengthOk("Adf.Btr", Len(tmAdf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Agf.Btr", Len(tmAgf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Slf.Btr", Len(tmSlf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Prf.Btr", Len(tmPrf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Sif.Btr", Len(tmSif)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Bsf.Btr", Len(tmBsf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Bvf.Btr", Len(tmBvf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Chf.Btr", Len(tmChf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Clf.Btr", Len(tmClf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Cff.Btr", Len(tmCff)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Cgf.Btr", Len(tmCgf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Msf.Btr", Len(tmMsf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    'If Not gRecLengthOk("Cxf.Btr", Len(tmCxf)) Then
    '    mCheckRecLength = False
    '    Exit Function
    'End If
    'If Not gRecLengthOk("Cdf.Btr", Len(tmCdf)) Then
    '    mCheckRecLength = False
    '    Exit Function
    'End If
    If Not gRecLengthOk("Sbf.Btr", Len(tmSbf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Cff.Btr", Len(tmCff)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Cif.Btr", Len(tmCif)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Vsf.Btr", Len(tmVsf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Cff.Btr", Len(tmCff)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Vcf.Btr", Len(tmVcf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Vlf.Btr", Len(tmVlf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Sdf.Btr", Len(tmSdf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Smf.Btr", Len(tmSmf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Dlf.Btr", Len(tmDlf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Crf.Btr", Len(tmCrf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Cnf.Btr", Len(tmCnf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    'If Not gRecLengthOk("Csf.Btr", Len(tmCsf)) Then
    '    mCheckRecLength = False
    '    Exit Function
    'End If
    If Not gRecLengthOk("Cpf.Btr", Len(tmCpf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Vef.Btr", Len(tmVef)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Rif.Btr", Len(tmRif)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Lcf.Btr", Len(tmLcf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Tzf.Btr", Len(tmTzf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Drf.Btr", Len(tmDrf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Dpf.Btr", Len(tmDpf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Ltf.Btr", Len(tmLtf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Lvf.Btr", Len(tmLvf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Lef.Btr", Len(tmLef)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Rvf.Btr", Len(tmRvf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Vof.Btr", Len(tmVof)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Vpf.Btr", Len(tmVpf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Pvf.Btr", Len(tmPvf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Pjf.Btr", Len(tmPjf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Stf.Btr", Len(tmStf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Urf.Btr", Len(tmUrf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Ihf.Btr", Len(tmIhf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Isf.Btr", Len(tmIsf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Ghf.Btr", Len(tmGhf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Gsf.Btr", Len(tmGsf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    mCheckRecLength = True
End Function
