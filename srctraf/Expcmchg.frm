VERSION 5.00
Begin VB.Form ExpCmChg 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5475
   ClientLeft      =   600
   ClientTop       =   2460
   ClientWidth     =   8925
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
   ScaleHeight     =   5475
   ScaleWidth      =   8925
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
      Height          =   1290
      ItemData        =   "Expcmchg.frx":0000
      Left            =   285
      List            =   "Expcmchg.frx":0002
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3465
      Visible         =   0   'False
      Width           =   8190
   End
   Begin VB.TextBox edcNoDays 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3315
      MaxLength       =   10
      TabIndex        =   5
      Top             =   390
      Width           =   555
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
      Left            =   1965
      Picture         =   "Expcmchg.frx":0004
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   405
      Width           =   195
   End
   Begin VB.TextBox edcStartDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1035
      MaxLength       =   10
      TabIndex        =   2
      Top             =   390
      Width           =   930
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   1050
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   16
      Top             =   615
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
         TabIndex        =   19
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
         Picture         =   "Expcmchg.frx":00FE
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   17
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
            TabIndex        =   18
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
         TabIndex        =   6
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcScreen 
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
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   2685
      TabIndex        =   0
      Top             =   0
      Width           =   2685
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7350
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4950
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6735
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4950
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7005
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4815
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcVehicle 
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
      Height          =   2340
      Left            =   1860
      MultiSelect     =   2  'Extended
      TabIndex        =   7
      Top             =   1020
      Width           =   4800
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "&Export"
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
      Left            =   3195
      TabIndex        =   8
      Top             =   4965
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
      Left            =   4590
      TabIndex        =   9
      Top             =   4965
      Width           =   1050
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   195
      Top             =   4935
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacNoDays 
      Appearance      =   0  'Flat
      Caption         =   "# of Days"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2415
      TabIndex        =   4
      Top             =   375
      Width           =   960
   End
   Begin VB.Label lacStartDate 
      Appearance      =   0  'Flat
      Caption         =   "Air Date"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   75
      TabIndex        =   1
      Top             =   375
      Width           =   960
   End
   Begin VB.Label lacErrors 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   6525
      TabIndex        =   15
      Top             =   1815
      Width           =   1725
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   6540
      TabIndex        =   13
      Top             =   1365
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   3225
      TabIndex        =   11
      Top             =   1395
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "ExpCmChg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Expcmchg.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpCmChg.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Export Commercial Change input screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim hmExport As Integer   'file hanle
Dim hmMsg As Integer        'Message File Handle
'Contract record information
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
'Contract Line record information
Dim hmClf As Integer            'Contract Line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
'Short Title Vehicle Table record information
Dim hmVsf As Integer            'Short Title Vehicle Table file handle
Dim imVsfRecLen As Integer        'VSF record length
Dim tmVsf As VSF
'Spot record information
Dim hmSdf As Integer            'Spot Detail file handle
Dim tmSdfSrchKey3 As LONGKEY0            'SDF record image
Dim imSdfRecLen As Integer        'SDF record length
Dim tmSdf As SDF
'Short Title record information
Dim hmSif As Integer            'Short Title file handle
Dim imSifRecLen As Integer        'SIF record length
Dim tmSif As SIF
' Vehicle File
Dim hmVef As Integer            'Vehiclee file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim tmVefCode() As SELVEFCC
Dim tmStnCode() As STNCODECC      'Table of unique station code names
'Vehicle Link
Dim hmVlf As Integer            'Vehicle Link file handle
Dim tmVlf As VLF                'VEF record image
Dim tmVlfSrchKey As VLFKEY0            'VEF record image
Dim imVlfRecLen As Integer        'VEF record length
Dim hmStf As Integer            'MG and outside Times file handle
Dim tmStf As STF                'RPF record image
Dim imStfRecLen As Integer        'RPF record length
Dim tmAStf() As STF
'Dim tmRec As LPOPREC
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
'Advertiser name
Dim hmAdf As Integer
Dim tmAdf As ADF
Dim tmAdfSrchKey As INTKEY0 'ANF key record image
Dim imAdfRecLen As Integer  'ANF record length
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim tmSort() As TYPESORTCC
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
' MsgBox parameters
Const vbOkOnly = 0                 ' OK button only
Const vbCritical = 16          ' Critical message
Const vbApplicationModal = 0
Const INDEXKEY0 = 0
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub
Private Sub cmcCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcExport_Click()
    Dim ilRet As Integer
    Dim ilDBRet As Integer
    Dim ilLoop As Integer
    Dim slCreateStartDate As String
    Dim slCreateEndDate As String
    Dim slAirStartDate As String
    Dim slAirEndDate As String
    Dim slName As String
    Dim slExportFile As String
    Dim slMsgFile As String
    Dim slMsgFileName As String
    Dim slMsgLine As String
    Dim slTimeStamp As String
    Dim slNowDate As String
    Dim slStnCode As String
    Dim slGenDate As String
    Dim slNYear As String
    Dim slNMonth As String
    Dim slNDay As String
    Dim slStr As String
    Dim ilDays As Integer
    Dim ilAllSpots As Integer
    Dim ilStartIndex As Integer
    Dim ilEndIndex As Integer
    Dim ilLastStartIndex As Integer
    Dim slVehName As String
    Dim slLastVehName As String
    Dim ilPageNo As Integer
    Dim ilLineNo As Integer
    Dim slRecord As String
    Dim slBlank As String
    Dim slHeader As String
    Dim ilNoLines As Integer
    Dim ilSortIndex As Integer
    Dim ilPos As Integer
    Dim slTitle As String
    Dim ilShowTitle As Integer
    Dim ilIncludeBlank As Integer
    Dim ilVef As Integer
    Dim tlVef As VEF
    ReDim slFields(0 To 6) As String
    If imExporting Then
        Exit Sub
    End If
    
    On Error GoTo ExportError
    ilIncludeBlank = False
    slStr = edcStartDate.Text
    If Not gValidDate(slStr) Then
        Beep
        edcStartDate.SetFocus
        Exit Sub
    End If
    slAirStartDate = slStr
    ilDays = Val(edcNoDays.Text)
    If ilDays = 0 Then
        Beep
        edcNoDays.SetFocus
        Exit Sub
    End If
    slAirEndDate = Format$(gDateValue(slAirStartDate) + ilDays - 1, "m/d/yy")

    slBlank = ""
    slNowDate = Format$(Now, "mm/dd/yyyy")
    gObtainYearMonthDayStr slNowDate, True, slNYear, slNMonth, slNDay
    slGenDate = right$(slNYear, 2) & slNMonth & slNDay

    slTitle = "Date     Time       Action Product              Length"
    Screen.MousePointer = vbHourglass
    lbcErrors.Clear
    lbcErrors.Visible = True
    slCreateStartDate = ""   'Start date
    slCreateEndDate = ""   'End date
    ilAllSpots = False
    ReDim tmVefCode(0 To 0) As SELVEFCC
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            'Test if file exist- if so disallow vehicle selection
            slStnCode = Trim$(tmStnCode(ilLoop).sStnCode)
            slExportFile = sgExportPath & slStnCode & slGenDate & ".ccf"
            ilRet = 0
            'On Error GoTo cmcExportErr:
            'slTimeStamp = FileDateTime(slExportFile)
            ilRet = gFileExist(slExportFile)
            If ilRet = 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Commercial Change already generated for this date, Export terminated", vbOkOnly + vbCritical + vbApplicationModal, "Export"
                gAutomationAlertAndLogHandler "Commercial Change already generated for this date, Export terminated", vbOkOnly + vbCritical + vbApplicationModal, "Export"
                cmcCancel.SetFocus
                Exit Sub
            End If
            tmVefCode(UBound(tmVefCode)).iVefCode = tmStnCode(ilLoop).iVefCode
            tmVefCode(UBound(tmVefCode)).iWithData = False
            ReDim Preserve tmVefCode(0 To UBound(tmVefCode) + 1) As SELVEFCC
        End If
    Next ilLoop
    
    sgMessageFile = sgDBPath & "Messages\" & "ExptCmChng.Txt"
    
    gAutomationAlertAndLogHandler "** Export Commercial Changes ** "
    gAutomationAlertAndLogHandler "* StartDate = " & edcStartDate.Text
    gAutomationAlertAndLogHandler "* # Days = " & edcNoDays.Text
    
    gAutomationAlertAndLogHandler "Get stf.."
    'mObtainStf duplicated from RptGen.Bas then vehicle selection added
    mObtainStf ilAllSpots, slCreateStartDate, slCreateEndDate, slAirStartDate, slAirEndDate
    slLastVehName = ""
    ilStartIndex = LBound(tmSort)
    ilLastStartIndex = ilStartIndex
    gAutomationAlertAndLogHandler "determine stf index.."
    ilDBRet = mDetermineStfIndex(ilStartIndex, ilEndIndex)
    If ilDBRet Then
        Do While (ilDBRet = True)
            For ilLoop = 1 To 6 Step 1
                slFields(ilLoop) = ""
            Next ilLoop
            For ilLoop = LBound(tmAStf) To UBound(tmAStf) - 1 Step 1
                ilSortIndex = tmAStf(ilLoop).iLineNo
                ilRet = gParseItem(tmSort(ilSortIndex).sKey, 1, "|", slStr)
                If tmSort(ilSortIndex).iSelVefIndex >= 0 Then
                    tmVefCode(tmSort(ilSortIndex).iSelVefIndex).iWithData = True
                End If
                slFields(1) = Trim$(slStr)
                slName = Trim$(slStr)    'Save vehicle name- used for determining if new page required
                If slLastVehName <> slName Then
                    If slLastVehName <> "" Then 'Output end messages
                        'GoSub cmcExportHeader
                        'If Not mExportLine(slBlank, ilLineNo) Then
                        '    Exit Sub
                        'End If
                        ilIncludeBlank = True
                        slMsgFileName = slStnCode & "000000.ecc"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                            Exit Sub
                        End If
                        slMsgFileName = slStnCode & slGenDate & ".ecc"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                            Exit Sub
                        End If
                        slMsgFileName = "00" & slGenDate & ".ecc"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                            Exit Sub
                        End If
                        slMsgFileName = "00000000.ecc"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                            Exit Sub
                        End If
                        Close hmExport
                        slVehName = slName
                        ilIncludeBlank = False
                    Else
                        slVehName = slName
                    End If
                    slLastVehName = slVehName
                    slStnCode = Trim$(tmSort(ilSortIndex).sStnCode)
                    slExportFile = sgExportPath & slStnCode & slGenDate & ".ccf"
                    ilRet = 0
                    'On Error GoTo cmcExportErr:
                    'hmExport = FreeFile
                    ''Create file name based on vehicle name
                    'Open slExportFile For Output As hmExport
                    ilRet = gFileOpen(slExportFile, "Output", hmExport)
                    If ilRet <> 0 Then
                        Screen.MousePointer = vbDefault
                        ''MsgBox "Open " & slExportFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                        gAutomationAlertAndLogHandler "Open " & slExportFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                        cmcCancel.SetFocus
                        Exit Sub
                    End If
                    ilPageNo = 0
                    ilLineNo = 52
                    ilNoLines = 1
                    ilShowTitle = False
                    '6/3/16: Replaced GoSub
                    'GoSub cmcExportHeader
                    If Not mExportHeader(ilLineNo, ilNoLines, ilPageNo, slBlank, slHeader, slTitle, slVehName, slNowDate, ilShowTitle) Then
                        Exit Sub
                    End If
                    slMsgFileName = slStnCode & "000000.bcc"
                    '6/3/16: Replaced GoSub
                    'GoSub cmcExportSendMsg
                    If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                        Exit Sub
                    End If
                    slMsgFileName = slStnCode & slGenDate & ".bcc"
                    '6/3/16: Replaced GoSub
                    'GoSub cmcExportSendMsg
                    If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                        Exit Sub
                    End If
                    slMsgFileName = "00" & slGenDate & ".bcc"
                    '6/3/16: Replaced GoSub
                    'GoSub cmcExportSendMsg
                    If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                        Exit Sub
                    End If
                    slMsgFileName = "00000000.bcc"
                    '6/3/16: Replaced GoSub
                    'GoSub cmcExportSendMsg
                    If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                        Exit Sub
                    End If
                    If Not mExportLine(slBlank, ilLineNo) Then
                        Exit Sub
                    End If
                    If Not mExportLine(slTitle, ilLineNo) Then
                        Exit Sub
                    End If
                    If Not mExportLine(slBlank, ilLineNo) Then
                        Exit Sub
                    End If
                    ilShowTitle = True
                Else
                    'Add blank line between times
                    If ilLoop = LBound(tmAStf) Then
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportHeader
                        If Not mExportHeader(ilLineNo, ilNoLines, ilPageNo, slBlank, slHeader, slTitle, slVehName, slNowDate, ilShowTitle) Then
                            Exit Sub
                        End If
                        If Not mExportLine(slBlank, ilLineNo) Then
                            Exit Sub
                        End If
                    End If
                End If
                gUnpackDate tmAStf(ilLoop).iLogDate(0), tmAStf(ilLoop).iLogDate(1), slStr
                slFields(2) = slStr
                Do While Len(slFields(2)) < 8
                    slFields(2) = slFields(2) & " "
                Loop
                ilRet = gParseItem(tmSort(ilSortIndex).sKey, 7, "|", slStr)
                slFields(3) = Trim$(slStr)
                Do While Len(slFields(3)) < 10
                    slFields(3) = slFields(3) & " "
                Loop
                If tmAStf(ilLoop).sAction = "A" Then
                    slFields(4) = "Add   "
                Else
                    slFields(4) = "Remove"
                End If
                If tmAStf(ilLoop).lChfCode <> tmChf.lCode Then
                    tmChfSrchKey.lCode = tmAStf(ilLoop).lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                End If
                'slFields(5) = tmChf.sProduct
                If tmAStf(ilLoop).lSdfCode <= 0 Then    'ABC- old system
                    tmSdf.iVefCode = tmAStf(ilLoop).iVefCode
                    tmSdf.lChfCode = tmAStf(ilLoop).lChfCode
                    tmSdf.iLineNo = tmAStf(ilLoop).iLineNo
                    tmSdf.sPtType = "0"
                    tmSdf.iRotNo = 0
                Else
                    tmSdfSrchKey3.lCode = tmAStf(ilLoop).lSdfCode
                    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet <> BTRV_ERR_NONE Then
                        tmSdf.iVefCode = tmAStf(ilLoop).iVefCode
                        tmSdf.lChfCode = tmAStf(ilLoop).lChfCode
                        tmSdf.iLineNo = tmAStf(ilLoop).iLineNo
                        tmSdf.sPtType = "0"
                        tmSdf.iRotNo = 0
                    End If
                End If
                If tmChf.iAdfCode <> tmAdf.iCode Then
                    tmAdfSrchKey.iCode = tmChf.iAdfCode
                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                End If
                slFields(5) = gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf)  'tmChf.sProduct
                slFields(6) = Trim$(str$(tmAStf(ilLoop).iLen))
                Do While Len(slFields(6)) < 4
                    slFields(6) = " " & slFields(6)
                Loop
                'Output record
                slRecord = ""
                If ilLoop = LBound(tmAStf) Then 'include date and time
                    slRecord = slFields(2) & " " & slFields(3)
                Else
                    Do While Len(slRecord) < Len(slFields(2)) + Len(slFields(3)) + 1
                        slRecord = slRecord & " "
                    Loop
                End If
                slRecord = slRecord & " " & slFields(4) & " " & slFields(5) & " " & slFields(6)
                ilNoLines = 1   'UBound(tmAStf)
                '6/3/16: Replaced GoSub
                'GoSub cmcExportHeader
                If Not mExportHeader(ilLineNo, ilNoLines, ilPageNo, slBlank, slHeader, slTitle, slVehName, slNowDate, ilShowTitle) Then
                    Exit Sub
                End If
                If Not mExportLine(slRecord, ilLineNo) Then
                    Exit Sub
                End If

            Next ilLoop
            ilStartIndex = ilEndIndex + 1
            ilLastStartIndex = ilStartIndex
            ilDBRet = mDetermineStfIndex(ilStartIndex, ilEndIndex)
        Loop
        'GoSub cmcExportHeader
        'If Not mExportLine(slBlank, ilLineNo) Then
        '    Exit Sub
        'End If
        ilIncludeBlank = True
        slMsgFileName = slStnCode & "000000.ecc"
        '6/3/16: Replaced GoSub
        'GoSub cmcExportSendMsg
        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
            Exit Sub
        End If
        slMsgFileName = slStnCode & slGenDate & ".ecc"
        '6/3/16: Replaced GoSub
        'GoSub cmcExportSendMsg
        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
            Exit Sub
        End If
        slMsgFileName = "00" & slGenDate & ".ecc"
        '6/3/16: Replaced GoSub
        'GoSub cmcExportSendMsg
        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
            Exit Sub
        End If
        slMsgFileName = "00000000.ecc"
        '6/3/16: Replaced GoSub
        'GoSub cmcExportSendMsg
        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
            Exit Sub
        End If
        Close hmExport
        ilIncludeBlank = False
    End If
    
    gAutomationAlertAndLogHandler "Output vehicle without data.."
    'Output any vehicle without data
    For ilLoop = 0 To UBound(tmVefCode) - 1 Step 1
        If Not tmVefCode(ilLoop).iWithData Then
            For ilVef = 0 To UBound(tmStnCode) - 1 Step 1
                If tmStnCode(ilVef).iVefCode = tmVefCode(ilLoop).iVefCode Then
                    tmVefSrchKey.iCode = tmVefCode(ilLoop).iVefCode
                    ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        slVehName = Trim$(tlVef.sName)
                        slStnCode = Trim$(tmStnCode(ilVef).sStnCode)
                        slExportFile = sgExportPath & slStnCode & slGenDate & ".ccf"
                        ilRet = 0
                        'On Error GoTo cmcExportErr:
                        'hmExport = FreeFile
                        ''Create file name based on vehicle name
                        'Open slExportFile For Output As hmExport
                        ilRet = gFileOpen(slExportFile, "Output", hmExport)
                        If ilRet <> 0 Then
                            Screen.MousePointer = vbDefault
                            ''MsgBox "Open " & slExportFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                            gAutomationAlertAndLogHandler "Open " & slExportFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                            cmcCancel.SetFocus
                            Exit Sub
                        End If
                        ilPageNo = 0
                        ilLineNo = 52
                        ilNoLines = 1
                        ilShowTitle = False
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportHeader
                        If Not mExportHeader(ilLineNo, ilNoLines, ilPageNo, slBlank, slHeader, slTitle, slVehName, slNowDate, ilShowTitle) Then
                            Exit Sub
                        End If
                        slMsgFileName = slStnCode & "000000.bcc"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                            Exit Sub
                        End If
                        slMsgFileName = slStnCode & slGenDate & ".bcc"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                            Exit Sub
                        End If
                        slMsgFileName = "00" & slGenDate & ".bcc"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                            Exit Sub
                        End If
                        slMsgFileName = "00000000.bcc"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                            Exit Sub
                        End If
                        If Not mExportLine(slBlank, ilLineNo) Then
                            Exit Sub
                        End If
                        If Not mExportLine(slTitle, ilLineNo) Then
                            Exit Sub
                        End If
                        If Not mExportLine(slBlank, ilLineNo) Then
                            Exit Sub
                        End If
                        'Add Message- No changes
                        slRecord = "**** There are no commercial changes today ****"
                        ilNoLines = 1
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportHeader
                        If Not mExportHeader(ilLineNo, ilNoLines, ilPageNo, slBlank, slHeader, slTitle, slVehName, slNowDate, ilShowTitle) Then
                            Exit Sub
                        End If
                        If Not mExportLine(slRecord, ilLineNo) Then
                            Exit Sub
                        End If

                        'GoSub cmcExportHeader
                        'If Not mExportLine(slBlank, ilLineNo) Then
                        '    Exit Sub
                        'End If
                        ilIncludeBlank = True
                        slMsgFileName = slStnCode & "000000.ecc"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                            Exit Sub
                        End If
                        slMsgFileName = slStnCode & slGenDate & ".ecc"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                            Exit Sub
                        End If
                        slMsgFileName = "00" & slGenDate & ".ecc"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                            Exit Sub
                        End If
                        slMsgFileName = "00000000.ecc"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        If Not mExportSendMsg(slMsgFileName, slMsgFile, slMsgLine, slNowDate, ilNoLines, ilLineNo, ilPageNo, ilIncludeBlank, slBlank, slHeader, slTitle, slVehName, ilShowTitle) Then
                            Exit Sub
                        End If
                        Close hmExport
                        ilIncludeBlank = False
                    End If
                End If
            Next ilVef
        End If
    Next ilLoop
    
    If Not ilAllSpots Then
        ilRet = btrBeginTrans(hmStf, 1000)
        If ilDBRet <> BTRV_ERR_NONE Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Export Not Completed, Remove Export Generated Files, Try Later", vbOkOnly + vbExclamation, "Export")
            imExporting = False
            Exit Sub
        End If
        For ilLoop = LBound(tmSort) To UBound(tmSort) - 1 Step 1
            Do
                ilDBRet = btrGetDirect(hmStf, tmStf, imStfRecLen, tmSort(ilLoop).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilDBRet <> BTRV_ERR_NONE Then
                    ilRet = btrAbortTrans(hmStf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Export Not Completed, Remove Export Generated Files, Try Later", vbOkOnly + vbExclamation, "Export")
                    imExporting = False
                    Exit Sub
                End If
                'tmRec = tmStf
                'ilRet = gGetByKeyForUpdate("STF", hmStf, tmRec)
                'tmStf = tmRec
                'If ilRet <> BTRV_ERR_NONE Then
                '    ilRet = btrAbortTrans(hmStf)
                '    Screen.MousePointer = vbDefault
                '    ilRet = MsgBox("Export Not Completed, Remove Export Generated Files, Try Later", vbOkOnly + vbExclamation, "Export")
                '    imExporting = False
                '    Exit Sub
                'End If
                tmStf.sPrint = "P"
                ilDBRet = btrUpdate(hmStf, tmStf, imStfRecLen)
            Loop While ilDBRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrAbortTrans(hmStf)
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Export Not Completed, Try Later", vbOkOnly + vbExclamation, "Export")
                imExporting = False
                Exit Sub
            End If
        Next ilLoop
        ilRet = btrEndTrans(hmStf)
    End If
    
    gAutomationAlertAndLogHandler "Export Complete"
    cmcCancel.Caption = "&Done"
    cmcCancel.SetFocus
    Screen.MousePointer = vbDefault
    imExporting = False
    Exit Sub
'cmcExportErr:
'    ilRet = Err.Number
'    Resume Next
'cmcExportHeader:
'    If ilLineNo + ilNoLines > 52 Then
'        If ilPageNo = 0 Then
'            'slRecord = ""
'            If Not mExportLine(slBlank, ilLineNo) Then
'                Exit Sub
'            End If
'        Else
'            slHeader = Chr(12)  'Form Feed
'            If Not mExportLine(slHeader, ilLineNo) Then
'                Exit Sub
'            End If
'        End If
'        ilPageNo = ilPageNo + 1
'        ilLineNo = 0
'        slHeader = " "
'        Do While Len(slHeader) < 35
'            slHeader = slHeader & " "
'        Loop
'        slHeader = slHeader & Trim$(tgSpf.sGClient)
'        If Not mExportLine(slHeader, ilLineNo) Then
'            Exit Sub
'        End If
'        slHeader = " "
'        Do While Len(slHeader) < 35
'            slHeader = slHeader & " "
'        Loop
'        slHeader = slHeader & "Commercial Changes "
'        If Not mExportLine(slHeader, ilLineNo) Then
'            Exit Sub
'        End If
'        slHeader = " "
'        Do While Len(slHeader) < 35
'            slHeader = slHeader & " "
'        Loop
'        slHeader = slHeader & slVehName
'        If Not mExportLine(slHeader, ilLineNo) Then
'            Exit Sub
'        End If
'        slHeader = " "
'        Do While Len(slHeader) < 35
'            slHeader = slHeader & " "
'        Loop
'        slHeader = slHeader & slNowDate & "  "
'        slHeader = slHeader & "Page:"
'        slStr = Trim$(str$(ilPageNo))
'        Do While Len(slStr) < 5
'            slStr = " " & slStr
'        Loop
'        slHeader = slHeader & slStr
'        If Not mExportLine(slHeader, ilLineNo) Then
'            Exit Sub
'        End If
'        If Not mExportLine(slBlank, ilLineNo) Then
'            Exit Sub
'        End If
'        slHeader = ""
'        If ilShowTitle Then
'            If Not mExportLine(slTitle, ilLineNo) Then
'                Exit Sub
'            End If
'            If Not mExportLine(slBlank, ilLineNo) Then
'                Exit Sub
'            End If
'        End If
'    End If
'    Return
'cmcExportSendMsg:
'    ilRet = 0
'    On Error GoTo cmcExportErr:
'    hmMsg = FreeFile
'    slMsgFile = sgExportPath & slMsgFileName
'    Open slMsgFile For Input Access Read As hmMsg
'    If ilRet = 0 Then
'        Do
'            On Error GoTo cmcExportErr:
'            Line Input #hmMsg, slMsgLine
'            On Error GoTo 0
'            If (ilRet <> 0) Then    'Ctrl Z
'                Exit Do
'            End If
'            If Len(slMsgLine) > 0 Then
'                If (Asc(slMsgLine) = 26) Then    'Ctrl Z
'                    Exit Do
'                End If
'                ilPos = InStr(UCase$(slMsgLine), "XX/XX/XXXX")
'                If ilPos > 0 Then
'                    Mid$(slMsgLine, ilPos) = slNowDate
'                End If
'            End If
'            ilNoLines = 1
'            If ilIncludeBlank Then
'                GoSub cmcExportHeader
'                If Not mExportLine(slBlank, ilLineNo) Then
'                    Exit Sub
'                End If
'                ilIncludeBlank = False
'            End If
'            GoSub cmcExportHeader
'            If Not mExportLine(slMsgLine, ilLineNo) Then
'                Exit Sub
'            End If
'        Loop
'        Close hmMsg
'    End If
'    Return
    Exit Sub
ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)


End Sub
Private Sub cmcExport_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub cmcStartDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub
Private Sub cmcStartDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcNoDays_GotFocus()
    plcCalendar.Visible = False
    If edcNoDays.Text = "" Then
        edcNoDays.Text = 1
    End If
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcNoDays_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcStartDate_Change()
    Dim slStr As String
    slStr = edcStartDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcStartDate_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    gCtrlGotFocus edcStartDate
End Sub
Private Sub edcStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcStartDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcStartDate.SelLength <> 0 Then    'avoid deleting two characters
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
Private Sub edcStartDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDate.Text = slDate
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDate.Text = slDate
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
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
    DoEvents    'Process events so pending keys are not sent to this
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
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
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tmVefCode
    Erase tmStnCode
    Erase tmAStf
    Erase tmSort
    ilRet = btrClose(hmStf)
    ilRet = btrClose(hmVlf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmSif)
    ilRet = btrClose(hmAdf)
    btrDestroy hmVef
    btrDestroy hmVlf
    btrDestroy hmCHF
    btrDestroy hmClf
    btrDestroy hmVsf
    btrDestroy hmSdf
    btrDestroy hmSif
    btrDestroy hmAdf
    btrDestroy hmStf
    
    Set ExpCmChg = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcVehicle_GotFocus()
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
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    slStr = edcStartDate.Text
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
'*      Procedure Name:mDetermineStfIndex               *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine next Stf index values *
'*                     Build tmAStf with records       *
'*                                                     *
'*******************************************************
Private Function mDetermineStfIndex(ilStartIndex As Integer, ilEndIndex As Integer) As Integer
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim ilLoop1 As Integer
    Dim ilLoop2 As Integer
    Dim ilIndex As Integer
    Dim ilAnyRemoved As Integer
    Dim llTime0 As Long
    Dim llTime1 As Long
    Dim slVeh0 As String
    Dim slVeh1 As String
    Dim slTime As String
    Do
        ReDim tmAStf(0 To 0) As STF
        If ilStartIndex >= UBound(tmSort) Then
            mDetermineStfIndex = False
            Exit Function
        End If
        ilRet = btrGetDirect(hmStf, tmAStf(0), imStfRecLen, tmSort(ilStartIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        tmAStf(0).iLineNo = ilStartIndex 'Save index
        ilRet = gParseItem(tmSort(ilStartIndex).sKey, 1, "|", slVeh0)
        ilRet = gParseItem(tmSort(ilStartIndex).sKey, 3, "|", slTime)
        slVeh0 = Trim$(slVeh0)
        slTime = Trim$(slTime)
        llTime0 = Val(slTime)
        ilUpper = 1
        ReDim Preserve tmAStf(0 To ilUpper) As STF
        ilEndIndex = ilStartIndex + 1
        Do While ilEndIndex < UBound(tmSort)
            ilRet = btrGetDirect(hmStf, tmAStf(ilUpper), imStfRecLen, tmSort(ilEndIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            tmAStf(ilUpper).iLineNo = ilEndIndex 'Save index
            ilRet = gParseItem(tmSort(ilEndIndex).sKey, 1, "|", slVeh1)
            ilRet = gParseItem(tmSort(ilEndIndex).sKey, 3, "|", slTime)
            slVeh1 = Trim$(slVeh1)
            slTime = Trim$(slTime)
            llTime1 = Val(slTime)
            If (StrComp(slVeh0, slVeh1, 0) <> 0) Or (tmAStf(0).iLogDate(0) <> tmAStf(ilUpper).iLogDate(0)) Or (tmAStf(0).iLogDate(1) <> tmAStf(ilUpper).iLogDate(1)) Or (llTime0 <> llTime1) Then
                Exit Do
            End If
            ilUpper = ilUpper + 1
            ilEndIndex = ilEndIndex + 1
            ReDim Preserve tmAStf(0 To ilUpper) As STF
        Loop
        ilEndIndex = ilEndIndex - 1
        'Remove pars -Any matching Contract # and length and Add with remove
        ilAnyRemoved = False
        Do
            ilAnyRemoved = False
            For ilLoop1 = LBound(tmAStf) To UBound(tmAStf) - 1 Step 1
                For ilLoop2 = LBound(tmAStf) To UBound(tmAStf) - 1 Step 1
                    If ilLoop1 <> ilLoop2 Then
                        If (tmAStf(ilLoop1).lChfCode = tmAStf(ilLoop2).lChfCode) And (tmAStf(ilLoop1).iLen = tmAStf(ilLoop2).iLen) And (tmAStf(ilLoop1).sAction <> tmAStf(ilLoop2).sAction) Then
                            ilAnyRemoved = True
                            ilUpper = 0
                            For ilIndex = LBound(tmAStf) To UBound(tmAStf) - 1 Step 1
                                If (ilIndex <> ilLoop1) And (ilIndex <> ilLoop2) Then
                                    tmAStf(ilUpper) = tmAStf(ilIndex)
                                    ilUpper = ilUpper + 1
                                End If
                            Next ilIndex
                            ReDim Preserve tmAStf(0 To ilUpper) As STF
                            Exit For
                        End If
                    End If
                Next ilLoop2
                If ilAnyRemoved Then
                    Exit For
                End If
            Next ilLoop1
        Loop While ilAnyRemoved
        If ilUpper <= 0 Then
            ilStartIndex = ilEndIndex + 1
        End If
    Loop While ilUpper <= 0
    mDetermineStfIndex = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mExportLine                     *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Send line to output             *
'*                                                     *
'*******************************************************
Private Function mExportLine(slRecord As String, ilLineNo As Integer) As Integer
    Dim ilRet As Integer
    On Error GoTo mExportLineErr
    ilRet = 0
    Print #hmExport, slRecord
    If ilRet <> 0 Then
        imExporting = False
        Close #hmExport
        Screen.MousePointer = vbDefault
        ''MsgBox "Error writing to file" & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
        gAutomationAlertAndLogHandler "Error writing to file" & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
        cmcCancel.SetFocus
        mExportLine = False
        Exit Function
    End If
    ilLineNo = ilLineNo + 1
    mExportLine = True
    Exit Function
mExportLineErr:
    ilRet = err.Number
    Resume Next
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
    Dim slStr As String
    Dim slSunStr As String
    imLBCDCtrls = 1
    imTerminate = False
    imFirstActivate = True
    'mParseCmmdLine
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoBytes = 0
    lmProcessedNoBytes = 0
    Screen.MousePointer = vbHourglass
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCmChg
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCmChg
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCmChg
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCmChg
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    hmSif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCmChg
    On Error GoTo 0
    imSifRecLen = Len(tmSif)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCmChg
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCmChg
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmVlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVlf, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCmChg
    On Error GoTo 0
    imVlfRecLen = Len(tmVlf)
    hmStf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmStf, "", sgDBPath & "Stf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCmChg
    On Error GoTo 0
    imStfRecLen = Len(tmStf)
    'Populate arrays to determine if records exist
    mVehPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        'mTerminate
        Exit Sub
    End If
    'plcGauge.Move ExpCmChg.Width / 2 - plcGauge.Width / 2
    'cmcFileConv.Move ExpCmChg.Width / 2 - cmcFileConv.Width / 2
    'cmcCancel.Move ExpCmChg.Width / 2 - cmcCancel.Width / 2 - cmcCancel.Width
    'cmcReport.Move ExpCmChg.Width / 2 - cmcReport.Width / 2 + cmcReport.Width
    imBSMode = False
    imCalType = 0   'Standard
    mInitBox
    slStr = Format$(gNow(), "m/d/yy")
    slSunStr = gObtainNextSunday(slStr)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    edcStartDate.Text = slStr
    edcNoDays.Text = Trim$(str$(gDateValue(slSunStr) - gDateValue(slStr) + 1))
    lacDate.Visible = False
    gCenterStdAlone ExpCmChg
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
        
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
    Dim ilLoop As Integer
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    plcCalendar.Move edcStartDate.Left, edcStartDate.Top + edcStartDate.Height
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainStf                      *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the Stf records to be    *
'*                     reported                        *
'*                                                     *
'*******************************************************
Private Sub mObtainStf(ilAllRec As Integer, slCreateStartDate As String, slCreateEndDate As String, slAirStartDate As String, slAirEndDate As String)
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim slDate As String
    Dim slAirDate As String
    Dim slEffDate As String
    Dim slTermDate As String
    Dim llDate As Long
    Dim llTime As Long
    Dim slTime As String
    Dim ilDay As Integer
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilUpper As Integer
    Dim tlVef As VEF
    Dim tlVefL As VEF
    Dim ilEffDate0 As Integer
    Dim ilEffDate1 As Integer
    Dim ilVlfFd As Integer
    Dim slAirTime As String
    Dim slActionDate As String
    Dim slActionTime As String
    Dim slActionType As String  '1=Removed; 2=Added
    Dim ilFound As Integer
    Dim ilVef As Integer
    Dim ilSelVefIndex As Integer
    Dim ilTerminated As Integer
    ReDim tmSort(0 To 0) As TYPESORTCC
    tlVef.iCode = 0
    ilUpper = LBound(tmSort)
    btrExtClear hmStf   'Clear any previous extend operation
    ilExtLen = Len(tmStf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmStf   'Clear any previous extend operation
    'Screen.MousePointer = vbDefault
    'Using key 1 (stfCode as key 0 failed- it did not get all records)
    ilRet = btrGetFirst(hmStf, tmStf, imStfRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hmStf, llNoRec, -1, "UC", "STF", "") 'Set extract limits (all records)
        If Not ilAllRec Then
            tlCharTypeBuff.sType = "R"    'Extract all matching records
            ilOffSet = gFieldOffset("Stf", "StfPrint")
            If (slCreateStartDate = "") And (slCreateEndDate = "") And (slAirStartDate = "") And (slAirEndDate = "") Then
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
            Else
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
            End If
        End If
        If slCreateStartDate <> "" Then
            gPackDate slCreateStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Stf", "StfCreateDate")
            If (slCreateEndDate = "") And (slAirStartDate = "") And (slAirEndDate = "") Then
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            End If
        End If
        If slCreateEndDate <> "" Then
            gPackDate slCreateEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Stf", "StfCreateDate")
            If (slAirStartDate = "") And (slAirEndDate = "") Then
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            End If
        End If
        If slAirStartDate <> "" Then
            gPackDate slAirStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Stf", "StfLogDate")
            If (slAirEndDate = "") Then
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            End If
        End If
        If slAirEndDate <> "" Then
            gPackDate slAirEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Stf", "StfLogDate")
            ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        End If
        ilRet = btrExtAddField(hmStf, 0, ilExtLen)  'Extract Name
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        ilRet = btrExtGetNext(hmStf, tmStf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(ilUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmStf, tmStf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                'Build sort record
                gUnpackDateForSort tmStf.iCreateDate(0), tmStf.iCreateDate(1), slActionDate
                gUnpackTimeLong tmStf.iCreateTime(0), tmStf.iCreateTime(1), False, llTime
                slActionTime = Trim$(str$(llTime))
                Do While Len(slActionTime) < 5
                    slActionTime = "0" & slActionTime
                Loop
                If tmStf.sAction = "R" Then
                    slActionType = "1"
                Else
                    slActionType = "2"
                End If
                'Obtain vehicle and determine if selling or conventional
                If tmVef.iCode <> tmStf.iVefCode Then
                    tmVefSrchKey.iCode = tmStf.iVefCode
                    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                If ilRet = BTRV_ERR_NONE Then
                    If tmVef.sType = "S" Then
                        ilVlfFd = False
                        'Map selling to airing vehicle
                        'First obtain effective date
                        gUnpackDate tmStf.iLogDate(0), tmStf.iLogDate(1), slAirDate
                        ilDay = gWeekDayStr(slAirDate)
                        If ilDay <= 4 Then
                            ilDay = 0
                        ElseIf ilDay = 5 Then
                            ilDay = 6
                        ElseIf ilDay = 6 Then
                            ilDay = 7
                        End If
                        ilEffDate0 = 0  'tmStf.iLogDate(0)
                        ilEffDate1 = 0  'tmStf.iLogDate(1)
                        tmVlfSrchKey.iSellCode = tmVef.iCode
                        tmVlfSrchKey.iSellDay = ilDay
                        tmVlfSrchKey.iEffDate(0) = tmStf.iLogDate(0) 'ilEffDate0
                        tmVlfSrchKey.iEffDate(1) = tmStf.iLogDate(1) 'ilEffDate1
                        tmVlfSrchKey.iSellTime(0) = 0
                        tmVlfSrchKey.iSellTime(1) = 6144  '24*256
                        tmVlfSrchKey.iSellPosNo = 32000
                        ilRet = btrGetLessOrEqual(hmVlf, tmVlf, imVlfRecLen, tmVlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = tmVef.iCode)
                            ilTerminated = False
                            'Check for CBS
                            If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                                If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                                    ilTerminated = True
                                End If
                            End If
                            If (tmVlf.sStatus <> "P") And (tmVlf.iSellDay = ilDay) And (Not ilTerminated) Then
                                ilEffDate0 = tmVlf.iEffDate(0)
                                ilEffDate1 = tmVlf.iEffDate(1)
                                Exit Do
                            End If
                            ilRet = btrGetPrevious(hmVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        'If (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = tmVef.iCode) And (tmVlf.iSellDay = ilDay) And (tmVlf.sStatus = "C") Then
                        '    ilEffDate0 = tmVlf.iEffDate(0)
                        '    ilEffDate1 = tmVlf.iEffDate(1)
                        'Else
                        '    ilEffDate0 = 0
                        '    ilEffDate1 = 0
                        'End If
                        tmVlfSrchKey.iSellCode = tmVef.iCode    'selling vehicle code number
                        tmVlfSrchKey.iSellDay = ilDay     '0=M-F, 6= Sa, 7=Su
                        tmVlfSrchKey.iEffDate(0) = ilEffDate0 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
                        tmVlfSrchKey.iEffDate(1) = ilEffDate1 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
                        tmVlfSrchKey.iSellTime(0) = tmStf.iLogTime(0) 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
                        tmVlfSrchKey.iSellTime(1) = tmStf.iLogTime(1) 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
                        tmVlfSrchKey.iSellPosNo = 0   'Unit (spot) no- currently zero
                        tmVlfSrchKey.iSellSeq = 0     'Sequence number start at 1
                        ilRet = btrGetGreaterOrEqual(hmVlf, tmVlf, imVlfRecLen, tmVlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                        Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = tmVef.iCode) And (tmVlf.iSellDay = ilDay)
                            If tmVlf.sStatus = "C" Then
                                If (tmVlf.iSellTime(0) = tmStf.iLogTime(0)) And (tmVlf.iSellTime(1) = tmStf.iLogTime(1)) Then
                                    gUnpackDate tmVlf.iEffDate(0), tmVlf.iEffDate(1), slEffDate
                                    If gDateValue(slEffDate) <= gDateValue(slAirDate) Then
                                        gUnpackDate tmVlf.iTermDate(0), tmVlf.iTermDate(1), slTermDate
                                        If ((tmVlf.iTermDate(0) = 0) And (tmVlf.iTermDate(1) = 0)) Or (gDateValue(slAirDate) <= gDateValue(slTermDate)) Then
                                            ilFound = False
                                            For ilVef = 0 To UBound(tmVefCode) - 1 Step 1
                                                If tmVlf.iAirCode = tmVefCode(ilVef).iVefCode Then
                                                    ilSelVefIndex = ilVef
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Next ilVef
                                            ilVlfFd = True
                                            If ilFound Then
                                                If tlVef.iCode <> tmVlf.iAirCode Then
                                                    tmVefSrchKey.iCode = tmVlf.iAirCode
                                                    ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                Else
                                                    ilRet = BTRV_ERR_NONE
                                                End If
                                                If ilRet = BTRV_ERR_NONE Then
                                                    'Create one sort record
                                                    gUnpackDateForSort tmStf.iLogDate(0), tmStf.iLogDate(1), slDate
                                                    llDate = gDateValue(slDate)
                                                    gUnpackTime tmVlf.iAirTime(0), tmVlf.iAirTime(1), "A", "1", slAirTime
                                                    gUnpackTimeLong tmVlf.iAirTime(0), tmVlf.iAirTime(1), False, llTime
                                                    slTime = Trim$(str$(llTime))
                                                    Do While Len(slTime) < 5
                                                        slTime = "0" & slTime
                                                    Loop
                                                    tmSort(ilUpper).sKey = tlVef.sName & "|" & slDate & "|" & slTime & "|" & slActionDate & "|" & slActionTime & "|" & slActionType & "|" & slAirTime
                                                    For ilVef = 0 To UBound(tmStnCode) - 1 Step 1
                                                        If tmStnCode(ilVef).iVefCode = tlVef.iCode Then
                                                            tmSort(ilUpper).sStnCode = tmStnCode(ilVef).sStnCode
                                                            Exit For
                                                        End If
                                                    Next ilVef
                                                    tmSort(ilUpper).iSelVefIndex = ilSelVefIndex
                                                    tmSort(ilUpper).lRecPos = llRecPos
                                                    ilUpper = ilUpper + 1
                                                    ReDim Preserve tmSort(0 To ilUpper) As TYPESORTCC
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            ilRet = btrGetNext(hmVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        If Not ilVlfFd Then
                            gUnpackDateForSort tmStf.iLogDate(0), tmStf.iLogDate(1), slDate
                            llDate = gDateValue(slDate)
                            gUnpackTime tmStf.iLogTime(0), tmStf.iLogTime(1), "A", "1", slAirTime
                            gUnpackTimeLong tmStf.iLogTime(0), tmStf.iLogTime(1), False, llTime
                            slTime = Trim$(str$(llTime))
                            Do While Len(slTime) < 5
                                slTime = "0" & slTime
                            Loop
                            tmSort(ilUpper).sKey = "~" & tmVef.sName & "|" & slDate & "|" & slTime & "|" & slActionDate & "|" & slActionTime & "|" & slActionType & "|" & slAirTime
                            tmSort(ilUpper).sStnCode = ""
                            For ilVef = 0 To UBound(tmStnCode) - 1 Step 1
                                If tmStnCode(ilVef).iVefCode = tmVef.iCode Then
                                    tmSort(ilUpper).sStnCode = tmStnCode(ilVef).sStnCode
                                    Exit For
                                End If
                            Next ilVef
                            If Trim$(tmSort(ilUpper).sStnCode) <> "" Then
                                tmSort(ilUpper).iSelVefIndex = -1
                                tmSort(ilUpper).lRecPos = llRecPos
                                ilUpper = ilUpper + 1
                                ReDim Preserve tmSort(0 To ilUpper) As TYPESORTCC
                            Else
                               lbcErrors.AddItem "Link Missing for " & Trim$(tmVef.sName) & " on " & slAirDate & ", " & slAirTime
                            End If
                        End If
                    Else
                        ilFound = False
                        For ilVef = 0 To UBound(tmVefCode) - 1 Step 1
                            If tmVef.iVefCode > 0 Then  'Log vehicle defined
                                If tmVef.iVefCode = tmVefCode(ilVef).iVefCode Then
                                    ilFound = True
                                    ilSelVefIndex = ilVef
                                    Exit For
                                End If
                            Else
                                If tmVef.iCode = tmVefCode(ilVef).iVefCode Then
                                    ilFound = True
                                    ilSelVefIndex = ilVef
                                    Exit For
                                End If
                            End If
                        Next ilVef
                        If ilFound Then
                            'Create one sort record
                            gUnpackDateForSort tmStf.iLogDate(0), tmStf.iLogDate(1), slDate
                            llDate = gDateValue(slDate)
                            gUnpackTime tmStf.iLogTime(0), tmStf.iLogTime(1), "A", "1", slAirTime
                            gUnpackTimeLong tmStf.iLogTime(0), tmStf.iLogTime(1), False, llTime
                            slTime = Trim$(str$(llTime))
                            Do While Len(slTime) < 5
                                slTime = "0" & slTime
                            Loop
                            If tmVef.iVefCode > 0 Then
                                If tlVefL.iCode <> tmVef.iVefCode Then
                                    tmVefSrchKey.iCode = tmVef.iVefCode
                                    ilRet = btrGetEqual(hmVef, tlVefL, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                Else
                                    ilRet = BTRV_ERR_NONE
                                End If
                                If ilRet = BTRV_ERR_NONE Then
                                    tmSort(ilUpper).sKey = tlVefL.sName & "|" & slDate & "|" & slTime & "|" & slActionDate & "|" & slActionTime & "|" & slActionType & "|" & slAirTime
                                    For ilVef = 0 To UBound(tmStnCode) - 1 Step 1
                                        If tmStnCode(ilVef).iVefCode = tlVefL.iCode Then
                                            tmSort(ilUpper).sStnCode = tmStnCode(ilVef).sStnCode
                                            Exit For
                                        End If
                                    Next ilVef
                                    tmSort(ilUpper).iSelVefIndex = ilSelVefIndex
                                    tmSort(ilUpper).lRecPos = llRecPos
                                Else
                                    tmSort(ilUpper).sKey = Left$(tmVef.sName, 8) & " Log Veh Missing" & "|" & slDate & "|" & slTime & "|" & slActionDate & "|" & slActionTime & "|" & slActionType & "|" & slAirTime
                                    For ilVef = 0 To UBound(tmStnCode) - 1 Step 1
                                        If tmStnCode(ilVef).iVefCode = tmVef.iCode Then
                                            tmSort(ilUpper).sStnCode = tmStnCode(ilVef).sStnCode
                                            Exit For
                                        End If
                                    Next ilVef
                                    tmSort(ilUpper).iSelVefIndex = ilSelVefIndex
                                    tmSort(ilUpper).lRecPos = llRecPos
                                End If
                            Else
                                tmSort(ilUpper).sKey = tmVef.sName & "|" & slDate & "|" & slTime & "|" & slActionDate & "|" & slActionTime & "|" & slActionType & "|" & slAirTime
                                For ilVef = 0 To UBound(tmStnCode) - 1 Step 1
                                    If tmStnCode(ilVef).iVefCode = tmVef.iCode Then
                                        tmSort(ilUpper).sStnCode = tmStnCode(ilVef).sStnCode
                                        Exit For
                                    End If
                                Next ilVef
                                tmSort(ilUpper).iSelVefIndex = ilSelVefIndex
                                tmSort(ilUpper).lRecPos = llRecPos
                            End If
                            ilUpper = ilUpper + 1
                            ReDim Preserve tmSort(0 To ilUpper) As TYPESORTCC
                        End If
                    End If
                End If
                ilRet = btrExtGetNext(hmStf, tmStf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmStf, tmStf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    If ilUpper > 0 Then
        ArraySortTyp fnAV(tmSort(), 0), ilUpper, 0, LenB(tmSort(0)), 0, LenB(tmSort(0).sKey), 0 '100, 0
    End If
    Exit Sub

    ilRet = err.Number
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

    Screen.MousePointer = vbDefault

    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExpCmChg
    igManUnload = NO
End Sub
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
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilTest As Integer
    Dim ilVefCode As Integer
    Dim ilUsed As Integer
    Dim ilCount As Integer
    Dim ilStnCodeLen As Integer
    
    'ilRet = gPopUserVehicleBox(ExpCmChg, 7, lbcVehicle, Traffic!lbcUserVehicle)
    'ilRet = gPopUserVehicleBox(ExpCmChg, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH, lbcVehicle, Traffic!lbcUserVehicle)
    ilRet = gPopUserVehicleBox(ExpCmChg, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ExpCmChg
        On Error GoTo 0
    End If
    'Select on all vehicles that have Clearance as the format
    ReDim tmStnCode(0 To lbcVehicle.ListCount - 1) As STNCODECC
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        tmVefSrchKey.iCode = Val(slCode)
        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        tmStnCode(ilLoop).iVefCode = ilVefCode
        tmStnCode(ilLoop).sStnCode = tmVef.sCodeStn
        If Len(Trim$(tmStnCode(ilLoop).sStnCode)) = 0 Then
            tmStnCode(ilLoop).sStnCode = Left$(tmVef.sName, 5)
        End If
        ilStnCodeLen = Len(Trim$(tmStnCode(ilLoop).sStnCode)) - 1
        'Test if Name Used
        ilCount = 0
        Do
            ilUsed = False
            For ilTest = 0 To ilLoop - 1 Step 1
                If tmStnCode(ilLoop).sStnCode = tmStnCode(ilTest).sStnCode Then
                    If ilCount <= 9 Then
                        tmStnCode(ilLoop).sStnCode = Left$(tmStnCode(ilLoop).sStnCode, ilStnCodeLen) & Trim$(str$(ilCount))
                    Else
                        tmStnCode(ilLoop).sStnCode = Left$(tmStnCode(ilLoop).sStnCode, ilStnCodeLen) & Chr$(Asc("A") + ilCount - 10)
                    End If
                    ilCount = ilCount + 1
                    ilUsed = True
                    Exit For
                End If
            Next ilTest
        Loop While ilUsed
        'For ilTest = 0 To UBound(tgVpf) Step 1
        '    If ilVefCode = tgVpf(ilTest).iVefKCode Then
            ilTest = gBinarySearchVpf(ilVefCode)
            If ilTest <> -1 Then
                If tgVpf(ilTest).sExpHiCmmlChg = "Y" Then
                    lbcVehicle.Selected(ilLoop) = True
                End If
        '        Exit For
            End If
        'Next ilTest
    Next ilLoop
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
                edcStartDate.Text = Format$(llDate, "m/d/yy")
                edcStartDate.SelStart = 0
                edcStartDate.SelLength = Len(edcStartDate.Text)
                imBypassFocus = True
                edcStartDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcStartDate.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Export Commercial Changes"
End Sub

Private Function mExportSendMsg(slMsgFileName As String, slMsgFile As String, slMsgLine As String, slNowDate As String, ilNoLines As Integer, ilLineNo As Integer, ilPageNo As Integer, ilIncludeBlank As Integer, slBlank As String, slHeader As String, slTitle As String, slVehName As String, ilShowTitle As Integer) As Integer
    Dim ilRet As Integer
    Dim ilPos As Integer
    
    ilRet = 0
    'On Error GoTo mExportSendMsgErr:
    'hmMsg = FreeFile
    slMsgFile = sgExportPath & slMsgFileName
    'Open slMsgFile For Input Access Read As hmMsg
    ilRet = gFileOpen(slMsgFile, "Input Access Read", hmMsg)
    If ilRet = 0 Then
        err.Clear
        Do
            'On Error GoTo mExportSendMsgErr:
            Line Input #hmMsg, slMsgLine
            On Error GoTo 0
            ilRet = err.Number
            If (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            End If
            If Len(slMsgLine) > 0 Then
                If (Asc(slMsgLine) = 26) Then    'Ctrl Z
                    Exit Do
                End If
                ilPos = InStr(UCase$(slMsgLine), "XX/XX/XXXX")
                If ilPos > 0 Then
                    Mid$(slMsgLine, ilPos) = slNowDate
                End If
            End If
            ilNoLines = 1
            If ilIncludeBlank Then
                '6/3/16: Replaced GoSub
                'GoSub cmcExportHeader
                If Not mExportHeader(ilLineNo, ilNoLines, ilPageNo, slBlank, slHeader, slTitle, slVehName, slNowDate, ilShowTitle) Then
                    mExportSendMsg = False
                    Exit Function
                End If
                If Not mExportLine(slBlank, ilLineNo) Then
                    mExportSendMsg = False
                    Exit Function
                End If
                ilIncludeBlank = False
            End If
            '6/3/16: Replaced GoSub
            'GoSub cmcExportHeader
            If Not mExportHeader(ilLineNo, ilNoLines, ilPageNo, slBlank, slHeader, slTitle, slVehName, slNowDate, ilShowTitle) Then
                mExportSendMsg = False
                Exit Function
            End If
            If Not mExportLine(slMsgLine, ilLineNo) Then
                mExportSendMsg = False
                Exit Function
            End If
        Loop
        Close hmMsg
    End If
    mExportSendMsg = True
    Exit Function
'mExportSendMsgErr:
'    ilRet = Err.Number
'    Resume Next
End Function

Private Function mExportHeader(ilLineNo As Integer, ilNoLines As Integer, ilPageNo As Integer, slBlank As String, slHeader As String, slTitle As String, slVehName As String, slNowDate As String, ilShowTitle As Integer) As Integer
    Dim slStr As String
    
    If ilLineNo + ilNoLines > 52 Then
        If ilPageNo = 0 Then
            'slRecord = ""
            If Not mExportLine(slBlank, ilLineNo) Then
                mExportHeader = False
                Exit Function
            End If
        Else
            slHeader = Chr(12)  'Form Feed
            If Not mExportLine(slHeader, ilLineNo) Then
                mExportHeader = False
                Exit Function
            End If
        End If
        ilPageNo = ilPageNo + 1
        ilLineNo = 0
        slHeader = " "
        Do While Len(slHeader) < 35
            slHeader = slHeader & " "
        Loop
        slHeader = slHeader & Trim$(tgSpf.sGClient)
        If Not mExportLine(slHeader, ilLineNo) Then
            Exit Function
        End If
        slHeader = " "
        Do While Len(slHeader) < 35
            slHeader = slHeader & " "
        Loop
        slHeader = slHeader & "Commercial Changes "
        If Not mExportLine(slHeader, ilLineNo) Then
            mExportHeader = False
            Exit Function
        End If
        slHeader = " "
        Do While Len(slHeader) < 35
            slHeader = slHeader & " "
        Loop
        slHeader = slHeader & slVehName
        If Not mExportLine(slHeader, ilLineNo) Then
            mExportHeader = False
            Exit Function
        End If
        slHeader = " "
        Do While Len(slHeader) < 35
            slHeader = slHeader & " "
        Loop
        slHeader = slHeader & slNowDate & "  "
        slHeader = slHeader & "Page:"
        slStr = Trim$(str$(ilPageNo))
        Do While Len(slStr) < 5
            slStr = " " & slStr
        Loop
        slHeader = slHeader & slStr
        If Not mExportLine(slHeader, ilLineNo) Then
            mExportHeader = False
            Exit Function
        End If
        If Not mExportLine(slBlank, ilLineNo) Then
            mExportHeader = False
            Exit Function
        End If
        slHeader = ""
        If ilShowTitle Then
            If Not mExportLine(slTitle, ilLineNo) Then
                mExportHeader = False
                Exit Function
            End If
            If Not mExportLine(slBlank, ilLineNo) Then
                mExportHeader = False
                Exit Function
            End If
        End If
    End If
    mExportHeader = True
End Function
