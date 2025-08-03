VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExptEnco 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5130
   ClientLeft      =   225
   ClientTop       =   1620
   ClientWidth     =   9135
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
   ScaleHeight     =   5130
   ScaleWidth      =   9135
   Begin VB.ListBox lbcSort 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "exptenco.frx":0000
      Left            =   6000
      List            =   "exptenco.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox plcCalendarTo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   2880
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendarTo 
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
         Picture         =   "exptenco.frx":0014
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDateTo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   21
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.CommandButton cmcCalDnTo 
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
      Begin VB.CommandButton cmcCalUpTo 
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
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalNameTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   30
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   255
      Visible         =   0   'False
      Width           =   525
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
      Left            =   4680
      Picture         =   "exptenco.frx":2E2E
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   390
      Width           =   195
   End
   Begin VB.TextBox edcEndDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   4
      Top             =   390
      Width           =   930
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   840
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   18
      Top             =   600
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
         TabIndex        =   10
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
         TabIndex        =   22
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
         Picture         =   "exptenco.frx":2F28
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   19
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
            TabIndex        =   20
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
         TabIndex        =   11
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   165
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3570
      Width           =   1410
   End
   Begin VB.Timer tmcCancel 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2475
      Top             =   4545
   End
   Begin VB.ListBox lbcMsg 
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
      Height          =   2760
      Left            =   3720
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   780
      Width           =   5235
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
      Left            =   2905
      Picture         =   "exptenco.frx":5D42
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   390
      Width           =   195
   End
   Begin VB.TextBox edcStartDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   1
      Top             =   390
      Width           =   930
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   1545
      Top             =   4485
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   240
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
      Height          =   2760
      ItemData        =   "exptenco.frx":5E3C
      Left            =   165
      List            =   "exptenco.frx":5E3E
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   780
      Width           =   3375
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "&Export"
      Enabled         =   0   'False
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
      Left            =   3240
      TabIndex        =   8
      Top             =   4575
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
      Left            =   4680
      TabIndex        =   9
      Top             =   4575
      Width           =   1050
   End
   Begin VB.Label lacEncoText 
      Appearance      =   0  'Flat
      Caption         =   "Export Enco"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   32
      Top             =   0
      Width           =   1785
   End
   Begin VB.Label lacEndDate 
      Appearance      =   0  'Flat
      Caption         =   "To"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3480
      TabIndex        =   3
      Top             =   375
      Width           =   465
   End
   Begin VB.Label lacMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   180
      TabIndex        =   25
      Top             =   4185
      Width           =   8730
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      Top             =   4500
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   195
      TabIndex        =   23
      Top             =   3855
      Width           =   8730
   End
   Begin VB.Label lacStartDate 
      Appearance      =   0  'Flat
      Caption         =   "Export Date- From"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   375
      Width           =   1785
   End
   Begin VB.Label lacErrors 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5055
      TabIndex        =   17
      Top             =   1800
      Width           =   1725
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   5175
      TabIndex        =   15
      Top             =   1395
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   3225
      TabIndex        =   13
      Top             =   1395
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "ExptEnco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Copyright 1993 Counterpoint Software®. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExptEnco.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Export feed (for Dalet, Scott, Drake & Prophet) input screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim hmTo As Integer   'From file handle
Dim hmMsg As Integer   'From file handle
Dim lmNowDate As Long   'Todays date

'ODF name
Dim hmOdf As Integer
Dim tmTOdf As ODF
Dim tmOdfExt() As ODF
Dim imOdfRecLen As Integer  'ODF record length
Dim tmOdfSrchKey As ODFKEY0
'Advertiser name
Dim hmAdf As Integer
Dim tmAdf As ADF
Dim tmAdfSrchKey As INTKEY0 'ANF key record image
Dim imAdfRecLen As Integer  'ANF record length

'Copy/Product
Dim hmCpf As Integer
Dim tmCpf As CPF
Dim tmCpfSrchKey0 As LONGKEY0
Dim imCpfRecLen As Integer  'CPF record length

'Copy inventory record information
Dim hmCif As Integer        'Copy line file handle
Dim tmCifSrchKey0 As LONGKEY0
Dim imCifRecLen As Integer  'CIF record length
Dim tmCif As CIF            'CIF record image

' Vehicle File
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer     'VEF record length
Dim smVehName As String
'Vehicle Options
Dim tmVpf As VPF                'VPF record image
Dim imVpfRecLen As Integer      'VPF record length
Dim hmVpf As Integer            'Vehicle preference file handle

Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim imEvtType(0 To 14) As Integer

Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)

Dim imFoundspot As Integer  '1-5-05 flag to indicate if at least one spot found; if not, dont retain export file
' MsgBox parameters
Const vbOkOnly = 0                 ' OK button only
Const vbCritical = 16          ' Critical message
Const vbApplicationModal = 0
Const INDEXKEY0 = 0
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadODFSpots                   *
'*          Extended read on ODF file for given vehicle
'*          and date/time.  Build spots only in array
'*                                                     *
'*             Created:7/9/99       By:D. Hosaka       *
'*            Modified:              By:               *
'*                                                     *
'*                                                     *
'*******************************************************
Sub mReadODFSpots(hlODF As Integer, ilVefCode As Integer, slZone As String, slDate As String)


'
'   mReadODFSpots
'   Where:
'       hlOdf - ODF handle
'       ilVef - vehicle to obtain
'       slzone - zone to obtain
'       slDate - process one day at a time
'   <output>
'       tmOdfExt() array of ODF records
'
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpper As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlLongTypeBuff As POPLCODE          '6-12-02 long
    Dim tlOdfSrchKey0 As ODFKEY0
    Dim tlOdf As ODF
    ReDim tmOdfExt(0 To 0) As ODF
    Dim blAvailOk As Boolean
    Dim ilAnf As Integer
    ilRecLen = Len(tmOdfExt(0))
    llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlAdf) 'Obtain number of records
    btrExtClear hlODF   'Clear any previous extend operation
    'tlodfSrchKey2.iGenDate(0) = igGenDate(0)
    'tlodfSrchKey2.iGenDate(1) = igGenDate(1)
    'tlodfSrchKey2.lGenTime = lgGenTime

    tlOdfSrchKey0.iVefCode = ilVefCode
    gPackDate slDate, tlOdfSrchKey0.iAirDate(0), tlOdfSrchKey0.iAirDate(1)
    gPackTime "12M", tlOdfSrchKey0.iLocalTime(0), tlOdfSrchKey0.iLocalTime(1)
    tlOdfSrchKey0.sZone = "" 'slZone
    tlOdfSrchKey0.iSeqNo = 0

    ilRet = btrGetGreaterOrEqual(hlODF, tlOdf, Len(tlOdf), tlOdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_END_OF_FILE Then
        Exit Sub
    Else
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
    End If
    Call btrExtSetBounds(hlODF, llNoRec, -1, "UC", "ODF", "")  'Set extract limits (all records)

    tlIntTypeBuff.iType = ilVefCode
    ilOffSet = gFieldOffsetExtra("ODF", "OdfVefCode")
    ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)

    tlDateTypeBuff.iDate0 = igGenDate(0)
    tlDateTypeBuff.iDate1 = igGenDate(1)
    ilOffSet = gFieldOffsetExtra("ODF", "ODFGenDate")
    ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlDateTypeBuff, 4)
    tlLongTypeBuff.lCode = lgGenTime            '6-12-02
    ilOffSet = gFieldOffsetExtra("ODF", "ODFGenTime")
    ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLongTypeBuff, 4)

    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffsetExtra("ODF", "OdfAirDate")
    ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
    ilUpper = UBound(tmOdfExt)
    ilRet = btrExtAddField(hlODF, 0, ilRecLen)  'Extract the whole record

    ilRet = btrExtGetNext(hlODF, tlOdf, ilRecLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            Exit Sub
        End If
    End If
    Do While ilRet = BTRV_ERR_REJECT_COUNT
        ilRet = btrExtGetNext(hlODF, tlOdf, ilRecLen, llRecPos)
    Loop
    Do While ilRet = BTRV_ERR_NONE
        'gather spots only for matching vehicle, date and time zone
        If tlOdf.iMnfSubFeed = 0 And tlOdf.iType = 4 And Trim$(tlOdf.sZone) = Trim$(slZone) Then    'bypass records with subfeed
            '4/26/11: Test if avail spot should be exported
            blAvailOk = True
            ilAnf = gBinarySearchAnf(tlOdf.ianfCode, tgAvailAnf())
            If ilAnf <> -1 Then
                If tgAvailAnf(ilAnf).sAutomationExport = "N" Then
                    blAvailOk = False
                End If
            End If
            If blAvailOk Then
                tmOdfExt(ilUpper) = tlOdf
                ReDim Preserve tmOdfExt(0 To ilUpper + 1) As ODF
                ilUpper = ilUpper + 1
            End If
        End If
        ilRet = btrExtGetNext(hlODF, tlOdf, ilRecLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlODF, tlOdf, ilRecLen, llRecPos)
        Loop
    Loop
    Exit Sub
End Sub


Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    If lbcVehicle.ListCount <= 0 Then
        Exit Sub
    End If
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
        llRg = CLng(lbcVehicle.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcVehicle.HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcAll_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

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

Private Sub cmcCalDnTo_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendarTo_Paint
    edcEndDate.SelStart = 0
    edcEndDate.SelLength = Len(edcEndDate.Text)
    edcEndDate.SetFocus
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

Private Sub cmcCalUpTo_Click()
    plcCalendar.Visible = False
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendarTo_Paint
    edcEndDate.SelStart = 0
    edcEndDate.SelLength = Len(edcEndDate.Text)
    edcEndDate.SetFocus
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
    plcCalendarTo.Visible = False
End Sub

Private Sub cmcEndDate_Click()
    plcCalendarTo.Visible = Not plcCalendarTo.Visible
    edcEndDate.SelStart = 0
    edcEndDate.SelLength = Len(edcEndDate.Text)
    edcEndDate.SetFocus
    mSetCommands
End Sub

Private Sub cmcEndDate_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
    mSetCommands
End Sub

Private Sub cmcExport_Click()
    Dim slToFile As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim slStr As String
    Dim slFYear As String
    Dim slFMonth As String
    Dim slFDay As String
    Dim slLetter As String
    Dim slFileName As String
    Dim slDateTime As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llDate As Long
    Dim slInputStartDate As String
    Dim slInputEndDate As String
    Dim slDate As String
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim slExt As String
    Dim ilVpfIndex As Integer
    Dim ilLoZone As Integer
    Dim ilHiZone As Integer
    Dim ilUseZone As Integer
    Dim ilZoneLoop As Integer
    Dim ilDayOfWeek As Integer
    Dim ilZones As Integer
    Dim slZone As String
    Dim ilFoundZone As Integer
    Dim ilZoneToUse As Integer
    Dim slPDFFileName As String
    Dim slPDFDate As String
    Dim slPDFTime As String
    Dim slOutput As String

    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    
    lacProcessing.Caption = ""
    lacMsg.Caption = ""
    slInputStartDate = edcStartDate.Text
    If Not gValidDate(slInputStartDate) Then
        Beep
        edcStartDate.SetFocus
        Exit Sub
    End If
    llStartDate = gDateValue(slInputStartDate)

    slInputEndDate = edcEndDate.Text
    If Not gValidDate(slInputEndDate) And slInputEndDate <> "" Then
        Beep
        edcEndDate.SetFocus
        Exit Sub
    End If
    If slInputEndDate = "" Then
        slInputEndDate = slInputStartDate
    End If
    llEndDate = gDateValue(slInputEndDate)

    imExporting = True
    lbcMsg.Clear
    If Not mOpenMsgFile() Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    'get current date and time for filtering of odf records
    gCurrDateTime slPDFDate, slPDFTime, slMonth, slDay, slYear    'get system date and time for prepass filtering

    'creating log events
    lacProcessing.Caption = "Generating log days " & slInputStartDate & "-" & slInputEndDate
    ilRet = gGenODFDay(ExptEnco, slInputStartDate, slInputEndDate, lbcVehicle, lbcSort, tgUserVehicle(), "F", "E")

    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            smVehName = Trim$(slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            tmVefSrchKey.iCode = ilVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If ilRet = BTRV_ERR_NONE Then

                ilVpfIndex = -1
                ilVpfIndex = gVpfFind(ExptEnco, ilVefCode)              'determine vehicle options index
                'determine zones from vehicle options table
                ilUseZone = False                                       'assume not using zones until one is found in the vehicle options table
                ilZones = 0                                       'save zones requested by user : 0=none, 1 =est, 2= cst, 3 =mst, 4 = pst
                ilLoZone = 1                                            'low loop factor to process zones
                ilHiZone = 4                                            'hi loop factor to process zones
                If ilZones <> 0 Then                                    'user has requested one zone in particular
                    ilLoZone = ilZones
                    ilHiZone = ilZones
                End If
                If ilVpfIndex >= 0 Then                                  'associated vehicle options record exists
                    'If tgVpf(ilVpfIndex).sGZone(1) <> "   " Then
                    If tgVpf(ilVpfIndex).sGZone(0) <> "   " Then
                        ilUseZone = True
                    Else
                        'Zones not used, fake out flag to do 1 zone (EST)
                        ilZones = 1
                        ilLoZone = 1
                        ilHiZone = 1
                    End If
                Else
                    'no vehicle options table
                    'Zones not used, fake out flag to do 1 zone (EST)
                    ilZones = 1
                    ilLoZone = 1
                    ilHiZone = 1
                End If


                For llDate = llStartDate To llEndDate         'loop on all days within the selected vehicle
                    slDate = Format(llDate, "m/d/yy")
                    ilDayOfWeek = gWeekDayLong(llDate)

                    'Process one zone at a time, for all days in that zone
                    For ilZoneLoop = ilLoZone To ilHiZone               'loop on all time zones (or just the selective one,  variety of zones are not allowed)
                        Select Case ilZoneLoop
                            Case 1  'Eastern
                                If ilUseZone Then
                                    slZone = "EST"
                                Else
                                    slZone = "   "
                                End If
                            Case 2  'Central
                                slZone = "CST"
                            Case 3  'Mountain
                                slZone = "MST"
                            Case 4  'Pacific
                                slZone = "PST"
                        End Select

                        'test if zone is in the zone conversion table; if not bypass the zone
                        ilFoundZone = False

                        'For ilZoneToUse = 1 To 4
                        For ilZoneToUse = 0 To 3
                            If slZone = "   " Then
                                ilFoundZone = True
                                Exit For
                            Else
                                If Trim$(slZone) = tgVpf(ilVpfIndex).sGZone(ilZoneToUse) And tgVpf(ilVpfIndex).sGFed(ilZoneToUse) = "*" Then
                                    ilFoundZone = True
                                    Exit For
                                End If
                            End If
                        Next ilZoneToUse
                        If ilFoundZone Then

                            Select Case ilDayOfWeek
                                Case 0
                                    mReadODFSpots hmOdf, ilVefCode, slZone, slDate
                                Case 1
                                    mReadODFSpots hmOdf, ilVefCode, slZone, slDate
                                Case 2
                                    mReadODFSpots hmOdf, ilVefCode, slZone, slDate
                                Case 3
                                    mReadODFSpots hmOdf, ilVefCode, slZone, slDate
                                Case 4
                                    mReadODFSpots hmOdf, ilVefCode, slZone, slDate
                                Case 5
                                    mReadODFSpots hmOdf, ilVefCode, slZone, slDate
                                Case 6
                                    mReadODFSpots hmOdf, ilVefCode, slZone, slDate
                            End Select

                            'Determine file name from date: SSSSZMMDD.txt
                            'where SSSS = vehicle station code (max 5 char)
                            'Z = zone (W = PST, E = EST, C = CST, M = MST)
                            'MM = 2 char month
                            'DD = 2 char day
                            gObtainYearMonthDayStr slDate, True, slFYear, slFMonth, slFDay
                            'Print #hmMsg, " "
                            gAutomationAlertAndLogHandler " "
                            'Print #hmMsg, "** Generating Data for " & Trim$(smVehName) & " for " & slDate & " **"
                            gAutomationAlertAndLogHandler "** Generating Data for " & Trim$(smVehName) & " for " & slDate & " **"
                            lacProcessing.Caption = "Generating Data for " & Trim$(smVehName) & " for " & slDate


                            'slFileName = right$(slFYear, 2) & slFMonth & slFDay & (slZoneAbbrev)
                            slFileName = Trim$(Left(slZone, 1)) & slFMonth & slFDay
                            slExt = ".txt"
                            slLetter = tmVef.sCodeStn
                            ilRet = 0
                            'On Error GoTo cmcExportErr:

                            slToFile = sgExportPath & Trim$(slLetter) & Trim$(slFileName) & slExt   'ssssZmmdd.ext

                            'slDateTime = FileDateTime(slToFile)     '1-6-05 chged from sgExportPath to new sgProphetExportPath
                            ilRet = gFileExist(slToFile)
                            If ilRet = 0 Then
                                Kill slToFile   '1-6-05 chg from sgExportpath to new sgProphetExportPath
                            End If
                            On Error GoTo 0


                            ilRet = 0
                            'On Error GoTo cmcExportErr:
                            'hmTo = FreeFile
                            'Open slToFile For Output As hmTo
                            ilRet = gFileOpen(slToFile, "Output", hmTo)
                            If ilRet <> 0 Then
                                gClearODF                   'remove all the ODFs for the logs just created
                                'Print #hmMsg, "** Terminated **"
                                gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                                'Print #hmMsg, "Open error " & slToFile & " Error #" & str$(ilRet)
                                gAutomationAlertAndLogHandler "Open error " & slToFile & " Error #" & str$(ilRet)
                                Close #hmMsg
                                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                                Exit Sub
                            End If
                            'Print #hmMsg, "** Storing Output into " & slToFile & " **"
                            lacProcessing.Caption = "Writing Data to " & slLetter & slFileName & slExt
                            If Not mCreateEncoText() Then      'gather matching spot events in ODF and create the export record for day/vehicle
                                'Print #hmMsg, "** Terminated **"
                                gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                                Close #hmMsg
                                Close #hmTo
                                imExporting = False
                                gClearODF                   'remove all the ODFs for the logs just created
                                Screen.MousePointer = vbDefault
                                cmcCancel.SetFocus
                                Exit Sub
                            End If
                            If igDOE > 10 Then
                                DoEvents
                                igDOE = 0
                            End If
                            igDOE = igDOE + 1
                            ilRet = 0
                            'On Error GoTo cmcExportErr:

                            Close #hmTo         'close the output file

                            If imFoundspot Then         '1-5-05 if no spots found, dont retain export file and dont show file sent to message
                                lacProcessing.Caption = "Output for " & smVehName & " sent to " & slToFile
                                'Print #hmMsg, "** Output for " & smVehName & " sent to " & slToFile & " **"
                                gAutomationAlertAndLogHandler "** Output for " & smVehName & " sent to " & slToFile & " **"
                            Else
                                lacProcessing.Caption = "No spots found for " & smVehName & " on " & slDate & ", no export created"
                                Kill (slToFile)
                                'Print #hmMsg, "No spots found for " & smVehName & " on " & slDate & ", no export created"
                                gAutomationAlertAndLogHandler "No spots found for " & smVehName & " on " & slDate & ", no export created"
                            End If
                        End If
                    Next ilZoneLoop         'next zone
                    'create pdf, use station code and mmdd (no zone since all zones combined in 1 report )
                    slPDFFileName = Trim$(tmVef.sCodeStn) & slFMonth & slFDay & ".pdf"
                    igRptCallType = EXPORTENCO
                    igRptType = 2
                    slOutput = "2"
                    If (Not igStdAloneMode) Then
                        If igTestSystem Then
                            slStr = "ExptEnco^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & slOutput & "\0\" & slPDFFileName & "\" & slPDFDate & "\" & slPDFTime & "\" & str$(ilVefCode) & "\" & slDate
                        Else
                            slStr = "ExptEnco^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & slOutput & "\0\" & slPDFFileName & "\" & slPDFDate & "\" & slPDFTime & "\" & str$(ilVefCode) & "\" & slDate
                        End If
                    Else
                        If igTestSystem Then
                            slStr = "ExptEnco^Test^NOHELP\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & slOutput & "\0\" & slPDFFileName & "\" & slPDFDate & "\" & slPDFTime & "\" & str$(ilVefCode) & "\" & slDate
                        Else
                            slStr = "ExptEnco^Prod^NOHELP\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType)) & "\" & slOutput & "\0\" & slPDFFileName & "\" & slPDFDate & "\" & slPDFTime & "\" & str$(ilVefCode) & "\" & slDate
                        End If
                    End If
                    sgCommandStr = slStr
                    RptSelEx.Show vbModal
                    'Print #hmMsg, "** Output for " & smVehName & " sent to " & sgExportPath & slPDFFileName & " **"
                    gAutomationAlertAndLogHandler "** Output for " & smVehName & " sent to " & sgExportPath & slPDFFileName & " **"
                Next llDate                 'next date
                'Print #hmMsg, "** Completed " & Trim$(tmVef.sName) & " for " & slInputStartDate & "-"; slInputEndDate & " **"
                gAutomationAlertAndLogHandler "** Completed " & Trim$(tmVef.sName) & " for " & slInputStartDate & "-" & slInputEndDate & " **"
            Else                    'error, vehicle not found
                'Print #hmMsg, " "
                gAutomationAlertAndLogHandler " "
                'Print #hmMsg, "Name: " & slName & " not found"
                gAutomationAlertAndLogHandler "Name: " & slName & " not found"
                lacProcessing.Caption = slName & " not found: vehicle aborted"
            End If                  'ilret <> BTRV_err_none
        End If                  'lbcVehicle.Selected(ilLoop)

    Next ilLoop                 'next vehicle
    gClearODF                   'remove all the ODFs for the logs just created
    'Print #hmMsg, "** Completed Enco Export: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    gAutomationAlertAndLogHandler "** Completed Enco Export: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    Close #hmTo
    'Print #hmMsg, "** Completed " & smScreenCaption & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    Close #hmMsg
    On Error GoTo 0
    lacMsg.Caption = "Messages sent to " & sgDBPath & "Messages\" & "ExptEnco.Txt"
    Screen.MousePointer = vbDefault
    imExporting = False
    cmcCancel.Caption = "&Done"
    cmcCancel.SetFocus
    Erase tmOdfExt
    Exit Sub
'cmcExportErr:
'    ilRet = Err.Number
'    Resume Next

ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    
End Sub
Private Sub cmcExport_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub cmcStartDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
    mSetCommands
End Sub
Private Sub cmcStartDate_GotFocus()
    plcCalendarTo.Visible = False
    gCtrlGotFocus ActiveControl
    mSetCommands
End Sub

Private Sub edcEndDate_Change()
    Dim slStr As String
    plcCalendar.Visible = False

    slStr = edcEndDate.Text
    If Not gValidDate(slStr) Then
        lacDateTo.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendarTo_Paint   'mBoxCalDate called within paint
    mSetCommands
End Sub
Private Sub edcEndDate_Click()
    plcCalendar.Visible = False
    mSetCommands
End Sub

Private Sub edcEndDate_GotFocus()
    plcCalendar.Visible = False

    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    gCtrlGotFocus edcStartDate
    mSetCommands
End Sub

Private Sub edcEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
    mSetCommands
End Sub

Private Sub edcEndDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcEndDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    mSetCommands
End Sub

Private Sub edcEndDate_KeyUp(KeyCode As Integer, Shift As Integer)
Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendarTo.Visible = Not plcCalendarTo.Visible
        Else
            slDate = edcEndDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcEndDate.Text = slDate
            End If
        End If
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcEndDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcEndDate.Text = slDate
            End If
        End If
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
    End If
    mSetCommands
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
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
    mSetCommands
End Sub
Private Sub edcStartDate_Click()
    mSetCommands
End Sub
Private Sub edcStartDate_GotFocus()
    plcCalendarTo.Visible = False
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    gCtrlGotFocus edcStartDate
    mSetCommands
End Sub
Private Sub edcStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
    mSetCommands
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
    mSetCommands
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
    mSetCommands
End Sub

Private Sub Form_Activate()

    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    'Me.Visible = False
    'Me.Visible = True
    DoEvents    'Process events so pending keys are not sent to this
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        plcCalendarTo.Visible = False
        gFunctionKeyBranch KeyCode
'        If plcAutoType.Visible = True Then
'            plcAutoType.Visible = False
'            plcAutoType.Visible = True
'        End If
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
        'cmcCancel_Click
        tmcCancel.Enabled = True
        Me.Left = 2 * Screen.Width  'move off the screen so screen won't flash
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf

    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf


    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmVpf)
    btrDestroy hmVpf
    ilRet = btrClose(hmOdf)
    btrDestroy hmOdf
    
    Set ExptEnco = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lacEndDate_Click()
    mSetCommands
End Sub

Private Sub lacStartDate_Click()
    mSetCommands
End Sub

Private Sub lbcVehicle_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked  '9-12-02 False
        imSetAll = True
    End If
    mSetCommands
End Sub
Private Sub lbcVehicle_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
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
Private Sub mBoxCalDate(EditDate As Control, LabelDate As Control)
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    slStr = EditDate.Text   'edcStartDate.Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    LabelDate.Caption = slDay
                    LabelDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    LabelDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            LabelDate.Visible = False
        Else
            LabelDate.Visible = False
        End If
    Else
        lacDate.Visible = False
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
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    imTerminate = False
    imFirstActivate = True
    'mParseCmmdLine
    Screen.MousePointer = vbHourglass
    imAllClicked = False
    imSetAll = True
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False

    hmOdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmOdf, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptEnco
    On Error GoTo 0
    imOdfRecLen = Len(tmTOdf)

    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptEnco
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)

    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptEnco
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)

    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptEnco
    On Error GoTo 0
    imCifRecLen = Len(tmCif)

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptEnco
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    
    hmVpf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmVpf, "", sgDBPath & "Vpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptEnco
    On Error GoTo 0
    imVpfRecLen = Len(tmVpf)

    'Populate arrays to determine if records exist
    mVehPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        'mTerminate
        Exit Sub
    End If

    '4/26/11: Add test of avail attribute
    ilRet = gObtainAvail()
    
    For ilLoop = LBound(imEvtType) To UBound(imEvtType) Step 1
        imEvtType(ilLoop) = True
    Next ilLoop
    imEvtType(0) = False 'Don't include library names
    'plcGauge.Move ExptEnco.Width / 2 - plcGauge.Width / 2
    'cmcFileConv.Move ExptEnco.Width / 2 - cmcFileConv.Width / 2
    'cmcCancel.Move ExptEnco.Width / 2 - cmcCancel.Width / 2 - cmcCancel.Width
    'cmcReport.Move ExptEnco.Width / 2 - cmcReport.Width / 2 + cmcReport.Width
    imBSMode = False
    imCalType = 0   'Standard
    mInitBox

    slStr = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(slStr)
    slStr = gObtainNextMonday(slStr)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    pbcCalendarTo_Paint
    lacDate.Visible = False
    lacDateTo.Visible = False
    gCenterStdAlone ExptEnco
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
    
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
    Dim ilLoop As Integer
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    plcCalendar.Move edcStartDate.Left, edcStartDate.Top + edcStartDate.Height
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
    'On Error GoTo mOpenMsgFileErr:
    ''slToFile = sgExportPath & "ExptEnco.Txt"
    '1/10/19" add user name to message folder name
    'slToFile = sgDBPath & "Messages\" & "ExptEnco.Txt"
    If Trim$(tgUrf(0).sRept) <> "" Then
        slToFile = sgDBPath & "Messages\" & "ExptEnco_" & Trim$(tgUrf(0).sRept) & ".Txt"
    Else
        slToFile = sgDBPath & "Messages\" & "ExptEnco_" & sgUserName & ".Txt"
    End If
    sgMessageFile = slToFile
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        If gDateValue(slFileDate) = lmNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            'ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Output As hmMsg
            'ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        'ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    'Print #hmMsg, ""
    
    'Print #hmMsg, "** Export Dalet Systems: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    'Print #hmMsg, "** Export Enco :" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    gAutomationAlertAndLogHandler "** Export Enco :" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function
Private Sub mSetCommands()
Dim ilEnabled As Integer
Dim ilLoop As Integer
    ilEnabled = False
    'at least one vehicle must be selected
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            ilEnabled = True
            Exit For
        End If
    Next ilLoop
    If ilEnabled Then
        ilEnabled = False
        If edcStartDate <> "" Then
            ilEnabled = True
        End If
    End If
    cmcExport.Enabled = ilEnabled
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
    Unload ExptEnco
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
    Dim ilVff As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    
    ilRet = gPopUserVehicleBox(ExptEnco, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHSPORTMINUELIVE + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)

    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ExptEnco
        On Error GoTo 0
    End If
    'Select on all vehicles that have Clearance as the format
    'For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
    '    slNameCode = Traffic!lbcUserVehicle.List(ilLoop)
    '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '    tmVefSrchKey.iCode = Val(slCode)
    '   ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    '    If (ilRet = BTRV_ERR_NONE) Then
    '        If StrComp(Trim$(tmVef.sFormat), "Clearance", 1) = 0 Then
    '            lbcVehicle.Selected(ilLoop) = True
    '        End If
    '    End If
    'Next ilLoop

    For ilLoop = LBound(tgUserVehicle) To UBound(tgUserVehicle) - 1 Step 1
        slNameCode = tgUserVehicle(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If ilVefCode = tgVff(ilVff).iVefCode Then
                If tgVff(ilVff).sExportEnco = "Y" Then
                    lbcVehicle.Selected(ilLoop) = True
                End If
                Exit For
            End If
        Next ilVff
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
    mBoxCalDate edcStartDate, lacDate
End Sub

Private Sub pbcCalendarTo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
                edcEndDate.Text = Format$(llDate, "m/d/yy")
                edcEndDate.SelStart = 0
                edcEndDate.SelLength = Len(edcEndDate.Text)
                imBypassFocus = True
                edcEndDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcEndDate.SetFocus
End Sub
Private Sub pbcCalendarTo_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalNameTo.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendarTo, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate edcEndDate, lacDateTo
End Sub


Private Sub tmcCancel_Timer()
    tmcCancel.Enabled = False       'screen has now been focused to show
    cmcCancel_Click         'simulate clicking of cancen button
End Sub
'
'       mCreateEncoText - tmOdfExt array contains all the spots to create
'       the days export , 1 vehicle, 1day
'       <return> imSpotFound :true if at least one spot found
'                 false if no spots found.  The export file is erased
'
Public Function mCreateEncoText() As Integer


Dim slRecord As String
Dim ilLoopODF As Integer
Dim slCopy As String
Dim slAdvt As String
Dim slProd As String
Dim slCreative As String
Dim slLen As String
Dim ilRet As Integer
Dim tlOdf As ODF
Dim llSpotTime As Long
Dim slMsg As String
Dim ilType As Integer        '0 = no error, 1 = cart # missing, 2 = creative title missing, 3 = advt missing
Dim slISCI As String
Dim llPrevAvailTime As Long
Dim llCurrAvailTime As Long
Dim llNextStartTime As Long
Dim slAvailSTime As String
Dim slAvailETime As String
Dim ilfirstTime As Integer
Dim slXMid As String
Dim slSpotTime As String

    mCreateEncoText = True
    imFoundspot = False
    ilType = 0                  'assume no errors found in retrieval of data
    ilfirstTime = True
    For ilLoopODF = LBound(tmOdfExt) To UBound(tmOdfExt) - 1
        imFoundspot = True          'found at least 1 spot, do not remove the file due to empty
        tlOdf = tmOdfExt(ilLoopODF)

        gUnpackTimeLong tlOdf.iFeedTime(0), tlOdf.iFeedTime(1), False, llSpotTime

        slRecord = ""
        slSpotTime = gFormatSpotTime(llSpotTime)     'obtain time time as string as HH:MM:SS and extract whats needed
        slRecord = slRecord & Left(slSpotTime, 8)    'spot time in military: pos 1 - 8
        slRecord = slRecord & "  "              '2 blanks
        'slRecord = slRecord & "                " '16 blanks:pos 9 -24
        slCopy = ""
        slCreative = ""
        slISCI = ""
        tmCifSrchKey0.lCode = tlOdf.lCifCode     'copy code
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            slCopy = tmCif.sName            'cart # only (no media definition)
            tmCpfSrchKey0.lCode = tmCif.lcpfCode     'product/isci/creative title
            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                slCreative = tmCpf.sCreative
                slISCI = tmCpf.sISCI
            Else                        'Creative Title missing
                ilType = 2
            End If
        Else                            'copy missing
            ilType = 1
        End If
        Do While Len(slCopy) < 5
            slCopy = slCopy & " "
        Loop

        Do While Len(slCreative) < 30
            slCreative = slCreative & " "
        Loop
        Do While Len(slISCI) < 20
            slISCI = slISCI & " "
        Loop

        'slRecord = slRecord & Left$(slCopy, 5)      'cart #: pos 25 - 29
        'slRecord = slRecord & "   "          '3 blanks: pos 30-32
        slRecord = slRecord & slISCI          'ISCI Pos 11-30
        slRecord = slRecord & "  "            '2 blanks pos 31 -32

        slAdvt = ""
        tmAdfSrchKey.iCode = tlOdf.iAdfCode
        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            slAdvt = Left(tmAdf.sName, 5)
        Else
            ilType = 3              'advt missing
        End If
        Do While Len(slAdvt) < 5
            slAdvt = slAdvt & " "
        Loop
        slProd = Left(tlOdf.sProduct, 5)
        Do While Len(slProd) < 5
            slProd = slProd & " "
        Loop
        slRecord = slRecord & slAdvt & "/" & slProd     'advt/prod: pos 33-43
        slRecord = slRecord & "     "            '5 blanks: pos 44-48
        slRecord = slRecord & slCreative         'creative title (30): Pos 49-78
        slRecord = slRecord & "  :"               '2 blanks + ":" Pos 79-81
        gUnpackLength tlOdf.iLen(0), tlOdf.iLen(1), "1", True, slLen        'length Pos 82-83
        Do While Len(slLen) < 2
            slLen = slLen & " "
        Loop
        slRecord = slRecord & slLen
        On Error GoTo cmcExportErr:

        Print #hmTo, Left(slRecord, 83)
        On Error GoTo 0
        If ilRet <> 0 Then
            imExporting = False
            'Print #hmMsg, "** Terminated **"
            gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Close #hmTo
            Screen.MousePointer = vbDefault
            ''MsgBox "Error writing to " & sgDBPath & "Messages\" & "ExptEnco.Txt" & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
            gAutomationAlertAndLogHandler "Error writing to " & sgDBPath & "Messages\" & "ExptEnco.Txt" & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
            cmcCancel.SetFocus
            mCreateEncoText = False
            Exit Function
        End If
        mGenerateErrorMsg tlOdf, ilType

        'update running time of spots within a break for the log report (pdf)
        'It needs to show which spots feed are at the same times
        tmOdfSrchKey.iVefCode = tlOdf.iVefCode
        tmOdfSrchKey.iAirDate(0) = tlOdf.iAirDate(0)
        tmOdfSrchKey.iAirDate(1) = tlOdf.iAirDate(1)
        tmOdfSrchKey.sZone = tlOdf.sZone
        tmOdfSrchKey.iSeqNo = tlOdf.iSeqNo
        tmOdfSrchKey.iLocalTime(0) = tlOdf.iLocalTime(0)
        tmOdfSrchKey.iLocalTime(1) = tlOdf.iLocalTime(1)
        ilRet = btrGetEqual(hmOdf, tlOdf, Len(tlOdf), tmOdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)

        gUnpackTimeLong tlOdf.iFeedTime(0), tlOdf.iFeedTime(1), False, llCurrAvailTime
        If ilfirstTime Then
            ilfirstTime = False
            llPrevAvailTime = llCurrAvailTime
            gUnpackLength tlOdf.iLen(0), tlOdf.iLen(1), "3", False, slLen
            gUnpackTime tlOdf.iFeedTime(0), tlOdf.iFeedTime(1), "A", "1", slAvailSTime
            gAddTimeLength slAvailSTime, slLen, "A", "1", slAvailETime, slXMid
            llNextStartTime = gTimeToLong(slAvailETime, True)
        Else
            If llPrevAvailTime = llCurrAvailTime Then      'same break
                'same break, keep running time
                gPackTimeLong llNextStartTime, tlOdf.iFeedTime(0), tlOdf.iFeedTime(1)
                gUnpackLength tlOdf.iLen(0), tlOdf.iLen(1), "3", False, slLen
                gUnpackTime tlOdf.iFeedTime(0), tlOdf.iFeedTime(1), "A", "1", slAvailSTime
                gAddTimeLength slAvailSTime, slLen, "A", "1", slAvailETime, slXMid
                llNextStartTime = gTimeToLong(slAvailETime, True)
                gPackTime slAvailSTime, tlOdf.iFeedTime(0), tlOdf.iFeedTime(1)
            Else                                        'different break
                llPrevAvailTime = llCurrAvailTime
                gUnpackLength tlOdf.iLen(0), tlOdf.iLen(1), "3", False, slLen
                gUnpackTime tlOdf.iFeedTime(0), tlOdf.iFeedTime(1), "A", "1", slAvailSTime
                gAddTimeLength slAvailSTime, slLen, "A", "1", slAvailETime, slXMid
                llNextStartTime = gTimeToLong(slAvailETime, True)
                gPackTime slAvailSTime, tlOdf.iFeedTime(0), tlOdf.iFeedTime(1)
            End If
        End If

        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrUpdate(hmOdf, tlOdf, Len(tlOdf))
            If ilRet <> BTRV_ERR_NONE Then
                'error, cant udpate the ODF record with time
                slMsg = "Update Error #" & Trim$(str(ilRet)) & " in " & smVehName & " for " & slAdvt & " @ " & slSpotTime
                'Print #hmMsg, slMsg
                gAutomationAlertAndLogHandler slMsg
                lbcMsg.AddItem slMsg

            End If
        Else
            'error, cant find the ODF record that was created
            slMsg = "BtrGetEqual Error #" & Trim$(str(ilRet)) & " in " & smVehName & " for " & slAdvt & " @ " & slSpotTime
            'Print #hmMsg, slMsg
            gAutomationAlertAndLogHandler slMsg
            lbcMsg.AddItem slMsg
        End If

    Next ilLoopODF
    Exit Function
cmcExportErr:
    ilRet = err.Number
    Resume Next

End Function
'
'              mGenerateErrorMsg - Generate error messages in Msg box on screen
'                                   and in Message file
'
'               <input> tlOdf - Spot image from ODF
'                       ilType - type of error:  1 = copy (cart # missing), 2 = Creative Title missing
'                               3 = Advertiser missing
Private Sub mGenerateErrorMsg(tlOdf As ODF, ilType As Integer)
Dim slTime As String
Dim slMsg As String

    gUnpackTime tlOdf.iFeedTime(0), tlOdf.iFeedTime(1), "A", "1", slTime
    If ilType = 1 Then              'copy missing
        slMsg = "Copy Missing for " & Trim$(tmAdf.sName) & " at Feed Time " & slTime & " on " & smVehName
    ElseIf ilType = 2 Then          'creative title missing
        slMsg = "Creative Title of ISCI missing for " & Trim$(tmAdf.sName) & " at FeedTime " & slTime & " on " & smVehName
    ElseIf ilType = 3 Then          'advt missing
        slMsg = "Advertiser Missing " & " at Feed Time " & slTime & " on " & smVehName
    Else                            'show no message, export record OK
        Exit Sub
    End If
    'Print #hmMsg, slMsg
    gAutomationAlertAndLogHandler slMsg
    lbcMsg.AddItem slMsg
    Exit Sub
End Sub
