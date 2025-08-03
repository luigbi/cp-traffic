VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExptAirWave 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5610
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
   ScaleHeight     =   5610
   ScaleWidth      =   9135
   Begin V81TrafficExports.CSI_Calendar cccDate 
      Height          =   285
      Left            =   1770
      TabIndex        =   13
      Top             =   345
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   503
      Text            =   "7/12/2022"
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   -1  'True
      CSI_InputBoxBoxAlignment=   0
      CSI_CalBackColor=   16777130
      CSI_CalDateFormat=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CSI_CurDayBackColor=   16777215
      CSI_CurDayForeColor=   0
      CSI_ForceMondaySelectionOnly=   -1  'True
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   2
   End
   Begin VB.CommandButton cmcGetPath 
      Appearance      =   0  'Flat
      Caption         =   "&Browse..."
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
      Left            =   7680
      TabIndex        =   16
      Top             =   750
      Width           =   1245
   End
   Begin VB.TextBox edcGetPath 
      Height          =   300
      Left            =   1515
      TabIndex        =   15
      Top             =   750
      Width           =   6030
   End
   Begin VB.ListBox lbcSort 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "ExptAirWave.frx":0000
      Left            =   6000
      List            =   "ExptAirWave.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5250
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   165
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4140
      Width           =   1410
   End
   Begin VB.Timer tmcCancel 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2475
      Top             =   5115
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1350
      Width           =   5235
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   1545
      Top             =   5055
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
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
      ItemData        =   "ExptAirWave.frx":0014
      Left            =   165
      List            =   "ExptAirWave.frx":0016
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   1350
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
      TabIndex        =   3
      Top             =   5145
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
      TabIndex        =   4
      Top             =   5145
      Width           =   1050
   End
   Begin VB.Label lacGetPath 
      Appearance      =   0  'Flat
      Caption         =   "Export File To"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   14
      Top             =   795
      Width           =   1395
   End
   Begin VB.Label lacEncoText 
      Appearance      =   0  'Flat
      Caption         =   "Export AirWave"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   1785
   End
   Begin VB.Label lacMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   180
      TabIndex        =   10
      Top             =   4755
      Width           =   8730
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      Top             =   5070
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   195
      TabIndex        =   8
      Top             =   4425
      Width           =   8730
   End
   Begin VB.Label lacStartDate 
      Appearance      =   0  'Flat
      Caption         =   "Export Start Date"
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
      TabIndex        =   7
      Top             =   2370
      Width           =   1725
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   5175
      TabIndex        =   6
      Top             =   1965
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   3225
      TabIndex        =   5
      Top             =   1965
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "ExptAirWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Copyright 1993 Counterpoint Software®. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExptAirWave.Frm
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
Dim smGenNowDate As String
Dim smGenNowTime As String

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

'Contract record information
Dim hmCHF As Integer        'Contract header file handle
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim imCHFRecLen As Integer  'CHF record length
Dim tmChf As CHF            'CHF record image

'Csf Type
Dim hmCsf As Integer        'comment for other type
Dim tmCsf As CSF
Dim imCsfRecLen As Integer
Dim tmCsfSrchKey0 As LONGKEY0

Dim hmPrf As Integer            'Product file handle
Dim tmPrfSrchKey As PRFKEY1            'PRF record image
Dim imPrfRecLen As Integer        'PRF record length
Dim tmPrf As PRF

Dim hmLcf As Integer            'Log calendar library file handle
Dim tmLcf As LCF               'LCF record image
Dim tmLcfSrchKey As LCFKEY0     'LCF key record image
Dim tmLcfSrchKey1 As LCFKEY1
Dim tmLcfSrchKey2 As LCFKEY2
Dim imLcfRecLen As Integer         'LCF record length

Dim hmLvf As Integer            'Log calendar library file handle
Dim tmLvf As LVF               'LCF record image
Dim tmLvfSrchKey As LONGKEY0     'LCF key record image
Dim imLvfRecLen As Integer         'LCF record length

Dim hmVff As Integer        'Vehicle file handle
Dim tmVff As VFF            'VfF record image
Dim imVffRecLen As Integer     'VfF record length

Dim hmAwf As Integer            'AirWave file handle
Dim tmAwf As AWF               'AWF record image
Dim tmAwfSrchKey1 As AWFKEY1
Dim imAwfRecLen As Integer         'AWF record length

Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control

Dim tmPPMnf() As MNF

Dim imEvtType(0 To 14) As Integer

Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)

Dim lmSpotSeqNo As Long

Dim imFoundspot As Integer  '1-5-05 flag to indicate if at least one spot found; if not, dont retain export file

Dim smDelimiter As String

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
Sub mReadODFSpots(hlODF As Integer, ilVefCode As Integer, slZone As String, slSDate As String, slEDate As String)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slTime                        tlodfSrchKey2                 ilAirDate0                *
'*  ilAirDate1                                                                            *
'******************************************************************************************

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
    gPackDate slSDate, tlOdfSrchKey0.iAirDate(0), tlOdfSrchKey0.iAirDate(1)
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

    If gDateValue(slSDate) = gDateValue(slEDate) Then
        gPackDate slSDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffsetExtra("ODF", "OdfAirDate")
        ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
    Else
        gPackDate slSDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffsetExtra("ODF", "OdfAirDate")
        ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        gPackDate slEDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffsetExtra("ODF", "OdfAirDate")
        ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
    End If
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

Private Sub cmcCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcExport_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStartDate                   slEndDate                     ilDays                    *
'*  ilLine                        ilRecdLen                     ilVef                     *
'*  slTime                        ilRec                         llZoneEndTimes            *
'*  llPDFDate                                                                             *
'******************************************************************************************
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
    Dim slSDate As String
    Dim slEDate As String
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
    Dim slExportPath As String
    Dim slAwfSeqNo As String
    Dim slAllOrPartial As String

    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    
    lacProcessing.Caption = ""
    lacMsg.Caption = ""
    slInputStartDate = cccDate.Text
    If Not gValidDate(slInputStartDate) Then
        Beep
        cccDate.SetFocus
        Exit Sub
    End If
    llStartDate = gDateValue(slInputStartDate)
    llEndDate = llStartDate + 6
    slInputEndDate = DateAdd("d", 6, slInputStartDate)

    imExporting = True
    lbcMsg.Clear
    If Not mOpenMsgFile() Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    lmSpotSeqNo = 0
    smGenNowDate = Format(gNow(), "m/d/yy")
    smGenNowTime = Format(gNow(), "h:mm:ssa/p")
    'creating log events
    lacProcessing.Caption = "Generating log days " & slInputStartDate & "-" & slInputEndDate

    mReadAWF slInputStartDate
    tmAwf.iSeqNo = tmAwf.iSeqNo + 1
    If tmAwf.iSeqNo <= 9 Then
        slAwfSeqNo = "0" & Trim$(str(tmAwf.iSeqNo))
    Else
        slAwfSeqNo = Trim$(str(tmAwf.iSeqNo))
    End If
    'slFileName = right$(slFYear, 2) & slFMonth & slFDay & (slZoneAbbrev)
    slAllOrPartial = "A"
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If Not lbcVehicle.Selected(ilLoop) Then
            slAllOrPartial = "P"
            Exit For
        End If
    Next ilLoop
    
    slFileName = Format$(slInputStartDate, "mmddyy") & slAllOrPartial & Format$(slInputEndDate, "mmddyy") & "V" & slAwfSeqNo
    slExt = ".psv"
    ilRet = 0
    'On Error GoTo cmcExportErr:

    slExportPath = edcGetPath.Text
    If slExportPath = "" Then
        slExportPath = sgExportPath
    End If
    slExportPath = gSetPathEndSlash(slExportPath, True)
    
    slToFile = slExportPath & Trim$(slFileName) & slExt   'ssssZmmdd.ext

    'slDateTime = FileDateTime(slToFile)     '1-6-05 chged from sgExportPath to new sgProphetExportPath
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        'Kill slToFile   '1-6-05 chg from sgExportpath to new sgProphetExportPath
        imExporting = False
        Screen.MousePointer = vbDefault
        'Print #hmMsg, "** Terminated **"
        gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        'Print #hmMsg, "Export currently being run by another user for the same week"
        gAutomationAlertAndLogHandler "Export currently being run by another user for the same week"
        Close #hmMsg
        ''MsgBox "AirWave export currently being generated by another User for the same week, Export disallowed"
        gAutomationAlertAndLogHandler "AirWave export currently being generated by another User for the same week, Export disallowed", vbOkOnly + vbCritical + vbApplicationModal, "Export Error"
        Exit Sub
    End If
    On Error GoTo 0


    ilRet = 0
    'On Error GoTo cmcExportErr:
    'hmTo = FreeFile
    'Open slToFile For Output As hmTo
    ilRet = gFileOpen(slToFile, "Output", hmTo)
    If ilRet <> 0 Then
        imExporting = False
        Screen.MousePointer = vbDefault
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
    
    'Print Header
    Print #hmTo, "Spot #" & smDelimiter & "Date stamp" & smDelimiter & "Time stamp" & smDelimiter & "Vehicle" & smDelimiter & "CSI vehicle ID" & smDelimiter & "Airwave Program Code" _
    & smDelimiter & "Program Start Date" & smDelimiter & "Program End date" & smDelimiter & "Program Start Time" & smDelimiter & "Program Days" & smDelimiter & "Break #" _
    & smDelimiter & "Position #" & smDelimiter & "Break Code" & smDelimiter & "Length" & smDelimiter & "Scheduled Date" & smDelimiter & "Scheduled Time" & smDelimiter & "Copy Inventory ID" _
    & smDelimiter & "Copy Creative Title" & smDelimiter & "Copy Rotation Comments" & smDelimiter & "ISCII Code" & smDelimiter & "Copy Inventory Start Date" _
    & smDelimiter & "Copy Inventory End Date" & smDelimiter & "Advertiser Name" & smDelimiter & "Advertiser Code" & smDelimiter & "Product Name" & smDelimiter & "Product Code" _
    & smDelimiter & "Primary Product Protection Name" & smDelimiter & "Primary Product Protection ID" & smDelimiter & "Secondary Product Protection Name" & smDelimiter & "Secondary Product Protection ID"
    
    
    'get current date and time for filtering of odf records
    gCurrDateTime slPDFDate, slPDFTime, slMonth, slDay, slYear    'get system date and time for prepass filtering
    
    ilRet = gGenODFDay(ExptAirWave, slInputStartDate, slInputEndDate, lbcVehicle, lbcSort, tgUserVehicle(), "W", "E")

    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If imTerminate Then
            Exit For
        End If
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
                ilVpfIndex = gVpfFind(ExptAirWave, ilVefCode)              'determine vehicle options index
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
                        'EST only
                        ilZones = 1
                        ilLoZone = 1
                        ilHiZone = 1
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
                    slSDate = Format(llDate, "m/d/yy")
                    If llDate = llEndDate Then
                        slEDate = Format(llDate + 1, "m/d/yy")
                    Else
                        slEDate = slSDate
                    End If
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

                            'Print #hmMsg, " "
                            gAutomationAlertAndLogHandler " "
                            'Print #hmMsg, "** Generating Data for " & Trim$(smVehName) & " for " & slSDate & " **"
                            gAutomationAlertAndLogHandler "** Generating Data for " & Trim$(smVehName) & " for " & slSDate & " **"
                            lacProcessing.Caption = "Generating Data for " & Trim$(smVehName) & " for " & slSDate

                            Select Case ilDayOfWeek
                                Case 0
                                    mReadODFSpots hmOdf, ilVefCode, slZone, slSDate, slEDate
                                Case 1
                                    mReadODFSpots hmOdf, ilVefCode, slZone, slSDate, slEDate
                                Case 2
                                    mReadODFSpots hmOdf, ilVefCode, slZone, slSDate, slEDate
                                Case 3
                                    mReadODFSpots hmOdf, ilVefCode, slZone, slSDate, slEDate
                                Case 4
                                    mReadODFSpots hmOdf, ilVefCode, slZone, slSDate, slEDate
                                Case 5
                                    mReadODFSpots hmOdf, ilVefCode, slZone, slSDate, slEDate
                                Case 6
                                    mReadODFSpots hmOdf, ilVefCode, slZone, slSDate, slEDate
                            End Select

                            'Determine file name from date: SSSSZMMDD.txt
                            'where SSSS = vehicle station code (max 5 char)
                            'Z = zone (W = PST, E = EST, C = CST, M = MST)
                            'MM = 2 char month
                            'DD = 2 char day
                            'Print #hmMsg, "** Storing Output into " & slToFile & " **"
                            lacProcessing.Caption = "Writing Data to " & Trim$(slFileName) & slExt
                            If Not mCreateAirWaveText(ilVefCode) Then      'gather matching spot events in ODF and create the export record for day/vehicle
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

                            DoEvents
                            ilRet = 0
                            'On Error GoTo cmcExportErr:

                            If imFoundspot Then         '1-5-05 if no spots found, dont retain export file and dont show file sent to message
                                lacProcessing.Caption = "Output for " & smVehName & " added to " & slToFile
                                'Print #hmMsg, "** Output for " & smVehName & " added to " & slToFile & " **"
                                gAutomationAlertAndLogHandler "** Output for " & smVehName & " added to " & slToFile & " **"
                            Else
                                lacProcessing.Caption = "No spots found for " & smVehName & " on " & slSDate
                                'Unable to kill file as file is open plus it might have spots from different vehicle.
                                'Kill (slToFile)
                                'Print #hmMsg, "No spots found for " & smVehName & " on " & slSDate
                                gAutomationAlertAndLogHandler "No spots found for " & smVehName & " on " & slSDate
                            End If
                        End If
                    Next ilZoneLoop         'next zone
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
    Close #hmTo         'close the output file

    gClearODF                   'remove all the ODFs for the logs just created
    
    'Create Control File
    slFileName = Format$(slInputStartDate, "mmddyy") & "C" & Format$(slInputEndDate, "mmddyy") & "V" & slAwfSeqNo
    slExt = ".psv"
    ilRet = 0
    'On Error GoTo cmcExportErr:

    slToFile = slExportPath & Trim$(slFileName) & slExt   'ssssZmmdd.ext

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
    Else
        Print #hmTo, smGenNowDate & smDelimiter & smGenNowTime & smDelimiter & smGenNowDate & smDelimiter & smGenNowTime & smDelimiter & smGenNowDate & smDelimiter & smGenNowTime & smDelimiter & lmSpotSeqNo
        Close #hmTo
    End If
    
    If imTerminate Then
        'Print #hmMsg, "** Completed AirWave Export: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        gAutomationAlertAndLogHandler "** Completed AirWave Export: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    Else
        'Print #hmMsg, "** Terminated AirWave Export: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        gAutomationAlertAndLogHandler "** Terminated AirWave Export: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    End If
    Close #hmMsg
    
    'Print #hmMsg, "** Completed " & smScreenCaption & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    Close #hmMsg
    If Not imTerminate Then
        ilRet = mUpdateAwf()
    End If
    On Error GoTo 0
    lacMsg.Caption = "Messages sent to " & sgDBPath & "Messages\" & "ExptAirWave.Txt"
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
Private Sub cmcGetPath_Click()
    Dim slCurDir As String
    
    slCurDir = CurDir
    igPathType = 0
    sgGetPath = edcGetPath.Text
    lgCallTop = ExptAirWave.Top
    lgCallLeft = ExptAirWave.Left
    lgCallHeight = ExptAirWave.Height
    lgCallWidth = ExptAirWave.Width
    GetPath.Show vbModal
    If igGetPath = 0 Then
        edcGetPath.Text = sgGetPath
    End If
    
    ChDir slCurDir
    
    Exit Sub
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
    'cccDate.Visible = False
    DoEvents    'Process events so pending keys are not sent to this
    'cccDate.Visible = True
    Me.KeyPreview = True
    mSetCommands
    'Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
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
    
    Erase tmPPMnf
    Erase tmOdfExt
    

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
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmCsf)
    btrDestroy hmCsf
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmLvf)
    btrDestroy hmLvf
    ilRet = btrClose(hmVff)
    btrDestroy hmVff
    ilRet = btrClose(hmAwf)
    btrDestroy hmAwf
    ilRet = btrClose(hmOdf)
    btrDestroy hmOdf
    Set ExptAirWave = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub



Private Sub lbcVehicle_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked  '9-12-02 False
        imSetAll = True
    End If
    mSetCommands
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
    Dim slMnfStamp As String
    
    imTerminate = False
    imFirstActivate = True
    'mParseCmmdLine
    Screen.MousePointer = vbHourglass
    smDelimiter = "|"
    imAllClicked = False
    imSetAll = True
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False

    hmOdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmOdf, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptAirWave
    On Error GoTo 0
    imOdfRecLen = Len(tmTOdf)

    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptAirWave
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)

    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptAirWave
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)

    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptAirWave
    On Error GoTo 0
    imCifRecLen = Len(tmCif)

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptAirWave
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    
    hmVpf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmVpf, "", sgDBPath & "Vpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptAirWave
    On Error GoTo 0
    imVpfRecLen = Len(tmVpf)

   
    hmCHF = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptAirWave
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
  
    hmCsf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmCsf, "", sgDBPath & "Csf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptAirWave
    On Error GoTo 0
    imCsfRecLen = Len(tmCsf)

    hmPrf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptAirWave
    On Error GoTo 0
    imPrfRecLen = Len(tmPrf)

    hmLcf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptAirWave
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf)

    hmLvf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmLvf, "", sgDBPath & "Lvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptAirWave
    On Error GoTo 0
    imLvfRecLen = Len(tmLvf)

    hmVff = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmVff, "", sgDBPath & "Vff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptAirWave
    On Error GoTo 0
    imVffRecLen = Len(tmVff)

    hmAwf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmAwf, "", sgDBPath & "Awf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptAirWave
    On Error GoTo 0
    imAwfRecLen = Len(tmAwf)

    'Populate arrays to determine if records exist
    mVehPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        'mTerminate
        Exit Sub
    End If

    slMnfStamp = ""
    'ReDim tmPPMnf(1 To 1) As MNF
    ReDim tmPPMnf(0 To 0) As MNF
    ilRet = gObtainMnfForType("C", slMnfStamp, tmPPMnf())

    '4/26/11: Add test of avail attribute
    ilRet = gObtainAvail()
    
    For ilLoop = LBound(imEvtType) To UBound(imEvtType) Step 1
        imEvtType(ilLoop) = True
    Next ilLoop
    imEvtType(0) = False 'Don't include library names
    'plcGauge.Move ExptAirWave.Width / 2 - plcGauge.Width / 2
    'cmcFileConv.Move ExptAirWave.Width / 2 - cmcFileConv.Width / 2
    'cmcCancel.Move ExptAirWave.Width / 2 - cmcCancel.Width / 2 - cmcCancel.Width
    'cmcReport.Move ExptAirWave.Width / 2 - cmcReport.Width / 2 + cmcReport.Width
    imBSMode = False
    mInitBox

    edcGetPath.Text = sgExportPath
    slStr = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(slStr)
    slStr = gObtainNextMonday(slStr)
    cccDate.Text = slStr
    gCenterStdAlone ExptAirWave
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
    ''slToFile = sgExportPath & "ExptAirWave.Txt"
    slToFile = sgDBPath & "Messages\" & "ExptAirWave.Txt"
    sgMessageFile = slToFile
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = FileDateTime(slToFile)
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
    'Print #hmMsg, "** Export AirWave :" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    gAutomationAlertAndLogHandler "** Export AirWave :" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
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
        If cccDate.Text <> "" Then
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

    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExptAirWave
    'Set ExptAirWave = Nothing   'Remove data segment
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
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim ilVefCode As Integer
    Dim slName As String
    Dim ilVff As Integer
    Dim blRetain As Boolean
    Dim ilLoop As Integer
    Dim slCode As String
    Dim llLen As Long
    
    ilRet = gPopUserVehicleBox(ExptAirWave, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHSPORTMINUELIVE + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)

    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ExptAirWave
        On Error GoTo 0
    End If
    ilIndex = UBound(tgUserVehicle) - 1
    Do
        slNameCode = tgUserVehicle(ilIndex).sKey    'Traffic!lbcUserVehicle.List(ilIndex)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 3, "|", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        blRetain = False
        'For ilVff = LBound(tgVff) To UBound(tgVff) - 1 Step 1
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If ilVefCode = tgVff(ilVff).iVefCode Then
                If tgVff(ilVff).sExportAirWave = "Y" Then
                    blRetain = True
                End If
                Exit For
            End If
        Next ilVff
        If Not blRetain Then
            For ilLoop = ilIndex To UBound(tgUserVehicle) - 1 Step 1
                tgUserVehicle(ilLoop) = tgUserVehicle(ilLoop + 1)
            Next ilLoop
            ReDim Preserve tgUserVehicle(LBound(tgUserVehicle) To UBound(tgUserVehicle) - 1) As SORTCODE
        End If
        ilIndex = ilIndex - 1
    Loop While ilIndex >= LBound(tgUserVehicle)
    
    llLen = 0
    lbcVehicle.Clear
    For ilLoop = 0 To UBound(tgUserVehicle) - 1 Step 1
        slNameCode = tgUserVehicle(ilLoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            Exit Sub
        End If
        ilRet = gParseItem(slName, 3, "|", slName)
        If ilRet <> CP_MSG_NONE Then
            Exit Sub
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcVehicle.AddItem slName  'Add ID to list box
    Next ilLoop
    'Select all
    ckcAll.Value = vbChecked
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub



Private Sub tmcCancel_Timer()
    tmcCancel.Enabled = False       'screen has now been focused to show
    cmcCancel_Click         'simulate clicking of cancen button
End Sub
'
'       mCreateAirWaveText - tmOdfExt array contains all the spots to create
'       the days export , 1 vehicle, 1day
'       <return> imSpotFound :true if at least one spot found
'                 false if no spots found.  The export file is erased
'
Public Function mCreateAirWaveText(ilVefCode As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         llSpotime                     slISIC                    *
'*  slTime                                                                                *
'******************************************************************************************

    Dim slRecord As String
    Dim ilLoopODF As Integer
    Dim ilRet As Integer
    Dim tlOdf As ODF
    Dim llSpotTime As Long

    Dim slPrgDate As String
    Dim slPrgTime As String
    Dim llAirTime As Long
    Dim llAvailTime As Long
    Dim llRunLength As Long
    Dim ilBreakNo As Integer
    Dim ilPositionNo As Integer
    Dim slISCI As String
    Dim slRotStartDate As String
    Dim slRotEndDate As String
    Dim slLength As String
    Dim ilIndex As Integer
    Dim llPrfCode As Long
    Dim slCreativeTitle As String
    Dim slAirDate As String
    
    On Error GoTo mCreateAirWaveTextErr

    mCreateAirWaveText = True
    imFoundspot = False
    llAvailTime = -1
    ilBreakNo = 0
    For ilLoopODF = LBound(tmOdfExt) To UBound(tmOdfExt) - 1
        imFoundspot = True          'found at least 1 spot, do not remove the file due to empty
        tlOdf = tmOdfExt(ilLoopODF)

        gUnpackTimeLong tlOdf.iAirTime(0), tlOdf.iAirTime(1), False, llAirTime
        If llAvailTime <> llAirTime Then
            ilBreakNo = ilBreakNo + 1
            ilPositionNo = 1
            llRunLength = 0
            llAvailTime = llAirTime
        Else
            ilPositionNo = ilPositionNo + 1
        End If
        'gUnpackTimeLong tlOdf.iFeedTime(0), tlOdf.iFeedTime(1), False, llSpotTime

        slRecord = ""
        lmSpotSeqNo = lmSpotSeqNo + 1
        slRecord = Trim$(str$(lmSpotSeqNo)) & smDelimiter
        slRecord = slRecord & smGenNowDate & smDelimiter
        slRecord = slRecord & smGenNowTime & smDelimiter
        slRecord = slRecord & smVehName & smDelimiter
        slRecord = slRecord & Trim$(str$(tmVef.iCode)) & smDelimiter
        'AirWave program ID
        slRecord = slRecord & Trim$(mGetAirWavePrgID(ilVefCode)) & smDelimiter
        'Program start date
        If tlOdf.iGameNo > 0 Then
            'Determine true program date
            gUnpackDate tlOdf.iAirDate(0), tlOdf.iAirDate(1), slPrgDate
            ilRet = mFindGameDate(tlOdf.iGameNo, slPrgDate, slPrgTime)
            If Not ilRet Then
                gUnpackDate tlOdf.iAirDate(0), tlOdf.iAirDate(1), slPrgDate
                slPrgTime = mFindLibraryStartTime(slPrgDate)
            End If
        Else
            gUnpackDate tlOdf.iAirDate(0), tlOdf.iAirDate(1), slPrgDate
            slPrgTime = mFindLibraryStartTime(slPrgDate)
        End If
        slRecord = slRecord & Format(DateAdd("d", -7, gObtainPrevMonday(slPrgDate)), "m/d/yy") & smDelimiter
        'Program end date
        slRecord = slRecord & Format(DateAdd("d", 7, gObtainNextSunday(slPrgDate)), "m/d/yy") & smDelimiter
        'Program start time
        slRecord = slRecord & slPrgTime & smDelimiter
        gUnpackDate tlOdf.iAirDate(0), tlOdf.iAirDate(1), slAirDate
        'Program day
        slRecord = slRecord & Left$(Format(slAirDate, "ddd"), 2) & smDelimiter
        'Break #
        slRecord = slRecord & Trim$(str$(ilBreakNo)) & smDelimiter
        'Position #
        slRecord = slRecord & Trim$(str$(ilPositionNo)) & smDelimiter
        'Break Type (C=Commercial; B=BB)
        If Trim$(tlOdf.sBBDesc) = "" Then
            slRecord = slRecord & "C" & smDelimiter
        Else
            slRecord = slRecord & "B" & smDelimiter
        End If
        'Spot Length
        gUnpackLength tlOdf.iLen(0), tlOdf.iLen(1), "1", True, slLength
        slRecord = slRecord & slLength & smDelimiter
        'Schedule date
        slRecord = slRecord & Format(slAirDate, "m/d/yy") & smDelimiter
        'Schedule Time
        llSpotTime = llAvailTime + llRunLength
        llRunLength = llRunLength + Val(slLength)
        slRecord = slRecord & Format(gFormatTimeLong(llSpotTime, "A", "1"), "h:mm:ssA/P") & smDelimiter
        'Copy Inventory code
        slRecord = slRecord & Trim$(str(tlOdf.lCifCode)) & smDelimiter
        'Creative Title
        slCreativeTitle = ""
        slISCI = ""
        slRotStartDate = ""
        slRotEndDate = ""
        tmCifSrchKey0.lCode = tlOdf.lCifCode     'copy code
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            gUnpackDate tmCif.iRotStartDate(0), tmCif.iRotStartDate(1), slRotStartDate
            gUnpackDate tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), slRotEndDate
            tmCpfSrchKey0.lCode = tmCif.lcpfCode     'product/isci/creative title
            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                slISCI = Trim$(tmCpf.sISCI)
                slCreativeTitle = Trim$(tmCpf.sCreative)
            End If
        End If
        slRecord = slRecord & slCreativeTitle & smDelimiter
        'Rotation comment
        If tlOdf.lCefCode <= 0 Then
            slRecord = slRecord & "" & smDelimiter
        Else
            tmCsfSrchKey0.lCode = tlOdf.lCefCode
            imCsfRecLen = Len(tmCsf)
            ilRet = btrGetEqual(hmCsf, tmCsf, imCsfRecLen, tmCsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                slRecord = slRecord & gStripChr0(tmCsf.sComment) & smDelimiter
            Else
                slRecord = slRecord & "" & smDelimiter
            End If
        End If
        'ISCI
        slRecord = slRecord & slISCI & smDelimiter
        'Inventory start date
        If slRotStartDate = "" Then
            slRecord = slRecord & Format(slPrgDate, "m/d/yy") & smDelimiter
        Else
            slRecord = slRecord & Format(slRotStartDate, "m/d/yy") & smDelimiter
        End If
        'Inventory end date
        If slRotEndDate = "" Then
            slRecord = slRecord & Format(slPrgDate, "m/d/yy") & smDelimiter
        Else
            slRecord = slRecord & Format(slRotEndDate, "m/d/yy") & smDelimiter
        End If
        'Advertiser name
        tmAdfSrchKey.iCode = tlOdf.iAdfCode
        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            slRecord = slRecord & Trim$(tmAdf.sName) & smDelimiter
        Else
            slRecord = slRecord & "Advertiser Missing" & smDelimiter
        End If
        'Advertiser code
        slRecord = slRecord & Trim$(str(tlOdf.iAdfCode)) & smDelimiter
        'Product Name
        slRecord = slRecord & Trim$(tlOdf.sProduct) & smDelimiter
        tmChfSrchKey.lCode = tlOdf.lFt1CefCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            'Product code
            llPrfCode = mFindProductCode(tmChf.iAdfCode, tlOdf.sProduct)
            If llPrfCode <> -1 Then
                slRecord = slRecord & Trim$(str(llPrfCode)) & smDelimiter
            Else
                slRecord = slRecord & "0" & smDelimiter
            End If
            'Product protection
            ilIndex = mFindPP(tmChf.iMnfComp(0))
            If ilIndex <> -1 Then
                slRecord = slRecord & Trim$(tmPPMnf(ilIndex).sName) & smDelimiter
            Else
                slRecord = slRecord & "" & smDelimiter
            End If
            'Product protection code
            slRecord = slRecord & Trim$(str(tmChf.iMnfComp(0))) & smDelimiter
            'Product protection
            ilIndex = mFindPP(tmChf.iMnfComp(1))
            If ilIndex <> -1 Then
                slRecord = slRecord & Trim$(tmPPMnf(ilIndex).sName) & smDelimiter
            Else
                slRecord = slRecord & "" & smDelimiter
            End If
            'Product protection code
            slRecord = slRecord & Trim$(str(tmChf.iMnfComp(1)))
        Else
            'Product code
            slRecord = slRecord & "0" & smDelimiter
            'Product protection
            slRecord = slRecord & "" & smDelimiter
            'Product protection code
            slRecord = slRecord & "0" & smDelimiter
            'Product protection
            slRecord = slRecord & "" & smDelimiter
            'Product protection code
            slRecord = slRecord & "0"
        End If
        
        
        On Error GoTo cmcExportErr:

        Print #hmTo, slRecord
        On Error GoTo 0
        If ilRet <> 0 Then
            imExporting = False
            'Print #hmMsg, "** Terminated **"
            gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Close #hmTo
            Screen.MousePointer = vbDefault
            ''MsgBox "Error writing to " & sgDBPath & "Messages\" & "ExptAirWave.Txt" & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
            gAutomationAlertAndLogHandler "Error writing to " & sgDBPath & "Messages\" & "ExptAirWave.Txt" & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
            cmcCancel.SetFocus
            mCreateAirWaveText = False
            Exit Function
        End If
        On Error GoTo mCreateAirWaveTextErr
        

    Next ilLoopODF
    Exit Function
cmcExportErr:
    ilRet = err.Number
    Resume Next
mCreateAirWaveTextErr:
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

Private Function mFindPP(ilPPCode As Integer) As Integer
    Dim ilLoop As Integer
    
    If ilPPCode > 0 Then
        For ilLoop = LBound(tmPPMnf) To UBound(tmPPMnf) - 1 Step 1
            If tmPPMnf(ilLoop).iCode = ilPPCode Then
                mFindPP = ilLoop
                Exit Function
            End If
        Next ilLoop
    End If
    mFindPP = -1
End Function

Private Function mFindProductCode(ilAdfCode As Integer, slProduct As String) As Long
    Dim ilRet As Integer
    
    tmPrfSrchKey.iAdfCode = tmChf.iAdfCode
    ilRet = btrGetGreaterOrEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
    Do While (ilRet = BTRV_ERR_NONE) And (tmPrf.iAdfCode = ilAdfCode)
        If StrComp(Trim$(Trim$(slProduct)), Trim$(tmPrf.sName), 1) = 0 Then
            mFindProductCode = tmPrf.lCode
            Exit Function
        End If
        ilRet = btrGetNext(hmPrf, tmPrf, imPrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mFindProductCode = -1


End Function

Private Function mFindLibraryStartTime(slDate As String) As String
    Dim ilType As Integer
    Dim ilRet As Integer
    Dim llDate As Long
    
    tmLcfSrchKey2.iVefCode = tmVef.iCode
    gPackDate slDate, tmLcfSrchKey2.iLogDate(0), tmLcfSrchKey2.iLogDate(1)
    ilRet = btrGetEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = tmVef.iCode)
        gUnpackDateLong tmLcf.iLogDate(0), tmLcf.iLogDate(1), llDate
        If gDateValue(slDate) <> llDate Then
            Exit Do
        End If
        If (tmLcf.sStatus = "C") Then
            'Find time
            mFindLibraryStartTime = mScanLcfForTime()
            Exit Function
        End If
        ilRet = btrGetNext(hmLcf, tmLcf, imLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    'See if TFN record exist
    tmLcfSrchKey2.iVefCode = tmVef.iCode
    tmLcfSrchKey2.iLogDate(0) = 0   'ilStartDate(0)
    tmLcfSrchKey2.iLogDate(1) = 0   'ilStartDate(1)
    ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = tmVef.iCode)
        If tmLcf.iLogDate(0) <= 7 And tmLcf.iLogDate(1) = 0 Then
            If (tmLcf.sStatus = "C") Then
                If gWeekDayStr(slDate) + 1 = tmLcf.iLogDate(0) Then
                    'Find Time
                    mFindLibraryStartTime = mScanLcfForTime()
                    Exit Function
                End If
            End If
        Else
            Exit Do
        End If
        ilRet = btrGetNext(hmLcf, tmLcf, imLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mFindLibraryStartTime = ""

End Function

Private Function mScanLcfForTime() As String
    Dim ilLcf As Integer
    Dim llPrgTime As Long
    Dim llTime As Long
    Dim ilRet As Integer
    Dim slStr As String
    
    llPrgTime = 999999
    For ilLcf = LBound(tmLcf.lLvfCode) To UBound(tmLcf.lLvfCode) Step 1
        If tmLcf.lLvfCode(ilLcf) > 0 Then
            gUnpackTimeLong tmLcf.iTime(0, ilLcf), tmLcf.iTime(1, ilLcf), False, llTime
            If (llTime < llPrgTime) And (llTime >= 0) Then
                llPrgTime = llTime
            End If
        End If
    Next ilLcf
    If llPrgTime <> 999999 Then
        slStr = Format(gFormatTimeLong(llPrgTime, "A", "1"), "h:mm:ssA/P")
        If slStr = "12M" Then
            slStr = "12A"
        End If
        mScanLcfForTime = slStr
    Else
        mScanLcfForTime = ""
    End If
End Function

Private Function mGetAirWavePrgID(ilVefCode As Integer) As String
    Dim ilRet As Integer
    
    mGetAirWavePrgID = ""
    ilRet = btrGetEqual(hmVff, tmVff, imVffRecLen, ilVefCode, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        mGetAirWavePrgID = Trim$(tmVff.sAirWavePrgID)
    End If
End Function


Private Sub mReadAWF(slDate As String)
    Dim ilRet As Integer
    
    gPackDate slDate, tmAwfSrchKey1.iDateExported(0), tmAwfSrchKey1.iDateExported(1)
    ilRet = btrGetEqual(hmAwf, tmAwf, imAwfRecLen, tmAwfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
    If ilRet <> BTRV_ERR_NONE Then
        tmAwf.lCode = 0
        gPackDate slDate, tmAwf.iDateExported(0), tmAwf.iDateExported(1)
        tmAwf.iSeqNo = 0
        tmAwf.iUrfCode = tgUrf(0).iCode
        gPackDate smGenNowDate, tmAwf.iEnteredDate(0), tmAwf.iEnteredDate(1)
        gPackTime smGenNowTime, tmAwf.iEnteredTime(0), tmAwf.iEnteredTime(1)
        tmAwf.sUnused = ""
    End If
End Sub

Private Function mUpdateAwf() As Integer
    If tmAwf.lCode = 0 Then
        mUpdateAwf = btrInsert(hmAwf, tmAwf, imAwfRecLen, INDEXKEY0)
    Else
        mUpdateAwf = btrUpdate(hmAwf, tmAwf, imAwfRecLen)
    End If
End Function

Private Function mFindGameDate(ilGameNo As Integer, slGameDate As String, slGameStartTime As String) As Integer
    Dim ilRet As Integer
    Dim llDate As Long
    
    tmLcfSrchKey1.iVefCode = tmVef.iCode
    tmLcfSrchKey1.iType = ilGameNo
    ilRet = btrGetEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = tmVef.iCode) And (tmLcf.iType = ilGameNo)
        If (tmLcf.sStatus = "C") Then
            gUnpackDateLong tmLcf.iLogDate(0), tmLcf.iLogDate(1), llDate
            If llDate = gDateValue(slGameDate) Then
                slGameStartTime = mScanLcfForTime()
                mFindGameDate = True
                Exit Function
            End If
        End If
        ilRet = btrGetNext(hmLcf, tmLcf, imLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mFindGameDate = False
    Exit Function
    
End Function
