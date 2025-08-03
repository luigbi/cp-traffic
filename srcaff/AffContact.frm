VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmContact 
   Caption         =   "Contact"
   ClientHeight    =   7200
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9480
   Icon            =   "AffContact.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   9480
   Begin VB.ListBox lbcVehicleSort 
      Height          =   255
      ItemData        =   "AffContact.frx":08CA
      Left            =   3195
      List            =   "AffContact.frx":08CC
      Sorted          =   -1  'True
      TabIndex        =   29
      Top             =   1290
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox pbcActionTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   20
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox pbcActionSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   17
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox pbcPostFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   26
      Top             =   0
      Width           =   60
   End
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox pbcActionFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Width           =   45
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   0
      Picture         =   "AffContact.frx":08CE
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.TextBox txtDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   6120
      TabIndex        =   19
      Top             =   2520
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox pbcActionArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   0
      Picture         =   "AffContact.frx":0BD8
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4560
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.CommandButton cmdMail 
      Caption         =   "Generate Mail List"
      Height          =   375
      Left            =   3360
      TabIndex        =   22
      Top             =   6660
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Comments"
      Height          =   375
      Left            =   1440
      TabIndex        =   21
      Top             =   6660
      Width           =   1575
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   8520
      TabIndex        =   8
      Top             =   50
      Width           =   900
   End
   Begin VB.Frame Frame8 
      Caption         =   "Selection by"
      Height          =   1260
      Left            =   120
      TabIndex        =   9
      Top             =   75
      Width           =   2835
      Begin VB.OptionButton optSort 
         Caption         =   "Advertiser, then Contracts"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   255
         Value           =   -1  'True
         Width           =   2550
      End
      Begin VB.OptionButton optSort 
         Caption         =   "Vehicle, then Stations"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2565
      End
      Begin VB.OptionButton optSort 
         Caption         =   "Station, then Vehicles"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   915
         Width           =   2595
      End
   End
   Begin VB.ListBox lbcContract 
      Height          =   1230
      ItemData        =   "AffContact.frx":0EE2
      Left            =   7185
      List            =   "AffContact.frx":0EE4
      MultiSelect     =   2  'Extended
      TabIndex        =   7
      Top             =   330
      Width           =   2100
   End
   Begin VB.ListBox lbcAdvertiser 
      Height          =   1230
      ItemData        =   "AffContact.frx":0EE6
      Left            =   4860
      List            =   "AffContact.frx":0EE8
      TabIndex        =   5
      Top             =   330
      Width           =   2220
   End
   Begin VB.Timer tmcFill 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   30
      Top             =   6045
   End
   Begin VB.TextBox txtNoWeeks 
      Height          =   240
      Left            =   3930
      TabIndex        =   3
      Top             =   810
      Width           =   525
   End
   Begin VB.TextBox txtWeek 
      Height          =   240
      Left            =   3930
      TabIndex        =   1
      Top             =   330
      Width           =   840
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   6240
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   7200
      FormDesignWidth =   9480
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   375
      Left            =   5340
      TabIndex        =   13
      Top             =   6660
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7620
      TabIndex        =   14
      Top             =   6660
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPost 
      Height          =   2250
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1620
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3969
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
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
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAction 
      Height          =   2505
      Left            =   120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4080
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4419
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      Appearance      =   0
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
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label4 
      Caption         =   "Magenta = Partially Posted"
      ForeColor       =   &H00FF00FF&
      Height          =   225
      Index           =   3
      Left            =   6360
      TabIndex        =   28
      Top             =   3870
      Width           =   1965
   End
   Begin VB.Label Label4 
      Caption         =   "Red = Outstanding"
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   0
      Left            =   4680
      TabIndex        =   27
      Top             =   3870
      Width           =   1395
   End
   Begin VB.Image imcPrt 
      Height          =   480
      Left            =   600
      Picture         =   "AffContact.frx":0EEA
      Stretch         =   -1  'True
      Top             =   6645
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "# Weeks"
      Height          =   255
      Index           =   1
      Left            =   3105
      TabIndex        =   2
      Top             =   840
      Width           =   900
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Contracts"
      Height          =   255
      Left            =   7155
      TabIndex        =   6
      Top             =   45
      Width           =   1110
   End
   Begin VB.Label lacTitle1 
      Alignment       =   2  'Center
      Caption         =   "Advertisers"
      Height          =   255
      Left            =   4470
      TabIndex        =   4
      Top             =   45
      Width           =   1890
   End
   Begin VB.Label Label1 
      Caption         =   "Start Week"
      Height          =   255
      Left            =   2985
      TabIndex        =   0
      Top             =   375
      Width           =   975
   End
End
Attribute VB_Name = "frmContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmContact - allows for selection of station/vehicle/advertiser for contact information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text

Private hmMail As Integer
Private tmContactInfo() As CONTACTINFO
Private smToFile As String
Private tmCDate() As CONTACTDATE
Private imCIndex As Integer
Private imCMax As Integer
Private imShttCode As Integer
Private imVefCode As Integer
Private imCommentChgd As Integer
'Dim iLoadedRow As Integer
Private imSIntegralSet As Integer
Private imCIntegralSet As Integer
Private imFirstTime As Integer
Private imHeaderClick As Integer
Private lmPostRow As Long

'Grid Controls
Private imShowGridBox As Integer
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on

Const STATIONINDEX = 0
Const VEHICLEINDEX = 1
Const CONTACTINDEX = 2
Const TELEPHONEINDEX = 3
Const DATE1INDEX = 4
Const DATE2INDEX = 5
Const DATE3INDEX = 6
Const EXTRAINDEX = 7

Const ACTIONDATEINDEX = 0
Const USERNAMEINDEX = 1
Const COMMENTINDEX = 2
Const CCTCODEINDEX = 3
Const USTCODEINDEX = 4
Const DATEENTEREDINDEX = 5
Const TIMEENTEREDINDEX = 6
Const CHANGEFLAGINDEX = 7


Private imAdfCode As Integer
Private smFWkDate As String    'First week start start
Private smLWkDate As String  'Last week end date
Private hmExport As Integer
'Private imShttCode As Integer
'Private imVefCode As Integer
Private imAllClick As Integer
Private smNowDate As String
Private smNowTime As String


Private Sub chkAll_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcContract.ListCount > 0 Then
        imAllClick = True
        lRg = CLng(lbcContract.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcContract.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllClick = False
    End If
End Sub

Private Sub cmdCancel_Click()
    tmcFill.Enabled = False
    txtWeek.Text = ""
    Unload frmContact
End Sub

Private Sub cmdGenerate_Click()
    Dim i, j As Integer
    Dim bm As Variant
    Dim iNoWeeks As Integer
    Dim dLWeek As Date
    Dim dFWeek As Date
    
    'If (smFWkDate = "") Or (smLWkDate = "") Then
    '    Beep
    '    gMsgBox "Dates must be specified.", vbOKOnly
    '    If smFWkDate = "" Then
    '        txtWeek.SetFocus
    '    Else
    '        txtNoWeeks.SetFocus
    '    End If
    '    Exit Sub
    'End If
    'If gIsDate(txtWeek.Text) = False Then
    '    Beep
    '    gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
    '    txtWeek.SetFocus
    'End If
    If txtWeek.Text = "" Then
        gMsgBox "Date must be specified.", vbOKOnly
        txtWeek.SetFocus
        Exit Sub
    End If
    If Trim$(txtNoWeeks.Text) = "" Then
        gMsgBox "# Weeks must be specified.", vbOKOnly
        txtNoWeeks.SetFocus
        Exit Sub
    End If
    If gIsDate(txtWeek.Text) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        txtWeek.SetFocus
        Exit Sub
    Else
        smFWkDate = Format(txtWeek.Text, sgShowDateForm)
    End If
    dFWeek = CDate(smFWkDate)
    iNoWeeks = 7 * CInt(txtNoWeeks.Text) - 1
    dLWeek = DateAdd("d", iNoWeeks, dFWeek)
    sLWeek = CStr(dLWeek)
    smLWkDate = Format$(sLWeek, sgShowDateForm)
    
    Screen.MousePointer = vbHourglass
    
    If lbcAdvertiser.ListIndex < 0 Then
        Screen.MousePointer = vbDefault
        If optSort(0).Value Then
            gMsgBox "Advertiser must be selected.", vbOKOnly
        ElseIf optSort(1).Value Then
            gMsgBox "Vehicle must be selected.", vbOKOnly
        Else
            gMsgBox "Station must be selected.", vbOKOnly
        End If
        lbcAdvertiser.SetFocus
        Exit Sub
    End If
    sAdvtDates = Trim$(lbcAdvertiser.List(lbcAdvertiser.ListIndex))
    sContracts = ""
    For i = 0 To lbcContract.ListCount - 1
        If lbcContract.Selected(i) Then
            If optSort(0).Value Then
                If sContracts = "" Then
                    'sContracts = "(lst.lstStartDate BETWEEN '" & smFWkDate & "' AND '" & smLWkDate & "')"
                    sContracts = "(lstLogDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(smLWkDate, sgSQLDateForm) & "')"
                    sContracts = sContracts & " AND ( cpttVefCode = lstLogVefCode " & "And cpttStartDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(smLWkDate, sgSQLDateForm) & "')"
                    sContracts = sContracts & " AND ((lstCntrNo = " & lbcContract.List(i) & ")"
                    sAdvtDates = sAdvtDates & " " & lbcContract.List(i)
                Else
                    sContracts = sContracts & " OR (lstCntrNo = " & lbcContract.List(i) & ")"
                    sAdvtDates = sAdvtDates & ", " & lbcContract.List(i)
                End If
            ElseIf optSort(1).Value Then
                imShttCode = lbcContract.ItemData(i)
                If sContracts = "" Then
                    sContracts = "(cpttStartDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(smLWkDate, sgSQLDateForm) & "')"
                    sContracts = sContracts & " AND (cpttVefCode = " & imVefCode & ")"
                    If chkAll.Value = vbUnchecked Then          'prevent extra SQL AND testing of stations
                        sContracts = sContracts & " AND ((cpttShfCode = " & imShttCode & ")"
                    End If
                    sAdvtDates = sAdvtDates & " " & lbcContract.List(i)
                Else
                    If chkAll.Value = vbUnchecked Then          'prevent extra SQL AND testing of stations
                        sContracts = sContracts & " OR (cpttShfCode = " & imShttCode & ")"
                    End If
                    sAdvtDates = sAdvtDates & ", " & lbcContract.List(i)
                End If
            Else
                imVefCode = lbcContract.ItemData(i)
                If sContracts = "" Then
                    sContracts = "(cpttStartDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(smLWkDate, sgSQLDateForm) & "')"
                    sContracts = sContracts & " AND (cpttShfCode = " & imShttCode & ")"
                    If chkAll.Value = vbUnchecked Then      'prevent extra SQL AND testing of vehicles
                        sContracts = sContracts & " AND ((cpttVefCode = " & imVefCode & ")"
                    End If
                    sAdvtDates = sAdvtDates & " " & lbcContract.List(i)
                Else
                    If chkAll.Value = vbUnchecked Then          'prevent extra SQL AND testing of vehicles
                        sContracts = sContracts & " OR (cpttVefCode = " & imVefCode & ")"
                    End If
                    sAdvtDates = sAdvtDates & ", " & lbcContract.List(i)
                End If
            End If
        End If
    Next i
    If sContracts = "" Then
        Screen.MousePointer = vbDefault
        If optSort(0).Value Then
            gMsgBox "Contracts must be selected.", vbOKOnly
        ElseIf optSort(1).Value Then
            gMsgBox "Stations must be selected.", vbOKOnly
        Else
            gMsgBox "Vehicles must be selected.", vbOKOnly
        End If
        lbcContract.SetFocus
        Exit Sub
    End If
    sContracts = sContracts & ")"
    If optSort(0).Value Then
        SQLQuery = "SELECT cpttStartDate, cpttPostingStatus, shttCallLetters, vefType, vefName, vefCode, shttACName, shttACPhone, shttCode, attACName, attACPhone"
        'SQLQuery = SQLQuery + " FROM cptt, lst, shtt, VEF_Vehicles vef, Att"
        SQLQuery = SQLQuery & " FROM cptt, lst, shtt, VEF_Vehicles, att"
        SQLQuery = SQLQuery + " WHERE ((cpttStatus = 0) AND (vefCode = lstLogVefCode)"
        SQLQuery = SQLQuery + " AND (shttCode = cpttShfCode)"
        SQLQuery = SQLQuery + " AND (attCode = cpttAtfCode)"
        SQLQuery = SQLQuery + " AND " & sContracts & ")"
        SQLQuery = SQLQuery + " ORDER BY shttCallLetters, vefName, cpttStartDate"
    ElseIf optSort(1).Value Then
        SQLQuery = "SELECT cpttStartDate, cpttPostingStatus, shttCallLetters, vefType, vefName, vefcode, shttACName, shttACPhone, shttCode, attACName, attACPhone"
        'SQLQuery = SQLQuery + " FROM cptt, shtt, VEF_Vehicles vef, att"
        SQLQuery = SQLQuery & " FROM cptt, shtt, VEF_Vehicles, att"
        SQLQuery = SQLQuery + " WHERE ((cpttStatus = 0) AND (vefCode = cpttVefCode)"
        SQLQuery = SQLQuery + " AND (shttCode = cpttShfCode)"
        SQLQuery = SQLQuery + " AND (attCode = cpttAtfCode)"
        If chkAll.Value = vbChecked Then            'prevent extra SQL testing of stations
            SQLQuery = SQLQuery + "AND " & sContracts
        Else
            SQLQuery = SQLQuery + " AND " & sContracts & ")"
        End If
        SQLQuery = SQLQuery + " ORDER BY shttCallLetters, vefName, cpttStartDate"
    Else
        SQLQuery = "SELECT cpttStartDate, cpttPostingStatus, shttCallLetters, vefType, vefName, vefCode, shttACName, shttACPhone, shttCode, attACName, attACPhone, attcode "
        'SQLQuery = SQLQuery + " FROM cptt, shtt, VEF_Vehicles vef, att"
        SQLQuery = SQLQuery & " FROM cptt, shtt, VEF_Vehicles, att"
        SQLQuery = SQLQuery + " WHERE ((cpttStatus = 0) AND (vefCode = cpttVefCode)"
        SQLQuery = SQLQuery + " AND (shttCode = cpttShfCode)"
        SQLQuery = SQLQuery + " AND (attCode = cpttAtfCode)"
        If chkAll.Value = vbChecked Then           'prevent extra SQL testing of vehicles
            SQLQuery = SQLQuery + "AND " & sContracts
        Else
            SQLQuery = SQLQuery + " AND " & sContracts & ")"
        End If
        SQLQuery = SQLQuery + " ORDER BY shttCallLetters, vefName, cpttStartDate"
   End If
    Screen.MousePointer = vbDefault
    'frmContactGrid.Show vbModal
    gGrid_Clear grdPost, True
    'gGrid_Clear grdAction, True
    mActionClearGrid
    mRefreshGrid
    
    '7-22-09 Update user, date and time added
    smNowDate = Format$(gNow(), "m/d/yy")
    smNowTime = Format$(gNow(), "h:mm:ssAM/PM")
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    
    If imFirstTime Then
        
        grdPost.ColWidth(EXTRAINDEX) = 0
        grdPost.ColWidth(STATIONINDEX) = grdPost.Width * 0.12
        grdPost.ColWidth(VEHICLEINDEX) = grdPost.Width * 0.25
        grdPost.ColWidth(TELEPHONEINDEX) = grdPost.Width * 0.15
        grdPost.ColWidth(DATE1INDEX) = grdPost.Width * 0.08
        grdPost.ColWidth(DATE2INDEX) = grdPost.Width * 0.08
        grdPost.ColWidth(DATE3INDEX) = grdPost.Width * 0.08
        grdPost.ColWidth(CONTACTINDEX) = grdPost.Width - grdPost.ColWidth(STATIONINDEX) - grdPost.ColWidth(VEHICLEINDEX) - grdPost.ColWidth(TELEPHONEINDEX) - grdPost.ColWidth(DATE1INDEX) - grdPost.ColWidth(DATE2INDEX) - grdPost.ColWidth(DATE3INDEX) - GRIDSCROLLWIDTH '(5 * grdStation.Columns(6).Width) / 6
        gGrid_AlignAllColsLeft grdPost
        grdPost.TextMatrix(0, STATIONINDEX) = "Station"
        grdPost.TextMatrix(0, VEHICLEINDEX) = "Vehicle"
        grdPost.TextMatrix(0, CONTACTINDEX) = "Contact"
        grdPost.TextMatrix(0, TELEPHONEINDEX) = "Telephone"
        grdPost.TextMatrix(0, DATE1INDEX) = Chr$(171)
        grdPost.TextMatrix(0, DATE2INDEX) = "Date"
        grdPost.TextMatrix(0, DATE3INDEX) = Chr$(187)
        gGrid_IntegralHeight grdPost
        gGrid_Clear grdPost, True
    
        'Hide column 2
        grdAction.ColWidth(CCTCODEINDEX) = 0
        grdAction.ColWidth(USTCODEINDEX) = 0
        grdAction.ColWidth(DATEENTEREDINDEX) = 0
        grdAction.ColWidth(TIMEENTEREDINDEX) = 0
        grdAction.ColWidth(CHANGEFLAGINDEX) = 0
        grdAction.ColWidth(ACTIONDATEINDEX) = grdAction.Width * 0.16
        grdAction.ColWidth(USERNAMEINDEX) = grdAction.Width * 0.16
        grdAction.ColWidth(COMMENTINDEX) = grdAction.Width - grdAction.ColWidth(ACTIONDATEINDEX) - grdAction.ColWidth(USERNAMEINDEX) - GRIDSCROLLWIDTH '(5 * grdStation.Columns(6).Width) / 6
        gGrid_AlignAllColsLeft grdAction
        grdAction.TextMatrix(0, ACTIONDATEINDEX) = "Action Date"
        grdAction.TextMatrix(0, USERNAMEINDEX) = "User Name"
        grdAction.TextMatrix(0, COMMENTINDEX) = "Comment"
        gGrid_IntegralHeight grdAction
        'gGrid_Clear grdAction, True
        mActionClearGrid
        GridPaint True
        imFirstTime = False
    End If


End Sub

Private Sub Form_Initialize()
'    Me.Width = Screen.Width / 1.05
'    Me.Height = Screen.Height / 1.7
'    Me.Top = (Screen.Height - Me.Height) / 2.5
'    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Width = Screen.Width / 1.02   '1.05  '1.15
    Me.Height = Screen.Height / 1.25 '15    '1.45    '1.25
    Me.Top = (Screen.Height - Me.Height) / 1.2
    Me.Left = (Screen.Width - Me.Width) / 1.2

    gSetFonts frmContact
    gCenterForm frmContact
End Sub




Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmContactInfo
    Erase tmCDate
    Set frmContact = Nothing
End Sub





Private Sub lbcAdvertiser_Click()
    Dim iNoWeeks As Integer
    Dim dLWeek As Date
    Dim dFWeek As Date
    Dim sMarket As String
    On Error GoTo ErrHand
    
    lbcContract.Clear
    chkAll.Value = vbUnchecked
    If optSort(0).Value Then
        'Select by Advertiser then Contracts
        If txtWeek.Text = "" Then
            gMsgBox "Date must be specified.", vbOKOnly
            txtWeek.SetFocus
            Exit Sub
        End If
        If Trim$(txtNoWeeks.Text) = "" Then
            gMsgBox "# Weeks must be specified.", vbOKOnly
            txtNoWeeks.SetFocus
            Exit Sub
        End If
        If gIsDate(txtWeek.Text) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
            txtWeek.SetFocus
            Exit Sub
        Else
            smFWkDate = Format(txtWeek.Text, sgShowDateForm)
        End If
        Screen.MousePointer = vbHourglass
        dFWeek = CDate(smFWkDate)
        iNoWeeks = 7 * CInt(txtNoWeeks.Text) - 1
        dLWeek = DateAdd("d", iNoWeeks, dFWeek)
        sLWeek = CStr(dLWeek)
        smLWkDate = Format$(sLWeek, sgShowDateForm)
        imAdfCode = lbcAdvertiser.ItemData(lbcAdvertiser.ListIndex)
        
        'SQLQuery = "SELECT DISTINCT lst.lstCntrNo from ADF_Advertisers adf, lst"
        SQLQuery = "SELECT DISTINCT lstCntrNo"
        SQLQuery = SQLQuery & " FROM ADF_Advertisers, lst"
        SQLQuery = SQLQuery + " WHERE (adfCode = lstAdfCode"
        SQLQuery = SQLQuery + " AND lstCntrNo <> 0"
        SQLQuery = SQLQuery + " AND lstStartDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "' AND lstStartDate <= '" & Format$(smLWkDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery + " AND adfCode = " & imAdfCode & ")"
        SQLQuery = SQLQuery + " ORDER BY lstCntrNo"
        
        Set rst = gSQLSelectCall(SQLQuery)
        While Not rst.EOF
            lbcContract.AddItem rst!lstCntrNo  ', " & rst(1).Value & ""
            rst.MoveNext
        Wend
    ElseIf optSort(1).Value Then
        'Select by Vehicle then Stations
        Screen.MousePointer = vbHourglass
        imVefCode = lbcAdvertiser.ItemData(lbcAdvertiser.ListIndex)
        'SQLQuery = "SELECT DISTINCT shttCallLetters, shttMarket, shttCode"
        SQLQuery = "SELECT DISTINCT shttCallLetters, mktName, shttCode"
        ''SQLQuery = SQLQuery + " FROM shtt, cptt"
        ''SQLQuery = SQLQuery + " WHERE (cptt.cpttVefCode = " & imVefCode
        ''SQLQuery = SQLQuery + " AND cptt.cpttStatus = 0"
        ''SQLQuery = SQLQuery + " AND cptt.cpttStartDate BETWEEN '" & smFWkDate & "' AND '" & smLWkDate & "'"
        ''SQLQuery = SQLQuery + " AND shtt.shttCode = cptt.cpttShfCode)"
        'SQLQuery = SQLQuery + " FROM shtt, att"
        SQLQuery = SQLQuery + " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode, att"
        SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode
        SQLQuery = SQLQuery + " AND shttCode = attShfCode)"
        SQLQuery = SQLQuery + " ORDER BY shttCallLetters"
        Set rst = gSQLSelectCall(SQLQuery)
        While Not rst.EOF
            'If IsNull(rst!shttMarket) = True Then
            '    sMarket = ""
            'Else
            '    sMarket = rst!shttMarket  'Trim$(rst!shttMarket)
            'End If
            'lbcContract.AddItem Trim$(rst!shttCallLetters) & ", " & Trim$(sMarket)  ', " & rst(1).Value & ""
            If IsNull(rst!mktName) = True Then
                sMarket = ""
                lbcContract.AddItem Trim$(rst!shttCallLetters)
            Else
                sMarket = rst!mktName  'Trim$(rst!shttMarket)
                lbcContract.AddItem Trim$(rst!shttCallLetters) & ", " & Trim$(sMarket)  ', " & rst(1).Value & ""
            End If
            lbcContract.ItemData(lbcContract.NewIndex) = rst!shttCode
            rst.MoveNext
        Wend
        'insure that the previous vehicles post info gets cleared
        ReDim tmContactInfo(0 To 0) As CONTACTINFO
        ReDim tmCDate(0 To 0) As CONTACTDATE

    Else
        'Select by Station then Vehicles
        Screen.MousePointer = vbHourglass
        imShttCode = lbcAdvertiser.ItemData(lbcAdvertiser.ListIndex)
        SQLQuery = "SELECT DISTINCT vefType, vefName, vefCode"
        ''SQLQuery = SQLQuery + " FROM vef, cptt"
        ''SQLQuery = SQLQuery + " WHERE (vef.vefCode = cptt.cpttVefCode"
        ''SQLQuery = SQLQuery + " AND cptt.cpttShfCode = " & imShttCode & ""
        ''SQLQuery = SQLQuery + " AND cptt.cpttStatus = 0"
        ''SQLQuery = SQLQuery + " AND cptt.cpttStartDate BETWEEN '" & smFWkDate & "' AND '" & smLWkDate & "'"
        ''SQLQuery = SQLQuery + " AND ((vef.vefvefCode = 0 AND vef.vefType = 'C') OR vef.vefType = 'L' OR vef.vefType = 'A'))"
        'SQLQuery = SQLQuery + " FROM VEF_Vehicles vef, att"
        SQLQuery = SQLQuery & " FROM VEF_Vehicles, att"
        SQLQuery = SQLQuery + " WHERE (attShfCode = " & imShttCode
        SQLQuery = SQLQuery + " AND vefCode = attVefCode)"
        SQLQuery = SQLQuery + " ORDER BY vefName"
        Set rst = gSQLSelectCall(SQLQuery)
        While Not rst.EOF
            If sgShowByVehType = "Y" Then
                lbcContract.AddItem Trim$(rst!vefType) & ":" & Trim$(rst!vefName)
            Else
                lbcContract.AddItem Trim$(rst!vefName)
            End If
            lbcContract.ItemData(lbcContract.NewIndex) = rst!vefCode
            rst.MoveNext
        Wend
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Contact-lbcAdvertiser"
End Sub

Private Sub lbcContract_Click()
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = vbChecked Then
        imAllClick = True
        chkAll.Value = vbUnchecked
        imAllClick = False
    End If
End Sub

Private Sub optSort_Click(Index As Integer)
    If optSort(Index).Value Then
        Screen.MousePointer = vbHourglass
        lbcAdvertiser.Clear
        lbcContract.Clear
        chkAll.Value = vbChecked
        If Index = 0 Then
            lacTitle1.Caption = "Advertisers"
            lacTitle2.Caption = "Contracts"
            mFillAdvt
        ElseIf Index = 1 Then
            lacTitle1.Caption = "Vehicles"
            lacTitle2.Caption = "Stations"
            mFillVehicle
        Else
            lacTitle1.Caption = "Stations"
            lacTitle2.Caption = "Vehicles"
            mFillStation
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub tmcFill_Timer()
    Dim ilLoop As Integer
    Dim slStr As String
    
    tmcFill.Enabled = False
    If lbcAdvertiser.ListCount <= 0 Then
        Exit Sub
    End If
    slStr = Trim$(txtWeek.Text)
    If slStr = "" Then
        Exit Sub
    End If
    slStr = Trim$(txtNoWeeks.Text)
    If slStr = "" Then
        Exit Sub
    End If
    'Single selection only
'    For ilLoop = 0 To lbcAdvertiser.ListCount - 1 Step 1
'        If lbcAdvertiser.Selected(ilLoop) Then
'            lbcAdvertiser_Click
'            Exit For
'        End If
'    Next ilLoop
    If lbcAdvertiser.ListIndex >= 0 Then
        lbcAdvertiser_Click
    End If
End Sub

Private Sub txtNoWeeks_Change()
    tmcFill.Enabled = False
    imAdfCode = -1
    'If optSort(0).Value Then
    '    lbcAdvertiser.Clear
    'End If
    lbcContract.Clear
    chkAll.Value = vbUnchecked
    tmcFill.Enabled = True
End Sub

Private Sub txtNoWeeks_GotFocus()
    tmcFill.Enabled = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtNoWeeks_LostFocus()
    'If Not tmcFill.Enabled Then
    '    tmcFill.Enabled = True
    'End If
End Sub


Private Sub txtWeek_Change()
    tmcFill.Enabled = False
    imAdfCode = -1
    txtNoWeeks.Text = ""
    'If optSort(0).Value Then
    '    lbcAdvertiser.Clear
    'End If
    lbcContract.Clear
    chkAll.Value = vbUnchecked
    tmcFill.Enabled = True
End Sub

Private Sub txtWeek_GotFocus()
    tmcFill.Enabled = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtWeek_LostFocus()
    'Dim iResponse As Integer

    'If txtWeek.Text = "" Then
    '    Exit Sub
    'End If
   '
    'If gIsDate(txtWeek.Text) = False Then
    '    Beep
    '    gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
    '    txtWeek.SetFocus
    'Else
    '    smFWkDate = Format(txtWeek.Text, "m/d/yy")
    'End If
End Sub



Private Sub mFillAdvt()
    Dim iNoWeeks As Integer
    Dim dLWeek As Date
    Dim dFWeek As Date
    Dim iFound As Integer
    Dim iLoop As Integer
    On Error GoTo ErrHand
    
    lbcAdvertiser.Clear
    lbcContract.Clear
    chkAll.Value = vbUnchecked
    ''If txtWeek.Text = "" Then
    ''    Exit Sub
    ''End If
    ''If Trim$(txtNoWeeks.Text) = "" Then
    ''    Exit Sub
    ''End If
    ''
    ''If gIsDate(txtWeek.Text) = False Then
    ''    Beep
    ''    gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
    ''    txtWeek.SetFocus
    ''Else
    ''    smFWkDate = Format(txtWeek.Text, "m/d/yy")
    ''End If
    
    ''dFWeek = CDate(smFWkDate)
    ''iNoWeeks = 7 * CInt(txtNoWeeks.Text) - 1
    ''dLWeek = DateAdd("d", iNoWeeks, dFWeek)
    ''sLWeek = CStr(dLWeek)
    ''smLWkDate = Format$(sLWeek, "m/d/yy")
    ''SQLQuery = "SELECT adf.adfName, adf.adfCode from adf, lst"
    ''SQLQuery = SQLQuery + " WHERE (adf.adfCode = lst.lstAdfCode"
    ''SQLQuery = SQLQuery + " AND lst.lstStartDate BETWEEN '" & smFWkDate & "' AND '" & smLWkDate & "')"
    ''SQLQuery = SQLQuery + " ORDER BY adf.adfName"
    'SQLQuery = "SELECT adf.adfName, adf.adfCode from ADF_Advertisers adf"
    SQLQuery = "SELECT adfName, adfCode"
    SQLQuery = SQLQuery & " FROM ADF_Advertisers"
    SQLQuery = SQLQuery + " ORDER BY adfName"

    

    'SQLQuery = "SELECT adf.adfName, adf.adfCode from adf"
    'SQLQuery = SQLQuery + " ORDER BY adf.adfName"
    
    Set rst = gSQLSelectCall(SQLQuery)
    'lbcAdvertiser.Clear
    While Not rst.EOF
        iFound = False
        'For iLoop = 0 To lbcAdvertiser.ListCount - 1 Step 1
        '    If lbcAdvertiser.ItemData(iLoop) = rst!adfCode Then
        '        iFound = True
        '        Exit For
        '    End If
        'Next iLoop
        If Not iFound Then
            lbcAdvertiser.AddItem rst!adfName '& ", " & rst(1).Value
            lbcAdvertiser.ItemData(lbcAdvertiser.NewIndex) = rst!adfCode
        End If
        rst.MoveNext
    Wend
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Contact-mFillAdvt"
End Sub

Private Sub mFillVehicle()
    Dim iLoop As Integer
    'lbcAdvertiser.Clear
    'lbcContract.Clear
    'chkAll.Value = vbUnchecked
    'For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
    '    'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
    '        lbcAdvertiser.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
    '        lbcAdvertiser.ItemData(lbcAdvertiser.NewIndex) = tgVehicleInfo(iLoop).iCode
    '    'End If
    'Next iLoop
    lbcAdvertiser.Clear
    lbcContract.Clear
    lbcVehicleSort.Clear
    chkAll.Value = vbUnchecked
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehicleSort.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
        lbcVehicleSort.ItemData(lbcVehicleSort.NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop
    For iLoop = 0 To lbcVehicleSort.ListCount - 1 Step 1
        lbcAdvertiser.AddItem Trim$(lbcVehicleSort.List(iLoop))
        lbcAdvertiser.ItemData(lbcAdvertiser.NewIndex) = lbcVehicleSort.ItemData(iLoop)
    Next iLoop
End Sub

Private Sub mFillStation()
    Dim iLoop As Integer
    lbcAdvertiser.Clear
    lbcContract.Clear
    chkAll.Value = vbUnchecked
    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(iLoop).sUsedForATT = "Y" Then
            If tgStationInfo(iLoop).iType = 0 Then
                lbcAdvertiser.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcAdvertiser.ItemData(lbcAdvertiser.NewIndex) = tgStationInfo(iLoop).iCode
            End If
        End If
    Next iLoop

End Sub

Private Sub GridPaint(iClear As Integer)
    Dim iTotRec As Integer
    Dim iSIndex As Integer
    Dim iEIndex As Integer
    Dim iRow As Integer
    Dim iLoop As Integer
    Dim iSLoop As Integer
    Dim iELoop As Integer
    Dim iStep As Integer
    Dim iCol As Integer
    Dim iTRow As Integer
    Dim sStation, sVehicle, sACName, sACPhone As String
    Dim sDate(0 To 2) As String
    Dim ilPostingStatus(0 To 2) As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim llTRow As Long
    
    grdPost.Redraw = False
    If iClear Then
        mClearGrid
        iSLoop = 0
        iELoop = UBound(tmContactInfo) - 1
        iStep = 1
    Else
        'iTRow = iLoadedRow - grdStation.VisibleRows
'        iELoop = 0
'        iSLoop = UBound(tmContactInfo) - 1
'        iStep = -1
        iSLoop = 0
        iELoop = UBound(tmContactInfo) - 1
        iStep = 1
    End If
    llRow = grdPost.FixedRows
    llTRow = grdPost.TopRow
    For iRow = iSLoop To iELoop Step iStep
        sStation = tmContactInfo(iRow).sStation
        sVehicle = tmContactInfo(iRow).sVehicle
        sACName = tmContactInfo(iRow).sACName
        sACPhone = tmContactInfo(iRow).sACPhone
        iSIndex = tmContactInfo(iRow).iCDateIndex
        If iRow < UBound(tmContactInfo) - 1 Then
            iEIndex = tmContactInfo(iRow + 1).iCDateIndex - 1
        Else
            iEIndex = UBound(tmCDate) - 1
        End If
        iTotRec = iEIndex - iSIndex + 1
        
        iSIndex = 3 * (imCIndex - 1) + iSIndex
        If iSIndex <= iEIndex Then
            If iSIndex + 2 < iEIndex Then
                iEIndex = iSIndex + 2
            End If
            iCol = 0
            sDate(0) = ""
            sDate(1) = ""
            sDate(2) = ""
            ilPostingStatus(0) = 0
            ilPostingStatus(1) = 0
            ilPostingStatus(2) = 0
            For iLoop = iSIndex To iEIndex Step 1
                sDate(iCol) = tmCDate(iLoop).sDate
                ilPostingStatus(iCol) = tmCDate(iLoop).iPostingStatus
                iCol = iCol + 1
            Next iLoop
        Else
            sDate(0) = ""
            sDate(1) = ""
            sDate(2) = ""
            ilPostingStatus(0) = 0
            ilPostingStatus(1) = 0
            ilPostingStatus(2) = 0
        End If
        If llRow + 1 > grdPost.Rows Then
            grdPost.AddItem ""
        End If
        grdPost.Row = llRow
        For llCol = STATIONINDEX To grdPost.Cols - 2 Step 1
            grdPost.Col = llCol
            grdPost.CellBackColor = LIGHTYELLOW
        Next llCol
        If iClear Then
            grdPost.TextMatrix(llRow, STATIONINDEX) = sStation
            grdPost.TextMatrix(llRow, VEHICLEINDEX) = sVehicle
            grdPost.TextMatrix(llRow, CONTACTINDEX) = sACName
            grdPost.TextMatrix(llRow, TELEPHONEINDEX) = sACPhone
            grdPost.Col = DATE1INDEX
            If ilPostingStatus(0) <= 0 Then
                grdPost.CellForeColor = vbRed
            Else
                grdPost.CellForeColor = vbMagenta
            End If
            grdPost.TextMatrix(llRow, DATE1INDEX) = sDate(0)
            grdPost.Col = DATE2INDEX
            If ilPostingStatus(1) <= 0 Then
                grdPost.CellForeColor = vbRed
            Else
                grdPost.CellForeColor = vbMagenta
            End If
            grdPost.TextMatrix(llRow, DATE2INDEX) = sDate(1)
            grdPost.Col = DATE3INDEX
            If ilPostingStatus(2) <= 0 Then
                grdPost.CellForeColor = vbRed
            Else
                grdPost.CellForeColor = vbMagenta
            End If
            grdPost.TextMatrix(llRow, DATE3INDEX) = sDate(2)
            grdPost.TextMatrix(llRow, 7) = iRow
        Else
            grdPost.Col = DATE1INDEX
            If ilPostingStatus(0) <= 0 Then
                grdPost.CellForeColor = vbRed
            Else
                grdPost.CellForeColor = vbMagenta
            End If
            grdPost.TextMatrix(llRow, DATE1INDEX) = sDate(0)
            grdPost.Col = DATE2INDEX
            If ilPostingStatus(1) <= 0 Then
                grdPost.CellForeColor = vbRed
            Else
                grdPost.CellForeColor = vbMagenta
            End If
            grdPost.TextMatrix(llRow, DATE2INDEX) = sDate(1)
            grdPost.Col = DATE3INDEX
            If ilPostingStatus(2) <= 0 Then
                grdPost.CellForeColor = vbRed
            Else
                grdPost.CellForeColor = vbMagenta
            End If
            grdPost.TextMatrix(llRow, DATE3INDEX) = sDate(2)
        End If
        llRow = llRow + 1
    Next iRow
    'Don't add extra row
'    If llRow >= grdPost.Rows Then
'        grdPost.AddItem ""
'    End If
    If Not iClear Then
        grdPost.TopRow = llTRow
    End If
    grdPost.Redraw = True
    Exit Sub
End Sub

Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    gGrid_Clear grdPost, True
    grdPost.Row = 0
    grdPost.Col = DATE1INDEX
    grdPost.CellAlignment = flexAlignCenterTop
    'grdPost.TextMatrix(0, 2) = Chr$(171)
    grdPost.Row = 0
    grdPost.Col = DATE2INDEX
    grdPost.CellAlignment = flexAlignCenterTop
    'grdPost.TextMatrix(0, 8) = "Dates*"
    grdPost.Row = 0
    grdPost.Col = DATE3INDEX
    grdPost.CellAlignment = flexAlignCenterTop
    For llRow = grdPost.FixedRows To grdPost.Rows - 1 Step 1
        grdPost.Row = llRow
        For llCol = STATIONINDEX To grdPost.Cols - 2 Step 1
            grdPost.Col = llCol
            grdPost.CellBackColor = LIGHTYELLOW
        Next llCol
    Next llRow

End Sub

Private Sub mActionSetShow()
    If (lmEnableRow >= grdAction.FixedRows) And (lmEnableRow < grdAction.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        grdAction.TextMatrix(lmEnableRow, USERNAMEINDEX) = sgUserName
    End If
    imShowGridBox = False
    pbcActionArrow.Visible = False
    txtDropdown.Visible = False
End Sub


Private Sub mActionEnableBox()
    If (grdAction.Row >= grdAction.FixedRows) And (grdAction.Row < grdAction.Rows) And (grdAction.Col >= ACTIONDATEINDEX) And (grdAction.Col <= COMMENTINDEX) Then
        lmEnableRow = grdAction.Row
        imShowGridBox = True
        pbcActionArrow.Move grdAction.Left - pbcActionArrow.Width - 15, grdAction.Top + grdAction.RowPos(grdAction.Row) + (grdAction.RowHeight(grdAction.Row) - pbcActionArrow.Height) / 2
        pbcActionArrow.Visible = True
        Select Case grdAction.Col
            Case ACTIONDATEINDEX  'Action Date
                txtDropdown.Move grdAction.Left + grdAction.ColPos(grdAction.Col) + 30, grdAction.Top + grdAction.RowPos(grdAction.Row) + 15, grdAction.ColWidth(grdAction.Col) - 30, grdAction.RowHeight(grdAction.Row) - 15
                If grdAction.Text <> "Missing" Then
                    txtDropdown.Text = grdAction.Text
                Else
                    txtDropdown.Text = ""
                End If
                If txtDropdown.Height > grdAction.RowHeight(grdAction.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdAction.RowHeight(grdAction.Row) - 15
                End If
                txtDropdown.Visible = True
                txtDropdown.SetFocus
            Case COMMENTINDEX  'Comment
                txtDropdown.Move grdAction.Left + grdAction.ColPos(grdAction.Col) + 30, grdAction.Top + grdAction.RowPos(grdAction.Row) + 15, grdAction.ColWidth(grdAction.Col) - 30, grdAction.RowHeight(grdAction.Row) - 15
                If grdAction.Text <> "Missing" Then
                    txtDropdown.Text = grdAction.Text
                Else
                    txtDropdown.Text = ""
                End If
                If txtDropdown.Height > grdAction.RowHeight(grdAction.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdAction.RowHeight(grdAction.Row) - 15
                End If
                txtDropdown.Visible = True
                txtDropdown.SetFocus
        End Select
    End If
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
Private Function OpenMsgFile() As Integer
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer
    Dim sLetter As String

    'On Error GoTo OpenMsgFileErr:
    'slToFile = "Mail.CSV"
    sLetter = "A"
    Do
        ilRet = 0
        smToFile = sgExportDirectory & "C" & Format$(gNow(), "mm") & Format$(gNow(), "dd") & Format$(gNow(), "yy") & sLetter & ".csv"
        'slDateTime = FileDateTime(smToFile)
        ilRet = gFileExist(smToFile)
        If ilRet = 0 Then
            sLetter = Chr$(Asc(sLetter) + 1)
        End If
    Loop While ilRet = 0
    On Error GoTo 0
    ilRet = 0
    'On Error GoTo OpenMsgFileErr:
    'hmMail = FreeFile
    'Open smToFile For Output As hmMail
    ilRet = gFileOpen(smToFile, "Output", hmMail)
    If ilRet <> 0 Then
        Close hmMail
        hmMail = -1
        gMsgBox "Open File " & smToFile & " error#" & Str$(Err.Number), vbOKOnly
        OpenMsgFile = False
        Exit Function
    End If
    On Error GoTo 0
    OpenMsgFile = True
    Exit Function
'OpenMsgFileErr:
'    ilRet = 1
'    Resume Next
End Function



Private Sub cmdCancel_GotFocus()
    mActionSetShow
End Sub

Private Sub cmdMail_Click()
    Dim iRet As Integer
    Dim sMail As String
    Dim iRow As Integer
    Dim iGetStation As Integer
    Dim sStationInfo As String
    Dim sStation As String
    Dim llRow As Long
    On Error GoTo ErrHand
    
    If grdPost.Rows - 1 <= grdPost.FixedRows Then
        Exit Sub
    End If
    iRet = OpenMsgFile()
    If iRet = False Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    iGetStation = True
    iRow = 0
    llRow = grdPost.FixedRows
    Do While llRow <= grdPost.Rows - 1
        If iGetStation = True Then
            iRow = Val(grdPost.TextMatrix(llRow, 7))
            sStation = Trim$(grdPost.TextMatrix(llRow, STATIONINDEX))
            SQLQuery = "SELECT shttAddress1, shttAddress2, shttCity, shttState, shttZip , shttFax FROM shtt WHERE (shttCode = " & tmContactInfo(iRow).iShttCode & ")"
            Set rst = gSQLSelectCall(SQLQuery)
            'sMail = "1-" & Trim$(rst(6).Value) & "," & Trim$(grdStations.Columns(3).Text) & "," & Trim$(grdStations.Columns(1).Text)
            'sMail = sMail & "," & Trim$(rst(0).Value) & "," & Trim$(rst(1).Value)
            'sMail = sMail & "," & Trim$(rst(2).Value) & "," & Trim$(rst(3).Value) & "," & Trim$(rst(4).Value)
            sMail = """" & sStation & """" & "," & """" & "1-" & Trim$(rst(5).Value) & """"
            'sStationInfo = sMail
            Print #hmMail, sMail
        End If
        'sMail = sStationInfo & "," & Trim$(grdStations.Columns(2).Text) & "," & Trim$(grdStations.Columns(0).Text)
        'Print #hmMail, sMail
        
        llRow = llRow + 1
        'If iRow <= grdStations.Rows - 1 Then
        '    If sStation <> Trim$(grdStations.Columns(1).Text) Then
        '        iGetStation = True
        '    End If
        'End If
    Loop
    Close hmMail
    Screen.MousePointer = vbDefault
    gMsgBox "Mail File Created Successfully, Its Name is " & smToFile, vbOKOnly
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Contact-cmdMail"
End Sub


Private Sub cmdMail_GotFocus()
    mActionSetShow
End Sub

Private Sub cmdSave_Click()
    Dim ilRet As Integer
    
    If (imCommentChgd = False) Then
        Exit Sub
    End If
    ilRet = mSave()

End Sub

Private Sub cmdSave_GotFocus()
    mActionSetShow
End Sub

Private Sub Form_Click()
    mActionSetShow
    If Not imCommentChgd Then
        lmPostRow = -1
        pbcArrow.Visible = False
    End If
End Sub



Private Sub Form_Load()
    
    Dim iCol As Integer
    Dim iRow As Integer
    Dim iUpper As Integer
    
    On Error GoTo ErrHand
    
    frmContact.Caption = "Contact - " & sgClientName
    smFWkDate = ""
    smLWkDate = ""
    imAllClick = False
    
    mFillAdvt

    'Me.Width = Screen.Width / 1.1   '1.3
    'Me.Height = Screen.Height / 1.5 '2.2
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    Screen.MousePointer = vbHourglass
    imCIndex = 1
    imCommentChgd = False
    imSIntegralSet = False
    imHeaderClick = False
    imFirstTime = True
    lmPostRow = -1
    imcPrt.Picture = frmDirectory!imcPrinter.Picture
    imCIntegralSet = False
    'SQLQuery = "SELECT cptt.cpttStartDate, shtt.shttCallLetters, vef.vefName, shttACName, shttACPhone, shttCode"
    'SQLQuery = SQLQuery + " FROM cptt, lst, shtt, vef"
    'SQLQuery = SQLQuery + " WHERE ((cptt.cpttStatus = 0) AND (vef.vefCode = lst.lstLogVefCode)"
    'SQLQuery = SQLQuery + " AND (shtt.shttCode = cptt.cpttShfCode)"
    'SQLQuery = SQLQuery + " AND " & sContracts & ")"
    'SQLQuery = SQLQuery + " ORDER BY shtt.shttCallLetters, vef.vefName, cptt.cpttStartDate"
    ReDim tmContactInfo(0 To 0) As CONTACTINFO
    iRow = 0
    iCol = 0
    iUpper = 0
    imCMax = 1
    ReDim tmCDate(0 To 0) As CONTACTDATE
    If iCol \ 3 + 1 > imCMax Then
        imCMax = iCol \ 3 + 1
    End If
    'GridPaint True
    If sgUstWin(8) <> "I" Then
        cmdSave.Enabled = False
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Contact - Form Load"
End Sub
Private Sub grdAction_Click()
    If sgUstWin(8) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdAction.Col > COMMENTINDEX Then
        Exit Sub
    End If
End Sub

Private Sub grdAction_EnterCell()
    mActionSetShow
    If sgUstWin(8) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
End Sub

Private Sub grdAction_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdAction.TopRow
    grdAction.Redraw = False
End Sub

Private Sub grdAction_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim ilFound As Integer
    
    If sgUstWin(8) <> "I" Then
        grdAction.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If (lmPostRow < grdPost.FixedRows) Or (lmPostRow >= grdPost.Rows) Then
        grdAction.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdAction, X, Y)
    If Not ilFound Then
        grdAction.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdAction.Col > COMMENTINDEX Then
        grdAction.Redraw = True
        Exit Sub
    End If
    
    If grdAction.Col = USERNAMEINDEX Then
        grdAction.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdAction.TopRow
    
    llRow = grdAction.Row
    If grdAction.TextMatrix(llRow, ACTIONDATEINDEX) = "" Then
        grdAction.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdAction.TextMatrix(llRow, ACTIONDATEINDEX) = ""
        grdAction.Row = llRow + 1
        grdAction.Col = ACTIONDATEINDEX
        grdAction.Redraw = True
    End If
    grdAction.Redraw = True
    mActionEnableBox
End Sub

Private Sub grdAction_Scroll()
    If grdAction.Redraw = False Then
        grdAction.Redraw = True
        grdAction.TopRow = lmTopRow
        grdAction.Refresh
        grdAction.Redraw = False
    End If
    If (imShowGridBox) And (grdAction.Row >= grdAction.FixedRows) And (grdAction.Col >= ACTIONDATEINDEX) And (grdAction.Col <= COMMENTINDEX) Then
        If grdAction.RowIsVisible(grdAction.Row) Then
            txtDropdown.Move grdAction.Left + grdAction.ColPos(grdAction.Col) + 30, grdAction.Top + grdAction.RowPos(grdAction.Row) + 30, grdAction.ColWidth(grdAction.Col) - 30, grdAction.RowHeight(grdAction.Row) - 30
            pbcActionArrow.Move grdAction.Left - pbcActionArrow.Width, grdAction.Top + grdAction.RowPos(grdAction.Row) + (grdAction.RowHeight(grdAction.Row) - pbcActionArrow.Height) / 2
            pbcActionArrow.Visible = True
            txtDropdown.Visible = True
            txtDropdown.SetFocus
        Else
            pbcActionFocus.SetFocus
            txtDropdown.Visible = False
            pbcActionArrow.Visible = False
        End If
    Else
        pbcActionFocus.SetFocus
        pbcActionArrow.Visible = False
        imFromArrow = False
    End If
End Sub

Private Sub grdPost_Click()
    Dim ilRet As Integer
    tmcDelay.Enabled = False
    If sgUstWin(8) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If imCommentChgd Then
        If gMsgBox("Save comment changes?", vbYesNo) = vbYes Then
            ilRet = mSave()
            If ilRet = False Then
                grdPost.Row = lmPostRow
                Exit Sub
            End If
        End If
    End If
    imCommentChgd = False
    DoEvents
    If (grdPost.Row - 1 < STATIONINDEX) Or (grdPost.Row - 1 >= UBound(tmContactInfo)) Then
        lmPostRow = -1
        pbcArrow.Visible = False
        Exit Sub
    End If
    If imHeaderClick Then
        imHeaderClick = False
        Exit Sub
    End If
    tmcDelay.Enabled = True
    Exit Sub
End Sub

Private Sub grdPost_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Determine if in header
    If sgUstWin(8) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If Y < grdPost.RowHeight(STATIONINDEX) Then
        imHeaderClick = True
        If (X >= grdPost.ColPos(DATE1INDEX)) And (X <= grdPost.ColPos(DATE1INDEX) + grdPost.ColWidth(DATE1INDEX)) Then
            imCIndex = imCIndex - 1
            If imCIndex < 1 Then
                imCIndex = 1
            Else
                GridPaint False
            End If
        End If
        If (X >= grdPost.ColPos(DATE3INDEX)) And (X <= grdPost.ColPos(DATE3INDEX) + grdPost.ColWidth(DATE3INDEX)) Then
            imCIndex = imCIndex + 1
            If imCIndex > imCMax Then
                imCIndex = imCMax
            Else
                GridPaint False
            End If
        End If
'        If (x >= grdPost.ColPos(8)) And (x <= grdPost.ColPos(8) + grdPost.ColWidth(8)) Then
'            mDateSort
'            mGridPaint False
'        End If
    End If
End Sub


Private Sub grdPost_Scroll()
    If (pbcArrow.Visible) And (grdPost.Row >= grdPost.FixedRows) And (grdPost.Col >= STATIONINDEX) And (grdPost.Col < grdPost.Cols - 1) Then
        If grdPost.RowIsVisible(grdPost.Row) Then
            pbcArrow.Move grdPost.Left - pbcActionArrow.Width - 15, grdPost.Top + grdPost.RowPos(grdPost.Row) + (grdPost.RowHeight(grdPost.Row) - pbcActionArrow.Height) / 2
            pbcArrow.Visible = True
        Else
            pbcPostFocus.SetFocus
            pbcArrow.Visible = False
        End If
    Else
        pbcPostFocus.SetFocus
        pbcArrow.Visible = False
    End If

End Sub

Private Sub imcPrt_Click()
    Dim llRow As Long
    Dim sContactPhone As String
    
    mActionSetShow
    If grdPost.Rows - 1 <= grdPost.FixedRows Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Printer.Print ""
    Printer.Print Tab(65); Format$(Now)
    Printer.Print ""
    Printer.Print " " & sAdvtDates
    Printer.Print ""
    Printer.Print "  Call Letters"; Tab(15); "Vehicle"; Tab(37); "Contact"; Tab(69); "Dates"
    
    For llRow = grdPost.FixedRows To grdPost.Rows - 1 Step 1
        sContactPhone = Trim$(grdPost.TextMatrix(llRow, CONTACTINDEX))
        If Len(sContactPhone) > 16 Then
            sContactPhone = Left$(sContactPhone, 30 - Len(Trim$(grdPost.TextMatrix(llRow, TELEPHONEINDEX)))) & " " & Trim$(grdPost.TextMatrix(llRow, TELEPHONEINDEX))
        Else
            sContactPhone = sContactPhone & " " & Trim$(grdPost.TextMatrix(llRow, TELEPHONEINDEX))
        End If
        Printer.Print "  " & Trim$(grdPost.TextMatrix(llRow, STATIONINDEX)); Tab(15); Trim$(grdPost.TextMatrix(llRow, VEHICLEINDEX)); Tab(37); sContactPhone; Tab(69); Trim$(grdPost.TextMatrix(llRow, DATE1INDEX)) & " " & Trim$(grdPost.TextMatrix(llRow, DATE2INDEX)) & " " & Trim$(grdPost.TextMatrix(llRow, DATE3INDEX))
    Next llRow
    Printer.EndDoc
    Screen.MousePointer = vbDefault
End Sub

Private Function mSave() As Integer
    Dim llRow As Long
    Dim sDate As String
    Dim sComment As String
    Dim ilError As Integer
    Dim sNowDate As String
    Dim sNowTime As String
    Dim iUstCode As Integer
    
    On Error GoTo ErrHand
    
    If sgUstWin(8) <> "I" Then
        gMsgBox "Not Allowed to Save.", vbOKOnly
        mSave = True
        Exit Function
    End If
    
    'D.S. 10/10/02 check for date and comments before saving
    ilError = False
    For llRow = grdAction.FixedRows To grdAction.Rows - 1 Step 1
        If (grdAction.TextMatrix(llRow, ACTIONDATEINDEX) <> "") Or (grdAction.TextMatrix(llRow, COMMENTINDEX) <> "") Then
            If (grdAction.TextMatrix(llRow, ACTIONDATEINDEX) = "") Then
                grdAction.TextMatrix(llRow, ACTIONDATEINDEX) = "Missing"
                ilError = True
            Else
                sDate = grdAction.TextMatrix(llRow, ACTIONDATEINDEX)
                If Not gIsDate(sDate) Then
                    grdAction.Row = llRow
                    grdAction.Col = ACTIONDATEINDEX
                    grdAction.CellForeColor = vbRed
                    ilError = True
                End If
            End If
            If (grdAction.TextMatrix(llRow, COMMENTINDEX) = "") Then
                grdAction.TextMatrix(llRow, COMMENTINDEX) = "Missing"
                ilError = True
            End If
        End If
    Next llRow
    If ilError Then
        mSave = False
        Exit Function
    End If
    SQLQuery = "DELETE "
    SQLQuery = SQLQuery + " FROM cct"
    SQLQuery = SQLQuery + " WHERE (cct.cctShfCode= " & imShttCode & ""
    SQLQuery = SQLQuery + " AND cct.cctVefCode = " & imVefCode & ")"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "Contact-mSave"
        cnn.RollbackTrans
        mSave = False
        Exit Function
    End If
    
    For llRow = grdAction.FixedRows To grdAction.Rows - 1 Step 1
        If (grdAction.TextMatrix(llRow, ACTIONDATEINDEX) <> "") Or (grdAction.TextMatrix(llRow, COMMENTINDEX) <> "") Then
            sDate = grdAction.TextMatrix(llRow, ACTIONDATEINDEX)
            sDate = Format$(sDate, sgShowDateForm)
            sComment = Trim$(grdAction.TextMatrix(llRow, COMMENTINDEX))
            sComment = gFixQuote(sComment) & Chr(0)
            'test for newly added comments that were added in existing rows (new lines not put out)
            sNowDate = grdAction.TextMatrix(llRow, DATEENTEREDINDEX)

            If sNowDate = "" Then
                sNowDate = smNowDate
                sNowTime = smNowTime
                iUstCode = igUstCode
            Else
                sNowDate = Format$(sNowDate, sgShowDateForm)
                sNowTime = grdAction.TextMatrix(llRow, TIMEENTEREDINDEX)
                sNowTime = Format$(sNowTime, sgShowTimeWSecForm)
                If grdAction.TextMatrix(llRow, CHANGEFLAGINDEX) = "Y" Then
                    iUstCode = igUstCode            'Current user change a previous comment
                Else
                    iUstCode = grdAction.TextMatrix(llRow, USTCODEINDEX)    'retain previous users code
                End If
            End If
            If (sDate <> "") And (sComment <> "") Then
            
                '7-22-09 Update user, date and time added
                SQLQuery = "INSERT INTO cct (cctShfCode, cctVefCode, cctActionDate, cctComment, cctUstCode, cctEnteredDate, cctEnteredTime)"
                SQLQuery = SQLQuery & " VALUES (" & imShttCode & ", " & imVefCode & ", '"
                SQLQuery = SQLQuery & Format$(sDate, sgSQLDateForm) & "','" & sComment & "', "
                SQLQuery = SQLQuery & iUstCode & ", '" & Format$(sNowDate, sgSQLDateForm) & "', '" & Format$(sNowTime, sgSQLTimeForm) & "' " & ")"
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "Contact-mSave"
                    cnn.RollbackTrans
                    mSave = False
                    Exit Function
                End If
            End If
        End If
    Next llRow
    cnn.CommitTrans
    mSave = True
    imCommentChgd = False
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Contact Grid-mSave"
    mSave = False
    Exit Function
End Function

Private Sub pbcActionSTab_GotFocus()
    If GetFocus() <> pbcActionSTab.hwnd Then
        Exit Sub
    End If
    If sgUstWin(8) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If (lmPostRow < grdPost.FixedRows) Or (lmPostRow >= grdPost.Rows) Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mActionEnableBox
        Exit Sub
    End If
    If txtDropdown.Visible Then
        mActionSetShow
        If grdAction.Col = ACTIONDATEINDEX Then
            If grdAction.Row > grdAction.FixedRows Then
                lmTopRow = -1
                grdAction.Row = grdAction.Row - 1
                If Not grdAction.RowIsVisible(grdAction.Row) Then
                    grdAction.TopRow = grdAction.TopRow - 1
                End If
                grdAction.Col = COMMENTINDEX
                mActionEnableBox
            Else
                pbcClickFocus.SetFocus
            End If
        Else
            grdAction.Col = grdAction.Col - 1
            If grdAction.Col = USERNAMEINDEX Then
                grdAction.Col = grdAction.Col - 1
            End If
            mActionEnableBox
        End If
    Else
        lmTopRow = -1
        grdAction.TopRow = grdAction.FixedRows
        grdAction.Col = ACTIONDATEINDEX
        grdAction.Row = grdAction.FixedRows
        mActionEnableBox
    End If

End Sub

Private Sub pbcActionTab_GotFocus()
    Dim llRow As Long
    
    If GetFocus() <> pbcActionTab.hwnd Then
        Exit Sub
    End If
    If sgUstWin(8) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If (lmPostRow < grdPost.FixedRows) Or (lmPostRow >= grdPost.Rows) Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If txtDropdown.Visible Then
        mActionSetShow
        If grdAction.Col = COMMENTINDEX Then
'            If grdAction.Row + 1 < grdAction.Rows Then
            llRow = grdAction.Rows
            Do
                llRow = llRow - 1
            Loop While grdAction.TextMatrix(llRow, ACTIONDATEINDEX) = ""
            llRow = llRow + 1
            If (grdAction.Row + 1 < llRow) Then
                lmTopRow = -1
                grdAction.Row = grdAction.Row + 1
                If Not grdAction.RowIsVisible(grdAction.Row) Then
                    grdAction.TopRow = grdAction.TopRow + 1
                End If
                grdAction.Col = ACTIONDATEINDEX
                If Trim$(grdAction.TextMatrix(grdAction.Row, ACTIONDATEINDEX)) <> "" Then
                    mActionEnableBox
                Else
                    imFromArrow = True
                    pbcActionArrow.Move grdAction.Left - pbcActionArrow.Width, grdAction.Top + grdAction.RowPos(grdAction.Row) + (grdAction.RowHeight(grdAction.Row) - pbcActionArrow.Height) / 2
                    pbcActionArrow.Visible = True
                    pbcActionArrow.SetFocus
                End If
            Else
                If txtDropdown.Text <> "" Then
                    lmTopRow = -1
                    If grdAction.Row + 1 >= grdAction.Rows Then
                        grdAction.AddItem ""
                        llRow = grdAction.Row + 1
                        grdAction.Row = llRow
                        grdAction.Col = USERNAMEINDEX
                        grdAction.CellBackColor = LIGHTYELLOW
                        grdAction.TextMatrix(llRow, USTCODEINDEX) = igUstCode
                        grdAction.TextMatrix(llRow, DATEENTEREDINDEX) = smNowDate
                        grdAction.TextMatrix(llRow, TIMEENTEREDINDEX) = smNowTime
                    Else
                        grdAction.Row = grdAction.Row + 1
                    End If
                    If Not grdAction.RowIsVisible(grdAction.Row) Then
                        grdAction.TopRow = grdAction.TopRow + 1
                    End If
                    grdAction.Col = ACTIONDATEINDEX
                    'mActionEnableBox
                    imFromArrow = True
                    pbcActionArrow.Move grdAction.Left - pbcActionArrow.Width, grdAction.Top + grdAction.RowPos(grdAction.Row) + (grdAction.RowHeight(grdAction.Row) - pbcActionArrow.Height) / 2
                    pbcActionArrow.Visible = True
                    pbcActionArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdAction.Col = grdAction.Col + 1
            If grdAction.Col = USERNAMEINDEX Then
                grdAction.Col = grdAction.Col + 1
            End If
            mActionEnableBox
        End If
    Else
        lmTopRow = -1
        grdAction.TopRow = grdAction.FixedRows
        grdAction.Col = ACTIONDATEINDEX
        grdAction.Row = grdAction.FixedRows
        mActionEnableBox
    End If
End Sub

Private Sub tmcDelay_Timer()
    Dim iIndex As Integer
    Dim llRow As Long
    
    tmcDelay.Enabled = False
    Screen.MousePointer = vbHourglass
    pbcArrow.Move grdPost.Left - pbcArrow.Width - 15, grdPost.Top + grdPost.RowPos(grdPost.Row) + (grdPost.RowHeight(grdPost.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    iIndex = Val(grdPost.TextMatrix(grdPost.Row, 7))
    imShttCode = tmContactInfo(iIndex).iShttCode
    imVefCode = tmContactInfo(iIndex).iVefCode
    SQLQuery = "SELECT cctActionDate, cctComment, cctCode, cctustCode, cctEnteredDate, cctEnteredTime, ustName FROM  cct LEFT OUTER JOIN ust on cctustcode = ustcode  WHERE (cctshfCode = " & imShttCode & " And cctVefCode = " & imVefCode & ")" & " ORDER By cctActionDate desc"
    Set rst = gSQLSelectCall(SQLQuery)
    'gGrid_Clear grdAction, True
    mActionClearGrid
    llRow = grdAction.FixedRows
    While Not rst.EOF
        If llRow + 1 > grdAction.Rows Then
            grdAction.AddItem ""
        End If
        grdAction.Row = llRow
        grdAction.TextMatrix(llRow, ACTIONDATEINDEX) = Trim$(rst!cctActionDate)
        grdAction.Col = USERNAMEINDEX
        grdAction.CellBackColor = LIGHTYELLOW
        
        grdAction.TextMatrix(llRow, COMMENTINDEX) = Trim$(rst!cctComment)
        grdAction.TextMatrix(llRow, CCTCODEINDEX) = rst!cctCode
        grdAction.TextMatrix(llRow, USTCODEINDEX) = rst!cctUstCode
        If rst!cctUstCode = 0 Then          'older comment, no user was saved with record
            grdAction.TextMatrix(llRow, USERNAMEINDEX) = ""
        Else
            grdAction.TextMatrix(llRow, USERNAMEINDEX) = rst!ustname
        End If
        If IsNull(rst!cctEnteredDate) Then
            grdAction.TextMatrix(llRow, DATEENTEREDINDEX) = "1/1/1970"
        Else
            grdAction.TextMatrix(llRow, DATEENTEREDINDEX) = Format$(rst!cctEnteredDate, sgShowDateForm)
        End If
        If IsNull(rst!cctEnteredTime) Then
            grdAction.TextMatrix(llRow, TIMEENTEREDINDEX) = "12:00:00AM"
        Else
            grdAction.TextMatrix(llRow, TIMEENTEREDINDEX) = Format$(rst!cctEnteredTime, sgShowTimeWSecForm)
        End If
        grdAction.TextMatrix(llRow, CHANGEFLAGINDEX) = ""   'set to Y when change found
        llRow = llRow + 1
        rst.MoveNext
    Wend
    If llRow >= grdAction.Rows Then
        grdAction.AddItem ""
        '3/19/10
        'grdAction.Row = llRow + 1
        'grdAction.Col = USERNAMEINDEX
        'grdAction.CellBackColor = LIGHTYELLOW
        'grdAction.TextMatrix(llRow + 1, USTCODEINDEX) = igUstCode
        'grdAction.TextMatrix(llRow + 1, DATEENTEREDINDEX) = smNowDate
        'grdAction.TextMatrix(llRow + 1, TIMEENTEREDINDEX) = smNowTime
        grdAction.Row = llRow
        grdAction.Col = USERNAMEINDEX
        grdAction.CellBackColor = LIGHTYELLOW
        grdAction.TextMatrix(llRow, USTCODEINDEX) = igUstCode
        grdAction.TextMatrix(llRow, DATEENTEREDINDEX) = smNowDate
        grdAction.TextMatrix(llRow, TIMEENTEREDINDEX) = smNowTime

    End If
    lmPostRow = grdPost.Row
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Contact Grid-grdPost"
End Sub

Private Sub txtDropdown_Change()
    Dim slStr As String
    Dim llRow As Long
    Select Case grdAction.Col
        Case ACTIONDATEINDEX
            slStr = Trim$(txtDropdown.Text)
            If (gIsDate(slStr)) And (slStr <> "") Then
                grdAction.CellForeColor = vbBlack
                If grdAction.Text <> txtDropdown.Text Then
                    imCommentChgd = True
                End If
                grdAction.Text = txtDropdown.Text
            End If
        Case COMMENTINDEX
            slStr = txtDropdown.Text
            grdAction.CellForeColor = vbBlack
            If grdAction.Text <> txtDropdown.Text Then
                imCommentChgd = True
                llRow = grdAction.RowSel
                grdAction.TextMatrix(llRow, CHANGEFLAGINDEX) = "Y"
            End If
            grdAction.Text = txtDropdown.Text
    End Select
End Sub

Private Sub txtDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub mRefreshGrid()

    Dim sPrevCall As String
    Dim sPrevDate As String
    Dim sPrevVeh As String
    Dim iCol As Integer
    Dim iRow As Integer
    Dim iUpper As Integer
    Dim iAddDate As Integer
    Dim rstContact As ADODB.Recordset      '7-22-09


    sPrevCall = ""
    sPrevDate = ""
    sPrevVeh = ""
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        ''If (sPrevCall <> Trim$(rst(1).Value)) Or (sPrevDate <> rst(0).Value) Or (sPrevVeh <> Trim$(rst(2).Value)) Then
        'If (sPrevCall <> Trim$(rst!shttCallLetters)) Or (sPrevDate <> rst!cpttStartDate) Or (sPrevVeh <> Trim$(rst!vefName)) Then
        If (sPrevCall <> Trim$(rst!shttCallLetters)) Or (sPrevVeh <> Trim$(rst!vefName)) Then
            'grdStations.AddItem "" & rst(0).Value & ", " & rst(1).Value & ", " & rst(2).Value & ", " & rst(3).Value & ", " & rst(4).Value & ""
            If iCol \ 3 + 1 > imCMax Then
                imCMax = iCol \ 3 + 1
            End If
            tmContactInfo(iRow).sStation = rst!shttCallLetters
            If sgShowByVehType = "Y" Then
                tmContactInfo(iRow).sVehicle = Trim$(rst!vefType) & ":" & rst!vefName
            Else
                tmContactInfo(iRow).sVehicle = rst!vefName
            End If
            
            '7-22-09 Pull contact from ARTT, not the agreement
            SQLQuery = "Select arttFirstName, arttLastName, arttPhone, arttaffcontact from artt where arttshttcode = " & Str$(rst!shttCode) & " and arttaffContact = " & "'1'"
            Set rstContact = gSQLSelectCall(SQLQuery)
            If Not rstContact.EOF Then
                If Not IsNull(rstContact!arttLastName) Then
                
                    'If Trim$(rst!attACName) <> "" Then
                    If Trim$(rstContact!arttLastName <> "" Or Trim$(rstContact!arttFirstName) <> "") Then
                        tmContactInfo(iRow).sACName = Trim$(rstContact!arttFirstName) + " " + Trim$(rstContact!arttLastName)   'rst!attACName
                        tmContactInfo(iRow).sACPhone = rstContact!arttPhone  'rst!attACPhone
                    Else
                        tmContactInfo(iRow).sACName = rst!shttACName
                        tmContactInfo(iRow).sACPhone = rst!shttACPhone
                    End If
                End If
            End If
            
            tmContactInfo(iRow).iShttCode = rst!shttCode
            tmContactInfo(iRow).iVefCode = rst!vefCode
            tmContactInfo(iRow).iCDateIndex = iUpper
            iCol = 0
            iRow = iRow + 1
            ReDim Preserve tmContactInfo(0 To iRow) As CONTACTINFO
            sPrevCall = Trim$(rst!shttCallLetters)
            sPrevDate = Format$(Trim$(rst!CpttStartDate), sgShowDateForm)
            sPrevVeh = Trim$(rst!vefName)
            iAddDate = True
        Else
            'When filter is by Advertiser/Contract, multi records for same date is caused by multi-spots within same week
            If DateValue(gAdjYear(sPrevDate)) <> DateValue(gAdjYear(Trim$(rst!CpttStartDate))) Then
                iAddDate = True
                sPrevDate = Format$(Trim$(rst!CpttStartDate), sgShowDateForm)
            Else
                iAddDate = False
            End If
        End If
        If iAddDate Then
            tmCDate(iUpper).iCol = iCol
            tmCDate(iUpper).sDate = Format$(rst!CpttStartDate, sgShowDateForm)
            tmCDate(iUpper).iPostingStatus = rst!cpttPostingStatus
            iUpper = iUpper + 1
            iCol = iCol + 1
            ReDim Preserve tmCDate(0 To iUpper) As CONTACTDATE
        End If
        rst.MoveNext
    Wend
        GridPaint True

End Sub
Public Sub mActionClearGrid()
    Dim llRow As Long
    
    gGrid_Clear grdAction, True
    
    For llRow = grdPost.FixedRows To grdAction.Rows - 1 Step 1
        grdAction.Row = llRow
        grdAction.Col = USERNAMEINDEX
        grdAction.CellBackColor = LIGHTYELLOW
    Next llRow
End Sub
