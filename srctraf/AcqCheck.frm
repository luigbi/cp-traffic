VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form AcqCheck 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   9105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "AcqCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8730
      Top             =   4440
   End
   Begin VB.CommandButton cmcUndo 
      Caption         =   "Undo All"
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
      Height          =   375
      Left            =   7200
      TabIndex        =   25
      Top             =   4545
      Width           =   1335
   End
   Begin VB.CheckBox ckcSet 
      Caption         =   "Set All Acquisition Spot Counts to Invoice Spot Counts"
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
      Height          =   210
      Left            =   240
      TabIndex        =   24
      Top             =   4290
      Width           =   5175
   End
   Begin VB.CommandButton cmcCSV 
      Caption         =   "Save as CSV"
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
      Height          =   375
      Left            =   5520
      TabIndex        =   22
      Top             =   4545
      Width           =   1335
   End
   Begin VB.CommandButton cmcSave 
      Caption         =   "Save"
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
      Height          =   375
      Left            =   3915
      TabIndex        =   20
      Top             =   4545
      Width           =   1335
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1845
      MaxLength       =   10
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2670
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.ComboBox cbcMonth 
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
      ItemData        =   "AcqCheck.frx":08CA
      Left            =   810
      List            =   "AcqCheck.frx":08F2
      TabIndex        =   2
      Top             =   315
      Width           =   1500
   End
   Begin VB.Frame frcInclude 
      Caption         =   "Include Acquisition (Select at least one from each Column)"
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
      Height          =   1065
      Left            =   2565
      TabIndex        =   5
      Top             =   225
      Width           =   5715
      Begin VB.CheckBox ckcInclude 
         Caption         =   "Fully Paid in Collection"
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
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   255
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox ckcInclude 
         Caption         =   "Not Fully Paid in Collections"
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
         Height          =   195
         Index           =   1
         Left            =   2610
         TabIndex        =   7
         Top             =   255
         Value           =   1  'Checked
         Width           =   2790
      End
      Begin VB.CheckBox ckcInclude 
         Caption         =   "Non-Zero Dollars"
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
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   495
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin VB.CheckBox ckcInclude 
         Caption         =   "Zero Dollars"
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
         Height          =   195
         Index           =   3
         Left            =   2610
         TabIndex        =   9
         Top             =   495
         Value           =   1  'Checked
         Width           =   2025
      End
      Begin VB.CheckBox ckcInclude 
         Caption         =   "Spot Count of Zero"
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
         Height          =   195
         Index           =   4
         Left            =   2595
         TabIndex        =   11
         Top             =   735
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.CheckBox ckcInclude 
         Caption         =   "Spot Count of Non-Zero"
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
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   735
         Value           =   1  'Checked
         Width           =   2475
      End
   End
   Begin VB.TextBox edcYear 
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
      Height          =   300
      Left            =   810
      MaxLength       =   4
      TabIndex        =   4
      Top             =   765
      Width           =   1170
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   -30
      ScaleHeight     =   45
      ScaleWidth      =   60
      TabIndex        =   12
      Top             =   285
      Width           =   60
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   15
      Top             =   4860
      Width           =   45
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2220
      TabIndex        =   19
      Top             =   4545
      Width           =   1335
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4725
      Width           =   45
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   60
      Picture         =   "AcqCheck.frx":0958
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.CommandButton cmcCheck 
      Caption         =   "Check"
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
      Height          =   375
      Left            =   555
      TabIndex        =   18
      Top             =   4545
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAcqCheckGrid 
      Height          =   2775
      Left            =   225
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1500
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   3
      Cols            =   16
      FixedRows       =   2
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
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
      _Band(0).Cols   =   16
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lacAdjTotal 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   3975
      TabIndex        =   23
      Top             =   4320
      Width           =   5085
   End
   Begin VB.Label lacTotal 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   6465
      TabIndex        =   21
      Top             =   0
      Width           =   1770
   End
   Begin VB.Label lacMonth 
      Appearance      =   0  'Flat
      Caption         =   "Month"
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   375
      Width           =   885
   End
   Begin VB.Label lacYear 
      Appearance      =   0  'Flat
      Caption         =   "Year"
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
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   810
      Width           =   705
   End
   Begin VB.Label plcScreen 
      Caption         =   "Acquisition Check"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   2025
   End
End
Attribute VB_Name = "AcqCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************************************
'*  AcqCheck
'*
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFirstTime As Integer
Private imBSMode As Integer
Private imMouseDown As Integer
Private imCtrlKey As Integer
Private imShiftKey As Integer
Private imTerminate As Integer
Private lmLastClickedRow As Long
Private lmScrollTop As Long
Private imIgnoreScroll As Integer
Private lmEnableRow As Long
Private lmEnableCol As Long
Private imCtrlVisible As Integer
Private smNowDate As String
Private lmNowDate As Long
Private smMonthStart As String
Private smMonthEnd As String
Private lmMonthStart As Long
Private lmMonthEnd As Long
Private imChg As Integer
Private lmTopRow As Long
Private imFromArrow As Integer
Private lmLastStdMnthBilled As Long
Private smTotalAdjDollar As String
Private bmInSetAll As Boolean
Private bmCSVPressed As Boolean
Private lmNoRowsShown As Long
Private imPasswordOk As Integer

Private hmAddApf As Integer
Private hmAdjOrder As Integer
Private hmSave As Integer

Private imColSorted As Integer
Private imSort As Integer

Private rst_Temp As ADODB.Recordset
Private rst_Apf As ADODB.Recordset
Private rst_Sdf As ADODB.Recordset
Private rst_Iihf As ADODB.Recordset
Private rst_Rvf As ADODB.Recordset
Private rst_Prf As ADODB.Recordset
Private rst_Cff As ADODB.Recordset

Dim hmApf As Integer
Dim tmApf As APF        'CFF record image
Dim tmApfSrchKey0 As LONGKEY0    'CFF key record image
Dim tmApfSrchKey4 As APFKEY4
Dim imApfRecLen As Integer        'CFF record length

Dim hmSdf As Integer    'file handle
Dim imSdfRecLen As Integer  'Record length
Dim tmSdf As SDF
Dim tmSdfSrchKey0 As SDFKEY0
Dim tmSdfSrchKey3 As LONGKEY0

'Contract
Dim tmChf As CHF        'Chf record image
Dim tmChfSrchKey0 As LONGKEY0    'Chf key record image
Dim hmCHF As Integer    'Contract file handle
Dim imCHFRecLen As Integer        'CHF record length

Dim hmClf As Integer        'Contract line file handle
Dim tmClf As CLF            'CLF record image
Dim tmClfSrchKey0 As CLFKEY0 'CLF key record image
Dim tmClfSrchKey1 As CLFKEY1 'CLF key record image
Dim imClfRecLen As Integer     'CLF record length

Dim hmCff As Integer
Dim tmCff As CFF        'CFF record image
Dim tmCffSrchKey0 As CFFKEY0    'CFF key record image
Dim tmCffSrchKey1 As LONGKEY0
Dim imCffRecLen As Integer        'CFF record length

Dim hmSsf As Integer        'Spot summary file handle
Dim lmSsfDate(0 To 6) As Long    'Dates of the days stored into tmSsf
Dim lmSsfRecPos(0 To 6) As Long  'Record positions
Dim tmSsf(0 To 6) As SSF         'Spot summary for one week (0 index for monday; 1 for tuesday;...; 6 for sunday)
Dim tmSsfSrchKey As SSFKEY0 'SSF key record image
Dim imSsfRecLen As Integer     'SSF record length
Dim imSelectedDay As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS

Dim hmSmf As Integer
Dim tmSmf As SMF
Dim imSmfRecLen As Integer
Dim tmSmfSrchKey2 As LONGKEY0   'SdfCode

Dim hmCrf As Integer
Dim hmGsf As Integer
Dim hmGhf As Integer
Dim hmSxf As Integer

Dim tmRdf As RDF
Private Type APFORDERADJ
    lApfCode As Long
    iLineCount As Integer
    sStationInvNo As String * 20
    lCntrNo As Long
    iAdfCode As Integer
    lAcqCost As Long
    sVehicleName As String * 40
End Type
Dim tmApfOrderAdj() As APFORDERADJ

Const CNTRNOINDEX = 0
Const ADVTNAMEINDEX = 1
Const STATIONINDEX = 2
Const STATIONINVNOINDEX = 3
'Const ORDERCOUNTINDEX = 4
Const FULLYPAIDDATEINDEX = 4    '5
Const ACQCOSTINDEX = 5  '6
Const ORDERCOUNTINDEX = 6   '4
Const INVCOUNTINDEX = 7 '8
Const ACQSPOTCOUNTINDEX = 8 '7
'Const INVCOUNTINDEX = 8
Const ADJDOLLARINDEX = 9
Const SELECTEDINDEX = 10
Const ORIGACQSPOTCOUNTINDEX = 11
Const ORIGINVCOUNTINDEX = 12
Const SORTINDEX = 13
Const CHFCODEINDEX = 14
Const APFCODEINDEX = 15

Private Sub cbcMonth_Change()
    mClearGrid
    mSetCommands
End Sub

Private Sub cbcMonth_Click()
    cbcMonth_Change
End Sub

Private Sub cbcMonth_GotFocus()
    mSetShow
End Sub

Private Sub ckcInclude_GotFocus(Index As Integer)
    mSetShow
End Sub

Private Sub ckcSet_Click()
    Dim llRow As Long
    Dim slStr As String
    Dim llEnableRow As Long
    Dim llEnableCol As Long
    
    If ckcSet.Value = vbUnchecked Then
        Exit Sub
    End If
    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbHourglass
    bmInSetAll = True
    grdAcqCheckGrid.Redraw = False
    For llRow = grdAcqCheckGrid.FixedRows To grdAcqCheckGrid.Rows - 1 Step 1
        slStr = Trim$(grdAcqCheckGrid.TextMatrix(llRow, CNTRNOINDEX))
        If slStr <> "" Then
            If grdAcqCheckGrid.TextMatrix(llRow, ACQSPOTCOUNTINDEX) = grdAcqCheckGrid.TextMatrix(llRow, ORIGACQSPOTCOUNTINDEX) Then
                lmEnableRow = llRow
                lmEnableCol = ACQSPOTCOUNTINDEX
                edcDropDown.Text = grdAcqCheckGrid.TextMatrix(llRow, INVCOUNTINDEX)
                mSetShow
            End If
            If grdAcqCheckGrid.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                grdAcqCheckGrid.TextMatrix(llRow, SELECTEDINDEX) = "0"
                mPaintRowColor llRow
            End If
        End If
    Next llRow
    bmInSetAll = False
    imChg = True
    mSetCommands
    grdAcqCheckGrid.Redraw = True
    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbDefault

End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    mSetShow
End Sub

Private Sub cmcCheck_Click()
    Dim ilRet As Integer
    Dim slDate As String
    Dim llLastStdMnth As Long
    
    If cbcMonth.ListIndex < 0 Then
        ilRet = MsgBox("Month must be specified", vbOKOnly + vbExclamation, "Incomplete")
        cbcMonth.SetFocus
        Exit Sub
    End If
    If Trim$(edcYear.Text) = "" Then
        ilRet = MsgBox("Year must be specified", vbOKOnly + vbExclamation, "Incomplete")
        edcYear.SetFocus
        Exit Sub
    End If
    slDate = cbcMonth.ListIndex + 1 & "/15/" & edcYear.Text
    smMonthStart = gObtainStartStd(slDate)
    smMonthEnd = gObtainEndStd(slDate)
    lmMonthStart = gDateValue(smMonthStart)
    lmMonthEnd = gDateValue(smMonthEnd)
    gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), llLastStdMnth
    If lmMonthEnd > llLastStdMnth Then
        ilRet = MsgBox("The Month Year must be on or before " & Format(llLastStdMnth, "mmmm yyyy"), vbOKOnly + vbExclamation, "Incomplete")
        cbcMonth.SetFocus
        Exit Sub
    End If
    mPopulate
End Sub

Private Sub cmcCheck_GotFocus()
    mSetShow
End Sub

Private Sub cmcCSV_Click()
    Dim hlCSV As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim ilCol As Integer
    Dim slFileName As String
    Dim ilRet As Integer
    
    If cbcMonth.ListIndex < 0 Then
        ilRet = MsgBox("Month must be specified", vbOKOnly + vbExclamation, "Incomplete")
        cmcCheck.SetFocus
        Exit Sub
    End If
    If Trim$(edcYear.Text) = "" Then
        ilRet = MsgBox("Year must be specified", vbOKOnly + vbExclamation, "Incomplete")
        edcYear.SetFocus
        Exit Sub
    End If
    slFileName = sgDBPath & "Messages\" & "Acquisition_Check_" & cbcMonth.List(cbcMonth.ListIndex) & "-" & Trim$(edcYear.Text) & ".csv"
    gLogMsgWODT "OA", hlCSV, slFileName
    
    gLogMsgWODT "W", hlCSV, "Check ran on " & Format(Now, "m/d/yy") & " at " & Format(Now, "h:mm:ssAM/PM")
    gLogMsgWODT "W", hlCSV, "Contract #,Advertiser Name,Call Letters,Station Invoice #,Agency Fully Paid Date,Acquisition Cost,Ordered Count,Invoice Spot Count,Acquisition Spot Count,Adjusted Dollars"
    For llRow = grdAcqCheckGrid.FixedRows To grdAcqCheckGrid.Rows - 1 Step 1
        slStr = Trim$(grdAcqCheckGrid.TextMatrix(llRow, CNTRNOINDEX))
        If slStr <> "" Then
            For ilCol = ADVTNAMEINDEX To ADJDOLLARINDEX Step 1
                If ilCol = ADVTNAMEINDEX Then
                    slStr = slStr & "," & """" & Trim$(grdAcqCheckGrid.TextMatrix(llRow, ilCol)) & """"
                Else
                    slStr = slStr & "," & Trim$(grdAcqCheckGrid.TextMatrix(llRow, ilCol))
                End If
            Next ilCol
            gLogMsgWODT "W", hlCSV, slStr
        End If
    Next llRow

    gLogMsgWODT "C", hlCSV, ""
    bmCSVPressed = True
    ilRet = MsgBox("CSV created File: " & slFileName, vbOKOnly + vbExclamation, "Information")
End Sub

Private Sub cmcCSV_GotFocus()
    mSetShow
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim slFileName As String
    
    'If Not bmCSVPressed Then
    '    ilRet = MsgBox("Press 'Save to CSV' button prior to 'Save' button to retain reconcile information.  Continue without 'Saving to CSV'?", vbYesNo + vbDefaultButton2 + vbQuestion, "Warning")
    '    If ilRet = vbNo Then
    '        Exit Sub
    '    End If
    'End If
    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbHourglass
    smTotalAdjDollar = ""
    lacAdjTotal.Caption = ""
    
    slFileName = sgDBPath & "Messages\" & "Acquisition_Save_" & cbcMonth.List(cbcMonth.ListIndex) & "-" & Trim$(edcYear.Text) & ".csv"
    gLogMsgWODT "OA", hmSave, slFileName
    
    gLogMsgWODT "W", hmSave, "Check ran on " & Format(Now, "m/d/yy") & " at " & Format(Now, "h:mm:ssAM/PM")
    gLogMsgWODT "W", hmSave, "Contract #,Advertiser Name,Call Letters,Station Invoice #,Agency Fully Paid Date,Acquisition Cost,Ordered Count,Invoice Spot Count,Acquisition Spot Count,Adjusted Dollars"
    ilRet = mSaveRec()
    gLogMsgWODT "C", hmSave, ""
    
    imChg = False
    ckcSet.Value = vbUnchecked
    mPopulate
    mSetCommands
    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbDefault
    ilRet = MsgBox("Saved rows placed into: " & slFileName, vbOKOnly + vbExclamation, "Information")
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub cmcUndo_Click()
    Dim llRow As Long
    Dim slStr As String
    
    smTotalAdjDollar = ""
    lacAdjTotal.Caption = ""
    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbHourglass
    grdAcqCheckGrid.Redraw = False
    For llRow = grdAcqCheckGrid.FixedRows To grdAcqCheckGrid.Rows - 1 Step 1
        slStr = Trim$(grdAcqCheckGrid.TextMatrix(llRow, CNTRNOINDEX))
        If slStr <> "" Then
            grdAcqCheckGrid.TextMatrix(llRow, ACQSPOTCOUNTINDEX) = grdAcqCheckGrid.TextMatrix(llRow, ORIGACQSPOTCOUNTINDEX)
            grdAcqCheckGrid.TextMatrix(llRow, ADJDOLLARINDEX) = ""
            lacAdjTotal.Caption = ""
            If grdAcqCheckGrid.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                grdAcqCheckGrid.TextMatrix(llRow, SELECTEDINDEX) = "0"
                mPaintRowColor llRow
            End If
        End If
    Next llRow
    ckcSet.Value = vbUnchecked
    imChg = False
    mSetCommands
    grdAcqCheckGrid.Redraw = True
    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbDefault
End Sub

Private Sub cmcUndo_GotFocus()
    mSetShow
End Sub

Private Sub edcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim slMaxValue As String
    Dim slStr As String
    Dim ilSoldAdj As Integer
    
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case lmEnableCol
        Case ACQSPOTCOUNTINDEX, INVCOUNTINDEX
            slMaxValue = grdAcqCheckGrid.TextMatrix(lmEnableRow, ORDERCOUNTINDEX)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            If gCompNumberStr(slStr, slMaxValue) > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select

End Sub

Private Sub edcYear_Change()
    mClearGrid
    mSetCommands
End Sub

Private Sub edcYear_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcYear_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    
    If imFirstTime Then
        imFirstTime = False
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = (CLng(60) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    Me.Height = (CLng(75) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    gCenterStdAlone AcqCheck
    DoEvents
    mSetControls
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    mInit
    sgPasswordAddition = "Replace: Obtain the New Keycode from Counterpoint Software"
    If (Trim$(tgUrf(0).sName) = sgCPName) Then
        imPasswordOk = True
    Else
        CSPWord.Show vbModal
        imPasswordOk = igPasswordOk
        If Not imPasswordOk Then
            imTerminate = True
        End If
    End If
    If imTerminate = True Then
        tmcTerminate.Enabled = True
    End If
    Screen.MousePointer = vbDefault
    Exit Sub

End Sub

Private Sub Form_Terminate()
    Dim ilRet As Integer
    
    On Error Resume Next
    
    'Erase tmUnpostedCntrInfo
    Erase tmApfOrderAdj
    
    ilRet = btrClose(hmApf)
    btrDestroy hmApf
    
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    
    ilRet = btrClose(hmSxf)
    btrDestroy hmSxf
    
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf
    
    rst_Sdf.Close
    rst_Apf.Close
    rst_Temp.Close
    rst_Iihf.Close
    rst_Rvf.Close
    rst_Prf.Close
    rst_Cff.Close
    igManUnload = YES
    Set AcqCheck = Nothing   'Remove data segment
    igManUnload = NO
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set AcqCheck = Nothing
End Sub


Private Sub grdAcqCheckGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If
    If (Shift And SHIFTMASK) > 0 Then
        imShiftKey = True
    Else
        imShiftKey = False
    End If
End Sub

Private Sub grdAcqCheckGrid_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
    imShiftKey = False
End Sub

Private Sub mInit()
    Dim ilRet As Integer
    Dim ilAdf As Integer
    Dim blRet As Boolean

    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbHourglass
    imTerminate = False
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmLastClickedRow = -1
    lmScrollTop = grdAcqCheckGrid.FixedRows
    imColSorted = -1
    imSort = -1
    imCtrlVisible = False
    lmEnableRow = -1
    lmEnableCol = -1
    imChg = False
    imIgnoreScroll = False
    bmInSetAll = False
    bmCSVPressed = False
    lmNoRowsShown = -1

    hmApf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmApf, "", sgDBPath & "Apf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", AcqCheck
    On Error GoTo 0
    imApfRecLen = Len(tmApf)  'Get and save ARF record length

    hmSdf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", AcqCheck
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)  'Get and save ARF record length
    
    hmCHF = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", AcqCheck
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)  'Get and save ARF record length
    
    hmClf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", AcqCheck
    On Error GoTo 0
    imClfRecLen = Len(tmClf)  'Get and save ARF record length
    
    hmCff = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", AcqCheck
    On Error GoTo 0
    imCffRecLen = Len(tmCff)  'Get and save ARF record length
    
    hmSsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", AcqCheck
    On Error GoTo 0
    imSsfRecLen = Len(tmSsf(0))  'Get and save ARF record length

    hmSmf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", AcqCheck
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)

    hmCrf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", AcqCheck
    On Error GoTo 0

    hmGsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", AcqCheck
    On Error GoTo 0

    hmGhf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", AcqCheck
    On Error GoTo 0

    hmSxf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSxf, "", sgDBPath & "Sxf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", AcqCheck
    On Error GoTo 0
    
    gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), lmLastStdMnthBilled
    
    ilRet = gObtainRdf(sgMRdfStamp, tgMRdf())
    
    blRet = gBuildAcqCommInfo(AcqCheck)

    mClearGrid

    Screen.MousePointer = vbDefault
    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbDefault
    Exit Sub

End Sub

Private Sub mPopulate()
    Dim llCol As Long
    Dim slSQLQuery As String
    Dim llRow As Long
    Dim slDate As String
    Dim slAcqCost As String
    Dim ilAdf As Integer
    Dim blInclude As Boolean
    Dim llDate01011970 As Long
    Dim ilTotal As Integer
    Dim ilApf As Integer
    Dim slStr As String
    Dim ilCount As Integer
    Dim slStnInvoiceNo As String
    Dim ilVff As Integer
    Dim slPostLogSource As String
    Dim ilOrderCountUpdated As Integer
    Dim blRet As Boolean
    ReDim tmApfOrderAdj(0 To 0) As APFORDERADJ
Dim ilRet As Integer
    On Error GoTo ErrHand:

    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbHourglass

    grdAcqCheckGrid.Redraw = False

    blRet = mFindAndFixMissingApf()

    llDate01011970 = gDateValue("1/1/1970")
    ilTotal = 0
    lacTotal.Caption = ""
    lacAdjTotal.Caption = ""
    On Error Resume Next
    slSQLQuery = "Drop Table #sdfCount"
    'Set rst_Temp = cnn.Execute(slSQLQuery)
    gSQLCallIgnoreError slSQLQuery
    On Error GoTo ErrHand:

    mClearGrid
    'grdAcqCheckGrid.Row = 0
    'For llCol = CNTRNOINDEX To ADJDOLLARINDEX Step 1
    '    grdAcqCheckGrid.Col = llCol
    '    grdAcqCheckGrid.CellBackColor = vbBlack
    '    grdAcqCheckGrid.CellBackColor = LIGHTBLUE
    'Next llCol

    ilOrderCountUpdated = 0

    slSQLQuery = "Select chfCntrNo, chfCode, chfAdfCode, sdfVefCode, clfAcquisitionCost, Count(sdfCode) As SpotCount, "
    slSQLQuery = slSQLQuery & "Count(If((sdfSchStatus = 'S') Or (sdfSchStatus = 'G' And smfActualDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' And smfActualDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "'" & ") Or (sdfSchStatus = 'O' And smfActualDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' And smfActualDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "'" & "), 1, Null)) as SpotAiredCount"
    slSQLQuery = slSQLQuery & " Into #sdfCount From sdf_Spot_Detail"
    slSQLQuery = slSQLQuery & " Inner Join chf_Contract_Header On sdfChfCode = chfCode"
    slSQLQuery = slSQLQuery & " Inner Join clf_Contract_Line On sdfChfCode = clfChfCode and clfLine = sdfLineNo and clfCntRevNo = chfCntRevNo"
    slSQLQuery = slSQLQuery & " Inner Join vff_Vehicle_Features On sdfVefCode = vffVefCode"
    slSQLQuery = slSQLQuery & " Left Outer Join smf_Spot_MG_Specs On sdfSmfCode = smfCode"
    'slSQLQuery = slSQLQuery & " where vffPostLogSource = 'S'"
    'slSQLQuery = slSQLQuery & " And (sdfDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' And sdfDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "')"
    slSQLQuery = slSQLQuery & " where "
    'slSQLQuery = slSQLQuery & " (sdfDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' And sdfDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "')"
    'slSQLQuery = slSQLQuery & " And (sdfSpotType <> 'X')"
    slSQLQuery = slSQLQuery & "(sdfSpotType <> 'X')"
    slSQLQuery = slSQLQuery & " And (((sdfSchStatus = 'S') and (sdfDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' And sdfDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "')" & ")"
    slSQLQuery = slSQLQuery & " Or ((sdfSchStatus = 'M') and (sdfDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' And sdfDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "')" & ")"
    slSQLQuery = slSQLQuery & " Or ((sdfSchStatus = 'O' ) and (smfMissedDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' And smfMissedDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "')" & ")"
    slSQLQuery = slSQLQuery & " Or ((sdfSchStatus = 'G' ) and (smfMissedDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' And smfMissedDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "')" & ")" & ")"
    slSQLQuery = slSQLQuery & " group by chfCntrNo, chfCode, chfAdfCode, sdfVefCode, clfAcquisitionCost "   'Having SpotAiredCount > 0"

    'Set rst_Temp = cnn.Execute(slSQLQuery)
    Set rst_Temp = gSQLSelectCall(slSQLQuery)

    'slSQLQuery = "Select apfCntrNo, sdfCount.chfCode as chfCode, sdfCount.chfAdfCode as chfAdfCode, vefName, vefCode, sdfCount.clfAcquisitionCost, apfAiredSpotCount, sdfCount.SpotAiredCount as InvSpotCount, apfOrderSpotCount, sdfCount.SpotCount as LineCount, iihfSourceForm, iihfStnInvoiceNo, apfInvDate, apfFullyPaidDate, apfCode"
    slSQLQuery = "Select apfCntrNo, apfStationInvNo, #sdfCount.chfCode as chfCode, #sdfCount.chfAdfCode as chfAdfCode, vefName, vefCode, #sdfCount.clfAcquisitionCost, apfAiredSpotCount, apfAcquisitionCost, #sdfCount.SpotAiredCount as InvSpotCount, apfOrderSpotCount, #sdfCount.SpotCount as LineCount, apfInvDate, apfFullyPaidDate, apfCode"
    slSQLQuery = slSQLQuery & " From #SdfCount"
    'slSQLQuery = slSQLQuery & " Inner Join apf_acq_payable On apfCntrNo = sdfCount.ChfCntrNo And apfVefCode = sdfCount.sdfVefCode And apfAcquisitionCost = sdfCount.clfAcquisitionCost and apfOrderSpotCount = SdfCount.SpotCount"
    slSQLQuery = slSQLQuery & " Inner Join apf_acq_payable On apfCntrNo = #sdfCount.ChfCntrNo And apfVefCode = #sdfCount.sdfVefCode And apfAcquisitionCost = #sdfCount.clfAcquisitionCost"
    slSQLQuery = slSQLQuery & " Inner Join vef_Vehicles On apfVefCode = vefCode"
    'slSQLQuery = slSQLQuery & " Inner Join iihf_ImptInvHeader On iihfChfCode = sdfCount.chfCode And iihfVefCode = apfVefCode And iihfInvStartDate = '" & Format(smMonthStart, "yyyy-mm-dd") & "'"
    '4/20/20: added date range
    slSQLQuery = slSQLQuery & " where apfInvDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "'"
    slSQLQuery = slSQLQuery & " And apfInvDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "'"
    slSQLQuery = slSQLQuery & " And apfMnfItem = 0"
    slSQLQuery = slSQLQuery & " And ((apfAiredSpotCount <> #sdfCount.SpotAiredCount)"
    slSQLQuery = slSQLQuery & " Or (apfOrderSpotCount <> #sdfCount.SpotCount))"
    slSQLQuery = slSQLQuery & " Order By apfCntrNo, vefName, #sdfCount.clfAcquisitionCost"
    'Set rst_Apf = cnn.Execute(slSQLQuery)
    Set rst_Apf = gSQLSelectCall(slSQLQuery)
    llRow = grdAcqCheckGrid.FixedRows
    Do While Not rst_Apf.EOF
        blInclude = True
        slDate = Format(rst_Apf!apfFullyPaidDate, "m/d/yy")
        slAcqCost = gLongToStrDec(rst_Apf!clfAcquisitionCost, 2)
        If ckcInclude(0).Value = vbUnchecked Then 'Fully Paid
            If gDateValue(slDate) <> llDate01011970 Then
                blInclude = False
            End If
        End If
        If ckcInclude(1).Value = vbUnchecked Then  'Not Fully Paid
            If gDateValue(slDate) = llDate01011970 Then
                blInclude = False
            End If
        End If
        If ckcInclude(2).Value = vbUnchecked Then  'Non-zero dollars
            If rst_Apf!clfAcquisitionCost > 0 Then
                blInclude = False
            End If
        End If
        If ckcInclude(3).Value = vbUnchecked Then  'Zero dollars
            If rst_Apf!clfAcquisitionCost = 0 Then
                blInclude = False
            End If
        End If
        If ckcInclude(4).Value = vbUnchecked Then  'Spot count of zero
            If rst_Apf!apfAiredSpotCount = 0 Then
                blInclude = False
            End If
        End If
        If ckcInclude(5).Value = vbUnchecked Then  'Spot count of non-zero
            If rst_Apf!apfAiredSpotCount <> 0 Then
                blInclude = False
            End If
        End If
        If blInclude Then
            slStnInvoiceNo = ""
            ilVff = gBinarySearchVff(rst_Apf!vefCode)
            If ilVff <> -1 Then
                slPostLogSource = tgVff(ilVff).sPostLogSource
            Else
                slPostLogSource = "S"
            End If
            If slPostLogSource = "S" Then
                slSQLQuery = "Select iihfSourceForm, iihfStnInvoiceNo"
                slSQLQuery = slSQLQuery & " From iihf_ImptInvHeader"
                slSQLQuery = slSQLQuery & " Where iihfChfCode = " & rst_Apf!chfCode
                slSQLQuery = slSQLQuery & " And iihfVefCode = " & rst_Apf!vefCode
                slSQLQuery = slSQLQuery & " And iihfInvStartDate = '" & Format(smMonthStart, "yyyy-mm-dd") & "'"
                'Set rst_Iihf = cnn.Execute(slSQLQuery)
                Set rst_Iihf = gSQLSelectCall(slSQLQuery)
                If Not rst_Iihf.EOF Then
                    slStnInvoiceNo = rst_Iihf!iihfStnInvoiceNo
                End If
            End If
            If (rst_Apf!apfOrderSpotCount <> rst_Apf!LineCount) Or (rst_Apf!apfStationInvNo <> slStnInvoiceNo) Then
                If rst_Apf!apfOrderSpotCount <> rst_Apf!LineCount Then
                    ilOrderCountUpdated = ilOrderCountUpdated + 1
                End If
                tmApfOrderAdj(UBound(tmApfOrderAdj)).lApfCode = rst_Apf!apfCode
                tmApfOrderAdj(UBound(tmApfOrderAdj)).iLineCount = rst_Apf!LineCount
                tmApfOrderAdj(UBound(tmApfOrderAdj)).sStationInvNo = slStnInvoiceNo
                tmApfOrderAdj(UBound(tmApfOrderAdj)).lCntrNo = rst_Apf!apfCntrno
                tmApfOrderAdj(UBound(tmApfOrderAdj)).iAdfCode = rst_Apf!chfAdfCode
                tmApfOrderAdj(UBound(tmApfOrderAdj)).lAcqCost = rst_Apf!apfAcquisitionCost
                tmApfOrderAdj(UBound(tmApfOrderAdj)).sVehicleName = rst_Apf!VEFNAME
                ReDim Preserve tmApfOrderAdj(0 To UBound(tmApfOrderAdj) + 1) As APFORDERADJ
            End If
            If rst_Apf!apfAiredSpotCount = rst_Apf!InvSpotCount Then
                blInclude = False
            End If
        End If
        If blInclude Then
            ilTotal = ilTotal + 1
            If llRow >= grdAcqCheckGrid.Rows Then
                grdAcqCheckGrid.AddItem ""
            End If
            grdAcqCheckGrid.RowHeight(llRow) = fgFlexGridRowH
            grdAcqCheckGrid.TextMatrix(llRow, SELECTEDINDEX) = "0"
            mPaintRowColor llRow
            grdAcqCheckGrid.TextMatrix(llRow, CNTRNOINDEX) = rst_Apf!apfCntrno
            ilAdf = gBinarySearchAdf(rst_Apf!chfAdfCode)
            If ilAdf <> -1 Then
                grdAcqCheckGrid.TextMatrix(llRow, ADVTNAMEINDEX) = Trim$(tgCommAdf(ilAdf).sName)
            End If
            grdAcqCheckGrid.TextMatrix(llRow, STATIONINDEX) = Trim$(rst_Apf!VEFNAME)

            If slPostLogSource = "S" Then
                slStr = ""
                ilCount = 0
                'moved above
                'slSQLQuery = "Select iihfSourceForm, iihfStnInvoiceNo"
                'slSQLQuery = slSQLQuery & " From iihf_ImptInvHeader"
                'slSQLQuery = slSQLQuery & " Where iihfChfCode = " & rst_Apf!chfCode
                'slSQLQuery = slSQLQuery & " And iihfVefCode = " & rst_Apf!vefCode
                'slSQLQuery = slSQLQuery & " And iihfInvStartDate = '" & Format(smMonthStart, "yyyy-mm-dd") & "'"
                'Set rst_Iihf = cnn.Execute(slSQLQuery)
                Do While Not rst_Iihf.EOF
                    If ilCount = 0 Then
                        slStr = Trim$(rst_Iihf!iihfStnInvoiceNo) & " " & Trim$(rst_Iihf!iihfSourceForm)
                        slStnInvoiceNo = rst_Iihf!iihfStnInvoiceNo
                    ElseIf ilCount = 1 Then
                        slStr = "+" & slStr & ";" & Trim$(rst_Iihf!iihfStnInvoiceNo) & " " & Trim$(rst_Iihf!iihfSourceForm)
                    Else
                        slStr = slStr & ";" & Trim$(rst_Iihf!iihfStnInvoiceNo) & " " & Trim$(rst_Iihf!iihfSourceForm)
                    End If
                    ilCount = ilCount + 1
                    rst_Iihf.MoveNext
                Loop
            Else
                slStr = "Wired"
            End If
            grdAcqCheckGrid.TextMatrix(llRow, STATIONINVNOINDEX) = slStr

            'If rst_Apf!apfOrderSpotCount <> rst_Apf!LineCount Then
            '    grdAcqCheckGrid.TextMatrix(llRow, ORDERCOUNTINDEX) = rst_Apf!apfOrderSpotCount & " / " & rst_Apf!LineCount
            'Else
            '    grdAcqCheckGrid.TextMatrix(llRow, ORDERCOUNTINDEX) = rst_Apf!apfOrderSpotCount
            'End If

            grdAcqCheckGrid.TextMatrix(llRow, ORDERCOUNTINDEX) = rst_Apf!LineCount
            If gDateValue(slDate) <> gDateValue("1/1/1970") Then
                grdAcqCheckGrid.TextMatrix(llRow, FULLYPAIDDATEINDEX) = slDate
            End If
            grdAcqCheckGrid.TextMatrix(llRow, ACQCOSTINDEX) = slAcqCost
            grdAcqCheckGrid.TextMatrix(llRow, ACQSPOTCOUNTINDEX) = rst_Apf!apfAiredSpotCount
            grdAcqCheckGrid.TextMatrix(llRow, INVCOUNTINDEX) = rst_Apf!InvSpotCount
            grdAcqCheckGrid.TextMatrix(llRow, ORIGACQSPOTCOUNTINDEX) = rst_Apf!apfAiredSpotCount
            grdAcqCheckGrid.TextMatrix(llRow, ORIGINVCOUNTINDEX) = rst_Apf!InvSpotCount
            grdAcqCheckGrid.TextMatrix(llRow, CHFCODEINDEX) = rst_Apf!chfCode
            grdAcqCheckGrid.TextMatrix(llRow, APFCODEINDEX) = rst_Apf!apfCode
            llRow = llRow + 1
        End If
        rst_Apf.MoveNext
    Loop
    On Error Resume Next
    rst_Apf.Close
    rst_Temp.Close
    slSQLQuery = "Drop Table #sdfCount"
    'Set rst_Temp = cnn.Execute(slSQLQuery)
    gSQLCallIgnoreError slSQLQuery
    If UBound(tmApfOrderAdj) > LBound(tmApfOrderAdj) Then
        If ilOrderCountUpdated > 0 Then
            gLogMsgWODT "O", hmAdjOrder, sgDBPath & "Messages\" & "AcqCheck_Adj_Order_" & Format(Now, "mmddyyyy") & ".csv"
            gLogMsgWODT "W", hmAdjOrder, "Acquisition Check: Adjusting Order Count for the Month of " & smMonthStart & "-" & smMonthEnd & " on " & Format(Now, "m/d/yy") & " " & Format(Now, "h:mm:ssAM/PM")
            gLogMsgWODT "W", hmAdjOrder, "Contract #,Advertiser Name,Vehicle Name,Acquisition Cost"
        End If
        For ilApf = 0 To UBound(tmApfOrderAdj) - 1 Step 1
            slSQLQuery = "Update Apf_Acq_Payable Set apfOrderSpotCount = " & tmApfOrderAdj(ilApf).iLineCount & ", "
            slSQLQuery = slSQLQuery & "apfStationInvNo = '" & tmApfOrderAdj(ilApf).sStationInvNo & "'"
            slSQLQuery = slSQLQuery & " Where apfCode = " & tmApfOrderAdj(ilApf).lApfCode
            'Set rst_Apf = cnn.Execute(slSQLQuery)
            If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                gHandleError "TrafficErrors.txt", "AcqCheck: mPopulate"
            End If
            slStr = ""
            ilAdf = gBinarySearchAdf(tmApfOrderAdj(ilApf).iAdfCode)
            If ilAdf <> -1 Then
                slStr = Trim$(tgCommAdf(ilAdf).sName)
            End If
            If ilOrderCountUpdated > 0 Then
                gLogMsgWODT "W", hmAdjOrder, tmApfOrderAdj(ilApf).lCntrNo & "," & """" & slStr & """" & "," & """" & Trim$(tmApfOrderAdj(ilApf).sVehicleName) & """" & "," & tmApfOrderAdj(ilApf).lAcqCost
            End If
        Next ilApf
        rst_Apf.Close
        If ilOrderCountUpdated > 0 Then
            gLogMsgWODT "C", hmAdjOrder, ""
            MsgBox ilOrderCountUpdated & " Ordered counts automatically adjusted, which will affect the aquisition reports", vbInformation + vbOKOnly, "Fixed Count"
        End If
    End If

    mGridSortCol STATIONINDEX
    mGridSortCol CNTRNOINDEX
    grdAcqCheckGrid.Redraw = True
    lacTotal.Caption = "Total: " & ilTotal
    lmNoRowsShown = ilTotal
    If ilTotal > 0 Then
        cmcCSV.Enabled = True
        If imPasswordOk Then
            ckcSet.Enabled = True
        Else
            ckcSet.Enabled = False
        End If
    Else
        cmcCSV.Enabled = False
        ckcSet.Enabled = False
    End If
    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbDefault
    On Error GoTo 0
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then
        ElseIf gErrSQL.Number <> 0 Then
        End If
    Next gErrSQL
End Sub

Private Sub mPopulateWOTempTable()
    Dim llCol As Long
    Dim slSQLQuery As String
    Dim llRow As Long
    Dim slDate As String
    Dim slAcqCost As String
    Dim ilAdf As Integer
    Dim blInclude As Boolean
    Dim llDate01011970 As Long
    Dim ilTotal As Integer
    Dim ilApf As Integer
    Dim slStr As String
    Dim ilCount As Integer
    Dim slStnInvoiceNo As String
    Dim ilVff As Integer
    Dim slPostLogSource As String
    Dim ilOrderCountUpdated As Integer
    Dim blRet As Boolean
    ReDim tmApfOrderAdj(0 To 0) As APFORDERADJ

    On Error GoTo ErrHand:

    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbHourglass

    grdAcqCheckGrid.Redraw = False
    
    blRet = mFindAndFixMissingApf()
    
    llDate01011970 = gDateValue("1/1/1970")
    ilTotal = 0
    lacTotal.Caption = ""
    lacAdjTotal.Caption = ""
    On Error Resume Next
    On Error GoTo ErrHand:
    
    mClearGrid
    grdAcqCheckGrid.Redraw = False
    'grdAcqCheckGrid.Row = 0
    'For llCol = CNTRNOINDEX To ADJDOLLARINDEX Step 1
    '    grdAcqCheckGrid.Col = llCol
    '    grdAcqCheckGrid.CellBackColor = vbBlack
    '    grdAcqCheckGrid.CellBackColor = LIGHTBLUE
    'Next llCol

    ilOrderCountUpdated = 0
    
    slSQLQuery = "Select chfCntrNo, chfCode, chfAdfCode, sdfVefCode, clfAcquisitionCost, Count(sdfCode) As SpotCount, "
    slSQLQuery = slSQLQuery & "Count(If((sdfSchStatus = 'S') Or (sdfSchStatus = 'G' And smfActualDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' And smfActualDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "'" & ") Or (sdfSchStatus = 'O' And smfActualDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' And smfActualDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "'" & "), 1, Null)) as SpotAiredCount"
    slSQLQuery = slSQLQuery & " From sdf_Spot_Detail"
    slSQLQuery = slSQLQuery & " Inner Join chf_Contract_Header On sdfChfCode = chfCode"
    slSQLQuery = slSQLQuery & " Inner Join clf_Contract_Line On sdfChfCode = clfChfCode and clfLine = sdfLineNo and clfCntRevNo = chfCntRevNo"
    slSQLQuery = slSQLQuery & " Inner Join vff_Vehicle_Features On sdfVefCode = vffVefCode"
    slSQLQuery = slSQLQuery & " Left Outer Join smf_Spot_MG_Specs On sdfSmfCode = smfCode"
    slSQLQuery = slSQLQuery & " where "
    slSQLQuery = slSQLQuery & "(sdfSpotType <> 'X')"
    slSQLQuery = slSQLQuery & " And (((sdfSchStatus = 'S') and (sdfDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' And sdfDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "')" & ")"
    slSQLQuery = slSQLQuery & " Or ((sdfSchStatus = 'M') and (sdfDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' And sdfDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "')" & ")"
    slSQLQuery = slSQLQuery & " Or ((sdfSchStatus = 'O' ) and (smfMissedDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' And smfMissedDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "')" & ")"
    slSQLQuery = slSQLQuery & " Or ((sdfSchStatus = 'G' ) and (smfMissedDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' And smfMissedDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "')" & ")" & ")"
    slSQLQuery = slSQLQuery & " group by chfCntrNo, chfCode, chfAdfCode, sdfVefCode, clfAcquisitionCost "   'Having SpotAiredCount > 0"

    Set rst_Temp = gSQLSelectCall(slSQLQuery)
    
    llRow = grdAcqCheckGrid.FixedRows
    
    Do While Not rst_Temp.EOF
        slSQLQuery = "Select apfCntrNo, apfStationInvNo, vefName, vefCode, apfAiredSpotCount, apfAcquisitionCost, apfOrderSpotCount, apfInvDate, apfFullyPaidDate, apfCode"
        slSQLQuery = slSQLQuery & " From apf_acq_payable"
        slSQLQuery = slSQLQuery & " Inner Join vef_Vehicles On apfVefCode = vefCode"
        slSQLQuery = slSQLQuery & " where "
        slSQLQuery = slSQLQuery & " apfCntrNo = " & rst_Temp!chfCntrNo & " And apfVefCode = " & rst_Temp!sdfVefCode & " And apfAcquisitionCost = " & rst_Temp!clfAcquisitionCost
        '4/20/20: added date range
        slSQLQuery = slSQLQuery & " And apfInvDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "'"
        slSQLQuery = slSQLQuery & " And apfInvDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "'"
        slSQLQuery = slSQLQuery & " And apfMnfItem = 0"
        slSQLQuery = slSQLQuery & " And ((apfAiredSpotCount <> " & rst_Temp!SpotAiredCount & ")"
        slSQLQuery = slSQLQuery & " Or (apfOrderSpotCount <> " & rst_Temp!SpotCount & "))"
        slSQLQuery = slSQLQuery & " Order By apfCntrNo, vefName, apfAcquisitionCost"
        Set rst_Apf = gSQLSelectCall(slSQLQuery)
        Do While Not rst_Apf.EOF
            blInclude = True
            slDate = Format(rst_Apf!apfFullyPaidDate, "m/d/yy")
            slAcqCost = gLongToStrDec(rst_Temp!clfAcquisitionCost, 2)
            If ckcInclude(0).Value = vbUnchecked Then 'Fully Paid
                If gDateValue(slDate) <> llDate01011970 Then
                    blInclude = False
                End If
            End If
            If ckcInclude(1).Value = vbUnchecked Then  'Not Fully Paid
                If gDateValue(slDate) = llDate01011970 Then
                    blInclude = False
                End If
            End If
            If ckcInclude(2).Value = vbUnchecked Then  'Non-zero dollars
                If rst_Temp!clfAcquisitionCost > 0 Then
                    blInclude = False
                End If
            End If
            If ckcInclude(3).Value = vbUnchecked Then  'Zero dollars
                If rst_Temp!clfAcquisitionCost = 0 Then
                    blInclude = False
                End If
            End If
            If ckcInclude(4).Value = vbUnchecked Then  'Spot count of zero
                If rst_Apf!apfAiredSpotCount = 0 Then
                    blInclude = False
                End If
            End If
            If ckcInclude(5).Value = vbUnchecked Then  'Spot count of non-zero
                If rst_Apf!apfAiredSpotCount <> 0 Then
                    blInclude = False
                End If
            End If
            If blInclude Then
                slStnInvoiceNo = ""
                ilVff = gBinarySearchVff(rst_Apf!vefCode)
                If ilVff <> -1 Then
                    slPostLogSource = tgVff(ilVff).sPostLogSource
                Else
                    slPostLogSource = "S"
                End If
                If slPostLogSource = "S" Then
                    slSQLQuery = "Select iihfSourceForm, iihfStnInvoiceNo"
                    slSQLQuery = slSQLQuery & " From iihf_ImptInvHeader"
                    slSQLQuery = slSQLQuery & " Where iihfChfCode = " & rst_Temp!chfCode
                    slSQLQuery = slSQLQuery & " And iihfVefCode = " & rst_Apf!vefCode
                    slSQLQuery = slSQLQuery & " And iihfInvStartDate = '" & Format(smMonthStart, "yyyy-mm-dd") & "'"
                    'Set rst_Iihf = cnn.Execute(slSQLQuery)
                    Set rst_Iihf = gSQLSelectCall(slSQLQuery)
                    If Not rst_Iihf.EOF Then
                        slStnInvoiceNo = rst_Iihf!iihfStnInvoiceNo
                    End If
                End If
                If (rst_Apf!apfOrderSpotCount <> rst_Temp!SpotCount) Or (rst_Apf!apfStationInvNo <> slStnInvoiceNo) Then
                    If rst_Apf!apfOrderSpotCount <> rst_Temp!SpotCount Then
                        ilOrderCountUpdated = ilOrderCountUpdated + 1
                    End If
                    tmApfOrderAdj(UBound(tmApfOrderAdj)).lApfCode = rst_Apf!apfCode
                    tmApfOrderAdj(UBound(tmApfOrderAdj)).iLineCount = rst_Temp!SpotCount
                    tmApfOrderAdj(UBound(tmApfOrderAdj)).sStationInvNo = slStnInvoiceNo
                    tmApfOrderAdj(UBound(tmApfOrderAdj)).lCntrNo = rst_Apf!apfCntrno
                    tmApfOrderAdj(UBound(tmApfOrderAdj)).iAdfCode = rst_Temp!chfAdfCode
                    tmApfOrderAdj(UBound(tmApfOrderAdj)).lAcqCost = rst_Apf!apfAcquisitionCost
                    tmApfOrderAdj(UBound(tmApfOrderAdj)).sVehicleName = rst_Apf!VEFNAME
                    ReDim Preserve tmApfOrderAdj(0 To UBound(tmApfOrderAdj) + 1) As APFORDERADJ
                End If
                If rst_Apf!apfAiredSpotCount = rst_Temp!SpotAiredCount Then
                    blInclude = False
                End If
            End If
            If blInclude Then
                ilTotal = ilTotal + 1
                If llRow >= grdAcqCheckGrid.Rows Then
                    grdAcqCheckGrid.AddItem ""
                End If
                grdAcqCheckGrid.RowHeight(llRow) = fgFlexGridRowH
                grdAcqCheckGrid.TextMatrix(llRow, SELECTEDINDEX) = "0"
                mPaintRowColor llRow
                grdAcqCheckGrid.TextMatrix(llRow, CNTRNOINDEX) = rst_Apf!apfCntrno
                ilAdf = gBinarySearchAdf(rst_Temp!chfAdfCode)
                If ilAdf <> -1 Then
                    grdAcqCheckGrid.TextMatrix(llRow, ADVTNAMEINDEX) = Trim$(tgCommAdf(ilAdf).sName)
                End If
                grdAcqCheckGrid.TextMatrix(llRow, STATIONINDEX) = Trim$(rst_Apf!VEFNAME)
                
                If slPostLogSource = "S" Then
                    slStr = ""
                    ilCount = 0
                    Do While Not rst_Iihf.EOF
                        If ilCount = 0 Then
                            slStr = Trim$(rst_Iihf!iihfStnInvoiceNo) & " " & Trim$(rst_Iihf!iihfSourceForm)
                            slStnInvoiceNo = rst_Iihf!iihfStnInvoiceNo
                        ElseIf ilCount = 1 Then
                            slStr = "+" & slStr & ";" & Trim$(rst_Iihf!iihfStnInvoiceNo) & " " & Trim$(rst_Iihf!iihfSourceForm)
                        Else
                            slStr = slStr & ";" & Trim$(rst_Iihf!iihfStnInvoiceNo) & " " & Trim$(rst_Iihf!iihfSourceForm)
                        End If
                        ilCount = ilCount + 1
                        rst_Iihf.MoveNext
                    Loop
                Else
                    slStr = "Wired"
                End If
                grdAcqCheckGrid.TextMatrix(llRow, STATIONINVNOINDEX) = slStr
                
                grdAcqCheckGrid.TextMatrix(llRow, ORDERCOUNTINDEX) = rst_Temp!SpotCount
                If gDateValue(slDate) <> gDateValue("1/1/1970") Then
                    grdAcqCheckGrid.TextMatrix(llRow, FULLYPAIDDATEINDEX) = slDate
                End If
                grdAcqCheckGrid.TextMatrix(llRow, ACQCOSTINDEX) = slAcqCost
                grdAcqCheckGrid.TextMatrix(llRow, ACQSPOTCOUNTINDEX) = rst_Apf!apfAiredSpotCount
                grdAcqCheckGrid.TextMatrix(llRow, INVCOUNTINDEX) = rst_Temp!SpotAiredCount
                grdAcqCheckGrid.TextMatrix(llRow, ORIGACQSPOTCOUNTINDEX) = rst_Apf!apfAiredSpotCount
                grdAcqCheckGrid.TextMatrix(llRow, ORIGINVCOUNTINDEX) = rst_Temp!SpotAiredCount
                grdAcqCheckGrid.TextMatrix(llRow, CHFCODEINDEX) = rst_Temp!chfCode
                grdAcqCheckGrid.TextMatrix(llRow, APFCODEINDEX) = rst_Apf!apfCode
                llRow = llRow + 1
            End If
            rst_Apf.MoveNext
        Loop
        rst_Temp.MoveNext
    Loop
    On Error Resume Next
    rst_Apf.Close
    rst_Temp.Close
    If UBound(tmApfOrderAdj) > LBound(tmApfOrderAdj) Then
        If ilOrderCountUpdated > 0 Then
            gLogMsgWODT "O", hmAdjOrder, sgDBPath & "Messages\" & "AcqCheck_Adj_Order_" & Format(Now, "mmddyyyy") & ".csv"
            gLogMsgWODT "W", hmAdjOrder, "Acquisition Check: Adjusting Order Count for the Month of " & smMonthStart & "-" & smMonthEnd & " on " & Format(Now, "m/d/yy") & " " & Format(Now, "h:mm:ssAM/PM")
            gLogMsgWODT "W", hmAdjOrder, "Contract #,Advertiser Name,Vehicle Name,Acquisition Cost"
        End If
        For ilApf = 0 To UBound(tmApfOrderAdj) - 1 Step 1
            slSQLQuery = "Update Apf_Acq_Payable Set apfOrderSpotCount = " & tmApfOrderAdj(ilApf).iLineCount & ", "
            slSQLQuery = slSQLQuery & "apfStationInvNo = '" & tmApfOrderAdj(ilApf).sStationInvNo & "'"
            slSQLQuery = slSQLQuery & " Where apfCode = " & tmApfOrderAdj(ilApf).lApfCode
            'Set rst_Apf = cnn.Execute(slSQLQuery)
            If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                gHandleError "TrafficErrors.txt", "AcqCheck: mPopulate"
            End If
            slStr = ""
            ilAdf = gBinarySearchAdf(tmApfOrderAdj(ilApf).iAdfCode)
            If ilAdf <> -1 Then
                slStr = Trim$(tgCommAdf(ilAdf).sName)
            End If
            If ilOrderCountUpdated > 0 Then
                gLogMsgWODT "W", hmAdjOrder, tmApfOrderAdj(ilApf).lCntrNo & "," & """" & slStr & """" & "," & """" & Trim$(tmApfOrderAdj(ilApf).sVehicleName) & """" & "," & tmApfOrderAdj(ilApf).lAcqCost
            End If
        Next ilApf
        rst_Apf.Close
        If ilOrderCountUpdated > 0 Then
            gLogMsgWODT "C", hmAdjOrder, ""
            MsgBox ilOrderCountUpdated & " Ordered counts automatically adjusted, which will affect the aquisition reports", vbInformation + vbOKOnly, "Fixed Count"
        End If
    End If
    
    mGridSortCol STATIONINDEX
    mGridSortCol CNTRNOINDEX
    grdAcqCheckGrid.Redraw = True
    lacTotal.Caption = "Total: " & ilTotal
    lmNoRowsShown = ilTotal
    If ilTotal > 0 Then
        cmcCSV.Enabled = True
        If imPasswordOk Then
            ckcSet.Enabled = True
        Else
            ckcSet.Enabled = False
        End If
    Else
        cmcCSV.Enabled = False
        ckcSet.Enabled = False
    End If
    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbDefault
    On Error GoTo 0
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then
        ElseIf gErrSQL.Number <> 0 Then
        End If
    Next gErrSQL
End Sub
Private Sub mSetGridColumns()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer
    
    'Copy of grdNoMatchResult
    grdAcqCheckGrid.ColWidth(ORIGACQSPOTCOUNTINDEX) = 0
    grdAcqCheckGrid.ColWidth(ORIGINVCOUNTINDEX) = 0
    grdAcqCheckGrid.ColWidth(SELECTEDINDEX) = 0
    grdAcqCheckGrid.ColWidth(SORTINDEX) = 0
    grdAcqCheckGrid.ColWidth(CHFCODEINDEX) = 0
    grdAcqCheckGrid.ColWidth(APFCODEINDEX) = 0
    grdAcqCheckGrid.ColWidth(CNTRNOINDEX) = 0.07 * grdAcqCheckGrid.Width
    grdAcqCheckGrid.ColWidth(ADVTNAMEINDEX) = 0.2 * grdAcqCheckGrid.Width
    grdAcqCheckGrid.ColWidth(STATIONINVNOINDEX) = 0.1 * grdAcqCheckGrid.Width
    grdAcqCheckGrid.ColWidth(STATIONINDEX) = 0.07 * grdAcqCheckGrid.Width
    grdAcqCheckGrid.ColWidth(ORDERCOUNTINDEX) = 0.07 * grdAcqCheckGrid.Width
    grdAcqCheckGrid.ColWidth(FULLYPAIDDATEINDEX) = 0.07 * grdAcqCheckGrid.Width
    grdAcqCheckGrid.ColWidth(ACQCOSTINDEX) = 0.07 * grdAcqCheckGrid.Width
    grdAcqCheckGrid.ColWidth(ACQSPOTCOUNTINDEX) = 0.07 * grdAcqCheckGrid.Width
    grdAcqCheckGrid.ColWidth(INVCOUNTINDEX) = 0.07 * grdAcqCheckGrid.Width
    grdAcqCheckGrid.ColWidth(ADJDOLLARINDEX) = 0.07 * grdAcqCheckGrid.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = 0  'grdAcqCheckGrid.Width
    For ilCol = 0 To grdAcqCheckGrid.Cols - 1 Step 1
        llWidth = llWidth + grdAcqCheckGrid.ColWidth(ilCol)
        If (grdAcqCheckGrid.ColWidth(ilCol) > 15) And (grdAcqCheckGrid.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdAcqCheckGrid.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdAcqCheckGrid.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdAcqCheckGrid.Width
            For ilCol = 0 To grdAcqCheckGrid.Cols - 1 Step 1
                If (grdAcqCheckGrid.ColWidth(ilCol) > 15) And (grdAcqCheckGrid.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdAcqCheckGrid.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdAcqCheckGrid.FixedCols To grdAcqCheckGrid.Cols - 1 Step 1
                If grdAcqCheckGrid.ColWidth(ilCol) > 15 Then
                    ilColInc = grdAcqCheckGrid.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdAcqCheckGrid.ColWidth(ilCol) = grdAcqCheckGrid.ColWidth(ilCol) + 15
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

Private Sub mSetGridTitles()
    Dim llRow As Long
    For llRow = 0 To 1 Step 1
        grdAcqCheckGrid.Row = llRow
        grdAcqCheckGrid.Col = CNTRNOINDEX
        grdAcqCheckGrid.CellFontBold = False
        grdAcqCheckGrid.CellFontName = "Arial"
        grdAcqCheckGrid.CellFontSize = 6.75
        grdAcqCheckGrid.CellForeColor = vbBlue
        grdAcqCheckGrid.CellBackColor = LIGHTBLUE
        If llRow = 0 Then
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Contract"
        Else
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Number"
        End If
        grdAcqCheckGrid.Col = ADVTNAMEINDEX
        grdAcqCheckGrid.CellFontBold = False
        grdAcqCheckGrid.CellFontName = "Arial"
        grdAcqCheckGrid.CellFontSize = 6.75
        grdAcqCheckGrid.CellForeColor = vbBlue
        grdAcqCheckGrid.CellBackColor = LIGHTBLUE
        If llRow = 0 Then
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Advertiser"
        Else
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Name"
        End If
        grdAcqCheckGrid.Col = STATIONINDEX
        grdAcqCheckGrid.CellFontBold = False
        grdAcqCheckGrid.CellFontName = "Arial"
        grdAcqCheckGrid.CellFontSize = 6.75
        grdAcqCheckGrid.CellForeColor = vbBlue
        grdAcqCheckGrid.CellBackColor = LIGHTBLUE
        If llRow = 0 Then
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Call"
        Else
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Letters"
        End If
        grdAcqCheckGrid.Col = STATIONINVNOINDEX
        grdAcqCheckGrid.CellFontBold = False
        grdAcqCheckGrid.CellFontName = "Arial"
        grdAcqCheckGrid.CellFontSize = 6.75
        grdAcqCheckGrid.CellForeColor = vbBlue
        grdAcqCheckGrid.CellBackColor = LIGHTBLUE
        If llRow = 0 Then
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Station"
        Else
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Invoice #"
        End If
        grdAcqCheckGrid.Col = ORDERCOUNTINDEX
        grdAcqCheckGrid.CellFontBold = False
        grdAcqCheckGrid.CellFontName = "Arial"
        grdAcqCheckGrid.CellFontSize = 6.75
        grdAcqCheckGrid.CellForeColor = vbBlue
        'grdAcqCheckGrid.CellBackColor = LIGHTBLUE
        If llRow = 0 Then
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Ordered"
        Else
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Count"
        End If
        grdAcqCheckGrid.Col = FULLYPAIDDATEINDEX
        grdAcqCheckGrid.CellFontBold = False
        grdAcqCheckGrid.CellFontName = "Arial"
        grdAcqCheckGrid.CellFontSize = 6.75
        grdAcqCheckGrid.CellForeColor = vbBlue
        'grdAcqCheckGrid.CellBackColor = LIGHTBLUE
        If llRow = 0 Then
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Agency Fully"
        Else
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Paid Date"
        End If
        grdAcqCheckGrid.Col = ACQCOSTINDEX
        grdAcqCheckGrid.CellFontBold = False
        grdAcqCheckGrid.CellFontName = "Arial"
        grdAcqCheckGrid.CellFontSize = 6.75
        grdAcqCheckGrid.CellForeColor = vbBlue
        grdAcqCheckGrid.CellBackColor = LIGHTBLUE
        If llRow = 0 Then
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Acquisition"
        Else
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Cost"
        End If
        grdAcqCheckGrid.Col = ACQSPOTCOUNTINDEX
        grdAcqCheckGrid.CellFontBold = False
        grdAcqCheckGrid.CellFontName = "Arial"
        grdAcqCheckGrid.CellFontSize = 6.75
        grdAcqCheckGrid.CellForeColor = vbBlue
        'grdAcqCheckGrid.CellBackColor = LIGHTBLUE
        If llRow = 0 Then
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Acquisition"
        Else
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Spot Count"
        End If
        grdAcqCheckGrid.Col = INVCOUNTINDEX
        grdAcqCheckGrid.CellFontBold = False
        grdAcqCheckGrid.CellFontName = "Arial"
        grdAcqCheckGrid.CellFontSize = 6.75
        grdAcqCheckGrid.CellForeColor = vbBlue
        'grdAcqCheckGrid.CellBackColor = LIGHTBLUE
        If llRow = 0 Then
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Invoiced"
        Else
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Spot Count"
        End If
        grdAcqCheckGrid.Col = ADJDOLLARINDEX
        grdAcqCheckGrid.CellFontBold = False
        grdAcqCheckGrid.CellFontName = "Arial"
        grdAcqCheckGrid.CellFontSize = 6.75
        grdAcqCheckGrid.CellForeColor = vbBlue
        'grdAcqCheckGrid.CellBackColor = LIGHTBLUE
        If llRow = 0 Then
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Adjusted"
        Else
            grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, grdAcqCheckGrid.Col) = "Dollars"
        End If
    Next llRow
End Sub

Private Sub mGridSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim llEnableRow As Long
    Dim llEnableCol As Long

    llEnableRow = grdAcqCheckGrid.Row
    llEnableCol = grdAcqCheckGrid.Col
    grdAcqCheckGrid.Row = 0
    grdAcqCheckGrid.Col = ilCol
    If grdAcqCheckGrid.CellBackColor <> LIGHTBLUE Then
        grdAcqCheckGrid.Row = llEnableRow
        grdAcqCheckGrid.Col = llEnableCol
        Exit Sub
    End If
    grdAcqCheckGrid.Row = llEnableRow
    grdAcqCheckGrid.Col = llEnableCol
    For llRow = grdAcqCheckGrid.FixedRows To grdAcqCheckGrid.Rows - 1 Step 1
        slStr = Trim$(grdAcqCheckGrid.TextMatrix(llRow, CNTRNOINDEX))
        If slStr <> "" Then
            If ilCol = CNTRNOINDEX Then
                slSort = grdAcqCheckGrid.TextMatrix(llRow, CNTRNOINDEX)
                Do While Len(slSort) < 10
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = STATIONINVNOINDEX Then
                slSort = grdAcqCheckGrid.TextMatrix(llRow, STATIONINVNOINDEX)
                Do While Len(slSort) < 10
                    slSort = " " & slSort
                Loop
            ElseIf ilCol = ACQCOSTINDEX Then
                slSort = grdAcqCheckGrid.TextMatrix(llRow, ACQCOSTINDEX)
                Do While Len(slSort) < 10
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdAcqCheckGrid.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdAcqCheckGrid.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imColSorted) Or ((ilCol = imColSorted) And (imSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdAcqCheckGrid.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdAcqCheckGrid.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imColSorted Then
        imColSorted = SORTINDEX
    Else
        imColSorted = -1
        imSort = -1
    End If
    gGrid_SortByCol grdAcqCheckGrid, CNTRNOINDEX, SORTINDEX, imColSorted, imSort
    imColSorted = ilCol
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

    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbDefault
    Unload AcqCheck
End Sub



Private Sub mSetControls()
    Dim ilGap As Integer

    ilGap = cmcCancel.Left - (cmcCheck.Left + cmcCheck.Width)
    cmcCheck.Top = Me.Height - cmcCheck.Height - 120
    cmcCancel.Top = cmcCheck.Top
    cmcSave.Top = cmcCheck.Top
    cmcCSV.Top = cmcCheck.Top
    cmcUndo.Top = cmcCheck.Top
    grdAcqCheckGrid.Move 120, frcInclude.Top + frcInclude.Height + 60, AcqCheck.Width - 360, cmcCheck.Top - frcInclude.Top - frcInclude.Height - 255 - 120
    lacTotal.Move grdAcqCheckGrid.Left + grdAcqCheckGrid.Width - lacTotal.Width, grdAcqCheckGrid.Top - lacTotal.Height
    ckcSet.Move grdAcqCheckGrid.Left, grdAcqCheckGrid.Top + grdAcqCheckGrid.Height + 60
    lacAdjTotal.Move grdAcqCheckGrid.Left + grdAcqCheckGrid.Width - lacAdjTotal.Width, grdAcqCheckGrid.Top + grdAcqCheckGrid.Height + 60
    cmcCheck.Left = grdAcqCheckGrid.Left + grdAcqCheckGrid.Width / 2 - (5 * cmcCheck.Width) / 2 - (2 * ilGap)
    cmcCancel.Left = cmcCheck.Left + cmcCheck.Width + ilGap
    cmcSave.Left = cmcCancel.Left + cmcCancel.Width + ilGap
    cmcCSV.Left = cmcSave.Left + cmcSave.Width + ilGap
    cmcUndo.Left = cmcCSV.Left + cmcCSV.Width + ilGap
    mSetGridColumns
    mSetGridTitles
    gGrid_IntegralHeight grdAcqCheckGrid, fgBoxGridH + 15

End Sub


Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long
    Dim llEnableRow As Long
    Dim llEnableCol As Long
    
    llEnableRow = grdAcqCheckGrid.Row
    llEnableCol = grdAcqCheckGrid.Col
    grdAcqCheckGrid.Row = llRow
    For llCol = CNTRNOINDEX To ADJDOLLARINDEX Step 1
        grdAcqCheckGrid.Col = llCol
        If grdAcqCheckGrid.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
            'If (llCol <= ACQCOSTINDEX) Or (llCol = ADJDOLLARINDEX) Then
            If (llCol <> ACQSPOTCOUNTINDEX) Then
                grdAcqCheckGrid.CellBackColor = LIGHTYELLOW
            Else
                grdAcqCheckGrid.CellBackColor = vbWhite
            End If
        Else
            grdAcqCheckGrid.CellBackColor = GRAY    'vbBlue
        End If
    Next llCol
    grdAcqCheckGrid.Row = llEnableRow
    grdAcqCheckGrid.Col = llEnableCol
End Sub

Private Sub grdAcqCheckGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim llHeight As Long
    Dim ilSoldAdj As Integer

    llHeight = 0
    For llRow = 0 To grdAcqCheckGrid.FixedRows - 1 Step 1
        llHeight = llHeight + grdAcqCheckGrid.RowHeight(llRow)
    Next llRow
    'grdAcqCheckGrid.ToolTipText = ""
    If Y <= llHeight Then
        grdAcqCheckGrid.ToolTipText = ""
        Exit Sub
    End If
    'If y > lmMaxHeight Then
    '    Exit Sub
    'End If
    If grdAcqCheckGrid.MouseRow >= grdAcqCheckGrid.Rows Then
        grdAcqCheckGrid.ToolTipText = ""
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdAcqCheckGrid, X, Y, llRow, llCol)
    If Not ilFound Then
        grdAcqCheckGrid.ToolTipText = ""
        Exit Sub
    End If
    grdAcqCheckGrid.ToolTipText = Trim$(grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.MouseRow, grdAcqCheckGrid.MouseCol))

End Sub

Private Sub grdAcqCheckGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    imIgnoreScroll = False
    mSetShow
    If Y < 2 * grdAcqCheckGrid.RowHeight(0) Then
        grdAcqCheckGrid.Col = grdAcqCheckGrid.MouseCol
        mGridSortCol grdAcqCheckGrid.Col
        grdAcqCheckGrid.Row = 0
        grdAcqCheckGrid.Col = APFCODEINDEX
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdAcqCheckGrid, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdAcqCheckGrid.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdAcqCheckGrid.FixedRows Then
        If grdAcqCheckGrid.TextMatrix(llCurrentRow, CNTRNOINDEX) <> "" Then
            llTopRow = grdAcqCheckGrid.TopRow
            'For llRow = grdAcqCheckGrid.FixedRows To grdAcqCheckGrid.Rows - 1 Step 1
            '    If grdAcqCheckGrid.TextMatrix(llRow, CNTRNOINDEX) <> "" Then
            '        If llRow = llCurrentRow Then
            '            grdAcqCheckGrid.TextMatrix(llRow, SELECTEDINDEX) = "1"
            '        Else
            '            grdAcqCheckGrid.TextMatrix(llRow, SELECTEDINDEX) = "0"
            '        End If
            '        mPaintRowColor llRow
            '    End If
            'Next llRow
            grdAcqCheckGrid.TopRow = llTopRow
            grdAcqCheckGrid.Row = llCurrentRow
            grdAcqCheckGrid.Row = llCurrentRow
            grdAcqCheckGrid.Col = llCol
            If Not mColOk() Then
                pbcClickFocus.SetFocus
            Else
                mEnableBox
            End If
            Exit Sub
        End If
    End If

End Sub

Private Sub mClearGrid()
    Dim llRow As Long
    Dim ilCol As Integer
    
    On Error Resume Next
    cmcCSV.Enabled = False
    ckcSet.Enabled = False
    ckcSet.Value = vbUnchecked
    cmcUndo.Enabled = False
    smTotalAdjDollar = ""
    grdAcqCheckGrid.Redraw = False
    grdAcqCheckGrid.RowHeight(0) = fgFlexGridRowH
    grdAcqCheckGrid.RowHeight(1) = fgFlexGridRowH
    grdAcqCheckGrid.Rows = grdAcqCheckGrid.FixedRows + 1
    For llRow = grdAcqCheckGrid.FixedRows To grdAcqCheckGrid.Rows - 1 Step 1
        grdAcqCheckGrid.RowHeight(llRow) = fgFlexGridRowH
        grdAcqCheckGrid.TextMatrix(llRow, SELECTEDINDEX) = "0"
        mPaintRowColor llRow
        For ilCol = CNTRNOINDEX To APFCODEINDEX Step 1
            grdAcqCheckGrid.TextMatrix(llRow, ilCol) = ""
        Next ilCol
    Next llRow
    gGrid_AlignAllColsLeft grdAcqCheckGrid
    grdAcqCheckGrid.Redraw = True
End Sub
Private Function mColOk() As Integer
    mColOk = True
    If grdAcqCheckGrid.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
End Function

Private Sub mEnableBox()
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'

    If Not imPasswordOk Then
        Exit Sub
    End If
    If (grdAcqCheckGrid.Row < grdAcqCheckGrid.FixedRows) Or (grdAcqCheckGrid.Row >= grdAcqCheckGrid.Rows) Or (grdAcqCheckGrid.Col < grdAcqCheckGrid.FixedCols) Or (grdAcqCheckGrid.Col >= grdAcqCheckGrid.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdAcqCheckGrid.Row
    lmEnableCol = grdAcqCheckGrid.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdAcqCheckGrid.Left - pbcArrow.Width - 30, grdAcqCheckGrid.Top + grdAcqCheckGrid.RowPos(grdAcqCheckGrid.Row) + (grdAcqCheckGrid.RowHeight(grdAcqCheckGrid.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    grdAcqCheckGrid.TextMatrix(lmEnableRow, SELECTEDINDEX) = "1"
    mPaintRowColor lmEnableRow
    Select Case grdAcqCheckGrid.Col

        Case ACQSPOTCOUNTINDEX
            edcDropDown.MaxLength = 5
            edcDropDown.Text = grdAcqCheckGrid.Text
        Case INVCOUNTINDEX
            edcDropDown.MaxLength = 5
            edcDropDown.Text = grdAcqCheckGrid.Text
    End Select
    mSetFocus
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
    Dim ilUnits As Integer
    Dim llEnableRow As Long
    Dim llEnableCol As Long
    Dim llRow As Long
    Dim blFound As Boolean
    Dim slOldAdjDollar As String
    Dim slNewAdjDollar As String

    pbcArrow.Visible = False
    llEnableRow = grdAcqCheckGrid.Row
    llEnableCol = grdAcqCheckGrid.Col
    If (lmEnableRow >= grdAcqCheckGrid.FixedRows) And (lmEnableRow < grdAcqCheckGrid.Rows) Then
        If (grdAcqCheckGrid.TextMatrix(lmEnableRow, lmEnableCol) <> edcDropDown.Text) Then
            imChg = True
            bmCSVPressed = False
            If Not bmInSetAll Then
                ckcSet.Value = vbUnchecked
            End If
        End If
        Select Case lmEnableCol
            Case ACQSPOTCOUNTINDEX
                edcDropDown.Visible = False
                grdAcqCheckGrid.TextMatrix(lmEnableRow, lmEnableCol) = edcDropDown.Text
                If grdAcqCheckGrid.TextMatrix(lmEnableRow, FULLYPAIDDATEINDEX) <> "" Then
                    slOldAdjDollar = grdAcqCheckGrid.TextMatrix(lmEnableRow, ADJDOLLARINDEX)
                    If grdAcqCheckGrid.TextMatrix(lmEnableRow, lmEnableCol) <> grdAcqCheckGrid.TextMatrix(lmEnableRow, ORIGACQSPOTCOUNTINDEX) Then
                        grdAcqCheckGrid.TextMatrix(lmEnableRow, ADJDOLLARINDEX) = gMulStr(grdAcqCheckGrid.TextMatrix(lmEnableRow, ACQCOSTINDEX), gSubStr(grdAcqCheckGrid.TextMatrix(lmEnableRow, lmEnableCol), grdAcqCheckGrid.TextMatrix(lmEnableRow, ORIGACQSPOTCOUNTINDEX)))
                    Else
                        grdAcqCheckGrid.TextMatrix(lmEnableRow, ADJDOLLARINDEX) = ""
                    End If
                    If slOldAdjDollar = "" Then
                        slOldAdjDollar = ".00"
                    End If
                    slNewAdjDollar = grdAcqCheckGrid.TextMatrix(lmEnableRow, ADJDOLLARINDEX)
                    If slNewAdjDollar = "" Then
                        slNewAdjDollar = ".00"
                    End If
                    If smTotalAdjDollar = "" Then
                        smTotalAdjDollar = ".00"
                    End If
                    smTotalAdjDollar = gAddStr(smTotalAdjDollar, gSubStr(slNewAdjDollar, slOldAdjDollar))
                    lacAdjTotal.Caption = "Agency Fully Paid Adjustment Total: " & smTotalAdjDollar
                End If
            Case INVCOUNTINDEX
                edcDropDown.Visible = False
                grdAcqCheckGrid.TextMatrix(lmEnableRow, lmEnableCol) = edcDropDown.Text
        End Select
        grdAcqCheckGrid.TextMatrix(lmEnableRow, SELECTEDINDEX) = "0"
        mPaintRowColor lmEnableRow
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    mSetCommands
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
    Dim llColWidth As Long

    If (grdAcqCheckGrid.Row < grdAcqCheckGrid.FixedRows) Or (grdAcqCheckGrid.Row >= grdAcqCheckGrid.Rows) Or (grdAcqCheckGrid.Col < grdAcqCheckGrid.FixedCols) Or (grdAcqCheckGrid.Col >= grdAcqCheckGrid.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdAcqCheckGrid.Col - 1 Step 1
        llColPos = llColPos + grdAcqCheckGrid.ColWidth(ilCol)
    Next ilCol
    llColWidth = grdAcqCheckGrid.ColWidth(grdAcqCheckGrid.Col)
    ilCol = grdAcqCheckGrid.Col
    Do While ilCol < grdAcqCheckGrid.Cols - 1
        If (Trim$(grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row - 1, grdAcqCheckGrid.Col)) <> "") And (Trim$(grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row - 1, grdAcqCheckGrid.Col)) = Trim$(grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row - 1, ilCol + 1))) Then
            llColWidth = llColWidth + grdAcqCheckGrid.ColWidth(ilCol + 1)
            ilCol = ilCol + 1
        Else
            Exit Do
        End If
    Loop
    Select Case grdAcqCheckGrid.Col
        Case ACQSPOTCOUNTINDEX
            edcDropDown.Move grdAcqCheckGrid.Left + llColPos + 30, grdAcqCheckGrid.Top + grdAcqCheckGrid.RowPos(grdAcqCheckGrid.Row) + 15, grdAcqCheckGrid.ColWidth(grdAcqCheckGrid.Col) - 30, grdAcqCheckGrid.RowHeight(grdAcqCheckGrid.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case INVCOUNTINDEX
            edcDropDown.Move grdAcqCheckGrid.Left + llColPos + 30, grdAcqCheckGrid.Top + grdAcqCheckGrid.RowPos(grdAcqCheckGrid.Row) + 15, grdAcqCheckGrid.ColWidth(grdAcqCheckGrid.Col) - 30, grdAcqCheckGrid.RowHeight(grdAcqCheckGrid.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
    End Select
End Sub

Private Sub grdAcqCheckGrid_Scroll()
    'other scroll logic: see copy->grdCopy_Scroll
    If imIgnoreScroll Then  'Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdAcqCheckGrid.Redraw = False Then
        grdAcqCheckGrid.Redraw = True
        If lmTopRow < grdAcqCheckGrid.FixedRows Then
            grdAcqCheckGrid.TopRow = grdAcqCheckGrid.FixedRows
        Else
            grdAcqCheckGrid.TopRow = lmTopRow
        End If
        grdAcqCheckGrid.Refresh
        grdAcqCheckGrid.Redraw = False
    End If
    If (imCtrlVisible) And (grdAcqCheckGrid.Row >= grdAcqCheckGrid.FixedRows) And (grdAcqCheckGrid.Col >= 0) And (grdAcqCheckGrid.Col < grdAcqCheckGrid.Cols - 1) Then
        If grdAcqCheckGrid.RowIsVisible(grdAcqCheckGrid.Row) Then
            pbcArrow.Move grdAcqCheckGrid.Left - pbcArrow.Width - 30, grdAcqCheckGrid.Top + grdAcqCheckGrid.RowPos(grdAcqCheckGrid.Row) + (grdAcqCheckGrid.RowHeight(grdAcqCheckGrid.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            mSetFocus
        Else
            pbcClickFocus.SetFocus
            edcDropDown.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
    End If
End Sub

Private Sub mSetCommands()
'
'   mSetCommands
'   Where:
'
    cmcCheck.Enabled = True
    If imChg And imPasswordOk Then
        cmcCheck.Enabled = False
        frcInclude.Enabled = False
        cbcMonth.Enabled = False
        edcYear.Enabled = False
        cmcSave.Enabled = True
        cmcUndo.Enabled = True
    Else
        cmcSave.Enabled = False
        cmcUndo.Enabled = False
        If (cbcMonth.ListIndex >= 0) And ((Len(edcYear.Text) = 2) Or (Len(edcYear.Text) = 4)) Then
            cmcCheck.Enabled = True
        Else
            cmcCheck.Enabled = False
        End If
        frcInclude.Enabled = True
        cbcMonth.Enabled = True
        edcYear.Enabled = True
    End If
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilPrev As Integer

    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mEnableBox
        Exit Sub
    End If
    If imCtrlVisible Then
        mSetShow
        Do
            ilPrev = False
            If grdAcqCheckGrid.Col = ACQSPOTCOUNTINDEX Then
                If grdAcqCheckGrid.Row > grdAcqCheckGrid.FixedRows Then
                    lmTopRow = -1
                    grdAcqCheckGrid.Row = grdAcqCheckGrid.Row - 1
                    If Not grdAcqCheckGrid.RowIsVisible(grdAcqCheckGrid.Row) Then
                        grdAcqCheckGrid.TopRow = grdAcqCheckGrid.TopRow - 1
                    End If
                    'grdAcqCheckGrid.Col = INVCOUNTINDEX
                    grdAcqCheckGrid.Col = ACQSPOTCOUNTINDEX
                    mEnableBox
                Else
                    cmcCancel.SetFocus
                End If
            Else
                grdAcqCheckGrid.Col = grdAcqCheckGrid.Col - 1
                If Not mColOk() Then
                    ilPrev = True
                Else
                    mEnableBox
                End If
            End If
        Loop While ilPrev
    Else
        If grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.FixedRows, CNTRNOINDEX) = "" Then
            cmcCancel.SetFocus
            Exit Sub
        End If
        lmTopRow = -1
        grdAcqCheckGrid.TopRow = grdAcqCheckGrid.FixedRows
        'grdAcqCheckGrid.Col = INVCOUNTINDEX
        grdAcqCheckGrid.Col = ACQSPOTCOUNTINDEX
        grdAcqCheckGrid.Row = grdAcqCheckGrid.FixedRows
        mEnableBox
    End If
End Sub

Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    Dim ilNext As Integer
    Dim llEnableRow As Long

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        llEnableRow = lmEnableRow
        mSetShow
        Do
            ilNext = False
            'If grdAcqCheckGrid.Col = INVCOUNTINDEX Then
            If grdAcqCheckGrid.Col = ACQSPOTCOUNTINDEX Then
                llRow = grdAcqCheckGrid.Rows
                Do
                    llRow = llRow - 1
                Loop While grdAcqCheckGrid.TextMatrix(llRow, CNTRNOINDEX) = ""
                llRow = llRow + 1
                If (grdAcqCheckGrid.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdAcqCheckGrid.Row = grdAcqCheckGrid.Row + 1
                    If Not grdAcqCheckGrid.RowIsVisible(grdAcqCheckGrid.Row) Then
                        imIgnoreScroll = True
                        grdAcqCheckGrid.TopRow = grdAcqCheckGrid.TopRow + 1
                    End If
                    grdAcqCheckGrid.Col = ACQSPOTCOUNTINDEX
                    If Trim$(grdAcqCheckGrid.TextMatrix(grdAcqCheckGrid.Row, CNTRNOINDEX)) <> "" Then
                        mEnableBox
                    Else
                        imFromArrow = True
                        pbcArrow.Move grdAcqCheckGrid.Left - pbcArrow.Width - 30, grdAcqCheckGrid.Top + grdAcqCheckGrid.RowPos(grdAcqCheckGrid.Row) + (grdAcqCheckGrid.RowHeight(grdAcqCheckGrid.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    End If
                Else
                    pbcClickFocus.SetFocus
                End If
            Else
                grdAcqCheckGrid.Col = grdAcqCheckGrid.Col + 1
                mEnableBox
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdAcqCheckGrid.TopRow = grdAcqCheckGrid.FixedRows
        grdAcqCheckGrid.Col = ACQSPOTCOUNTINDEX
        grdAcqCheckGrid.Row = grdAcqCheckGrid.FixedRows
        mEnableBox
    End If
End Sub

Private Function mSaveRec() As Integer
    Dim llRow As Long
    Dim ilCol As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim llChfCode As Long
    Dim ilVefCode As Integer
    Dim llSdfDate As Long
    Dim llMissedDate As Long
    Dim blIncludeSpot As Boolean

    mSaveRec = True
    For llRow = grdAcqCheckGrid.FixedRows To grdAcqCheckGrid.Rows - 1 Step 1
        slStr = Trim$(grdAcqCheckGrid.TextMatrix(llRow, CNTRNOINDEX))
        If slStr <> "" Then
            tmApfSrchKey0.lCode = Val(grdAcqCheckGrid.TextMatrix(llRow, APFCODEINDEX))
            ilRet = btrGetEqual(hmApf, tmApf, imApfRecLen, tmApfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                If (grdAcqCheckGrid.TextMatrix(llRow, ACQSPOTCOUNTINDEX) <> grdAcqCheckGrid.TextMatrix(llRow, ORIGACQSPOTCOUNTINDEX)) Then
                    llChfCode = Val(grdAcqCheckGrid.TextMatrix(llRow, CHFCODEINDEX))
                    ilVefCode = tmApf.iVefCode
                    'If grdAcqCheckGrid.TextMatrix(llRow, INVCOUNTINDEX) < grdAcqCheckGrid.TextMatrix(llRow, ORIGINVCOUNTINDEX) Then
                    If grdAcqCheckGrid.TextMatrix(llRow, ACQSPOTCOUNTINDEX) < grdAcqCheckGrid.TextMatrix(llRow, ORIGINVCOUNTINDEX) Then
                        'Change schedule to missed spot
                        ilRet = mFindAndFixSpot(llChfCode, ilVefCode, "S", grdAcqCheckGrid.TextMatrix(llRow, ACQCOSTINDEX), Val(grdAcqCheckGrid.TextMatrix(llRow, ORIGINVCOUNTINDEX)) - Val(grdAcqCheckGrid.TextMatrix(llRow, ACQSPOTCOUNTINDEX)))
                    'ElseIf grdAcqCheckGrid.TextMatrix(llRow, INVCOUNTINDEX) > grdAcqCheckGrid.TextMatrix(llRow, ORIGINVCOUNTINDEX) Then
                    ElseIf grdAcqCheckGrid.TextMatrix(llRow, ACQSPOTCOUNTINDEX) > grdAcqCheckGrid.TextMatrix(llRow, ORIGINVCOUNTINDEX) Then
                        'Schedule missed spot
                        ilRet = mFindAndFixSpot(llChfCode, ilVefCode, "M", grdAcqCheckGrid.TextMatrix(llRow, ACQCOSTINDEX), Val(grdAcqCheckGrid.TextMatrix(llRow, ACQSPOTCOUNTINDEX)) - Val(grdAcqCheckGrid.TextMatrix(llRow, ORIGINVCOUNTINDEX)))
                    End If
                    'If (grdAcqCheckGrid.TextMatrix(llRow, INVCOUNTINDEX) <> grdAcqCheckGrid.TextMatrix(llRow, ORIGINVCOUNTINDEX)) Or (grdAcqCheckGrid.TextMatrix(llRow, ACQSPOTCOUNTINDEX) <> grdAcqCheckGrid.TextMatrix(llRow, ORIGACQSPOTCOUNTINDEX)) Then
                    'Find lines
                    tmApf.iAiredSpotCount = 0
                    tmClfSrchKey1.lChfCode = llChfCode
                    tmClfSrchKey1.iVefCode = ilVefCode
                    ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode) And (tmClf.iVefCode = ilVefCode)
                        If tmApf.lAcquisitionCost = tmClf.lAcquisitionCost Then
                            'Update aired count
                            tmSdfSrchKey0.iVefCode = ilVefCode
                            tmSdfSrchKey0.lChfCode = llChfCode
                            tmSdfSrchKey0.iLineNo = tmClf.iLine
                            tmSdfSrchKey0.lFsfCode = 0
                            gPackDateLong lmMonthStart, tmSdfSrchKey0.iDate(0), tmSdfSrchKey0.iDate(1)
                            tmSdfSrchKey0.sSchStatus = ""
                            gPackTime "12AM", tmSdfSrchKey0.iTime(0), tmSdfSrchKey0.iTime(1)
                            ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.lChfCode = llChfCode) And (tmSdf.iLineNo = tmClf.iLine)
                                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSdfDate
                                If llSdfDate > lmMonthEnd Then
                                    Exit Do
                                End If
                                blIncludeSpot = True
                                If tmSdf.sSpotType = "X" Then
                                    blIncludeSpot = False
                                ElseIf tmSdf.sSchStatus = "M" Then
                                    blIncludeSpot = False
                                Else
                                    If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                                        If tmSdf.lSmfCode > 0 Then
                                            tmSmfSrchKey2.lCode = tmSdf.lCode
                                            ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            If ilRet = BTRV_ERR_NONE Then
                                                gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llMissedDate
                                                If (llMissedDate < lmMonthStart) Or (llMissedDate > lmMonthEnd) Then
                                                    blIncludeSpot = False
                                                End If
                                            Else
                                                blIncludeSpot = False
                                            End If
                                        Else
                                            blIncludeSpot = False
                                        End If
                                    End If
                                End If
                                If blIncludeSpot Then
                                    tmApf.iAiredSpotCount = tmApf.iAiredSpotCount + 1
                                End If
                                ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                        End If
                        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    ilRet = btrUpdate(hmApf, tmApf, imApfRecLen)
                    
                    For ilCol = ADVTNAMEINDEX To ADJDOLLARINDEX Step 1
                        If ilCol = ADVTNAMEINDEX Then
                            slStr = slStr & "," & """" & Trim$(grdAcqCheckGrid.TextMatrix(llRow, ilCol)) & """"
                        Else
                            slStr = slStr & "," & Trim$(grdAcqCheckGrid.TextMatrix(llRow, ilCol))
                        End If
                    Next ilCol
                    gLogMsgWODT "W", hmSave, slStr

                    grdAcqCheckGrid.TextMatrix(llRow, ACQSPOTCOUNTINDEX) = tmApf.iAiredSpotCount
                    grdAcqCheckGrid.TextMatrix(llRow, ORIGACQSPOTCOUNTINDEX) = grdAcqCheckGrid.TextMatrix(llRow, ACQSPOTCOUNTINDEX)
                    grdAcqCheckGrid.TextMatrix(llRow, INVCOUNTINDEX) = grdAcqCheckGrid.TextMatrix(llRow, ACQSPOTCOUNTINDEX)
                    grdAcqCheckGrid.TextMatrix(llRow, ORIGINVCOUNTINDEX) = grdAcqCheckGrid.TextMatrix(llRow, ACQSPOTCOUNTINDEX)
                End If
            End If
        End If
    Next llRow

End Function

Private Function mFindAndFixSpot(llChfCode As Long, ilVefCode As Integer, slSpotType As String, slAcqCost As String, ilHowMany As Integer) As Integer
    'slSpotType: M=Missed; S=Schedule
    Dim slSQLQuery As String
    Dim slSchDate As String
    Dim ilRet As Integer
    Dim blSpotOk As Boolean
    Dim llMissedDate As Long
    
    On Error GoTo ErrHand:
    mFindAndFixSpot = True
    If ilHowMany <= 0 Then
        Exit Function
    End If
    slSQLQuery = "Select sdfCode"
    slSQLQuery = slSQLQuery & " From sdf_Spot_Detail"
    slSQLQuery = slSQLQuery & " Inner Join clf_Contract_Line On sdfChfCode = clfChfCode And clfLine = sdfLineNo"
    If slSpotType = "M" Then
        slSQLQuery = slSQLQuery & " where sdfSchStatus = 'M'"
    Else
        slSQLQuery = slSQLQuery & " where sdfSchStatus <> 'M'"
    End If
    slSQLQuery = slSQLQuery & " And sdfChfCode = " & llChfCode
    slSQLQuery = slSQLQuery & " And sdfVefCode = " & ilVefCode
    slSQLQuery = slSQLQuery & " And sdfSpotType <> 'X'"
    slSQLQuery = slSQLQuery & " And clfAcquisitionCost = " & gStrDecToLong(slAcqCost, 2)
    slSQLQuery = slSQLQuery & " And (sdfDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' And sdfDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "')"
    slSQLQuery = slSQLQuery & " Order By sdfDate"
    'Set rst_Sdf = cnn.Execute(slSQLQuery)
    Set rst_Sdf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Sdf.EOF
        tmSdfSrchKey3.lCode = rst_Sdf!SdfCode
        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            blSpotOk = True
            If tmSdf.lSmfCode > 0 Then
                tmSmfSrchKey2.lCode = tmSdf.lCode
                ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llMissedDate
                    If (llMissedDate < lmMonthStart) Or (llMissedDate > lmMonthEnd) Then
                        blSpotOk = False
                    End If
                Else
                    blSpotOk = False
                End If
            End If
            If blSpotOk Then
                If slSpotType = "M" Then
                    ilRet = mBookSpot(tmSdf)
                    If ilRet Then
                        ilHowMany = ilHowMany - 1
                        If ilHowMany <= 0 Then
                            Exit Function
                        End If
                    End If
                Else
                    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slSchDate
                    imSelectedDay = gWeekDayStr(slSchDate)
                    ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, tmSdf.iGameNo, tmSmf, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay), hmSxf, hmGsf, hmGhf)
                    If ilRet Then
                        ilHowMany = ilHowMany - 1
                        If ilHowMany <= 0 Then
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        rst_Sdf.MoveNext
    Loop
    Exit Function
ErrHand:
    gSetMousePointer grdAcqCheckGrid, grdAcqCheckGrid, vbDefault
    On Error GoTo 0
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then
        ElseIf gErrSQL.Number <> 0 Then
        End If
    Next gErrSQL
    mFindAndFixSpot = False
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mFindAvail                      *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get avail within Ssf           *
'*                                                     *
'*******************************************************
Private Function mFindAvail(ilVefCode As Integer, slSchDate As String, slFindTime As String, ilGameNo As Integer, ilFindAdjAvail As Integer, ilAvailIndex As Integer) As Integer
'
'   ilRet = mFindAvail(slSchDate, slSchTime, ilAvailIndex)
'   Where:
'       slSchDate(I)- Scheduled Date
'       slSchTime(I)- Time that avail is to be found at
'       ilFindAdjAvail(I)- Find closest avail to specified time
'       llSsfRecPos(O)- Ssf record position
'       ilAvailIndex(O)- Index into Ssf where avail is located
'       ilRet(O)- True=Avail found; False=Avail not found
'       lmSsfRecPos(O)- Ssf record position
'
    Dim ilRet As Integer
    Dim llSchDate As Long
    Dim llTime As Long
    Dim llTstTime As Long
    Dim llFndAdjTime As Long
    Dim ilLoop As Integer
    llTime = CLng(gTimeToCurrency(slFindTime, False))
    llSchDate = gDateValue(slSchDate)
    imSelectedDay = gWeekDayStr(slSchDate)
    lmSsfDate(imSelectedDay) = 0
    ilRet = gObtainSsfForDateOrGame(ilVefCode, llSchDate, slFindTime, ilGameNo, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay))
    llFndAdjTime = -1
    For ilLoop = 1 To tmSsf(imSelectedDay).iCount Step 1
       LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilLoop)
        If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
            gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTstTime
            If llTime = llTstTime Then 'Replace
                ilAvailIndex = ilLoop
                mFindAvail = True
                Exit Function
            ElseIf (llTstTime < llTime) And (ilFindAdjAvail) Then
                ilAvailIndex = ilLoop
                llFndAdjTime = llTstTime
            ElseIf (llTime < llTstTime) And (ilFindAdjAvail) Then
                If llFndAdjTime = -1 Then
                    ilAvailIndex = ilLoop
                    mFindAvail = True
                    Exit Function
                Else
                    If (llTime - llFndAdjTime) < (llTstTime - llTime) Then
                        mFindAvail = True
                        Exit Function
                    Else
                        ilAvailIndex = ilLoop
                        mFindAvail = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ilLoop
    If (llFndAdjTime <> -1) And (ilFindAdjAvail) Then
        mFindAvail = True
        Exit Function
    End If
    mFindAvail = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAvailRoom                      *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if room exist for    *
'*                      spot within avail              *
'*                                                     *
'*******************************************************
Private Function mAvailRoom(ilVefCode As Integer, ilAvailIndex As Integer) As Integer
'
'   ilRet = mAvailRoom(ilAvailIndex)
'   where:
'       ilAvailIndex(I)- location of avail within Ssf (use mFindAvail)
'       ilRet(O)- True=Avail has room; False=insufficient room within avail
'
'       tmSdf(I)- spot records
'
'       Code later: ask if avail should be overbooked
'                   If so, create a version zero (0) of the library with the new
'                   units/seconds
'
    Dim ilAvailUnits As Integer
    Dim ilAvailSec As Integer
    Dim ilUnitsSold As Integer
    Dim ilSecSold As Integer
    Dim ilSpotLen As Integer
    Dim ilSpotUnits As Integer
    Dim ilSpotIndex As Integer
    Dim ilNewUnit As Integer
    Dim ilNewSec As Integer
    Dim ilRet As Integer
    Dim ilVpfIndex As Integer
    
    ilVpfIndex = gBinarySearchVpfPlus(ilVefCode)    'gVpfFind(PostLog, imVefCode)
    If ilVpfIndex = -1 Then
        mAvailRoom = False
        Exit Function
    End If
   LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilAvailIndex)
    ilAvailUnits = tmAvail.iAvInfo And &H1F
    ilAvailSec = tmAvail.iLen
    '10/27/11: Disallow more then 31 spots in any avail
    If tmAvail.iNoSpotsThis >= 31 Then
        'ilRet = MsgBox("Move not allowed because Avail contains the maximum number of spots (31).", vbOkOnly + vbExclamation, "Save")
        mAvailRoom = False
        Exit Function
    End If
    For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tmAvail.iNoSpotsThis Step 1
       LSet tmSpot = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilSpotIndex)
        If tmSpot.lSdfCode = tmSdf.lCode Then
            mAvailRoom = True
            Exit Function
        End If
        If (tmSpot.iRecType And &HF) >= 10 Then
            ilSpotLen = tmSpot.iPosLen And &HFFF
            If (tgVpf(ilVpfIndex).sSSellOut = "T") Then
                ilSpotUnits = ilSpotLen \ 30
                If ilSpotUnits <= 0 Then
                    ilSpotUnits = 1
                End If
                ilSpotLen = 0
            Else
                ilSpotUnits = 1
                'If (tgVpf(ilVpfIndex).sSSellOut = "U") Then
                '    ilSpotLen = 0
                'End If
            End If
            If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                ilUnitsSold = ilUnitsSold + ilSpotUnits
                ilSecSold = ilSecSold + ilSpotLen
            End If
        End If
    Next ilSpotIndex
    ilSpotLen = tmSdf.iLen
    If (tgVpf(ilVpfIndex).sSSellOut = "T") Then
        ilSpotUnits = ilSpotLen \ 30
        If ilSpotUnits <= 0 Then
            ilSpotUnits = 1
        End If
        ilSpotLen = 0
    Else
        ilSpotUnits = 1
        'If (tgVpf(ilVpfIndex).sSSellOut = "U") Then
        '    ilSpotLen = 0
        'End If
    End If
    ilNewUnit = 0
    ilNewSec = 0
    If (tgVpf(ilVpfIndex).sSSellOut = "M") Then
        If (ilSpotLen + ilSecSold <> ilAvailSec) Or (ilSpotUnits + ilUnitsSold <> ilAvailUnits) Then
            ilNewSec = ilSpotLen + ilSecSold
            ilNewUnit = ilSpotUnits + ilUnitsSold
        Else
            mAvailRoom = True
            Exit Function
        End If
    Else
        If (ilSpotLen + ilSecSold > ilAvailSec) Or (ilSpotUnits + ilUnitsSold > ilAvailUnits) Then
            ilNewSec = ilSpotLen + ilSecSold
            ilNewUnit = ilSpotUnits + ilUnitsSold
        Else
            mAvailRoom = True
            Exit Function
        End If
    End If
    If (tgVpf(ilVpfIndex).sSOverBook <> "Y") Then
        'ilRet = MsgBox("Move not allowed because Avail would be Overbooked.", vbOkOnly + vbExclamation, "Save")
        mAvailRoom = False
        Exit Function
    End If
    Do
        imSsfRecLen = Len(tmSsf(imSelectedDay))
        ilRet = gSSFGetDirect(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, lmSsfRecPos(imSelectedDay), INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mAvailRoom = False
            Exit Function
        End If
        ilRet = gGetByKeyForUpdateSSF(hmSsf, tmSsf(imSelectedDay))
        If ilRet <> BTRV_ERR_NONE Then
            mAvailRoom = False
            Exit Function
        End If
        '5/20/11
        If (tmAvail.iOrigUnit = 0) And (tmAvail.iOrigLen = 0) Then
            tmAvail.iOrigUnit = tmAvail.iAvInfo And &H1F
            tmAvail.iOrigLen = tmAvail.iLen
        End If
        tmAvail.iAvInfo = (tmAvail.iAvInfo And (Not &H1F)) + ilNewUnit
        tmAvail.iLen = ilNewSec
        tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilAvailIndex) = tmAvail
        imSsfRecLen = igSSFBaseLen + tmSsf(imSelectedDay).iCount * Len(tmProg)
        ilRet = gSSFUpdate(hmSsf, tmSsf(imSelectedDay), imSsfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        mAvailRoom = False
        Exit Function
    End If
    mAvailRoom = True
    Exit Function
End Function
Private Function mBookSpot(tlSdf As SDF) As Boolean
    Dim slAirDate As String
    Dim llAirDate As Long
    Dim slAirTime As String
    Dim ilAvailIndex As Integer
    Dim ilBkQH As Integer
    Dim ilRet As Integer
    Dim llSdfRecPos As Long
    Dim slRet As String
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilPriceLevel As Integer
    Dim ilRdf As Integer
    Dim slAgyCompliant As String
    Dim blFound As Boolean
    Dim ilDay As Integer
    Dim slAllowedDays As String
    Dim slWkSDate As String
    Dim ilTBIndex As Integer
    'Dim llTBStartTime(1 To 7) As Long  'Allowed times if time buy
    Dim llTBStartTime(0 To 6) As Long  'Allowed times if time buy
    Dim llTBEndTime(0 To 6) As Long
    Dim slWkEDate As String
    Dim ilTest As Integer
    Dim llMissedTime As Long
    Dim slClfStartTime As String
    Dim slClfEndTime As String
    Dim llLnStartDate As Long
    Dim llCffStartDate As Long
    Dim llCffEndDate As Long
    Dim ilCount As Integer
    Dim ilLoop As Integer
    
    mBookSpot = False
    
    ilVefCode = tlSdf.iVefCode
    ilVpfIndex = gBinarySearchVpfPlus(ilVefCode)    'gVpfFind(PostLog, imVefCode)
    If ilVpfIndex = -1 Then
        Exit Function
    End If
    
    gUnpackDate tlSdf.iDate(0), tlSdf.iDate(1), slAirDate
    gUnpackTime tlSdf.iTime(0), tlSdf.iTime(1), "A", "1", slAirTime
    llAirDate = gDateValue(slAirDate)
    tmChfSrchKey0.lCode = tlSdf.lChfCode
    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        tmClfSrchKey0.lChfCode = tlSdf.lChfCode
        tmClfSrchKey0.iLine = tlSdf.iLineNo
        tmClfSrchKey0.iCntRevNo = tmChf.iCntRevNo
        tmClfSrchKey0.iPropVer = tmChf.iPropVer
        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Function
        End If
    Else
        Exit Function
    End If
    
    ilTBIndex = 0
    If (tmClf.iStartTime(0) = 1) And (tmClf.iStartTime(1) = 0) Then
        ilRdf = gBinarySearchRdf(tmClf.iRdfCode)
        If ilRdf <> -1 Then
            For ilTest = UBound(tgMRdf(ilRdf).iStartTime, 2) To LBound(tgMRdf(ilRdf).iStartTime, 2) Step -1
                If (tgMRdf(ilRdf).iStartTime(0, ilTest) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilTest) <> 0) Then
                    gUnpackTime tgMRdf(ilRdf).iStartTime(0, ilTest), tgMRdf(ilRdf).iStartTime(1, ilTest), "A", "1", slClfStartTime
                    gUnpackTime tgMRdf(ilRdf).iEndTime(0, ilTest), tgMRdf(ilRdf).iEndTime(1, ilTest), "A", "1", slClfEndTime
                    If ilTBIndex = 0 Then
                        'ilTBIndex = 1
                        ilTBIndex = 0
                        llTBStartTime(ilTBIndex) = gTimeToLong(slClfStartTime, False)
                        llTBEndTime(ilTBIndex) = gTimeToLong(slClfEndTime, True)
                    Else
                        ilTBIndex = ilTBIndex + 1
                        llTBStartTime(ilTBIndex) = gTimeToLong(slClfStartTime, False)
                        llTBEndTime(ilTBIndex) = gTimeToLong(slClfEndTime, True)
                    End If
                End If
            Next ilTest
        Else
            Exit Function
        End If
    Else
        gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slClfStartTime
        gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slClfEndTime
        'ilTBIndex = 1
        ilTBIndex = 0
        llTBStartTime(ilTBIndex) = gTimeToLong(slClfStartTime, False)
        llTBEndTime(ilTBIndex) = gTimeToLong(slClfEndTime, True)
    End If


    slWkSDate = gObtainPrevMonday(slAirDate)
    blFound = False
    tmCffSrchKey0.lChfCode = tmClf.lChfCode
    tmCffSrchKey0.iClfLine = tmClf.iLine
    tmCffSrchKey0.iCntRevNo = tmClf.iCntRevNo
    tmCffSrchKey0.iPropVer = tmClf.iPropVer
    gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llLnStartDate
    gPackDateLong llLnStartDate, tmCffSrchKey0.iStartDate(0), tmCffSrchKey0.iStartDate(1)
    ilRet = btrGetGreaterOrEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmCff.iClfLine = tmClf.iLine) And (tmCff.lChfCode = tmClf.lChfCode)
        gUnpackDateLong tmCff.iStartDate(0), tmCff.iStartDate(1), llCffStartDate
        gUnpackDateLong tmCff.iEndDate(0), tmCff.iEndDate(1), llCffEndDate
        If (llAirDate >= llCffStartDate) And (llAirDate <= llCffStartDate) Then
            blFound = True
            Exit Do
        End If
        ilRet = btrGetNext(hmCff, tmCff, imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If Not blFound Then
        Exit Function
    End If
    If tmCff.sDyWk <> "D" Then
        slAllowedDays = ""
        For ilDay = 0 To 6 Step 1
            If (tmCff.iDay(ilDay) > 0) Or (tmCff.sXDay(ilDay) = "1") Then
                slAllowedDays = slAllowedDays & "Y"
            Else
                slAllowedDays = slAllowedDays & "N"
            End If
        Next ilDay
    Else    'Daily
        slAllowedDays = ""
        For ilDay = 0 To 6 Step 1
            If (tmCff.iDay(ilDay) > 0) Then
                slAllowedDays = slAllowedDays & "Y"
            Else
                slAllowedDays = slAllowedDays & "N"
            End If
        Next ilDay
    End If
    
    'Radomily pick day and time
    Do
        ilDay = Int((7) * Rnd)  '0 to 6
        If Mid(slAllowedDays, ilDay + 1, 1) = "Y" Then
            slAirDate = DateAdd("d", ilDay, slWkSDate)
            Exit Do
        End If
    Loop
    ilCount = 0
    Do
        'If ((ilTBIndex = 1) And (llTBEndTime(1) - llTBStartTime(1) < 3600)) Or (ilCount >= 100) Then
        If ((ilTBIndex = 0) And (llTBEndTime(0) - llTBStartTime(0) < 3600)) Or (ilCount >= 100) Then
            llMissedTime = llTBStartTime(0)
            Exit Do
        Else
            llMissedTime = 60 * CLng(Int(1440 * Rnd))  '0-1439
            ilCount = ilCount + 1
            'For ilLoop = 1 To ilTBIndex Step 1
            For ilLoop = 0 To ilTBIndex Step 1
                If (llMissedTime >= llTBStartTime(ilLoop)) And (llMissedTime < llTBEndTime(ilLoop)) Then
                    Exit Do
                End If
            Next ilLoop
        End If
    Loop
    slAirTime = gFormatTimeLong(llMissedTime, "A", "1")

    slAgyCompliant = "A"

    If Not mFindAvail(ilVefCode, slAirDate, slAirTime, 0, True, ilAvailIndex) Then
        Exit Function
    End If
    If Not mAvailRoom(ilVefCode, ilAvailIndex) Then
        Exit Function
    End If
    If Not mFindAvail(ilVefCode, slAirDate, slAirTime, 0, True, ilAvailIndex) Then
        Exit Function
    End If
    tmSdfSrchKey3.lCode = tlSdf.lCode
    ilRet = btrGetEqual(hmSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Function
    End If
    ilRet = btrGetPosition(hmSdf, llSdfRecPos)
    If slAgyCompliant = "O" Then
        slRet = "O"
    Else
        slRet = "S"
    End If
    'Test if time within daypart, if not set to Outside
    ilBkQH = IMPORTINVOICESPOT
    ilPriceLevel = 0
    ilRet = gBookSpot(slRet, hmSdf, tlSdf, llSdfRecPos, ilBkQH, hmSsf, tmSsf(imSelectedDay), lmSsfRecPos(imSelectedDay), ilAvailIndex, -1, tmChf, tmClf, tmRdf, ilVpfIndex, hmSmf, tmSmf, hmClf, hmCrf, ilPriceLevel, False, hmSxf, hmGsf)
    mBookSpot = ilRet
    If ilRet Then
        tmSdfSrchKey3.lCode = tlSdf.lCode
        ilRet = btrGetEqual(hmSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            If gDateValue(slAirDate) <= lmLastStdMnthBilled Then
                tlSdf.sBill = "Y"
            End If
            tlSdf.sAffChg = "Y"
            gPackDate slAirDate, tlSdf.iDate(0), tlSdf.iDate(1)
            gPackTime slAirTime, tlSdf.iTime(0), tlSdf.iTime(1)
            ilRet = btrUpdate(hmSdf, tlSdf, imSdfRecLen)
        End If
    Else
    End If
End Function

Private Function mFindAndFixMissingApf() As Boolean
    Dim slSQLQuery As String
    Dim blRet As Boolean
    Dim blAddApfFileOpened As Boolean
'    Dim rst_Tmp As ADODB.Recordset
    
    On Error GoTo ErrHand:
    blAddApfFileOpened = False
'    slSQLQuery = "Select distinct chfCntrNo, chfAdfCode, chfAgfCode, chfAgyEst, chfTitle, chfPctTrade, chfSlfCode1, vefName, vefCode, clfAcquisitionCost, clfLine, apfCode "
'    slSQLQuery = slSQLQuery & "from CFF_Contract_Flight Left Outer Join clf_Contract_Line On (cffChfCode = clfChfCode And cffClfLine = clfLine )"
'    slSQLQuery = slSQLQuery & "Left Outer Join chf_Contract_Header On cffChfCode = chfCode "
'    slSQLQuery = slSQLQuery & "Left Outer Join vff_Vehicle_Features On clfVefCode = vffVefCode "
'    slSQLQuery = slSQLQuery & "Left Outer Join vef_Vehicles On clfVefCode = vefCode "
'    slSQLQuery = slSQLQuery & "Left Outer Join apf_Acq_Payable On (chfCntrNo = apfCntrNo And clfAcquisitionCost = apfAcquisitionCost And clfVefCode = apfVefCode And apfInvDate = '" & Format(smMonthEnd, "yyyy-mm-dd") & "') "
'    slSQLQuery = slSQLQuery & "where cffStartDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "' And cffEndDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' and cffStartDate <= cffEndDate "
'    slSQLQuery = slSQLQuery & "And (((cffDyWk = 'W') And (cffSpotsWk > 0)) Or ((cffDyWk = 'D') and ((cffMo > 0) or (cffTu > 0) or (cffWe > 0) or (cffTh > 0) or (cffFr > 0) or (cffSa > 0) or (cffSu > 0)))) "
'    slSQLQuery = slSQLQuery & "And chfSchStatus = 'F' And chfDelete = 'N' And vffPostLogSource = 'S' And apfCode Is Null Order By chfCntrNo"
'    Set rst_Tmp = cnn.Execute(slSQLQuery)
'    Do While Not rst_Tmp.EOF
'        'Add apf if sdf exist?
'        blRet = mAddApf()
'        rst_Tmp.MoveNext
'    Loop
    slSQLQuery = "Select distinct chfCntrNo, chfAdfCode, chfAgfCode, chfProduct, chfAgyEst, chfTitle, chfPctTrade, chfSlfCode1, clfVefCode, clfAcquisitionCost "
    slSQLQuery = slSQLQuery & "from CFF_Contract_Flight Left Outer Join clf_Contract_Line On cffChfCode = clfChfCode And cffClfLine = clfLine "
    slSQLQuery = slSQLQuery & "Left Outer Join chf_Contract_Header On cffChfCode = chfCode "
    slSQLQuery = slSQLQuery & "Left Outer Join vff_Vehicle_Features On clfVefCode = vffVefCode "
    slSQLQuery = slSQLQuery & "where cffStartDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "' And cffEndDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' and cffStartDate <= cffEndDate "
    slSQLQuery = slSQLQuery & "And (((cffDyWk = 'W') And (cffSpotsWk > 0)) Or ((cffDyWk = 'D') and ((cffMo > 0) or (cffTu > 0) or (cffWe > 0) or (cffTh > 0) or (cffFr > 0) or (cffSa > 0) or (cffSu > 0)))) "
    slSQLQuery = slSQLQuery & "And chfSchStatus = 'F' And chfDelete = 'N' And vffPostLogSource = 'S' Order By chfCntrNo"
    'Set rst_Cff = cnn.Execute(slSQLQuery)
    Set rst_Cff = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Cff.EOF
        'Add apf if sdf exist?
        slSQLQuery = "Select apfCode From apf_Acq_Payable"
        '4/20/20: added date range
        slSQLQuery = slSQLQuery & " Where apfCntrNo = " & rst_Cff!chfCntrNo & " And apfAcquisitionCost = " & rst_Cff!clfAcquisitionCost & " And apfVefCode = " & rst_Cff!clfVefCode & " And apfInvDate >= '" & Format(smMonthStart, "yyyy-mm-dd") & "' " & " And apfInvDate <= '" & Format(smMonthEnd, "yyyy-mm-dd") & "' "
        'Set rst_Apf = cnn.Execute(slSQLQuery)
        Set rst_Apf = gSQLSelectCall(slSQLQuery)
        If rst_Apf.EOF Then
            If Not blAddApfFileOpened Then
                gLogMsgWODT "O", hmAddApf, sgDBPath & "Messages\" & "AcqCheck_Add_Apf_" & Format(Now, "mmddyyyy") & ".csv"
                gLogMsgWODT "W", hmAddApf, "Acquisition Check: Adding Apf records for the Month of " & smMonthStart & "-" & smMonthEnd & " on " & Format(Now, "m/d/yy") & " " & Format(Now, "h:mm:ssAM/PM")
                gLogMsgWODT "W", hmAddApf, "Contract #,Advertiser Name,Vehicle Name,Acquisition Cost"
                blAddApfFileOpened = True
            End If
            blRet = mAddApf()
        End If
        rst_Cff.MoveNext
    Loop
    If blAddApfFileOpened Then
        gLogMsgWODT "C", hmAddApf, ""
    End If

    mFindAndFixMissingApf = True
    Exit Function
ErrHand:
    'gDbg_HandleError "AcqCheck: mFindAndFixMissingApf"
    Resume Next
    mFindAndFixMissingApf = False
End Function

Private Function mAddApf() As Boolean
    Dim ilAcqCommPct As Integer
    Dim slFullyPaidDate As String
    Dim ilPrfCode As Integer
    Dim ilLoInx As Integer
    Dim ilHiInx As Integer
    Dim slSQLQuery As String
    Dim llInvNo As Long
    Dim blOk As Boolean
    Dim llRet As Long
    Dim ilAdf As Integer
    Dim slAdvtName As String
    Dim ilVef As Integer
    Dim slVehName As String
    Dim rst_Apf As ADODB.Recordset
    
    On Error GoTo ErrHand:
    llInvNo = -1
    slFullyPaidDate = "1/1/1970"
    slSQLQuery = "Select rvfInvNo from RVF_Receivables Where rvfCntrNo = " & rst_Cff!chfCntrNo & " And rvfAirVefCode = " & rst_Cff!clfVefCode & " And rvfTranType = 'IN'" & " And rvfInvDate = '" & Format(smMonthEnd, "yyyy-mm-dd") & "'"
    'Set rst_Rvf = cnn.Execute(slSQLQuery)
    Set rst_Rvf = gSQLSelectCall(slSQLQuery)
    If rst_Rvf.EOF Then
        slSQLQuery = "Select phfInvNo, phfTranDate from PHF_Payment_History Where phfCntrNo = " & rst_Cff!chfCntrNo & " And phfAirVefCode = " & rst_Cff!clfVefCode & " And phfTranType = 'IN'" & " And phfInvDate = '" & Format(smMonthEnd, "yyyy-mm-dd") & "'"
        'Set rst_Rvf = cnn.Execute(slSQLQuery)
        Set rst_Rvf = gSQLSelectCall(slSQLQuery)
        If rst_Rvf.EOF Then
            mAddApf = False
            Exit Function
        Else
            llInvNo = rst_Rvf!phfInvNo
            slSQLQuery = "Select phfInvNo, phfTranDate from PHF_Payment_History Where phfCntrNo = " & rst_Cff!chfCntrNo & " And phfAirVefCode = " & rst_Cff!clfVefCode & " And phfTranType = 'PI'" & " And phfInvDate = '" & Format(smMonthEnd, "yyyy-mm-dd") & "'"
            'Set rst_Rvf = cnn.Execute(slSQLQuery)
            Set rst_Rvf = gSQLSelectCall(slSQLQuery)
            If Not rst_Rvf.EOF Then
                slFullyPaidDate = Format(rst_Rvf!phfTranDate, "m/d/yy")
            End If
        End If
    Else
        llInvNo = rst_Rvf!rvfInvNo
    End If
    ilPrfCode = 0
    slSQLQuery = "Select prfCode from PRF_Product_Names Where prfAdfCode = " & rst_Cff!chfAdfCode & " And prfName = '" & Trim$(rst_Cff!chfProduct) & "'"
    'Set rst_Prf = cnn.Execute(slSQLQuery)
    Set rst_Prf = gSQLSelectCall(slSQLQuery)
    If Not rst_Prf.EOF Then
        ilPrfCode = rst_Prf!prfCode
    End If

    ilAcqCommPct = 0
    If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then 'Acquisition Commissionable then
        If (rst_Cff!chfPctTrade = 0) Then
            blOk = gGetAcqCommInfoByVehicle(rst_Cff!clfVefCode, ilLoInx, ilHiInx)
            If blOk Then
                ilAcqCommPct = gGetEffectiveAcqComm(gDateValue(smMonthEnd), ilLoInx, ilHiInx)
            End If
        End If
    End If

    slSQLQuery = "Insert Into apf_Acq_Payable ( "
    slSQLQuery = slSQLQuery & "apfCode, "
    slSQLQuery = slSQLQuery & "apfAgfCode, "
    slSQLQuery = slSQLQuery & "apfAdfCode, "
    slSQLQuery = slSQLQuery & "apfPrfCode, "
    slSQLQuery = slSQLQuery & "apfSlfCode, "
    slSQLQuery = slSQLQuery & "apfCntrNo, "
    slSQLQuery = slSQLQuery & "apfInvNo, "
    slSQLQuery = slSQLQuery & "apfAgyEst, "
    slSQLQuery = slSQLQuery & "apfMnfItem, "
    slSQLQuery = slSQLQuery & "apfSbfCode, "
    slSQLQuery = slSQLQuery & "apfInvDate, "
    slSQLQuery = slSQLQuery & "apfOrderSpotCount, "
    slSQLQuery = slSQLQuery & "apfAiredSpotCount, "
    slSQLQuery = slSQLQuery & "apfAcquisitionCost, "
    slSQLQuery = slSQLQuery & "apfAcqCommPct, "
    slSQLQuery = slSQLQuery & "apfFullyPaidDate, "
    slSQLQuery = slSQLQuery & "apfStationInvNo, "
    slSQLQuery = slSQLQuery & "apfStationCntrNo, "
    slSQLQuery = slSQLQuery & "apfVefCode, "
    slSQLQuery = slSQLQuery & "apfUnused "
    slSQLQuery = slSQLQuery & ") "
    slSQLQuery = slSQLQuery & "Values ( "
    slSQLQuery = slSQLQuery & 0 & ", "                          'tlAPF.lCode
    slSQLQuery = slSQLQuery & rst_Cff!chfAgfCode & ", "        'tlAPF.iAgfCode
    slSQLQuery = slSQLQuery & rst_Cff!chfAdfCode & ", "        'tlAPF.iAdfCode & ", "
    slSQLQuery = slSQLQuery & ilPrfCode & ", "                  'tlAPF.lPrfCode
    slSQLQuery = slSQLQuery & rst_Cff!chfSlfCode1 & ", "       'tlAPF.iSlfCode & ", "
    slSQLQuery = slSQLQuery & rst_Cff!chfCntrNo & ", "         'tlAPF.lCntrNo
    slSQLQuery = slSQLQuery & llInvNo & ", "                          'tlAPF.lInvNo & ", "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(Trim$(rst_Cff!chfAgyEst) & Trim(rst_Cff!chfTitle)) & "', "    'gFixQuote(tlAPF.sAgyEst)
    slSQLQuery = slSQLQuery & 0 & ", "                          'tlAPF.iMnfItem
    slSQLQuery = slSQLQuery & 0 & ", "                          'tlAPF.lSbfCode
    slSQLQuery = slSQLQuery & "'" & Format$(smMonthEnd, sgSQLDateForm) & "', " 'tlAPF.sInvDate
    slSQLQuery = slSQLQuery & 0 & ", "                          'tlAPF.iOrderSpotCount
    slSQLQuery = slSQLQuery & 0 & ", "                          'tlAPF.iAiredSpotCount
    slSQLQuery = slSQLQuery & rst_Cff!clfAcquisitionCost & ", "    'tlAPF.lAcquisitionCost
    slSQLQuery = slSQLQuery & ilAcqCommPct & ", "               'tlAPF.iAcqCommPct
    slSQLQuery = slSQLQuery & "'" & Format$(slFullyPaidDate, sgSQLDateForm) & "', "
    slSQLQuery = slSQLQuery & "'" & "" & "', "                  'gFixQuote(tlAPF.sStationInvNo)
    slSQLQuery = slSQLQuery & "'" & "" & "', "                  'gFixQuote(tlAPF.sStationCntrNo)
    slSQLQuery = slSQLQuery & rst_Cff!clfVefCode & ", "           'tlAPF.iVefCode
    slSQLQuery = slSQLQuery & "'" & "" & "' "                   'gFixQuote(tlAPF.sUnused)
    slSQLQuery = slSQLQuery & ") "
    On Error GoTo ErrHand1
    llRet = 0
    'cnn.Execute slSQLQuery
    llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    If llRet <> 0 Then
        mAddApf = False
        Exit Function
    End If
    slAdvtName = ""
    ilAdf = gBinarySearchAdf(rst_Cff!chfAdfCode)
    If ilAdf <> -1 Then
        slAdvtName = Trim$(tgCommAdf(ilAdf).sName)
    End If
    ilVef = gBinarySearchVef(rst_Cff!clfVefCode)
    If ilVef <> -1 Then
        slVehName = Trim$(tgMVef(ilVef).sName)
    End If
    gLogMsgWODT "W", hmAddApf, rst_Cff!chfCntrNo & "," & """" & slAdvtName & """" & "," & """" & slVehName & """" & "," & rst_Cff!clfAcquisitionCost
    mAddApf = True
    Exit Function
ErrHand1:
    llRet = 1
    Resume Next
ErrHand:
    'gDbg_HandleError "AcqCheck: mAddApf"
    mAddApf = False
Resume Next
End Function

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    If imTerminate = True Then
        mTerminate
    End If
End Sub
