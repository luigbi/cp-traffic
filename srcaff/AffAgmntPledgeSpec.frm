VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAgmntPledgeSpec 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   9465
   ControlBox      =   0   'False
   Icon            =   "AffAgmntPledgeSpec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      ItemData        =   "AffAgmntPledgeSpec.frx":08CA
      Left            =   90
      List            =   "AffAgmntPledgeSpec.frx":08CC
      TabIndex        =   24
      Top             =   255
      Visible         =   0   'False
      Width           =   3465
   End
   Begin V81Affiliate.CSI_DayPicker dpcAirDay 
      Height          =   195
      Left            =   1545
      TabIndex        =   9
      Top             =   4950
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   344
      BackColor       =   16777088
      ForeColor       =   -2147483640
      BorderStyle     =   0
      CSI_ShowSelectRangeButtons=   -1  'True
      CSI_AllowMultiSelection=   -1  'True
      CSI_ShowDropDownOnFocus=   -1  'True
      CSI_InputBoxBoxAlignment=   0
      CSI_DayOnColor  =   4638790
      CSI_DayOffColor =   -2147483633
      CSI_RangeFGColor=   0
      CSI_RangeBGColor=   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   0
      Picture         =   "AffAgmntPledgeSpec.frx":08CE
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1140
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.ListBox lbcMultiSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffAgmntPledgeSpec.frx":0BD8
      Left            =   4665
      List            =   "AffAgmntPledgeSpec.frx":0BDA
      MultiSelect     =   2  'Extended
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1515
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Frame frcAirPlayType 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   240
      Left            =   4245
      TabIndex        =   21
      Top             =   375
      Width           =   2355
      Begin VB.OptionButton rbcAirPlayType 
         Caption         =   "Replace"
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton rbcAirPlayType 
         Caption         =   "Add to"
         Height          =   240
         Index           =   1
         Left            =   1110
         TabIndex        =   22
         Top             =   0
         Width           =   930
      End
   End
   Begin VB.ListBox lbcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffAgmntPledgeSpec.frx":0BDC
      Left            =   3195
      List            =   "AffAgmntPledgeSpec.frx":0BDE
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1515
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Frame frcStatus 
      BorderStyle     =   0  'None
      Caption         =   "Pledge Status"
      Height          =   240
      Left            =   2190
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   5145
      Begin VB.OptionButton rbcStatus 
         Caption         =   "Air Cmml Only"
         Height          =   225
         Index           =   2
         Left            =   3345
         TabIndex        =   16
         Top             =   0
         Width           =   1470
      End
      Begin VB.OptionButton rbcStatus 
         Caption         =   "Air in Daypart"
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   18
         Top             =   0
         Value           =   -1  'True
         Width           =   1470
      End
      Begin VB.OptionButton rbcStatus 
         Caption         =   "Delay Cmml/Prg"
         Height          =   225
         Index           =   1
         Left            =   1650
         TabIndex        =   17
         Top             =   0
         Width           =   1830
      End
   End
   Begin VB.PictureBox pbcAirPlayType 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4410
      ScaleHeight     =   240
      ScaleWidth      =   2430
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   375
      Width           =   2430
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   -15
      Width           =   120
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   540
      Width           =   60
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   15
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   10
      Top             =   4665
      Width           =   60
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
      Left            =   2280
      Picture         =   "AffAgmntPledgeSpec.frx":0BE0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1785
      TabIndex        =   5
      Top             =   1620
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Timer tmcFillGrid 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8775
      Top             =   3195
   End
   Begin VB.PictureBox pbcArial 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7755
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   12
      Top             =   3195
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Continue"
      Height          =   315
      Left            =   3120
      TabIndex        =   0
      Top             =   4860
      Width           =   1380
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4920
      TabIndex        =   1
      Top             =   4860
      Width           =   1380
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8235
      Top             =   3135
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5355
      FormDesignWidth =   9465
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAirPlay 
      Height          =   3855
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   795
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   8
      Cols            =   9
      FixedRows       =   3
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
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
      _Band(0).Cols   =   9
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8820
      Picture         =   "AffAgmntPledgeSpec.frx":0CDA
      Top             =   4770
      Width           =   480
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      Picture         =   "AffAgmntPledgeSpec.frx":0FE4
      Top             =   75
      Width           =   480
   End
   Begin VB.Label lacAirPlayType 
      Caption         =   "Current Defined Pledge:"
      Height          =   270
      Left            =   2100
      TabIndex        =   20
      Top             =   375
      Width           =   2145
   End
   Begin VB.Label lacStatus 
      Caption         =   "Pledge Status"
      Height          =   240
      Left            =   1035
      TabIndex        =   19
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lacTitle 
      Alignment       =   2  'Center
      ForeColor       =   &H80000001&
      Height          =   210
      Left            =   660
      TabIndex        =   2
      Top             =   60
      Width           =   8430
   End
End
Attribute VB_Name = "frmAgmntPledgeSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private imBSMode As Integer

'Grid Controls
Private imCtrlVisible As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on

Private smTimeType As String 'A=Avail; D=Daypart; C=CD/Tape
Private smGridfocus As String   'A=Air Play; B=Breakout; D=Daypart
Private smPrevTip As String
Private lmLastClickedRow As Long
Private bmIgnoreChg As Boolean
Private bFormWasAlreadyResized As Boolean
Private imFromArrow As Integer
Private imAirPlayAdj As Integer

Private imFieldChgd As Integer
Private imState As Integer
Private tmSaveDat() As DAT
Private tmBuildDat() As DAT
Private tmTempDat() As DAT
Private tmSvAvailDat() As DAT
Private imColPos(0 To 8) As Integer 'Save column position because of merge

Private tmAirPlaySpec() As AIRPLAYSPEC
Private tmBreakoutSpec() As BREAKOUTSPEC

Private Enum eWeekDays
                MON = 0
                TUE = 1
                WED = 2
                THU = 3
                FRI = 4
                SAT = 5
                SUN = 6
            End Enum

Const AIRDAYSINDEX = 0
Const PARTIALSTARTTIMEINDEX = 1
Const PARTIALENDTIMEINDEX = 2
Const STATUSINDEX = 3
Const AIRPLAYNOINDEX = 4
Const PLEDGESTARTTIMEINDEX = 5
Const PLEDGEENDTIMEINDEX = 6
Const OFFSETDAYINDEX = 7
Const PLEDGEOFFSETTIMEINDEX = 8


Private Sub cmcCancel_Click()
    igAgmntReturn = False
    Unload frmAgmntPledgeSpec
End Sub

Private Sub cmcCancel_GotFocus()
    mSetShow
End Sub

Private Sub cmcDone_Click()
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim blDaybreakDefined As Boolean
    Dim ilWkDayIdx As Integer
    Dim ilLoop As Integer
    Dim llVpf As Long
    Dim sSDate As String
    Dim slAirDays As String
    Dim slStr As String
    Dim ilAirPlaySpec As Integer
    Dim ilNext As Integer
    Dim ilNextDP As Integer
    Dim llDPRow As Long
    Dim ilIdx As Integer
    Dim llPledgeTime As Long
    Dim llVehProgStartTime As Long
    Dim llPartialStartTime As Long
    Dim llPartialEndTime As Long
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim ilDat As Integer
    Dim ilDayIndex As Integer
    Dim ilDayFound As Boolean
    Dim ilSplitDay As Integer
    Dim ilError As Integer
    Dim blByAvails As Boolean
    Dim blByDayparts As Boolean
    Dim slSvFdSTime As String
    Dim slSvFdETime As String
    Dim blCheckTimes As Boolean
    Dim ilPdDay As Integer
    ReDim ilFdDay(0 To 6) As Integer

    
    If (rbcAirPlayType(0).Value = False) And (rbcAirPlayType(1).Value = False) Then
        MsgBox "'Replace' or 'Add To' must be selected", vbOKOnly
        Exit Sub
    End If
    grdAirPlay.Redraw = False
    ilError = False
    For ilRow = grdAirPlay.FixedRows To grdAirPlay.Rows - 1 Step 1
        If grdAirPlay.TextMatrix(ilRow, AIRDAYSINDEX) <> "" Then
            If (grdAirPlay.TextMatrix(ilRow, STATUSINDEX) = "") Or (grdAirPlay.TextMatrix(ilRow, STATUSINDEX) = "Missing") Then
                grdAirPlay.TextMatrix(ilRow, STATUSINDEX) = "Missing"
                grdAirPlay.Row = ilRow
                grdAirPlay.Col = STATUSINDEX
                grdAirPlay.CellForeColor = vbRed
                ilError = True
            End If
            If (grdAirPlay.TextMatrix(ilRow, AIRDAYSINDEX) = "") Or (grdAirPlay.TextMatrix(ilRow, AIRDAYSINDEX) = "Missing") Then
                grdAirPlay.TextMatrix(ilRow, AIRDAYSINDEX) = "Missing"
                grdAirPlay.Row = ilRow
                grdAirPlay.Col = AIRDAYSINDEX
                grdAirPlay.CellForeColor = vbRed
                ilError = True
            End If
        End If
    Next ilRow
    grdAirPlay.Redraw = True
    If ilError Then
        MsgBox "Field data missing, Continue stopped", vbOKOnly
        cmcCancel.SetFocus
        Exit Sub
    End If
    sSDate = Format$(frmAgmnt!txtOnAirDate.Text, sgShowDateForm)
    bmIgnoreChg = True
    llVehProgStartTime = gTimeToLong(sgVehProgStartTime, False)
    'Save current
    gSetMousePointer grdAirPlay, grdAirPlay, vbHourglass
    ilRet = mRetainAirPlay()
    If Not ilRet Then
        gSetMousePointer grdAirPlay, grdAirPlay, vbDefault
        bmIgnoreChg = False
        Exit Sub
    End If
    ReDim tmBuildDat(0 To 0) As DAT
    
    '5/19/16: Moved here to speed-up
    ReDim tgDat(0 To 0) As DAT
    gGetAvails lgPledgeAttCode, igDayPartShttCode, igDayPartVefCode, igDayPartVefCombo, sSDate, True
    ReDim tmSvAvailDat(0 To UBound(tgDat)) As DAT
    For ilDat = 0 To UBound(tgDat) - 1 Step 1
        tmSvAvailDat(ilDat) = tgDat(ilDat)
    Next ilDat

    For ilAirPlaySpec = 0 To UBound(tmAirPlaySpec) - 1 Step 1
        blByAvails = False
        blByDayparts = False
        ilNext = tmAirPlaySpec(ilAirPlaySpec).iFirstBO
        Do While ilNext <> -1
            ReDim tgDat(0 To 0) As DAT
            blByAvails = True
            '5/19/16: moved above
            'gGetAvails lgPledgeAttCode, igDayPartShttCode, igDayPartVefCode, igDayPartVefCombo, sSDate, True
            ReDim tgDat(0 To UBound(tmSvAvailDat)) As DAT
            For ilDat = 0 To UBound(tmSvAvailDat) - 1 Step 1
                tgDat(ilDat) = tmSvAvailDat(ilDat)
            Next ilDat
            
            'Move tgDat into tmBuildDat
            For ilDat = 0 To UBound(tgDat) - 1 Step 1
                tgDat(ilDat).iAirPlayNo = tmAirPlaySpec(ilAirPlaySpec).iAirPlayNo
                tgDat(ilDat).iFirstET = -1
                If Trim$(tmBreakoutSpec(ilNext).sStatus) <> "Live" Then
                    tgDat(ilDat).sEstimatedTime = "Y"
                    If Trim$(tmBreakoutSpec(ilNext).sPledgeStartTime) <> "" Then
                        tgDat(ilDat).sPdSTime = Trim$(tmBreakoutSpec(ilNext).sPledgeStartTime)
                    End If
                    If Trim$(tmBreakoutSpec(ilNext).sPledgeEndTime) <> "" Then
                        tgDat(ilDat).sPdETime = Trim$(tmBreakoutSpec(ilNext).sPledgeEndTime)
                    End If
                Else
                    tgDat(ilDat).sEstimatedTime = "N"
                End If
                '7/15/14
                tgDat(ilDat).sEmbeddedOrROS = "R"
                llVpf = gBinarySearchVpf(CLng(igDayPartVefCode))
                If llVpf <> -1 Then
                    tgDat(ilDat).sEmbeddedOrROS = tgVpfOptions(llVpf).sEmbeddedOrROS
                End If
                If Trim$(tgDat(ilDat).sEmbeddedOrROS) = "" Then
                    tgDat(ilDat).sEmbeddedOrROS = "R"
                End If
                '11/21/14: Retain Delivery Link times and status unless user specified pledge information
                If (Not bgDlfExist) Or (Trim(tmBreakoutSpec(ilNext).sPledgeTime) <> "") Or (Trim$(tmBreakoutSpec(ilNext).sPledgeStartTime) <> "") Or (Trim$(tmBreakoutSpec(ilNext).sPledgeEndTime) <> "") Or (tmBreakoutSpec(ilNext).iPledgeOffsetDay > 0) Then
                    If Trim$(tmBreakoutSpec(ilNext).sStatus) = "Delay" Then
                        tgDat(ilDat).iFdStatus = tgStatusTypes(1).iStatus
                    ElseIf Trim$(tmBreakoutSpec(ilNext).sStatus) = "Delay Cmml/Prg" Then
                        tgDat(ilDat).iFdStatus = tgStatusTypes(9).iStatus
                    ElseIf Trim$(tmBreakoutSpec(ilNext).sStatus) = "Delay Cmml Only" Then
                        tgDat(ilDat).iFdStatus = tgStatusTypes(10).iStatus
                    Else
                        tgDat(ilDat).iFdStatus = tgStatusTypes(0).iStatus
                    End If
                    For ilWkDayIdx = MON To SUN Step 1
                        ilFdDay(ilWkDayIdx) = tgDat(ilDat).iFdDay(ilWkDayIdx)
                        tgDat(ilDat).iFdDay(ilWkDayIdx) = 0
                        tgDat(ilDat).iPdDay(ilWkDayIdx) = 0
                    Next ilWkDayIdx
                    slAirDays = Trim$(tmBreakoutSpec(ilNext).sDays)
                    ilSplitDay = InStr(1, slAirDays, "Breakout:", vbBinaryCompare)
                    If ilSplitDay > 0 Then
                        slAirDays = Trim$(Mid$(slAirDays, ilSplitDay + 9))
                    End If
                    slAirDays = gCreateDayStr(slAirDays)
                    ilDayFound = False
                    For ilWkDayIdx = 0 To 6 Step 1
                        'Select Case ilWkDayIdx
                        '    Case 0
                        '        slStr = "Mo"
                        '    Case 1
                        '        slStr = "Tu"
                        '    Case 2
                        '        slStr = "We"
                        '    Case 3
                        '        slStr = "Th"
                        '    Case 4
                        '        slStr = "Fr"
                        '    Case 5
                        '        slStr = "Sa"
                        '    Case 6
                        '        slStr = "Su"
                        'End Select
                        'If InStr(1, slAirDays, slStr, vbTextCompare) > 0 Then
                        If Mid$(slAirDays, ilWkDayIdx + 1, 1) = "Y" Then
                            If ilFdDay(ilWkDayIdx) = 1 Then
                                ilDayFound = True
                                tgDat(ilDat).iFdDay(ilWkDayIdx) = 1
                                If tmBreakoutSpec(ilNext).iPledgeOffsetDay = 0 Then
                                    tgDat(ilDat).iPdDay(ilWkDayIdx) = 1
                                Else
                                    ilDayIndex = ilWkDayIdx + tmBreakoutSpec(ilNext).iPledgeOffsetDay
                                    If tmBreakoutSpec(ilNext).iPledgeOffsetDay < 0 Then
                                        tgDat(ilDat).sPdDayFed = "B"
                                        If ilDayIndex < 0 Then
                                            tgDat(ilDat).iPdDay(7 + ilDayIndex) = 1
                                        Else
                                            tgDat(ilDat).iPdDay(ilDayIndex) = 1
                                        End If
                                    Else
                                        tgDat(ilDat).sPdDayFed = "A"
                                        If ilDayIndex > 6 Then
                                            tgDat(ilDat).iPdDay(ilDayIndex - 7) = 1
                                        Else
                                            tgDat(ilDat).iPdDay(ilDayIndex) = 1
                                        End If
                                    End If
                                End If
                                blDaybreakDefined = True
                            End If
                        End If
                    Next ilWkDayIdx
                    If Not ilDayFound Then
                        tgDat(ilDat).iAirPlayNo = -1
                    End If
                    llPledgeTime = 0    'llVehProgStartTime
                    If tmBreakoutSpec(ilNext).sPledgeTime <> "" Then
                        'llPledgeTime = gTimeToLong(tmBreakoutSpec(ilNext).sPledgeTime, False)
                        ilPos = InStr(1, tmBreakoutSpec(ilNext).sPledgeTime, "-", vbBinaryCompare)
                        If ilPos > 0 Then
                            llPledgeTime = -gLengthToLong(Mid$(tmBreakoutSpec(ilNext).sPledgeTime, ilPos + 1))
                        Else
                            llPledgeTime = gLengthToLong(Mid$(tmBreakoutSpec(ilNext).sPledgeTime, ilPos + 1))
                        End If
                        'tgDat(ilDat).sPdSTime = Format$(gLongToTime(gTimeToLong(tgDat(ilDat).sPdSTime, False) + (llPledgeTime - llVehProgStartTime)), sgShowTimeWOSecForm)
                        'tgDat(ilDat).sPdETime = Format$(gLongToTime(gTimeToLong(tgDat(ilDat).sPdETime, True) + (llPledgeTime - llVehProgStartTime)), sgShowTimeWOSecForm)
                        tgDat(ilDat).sPdSTime = Format$(gLongToTime(gTimeToLong(tgDat(ilDat).sPdSTime, False) + (llPledgeTime)), sgShowTimeWSecForm)
                        tgDat(ilDat).sPdETime = Format$(gLongToTime(gTimeToLong(tgDat(ilDat).sPdETime, True) + (llPledgeTime)), sgShowTimeWSecForm)
                    End If
                    If (Trim$(tmBreakoutSpec(ilNext).sPartialStartTime) <> "") Or (Trim$(tmBreakoutSpec(ilNext).sPartialEndTime) <> "") Then
                        If (Trim$(tmBreakoutSpec(ilNext).sPartialStartTime) <> "") Then
                            llPartialStartTime = gTimeToLong(tmBreakoutSpec(ilNext).sPartialStartTime, False)
                        Else
                            llPartialStartTime = 0
                        End If
                        If (Trim$(tmBreakoutSpec(ilNext).sPartialEndTime) <> "") Then
                            llPartialEndTime = gTimeToLong(tmBreakoutSpec(ilNext).sPartialEndTime, True)
                        Else
                            llPartialEndTime = 86400
                        End If
                        If llPartialStartTime < llPartialEndTime Then
                            If (gTimeToLong(tgDat(ilDat).sFdETime, True) < llPartialStartTime) Or (gTimeToLong(tgDat(ilDat).sFdSTime, False) >= llPartialEndTime) Then
                                tgDat(ilDat).iFdStatus = tgStatusTypes(8).iStatus
                            End If
                        Else
                            If (gTimeToLong(tgDat(ilDat).sFdETime, True) < llPartialStartTime) And (gTimeToLong(tgDat(ilDat).sFdSTime, False) >= llPartialEndTime) Then
                                tgDat(ilDat).iFdStatus = tgStatusTypes(8).iStatus
                            End If
                        End If
                    End If
                End If
            Next ilDat
            'Move tgDat into tmBuildDat
            For ilDat = 0 To UBound(tgDat) - 1 Step 1
                If tgDat(ilDat).iAirPlayNo <> -1 Then
                    If tgStatusTypes(tgDat(ilDat).iFdStatus).iPledged = 2 Then
                        For ilWkDayIdx = MON To SUN Step 1
                            tgDat(ilDat).iPdDay(ilWkDayIdx) = 0
                        Next ilWkDayIdx
                        tgDat(ilDat).sPdSTime = ""
                        tgDat(ilDat).sPdETime = ""
                        tgDat(ilDat).sPdDayFed = ""
                    End If
                    slAirDays = Trim$(tmBreakoutSpec(ilNext).sDays)
                    ilSplitDay = InStr(1, slAirDays, "Breakout:", vbBinaryCompare)
                    If (ilSplitDay > 0) And (tgStatusTypes(tgDat(ilDat).iFdStatus).iPledged <> 2) Then
                        For ilWkDayIdx = MON To SUN Step 1
                            If tgDat(ilDat).iFdDay(ilWkDayIdx) = 1 Then
                                tmBuildDat(UBound(tmBuildDat)) = tgDat(ilDat)
                                For ilCol = MON To SUN Step 1
                                    If ilCol <> ilWkDayIdx Then
                                        tmBuildDat(UBound(tmBuildDat)).iFdDay(ilCol) = 0
                                    End If
                                    tmBuildDat(UBound(tmBuildDat)).iPdDay(ilCol) = 0
                                Next ilCol
                                If tmBreakoutSpec(ilNext).iPledgeOffsetDay = 0 Then
                                    tmBuildDat(UBound(tmBuildDat)).iPdDay(ilWkDayIdx) = 1
                                Else
                                    ilDayIndex = ilWkDayIdx + tmBreakoutSpec(ilNext).iPledgeOffsetDay
                                    If tmBreakoutSpec(ilNext).iPledgeOffsetDay < 0 Then
                                        tmBuildDat(UBound(tmBuildDat)).sPdDayFed = "B"
                                        If ilDayIndex < 0 Then
                                            tmBuildDat(UBound(tmBuildDat)).iPdDay(7 + ilDayIndex) = 1
                                        Else
                                            tmBuildDat(UBound(tmBuildDat)).iPdDay(ilDayIndex) = 1
                                        End If
                                    Else
                                        tmBuildDat(UBound(tmBuildDat)).sPdDayFed = "A"
                                        If ilDayIndex > 6 Then
                                            tmBuildDat(UBound(tmBuildDat)).iPdDay(ilDayIndex - 7) = 1
                                        Else
                                            tmBuildDat(UBound(tmBuildDat)).iPdDay(ilDayIndex) = 1
                                        End If
                                    End If
                                End If
                                ReDim Preserve tmBuildDat(0 To UBound(tmBuildDat) + 1) As DAT
                            End If
                        Next ilWkDayIdx
                    Else
                        tmBuildDat(UBound(tmBuildDat)) = tgDat(ilDat)
                        ReDim Preserve tmBuildDat(0 To UBound(tmBuildDat) + 1) As DAT
                    End If
                    
                End If
            Next ilDat
            ilNext = tmBreakoutSpec(ilNext).iNextBO
        Loop
        If tmAirPlaySpec(ilAirPlaySpec).sAction = "R" Then
            For ilLoop = 0 To UBound(tmSaveDat) - 1 Step 1
                If tmAirPlaySpec(ilAirPlaySpec).iAirPlayNo = tmSaveDat(ilLoop).iAirPlayNo Then
                    tmSaveDat(ilLoop).iAirPlayNo = -1
                End If
            Next ilLoop
        End If
        ReDim tmTempDat(0 To UBound(tmBuildDat)) As DAT
        For ilDat = 0 To UBound(tmBuildDat) - 1 Step 1
            tmTempDat(ilDat) = tmBuildDat(ilDat)
        Next ilDat
        If blByAvails Then
            ReDim tgDat(0 To 0) As DAT
            gGetAvails lgPledgeAttCode, igDayPartShttCode, igDayPartVefCode, igDayPartVefCombo, sSDate, True
            For ilDat = 0 To UBound(tgDat) - 1 Step 1
                For ilLoop = 0 To UBound(tmTempDat) - 1 Step 1
                    If tmAirPlaySpec(ilAirPlaySpec).iAirPlayNo = tmTempDat(ilLoop).iAirPlayNo Then
                        If (gTimeToLong(tgDat(ilDat).sFdSTime, True) = gTimeToLong(tmTempDat(ilLoop).sFdSTime, True)) Then
                            If (gTimeToLong(tgDat(ilDat).sFdETime, True) = gTimeToLong(tmTempDat(ilLoop).sFdETime, True)) Then
                                For ilWkDayIdx = MON To SUN Step 1
                                    If tmTempDat(ilLoop).iFdDay(ilWkDayIdx) = 1 Then
                                        tgDat(ilDat).iFdDay(ilWkDayIdx) = 0
                                    End If
                                Next ilWkDayIdx
                            End If
                        End If
                    End If
                Next ilLoop
                If tmAirPlaySpec(ilAirPlaySpec).sAction = "A" Then
                    For ilLoop = 0 To UBound(tmSaveDat) - 1 Step 1
                        If tmAirPlaySpec(ilAirPlaySpec).iAirPlayNo = tmSaveDat(ilLoop).iAirPlayNo Then
                            If (gTimeToLong(tgDat(ilDat).sFdSTime, True) = gTimeToLong(tmSaveDat(ilLoop).sFdSTime, True)) Then
                                If (gTimeToLong(tgDat(ilDat).sFdETime, True) = gTimeToLong(tmSaveDat(ilLoop).sFdETime, True)) Then
                                    For ilWkDayIdx = MON To SUN Step 1
                                        If tmSaveDat(ilLoop).iFdDay(ilWkDayIdx) = 1 Then
                                            tgDat(ilDat).iFdDay(ilWkDayIdx) = 0
                                        End If
                                    Next ilWkDayIdx
                                End If
                            End If
                        End If
                    Next ilLoop
                End If
                For ilWkDayIdx = MON To SUN Step 1
                    If tgDat(ilDat).iFdDay(ilWkDayIdx) = 1 Then
                        tgDat(ilDat).iAirPlayNo = tmAirPlaySpec(ilAirPlaySpec).iAirPlayNo
                        tgDat(ilDat).iFdStatus = tgStatusTypes(8).iStatus
                        For ilPdDay = MON To SUN Step 1
                            tgDat(ilDat).iPdDay(ilPdDay) = 0
                        Next ilPdDay
                        tgDat(ilDat).sPdSTime = ""
                        tgDat(ilDat).sPdETime = ""
                        tgDat(ilDat).sPdDayFed = ""
                        tmTempDat(UBound(tmTempDat)) = tgDat(ilDat)
                        ReDim Preserve tmTempDat(0 To UBound(tmTempDat) + 1) As DAT
                        Exit For
                    End If
                Next ilWkDayIdx
            Next ilDat
        End If
        ReDim tmBuildDat(0 To UBound(tmTempDat)) As DAT
        For ilDat = 0 To UBound(tmTempDat) - 1 Step 1
            tmBuildDat(ilDat) = tmTempDat(ilDat)
        Next ilDat
    Next ilAirPlaySpec
    mMergeDat
    'Move tmBuildDat into tgDat
    ReDim tgDat(0 To UBound(tmBuildDat)) As DAT
    For ilDat = 0 To UBound(tmBuildDat) - 1 Step 1
        tgDat(ilDat) = tmBuildDat(ilDat)
    Next ilDat
    For ilLoop = 0 To UBound(tmSaveDat) - 1 Step 1
        If tmSaveDat(ilLoop).iAirPlayNo >= 0 Then
            tgDat(UBound(tgDat)) = tmSaveDat(ilLoop)
            ReDim Preserve tgDat(0 To UBound(tgDat) + 1) As DAT
        End If
    Next ilLoop
    gSetMousePointer grdAirPlay, grdAirPlay, vbDefault
    igAgmntReturn = True
    Unload frmAgmntPledgeSpec
    
End Sub

Private Sub cmcDone_GotFocus()
    mSetShow
End Sub

Private Sub cmcDropDown_Click()
    lbcDropdown.Visible = Not lbcDropdown.Visible
End Sub



Private Sub dpcAirDay_OnChange()
    Dim slStr As String
    
    slStr = Trim$(dpcAirDay.Text)
    If dpcAirDay.CSI_BreakoutDays Then
        slStr = "Breakout: " & slStr
    End If
    If StrComp(Trim$(grdAirPlay.Text), slStr, vbTextCompare) <> 0 Then
        grdAirPlay.Text = slStr
    End If

End Sub

Private Sub edcDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As Integer
    
    slStr = edcDropdown.Text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    If (lmEnableCol = AIRPLAYNOINDEX) Or (lmEnableCol = STATUSINDEX) Then
        llRow = SendMessageByString(lbcDropdown.hwnd, LB_FINDSTRING, -1, slStr)
        If llRow >= 0 Then
            lbcDropdown.ListIndex = llRow
            edcDropdown.Text = lbcDropdown.List(lbcDropdown.ListIndex)
            edcDropdown.SelStart = ilLen
            edcDropdown.SelLength = Len(edcDropdown.Text)
            grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) = lbcDropdown.List(lbcDropdown.ListIndex)
        Else
            lbcDropdown.ListIndex = 0
            edcDropdown.Text = lbcDropdown.List(lbcDropdown.ListIndex)
            edcDropdown.SelStart = 0
            edcDropdown.SelLength = Len(edcDropdown.Text)
            grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) = lbcDropdown.List(lbcDropdown.ListIndex)
        End If
    End If
End Sub

Private Sub edcDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub edcDropdown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If edcDropdown.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub edcDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        gProcessArrowKey Shift, KeyCode, lbcDropdown, True
    End If
End Sub

Private Sub Form_Click()
    mSetShow
    cmcCancel.SetFocus
End Sub

Private Sub Form_Initialize()
    bFormWasAlreadyResized = False
    Me.Width = (Screen.Width) / 1.2
    Me.Height = (Screen.Height) / 1.4
    gSetFonts Me
    gCenterStdAlone Me
End Sub

Private Sub Form_Load()
    Dim ilLoop As Integer

    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAirPlay, grdAirPlay, vbHourglass
    'Load Grid
    imBSMode = False
    imCtrlVisible = False
    lmLastClickedRow = -1
    bmIgnoreChg = False
    smTimeType = "A"
    If igNoAirPlays > 1 Then
        imAirPlayAdj = 1
    Else
        imAirPlayAdj = 0
    End If
    frcStatus.Visible = False
    lacStatus.Visible = False
    lacTitle.Caption = "Pledge for:"
    For ilLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(ilLoop).iCode = igDayPartShttCode Then
            lacTitle.Caption = lacTitle.Caption & " " & Trim$(tgStationInfo(ilLoop).sCallLetters)
            Exit For
        End If
    Next ilLoop
    For ilLoop = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
        If tgVehicleInfo(ilLoop).iCode = igDayPartVefCode Then
            lacTitle.Caption = lacTitle.Caption & " " & Trim$(tgVehicleInfo(ilLoop).sVehicle)
            Exit For
        End If
    Next ilLoop
    lacTitle.Caption = lacTitle.Caption & " " & sgVehProgStartTime & "-" & sgVehProgEndTime   'Trim$(frmAgmnt!lacPrgTimes.Caption)
    imcTrash.Picture = frmDirectory!imcTrashClosed.Picture
    imcTrash.Enabled = False
    smGridfocus = ""
    imFromArrow = False
    pbcSTab.Left = -240
    pbcTab.Left = -240
    pbcClickFocus.Left = -240
    tmcFillGrid.Enabled = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAirPlay, grdAirPlay, vbDefault

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcTrash.Picture = frmDirectory!imcTrashClosed.Picture
End Sub

Private Sub Form_Resize()
    If bFormWasAlreadyResized Then
        Exit Sub
    End If
    bFormWasAlreadyResized = True
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    mSetGridColumns
    mSetGridTitles
    mClearGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmSaveDat
    Erase tmTempDat
    Erase tmBuildDat
    Erase tmSvAvailDat
    Erase tmAirPlaySpec
    Erase tmBreakoutSpec
    gSetMousePointer grdAirPlay, grdAirPlay, vbDefault
    Set frmAgmntPledgeSpec = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    
    grdAirPlay.ColWidth(AIRPLAYNOINDEX) = grdAirPlay.Width * 0.07
    grdAirPlay.ColWidth(PLEDGESTARTTIMEINDEX) = grdAirPlay.Width * 0.09
    grdAirPlay.ColWidth(PLEDGEENDTIMEINDEX) = grdAirPlay.Width * 0.09
    grdAirPlay.ColWidth(AIRDAYSINDEX) = grdAirPlay.Width * 0.16
    grdAirPlay.ColWidth(PARTIALSTARTTIMEINDEX) = grdAirPlay.Width * 0.1
    grdAirPlay.ColWidth(PARTIALENDTIMEINDEX) = grdAirPlay.Width * 0.1
    grdAirPlay.ColWidth(PLEDGEOFFSETTIMEINDEX) = grdAirPlay.Width * 0.12
    grdAirPlay.ColWidth(OFFSETDAYINDEX) = grdAirPlay.Width * 0.06
    grdAirPlay.ColWidth(STATUSINDEX) = grdAirPlay.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To PLEDGEOFFSETTIMEINDEX Step 1
        If (ilCol <> STATUSINDEX) Then
            grdAirPlay.ColWidth(STATUSINDEX) = grdAirPlay.ColWidth(STATUSINDEX) - grdAirPlay.ColWidth(ilCol)
        End If
    Next ilCol
End Sub

Private Sub mSetGridTitles()
    Dim ilCol As Integer
    
    For ilCol = 0 To grdAirPlay.Cols - 1 Step 1
        imColPos(ilCol) = grdAirPlay.ColPos(ilCol)
    Next ilCol
    
    grdAirPlay.TextMatrix(0, AIRDAYSINDEX) = "Feed"
    grdAirPlay.TextMatrix(1, AIRDAYSINDEX) = "Days"
    grdAirPlay.TextMatrix(0, PARTIALSTARTTIMEINDEX) = "Feed"
    grdAirPlay.TextMatrix(1, PARTIALSTARTTIMEINDEX) = "Partial"
    grdAirPlay.TextMatrix(2, PARTIALSTARTTIMEINDEX) = "Start Time"
    grdAirPlay.TextMatrix(0, PARTIALENDTIMEINDEX) = "Feed"
    grdAirPlay.TextMatrix(1, PARTIALENDTIMEINDEX) = "Partial"
    grdAirPlay.TextMatrix(2, PARTIALENDTIMEINDEX) = "End Time"
    
    grdAirPlay.TextMatrix(0, STATUSINDEX) = "Pledge"
    grdAirPlay.TextMatrix(1, STATUSINDEX) = "Status"
    grdAirPlay.TextMatrix(0, AIRPLAYNOINDEX) = "Pledge"
    grdAirPlay.TextMatrix(1, AIRPLAYNOINDEX) = "Air"
    grdAirPlay.TextMatrix(2, AIRPLAYNOINDEX) = "Play"
    grdAirPlay.TextMatrix(0, PLEDGESTARTTIMEINDEX) = "Pledge"
    grdAirPlay.TextMatrix(1, PLEDGESTARTTIMEINDEX) = "Start Time"
    grdAirPlay.TextMatrix(2, PLEDGESTARTTIMEINDEX) = ""
    grdAirPlay.TextMatrix(0, PLEDGEENDTIMEINDEX) = "Pledge"
    grdAirPlay.TextMatrix(1, PLEDGEENDTIMEINDEX) = "End Time"
    grdAirPlay.TextMatrix(2, PLEDGEENDTIMEINDEX) = ""
    grdAirPlay.TextMatrix(0, OFFSETDAYINDEX) = "Pledge"
    grdAirPlay.TextMatrix(1, OFFSETDAYINDEX) = "Offset"
    grdAirPlay.TextMatrix(2, OFFSETDAYINDEX) = "Day +-"
    grdAirPlay.TextMatrix(0, PLEDGEOFFSETTIMEINDEX) = "Pledge"
    grdAirPlay.TextMatrix(1, PLEDGEOFFSETTIMEINDEX) = "Offset"
    grdAirPlay.TextMatrix(2, PLEDGEOFFSETTIMEINDEX) = "Time +-HMS"
    
    grdAirPlay.Row = 0
    grdAirPlay.MergeCells = 2    'flexMergeRestrictColumns
    grdAirPlay.MergeRow(0) = True
    
    grdAirPlay.Row = 0
    grdAirPlay.Col = AIRDAYSINDEX
    grdAirPlay.CellAlignment = flexAlignRightTop
    grdAirPlay.Row = 0
    grdAirPlay.Col = STATUSINDEX
    grdAirPlay.CellAlignment = flexAlignRightTop
    

End Sub

Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    'Blank rows within grid
    'gGrid_Clear grdAirPlay, False
    'Set color within cells
    For llRow = grdAirPlay.FixedRows To grdAirPlay.Rows - 1 Step 1
        For llCol = AIRDAYSINDEX To PLEDGEOFFSETTIMEINDEX Step 1
            grdAirPlay.TextMatrix(llRow, llCol) = ""
            grdAirPlay.Row = llRow
            grdAirPlay.Col = llCol
            grdAirPlay.CellBackColor = vbWhite
        Next llCol
    Next llRow
End Sub


Private Sub grdAirPlay_EnterCell()
    mSetShow
End Sub

Private Sub grdAirPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim slStr As String
    slStr = ""
    If Not imCtrlVisible Then
        If (grdAirPlay.MouseRow >= grdAirPlay.FixedRows) And (grdAirPlay.TextMatrix(grdAirPlay.MouseRow, grdAirPlay.MouseCol)) <> "" Then
            If (grdAirPlay.MouseCol >= AIRDAYSINDEX) And (grdAirPlay.MouseCol <= PLEDGEOFFSETTIMEINDEX) Then
                slStr = Trim$(grdAirPlay.TextMatrix(grdAirPlay.MouseRow, grdAirPlay.MouseCol))
            End If
        End If
    End If
    If smPrevTip <> slStr Then
        grdAirPlay.ToolTipText = slStr
    End If
    smPrevTip = slStr
End Sub

Private Sub grdAirPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer
    If Y < grdAirPlay.RowHeight(0) Then
        Exit Sub
    End If
    ilCol = grdAirPlay.MouseCol
    ilRow = grdAirPlay.MouseRow
    If ilCol < grdAirPlay.FixedCols Then
        grdAirPlay.Redraw = True
        Exit Sub
    End If
    If ilRow < grdAirPlay.FixedRows Then
        grdAirPlay.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdAirPlay.TopRow
    DoEvents
    If ilRow > grdAirPlay.FixedRows Then
        If grdAirPlay.TextMatrix(ilRow - 1, AIRDAYSINDEX) = "" Then
            grdAirPlay.Redraw = False
            Do
                ilRow = ilRow - 1
                If ilRow < grdAirPlay.FixedRows Then
                    Exit Do
                End If
            Loop While Trim(grdAirPlay.TextMatrix(ilRow, AIRDAYSINDEX)) = ""
            ilRow = ilRow + 1
            ilCol = AIRDAYSINDEX
        End If
    End If
    grdAirPlay.Col = ilCol
    grdAirPlay.Row = ilRow
    smGridfocus = "A"
    If Not mColOk() Then
        grdAirPlay.Redraw = True
        Exit Sub
    End If
    grdAirPlay.Redraw = True
    mEnableBox
End Sub

Private Sub grdAirPlay_Scroll()
    mSetShow
End Sub



Private Sub imcKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbcKey.Visible = True
    lbcKey.ZOrder
End Sub

Private Sub imcKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbcKey.Visible = False
End Sub

Private Sub imcTrash_Click()
    Dim iLoop As Integer
    Dim llRow As Long
    Dim llRows As Long
    
    mSetShow
    llRow = grdAirPlay.Row
    llRows = grdAirPlay.Rows
    If (llRow <= grdAirPlay.FixedRows) Or (llRow > grdAirPlay.Rows - 1) Then
        Exit Sub
    End If
    lmTopRow = -1
    If grdAirPlay.TextMatrix(llRow, AIRDAYSINDEX) <> "" Then
        imFieldChgd = True
        grdAirPlay.RemoveItem llRow
        gGrid_FillWithRows grdAirPlay
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
    End If
End Sub

Private Sub imcTrash_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imcTrash.Enabled Then
        imcTrash.Picture = frmDirectory!imcTrashOpened.Picture
    End If
End Sub

Private Sub lbcDropdown_Click()
    edcDropdown.Text = lbcDropdown.List(lbcDropdown.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcDropdown.Visible = False
    End If
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow
End Sub

Private Sub pbcTab_GotFocus()
    Dim slStr As String
    Dim ilNext As Integer
    Dim ilTestValue As Integer
    Dim llEnableRow As Long
    Dim llEnableCol As Long
    Dim llRow As Long

    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If smGridfocus = "A" Then
        If imCtrlVisible Then
            'Branch
            If (grdAirPlay.Row >= grdAirPlay.FixedRows) And (grdAirPlay.Row < grdAirPlay.Rows) And (grdAirPlay.Col >= AIRDAYSINDEX) And (grdAirPlay.Col <= PLEDGEOFFSETTIMEINDEX) Then
                llEnableRow = lmEnableRow
                llEnableCol = lmEnableCol
                mSetShow
                lmEnableRow = llEnableRow
                lmEnableCol = llEnableCol
                ilTestValue = True
                Do
                    ilNext = False
                    Select Case grdAirPlay.Col
                        Case PLEDGEOFFSETTIMEINDEX
                            llRow = grdAirPlay.Rows
                            Do
                                llRow = llRow - 1
                            Loop While grdAirPlay.TextMatrix(llRow, AIRDAYSINDEX) = ""
                            llRow = llRow + 1
                            If (grdAirPlay.Row + 1 < llRow) Then
                                lmTopRow = -1
                                grdAirPlay.Row = grdAirPlay.Row + 1
                                'If Not grdAirPlay.RowIsVisible(grdAirPlay.Row) Then
                                If grdAirPlay.Top + grdAirPlay.RowPos(grdAirPlay.Row) + 15 + edcDropdown.Height > grdAirPlay.Top + grdAirPlay.Height Then
                                    grdAirPlay.TopRow = grdAirPlay.TopRow + 1
                                End If
                                grdAirPlay.Col = AIRDAYSINDEX
                                mEnableBox
                            Else
                                lmTopRow = -1
                                If grdAirPlay.Row + 1 >= grdAirPlay.Rows Then
                                    grdAirPlay.AddItem ""
                                    grdAirPlay.Row = grdAirPlay.Row + 1
                                    'grdAirPlay.Col = APNUMBERINDEX
                                    'grdAirPlay.CellBackColor = LIGHTYELLOW
                                    'grdAirPlay.Text = grdAirPlay.Row
                                Else
                                    grdAirPlay.Row = grdAirPlay.Row + 1
                                End If
                                
                                'If Not grdAirPlay.RowIsVisible(grdAirPlay.Row) Then
                                If grdAirPlay.Top + grdAirPlay.RowPos(grdAirPlay.Row) + 15 + edcDropdown.Height > grdAirPlay.Top + grdAirPlay.Height Then
                                    grdAirPlay.TopRow = grdAirPlay.TopRow + 1
                                End If
                                If grdAirPlay.TextMatrix(grdAirPlay.Row, AIRDAYSINDEX) = "" Then
                                    imFromArrow = True
                                    pbcArrow.Move grdAirPlay.Left - pbcArrow.Width, grdAirPlay.Top + grdAirPlay.RowPos(grdAirPlay.Row) + (grdAirPlay.RowHeight(grdAirPlay.Row) - pbcArrow.Height) / 2
                                    pbcArrow.Visible = True
                                    pbcArrow.SetFocus
                                Else
                                    grdAirPlay.Col = AIRDAYSINDEX
                                    mEnableBox
                                End If
                            End If
                            Exit Sub
                        Case Else
                            grdAirPlay.Col = grdAirPlay.Col + 1
                    End Select
                    If mColOk() Then
                        Exit Do
                    Else
                        ilNext = True
                    End If
                Loop While ilNext
            Else
                grdAirPlay.TopRow = grdAirPlay.FixedRows
                grdAirPlay.Col = PLEDGEOFFSETTIMEINDEX
                Do
                    If grdAirPlay.Row <= grdAirPlay.FixedRows Then
                        cmcDone.SetFocus
                        Exit Sub
                    End If
                    grdAirPlay.Row = grdAirPlay.Rows - 1
                    Do
                        If Not grdAirPlay.RowIsVisible(grdAirPlay.Row) Then
                            grdAirPlay.TopRow = grdAirPlay.TopRow + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    If mColOk() Then
                        Exit Do
                    End If
                Loop
            End If
            lmTopRow = grdAirPlay.TopRow
            mEnableBox
        End If
    ElseIf smGridfocus = "D" Then
    End If
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilNext As Integer

    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If smGridfocus = "A" Then
        If imFromArrow Then
            imFromArrow = False
            grdAirPlay.Col = AIRDAYSINDEX
            mEnableBox
            Exit Sub
        End If
        If imCtrlVisible Then
            'Branch
            Do
                ilNext = False
                Select Case grdAirPlay.Col
                    Case AIRDAYSINDEX
                        If grdAirPlay.Row = grdAirPlay.FixedRows Then
                            mSetShow
                            cmcDone.SetFocus
                            Exit Sub
                        End If
                        lmTopRow = -1
                        grdAirPlay.Row = grdAirPlay.Row - 1
                        If Not grdAirPlay.RowIsVisible(grdAirPlay.Row) Then
                            grdAirPlay.TopRow = grdAirPlay.TopRow - 1
                        End If
                        grdAirPlay.Col = PLEDGEOFFSETTIMEINDEX
                    Case Else
                        grdAirPlay.Col = grdAirPlay.Col - 1
                End Select
                If mColOk() Then
                    Exit Do
                Else
                    ilNext = True
                End If
            Loop While ilNext
            mSetShow
        Else
            lmTopRow = -1
            grdAirPlay.TopRow = grdAirPlay.FixedRows
            grdAirPlay.Row = grdAirPlay.FixedRows
            grdAirPlay.Col = PLEDGEOFFSETTIMEINDEX
            Do
                If mColOk() Then
                    Exit Do
                End If
                If grdAirPlay.Row + 1 >= grdAirPlay.Rows Then
                    cmcDone.SetFocus
                    Exit Sub
                End If
                grdAirPlay.Row = grdAirPlay.Row + 1
                Do
                    If Not grdAirPlay.RowIsVisible(grdAirPlay.Row) Then
                        grdAirPlay.TopRow = grdAirPlay.TopRow + 1
                    Else
                        Exit Do
                    End If
                Loop
            Loop
        End If
        lmTopRow = grdAirPlay.TopRow
    ElseIf smGridfocus = "B" Then
    ElseIf smGridfocus = "D" Then
    End If
    mEnableBox
End Sub


Private Sub tmcFillGrid_Timer()
    Dim llLoop As Long
    Dim llRow As Long
    Dim llCol As Long
    Dim ilTestCol As Integer
    Dim ilSortCol As Integer
    Dim ilPrevSortCol As Integer
    Dim ilPrevSortDirection As Integer
    Dim ilFlt As Integer
    
    tmcFillGrid.Enabled = False
    
    mPopListKey
    lbcKey.FontBold = False
    lbcKey.FontName = "Arial"
    lbcKey.FontBold = False
    lbcKey.FontSize = 8
    lbcKey.Height = (lbcKey.ListCount - 1) * 225 + 225
    lbcKey.Move imcKey.Left, imcKey.Top + imcKey.Height / 2
    
    ReSize1.Enabled = False
    lacTitle.FontBold = True
    grdAirPlay.Height = grdAirPlay.RowPos(grdAirPlay.Rows - 1) + grdAirPlay.RowHeight(grdAirPlay.Rows - 1) + 45
    grdAirPlay.Row = grdAirPlay.FixedRows
    ReDim tmAirPlaySpec(0 To UBound(tgAirPlaySpec)) As AIRPLAYSPEC
    For llLoop = 0 To UBound(tgAirPlaySpec) - 1 Step 1
        tmAirPlaySpec(llLoop) = tgAirPlaySpec(llLoop)
    Next llLoop
    
    ReDim tmBreakoutSpec(0 To UBound(tgBreakoutSpec)) As BREAKOUTSPEC
    For llLoop = 0 To UBound(tgBreakoutSpec) - 1 Step 1
        tmBreakoutSpec(llLoop) = tgBreakoutSpec(llLoop)
    Next llLoop
    
    ReDim tmDPSelection(0 To UBound(tgDPSelection)) As DPSELECTION
    For llLoop = 0 To UBound(tgDPSelection) - 1 Step 1
        tmDPSelection(llLoop) = tgDPSelection(llLoop)
    Next llLoop
    
    ReDim tmSaveDat(0 To UBound(tgDat)) As DAT
    For llLoop = 0 To UBound(tgDat) - 1 Step 1
        tmSaveDat(llLoop) = tgDat(llLoop)
    Next llLoop
    
    
    frmAgmntPledgeSpec.MousePointer = vbHourglass
    gSetMousePointer grdAirPlay, grdAirPlay, vbHourglass
    
    grdAirPlay.Height = cmcDone.Top - grdAirPlay.Top
    gGrid_FillWithRows grdAirPlay
    gGrid_AlignAllColsLeft grdAirPlay
    grdAirPlay.Height = grdAirPlay.RowPos(grdAirPlay.Rows - 1) + grdAirPlay.RowHeight(grdAirPlay.Rows - 1) + 45
    
    grdAirPlay.TextMatrix(grdAirPlay.FixedRows, AIRPLAYNOINDEX) = "1"
    grdAirPlay.TextMatrix(grdAirPlay.FixedRows, STATUSINDEX) = "Live"
    grdAirPlay.Row = grdAirPlay.FixedRows
    grdAirPlay.Col = PLEDGESTARTTIMEINDEX
    grdAirPlay.CellBackColor = LIGHTYELLOW
    grdAirPlay.Col = PLEDGEENDTIMEINDEX
    grdAirPlay.CellBackColor = LIGHTYELLOW
    grdAirPlay.Col = PLEDGEOFFSETTIMEINDEX
    grdAirPlay.CellBackColor = LIGHTYELLOW
    grdAirPlay.Col = OFFSETDAYINDEX
    grdAirPlay.CellBackColor = LIGHTYELLOW
    grdAirPlay.TextMatrix(grdAirPlay.FixedRows, AIRDAYSINDEX) = "M-Su"  '"Mo Tu We Th Fr Sa Su"
    grdAirPlay.TextMatrix(grdAirPlay.FixedRows, PLEDGEOFFSETTIMEINDEX) = ""   'sgVehProgStartTime
    grdAirPlay.TextMatrix(grdAirPlay.FixedRows, OFFSETDAYINDEX) = ""
    grdAirPlay.TextMatrix(grdAirPlay.FixedRows, PARTIALSTARTTIMEINDEX) = ""
    grdAirPlay.TextMatrix(grdAirPlay.FixedRows, PARTIALENDTIMEINDEX) = ""
    frmAgmntPledgeSpec.MousePointer = vbDefault
    gSetMousePointer grdAirPlay, grdAirPlay, vbDefault
    
    cmcDone.Top = grdAirPlay.Top + grdAirPlay.Height + (frmAgmntPledgeSpec.Height - (grdAirPlay.Top + grdAirPlay.Height) - cmcDone.Height) / 2
    cmcCancel.Top = cmcDone.Top
    imcTrash.Top = grdAirPlay.Top + grdAirPlay.Height + (frmAgmntPledgeSpec.Height - (grdAirPlay.Top + grdAirPlay.Height) - imcTrash.Height) / 2
End Sub

Private Function mColOk() As Integer
    mColOk = True
    If grdAirPlay.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
End Function

Private Sub mEnableBox()
    Dim ilIndex As Integer
    Dim slStr As String
    Dim slAirDays As String
    Dim ilWkDayIdx As Integer
    Dim llDP As Long
    Dim ilPos As Integer
    Dim llRow As Long
    
    If smGridfocus = "A" Then
        If (grdAirPlay.Row >= grdAirPlay.FixedRows) And (grdAirPlay.Row < grdAirPlay.Rows) And (grdAirPlay.Col >= AIRDAYSINDEX) And (grdAirPlay.Col <= PLEDGEOFFSETTIMEINDEX) Then
            If grdAirPlay.TextMatrix(grdAirPlay.Row, grdAirPlay.Col) = "Missing" Then
                grdAirPlay.TextMatrix(grdAirPlay.Row, grdAirPlay.Col) = ""
                grdAirPlay.CellForeColor = vbBlack
            End If
            lmEnableRow = grdAirPlay.Row
            lmEnableCol = grdAirPlay.Col
            imCtrlVisible = True
            pbcArrow.Move grdAirPlay.Left - pbcArrow.Width, grdAirPlay.Top + grdAirPlay.RowPos(grdAirPlay.Row) + (grdAirPlay.RowHeight(grdAirPlay.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            Select Case grdAirPlay.Col
                Case AIRPLAYNOINDEX
                    lbcDropdown.Clear
                    For llRow = 1 To igNoAirPlays Step 1
                        lbcDropdown.AddItem llRow
                        lbcDropdown.ItemData(lbcDropdown.NewIndex) = llRow
                    Next llRow
                    If igNoAirPlays > 1 Then
                        lbcDropdown.AddItem "[All]", 0
                        lbcDropdown.ItemData(lbcDropdown.NewIndex) = 0
                    End If
                    If grdAirPlay.TextMatrix(lmEnableRow, AIRPLAYNOINDEX) = "" Then
                        If lmEnableRow = grdAirPlay.FixedRows Then
                            lbcDropdown.ListIndex = imAirPlayAdj
                        Else
                            lbcDropdown.ListIndex = Val(grdAirPlay.TextMatrix(lmEnableRow - 1, AIRPLAYNOINDEX)) - 1 + imAirPlayAdj
                        End If
                    Else
                        lbcDropdown.ListIndex = Val(grdAirPlay.TextMatrix(lmEnableRow, AIRPLAYNOINDEX)) - 1 + imAirPlayAdj
                    End If
                    mSetLbcGridControl
                Case STATUSINDEX
                    lbcDropdown.Clear
                    lbcDropdown.AddItem "Live"
                    lbcDropdown.AddItem "Delay"
                    lbcDropdown.AddItem "Delay Cmml/Prg"
                    lbcDropdown.AddItem "Delay Cmml Only"
                    If grdAirPlay.TextMatrix(lmEnableRow, STATUSINDEX) = "Delay Cmml Only" Then
                        lbcDropdown.ListIndex = 3
                    ElseIf grdAirPlay.TextMatrix(lmEnableRow, STATUSINDEX) = "Delay Cmml/Prg" Then
                        lbcDropdown.ListIndex = 2
                    ElseIf grdAirPlay.TextMatrix(lmEnableRow, STATUSINDEX) = "Delay" Then
                        lbcDropdown.ListIndex = 1
                    ElseIf grdAirPlay.TextMatrix(lmEnableRow, STATUSINDEX) = "Live" Then
                        lbcDropdown.ListIndex = 0
                    Else
                        lbcDropdown.ListIndex = 0
                    End If
                    mSetLbcGridControl
                Case PLEDGESTARTTIMEINDEX
                    edcDropdown.Move grdAirPlay.Left + imColPos(grdAirPlay.Col) + 30, grdAirPlay.Top + grdAirPlay.RowPos(grdAirPlay.Row) + 15, grdAirPlay.ColWidth(grdAirPlay.Col) - 30, grdAirPlay.RowHeight(grdAirPlay.Row) - 15
                    If grdAirPlay.Text = "" Then
                        edcDropdown.Text = ""   'sgVehProgStartTime
                    Else
                        edcDropdown.Text = grdAirPlay.Text
                    End If
                    If edcDropdown.Height > grdAirPlay.RowHeight(grdAirPlay.Row) - 15 Then
                        edcDropdown.FontName = "Arial"
                        edcDropdown.Height = grdAirPlay.RowHeight(grdAirPlay.Row) - 15
                    End If
                    edcDropdown.Visible = True
                    edcDropdown.SetFocus
                Case PLEDGEENDTIMEINDEX
                    edcDropdown.Move grdAirPlay.Left + imColPos(grdAirPlay.Col) + 30, grdAirPlay.Top + grdAirPlay.RowPos(grdAirPlay.Row) + 15, grdAirPlay.ColWidth(grdAirPlay.Col) - 30, grdAirPlay.RowHeight(grdAirPlay.Row) - 15
                    If grdAirPlay.Text = "" Then
                        edcDropdown.Text = ""   'sgVehProgEndTime
                    Else
                        edcDropdown.Text = grdAirPlay.Text
                    End If
                    If edcDropdown.Height > grdAirPlay.RowHeight(grdAirPlay.Row) - 15 Then
                        edcDropdown.FontName = "Arial"
                        edcDropdown.Height = grdAirPlay.RowHeight(grdAirPlay.Row) - 15
                    End If
                    edcDropdown.Visible = True
                    edcDropdown.SetFocus
                Case AIRDAYSINDEX
                    slAirDays = grdAirPlay.TextMatrix(lmEnableRow, AIRDAYSINDEX)
                    dpcAirDay.MaxLength = 0
                    'dpcAirDay.Text = ""
                    dpcAirDay.CSI_AllowBreakoutDays = True
                    ilPos = InStr(1, slAirDays, "Breakout:", vbBinaryCompare)
                    If ilPos > 0 Then
                        slAirDays = Trim$(Mid$(slAirDays, ilPos + 9))
                        dpcAirDay.CSI_BreakoutDays = True
                    Else
                        dpcAirDay.CSI_BreakoutDays = False
                    End If
                    If slAirDays = "" Then
                        dpcAirDay.Text = ""
                        slAirDays = "M-Su"
                    End If
                    dpcAirDay.Text = slAirDays
                    dpcAirDay.Move grdAirPlay.Left + imColPos(grdAirPlay.Col) + 30, grdAirPlay.Top + grdAirPlay.RowPos(grdAirPlay.Row) + 15, grdAirPlay.ColWidth(grdAirPlay.Col) - 30, grdAirPlay.RowHeight(grdAirPlay.Row) - 15
                    dpcAirDay.Visible = True
                    dpcAirDay.SetFocus
                Case PLEDGEOFFSETTIMEINDEX
                    edcDropdown.Move grdAirPlay.Left + imColPos(grdAirPlay.Col) + 30, grdAirPlay.Top + grdAirPlay.RowPos(grdAirPlay.Row) + 15, grdAirPlay.ColWidth(grdAirPlay.Col) - 30, grdAirPlay.RowHeight(grdAirPlay.Row) - 15
                    If grdAirPlay.Text = "" Then
                        If grdAirPlay.Row = grdAirPlay.FixedRows Then
                            edcDropdown.Text = ""   'sgVehProgStartTime
                        Else
                            edcDropdown.Text = grdAirPlay.TextMatrix(grdAirPlay.Row - 1, PLEDGEOFFSETTIMEINDEX)
                        End If
                    Else
                        edcDropdown.Text = grdAirPlay.Text
                    End If
                    If edcDropdown.Height > grdAirPlay.RowHeight(grdAirPlay.Row) - 15 Then
                        edcDropdown.FontName = "Arial"
                        edcDropdown.Height = grdAirPlay.RowHeight(grdAirPlay.Row) - 15
                    End If
                    edcDropdown.Visible = True
                    edcDropdown.SetFocus
                Case OFFSETDAYINDEX
                    edcDropdown.Move grdAirPlay.Left + imColPos(grdAirPlay.Col) + 30, grdAirPlay.Top + grdAirPlay.RowPos(grdAirPlay.Row) + 15, grdAirPlay.ColWidth(grdAirPlay.Col) - 30, grdAirPlay.RowHeight(grdAirPlay.Row) - 15
                    edcDropdown.Text = grdAirPlay.Text
                    If grdAirPlay.Text = "" Then
                        If grdAirPlay.Row = grdAirPlay.FixedRows Then
                            edcDropdown.Text = "0"
                        Else
                            edcDropdown.Text = grdAirPlay.TextMatrix(grdAirPlay.Row - 1, OFFSETDAYINDEX)
                        End If
                    Else
                        edcDropdown.Text = grdAirPlay.Text
                    End If
                    If edcDropdown.Height > grdAirPlay.RowHeight(grdAirPlay.Row) - 15 Then
                        edcDropdown.FontName = "Arial"
                        edcDropdown.Height = grdAirPlay.RowHeight(grdAirPlay.Row) - 15
                    End If
                    edcDropdown.Visible = True
                    edcDropdown.SetFocus
                Case PARTIALSTARTTIMEINDEX
                    edcDropdown.Move grdAirPlay.Left + imColPos(grdAirPlay.Col) + 30, grdAirPlay.Top + grdAirPlay.RowPos(grdAirPlay.Row) + 15, grdAirPlay.ColWidth(grdAirPlay.Col) - 30, grdAirPlay.RowHeight(grdAirPlay.Row) - 15
                    If grdAirPlay.Text = "" Then
                        edcDropdown.Text = ""   'sgVehProgStartTime
                    Else
                        edcDropdown.Text = grdAirPlay.Text
                    End If
                    If edcDropdown.Height > grdAirPlay.RowHeight(grdAirPlay.Row) - 15 Then
                        edcDropdown.FontName = "Arial"
                        edcDropdown.Height = grdAirPlay.RowHeight(grdAirPlay.Row) - 15
                    End If
                    edcDropdown.Visible = True
                    edcDropdown.SetFocus
                Case PARTIALENDTIMEINDEX
                    edcDropdown.Move grdAirPlay.Left + imColPos(grdAirPlay.Col) + 30, grdAirPlay.Top + grdAirPlay.RowPos(grdAirPlay.Row) + 15, grdAirPlay.ColWidth(grdAirPlay.Col) - 30, grdAirPlay.RowHeight(grdAirPlay.Row) - 15
                    If grdAirPlay.Text = "" Then
                        edcDropdown.Text = ""   'sgVehProgEndTime
                    Else
                        edcDropdown.Text = grdAirPlay.Text
                    End If
                    If edcDropdown.Height > grdAirPlay.RowHeight(grdAirPlay.Row) - 15 Then
                        edcDropdown.FontName = "Arial"
                        edcDropdown.Height = grdAirPlay.RowHeight(grdAirPlay.Row) - 15
                    End If
                    edcDropdown.Visible = True
                    edcDropdown.SetFocus
            End Select
        End If
    ElseIf smGridfocus = "D" Then
    End If
    imcTrash.Enabled = True
End Sub

Private Sub mSetShow()
    Dim slStr As String
    Dim llSvRow As Long
    Dim llSvCol As Long
    Dim llRow As Long
    Dim llCol As Long
    Dim ilCount As Integer
    
    If smGridfocus = "A" Then
        llSvRow = grdAirPlay.Row
        llSvCol = grdAirPlay.Col
        If (lmEnableRow >= grdAirPlay.FixedRows) And (lmEnableRow < grdAirPlay.Rows) Then
            'Set any field that that should only be set after user leaves the cell
            Select Case lmEnableCol
                Case AIRPLAYNOINDEX
                    slStr = edcDropdown.Text
                    If grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                        grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    End If
                Case STATUSINDEX
                    If grdAirPlay.TextMatrix(lmEnableRow, STATUSINDEX) = "Live" Then
                        grdAirPlay.Row = lmEnableRow
                        grdAirPlay.Col = PLEDGESTARTTIMEINDEX
                        grdAirPlay.CellBackColor = LIGHTYELLOW
                        grdAirPlay.Col = PLEDGEENDTIMEINDEX
                        grdAirPlay.CellBackColor = LIGHTYELLOW
                        grdAirPlay.Col = PLEDGEOFFSETTIMEINDEX
                        grdAirPlay.CellBackColor = LIGHTYELLOW
                        grdAirPlay.Col = OFFSETDAYINDEX
                        grdAirPlay.CellBackColor = LIGHTYELLOW
                        grdAirPlay.TextMatrix(lmEnableRow, PLEDGESTARTTIMEINDEX) = ""
                        grdAirPlay.TextMatrix(lmEnableRow, PLEDGEENDTIMEINDEX) = ""
                        grdAirPlay.TextMatrix(lmEnableRow, OFFSETDAYINDEX) = ""
                        grdAirPlay.TextMatrix(lmEnableRow, PLEDGEOFFSETTIMEINDEX) = ""    'sgVehProgStartTime
                        grdAirPlay.TextMatrix(lmEnableRow, OFFSETDAYINDEX) = "0"
                    Else
                        grdAirPlay.Row = lmEnableRow
                        grdAirPlay.Col = PLEDGESTARTTIMEINDEX
                        grdAirPlay.CellBackColor = vbWhite
                        grdAirPlay.Col = PLEDGEENDTIMEINDEX
                        grdAirPlay.CellBackColor = vbWhite
                        grdAirPlay.Col = PLEDGEOFFSETTIMEINDEX
                        grdAirPlay.CellBackColor = vbWhite
                        grdAirPlay.Col = OFFSETDAYINDEX
                        grdAirPlay.CellBackColor = vbWhite
                    End If
                Case PLEDGESTARTTIMEINDEX
                    slStr = edcDropdown.Text
                    If grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                        grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    End If
                    If slStr <> "" Then
                        grdAirPlay.Row = lmEnableRow
                        grdAirPlay.Col = PLEDGEOFFSETTIMEINDEX
                        grdAirPlay.CellBackColor = LIGHTYELLOW
                    Else
                        grdAirPlay.Row = lmEnableRow
                        grdAirPlay.Col = PLEDGEOFFSETTIMEINDEX
                        grdAirPlay.CellBackColor = vbWhite
                    End If
                Case PLEDGEENDTIMEINDEX
                    slStr = edcDropdown.Text
                    If grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                        grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    End If
                Case AIRDAYSINDEX
                Case PLEDGEOFFSETTIMEINDEX
                    slStr = edcDropdown.Text
                    If grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                        grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    End If
                    If slStr <> "" Then
                        grdAirPlay.Row = lmEnableRow
                        grdAirPlay.Col = PLEDGESTARTTIMEINDEX
                        grdAirPlay.CellBackColor = LIGHTYELLOW
                        grdAirPlay.Col = PLEDGEENDTIMEINDEX
                        grdAirPlay.CellBackColor = LIGHTYELLOW
                    Else
                        grdAirPlay.Row = lmEnableRow
                        grdAirPlay.Col = PLEDGESTARTTIMEINDEX
                        grdAirPlay.CellBackColor = vbWhite
                        grdAirPlay.Col = PLEDGEENDTIMEINDEX
                        grdAirPlay.CellBackColor = vbWhite
                    End If
                Case OFFSETDAYINDEX
                    slStr = edcDropdown.Text
                    If grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                        grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    End If
                Case PARTIALSTARTTIMEINDEX
                    slStr = edcDropdown.Text
                    If grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                        grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    End If
                Case PARTIALENDTIMEINDEX
                    slStr = edcDropdown.Text
                    If grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                        grdAirPlay.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    End If
            End Select
        End If
        grdAirPlay.Row = llSvRow
        grdAirPlay.Col = llSvCol
    ElseIf smGridfocus = "B" Then
    ElseIf smGridfocus = "D" Then
    End If
    mExpandAirPlay lmEnableRow, lmEnableCol
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    pbcArrow.Visible = False
    edcDropdown.Visible = False
    cmcDropDown.Visible = False
    lbcDropdown.Visible = False
    lbcMultiSelect.Visible = False
    dpcAirDay.Visible = False
    imcTrash.Enabled = False
End Sub


Private Sub mSetCommands()
    Dim ilLoop As Integer
    
    rbcAirPlayType(0).Enabled = False
    If UBound(tmSaveDat) <= LBound(tmSaveDat) Then
        rbcAirPlayType(1).Value = True
    Else
        rbcAirPlayType(0).Enabled = True
    End If
End Sub

Private Function mRetainAirPlay() As Integer
    Dim llRow As Long
    Dim ilAirPlaySpec As Integer
    Dim ilBreakout As Integer
    Dim ilDay As Integer
    Dim ilPrevIndex As Integer
    Dim llDP As Long
    Dim slStr As String
    Dim ilPrevBO As Integer
    Dim ilAirPlay As Integer
    
    'Move spec to arrays
    ReDim tmAirPlaySpec(0 To 0) As AIRPLAYSPEC
    ReDim tmBreakoutSpec(0 To 0) As BREAKOUTSPEC
    For ilAirPlay = 1 To igNoAirPlays Step 1
        ilAirPlaySpec = UBound(tmAirPlaySpec)
        tmAirPlaySpec(ilAirPlaySpec).iAirPlayNo = ilAirPlay
        tmAirPlaySpec(ilAirPlaySpec).iFirstBO = -1
        If rbcAirPlayType(0).Value Then
            'Replace
            tmAirPlaySpec(ilAirPlaySpec).sAction = "R"
        Else
            'Add
            tmAirPlaySpec(ilAirPlaySpec).sAction = "A"
        End If
        ilPrevBO = -1
        For llRow = grdAirPlay.FixedRows To grdAirPlay.Rows - 1 Step 1
            If grdAirPlay.TextMatrix(llRow, AIRDAYSINDEX) <> "" Then
                If ilAirPlay = Val(grdAirPlay.TextMatrix(llRow, AIRPLAYNOINDEX)) Then
                    ilBreakout = UBound(tmBreakoutSpec)
                    If tmAirPlaySpec(ilAirPlaySpec).iFirstBO = -1 Then
                        tmAirPlaySpec(ilAirPlaySpec).iFirstBO = ilBreakout
                    End If
                    tmBreakoutSpec(ilBreakout).sType = "Avails"  'grdAirPlay.TextMatrix(llRow, TYPEINDEX)
                    tmBreakoutSpec(ilBreakout).sStatus = grdAirPlay.TextMatrix(llRow, STATUSINDEX)
                    tmBreakoutSpec(ilBreakout).iFirstDP = -1
                    tmBreakoutSpec(ilBreakout).sPledgeStartTime = Trim(grdAirPlay.TextMatrix(llRow, PLEDGESTARTTIMEINDEX))
                    tmBreakoutSpec(ilBreakout).sPledgeEndTime = Trim(grdAirPlay.TextMatrix(llRow, PLEDGEENDTIMEINDEX))
                    tmBreakoutSpec(ilBreakout).sEstimatedTime = "Y"
                    tmBreakoutSpec(ilBreakout).sDays = grdAirPlay.TextMatrix(llRow, AIRDAYSINDEX)
                    tmBreakoutSpec(ilBreakout).sPledgeTime = Trim(grdAirPlay.TextMatrix(llRow, PLEDGEOFFSETTIMEINDEX))
                    tmBreakoutSpec(ilBreakout).iPledgeOffsetDay = Val(Trim(grdAirPlay.TextMatrix(llRow, OFFSETDAYINDEX)))
                    tmBreakoutSpec(ilBreakout).sPartialStartTime = Trim(grdAirPlay.TextMatrix(llRow, PARTIALSTARTTIMEINDEX))
                    tmBreakoutSpec(ilBreakout).sPartialEndTime = Trim(grdAirPlay.TextMatrix(llRow, PARTIALENDTIMEINDEX))
                    tmBreakoutSpec(ilBreakout).iNextBO = -1
                    If ilPrevBO <> -1 Then
                        tmBreakoutSpec(ilBreakout - 1).iNextBO = ilBreakout
                    End If
                    ilPrevBO = ilBreakout
                    ReDim Preserve tmBreakoutSpec(0 To ilBreakout + 1) As BREAKOUTSPEC
                End If
            End If
        Next llRow
        If tmAirPlaySpec(ilAirPlaySpec).iFirstBO <> -1 Then
            ReDim Preserve tmAirPlaySpec(0 To ilAirPlaySpec + 1) As AIRPLAYSPEC
        End If
    Next ilAirPlay
    mRetainAirPlay = True
            
End Function

Private Sub mSetCtrlPositions()
    lacAirPlayType.Top = 375
    frcAirPlayType.Top = lacAirPlayType.Top
    rbcAirPlayType(0).Top = 0
    rbcAirPlayType(1).Top = 0
End Sub


Private Sub mSetLbcGridControl()
    Dim slStr As String
    Dim ilIndex As Integer
    
     edcDropdown.Move grdAirPlay.Left + imColPos(grdAirPlay.Col) + 30, grdAirPlay.Top + grdAirPlay.RowPos(grdAirPlay.Row) + 15, grdAirPlay.ColWidth(grdAirPlay.Col) - cmcDropDown.Width - 30, grdAirPlay.RowHeight(grdAirPlay.Row) - 15
    cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
    lbcDropdown.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
    gSetListBoxHeight lbcDropdown, 6
    slStr = grdAirPlay.Text
    ilIndex = SendMessageByString(lbcDropdown.hwnd, LB_FINDSTRING, -1, slStr)
    If ilIndex >= 0 Then
        lbcDropdown.ListIndex = ilIndex
        edcDropdown.Text = lbcDropdown.List(lbcDropdown.ListIndex)
    Else
        lbcDropdown.ListIndex = -1
        edcDropdown.Text = ""
    End If
    If edcDropdown.Height > grdAirPlay.RowHeight(grdAirPlay.Row) - 15 Then
        edcDropdown.FontName = "Arial"
        edcDropdown.Height = grdAirPlay.RowHeight(grdAirPlay.Row) - 15
    End If
    edcDropdown.Visible = True
    cmcDropDown.Visible = True
    lbcDropdown.Visible = True
    edcDropdown.SetFocus
End Sub

Private Sub mSetLbcMSGridControl()
    Dim slStr As String
    Dim ilIndex As Integer
    
    lbcMultiSelect.Move grdAirPlay.Left + imColPos(grdAirPlay.Col) + 30, grdAirPlay.Top + grdAirPlay.RowPos(grdAirPlay.Row) + 15, grdAirPlay.ColWidth(grdAirPlay.Col) - cmcDropDown.Width - 30, grdAirPlay.RowHeight(grdAirPlay.Row) - 15
    gSetListBoxHeight lbcMultiSelect, 7
    lbcMultiSelect.Visible = True
    lbcMultiSelect.SetFocus
End Sub

Private Sub mPopListKey()
    Dim slStr As String
    lbcKey.Clear
    lbcKey.AddItem "Status: Indicate how pledge should be generated"
    lbcKey.AddItem "Feed Days:  Indicate which days to combine"
    lbcKey.AddItem "Partial Start and End Time: Indicate which Spots"
    lbcKey.AddItem "      should be Extracted and sent to Station"
    lbcKey.AddItem "Pledge Start and End Time: Indicates time range"
    lbcKey.AddItem "      that the Station will air the spots"
    lbcKey.AddItem "Offset Day: Indicate displacement of Feed day"
    lbcKey.AddItem "      Positive and Negative values allowed"
    lbcKey.AddItem "Pledge Offset Time: Indicate the displacement"
    lbcKey.AddItem "      time of the avails"
End Sub

Private Sub mMergeDat()
    Dim ilDatCarry As Integer
    Dim ilDatNot As Integer
    Dim blCheckTimes As Boolean
    Dim slSvFdETime As String
    Dim ilDay As Integer
    Dim ilBuildDat As Integer
    Dim ilDat As Integer
    Dim ilFound As Integer
    Dim ilTest As Integer
    ReDim tlDat(0 To 0) As DAT
    
'    ilBuildDat = UBound(tmBuildDat)
'    For ilDatCarry = 0 To UBound(tmBuildDat) - 1 Step 1
'        If (tmBuildDat(ilDatCarry).iRdfCode > 0) And (tmBuildDat(ilDatCarry).iFdStatus <> tgStatusTypes(8).iStatus) Then
'            ilDatNot = 0
'            Do While ilDatNot <= ilBuildDat - 1
'                If (tmBuildDat(ilDatCarry).iRdfCode = tmBuildDat(ilDatNot).iRdfCode) And (tmBuildDat(ilDatNot).iFdStatus = tgStatusTypes(8).iStatus) Then
'                    If (tmBuildDat(ilDatCarry).iAirPlayNo = tmBuildDat(ilDatNot).iAirPlayNo) Then
'                        If (gTimeToLong(tmBuildDat(ilDatNot).sFdSTime, False) < gTimeToLong(tmBuildDat(ilDatCarry).sFdSTime, False)) Then
'                            If (gTimeToLong(tmBuildDat(ilDatNot).sFdETime, True) > gTimeToLong(tmBuildDat(ilDatCarry).sFdSTime, False)) Then
'                                slSvFdETime = tmBuildDat(ilDatNot).sFdETime
'                                tmBuildDat(ilDatNot).sFdETime = tmBuildDat(ilDatCarry).sFdSTime
'                                If (gTimeToLong(slSvFdETime, True) > gTimeToLong(tmBuildDat(ilDatCarry).sFdETime, True)) Then
'                                    tmBuildDat(UBound(tmBuildDat)) = tmBuildDat(ilDatNot)
'                                    tmBuildDat(UBound(tmBuildDat)).sFdSTime = tmBuildDat(ilDatCarry).sFdSTime
'                                    tmBuildDat(UBound(tmBuildDat)).sFdETime = slSvFdETime
'                                    ReDim Preserve tmBuildDat(0 To UBound(tmBuildDat) + 1) As DAT
'                                    ilBuildDat = ilBuildDat + 1
'                                End If
'                            End If
'                        Else
'                            If (gTimeToLong(tmBuildDat(ilDatNot).sFdSTime, False) > gTimeToLong(tmBuildDat(ilDatCarry).sFdSTime, False)) Then
'                                If (gTimeToLong(tmBuildDat(ilDatNot).sFdSTime, False) < gTimeToLong(tmBuildDat(ilDatCarry).sFdETime, True)) Then
'                                    If (gTimeToLong(tmBuildDat(ilDatNot).sFdETime, True) > gTimeToLong(tmBuildDat(ilDatCarry).sFdETime, True)) Then
'                                        tmBuildDat(ilDatNot).sFdSTime = tmBuildDat(ilDatCarry).sFdETime
'                                    ElseIf (gTimeToLong(tmBuildDat(ilDatNot).sFdETime, True) <= gTimeToLong(tmBuildDat(ilDatCarry).sFdETime, True)) Then
'                                        'Remove
'                                        tmBuildDat(ilDatNot).iStatus = -1
'                                    End If
'                                End If
'                            Else
'                                If (gTimeToLong(tmBuildDat(ilDatNot).sFdETime, True) > gTimeToLong(tmBuildDat(ilDatCarry).sFdETime, True)) Then
'                                    tmBuildDat(ilDatNot).sFdSTime = tmBuildDat(ilDatCarry).sFdETime
'                                End If
'                            End If
'                        End If
'
'                    End If
'                End If
'                ilDatNot = ilDatNot + 1
'            Loop
'        End If
'    Next ilDatCarry
    'Combine matching
    ReDim tlDat(0 To 0) As DAT
    For ilDat = 0 To UBound(tmBuildDat) - 1 Step 1
        If tmBuildDat(ilDat).iStatus <> -1 Then
            ilFound = False
            If (tmBuildDat(ilDat).iFdStatus = tgStatusTypes(8).iStatus) Then
                For ilTest = ilDat + 1 To UBound(tmBuildDat) - 1 Step 1
                    If (tmBuildDat(ilDat).iRdfCode = tmBuildDat(ilTest).iRdfCode) And (tmBuildDat(ilDat).iFdStatus = tmBuildDat(ilTest).iFdStatus) And (tmBuildDat(ilDat).iAirPlayNo = tmBuildDat(ilTest).iAirPlayNo) Then
                        If (gTimeToLong(tmBuildDat(ilDat).sFdSTime, False) = gTimeToLong(tmBuildDat(ilTest).sFdSTime, False)) Then
                            If (gTimeToLong(tmBuildDat(ilDat).sFdETime, True) = gTimeToLong(tmBuildDat(ilTest).sFdETime, True)) Then
                                ilFound = True
                                For ilDay = MON To SUN Step 1
                                    If (tmBuildDat(ilDat).iFdDay(ilDay) = 1) Then
                                        tmBuildDat(ilTest).iFdDay(ilDay) = 1
                                    End If
                                Next ilDay
                            End If
                        End If
                    End If
                Next ilTest
            End If
            If Not ilFound Then
                If (tmBuildDat(ilDat).iFdStatus = tgStatusTypes(8).iStatus) Then
                    ilFound = True
                    For ilDay = MON To SUN Step 1
                        If (tmBuildDat(ilDat).iFdDay(ilDay) = 1) Then
                            ilFound = False
                            Exit For
                        End If
                    Next ilDay
                End If
            End If
            If Not ilFound Then
                tlDat(UBound(tlDat)) = tmBuildDat(ilDat)
                ReDim Preserve tlDat(0 To UBound(tlDat) + 1) As DAT
            End If
        End If
    Next ilDat
    ReDim tmBuildDat(0 To UBound(tlDat)) As DAT
    For ilDat = 0 To UBound(tlDat) - 1 Step 1
        tmBuildDat(ilDat) = tlDat(ilDat)
    Next ilDat
    'Remove Not Carried that Match Carry
    For ilDat = 0 To UBound(tmBuildDat) - 1 Step 1
        If tmBuildDat(ilDat).iStatus <> -1 Then
            If (tmBuildDat(ilDat).iFdStatus <> tgStatusTypes(8).iStatus) Then
                For ilTest = 0 To UBound(tmBuildDat) - 1 Step 1
                    '5/19/16: added = 8 status
                    If (ilDat <> ilTest) And (tmBuildDat(ilDat).iRdfCode = tmBuildDat(ilTest).iRdfCode) And (tmBuildDat(ilDat).iFdStatus <> tmBuildDat(ilTest).iFdStatus) And (tmBuildDat(ilDat).iAirPlayNo = tmBuildDat(ilTest).iAirPlayNo) And (tmBuildDat(ilTest).iFdStatus = tgStatusTypes(8).iStatus) Then
                        If (gTimeToLong(tmBuildDat(ilDat).sFdSTime, False) = gTimeToLong(tmBuildDat(ilTest).sFdSTime, False)) Then
                            If (gTimeToLong(tmBuildDat(ilDat).sFdETime, True) = gTimeToLong(tmBuildDat(ilTest).sFdETime, True)) Then
                                ilFound = True
                                For ilDay = MON To SUN Step 1
                                    If (tmBuildDat(ilDat).iFdDay(ilDay) <> tmBuildDat(ilTest).iFdDay(ilDay)) Then
                                        ilFound = False
                                    '5/19/16: Remove days in conflict
                                    Else
                                        If tmBuildDat(ilDat).iFdDay(ilDay) <> 0 Then
                                            tmBuildDat(ilTest).iFdDay(ilDay) = 0
                                        End If
                                    End If
                                Next ilDay
                                '5/19/16
                                If Not ilFound Then
                                    ilFound = True
                                    For ilDay = MON To SUN Step 1
                                        If tmBuildDat(ilTest).iFdDay(ilDay) <> 0 Then
                                            ilFound = False
                                            Exit For
                                        End If
                                    Next ilDay
                                End If
                                If ilFound Then
                                    tmBuildDat(ilTest).iStatus = -1
                                End If
                            End If
                        End If
                    End If
                Next ilTest
            End If
        End If
    Next ilDat
    ReDim tlDat(0 To 0) As DAT
    For ilDat = 0 To UBound(tmBuildDat) - 1 Step 1
        If tmBuildDat(ilDat).iStatus <> -1 Then
            tlDat(UBound(tlDat)) = tmBuildDat(ilDat)
            ReDim Preserve tlDat(0 To UBound(tlDat) + 1) As DAT
        End If
    Next ilDat
    ReDim tmBuildDat(0 To UBound(tlDat)) As DAT
    For ilDat = 0 To UBound(tlDat) - 1 Step 1
        tmBuildDat(ilDat) = tlDat(ilDat)
    Next ilDat
End Sub

Private Sub mExpandAirPlay(llRow As Long, llCol As Long)
    Dim ilLoop As Integer
    Dim ilCol As Integer
    Dim ilLastCol As Integer
    Dim llSvRow As Long
    Dim llSvCol As Long
    ReDim llColor(AIRDAYSINDEX To PLEDGEOFFSETTIMEINDEX) As Long
    
    If (llRow >= grdAirPlay.FixedRows) And (llRow < grdAirPlay.Rows) Then
        If grdAirPlay.TextMatrix(llRow, AIRPLAYNOINDEX) = "[All]" Then
            llSvRow = grdAirPlay.Row
            llSvCol = grdAirPlay.Col
            ilLastCol = -1
            grdAirPlay.Row = llRow
            For ilCol = AIRDAYSINDEX To PLEDGEOFFSETTIMEINDEX Step 1
                grdAirPlay.Col = ilCol
                llColor(ilCol) = grdAirPlay.CellBackColor
                If grdAirPlay.CellBackColor <> LIGHTYELLOW Then
                    ilLastCol = ilCol
                End If
            Next ilCol
            If ilLastCol = llCol Then
                grdAirPlay.TextMatrix(llRow, AIRPLAYNOINDEX) = 1
                For ilLoop = 2 To igNoAirPlays Step 1
                    grdAirPlay.AddItem "", llRow + ilLoop - 1
                    grdAirPlay.Row = llRow + ilLoop - 1
                    For ilCol = AIRDAYSINDEX To PLEDGEOFFSETTIMEINDEX Step 1
                        If ilCol <> AIRPLAYNOINDEX Then
                            grdAirPlay.TextMatrix(llRow + ilLoop - 1, ilCol) = grdAirPlay.TextMatrix(llRow, ilCol)
                        Else
                            grdAirPlay.TextMatrix(llRow + ilLoop - 1, ilCol) = ilLoop
                        End If
                        grdAirPlay.Col = ilCol
                        grdAirPlay.CellBackColor = llColor(ilCol)
                    Next ilCol
                Next ilLoop
            End If
            grdAirPlay.Row = llSvRow
            grdAirPlay.Col = llSvCol
        End If
    End If
End Sub
