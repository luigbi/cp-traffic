VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form EngrAsAirCompare 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrAsAirCompare.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11790
   Begin VB.Frame frcNextDiscrepancy 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   7290
      TabIndex        =   13
      Top             =   45
      Visible         =   0   'False
      Width           =   2595
      Begin VB.CommandButton cmcMoveDown 
         Appearance      =   0  'Flat
         Caption         =   "t"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2025
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   90
         Width           =   375
      End
      Begin VB.CommandButton cmcMoveUp 
         Appearance      =   0  'Flat
         Caption         =   "s"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1545
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   90
         Width           =   375
      End
      Begin VB.Label lacDiscrepancy 
         Caption         =   "Next Discrepancy"
         Height          =   255
         Left            =   135
         TabIndex        =   14
         Top             =   120
         Width           =   1440
      End
   End
   Begin VB.CheckBox ckcShow 
      Caption         =   "Show Discrepancies Only"
      Height          =   210
      Left            =   4635
      TabIndex        =   12
      Top             =   165
      Value           =   1  'Checked
      Width           =   2235
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10215
      Top             =   6795
   End
   Begin VB.CheckBox cbcApplyFilter 
      Caption         =   "Apply Filter"
      Enabled         =   0   'False
      Height          =   210
      Left            =   8520
      TabIndex        =   11
      Top             =   6480
      Width           =   1065
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   11700
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   30
      Width           =   45
   End
   Begin VB.ListBox lbcKey 
      BackColor       =   &H00C0FFFF&
      Height          =   2400
      ItemData        =   "EngrAsAirCompare.frx":030A
      Left            =   180
      List            =   "EngrAsAirCompare.frx":032F
      TabIndex        =   9
      Top             =   555
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.CommandButton cmcFilter 
      Caption         =   "&Filter"
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   6390
      Width           =   1200
   End
   Begin VB.PictureBox pbcETab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   5
      Top             =   6675
      Width           =   60
   End
   Begin VB.PictureBox pbcESTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   45
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   540
      Width           =   60
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   15
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   270
      Width           =   60
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   -15
      Picture         =   "EngrAsAirCompare.frx":0427
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4380
      Visible         =   0   'False
      Width           =   90
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   11235
      Top             =   6825
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   7290
      FormDesignWidth =   11790
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5145
      TabIndex        =   7
      Top             =   6390
      Width           =   1200
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   6390
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdEvents 
      Height          =   5655
      Left            =   165
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   630
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   9975
      _Version        =   393216
      Rows            =   4
      Cols            =   44
      FixedRows       =   2
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
      _Band(0).Cols   =   44
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   165
      Picture         =   "EngrAsAirCompare.frx":0731
      Top             =   315
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "As Air Compare"
      Height          =   270
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   3165
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   10845
      Picture         =   "EngrAsAirCompare.frx":0A3B
      Top             =   6555
      Width           =   480
   End
End
Attribute VB_Name = "EngrAsAirCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrAsAirCompare - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private hmSEE As Integer
Private hmCME As Integer
Private hmSOE As Integer
Private hmCTE As Integer


Private imFieldChgd As Integer
Private smState As String
Private imInChg As Integer
Private imBSMode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer
Private lmCharacterWidth As Long
Private imMaxColChars As Integer
Private smAirDate As String
Private hmMsg As Integer
Private imSpotETECode As Integer
Private smSpotEventTypeName As String

Private bmPrinting As Boolean

Private tmFilterValues() As FILTERVALUES    'Same as tgFilterValues except the equals and Not Equals placed first

Private tmSHE As SHE
Private smCurrDEEStamp
Private tmCurrSEE() As SEE
Private tmCTE As CTE

Private tmDee As DEE
Private tmDHE As DHE

Private smCurrAAEStamp As String
Private tmCurrAAE() As AAE
Private imAAEMatchFound() As Integer    'True=Match Event ID found

Private smT1Comment() As String
Private smT2Comment() As String
'Private smEBuses() As String

Private fmUsedWidth As Single
Private fmUnusedWidth As Single
Private imUnusedCount As Integer


'Grid Controls
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private imShowGridBox As Integer
Private imLastColSorted As Integer
Private imLastSort As Integer



Const AIRDATEERRORINDEX = 0
Const AIRDATEINDEX = 1
Const AUTOOFFINDEX = 2
Const DATAERRORINDEX = 3
Const SCHEDULEERRORINDEX = 4
Const EVENTTYPEINDEX = 5
Const EVENTIDINDEX = 6
Const BUSNAMEINDEX = 7
Const BUSCTRLINDEX = 8
Const TIMEINDEX = 9
Const STARTTYPEINDEX = 10
Const FIXEDINDEX = 11
Const ENDTYPEINDEX = 12
Const DURATIONINDEX = 13
Const MATERIALINDEX = 14
Const AUDIONAMEINDEX = 15
Const AUDIOITEMIDINDEX = 16
Const AUDIOISCIINDEX = 17
Const AUDIOCTRLINDEX = 18
Const BACKUPNAMEINDEX = 19
Const BACKUPCTRLINDEX = 20
Const PROTNAMEINDEX = 21
Const PROTITEMIDINDEX = 22
Const PROTISCIINDEX = 23
Const PROTCTRLINDEX = 24
Const RELAY1INDEX = 25
Const RELAY2INDEX = 26
Const FOLLOWINDEX = 27
Const SILENCETIMEINDEX = 28
Const SILENCE1INDEX = 29
Const SILENCE2INDEX = 30
Const SILENCE3INDEX = 31
Const SILENCE4INDEX = 32
Const NETCUE1INDEX = 33
Const NETCUE2INDEX = 34
Const TITLE1INDEX = 35
Const TITLE2INDEX = 36
Const PCODEINDEX = 37
Const ROWTYPEINDEX = 38
Const ROWSORTINDEX = 39
Const SORTTIMEINDEX = 40
Const LIBNAMEINDEX = 41
Const TMCURRSEEINDEX = 42
Const DISCREPANCYINDEX = 43






Private Sub cbcApplyFilter_Click()
    Dim ilCol As Integer
    
    gSetMousePointer grdEvents, grdEvents, vbHourglass
    grdEvents.Redraw = False
    mMoveSEERecToCtrls
    grdEvents.Redraw = False
    If imLastColSorted >= 0 Then
        If imLastSort = flexSortStringNoCaseDescending Then
            imLastSort = flexSortStringNoCaseAscending
        Else
            imLastSort = flexSortStringNoCaseDescending
        End If
        ilCol = imLastColSorted
        mSortCol ilCol
    Else
        imLastSort = -1
        mSortCol TIMEINDEX
    End If
    grdEvents.Redraw = True
    gSetMousePointer grdEvents, grdEvents, vbDefault
End Sub




Private Sub mSortCol(ilCol As Integer)
    Dim llEndRow As Long
    Dim llRow As Long
    Dim slStr As String
    Dim slBus As String
    Dim slTime As String
    Dim slType As String
    Dim ilLen As Integer
    Dim ilETE As Integer
    Dim slEventCategory As String
    
    For llRow = grdEvents.FixedRows To grdEvents.Rows - 1 Step 1
        slStr = Trim$(grdEvents.TextMatrix(llRow, ROWSORTINDEX))
        If slStr <> "" Then
            slStr = Trim$(grdEvents.TextMatrix(llRow, ROWTYPEINDEX))
            If slStr = "SEE" Then
                If (ilCol = TIMEINDEX) Then
                    'slEventCategory = ""
                    'For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                    '    If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                    '        slEventCategory = tgCurrETE(ilETE).sCategory
                     '       If slEventCategory = "A" Then
                     '           slType = "C"    '"B"
                    '        ElseIf slEventCategory = "S" Then
                    '            slType = "B"    '"C"
                    '        Else
                    '            slType = "A"
                    '        End If
                    '        Exit For
                    '    End If
                    'Next ilETE
                    slTime = grdEvents.TextMatrix(llRow, TIMEINDEX)
                    slStr = grdEvents.TextMatrix(llRow, EVENTIDINDEX)
                    If StrComp(slStr, "Missing", vbTextCompare) = 0 Then
                        If llRow + 1 <= grdEvents.Rows Then
                            If Val(grdEvents.TextMatrix(llRow, ROWSORTINDEX)) + 1 = Val(grdEvents.TextMatrix(llRow + 1, ROWSORTINDEX)) Then
                                slTime = grdEvents.TextMatrix(llRow + 1, TIMEINDEX)
                            Else
                                slTime = grdEvents.TextMatrix(llRow - 1, TIMEINDEX)
                            End If
                        Else
                            slTime = grdEvents.TextMatrix(llRow - 1, TIMEINDEX)
                        End If
                    End If
                    slTime = Trim$(Str$(gStrTimeInTenthToLong(slTime, False)))
                    Do While Len(slTime) < 8
                        slTime = "0" & slTime
                    Loop
                    grdEvents.TextMatrix(llRow, SORTTIMEINDEX) = slTime '& slType
                ElseIf (ilCol = DURATIONINDEX) Then
                    slStr = grdEvents.TextMatrix(llRow, DURATIONINDEX)
                    slStr = Trim$(Str$(gStrLengthInTenthToLong(slStr)))
                    Do While Len(slStr) < 8
                        slStr = "0" & slStr
                    Loop
                    grdEvents.TextMatrix(llRow, SORTTIMEINDEX) = slStr
                ElseIf (ilCol = SILENCETIMEINDEX) Then
                    slStr = grdEvents.TextMatrix(llRow, SILENCETIMEINDEX)
                    slStr = Trim$(Str$(gLengthToLong(slStr))) 'Trim$(Str$(gStrLengthInTenthToLong(slStr)))
                    Do While Len(slStr) < 8
                        slStr = "0" & slStr
                    Loop
                    grdEvents.TextMatrix(llRow, SORTTIMEINDEX) = slStr
                Else
                    grdEvents.TextMatrix(llRow, SORTTIMEINDEX) = grdEvents.TextMatrix(llRow, ilCol)
                End If
                slStr = grdEvents.TextMatrix(llRow, SORTTIMEINDEX)
                grdEvents.TextMatrix(llRow, SORTTIMEINDEX) = slStr & grdEvents.TextMatrix(llRow, ROWSORTINDEX)
                If llRow + 1 <= grdEvents.Rows Then
                    If Val(grdEvents.TextMatrix(llRow, ROWSORTINDEX)) + 1 = Val(grdEvents.TextMatrix(llRow + 1, ROWSORTINDEX)) Then
                        grdEvents.TextMatrix(llRow + 1, SORTTIMEINDEX) = slStr & grdEvents.TextMatrix(llRow + 1, ROWSORTINDEX)
                    Else
                        grdEvents.TextMatrix(llRow - 1, SORTTIMEINDEX) = slStr & grdEvents.TextMatrix(llRow - 1, ROWSORTINDEX)
                    End If
                Else
                    grdEvents.TextMatrix(llRow - 1, SORTTIMEINDEX) = slStr & grdEvents.TextMatrix(llRow - 1, ROWSORTINDEX)
                End If
            End If
        End If
    Next llRow
    gGrid_SortByCol grdEvents, ROWSORTINDEX, SORTTIMEINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
End Sub

Private Sub mSetCommands()
    Dim ilRet As Integer
    Dim llRow As Long
    
    If imInChg Then
        Exit Sub
    End If
    
End Sub







Private Sub mGridColumns()
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    
    gGrid_AlignAllColsLeft grdEvents
    mGridColumnWidth
    'Set Titles
    'Set Titles
    For ilCol = BUSNAMEINDEX To BUSCTRLINDEX Step 1
        grdEvents.TextMatrix(0, ilCol) = "Bus"
    Next ilCol
    For ilCol = AUDIONAMEINDEX To AUDIOCTRLINDEX Step 1
        grdEvents.TextMatrix(0, ilCol) = "Audio"
    Next ilCol
    For ilCol = BACKUPNAMEINDEX To BACKUPCTRLINDEX Step 1
        grdEvents.TextMatrix(0, ilCol) = "B'kup"
    Next ilCol
    For ilCol = PROTNAMEINDEX To PROTCTRLINDEX Step 1
        grdEvents.TextMatrix(0, ilCol) = "Protection"
    Next ilCol
    For ilCol = RELAY1INDEX To RELAY2INDEX Step 1
        grdEvents.TextMatrix(0, ilCol) = "Relay"
    Next ilCol
    For ilCol = SILENCETIMEINDEX To SILENCE4INDEX Step 1
        grdEvents.TextMatrix(0, ilCol) = "Silence"
    Next ilCol
    For ilCol = NETCUE1INDEX To NETCUE2INDEX Step 1
        grdEvents.TextMatrix(0, ilCol) = "Netcue"
    Next ilCol
    For ilCol = TITLE1INDEX To TITLE2INDEX Step 1
        grdEvents.TextMatrix(0, ilCol) = "Title"
    Next ilCol
    grdEvents.TextMatrix(1, BUSNAMEINDEX) = "Name"
    grdEvents.TextMatrix(1, BUSCTRLINDEX) = "C"
    grdEvents.TextMatrix(0, AIRDATEERRORINDEX) = "A"
    grdEvents.TextMatrix(1, AIRDATEERRORINDEX) = "D"
    grdEvents.TextMatrix(0, AIRDATEINDEX) = "D"
    grdEvents.TextMatrix(1, AIRDATEINDEX) = ""
    grdEvents.TextMatrix(0, AUTOOFFINDEX) = "A"
    grdEvents.TextMatrix(1, AUTOOFFINDEX) = "O"
    grdEvents.TextMatrix(0, DATAERRORINDEX) = "D"
    grdEvents.TextMatrix(1, DATAERRORINDEX) = "E"
    grdEvents.TextMatrix(0, SCHEDULEERRORINDEX) = "S"
    grdEvents.TextMatrix(1, SCHEDULEERRORINDEX) = "E"
    grdEvents.TextMatrix(0, EVENTTYPEINDEX) = "Event"
    grdEvents.TextMatrix(1, EVENTTYPEINDEX) = "Type"
    grdEvents.TextMatrix(0, EVENTIDINDEX) = "Event"
    grdEvents.TextMatrix(1, EVENTIDINDEX) = "ID"
    grdEvents.TextMatrix(0, TIMEINDEX) = "Time"
    grdEvents.TextMatrix(1, TIMEINDEX) = ""
    grdEvents.TextMatrix(0, STARTTYPEINDEX) = "Start "
    grdEvents.TextMatrix(1, STARTTYPEINDEX) = "Type"
    grdEvents.TextMatrix(0, FIXEDINDEX) = "Fix"
    grdEvents.TextMatrix(0, ENDTYPEINDEX) = "End"
    grdEvents.TextMatrix(1, ENDTYPEINDEX) = "Type"
    grdEvents.TextMatrix(0, DURATIONINDEX) = "Duration"
    grdEvents.TextMatrix(0, MATERIALINDEX) = "Mat"
    grdEvents.TextMatrix(1, MATERIALINDEX) = "Type"
    grdEvents.TextMatrix(1, AUDIONAMEINDEX) = "Name"
    grdEvents.TextMatrix(1, AUDIOITEMIDINDEX) = "Item"
    grdEvents.TextMatrix(1, AUDIOISCIINDEX) = "ISCI"
    grdEvents.TextMatrix(1, AUDIOCTRLINDEX) = "C"
    grdEvents.TextMatrix(1, BACKUPNAMEINDEX) = "Name"
    grdEvents.TextMatrix(1, BACKUPCTRLINDEX) = "C"
    grdEvents.TextMatrix(1, PROTNAMEINDEX) = "Name"
    grdEvents.TextMatrix(1, PROTITEMIDINDEX) = "Item"
    grdEvents.TextMatrix(1, PROTISCIINDEX) = "ISCI"
    grdEvents.TextMatrix(1, PROTCTRLINDEX) = "C"
    grdEvents.TextMatrix(1, RELAY1INDEX) = "1"
    grdEvents.TextMatrix(1, RELAY2INDEX) = "2"
    grdEvents.TextMatrix(0, FOLLOWINDEX) = "Fol-"
    grdEvents.TextMatrix(1, FOLLOWINDEX) = "low"
    grdEvents.TextMatrix(1, SILENCETIMEINDEX) = "Time"
    grdEvents.TextMatrix(1, SILENCE1INDEX) = "1"
    grdEvents.TextMatrix(1, SILENCE2INDEX) = "2"
    grdEvents.TextMatrix(1, SILENCE3INDEX) = "3"
    grdEvents.TextMatrix(1, SILENCE4INDEX) = "4"
    grdEvents.TextMatrix(1, NETCUE1INDEX) = "Start"
    grdEvents.TextMatrix(1, NETCUE2INDEX) = "Stop"
    grdEvents.TextMatrix(1, TITLE1INDEX) = "1"
    grdEvents.TextMatrix(1, TITLE2INDEX) = "2"
    
    grdEvents.Row = 1
    For ilCol = 0 To grdEvents.Cols - 1 Step 1
        grdEvents.Col = ilCol
        grdEvents.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdEvents.Row = 0
    grdEvents.MergeCells = flexMergeRestrictRows
    grdEvents.MergeRow(0) = True
    grdEvents.Row = 0
    grdEvents.Col = EVENTTYPEINDEX
    grdEvents.CellAlignment = flexAlignCenterCenter
    grdEvents.Row = 0
    grdEvents.Col = BUSNAMEINDEX
    grdEvents.CellAlignment = flexAlignCenterCenter
    grdEvents.Row = 0
    grdEvents.Row = 0
    grdEvents.Col = AUDIONAMEINDEX
    grdEvents.CellAlignment = flexAlignCenterCenter
    grdEvents.Row = 0
    grdEvents.Col = BACKUPNAMEINDEX
    grdEvents.CellAlignment = flexAlignCenterCenter
    grdEvents.Row = 0
    grdEvents.Col = PROTNAMEINDEX
    grdEvents.CellAlignment = flexAlignCenterCenter
    grdEvents.Row = 0
    grdEvents.Col = RELAY1INDEX
    grdEvents.CellAlignment = flexAlignCenterCenter
    grdEvents.Row = 0
    grdEvents.Col = SILENCETIMEINDEX
    grdEvents.CellAlignment = flexAlignCenterCenter
    grdEvents.Row = 0
    grdEvents.Col = NETCUE1INDEX
    grdEvents.CellAlignment = flexAlignCenterCenter
    grdEvents.Row = 0
    grdEvents.Col = TITLE1INDEX
    grdEvents.CellAlignment = flexAlignCenterCenter
    grdEvents.Height = cmcCancel.Top - grdEvents.Top - 240    '4 * grdEvents.RowHeight(0) + 15
    gGrid_IntegralHeight grdEvents
    gGrid_Clear grdEvents, True
    
    mGridColumnWidth

End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    Dim ilPass As Integer
    
    
    grdEvents.ColWidth(PCODEINDEX) = 0
    grdEvents.ColWidth(SORTTIMEINDEX) = 0
    grdEvents.ColWidth(LIBNAMEINDEX) = 0
    grdEvents.ColWidth(TMCURRSEEINDEX) = 0
    grdEvents.ColWidth(ROWTYPEINDEX) = 0
    grdEvents.ColWidth(ROWSORTINDEX) = 0
    grdEvents.ColWidth(DISCREPANCYINDEX) = 0
    imUnusedCount = 0
    fmUsedWidth = 0
    fmUnusedWidth = 0
    For ilPass = 0 To 1 Step 1
        grdEvents.ColWidth(AIRDATEERRORINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(AIRDATEERRORINDEX), 100, "Y")
        grdEvents.ColWidth(AIRDATEINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(AIRDATEINDEX), 100, "N")
        grdEvents.ColWidth(AUTOOFFINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(AUTOOFFINDEX), 100, "Y")
        grdEvents.ColWidth(DATAERRORINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(DATAERRORINDEX), 100, "Y")
        grdEvents.ColWidth(SCHEDULEERRORINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(SCHEDULEERRORINDEX), 100, "Y")
        grdEvents.ColWidth(EVENTTYPEINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(EVENTTYPEINDEX), 65, "Y")
        grdEvents.ColWidth(EVENTIDINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(EVENTIDINDEX), 30, "Y")
        grdEvents.ColWidth(BUSNAMEINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(BUSNAMEINDEX), 32, tgSchUsedSumEPE.sBus)
        grdEvents.ColWidth(BUSCTRLINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(BUSCTRLINDEX), 65, tgSchUsedSumEPE.sBusControl)
        grdEvents.ColWidth(TIMEINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(TIMEINDEX), 24, tgSchUsedSumEPE.sTime)  '21
        grdEvents.ColWidth(STARTTYPEINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(STARTTYPEINDEX), 40, tgSchUsedSumEPE.sStartType)
        grdEvents.ColWidth(FIXEDINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(FIXEDINDEX), 50, tgSchUsedSumEPE.sFixedTime)
        grdEvents.ColWidth(ENDTYPEINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(ENDTYPEINDEX), 40, tgSchUsedSumEPE.sEndType)
        grdEvents.ColWidth(DURATIONINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(DURATIONINDEX), 24, tgSchUsedSumEPE.sDuration)  '25
        grdEvents.ColWidth(MATERIALINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(MATERIALINDEX), 40, tgSchUsedSumEPE.sMaterialType)
        grdEvents.ColWidth(AUDIONAMEINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(AUDIONAMEINDEX), 27, tgSchUsedSumEPE.sAudioName)
        grdEvents.ColWidth(AUDIOITEMIDINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(AUDIOITEMIDINDEX), 27, tgSchUsedSumEPE.sAudioItemID)
        grdEvents.ColWidth(AUDIOISCIINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(AUDIOISCIINDEX), 27, tgSchUsedSumEPE.sAudioISCI)
        grdEvents.ColWidth(AUDIOCTRLINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(AUDIOCTRLINDEX), 65, tgSchUsedSumEPE.sAudioControl)
        grdEvents.ColWidth(BACKUPNAMEINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(BACKUPNAMEINDEX), 27, tgSchUsedSumEPE.sBkupAudioName)
        grdEvents.ColWidth(BACKUPCTRLINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(BACKUPCTRLINDEX), 65, tgSchUsedSumEPE.sBkupAudioControl)
        grdEvents.ColWidth(PROTNAMEINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(PROTNAMEINDEX), 27, tgSchUsedSumEPE.sProtAudioName)
        grdEvents.ColWidth(PROTITEMIDINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(PROTITEMIDINDEX), 27, tgSchUsedSumEPE.sProtAudioItemID)
        grdEvents.ColWidth(PROTISCIINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(PROTISCIINDEX), 27, tgSchUsedSumEPE.sProtAudioISCI)
        grdEvents.ColWidth(PROTCTRLINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(PROTCTRLINDEX), 65, tgSchUsedSumEPE.sProtAudioControl)
        grdEvents.ColWidth(RELAY1INDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(RELAY1INDEX), 32, tgSchUsedSumEPE.sRelay1)
        grdEvents.ColWidth(RELAY2INDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(RELAY2INDEX), 32, tgSchUsedSumEPE.sRelay2)
        grdEvents.ColWidth(FOLLOWINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(FOLLOWINDEX), 40, tgSchUsedSumEPE.sFollow)
        grdEvents.ColWidth(SILENCETIMEINDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(SILENCETIMEINDEX), 30, tgSchUsedSumEPE.sSilenceTime)
        grdEvents.ColWidth(SILENCE1INDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(SILENCE1INDEX), 65, tgSchUsedSumEPE.sSilence1)
        grdEvents.ColWidth(SILENCE2INDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(SILENCE2INDEX), 65, tgSchUsedSumEPE.sSilence2)
        grdEvents.ColWidth(SILENCE3INDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(SILENCE3INDEX), 65, tgSchUsedSumEPE.sSilence3)
        grdEvents.ColWidth(SILENCE4INDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(SILENCE4INDEX), 65, tgSchUsedSumEPE.sSilence4)
        grdEvents.ColWidth(NETCUE1INDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(NETCUE1INDEX), 40, tgSchUsedSumEPE.sStartNetcue)
        grdEvents.ColWidth(NETCUE2INDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(NETCUE2INDEX), 40, tgSchUsedSumEPE.sStopNetcue)
        'grdEvents.ColWidth(TITLE1INDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(TITLE1INDEX), 53, tgSchUsedSumEPE.sTitle1)
        'grdEvents.ColWidth(TITLE2INDEX) = mComputeWidth(ilPass, grdEvents.ColWidth(TITLE2INDEX), 53, tgSchUsedSumEPE.sTitle2)
        If imUnusedCount = 0 Then
            Exit For
        End If
    Next ilPass
    
    grdEvents.ColWidth(TITLE1INDEX) = grdEvents.Width - GRIDSCROLLWIDTH
    For ilCol = AIRDATEERRORINDEX To TITLE2INDEX Step 1
        If ilCol <> TITLE1INDEX And ilCol <> TITLE2INDEX Then
            If grdEvents.ColWidth(TITLE1INDEX) > grdEvents.ColWidth(ilCol) Then
                grdEvents.ColWidth(TITLE1INDEX) = grdEvents.ColWidth(TITLE1INDEX) - grdEvents.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
    grdEvents.ColWidth(TITLE2INDEX) = grdEvents.ColWidth(TITLE1INDEX) / 3
    grdEvents.ColWidth(TITLE1INDEX) = grdEvents.ColWidth(TITLE1INDEX) - grdEvents.ColWidth(TITLE2INDEX)


End Sub


Private Sub mClearControls()
    gGrid_Clear grdEvents, True
    grdEvents.BackColor = vbWhite
End Sub

Private Sub mPopulate()
'    Dim ilRet As Integer
'    Dim ilLoop As Integer
'    Dim llRow As Long
'
'
'    ilRet = gGetRec_DHE_DayHeaderInfo(lgLibCallCode, "EngrAsAirCompare-mPopulation", tmDHE)
'    ilRet = gGetRecs_DEE_DayEvent(sgCurrDEEStamp, lgLibCallCode, "EngrAsAirCompare-mPopulate", tgCurrDEE())
'    If lgLibCallCode <= 0 Then
'        tmDHE.lCode = 0
'    End If
End Sub

Private Sub ckcShow_Click()
    If ckcShow.Value = vbChecked Then
        frcNextDiscrepancy.Visible = False
    Else
        frcNextDiscrepancy.Visible = True
    End If
    gSetMousePointer grdEvents, grdEvents, vbHourglass
    mMoveSEERecToCtrls
    imLastSort = -1
    imLastColSorted = -1
    grdEvents.Redraw = False
    mSortCol TIMEINDEX
    grdEvents.Redraw = True
    gSetMousePointer grdEvents, grdEvents, vbDefault
End Sub

Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrAsAirCompare
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdEvents, grdEvents, vbDefault
    Unload EngrAsAirCompare
    Exit Sub

End Sub

Private Sub mShowEvents(slAsAirDate As String)
    Dim ilRet As Integer
    Dim llAirDate As Long
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim llLoop As Long
    Dim llRow As Long
    Dim ilETE As Integer
    Dim slCategory As String
    Dim llAvailTest As Long
    Dim llTimeTest As Long
    Dim llAvailLength As Long
    
    smAirDate = slAsAirDate
    If Not gIsDate(smAirDate) Then
        Beep
        gSetMousePointer grdEvents, grdEvents, vbDefault
        MsgBox "Invalid Date Specified"
        cmcCancel.SetFocus
        Exit Sub
    End If
    gSetMousePointer grdEvents, grdEvents, vbHourglass
    'DoEvents
    llAirDate = gDateValue(smAirDate)
    slDateTime = gNow()
    slNowDate = Format(slDateTime, "ddddd")
    slNowTime = Format(slDateTime, "ttttt")
    llNowDate = gDateValue(slNowDate)
    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
    ilRet = gGetRec_SHE_ScheduleHeaderByDate(smAirDate, "EngrSchedule-Get Schedule by Date", tmSHE)
    If Not ilRet Then
        gSetMousePointer grdEvents, grdEvents, vbDefault
        MsgBox "Schedule has not beed created for specified date"
        cmcCancel.SetFocus
        Exit Sub
    Else
        ilRet = gGetRecs_SEE_ScheduleEventsAPI(hmSEE, sgCurrSEEStamp, -1, tmSHE.lCode, "EngrAsAirCompare-Get Events", tgCurrSEE())
    End If
    ilRet = gGetRecs_AAE_As_Aired(smCurrAAEStamp, tmSHE.lCode, "EngrAsAirCompare-Get Events", tmCurrAAE())
    If Not ilRet Then
        gSetMousePointer grdEvents, grdEvents, vbDefault
        MsgBox "Error Retrieving As Aired for specified date"
        cmcCancel.SetFocus
        Exit Sub
    End If
    'If UBound(tmCurrAAE) <= LBound(tmCurrAAE) Then
    '    gSetMousePointer grdEvents, grdEvents, vbDefault
    '    MsgBox "As Aired not Imported for specified date"
    '    cmcCancel.SetFocus
    '    Exit Sub
    'End If
    mPopGrid
    grdEvents.Redraw = True
    If (imSpotETECode > 0) And (llAirDate >= llNowDate) Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igJobStatus(SCHEDULEJOB) = 2) Then
            cmcDone.Enabled = True
        Else
            cmcDone.Enabled = False
        End If
    Else
        'cmcDone.Enabled = False
        'imcInsert.Enabled = False
        'imcTrash.Enabled = False
    End If
    mSetCommands
    gSetMousePointer grdEvents, grdEvents, vbDefault
End Sub



Private Sub cmcFilter_Click()
    Dim ilCol As Integer
    Dim ilFilter As Integer
    Dim ilIndex As Integer
    
    gSetMousePointer grdEvents, grdEvents, vbHourglass
    '6/14/06- remove check as done with save
    'If Not mCheckFields(True) Then
    '    gSetMousePointer grdEvents, grdEvents, vbDefault
    '    MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Schedule"
    '    mSortErrorsToTop
    '    Exit Sub
    'End If
    mCreateUsedArrays
    mInitFilterInfo
    igAnsFilter = 0
    gSetMousePointer grdEvents, grdEvents, vbDefault
    EngrSchdFilter.Show vbModal
    If igAnsFilter = CALLDONE Then 'Apply
        gSetMousePointer grdEvents, grdEvents, vbHourglass
        'Reorder, Place Equal and Not Equal at Top
        ReDim tmFilterValues(0 To UBound(tgFilterValues)) As FILTERVALUES
        For ilFilter = LBound(tgFilterValues) To UBound(tgFilterValues) - 1 Step 1
            tgFilterValues(ilFilter).iUsed = False
        Next ilFilter
        ilIndex = LBound(tmFilterValues)
        For ilFilter = LBound(tgFilterValues) To UBound(tgFilterValues) - 1 Step 1
            If tgFilterValues(ilFilter).iUsed = False Then
                If (tgFilterValues(ilFilter).iOperator = 1) Or (tgFilterValues(ilFilter).iOperator = 2) Then
                    LSet tmFilterValues(ilIndex) = tgFilterValues(ilFilter)
                    ilIndex = ilIndex + 1
                    tgFilterValues(ilFilter).iUsed = True
                End If
            End If
        Next ilFilter
        For ilFilter = LBound(tgFilterValues) To UBound(tgFilterValues) - 1 Step 1
            If tgFilterValues(ilFilter).iUsed = False Then
                LSet tmFilterValues(ilIndex) = tgFilterValues(ilFilter)
                ilIndex = ilIndex + 1
                tgFilterValues(ilFilter).iUsed = True
            End If
        Next ilFilter
        If UBound(tgFilterValues) > LBound(tgFilterValues) Then
            cbcApplyFilter.Enabled = True
        End If
        If cbcApplyFilter.Value = vbChecked Then
            grdEvents.Redraw = False
            mMoveSEERecToCtrls
            grdEvents.Redraw = False
            If imLastColSorted >= 0 Then
                If imLastSort = flexSortStringNoCaseDescending Then
                    imLastSort = flexSortStringNoCaseAscending
                Else
                    imLastSort = flexSortStringNoCaseDescending
                End If
                ilCol = imLastColSorted
                mSortCol ilCol
            Else
                imLastSort = -1
                mSortCol TIMEINDEX
            End If
            grdEvents.Redraw = True
        Else
            cbcApplyFilter.Value = vbChecked
        End If
    End If
    gSetMousePointer grdEvents, grdEvents, vbDefault
End Sub

Private Sub Form_Activate()
    'mGridColumns
    If imFirstActivate Then
    End If
    imFirstActivate = False
    Me.KeyPreview = True
End Sub

Private Sub Form_Click()
    cmcCancel.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrAsAirCompare
    gCenterFormModal EngrAsAirCompare
'    Unload EngrLib
End Sub

Private Sub Form_Load()
    mGridColumns
    mInit
    igJobShowing(SCHEDULEJOB) = 3
End Sub

Private Sub Form_Resize()
    Dim llRow As Long
    'These call are here and in form_Active (call to mGridColumns)
    'They are in mGridColumn in case the For_Initialize size chage does not cause a resize event
    mGridColumnWidth
    grdEvents.Height = cmcCancel.Top - grdEvents.Top - 240    '4 * grdEvents.RowHeight(0) + 15
    gGrid_IntegralHeight grdEvents
    gGrid_FillWithRows grdEvents
    imcPrint.Top = cmcCancel.Top
    lmCharacterWidth = CLng(pbcETab.TextWidth("n"))
    gSetListBoxHeight lbcKey, 2 * grdEvents.Height
    lbcKey.Height = lbcKey.Height + lbcKey.Height / 10

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    btrDestroy hmSEE
    btrDestroy hmCME
    btrDestroy hmSOE
    btrDestroy hmCTE
    
    Erase smT1Comment
    Erase smT2Comment
    Erase tmCurrSEE
    
    Erase tmCurrAAE
    
    Erase tmFilterValues
    
    Erase imAAEMatchFound
    
    
    Set EngrAsAirCompare = Nothing
    EngrSchd.Show vbModeless
End Sub





Private Sub mInit()
    Dim llRow As Long
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilETE As Integer
    
    On Error GoTo ErrHand
    
    gSetMousePointer grdEvents, grdEvents, vbHourglass
    imcPrint.Picture = EngrMain!imcPrinter.Picture
    ReDim tgFilterValues(0 To 0) As FILTERVALUES
    ReDim tgFilterFields(0 To 0) As FIELDSELECTION
    ReDim tmFilterValues(0 To 0) As FILTERVALUES
    ReDim tmCurrSEE(0 To 0) As SEE
    'Can't be 0 to 0 because of index in grid
'    cmcSearch.Top = 30
'    edcSearch.Top = cmcSearch.Top
    hmSEE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSEE, "", sgDBPath & "SEE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmCME = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCME, "", sgDBPath & "CME.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmSOE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSOE, "", sgDBPath & "SOE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmCTE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCTE, "", sgDBPath & "CTE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    tmSHE.lCode = 0
    imIgnoreScroll = False
    imLastColSorted = -1
    imLastSort = -1
    imFirstActivate = True
    imInChg = True
    bmPrinting = False
    mPopANE
    mPopASE
    mPopBDE
    mPopCCE_Audio
    mPopCCE_Bus
    mPopCTE
    mPopDNE
    mPopDSE
    mPopETE
    mPopFNE
    mPopMTE
    mPopNNE
    mPopRNE
    mPopSCE
    mPopTTE_EndType
    mPopTTE_StartType
    mPopARE
    ilRet = gGetTypeOfRecs_ATE_AudioType("C", sgCurrATEStamp, "EngrAudioType-mPopulate", tgCurrATE())
    imSpotETECode = 0
    smSpotEventTypeName = "Spot"
    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
        If tgCurrETE(ilETE).sCategory = "S" Then
            imSpotETECode = tgCurrETE(ilETE).iCode
            smSpotEventTypeName = Trim$(tgCurrETE(ilETE).sName)
            Exit For
        End If
    Next ilETE
    If imSpotETECode <= 0 Then
        gSetMousePointer grdEvents, grdEvents, vbDefault
        MsgBox "Spot Event Type not defined", vbCritical + vbOKOnly, "Schedule"
    End If
    lacScreen.Caption = "As Air Compare Date: " & sgAsAirCompareDate
    imInChg = False
    imFieldChgd = False
    If imSpotETECode > 0 Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igJobStatus(SCHEDULEJOB) = 2) Then
            cmcDone.Enabled = True
        Else
            cmcDone.Enabled = False
        End If
    Else
        cmcDone.Enabled = False
    End If
    
    gSetListBoxHeight lbcKey, grdEvents.Height
    mSetCommands
    tmcStart.Enabled = True
    gSetMousePointer grdEvents, grdEvents, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdEvents, grdEvents, vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Relay Definition-Form Load: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Relay Definition-Form Load: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    Resume Next
End Sub

Private Sub grdEvents_Click()
    If grdEvents.Col >= grdEvents.Cols - 1 Then
        Exit Sub
    End If

End Sub

Private Sub grdEvents_GotFocus()
    If grdEvents.Col >= grdEvents.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdEvents_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdEvents.TopRow
    grdEvents.Redraw = False
End Sub

Private Sub grdEvents_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilFound As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    Dim llAAE As Long
    
    grdEvents.ToolTipText = ""
    If (y > grdEvents.RowHeight(0)) And (y < grdEvents.RowHeight(0) + grdEvents.RowHeight(1)) Then
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdEvents, x, y, llRow, llCol)
    If (ilFound) And (llCol = EVENTIDINDEX) Then
        slStr = Trim$(grdEvents.TextMatrix(llRow, LIBNAMEINDEX))
        grdEvents.TextMatrix(llRow, LIBNAMEINDEX) = mGetLibName(slStr)
        grdEvents.ToolTipText = Trim$(grdEvents.TextMatrix(llRow, LIBNAMEINDEX))
    ElseIf (ilFound) And (llCol = AIRDATEERRORINDEX) And (Trim$(grdEvents.TextMatrix(llRow, ROWTYPEINDEX)) = "AAE") Then
        grdEvents.ToolTipText = Trim$(grdEvents.TextMatrix(llRow, AIRDATEINDEX))
    ElseIf (ilFound) And (llCol = AUDIONAMEINDEX) And (Trim$(grdEvents.TextMatrix(llRow, ROWTYPEINDEX)) = "AAE") Then
        llAAE = Val(grdEvents.TextMatrix(llRow, TMCURRSEEINDEX))
        grdEvents.ToolTipText = Trim$(tmCurrAAE(llAAE).sSourceConflict) & Trim$(tmCurrAAE(llAAE).sSourceUnavail)
    Else
        grdEvents.ToolTipText = Trim$(grdEvents.TextMatrix(llRow, llCol))
    End If
End Sub

Private Sub grdEvents_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Determine if in header
    If y < grdEvents.RowHeight(0) Then
        mSortCol grdEvents.Col
        Exit Sub
    End If
    If (y > grdEvents.RowHeight(0)) And (y < grdEvents.RowHeight(0) + grdEvents.RowHeight(1)) Then
        mSortCol grdEvents.Col
        Exit Sub
    End If
End Sub

Private Sub imcKey_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lbcKey.Visible = True
End Sub

Private Sub imcKey_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    lbcKey.Visible = False
End Sub

'
'               Snapshot of the Schedule events
'         Generate list of schedule events from the grid on the  screen
'
Private Sub imcPrint_Click()
Dim ilRptDest As Integer            'disply, print, save as file
Dim slRptName As String
Dim slExportName As String
Dim slRptType As String
Dim llResult As Long
Dim ilExportType As Integer
Dim llGridRow As Long
Dim slStr As String
Dim llTime As Long
Dim llAirDate As Long
Dim slFilter As String              'filters selected by user
Dim ilLoop As Integer
Dim slOperator As String * 2        'operator for filter
Dim llSequence As Long              'sequence for sorting crystal records

    If bmPrinting Then
        Exit Sub
    End If
    bmPrinting = True
    igRptIndex = ASAIRCOMPARE_RPT
    igRptSource = vbModal
    slRptName = "AsAirCompare.rpt"      'concatenate the crystal report name plus extension

    slExportName = ""               'no export for now
    
    Set rstSchedRpt = New Recordset
    gGenerateRstSchedule     'generate the ddfs for report
    
    rstSchedRpt.Open
    'build the data definition (.ttx) file in the database path for crystal to access
    llResult = CreateFieldDefFile(rstSchedRpt, sgDBPath & "\SchedRpt.ttx", True)
    
    If smAirDate = "" Then
        bmPrinting = False
        MsgBox "Enter valid date"
        Exit Sub
    End If
    
    llAirDate = gDateValue(smAirDate)
    llSequence = 0
    'loop thru the ItemID grid to print whats shown on the screen
    For llGridRow = grdEvents.FixedRows To grdEvents.Rows - 1
        slStr = Trim$(grdEvents.TextMatrix(llGridRow, ROWSORTINDEX))
        If slStr = "" Then
            Exit For
        Else
            rstSchedRpt.AddNew
            'Sequence # to keep in same order as screen output
            llSequence = llSequence + 1
            rstSchedRpt.Fields("SequenceID") = llSequence
            'error codes
            rstSchedRpt.Fields("AirDateError") = grdEvents.TextMatrix(llGridRow, AIRDATEERRORINDEX)
            rstSchedRpt.Fields("AutoOffError") = grdEvents.TextMatrix(llGridRow, AUTOOFFINDEX)
            rstSchedRpt.Fields("DataError") = grdEvents.TextMatrix(llGridRow, DATAERRORINDEX)
            rstSchedRpt.Fields("ScheduleError") = grdEvents.TextMatrix(llGridRow, SCHEDULEERRORINDEX)
            rstSchedRpt.Fields("StartDateSort") = llAirDate             'schedule date
            rstSchedRpt.Fields("EventType") = Left(grdEvents.TextMatrix(llGridRow, EVENTTYPEINDEX), 1)   'program, spot, avail
            rstSchedRpt.Fields("Event ID") = grdEvents.TextMatrix(llGridRow, EVENTIDINDEX)               'event ID
            rstSchedRpt.Fields("EvBusName") = grdEvents.TextMatrix(llGridRow, BUSNAMEINDEX)              'Bus name
            rstSchedRpt.Fields("EvBusCtl") = grdEvents.TextMatrix(llGridRow, BUSCTRLINDEX)                'Bus Control index
            rstSchedRpt.Fields("EvStarttime") = grdEvents.TextMatrix(llGridRow, TIMEINDEX)               'Event start time
            slStr = grdEvents.TextMatrix(llGridRow, TIMEINDEX)
            llTime = gStrTimeInTenthToLong(slStr, False)                'convert the start time of event to long for sorting
            rstSchedRpt.Fields("EvStartTimeSort") = llTime
            rstSchedRpt.Fields("EvStartType") = grdEvents.TextMatrix(llGridRow, STARTTYPEINDEX)          'start type
            rstSchedRpt.Fields("EvFix") = grdEvents.TextMatrix(llGridRow, FIXEDINDEX)                    'Fixed type
            rstSchedRpt.Fields("EvEndType") = grdEvents.TextMatrix(llGridRow, ENDTYPEINDEX)              'end type
            rstSchedRpt.Fields("EvDur") = grdEvents.TextMatrix(llGridRow, DURATIONINDEX)                 'duration
            rstSchedRpt.Fields("EvMatType") = grdEvents.TextMatrix(llGridRow, MATERIALINDEX)             'material type
            rstSchedRpt.Fields("EvAudName1") = grdEvents.TextMatrix(llGridRow, AUDIONAMEINDEX)           'primary audio name
            rstSchedRpt.Fields("EvItem1") = grdEvents.TextMatrix(llGridRow, AUDIOITEMIDINDEX)            'primary audio item id
            'rstSchedRpt.Fields("EvISCI1") = grdEvents.TextMatrix(llGridRow, AUDIOISCIINDEX)            'primary audio item id
            rstSchedRpt.Fields("EvCtl1") = grdEvents.TextMatrix(llGridRow, AUDIOCTRLINDEX)               'primary audio control
            rstSchedRpt.Fields("EvAudName2") = grdEvents.TextMatrix(llGridRow, BACKUPNAMEINDEX)          'backup audio name
            rstSchedRpt.Fields("EvCtl2") = grdEvents.TextMatrix(llGridRow, BACKUPCTRLINDEX)              'back control char
            rstSchedRpt.Fields("EvAudName3") = grdEvents.TextMatrix(llGridRow, PROTNAMEINDEX)            'protection audio name
            rstSchedRpt.Fields("EvItem3") = grdEvents.TextMatrix(llGridRow, PROTITEMIDINDEX)             'protection item id
            'rstSchedRpt.Fields("EvISCI3") = grdEvents.TextMatrix(llGridRow, PROTISCIINDEX)             'protection item id
            rstSchedRpt.Fields("EvCtl3") = grdEvents.TextMatrix(llGridRow, PROTCTRLINDEX)                'protection control
            rstSchedRpt.Fields("EvRelay1") = grdEvents.TextMatrix(llGridRow, RELAY1INDEX)                'relay 1 of 2
            rstSchedRpt.Fields("EvRelay2") = grdEvents.TextMatrix(llGridRow, RELAY2INDEX)                'relay 2 of 2
            rstSchedRpt.Fields("EvFollow") = grdEvents.TextMatrix(llGridRow, FOLLOWINDEX)                'follow name
            rstSchedRpt.Fields("EvSilenceTime") = grdEvents.TextMatrix(llGridRow, SILENCETIMEINDEX)      'silence time
            rstSchedRpt.Fields("EvSilence1") = grdEvents.TextMatrix(llGridRow, SILENCE1INDEX)            'silence name 1 of 4
            rstSchedRpt.Fields("EvSilence2") = grdEvents.TextMatrix(llGridRow, SILENCE2INDEX)            'silence name 2 of 4
            rstSchedRpt.Fields("EvSilence3") = grdEvents.TextMatrix(llGridRow, SILENCE3INDEX)            'silence name 3 of 4
            rstSchedRpt.Fields("EvSilence4") = grdEvents.TextMatrix(llGridRow, SILENCE4INDEX)            'silence name 4 of 4
            rstSchedRpt.Fields("EvNetCue1") = grdEvents.TextMatrix(llGridRow, NETCUE1INDEX)              'netcue name 1 of 2
            rstSchedRpt.Fields("EvNetCue2") = grdEvents.TextMatrix(llGridRow, NETCUE2INDEX)              'netcue name 2 of 2
            rstSchedRpt.Fields("EvTitle1") = grdEvents.TextMatrix(llGridRow, TITLE1INDEX)               'title 1 of 2
            rstSchedRpt.Fields("EvTitle2") = grdEvents.TextMatrix(llGridRow, TITLE2INDEX)               'title 2 of 2
            
        End If
    Next llGridRow
    
    slFilter = ""
    For ilLoop = 0 To UBound(tgFilterValues) - 1
        If slFilter <> "" Then               'not first time
            slFilter = slFilter & ", "
        End If
        If tgFilterValues(ilLoop).iOperator = 1 Then
            slOperator = "="
        ElseIf tgFilterValues(ilLoop).iOperator = 2 Then
            slOperator = "<>"
         ElseIf tgFilterValues(ilLoop).iOperator = 3 Then
            slOperator = ">"
        ElseIf tgFilterValues(ilLoop).iOperator = 4 Then
            slOperator = "<"
        ElseIf tgFilterValues(ilLoop).iOperator = 5 Then
            slOperator = ">="
        ElseIf tgFilterValues(ilLoop).iOperator = 6 Then
            slOperator = "<="
        End If
        slFilter = slFilter & Trim$(tgFilterValues(ilLoop).sFieldName) & " " & slOperator & " " & Trim$(tgFilterValues(ilLoop).sValue)
    
    Next ilLoop
    'pass the description of filters selected
    If slFilter = "" Then
        slFilter = "Filters: all events"
    Else
        slFilter = "Filters: " & slFilter
    End If
    
    sgCrystlFormula2 = "'" & Format$(llAirDate, "ddddd") & "'"         'air date for heading
    sgCrystlFormula3 = "'" & Trim$(slFilter) & "'"             'filter selected for heading

    gObtainReportforCrystal slRptName, slExportName     'determine which .rpt to call and setup an export name is user selected output to export
    igRptSource = vbModal
        
    EngrSchedRpt.Show vbModal
    
    'determine how the user responded, either cancel or produce output
    If igReturnCallStatus = CALLDONE Then           'produce the report flag
    
        slExportName = sgReturnCallName     'if exporting path and filename, this is filename; otherwise blank
        slRptType = ""
        'determine which version (condensed # of fields or all fields)
        If sgReturnOption = "ALL" Then
            slRptName = "AsAirCompareAll.rpt"
            EngrCrystal.gActiveCrystalReports igExportType, igRptDest, Trim$(slRptName) & Trim$(slRptType), slExportName, rstSchedRpt
        Else
            slRptName = "AsAirCompare.rpt"
            EngrCrystal.gActiveCrystalReports igExportType, igRptDest, Trim$(slRptName) & Trim$(slRptType), slExportName, rstSchedRpt
        End If
    End If
    Screen.MousePointer = vbDefault
    Set rstSchedRpt = Nothing
    cmcCancel.SetFocus
    bmPrinting = False
    Exit Sub
End Sub












Private Sub mMoveSEERecToCtrls()
    Dim llRow As Long
    Dim slStr As String
    Dim ilDNE As Integer
    Dim ilDSE As Integer
    Dim ilEBE As Integer
    Dim ilBDE As Integer
    Dim ilCCE As Integer
    Dim ilETE As Integer
    Dim ilTTE As Integer
    Dim ilMTE As Integer
    Dim ilASE As Integer
    Dim ilANE As Integer
    Dim ilRNE As Integer
    Dim ilFNE As Integer
    Dim ilSCE As Integer
    Dim ilNNE As Integer
    Dim llCTE As Long
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim slHours As String
    Dim llRet As Long
    Dim ilRowOk As Integer
    Dim slCategory As String
    Dim llAvailLength As Long
    Dim llTest As Long
    Dim llAirDate As Long
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim ilCol As Integer
    Dim llChg As Long
    Dim llARE As Long
    Dim llAAE As Long
    
    grdEvents.Rows = 4
    mClearControls
    If smAirDate = "" Then
        Exit Sub
    End If
    llAirDate = gDateValue(smAirDate)
    slDateTime = gNow()
    slNowDate = Format(slDateTime, "ddddd")
    slNowTime = Format(slDateTime, "ttttt")
    llNowDate = gDateValue(slNowDate)
    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
    ReDim imAAEMatchFound(0 To UBound(tmCurrAAE)) As Integer
    For llLoop = 0 To UBound(imAAEMatchFound) Step 1
        imAAEMatchFound(llLoop) = False
    Next llLoop
    llRow = grdEvents.FixedRows
    grdEvents.Rows = 2 * UBound(tmCurrSEE)
    grdEvents.Redraw = False
    For llLoop = 0 To UBound(tmCurrSEE) - 1 Step 1
        slCategory = ""
        ilETE = gBinarySearchETE(tmCurrSEE(llLoop).iEteCode, tgCurrETE)
        If ilETE <> -1 Then
            slCategory = tgCurrETE(ilETE).sCategory
        End If
        'Bypass Avails
        ilRowOk = True
        If slCategory = "A" Then
            ilRowOk = False
        End If
        'Add Criteria test here
        If ilRowOk Then
            ilRowOk = mCheckFilter(tmCurrSEE(llLoop), smT1Comment(llLoop))
        End If
        If ilRowOk Then
            If llRow + 1 > grdEvents.Rows Then
                grdEvents.AddItem ""
            End If
            grdEvents.Row = llRow
            For ilCol = AIRDATEERRORINDEX To TITLE2INDEX Step 1
                grdEvents.Col = ilCol
                grdEvents.CellBackColor = LIGHTBLUE
            Next ilCol
            grdEvents.TextMatrix(llRow, TMCURRSEEINDEX) = Trim$(Str$(llLoop))
            If tmCurrSEE(llLoop).lEventID > 0 Then
                grdEvents.TextMatrix(llRow, EVENTIDINDEX) = Trim$(Str$(tmCurrSEE(llLoop).lEventID))
            Else
                grdEvents.TextMatrix(llRow, EVENTIDINDEX) = ""
            End If
            slStr = ""
            If tmCurrSEE(llLoop).lDeeCode > 0 Then
                slStr = "DEE=" & tmCurrSEE(llLoop).lDeeCode
                If tmCurrSEE(llLoop).iEteCode = imSpotETECode Then
                    slStr = slStr & "/" & gLongToStrTimeInTenth(tmCurrSEE(llLoop).lTime)
                End If
            End If
            grdEvents.TextMatrix(llRow, LIBNAMEINDEX) = slStr
            slStr = ""
            ilBDE = gBinarySearchBDE(tmCurrSEE(llLoop).iBdeCode, tgCurrBDE)
            If ilBDE <> -1 Then
                slStr = slStr & Trim$(tgCurrBDE(ilBDE).sName)
            End If
            grdEvents.TextMatrix(llRow, BUSNAMEINDEX) = slStr
            grdEvents.TextMatrix(llRow, BUSCTRLINDEX) = ""
            ilCCE = gBinarySearchCCE(tmCurrSEE(llLoop).iBusCceCode, tgCurrBusCCE)
            If ilCCE <> -1 Then
                grdEvents.TextMatrix(llRow, BUSCTRLINDEX) = Trim$(tgCurrBusCCE(ilCCE).sAutoChar)
            End If
            grdEvents.TextMatrix(llRow, EVENTTYPEINDEX) = ""
            ilETE = gBinarySearchETE(tmCurrSEE(llLoop).iEteCode, tgCurrETE)
            If ilETE <> -1 Then
                'Auto Code sent and returned
                grdEvents.TextMatrix(llRow, EVENTTYPEINDEX) = Trim$(tgCurrETE(ilETE).sAutoCodeChar) 'Trim$(tgCurrETE(ilETE).sName)
            End If
            If tmCurrSEE(llLoop).iEteCode <> imSpotETECode Then
                If slCategory = "A" Then
                    grdEvents.TextMatrix(llRow, TIMEINDEX) = gLongToStrTimeInTenth(tmCurrSEE(llLoop).lSpotTime)
                Else
                    grdEvents.TextMatrix(llRow, TIMEINDEX) = gLongToStrTimeInTenth(tmCurrSEE(llLoop).lTime)
                End If
            Else
                grdEvents.Col = TIMEINDEX
                grdEvents.TextMatrix(llRow, TIMEINDEX) = gLongToStrTimeInTenth(tmCurrSEE(llLoop).lSpotTime)
            End If
            grdEvents.TextMatrix(llRow, STARTTYPEINDEX) = ""
            ilTTE = gBinarySearchTTE(tmCurrSEE(llLoop).iStartTteCode, tgCurrStartTTE)
            If ilTTE <> -1 Then
                grdEvents.TextMatrix(llRow, STARTTYPEINDEX) = Trim$(tgCurrStartTTE(ilTTE).sName)
            End If
            If Trim$(tmCurrSEE(llLoop).sFixedTime) = "Y" Then
                grdEvents.TextMatrix(llRow, FIXEDINDEX) = Trim$(tgAEE.sFixedTimeChar)
            End If
            grdEvents.TextMatrix(llRow, ENDTYPEINDEX) = ""
            ilTTE = gBinarySearchTTE(tmCurrSEE(llLoop).iEndTteCode, tgCurrEndTTE)
            If ilTTE <> -1 Then
                grdEvents.TextMatrix(llRow, ENDTYPEINDEX) = Trim$(tgCurrEndTTE(ilTTE).sName)
            End If
            If slCategory = "A" Then
                grdEvents.TextMatrix(llRow, DURATIONINDEX) = gLongToStrLengthInTenth(llAvailLength, False)
            Else
                If (tmCurrSEE(llLoop).lDuration > 0) Then
                    grdEvents.TextMatrix(llRow, DURATIONINDEX) = gLongToStrLengthInTenth(tmCurrSEE(llLoop).lDuration, True)
                Else
                    grdEvents.TextMatrix(llRow, DURATIONINDEX) = gLongToStrLengthInTenth(tmCurrSEE(llLoop).lDuration, True)    '""
                End If
            End If
            If tmCurrSEE(llLoop).iEteCode = imSpotETECode Then
                grdEvents.Col = DURATIONINDEX
            End If
            grdEvents.TextMatrix(llRow, MATERIALINDEX) = ""
            ilMTE = gBinarySearchMTE(tmCurrSEE(llLoop).iMteCode, tgCurrMTE)
            If ilMTE <> -1 Then
                grdEvents.TextMatrix(llRow, MATERIALINDEX) = Trim$(tgCurrMTE(ilMTE).sName)
            End If
            grdEvents.TextMatrix(llRow, AUDIONAMEINDEX) = ""
            ilASE = gBinarySearchASE(tmCurrSEE(llLoop).iAudioAseCode, tgCurrASE())
            If ilASE <> -1 Then
                ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                If ilANE <> -1 Then
                    grdEvents.TextMatrix(llRow, AUDIONAMEINDEX) = Trim$(tgCurrANE(ilANE).sName)
                End If
            End If
            grdEvents.TextMatrix(llRow, AUDIOITEMIDINDEX) = Trim$(tmCurrSEE(llLoop).sAudioItemID)
            grdEvents.TextMatrix(llRow, AUDIOISCIINDEX) = Trim$(tmCurrSEE(llLoop).sAudioISCI)
            grdEvents.TextMatrix(llRow, AUDIOCTRLINDEX) = ""
            ilCCE = gBinarySearchCCE(tmCurrSEE(llLoop).iAudioCceCode, tgCurrAudioCCE)
            If ilCCE <> -1 Then
                grdEvents.TextMatrix(llRow, AUDIOCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
            End If
            ilANE = gBinarySearchANE(tmCurrSEE(llLoop).iBkupAneCode, tgCurrANE())
            If ilANE <> -1 Then
                grdEvents.TextMatrix(llRow, BACKUPNAMEINDEX) = Trim$(tgCurrANE(ilANE).sName)
            End If
            grdEvents.TextMatrix(llRow, BACKUPCTRLINDEX) = ""
            ilCCE = gBinarySearchCCE(tmCurrSEE(llLoop).iBkupCceCode, tgCurrAudioCCE)
            If ilCCE <> -1 Then
                grdEvents.TextMatrix(llRow, BACKUPCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
            End If
            grdEvents.TextMatrix(llRow, PROTNAMEINDEX) = ""
            ilANE = gBinarySearchANE(tmCurrSEE(llLoop).iProtAneCode, tgCurrANE())
            If ilANE <> -1 Then
                grdEvents.TextMatrix(llRow, PROTNAMEINDEX) = Trim$(tgCurrANE(ilANE).sName)
            End If
            grdEvents.TextMatrix(llRow, PROTITEMIDINDEX) = Trim$(tmCurrSEE(llLoop).sProtItemID)
            grdEvents.TextMatrix(llRow, PROTISCIINDEX) = Trim$(tmCurrSEE(llLoop).sProtISCI)
            grdEvents.TextMatrix(llRow, PROTCTRLINDEX) = ""
            ilCCE = gBinarySearchCCE(tmCurrSEE(llLoop).iProtCceCode, tgCurrAudioCCE)
            If ilCCE <> -1 Then
                grdEvents.TextMatrix(llRow, PROTCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
            End If
            grdEvents.TextMatrix(llRow, RELAY1INDEX) = ""
            ilRNE = gBinarySearchRNE(tmCurrSEE(llLoop).i1RneCode, tgCurrRNE)
            If ilRNE <> -1 Then
                grdEvents.TextMatrix(llRow, RELAY1INDEX) = Trim$(tgCurrRNE(ilRNE).sName)
            End If
            grdEvents.TextMatrix(llRow, RELAY2INDEX) = ""
            ilRNE = gBinarySearchRNE(tmCurrSEE(llLoop).i2RneCode, tgCurrRNE)
            If ilRNE <> -1 Then
                grdEvents.TextMatrix(llRow, RELAY2INDEX) = Trim$(tgCurrRNE(ilRNE).sName)
            End If
            grdEvents.TextMatrix(llRow, FOLLOWINDEX) = ""
            ilFNE = gBinarySearchFNE(tmCurrSEE(llLoop).iFneCode, tgCurrFNE)
            If ilFNE <> -1 Then
                grdEvents.TextMatrix(llRow, FOLLOWINDEX) = Trim$(tgCurrFNE(ilFNE).sName)
            End If
            If tmCurrSEE(llLoop).lSilenceTime > 0 Then
                grdEvents.TextMatrix(llRow, SILENCETIMEINDEX) = gLongToLength(tmCurrSEE(llLoop).lSilenceTime, False)   'gLongToStrLengthInTenth(tmCurrSEE(llLoop).lSilenceTime, False)
            Else
                grdEvents.TextMatrix(llRow, SILENCETIMEINDEX) = ""
            End If
            grdEvents.TextMatrix(llRow, SILENCE1INDEX) = ""
            ilSCE = gBinarySearchSCE(tmCurrSEE(llLoop).i1SceCode, tgCurrSCE)
            If ilSCE <> -1 Then
                grdEvents.TextMatrix(llRow, SILENCE1INDEX) = Trim$(tgCurrSCE(ilSCE).sAutoChar)
            End If
            grdEvents.TextMatrix(llRow, SILENCE2INDEX) = ""
            ilSCE = gBinarySearchSCE(tmCurrSEE(llLoop).i2SceCode, tgCurrSCE)
            If ilSCE <> -1 Then
                grdEvents.TextMatrix(llRow, SILENCE2INDEX) = Trim$(tgCurrSCE(ilSCE).sAutoChar)
            End If
            grdEvents.TextMatrix(llRow, SILENCE3INDEX) = ""
            ilSCE = gBinarySearchSCE(tmCurrSEE(llLoop).i3SceCode, tgCurrSCE)
            If ilSCE <> -1 Then
                grdEvents.TextMatrix(llRow, SILENCE3INDEX) = Trim$(tgCurrSCE(ilSCE).sAutoChar)
            End If
            grdEvents.TextMatrix(llRow, SILENCE4INDEX) = ""
            ilSCE = gBinarySearchSCE(tmCurrSEE(llLoop).i4SceCode, tgCurrSCE)
            If ilSCE <> -1 Then
                grdEvents.TextMatrix(llRow, SILENCE4INDEX) = Trim$(tgCurrSCE(ilSCE).sAutoChar)
            End If
            grdEvents.TextMatrix(llRow, NETCUE1INDEX) = ""
            ilNNE = gBinarySearchNNE(tmCurrSEE(llLoop).iStartNneCode, tgCurrNNE)
            If ilNNE <> -1 Then
                grdEvents.TextMatrix(llRow, NETCUE1INDEX) = Trim$(tgCurrNNE(ilNNE).sName)
            End If
            grdEvents.TextMatrix(llRow, NETCUE2INDEX) = ""
            ilNNE = gBinarySearchNNE(tmCurrSEE(llLoop).iEndNneCode, tgCurrNNE)
            If ilNNE <> -1 Then
                grdEvents.TextMatrix(llRow, NETCUE2INDEX) = Trim$(tgCurrNNE(ilNNE).sName)
            End If
            If tmCurrSEE(llLoop).iEteCode <> imSpotETECode Then
                grdEvents.TextMatrix(llRow, TITLE1INDEX) = smT1Comment(llLoop)
            Else
                llARE = gBinarySearchARE(tmCurrSEE(llLoop).lAreCode, tgCurrARE())
                If llARE <> -1 Then
                    grdEvents.TextMatrix(llRow, TITLE1INDEX) = Trim$(tgCurrARE(llARE).sName)
                End If
            End If
            '7/8/11: Make T2 work like T1
            'grdEvents.TextMatrix(llRow, TITLE2INDEX) = ""
            'llCTE = gBinarySearchCTE(tmCurrSEE(llLoop).l2CteCode, tgCurrCTE)
            'If llCTE <> -1 Then
            '    grdEvents.TextMatrix(llRow, TITLE2INDEX) = Trim$(tgCurrCTE(llCTE).sName)
            'End If
            grdEvents.TextMatrix(llRow, TITLE2INDEX) = smT2Comment(llLoop)
            grdEvents.TextMatrix(llRow, PCODEINDEX) = tmCurrSEE(llLoop).lCode
            slStr = Trim$(Str$(llRow))
            Do While Len(slStr) < 6
                slStr = "0" & slStr
            Loop
            grdEvents.TextMatrix(llRow, ROWSORTINDEX) = slStr
            grdEvents.TextMatrix(llRow, ROWTYPEINDEX) = "SEE"
            llRow = llRow + 1
            If llRow + 1 > grdEvents.Rows Then
                grdEvents.AddItem ""
            End If
            grdEvents.Row = llRow
            For ilCol = AIRDATEERRORINDEX To TITLE2INDEX Step 1
                grdEvents.Col = ilCol
                grdEvents.CellBackColor = LIGHTYELLOW    'LIGHTGREEN 'LIGHTRED
            Next ilCol
            'Find AAE match
            llAAE = gBinarySearchAAEbyEventID(tmCurrSEE(llLoop).lEventID, tmCurrAAE())
            If llAAE <> -1 Then
                imAAEMatchFound(llAAE) = True
                mDisplayAAE llAAE, llRow
                If Not mIsDiscrepant(llRow - 1, tmCurrAAE(llAAE)) Then
                    If ckcShow.Value = vbChecked Then
                        llRow = llRow - 1
                    End If
                Else
                    grdEvents.TextMatrix(llRow - 1, DISCREPANCYINDEX) = "Y"
                    grdEvents.TextMatrix(llRow, DISCREPANCYINDEX) = "Y"
                End If
            Else
                grdEvents.TextMatrix(llRow, EVENTIDINDEX) = "Missing"
                grdEvents.Row = llRow
                grdEvents.Col = EVENTIDINDEX
                grdEvents.CellForeColor = vbRed
                slStr = Trim$(Str$(llRow))
                Do While Len(slStr) < 6
                    slStr = "0" & slStr
                Loop
                grdEvents.TextMatrix(llRow, ROWSORTINDEX) = slStr
                grdEvents.TextMatrix(llRow, ROWTYPEINDEX) = "AAE"
                grdEvents.TextMatrix(llRow - 1, DISCREPANCYINDEX) = "Y"
                grdEvents.TextMatrix(llRow, DISCREPANCYINDEX) = "Y"
            End If
            llRow = llRow + 1
        End If
    Next llLoop
    For llAAE = 0 To UBound(imAAEMatchFound) - 1 Step 1
        If imAAEMatchFound(llAAE) = False Then
            'Add Rows
            If llRow + 1 > grdEvents.Rows Then
                grdEvents.AddItem ""
            End If
            grdEvents.Row = llRow
            For ilCol = AIRDATEERRORINDEX To TITLE2INDEX Step 1
                grdEvents.Col = ilCol
                grdEvents.CellBackColor = LIGHTBLUE
            Next ilCol
            grdEvents.TextMatrix(llRow, EVENTIDINDEX) = "Missing"
            grdEvents.Row = llRow
            grdEvents.Col = EVENTIDINDEX
            grdEvents.CellForeColor = vbRed
            slStr = Trim$(Str$(llRow))
            Do While Len(slStr) < 6
                slStr = "0" & slStr
            Loop
            grdEvents.TextMatrix(llRow, ROWSORTINDEX) = slStr
            grdEvents.TextMatrix(llRow, ROWTYPEINDEX) = "SEE"
            llRow = llRow + 1
            If llRow + 1 > grdEvents.Rows Then
                grdEvents.AddItem ""
            End If
            grdEvents.Row = llRow
            For ilCol = AIRDATEERRORINDEX To TITLE2INDEX Step 1
                grdEvents.Col = ilCol
                grdEvents.CellBackColor = LIGHTYELLOW    'LIGHTGREEN 'LIGHTRED
            Next ilCol
            mDisplayAAE llAAE, llRow
            grdEvents.TextMatrix(llRow - 1, DISCREPANCYINDEX) = "Y"
            grdEvents.TextMatrix(llRow, DISCREPANCYINDEX) = "Y"
            llRow = llRow + 1
        End If
    Next llAAE
    'If llRow >= grdEvents.Rows Then
    '    grdEvents.AddItem ""
    'End If
    grdEvents.Rows = llRow
    If llAirDate < llNowDate Then
        For llRow = grdEvents.FixedRows To grdEvents.Rows - 1 Step 1
            grdEvents.Row = llRow
            grdEvents.Col = EVENTIDINDEX
            grdEvents.CellAlignment = flexAlignRightCenter
        Next llRow
    Else
        For llRow = grdEvents.FixedRows To grdEvents.Rows - 1 Step 1
            grdEvents.Row = llRow
            grdEvents.Col = EVENTIDINDEX
            grdEvents.CellAlignment = flexAlignRightCenter
        Next llRow
    End If
    grdEvents.Redraw = True
End Sub



Private Sub mPopDNE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_DNE_DayName("C", "L", sgCurrLibDNEStamp, "EngrAsAirCompare-mPopulate Library Names", tgCurrLibDNE())
    ilRet = gGetTypeOfRecs_DNE_DayName("C", "T", sgCurrTempDNEStamp, "EngrAsAirCompare-mPopulate Template Names", tgCurrTempDNE())
End Sub

Private Sub mPopDSE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_DSE_DaySubName("C", sgCurrDSEStamp, "EngrAsAirCompare-mPopDSE Day Subname", tgCurrDSE())
End Sub


Private Sub mPopBDE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrAsAirCompare-mPopBDE Bus Definition", tgCurrBDE())
End Sub





Private Sub mPopCCE_Audio()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "A", sgCurrAudioCCEStamp, "EngrAsAirCompare-mPopCCE_Audio Control Character", tgCurrAudioCCE())
End Sub

Private Sub mPopCCE_Bus()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "B", sgCurrBusCCEStamp, "EngrAsAirCompare-mPopCCE_Bus Control Character", tgCurrBusCCE())
End Sub

Private Sub mPopTTE_StartType()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_TTE_TimeType("C", "S", sgCurrStartTTEStamp, "EngrAsAirCompare-mPopTTE_StartType Start Type", tgCurrStartTTE())
End Sub

Private Sub mPopTTE_EndType()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_TTE_TimeType("C", "E", sgCurrEndTTEStamp, "EngrAsAirCompare-mPopTTE_EndType End Type", tgCurrEndTTE())
End Sub

Private Sub mPopASE()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilANE As Integer

    mPopANE
    ilRet = gGetTypeOfRecs_ASE_AudioSource("C", sgCurrASEStamp, "EngrAsAirCompare-mPopASE Audio Source", tgCurrASE())
End Sub

Private Sub mPopSCE()

    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_SCE_SilenceChar("C", sgCurrSCEStamp, "EngrAsAirCompare-mPopSCE Silence Character", tgCurrSCE())
End Sub

Private Sub mPopNNE()

    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_NNE_NetcueName("C", sgCurrNNEStamp, "EngrAsAirCompare-mPopNNE Netcue", tgCurrNNE())
End Sub

Private Sub mPopCTE()

    '7/8/11: Make T2 work like T1
    'Dim ilRet As Integer
    'Dim ilLoop As Integer

    'ilRet = gGetTypeOfRecs_CTE_CommtsTitle("C", "T2", sgCurrCTEStamp, "EngrAsAirCompare-mPopCTE Title 2", tgCurrCTE())
End Sub

Private Sub mPopANE()

    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrAsAirCompare-mPopANE Audio Audio Names", tgCurrANE())
End Sub

Private Sub mPopARE()

    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetRecs_ARE_AdvertiserRefer(sgCurrAREStamp, "EngrAsAirCompare-mPopARE Advertiser Names", tgCurrARE())
End Sub

Private Sub mPopETE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_ETE_EventType("C", sgCurrETEStamp, "EngrLibETE-mPopETE Event Types", tgCurrETE())
    ilRet = gGetTypeOfRecs_EPE_EventProperties("C", sgCurrEPEStamp, "EngrAsAirCompare-mPopETE Event Properties", tgCurrEPE())
End Sub

Private Sub mPopMTE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_MTE_MaterialType("C", sgCurrMTEStamp, "EngrAsAirCompare-mPopMTE Material Type", tgCurrMTE())
End Sub

Private Sub mPopRNE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_RNE_RelayName("C", sgCurrRNEStamp, "EngrAsAirCompare-mPopRNE Relay", tgCurrRNE())
End Sub

Private Sub mPopFNE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_FNE_FollowName("C", sgCurrFNEStamp, "EngrAsAirCompare-mPopFNE Follow", tgCurrFNE())
End Sub











Private Function mComputeWidth(ilPass As Integer, CtrlWidth As Single, ilAdjValue As Integer, slUsedFlag As String) As Single
    If ilPass = 0 Then
        CtrlWidth = grdEvents.Width / ilAdjValue
        If slUsedFlag <> "Y" Then
            imUnusedCount = imUnusedCount + 1
            fmUnusedWidth = fmUnusedWidth + CtrlWidth
            CtrlWidth = 0
        Else
            fmUsedWidth = fmUsedWidth + CtrlWidth
        End If
    Else
        CtrlWidth = CtrlWidth + ((fmUnusedWidth * CtrlWidth) / fmUsedWidth)
    End If
    mComputeWidth = CtrlWidth
End Function
























Private Sub mInitFilterInfo()
    Dim ilUpper As Integer
    ReDim tgFilterFields(0 To 0) As FIELDSELECTION
    
    ilUpper = 0
    If (UBound(tgUsedATE) > 0) And ((tgSchUsedSumEPE.sAudioName <> "N") Or (tgSchUsedSumEPE.sProtAudioName <> "N") Or (tgSchUsedSumEPE.sBkupAudioName <> "N")) Then
        tgFilterFields(ilUpper).sFieldName = "Audio Types"
        tgFilterFields(ilUpper).iFieldType = 5
        If Len(tgATE.sName) >= 6 Then
            tgFilterFields(ilUpper).iMaxNoChar = Len(tgATE.sName)
        Else
            tgFilterFields(ilUpper).iMaxNoChar = 6
        End If
        tgFilterFields(ilUpper).sListFile = "ATE"
        tgFilterFields(ilUpper).sMandatory = "Y"
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedANE) > 0) And ((tgSchUsedSumEPE.sAudioName <> "N") Or (tgSchUsedSumEPE.sProtAudioName <> "N") Or (tgSchUsedSumEPE.sBkupAudioName <> "N")) Then
        tgFilterFields(ilUpper).sFieldName = "Audio Name"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("AudioName", 6)
        tgFilterFields(ilUpper).sListFile = "ANE"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sAudioName
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedBDE) > 0) And (tgSchUsedSumEPE.sBus <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "Bus"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("BusName", 6)
        tgFilterFields(ilUpper).sListFile = "BDE"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sBus
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If UBound(tgUsedETE) > 0 Then
        tgFilterFields(ilUpper).sFieldName = "Event Types"
        tgFilterFields(ilUpper).iFieldType = 5
        If Len(tgETE.sName) >= 6 Then
            tgFilterFields(ilUpper).iMaxNoChar = Len(tgETE.sName)
        Else
            tgFilterFields(ilUpper).iMaxNoChar = 6
        End If
        tgFilterFields(ilUpper).sListFile = "ETE"
        tgFilterFields(ilUpper).sMandatory = "Y"
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedFNE) > 0) And (tgSchUsedSumEPE.sFollow <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "Follow"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Follow", 6)
        tgFilterFields(ilUpper).sListFile = "FNE"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sFollow
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedMTE) > 0) And (tgSchUsedSumEPE.sMaterialType <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "Material"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Material", 6)
        tgFilterFields(ilUpper).sListFile = "MTE"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sMaterialType
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedNNE) > 0) And ((tgSchUsedSumEPE.sStartNetcue <> "N") Or (tgSchUsedSumEPE.sStopNetcue <> "N")) Then
        tgFilterFields(ilUpper).sFieldName = "Netcue"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Netcue1", 6)
        tgFilterFields(ilUpper).sListFile = "NNE"
        If (tgSchManSumEPE.sStartNetcue = "Y") Or (tgSchManSumEPE.sStopNetcue = "Y") Then
            tgFilterFields(ilUpper).sMandatory = "Y"
        Else
            tgFilterFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedRNE) > 0) And ((tgSchUsedSumEPE.sRelay1 <> "N") Or (tgSchUsedSumEPE.sRelay2 <> "N")) Then
        tgFilterFields(ilUpper).sFieldName = "Relay"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Relay1", 6)
        tgFilterFields(ilUpper).sListFile = "RNE"
        If (tgSchManSumEPE.sRelay1 = "Y") Or (tgSchManSumEPE.sRelay2 = "Y") Then
            tgFilterFields(ilUpper).sMandatory = "Y"
        Else
            tgFilterFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedStartTTE) > 0) And (tgSchUsedSumEPE.sStartType <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "Start Type"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("StartType", 6)
        tgFilterFields(ilUpper).sListFile = "TTES"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sStartType
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedEndTTE) > 0) And (tgSchUsedSumEPE.sEndType <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "End Type"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("EndType", 6)
        tgFilterFields(ilUpper).sListFile = "TTEE"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sEndType
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedAudioCCE) > 0) And ((tgSchUsedSumEPE.sAudioControl <> "N") Or (tgSchUsedSumEPE.sProtAudioControl <> "N") Or (tgSchUsedSumEPE.sBkupAudioControl <> "N")) Then
        tgFilterFields(ilUpper).sFieldName = "Audio Control"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("AudioCtrl", 6)
        tgFilterFields(ilUpper).sListFile = "CCEA"
        If (tgSchManSumEPE.sAudioControl = "Y") Or (tgSchManSumEPE.sProtAudioControl = "Y") Or (tgSchManSumEPE.sBkupAudioControl = "Y") Then
            tgFilterFields(ilUpper).sMandatory = "Y"
        Else
            tgFilterFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedBusCCE) > 0) And (tgSchUsedSumEPE.sBusControl <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "Bus Control"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("BusCtrl", 6)
        tgFilterFields(ilUpper).sListFile = "CCEB"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sBusControl
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    '7/8/11: Make T2 work like T1
    'If (UBound(tgUsedT2CTE) > 0) And (tgSchUsedSumEPE.sTitle2 <> "N") Then
    '    tgFilterFields(ilUpper).sFieldName = "Title 2"
    '    tgFilterFields(ilUpper).iFieldType = 5
    '    tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Title2", 6)
    '    tgFilterFields(ilUpper).sListFile = "CTE2"
    '    tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sTitle2
    '    ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
    '    ilUpper = ilUpper + 1
    'End If
    If (UBound(tgT2MatchList) > 0) And (tgSchUsedSumEPE.sTitle2 <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "Title 2"
        tgFilterFields(ilUpper).iFieldType = 9
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Title2", 6)
        tgFilterFields(ilUpper).sListFile = "CTE2"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sTitle2
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgT1MatchList) > 0) And (tgSchUsedSumEPE.sTitle1 <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "Title 1"
        tgFilterFields(ilUpper).iFieldType = 9
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Title1", 6)
        tgFilterFields(ilUpper).sListFile = "CTE1"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sTitle1
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedSCE) > 0) And ((tgSchUsedSumEPE.sSilence1 <> "N") Or (tgSchUsedSumEPE.sSilence2 <> "N") Or (tgSchUsedSumEPE.sSilence3 <> "N") Or (tgSchUsedSumEPE.sSilence4 <> "N")) Then
        tgFilterFields(ilUpper).sFieldName = "Silence Control"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Silence1", 6)
        tgFilterFields(ilUpper).sListFile = "SCE"
        If (tgSchManSumEPE.sSilence1 = "Y") Or (tgSchManSumEPE.sSilence2 = "Y") Or (tgSchManSumEPE.sSilence3 = "Y") Or (tgSchManSumEPE.sSilence4 = "Y") Then
            tgFilterFields(ilUpper).sMandatory = "Y"
        Else
            tgFilterFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If tgSchUsedSumEPE.sFixedTime <> "N" Then
        tgFilterFields(ilUpper).sFieldName = "Fixed Time"
        tgFilterFields(ilUpper).iFieldType = 9
        tgFilterFields(ilUpper).iMaxNoChar = 1
        tgFilterFields(ilUpper).sListFile = "FTYN"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sFixedTime
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If tgSchUsedSumEPE.sTime <> "N" Then
        tgFilterFields(ilUpper).sFieldName = "Time"
        tgFilterFields(ilUpper).iFieldType = 6
        tgFilterFields(ilUpper).iMaxNoChar = 10
        tgFilterFields(ilUpper).sListFile = ""
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sTime
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If tgSchUsedSumEPE.sDuration <> "N" Then
        tgFilterFields(ilUpper).sFieldName = "Duration"
        tgFilterFields(ilUpper).iFieldType = 8
        tgFilterFields(ilUpper).iMaxNoChar = 10
        tgFilterFields(ilUpper).sListFile = ""
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sDuration
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedANE) > 0) And ((tgSchUsedSumEPE.sAudioItemID <> "N") Or (tgSchUsedSumEPE.sProtAudioItemID <> "N")) Then
        tgFilterFields(ilUpper).sFieldName = "Item ID"
        tgFilterFields(ilUpper).iFieldType = 2
        tgFilterFields(ilUpper).iMaxNoChar = 0
        tgFilterFields(ilUpper).sListFile = ""
        If (tgSchManSumEPE.sAudioItemID = "Y") Or (tgSchManSumEPE.sProtAudioItemID = "Y") Then
            tgFilterFields(ilUpper).sMandatory = "Y"
        Else
            tgFilterFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sAudioISCI <> "N") Or (tgSchUsedSumEPE.sProtAudioISCI <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "ISCI"
        tgFilterFields(ilUpper).iFieldType = 2
        tgFilterFields(ilUpper).iMaxNoChar = 0
        tgFilterFields(ilUpper).sListFile = ""
        If (tgSchManSumEPE.sAudioISCI = "Y") Or (tgSchManSumEPE.sProtAudioISCI = "Y") Then
            tgFilterFields(ilUpper).sMandatory = "Y"
        Else
            tgFilterFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If tgSchUsedSumEPE.sSilenceTime <> "N" Then
        tgFilterFields(ilUpper).sFieldName = "Silence Time"
        tgFilterFields(ilUpper).iFieldType = 8
        tgFilterFields(ilUpper).iMaxNoChar = 5
        tgFilterFields(ilUpper).sListFile = ""
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sSilenceTime
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
    End If
    If tmSHE.lCode <> 0 Then
        tgFilterFields(ilUpper).sFieldName = "Event ID"
        tgFilterFields(ilUpper).iFieldType = 1
        tgFilterFields(ilUpper).iMaxNoChar = 0
        tgFilterFields(ilUpper).sListFile = ""
        tgFilterFields(ilUpper).sMandatory = "Y"
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    
End Sub


Private Function mCheckFilter(tlCurrSEE As SEE, slComment As String) As Integer
    Dim ilFilter As Integer
    Dim ilField As Integer
    Dim ilFilterType As Integer
    Dim slFileName As String
    Dim ilOrTest As Integer
    Dim ilMatch As Integer
    Dim ilASE As Integer
    Dim ilANE As Integer
    
    mCheckFilter = True
    If cbcApplyFilter.Value = vbUnchecked Then
        Exit Function
    End If
    For ilFilter = LBound(tmFilterValues) To UBound(tmFilterValues) - 1 Step 1
        tmFilterValues(ilFilter).iUsed = False
    Next ilFilter
    For ilFilter = LBound(tmFilterValues) To UBound(tmFilterValues) - 1 Step 1
        If tmFilterValues(ilFilter).iUsed = False Then
            For ilField = LBound(tgFilterFields) To UBound(tgFilterFields) - 1 Step 1
                If tgFilterFields(ilField).sFieldName = tmFilterValues(ilFilter).sFieldName Then
                    ilFilterType = tgFilterFields(ilField).iFieldType
                    slFileName = tgFilterFields(ilField).sListFile
                    If ilFilterType = 5 Then
                        Select Case UCase$(Trim$(slFileName))
                            Case "ATE"
                                'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                                '    If tlCurrSEE.iAudioAseCode = tgCurrASE(ilASE).iCode Then
                                    ilASE = gBinarySearchASE(tlCurrSEE.iAudioAseCode, tgCurrASE())
                                    If ilASE <> -1 Then
                                        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                                        '    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                                            ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                                            If ilANE <> -1 Then
                                                ilMatch = mMatchTestFile(ilFilter, CLng(tgCurrANE(ilANE).iAteCode))
                                        '        Exit For
                                            End If
                                        'Next ilANE
                                '        Exit For
                                    End If
                                'Next ilASE
                                If Not ilMatch Then
                                    'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                                    '    If tlCurrSEE.iProtAneCode = tgCurrANE(ilANE).iCode Then
                                        ilANE = gBinarySearchANE(tlCurrSEE.iProtAneCode, tgCurrANE())
                                        If ilANE <> -1 Then
                                            ilMatch = mMatchTestFile(ilFilter, CLng(tgCurrANE(ilANE).iAteCode))
                                    '        Exit For
                                        End If
                                    'Next ilANE
                                End If
                                If Not ilMatch Then
                                    'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                                    '    If tlCurrSEE.iBkupAneCode = tgCurrANE(ilANE).iCode Then
                                        ilANE = gBinarySearchANE(tlCurrSEE.iBkupAneCode, tgCurrANE())
                                        If ilANE <> -1 Then
                                            ilMatch = mMatchTestFile(ilFilter, CLng(tgCurrANE(ilANE).iAteCode))
                                    '        Exit For
                                        End If
                                    'Next ilANE
                                End If
                                If Not ilMatch Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                                
                            Case "ANE"
                                'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                                '    If tlCurrSEE.iAudioAseCode = tgCurrASE(ilASE).iCode Then
                                    ilASE = gBinarySearchASE(tlCurrSEE.iAudioAseCode, tgCurrASE())
                                    If ilASE <> -1 Then
                                        ilMatch = mMatchTestFile(ilFilter, CLng(tgCurrASE(ilASE).iPriAneCode))
                                '        Exit For
                                    End If
                                'Next ilASE
                                If Not ilMatch Then
                                    ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.iProtAneCode))
                                End If
                                If Not ilMatch Then
                                    ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.iBkupAneCode))
                                End If
                                If Not ilMatch Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "BDE"
                                If Not mMatchTestFile(ilFilter, CLng(tlCurrSEE.iBdeCode)) Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "ETE"
                                If Not mMatchTestFile(ilFilter, CLng(tlCurrSEE.iEteCode)) Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "FNE"
                                If Not mMatchTestFile(ilFilter, CLng(tlCurrSEE.iFneCode)) Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "MTE"
                                If Not mMatchTestFile(ilFilter, CLng(tlCurrSEE.iMteCode)) Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "NNE"
                                ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.iStartNneCode))
                                If Not ilMatch Then
                                    ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.iEndNneCode))
                                End If
                                If Not ilMatch Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "RNE"
                                ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.i1RneCode))
                                If Not ilMatch Then
                                    ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.i2RneCode))
                                End If
                                If Not ilMatch Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "TTES"
                                If Not mMatchTestFile(ilFilter, CLng(tlCurrSEE.iStartTteCode)) Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "TTEE"
                                If Not mMatchTestFile(ilFilter, CLng(tlCurrSEE.iEndTteCode)) Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "CCEA"
                                ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.iAudioCceCode))
                                If Not ilMatch Then
                                    ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.iBkupCceCode))
                                    If Not ilMatch Then
                                        ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.iProtCceCode))
                                    End If
                                End If
                                If Not ilMatch Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "CCEB"
                                If Not mMatchTestFile(ilFilter, CLng(tlCurrSEE.iBusCceCode)) Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "SCE"
                                ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.i1SceCode))
                                If Not ilMatch Then
                                    ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.i2SceCode))
                                    If Not ilMatch Then
                                        ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.i3SceCode))
                                        If Not ilMatch Then
                                            ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.i4SceCode))
                                        End If
                                    End If
                                End If
                                If Not ilMatch Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "FTYN"
                                If Trim$(tlCurrSEE.sFixedTime) = "Y" Then
                                    If Not mMatchTestFile(ilFilter, 0) Then
                                        mCheckFilter = False
                                        Exit Function
                                    End If
                                ElseIf Trim$(tlCurrSEE.sFixedTime) = "N" Then
                                    If Not mMatchTestFile(ilFilter, 1) Then
                                        mCheckFilter = False
                                        Exit Function
                                    End If
                                End If
                            '7/8/11: Make T2 work like T1
                            'Case "CTE2"
                            '    If Not mMatchTestFile(ilFilter, tlCurrSEE.l2CteCode) Then
                            '        mCheckFilter = False
                            '        Exit Function
                            '    End If
                        End Select
                    ElseIf ilFilterType = 9 Then
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Fixed Time" Then
                            If Not mMatchTestList(ilFilter, tlCurrSEE.sFixedTime) Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Title 1" Then
                            If Not mMatchTestList(ilFilter, slComment) Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Title 2" Then
                            If Not mMatchTestList(ilFilter, slComment) Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                    ElseIf ilFilterType = 1 Then    'Event ID
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Event ID" Then
                            If Not mMatchTestValue(ilFilter, tlCurrSEE.lEventID) Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                    ElseIf ilFilterType = 2 Then
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Item ID" Then
                            ilMatch = mMatchTestString(ilFilter, tlCurrSEE.sAudioItemID)
                            If Not ilMatch Then
                                ilMatch = mMatchTestString(ilFilter, tlCurrSEE.sProtItemID)
                            End If
                            If Not ilMatch Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "ISCI" Then
                            ilMatch = mMatchTestString(ilFilter, tlCurrSEE.sAudioISCI)
                            If Not ilMatch Then
                                ilMatch = mMatchTestString(ilFilter, tlCurrSEE.sProtISCI)
                            End If
                            If Not ilMatch Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                    ElseIf ilFilterType = 6 Then    'Time
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Time" Then
                            If Not mMatchTestValue(ilFilter, tlCurrSEE.lTime) Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                    ElseIf ilFilterType = 8 Then    'Length
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Silence" Then
                            If Not mMatchTestValue(ilFilter, tlCurrSEE.lSilenceTime) Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Duration" Then
                            If Not mMatchTestValue(ilFilter, tlCurrSEE.lDuration) Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                    End If
                    tmFilterValues(ilFilter).iUsed = True
                    Exit For
                End If
            Next ilField
        End If
    Next ilFilter
    
    
End Function

Private Function mMatchTestFile(ilFilter As Integer, llFileCode As Long) As Integer
    Dim ilMatch As Integer
    Dim ilOrTest As Integer
    Dim ilAndTest As Integer
    
    If tmFilterValues(ilFilter).iOperator = 1 Then
        ilMatch = False
        For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilOrTest).iOperator) Then
                If llFileCode = tmFilterValues(ilOrTest).lCode Then
                    ilMatch = True
                    Exit For
                End If
            End If
        Next ilOrTest
    ElseIf tmFilterValues(ilFilter).iOperator = 2 Then
        ilMatch = True
        For ilAndTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilAndTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilAndTest).iOperator) Then
                If llFileCode = tmFilterValues(ilAndTest).lCode Then
                    ilMatch = False
                    Exit For
                End If
            End If
        Next ilAndTest
    Else
        ilMatch = False
    End If
    For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
        If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilOrTest).iOperator) Then
            tmFilterValues(ilOrTest).iUsed = True
        End If
    Next ilOrTest
    mMatchTestFile = ilMatch
End Function
Private Function mMatchTestList(ilFilter As Integer, slValue As String) As Integer
    Dim ilMatch As Integer
    Dim ilOrTest As Integer
    Dim ilAndTest As Integer
    
    If tmFilterValues(ilFilter).iOperator = 1 Then
        ilMatch = False
        For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilOrTest).iOperator) Then
                If StrComp(Trim$(slValue), Trim$(tmFilterValues(ilOrTest).sValue), vbTextCompare) = 0 Then
                    ilMatch = True
                    Exit For
                End If
            End If
        Next ilOrTest
    ElseIf tmFilterValues(ilFilter).iOperator = 2 Then
        ilMatch = True
        For ilAndTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilAndTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilAndTest).iOperator) Then
                If StrComp(Trim$(slValue), Trim$(tmFilterValues(ilOrTest).sValue), vbTextCompare) = 0 Then
                    ilMatch = False
                    Exit For
                End If
            End If
        Next ilAndTest
    Else
        ilMatch = False
    End If
    For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
        If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilOrTest).iOperator) Then
            tmFilterValues(ilOrTest).iUsed = True
        End If
    Next ilOrTest
    mMatchTestList = ilMatch
End Function

Private Function mMatchTestString(ilFilter As Integer, slString As String) As Integer
    Dim ilMatch As Integer
    Dim ilOrTest As Integer
    Dim ilAndTest As Integer
    
    If tmFilterValues(ilFilter).iOperator = 1 Then
        ilMatch = False
        For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilOrTest).iOperator) Then
                If StrComp(Trim$(slString), Trim$(tmFilterValues(ilOrTest).sValue), vbTextCompare) = 0 Then
                    ilMatch = True
                    Exit For
                End If
            End If
        Next ilOrTest
    ElseIf tmFilterValues(ilFilter).iOperator = 2 Then
        ilMatch = True
        For ilAndTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilAndTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilAndTest).iOperator) Then
                If StrComp(Trim$(slString), Trim$(tmFilterValues(ilAndTest).sValue), vbTextCompare) = 0 Then
                    ilMatch = False
                    Exit For
                End If
            End If
        Next ilAndTest
    Else
        ilMatch = False
    End If
    For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
        If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilOrTest).iOperator) Then
            tmFilterValues(ilOrTest).iUsed = True
        End If
    Next ilOrTest
    mMatchTestString = ilMatch
End Function

Private Function mMatchTestValue(ilFilter As Integer, llValue As Long) As Integer
    Dim ilMatch As Integer
    Dim ilOrTest As Integer
    Dim ilAndTest As Integer
    Dim ilBetween As Integer
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilNameMatch As Integer
    
    ilMatch = False
    If tmFilterValues(ilFilter).iOperator = 1 Then   'Equal Match
        ilMatch = False
        For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilOrTest).iOperator) Then
                If llValue = tmFilterValues(ilOrTest).lCode Then
                    ilMatch = True
                    Exit For
                End If
            End If
        Next ilOrTest
    ElseIf tmFilterValues(ilFilter).iOperator = 2 Then   'Not Equal Match
        ilMatch = True
        For ilAndTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilAndTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilAndTest).iOperator) Then
                If llValue = tmFilterValues(ilAndTest).lCode Then
                    ilMatch = False
                    Exit For
                End If
            End If
        Next ilAndTest
    End If
    If tmFilterValues(ilFilter).iOperator <> 2 Then
        'Look for Greater Than
        If (Not ilMatch) Then
            For ilLoop = ilFilter To UBound(tmFilterValues) - 1 Step 1
                If (tmFilterValues(ilLoop).iOperator = 3) And (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilLoop).sFieldName) Then   'Greater Than
                    ilIndex = -1
                    For ilBetween = ilLoop + 1 To UBound(tmFilterValues) - 1 Step 1
                        If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilBetween).sFieldName) And ((tmFilterValues(ilBetween).iOperator = 4) Or (tmFilterValues(ilBetween).iOperator = 6)) Then
                            If ilIndex = -1 Then
                                If tmFilterValues(ilBetween).lCode - tmFilterValues(ilLoop).lCode > 0 Then
                                    ilIndex = ilBetween
                                End If
                            Else
                                If tmFilterValues(ilBetween).lCode - tmFilterValues(ilFilter).lCode > 0 Then
                                    If (tmFilterValues(ilBetween).lCode - tmFilterValues(ilLoop).lCode) < (tmFilterValues(ilIndex).lCode - tmFilterValues(ilLoop).lCode) Then
                                        ilIndex = ilBetween
                                    End If
                                End If
                            End If
                        End If
                    Next ilBetween
                    If ilIndex = -1 Then
'                        If llValue > tmFilterValues(ilLoop).lCode Then
'                            ilMatch = True
'                            Exit For
'                        End If
                    Else
                        tmFilterValues(ilLoop).iUsed = True
                        tmFilterValues(ilIndex).iUsed = True
                        If llValue > tmFilterValues(ilLoop).lCode Then
                            If tmFilterValues(ilIndex).iOperator = 4 Then
                                If llValue < tmFilterValues(ilIndex).lCode Then
                                    ilMatch = True
                                    Exit For
                                End If
                            End If
                            If tmFilterValues(ilIndex).iOperator = 6 Then
                                If llValue <= tmFilterValues(ilIndex).lCode Then
                                    ilMatch = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            Next ilLoop
        End If
        'Look for Less than
        If (Not ilMatch) Then
            For ilLoop = ilFilter To UBound(tmFilterValues) - 1 Step 1
                If (tmFilterValues(ilLoop).iOperator = 5) And (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilLoop).sFieldName) Then   'Greater Than
                    ilIndex = -1
                    For ilBetween = ilLoop + 1 To UBound(tmFilterValues) - 1 Step 1
                        If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilBetween).sFieldName) And ((tmFilterValues(ilBetween).iOperator = 4) Or (tmFilterValues(ilBetween).iOperator = 6)) Then
                            If ilIndex = -1 Then
                                If tmFilterValues(ilBetween).lCode - tmFilterValues(ilLoop).lCode > 0 Then
                                    ilIndex = ilBetween
                                End If
                            Else
                                If tmFilterValues(ilBetween).lCode - tmFilterValues(ilFilter).lCode > 0 Then
                                    If (tmFilterValues(ilBetween).lCode - tmFilterValues(ilLoop).lCode) < (tmFilterValues(ilIndex).lCode - tmFilterValues(ilLoop).lCode) Then
                                        ilIndex = ilBetween
                                    End If
                                End If
                            End If
                        End If
                    Next ilBetween
                    If ilIndex = -1 Then
'                        If llValue >= tmFilterValues(ilLoop).lCode Then
'                            ilMatch = True
'                            Exit For
'                        End If
                    Else
                        tmFilterValues(ilLoop).iUsed = True
                        tmFilterValues(ilIndex).iUsed = True
                        If llValue >= tmFilterValues(ilLoop).lCode Then
                            If tmFilterValues(ilIndex).iOperator = 4 Then
                                If llValue < tmFilterValues(ilIndex).lCode Then
                                    ilMatch = True
                                    Exit For
                                End If
                            End If
                            If tmFilterValues(ilIndex).iOperator = 6 Then
                                If llValue <= tmFilterValues(ilIndex).lCode Then
                                    ilMatch = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            Next ilLoop
        End If
    End If
    If ((Not ilMatch) And (tmFilterValues(ilFilter).iOperator <> 2)) Or ((ilMatch) And (tmFilterValues(ilFilter).iOperator = 2)) Then
        ilNameMatch = False
        ilMatch = False
        For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilOrTest).iUsed = False) And ((tmFilterValues(ilOrTest).iOperator <> 1) And (tmFilterValues(ilOrTest).iOperator <> 2)) Then
                If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) Then
                    If ilOrTest <> ilFilter Then
                        ilNameMatch = True
                    End If
                    If tmFilterValues(ilOrTest).iOperator = 3 Then   'Greater Than
                        If llValue > tmFilterValues(ilOrTest).lCode Then
                            ilMatch = True
                            Exit For
                        End If
                    End If
                    If tmFilterValues(ilOrTest).iOperator = 4 Then   'Less Than
                        If llValue < tmFilterValues(ilOrTest).lCode Then
                            ilMatch = True
                            Exit For
                        End If
                    End If
                    If tmFilterValues(ilOrTest).iOperator = 5 Then   'Greater Than or Equal
                        If llValue >= tmFilterValues(ilOrTest).lCode Then
                            ilMatch = True
                            Exit For
                        End If
                    End If
                    If tmFilterValues(ilOrTest).iOperator = 6 Then   'Less Than or Equal
                        If llValue <= tmFilterValues(ilOrTest).lCode Then
                            ilMatch = True
                            Exit For
                        End If
                    End If
                End If
            End If
        Next ilOrTest
        If Not ilNameMatch Then
            If tmFilterValues(ilFilter).iOperator = 2 Then
                ilMatch = True
            End If
        End If
    End If
    For ilLoop = ilFilter To UBound(tmFilterValues) - 1 Step 1
        If tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilLoop).sFieldName Then
            tmFilterValues(ilLoop).iUsed = True
        End If
    Next ilLoop
    mMatchTestValue = ilMatch
End Function

Private Sub mCreateUsedArrays()
    Dim llLoop As Long
    Dim ilTest As Integer
    Dim ilFound As Integer
    Dim ilBDE As Integer
    Dim ilASE As Integer
    Dim ilANE As Integer
    Dim ilATE As Integer
    Dim ilETE As Integer
    Dim ilFNE As Integer
    Dim ilMTE As Integer
    Dim ilNNE As Integer
    Dim ilRNE As Integer
    Dim ilTTE As Integer
    Dim ilCCE As Integer
    Dim ilSCE As Integer
    Dim ilCTE As Integer
    
    ReDim tgYNMatchList(0 To 2) As MATCHLIST
    tgYNMatchList(0).sValue = "Y"
    tgYNMatchList(0).lValue = 0
    tgYNMatchList(1).sValue = "N"
    tgYNMatchList(1).lValue = 1
    If UBound(tmCurrSEE) <= 0 Then
        ReDim tgUsedBDE(0 To UBound(tgCurrBDE)) As BDE
        For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
            LSet tgUsedBDE(ilBDE) = tgCurrBDE(ilBDE)
        Next ilBDE
        ReDim tgUsedANE(0 To UBound(tgCurrANE)) As ANE
        For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            LSet tgUsedANE(ilANE) = tgCurrANE(ilANE)
        Next ilANE
        ReDim tgUsedATE(0 To UBound(tgCurrATE)) As ATE
        For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
            LSet tgUsedATE(ilATE) = tgCurrATE(ilATE)
        Next ilATE
        ReDim tgUsedETE(0 To UBound(tgCurrETE)) As ETE
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            LSet tgUsedETE(ilETE) = tgCurrETE(ilETE)
        Next ilETE
        ReDim tgUsedFNE(0 To UBound(tgCurrFNE)) As FNE
        For ilFNE = 0 To UBound(tgCurrFNE) - 1 Step 1
            LSet tgUsedFNE(ilFNE) = tgCurrFNE(ilFNE)
        Next ilFNE
        ReDim tgUsedMTE(0 To UBound(tgCurrMTE)) As MTE
        For ilMTE = 0 To UBound(tgCurrMTE) - 1 Step 1
            LSet tgUsedMTE(ilMTE) = tgCurrMTE(ilMTE)
        Next ilMTE
        ReDim tgUsedNNE(0 To UBound(tgCurrNNE)) As NNE
        For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
            LSet tgUsedNNE(ilNNE) = tgCurrNNE(ilNNE)
        Next ilNNE
        ReDim tgUsedRNE(0 To UBound(tgCurrRNE)) As RNE
        For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
            LSet tgUsedRNE(ilRNE) = tgCurrRNE(ilRNE)
        Next ilRNE
        ReDim tgUsedStartTTE(0 To UBound(tgCurrStartTTE)) As TTE
        For ilTTE = 0 To UBound(tgCurrStartTTE) - 1 Step 1
            LSet tgUsedStartTTE(ilTTE) = tgCurrStartTTE(ilTTE)
        Next ilTTE
        ReDim tgUsedEndTTE(0 To UBound(tgCurrEndTTE)) As TTE
        For ilTTE = 0 To UBound(tgCurrEndTTE) - 1 Step 1
            LSet tgUsedEndTTE(ilTTE) = tgCurrEndTTE(ilTTE)
        Next ilTTE
        ReDim tgUsedAudioCCE(0 To UBound(tgCurrAudioCCE)) As CCE
        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            LSet tgUsedAudioCCE(ilCCE) = tgCurrAudioCCE(ilCCE)
        Next ilCCE
        ReDim tgUsedBusCCE(0 To UBound(tgCurrBusCCE)) As CCE
        For ilCCE = 0 To UBound(tgCurrBusCCE) - 1 Step 1
            LSet tgUsedBusCCE(ilCCE) = tgCurrBusCCE(ilCCE)
        Next ilCCE
        ReDim tgUsedSCE(0 To UBound(tgCurrSCE)) As SCE
        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            LSet tgUsedSCE(ilSCE) = tgCurrSCE(ilSCE)
        Next ilSCE
        '7/8/11: Make T2 work like T1
        'ReDim tgUsedT2CTE(0 To UBound(tgCurrCTE)) As CTE
        'For ilCTE = 0 To UBound(tgCurrCTE) - 1 Step 1
        '    LSet tgUsedT2CTE(ilCTE) = tgCurrCTE(ilCTE)
        'Next ilCTE
        ReDim tgT1MatchList(0 To 0) As MATCHLIST
        ReDim tgT2MatchList(0 To 0) As MATCHLIST
        Exit Sub
    End If
    ReDim tgUsedBDE(0 To 0) As BDE
    ReDim tgUsedANE(0 To 0) As ANE
    ReDim tgUsedATE(0 To 0) As ATE
    ReDim tgUsedETE(0 To 0) As ETE
    ReDim tgUsedFNE(0 To 0) As FNE
    ReDim tgUsedMTE(0 To 0) As MTE
    ReDim tgUsedNNE(0 To 0) As NNE
    ReDim tgUsedRNE(0 To 0) As RNE
    ReDim tgUsedStartTTE(0 To 0) As TTE
    ReDim tgUsedEndTTE(0 To 0) As TTE
    ReDim tgUsedAudioCCE(0 To 0) As CCE
    ReDim tgUsedBusCCE(0 To 0) As CCE
    ReDim tgUsedSCE(0 To 0) As SCE
    '7/8/11: Make T2 work like T1
    'ReDim tgUsedT2CTE(0 To 0) As CTE
    ReDim tgT1MatchList(0 To 0) As MATCHLIST
    '7/8/11: Make T2 work like T1
    ReDim tgT2MatchList(0 To 0) As MATCHLIST
    For llLoop = 0 To UBound(tmCurrSEE) - 1 Step 1
        'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        '    If tmCurrSEE(llLoop).iBdeCode = tgCurrBDE(ilBDE).iCode Then
            ilBDE = gBinarySearchBDE(tmCurrSEE(llLoop).iBdeCode, tgCurrBDE())
            If ilBDE <> -1 Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedBDE) - 1 Step 1
                    If tgUsedBDE(ilTest).iCode = tgCurrBDE(ilBDE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedBDE(UBound(tgUsedBDE)) = tgCurrBDE(ilBDE)
                    ReDim Preserve tgUsedBDE(0 To UBound(tgUsedBDE) + 1) As BDE
                End If
        '        Exit For
            End If
        'Next ilBDE
        'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
        '    If tmCurrSEE(llLoop).iAudioAseCode = tgCurrASE(ilASE).iCode Then
            ilASE = gBinarySearchASE(tmCurrSEE(llLoop).iAudioAseCode, tgCurrASE())
            If ilASE <> -1 Then
                'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                '    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                    ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                    If ilANE <> -1 Then
                        ilFound = False
                        For ilTest = 0 To UBound(tgUsedANE) - 1 Step 1
                            If tgUsedANE(ilTest).iCode = tgCurrANE(ilANE).iCode Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilTest
                        If Not ilFound Then
                            LSet tgUsedANE(UBound(tgUsedANE)) = tgCurrANE(ilANE)
                            ReDim Preserve tgUsedANE(0 To UBound(tgUsedANE) + 1) As ANE
                        End If
                    End If
                'Next ilANE
        '        Exit For
            End If
        'Next ilASE
        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        '    If tmCurrSEE(llLoop).iProtAneCode = tgCurrANE(ilANE).iCode Then
            ilANE = gBinarySearchANE(tmCurrSEE(llLoop).iProtAneCode, tgCurrANE())
            If ilANE <> -1 Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedANE) - 1 Step 1
                    If tgUsedANE(ilTest).iCode = tgCurrANE(ilANE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedANE(UBound(tgUsedANE)) = tgCurrANE(ilANE)
                    ReDim Preserve tgUsedANE(0 To UBound(tgUsedANE) + 1) As ANE
                End If
            End If
        'Next ilANE
        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        '    If tmCurrSEE(llLoop).iBkupAneCode = tgCurrANE(ilANE).iCode Then
            ilANE = gBinarySearchANE(tmCurrSEE(llLoop).iBkupAneCode, tgCurrANE())
            If ilANE <> -1 Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedANE) - 1 Step 1
                    If tgUsedANE(ilTest).iCode = tgCurrANE(ilANE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedANE(UBound(tgUsedANE)) = tgCurrANE(ilANE)
                    ReDim Preserve tgUsedANE(0 To UBound(tgUsedANE) + 1) As ANE
                End If
            End If
        'Next ilANE
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If tmCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedETE) - 1 Step 1
                    If tgUsedETE(ilTest).iCode = tgCurrETE(ilETE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedETE(UBound(tgUsedETE)) = tgCurrETE(ilETE)
                    ReDim Preserve tgUsedETE(0 To UBound(tgUsedETE) + 1) As ETE
                End If
                Exit For
            End If
        Next ilETE
        For ilFNE = 0 To UBound(tgCurrFNE) - 1 Step 1
            If tmCurrSEE(llLoop).iFneCode = tgCurrFNE(ilFNE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedFNE) - 1 Step 1
                    If tgUsedFNE(ilTest).iCode = tgCurrFNE(ilFNE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedFNE(UBound(tgUsedFNE)) = tgCurrFNE(ilFNE)
                    ReDim Preserve tgUsedFNE(0 To UBound(tgUsedFNE) + 1) As FNE
                End If
                Exit For
            End If
        Next ilFNE
        For ilMTE = 0 To UBound(tgCurrMTE) - 1 Step 1
            If tmCurrSEE(llLoop).iMteCode = tgCurrMTE(ilMTE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedMTE) - 1 Step 1
                    If tgUsedMTE(ilTest).iCode = tgCurrMTE(ilMTE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedMTE(UBound(tgUsedMTE)) = tgCurrMTE(ilMTE)
                    ReDim Preserve tgUsedMTE(0 To UBound(tgUsedMTE) + 1) As MTE
                End If
                Exit For
            End If
        Next ilMTE
        'For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
        '    If tmCurrSEE(llLoop).iStartNneCode = tgCurrNNE(ilNNE).iCode Then
            ilNNE = gBinarySearchNNE(tmCurrSEE(llLoop).iStartNneCode, tgCurrNNE())
            If ilNNE <> -1 Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedNNE) - 1 Step 1
                    If tgUsedNNE(ilTest).iCode = tgCurrNNE(ilNNE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedNNE(UBound(tgUsedNNE)) = tgCurrNNE(ilNNE)
                    ReDim Preserve tgUsedNNE(0 To UBound(tgUsedNNE) + 1) As NNE
                End If
        '        Exit For
            End If
        'Next ilNNE
        'For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
        '    If tmCurrSEE(llLoop).iEndNneCode = tgCurrNNE(ilNNE).iCode Then
            ilNNE = gBinarySearchNNE(tmCurrSEE(llLoop).iEndNneCode, tgCurrNNE())
            If ilNNE <> -1 Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedNNE) - 1 Step 1
                    If tgUsedNNE(ilTest).iCode = tgCurrNNE(ilNNE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedNNE(UBound(tgUsedNNE)) = tgCurrNNE(ilNNE)
                    ReDim Preserve tgUsedNNE(0 To UBound(tgUsedNNE) + 1) As NNE
                End If
        '        Exit For
            End If
        'Next ilNNE
        'For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
        '    If tmCurrSEE(llLoop).i1RneCode = tgCurrNNE(ilRNE).iCode Then
            ilRNE = gBinarySearchRNE(tmCurrSEE(llLoop).i1RneCode, tgCurrRNE())
            If ilRNE <> -1 Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedRNE) - 1 Step 1
                    If tgUsedRNE(ilTest).iCode = tgCurrRNE(ilRNE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedRNE(UBound(tgUsedRNE)) = tgCurrRNE(ilRNE)
                    ReDim Preserve tgUsedRNE(0 To UBound(tgUsedRNE) + 1) As RNE
                End If
        '        Exit For
            End If
        'Next ilRNE
        'For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
        '    If tmCurrSEE(llLoop).i2RneCode = tgCurrNNE(ilRNE).iCode Then
            ilRNE = gBinarySearchRNE(tmCurrSEE(llLoop).i2RneCode, tgCurrRNE())
            If ilRNE <> -1 Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedRNE) - 1 Step 1
                    If tgUsedRNE(ilTest).iCode = tgCurrRNE(ilRNE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedRNE(UBound(tgUsedRNE)) = tgCurrRNE(ilRNE)
                    ReDim Preserve tgUsedRNE(0 To UBound(tgUsedRNE) + 1) As RNE
                End If
        '        Exit For
            End If
        'Next ilRNE
        For ilTTE = 0 To UBound(tgCurrStartTTE) - 1 Step 1
            If tmCurrSEE(llLoop).iStartTteCode = tgCurrStartTTE(ilTTE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedStartTTE) - 1 Step 1
                    If tgUsedStartTTE(ilTest).iCode = tgCurrStartTTE(ilTTE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedStartTTE(UBound(tgUsedStartTTE)) = tgCurrStartTTE(ilTTE)
                    ReDim Preserve tgUsedStartTTE(0 To UBound(tgUsedStartTTE) + 1) As TTE
                End If
                Exit For
            End If
        Next ilTTE
        For ilTTE = 0 To UBound(tgCurrEndTTE) - 1 Step 1
            If tmCurrSEE(llLoop).iEndTteCode = tgCurrEndTTE(ilTTE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedEndTTE) - 1 Step 1
                    If tgUsedEndTTE(ilTest).iCode = tgCurrEndTTE(ilTTE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedEndTTE(UBound(tgUsedEndTTE)) = tgCurrEndTTE(ilTTE)
                    ReDim Preserve tgUsedEndTTE(0 To UBound(tgUsedEndTTE) + 1) As TTE
                End If
                Exit For
            End If
        Next ilTTE
        For ilCCE = 0 To UBound(tgCurrBusCCE) - 1 Step 1
            If tmCurrSEE(llLoop).iBusCceCode = tgCurrBusCCE(ilCCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedBusCCE) - 1 Step 1
                    If tgUsedBusCCE(ilTest).iCode = tgCurrBusCCE(ilCCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedBusCCE(UBound(tgUsedBusCCE)) = tgCurrBusCCE(ilCCE)
                    ReDim Preserve tgUsedBusCCE(0 To UBound(tgUsedBusCCE) + 1) As CCE
                End If
                Exit For
            End If
        Next ilCCE
        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If tmCurrSEE(llLoop).iAudioCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedAudioCCE) - 1 Step 1
                    If tgUsedAudioCCE(ilTest).iCode = tgCurrAudioCCE(ilCCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedAudioCCE(UBound(tgUsedAudioCCE)) = tgCurrAudioCCE(ilCCE)
                    ReDim Preserve tgUsedAudioCCE(0 To UBound(tgUsedAudioCCE) + 1) As CCE
                End If
                Exit For
            End If
        Next ilCCE
         For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If tmCurrSEE(llLoop).iBkupCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedAudioCCE) - 1 Step 1
                    If tgUsedAudioCCE(ilTest).iCode = tgCurrAudioCCE(ilCCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedAudioCCE(UBound(tgUsedAudioCCE)) = tgCurrAudioCCE(ilCCE)
                    ReDim Preserve tgUsedAudioCCE(0 To UBound(tgUsedAudioCCE) + 1) As CCE
                End If
                Exit For
            End If
        Next ilCCE
        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If tmCurrSEE(llLoop).iProtCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedAudioCCE) - 1 Step 1
                    If tgUsedAudioCCE(ilTest).iCode = tgCurrAudioCCE(ilCCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedAudioCCE(UBound(tgUsedAudioCCE)) = tgCurrAudioCCE(ilCCE)
                    ReDim Preserve tgUsedAudioCCE(0 To UBound(tgUsedAudioCCE) + 1) As CCE
                End If
                Exit For
            End If
        Next ilCCE
        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            If tmCurrSEE(llLoop).i1SceCode = tgCurrSCE(ilSCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedSCE) - 1 Step 1
                    If tgUsedSCE(ilTest).iCode = tgCurrSCE(ilSCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedSCE(UBound(tgUsedSCE)) = tgCurrSCE(ilSCE)
                    ReDim Preserve tgUsedSCE(0 To UBound(tgUsedSCE) + 1) As SCE
                End If
                Exit For
            End If
        Next ilSCE
        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            If tmCurrSEE(llLoop).i2SceCode = tgCurrSCE(ilSCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedSCE) - 1 Step 1
                    If tgUsedSCE(ilTest).iCode = tgCurrSCE(ilSCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedSCE(UBound(tgUsedSCE)) = tgCurrSCE(ilSCE)
                    ReDim Preserve tgUsedSCE(0 To UBound(tgUsedSCE) + 1) As SCE
                End If
                Exit For
            End If
        Next ilSCE
        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            If tmCurrSEE(llLoop).i3SceCode = tgCurrSCE(ilSCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedSCE) - 1 Step 1
                    If tgUsedSCE(ilTest).iCode = tgCurrSCE(ilSCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedSCE(UBound(tgUsedSCE)) = tgCurrSCE(ilSCE)
                    ReDim Preserve tgUsedSCE(0 To UBound(tgUsedSCE) + 1) As SCE
                End If
                Exit For
            End If
        Next ilSCE
        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            If tmCurrSEE(llLoop).i4SceCode = tgCurrSCE(ilSCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedSCE) - 1 Step 1
                    If tgUsedSCE(ilTest).iCode = tgCurrSCE(ilSCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedSCE(UBound(tgUsedSCE)) = tgCurrSCE(ilSCE)
                    ReDim Preserve tgUsedSCE(0 To UBound(tgUsedSCE) + 1) As SCE
                End If
                Exit For
            End If
        Next ilSCE
        '7/8/11: Make T2 work like T1
        'For ilCTE = 0 To UBound(tgCurrCTE) - 1 Step 1
        '    If tmCurrSEE(llLoop).l2CteCode = tgCurrCTE(ilCTE).lCode Then
        '        ilFound = False
        '        For ilTest = 0 To UBound(tgUsedT2CTE) - 1 Step 1
        '            If tgUsedT2CTE(ilTest).lCode = tgCurrCTE(ilCTE).lCode Then
        '                ilFound = True
        '                Exit For
        '            End If
        '        Next ilTest
        '        If Not ilFound Then
        '            LSet tgUsedT2CTE(UBound(tgUsedT2CTE)) = tgCurrCTE(ilCTE)
        '            ReDim Preserve tgUsedT2CTE(0 To UBound(tgUsedT2CTE) + 1) As CTE
        '        End If
        '        Exit For
        '    End If
        'Next ilCTE
        ilFound = False
        For ilTest = 0 To UBound(tgT2MatchList) - 1 Step 1
            If StrComp(Trim$(tgT2MatchList(ilTest).sValue), Trim$(smT2Comment(llLoop)), vbTextCompare) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilTest
        If Not ilFound Then
            tgT2MatchList(UBound(tgT2MatchList)).sValue = smT2Comment(llLoop)
            tgT2MatchList(UBound(tgT2MatchList)).lValue = llLoop
            ReDim Preserve tgT2MatchList(0 To UBound(tgT2MatchList) + 1) As MATCHLIST
        End If
        
        ilFound = False
        For ilTest = 0 To UBound(tgT1MatchList) - 1 Step 1
            If StrComp(Trim$(tgT1MatchList(ilTest).sValue), Trim$(smT1Comment(llLoop)), vbTextCompare) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilTest
        If Not ilFound Then
            tgT1MatchList(UBound(tgT1MatchList)).sValue = smT1Comment(llLoop)
            tgT1MatchList(UBound(tgT1MatchList)).lValue = llLoop
            ReDim Preserve tgT1MatchList(0 To UBound(tgT1MatchList) + 1) As MATCHLIST
        End If
   Next llLoop
    For ilANE = 0 To UBound(tgUsedANE) - 1 Step 1
        For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
            If tgUsedANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedATE) - 1 Step 1
                    If tgUsedATE(ilTest).iCode = tgCurrATE(ilATE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedATE(UBound(tgUsedATE)) = tgCurrATE(ilATE)
                    ReDim Preserve tgUsedATE(0 To UBound(tgUsedATE) + 1) As ATE
                End If
                Exit For
            End If
        Next ilATE
    Next ilANE
End Sub







Private Function mExportCol(llRow As Long, llCol As Long) As Integer
    Dim slStr As String
    Dim ilETE As Integer
    Dim ilEPE As Integer
    Dim ilUsed As Integer
    
    mExportCol = True
    If Trim$(grdEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
        slStr = Trim$(grdEvents.TextMatrix(llRow, EVENTTYPEINDEX))
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                For ilEPE = 0 To UBound(tgCurrEPE) - 1 Step 1
                    If tgCurrEPE(ilEPE).sType = "U" Then
                        If tgCurrETE(ilETE).iCode = tgCurrEPE(ilEPE).iEteCode Then
                            Select Case llCol
                                Case BUSNAMEINDEX
                                    If tgCurrEPE(ilEPE).sBus <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case BUSCTRLINDEX
                                    If tgCurrEPE(ilEPE).sBusControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case EVENTTYPEINDEX
                                    'Event Type exported if any other column exported and tgStartColAFE.iEventType >0
                                Case EVENTIDINDEX
                                    'Event ID exported if any other column is exported and tgStartColAFE.iEventID > 0
                                Case TIMEINDEX
                                    If tgCurrEPE(ilEPE).sTime <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case STARTTYPEINDEX
                                    If tgCurrEPE(ilEPE).sStartType <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case FIXEDINDEX
                                    If tgCurrEPE(ilEPE).sFixedTime <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case ENDTYPEINDEX
                                    If tgCurrEPE(ilEPE).sEndType <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case DURATIONINDEX
                                    If tgCurrEPE(ilEPE).sDuration <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case MATERIALINDEX
                                    If tgCurrEPE(ilEPE).sMaterialType <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIONAMEINDEX
                                    If tgCurrEPE(ilEPE).sAudioName <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIOITEMIDINDEX
                                    If tgCurrEPE(ilEPE).sAudioItemID <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIOISCIINDEX
                                    If tgCurrEPE(ilEPE).sAudioISCI <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIOCTRLINDEX
                                    If tgCurrEPE(ilEPE).sAudioControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case BACKUPNAMEINDEX
                                    If tgCurrEPE(ilEPE).sBkupAudioName <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case BACKUPCTRLINDEX
                                    If tgCurrEPE(ilEPE).sBkupAudioControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTNAMEINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioName <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTITEMIDINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioItemID <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTISCIINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioISCI <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTCTRLINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case RELAY1INDEX
                                    If tgCurrEPE(ilEPE).sRelay1 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case RELAY2INDEX
                                    If tgCurrEPE(ilEPE).sRelay2 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case FOLLOWINDEX
                                    If tgCurrEPE(ilEPE).sFollow <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCETIMEINDEX
                                    If tgCurrEPE(ilEPE).sSilenceTime <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE1INDEX
                                    If tgCurrEPE(ilEPE).sSilence1 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE2INDEX
                                    If tgCurrEPE(ilEPE).sSilence2 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE3INDEX
                                    If tgCurrEPE(ilEPE).sSilence3 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE4INDEX
                                    If tgCurrEPE(ilEPE).sSilence4 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case NETCUE1INDEX
                                    If tgCurrEPE(ilEPE).sStartNetcue <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case NETCUE2INDEX
                                    If tgCurrEPE(ilEPE).sStopNetcue <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case TITLE1INDEX
                                    If tgCurrEPE(ilEPE).sTitle1 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case TITLE2INDEX
                                    If tgCurrEPE(ilEPE).sTitle2 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                            End Select
                            Exit For
                        End If
                    End If
                Next ilEPE
                For ilEPE = 0 To UBound(tgCurrEPE) - 1 Step 1
                    If tgCurrEPE(ilEPE).sType = "E" Then
                        If tgCurrETE(ilETE).iCode = tgCurrEPE(ilEPE).iEteCode Then
                            Select Case llCol
                                Case BUSNAMEINDEX
                                    If tgCurrEPE(ilEPE).sBus <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case BUSCTRLINDEX
                                    If tgCurrEPE(ilEPE).sBusControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case EVENTTYPEINDEX
                                    'Always exported if any other col is exported
                                Case EVENTIDINDEX
                                    'Always exported if any other col is exported
                                Case TIMEINDEX
                                    If tgCurrEPE(ilEPE).sTime <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case STARTTYPEINDEX
                                    If tgCurrEPE(ilEPE).sStartType <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case FIXEDINDEX
                                    If tgCurrEPE(ilEPE).sFixedTime <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case ENDTYPEINDEX
                                    If tgCurrEPE(ilEPE).sEndType <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case DURATIONINDEX
                                    If tgCurrEPE(ilEPE).sDuration <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case MATERIALINDEX
                                    If tgCurrEPE(ilEPE).sMaterialType <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIONAMEINDEX
                                    If tgCurrEPE(ilEPE).sAudioName <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIOITEMIDINDEX
                                    If tgCurrEPE(ilEPE).sAudioItemID <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIOISCIINDEX
                                    If tgCurrEPE(ilEPE).sAudioISCI <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIOCTRLINDEX
                                    If tgCurrEPE(ilEPE).sAudioControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case BACKUPNAMEINDEX
                                    If tgCurrEPE(ilEPE).sBkupAudioName <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case BACKUPCTRLINDEX
                                    If tgCurrEPE(ilEPE).sBkupAudioControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTNAMEINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioName <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTITEMIDINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioItemID <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTISCIINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioISCI <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTCTRLINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case RELAY1INDEX
                                    If tgCurrEPE(ilEPE).sRelay1 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case RELAY2INDEX
                                    If tgCurrEPE(ilEPE).sRelay2 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case FOLLOWINDEX
                                    If tgCurrEPE(ilEPE).sFollow <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCETIMEINDEX
                                    If tgCurrEPE(ilEPE).sSilenceTime <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE1INDEX
                                    If tgCurrEPE(ilEPE).sSilence1 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE2INDEX
                                    If tgCurrEPE(ilEPE).sSilence2 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE3INDEX
                                    If tgCurrEPE(ilEPE).sSilence3 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE4INDEX
                                    If tgCurrEPE(ilEPE).sSilence4 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case NETCUE1INDEX
                                    If tgCurrEPE(ilEPE).sStartNetcue <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case NETCUE2INDEX
                                    If tgCurrEPE(ilEPE).sStopNetcue <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case TITLE1INDEX
                                    If tgCurrEPE(ilEPE).sTitle1 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case TITLE2INDEX
                                    If tgCurrEPE(ilEPE).sTitle2 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                            End Select
                            Exit For
                        End If
                    End If
                Next ilEPE
                Exit For
            End If
        Next ilETE
    End If
End Function
Private Function mGetLibName(slLibNameOrDDECode As String) As String
    Dim slStr As String
    Dim llDNE As Long
    Dim llDSE As Long
    Dim ilRet As Integer
    Dim llDeeCode As Long
    Dim ilPos As Integer
    Dim slSpotTime As String
    Dim slDEECode As String
    
    slStr = slLibNameOrDDECode
    ilPos = InStr(1, slLibNameOrDDECode, "DEE=", vbTextCompare)
    If ilPos > 0 Then
        slSpotTime = ""
        slDEECode = Val(Mid$(slStr, ilPos + 4))
        ilPos = InStr(1, slDEECode, "/", vbTextCompare)
        If ilPos > 0 Then
            slSpotTime = Mid$(slDEECode, ilPos)
            slDEECode = Left$(slDEECode, ilPos - 1)
        End If
        llDeeCode = Val(slDEECode)
        If llDeeCode > 0 Then
            ilRet = gGetRec_DEE_DayEvent(llDeeCode, "EngrAsAirCompare-mMoveSEERecToCtrls: DEE", tmDee)
            ilRet = gGetRec_DHE_DayHeaderInfo(tmDee.lDheCode, "EngrAsAirCompare-mMoveSEERecToCtrls: DHE", tmDHE)
            
            If tmDHE.sType <> "T" Then
                llDNE = gBinarySearchDNE(tmDHE.lDneCode, tgCurrLibDNE)
                If llDNE <> -1 Then
                    slStr = Trim$(tgCurrLibDNE(llDNE).sName)
                Else
                    slStr = "Name Missing"
                End If
            Else
                llDNE = gBinarySearchDNE(tmDHE.lDneCode, tgCurrTempDNE)
                If llDNE <> -1 Then
                    slStr = Trim$(tgCurrTempDNE(llDNE).sName)
                End If
            End If
            llDSE = gBinarySearchDSE(tmDHE.lDseCode, tgCurrDSE)
            If llDSE <> -1 Then
                slStr = slStr & "/" & Trim$(tgCurrDSE(llDSE).sName)
            End If
            slStr = slStr & slSpotTime
        End If
    End If
    mGetLibName = slStr
End Function

Private Sub mSortErrorsToTop()
    Dim ilCol As Integer
    
    gSetMousePointer grdEvents, grdEvents, vbHourglass
    grdEvents.Redraw = False
    If imLastColSorted >= 0 Then
        If imLastSort = flexSortStringNoCaseDescending Then
            imLastSort = flexSortStringNoCaseAscending
        Else
            imLastSort = flexSortStringNoCaseDescending
        End If
        ilCol = imLastColSorted
        mSortCol ilCol
    Else
        imLastSort = -1
        mSortCol TIMEINDEX
    End If
    grdEvents.Redraw = True
    gSetMousePointer grdEvents, grdEvents, vbDefault
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    gSetMousePointer grdEvents, grdEvents, vbHourglass
    mShowEvents sgAsAirCompareDate
    cmcCancel.SetFocus
    gSetMousePointer grdEvents, grdEvents, vbDefault
    imFieldChgd = False
End Sub

Public Sub mPopGrid()
    Dim llRow As Long
    Dim llLoop As Long
    Dim slCategory As String
    Dim ilETE As Integer
    Dim llAvailLength As Long
    Dim llAvailTest As Long
    Dim llTimeTest As Long
    Dim ilRet As Integer
    
    ReDim tmCurrSEE(0 To UBound(tgCurrSEE)) As SEE
    ReDim smT1Comment(0 To UBound(tmCurrSEE)) As String
    ReDim smT2Comment(0 To UBound(tmCurrSEE)) As String
    llRow = 0
    For llLoop = 0 To UBound(tgCurrSEE) - 1 Step 1
        If tgCurrSEE(llLoop).sAction <> "D" Then
            slCategory = ""
            ilETE = gBinarySearchETE(tgCurrSEE(llLoop).iEteCode, tgCurrETE)
            If ilETE <> -1 Then
                slCategory = tgCurrETE(ilETE).sCategory
            End If
            If slCategory <> "A" Then
                LSet tmCurrSEE(llRow) = tgCurrSEE(llLoop)
                tmCurrSEE(llRow).sInsertFlag = "N"
                tmCurrSEE(llRow).lAvailLength = tmCurrSEE(llRow).lDuration
                smT1Comment(llRow) = ""
                If tgCurrSEE(llLoop).l1CteCode > 0 Then
                    ilRet = gGetRec_CTE_CommtsTitleAPI(hmCTE, tgCurrSEE(llLoop).l1CteCode, "EngrLib- mMoveRecToCtrl for CTE", tmCTE)
                    smT1Comment(llRow) = Trim$(tmCTE.sComment)
                End If
                smT2Comment(llRow) = ""
                If tgCurrSEE(llLoop).l2CteCode > 0 Then
                    ilRet = gGetRec_CTE_CommtsTitleAPI(hmCTE, tgCurrSEE(llLoop).l2CteCode, "EngrLib- mMoveRecToCtrl for CTE", tmCTE)
                    smT2Comment(llRow) = Trim$(tmCTE.sComment)
                End If
                llRow = llRow + 1
            End If
        End If
    Next llLoop
    ReDim Preserve tmCurrSEE(0 To llRow) As SEE
    ReDim Preserve smT1Comment(0 To llRow) As String
    ReDim Preserve smT2Comment(0 To llRow) As String
    grdEvents.Redraw = False
    mMoveSEERecToCtrls
    imLastColSorted = -1
    mSortCol TIMEINDEX
    grdEvents.Redraw = True

End Sub

Private Function mIsDiscrepant(llInRow As Long, tlAAE As AAE) As Integer
    Dim llSEERow As Long
    Dim llAAERow As Long
    Dim llStartToleranceTime As Long
    Dim llEndToleranceTime As Long
    Dim llAAETime As Long
    Dim llStartToleranceLength As Long
    Dim llEndToleranceLength As Long
    Dim llAAELength As Long
    
    'Compare columns
    llSEERow = llInRow
    llAAERow = llInRow + 1
    mIsDiscrepant = False
    grdEvents.Row = llAAERow
    If grdEvents.TextMatrix(llAAERow, AIRDATEERRORINDEX) = "Y" Then
        mIsDiscrepant = True
        grdEvents.Col = AIRDATEERRORINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llAAERow, AUTOOFFINDEX)) <> "" Then
        mIsDiscrepant = True
        grdEvents.Col = AUTOOFFINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llAAERow, DATAERRORINDEX)) <> "" Then
        mIsDiscrepant = True
        grdEvents.Col = DATAERRORINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llAAERow, SCHEDULEERRORINDEX)) <> "" Then
        mIsDiscrepant = True
        grdEvents.Col = SCHEDULEERRORINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, BUSNAMEINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, BUSNAMEINDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = BUSNAMEINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, BUSCTRLINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, BUSCTRLINDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = BUSCTRLINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, EVENTTYPEINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, EVENTTYPEINDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = EVENTTYPEINDEX
        grdEvents.CellForeColor = vbRed
    End If
    'Add Time window
    If Trim$(grdEvents.TextMatrix(llAAERow, TIMEINDEX)) <> "" Then
        llStartToleranceTime = gStrLengthInTenthToLong(grdEvents.TextMatrix(llSEERow, TIMEINDEX)) - tgSOE.lTimeTolerance
        llEndToleranceTime = gStrLengthInTenthToLong(grdEvents.TextMatrix(llSEERow, TIMEINDEX)) + tgSOE.lTimeTolerance
        llAAETime = gStrLengthInTenthToLong(grdEvents.TextMatrix(llAAERow, TIMEINDEX))
        If llStartToleranceTime > 0 Then
            If llEndToleranceTime <= 864000 Then
                If (llAAETime < llStartToleranceTime) Or (llAAETime > llEndToleranceTime) Then
                    mIsDiscrepant = True
                    grdEvents.Col = TIMEINDEX
                    grdEvents.CellForeColor = vbRed
                End If
            Else
                If ((llAAETime >= llStartToleranceTime) And (llAAETime <= 864000)) Or ((llAAETime >= 0) And (llAAETime <= (llEndToleranceTime - 864000))) Then
                Else
                    mIsDiscrepant = True
                    grdEvents.Col = TIMEINDEX
                    grdEvents.CellForeColor = vbRed
                End If
            End If
        Else
                If ((llAAETime >= (864000 + llStartToleranceTime)) And (llAAETime <= 864000)) Or ((llAAETime >= 0) And (llAAETime <= llEndToleranceTime)) Then
                Else
                    mIsDiscrepant = True
                    grdEvents.Col = TIMEINDEX
                    grdEvents.CellForeColor = vbRed
                End If
        End If
    Else
        grdEvents.Row = llSEERow
        mIsDiscrepant = True
        grdEvents.Col = TIMEINDEX
        grdEvents.CellForeColor = vbRed
    End If
    grdEvents.Row = llAAERow
    If Trim$(grdEvents.TextMatrix(llSEERow, STARTTYPEINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, STARTTYPEINDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = STARTTYPEINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If (Trim$(grdEvents.TextMatrix(llSEERow, FIXEDINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, FIXEDINDEX))) Or (Trim$(tlAAE.sTrueTime) <> "") Then
        mIsDiscrepant = True
        grdEvents.Col = FIXEDINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, ENDTYPEINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, ENDTYPEINDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = ENDTYPEINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llAAERow, DURATIONINDEX)) <> "" Then
        llStartToleranceLength = gStrLengthInTenthToLong(grdEvents.TextMatrix(llSEERow, DURATIONINDEX)) - tgSOE.lLengthTolerance
        llEndToleranceLength = gStrLengthInTenthToLong(grdEvents.TextMatrix(llSEERow, DURATIONINDEX)) + tgSOE.lLengthTolerance
        llAAELength = gStrLengthInTenthToLong(grdEvents.TextMatrix(llAAERow, DURATIONINDEX))
        If (llAAELength >= llStartToleranceLength) And (llAAELength <= llEndToleranceLength) Then
        Else
            mIsDiscrepant = True
            grdEvents.Col = DURATIONINDEX
            grdEvents.CellForeColor = vbRed
        End If
    Else
        grdEvents.Row = llSEERow
        mIsDiscrepant = True
        grdEvents.Col = DURATIONINDEX
        grdEvents.CellForeColor = vbRed
    End If
    grdEvents.Row = llAAERow
    If Trim$(grdEvents.TextMatrix(llSEERow, MATERIALINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, MATERIALINDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = MATERIALINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, AUDIONAMEINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, AUDIONAMEINDEX)) Or (Trim$(tlAAE.sSourceConflict) <> "") Or (Trim$(tlAAE.sSourceUnavail) <> "") Then
        mIsDiscrepant = True
        grdEvents.Col = AUDIONAMEINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, AUDIOITEMIDINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, AUDIOITEMIDINDEX)) Or (Trim$(tlAAE.sSourceItem) <> "") Then
        mIsDiscrepant = True
        grdEvents.Col = AUDIOITEMIDINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, AUDIOISCIINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, AUDIOISCIINDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = AUDIOISCIINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, AUDIOCTRLINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, AUDIOCTRLINDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = AUDIOCTRLINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, BACKUPNAMEINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, BACKUPNAMEINDEX)) Or (Trim$(tlAAE.sBkupSrceUnavail) <> "") Or (Trim$(tlAAE.sBkupSrceItem) <> "") Then
        mIsDiscrepant = True
        grdEvents.Col = BACKUPNAMEINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, BACKUPCTRLINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, BACKUPCTRLINDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = BACKUPCTRLINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, PROTNAMEINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, PROTNAMEINDEX)) Or (Trim$(tlAAE.sProtSrceUnavail) <> "") Then
        mIsDiscrepant = True
        grdEvents.Col = PROTNAMEINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, PROTITEMIDINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, PROTITEMIDINDEX)) Or (Trim$(tlAAE.sProtSrceItem) <> "") Then
        mIsDiscrepant = True
        grdEvents.Col = PROTITEMIDINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, PROTISCIINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, PROTISCIINDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = PROTISCIINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, PROTCTRLINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, PROTCTRLINDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = PROTCTRLINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, RELAY1INDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, RELAY1INDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = RELAY1INDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, RELAY2INDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, RELAY2INDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = RELAY2INDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, FOLLOWINDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, FOLLOWINDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = FOLLOWINDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llAAERow, SILENCETIMEINDEX)) <> "" Then
        llStartToleranceLength = gStrLengthInTenthToLong(grdEvents.TextMatrix(llSEERow, SILENCETIMEINDEX)) - tgSOE.lLengthTolerance
        llEndToleranceLength = gStrLengthInTenthToLong(grdEvents.TextMatrix(llSEERow, SILENCETIMEINDEX)) + tgSOE.lLengthTolerance
        llAAELength = gStrLengthInTenthToLong(grdEvents.TextMatrix(llAAERow, SILENCETIMEINDEX))
        If (llAAELength >= llStartToleranceLength) And (llAAELength <= llEndToleranceLength) Then
        Else
            mIsDiscrepant = True
            grdEvents.Col = SILENCETIMEINDEX
            grdEvents.CellForeColor = vbRed
        End If
    Else
        grdEvents.Row = llSEERow
        mIsDiscrepant = True
        grdEvents.Col = SILENCETIMEINDEX
        grdEvents.CellForeColor = vbRed
    End If
    grdEvents.Row = llAAERow
    If Trim$(grdEvents.TextMatrix(llSEERow, SILENCE1INDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, SILENCE1INDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = SILENCE1INDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, SILENCE2INDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, SILENCE2INDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = SILENCE2INDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, SILENCE3INDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, SILENCE3INDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = SILENCE3INDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, SILENCE4INDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, SILENCE4INDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = SILENCE4INDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, NETCUE1INDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, NETCUE1INDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = NETCUE1INDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, NETCUE2INDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, NETCUE2INDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = NETCUE2INDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, TITLE1INDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, TITLE1INDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = TITLE1INDEX
        grdEvents.CellForeColor = vbRed
    End If
    If Trim$(grdEvents.TextMatrix(llSEERow, TITLE2INDEX)) <> Trim$(grdEvents.TextMatrix(llAAERow, TITLE2INDEX)) Then
        mIsDiscrepant = True
        grdEvents.Col = TITLE2INDEX
        grdEvents.CellForeColor = vbRed
    End If
End Function

Private Sub mDisplayAAE(llAAE As Long, llRow As Long)
    Dim slStr As String
    
    grdEvents.TextMatrix(llRow, TMCURRSEEINDEX) = Trim$(Str$(llAAE))
    
    If DateValue(tmSHE.sAirDate) <> DateValue(tmCurrAAE(llAAE).sAirDate) Then
        grdEvents.TextMatrix(llRow, AIRDATEERRORINDEX) = "Y"
    End If
    grdEvents.TextMatrix(llRow, AIRDATEINDEX) = Trim$(tmCurrAAE(llAAE).sAirDate)
    grdEvents.TextMatrix(llRow, AUTOOFFINDEX) = Trim$(tmCurrAAE(llAAE).sAutoOff)
    grdEvents.TextMatrix(llRow, DATAERRORINDEX) = Trim$(tmCurrAAE(llAAE).sData)
    grdEvents.TextMatrix(llRow, SCHEDULEERRORINDEX) = Trim$(tmCurrAAE(llAAE).sSchedule)
    
    grdEvents.TextMatrix(llRow, EVENTIDINDEX) = Trim$(Str$(tmCurrAAE(llAAE).lEventID))
    grdEvents.TextMatrix(llRow, LIBNAMEINDEX) = ""
    grdEvents.TextMatrix(llRow, BUSNAMEINDEX) = Trim$(tmCurrAAE(llAAE).sBusName)
    grdEvents.TextMatrix(llRow, BUSCTRLINDEX) = Trim$(tmCurrAAE(llAAE).sBusControl)
    grdEvents.TextMatrix(llRow, EVENTTYPEINDEX) = Trim$(tmCurrAAE(llAAE).sEventType)
    grdEvents.TextMatrix(llRow, TIMEINDEX) = gLongToStrLengthInTenth(tmCurrAAE(llAAE).lAirTime, True)
    grdEvents.TextMatrix(llRow, STARTTYPEINDEX) = Trim$(tmCurrAAE(llAAE).sStartType)
    grdEvents.TextMatrix(llRow, FIXEDINDEX) = Trim$(tmCurrAAE(llAAE).sFixedTime)
    grdEvents.TextMatrix(llRow, ENDTYPEINDEX) = Trim$(tmCurrAAE(llAAE).sEndType)
    grdEvents.TextMatrix(llRow, DURATIONINDEX) = Trim$(tmCurrAAE(llAAE).sDuration)
    grdEvents.TextMatrix(llRow, MATERIALINDEX) = Trim$(tmCurrAAE(llAAE).sMaterialType)
    grdEvents.TextMatrix(llRow, AUDIONAMEINDEX) = Trim$(tmCurrAAE(llAAE).sAudioName)
    grdEvents.TextMatrix(llRow, AUDIOITEMIDINDEX) = Trim$(tmCurrAAE(llAAE).sAudioItemID)
    grdEvents.TextMatrix(llRow, AUDIOISCIINDEX) = Trim$(tmCurrAAE(llAAE).sAudioISCI)
    grdEvents.TextMatrix(llRow, AUDIOCTRLINDEX) = Trim$(tmCurrAAE(llAAE).sAudioCrtlChar)
    grdEvents.TextMatrix(llRow, BACKUPNAMEINDEX) = Trim$(tmCurrAAE(llAAE).sBkupAudioName)
    grdEvents.TextMatrix(llRow, BACKUPCTRLINDEX) = Trim$(tmCurrAAE(llAAE).sBkupCtrlChar)
    grdEvents.TextMatrix(llRow, PROTNAMEINDEX) = Trim$(tmCurrAAE(llAAE).sProtAudioName)
    grdEvents.TextMatrix(llRow, PROTITEMIDINDEX) = Trim$(tmCurrAAE(llAAE).sProtItemID)
    grdEvents.TextMatrix(llRow, PROTISCIINDEX) = Trim$(tmCurrAAE(llAAE).sProtISCI)
    grdEvents.TextMatrix(llRow, PROTCTRLINDEX) = Trim$(tmCurrAAE(llAAE).sProtCtrlChar)
    grdEvents.TextMatrix(llRow, RELAY1INDEX) = Trim$(tmCurrAAE(llAAE).sRelay1)
    grdEvents.TextMatrix(llRow, RELAY2INDEX) = Trim$(tmCurrAAE(llAAE).sRelay2)
    grdEvents.TextMatrix(llRow, FOLLOWINDEX) = Trim$(tmCurrAAE(llAAE).sFollow)
    grdEvents.TextMatrix(llRow, SILENCETIMEINDEX) = Trim$(tmCurrAAE(llAAE).sSilenceTime)
    grdEvents.TextMatrix(llRow, SILENCE1INDEX) = Trim$(tmCurrAAE(llAAE).sSilence1)
    grdEvents.TextMatrix(llRow, SILENCE2INDEX) = Trim$(tmCurrAAE(llAAE).sSilence2)
    grdEvents.TextMatrix(llRow, SILENCE3INDEX) = Trim$(tmCurrAAE(llAAE).sSilence3)
    grdEvents.TextMatrix(llRow, SILENCE4INDEX) = Trim$(tmCurrAAE(llAAE).sSilence4)
    grdEvents.TextMatrix(llRow, NETCUE1INDEX) = Trim$(tmCurrAAE(llAAE).sNetcueStart)
    grdEvents.TextMatrix(llRow, NETCUE2INDEX) = Trim$(tmCurrAAE(llAAE).sNetcueEnd)
    grdEvents.TextMatrix(llRow, TITLE1INDEX) = Trim$(tmCurrAAE(llAAE).sTitle1)
    grdEvents.TextMatrix(llRow, TITLE2INDEX) = Trim$(tmCurrAAE(llAAE).sTitle2)
    grdEvents.TextMatrix(llRow, PCODEINDEX) = tmCurrAAE(llAAE).lCode
    slStr = Trim$(Str$(llRow))
    Do While Len(slStr) < 6
        slStr = "0" & slStr
    Loop
    grdEvents.TextMatrix(llRow, ROWSORTINDEX) = slStr
    grdEvents.TextMatrix(llRow, ROWTYPEINDEX) = "AAE"
    grdEvents.TextMatrix(llRow - 1, DISCREPANCYINDEX) = ""
    grdEvents.TextMatrix(llRow, DISCREPANCYINDEX) = ""
End Sub
