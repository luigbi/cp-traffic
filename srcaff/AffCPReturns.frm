VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCPReturns 
   Caption         =   "Post C.P."
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   9480
   Icon            =   "AffCPReturns.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   9480
   Begin V81Affiliate.CSI_Calendar edcStartDate 
      Height          =   240
      Left            =   810
      TabIndex        =   35
      Top             =   300
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   397
      Text            =   "4/15/2020"
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
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin V81Affiliate.CSI_Calendar edcEndDate 
      Height          =   240
      Left            =   810
      TabIndex        =   36
      Top             =   930
      Width           =   1035
      _ExtentX        =   1746
      _ExtentY        =   476
      Text            =   "4/15/2020"
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   0   'False
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
      CSI_CurDayForeColor=   51200
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin VB.Timer tmcGrid 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8460
      Top             =   5265
   End
   Begin VB.PictureBox pbcPostFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   17
      Top             =   1095
      Width           =   60
   End
   Begin VB.Timer tmcSSSort 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   9375
      Top             =   4725
   End
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9330
      Top             =   5490
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "3) Display"
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
      Left            =   135
      TabIndex        =   11
      Top             =   1365
      Width           =   4485
      Begin VB.OptionButton optStatus 
         Caption         =   "Did Not Air"
         Height          =   255
         Index           =   2
         Left            =   3015
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   30
         Width           =   1215
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Received"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   30
         Width           =   1260
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Outstanding"
         Height          =   255
         Index           =   0
         Left            =   645
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   30
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "Display"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   45
         Width           =   630
      End
   End
   Begin VB.ListBox lbcLookup2 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "AffCPReturns.frx":08CA
      Left            =   4050
      List            =   "AffCPReturns.frx":08CC
      TabIndex        =   26
      Top             =   5940
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.ListBox lbcLookup1 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "AffCPReturns.frx":08CE
      Left            =   4440
      List            =   "AffCPReturns.frx":08D0
      TabIndex        =   25
      Top             =   6045
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<--- Previous Dates"
      Height          =   330
      Left            =   120
      TabIndex        =   18
      Top             =   6105
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Dates ---->"
      Height          =   330
      Left            =   1845
      TabIndex        =   19
      Top             =   6135
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   9240
      Begin VB.Frame frmVeh 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   5685
         TabIndex        =   30
         Top             =   225
         Visible         =   0   'False
         Width           =   2715
         Begin VB.OptionButton rbcVeh 
            Caption         =   "All Veh"
            Height          =   195
            Index           =   1
            Left            =   1650
            TabIndex        =   32
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton rbcVeh 
            Caption         =   "Active Veh"
            Height          =   195
            Index           =   0
            Left            =   345
            TabIndex        =   31
            Top             =   0
            Value           =   -1  'True
            Width           =   1200
         End
         Begin VB.Label lblSort 
            Caption         =   "By:"
            Height          =   255
            Left            =   0
            TabIndex        =   37
            Top             =   -15
            Width           =   360
         End
      End
      Begin VB.ComboBox cboPSSort 
         Height          =   315
         ItemData        =   "AffCPReturns.frx":08D2
         Left            =   2040
         List            =   "AffCPReturns.frx":08D4
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   855
         Width           =   3405
      End
      Begin VB.ComboBox cboSSSort 
         Height          =   315
         ItemData        =   "AffCPReturns.frx":08D6
         Left            =   5655
         List            =   "AffCPReturns.frx":08D8
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   855
         Width           =   3405
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5685
         TabIndex        =   7
         Top             =   525
         Width           =   2985
         Begin VB.OptionButton optSSSort 
            Caption         =   "Stations"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   345
            TabIndex        =   8
            Top             =   0
            Width           =   990
         End
         Begin VB.OptionButton optSSSort 
            Caption         =   "DMA"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1650
            TabIndex        =   9
            Top             =   0
            Width           =   765
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   480
         Left            =   2055
         TabIndex        =   1
         Top             =   225
         Width           =   3225
         Begin VB.OptionButton optPSSort 
            Caption         =   "All Veh"
            Height          =   195
            Index           =   3
            Left            =   1545
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   285
            Width           =   915
         End
         Begin VB.OptionButton optPSSort 
            Caption         =   "Active Veh"
            Height          =   195
            Index           =   2
            Left            =   345
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   285
            Width           =   1260
         End
         Begin VB.OptionButton optPSSort 
            Caption         =   "Stations"
            Height          =   195
            Index           =   0
            Left            =   345
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton optPSSort 
            Caption         =   "DMA"
            Height          =   210
            Index           =   1
            Left            =   1545
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   990
         End
         Begin VB.Label Label1 
            Caption         =   "By:"
            Height          =   255
            Left            =   0
            TabIndex        =   2
            Top             =   -15
            Width           =   1350
         End
      End
      Begin VB.Label lacEndDate 
         Appearance      =   0  'Flat
         Caption         =   "End"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         TabIndex        =   34
         Top             =   870
         Width           =   585
      End
      Begin VB.Label lacStartDate 
         Appearance      =   0  'Flat
         Caption         =   "Start"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   33
         Top             =   240
         Width           =   570
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   5850
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6255
      FormDesignWidth =   9480
   End
   Begin VB.CommandButton cmdBarCode 
      Caption         =   "Bar Code Scan"
      Enabled         =   0   'False
      Height          =   330
      Left            =   6375
      TabIndex        =   20
      Top             =   5895
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   330
      Left            =   8040
      TabIndex        =   21
      Top             =   5805
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPost 
      Height          =   3690
      Left            =   135
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1755
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   6509
      _Version        =   393216
      Cols            =   17
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
      _Band(0).Cols   =   17
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label4 
      Caption         =   "* Click to Change Date Sort Order"
      Height          =   225
      Index           =   4
      Left            =   165
      TabIndex        =   28
      Top             =   5850
      Width           =   2715
   End
   Begin VB.Label Label4 
      Caption         =   "Magenta = Partially Posted"
      ForeColor       =   &H00FF00FF&
      Height          =   225
      Index           =   3
      Left            =   4635
      TabIndex        =   27
      Top             =   5550
      Width           =   1965
   End
   Begin VB.Label Label4 
      Caption         =   "Blue = Did Not Air"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   2
      Left            =   3195
      TabIndex        =   24
      Top             =   5550
      Width           =   1410
   End
   Begin VB.Label Label4 
      Caption         =   "Green = Received"
      ForeColor       =   &H00008000&
      Height          =   225
      Index           =   1
      Left            =   1695
      TabIndex        =   23
      Top             =   5550
      Width           =   1395
   End
   Begin VB.Label Label4 
      Caption         =   "Red = Outstanding"
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   0
      Left            =   180
      TabIndex        =   22
      Top             =   5550
      Width           =   1395
   End
End
Attribute VB_Name = "frmCPReturns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmCPRturns - allows for tracking of returns of CPs
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private rst_att As ADODB.Recordset

Private imShttCode As Integer
Private imVefCode As Integer
Private imCPIndex As Integer
Private imCPMax As Integer
Private imFirstTime As Integer
Private imSetAlign As Integer
Private imStatus As Integer
Private imIntegralSet As Integer
Private imInChg As Integer
Private imBSMode As Integer
Private imHeaderClick As Integer
Private imDateClick As Integer
Private lmSSSortIndex As Long
Private tmCPDate() As CPDATE
Private tmCPVehicle() As CPVEHICLE
Private hmAst As Integer
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO

'Grid Controls
Private imShowGridBox As Integer
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on

Const VEHICLEINDEX = 0
Const VEHICLECODEINDEX = 1
Const DATE1INDEX = 2
Const CPTTCODE1INDEX = 3
Const DATE2INDEX = 4
Const CPTTCODE2INDEX = 5
Const DATE3INDEX = 6
Const CPTTCODE3INDEX = 7
Const DATE4INDEX = 8
Const CPTTCODE4INDEX = 9
Const DATE5INDEX = 10
Const CPTTCODE5INDEX = 11
Const DATE6INDEX = 12
Const CPTTCODE6INDEX = 13
Const DATE7INDEX = 14
Const CPTTCODE7INDEX = 15
Const NOMISSINGINDEX = 16



Private Sub mClearGrid()
    Dim llRow As Long
    gGrid_Clear grdPost, True
    grdPost.Row = 0
    grdPost.Col = DATE1INDEX
    grdPost.CellAlignment = flexAlignCenterTop
    'grdPost.TextMatrix(0, 2) = Chr$(171)
    grdPost.Row = 0
    grdPost.Col = DATE4INDEX
    grdPost.CellAlignment = flexAlignCenterTop
    'grdPost.TextMatrix(0, 8) = "Dates*"
    grdPost.Row = 0
    grdPost.Col = DATE7INDEX
    grdPost.CellAlignment = flexAlignCenterTop
    grdPost.Row = 0
    grdPost.Col = NOMISSINGINDEX
    grdPost.CellAlignment = flexAlignRightTop
    For llRow = grdPost.FixedRows To grdPost.Rows - 1 Step 1
        grdPost.Row = llRow
        grdPost.Col = VEHICLEINDEX
        grdPost.CellBackColor = vbWhite
        grdPost.Col = NOMISSINGINDEX
        grdPost.CellBackColor = vbWhite
    Next llRow
End Sub


Private Sub mSort(optSort As OptionButton, cboCtrl As ComboBox)
    Dim iLoop As Integer
    Dim iIndex As Integer
    
    cboCtrl.Clear
    If optSort.Value = True Then
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                cboCtrl.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                cboCtrl.ItemData(cboCtrl.NewIndex) = tgStationInfo(iLoop).iCode
            End If
        Next iLoop
        
    Else
        
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                cboCtrl.AddItem Trim$(tgStationInfo(iLoop).sMarket) & ", " & Trim$(tgStationInfo(iLoop).sCallLetters)
                cboCtrl.ItemData(cboCtrl.NewIndex) = tgStationInfo(iLoop).iCode
            End If
        Next iLoop
    End If

End Sub

Private Sub mGridPaint(iClear As Integer)
    Dim iTotRec As Integer
    Dim lSIndex As Long
    Dim lEIndex As Long
    Dim iRow As Integer
    Dim lLoop As Long
    Dim iCol As Integer
    Dim sStr As String
    Dim sName As String
    Dim iCode As Integer
    Dim iRowIndex As Integer
    'ReDim sDates(1 To 7) As String
    ReDim sDates(0 To 6) As String
    'ReDim lCodes(1 To 7) As Long
    ReDim lCodes(0 To 6) As Long
    Dim llRow As Long
    Dim llTRow As Long
    
    'Dim a, c, e, g, j, l, n, p As String
    'Dim b As Integer
    'Dim d, f, h, k, m, o, q As Long
    
    On Error GoTo ErrHand
    
    llTRow = grdPost.TopRow
    If iClear Then
        mClearGrid
    End If
    grdPost.Redraw = False
    llRow = grdPost.FixedRows
    For iRow = 0 To UBound(tmCPVehicle) - 1 Step 1
        sName = Trim$(tmCPVehicle(iRow).sName)
        If ((optPSSort(2).Value = True) Or (optPSSort(3).Value = True)) And (Trim$(tmCPVehicle(iRow).sMarket) <> "") Then
            sName = sName & ", " & Trim$(tmCPVehicle(iRow).sMarket)
        End If
        'iCode = tmCPVehicle(iRow).iCode
        iRowIndex = iRow
        lSIndex = tmCPVehicle(iRow).lCPDateIndex
        If iRow < UBound(tmCPVehicle) - 1 Then
            lEIndex = tmCPVehicle(iRow + 1).lCPDateIndex - 1
        Else
            lEIndex = UBound(tmCPDate) - 1
        End If
        iTotRec = lEIndex - lSIndex + 1
        'For lLoop = 1 To 7 Step 1
        For lLoop = 0 To 6 Step 1
            sDates(lLoop) = ""
            lCodes(lLoop) = 0
        Next lLoop
        If llRow + 1 > grdPost.Rows Then
            grdPost.AddItem ""
        End If
        grdPost.Row = llRow
        lSIndex = 7 * (imCPIndex - 1) + lSIndex
        If lSIndex <= lEIndex Then
            If lSIndex + 6 < lEIndex Then
                lEIndex = lSIndex + 6
            End If
            'iCol = 1
            iCol = 0
            For lLoop = lSIndex To lEIndex Step 1
                If tmCPDate(lLoop).iAttPostingType = 3 Then
                    sDates(iCol) = Format$(tmCPDate(lLoop).sDate, "mmm,yy")
                ElseIf tmCPDate(lLoop).iAttPostingType = 1 Then
                    If tmCPDate(lLoop).iNoSpotsGen > 0 Then
                        sDates(iCol) = tmCPDate(lLoop).sDate & Str$(tmCPDate(lLoop).iNoSpotsAired) & " of" & Str$(tmCPDate(lLoop).iNoSpotsGen)
                   Else
                        sDates(iCol) = tmCPDate(lLoop).sDate
                    End If
                Else
                    sDates(iCol) = tmCPDate(lLoop).sDate
                End If
                lCodes(iCol) = tmCPDate(lLoop).lCpttCode
                'Select Case iCol
                '    Case 0
                '        c = tmCPDate(lLoop).sDate
                '        d = tmCPDate(lLoop).lCode
                '    Case 1
                '        e = tmCPDate(lLoop).sDate
                '        f = tmCPDate(lLoop).lCode
                '    Case 2
                '        g = tmCPDate(lLoop).sDate
                '        h = tmCPDate(lLoop).lCode
                '    Case 3
                '        j = tmCPDate(lLoop).sDate
                '        k = tmCPDate(lLoop).lCode
                '    Case 4
                '        l = tmCPDate(lLoop).sDate
                '        m = tmCPDate(lLoop).lCode
                '    Case 5
                '        n = tmCPDate(lLoop).sDate
                '        o = tmCPDate(lLoop).lCode
                '    Case 6
                '        p = tmCPDate(lLoop).sDate
                '        q = tmCPDate(lLoop).lCode
                'End Select
                iCol = iCol + 1
            Next lLoop
            iCol = 2
            For lLoop = lSIndex To lEIndex Step 1
                grdPost.Col = iCol
                If tmCPDate(lLoop).iStatus = 0 Then
                    If ((tmCPDate(lLoop).iAttPostingType = 2) Or (tmCPDate(lLoop).iAttPostingType = 3)) And (tmCPDate(lLoop).iPostingStatus = 1) Then
                        grdPost.CellForeColor = vbMagenta
                    Else
                        grdPost.CellForeColor = vbRed
                    End If
                ElseIf tmCPDate(lLoop).iStatus = 1 Then
                    grdPost.CellForeColor = RGB(0, 128, 0)  '64)
                Else
                    grdPost.CellForeColor = vbBlue
                End If
                iCol = iCol + 2
            Next lLoop
        End If
        If iClear Then
            sStr = sName & "|" & iRowIndex  'iCode
            grdPost.TextMatrix(llRow, VEHICLEINDEX) = sName
            grdPost.TextMatrix(llRow, VEHICLECODEINDEX) = iRowIndex
            grdPost.Row = llRow
            grdPost.Col = VEHICLEINDEX
            grdPost.CellBackColor = LIGHTYELLOW
            'For lLoop = 1 To 7 Step 1
            For lLoop = 0 To 6 Step 1
                sStr = sStr & "|" & sDates(lLoop) & "|" & lCodes(lLoop)
                grdPost.TextMatrix(llRow, 2 * (lLoop + 1)) = sDates(lLoop)
                grdPost.TextMatrix(llRow, 2 * (lLoop + 1) + 1) = lCodes(lLoop)
            Next lLoop
            grdPost.TextMatrix(llRow, 16) = iTotRec
            grdPost.Row = llRow
            grdPost.Col = 16
            grdPost.CellBackColor = LIGHTYELLOW
        Else
            'For lLoop = 1 To 7 Step 1
            For lLoop = 0 To 6 Step 1
                grdPost.TextMatrix(llRow, 2 * (lLoop + 1)) = sDates(lLoop)
                grdPost.TextMatrix(llRow, 2 * (lLoop + 1) + 1) = lCodes(lLoop)
            Next lLoop
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
    tmcGrid.Enabled = True
    Exit Sub
ErrHand:

    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CPReturn-mGridPaint"
End Sub

Private Sub cboPSSort_Change()
    Dim iLoop As Integer
    Dim sName As String
    Dim iLen As Integer
    Dim iSel As Integer
    Dim lRow As Long
    Dim sVefTypes As String
    Dim sRange As String
    Dim sMarket As String
    Dim llStaCode As Long

    On Error GoTo ErrHand
    
    
    If imInChg Then
        Exit Sub
    End If
    imInChg = True
    Screen.MousePointer = vbHourglass
    sName = LTrim$(cboPSSort.Text)
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    If (optPSSort(2).Value = False) And (optPSSort(3).Value = False) Then       'Sorting by stations/markets
        lRow = SendMessageByString(cboPSSort.hwnd, CB_FINDSTRING, -1, sName)
    Else
        lRow = SendMessageByString(cboPSSort.hwnd, CB_FINDSTRING, -1, sName)
    End If
    If lRow >= 0 Then
        'If optPSSort(2).Value = False Then       'Sorting by stations/markets
        '    cboPSSort.Bookmark = lRow ' + 1
        'Else
        '    cboPSSort.Bookmark = lRow
        'End If
        'cboPSSort.Text = cboPSSort.Columns(0).Text
        cboPSSort.ListIndex = lRow
        cboPSSort.SelStart = iLen
        cboPSSort.SelLength = Len(cboPSSort.Text)
        mClearGrid
        grdPost.Enabled = False
        ReDim tmCPDate(0 To 0) As CPDATE
        imCPIndex = 1
        ReDim tmCPVehicle(0 To 0) As CPVEHICLE
            
        If (optPSSort(0).Value = True) Or (optPSSort(1).Value = True) Then
            imShttCode = CInt(cboPSSort.ItemData(cboPSSort.ListIndex))
            SQLQuery = "SELECT DISTINCT vefType, vefName, vefCode, vefState"
            'SQLQuery = SQLQuery + " FROM VEF_Vehicles vef, att"
            SQLQuery = SQLQuery & " FROM VEF_Vehicles, att"
            SQLQuery = SQLQuery + " WHERE (vefCode = attVefCode "
            SQLQuery = SQLQuery + " AND attShfCode = " & imShttCode & ")"
            SQLQuery = SQLQuery + " ORDER BY vefName"
        
            Set rst = gSQLSelectCall(SQLQuery)
            cboSSSort.Clear
            imVefCode = -1
            While Not rst.EOF
                If rbcVeh(0).Value = True Then 'Only show active veh
                    If rst!vefState = "A" Then
                        If sgShowByVehType = "Y" Then
                            cboSSSort.AddItem Trim$(rst!vefType) & ":" & Trim$(rst!vefName)
                        Else
                            cboSSSort.AddItem Trim$(rst!vefName)
                        End If
                        cboSSSort.ItemData(cboSSSort.NewIndex) = rst!vefCode
                    End If
                Else                           'Show all veh
                    If sgShowByVehType = "Y" Then
                        cboSSSort.AddItem Trim$(rst!vefType) & ":" & Trim$(rst!vefName)
                    Else
                        cboSSSort.AddItem Trim$(rst!vefName)
                    End If
                    cboSSSort.ItemData(cboSSSort.NewIndex) = rst!vefCode
                End If
                rst.MoveNext
            Wend
            cboSSSort.AddItem "[All Vehicles]", 0
            cboSSSort.ItemData(cboSSSort.NewIndex) = 0
        Else
            imVefCode = CInt(cboPSSort.ItemData(cboPSSort.ListIndex))
            cboSSSort.Clear
            imShttCode = -1
            SQLQuery = "SELECT distinct shttCode, shttCallLetters"
            SQLQuery = SQLQuery + " FROM shtt, att"
            SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode
            SQLQuery = SQLQuery + " AND shttCode = attShfCode AND shttType = 0)"
            Set rst = gSQLSelectCall(SQLQuery)
            
            If optSSSort(0).Value Then
                While Not rst.EOF
                    llStaCode = gBinarySearchStation(rst!shttCallLetters)
                    If llStaCode <> -1 Then
                        sMarket = Trim$(tgStationInfo(llStaCode).sMarket)
                    Else
                        sMarket = ""
                    End If
                    If sMarket = "" Then
                        cboSSSort.AddItem Trim$(rst!shttCallLetters)
                    Else
                        cboSSSort.AddItem Trim$(rst!shttCallLetters) & ", " & Trim$(sMarket)
                    End If
                    cboSSSort.ItemData(cboSSSort.NewIndex) = rst!shttCode
                    rst.MoveNext
                Wend
            Else
                While Not rst.EOF
                    llStaCode = gBinarySearchStation(rst!shttCallLetters)
                    If llStaCode <> -1 Then
                        sMarket = Trim$(tgStationInfo(llStaCode).sMarket)
                    Else
                        sMarket = ""
                    End If
                    If sMarket = "" Then
                        cboSSSort.AddItem Trim$(rst!shttCallLetters)
                    Else
                        cboSSSort.AddItem Trim$(sMarket) & ", " & Trim$(rst!shttCallLetters)
                    End If
                    cboSSSort.ItemData(cboSSSort.NewIndex) = rst!shttCode
                    rst.MoveNext
                Wend
            End If
            
            cboSSSort.AddItem "[All Stations]", 0
            cboSSSort.ItemData(cboSSSort.NewIndex) = 0
        End If
    End If
    Screen.MousePointer = vbDefault
    imInChg = False
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CP Return-cboPSSort"
End Sub

Private Sub cboPSSort_GotFocus()
'    If imSetAlign Then
'        mClearGrid
'        imSetAlign = True
'    End If
End Sub

Private Sub cboPSSort_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboPSSort_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboPSSort.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cboSSSort_Change()
    Dim iLoop As Integer
    Dim sName As String
    Dim iLen As Integer
    Dim iSel As Integer
    Dim lRow As Long
    
    
    If imInChg Then
        Exit Sub
    End If
    imInChg = True
    Screen.MousePointer = vbHourglass
    tmcSSSort.Enabled = False
    sName = LTrim$(cboSSSort.Text)
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    If (optPSSort(2).Value = True) Or (optPSSort(3).Value = True) Then       'Sorting by stations/markets
        lRow = SendMessageByString(cboSSSort.hwnd, CB_FINDSTRING, -1, sName)
    Else
        lRow = SendMessageByString(cboSSSort.hwnd, CB_FINDSTRING, -1, sName)
    End If
    If lRow >= 0 Then
        'If optPSSort(2).Value = True Then       'Sorting by stations/markets
        '    cboSSSort.Bookmark = lRow ' + 1
        'Else
        '    cboSSSort.Bookmark = lRow
        'End If
        'cboSSSort.Text = cboSSSort.Columns(0).Text
        grdPost.Enabled = False
        cboSSSort.ListIndex = lRow
        cboSSSort.SelStart = iLen
        cboSSSort.SelLength = Len(cboSSSort.Text)
        If (optPSSort(0).Value = True) Or (optPSSort(1).Value = True) Then
            imVefCode = CInt(cboSSSort.ItemData(cboSSSort.ListIndex))
        Else
            imShttCode = CInt(cboSSSort.ItemData(cboSSSort.ListIndex))
        End If
        tmcSSSort.Enabled = True
    End If
    Screen.MousePointer = vbDefault
    imInChg = False

End Sub

Private Sub cboSSSort_Click()

    cboSSSort_Change

End Sub

Private Sub cboSSSort_GotFocus()
'Removed 6/29/04- Jim request
'    'Reinstalled 3/26/04- Jim request
'    If cboSSSort.ListIndex < 0 Then
'        If cboSSSort.ListCount > 0 Then
'            cboSSSort.ListIndex = 0
'            cboSSSort.SelStart = 0
'            cboSSSort.SelLength = Len(cboSSSort.Text)
'        End If
'    End If
End Sub

Private Sub cboSSSort_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
    tmcSSSort.Enabled = False
End Sub

Private Sub cboSSSort_KeyPress(KeyAscii As Integer)
    tmcSSSort.Enabled = False
    If KeyAscii = 8 Then
        If cboSSSort.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cboSSSort_LostFocus()
    If tmcSSSort.Enabled Then
        tmcSSSort.Enabled = False
        Screen.MousePointer = vbHourglass
        mGetCPTT
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub cmdNext_Click()
    Dim iMax As Integer
    Dim iRow As Integer
    
    imCPIndex = imCPIndex + 1
    If imCPIndex > imCPMax Then
        imCPIndex = imCPMax
    Else
        mGridPaint False
    End If
End Sub

Private Sub cmdPrevious_Click()

    imCPIndex = imCPIndex - 1
    If imCPIndex < 1 Then
        imCPIndex = 1
    Else
        mGridPaint False
    End If
End Sub


Private Sub edcEndDate_GotFocus()
    tmcSSSort.Enabled = False
End Sub

Private Sub edcEndDate_LostFocus()
    tmcSSSort.Enabled = True
End Sub

Private Sub edcStartDate_GotFocus()
    tmcSSSort.Enabled = False
End Sub

Private Sub edcStartDate_LostFocus()
    tmcSSSort.Enabled = True
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    
    If imFirstTime Then
        bgAffidavitVisible = True
        'Define stylesheets for changing colors of grdCPReturns based on status
'        grdCPReturns.StyleSets("Missed").ForeColor = vbRed  'RGB(255, 0, 0)
'        grdCPReturns.StyleSets("Partial").ForeColor = vbMagenta  'RGB(255, 0, 0)
'        grdCPReturns.StyleSets("Complete").ForeColor = RGB(0, 128, 64)
'        grdCPReturns.StyleSets("Drop").ForeColor = vbBlue   'RGB(0, 0, 255)
        'Vehicle
        grdPost.ColWidth(VEHICLEINDEX) = grdPost.Width * 0.27
        'Vehoicle Code
        grdPost.ColWidth(VEHICLECODEINDEX) = 0
        'Date 1
        grdPost.ColWidth(DATE1INDEX) = grdPost.Width * 0.08
        'CPTT Date 1 Code
        grdPost.ColWidth(CPTTCODE1INDEX) = 0
        'Date 2
        grdPost.ColWidth(DATE2INDEX) = grdPost.Width * 0.08
        'CPTT Date 2 Code
        grdPost.ColWidth(CPTTCODE2INDEX) = 0
        'Date 3
        grdPost.ColWidth(DATE3INDEX) = grdPost.Width * 0.08
        'CPTT Date 3 Code
        grdPost.ColWidth(CPTTCODE3INDEX) = 0
        'Date 4
        grdPost.ColWidth(DATE4INDEX) = grdPost.Width * 0.08
        'CPTT Date 4 Code
        grdPost.ColWidth(CPTTCODE4INDEX) = 0
        'Date 5
        grdPost.ColWidth(DATE5INDEX) = grdPost.Width * 0.08
        'CPTT Date 5 Code
        grdPost.ColWidth(CPTTCODE5INDEX) = 0
        'Date 6
        grdPost.ColWidth(DATE6INDEX) = grdPost.Width * 0.08
        'CPTT Date 6 Code
        grdPost.ColWidth(CPTTCODE6INDEX) = 0
        'Date 7
        grdPost.ColWidth(DATE7INDEX) = grdPost.Width * 0.08
        'CPTT Date 7 Code
        grdPost.ColWidth(CPTTCODE7INDEX) = 0
        '# Missing
        grdPost.ColWidth(NOMISSINGINDEX) = grdPost.Width * 0.07
        grdPost.ColWidth(VEHICLEINDEX) = grdPost.Width - grdPost.ColWidth(DATE1INDEX) - grdPost.ColWidth(DATE2INDEX) - grdPost.ColWidth(DATE3INDEX) - grdPost.ColWidth(DATE4INDEX) - grdPost.ColWidth(DATE5INDEX) - grdPost.ColWidth(DATE6INDEX) - grdPost.ColWidth(DATE7INDEX) - grdPost.ColWidth(NOMISSINGINDEX) - GRIDSCROLLWIDTH '(5 * grdStation.Columns(6).Width) / 6 ' - 120
        gGrid_AlignAllColsLeft grdPost
        grdPost.TextMatrix(0, VEHICLEINDEX) = "Vehicles"
        'grdPost.Row = 0
        'grdPost.Col = 2
        'grdPost.CellAlignment = flexAlignCenterTop
        grdPost.TextMatrix(0, DATE1INDEX) = Chr$(171)
        grdPost.Row = 0
        grdPost.Col = DATE1INDEX
        grdPost.CellBackColor = LIGHTBLUE
        'grdPost.Row = 0
        'grdPost.Col = 8
        'grdPost.CellAlignment = flexAlignCenterTop
        grdPost.TextMatrix(0, DATE4INDEX) = "Dates*"
        grdPost.Row = 0
        grdPost.Col = DATE4INDEX
        grdPost.CellBackColor = LIGHTBLUE
        'grdPost.Row = 0
        'grdPost.Col = 14
        'grdPost.CellAlignment = flexAlignCenterTop
        grdPost.TextMatrix(0, DATE7INDEX) = Chr$(187)
        grdPost.Row = 0
        grdPost.Col = DATE7INDEX
        grdPost.CellBackColor = LIGHTBLUE
        
        grdPost.ColAlignment(16) = flexAlignRightTop
        grdPost.TextMatrix(0, NOMISSINGINDEX) = "#"
        gGrid_IntegralHeight grdPost
        'Moved to cboPSSort because setting cell cause grid to show prior to form
        'mClearGrid
        mSort optPSSort(0), cboPSSort
        imFirstTime = False
        If cboPSSort.Visible Then
            cboPSSort.SetFocus
        End If
    End If
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.25
    Me.Top = (Screen.Height - Me.Height) / 1.8
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmCPReturns
    gCenterForm frmCPReturns
End Sub

Private Sub Form_Resize()
    Dim iLoop As Integer
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    tmcDelay.Enabled = False
    tmcGrid.Enabled = False
    tmcSSSort.Enabled = False
    bgAffidavitVisible = False
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    
    Erase tmCPDate
    Erase tmCPVehicle
    Erase tgCPPosting
    Erase tmCPDat
    Erase tmAstInfo
    Set frmCPReturns = Nothing
End Sub


Private Sub cmdDone_Click()
    Unload frmCPReturns
    Set frmCPReturns = Nothing
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim ilRet As Integer

    On Error GoTo ErrHand

    Screen.MousePointer = vbHourglass
    frmCPReturns.Caption = "Affidavits - " & sgClientName
    'Me.Width = Screen.Width / 1.05
    'Me.Height = Screen.Height / 1.25
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2

    iCellColor = 0       'Associates iCellColor = 0 with default show missing dates (red text)
       
    'Add stations
    'SQLQuery = "SELECT shttCallLetters, shttMarket, shttCode FROM shtt ORDER BY shttCallLetters, shttMarket"
    'Set rst = gSQLSelectCall(SQLQuery, rdUseOdbc, rdOpenKeyset)
    'While Not rst.EOF
    '    cboStations.AddItem "" & rst(0).Value & ", " & rst(1).Value & ", " & rst(2).Value & ""
    '    rst.MoveNext
    'Wend
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    
    lmSSSortIndex = -1
    imFirstTime = True
    imSetAlign = True
    imCPIndex = 1
    imVefCode = -1
    imShttCode = -1
    imIntegralSet = False
    imHeaderClick = False
    imDateClick = False
    lmEnableRow = -1
    lmEnableCol = -1
    lmTopRow = -1
    imBSMode = False
    imInChg = False
    ReDim tmCPVehicle(0 To 0) As CPVEHICLE
    imStatus = 0
    Screen.MousePointer = vbDefault
    frmVeh.Visible = True
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CPReturns-Load"
End Sub

Private Sub cboPSSort_Click()
    
    cboPSSort_Change
    Exit Sub
End Sub



Private Sub grdPost_Click()
    tmcDelay.Enabled = False
    '9/19/06: Allow viewing of spots
    'If sgUstWin(7) <> "I" Then
    If (sgUstWin(7) <> "I") And (sgUstWin(7) <> "V") Then
        cmdDone.SetFocus
        Exit Sub
    End If
    DoEvents
    If imDateClick Then
        Exit Sub
    End If
    imDateClick = True
    If (grdPost.Row - 1 < VEHICLEINDEX) Or (grdPost.Row - 1 >= UBound(tmCPVehicle)) Then
        imDateClick = False
        Exit Sub
    End If
    If imHeaderClick Then
        imHeaderClick = False
        imDateClick = False
        Exit Sub
    End If
    tmcDelay.Enabled = True
    Exit Sub
End Sub

Private Sub grdPost_EnterCell()
    Dim llRow As Long
    
    llRow = grdPost.Row
End Sub

Private Sub grdPost_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Determine if in header
    If Y < grdPost.RowHeight(0) Then
        imHeaderClick = True
        If (X >= grdPost.ColPos(DATE1INDEX)) And (X <= grdPost.ColPos(DATE1INDEX) + grdPost.ColWidth(DATE1INDEX)) Then
            imCPIndex = imCPIndex - 1
            If imCPIndex < 1 Then
                imCPIndex = 1
            Else
                mGridPaint False
            End If
        End If
        If (X >= grdPost.ColPos(DATE7INDEX)) And (X <= grdPost.ColPos(DATE7INDEX) + grdPost.ColWidth(DATE7INDEX)) Then
            imCPIndex = imCPIndex + 1
            If imCPIndex > imCPMax Then
                imCPIndex = imCPMax
            Else
                mGridPaint False
            End If
        End If
        If (X >= grdPost.ColPos(DATE4INDEX)) And (X <= grdPost.ColPos(DATE4INDEX) + grdPost.ColWidth(DATE4INDEX)) Then
            mDateSort
            mGridPaint False
        End If
    End If
End Sub

Private Sub optPSSort_Click(Index As Integer)
    Dim iLoop As Integer
    
    On Error GoTo ErrHand
    
    If optPSSort(Index).Value = False Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    mClearGrid
    cboPSSort.Clear
    cboSSSort.Clear
    imShttCode = -1
    imVefCode = -1
    ReDim tmCPDate(0 To 0) As CPDATE
    imCPIndex = 1
    ReDim tmCPVehicle(0 To 0) As CPVEHICLE
    'cboStations.Text = ""
    'If optSort(0).Value = True Then
    '    cboStations.RemoveAll
    '    cboStations.Columns(0).Caption = "Stations"
    '    cboStations.Columns(0).Width = 1154
    '    cboStations.Columns(1).Caption = "Markets"
    '    cboStations.Columns(1).Width = 1964
    '    SQLQuery = "SELECT shttCallLetters, shttMarket, shttCode FROM shtt ORDER BY shttCallLetters, shttMarket"
    '    Set rst = gSQLSelectCall(SQLQuery, rdUseOdbc, rdOpenKeyset)
    '    'cboStations.AddItem "New,,0", 0
    '    While Not rst.EOF
    '        cboStations.AddItem "" & rst(0).Value & ", " & rst(1).Value & ", " & rst(2).Value & ""
    '        rst.MoveNext
    '    Wend
    'Else
    '    cboStations.RemoveAll
    '    cboStations.Columns(0).Caption = "Markets"
    '    cboStations.Columns(0).Width = 1964
    '    cboStations.Columns(1).Caption = "Stations"
    '    cboStations.Columns(1).Width = 1154
    '    SQLQuery = "SELECT shttCallLetters, shttMarket, shttCode FROM shtt ORDER BY shttMarket, shttCallLetters"
    '    Set rst = gSQLSelectCall(SQLQuery, rdUseOdbc, rdOpenKeyset)
    '    'cboStations.AddItem ",New,0", 0
    '    While Not rst.EOF
    '        cboStations.AddItem "" & rst(1).Value & ", " & rst(0).Value & ", " & rst(2).Value & ""
    '        rst.MoveNext
    '    Wend
    'End If
    If optPSSort(0).Value = True Then
        'Frame6.Visible = False
        optSSSort(0).Enabled = False
        optSSSort(1).Enabled = False
        optSSSort(0).Value = False
        optSSSort(1).Value = False
        'frmVeh.Visible = True
        rbcVeh(0).Enabled = True
        rbcVeh(1).Enabled = True
        If (rbcVeh(0).Value = False And rbcVeh(1).Value = False) Then rbcVeh(0).Value = True
        mSort optPSSort(0), cboPSSort
        grdPost.TextMatrix(0, VEHICLEINDEX) = "Vehicles"
    ElseIf optPSSort(1).Value = True Then
        'Frame6.Visible = False
        optSSSort(0).Enabled = False
        optSSSort(1).Enabled = False
        optSSSort(0).Value = False
        optSSSort(1).Value = False
        'frmVeh.Visible = True
        rbcVeh(0).Enabled = True
        rbcVeh(1).Enabled = True
        If (rbcVeh(0).Value = False And rbcVeh(1).Value = False) Then rbcVeh(0).Value = True
        mSort optPSSort(0), cboPSSort
        grdPost.TextMatrix(0, VEHICLEINDEX) = "Vehicles"
    ElseIf (optPSSort(2).Value = True) Or (optPSSort(3).Value = True) Then
        'frmVeh.Visible = False
        rbcVeh(0).Value = False
        rbcVeh(1).Value = False
        rbcVeh(0).Enabled = False
        rbcVeh(1).Enabled = False
        'Frame6.Visible = True
        optSSSort(0).Enabled = True
        optSSSort(1).Enabled = True
        If (optSSSort(0).Value = False And optSSSort(1).Value = False) Then optSSSort(0).Value = True
        For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
                If optPSSort(2).Value = True Then
                    'D.S. 09/13/02 Added support for
                    If tgVehicleInfo(iLoop).sState = "A" Then 'Active veh. only
                        cboPSSort.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
                        cboPSSort.ItemData(cboPSSort.NewIndex) = tgVehicleInfo(iLoop).iCode
                    End If
                Else
                    cboPSSort.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle) 'All veh.
                    cboPSSort.ItemData(cboPSSort.NewIndex) = tgVehicleInfo(iLoop).iCode
                End If
            'End If
        Next iLoop
        grdPost.TextMatrix(0, VEHICLEINDEX) = "Stations"

    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CPReturns-Click"
End Sub

Private Sub optSSSort_Click(Index As Integer)
    If optSSSort(Index).Value = False Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    mClearGrid
    'cboPSSort.RemoveAll
    'cboPSSort.Text = ""
    cboSSSort.Clear
    imShttCode = -1
    'imVefCode = -1
    ReDim tmCPDate(0 To 0) As CPDATE
    imCPIndex = 1
    ReDim tmCPVehicle(0 To 0) As CPVEHICLE
   
    'If optSSSort(0).Value = True Then
    '    mSort optSSSort(0), cboSSSort
    'Else
    '    mSort optSSSort(0), cboSSSort
    'End If
    cboPSSort_Change
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub optStatus_Click(Index As Integer)
    Dim iLoop As Integer
    
    If optStatus(Index).Value = False Then
        Exit Sub
    End If
    
    
    'RedText = vbRed 'RGB(255, 0, 0)
    'GreenText = RGB(0, 128, 0)
    'BlueText = vbBlue   'RGB(0, 128, 255)
    'MagentaText = vbMagenta
    
    Screen.MousePointer = vbHourglass
    
    If optStatus(0).Value = True Then       'Missing
'        For iLoop = 2 To 15 Step 2
'            grdCPReturns.Columns(iLoop).ForeColor = vbRed
'        Next iLoop
        imStatus = 0
    ElseIf optStatus(1).Value = True Then   'Complete
'        For iLoop = 2 To 15 Step 2
'            grdCPReturns.Columns(iLoop).ForeColor = RGB(0, 128, 0)
'        Next iLoop
        imStatus = 1
    Else                                                    'Dropped
'        For iLoop = 2 To 15 Step 2
'            grdCPReturns.Columns(iLoop).ForeColor = vbBlue
'        Next iLoop
        imStatus = 2
    End If
    If (cboSSSort.ListIndex = 0) And (cboSSSort.List(0) = "[All Vehicles]") Then
        imVefCode = 0
    End If
    If (cboSSSort.ListIndex = 0) And (cboSSSort.List(0) = "[All Stations]") Then
        imShttCode = 0
    End If
    mGetCPTT
    
    'grdCPReturns.SetFocus
    pbcPostFocus.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub


End Sub

Private Sub mSort_Old()
    Dim iLoop As Integer
    Dim iIndex As Integer
    
    'cboStations.Text = ""
    'cboStations.RemoveAll
    'frmDirectory!lbcStationInfo.Clear
    'If optSort(0).Value = True Then
    '    cboStations.Columns(0).Caption = "Stations"
    '    cboStations.Columns(0).Width = cboStations.Width / 3
    '    cboStations.Columns(1).Caption = "Markets"
    '    cboStations.Columns(1).Width = (2 * cboStations.Width) / 3
    '    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
    '        If tgStationInfo(iLoop).iType = 0 Then
    '            frmDirectory!lbcStationInfo.AddItem tgStationInfo(iLoop).sCallLetters & tgStationInfo(iLoop).sMarket
    '            frmDirectory!lbcStationInfo.ItemData(frmDirectory!lbcStationInfo.NewIndex) = iLoop
    '        End If
    '    Next iLoop
    '    'cboStations.AddItem "New,,-1"
    '    For iLoop = 0 To frmDirectory!lbcStationInfo.ListCount - 1 Step 1
    '        iIndex = frmDirectory!lbcStationInfo.ItemData(iLoop)
    '        cboStations.AddItem Trim$(tgStationInfo(iIndex).sCallLetters) & "|" & Trim$(tgStationInfo(iIndex).sMarket) & "|" & tgStationInfo(iIndex).iCode
    '    Next iLoop
    '
    'Else
    '
    '    cboStations.Columns(0).Caption = "Markets"
    '    cboStations.Columns(0).Width = (2 * cboStations.Width) / 3
    '    cboStations.Columns(1).Caption = "Stations"
    '    cboStations.Columns(1).Width = cboStations.Width / 3
    '    frmDirectory!lbcStationInfo.Clear
    '    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
    '        If tgStationInfo(iLoop).iType = 0 Then
    '            frmDirectory!lbcStationInfo.AddItem tgStationInfo(iLoop).sMarket & tgStationInfo(iLoop).sCallLetters
    '            frmDirectory!lbcStationInfo.ItemData(frmDirectory!lbcStationInfo.NewIndex) = iLoop
    '        End If
    '    Next iLoop
    '    'cboStations.AddItem ",New,-1"
    '    For iLoop = 0 To frmDirectory!lbcStationInfo.ListCount - 1 Step 1
    '        iIndex = frmDirectory!lbcStationInfo.ItemData(iLoop)
    '        cboStations.AddItem Trim$(tgStationInfo(iIndex).sMarket) & "|" & Trim$(tgStationInfo(iIndex).sCallLetters) & "|" & tgStationInfo(iIndex).iCode
    '    Next iLoop
    'End If
End Sub

Private Sub mClick()
    Dim iCol As Integer
    Dim iRow As Integer
    Dim lCode As Long
    Dim iColNum As Integer
    Dim llLoop As Long
    Dim iRowIndex As Integer
    Dim sDate As String
    Dim iRet As Integer
    Dim slDate As String
    Dim slPostSDate As String
    Dim slPostEDate As String
    Dim ilPos As Integer
    Dim slSvUstWin7 As String
    Dim llVeh As Long
    '7/20/12: Set for receipt only agreements
    Dim ilCPPostingStatus As Integer
    '12/12/18: save and reset imstatus as this is used by mGetCPTT
    Dim ilSvStatus As Integer
    
    On Error GoTo ErrHand
    
    DoEvents
    ilSvStatus = imStatus
    If (lmEnableCol <= 0) Or (lmEnableCol >= 16) Or imHeaderClick Then       'Don't change color of vehicle or #missing columns
        cmdDone.SetFocus
        imHeaderClick = False
        Exit Sub
    End If
    If (lmEnableRow - 1 < 0) Or (lmEnableRow - 1 >= UBound(tmCPVehicle)) Then
        cmdDone.SetFocus
        imHeaderClick = False
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    imHeaderClick = False
    iCol = lmEnableCol
    iColNum = lmEnableCol + 1
    lCode = Val(grdPost.TextMatrix(lmEnableRow, iColNum))
    If lCode <= 0 Then
        Screen.MousePointer = vbDefault
        cmdDone.SetFocus
        Exit Sub
    End If
    iRowIndex = Val(grdPost.TextMatrix(lmEnableRow, VEHICLECODEINDEX))
    For llLoop = 0 To UBound(tmCPDate) - 1 Step 1
        If lCode = tmCPDate(llLoop).lCpttCode Then
            If (optPSSort(0).Value = True) Or (optPSSort(1).Value = True) Then
                imVefCode = tmCPVehicle(iRowIndex).iCode
            Else
                imShttCode = tmCPVehicle(iRowIndex).iCode
            End If
            If tmCPDate(llLoop).iAttPostingType = 1 Then 'By spot Count
                If sgUstWin(7) <> "I" Then
                    Screen.MousePointer = vbDefault
                    cmdDone.SetFocus
                    Exit Sub
                End If
                ReDim tgCPPosting(0 To 1) As CPPOSTING
                tgCPPosting(0).lCpttCode = lCode
                tgCPPosting(0).lAttCode = tmCPDate(llLoop).lAttCode
                tgCPPosting(0).iAttTimeType = tmCPDate(llLoop).iAttTimeType
                If (optPSSort(0).Value = True) Or (optPSSort(1).Value = True) Then
                    tgCPPosting(0).iVefCode = tmCPVehicle(iRowIndex).iCode
                    tgCPPosting(0).iShttCode = imShttCode
                    tgCPPosting(0).sZone = tmCPVehicle(iRowIndex).sZone
                Else
                    tgCPPosting(0).iVefCode = imVefCode
                    tgCPPosting(0).iShttCode = tmCPVehicle(iRowIndex).iCode
                    tgCPPosting(0).sZone = tmCPVehicle(iRowIndex).sZone
                End If
                tgCPPosting(0).sDate = grdPost.TextMatrix(lmEnableRow, lmEnableCol)
                Screen.MousePointer = vbDefault
                igCPStatus = tmCPDate(llLoop).iStatus
                igCPPostingStatus = tmCPDate(llLoop).iPostingStatus
                frmCPCount.Show vbModal
                Screen.MousePointer = vbHourglass
                If (igCPStatus <> imStatus) Or (tmCPDate(llLoop).iPostingStatus <> igCPPostingStatus) Then
                    If (igCPStatus = 0) Then
                        If igCPPostingStatus = 1 Then
                            grdPost.Col = lmEnableCol
                            grdPost.CellForeColor = vbMagenta
                        Else
                            grdPost.Col = lmEnableCol
                            grdPost.CellForeColor = vbRed
                        End If
                        If imStatus <> 0 Then
                            If optStatus(0).Value = True Then
                                grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) = grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) + 1
                            ElseIf optStatus(1).Value = True Then
                                grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) = grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) - 1
                            End If
                        End If
                        imStatus = igCPStatus
                    ElseIf (igCPStatus = 1) Then
                        grdPost.Col = lmEnableCol
                        grdPost.CellForeColor = RGB(0, 128, 0)
                        If imStatus = 0 Then
                            If optStatus(0).Value = True Then
                                grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) = grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) - 1
                            ElseIf optStatus(1).Value = True Then
                                grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) = grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) + 1
                            End If
                        End If
                        imStatus = igCPStatus
                    End If
                    SQLQuery = "SELECT cpttNoSpotsGen, cpttNoSpotsAired FROM cptt"
                    SQLQuery = SQLQuery + " WHERE (cpttCode= " & lCode & ")"
                    
                    Set rst = gSQLSelectCall(SQLQuery)
                    If Not rst.EOF Then
                        tmCPDate(llLoop).iNoSpotsGen = rst!cpttNoSpotsGen
                        tmCPDate(llLoop).iNoSpotsAired = rst!cpttNoSpotsAired
                        grdPost.TextMatrix(lmEnableRow, lmEnableCol) = tmCPDate(llLoop).sDate & Str$(tmCPDate(llLoop).iNoSpotsAired) & " of" & Str$(tmCPDate(llLoop).iNoSpotsGen)
                    End If
                End If
                tmCPDate(llLoop).iStatus = imStatus
                tmCPDate(llLoop).iPostingStatus = igCPPostingStatus
                mGridPaint False
            ElseIf (tmCPDate(llLoop).iAttPostingType = 2) Or (tmCPDate(llLoop).iAttPostingType = 3) Then 'Posting by Date
                'AttExportTYpe = 1
                slSvUstWin7 = sgUstWin(7)
                SQLQuery = "SELECT attExportType FROM att WHERE (attCode = " & tmCPDate(llLoop).lAttCode & ")"
                Set rst_att = gSQLSelectCall(SQLQuery)
                If Not rst_att.EOF Then
                    'D.S. 11/12/08
                    'If rst_att!attExportType = 1 Then
                    '    sgUstWin(7) = "V"
                    'End If
                End If
                lgSelGameGsfCode = -1
                llVeh = gBinarySearchVef(CLng(imVefCode))
                If llVeh <> -1 Then
                    If tgVehicleInfo(llVeh).sVehType = "G" Then
                        igGameVefCode = imVefCode
                        If (optPSSort(0).Value = True) Or (optPSSort(1).Value = True) Then
                            igGameVefCode = tmCPVehicle(iRowIndex).iCode
                        Else
                            igGameVefCode = imVefCode
                        End If
                        sgGameStartDate = grdPost.TextMatrix(lmEnableRow, lmEnableCol)
                        sgGameEndDate = DateAdd("d", 6, sgGameStartDate)
                        lgGameAttCode = tmCPDate(llLoop).lAttCode
                        frmGetGame.Show vbModal
                        If lgSelGameGsfCode <= 0 Then
                            imStatus = ilSvStatus
                            Screen.MousePointer = vbDefault
                            cmdDone.SetFocus
                            Exit Sub
                        End If
                    End If
                Else
                    imStatus = ilSvStatus
                    Screen.MousePointer = vbDefault
                    cmdDone.SetFocus
                    Exit Sub
                End If
                If sgUstWin(7) = "I" Then
                    '11/17/18: With the new pop-up screen, have all posting go thru it
                    'D.S. 05/11/12  added or to IF statement below - force the question
                    'If tmCPDate(llLoop).iPostingStatus = 0 Or tmCPDate(llLoop).iPostingStatus = 2 Then
                        'D.S. 09/10/02
                        'iRet = gMsgBox("All Spots Ran as Pledged", vbYesNoCancel + vbQuestion)
                        Screen.MousePointer = vbDefault
                        sgCPRetStatus = Trim$(cboPSSort.Text) & " " & Trim$(cboSSSort.Text) & " Week: " & grdPost.TextMatrix(lmEnableRow, lmEnableCol)
                        frmCPRetStatus.Show vbModal
                        If (sgCPRetStatus = "Yes") Or (sgCPRetStatus = "None") Then
                            If sgCPRetStatus = "None" Then
                                iRet = gMsgBox("Please Confirm that ""None Aired"" as Pledged.", vbYesNo)
                                If iRet = vbNo Then
                                    imStatus = ilSvStatus
                                    Screen.MousePointer = vbDefault
                                    cmdDone.SetFocus
                                    Exit Sub
                                End If
                            Else
                                iRet = vbYes
                            End If
                        ElseIf (sgCPRetStatus = "View") Then
                            iRet = vbNo
                        ElseIf (sgCPRetStatus = "Unpost") Then
                            'iRet = vbNo
                            iRet = gMsgBox("Please Confirm the selected action of ""Un-post"" Affiliate Spots.", vbYesNo)
                            If iRet = vbNo Then
                                imStatus = ilSvStatus
                                Screen.MousePointer = vbDefault
                                cmdDone.SetFocus
                                Exit Sub
                            End If
                            iRet = vbNo
                        ElseIf sgCPRetStatus = "No" Then
                            iRet = vbNo
                        End If
                    'Else
                    '    iRet = vbNo
                    'End If
                Else
                    iRet = vbNo
                End If
                ReDim tgCPPosting(0 To 1) As CPPOSTING
                tgCPPosting(0).lCpttCode = lCode
                tgCPPosting(0).lAttCode = tmCPDate(llLoop).lAttCode
                tgCPPosting(0).iAttTimeType = tmCPDate(llLoop).iAttTimeType
                If (optPSSort(0).Value = True) Or (optPSSort(1).Value = True) Then
                    tgCPPosting(0).iVefCode = tmCPVehicle(iRowIndex).iCode
                    tgCPPosting(0).iShttCode = imShttCode
                    tgCPPosting(0).sZone = tmCPVehicle(iRowIndex).sZone
                Else
                    tgCPPosting(0).iVefCode = imVefCode
                    tgCPPosting(0).iShttCode = tmCPVehicle(iRowIndex).iCode
                    tgCPPosting(0).sZone = tmCPVehicle(iRowIndex).sZone
                End If
                tgCPPosting(0).sDate = grdPost.TextMatrix(lmEnableRow, lmEnableCol)
                tgCPPosting(0).iStatus = tmCPDate(llLoop).iStatus
                tgCPPosting(0).iPostingStatus = tmCPDate(llLoop).iPostingStatus
                tgCPPosting(0).sAstStatus = tmCPDate(llLoop).sAstStatus
                Screen.MousePointer = vbDefault
                If tmCPDate(llLoop).iAttPostingType = 2 Then
                    igTimes = 1
                Else
                    igTimes = 0
                    slDate = tgCPPosting(0).sDate
                    ilPos = InStr(1, slDate, ",", vbTextCompare)
                    slDate = Left$(slDate, ilPos - 1) & ". 15" & Mid$(slDate, ilPos)
                    tgCPPosting(0).sDate = gObtainStartStd(Format$(slDate, "mm/dd/yyyy"))
                End If
                If iRet = vbNo Then
                    If (modAffiliate.sgCPRetStatus = "Unpost") Then
                        mUpdateAstAsNotPosted modCPReturns.tgCPPosting(0).lAttCode, modCPReturns.tgCPPosting(0).sDate
                        mUpdateCpttAsNotPosted modCPReturns.tgCPPosting(0).lAttCode, modCPReturns.tgCPPosting(0).sDate
                        tmCPDate(llLoop).iStatus = 0
                        tmCPDate(llLoop).iPostingStatus = 0
                        modCPReturns.tgCPPosting(0).iStatus = tmCPDate(llLoop).iStatus
                        modCPReturns.tgCPPosting(0).iPostingStatus = tmCPDate(llLoop).iPostingStatus
                    End If
                    igCPStatus = tmCPDate(llLoop).iStatus
                    igCPPostingStatus = tmCPDate(llLoop).iPostingStatus
                    If (sgCPRetStatus = "View") Then
                        sgUstWin(7) = "V"
                    End If
                    frmDateTimes.Show vbModal
                    If sgUstWin(7) <> "I" Then
                        imStatus = ilSvStatus
                        sgUstWin(7) = slSvUstWin7
                        Screen.MousePointer = vbDefault
                        cmdDone.SetFocus
                        Exit Sub
                    End If
                ElseIf iRet = vbYes Then
                    If lgSelGameGsfCode <= 0 Then
                        iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), -1, True, False, False, False, , , , , True)
                    Else
                        'iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), -1, True, False, False, False, lgSelGameGsfCode)
                        iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), -1, True, False, False, False, , , , , True)
                    End If
                    If sgCPRetStatus = "Yes" Then
                        igCPStatus = 1          '1 = Received
                    ElseIf sgCPRetStatus = "None" Then
                        igCPStatus = 2          'Not Aired
                    End If
                    If lgSelGameGsfCode <= 0 Then
                        'D.S. 09/11/02 Added If statement to support None Aired
                        igCPPostingStatus = 2   'Completed
                        If tmCPDate(llLoop).iAttPostingType = 2 Then 'By Time
                            SQLQuery = "UPDATE cptt SET "
                            SQLQuery = SQLQuery + "cpttStatus = " & igCPStatus & ", "
                            SQLQuery = SQLQuery + "cpttPostingStatus = " & igCPPostingStatus
                            '10/19/18: added setting user
                            SQLQuery = SQLQuery + ", " + "cpttUsfCode = " & igUstCode
                            SQLQuery = SQLQuery + " WHERE cpttCode = " & lCode & ""
                        Else    'By Advertiser
                            slDate = grdPost.TextMatrix(lmEnableRow, lmEnableCol)
                            ilPos = InStr(1, slDate, ",", vbTextCompare)
                            slDate = Left$(slDate, ilPos - 1) & ". 15" & Mid$(slDate, ilPos)
                            slPostSDate = gObtainStartStd(Format$(slDate, "mm/dd/yyyy"))
                            slPostEDate = Format$(gObtainEndStd(slDate), sgShowDateForm)
                            SQLQuery = "UPDATE cptt SET "
                            SQLQuery = SQLQuery + "cpttStatus = " & igCPStatus & ", "
                            SQLQuery = SQLQuery + "cpttPostingStatus = " & igCPPostingStatus
                            '10/19/18: added setting user
                            SQLQuery = SQLQuery + ", " + "cpttUsfCode = " & igUstCode
                            If (optPSSort(0).Value = True) Or (optPSSort(1).Value = True) Then
                                SQLQuery = SQLQuery + " WHERE (cpttVefCode = " & tmCPVehicle(iRowIndex).iCode
                                SQLQuery = SQLQuery + " AND cpttShfCode = " & imShttCode
                            Else
                                SQLQuery = SQLQuery + " WHERE (cpttVefCode = " & imVefCode
                                SQLQuery = SQLQuery + " AND cpttShfCode = " & tmCPVehicle(iRowIndex).iCode
                            End If
                            SQLQuery = SQLQuery + " AND (cpttStartDate >= '" & Format$(slPostSDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slPostEDate, sgSQLDateForm) & "')" & ")"
                        End If
                        'cnn.BeginTrans
                        'cnn.Execute SQLQuery, rdExecDirect
                        'cnn.CommitTrans
                        If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                            'Screen.MousePointer = vbDefault
                            'Exit Sub
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand:
                            imStatus = ilSvStatus
                            Screen.MousePointer = vbDefault
                            gHandleError "AffErrorLog.txt", "CPReturns-mClick"
                            Exit Sub
                        End If
                        gFileChgdUpdate "cptt.mkd", True
                        'Update AST for reports
                        If tmCPDate(llLoop).iAttPostingType = 2 Then
                            slDate = grdPost.TextMatrix(lmEnableRow, lmEnableCol)
                            slPostSDate = Format$(gObtainPrevMonday(slDate), sgShowDateForm)
                            slPostEDate = Format$(gObtainNextSunday(slDate), sgShowDateForm)
                        Else
                            slDate = grdPost.TextMatrix(lmEnableRow, lmEnableCol)
                            slPostSDate = Format$(gObtainPrevMonday(slDate), sgShowDateForm)
                            slPostEDate = Format$(gObtainEndStd(slDate), sgShowDateForm)
                        End If
                        SQLQuery = "UPDATE ast SET "
                        If igCPStatus = 2 Then
                            'Set status to NA-Other
                            '12/6/13: Moved below as status should NOT be set if pledge is Not Carried
                            'SQLQuery = SQLQuery + "astStatus = " & "4" & ", "
                            SQLQuery = SQLQuery + "astCPStatus = " & "2"
                        Else
                            SQLQuery = SQLQuery + "astCPStatus = " & "1"
                        End If
                        '10/19/18: added setting user
                        SQLQuery = SQLQuery + ", " & "astUstCode = " & igUstCode
                        SQLQuery = SQLQuery + " WHERE (astAtfCode = " & tmCPDate(llLoop).lAttCode
                        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slPostSDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slPostEDate, sgSQLDateForm) & "')" & ")"
                        'cnn.BeginTrans
                        'cnn.Execute SQLQuery, rdExecDirect
                        'cnn.CommitTrans
                        If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                            'Screen.MousePointer = vbDefault
                            'Exit Sub
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand:
                            imStatus = ilSvStatus
                            Screen.MousePointer = vbDefault
                            gHandleError "AffErrorLog.txt", "CPReturns-mClick"
                            Exit Sub
                        End If
                        '12/6/13: Moved here from above because the PledgeStatus of 8 should NOT have the astStatus set
                        If igCPStatus = 2 Then
                            '12/13/13: Pledge obtained from DAT
                            'SQLQuery = "UPDATE ast SET "
                            'SQLQuery = SQLQuery + "astStatus = " & "4"
                            'SQLQuery = SQLQuery + " WHERE (astAtfCode = " & tmCPDate(llLoop).lAttCode
                            'SQLQuery = SQLQuery + " AND astPledgeStatus <> 8"
                            'SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slPostSDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slPostEDate, sgSQLDateForm) & "')" & ")"
                            ''cnn.BeginTrans
                            ''cnn.Execute SQLQuery, rdExecDirect
                            ''cnn.CommitTrans
                            'If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                            '    'Screen.MousePointer = vbDefault
                            '    'Exit Sub
                            '    GoSub ErrHand
                            'End If
                            mSetStatusNotAired tmCPDate(llLoop).lAttCode, 0, slPostSDate, slPostEDate
                        Else
                            mSetToPledgeStatus tmCPDate(llLoop).lAttCode, 0, slPostSDate, slPostEDate
                        End If
                        gSetCpttCount tmCPDate(llLoop).lAttCode, slPostSDate, slPostEDate
                    Else
                        slDate = grdPost.TextMatrix(lmEnableRow, lmEnableCol)
                        slPostSDate = Format$(gObtainPrevMonday(slDate), sgShowDateForm)
                        slPostEDate = Format$(gObtainNextSunday(slDate), sgShowDateForm)
                        SQLQuery = "UPDATE ast SET "
                        If igCPStatus = 2 Then
                            'Set status to NA-Other
                            '12/6/13: Moved below as status should NOT be set if pledge is Not Carried
                            'SQLQuery = SQLQuery + "astStatus = " & "4" & ", "
                            SQLQuery = SQLQuery + "astCPStatus = " & "2"
                        Else
                            SQLQuery = SQLQuery + "astCPStatus = " & "1"
                        End If
                        '5/29/18: replace sql call as the call below failed because it retrieve too many records
                        'SQLQuery = SQLQuery + " FROM ast LEFT OUTER JOIN lst On astLsfCode = lstCode"
                        'SQLQuery = SQLQuery + " WHERE (astAtfCode = " & tmCPDate(llLoop).lAttCode
                        'SQLQuery = SQLQuery + " AND lstGsfCode = " & lgSelGameGsfCode
                        'SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slPostSDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slPostEDate, sgSQLDateForm) & "')" & ")"
                        '10/19/18: added setting user
                        SQLQuery = SQLQuery + ", " & "astUstCode = " & igUstCode
                        SQLQuery = SQLQuery + " Where astAtfCode = " & tmCPDate(llLoop).lAttCode
                        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slPostSDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slPostEDate, sgSQLDateForm) & "')"
                        SQLQuery = SQLQuery + " And astCode In (Select astCode From ast Left Outer Join lst on astlsfCode = lstCode Where (astAtfCode = " & tmCPDate(llLoop).lAttCode
                        SQLQuery = SQLQuery + " AND lstGsfCode = " & lgSelGameGsfCode
                        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slPostSDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slPostEDate, sgSQLDateForm) & "')" & "))"
                        
                        If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                            'Exit Sub
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand:
                            imStatus = ilSvStatus
                            Screen.MousePointer = vbDefault
                            gHandleError "AffErrorLog.txt", "CPReturns-mClick"
                            Exit Sub
                        End If
                        '12/6/13: Moved here from above because the PledgeStatus of 8 should NOT have the astStatus set
                        If igCPStatus = 2 Then
                            '12/13/13: Pledge obtained from DAT
                            'SQLQuery = "UPDATE ast SET "
                            'SQLQuery = SQLQuery + "astStatus = " & "4"
                            'SQLQuery = SQLQuery + " FROM ast LEFT OUTER JOIN lst On astLsfCode = lstCode"
                            'SQLQuery = SQLQuery + " WHERE (astAtfCode = " & tmCPDate(llLoop).lAttCode
                            'SQLQuery = SQLQuery + " AND astPledgeStatus <> 8"
                            'SQLQuery = SQLQuery + " AND lstGsfCode = " & lgSelGameGsfCode
                            'SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slPostSDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slPostEDate, sgSQLDateForm) & "')" & ")"
                            'If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                            '    'Exit Sub
                            '    GoSub ErrHand
                            'End If
                            mSetStatusNotAired tmCPDate(llLoop).lAttCode, lgSelGameGsfCode, slPostSDate, slPostEDate
                        Else
                            mSetToPledgeStatus tmCPDate(llLoop).lAttCode, lgSelGameGsfCode, slPostSDate, slPostEDate
                        End If
                        iRet = mUpdateCptt(tmCPDate(llLoop).lAttCode, slPostSDate, slPostEDate)
                        gSetCpttCount tmCPDate(llLoop).lAttCode, slPostSDate, slPostEDate
                    End If
                Else
                    igCPStatus = imStatus
                    igCPPostingStatus = tmCPDate(llLoop).iPostingStatus
                End If
                'iRet = gMsgBox("Posting Completed", vbYesNo + vbDefaultButton1 + vbQuestion, "Posting Completed")
                'If iRet = vbYes Then    ' User chose Yes.
                '    imStatus = 1
                'Else
                '    imStatus = 0
                'End If
                Screen.MousePointer = vbHourglass
                'If tmCPDate(llLoop).iStatus <> imStatus Then
                If (igCPStatus <> imStatus) Or (tmCPDate(llLoop).iPostingStatus <> igCPPostingStatus) Then
                    If (igCPStatus = 0) Then
                        If igCPPostingStatus = 1 Then
                            grdPost.Col = lmEnableCol
                            grdPost.CellForeColor = vbMagenta
                        Else
                            grdPost.Col = lmEnableCol
                            grdPost.CellForeColor = vbRed
                        End If
                        If imStatus <> 0 Then
                            If optStatus(0).Value = True Then
                                grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) = grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) + 1
                            ElseIf optStatus(1).Value = True Then
                                grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) = grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) - 1
                            End If
                        End If
                        imStatus = igCPStatus
                    'D.S. 09/11/02 Added igCPStatus = 2 to support None Aired
                    ElseIf (igCPStatus = 1) Or (igCPStatus = 2) Then
                        grdPost.Col = lmEnableCol
                        grdPost.CellForeColor = RGB(0, 128, 0)
                        If imStatus = 0 Then
                            If optStatus(0).Value = True Then
                                grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) = grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) - 1
                            ElseIf optStatus(1).Value = True Then
                                grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) = grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) + 1
                            End If
                        End If
                        imStatus = igCPStatus
                    End If
                End If
                tmCPDate(llLoop).iStatus = imStatus
                tmCPDate(llLoop).iPostingStatus = igCPPostingStatus
                mGridPaint False
            Else
                If sgUstWin(7) <> "I" Then
                    imStatus = ilSvStatus
                    Screen.MousePointer = vbDefault
                    cmdDone.SetFocus
                    Exit Sub
                End If
                If tmCPDate(llLoop).iStatus = 0 Then          'iCellColor = 0 Then
                    grdPost.Col = lmEnableCol
                    grdPost.CellForeColor = RGB(0, 128, 0)
                    imStatus = 1
                    '7/20/12: Set for receipt only agreements
                    ilCPPostingStatus = 2
                    If optStatus(0).Value = True Then
                        grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) = grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) - 1
                        sDate = tmCPDate(llLoop).sDate
                    ElseIf optStatus(1).Value = True Then
                        grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) = grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) + 1
                        sDate = tmCPDate(llLoop).sDate
                    Else
                        sDate = tmCPDate(llLoop).sDate
                    End If
                ElseIf tmCPDate(llLoop).iStatus = 1 Then
                    grdPost.Col = lmEnableCol
                    grdPost.CellForeColor = vbBlue
                    imStatus = 2
                    '7/20/12: Set for receipt only agreements
                    ilCPPostingStatus = 2
                    If optStatus(1).Value = True Then
                        grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) = grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) - 1
                        sDate = tmCPDate(llLoop).sDate
                    ElseIf optStatus(2).Value = True Then
                        grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) = grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) + 1
                        sDate = tmCPDate(llLoop).sDate
                    Else
                        sDate = tmCPDate(llLoop).sDate
                    End If
                ElseIf tmCPDate(llLoop).iStatus = 2 Then
                    grdPost.Col = lmEnableCol
                    grdPost.CellForeColor = vbRed
                    imStatus = 0
                    '7/20/12: Set for receipt only agreements
                    ilCPPostingStatus = 0
                    If optStatus(0).Value = True Then
                        grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) = grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) + 1
                        sDate = tmCPDate(llLoop).sDate
                    ElseIf optStatus(2).Value = True Then
                        grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) = grdPost.TextMatrix(lmEnableRow, NOMISSINGINDEX) - 1
                        sDate = tmCPDate(llLoop).sDate
                    Else
                        sDate = tmCPDate(llLoop).sDate
                    End If
                End If
                '7/20/12: Set for receipt only agreements
                SQLQuery = "UPDATE cptt SET "
                SQLQuery = SQLQuery + "cpttStatus = " & imStatus & ", "
                SQLQuery = SQLQuery + "cpttPostingStatus = " & ilCPPostingStatus
                '10/19/18: added setting user
                SQLQuery = SQLQuery + ", " + "cpttUsfCode = " & igUstCode
                SQLQuery = SQLQuery + " WHERE cpttCode = " & lCode & ""
                
                'cnn.BeginTrans
                'cnn.Execute SQLQuery, rdExecDirect
                'cnn.CommitTrans
                If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                    'Screen.MousePointer = vbDefault
                    'Exit Sub
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    imStatus = ilSvStatus
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "CPReturns-mClick"
                    Exit Sub
                End If
                gFileChgdUpdate "cptt.mkd", True
                tmCPDate(llLoop).iStatus = imStatus
                '7/20/12: Set for receipt only agreements
                tmCPDate(llLoop).iPostingStatus = ilCPPostingStatus
                grdPost.TextMatrix(lmEnableRow, lmEnableCol) = sDate
                Screen.MousePointer = vbDefault
            End If
            Exit For
        End If
    Next llLoop
    imStatus = ilSvStatus
    Screen.MousePointer = vbDefault
    cmdDone.SetFocus
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CPReturns-mClick"
    Resume Next
End Sub

Private Sub rbcVeh_Click(Index As Integer)
    cboPSSort_Change
End Sub

Private Sub tmcDelay_Timer()
    tmcDelay.Enabled = False
    lmEnableRow = grdPost.Row
    lmEnableCol = grdPost.Col
    mClick
    imDateClick = False
End Sub

Private Sub mDateSort()
    Dim llTotRec As Long
    Dim llSIndex As Long
    Dim llEIndex As Long
    Dim iCol As Integer
    Dim llLoop As Long
    Dim llRow As Long
    'ReDim tCPDate(1 To 1) As CPDATE
    ReDim tCPDate(0 To 0) As CPDATE
    
    For llRow = 0 To UBound(tmCPVehicle) - 1 Step 1
        llSIndex = tmCPVehicle(llRow).lCPDateIndex
        If llRow < UBound(tmCPVehicle) - 1 Then
            llEIndex = tmCPVehicle(llRow + 1).lCPDateIndex - 1
        Else
            llEIndex = UBound(tmCPDate) - 1
        End If
        llTotRec = llEIndex - llSIndex + 1
        If llTotRec > 1 Then
            'ReDim tCPDate(1 To llTotRec) As CPDATE
            ReDim tCPDate(0 To llTotRec - 1) As CPDATE
            'iCol = 1
            iCol = 0
            For llLoop = llEIndex To llSIndex Step -1
                tCPDate(iCol) = tmCPDate(llLoop)
                iCol = iCol + 1
            Next llLoop
            'iCol = 1
            iCol = 0
            For llLoop = llSIndex To llEIndex Step 1
                tmCPDate(llLoop) = tCPDate(iCol)
                iCol = iCol + 1
            Next llLoop
        End If
    Next llRow
    Erase tCPDate
End Sub

Private Sub tmcGrid_Timer()
    tmcGrid.Enabled = False
    grdPost.Enabled = True
End Sub

Private Sub tmcSSSort_Timer()
    tmcSSSort.Enabled = False
    If optStatus(0).Value = True Then
        imStatus = 0
    ElseIf optStatus(1).Value = True Then
        imStatus = 1
    Else
        imStatus = 2
    End If
    imCPIndex = 1
    Screen.MousePointer = vbHourglass
    mGetCPTT
    Screen.MousePointer = vbDefault
End Sub

Private Sub mGetCPTT()
    Dim X, Y As Integer
    Dim sVehicles As String
    Dim sStatus As String
    Dim sStations As String
    Dim iMonth As Integer
    Dim sAirTime As String
    Dim iTotRec As Integer
    Dim iMissRec As Integer
    Dim iRemainder As Integer
    Dim iNumSets As Integer
    Dim iDateSet As Integer
    Dim iCode As Integer
    Dim iTestCode As Integer
    Dim bNewVehicle As Boolean
    Dim lRow As Long
    Dim iCol As Integer
    Dim lUpper As Long
    Dim i As Integer
    Dim iAdd As Integer
    Dim iFirst As Integer
    Dim sZone As String
    Dim sTestDate As String
    Dim sMonthDate As String
    Dim sMarket As String
    Dim sShttZone As String
    Dim llStaCode As Long
    Dim llRet As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slSQLDate As String
    
    On Error GoTo ErrHand
    
    
    If (imShttCode < 0) Or (imVefCode < 0) Then
        Exit Sub
    End If
    
    slStartDate = edcStartDate.Text
    slEndDate = edcEndDate.Text
    slSQLDate = ""
    If (slStartDate <> "") Then
        If (gIsDate(slStartDate) = False) Then
            Beep
            edcStartDate.SetFocus
            Exit Sub
        End If
        slSQLDate = "cpttStartDate >= '" & Format(slStartDate, sgSQLDateForm) & "'"
    End If
    If (slEndDate <> "") Then
        If (gIsDate(slEndDate) = False) Then
            Beep
            edcEndDate.SetFocus
            Exit Sub
        End If
        If (slSQLDate = "") Then
            slSQLDate = "cpttStartDate <= '" & Format(slEndDate, sgSQLDateForm) & "'"
        Else
            slSQLDate = slSQLDate & " And " & "cpttStartDate <= '" + Format(slEndDate, modAffiliate.sgSQLDateForm) & "'"
        
        End If
    End If
   
    If (optPSSort(0).Value = True) Or (optPSSort(1).Value = True) Then
        SQLQuery = "SELECT shttTimeZone FROM shtt WHERE (shttCode = " & imShttCode & ")"
        Set rst = gSQLSelectCall(SQLQuery)
        If rst.EOF = True Then
            sZone = ""
        Else
            sZone = Trim$(rst!shttTimeZone)
        End If
        If imVefCode = 0 Then    'Select all vehicles
            SQLQuery = "SELECT vefType, vefName, vefCode, cpttStartDate, cpttAtfCode, cpttPostingStatus, cpttNoSpotsGen, cpttNoSpotsAired, cpttAstStatus, cpttCode, attPostingType, attTimeType"
            'SQLQuery = SQLQuery + " FROM VEF_Vehicles vef, cptt, att"
            SQLQuery = SQLQuery & " FROM VEF_Vehicles, cptt, att"
            SQLQuery = SQLQuery + " WHERE (vefCode = cpttVefCode"
            SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
            SQLQuery = SQLQuery + " AND cpttShfCode = " & imShttCode
            SQLQuery = SQLQuery + " AND cpttStatus = " & imStatus
            SQLQuery = SQLQuery + ")"
            If imStatus = 0 Then
                SQLQuery = SQLQuery + " ORDER BY vefName, cpttStartDate"
            Else
                SQLQuery = SQLQuery + " ORDER BY vefName, cpttStartDate desc"
            End If
        Else                        'Select one vehicles
            SQLQuery = "SELECT vefType, vefName, vefCode, cpttStartDate, cpttAtfCode, cpttPostingStatus, cpttNoSpotsGen, cpttNoSpotsAired, cpttAstStatus, cpttCode, attPostingType, attTimeType"
            'SQLQuery = SQLQuery + " FROM VEF_Vehicles vef, cptt, att"
            SQLQuery = SQLQuery & " FROM VEF_Vehicles, cptt, att"
            SQLQuery = SQLQuery + " WHERE (cpttVefCode = " & imVefCode
            SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
            SQLQuery = SQLQuery + " AND vefCode = " & imVefCode
            SQLQuery = SQLQuery + " AND cpttShfCode = " & imShttCode & ""
            SQLQuery = SQLQuery + " AND cpttStatus = " & imStatus
            SQLQuery = SQLQuery + ")"
            If imStatus = 0 Then
                SQLQuery = SQLQuery + " ORDER BY vefName, cpttStartDate"
            Else
                SQLQuery = SQLQuery + " ORDER BY vefName, cpttStartDate desc"
            End If
        End If
    Else
        If imShttCode = 0 Then    'Select all Stations
            SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttPostingStatus, cpttNoSpotsGen, cpttNoSpotsAired, cpttAstStatus, cpttCode, attPostingType, attTimeType"
            SQLQuery = SQLQuery + " FROM shtt, cptt, att"
            SQLQuery = SQLQuery + " WHERE (ShttCode = cpttShfCode"
            SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
            SQLQuery = SQLQuery + " AND cpttVefCode = " & imVefCode
            SQLQuery = SQLQuery + " AND cpttStatus = " & imStatus & " AND shttType = 0"
            SQLQuery = SQLQuery + ")"
            If imStatus = 0 Then
                SQLQuery = SQLQuery + " ORDER BY shttCallLetters, cpttStartDate"
            Else
                SQLQuery = SQLQuery + " ORDER BY shttCallLetters, cpttStartDate desc"
            End If
        Else                        'Select one vehicles
            SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttPostingStatus, cpttNoSpotsGen, cpttNoSpotsAired, cpttAstStatus, cpttCode, attPostingType, attTimeType"
            SQLQuery = SQLQuery + " FROM shtt, cptt, att"
            SQLQuery = SQLQuery + " WHERE (cpttVefCode = " & imVefCode
            SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
            SQLQuery = SQLQuery + " AND shttCode = " & imShttCode
            SQLQuery = SQLQuery + " AND cpttShfCode = " & imShttCode
            SQLQuery = SQLQuery + " AND cpttStatus = " & imStatus & " AND shttType = 0"
            SQLQuery = SQLQuery + ")"
            If imStatus = 0 Then
                SQLQuery = SQLQuery + " ORDER BY shttCallLetters, cpttStartDate"
            Else
                SQLQuery = SQLQuery + " ORDER BY shttCallLetters, cpttStartDate desc"
            End If
        End If
    End If
    If slSQLDate <> "" Then
        SQLQuery = Replace(SQLQuery, ")", " And " & slSQLDate & ")")
    End If
    Set rst = gSQLSelectCall(SQLQuery)
    
    mClearGrid
    imCPMax = 1
    i = 0
    lRow = 0
    iCol = 0
    ReDim tmCPDate(0 To 1000) As CPDATE
    ReDim tmCPVehicle(0 To 0) As CPVEHICLE
    imCPIndex = 1
    lUpper = 0
    If rst.EOF = True Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If (optPSSort(0).Value = True) Or (optPSSort(1).Value = True) Then
        iCode = rst!vefCode 'rst(1).Value
        If sgShowByVehType = "Y" Then
            tmCPVehicle(0).sName = Trim$(rst!vefType) & ":" & rst!vefName   'rst(0).Value
        Else
            tmCPVehicle(0).sName = rst!vefName   'rst(0).Value
        End If
        tmCPVehicle(0).sMarket = ""
        tmCPVehicle(0).sZone = sZone
    Else
        iCode = rst!shttCode 'rst(1).Value
        tmCPVehicle(0).sName = rst!shttCallLetters  'rst(0).Value
        llStaCode = gBinarySearchStation(Trim$(rst!shttCallLetters))
        If llStaCode <> -1 Then
            sMarket = tgStationInfo(llStaCode).sMarket
        Else
            sMarket = ""
        End If
        tmCPVehicle(0).sMarket = sMarket
        If IsNull(rst!shttTimeZone) = True Then
            sShttZone = ""
        Else
            sShttZone = rst!shttTimeZone  'Trim$(rst!shttMarket)
        End If
        tmCPVehicle(0).sZone = sShttZone
    End If
    iFirst = True
    tmCPVehicle(0).lRow = lRow
    tmCPVehicle(0).lCPDateIndex = lUpper
    tmCPVehicle(0).iCode = iCode
    ReDim Preserve tmCPVehicle(0 To 1) As CPVEHICLE
    While Not rst.EOF
        If (optPSSort(0).Value = True) Or (optPSSort(1).Value = True) Then
            iTestCode = rst!vefCode 'rst(1).Value
        Else
            iTestCode = rst!shttCode 'rst(1).Value
            llStaCode = gBinarySearchStation(Trim$(rst!shttCallLetters))
            If llStaCode <> -1 Then
                sMarket = tgStationInfo(llStaCode).sMarket
            Else
                sMarket = ""
            End If
        End If
        If iCode <> iTestCode Then
            iFirst = True
            If iCol \ 7 + 1 > imCPMax Then
                imCPMax = iCol \ 7 + 1
            End If
            iCode = iTestCode   'rst(1).Value
            lRow = lRow + 1
            iCol = 0
            tmCPVehicle(UBound(tmCPVehicle)).lRow = lRow
            tmCPVehicle(UBound(tmCPVehicle)).lCPDateIndex = lUpper
            If (optPSSort(0).Value = True) Or (optPSSort(1).Value = True) Then
                If sgShowByVehType = "Y" Then
                    tmCPVehicle(UBound(tmCPVehicle)).sName = Trim$(rst!vefType) & ":" & rst!vefName  'rst(0).Value
                Else
                    tmCPVehicle(UBound(tmCPVehicle)).sName = rst!vefName  'rst(0).Value
                End If
                tmCPVehicle(UBound(tmCPVehicle)).sMarket = ""
                tmCPVehicle(UBound(tmCPVehicle)).sZone = sZone
            Else
                tmCPVehicle(UBound(tmCPVehicle)).sName = rst!shttCallLetters  'rst(0).Value
                tmCPVehicle(UBound(tmCPVehicle)).sMarket = sMarket
                If IsNull(rst!shttTimeZone) = True Then
                    sShttZone = ""
                Else
                    sShttZone = rst!shttTimeZone  'Trim$(rst!shttMarket)
                End If
                tmCPVehicle(UBound(tmCPVehicle)).sZone = sShttZone
            End If
            tmCPVehicle(UBound(tmCPVehicle)).iCode = iCode    'rst(1).Value
            ReDim Preserve tmCPVehicle(0 To UBound(tmCPVehicle) + 1) As CPVEHICLE
        End If
        If rst!attPostingType = 3 Then  '3=spots by advertiser
            'Combine into months
            If Not iFirst Then
                sTestDate = gObtainEndStd(rst!CpttStartDate)
                If DateValue(gAdjYear(sTestDate)) <> DateValue(gAdjYear(sMonthDate)) Then
                    sMonthDate = gObtainEndStd(rst!CpttStartDate)
                    iAdd = True
                Else
                    iAdd = False
                End If
            Else
                sMonthDate = gObtainEndStd(rst!CpttStartDate)
                iAdd = True
            End If
            iFirst = False
        Else
            iAdd = True
        End If
        If iAdd Then
            tmCPDate(lUpper).iCol = iCol
            tmCPDate(lUpper).iStatus = imStatus
            tmCPDate(lUpper).lCpttCode = rst!cpttCode    'rst(3).Value
            If rst!attPostingType = 3 Then  '3=spots by advertiser
                tmCPDate(lUpper).sDate = sMonthDate   'rst(2).Value
            Else
                tmCPDate(lUpper).sDate = Format$(rst!CpttStartDate, sgShowDateForm)   'rst(2).Value
            End If
            tmCPDate(lUpper).iAttPostingType = rst!attPostingType
            tmCPDate(lUpper).iAttTimeType = rst!attTimeType
            tmCPDate(lUpper).lAttCode = rst!cpttatfCode    'rst(3).Value
            tmCPDate(lUpper).iPostingStatus = rst!cpttPostingStatus    'rst(3).Value
            tmCPDate(lUpper).iNoSpotsGen = rst!cpttNoSpotsGen
            tmCPDate(lUpper).iNoSpotsAired = rst!cpttNoSpotsAired
            tmCPDate(lUpper).sAstStatus = rst!cpttAstStatus
            lUpper = lUpper + 1
            iCol = iCol + 1
            If lUpper > UBound(tmCPDate) Then
                ReDim Preserve tmCPDate(0 To lUpper + 999) As CPDATE
            End If
        End If
        rst.MoveNext
    Wend
    ReDim Preserve tmCPDate(0 To lUpper) As CPDATE
    If iCol \ 7 + 1 > imCPMax Then
        imCPMax = iCol \ 7 + 1
    End If

    mGridPaint True
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CPReturns-mGetCptt"
    Exit Sub
End Sub

Private Function mUpdateCptt(llAttCode As Long, slSDate As String, slEDate As String) As Integer
    Dim slMoDate As String
    Dim slSuDate As String
    Dim ilSpotsAired As Integer
    
    On Error GoTo ErrHand:
    slMoDate = gAdjYear(gObtainPrevMonday(slSDate))
    Do
        slSuDate = DateAdd("d", 6, slMoDate)
        'Test to see if any spots aired or were they all not aired
        ilSpotsAired = gDidAnySpotsAir(llAttCode, slMoDate, slSuDate)
        If ilSpotsAired Then
            'We know at least one spot aired
           ilSpotsAired = True
        Else
            'no aired spots were found
            ilSpotsAired = False
        End If
        
        'Check for any spots that have not aired - astCPStatus = 0 = not aired
        SQLQuery = "Select astCode FROM ast WHERE astCPStatus = 0"
        SQLQuery = SQLQuery + " AND astAtfCode = " & llAttCode
        'SQLQuery = SQLQuery + " AND astShfCode = " & tmAirSpotInfo(ilLoop).iShfCode
        'SQLQuery = SQLQuery + " AND astVefCode = " & tmAirSpotInfo(ilLoop).iVefCode
        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
        Set rst = gSQLSelectCall(SQLQuery)
        If rst.EOF Then
            'Set CPTT as complete
            SQLQuery = "UPDATE cptt SET "
            If ilSpotsAired Then
                SQLQuery = SQLQuery + "cpttStatus = 1" & ", " 'Complete spots aired
            Else
                SQLQuery = SQLQuery + "cpttStatus = 2" & ", " 'Complete NO spots aired
            End If
            SQLQuery = SQLQuery + "cpttPostingStatus = 2"  'Complete
            '10/19/18: added setting user
            SQLQuery = SQLQuery + ", " + "cpttUsfCode = " & igUstCode
            SQLQuery = SQLQuery + " WHERE cpttAtfCode = " & llAttCode
            'SQLQuery = SQLQuery + " AND cpttShfCode = " & tmAirSpotInfo(ilLoop).iShfCode
            'SQLQuery = SQLQuery + " AND cpttVefCode = " & tmAirSpotInfo(ilLoop).iVefCode
            SQLQuery = SQLQuery + " AND (cpttStartDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            'cnn.BeginTrans
            ''cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                'GoSub ErrHand:
                gHandleError "AffErrorLog.txt", "Post CP-mUpdateCPTT"
                mUpdateCptt = False
                Exit Function
            End If
            'cnn.CommitTrans
        Else
            'Set CPTT as Partial
            igCPStatus = 0
            igCPPostingStatus = 1
            SQLQuery = "UPDATE cptt SET "
            SQLQuery = SQLQuery + "cpttStatus = 0" & ", " 'Partial
            SQLQuery = SQLQuery + "cpttPostingStatus = 1" 'Partial
            '10/19/18: added setting user
            SQLQuery = SQLQuery + ", " + "cpttUsfCode = " & igUstCode
            SQLQuery = SQLQuery + " WHERE cpttAtfCode = " & llAttCode
            'SQLQuery = SQLQuery + " AND cpttShfCode = " & tmAirSpotInfo(ilLoop).iShfCode
            'SQLQuery = SQLQuery + " AND cpttVefCode = " & tmAirSpotInfo(ilLoop).iVefCode
            SQLQuery = SQLQuery + " AND (cpttStartDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            'cnn.BeginTrans
            ''cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                'GoSub ErrHand1:
                gHandleError "AffErrorLog.txt", "Post CP-mUpdateCPTT"
                mUpdateCptt = False
                Exit Function
            End If
            'cnn.CommitTrans
        End If
        slMoDate = DateAdd("d", 7, slMoDate)
    Loop While DateValue(gAdjYear(slMoDate)) < DateValue(gAdjYear(slEDate))
    gFileChgdUpdate "cptt.mkd", True
    mUpdateCptt = True
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Post CP-mUpdateCPTT"
    mUpdateCptt = False
    Exit Function
ErrHand1:
    gHandleError "AffErrorLog.txt", "Post CP-mUpdateCPTT"
    mUpdateCptt = False
    Exit Function
End Function

Private Sub mSetStatusNotAired(llAttCode As Long, llGsfCode As Long, slStartDate As String, slEndDate As String)
    Dim tlDatPledgeInfo As DATPLEDGEINFO
    Dim ilRet As Integer
    
    On Error GoTo ErrHand:

    If llGsfCode <= 0 Then
        SQLQuery = "Select * FROM ast WHERE "
        SQLQuery = SQLQuery + " astAtfCode = " & llAttCode
        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slEndDate, sgSQLDateForm) & "')"
    Else
        SQLQuery = "Select * FROM ast LEFT OUTER JOIN lst On astLsfCode = lstCode"
        SQLQuery = SQLQuery + " WHERE (astAtfCode = " & llAttCode
        SQLQuery = SQLQuery + " AND lstGsfCode = " & llGsfCode
        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slEndDate, sgSQLDateForm) & "')" & ")"
    End If
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
    
        '12/13/13: Obtain Pledge information from Dat
        tlDatPledgeInfo.lAttCode = rst!astAtfCode
        tlDatPledgeInfo.lDatCode = rst!astDatCode
        tlDatPledgeInfo.iVefCode = rst!astVefCode
        tlDatPledgeInfo.sFeedDate = Format(rst!astFeedDate, "m/d/yy")
        tlDatPledgeInfo.sFeedTime = Format(rst!astFeedTime, "hh:mm:ssam/pm")
        ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)
        If tlDatPledgeInfo.iPledgeStatus <> 8 Then
            SQLQuery = "UPDATE ast SET "
            SQLQuery = SQLQuery + "astStatus = " & "4"
            '10/19/18: added setting user
            SQLQuery = SQLQuery + ", " & "astUstCode = " & igUstCode
            SQLQuery = SQLQuery + " WHERE (astCode = " & rst!astCode & ")"
            If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                'GoSub ErrHand1
                gHandleError "AffErrorLog.txt", "Post CP-mSetStatus"
            End If
        End If
        rst.MoveNext
    Loop
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Post CP-mSetStatus"
    Exit Sub
ErrHand1:
    gHandleError "AffErrorLog.txt", "Post CP-mSetStatus"
    Return
End Sub
Private Sub mSetToPledgeStatus(llAttCode As Long, llGsfCode As Long, slStartDate As String, slEndDate As String)
    Dim tlDatPledgeInfo As DATPLEDGEINFO
    Dim ilRet As Integer
    
    On Error GoTo ErrHand:

    If llGsfCode <= 0 Then
        SQLQuery = "Select * FROM ast WHERE "
        SQLQuery = SQLQuery + " astAtfCode = " & llAttCode
        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slEndDate, sgSQLDateForm) & "')"
    Else
        SQLQuery = "Select * FROM ast LEFT OUTER JOIN lst On astLsfCode = lstCode"
        SQLQuery = SQLQuery + " WHERE (astAtfCode = " & llAttCode
        SQLQuery = SQLQuery + " AND lstGsfCode = " & llGsfCode
        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slEndDate, sgSQLDateForm) & "')" & ")"
    End If
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
    
        '12/13/13: Obtain Pledge information from Dat
        tlDatPledgeInfo.lAttCode = rst!astAtfCode
        tlDatPledgeInfo.lDatCode = rst!astDatCode
        tlDatPledgeInfo.iVefCode = rst!astVefCode
        tlDatPledgeInfo.sFeedDate = Format(rst!astFeedDate, "m/d/yy")
        tlDatPledgeInfo.sFeedTime = Format(rst!astFeedTime, "hh:mm:ssam/pm")
        ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)
        If tlDatPledgeInfo.iPledgeStatus <> 8 Then
            SQLQuery = "UPDATE ast SET "
            SQLQuery = SQLQuery & " astStatus = " & "Case When astDatCode <= 0 Then 0 Else (Select datFdStatus From dat Where datCode = astDatCode) End"
            '10/19/18: added setting user
            SQLQuery = SQLQuery + ", " & "astUstCode = " & igUstCode
            SQLQuery = SQLQuery + " WHERE (astCode = " & rst!astCode & ")"
            If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                'GoSub ErrHand1
                gHandleError "AffErrorLog.txt", "Post CP-mSetStatus"
            End If
        End If
        rst.MoveNext
    Loop
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Post CP-mSetStatus"
    Exit Sub
ErrHand1:
    gHandleError "AffErrorLog.txt", "Post CP-mSetStatus"
    Return
End Sub
Private Sub mUpdateCpttAsNotPosted(llAttCode As Long, slStartDate As String)
    Dim SQLQuery As String
    '7895
    SQLQuery = "UPDATE cptt set "
    SQLQuery = SQLQuery & " cpttStatus = 0,"
    SQLQuery = SQLQuery & " cpttPostingStatus = 0,"
    SQLQuery = SQLQuery & " cpttNoSpotsGen = 0,"
    SQLQuery = SQLQuery & " cpttNoSpotsAired = 0,"
    SQLQuery = SQLQuery & " cpttUsfCode = " & igUstCode
    SQLQuery = SQLQuery & " WHERE cpttAtfCode = " & llAttCode
    SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(slStartDate, sgSQLDateForm) & "'"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError "AffErrorLog.txt", "Post CP-mUpdateCpttAsNotPosted"
        Exit Sub
    End If
    gFileChgdUpdate "cptt.mkd", True
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Post CP-mUpdateCpttAsNotPosted"
End Sub
Private Sub mUpdateAstAsNotPosted(llAttCode As Long, slStartDate As String)
    Dim SQLQuery As String
    '7895
    SQLQuery = "UPDATE ast set astCPStatus = 0,"
    SQLQuery = SQLQuery & " astStationCompliant = '" & "" & "',"
    SQLQuery = SQLQuery & " astAgencyCompliant = '" & "" & "',"
    SQLQuery = SQLQuery & " astAffidavitSource = '" & "" & "',"
    SQLQuery = SQLQuery & " astUstCode = " & igUstCode & ","
    SQLQuery = SQLQuery & " astStatus = " & "Case When astDatCode <= 0 Then 0 Else (Select datFdStatus From dat Where datCode = astDatCode) End"
    SQLQuery = SQLQuery & " WHERE astAtfCode = " & llAttCode
    SQLQuery = SQLQuery & " AND astFeedDate >= '" & Format$(slStartDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " AND astFeedDate <= '" & Format$(DateAdd("d", 6, slStartDate), sgSQLDateForm) & "'"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError "AffErrorLog.txt", "Post CP-tmUpdateAstAsNotPosted"
        Exit Sub
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Post CP-mUpdateAstAsNotPosted"
End Sub
