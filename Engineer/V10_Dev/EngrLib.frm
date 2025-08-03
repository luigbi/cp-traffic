VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form EngrLib 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrLib.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11790
   Begin V10EngineeringDev.CSI_Calendar cccEffEndDate 
      Height          =   285
      Left            =   6795
      TabIndex        =   4
      Top             =   120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
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
      CSI_CurDayForeColor=   0
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   -1  'True
      CSI_DefaultDateType=   1
   End
   Begin VB.CheckBox ckcDormant 
      Caption         =   "Include Dormant"
      Height          =   210
      Left            =   150
      TabIndex        =   12
      Top             =   6600
      Width           =   1770
   End
   Begin V10EngineeringDev.CSI_Calendar cccEffStartDate 
      Height          =   285
      Left            =   5115
      TabIndex        =   2
      Top             =   120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
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
      CSI_CurDayForeColor=   0
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   -1  'True
      CSI_DefaultDateType=   1
   End
   Begin VB.PictureBox pbcEffTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   7890
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   -15
      Width           =   60
   End
   Begin VB.CommandButton cmcChange 
      Caption         =   "&Change"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   6630
      Width           =   1335
   End
   Begin VB.CommandButton cmcNew 
      Caption         =   "&New from Scratch"
      Height          =   375
      Left            =   6090
      TabIndex        =   10
      Top             =   6630
      Width           =   2400
   End
   Begin VB.TextBox edcSearch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8340
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmcSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10035
      TabIndex        =   9
      Top             =   120
      Width           =   1140
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   11445
      Top             =   5760
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   -15
      ScaleHeight     =   90
      ScaleWidth      =   45
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   300
      Width           =   45
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   10710
      Top             =   6465
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
      Left            =   4350
      TabIndex        =   7
      Top             =   6630
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLib 
      Height          =   5880
      Left            =   150
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   10372
      _Version        =   393216
      Cols            =   11
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
      _Band(0).Cols   =   11
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lacEffEndDate 
      Caption         =   "End"
      Height          =   225
      Left            =   6315
      TabIndex        =   3
      Top             =   165
      Width           =   450
   End
   Begin VB.Label lacEffDate 
      Caption         =   "Effective Date- Start"
      Height          =   225
      Left            =   3435
      TabIndex        =   1
      Top             =   165
      Width           =   1575
   End
   Begin VB.Label lacScreen 
      Caption         =   "Libraries"
      Height          =   255
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   2625
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   9555
      Picture         =   "EngrLib.frx":030A
      Top             =   6540
      Width           =   480
   End
End
Attribute VB_Name = "EngrLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrLib - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imInChg As Integer
Private smNowDate As String
Private lmNowDate As Long
Private imRowSelectFlag As Integer '0=Zero rows selected; 1=One row selected; 2=2 or more rows selected
Private imSelectInFuture As Integer
Private smHours() As String
Private smBuses() As String
Private smReplaceValues() As String
Private smGridValues() As String
Private imOverlapCase As Integer    '1=Replace; 2=Terminate; 3=Change Start Date; 4=Split
Private smEffStartDate As String
Private lmEffStartDate As Long
Private smEffEndDate As String
Private lmEffEndDate As Long

Private tmCTE As CTE
Private tmCurrDBE() As DBE
Private smCurrDBE As String

Private smCurrEBEStamp As String
Private tmCurrEBE() As EBE
Private smEBEBuses() As String

Private smCurrDEEStamp As String
Private tmCurrDEE() As DEE
Private tmGridDEE As DEE

Private tmDHE As DHE

'Grid Controls
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private imShowGridBox As Integer
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on
Private imLastColSorted As Integer
Private imLastSort As Integer

Private lmLastClickedRow As Long

Const NAMEINDEX = 0
'Const SUBLIBNAMEINDEX = 1
Const DESCRIPTIONINDEX = 1
Const STARTDATEINDEX = 2
Const ENDDATEINDEX = 3
Const DAYSINDEX = 4
'Const STARTTIMEINDEX = 5
'Const LENGTHINDEX = 6
Const HOURSINDEX = 5    '7
Const BUSESINDEX = 6    '8
Const STATEINDEX = 7    '9
Const SELECTEDINDEX = 8
Const SORTINDEX = 9
Const CODEINDEX = 10 '10

Private Sub cccEffStartDate_Change()
    Dim slStr As String
    
    tmcClick.Enabled = False
    slStr = cccEffStartDate.text
    If slStr <> "" Then
        If gIsDate(slStr) Then
            tmcClick.Enabled = True
        End If
    Else
        tmcClick.Enabled = True
    End If
End Sub

Private Sub ckcDormant_Click()
    If UBound(tgCurrLibDHE) > LBound(tgCurrLibDHE) Then
        tmcClick_Timer
    End If
End Sub

Private Sub cmcChange_Click()
    Dim ilFound As Integer
    Dim llRow As Long
    
    If imRowSelectFlag >= 2 Then
        smEffStartDate = cccEffStartDate.text   'edcEffStartDate.Text
        If Not gIsDate(smEffStartDate) Then
            cccEffStartDate.SetFocus    'edcEffStartDate.SetFocus
            Exit Sub
        End If
        lmEffStartDate = gDateValue(smEffStartDate)
        smEffEndDate = cccEffEndDate.text
        If Trim$(smEffEndDate) <> "" Then
            If Not gIsDate(smEffEndDate) Then
                cccEffEndDate.SetFocus
                Exit Sub
            End If
        Else
            smEffEndDate = "12/31/2069"
        End If
        lmEffEndDate = gDateValue(smEffEndDate)
        ReDim tgLibReplaceValues(0 To 0) As LIBREPLACEVALUES
        mCreateUsedArrays
        mInitReplaceInfo
        igAnsReplace = 0
        igReplaceCallInfo = 3
        gSetMousePointer grdLib, grdLib, vbHourglass
        EngrReplaceLib.Show vbModal
        If igAnsReplace = CALLDONE Then 'Apply
            gSetMousePointer grdLib, grdLib, vbHourglass
'            grdLibEvents.Redraw = False
'            mReplaceValues
'            grdLibEvents.Redraw = True
            gSetMousePointer grdLib, grdLib, vbDefault
        End If
    Else
        lgLibCallCode = -1
        For llRow = grdLib.FixedRows To grdLib.Rows - 1 Step 1
            If grdLib.TextMatrix(llRow, NAMEINDEX) <> "" Then
                If grdLib.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                    lgLibCallCode = grdLib.TextMatrix(llRow, CODEINDEX)
                    Exit For
                End If
            End If
        Next llRow
        If lgLibCallCode > 0 Then
            If (Not imSelectInFuture) Or (cmcChange.Caption = "&View") Then
                If cmcChange.Caption = "&View" Then
                    igLibCallType = 3   'View
                Else
                    igLibCallType = 4   'Terminate
                End If
            Else
                igLibCallType = 1   'Change
            End If
            EngrLibDef.Show vbModeless
            Unload EngrLib
        End If
    End If
End Sub


Private Sub cmcNew_Click()
    Dim llRow As Long
    
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(LIBRARYJOB) = 2) Then
        If imRowSelectFlag = 1 Then
            lgLibCallCode = -1
            For llRow = grdLib.FixedRows To grdLib.Rows - 1 Step 1
                If grdLib.TextMatrix(llRow, NAMEINDEX) <> "" Then
                    If grdLib.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                        lgLibCallCode = grdLib.TextMatrix(llRow, CODEINDEX)
                        Exit For
                    End If
                End If
            Next llRow
            If lgLibCallCode <= 0 Then
                Exit Sub
            End If
            igLibCallType = 2   'Model
        Else
            lgLibCallCode = 0
            igLibCallType = 0   'New
        End If
        EngrLibDef.Show vbModeless
        Unload EngrLib
    End If
End Sub

Private Sub cmcSearch_Click()
    Dim llRow As Long
    Dim slStr As String
    For llRow = grdLib.FixedRows To grdLib.Rows - 1 Step 1
        If grdLib.TextMatrix(llRow, NAMEINDEX) <> "" Then
            If grdLib.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                grdLib.TextMatrix(llRow, SELECTEDINDEX) = "0"
                mPaintRowColor llRow
            End If
        End If
    Next llRow
    slStr = UCase(Trim$(edcSearch.text))
    For llRow = grdLib.FixedRows To grdLib.Rows - 1 Step 1
        If grdLib.TextMatrix(llRow, NAMEINDEX) <> "" Then
            If InStr(1, UCase(Trim$(grdLib.TextMatrix(llRow, NAMEINDEX))), slStr, vbBinaryCompare) > 0 Then
                grdLib.TextMatrix(llRow, SELECTEDINDEX) = "1"
                mPaintRowColor llRow
                grdLib.TopRow = llRow
                lmLastClickedRow = llRow
                mSetRowSelectedCount
                Exit For
            End If
        End If
    Next llRow
    mSetRowSelectedCount
    mSetCommands
End Sub




Private Sub mSortCol(ilCol As Integer)
    Dim slStr As String
    Dim llRow As Long
    
    
    For llRow = grdLib.FixedRows To grdLib.Rows - 1 Step 1
        slStr = Trim$(grdLib.TextMatrix(llRow, NAMEINDEX))
        If slStr <> "" Then
            If (ilCol = STARTDATEINDEX) Then
                slStr = grdLib.TextMatrix(llRow, STARTDATEINDEX)
                slStr = Trim$(Str$(gDateValue(slStr)))
                Do While Len(slStr) < 6
                    slStr = "0" & slStr
                Loop
            ElseIf (ilCol = ENDDATEINDEX) Then
                slStr = grdLib.TextMatrix(llRow, ENDDATEINDEX)
                slStr = Trim$(Str$(gDateValue(slStr)))
                Do While Len(slStr) < 6
                    slStr = "0" & slStr
                Loop
            Else
                slStr = grdLib.TextMatrix(llRow, ilCol)
            End If
            grdLib.TextMatrix(llRow, SORTINDEX) = slStr & grdLib.TextMatrix(llRow, SORTINDEX)
        End If
    Next llRow
    If imLastColSorted = ilCol Then
        gGrid_SortByCol grdLib, NAMEINDEX, SORTINDEX, SORTINDEX, imLastSort
    Else
        gGrid_SortByCol grdLib, NAMEINDEX, SORTINDEX, imLastColSorted, imLastSort
    End If
    imLastColSorted = ilCol

End Sub

Private Sub mSetCommands()
    Dim ilRet As Integer
    If UBound(tgCurrLibDHE) <= LBound(tgCurrLibDHE) Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(LIBRARYJOB) = 2) Then
            cmcChange.Enabled = False
            cmcChange.Caption = "&Change"
            cmcCancel.Enabled = True
            cmcNew.Enabled = True
            cmcNew.Caption = "&New from Scratch"
        Else
            cmcChange.Enabled = False
            cmcChange.Caption = "&View"
            cmcNew.Enabled = False
            cmcNew.Caption = "&New from Scratch"
            cmcCancel.Enabled = True
        End If
    Else
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(LIBRARYJOB) = 2) Then
            If imRowSelectFlag = 1 Then
                If imSelectInFuture Then
                    cmcChange.Caption = "&Change"
                Else
                    cmcChange.Caption = "&Terminate"
                End If
                cmcChange.Enabled = True
                cmcCancel.Enabled = True
                cmcNew.Enabled = True
                cmcNew.Caption = "&New with Modelling"
            ElseIf imRowSelectFlag >= 2 Then
                If imSelectInFuture Then
                    cmcChange.Caption = "&Replace"
                    cmcChange.Enabled = True
                Else
                    cmcChange.Caption = "&Change"
                    cmcChange.Enabled = False
                End If
                cmcCancel.Enabled = True
                cmcNew.Enabled = False
                cmcNew.Caption = "&New from Scratch"
            Else
                cmcChange.Enabled = False
                cmcChange.Caption = "&Change"
                cmcCancel.Enabled = True
                cmcNew.Enabled = True
                cmcNew.Caption = "&New from Scratch"
            End If
        Else
            If imRowSelectFlag = 1 Then
                cmcChange.Enabled = True
                cmcChange.Caption = "&View"
                cmcNew.Enabled = False
                cmcCancel.Enabled = True
            Else
                cmcChange.Enabled = False
                cmcChange.Caption = "View"
                cmcNew.Enabled = False
                cmcNew.Caption = "&New from Scratch"
                cmcCancel.Enabled = True
            End If
        End If
    End If
End Sub




Private Sub mGridColumns()
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    gGrid_AlignAllColsLeft grdLib
    mGridColumnWidth
    'Set Titles
    grdLib.TextMatrix(0, NAMEINDEX) = "Name/Subname"
    'grdLib.TextMatrix(SUBLIBNAMEINDEX + 1) = "Subname"
    grdLib.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdLib.TextMatrix(0, STARTDATEINDEX) = "Start Date"
    grdLib.TextMatrix(0, ENDDATEINDEX) = "End Date"
    grdLib.TextMatrix(0, DAYSINDEX) = "Days"
    'grdLib.TextMatrix(STARTTIMEINDEX + 1) = "Start Hour"
    'grdLib.TextMatrix(LENGTHINDEX + 1) = "Length"
    grdLib.TextMatrix(0, HOURSINDEX) = "Hours"
    grdLib.TextMatrix(0, BUSESINDEX) = "Bus"
    grdLib.TextMatrix(0, STATEINDEX) = "State"
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    Dim llWidth As Single
    
    grdLib.Width = EngrLib.Width - 2 * grdLib.Left - 120
    grdLib.ColWidth(SELECTEDINDEX) = 0
    grdLib.ColWidth(SORTINDEX) = 0
    grdLib.ColWidth(CODEINDEX) = 0
    grdLib.ColWidth(NAMEINDEX) = grdLib.Width / 5
    'grdLib.ColWidth.Item(SUBLIBNAMEINDEX + 1).Width = grdLib.Width / 9
    grdLib.ColWidth(STARTDATEINDEX) = grdLib.Width / 13
    grdLib.ColWidth(ENDDATEINDEX) = grdLib.Width / 13
    grdLib.ColWidth(DAYSINDEX) = grdLib.Width / 18
    'grdLib.ColWidth.Item(STARTTIMEINDEX + 1).Width = grdLib.Width / 18
    'grdLib.ColWidth.Item(LENGTHINDEX + 1).Width = grdLib.Width / 18
    grdLib.ColWidth(HOURSINDEX) = grdLib.Width / 8
    grdLib.ColWidth(BUSESINDEX) = grdLib.Width / 7
    grdLib.ColWidth(STATEINDEX) = grdLib.Width / 20
    llWidth = 0
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            llWidth = llWidth + grdLib.ColWidth(ilCol)
        End If
    Next ilCol
    grdLib.ColWidth(DESCRIPTIONINDEX) = grdLib.Width - llWidth - GRIDSCROLLWIDTH '- 10 * 240 '- 30

End Sub


Private Sub mClearControls()
    
End Sub


Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilDNE As Integer
    Dim ilDSE As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilDBE As Integer
    Dim ilBGE As Integer
    Dim ilBDE As Integer
    Dim slBus As String
    Dim ilHour As Integer
    Dim slHour As String
    Dim ilSIndex As Integer
    Dim ilEIndex As Integer
    Dim slSDate As String
    Dim llSdate As Long
    Dim slEDate As String
    Dim llEDate As Long
    Dim ilShow As Integer
    Dim llCol As Long
    
'    Set mItem = grdLib.ListItems.Add()
'    llRow = mItem.Height
    grdLib.Redraw = False
    grdLib.Rows = grdLib.FixedRows + 1
    
    For llCol = NAMEINDEX To STATEINDEX Step 1
        grdLib.Row = grdLib.FixedRows
        grdLib.Col = llCol
        grdLib.CellBackColor = vbWhite
    Next llCol
    lmLastClickedRow = -1
    slSDate = cccEffStartDate.text  'edcEffStartDate.Text
    If slSDate <> "" Then
        If Not gIsDate(slSDate) Then
            Beep
            cccEffStartDate.SetFocus    'edcEffStartDate.SetFocus
            Exit Sub
        End If
        llSdate = gDateValue(slSDate)
    Else
        llSdate = 0
    End If
    slEDate = cccEffEndDate.text
    If slEDate <> "" Then
        If Not gIsDate(slEDate) Then
            Beep
            cccEffEndDate.SetFocus
            Exit Sub
        End If
        llEDate = gDateValue(slEDate)
    Else
        llEDate = 2000000000
    End If
    
    llRow = grdLib.FixedRows
    For ilLoop = 0 To UBound(tgCurrLibDHE) - 1 Step 1
        If gDateValue(tgCurrLibDHE(ilLoop).sEndDate) < (llSdate) Then
            ilShow = False
        Else
            If gDateValue(tgCurrLibDHE(ilLoop).sStartDate) > (llEDate) Then
                ilShow = False
            Else
                ilShow = True
            End If
        End If
        If ckcDormant.Value = vbUnchecked Then
            If tgCurrLibDHE(ilLoop).sState = "D" Then
                ilShow = False
            End If
        End If
        If ilShow Then
            If llRow >= grdLib.Rows Then
                grdLib.AddItem ""
            End If
            
            slStr = ""
            For ilDNE = 0 To UBound(tgCurrLibDNE) - 1 Step 1
                If tgCurrLibDHE(ilLoop).lDneCode = tgCurrLibDNE(ilDNE).lCode Then
                    slStr = Trim$(tgCurrLibDNE(ilDNE).sName)
                    Exit For
                End If
            Next ilDNE
            For ilDSE = 0 To UBound(tgCurrDSE) - 1 Step 1
                If tgCurrLibDHE(ilLoop).lDseCode = tgCurrDSE(ilDSE).lCode Then
                    slStr = slStr & "/" & Trim$(tgCurrDSE(ilDSE).sName)
                    Exit For
                End If
            Next ilDSE
            grdLib.TextMatrix(llRow, NAMEINDEX) = slStr
            ilRet = gGetRec_CTE_CommtsTitle(tgCurrLibDHE(ilLoop).lCteCode, "EngrLib- mMoveRecToCtrl for CTE", tmCTE)
            grdLib.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tmCTE.sComment)
            grdLib.TextMatrix(llRow, STARTDATEINDEX) = Trim$(tgCurrLibDHE(ilLoop).sStartDate)
            If gDateValue(Trim$(tgCurrLibDHE(ilLoop).sEndDate)) <> gDateValue("12/31/2069") Then
                grdLib.TextMatrix(llRow, ENDDATEINDEX) = Trim$(tgCurrLibDHE(ilLoop).sEndDate)
            Else
                grdLib.TextMatrix(llRow, ENDDATEINDEX) = ""
            End If
            slStr = gDayMap(tgCurrLibDHE(ilLoop).sDays)
            grdLib.TextMatrix(llRow, DAYSINDEX) = slStr
            'mItem.SubItems(STARTTIMEINDEX) = Trim$(tgCurrLibDHE(ilLoop).sStartTime)
            'mItem.SubItems(LENGTHINDEX) = gLongToLength(tgCurrLibDHE(ilLoop).lLength)
            slStr = gHourMap(tgCurrLibDHE(ilLoop).sHours)
            grdLib.TextMatrix(llRow, HOURSINDEX) = Trim$(slStr)
            '6/27/11: Replace with string in header
            'ilRet = gGetRecs_DBE_DayBusSel(smCurrDBE, tgCurrLibDHE(ilLoop).lCode, "EngrLib- mMoveRecToCtrl for DBE", tmCurrDBE())
            'slStr = ""
            '6/27/11: Replace getting bus name with string in header
            'For ilDBE = 0 To UBound(tmCurrDBE) - 1 Step 1
            '    If tmCurrDBE(ilDBE).sType = "B" Then
            '        'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
            '        '    If tmCurrDBE(ilDBE).iBdeCode = tgCurrBDE(ilBDE).iCode Then
            '            ilBDE = gBinarySearchBDE(tmCurrDBE(ilDBE).iBdeCode, tgCurrBDE())
            '            If ilBDE <> -1 Then
            '                If slStr = "" Then
            '                    slStr = Trim$(tgCurrBDE(ilBDE).sName)
            '                Else
            '                    slStr = slStr & ", " & Trim$(tgCurrBDE(ilBDE).sName)
            '                End If
            '        '        Exit For
            '            End If
            '        'Next ilBDE
            '    End If
            'Next ilDBE
            '
            'mItem.SubItems(BUSESINDEX) = slStr
            grdLib.TextMatrix(llRow, BUSESINDEX) = Trim$(tgCurrLibDHE(ilLoop).sBusNames)
            If tgCurrLibDHE(ilLoop).sState = "A" Then
                grdLib.TextMatrix(llRow, STATEINDEX) = "Active"
            ElseIf tgCurrLibDHE(ilLoop).sState = "L" Then
                grdLib.TextMatrix(llRow, STATEINDEX) = "Limbo"
            Else
                grdLib.TextMatrix(llRow, STATEINDEX) = "Dormant"
            End If
            grdLib.TextMatrix(llRow, CODEINDEX) = tgCurrLibDHE(ilLoop).lCode
            grdLib.TextMatrix(llRow, SELECTEDINDEX) = "0"
            llRow = llRow + 1
        End If
    Next ilLoop
    gGrid_IntegralHeight grdLib
    gGrid_FillWithRows grdLib
    grdLib.Height = grdLib.Height
    grdLib.Visible = True
    grdLib.Redraw = True
    If lgLibTopRow <> -1 Then
        grdLib.TopRow = lgLibTopRow
    End If
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    ilRet = gGetTypeOfRecs_DHE_DayHeaderInfo("C", "L", sgCurrLibDHEStamp, "EngrLib-mPopulate", tgCurrLibDHE())
    ilRet = gGetTypeOfRecs_DNE_DayName("C", "L", sgCurrLibDNEStamp, "EngrLib-mPopulate Library Names", tgCurrLibDNE())
    ilRet = gGetTypeOfRecs_DSE_DaySubName("C", sgCurrDSEStamp, "EngrLib-mPopulate SubNames", tgCurrDSE())
    ilRet = gGetTypeOfRecs_BGE_BusGroup("C", sgCurrBGEStamp, "EngrLib-mPopulate Bus Groups", tgCurrBGE())
    ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrLib-Bus Definitions", tgCurrBDE())
    
    
End Sub
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    igJobShowing(LIBRARYJOB) = 0
    Unload EngrLib
End Sub

Private Sub cccEffEndDate_Change()
    Dim slStr As String
    
    tmcClick.Enabled = False
    slStr = cccEffEndDate.text
    If slStr <> "" Then
        If gIsDate(slStr) Then
            tmcClick.Enabled = True
        End If
    Else
        tmcClick.Enabled = True
    End If
End Sub

Private Sub edcSearch_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Click()
    cmcCancel.SetFocus
End Sub

Private Sub Form_Initialize()
    'Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    'Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    Me.Move Me.Left, Me.Top, 0.97 * Screen.Width, 0.82 * Screen.Height
    Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrLib
    'gCenterFormModal EngrLib
    gCenterForm EngrLib
End Sub

Private Sub Form_Load()
    mGridColumns
    mInit
    igJobShowing(LIBRARYJOB) = 1
End Sub

Private Sub Form_Resize()
    'These call are here and in form_Active (call to mGridColumns)
    'They are in mGridColumn in case the For_Initialize size chage does not cause a resize event
    grdLib.Visible = False
    mGridColumnWidth
    grdLib.Height = cmcCancel.Top - grdLib.Top - 240    '8 * grdLib.RowHeight(0) + 30
    gGrid_IntegralHeight grdLib
    gGrid_FillWithRows grdLib
    grdLib.Height = grdLib.Height
    grdLib.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lgLibTopRow = grdLib.TopRow
    Erase smHours
    Erase smBuses
    Erase smReplaceValues
    Erase smGridValues
    Erase tmCurrDBE
    Erase tmCurrEBE
    Erase smEBEBuses
    Erase tmCurrDEE
    Set EngrLib = Nothing
End Sub





Private Sub mInit()
    Dim llRet As Long
    On Error GoTo ErrHand
    
    gSetMousePointer grdLib, grdLib, vbHourglass
    imcPrint.Picture = EngrMain!imcPrinter.Picture
    cmcSearch.Top = 30
    edcSearch.Top = cmcSearch.Top
    smNowDate = Format$(gNow(), "ddddd")
    lmNowDate = gDateValue(smNowDate)
    'edcEffStartDate.Text = smNowDate
    cccEffStartDate.text = smNowDate
    tmcClick.Enabled = False
    imLastColSorted = -1
    imLastSort = -1
    lmEnableRow = -1
    imRowSelectFlag = -1
    imSelectInFuture = False
    imInChg = True
    mPopulate
    mMoveRecToCtrls
    mSortCol NAMEINDEX
    imInChg = False
    mSetCommands
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igJobStatus(LIBRARYJOB) = 2) Then
    Else
        cmcChange.Caption = "&View"
        cmcNew.Enabled = False
    End If
    gSetMousePointer grdLib, grdLib, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdLib, grdLib, vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetCoDHEctAttr vs. SQLSetOpenCoDHEction
            gMsg = "A SQL error has occured in Relay Definition-Form Load: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Relay Definition-Form Load: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub





















Private Sub mFindMatch(ilCreateNew As Integer)
    Dim llRow As Long
    Dim slStr As String
    
    If igInitCallInfo = 0 Then
        If UBound(tgCurrLibDHE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    
End Sub


Private Sub grdLib_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer

    If y < grdLib.RowHeight(0) Then
        grdLib.Col = grdLib.MouseCol
        mSortCol grdLib.Col
        grdLib.Row = 0
        grdLib.Col = CODEINDEX
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdLib, x, y, llCurrentRow, llCol)
    If llCurrentRow < grdLib.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdLib.FixedRows Then
        imRowSelectFlag = 0
        If grdLib.TextMatrix(llCurrentRow, NAMEINDEX) <> "" Then
            'grdLib.TopRow = lmScrollTop
            llTopRow = grdLib.TopRow
            If (Shift And CTRLMASK) > 0 Then
                If grdLib.TextMatrix(llCurrentRow, SELECTEDINDEX) <> 1 Then
                    grdLib.TextMatrix(llCurrentRow, SELECTEDINDEX) = 1
                    '7/10/11: Added
                    lmLastClickedRow = llCurrentRow
                Else
                    grdLib.TextMatrix(llCurrentRow, SELECTEDINDEX) = 0
                    '7/10/11: Added
                    lmLastClickedRow = -1
                End If
                mPaintRowColor llCurrentRow
                mSetRowSelectedCount
            Else
                '7/10/11: Disallow multi-row selection.  i.e. removing Replace.  Support routines defined in libDef
                'For llRow = grdLib.FixedRows To grdLib.Rows - 1 Step 1
                '    If grdLib.TextMatrix(llRow, NAMEINDEX) <> "" Then
                '        grdLib.TextMatrix(llRow, SELECTEDINDEX) = "0"
                '        If (lmLastClickedRow = -1) Or ((Shift And SHIFTMASK) <= 0) Then
                '            If llRow = llCurrentRow Then
                '                grdLib.TextMatrix(llRow, SELECTEDINDEX) = "1"
                '            Else
                '                grdLib.TextMatrix(llRow, SELECTEDINDEX) = "0"
                '            End If
                '        ElseIf lmLastClickedRow < llCurrentRow Then
                '            If (llRow >= lmLastClickedRow) And (llRow <= llCurrentRow) Then
                '                grdLib.TextMatrix(llRow, SELECTEDINDEX) = "1"
                '            End If
                '        Else
                '            If (llRow >= llCurrentRow) And (llRow <= lmLastClickedRow) Then
                '                grdLib.TextMatrix(llRow, SELECTEDINDEX) = "1"
                '            End If
                '        End If
                '        mPaintRowColor llRow
                '    End If
                'Next llRow
                '7/10/11: Added
                If lmLastClickedRow >= grdLib.FixedRows Then
                    grdLib.TextMatrix(lmLastClickedRow, SELECTEDINDEX) = "0"
                    mPaintRowColor lmLastClickedRow
                End If
                grdLib.TextMatrix(llCurrentRow, SELECTEDINDEX) = "1"
                mPaintRowColor llCurrentRow
                lmLastClickedRow = llCurrentRow
                '7/10/11: End of add
                mSetRowSelectedCount
                grdLib.TopRow = llTopRow
                grdLib.Row = llCurrentRow
            End If
            '7/10/11: Removed
            'lmLastClickedRow = llCurrentRow
        End If
        mSetCommands
    End If

End Sub

Private Sub imcPrint_Click()
    igRptIndex = LIBRARY_RPT
    igRptSource = vbModal
    EngrLibRpt.Show vbModal
End Sub


Private Sub pbcEffTab_GotFocus()
    If tmcClick.Enabled Then
        tmcClick.Enabled = False
        gSetMousePointer grdLib, grdLib, vbHourglass
        mMoveRecToCtrls
        mSortCol NAMEINDEX
        mSetCommands
        gSetMousePointer grdLib, grdLib, vbDefault
    End If
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    gSetMousePointer grdLib, grdLib, vbHourglass
    mMoveRecToCtrls
    mSortCol NAMEINDEX
    mSetCommands
    gSetMousePointer grdLib, grdLib, vbDefault
End Sub

Private Sub mInitReplaceInfo()
    Dim ilUpper As Integer
    ReDim tgReplaceFields(0 To 0) As FIELDSELECTION
    
    ilUpper = 0
    If ((tgUsedSumEPE.sAudioName <> "N") Or (tgUsedSumEPE.sProtAudioName <> "N") Or (tgUsedSumEPE.sBkupAudioName <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Audio Name"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("AudioName", 6)
        tgReplaceFields(ilUpper).sListFile = "ANE"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sAudioName
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgUsedSumEPE.sBus <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Bus"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("BusName", 6)
        tgReplaceFields(ilUpper).sListFile = "BDE"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sBus
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgUsedSumEPE.sFollow <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Follow"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Follow", 6)
        tgReplaceFields(ilUpper).sListFile = "FNE"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sFollow
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgUsedSumEPE.sMaterialType <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Material"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Material", 6)
        tgReplaceFields(ilUpper).sListFile = "MTE"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sMaterialType
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgUsedSumEPE.sStartNetcue <> "N") Or (tgUsedSumEPE.sStopNetcue <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Netcue"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Netcue1", 6)
        tgReplaceFields(ilUpper).sListFile = "NNE"
        If (tgManSumEPE.sStartNetcue = "Y") Or (tgManSumEPE.sStopNetcue = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgUsedSumEPE.sRelay1 <> "N") Or (tgUsedSumEPE.sRelay2 <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Relay"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Relay1", 6)
        tgReplaceFields(ilUpper).sListFile = "RNE"
        If (tgManSumEPE.sRelay1 = "Y") Or (tgManSumEPE.sRelay2 = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgUsedSumEPE.sStartType <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Start Type"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("StartType", 6)
        tgReplaceFields(ilUpper).sListFile = "TTES"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sStartType
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgUsedSumEPE.sEndType <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "End Type"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("EndType", 6)
        tgReplaceFields(ilUpper).sListFile = "TTEE"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sEndType
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgUsedSumEPE.sAudioControl <> "N") Or (tgUsedSumEPE.sProtAudioControl <> "N") Or (tgUsedSumEPE.sBkupAudioControl <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Audio Control"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("AudioCtrl", 6)
        tgReplaceFields(ilUpper).sListFile = "CCEA"
        If (tgManSumEPE.sAudioControl = "Y") Or (tgManSumEPE.sProtAudioControl = "Y") Or (tgManSumEPE.sBkupAudioControl = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgUsedSumEPE.sBusControl <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Bus Control"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("BusCtrl", 6)
        tgReplaceFields(ilUpper).sListFile = "CCEB"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sBusControl
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    '7/8/11: Make T2 work like T1
    'If (tgUsedSumEPE.sTitle2 <> "N") Then
    '    tgReplaceFields(ilUpper).sFieldName = "Title 2"
    '    'tgReplaceFields(ilUpper).iFieldType = 5
    '    tgReplaceFields(ilUpper).iFieldType = 9
    '    tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Title2", 6)
    '    tgReplaceFields(ilUpper).sListFile = "CTE2"
    '    tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sTitle2
    '    ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
    '    ilUpper = ilUpper + 1
    'End If
    If ((tgUsedSumEPE.sSilence1 <> "N") Or (tgUsedSumEPE.sSilence2 <> "N") Or (tgUsedSumEPE.sSilence3 <> "N") Or (tgUsedSumEPE.sSilence4 <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Silence Control"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Silence1", 6)
        tgReplaceFields(ilUpper).sListFile = "SCE"
        If (tgManSumEPE.sSilence1 = "Y") Or (tgManSumEPE.sSilence2 = "Y") Or (tgManSumEPE.sSilence3 = "Y") Or (tgManSumEPE.sSilence4 = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If tgUsedSumEPE.sFixedTime <> "N" Then
        tgReplaceFields(ilUpper).sFieldName = "Fixed Time"
        tgReplaceFields(ilUpper).iFieldType = 9
        tgReplaceFields(ilUpper).iMaxNoChar = 1
        tgReplaceFields(ilUpper).sListFile = "FTYN"
        tgReplaceFields(ilUpper).sMandatory = tgManSumEPE.sFixedTime
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgUsedSumEPE.sAudioItemID <> "N") Or (tgUsedSumEPE.sProtAudioItemID <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Item ID"
        tgReplaceFields(ilUpper).iFieldType = 2
        tgReplaceFields(ilUpper).iMaxNoChar = 0
        tgReplaceFields(ilUpper).sListFile = ""
        If (tgManSumEPE.sAudioItemID = "Y") Or (tgManSumEPE.sProtAudioItemID = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgUsedSumEPE.sAudioISCI <> "N") Or (tgUsedSumEPE.sProtAudioISCI <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "ISCI"
        tgReplaceFields(ilUpper).iFieldType = 2
        tgReplaceFields(ilUpper).iMaxNoChar = 0
        tgReplaceFields(ilUpper).sListFile = ""
        If (tgManSumEPE.sAudioISCI = "Y") Or (tgManSumEPE.sProtAudioISCI = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    
End Sub

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
    Dim slStr As String
    Dim ilBus As Integer
    Dim slBuses As String
    Dim slHours As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    
    ilRet = gGetTypeOfRecs_ASE_AudioSource("C", sgCurrASEStamp, "EngrLibDef-mPopASE Audio Source", tgCurrASE())
    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrLibDef-mPopASE Audio Audio Names", tgCurrANE())
    ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrLibDef-mPopBDE Bus Definition", tgCurrBDE())
    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "A", sgCurrAudioCCEStamp, "EngrLibDef-mPopCCE_Audio Control Character", tgCurrAudioCCE())
    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "B", sgCurrBusCCEStamp, "EngrLibDef-mPopCCE_Bus Control Character", tgCurrBusCCE())
    '7/8/11: Make T@ work like T1
    'ilRet = gGetTypeOfRecs_CTE_CommtsTitle("C", "T2", sgCurrCTEStamp, "EngrLibDef-mPopCTE Title 2", tgCurrCTE())
    ilRet = gGetTypeOfRecs_ETE_EventType("C", sgCurrETEStamp, "EngrLibETE-mPopETE Event Types", tgCurrETE())
    ilRet = gGetTypeOfRecs_EPE_EventProperties("C", sgCurrEPEStamp, "EngrLibDef-mPopETE Event Properties", tgCurrEPE())
    ilRet = gGetTypeOfRecs_FNE_FollowName("C", sgCurrFNEStamp, "EngrLibDef-mPopFNE Follow", tgCurrFNE())
    ilRet = gGetTypeOfRecs_MTE_MaterialType("C", sgCurrMTEStamp, "EngrLibDef-mPopMTE Material Type", tgCurrMTE())
    ilRet = gGetTypeOfRecs_NNE_NetcueName("C", sgCurrNNEStamp, "EngrLibDef-mPopNNE Netcue", tgCurrNNE())
    ilRet = gGetTypeOfRecs_RNE_RelayName("C", sgCurrRNEStamp, "EngrLibDef-mPopRNE Relay", tgCurrRNE())
    ilRet = gGetTypeOfRecs_SCE_SilenceChar("C", sgCurrSCEStamp, "EngrLibDef-mPopSCE Silence Character", tgCurrSCE())
    ilRet = gGetTypeOfRecs_TTE_TimeType("C", "E", sgCurrEndTTEStamp, "EngrLibDef-mPopTTE_EndType End Type", tgCurrEndTTE())
    ilRet = gGetTypeOfRecs_TTE_TimeType("C", "S", sgCurrStartTTEStamp, "EngrLibDef-mPopTTE_StartType Start Type", tgCurrStartTTE())
    ReDim tgYNMatchList(0 To 2) As MATCHLIST
    tgYNMatchList(0).sValue = "Y"
    tgYNMatchList(0).lValue = 0
    tgYNMatchList(1).sValue = "N"
    tgYNMatchList(1).lValue = 1
    sgReplaceDefaultHours = ""
    slHours = String(24, "N")
    ReDim tgUsedBDE(0 To 0) As BDE
    For ilLoop = grdLib.FixedRows To grdLib.Rows - 1 Step 1
        If grdLib.TextMatrix(ilLoop, NAMEINDEX) <> "" Then
            'Fix
            If grdLib.TextMatrix(ilLoop, SELECTEDINDEX) = "1" Then
                slBuses = grdLib.TextMatrix(ilLoop, BUSESINDEX)
                gParseCDFields slBuses, False, smBuses()
                For ilBus = LBound(smBuses) To UBound(smBuses) Step 1
                    slStr = Trim$(smBuses(ilBus))
                    For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
                        If StrComp(slStr, Trim$(tgCurrBDE(ilBDE).sName), vbTextCompare) = 0 Then
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
                            Exit For
                        End If
                    Next ilBDE
                Next ilBus
                slStr = grdLib.TextMatrix(ilLoop, HOURSINDEX)
                'mCreateHourStr slStr, slHours
                slHours = gCreateHourStr(slStr)
            End If
        End If
    Next ilLoop
    sgReplaceDefaultHours = gHourMap(slHours)

    ReDim tgUsedANE(0 To UBound(tgCurrANE)) As ANE
    For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        LSet tgUsedANE(ilANE) = tgCurrANE(ilANE)
    Next ilANE
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

End Sub

Private Sub mCreateHourStr(slHourName As String, slHourMap As String)
    Dim slStr As String
    Dim ilHours As Integer
    Dim ilPos As Integer
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim ilSet As Integer
    
    slStr = Trim$(slHourName)
    gParseCDFields slStr, False, smHours()
    For ilHours = LBound(smHours) To UBound(smHours) Step 1
        ilPos = InStr(1, smHours(ilHours), "-", vbTextCompare)
        If ilPos <= 0 Then
            ilStart = Val(smHours(ilHours))
            ilEnd = ilStart
        Else
            ilStart = Val(Left$(smHours(ilHours), ilPos - 1))
            ilEnd = Val(Mid$(smHours(ilHours), ilPos + 1))
        End If
        If (ilStart < 0) Or (ilStart > 23) Or (ilEnd < 0) Or (ilEnd > 23) Or (ilEnd < ilStart) Then
            Exit For
        Else
            For ilSet = ilStart To ilEnd Step 1
                Mid$(slHourMap, ilSet + 1, 1) = "Y"
            Next ilSet
        End If
    Next ilHours
End Sub


Private Sub mReplaceValues()
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim ilReplace As Integer
    Dim ilField As Integer
    Dim ilFieldType As Integer
    Dim slGridBuses As String
    Dim slGridHours As String
    Dim ilGLoop As Integer
    Dim ilRLoop As Integer
    Dim ilBusMatch As Integer
    Dim ilHourMatch As Integer
    Dim slReplaceBuses As String
    Dim slReplaceHours As String
    Dim slFileName As String
    Dim ilColumn As Integer
    Dim ilSet As Integer
    Dim slNewValue As String
    Dim slOldValue As String
    Dim ilAllBusesMatch As Integer
    Dim ilAllHoursMatch As Integer
    Dim slFromHours As String
    Dim slToHours As String
    Dim llFromRow As Long
    Dim llToRow As Long
    Dim ilFieldChanged As Integer
    Dim ilLib As Integer
    Dim ilPass As Integer
    Dim ilASE As Integer
    Dim ilEBE As Integer
    Dim ilBDE As Integer
    Dim llLibCode As Long
    Dim ilRet As Integer
    Dim ilSplit As Integer
    Dim ilSkipPass1 As Integer
    
    For ilLoop = grdLib.FixedRows To grdLib.Rows - 1 Step 1
        If grdLib.TextMatrix(ilLoop, NAMEINDEX) <> "" Then
            'Fix
            If grdLib.TextMatrix(ilLoop, SELECTEDINDEX) = "1" Then
                llLibCode = grdLib.TextMatrix(ilLoop, CODEINDEX)
                ilRet = gGetRec_DHE_DayHeaderInfo(llLibCode, "EngrLibDef-mPopulation", tmDHE)
                ilRet = gGetRecs_DEE_DayEvent(smCurrDEEStamp, llLibCode, "EngrLibDef-mPopulate", tmCurrDEE())
                ReDim smEBEBuses(0 To UBound(tmCurrDEE)) As String
                ilSkipPass1 = True
                For ilPass = 0 To 1 Step 1
                    'Pass = 0 => Check if libary changed
                    'Pass = 1 => Change library
                    If ilPass = 1 Then
                        'See OverlapCase in EngrLibDef
                        'Update Header- Determine how header should be handled
                        'If Effective date overlap library dates, then replace
                        'If Effective dates slpit within library dates, then terminate and duplicate after effective end date
                        'If Effective date partially overlap, then either adjust start date or end date of library
                        mSetOverlapFlag 'Note:  if multi-libraries overlapped, then each one will be adjusted separately
                        'Library defined 3/1-3/28 and another 3/29-TFN.
                        'Effective dates 3/15-4/4, then terminate first on 3/14 and have the new one only run thru 3/31.
                        'Change the start date of the other library to be 4/5 and the effective start date chaged to 3/29
                        'Don't add dates to libraries.  only use the current dates of the libraries
                        'Library started on 8/2 thru TFN and Effective dates are 7/19 to TFN, then
                        'the effective start date would be changed to 8/2
                        
                        'Note: Determine latest version number (gGetLatestVersion_DHE)
                        '      Retain end date of library
                    End If
                    For llRow = LBound(tmCurrDEE) To UBound(tmCurrDEE) - 1 Step 1
                        'Check if Bus and Hour filter matched
                        If ilPass = 0 Then
                            Erase tmCurrEBE
                            ilRet = gGetRecs_EBE_EventBusSel(smCurrEBEStamp, tmCurrDEE(llRow).lCode, "Bus Definition-mDEEMoveRecToCtrls", tmCurrEBE())
                            For ilEBE = 0 To UBound(tmCurrEBE) - 1 Step 1
                                'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
                                '    If tmCurrEBE(ilEBE).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                                    ilBDE = gBinarySearchBDE(tmCurrEBE(ilEBE).iBdeCode, tgCurrBDE())
                                    If ilBDE <> -1 Then
                                        slStr = slStr & Trim$(tgCurrBDE(ilBDE).sName) & ","
                                '        Exit For
                                    End If
                                'Next ilBDE
                            Next ilEBE
                            If slStr <> "" Then
                                slStr = Left$(slStr, Len(slStr) - 1)
                            End If
                            slGridBuses = slStr
                            smEBEBuses(llRow) = slStr
                        Else
                            slGridBuses = smEBEBuses(llRow)
                        End If
                        slGridHours = Trim$(tmCurrDEE(llRow).sHours)
                        LSet tmGridDEE = tmCurrDEE(llRow)
                        ilSplit = False
                        For ilReplace = LBound(tgLibReplaceValues) To UBound(tgLibReplaceValues) - 1 Step 1
                            For ilField = LBound(tgReplaceFields) To UBound(tgReplaceFields) - 1 Step 1
                                If tgReplaceFields(ilField).sFieldName = tgLibReplaceValues(ilReplace).sFieldName Then
                                    ilFieldType = tgReplaceFields(ilField).iFieldType
                                    slFileName = tgReplaceFields(ilField).sListFile
                                    slReplaceBuses = tgLibReplaceValues(ilReplace).sBuses
                                    
                                    ilBusMatch = 0
                                    gParseCDFields slGridBuses, False, smGridValues()
                                    gParseCDFields slReplaceBuses, False, smReplaceValues()
                                    For ilGLoop = LBound(smGridValues) To UBound(smGridValues) Step 1
                                        For ilRLoop = LBound(smReplaceValues) To UBound(smReplaceValues) Step 1
                                            If StrComp(Trim$(smGridValues(ilGLoop)), Trim$(smReplaceValues(ilRLoop)), vbTextCompare) = 0 Then
                                                ilBusMatch = ilBusMatch + 1
                                                Exit For
                                            End If
                                        Next ilRLoop
                                    Next ilGLoop
                                    If ilBusMatch = (UBound(smGridValues) - LBound(smGridValues) + 1) Then
                                        ilAllBusesMatch = True
                                    Else
                                        ilAllBusesMatch = False
                                    End If
                                    
                                    ilHourMatch = False
                                    slReplaceHours = Trim$(tgLibReplaceValues(ilReplace).sHours)
                                    If StrComp(slGridHours, slReplaceHours, vbTextCompare) = 0 Then
                                        ilHourMatch = True
                                        ilAllHoursMatch = True
                                    Else
                                        ilAllHoursMatch = True
                                        For ilGLoop = 1 To 24 Step 1
                                            If (Mid$(slGridHours, ilGLoop, 1) = "Y") And (Mid$(slReplaceHours, ilGLoop, 1) = "Y") Then
                                                ilHourMatch = True
                                            End If
                                            If (Mid$(slGridHours, ilGLoop, 1) = "Y") And (Mid$(slReplaceHours, ilGLoop, 1) = "N") Then
                                                ilAllHoursMatch = False
                                            End If
                                        Next ilGLoop
                                    End If
                                    ilFieldChanged = False
                                    If (ilBusMatch <> 0) And ilHourMatch Then
                                        If ilFieldType = 5 Then
                                            Select Case UCase$(Trim$(slFileName))
                                                Case "ANE"
                                                    'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                                                    '    If tmGridDEE.iAudioAseCode = tgCurrASE(ilASE).iCode Then
                                                        ilASE = gBinarySearchASE(tmGridDEE.iAudioAseCode, tgCurrASE())
                                                        If ilASE <> -1 Then
                                                            If tgCurrASE(ilASE).iPriAneCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                                ilFieldChanged = True
                                                                If ilPass = 1 Then
                                                                    tmCurrDEE(llRow).iProtAneCode = tgLibReplaceValues(ilReplace).lNewCode
                                                                End If
                                                            End If
                                                    '        Exit For
                                                        End If
                                                    'Next ilASE
                                                    If tmGridDEE.iProtAneCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).iProtAneCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                    If tmGridDEE.iBkupAneCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).iBkupAneCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                Case "BDE"
                                                    gParseCDFields slGridBuses, False, smGridValues()
                                                    For ilEBE = 0 To UBound(tmCurrEBE) - 1 Step 1
                                                        If tmCurrEBE(ilEBE).iBdeCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                            ilFieldChanged = True
                                                            If ilPass = 1 Then
                                                                tmCurrEBE(ilEBE).iBdeCode = tgLibReplaceValues(ilReplace).lNewCode
                                                            End If
                                                        End If
                                                    Next ilEBE
                                                Case "FNE"
                                                    If tmGridDEE.iFneCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).iFneCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                Case "MTE"
                                                    If tmGridDEE.iMteCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).iMteCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                Case "NNE"
                                                    If tmGridDEE.iStartNneCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).iStartNneCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                    If tmGridDEE.iEndNneCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).iEndNneCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                Case "RNE"
                                                    If tmGridDEE.i1RneCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).i1RneCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                    If tmGridDEE.i2RneCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).i2RneCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                Case "TTES"
                                                    If tmGridDEE.iStartTteCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).iStartTteCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                Case "TTEE"
                                                    If tmGridDEE.iEndTteCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).iEndTteCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                Case "CCEA"
                                                    If tmGridDEE.iAudioCceCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).iAudioCceCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                    If tmGridDEE.iProtCceCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).iProtCceCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                    If tmGridDEE.iBkupCceCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).iBkupCceCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                Case "CCEB"
                                                    If tmGridDEE.iCceCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).iCceCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                Case "SCE"
                                                    If tmGridDEE.i1SceCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).i1SceCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                    If tmGridDEE.i2SceCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).i2SceCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                    If tmGridDEE.i3SceCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).i3SceCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                    If tmGridDEE.i4SceCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                        ilFieldChanged = True
                                                        If ilPass = 1 Then
                                                            tmCurrDEE(llRow).i4SceCode = tgLibReplaceValues(ilReplace).lNewCode
                                                        End If
                                                    End If
                                                '7/8/11: Make T2 work like T1
                                                'Case "CTE2"
                                                '    If tmGridDEE.l2CteCode = tgLibReplaceValues(ilReplace).lOldCode Then
                                                '        ilFieldChanged = True
                                                '        If ilPass = 1 Then
                                                '            tmCurrDEE(llRow).l2CteCode = tgLibReplaceValues(ilReplace).lNewCode
                                                '        End If
                                                '    End If
                                            End Select
                                        ElseIf ilFieldType = 9 Then
                                            If Trim$(tgReplaceFields(ilField).sFieldName) = "Fixed Time" Then
                                                If StrComp(Trim$(tmGridDEE.sFixedTime), Trim$(tgLibReplaceValues(ilReplace).sOldValue), vbTextCompare) = 0 Then
                                                    ilFieldChanged = True
                                                    If ilPass = 1 Then
                                                        tmCurrDEE(llRow).sFixedTime = tgLibReplaceValues(ilReplace).sNewValue
                                                    End If
                                                End If
                                            End If
                                        ElseIf ilFieldType = 2 Then
                                            If Trim$(tgReplaceFields(ilField).sFieldName) = "Item ID" Then
                                                If StrComp(Trim$(tmGridDEE.sAudioItemID), Trim$(tgLibReplaceValues(ilReplace).sOldValue), vbTextCompare) = 0 Then
                                                    ilFieldChanged = True
                                                    If ilPass = 1 Then
                                                        tmCurrDEE(llRow).sAudioItemID = tgLibReplaceValues(ilReplace).sNewValue
                                                    End If
                                                End If
                                                If StrComp(Trim$(tmGridDEE.sProtItemID), Trim$(tgLibReplaceValues(ilReplace).sOldValue), vbTextCompare) = 0 Then
                                                    ilFieldChanged = True
                                                    If ilPass = 1 Then
                                                        tmCurrDEE(llRow).sProtItemID = tgLibReplaceValues(ilReplace).sNewValue
                                                    End If
                                                End If
                                            End If
                                            If Trim$(tgReplaceFields(ilField).sFieldName) = "ISCI" Then
                                                If StrComp(Trim$(tmGridDEE.sAudioISCI), Trim$(tgLibReplaceValues(ilReplace).sOldValue), vbTextCompare) = 0 Then
                                                    ilFieldChanged = True
                                                    If ilPass = 1 Then
                                                        tmCurrDEE(llRow).sAudioISCI = tgLibReplaceValues(ilReplace).sNewValue
                                                    End If
                                                End If
                                                If StrComp(Trim$(tmGridDEE.sProtISCI), Trim$(tgLibReplaceValues(ilReplace).sOldValue), vbTextCompare) = 0 Then
                                                    ilFieldChanged = True
                                                    If ilPass = 1 Then
                                                        tmCurrDEE(llRow).sProtISCI = tgLibReplaceValues(ilReplace).sNewValue
                                                    End If
                                                End If
                                            End If
                                        End If
                                        If ilFieldChanged Then
                                            ilSkipPass1 = False
                                            If (Not ilAllBusesMatch) Or (Not ilAllHoursMatch) Then
                                                'Remove Buses and Hours from Current record and make new row with buses and hours
                                                If ilPass = 0 Then
                                                    ilSplit = True
                                                    llFromRow = llRow
                                                    llToRow = UBound(tmCurrDEE)
                                                    ReDim Preserve tmCurrDEE(0 To UBound(tmCurrDEE) + 1) As DEE
                                                    ReDim Preserve smEBEBuses(0 To UBound(smEBEBuses) + 1) As String
                                                    LSet tmCurrDEE(llToRow) = tmCurrDEE(llFromRow)
                                                    tmCurrDEE(llToRow).lCode = 0
                                                    If (Not ilAllBusesMatch) Then
                                                        smEBEBuses(llFromRow) = ""
                                                        smEBEBuses(llToRow) = ""
                                                        For ilGLoop = LBound(smGridValues) To UBound(smGridValues) Step 1
                                                            ilBusMatch = False
                                                            For ilRLoop = LBound(smReplaceValues) To UBound(smReplaceValues) Step 1
                                                                If StrComp(Trim$(smGridValues(ilGLoop)), Trim$(smReplaceValues(ilRLoop)), vbTextCompare) = 0 Then
                                                                    ilBusMatch = True
                                                                    Exit For
                                                                End If
                                                            Next ilRLoop
                                                            If ilBusMatch Then
                                                                If smEBEBuses(llToRow) = "" Then
                                                                    smEBEBuses(llToRow) = smGridValues(ilGLoop)
                                                                Else
                                                                    smEBEBuses(llToRow) = smEBEBuses(llToRow) & "," & smGridValues(ilGLoop)
                                                                End If
                                                            Else
                                                                If smEBEBuses(llFromRow) = "" Then
                                                                    smEBEBuses(llFromRow) = smGridValues(ilGLoop)
                                                                Else
                                                                    smEBEBuses(llFromRow) = smEBEBuses(llFromRow) & "," & smGridValues(ilGLoop)
                                                                End If
                                                            End If
                                                        Next ilGLoop
                                                    End If
                                                    If (Not ilAllHoursMatch) Then
                                                        slFromHours = String(24, "N")
                                                        slToHours = String(24, "N")
                                                        For ilGLoop = 1 To 24 Step 1
                                                            If (Mid$(slGridHours, ilGLoop, 1) = "Y") And (Mid$(slReplaceHours, ilGLoop, 1) = "Y") Then
                                                                Mid$(slToHours, ilGLoop, 1) = "Y"
                                                            End If
                                                            If (Mid$(slGridHours, ilGLoop, 1) = "Y") And (Mid$(slReplaceHours, ilGLoop, 1) = "N") Then
                                                                Mid$(slFromHours, ilGLoop, 1) = "Y"
                                                            End If
                                                        Next ilGLoop
                                                        tmCurrDEE(llFromRow).sHours = slFromHours
                                                        tmCurrDEE(llToRow).sHours = slToHours
                                                    End If
                                                    Exit For
                                                End If
                                            Else
                                                'Insert DEE and AIE
                                            End If
                                        End If
                                    Else
                                        'Insert DEE only
                                    End If
                                End If
                            Next ilField
                            If ilSplit Then
                                Exit For
                            End If
                        Next ilReplace
                    Next llRow
                    If ilSkipPass1 Then
                        Exit For
                    End If
                Next ilPass
            End If
        End If
    Next ilLoop
    
End Sub



Private Sub mSetOverlapFlag()
    Dim slStartDate As String
    Dim slEndDate As String
    
    imOverlapCase = 0
    slStartDate = tmDHE.sStartDate
    slEndDate = tmDHE.sEndDate
    'Dtermine the case: 1: Total replace; 2: Terminate Old; 3: Change Start date of Old; 4: Split old into two parts
    If (lmEffStartDate <= gDateValue(slStartDate)) And (lmEffEndDate >= gDateValue(slEndDate)) Then
        'smOverlapMsg = "Replace Current with this version"
        imOverlapCase = 1
    ElseIf (lmEffStartDate > gDateValue(slStartDate)) And (lmEffEndDate >= gDateValue(slEndDate)) Then
        'smOverlapMsg = "Terminating Current as of " & Format$(llStartDate - 1, "ddddd")
        imOverlapCase = 2
    ElseIf (lmEffStartDate <= gDateValue(slStartDate)) And (lmEffEndDate >= gDateValue(slEndDate)) Then
        'smOverlapMsg = "Changing Start date of Current to " & Format$(llEndDate + 1, "ddddd")
        imOverlapCase = 3
    ElseIf (lmEffStartDate > gDateValue(slStartDate)) And (lmEffEndDate < gDateValue(slEndDate)) Then
        'smOverlapMsg = "Splitting Current into two pieces. One ending on " & Format$(llStartDate - 1, "ddddd") & " and the other start on " & Format$(llEndDate + 1, "ddddd")
        imOverlapCase = 4
    End If

End Sub

Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long
    
    If grdLib.TextMatrix(llRow, NAMEINDEX) <> "" Then
        For llCol = NAMEINDEX To STATEINDEX Step 1
            grdLib.Row = llRow
            grdLib.Col = llCol
            If grdLib.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
                grdLib.CellBackColor = vbWhite
            Else
                grdLib.CellBackColor = GRAY
            End If
        Next llCol
    End If
End Sub


Private Sub mSetRowSelectedCount()
    Dim llRow As Long
    Dim slDate As String
    Dim slLatestAirDate As String
    
    imRowSelectFlag = 0
    imSelectInFuture = True
    '7/10/11: Added
    If lmLastClickedRow < grdLib.FixedRows Then
        Exit Sub
    End If
    '7/10/11: End of add
    slLatestAirDate = gGetLatestSchdDate(True)
    '7/10/11: Handle single selection only
    '7/10/11: Removed for loop and set llRow = lmLastClickedRow
    'For llRow = grdLib.FixedRows To grdLib.Rows - 1 Step 1
        llRow = lmLastClickedRow
        If grdLib.TextMatrix(llRow, NAMEINDEX) <> "" Then
            If grdLib.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                imRowSelectFlag = imRowSelectFlag + 1
                slDate = grdLib.TextMatrix(llRow, STARTDATEINDEX)
                If (gDateValue(slDate) <= gDateValue(slLatestAirDate)) Or (gDateValue(slDate) <= lmNowDate) Then
                    If imSelectInFuture = True Then
                        imSelectInFuture = False
                    End If
                End If
            End If
        End If
    'Next llRow

End Sub
