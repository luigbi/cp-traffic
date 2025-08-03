VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form EngrTemp 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrTemp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11790
   Begin V10EngineeringDev.CSI_Calendar cccEffEndDate 
      Height          =   285
      Left            =   6855
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
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   10515
      Top             =   6585
   End
   Begin VB.CheckBox ckcDormant 
      Caption         =   "Include Dormant"
      Height          =   210
      Left            =   105
      TabIndex        =   12
      Top             =   6615
      Width           =   1770
   End
   Begin V10EngineeringDev.CSI_Calendar cccEffStartDate 
      Height          =   285
      Left            =   5040
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
      CSI_DefaultDateType=   0
   End
   Begin VB.PictureBox pbcEffTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   7800
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   -45
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
      Width           =   795
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   11115
      Top             =   5610
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   -15
      ScaleHeight     =   90
      ScaleWidth      =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   270
      Width           =   0
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   11175
      Top             =   6315
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTemp 
      Height          =   5880
      Left            =   45
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   10372
      _Version        =   393216
      Cols            =   8
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
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lacEffEndDate 
      Caption         =   "End"
      Height          =   225
      Left            =   6375
      TabIndex        =   3
      Top             =   165
      Width           =   450
   End
   Begin VB.Label lacEffDate 
      Caption         =   "Effective Date- Start"
      Height          =   225
      Left            =   3480
      TabIndex        =   1
      Top             =   165
      Width           =   1575
   End
   Begin VB.Label lacScreen 
      Caption         =   "Templates"
      Height          =   255
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   2625
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   9555
      Picture         =   "EngrTemp.frx":030A
      Top             =   6540
      Width           =   480
   End
End
Attribute VB_Name = "EngrTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrTemp - enters affiliate representative information
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
Private smDates() As String
Private smGridValues() As String
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

Private smCurrTSEStamp As String
Private tmCurrTSE() As TSE

Private tmDHE As DHE

'Grid Controls
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private imShowGridBox As Integer
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on
Private imLastColSorted As Integer
Private imLastSort As Integer
Private imTerminate As Integer

Private lmLastClickedRow As Long

Const NAMEINDEX = 0
Const DESCRIPTIONINDEX = 1
Const DATESINDEX = 2
Const HOURSINDEX = 3    '7
'Const BUSESINDEX = 4    '8
Const STATEINDEX = 4    '5    '9
Const SELECTEDINDEX = 5
Const SORTINDEX = 6
Const CODEINDEX = 7 '6 '10

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
    If UBound(tgCurrTempDHE) > LBound(tgCurrTempDHE) Then
        tmcClick_Timer
    End If
End Sub

Private Sub cmcChange_Click()
    Dim llRow As Long
    
    lgTempCallCode = -1
    For llRow = grdTemp.FixedRows To grdTemp.Rows - 1 Step 1
        If grdTemp.TextMatrix(llRow, NAMEINDEX) <> "" Then
            If grdTemp.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                lgTempCallCode = grdTemp.TextMatrix(llRow, CODEINDEX)
                Exit For
            End If
        End If
    Next llRow
    If lgTempCallCode <= 0 Then
        Exit Sub
    End If
    If (Not imSelectInFuture) Or (cmcChange.Caption = "&View") Then
        igTempCallType = 3   'View
    Else
        igTempCallType = 1   'Change
    End If
    EngrTempDef.Show vbModeless
    Unload EngrTemp
End Sub


Private Sub cmcNew_Click()
    Dim llRow As Long
    
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(TEMPLATEJOB) = 2) Then
        If imRowSelectFlag = 1 Then
            lgTempCallCode = -1
            For llRow = grdTemp.FixedRows To grdTemp.Rows - 1 Step 1
                If grdTemp.TextMatrix(llRow, NAMEINDEX) <> "" Then
                    If grdTemp.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                        lgTempCallCode = grdTemp.TextMatrix(llRow, CODEINDEX)
                        Exit For
                    End If
                End If
            Next llRow
            If lgTempCallCode <= 0 Then
                Exit Sub
            End If
            igTempCallType = 2   'Model
        Else
            lgTempCallCode = 0
            igTempCallType = 0   'New
        End If
        EngrTempDef.Show vbModeless
        Unload EngrTemp
    End If
End Sub

Private Sub cmcSearch_Click()
    Dim llRow As Long
    Dim slStr As String
    For llRow = grdTemp.FixedRows To grdTemp.Rows - 1 Step 1
        If grdTemp.TextMatrix(llRow, NAMEINDEX) <> "" Then
            If grdTemp.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                grdTemp.TextMatrix(llRow, SELECTEDINDEX) = "0"
                mPaintRowColor llRow
            End If
        End If
    Next llRow
    slStr = UCase(Trim$(edcSearch.text))
    For llRow = grdTemp.FixedRows To grdTemp.Rows - 1 Step 1
        If grdTemp.TextMatrix(llRow, NAMEINDEX) <> "" Then
            If InStr(1, UCase(Trim$(grdTemp.TextMatrix(llRow, NAMEINDEX))), slStr, vbBinaryCompare) > 0 Then
                grdTemp.TextMatrix(llRow, SELECTEDINDEX) = "1"
                mPaintRowColor llRow
                grdTemp.TopRow = llRow
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
    Dim ilPos As Integer
    
    
    For llRow = grdTemp.FixedRows To grdTemp.Rows - 1 Step 1
        slStr = Trim$(grdTemp.TextMatrix(llRow, NAMEINDEX))
        If slStr <> "" Then
            If (ilCol = DATESINDEX) Then
                slStr = grdTemp.TextMatrix(llRow, DATESINDEX)
                If InStr(1, slStr, "No Dates", vbTextCompare) <= 0 Then
                    ilPos = InStr(1, slStr, "-", vbTextCompare)
                    If ilPos > 0 Then
                        slStr = Left$(slStr, ilPos - 1)
                    End If
                    slStr = Trim$(Str$(gDateValue(slStr)))
                    Do While Len(slStr) < 6
                        slStr = "0" & slStr
                    Loop
                Else
                    slStr = "000000"
                End If
            Else
                slStr = grdTemp.TextMatrix(llRow, ilCol)
            End If
            grdTemp.TextMatrix(llRow, SORTINDEX) = slStr & grdTemp.TextMatrix(llRow, SORTINDEX)
        End If
    Next llRow
    If imLastColSorted = ilCol Then
        gGrid_SortByCol grdTemp, NAMEINDEX, SORTINDEX, SORTINDEX, imLastSort
    Else
        gGrid_SortByCol grdTemp, NAMEINDEX, SORTINDEX, imLastColSorted, imLastSort
    End If
    imLastColSorted = ilCol

End Sub

Private Sub mSetCommands()
    Dim ilRet As Integer
    If UBound(tgCurrTempDHE) <= LBound(tgCurrTempDHE) Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(TEMPLATEJOB) = 2) Then
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
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(TEMPLATEJOB) = 2) Then
            If imRowSelectFlag = 1 Then
                If imSelectInFuture Then
                    cmcChange.Caption = "&Change"
                Else
                    cmcChange.Caption = "&View"
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
    
    gGrid_AlignAllColsLeft grdTemp
    mGridColumnWidth
    'Set Titles
    grdTemp.TextMatrix(0, NAMEINDEX) = "Name/Subname"
    grdTemp.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdTemp.TextMatrix(0, DATESINDEX) = "Dates"
    grdTemp.TextMatrix(0, HOURSINDEX) = "Offset Hours"
    'grdTemp.TextMatrix(BUSESINDEX + 1) = "Bus"
    grdTemp.TextMatrix(0, STATEINDEX) = "State"
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    Dim llWidth As Single
    
    grdTemp.Width = EngrTemp.Width - 2 * grdTemp.Left - 120
    grdTemp.ColWidth(SELECTEDINDEX) = 0
    grdTemp.ColWidth(SORTINDEX) = 0
    grdTemp.ColWidth(CODEINDEX) = 0
    grdTemp.ColWidth(NAMEINDEX) = grdTemp.Width / 5
    grdTemp.ColWidth(DATESINDEX) = grdTemp.Width / 5    '9
    grdTemp.ColWidth(HOURSINDEX) = grdTemp.Width / 9
    'grdTemp.Colwidth(BUSESINDEX + 1).Width = grdTemp.Width / 9
    grdTemp.ColWidth(STATEINDEX) = grdTemp.Width / 20
    llWidth = 0
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If (ilCol <> DESCRIPTIONINDEX) Then
            llWidth = llWidth + grdTemp.ColWidth(ilCol)
        End If
    Next ilCol
    grdTemp.ColWidth(DESCRIPTIONINDEX) = grdTemp.Width - llWidth - GRIDSCROLLWIDTH '- 6 * 240 '- 30

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
    Dim slDates As String
    Dim llDate As Long
    Dim ilTest As Integer
    Dim llCol As Long
    
'    Set mItem = lbcTemp.ListItems.Add()
'    llRow = mItem.Height
    grdTemp.Redraw = False
    grdTemp.Rows = grdTemp.FixedRows + 1
    
    For llCol = NAMEINDEX To STATEINDEX Step 1
        grdTemp.Row = grdTemp.FixedRows
        grdTemp.Col = llCol
        grdTemp.CellBackColor = vbWhite
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
    slDates = ""
    llRow = grdTemp.FixedRows
    For ilLoop = 0 To UBound(tgCurrTempDHE) - 1 Step 1
        ilRet = gGetRecs_TSE_TemplateSchd(smCurrTSEStamp, tgCurrTempDHE(ilLoop).lCode, "EngrTemp- mMoveRecToCtrl for TSE", tmCurrTSE())
        ilShow = False
'        slDates = ""
'        If UBound(tmCurrTSE) <= LBound(tmCurrTSE) Then
'            ilShow = True
'        Else
'            For ilTest = 0 To UBound(tmCurrTSE) - 1 Step 1
'                llDate = gDateValue(tmCurrTSE(ilTest).sLogDate)
'                If (llDate >= llSdate) And (llDate <= llEDate) Then
'                    If slDates = "" Then
'                        slDates = tmCurrTSE(ilTest).sLogDate
'                    Else
'                        slDates = slDates & ", " & tmCurrTSE(ilTest).sLogDate
'                    End If
'                    ilShow = True
'                End If
'            Next ilTest
'        End If
        slDates = gGetTempDateRange(llSdate, llEDate, tmCurrTSE())
        If slDates = "" Then
            ilShow = False
        Else
            ilShow = True
        End If
        If ckcDormant.Value = vbUnchecked Then
            If tgCurrTempDHE(ilLoop).sState = "D" Then
                ilShow = False
            End If
        End If
        If ilShow Then
            If llRow >= grdTemp.Rows Then
                grdTemp.AddItem ""
            End If
            slStr = ""
            For ilDNE = 0 To UBound(tgCurrTempDNE) - 1 Step 1
                If tgCurrTempDHE(ilLoop).lDneCode = tgCurrTempDNE(ilDNE).lCode Then
                    slStr = Trim$(tgCurrTempDNE(ilDNE).sName)
                    Exit For
                End If
            Next ilDNE
            For ilDSE = 0 To UBound(tgCurrDSE) - 1 Step 1
                If tgCurrTempDHE(ilLoop).lDseCode = tgCurrDSE(ilDSE).lCode Then
                    slStr = slStr & "/" & Trim$(tgCurrDSE(ilDSE).sName)
                    Exit For
                End If
            Next ilDSE
            grdTemp.TextMatrix(llRow, NAMEINDEX) = slStr
            ilRet = gGetRec_CTE_CommtsTitle(tgCurrTempDHE(ilLoop).lCteCode, "EngrTemp- mMoveRecToCtrl for CTE", tmCTE)
            grdTemp.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tmCTE.sComment)
            grdTemp.TextMatrix(llRow, DATESINDEX) = slDates
            slStr = gHourMap(tgCurrTempDHE(ilLoop).sHours)
            grdTemp.TextMatrix(llRow, HOURSINDEX) = Trim$(slStr)
'            ilRet = gGetRecs_DBE_DayBusSel(smCurrDBE, tgCurrTempDHE(ilLoop).lCode, "EngrTemp- mMoveRecToCtrl for DBE", tmCurrDBE())
'            slStr = ""
'            For ilDBE = 0 To UBound(tmCurrDBE) - 1 Step 1
'                If tmCurrDBE(ilDBE).sType = "B" Then
'                    For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
'                        If tmCurrDBE(ilDBE).iBdeCode = tgCurrBDE(ilBDE).iCode Then
'                            If slStr = "" Then
'                                slStr = Trim$(tgCurrBDE(ilBDE).sName)
'                            Else
'                                slStr = slStr & ", " & Trim$(tgCurrBDE(ilBDE).sName)
'                            End If
'                            Exit For
'                        End If
'                    Next ilBDE
'                End If
'            Next ilDBE
'            mItem.SubItems(BUSESINDEX) = slStr
            If tgCurrTempDHE(ilLoop).sState = "A" Then
                grdTemp.TextMatrix(llRow, STATEINDEX) = "Active"
            ElseIf tgCurrTempDHE(ilLoop).sState = "L" Then
                grdTemp.TextMatrix(llRow, STATEINDEX) = "Limbo"
            Else
                grdTemp.TextMatrix(llRow, STATEINDEX) = "Dormant"
            End If
            grdTemp.TextMatrix(llRow, CODEINDEX) = tgCurrTempDHE(ilLoop).lCode
            grdTemp.TextMatrix(llRow, SELECTEDINDEX) = "0"
            llRow = llRow + 1
        End If
    Next ilLoop
    gGrid_IntegralHeight grdTemp
    gGrid_FillWithRows grdTemp
    grdTemp.Height = grdTemp.Height + 30
    grdTemp.Visible = True
    grdTemp.Redraw = True
    If lgTempTopRow <> -1 Then
        grdTemp.TopRow = lgTempTopRow
    End If
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    ilRet = gGetTypeOfRecs_DHE_DayHeaderInfo("C", "T", sgCurrTempDHEStamp, "EngrTemp-mPopulate", tgCurrTempDHE())
    ilRet = gGetTypeOfRecs_DNE_DayName("C", "T", sgCurrTempDNEStamp, "EngrTemp-mPopulate Template Names", tgCurrTempDNE())
    ilRet = gGetTypeOfRecs_DSE_DaySubName("C", sgCurrDSEStamp, "EngrTemp-mPopulate SubNames", tgCurrDSE())
    ilRet = gGetTypeOfRecs_BGE_BusGroup("C", sgCurrBGEStamp, "EngrTemp-mPopulate Bus Groups", tgCurrBGE())
    ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrTemp-Bus Definitions", tgCurrBDE())
    
    
End Sub
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    igJobShowing(TEMPLATEJOB) = 0
    Unload EngrTemp
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
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrTemp
    'gCenterFormModal EngrTemp
    gCenterForm EngrTemp
End Sub

Private Sub Form_Load()
    mGridColumns
    mInit
    igJobShowing(TEMPLATEJOB) = 1
End Sub

Private Sub Form_Resize()
    'These call are here and in form_Active (call to mGridColumns)
    'They are in mGridColumn in case the For_Initialize size chage does not cause a resize event
    grdTemp.Visible = False
    mGridColumnWidth
    grdTemp.Height = cmcCancel.Top - grdTemp.Top - 240    '8 * grdTemp.RowHeight(0) + 30
    grdTemp.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lgTempTopRow = grdTemp.TopRow
    Erase smHours
    Erase smBuses
    Erase smGridValues
    Erase tmCurrDBE
    Erase tmCurrEBE
    Erase smEBEBuses
    Erase tmCurrDEE
    Erase tmCurrTSE
    Set EngrTemp = Nothing
End Sub





Private Sub mInit()
    Dim llRet As Long
    
    On Error GoTo ErrHand
    
    gSetMousePointer grdTemp, grdTemp, vbHourglass
    imcPrint.Picture = EngrMain!imcPrinter.Picture
    cmcSearch.Top = 30
    edcSearch.Top = cmcSearch.Top
    smNowDate = Format$(gNow(), "ddddd")
    lmNowDate = gDateValue(smNowDate)
    imTerminate = False
    ''edcEffStartDate.Text = smNowDate
    'cccEffStartDate.text = smNowDate
    tmcClick.Enabled = False
    imLastColSorted = -1
    imLastSort = -1
    lmEnableRow = -1
    imRowSelectFlag = -1
    imSelectInFuture = False
    imInChg = True
    mPopulate
    'mMoveRecToCtrls
    'mSortCol NAMEINDEX
    imInChg = False
    mSetCommands
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igJobStatus(TEMPLATEJOB) = 2) Then
    Else
        cmcChange.Caption = "&View"
        cmcNew.Enabled = False
    End If
    
    tmcStart.Enabled = True
    gSetMousePointer grdTemp, grdTemp, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdTemp, grdTemp, vbDefault
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
        If UBound(tgCurrTempDHE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
'    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
'        For llRow = 1 To lbcTemp.ListCount - 1 Step 1
'            slStr = Trim$(lbcTemp.Text(llRow, NAMEINDEX))
'            If (slStr <> "") Then
'                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
'                    grdTemp.Row = llRow
'                    Do While Not grdTemp.RowIsVisible(grdTemp.Row)
'                        grdTemp.TopRow = grdTemp.TopRow + 1
'                    Loop
'                    Exit Sub
'                End If
'            End If
'        Next llRow
'    End If
    
End Sub


Private Sub grdTemp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer

    If y < grdTemp.RowHeight(0) Then
        grdTemp.Col = grdTemp.MouseCol
        mSortCol grdTemp.Col
        grdTemp.Row = 0
        grdTemp.Col = CODEINDEX
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdTemp, x, y, llCurrentRow, llCol)
    If llCurrentRow < grdTemp.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdTemp.FixedRows Then
        imRowSelectFlag = 0
        If grdTemp.TextMatrix(llCurrentRow, NAMEINDEX) <> "" Then
            'grdTemp.TopRow = lmScrollTop
            llTopRow = grdTemp.TopRow
            If (Shift And CTRLMASK) > 0 Then
                If grdTemp.TextMatrix(llCurrentRow, SELECTEDINDEX) <> 1 Then
                    grdTemp.TextMatrix(llCurrentRow, SELECTEDINDEX) = 1
                    '7/10/11: Added
                    lmLastClickedRow = llCurrentRow
                Else
                    grdTemp.TextMatrix(llCurrentRow, SELECTEDINDEX) = 0
                    '7/10/11: Added
                    lmLastClickedRow = -1
                End If
                mPaintRowColor llCurrentRow
                mSetRowSelectedCount
            Else
                '7/10/11: Disallow multi-row selection.  i.e. removing Replace.  Support routines missing but defined in libDef
                'For llRow = grdTemp.FixedRows To grdTemp.Rows - 1 Step 1
                '    If grdTemp.TextMatrix(llRow, NAMEINDEX) <> "" Then
                '        grdTemp.TextMatrix(llRow, SELECTEDINDEX) = "0"
                '        If (lmLastClickedRow = -1) Or ((Shift And SHIFTMASK) <= 0) Then
                '            If llRow = llCurrentRow Then
                '                grdTemp.TextMatrix(llRow, SELECTEDINDEX) = "1"
                '            Else
                '                grdTemp.TextMatrix(llRow, SELECTEDINDEX) = "0"
                '            End If
                '        ElseIf lmLastClickedRow < llCurrentRow Then
                '            If (llRow >= lmLastClickedRow) And (llRow <= llCurrentRow) Then
                '                grdTemp.TextMatrix(llRow, SELECTEDINDEX) = "1"
                '            End If
                '        Else
                '            If (llRow >= llCurrentRow) And (llRow <= lmLastClickedRow) Then
                '                grdTemp.TextMatrix(llRow, SELECTEDINDEX) = "1"
                '            End If
                '        End If
                '        mPaintRowColor llRow
                '    End If
                'Next llRow
                '7/10/11: Added
                If lmLastClickedRow >= grdTemp.FixedRows Then
                    grdTemp.TextMatrix(lmLastClickedRow, SELECTEDINDEX) = "0"
                    mPaintRowColor lmLastClickedRow
                End If
                grdTemp.TextMatrix(llCurrentRow, SELECTEDINDEX) = "1"
                mPaintRowColor llCurrentRow
                lmLastClickedRow = llCurrentRow
                '7/10/11: End of add
                mSetRowSelectedCount
                grdTemp.TopRow = llTopRow
                grdTemp.Row = llCurrentRow
            End If
            '7/10/11: Removed
            'lmLastClickedRow = llCurrentRow
        End If
        mSetCommands
    End If

End Sub

Private Sub imcPrint_Click()
    igRptIndex = TEMPLATE_RPT
    igRptSource = vbModal
    EngrTemplateRpt.Show vbModal
End Sub

Private Sub pbcEffTab_GotFocus()
    If tmcClick.Enabled Then
        tmcClick.Enabled = False
        gSetMousePointer grdTemp, grdTemp, vbHourglass
        mMoveRecToCtrls
        mSortCol 1
        mSetCommands
        gSetMousePointer grdTemp, grdTemp, vbDefault
    End If
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    gSetMousePointer grdTemp, grdTemp, vbHourglass
    mMoveRecToCtrls
    mSortCol NAMEINDEX
    mSetCommands
    gSetMousePointer grdTemp, grdTemp, vbDefault
End Sub

Private Sub mPaintRowColor(llRow As Long)

    Dim llCol As Long
    If grdTemp.TextMatrix(llRow, NAMEINDEX) <> "" Then
        For llCol = NAMEINDEX To STATEINDEX Step 1
            grdTemp.Row = llRow
            grdTemp.Col = llCol
            If grdTemp.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
                grdTemp.CellBackColor = vbWhite
            Else
                grdTemp.CellBackColor = GRAY
            End If
        Next llCol
    End If
End Sub

Private Sub mSetRowSelectedCount()
    Dim llRow As Long
    Dim slDate As String
    Dim slLatestAirDate As String
    Dim ilPos As Integer
    Dim ilDates As Long
    
    imRowSelectFlag = 0
    imSelectInFuture = True
    '7/10/11: Added
    If lmLastClickedRow < grdTemp.FixedRows Then
        Exit Sub
    End If
    '7/10/11: End of add
    slLatestAirDate = gGetLatestSchdDate(True)
    '7/10/11: Handle single selection only
    '7/10/11: Removed for loop and set llRow = lmLastClickedRow
    'For llRow = grdTemp.FixedRows To grdTemp.Rows - 1 Step 1
        llRow = lmLastClickedRow
        If grdTemp.TextMatrix(llRow, NAMEINDEX) <> "" Then
            If grdTemp.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                imRowSelectFlag = imRowSelectFlag + 1
                '7/11/11:  All Template to be changed or added regardless of current scheduled dates
                'slDate = grdTemp.TextMatrix(llRow, DATESINDEX)
                'If (StrComp(slDate, "No Dates", vbTextCompare) <> 0) Then
                '    If (gDateValue(slDate) <= gDateValue(slLatestAirDate)) Or (gDateValue(slDate) <= lmNowDate) Then
                '        If imSelectInFuture = True Then
                '            imSelectInFuture = False
                '        End If
                '    End If
                '    slDate = grdTemp.TextMatrix(llRow, DATESINDEX)
                '    If (slDate <> "") And (StrComp(slDate, "No Dates", vbTextCompare) <> 0) Then
                '        ilPos = InStr(1, slDate, "-", vbTextCompare)
                '        If ilPos > 0 Then
                '            Mid$(slDate, ilPos, 1) = ","
                '        End If
                '        gParseCDFields slDate, False, smDates()
                '        For ilDates = LBound(smDates) To UBound(smDates) Step 1
                '            If smDates(ilDates) <> "" Then
                '                If imSelectInFuture = True Then
                '                    imSelectInFuture = False
                '                End If
                '            End If
                '        Next ilDates
                '    End If
                'End If
            End If
        End If
    'Next llRow
    Erase smDates
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    If imTerminate Then
        cmcCancel_Click
    Else
        tmcClick_Timer
    End If
End Sub
