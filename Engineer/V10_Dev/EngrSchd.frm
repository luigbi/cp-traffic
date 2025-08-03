VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form EngrSchd 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrSchd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11790
   Begin VB.CommandButton cmcRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   8025
      TabIndex        =   8
      Top             =   6645
      Width           =   1335
   End
   Begin VB.CommandButton cmcImportAsAir 
      Caption         =   "&Import As Air Log"
      Height          =   375
      Left            =   7605
      TabIndex        =   6
      Top             =   7080
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.CommandButton cmcAsAirCompare 
      Caption         =   "&As Air Compare"
      Height          =   375
      Left            =   9810
      TabIndex        =   5
      Top             =   7080
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.CommandButton cmcChange 
      Caption         =   "&Change"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   6645
      Width           =   1335
   End
   Begin VB.CommandButton cmcNew 
      Caption         =   "&New from Scratch"
      Height          =   375
      Left            =   5820
      TabIndex        =   3
      Top             =   6645
      Width           =   1890
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   11115
      Top             =   5610
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   90
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   255
      Width           =   60
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   11520
      Top             =   4770
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
      Left            =   4170
      TabIndex        =   2
      Top             =   6645
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSchd 
      Height          =   5880
      Left            =   75
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   375
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   10372
      _Version        =   393216
      Cols            =   13
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
      _Band(0).Cols   =   13
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lacScreen 
      Caption         =   "Schedule"
      Height          =   270
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   2625
   End
End
Attribute VB_Name = "EngrSchd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrSchd - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private hmSEE As Integer

Private imInChg As Integer
Private smNowDate As String
Private lmNowDate As Long
Private imRowSelectFlag As Integer '0=Zero rows selected; 1=One row selected; 2=2 or more rows selected
Private imSelectInFuture As Integer

Private smEffStartDate As String
Private lmEffStartDate As Long
Private smEffEndDate As String
Private lmEffEndDate As Long

Private smSEEStamp As String
Private tmSHE() As SHE
Private tmSEE() As SEE

'Grid Controls
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private imShowGridBox As Integer
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on
Private imLastColSorted As Integer
Private imLastSort As Integer
Private lmHighlightRow As Long

Const DATEINDEX = 0
Const GETCOUNTSINDEX = 6
Const NOEVENTSINDEX = 7 '1
Const NOFILLEDAVAILSINDEX = 8   '2
Const NOSPLITAVAILSINDEX = 9   '3
Const NOEMPTYAVAILSINDEX = 10   '4
Const CONFLICTINDEX = 1 '5
Const ITEMCHECKINDEX = 2    '6
Const SPOTMERGESTATUSINDEX = 3  '7
'Const EXPORTEDINDEX = 4 '8
Const LOADSTATUSINDEX = 4   '9
Const IMPORTEDINDEX = 5 '10
Const SHECODEINDEX = 11
Const SORTINDEX = 12






Private Sub cmcAsAirCompare_Click()
    If lmHighlightRow < 0 Then
        Exit Sub
    End If
    sgAsAirCompareDate = grdSchd.TextMatrix(lmHighlightRow, DATEINDEX)
    EngrAsAirCompare.Show vbModeless
    Unload EngrSchd
End Sub

Private Sub cmcChange_Click()
    
    If lmHighlightRow < 0 Then
        Exit Sub
    End If
    igSchdCallType = 1
    sgSchdDate = grdSchd.TextMatrix(lmHighlightRow, DATEINDEX)
    ReDim tgFilterValues(0 To 0) As FILTERVALUES
    ReDim tgSchdReplaceValues(0 To 0) As SCHDREPLACEVALUES
    'mCreateUsedArrays
    'mInitFilterInfo
    EngrSchdFilter.Show vbModal
    If igAnsFilter = CALLCANCELLED Then
        Exit Sub
    End If
    EngrSchdDef.Show vbModeless
    Unload EngrSchd
End Sub



Private Sub cmcImportAsAir_Click()
    If lmHighlightRow < 0 Then
        Exit Sub
    End If
    sgAsAirLogDate = grdSchd.TextMatrix(lmHighlightRow, DATEINDEX)
    EngrImportAsAir.Show vbModeless
    Unload EngrSchd
End Sub

Private Sub cmcNew_Click()
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(SCHEDULEJOB) = 2) Then
        ReDim tgFilterValues(0 To 0) As FILTERVALUES
        ReDim tgSchdReplaceValues(0 To 0) As SCHDREPLACEVALUES
        igSchdCallType = 0
        sgSchdDate = ""
        EngrSchdDef.Show vbModeless
        Unload EngrSchd
    End If
End Sub






Private Sub mSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    For llRow = grdSchd.FixedRows To grdSchd.Rows - 1 Step 1
        slStr = Trim$(grdSchd.TextMatrix(llRow, DATEINDEX))
        If slStr <> "" Then
            If (ilCol = DATEINDEX) Then
                slStr = grdSchd.TextMatrix(llRow, DATEINDEX)
                slStr = Trim$(Str$(gDateValue(slStr)))
                Do While Len(slStr) < 6
                    slStr = "0" & slStr
                Loop
            ElseIf (ilCol = NOFILLEDAVAILSINDEX) Then
                slStr = grdSchd.TextMatrix(llRow, NOFILLEDAVAILSINDEX)
                Do While Len(slStr) < 6
                    slStr = "0" & slStr
                Loop
            ElseIf (ilCol = NOSPLITAVAILSINDEX) Then
                slStr = grdSchd.TextMatrix(llRow, NOSPLITAVAILSINDEX)
                Do While Len(slStr) < 6
                    slStr = "0" & slStr
                Loop
            ElseIf (ilCol = NOSPLITAVAILSINDEX) Then
                slStr = grdSchd.TextMatrix(llRow, NOSPLITAVAILSINDEX)
                Do While Len(slStr) < 6
                    slStr = "0" & slStr
                Loop
            ElseIf (ilCol = NOEMPTYAVAILSINDEX) Then
                slStr = grdSchd.TextMatrix(llRow, NOEMPTYAVAILSINDEX)
                Do While Len(slStr) < 6
                    slStr = "0" & slStr
                Loop
            ElseIf (ilCol = ITEMCHECKINDEX) Then
                slStr = grdSchd.TextMatrix(llRow, ITEMCHECKINDEX)
                slStr = Trim$(Str$(gDateValue(slStr)))
                Do While Len(slStr) < 6
                    slStr = "0" & slStr
                Loop
            'ElseIf (ilCol = EXPORTEDINDEX) Then
            '    slStr = grdSchd.TextMatrix(llRow, EXPORTEDINDEX)
            '    slStr = Trim$(Str$(gDateValue(slStr)))
            '    Do While Len(slStr) < 6
            '        slStr = "0" & slStr
            '    Loop
            ElseIf (ilCol = IMPORTEDINDEX) Then
                slStr = grdSchd.TextMatrix(llRow, IMPORTEDINDEX)
                Do While Len(slStr) < 6
                    slStr = "0" & slStr
                Loop
            Else
                slStr = grdSchd.TextMatrix(llRow, ilCol)
            End If
            grdSchd.TextMatrix(llRow, SORTINDEX) = slStr & grdSchd.TextMatrix(llRow, SORTINDEX)
        End If
    Next llRow
    If imLastColSorted = ilCol Then
        gGrid_SortByCol grdSchd, DATEINDEX, SORTINDEX, SORTINDEX, imLastSort
    Else
        gGrid_SortByCol grdSchd, DATEINDEX, SORTINDEX, imLastColSorted, imLastSort
    End If
    imLastColSorted = ilCol
End Sub

Private Sub mSetCommands()
    Dim ilRet As Integer
    cmcRemove.Enabled = False
    If UBound(tmSHE) <= LBound(tmSHE) Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(SCHEDULEJOB) = 2) Then
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
        cmcImportAsAir.Enabled = False
        cmcAsAirCompare.Enabled = False
    Else
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(SCHEDULEJOB) = 2) Then
            If imRowSelectFlag = 1 Then
                If imSelectInFuture Then
                    cmcChange.Caption = "&Change"
                Else
                    cmcChange.Caption = "&View"
                End If
                cmcChange.Enabled = True
                cmcCancel.Enabled = True
                cmcNew.Enabled = True
                cmcNew.Caption = "&New from Scratch"
                If (grdSchd.TextMatrix(lmHighlightRow + 1, DATEINDEX) = "") And (grdSchd.TextMatrix(lmHighlightRow, LOADSTATUSINDEX) = "") And (imSelectInFuture) Then
                    cmcRemove.Enabled = True
                End If
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
        If (imRowSelectFlag = 1) And (Not imSelectInFuture) Then
            cmcImportAsAir.Enabled = True
            cmcAsAirCompare.Enabled = True
        Else
            cmcImportAsAir.Enabled = False
            cmcAsAirCompare.Enabled = False
        End If
    End If
End Sub

Private Sub mGridColumns()
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    mGridColumnWidth
    'Set Titles
    grdSchd.TextMatrix(0, DATEINDEX) = "Date"
    grdSchd.TextMatrix(0, GETCOUNTSINDEX) = ""
    grdSchd.TextMatrix(0, NOEVENTSINDEX) = "# Events"
    grdSchd.TextMatrix(0, NOFILLEDAVAILSINDEX) = "# Filled Avails"
    grdSchd.TextMatrix(0, NOSPLITAVAILSINDEX) = "# Partial Avails"
    grdSchd.TextMatrix(0, NOEMPTYAVAILSINDEX) = "# Open Avails"
    grdSchd.TextMatrix(0, CONFLICTINDEX) = "Conflict"
    grdSchd.TextMatrix(0, ITEMCHECKINDEX) = "Items Checked"
    grdSchd.TextMatrix(0, SPOTMERGESTATUSINDEX) = "Merged"
    'grdSchd.TextMatrix(0, EXPORTEDINDEX) = "Exported"
    grdSchd.TextMatrix(0, LOADSTATUSINDEX) = "Load Created"
    grdSchd.TextMatrix(0, IMPORTEDINDEX) = "Imported"
    grdSchd.TextMatrix(0, SHECODEINDEX) = "SHE Code"
    grdSchd.TextMatrix(0, SORTINDEX) = "Sort"

End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    Dim llWidth As Single
    
    grdSchd.Left = 75
    grdSchd.Top = lacScreen.Top + lacScreen.Height + 120
    grdSchd.Width = EngrSchd.Width - 150
    grdSchd.ColWidth(SHECODEINDEX) = 0
    grdSchd.ColWidth(SORTINDEX) = 0
    grdSchd.ColWidth(DATEINDEX) = grdSchd.Width * 0.088
    grdSchd.ColWidth(CONFLICTINDEX) = grdSchd.Width * 0.06
    grdSchd.ColWidth(ITEMCHECKINDEX) = 0    'grdSchd.Width * 0.095
    'grdSchd.ColWidth(EXPORTEDINDEX) = grdSchd.Width * 0.088
    grdSchd.ColWidth(SPOTMERGESTATUSINDEX) = grdSchd.Width * 0.06
    grdSchd.ColWidth(LOADSTATUSINDEX) = grdSchd.Width * 0.14
    grdSchd.ColWidth(IMPORTEDINDEX) = 0 'grdSchd.Width * 0.088
    '7/8/11: All ways show events
    'If rbcShowEvtInfo(1).Value Then
    '    grdSchd.ColWidth(NOEVENTSINDEX) = 0
    '    grdSchd.ColWidth(NOFILLEDAVAILSINDEX) = 0
    '    grdSchd.ColWidth(NOSPLITAVAILSINDEX) = 0
    '    grdSchd.ColWidth(NOEMPTYAVAILSINDEX) = 0
    'Else
        grdSchd.ColWidth(GETCOUNTSINDEX) = grdSchd.Width * 0.05
        grdSchd.ColWidth(NOEVENTSINDEX) = grdSchd.Width * 0.09
        grdSchd.ColWidth(NOFILLEDAVAILSINDEX) = grdSchd.Width * 0.095
        grdSchd.ColWidth(NOSPLITAVAILSINDEX) = grdSchd.Width * 0.095
        grdSchd.ColWidth(NOEMPTYAVAILSINDEX) = grdSchd.Width * 0.095
    'End If
    llWidth = 0
    For ilCol = 0 To SORTINDEX Step 1
        llWidth = llWidth + grdSchd.ColWidth(ilCol)
    Next ilCol
    grdSchd.Width = llWidth + GRIDSCROLLWIDTH + 15
    grdSchd.Left = (EngrSchd.Width - grdSchd.Width) / 2
    'Align columns to left
    gGrid_AlignAllColsLeft grdSchd
    gGrid_IntegralHeight grdSchd
End Sub


Private Sub mClearControls()
    
End Sub


Private Sub mMoveRecToCtrls()
    Dim ilSHE As Integer
    Dim llSEE As Long
    Dim ilRet As Integer
    Dim slStr As String
    Dim llNoEvents As Long
    Dim llNoFilledAvails As Long
    Dim llNoSplitAvails As Long
    Dim llNoEmptyAvails As Long
    Dim ilNoSpotsInAvail As Integer
    Dim ilETE As Integer
    Dim slCategory As String
    Dim llAvailTest As Long
    Dim llAvailLength As Long
    Dim slItemCheckDate As String
    Dim mItem As ListItem
    Dim llCount As Long
    Dim llRow As Long
    Dim llCol As Long
    
    'lbcSchd.Visible = False
    grdSchd.Redraw = False
    DoEvents
    grdSchd.Visible = False
    mGridColumnWidth
    'lbcSchd.ListItems.Clear
    'lbcSchd.Visible = True
    
    grdSchd.Row = 0
    grdSchd.Rows = 2
    For llCol = DATEINDEX To NOEMPTYAVAILSINDEX Step 1
        grdSchd.Col = llCol
        'grdSchd.CellBackColor = LIGHTBLUE
    Next llCol
    'grdSchd.RowHeight(0) = fgBoxGridH + 15
    llRow = grdSchd.FixedRows
    
    For ilSHE = 0 To UBound(tmSHE) - 1 Step 1
        If gDateValue(tmSHE(ilSHE).sAirDate) <= lmEffEndDate Then
            llNoEvents = 0
            llNoFilledAvails = 0
            llNoSplitAvails = 0
            llNoEmptyAvails = 0
            smSEEStamp = ""
            'If rbcShowEvtInfo(1).Value Then
                ReDim tmSEE(0 To 0) As SEE
            'Else
            '    ilRet = gGetRecs_SEE_ScheduleEventsAPI(hmSEE, smSEEStamp, -1, tmSHE(ilSHE).lCode, "EngrSchdDef-Get Events", tmSEE())
            'End If
            'For llSEE = 0 To UBound(tmSEE) - 1 Step 1
            '    If tmSEE(llSEE).sAction <> "D" Then
            '        llNoEvents = llNoEvents + 1
            '        slCategory = ""
            '        ilETE = gBinarySearchETE(tmSEE(llSEE).iEteCode, tgCurrETE)
            '        If ilETE <> -1 Then
            '            slCategory = tgCurrETE(ilETE).sCategory
            '        End If
            '        If slCategory = "A" Then
            '            ilNoSpotsInAvail = 0
            '            llAvailLength = tmSEE(llSEE).lDuration
            '            For llAvailTest = 0 To UBound(tmSEE) - 1 Step 1
            '                If tmSEE(llAvailTest).sAction <> "D" Then
            '                    If (tmSEE(llAvailTest).lTime = tmSEE(llSEE).lTime) And (llSEE <> llAvailTest) Then
            '                        If tmSEE(llAvailTest).iBdeCode = tmSEE(llSEE).iBdeCode Then
            '                            ilETE = gBinarySearchETE(tmSEE(llAvailTest).iEteCode, tgCurrETE)
            '                            If ilETE <> -1 Then
            '                                If tgCurrETE(ilETE).sCategory = "S" Then
            '                                    ilNoSpotsInAvail = ilNoSpotsInAvail + 1
            '                                    llAvailLength = llAvailLength - tmSEE(llAvailTest).lDuration
            '                                End If
            '                            End If
            '                        End If
            '                    End If
            '                End If
            '            Next llAvailTest
            '            If llAvailLength > 0 Then
            '                llNoEmptyAvails = llNoEmptyAvails + 1
            '            Else
            '                llNoFilledAvails = llNoFilledAvails + 1
            '            End If
            '            If ilNoSpotsInAvail > 1 Then
            '                llNoSplitAvails = llNoSplitAvails + 1
            '            End If
            '        End If
            '    End If
            'Next llSEE
            
            'Set mItem = lbcSchd.ListItems.Add()
            'mItem.text = tmSHE(ilSHE).sAirDate
            'If tmSHE(ilSHE).sLoadedAutoStatus = "L" Then
            '    mItem.SubItems(EXPORTEDINDEX) = tmSHE(ilSHE).sLoadedAutoDate
            'Else
            '    mItem.SubItems(EXPORTEDINDEX) = ""
            'End If
            'llCount = gGetCount_AAE_As_Aired(tmSHE(ilSHE).lCode, "EngrSchd-mMoveRecToCtrls")
            'If llCount > 0 Then
            '    mItem.SubItems(IMPORTEDINDEX) = llCount
            'Else
            '    mItem.SubItems(IMPORTEDINDEX) = ""
            'End If
            'If rbcShowEvtInfo(0).Value Then
            '    mItem.SubItems(NOEVENTSINDEX) = Trim$(Str$(llNoEvents))
            '    mItem.SubItems(NOFILLEDAVAILSINDEX) = Trim$(Str$(llNoFilledAvails))
            '    mItem.SubItems(NOSPLITAVAILSINDEX) = Trim$(Str$(llNoSplitAvails))
            '    mItem.SubItems(NOEMPTYAVAILSINDEX) = Trim$(Str$(llNoEmptyAvails))
            '
            'End If
            'slItemCheckDate = tmSHE(ilSHE).sLastDateItemChk
            'If gDateValue(slItemCheckDate) <> gDateValue("12/31/2069") Then
            '    mItem.SubItems(ITEMCHECKINDEX) = slItemCheckDate
            'Else
            '    mItem.SubItems(ITEMCHECKINDEX) = ""
            'End If

            If llRow >= grdSchd.Rows Then
                grdSchd.AddItem ""
            End If
            'grdSchs.RowHeight(llRow) = fgBoxGridH + 15
            For llCol = DATEINDEX To NOEMPTYAVAILSINDEX Step 1
                grdSchd.Row = llRow
                grdSchd.Col = llCol
                grdSchd.CellBackColor = vbWhite
                If tmSHE(ilSHE).sConflictExist <> "Y" Then
                    grdSchd.CellForeColor = vbBlue
                Else
                    grdSchd.CellForeColor = vbRed
                End If
            Next llCol
            grdSchd.TextMatrix(llRow, DATEINDEX) = tmSHE(ilSHE).sAirDate
            If tmSHE(ilSHE).sConflictExist = "Y" Then
                grdSchd.TextMatrix(llRow, CONFLICTINDEX) = "Error"
            Else
                grdSchd.TextMatrix(llRow, CONFLICTINDEX) = ""
            End If
            If tmSHE(ilSHE).sSpotMergeStatus = "E" Then
                grdSchd.TextMatrix(llRow, SPOTMERGESTATUSINDEX) = "Error"
            ElseIf tmSHE(ilSHE).sSpotMergeStatus = "M" Then
                grdSchd.TextMatrix(llRow, SPOTMERGESTATUSINDEX) = "Merged"
            Else
                grdSchd.TextMatrix(llRow, SPOTMERGESTATUSINDEX) = ""
            End If
            If tmSHE(ilSHE).sLoadedAutoStatus = "L" Then
                If tmSHE(ilSHE).sLoadStatus = "E" Then
                    grdSchd.TextMatrix(llRow, LOADSTATUSINDEX) = tmSHE(ilSHE).sLoadedAutoDate & "-Error"
                Else
                    grdSchd.TextMatrix(llRow, LOADSTATUSINDEX) = tmSHE(ilSHE).sLoadedAutoDate
                End If
            Else
                If tmSHE(ilSHE).sLoadStatus = "E" Then
                    grdSchd.TextMatrix(llRow, LOADSTATUSINDEX) = "Error"
                Else
                    grdSchd.TextMatrix(llRow, LOADSTATUSINDEX) = ""
                End If
            End If
            llCount = gGetCount_AAE_As_Aired(tmSHE(ilSHE).lCode, "EngrSchd-mMoveRecToCtrls")
            If llCount > 0 Then
                grdSchd.TextMatrix(llRow, IMPORTEDINDEX) = llCount
            Else
                grdSchd.TextMatrix(llRow, IMPORTEDINDEX) = ""
            End If
            'If rbcShowEvtInfo(0).Value Then
            '    grdSchd.TextMatrix(llRow, NOEVENTSINDEX) = Trim$(Str$(llNoEvents))
            '    grdSchd.TextMatrix(llRow, NOFILLEDAVAILSINDEX) = Trim$(Str$(llNoFilledAvails))
            '    grdSchd.TextMatrix(llRow, NOSPLITAVAILSINDEX) = Trim$(Str$(llNoSplitAvails))
            '    grdSchd.TextMatrix(llRow, NOEMPTYAVAILSINDEX) = Trim$(Str$(llNoEmptyAvails))
            'End If
            grdSchd.Col = GETCOUNTSINDEX
            grdSchd.CellBackColor = GRAY
            grdSchd.TextMatrix(llRow, GETCOUNTSINDEX) = "Get #'s"
            grdSchd.TextMatrix(llRow, NOEVENTSINDEX) = ""
            grdSchd.TextMatrix(llRow, NOFILLEDAVAILSINDEX) = ""
            grdSchd.TextMatrix(llRow, NOSPLITAVAILSINDEX) = ""
            grdSchd.TextMatrix(llRow, NOEMPTYAVAILSINDEX) = ""
            slItemCheckDate = tmSHE(ilSHE).sLastDateItemChk
            If gDateValue(slItemCheckDate) <> gDateValue("12/31/2069") Then
                grdSchd.TextMatrix(llRow, ITEMCHECKINDEX) = slItemCheckDate
            Else
                grdSchd.TextMatrix(llRow, ITEMCHECKINDEX) = ""
            End If
            grdSchd.TextMatrix(llRow, SHECODEINDEX) = tmSHE(ilSHE).lCode
            llRow = llRow + 1
        End If
    Next ilSHE
    gGrid_IntegralHeight grdSchd
    gGrid_FillWithRows grdSchd
    grdSchd.Height = grdSchd.Height + 30
    grdSchd.Visible = True
    grdSchd.Redraw = True
    If lgSchTopRow <> -1 Then
        grdSchd.TopRow = lgSchTopRow
    End If
    
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim slStartDate As String
    
    
    ReDim tmSHE(0 To 0) As SHE
    'smEffStartDate = cccEffStartDate.text
    'If Not gIsDate(smEffStartDate) Then
    '    Beep
    '    cccEffStartDate.SetFocus    'edcEffStartDate.SetFocus
    '    Exit Sub
    'End If
    smEffStartDate = gGetEarlestSchdDate(True)
    lmEffStartDate = gDateValue(smEffStartDate)
    ''smEffEndDate = edcEffEndDate.text
    'smEffEndDate = cccEffEndDate.text
    'If smEffEndDate <> "" Then
    '    If Not gIsDate(smEffEndDate) Then
    '        Beep
    '        'edcEffEndDate.SetFocus
    '        cccEffEndDate.SetFocus
    '        Exit Sub
    '    End If
    '    lmEffEndDate = gDateValue(smEffEndDate)
    'Else
        lmEffEndDate = 99999999
    'End If
    ilRet = gGetTypeOfRecs_SHE_ScheduleHeaderByDate(smEffStartDate, "mPopulate SHE", tmSHE())
End Sub
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    igJobShowing(SCHEDULEJOB) = 0
    Unload EngrSchd
End Sub







Private Sub cmcRemove_Click()
    Dim ilRet As Integer
    Dim llSheCode As Long
    Dim ilSpotsExist As Integer
    Dim ilETE As Integer
    Dim ilSpotETECode As Integer
    Dim llSpotCount As Long
    Dim ilSpotExist As Integer
    
    If (lmHighlightRow < grdSchd.FixedRows) Or (lmHighlightRow >= grdSchd.Rows) Or (imRowSelectFlag <> 1) Then
        Exit Sub
    End If
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(SCHEDULEJOB) = 2) Then
        If (grdSchd.TextMatrix(lmHighlightRow + 1, DATEINDEX) = "") And (grdSchd.TextMatrix(lmHighlightRow, LOADSTATUSINDEX) = "") And (imSelectInFuture) Then
            'Determine if spots exist
            If grdSchd.TextMatrix(lmHighlightRow, SPOTMERGESTATUSINDEX) = "" Then
                llSheCode = grdSchd.TextMatrix(lmHighlightRow, SHECODEINDEX)
                ilSpotETECode = 0
                For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                    If tgCurrETE(ilETE).sCategory = "S" Then
                        ilSpotETECode = tgCurrETE(ilETE).iCode
                        Exit For
                    End If
                Next ilETE
                If ilSpotETECode > 0 Then
                    llSpotCount = gGetCount("SELECT count(seeETECode) FROM SEE_Schedule_Events WHERE seeSheCode = " & llSheCode & " AND seeEteCode = " & ilSpotETECode, "Schedule: Spot Count")
                    If llSpotCount > 0 Then
                        ilSpotExist = True
                    Else
                        ilSpotExist = False
                    End If
                Else
                    ilSpotExist = False
                End If
            Else
                ilSpotsExist = True
            End If
            If Not ilSpotsExist Then
                ilRet = MsgBox("This will remove the schedule for " & grdSchd.TextMatrix(lmHighlightRow, DATEINDEX) & ", Ok to Continue with removal", vbQuestion + vbYesNo, "Remove Schedule")
            Else
                ilRet = MsgBox("This will remove the schedule including spots for " & grdSchd.TextMatrix(lmHighlightRow, DATEINDEX) & ", Ok to Continue with removal", vbQuestion + vbYesNo, "Remove Schedule")
            End If
            If ilRet = vbYes Then
                gSetMousePointer grdSchd, grdSchd, vbHourglass
                ilRet = gExecGenSQLCall("DELETE FROM SEE_Schedule_Events Where seeSheCode = " & llSheCode)
                If ilRet Then
                    ilRet = gExecGenSQLCall("DELETE FROM SHE_Schedule_Header Where sheCode = " & llSheCode)
                    gSetMousePointer grdSchd, grdSchd, vbDefault
                    If ilRet Then
                        If Not ilSpotExist Then
                            MsgBox "Schedule successfully Removed", vbInformation + vbOKOnly, "Removed"
                        Else
                            MsgBox "Schedule including Spots successfully Removed", vbInformation + vbOKOnly, "Removed"
                        End If
                        grdSchd.RemoveItem lmHighlightRow
                        grdSchd.AddItem ""
                        lmHighlightRow = -1
                        imRowSelectFlag = -1
                        imSelectInFuture = False
                        mSetCommands
                    Else
                        MsgBox "Removing Event Headers was NOT successful", vbCritical + vbOKOnly, "Removal Error"
                    End If
                Else
                    gSetMousePointer grdSchd, grdSchd, vbDefault
                    MsgBox "Removing Events was NOT successful", vbCritical + vbOKOnly, "Removal Error"
                End If
            End If
        End If
    End If
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
    gSetFonts EngrSchd
    'gCenterFormModal EngrSchd
    gCenterForm EngrSchd
End Sub

Private Sub Form_Load()
    mGridColumns
    mInit
    igJobShowing(SCHEDULEJOB) = 1
End Sub

Private Sub Form_Resize()
    'These call are here and in form_Active (call to mGridColumns)
    'They are in mGridColumn in case the For_Initialize size chage does not cause a resize event
    mGridColumnWidth
    'lbcSchd.Height = cmcCancel.Top - lbcSchd.Top - 240    '8 * grdLib.RowHeight(0) + 30
    grdSchd.Height = cmcCancel.Top - grdSchd.Top - 240
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lgSchTopRow = grdSchd.TopRow
    btrDestroy hmSEE
    Erase tmSHE
    Erase tmSEE
    Set EngrSchd = Nothing
End Sub





Private Sub mInit()
    Dim llRet As Long
    Dim ilRet As Integer
    On Error GoTo ErrHand
    
    gSetMousePointer grdSchd, grdSchd, vbHourglass
    ReDim tgFilterValues(0 To 0) As FILTERVALUES
    ReDim tgFilterFields(0 To 0) As FIELDSELECTION
    ReDim tgSchdReplaceValues(0 To 0) As SCHDREPLACEVALUES
    ReDim tgReplaceFields(0 To 0) As FIELDSELECTION
    smNowDate = Format$(gNow(), "ddddd")
    lmNowDate = gDateValue(smNowDate)
    'edcEffStartDate.Text = smNowDate
    'tmcClick.Enabled = False
    imLastColSorted = -1
    imLastSort = -1
    lmEnableRow = -1
    imSelectInFuture = False
    imInChg = True
    hmSEE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSEE, "", sgDBPath & "SEE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    mPopETE
    mCreateUsedArrays
    mInitFilterInfo
    mInitReplaceInfo
    imInChg = False
    ReDim tmSHE(0 To 0) As SHE
    mSetCommands
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igJobStatus(SCHEDULEJOB) = 2) Then
    Else
        cmcChange.Caption = "&View"
        cmcNew.Enabled = False
    End If
    'llRet = SendMessageByNum(lbcSchd.hwnd, LV_SETEXTENDEDLISTVIEWSTYLE, 0, LV_FULLROWSSELECT + LV_GRIDLINES)
    tmcClick.Enabled = True
    gSetMousePointer grdSchd, grdSchd, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdSchd, grdSchd, vbDefault
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
'        For llRow = 1 To lbcSchd.ListCount - 1 Step 1
'            slStr = Trim$(lbcSchd.Text(llRow, NAMEINDEX))
'            If (slStr <> "") Then
'                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
'                    grdLib.Row = llRow
'                    Do While Not grdLib.RowIsVisible(grdLib.Row)
'                        grdLib.TopRow = grdLib.TopRow + 1
'                    Loop
'                    Exit Sub
'                End If
'            End If
'        Next llRow
'    End If
    
End Sub



Private Sub grdSchd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    imRowSelectFlag = 0
    imSelectInFuture = True
    If Y < grdSchd.RowHeight(0) Then
        If (grdSchd.MouseCol >= DATEINDEX) And (grdSchd.MouseCol <= NOEMPTYAVAILSINDEX) Then
            grdSchd.Col = grdSchd.MouseCol
            mSortCol grdSchd.Col
        End If
        grdSchd.Row = 0
        grdSchd.Col = SORTINDEX
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdSchd, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdSchd.FixedRows Then
        Exit Sub
    End If
    If grdSchd.TextMatrix(llCurrentRow, DATEINDEX) = "" Then
        lmHighlightRow = -1
    Else
        lmHighlightRow = llCurrentRow
        imRowSelectFlag = 1
        sgSchdDate = grdSchd.TextMatrix(llCurrentRow, DATEINDEX)
        If gDateValue(sgSchdDate) < lmNowDate Then
            imSelectInFuture = False
        End If
        If llCol = GETCOUNTSINDEX Then
            mGetSchCounts llCurrentRow
        End If
    End If
    mPaintRowColor
    mSetCommands
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    gSetMousePointer grdSchd, grdSchd, vbHourglass
    mPopulate
    mMoveRecToCtrls
    mSortCol DATEINDEX
    mSetCommands
    gSetMousePointer grdSchd, grdSchd, vbDefault
End Sub

Private Sub mPopETE()
    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_ETE_EventType("C", sgCurrETEStamp, "EngrLibETE-mPopETE Event Types", tgCurrETE())
End Sub

Private Sub mPaintRowColor()
    Dim llCol As Long
    Dim llRow As Long
    
    For llRow = grdSchd.FixedRows To grdSchd.Rows - 1 Step 1
        If grdSchd.TextMatrix(llRow, DATEINDEX) <> "" Then
            For llCol = DATEINDEX To NOEMPTYAVAILSINDEX Step 1
                grdSchd.Row = llRow
                grdSchd.Col = llCol
                If llCol <> GETCOUNTSINDEX Then
                    If lmHighlightRow <> llRow Then
                        grdSchd.CellBackColor = vbWhite
                    Else
                        grdSchd.CellBackColor = GRAY
                    End If
                End If
            Next llCol
        End If
    Next llRow
End Sub

Private Sub mGetSchCounts(llRow As Long)
    Dim llSEE As Long
    Dim ilRet As Integer
    Dim slStr As String
    Dim llNoEvents As Long
    Dim llNoFilledAvails As Long
    Dim llNoSplitAvails As Long
    Dim llNoEmptyAvails As Long
    Dim ilNoSpotsInAvail As Integer
    Dim slCategory As String
    Dim ilETE As Integer
    Dim llAvailTest As Long
    Dim llAvailLength As Long
    Dim llSheCode As Long
    
    gSetMousePointer grdSchd, grdSchd, vbHourglass
    llSheCode = grdSchd.TextMatrix(llRow, SHECODEINDEX)
    
    ilRet = gGetRecs_SEE_ScheduleEventsAPI(hmSEE, smSEEStamp, -1, llSheCode, "EngrSchdDef-Get Events", tmSEE())
    For llSEE = 0 To UBound(tmSEE) - 1 Step 1
        If tmSEE(llSEE).sAction <> "D" Then
            llNoEvents = llNoEvents + 1
            slCategory = ""
            ilETE = gBinarySearchETE(tmSEE(llSEE).iEteCode, tgCurrETE)
            If ilETE <> -1 Then
                slCategory = tgCurrETE(ilETE).sCategory
            End If
            If slCategory = "A" Then
                ilNoSpotsInAvail = 0
                llAvailLength = tmSEE(llSEE).lDuration
                For llAvailTest = 0 To UBound(tmSEE) - 1 Step 1
                    If tmSEE(llAvailTest).sAction <> "D" Then
                        If (tmSEE(llAvailTest).lTime = tmSEE(llSEE).lTime) And (llSEE <> llAvailTest) Then
                            If tmSEE(llAvailTest).iBdeCode = tmSEE(llSEE).iBdeCode Then
                                ilETE = gBinarySearchETE(tmSEE(llAvailTest).iEteCode, tgCurrETE)
                                If ilETE <> -1 Then
                                    If tgCurrETE(ilETE).sCategory = "S" Then
                                        ilNoSpotsInAvail = ilNoSpotsInAvail + 1
                                        llAvailLength = llAvailLength - tmSEE(llAvailTest).lDuration
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next llAvailTest
                'If llAvailLength > 0 Then
                If tmSEE(llSEE).lDuration > 0 Then
                    If ilNoSpotsInAvail = 0 Then
                        llNoEmptyAvails = llNoEmptyAvails + 1
                    ElseIf llAvailLength = 0 Then
                        llNoFilledAvails = llNoFilledAvails + 1
                    ElseIf llAvailLength > 0 Then
                        llNoSplitAvails = llNoSplitAvails + 1
                    End If
                End If
            End If
        End If
    Next llSEE
    grdSchd.TextMatrix(llRow, NOEVENTSINDEX) = Trim$(Str$(llNoEvents))
    grdSchd.TextMatrix(llRow, NOFILLEDAVAILSINDEX) = Trim$(Str$(llNoFilledAvails))
    grdSchd.TextMatrix(llRow, NOSPLITAVAILSINDEX) = Trim$(Str$(llNoSplitAvails))
    grdSchd.TextMatrix(llRow, NOEMPTYAVAILSINDEX) = Trim$(Str$(llNoEmptyAvails))
    gSetMousePointer grdSchd, grdSchd, vbDefault
End Sub

Private Sub mCreateUsedArrays()
    Dim ilRet As Integer
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
    ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrSchdDef-mPopBDE Bus Definition", tgCurrBDE())
    ReDim tgUsedBDE(0 To UBound(tgCurrBDE)) As BDE
    For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        LSet tgUsedBDE(ilBDE) = tgCurrBDE(ilBDE)
    Next ilBDE
    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrSchdDef-mPopANE Audio Audio Names", tgCurrANE())
    ReDim tgUsedANE(0 To UBound(tgCurrANE)) As ANE
    For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        LSet tgUsedANE(ilANE) = tgCurrANE(ilANE)
    Next ilANE
    ilRet = gGetTypeOfRecs_ATE_AudioType("C", sgCurrATEStamp, "EngrAudioType-mPopulate", tgCurrATE())
    ReDim tgUsedATE(0 To UBound(tgCurrATE)) As ATE
    For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
        LSet tgUsedATE(ilATE) = tgCurrATE(ilATE)
    Next ilATE
    ilRet = gGetTypeOfRecs_ETE_EventType("C", sgCurrETEStamp, "EngrLibETE-mPopETE Event Types", tgCurrETE())
    ilRet = gGetTypeOfRecs_EPE_EventProperties("C", sgCurrEPEStamp, "EngrSchdDef-mPopETE Event Properties", tgCurrEPE())
    ReDim tgUsedETE(0 To UBound(tgCurrETE)) As ETE
    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
        LSet tgUsedETE(ilETE) = tgCurrETE(ilETE)
    Next ilETE
    ilRet = gGetTypeOfRecs_FNE_FollowName("C", sgCurrFNEStamp, "EngrSchdDef-mPopFNE Follow", tgCurrFNE())
    ReDim tgUsedFNE(0 To UBound(tgCurrFNE)) As FNE
    For ilFNE = 0 To UBound(tgCurrFNE) - 1 Step 1
        LSet tgUsedFNE(ilFNE) = tgCurrFNE(ilFNE)
    Next ilFNE
    ilRet = gGetTypeOfRecs_MTE_MaterialType("C", sgCurrMTEStamp, "EngrSchdDef-mPopMTE Material Type", tgCurrMTE())
    ReDim tgUsedMTE(0 To UBound(tgCurrMTE)) As MTE
    For ilMTE = 0 To UBound(tgCurrMTE) - 1 Step 1
        LSet tgUsedMTE(ilMTE) = tgCurrMTE(ilMTE)
    Next ilMTE
    ilRet = gGetTypeOfRecs_NNE_NetcueName("C", sgCurrNNEStamp, "EngrSchdDef-mPopNNE Netcue", tgCurrNNE())
    ReDim tgUsedNNE(0 To UBound(tgCurrNNE)) As NNE
    For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
        LSet tgUsedNNE(ilNNE) = tgCurrNNE(ilNNE)
    Next ilNNE
    ilRet = gGetTypeOfRecs_RNE_RelayName("C", sgCurrRNEStamp, "EngrSchdDef-mPopRNE Relay", tgCurrRNE())
    ReDim tgUsedRNE(0 To UBound(tgCurrRNE)) As RNE
    For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
        LSet tgUsedRNE(ilRNE) = tgCurrRNE(ilRNE)
    Next ilRNE
    ilRet = gGetTypeOfRecs_TTE_TimeType("C", "S", sgCurrStartTTEStamp, "EngrSchdDef-mPopTTE_StartType Start Type", tgCurrStartTTE())
    ReDim tgUsedStartTTE(0 To UBound(tgCurrStartTTE)) As TTE
    For ilTTE = 0 To UBound(tgCurrStartTTE) - 1 Step 1
        LSet tgUsedStartTTE(ilTTE) = tgCurrStartTTE(ilTTE)
    Next ilTTE
    ilRet = gGetTypeOfRecs_TTE_TimeType("C", "E", sgCurrEndTTEStamp, "EngrSchdDef-mPopTTE_EndType End Type", tgCurrEndTTE())
    ReDim tgUsedEndTTE(0 To UBound(tgCurrEndTTE)) As TTE
    For ilTTE = 0 To UBound(tgCurrEndTTE) - 1 Step 1
        LSet tgUsedEndTTE(ilTTE) = tgCurrEndTTE(ilTTE)
    Next ilTTE
    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "A", sgCurrAudioCCEStamp, "EngrSchdDef-mPopCCE_Audio Control Character", tgCurrAudioCCE())
    ReDim tgUsedAudioCCE(0 To UBound(tgCurrAudioCCE)) As CCE
    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
        LSet tgUsedAudioCCE(ilCCE) = tgCurrAudioCCE(ilCCE)
    Next ilCCE
    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "B", sgCurrBusCCEStamp, "EngrSchdDef-mPopCCE_Bus Control Character", tgCurrBusCCE())
    ReDim tgUsedBusCCE(0 To UBound(tgCurrBusCCE)) As CCE
    For ilCCE = 0 To UBound(tgCurrBusCCE) - 1 Step 1
        LSet tgUsedBusCCE(ilCCE) = tgCurrBusCCE(ilCCE)
    Next ilCCE
    ilRet = gGetTypeOfRecs_SCE_SilenceChar("C", sgCurrSCEStamp, "EngrSchdDef-mPopSCE Silence Character", tgCurrSCE())
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
End Sub

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
    'If tmSHE.lCode <> 0 Then
        tgFilterFields(ilUpper).sFieldName = "Event ID"
        tgFilterFields(ilUpper).iFieldType = 1
        tgFilterFields(ilUpper).iMaxNoChar = 0
        tgFilterFields(ilUpper).sListFile = ""
        tgFilterFields(ilUpper).sMandatory = "Y"
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    'End If
    If (sgClientFields = "A") Then
        If (tgSchUsedSumEPE.sABCFormat <> "N") Then
            tgFilterFields(ilUpper).sFieldName = "ABC Format"
            tgFilterFields(ilUpper).iFieldType = 2
            tgFilterFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCFormat")
            tgFilterFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCFormat = "Y") Then
                tgFilterFields(ilUpper).sMandatory = "Y"
            Else
                tgFilterFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
        If (tgSchUsedSumEPE.sABCPgmCode <> "N") Then
            tgFilterFields(ilUpper).sFieldName = "ABC Pgm Code"
            tgFilterFields(ilUpper).iFieldType = 2
            tgFilterFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCPgmCode")
            tgFilterFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCPgmCode = "Y") Then
                tgFilterFields(ilUpper).sMandatory = "Y"
            Else
                tgFilterFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
        If (tgSchUsedSumEPE.sABCXDSMode <> "N") Then
            tgFilterFields(ilUpper).sFieldName = "ABC XDS Mode"
            tgFilterFields(ilUpper).iFieldType = 2
            tgFilterFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCXDSMODE")
            tgFilterFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCXDSMode = "Y") Then
                tgFilterFields(ilUpper).sMandatory = "Y"
            Else
                tgFilterFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
        If (tgSchUsedSumEPE.sABCRecordItem <> "N") Then
            tgFilterFields(ilUpper).sFieldName = "ABC Recd Item"
            tgFilterFields(ilUpper).iFieldType = 2
            tgFilterFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCRecordItem")
            tgFilterFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCRecordItem = "Y") Then
                tgFilterFields(ilUpper).sMandatory = "Y"
            Else
                tgFilterFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
    End If
    
End Sub

Private Sub mInitReplaceInfo()
    Dim ilUpper As Integer
    ReDim tgReplaceFields(0 To 0) As FIELDSELECTION
    
    ilUpper = 0
    If ((tgSchUsedSumEPE.sAudioName <> "N") Or (tgSchUsedSumEPE.sProtAudioName <> "N") Or (tgSchUsedSumEPE.sBkupAudioName <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Audio Name"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("AudioName", 6)
        tgReplaceFields(ilUpper).sListFile = "ANE"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sAudioName
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sBus <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Bus"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("BusName", 6)
        tgReplaceFields(ilUpper).sListFile = "BDE"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sBus
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sFollow <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Follow"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Follow", 6)
        tgReplaceFields(ilUpper).sListFile = "FNE"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sFollow
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sMaterialType <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Material"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Material", 6)
        tgReplaceFields(ilUpper).sListFile = "MTE"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sMaterialType
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgSchUsedSumEPE.sStartNetcue <> "N") Or (tgSchUsedSumEPE.sStopNetcue <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Netcue"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Netcue1", 6)
        tgReplaceFields(ilUpper).sListFile = "NNE"
        If (tgSchManSumEPE.sStartNetcue = "Y") Or (tgSchManSumEPE.sStopNetcue = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgSchUsedSumEPE.sRelay1 <> "N") Or (tgSchUsedSumEPE.sRelay2 <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Relay"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Relay1", 6)
        tgReplaceFields(ilUpper).sListFile = "RNE"
        If (tgSchManSumEPE.sRelay1 = "Y") Or (tgSchManSumEPE.sRelay2 = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sStartType <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Start Type"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("StartType", 6)
        tgReplaceFields(ilUpper).sListFile = "TTES"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sStartType
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sEndType <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "End Type"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("EndType", 6)
        tgReplaceFields(ilUpper).sListFile = "TTEE"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sEndType
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgSchUsedSumEPE.sAudioControl <> "N") Or (tgSchUsedSumEPE.sProtAudioControl <> "N") Or (tgSchUsedSumEPE.sBkupAudioControl <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Audio Control"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("AudioCtrl", 6)
        tgReplaceFields(ilUpper).sListFile = "CCEA"
        If (tgSchManSumEPE.sAudioControl = "Y") Or (tgSchManSumEPE.sProtAudioControl = "Y") Or (tgSchManSumEPE.sBkupAudioControl = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sBusControl <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Bus Control"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("BusCtrl", 6)
        tgReplaceFields(ilUpper).sListFile = "CCEB"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sBusControl
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sTitle2 <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Title 2"
        tgReplaceFields(ilUpper).iFieldType = 9
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Title2", 6)
        tgReplaceFields(ilUpper).sListFile = "CTE2"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sTitle2
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sTitle1 <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Title 1"
        tgReplaceFields(ilUpper).iFieldType = 9
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Title1", 6)
        tgReplaceFields(ilUpper).sListFile = "CTE1"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sTitle1
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgSchUsedSumEPE.sSilence1 <> "N") Or (tgSchUsedSumEPE.sSilence2 <> "N") Or (tgSchUsedSumEPE.sSilence3 <> "N") Or (tgSchUsedSumEPE.sSilence4 <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Silence Control"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Silence1", 6)
        tgReplaceFields(ilUpper).sListFile = "SCE"
        If (tgSchManSumEPE.sSilence1 = "Y") Or (tgSchManSumEPE.sSilence2 = "Y") Or (tgSchManSumEPE.sSilence3 = "Y") Or (tgSchManSumEPE.sSilence4 = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If tgSchUsedSumEPE.sFixedTime <> "N" Then
        tgReplaceFields(ilUpper).sFieldName = "Fixed Time"
        tgReplaceFields(ilUpper).iFieldType = 9
        tgReplaceFields(ilUpper).iMaxNoChar = 1
        tgReplaceFields(ilUpper).sListFile = "FTYN"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sFixedTime
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgSchUsedSumEPE.sAudioItemID <> "N") Or (tgSchUsedSumEPE.sProtAudioItemID <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Item ID"
        tgReplaceFields(ilUpper).iFieldType = 2
        tgReplaceFields(ilUpper).iMaxNoChar = 0
        tgReplaceFields(ilUpper).sListFile = ""
        If (tgSchManSumEPE.sAudioItemID = "Y") Or (tgSchManSumEPE.sProtAudioItemID = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgSchUsedSumEPE.sAudioISCI <> "N") Or (tgSchUsedSumEPE.sProtAudioISCI <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "ISCI"
        tgReplaceFields(ilUpper).iFieldType = 2
        tgReplaceFields(ilUpper).iMaxNoChar = 0
        tgReplaceFields(ilUpper).sListFile = ""
        If (tgSchManSumEPE.sAudioISCI = "Y") Or (tgSchManSumEPE.sProtAudioISCI = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (sgClientFields = "A") Then
        If (tgSchUsedSumEPE.sABCFormat <> "N") Then
            tgReplaceFields(ilUpper).sFieldName = "ABC Format"
            tgReplaceFields(ilUpper).iFieldType = 2
            tgReplaceFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCFormat")
            tgReplaceFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCFormat = "Y") Then
                tgReplaceFields(ilUpper).sMandatory = "Y"
            Else
                tgReplaceFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
        If (tgSchUsedSumEPE.sABCPgmCode <> "N") Then
            tgReplaceFields(ilUpper).sFieldName = "ABC Pgm Code"
            tgReplaceFields(ilUpper).iFieldType = 2
            tgReplaceFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCPgmCode")
            tgReplaceFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCPgmCode = "Y") Then
                tgReplaceFields(ilUpper).sMandatory = "Y"
            Else
                tgReplaceFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
        If (tgSchUsedSumEPE.sABCXDSMode <> "N") Then
            tgReplaceFields(ilUpper).sFieldName = "ABC XDS Mode"
            tgReplaceFields(ilUpper).iFieldType = 2
            tgReplaceFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCXDSMODE")
            tgReplaceFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCXDSMode = "Y") Then
                tgReplaceFields(ilUpper).sMandatory = "Y"
            Else
                tgReplaceFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
        If (tgSchUsedSumEPE.sABCRecordItem <> "N") Then
            tgReplaceFields(ilUpper).sFieldName = "ABC Recd Item"
            tgReplaceFields(ilUpper).iFieldType = 2
            tgReplaceFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCRecordItem")
            tgReplaceFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCRecordItem = "Y") Then
                tgReplaceFields(ilUpper).sMandatory = "Y"
            Else
                tgReplaceFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
    End If
    
End Sub

