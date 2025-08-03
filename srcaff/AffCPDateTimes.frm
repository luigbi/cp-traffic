VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDateTimes 
   Caption         =   "Times by Date"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   Icon            =   "AffCPDateTimes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   9105
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   -45
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4980
      Width           =   45
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
      Left            =   4980
      Picture         =   "AffCPDateTimes.frx":08CA
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1650
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   4035
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ListBox lbcAdvt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffCPDateTimes.frx":09C4
      Left            =   5865
      List            =   "AffCPDateTimes.frx":09C6
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2610
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox pbcPostSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   60
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   3
      Top             =   435
      Width           =   60
   End
   Begin VB.PictureBox pbcPostTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   8
      Top             =   4335
      Width           =   60
   End
   Begin VB.PictureBox pbcPostFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   30
      Width           =   60
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   60
      Picture         =   "AffCPDateTimes.frx":09C8
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   675
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.ListBox lbcStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffCPDateTimes.frx":0CD2
      Left            =   4290
      List            =   "AffCPDateTimes.frx":0CD4
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox txtKey 
      BackColor       =   &H80000018&
      Height          =   1785
      Left            =   5955
      MultiLine       =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "AffCPDateTimes.frx":0CD6
      Top             =   225
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.CommandButton cmdMG 
      Caption         =   "Add MG"
      Height          =   375
      Left            =   5295
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdBonus 
      Caption         =   "Add Bonus"
      Height          =   375
      Left            =   6855
      TabIndex        =   19
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame frcPosting 
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
      Height          =   555
      Left            =   210
      TabIndex        =   9
      Top             =   4755
      Width           =   3915
      Begin VB.OptionButton optStatus 
         Caption         =   "Outstanding"
         Height          =   225
         Index           =   2
         Left            =   690
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   300
         Width           =   1875
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Completed"
         Height          =   255
         Index           =   0
         Left            =   690
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   30
         Width           =   1260
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Partially Completed"
         Height          =   255
         Index           =   1
         Left            =   1965
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   30
         Width           =   1890
      End
      Begin VB.Label Label2 
         Caption         =   "Posting"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   30
         Width           =   585
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3210
      Top             =   4770
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5550
      FormDesignWidth =   9105
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4500
      TabIndex        =   14
      Top             =   4575
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPost 
      Height          =   4005
      Left            =   165
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   390
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   7064
      _Version        =   393216
      Cols            =   14
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
      _Band(0).Cols   =   14
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7620
      TabIndex        =   16
      Top             =   4575
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6060
      TabIndex        =   15
      Top             =   4575
      Width           =   1335
   End
   Begin VB.Label lacGame 
      Alignment       =   2  'Center
      Height          =   195
      Left            =   675
      TabIndex        =   21
      Top             =   120
      Width           =   7530
   End
   Begin VB.Image imcKey 
      Height          =   225
      Left            =   8235
      Picture         =   "AffCPDateTimes.frx":0CDA
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmDateTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmDateTimes - displays spot times sorted by date information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text
Private lmAttCode As Long
Private imVefCode As Integer
Private imShttCode As Integer
Private lmSdfCode As Long
Private lmAstCode As Long
Private lmLstCode As Long
Private smPostSDate As String
Private smPostEDate As String
Private imFieldChgd As Integer
Private imFirstTime As Integer
Private imIntegralSet As Integer
Private imFirstDrop As Integer
Private imMouseDown As Integer
Private imLastRow As Integer
Private imIgnoreRowChg As Integer
Private imBSMode As Integer
Private hmAst As Integer
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private tmGameAstInfo() As ASTINFO
Private smMGFeedDate As String
Private smMGFeedTime As String
Private bmMGCreated As Boolean

'Grid Controls
Private imShowGridBox As Integer
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on
Private imArrowKeyPressed As Integer    'Indicates arrow key pressed

Private imLastPostColSorted As Integer
Private imLastPostSort As Integer

Private lst_rst As ADODB.Recordset
Private dat_rst As ADODB.Recordset
Private ast_rst As ADODB.Recordset

Const VEHICLEINDEX = 0
Private imAdvtCol As Integer
Private imDateCol As Integer
Private imTimeCol As Integer
Const PLEDGEDAYSINDEX = 4
Const PLEDGETIMEINDEX = 5
Const AIRDATEINDEX = 7
Const AIRTIMEINDEX = 8
Const LENINDEX = 9
Const STATUSINDEX = 6
Const CARTISCIINDEX = 10
Const INFOINDEX = 11
Const ASTINDEX = 12
Const SORTINDEX = 13




Private Function mTestGridValues()
    Dim iLoop As Integer
    Dim iPack As Integer
    Dim sDate As String
    Dim sTime As String
    Dim iDay As Integer
    Dim iIndex As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim ilError As Integer
    Dim llRowIndex As Long
    
    grdPost.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdPost.FixedRows To grdPost.Rows - 1 Step 1
        slStr = Trim$(grdPost.TextMatrix(llRow, STATUSINDEX))
        If slStr <> "" Then
            slStr = grdPost.TextMatrix(llRow, STATUSINDEX)
            llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
            If llRowIndex >= 0 Then
                iIndex = lbcStatus.ItemData(llRowIndex)
                If tgStatusTypes(iIndex).iPledged <> 2 Then
                    sDate = grdPost.TextMatrix(llRow, AIRDATEINDEX)
                    If (gIsDate(sDate) = False) Or (Len(Trim$(sDate)) = 0) Then   'Date not valid.
                        ilError = True
                        If Len(Trim$(sTime)) = 0 Then
                            grdPost.TextMatrix(llRow, AIRDATEINDEX) = "Missing"
                        End If
                        grdPost.Row = llRow
                        grdPost.Col = AIRDATEINDEX
                        grdPost.CellForeColor = vbRed
                    End If
                    sTime = grdPost.TextMatrix(llRow, AIRTIMEINDEX)
                    If (gIsTime(sTime) = False) Or (Len(Trim$(sTime)) = 0) Then    'Time not valid.
                        ilError = True
                        If Len(Trim$(sTime)) = 0 Then
                            grdPost.TextMatrix(llRow, AIRTIMEINDEX) = "Missing"
                        End If
                        grdPost.Row = llRow
                        grdPost.Col = AIRTIMEINDEX
                        grdPost.CellForeColor = vbRed
                    End If
                End If
            End If
        End If
    Next llRow
    If ilError Then
        grdPost.Redraw = True
        mTestGridValues = False
        gSetMousePointer grdPost, grdPost, vbDefault
        Exit Function
    Else
        mTestGridValues = True
        Exit Function
    End If
End Function



Private Function mPostColAllowed(llCol As Long) As Integer
    Dim ilIndex As Integer
    Dim slStr As String
    Dim llRow As Long
    Dim ilLoop As Integer
    
    mPostColAllowed = False
    
    mPopStatus grdPost.Row
    
    '3/15/16:Disallow changing a missed with aMG.  Require that MG removed
    ilIndex = Val(grdPost.TextMatrix(grdPost.Row, ASTINDEX))
    If ilIndex >= 0 Then
        If tmAstInfo(ilIndex).lLkAstCode > 0 Then
            If tmAstInfo(ilIndex).iStatus Mod 100 <= 10 Then
                Exit Function
            End If
        End If
    End If
    slStr = grdPost.TextMatrix(grdPost.Row, STATUSINDEX)
    llRow = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
    If llRow >= 0 Then
        ilIndex = lbcStatus.ItemData(llRow)
        If (tgStatusTypes(ilIndex).iPledged = 0) Then   'Live
            mPostColAllowed = True
        ElseIf (tgStatusTypes(ilIndex).iPledged = 1) Then   'Delayed
            mPostColAllowed = True
        ElseIf (tgStatusTypes(ilIndex).iPledged = 2) Then   'Not Aired
            mPostColAllowed = True
        ElseIf (tgStatusTypes(ilIndex).iPledged = 3) Then   'No Pledged Times
            mPostColAllowed = True
        End If
    Else
        If (llCol = AIRDATEINDEX) Or (llCol = AIRTIMEINDEX) Then
            For ilLoop = 0 To UBound(tgStatusTypes) Step 1
                If StrComp(slStr, Trim$(tgStatusTypes(ilLoop).sName), 1) = 0 Then
                    If (tgStatusTypes(ilLoop).iStatus = 6) Or (tgStatusTypes(ilLoop).iStatus = 7) Then
                        mPostColAllowed = True
                    End If
                    Exit For
                End If
            Next ilLoop
        End If
        For ilLoop = 0 To UBound(tgStatusTypes) Step 1
            If StrComp(slStr, Trim$(tgStatusTypes(ilLoop).sName), 1) = 0 Then
                If gIsAstStatus(tmAstInfo(ilIndex).iStatus, ASTEXTENDED_MG) Or gIsAstStatus(tmAstInfo(ilIndex).iStatus, ASTEXTENDED_REPLACEMENT) Or gIsAstStatus(tmAstInfo(ilIndex).iStatus, ASTEXTENDED_BONUS) Then
                    mPostColAllowed = True
                End If
                Exit For
            End If
        Next ilLoop
    End If

End Function

Private Sub mPostSetShow()
    Dim slStr As String
    Dim ilIndex As Integer
    Dim iStatus As Integer
    Dim sStatus As String
    Dim iIndex As Integer
    Dim ilRowIndex As Integer
    Dim llAirDate As Long
    Dim slDate As String
    Dim llOldAirDate As Long
    Dim llFeedDate As Long
    Dim slTime As String
    Dim ilPos As Integer
    Dim llAirTime As Long
    Dim llOldAirTime As Long
    Dim llPledgeStartTime As Long
    Dim llPledgeEndTime As Long
    
    
    If (lmEnableRow >= grdPost.FixedRows) And (lmEnableRow < grdPost.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case AIRDATEINDEX
                slStr = txtDropdown.Text
                If (gIsDate(slStr)) And (slStr <> "") Then
                    If grdPost.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                        If optStatus(2).Value Then
                            optStatus(1).Value = True
                        End If
                        llAirDate = DateValue(gAdjYear(slStr))
                        slDate = grdPost.TextMatrix(lmEnableRow, lmEnableCol)
                        If slDate = "" Then
                            llOldAirDate = 0
                        Else
                            llOldAirDate = DateValue(gAdjYear(slDate))
                        End If
                        imFieldChgd = True
                        grdPost.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                        slDate = grdPost.TextMatrix(lmEnableRow, imDateCol)
                        llFeedDate = DateValue(gAdjYear(slDate))
                        iStatus = -1
                        sStatus = Trim$(grdPost.TextMatrix(lmEnableRow, STATUSINDEX))
                        If (StrComp(sStatus, "MG", vbTextCompare) <> 0) And (StrComp(sStatus, "Bonus", vbTextCompare) <> 0) Then
                            For iIndex = 0 To UBound(tgStatusTypes) Step 1
                                If StrComp(sStatus, Trim$(tgStatusTypes(iIndex).sName), 1) = 0 Then
                                    iStatus = tgStatusTypes(iIndex).iStatus
                                    ilRowIndex = iIndex
                                    Exit For
                                End If
                            Next iIndex
                            If (llAirDate <> llFeedDate) Then
                                If (iStatus <> 6) And (iStatus <> 7) Then
                                    For iIndex = 0 To UBound(tgStatusTypes) Step 1
                                        If tgStatusTypes(iIndex).iStatus = 6 Then
                                                grdPost.TextMatrix(lmEnableRow, STATUSINDEX) = Trim$(tgStatusTypes(iIndex).sName)
                                            Exit For
                                        End If
                                    Next iIndex
                                End If
                            Else
                                If llOldAirDate <> llFeedDate Then
'                                    If (iStatus <> 0) And (iStatus <> 1) Then
                                    If (iStatus <> 0) And (iStatus <> 1) And (iStatus <> 9) And (iStatus <> 10) Then
                                        For iIndex = 0 To UBound(tgStatusTypes) Step 1
                                            If tgStatusTypes(iIndex).iStatus = 1 Then
                                                grdPost.TextMatrix(lmEnableRow, STATUSINDEX) = Trim$(tgStatusTypes(iIndex).sName)
                                                Exit For
                                            End If
                                        Next iIndex
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Case AIRTIMEINDEX
                slStr = txtDropdown.Text
                If (gIsTime(slStr)) And (slStr <> "") Then
                    slStr = gConvertTime(slStr)
                    If Second(slStr) = 0 Then
                        slStr = Format$(slStr, sgShowTimeWOSecForm)
                    Else
                        slStr = Format$(slStr, sgShowTimeWSecForm)
                    End If
                    If grdPost.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                        If optStatus(2).Value Then
                            optStatus(1).Value = True
                        End If
                        llAirTime = gTimeToLong(slStr, False)
                        slTime = grdPost.TextMatrix(lmEnableRow, lmEnableCol)
                        llOldAirTime = gTimeToLong(slTime, False)
                        imFieldChgd = True
                        grdPost.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                        slStr = grdPost.TextMatrix(lmEnableRow, PLEDGETIMEINDEX)
                        ilPos = InStr(1, slStr, "-", vbTextCompare)
                        If ilPos > 0 Then
                            slTime = Trim$(Left$(slStr, ilPos - 1))
                            llPledgeStartTime = gTimeToLong(slTime, False)
                            slTime = Trim$(Mid$(slStr, ilPos + 1))
                            llPledgeEndTime = gTimeToLong(slTime, TIME)
                        Else
                            slTime = Trim$(slStr)
                            llPledgeStartTime = gTimeToLong(slTime, False)
                            llPledgeEndTime = llPledgeStartTime
                        End If
                        'If changed to outside of pledge time, then set status to "7"
                        'If inside of pledge, is if old time was outside of pledge, if so set to "2", otherwise set leave
                        
                        iStatus = -1
                        sStatus = Trim$(grdPost.TextMatrix(lmEnableRow, STATUSINDEX))
                        If (StrComp(sStatus, "MG", vbTextCompare) <> 0) And (StrComp(sStatus, "Bonus", vbTextCompare) <> 0) Then
                            For iIndex = 0 To UBound(tgStatusTypes) Step 1
                                If StrComp(sStatus, Trim$(tgStatusTypes(iIndex).sName), 1) = 0 Then
                                    iStatus = tgStatusTypes(iIndex).iStatus
                                    ilRowIndex = iIndex
                                    Exit For
                                End If
                            Next iIndex
                            If (llAirTime < llPledgeStartTime) Or (llAirTime > llPledgeEndTime) Then
                                If (iStatus <> 6) And (iStatus <> 7) Then
                                    For iIndex = 0 To UBound(tgStatusTypes) Step 1
                                        If tgStatusTypes(iIndex).iStatus = 6 Then
                                                grdPost.TextMatrix(lmEnableRow, STATUSINDEX) = Trim$(tgStatusTypes(iIndex).sName)
                                            Exit For
                                        End If
                                    Next iIndex
                                End If
                            Else
                                If (llOldAirTime < llPledgeStartTime) Or (llOldAirTime > llPledgeEndTime) Then
'                                    If (iStatus <> 0) And (iStatus <> 1) Then
                                    If (iStatus <> 0) And (iStatus <> 1) And (iStatus <> 9) And (iStatus <> 10) Then
                                        If llPledgeEndTime - llPledgeStartTime <= 300 Then
                                            For iIndex = 0 To UBound(tgStatusTypes) Step 1
                                                If tgStatusTypes(iIndex).iStatus = 0 Then
                                                        grdPost.TextMatrix(lmEnableRow, STATUSINDEX) = Trim$(tgStatusTypes(iIndex).sName)
                                                    Exit For
                                                End If
                                            Next iIndex
                                        Else
                                            For iIndex = 0 To UBound(tgStatusTypes) Step 1
                                                If tgStatusTypes(iIndex).iStatus = 1 Then
                                                        grdPost.TextMatrix(lmEnableRow, STATUSINDEX) = Trim$(tgStatusTypes(iIndex).sName)
                                                    Exit For
                                                End If
                                            Next iIndex
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Case STATUSINDEX
                slStr = txtDropdown.Text
                If grdPost.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                    If optStatus(2).Value Then
                        optStatus(1).Value = True
                    End If
                    imFieldChgd = True
                    grdPost.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    iStatus = -1
                    sStatus = Trim$(grdPost.TextMatrix(lmEnableRow, STATUSINDEX))
                    For iIndex = 0 To UBound(tgStatusTypes) Step 1
                        If StrComp(sStatus, Trim$(tgStatusTypes(iIndex).sName), 1) = 0 Then
                            iStatus = tgStatusTypes(iIndex).iStatus
                            Exit For
                        End If
                    Next iIndex
                    If iStatus <> -1 Then
                        If tgStatusTypes(iStatus).iPledged = 2 Then
                            grdPost.TextMatrix(lmEnableRow, AIRDATEINDEX) = ""
                            grdPost.TextMatrix(lmEnableRow, AIRTIMEINDEX) = ""
                        Else
                            If Trim$(grdPost.TextMatrix(lmEnableRow, AIRDATEINDEX)) = "" Then
                                iIndex = grdPost.TextMatrix(lmEnableRow, ASTINDEX)
                                grdPost.TextMatrix(lmEnableRow, AIRDATEINDEX) = Format$(tmAstInfo(iIndex).sAirDate, sgShowDateForm)
                            End If
                            If Trim$(grdPost.TextMatrix(lmEnableRow, AIRTIMEINDEX)) = "" Then
                                iIndex = grdPost.TextMatrix(lmEnableRow, ASTINDEX)
                                If Second(tmAstInfo(iIndex).sAirTime) <> 0 Then
                                    grdPost.TextMatrix(lmEnableRow, AIRTIMEINDEX) = Format$(tmAstInfo(iIndex).sAirTime, sgShowTimeWSecForm)
                                Else
                                    grdPost.TextMatrix(lmEnableRow, AIRTIMEINDEX) = Format$(tmAstInfo(iIndex).sAirTime, sgShowTimeWOSecForm)
                                End If
                            End If
                        End If
                    End If
                End If
        End Select
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imShowGridBox = False
    pbcArrow.Visible = False
    txtDropdown.Visible = False
    lbcStatus.Visible = False
    cmcDropDown.Visible = False
    If imFieldChgd Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
End Sub

Private Sub mPostEnableBox()
    Dim ilIndex As Integer
    Dim slStr As String
    Dim iLoop As Integer
    Dim iIndex As Integer
    
    If (grdPost.Row >= grdPost.FixedRows) And (grdPost.Row < grdPost.Rows) And (grdPost.Col >= 0) And (grdPost.Col < grdPost.Cols - 1) Then
        lmEnableRow = grdPost.Row
        lmEnableCol = grdPost.Col
        imShowGridBox = True
        pbcArrow.Move grdPost.Left - pbcArrow.Width, grdPost.Top + grdPost.RowPos(grdPost.Row) + (grdPost.RowHeight(grdPost.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        Select Case grdPost.Col
            Case AIRDATEINDEX  'Date
                txtDropdown.Move grdPost.Left + grdPost.ColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - 30, grdPost.RowHeight(grdPost.Row) - 15
                If grdPost.Text <> "Missing" Then
                    txtDropdown.Text = grdPost.Text
                Else
                    txtDropdown.Text = ""
                End If
                If txtDropdown.Height > grdPost.RowHeight(grdPost.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdPost.RowHeight(grdPost.Row) - 15
                End If
                txtDropdown.Visible = True
                txtDropdown.SetFocus
            Case AIRTIMEINDEX  'Time
                txtDropdown.Move grdPost.Left + grdPost.ColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - 30, grdPost.RowHeight(grdPost.Row) - 15
                If grdPost.Text <> "Missing" Then
                    txtDropdown.Text = grdPost.Text
                Else
                    txtDropdown.Text = ""
                End If
                If txtDropdown.Height > grdPost.RowHeight(grdPost.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdPost.RowHeight(grdPost.Row) - 15
                End If
                txtDropdown.Visible = True
                txtDropdown.SetFocus
            Case STATUSINDEX
                'txtDropdown.Move grdPost.Left + imColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - 30, grdPost.RowHeight(grdPost.Row) - 15
                txtDropdown.Move grdPost.Left + grdPost.ColPos(STATUSINDEX) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(STATUSINDEX) - cmcDropDown.Width - 30, grdPost.RowHeight(grdPost.Row) - 15
                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
                lbcStatus.Move txtDropdown.Left, txtDropdown.Top + txtDropdown.Height, txtDropdown.Width + (7 * txtDropdown.Width) \ 2
                'iIndex = grdPost.TextMatrix(grdPost.Row, ASTINDEX)
                'lbcStatus.Clear
                ''If tmAstInfo(iIndex).iStatus = 20 Then
                'If gIsAstStatus(tmAstInfo(iIndex).iStatus, ASTEXTENDED_MG) Then
                '    For iLoop = 0 To UBound(tgStatusTypes) Step 1
                '        If (tgStatusTypes(iLoop).iPledged = 2) Or (Trim$(tgStatusTypes(iLoop).sName) = "MG") Then
                '            lbcStatus.AddItem Trim$(tgStatusTypes(iLoop).sName)
                '            lbcStatus.ItemData(lbcStatus.NewIndex) = iLoop
                '        End If
                '    Next iLoop
                ''ElseIf tmAstInfo(iIndex).iStatus = 21 Then
                'ElseIf gIsAstStatus(tmAstInfo(iIndex).iStatus, ASTEXTENDED_BONUS) Then
                '    For iLoop = 0 To UBound(tgStatusTypes) Step 1
                '        If (tgStatusTypes(iLoop).iPledged = 2) Or (Trim$(tgStatusTypes(iLoop).sName) = "Bonus") Then
                '            lbcStatus.AddItem Trim$(tgStatusTypes(iLoop).sName)
                '            lbcStatus.ItemData(lbcStatus.NewIndex) = iLoop
                '        End If
                '    Next iLoop
                'ElseIf gIsAstStatus(tmAstInfo(iIndex).iStatus, ASTEXTENDED_REPLACEMENT) Then
                '    For iLoop = 0 To UBound(tgStatusTypes) Step 1
                '        If (tgStatusTypes(iLoop).iPledged = 2) Or (Trim$(tgStatusTypes(iLoop).sName) = "Replacement") Then
                '            lbcStatus.AddItem Trim$(tgStatusTypes(iLoop).sName)
                '            lbcStatus.ItemData(lbcStatus.NewIndex) = iLoop
                '        End If
                '    Next iLoop
                '
                'Else
                '    For iLoop = 0 To UBound(tgStatusTypes) Step 1
                '        'If tgStatusTypes(gGetAirStatus(iLoop)).iStatus < 20 Then
                '        If tgStatusTypes(iLoop).iStatus < ASTEXTENDED_MG Then
                '            '3/11/11: Remove 7-Air Outside Pledge and 8-Air not pledged
                '            If (tgStatusTypes(iLoop).iStatus <> 6) And (tgStatusTypes(iLoop).iStatus <> 7) Then
                '                lbcStatus.AddItem Trim$(tgStatusTypes(iLoop).sName)
                '                lbcStatus.ItemData(lbcStatus.NewIndex) = iLoop
                '            End If
                '        End If
                '    Next iLoop
                'End If
                mPopStatus lmEnableRow
                gSetListBoxHeight lbcStatus, 4
                slStr = grdPost.TextMatrix(grdPost.Row, STATUSINDEX)
                ilIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                If ilIndex >= 0 Then
                    lbcStatus.ListIndex = ilIndex
                Else
                    lbcStatus.ListIndex = 0
                End If
                txtDropdown.Text = lbcStatus.List(lbcStatus.ListIndex)
                If txtDropdown.Height > grdPost.RowHeight(grdPost.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdPost.RowHeight(grdPost.Row) - 15
                End If
                txtDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcStatus.Visible = True
                txtDropdown.SetFocus
        End Select
    End If
End Sub

Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    gGrid_Clear grdPost, True
    For llRow = grdPost.FixedRows To grdPost.Rows - 1 Step 1
        For llCol = VEHICLEINDEX To PLEDGETIMEINDEX Step 1
            grdPost.Row = llRow
            grdPost.Col = llCol
            grdPost.CellBackColor = LIGHTYELLOW
        Next llCol
        grdPost.Col = LENINDEX
        grdPost.CellBackColor = LIGHTYELLOW
    Next llRow
End Sub


Private Function mSaveRow(ilDelRow As Integer) As Integer
    Dim sStr As String
    Dim sDate As String
    Dim sTime As String
    Dim sAirDate As String
    Dim sAirTime As String
    Dim sIAirDate As String
    Dim sIAirTime As String
    Dim lCode As Long
    Dim iIndex As Integer
    Dim iStatus As Integer
    Dim sStatus As String
    Dim iRow As Integer
    Dim iChg As Integer
    Dim sFdDate As String
    Dim sFdTime As String
    Dim sPdDate As String
    Dim sPdSTime As String
    Dim sPdETime As String
    Dim iLoop As Integer
    Dim iTRow As Integer
    Dim lAstCode As Long
    Dim iRows As Integer
    Dim ilStatIdx As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim ilAdfCode As Integer
    Dim llDATCode As Long
    Dim llCpfCode As Long
    Dim llRsfCode As Long
    Dim slStationCompliant As String
    Dim slAgencyCompliant As String
    Dim slAffidavitSource As String
    Dim llCntrNo As Long
    Dim ilLen As Integer
    
    'D.S. 7/24/01 ilStatIdx replaces iStatus as an index into tgStatusTypes. iStatus
    'was causing errors whenever its value was 20 = MG or 21 = Bonus as the tgStatusTypes
    'array is 0 - 10 only
        
    On Error GoTo ErrHand
    
    mSaveRow = True
    If UBound(tmAstInfo) <= LBound(tmAstInfo) Then
        Exit Function
    End If
    If sgUstWin(7) <> "I" Then
        Exit Function
    End If
    llRow = grdPost.Row
    sStr = Trim$(grdPost.TextMatrix(llRow, AIRDATEINDEX))
    iStatus = -1
    sStatus = Trim$(grdPost.TextMatrix(llRow, STATUSINDEX))
    For iLoop = 0 To UBound(tgStatusTypes) Step 1
        If StrComp(sStatus, Trim$(tgStatusTypes(iLoop).sName), 1) = 0 Then
            iStatus = tgStatusTypes(iLoop).iStatus
            ilStatIdx = iLoop
            Exit For
        End If
    Next iLoop
    If iStatus >= 0 Then
        'If iStatus < 20 Then
        If (gIsAstStatus(iStatus, ASTEXTENDED_MG) = False) And (gIsAstStatus(iStatus, ASTEXTENDED_BONUS) = False) And (gIsAstStatus(iStatus, ASTEXTENDED_REPLACEMENT) = False) Then
            If tgStatusTypes(ilStatIdx).iPledged = 2 Then
                sAirDate = ""
                sAirTime = ""
            Else
                sDate = Trim$(grdPost.TextMatrix(llRow, AIRDATEINDEX))
                sTime = Trim$(grdPost.TextMatrix(llRow, AIRTIMEINDEX))
                sAirDate = Format$(sDate, sgShowDateForm)
                sAirTime = Format$(sTime, sgShowTimeWSecForm)
            End If
        Else
            sDate = Trim$(grdPost.TextMatrix(llRow, AIRDATEINDEX))
            sTime = Trim$(grdPost.TextMatrix(llRow, AIRTIMEINDEX))
            sAirDate = Format$(sDate, sgShowDateForm)
            sAirTime = Format$(sTime, sgShowTimeWSecForm)
        End If
    Else
        Exit Function
    End If
    iIndex = Val(grdPost.TextMatrix(llRow, ASTINDEX))
    lCode = tmAstInfo(iIndex).lCode
    lmAttCode = tmAstInfo(iIndex).lAttCode
    imShttCode = tmAstInfo(iIndex).iShttCode
    imVefCode = tmAstInfo(iIndex).iVefCode
    lmSdfCode = tmAstInfo(iIndex).lSdfCode
    iChg = False
    If iStatus >= 0 Then
        'If iStatus < 20 Then
        If (gIsAstStatus(iStatus, ASTEXTENDED_MG) = False) And (gIsAstStatus(iStatus, ASTEXTENDED_BONUS) = False) And (gIsAstStatus(iStatus, ASTEXTENDED_REPLACEMENT) = False) Then
            If (tgStatusTypes(ilStatIdx).iPledged = 0) Or (tgStatusTypes(ilStatIdx).iPledged = 1) Or (tgStatusTypes(ilStatIdx).iPledged = 3) Then
                If ((gIsDate(sAirDate) = True) And (gIsTime(sTime) = True)) And (Len(Trim$(sTime)) <> 0) Then
                    If (DateValue(gAdjYear(sAirDate)) <> DateValue(gAdjYear(tmAstInfo(iIndex).sAirDate))) Or (gTimeToLong(sTime, False) <> gTimeToLong(tmAstInfo(iIndex).sAirTime, False)) Or (gGetAirStatus(tmAstInfo(iIndex).iStatus) <> iStatus) Then
                        iChg = True
                    End If
                End If
            Else
                If (gGetAirStatus(tmAstInfo(iIndex).iStatus) <> iStatus) Then
                    iChg = True
                End If
            End If
        Else
            If ((gIsDate(sAirDate) = True) And (gIsTime(sTime) = True)) And (Len(Trim$(sTime)) <> 0) Then
                If (DateValue(gAdjYear(sAirDate)) <> DateValue(gAdjYear(tmAstInfo(iIndex).sAirDate))) Or (gTimeToLong(sTime, False) <> gTimeToLong(tmAstInfo(iIndex).sAirTime, False)) Or (gGetAirStatus(tmAstInfo(iIndex).iStatus) <> iStatus) Then
                    iChg = True
                End If
            End If
        End If
        If iChg Then
            'If iStatus < 20 Then
            If (gIsAstStatus(iStatus, ASTEXTENDED_MG) = False) And (gIsAstStatus(iStatus, ASTEXTENDED_BONUS) = False) And (gIsAstStatus(iStatus, ASTEXTENDED_REPLACEMENT) = False) Then
                If (tgStatusTypes(ilStatIdx).iPledged = 0) Or (tgStatusTypes(ilStatIdx).iPledged = 1) Or (tgStatusTypes(ilStatIdx).iPledged = 3) Then
                    If Not gIsDate(sAirDate) Then   'not valid.
                        iChg = False
                    End If
                    If (gIsTime(sTime) = False) Or (Len(Trim$(sTime)) = 0) Then   'Time not valid.
                        iChg = False
                    End If
                End If
            Else
                If Not gIsDate(sAirDate) Then   'not valid.
                    iChg = False
                End If
                If (gIsTime(sTime) = False) Or (Len(Trim$(sTime)) = 0) Then   'Time not valid.
                    iChg = False
                End If
            End If
        End If
    End If
    If iChg Then
        If tmAstInfo(iIndex).lCode > 0 Then
            'If (tmAstInfo(iIndex).iStatus <> iStatus) And (tmAstInfo(iIndex).iStatus = 20) Then
            If (ASTEXTENDED_MG <> iStatus) And (gIsAstStatus(tmAstInfo(iIndex).iStatus, ASTEXTENDED_MG)) Then
                'Remove MG and insert missed
                SQLQuery = "SELECT * FROM ast"
                SQLQuery = SQLQuery + " WHERE (astCode = " & lCode & ")"
                Set rst = gSQLSelectCall(SQLQuery)
                If Not rst.EOF Then
                    lAstCode = rst!astLkAstCode 'rst!astSdfCode
                    cnn.BeginTrans
                    SQLQuery = "DELETE FROM ast WHERE (astCode = " & lCode & ")"
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        gSetMousePointer grdPost, grdPost, vbDefault
                        gHandleError "AffErrorLog.txt", "CPDateTimes-mSaveRow"
                        cnn.RollbackTrans
                        mSaveRow = False
                        Exit Function
                    End If
                    lCode = lAstCode
                    SQLQuery = "SELECT * FROM ast"
                    SQLQuery = SQLQuery + " WHERE (astCode = " & lCode & ")"
                    Set rst = gSQLSelectCall(SQLQuery)
                    If Not rst.EOF Then
                        'tmAstInfo(iIndex).sFeedDate = Format$(rst!astFeedDate, sgShowDateForm)
                        'If Second(rst!astFeedTime) <> 0 Then
                        '    tmAstInfo(iIndex).sFeedTime = Format$(rst!astFeedTime, sgShowTimeWSecForm)
                        'Else
                        '    tmAstInfo(iIndex).sFeedTime = Format$(rst!astFeedTime, sgShowTimeWOSecForm)
                        'End If
                        'If igTimes = 0 Then
                        '    grdPost.TextMatrix(llRow, 2) = tmAstInfo(iIndex).sFeedDate
                        '    grdPost.TextMatrix(llRow, 3) = tmAstInfo(iIndex).sFeedTime
                        'Else
                        '    grdPost.TextMatrix(llRow, 1) = tmAstInfo(iIndex).sFeedDate
                        '    grdPost.TextMatrix(llRow, 2) = tmAstInfo(iIndex).sFeedTime
                        'End If
                        SQLQuery = "UPDATE ast SET "
                        SQLQuery = SQLQuery + "astStatus = " & iStatus & ", "
                        SQLQuery = SQLQuery & "astLkAstCode = 0" & ", "
                        SQLQuery = SQLQuery + "astCPStatus = " & "1"
                        '10/19/18: added setting user
                        SQLQuery = SQLQuery + ", " & "astUstCode = " & igUstCode
                        SQLQuery = SQLQuery + " WHERE (astCode = " & lCode & ")"
                        'cnn.Execute SQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand:
                            gSetMousePointer grdPost, grdPost, vbDefault
                            gHandleError "AffErrorLog.txt", "CPDateTimes-mSaveRow"
                            cnn.RollbackTrans
                            mSaveRow = False
                            Exit Function
                        End If
                        cnn.CommitTrans
                        'If (ilDelRow) And (DateValue(gAdjYear(tmAstInfo(iIndex).sFeedDate)) >= DateValue(gAdjYear(smPostSDate))) Or (DateValue(gAdjYear(tmAstInfo(iIndex).sFeedDate)) <= DateValue(gAdjYear(smPostEDate))) Then
                            iRows = grdPost.Rows - 1
                            grdPost.RemoveItem llRow
                            If iRows <> llRow Then
                                For iLoop = iIndex To UBound(tmAstInfo) - 1 Step 1
                                    tmAstInfo(iLoop) = tmAstInfo(iLoop + 1)
                                Next iLoop
                                ReDim Preserve tmAstInfo(0 To UBound(tmAstInfo) - 1) As ASTINFO
                                grdPost.Redraw = False
                                For iRow = 0 To UBound(tmAstInfo) - 1 Step 1
                                    If iIndex < Val(grdPost.TextMatrix(iRow + grdPost.FixedRows, ASTINDEX)) Then
                                        grdPost.TextMatrix(iRow + 1, ASTINDEX) = Val(grdPost.TextMatrix(iRow + grdPost.FixedRows, ASTINDEX)) - 1
                                    End If
                                Next iRow
                                grdPost.Redraw = True
                            End If
                        'End If
                    Else
                        cnn.RollbackTrans
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            'ElseIf (tmAstInfo(iIndex).iStatus <> iStatus) And (tmAstInfo(iIndex).iStatus = 21) Then
            ElseIf (ASTEXTENDED_BONUS <> iStatus) And (gIsAstStatus(tmAstInfo(iIndex).iStatus, ASTEXTENDED_BONUS)) Then
                'Remove spot
                'cnn.BeginTrans
                SQLQuery = "DELETE FROM ast WHERE (astCode = " & lCode & ")"
                'cnn.Execute SQLQuery, rdExecDirect
                'cnn.CommitTrans
                If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    gSetMousePointer grdPost, grdPost, vbDefault
                    gHandleError "AffErrorLog.txt", "CPDateTimes-mSaveRow"
                    mSaveRow = False
                    Exit Function
                End If
                'If ilDelRow Then
                    iRows = grdPost.Rows - 1
                    grdPost.RemoveItem llRow
                    If iRows <> llRow Then
                        For iLoop = iIndex To UBound(tmAstInfo) - 1 Step 1
                            tmAstInfo(iLoop) = tmAstInfo(iLoop + 1)
                        Next iLoop
                        ReDim Preserve tmAstInfo(0 To UBound(tmAstInfo) - 1) As ASTINFO
                        grdPost.Redraw = False
                        For iRow = 0 To UBound(tmAstInfo) - 1 Step 1
                            If iIndex < Val(grdPost.TextMatrix(iRow + grdPost.FixedRows, ASTINDEX)) Then
                                grdPost.TextMatrix(iRow + 1, ASTINDEX) = Val(grdPost.TextMatrix(iRow + grdPost.FixedRows, ASTINDEX)) - 1
                            End If
                        Next iRow
                        grdPost.Redraw = True
                    End If
                'End If
            Else
                SQLQuery = "UPDATE ast SET "
                If iStatus >= 0 Then
                    If (tgStatusTypes(ilStatIdx).iPledged = 0) Or (tgStatusTypes(ilStatIdx).iPledged = 1) Or (tgStatusTypes(ilStatIdx).iPledged = 3) Then
                        SQLQuery = SQLQuery + "astAirDate = '" & Format$(sAirDate, sgSQLDateForm) & "', "
                        SQLQuery = SQLQuery + "astAirTime = '" & Format$(sAirTime, sgSQLTimeForm) & "', "
                        '3/15/16
                        tmAstInfo(iIndex).sAirDate = sAirDate
                        tmAstInfo(iIndex).sAirTime = sAirTime
                    End If
                End If
                SQLQuery = SQLQuery + "astStatus = " & iStatus & ", "
                SQLQuery = SQLQuery + "astCPStatus = " & "1"
                '10/19/18: added setting user
                SQLQuery = SQLQuery + ", " & "astUstCode = " & igUstCode
                SQLQuery = SQLQuery + " WHERE (astCode = " & lCode & ")"
                'cnn.BeginTrans
                'cnn.Execute SQLQuery, rdExecDirect
                'cnn.CommitTrans
                If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    gSetMousePointer grdPost, grdPost, vbDefault
                    gHandleError "AffErrorLog.txt", "CPDateTimes-mSaveRow"
                    mSaveRow = False
                    Exit Function
                End If
                '3/15/16
                'If (iStatus <= 1) Or (iStatus = 9) Or (iStatus = 10) Then
                '    mClearMGs lCode
                'End If
                If ((iStatus <= 1) Or (iStatus = 9) Or (iStatus = 10)) And (tmAstInfo(iIndex).iStatus Mod 100 > 10) And (tmAstInfo(iIndex).iStatus Mod 100 < 14) Then
                    mClearMGs lCode
                End If
                '3/15/16
                If (gIsAstStatus(tmAstInfo(iIndex).iStatus, ASTEXTENDED_MG) = False) And (gIsAstStatus(tmAstInfo(iIndex).iStatus, ASTEXTENDED_BONUS) = False) And (gIsAstStatus(tmAstInfo(iIndex).iStatus, ASTEXTENDED_REPLACEMENT) = False) Then
                    tmAstInfo(iIndex).iStatus = iStatus
                    mCheckForMG tmAstInfo(iIndex)
                    iStatus = tmAstInfo(iIndex).iStatus
                    For iLoop = 0 To UBound(tgStatusTypes) Step 1
                        If iStatus = tgStatusTypes(iLoop).iStatus Then
                            grdPost.TextMatrix(llRow, STATUSINDEX) = Trim$(tgStatusTypes(iLoop).sName)
                            Exit For
                        End If
                    Next iLoop
                    If bmMGCreated = True Then
                        grdPost.TextMatrix(llRow, INFOINDEX) = "MG: " & grdPost.TextMatrix(llRow, AIRDATEINDEX) & " " & grdPost.TextMatrix(llRow, AIRTIMEINDEX)
                        grdPost.TextMatrix(llRow, AIRDATEINDEX) = ""
                        grdPost.TextMatrix(llRow, AIRTIMEINDEX) = ""
                    End If
                End If
                'Set as updated
                For llCol = VEHICLEINDEX To LENINDEX Step 1
                    grdPost.Row = llRow
                    grdPost.Col = llCol
                    grdPost.CellForeColor = DARKGREEN   'vbGreen
                Next llCol
            End If
        Else
            If iStatus >= 0 Then
                If (tgStatusTypes(ilStatIdx).iPledged = 0) Or (tgStatusTypes(ilStatIdx).iPledged = 1) Or (tgStatusTypes(ilStatIdx).iPledged = 3) Then
                    sIAirDate = sAirDate
                    sIAirTime = sAirTime
                Else
                    sIAirDate = Format$(tmAstInfo(iIndex).sAirDate, sgShowDateForm)
                    sIAirTime = Format$(tmAstInfo(iIndex).sAirTime, sgShowTimeWSecForm)
                End If
            Else
                sIAirDate = Format$(tmAstInfo(iIndex).sAirDate, sgShowDateForm)
                sIAirTime = Format$(tmAstInfo(iIndex).sAirTime, sgShowTimeWSecForm)
            End If
            sFdDate = Format$(tmAstInfo(iIndex).sFeedDate, sgShowDateForm)
            sFdTime = Format$(tmAstInfo(iIndex).sFeedTime, sgShowTimeWSecForm)
            sPdDate = Format$(tmAstInfo(iIndex).sPledgeDate, sgShowDateForm)
            sPdSTime = Format$(tmAstInfo(iIndex).sPledgeStartTime, sgShowTimeWSecForm)
            sPdETime = Format$(Trim$(tmAstInfo(iIndex).sPledgeEndTime), sgShowTimeWSecForm)
            If Len(Trim$(sPdSTime)) = 0 Then
                sPdSTime = sFdTime
            End If
            If Len(Trim$(sPdETime)) = 0 Then
                sPdETime = sPdSTime
            End If
            ilAdfCode = tmAstInfo(iIndex).iAdfCode
            llDATCode = tmAstInfo(iIndex).lDatCode
            llCpfCode = tmAstInfo(iIndex).lCpfCode
            llRsfCode = tmAstInfo(iIndex).lRRsfCode
            llCntrNo = tmAstInfo(iIndex).lCntrNo
            ilLen = tmAstInfo(iIndex).iLen
            slStationCompliant = ""
            slAgencyCompliant = ""
            slAffidavitSource = ""
            SQLQuery = "INSERT INTO ast"
            SQLQuery = SQLQuery + "(astAtfCode, astShfCode, astVefCode, "
            SQLQuery = SQLQuery + "astSdfCode, astLsfCode, astAirDate, astAirTime, "
            '12/13/13: New AST format
            'SQLQuery = SQLQuery + "astStatus, astCPStatus, astFeedDate, astFeedTime, astPledgeDate, "
            'SQLQuery = SQLQuery + "astPledgeStartTime, astPledgeEndTime, astPledgeStatus)"
            SQLQuery = SQLQuery + "astStatus, astCPStatus, astFeedDate, astFeedTime, "
            SQLQuery = SQLQuery + "astAdfCode, astDatCode, astCpfCode, astRsfCode, astStationCompliant, astAgencyCompliant, astAffidavitSource, astCntrNo, astLen, astLkAstCode, astMissedMnfCode, astUstCode)"
            SQLQuery = SQLQuery + " VALUES "
            SQLQuery = SQLQuery + "(" & lmAttCode & ", " & imShttCode & ", "
            SQLQuery = SQLQuery & imVefCode & ", " & lmSdfCode & ", " & tmAstInfo(iIndex).lLstCode & ", "
            SQLQuery = SQLQuery + "'" & Format$(sIAirDate, sgSQLDateForm) & "', '" & Format$(sIAirTime, sgSQLTimeForm) & "', "
            SQLQuery = SQLQuery & iStatus & ", " & "1" & ", '" & Format$(sFdDate, sgSQLDateForm) & "', "
            'SQLQuery = SQLQuery & "'" & Format$(sFdTime, sgSQLTimeForm) & "', '" & Format$(sPdDate, sgSQLDateForm) & "', "
            'SQLQuery = SQLQuery & "'" & Format$(sPdSTime, sgSQLTimeForm) & "', '" & Format$(sPdETime, sgSQLTimeForm) & "', " & tmAstInfo(iIndex).iPledgeStatus & ")"
            SQLQuery = SQLQuery & "'" & Format$(sFdTime, sgSQLTimeForm) & "', "
            SQLQuery = SQLQuery & ilAdfCode & ", " & llDATCode & ", " & llCpfCode & ", " & llRsfCode & ", "
            SQLQuery = SQLQuery & "'" & slStationCompliant & "', '" & slAgencyCompliant & "', '" & slAffidavitSource & "', " & llCntrNo & ", " & ilLen & ", " & 0 & ", " & 0 & ", " & igUstCode & ")"
            'cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            'cnn.CommitTrans
            If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                gSetMousePointer grdPost, grdPost, vbDefault
                gHandleError "AffErrorLog.txt", "CPDateTimes-mSaveRow"
                mSaveRow = False
                Exit Function
            End If
            SQLQuery = "Select MAX(astCode) from ast"
            Set rst = gSQLSelectCall(SQLQuery)
            tmAstInfo(iIndex).lCode = rst(0).Value
            'Set as updated
            For llCol = VEHICLEINDEX To LENINDEX Step 1
                grdPost.Row = llRow
                grdPost.Col = llCol
                grdPost.CellForeColor = DARKGREEN   'vbGreen
            Next llCol
        End If
        If iStatus >= 0 Then
            If (tgStatusTypes(ilStatIdx).iPledged = 0) Or (tgStatusTypes(ilStatIdx).iPledged = 1) Or (tgStatusTypes(ilStatIdx).iPledged = 3) Then
                tmAstInfo(iIndex).sAirDate = Format$(sAirDate, sgShowDateForm)
                If Second(sAirTime) <> 0 Then
                    tmAstInfo(iIndex).sAirTime = Format$(sAirTime, sgShowTimeWSecForm)
                Else
                    tmAstInfo(iIndex).sAirTime = Format$(sAirTime, sgShowTimeWOSecForm)
                End If
            End If
        End If
        tmAstInfo(iIndex).iStatus = iStatus
    End If
    On Error GoTo 0
    Exit Function
ErrHand:
    gSetMousePointer grdPost, grdPost, vbDefault
    gHandleError "AffErrorLog.txt", "Date/Time-mSaveRow"
    mSaveRow = False
End Function

Private Sub cmcDropDown_Click()
    Select Case grdPost.Col
        Case STATUSINDEX
            lbcStatus.Visible = Not lbcStatus.Visible
    End Select
End Sub

Private Sub cmdBonus_Click()
    gSetMousePointer grdPost, grdPost, vbHourglass
    
    If imFieldChgd = True Then
        If Not mSave() Then
            gSetMousePointer grdPost, grdPost, vbDefault
            Exit Sub
        End If
    End If
    gSetMousePointer grdPost, grdPost, vbDefault
    igUpdateDTGrid = False
    frmAddBonus.Show vbModal
    If igUpdateDTGrid Then
        gSetMousePointer grdPost, grdPost, vbHourglass
        mGetAst
        gSetMousePointer grdPost, grdPost, vbDefault
    End If
End Sub

Private Sub cmdBonus_GotFocus()
    mPostSetShow
End Sub

Private Sub cmdCancel_Click()
    Unload frmDateTimes
End Sub

Private Sub cmdCancel_GotFocus()
    mPostSetShow
End Sub

Private Sub cmdDone_Click()
    Dim iLoop As Integer
    Dim iPostingStatus As Integer
    Dim iPosted As Integer
    Dim sFWkDate As String
    Dim sLWkDate As String
    Dim iRet As Integer
    Dim iStatus As Integer
    Dim ilAst As Integer
    
    If imFieldChgd = True Then
        gSetMousePointer grdPost, grdPost, vbHourglass
        If sgUstWin(7) = "I" Then
            imIgnoreRowChg = True
            If Not mSave() Then
                gSetMousePointer grdPost, grdPost, vbDefault
                Exit Sub
            End If
            imIgnoreRowChg = False
            gSetMousePointer grdPost, grdPost, vbHourglass
'            If optStatus(0).Value Then
'                iPostingStatus = 2 'Completed
'                iStatus = 1
'            ElseIf optStatus(1).Value Then
'                iPostingStatus = 1 'Partially posted
'                iStatus = 0
'            Else
'                iPostingStatus = 0 'Not Posted
'                iStatus = 0
'            End If
'            If (igCPStatus <> iStatus) Or (igCPPostingStatus <> iPostingStatus) Then
'                On Error GoTo ErrHand
'                For iLoop = 0 To UBound(tgCPPosting) - 1 Step 1
'                    sFWkDate = Format$(gObtainPrevMonday(tgCPPosting(iLoop).sDate), sgShowDateForm)
'                    If igTimes = 0 Then
'                        sLWkDate = Format$(gObtainEndStd(tgCPPosting(iLoop).sDate), sgShowDateForm)
'                    Else
'                        sLWkDate = Format$(gObtainNextSunday(tgCPPosting(iLoop).sDate), sgShowDateForm)
'                    End If
'
'                    'SQLQuery = "UPDATE cptt SET "
'                    'SQLQuery = SQLQuery + "cpttStatus = " & iStatus & ", "
'                    'SQLQuery = SQLQuery + "cpttPostingStatus = " & iPostingStatus
'                    'SQLQuery = SQLQuery + " WHERE (cpttVefCode = " & tgCPPosting(iLoop).iVefCode
'                    'SQLQuery = SQLQuery + " AND cpttShfCode = " & tgCPPosting(iLoop).iShttCode
'                    'SQLQuery = SQLQuery + " AND (cpttStartDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')" & ")"
'                    ''cnn.BeginTrans
'                    ''cnn.Execute SQLQuery, rdExecDirect
'                    ''cnn.CommitTrans
'                    'If gSQLWait(SQLQuery, True) <> 0 Then
'                    '    gSetMousePointer grdPost, grdPost, vbDefault
'                    '    Unload frmDateTimes
'                    '    Exit Sub
'                    'End If
'                    'If optStatus(0).Value Or optStatus(1).Value Then
'                    If lgSelGameGsfCode <= 0 Then
'                        If iPostingStatus = 2 Then
'                            SQLQuery = "UPDATE ast SET "
'                            'If optStatus(0).Value Then
'                                SQLQuery = SQLQuery + "astCPStatus = " & "1"
'                            'Else
'                            '    SQLQuery = SQLQuery + "astCPStatus = " & "2"
'                            'End If
'                            SQLQuery = SQLQuery + " WHERE (astAtfCode = " & tgCPPosting(iLoop).lattCode
'                            SQLQuery = SQLQuery + " AND astCPStatus = 0"
'                            SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')" & ")"
'                            'cnn.BeginTrans
'                            'cnn.Execute SQLQuery, rdExecDirect
'                            'cnn.CommitTrans
'                            If gSQLWait(SQLQuery, True) <> 0 Then
'                                gSetMousePointer grdPost, grdPost, vbDefault
'                                Unload frmDateTimes
'                                Exit Sub
'                            End If
'                        ElseIf iPostingStatus = 0 Then
'                            SQLQuery = "UPDATE ast SET "
'                            'If optStatus(0).Value Then
'                                SQLQuery = SQLQuery + "astCPStatus = " & "0" & ", "
'                            'Else
'                            '    SQLQuery = SQLQuery + "astCPStatus = " & "2"
'                            'End If
'                            SQLQuery = SQLQuery + "astStatus = astPledgeStatus"
'                            SQLQuery = SQLQuery + " WHERE (astAtfCode = " & tgCPPosting(iLoop).lattCode
'                            SQLQuery = SQLQuery + " AND astCPStatus = 1"
'                            SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')" & ")"
'                            'cnn.BeginTrans
'                            'cnn.Execute SQLQuery, rdExecDirect
'                            'cnn.CommitTrans
'                            If gSQLWait(SQLQuery, True) <> 0 Then
'                                gSetMousePointer grdPost, grdPost, vbDefault
'                                Unload frmDateTimes
'                                Exit Sub
'                            End If
'                            mSetRowToUnposted
'                        End If
'                    Else
'                        'Change each game individually unless setup special link to lst Update ast set astCPStatus = 1 where field from (select ast, lst where lstcode = astlstcode and ........)
'                        If (iPostingStatus = 2) Or (iPostingStatus = 0) Then
'                            For ilAst = 0 To UBound(tmAstInfo) - 1 Step 1
'                                If iPostingStatus = 2 Then
'                                SQLQuery = "UPDATE ast SET "
'                                If iPostingStatus = 2 Then
'                                    SQLQuery = SQLQuery + "astCPStatus = " & "1"
'                                ElseIf iPostingStatus = 0 Then
'                                    SQLQuery = SQLQuery + "astCPStatus = " & "0"
'                                    SQLQuery = SQLQuery + "astStatus = astPledgeStatus"
'                                End If
'                                SQLQuery = SQLQuery + " WHERE (astCode = " & tmAstInfo(ilAst).lCode & ")"
'                                If gSQLWait(SQLQuery, True) <> 0 Then
'                                    gSetMousePointer grdPost, grdPost, vbDefault
'                                    Unload frmDateTimes
'                                    Exit Sub
'                                End If
'                            Next ilAst
'                            If iPostingStatus = 0 Then
'                                mSetRowToUnposted
'                            End If
'                        End If
'                    End If
'                    If lgSelGameGsfCode > 0 Then
'                        If DateValue(sLWkDate) < DateValue(Format$(gNow(), "m/d/yy")) Then
'                            If iPostingStatus = 2 Then
'                                SQLQuery = "Select astCode FROM ast WHERE astCPStatus = 0"
'                                SQLQuery = SQLQuery + " AND astAtfCode = " & lmAttCode
'                                SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')"
'                                Set rst_Ast = gSQLSelectCall(SQLQuery)
'                                If Not rst_Ast.EOF Then
'                                    iPostingStatus = 1 'Partially posted
'                                    iStatus = 0
'                                End If
'                            ElseIf iPostingStatus = 0 Then
'                                SQLQuery = "Select astCode FROM ast WHERE astCPStatus = 1"
'                                SQLQuery = SQLQuery + " AND astAtfCode = " & lmAttCode
'                                SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')"
'                                Set rst_Ast = gSQLSelectCall(SQLQuery)
'                                If Not rst_Ast.EOF Then
'                                    iPostingStatus = 1 'Partially posted
'                                    iStatus = 0
'                                End If
'                            End If
'                        Else
'                            If iPostingStatus = 2 Then
'                                iPostingStatus = 1 'Partially posted
'                                iStatus = 0
'                            End If
'                        End If
'                    End If
'                    SQLQuery = "UPDATE cptt SET "
'                    SQLQuery = SQLQuery + "cpttStatus = " & iStatus & ", "
'                    SQLQuery = SQLQuery + "cpttPostingStatus = " & iPostingStatus
'                    SQLQuery = SQLQuery + " WHERE (cpttVefCode = " & tgCPPosting(iLoop).iVefCode
'                    SQLQuery = SQLQuery + " AND cpttShfCode = " & tgCPPosting(iLoop).iShttCode
'                    SQLQuery = SQLQuery + " AND (cpttStartDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')" & ")"
'                    If gSQLWait(SQLQuery, True) <> 0 Then
'                        gSetMousePointer grdPost, grdPost, vbDefault
'                        Unload frmDateTimes
'                        Exit Sub
'                    End If
'                Next iLoop
'                igCPStatus = iStatus
'                igCPPostingStatus = iPostingStatus
'            End If
            mSetASTAndCPTTStatus
        End If
    End If
    On Error GoTo 0
    gSetMousePointer grdPost, grdPost, vbDefault
    Unload frmDateTimes
    Exit Sub
ErrHand:
    gSetMousePointer grdPost, grdPost, vbDefault
    gHandleError "AffErrorLog.txt", ""
End Sub

Private Sub cmdDone_GotFocus()
    mPostSetShow
End Sub

Private Sub cmdMG_Click()
    gSetMousePointer grdPost, grdPost, vbHourglass
    If imFieldChgd = True Then
        If Not mSave() Then
            gSetMousePointer grdPost, grdPost, vbDefault
            Exit Sub
        End If
    End If
    gSetMousePointer grdPost, grdPost, vbDefault
    igUpdateDTGrid = False
    frmAddMG.Show vbModal
    If igUpdateDTGrid Then
        gSetMousePointer grdPost, grdPost, vbHourglass
        mGetAst
        gSetMousePointer grdPost, grdPost, vbDefault
    End If
End Sub

Private Sub cmdMG_GotFocus()
    mPostSetShow
End Sub

Private Sub cmdSave_Click()
    Dim iLoop As Integer
    Dim iPostingStatus As Integer
    Dim iPosted As Integer
    Dim sFWkDate As String
    Dim sLWkDate As String
    Dim iRet As Integer
    Dim iStatus As Integer
    
    If imFieldChgd = True Then
        gSetMousePointer grdPost, grdPost, vbHourglass
        If sgUstWin(7) = "I" Then
            imIgnoreRowChg = True
            If Not mSave() Then
                gSetMousePointer grdPost, grdPost, vbDefault
                Exit Sub
            End If
            imIgnoreRowChg = False
            gSetMousePointer grdPost, grdPost, vbHourglass
'            If optStatus(0).Value Then
'                iPostingStatus = 2 'Completed
'                iStatus = 1
'            ElseIf optStatus(1).Value Then
'                iPostingStatus = 1 'Partially posted
'                iStatus = 0
'            Else
'                iPostingStatus = 0 'Not Posted
'                iStatus = 0
'            End If
'            If (igCPStatus <> iStatus) Or (igCPPostingStatus <> iPostingStatus) Then
'                On Error GoTo ErrHand
'                For iLoop = 0 To UBound(tgCPPosting) - 1 Step 1
'                    sFWkDate = Format$(gObtainPrevMonday(tgCPPosting(iLoop).sDate), sgShowDateForm)
'                    If igTimes = 0 Then
'                        sLWkDate = Format$(gObtainEndStd(tgCPPosting(iLoop).sDate), sgShowDateForm)
'                    Else
'                        sLWkDate = Format$(gObtainNextSunday(tgCPPosting(iLoop).sDate), sgShowDateForm)
'                    End If
'                    SQLQuery = "UPDATE cptt SET "
'                    SQLQuery = SQLQuery + "cpttStatus = " & iStatus & ", "
'                    SQLQuery = SQLQuery + "cpttPostingStatus = " & iPostingStatus
'                    SQLQuery = SQLQuery + " WHERE (cpttVefCode = " & tgCPPosting(iLoop).iVefCode
'                    SQLQuery = SQLQuery + " AND cpttShfCode = " & tgCPPosting(iLoop).iShttCode
'                    SQLQuery = SQLQuery + " AND (cpttStartDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')" & ")"
'                    'cnn.BeginTrans
'                    'cnn.Execute SQLQuery, rdExecDirect
'                    'cnn.CommitTrans
'                    If gSQLWait(SQLQuery, True) <> 0 Then
'                        gSetMousePointer grdPost, grdPost, vbDefault
'                        Unload frmDateTimes
'                        Exit Sub
'                    End If
'                    'If optStatus(0).Value Or optStatus(1).Value Then
'                    If optStatus(0).Value Then
'                        SQLQuery = "UPDATE ast SET "
'                        'If optStatus(0).Value Then
'                            SQLQuery = SQLQuery + "astCPStatus = " & "1"
'                        'Else
'                        '    SQLQuery = SQLQuery + "astCPStatus = " & "2"
'                        'End If
'                        SQLQuery = SQLQuery + " WHERE (astAtfCode = " & tgCPPosting(iLoop).lattCode
'                        SQLQuery = SQLQuery + " AND astCPStatus = 0"
'                        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')" & ")"
'                        'cnn.BeginTrans
'                        'cnn.Execute SQLQuery, rdExecDirect
'                        'cnn.CommitTrans
'                        If gSQLWait(SQLQuery, True) <> 0 Then
'                            gSetMousePointer grdPost, grdPost, vbDefault
'                            Unload frmDateTimes
'                            Exit Sub
'                        End If
'                    ElseIf iPostingStatus = 0 Then
'                        SQLQuery = "UPDATE ast SET "
'                        'If optStatus(0).Value Then
'                            SQLQuery = SQLQuery + "astCPStatus = " & "0" & ", "
'                        'Else
'                        '    SQLQuery = SQLQuery + "astCPStatus = " & "2"
'                        'End If
'                        SQLQuery = SQLQuery + "astStatus = astPledgeStatus"
'                        SQLQuery = SQLQuery + " WHERE (astAtfCode = " & tgCPPosting(iLoop).lattCode
'                        SQLQuery = SQLQuery + " AND astCPStatus = 1"
'                        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')" & ")"
'                        'cnn.BeginTrans
'                        'cnn.Execute SQLQuery, rdExecDirect
'                        'cnn.CommitTrans
'                        If gSQLWait(SQLQuery, True) <> 0 Then
'                            gSetMousePointer grdPost, grdPost, vbDefault
'                            Unload frmDateTimes
'                            Exit Sub
'                        End If
'                        mSetRowToUnposted
'                    End If
'                Next iLoop
'                igCPStatus = iStatus
'                igCPPostingStatus = iPostingStatus
'            End If
            mSetASTAndCPTTStatus
        End If
    End If
    imFieldChgd = False
    cmdSave.Enabled = False
    On Error GoTo 0
    gSetMousePointer grdPost, grdPost, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdPost, grdPost, vbDefault
    gHandleError "AffErrorLog.txt", "frmDateTime-cmdDone"
End Sub

Private Sub cmdSave_GotFocus()
    mPostSetShow
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    
    If imFirstTime Then
        If igTimes = 0 Then
            imAdvtCol = 1
            imDateCol = 2
            imTimeCol = 3
        Else
            imDateCol = 1
            imTimeCol = 2
            imAdvtCol = 3
        End If
        grdPost.ColWidth(SORTINDEX) = 0
        grdPost.ColWidth(ASTINDEX) = 0
        grdPost.ColWidth(CARTISCIINDEX) = 0
        grdPost.ColWidth(INFOINDEX) = 0
        grdPost.ColWidth(VEHICLEINDEX) = grdPost.Width * 0.12
        grdPost.ColWidth(imDateCol) = grdPost.Width * 0.08
        grdPost.ColWidth(imTimeCol) = grdPost.Width * 0.08
        grdPost.ColWidth(PLEDGEDAYSINDEX) = grdPost.Width * 0.06
        grdPost.ColWidth(PLEDGETIMEINDEX) = grdPost.Width * 0.14
        grdPost.ColWidth(AIRDATEINDEX) = grdPost.Width * 0.08
        grdPost.ColWidth(AIRTIMEINDEX) = grdPost.Width * 0.08
        grdPost.ColWidth(LENINDEX) = grdPost.Width * 0.04
        grdPost.ColWidth(STATUSINDEX) = grdPost.Width * 0.08

        grdPost.ColWidth(imAdvtCol) = grdPost.Width - GRIDSCROLLWIDTH  '(5 * grdStation.Columns(6).Width) / 6
        For ilCol = VEHICLEINDEX To LENINDEX Step 1
            If ilCol <> imAdvtCol Then
                grdPost.ColWidth(imAdvtCol) = grdPost.ColWidth(imAdvtCol) - grdPost.ColWidth(ilCol)
            End If
        Next ilCol
        gGrid_AlignAllColsLeft grdPost
        grdPost.TextMatrix(0, VEHICLEINDEX) = "Vehicle"
        grdPost.TextMatrix(0, imDateCol) = "Feed Date"
        grdPost.TextMatrix(0, imTimeCol) = "Feed Time"
        grdPost.TextMatrix(0, imAdvtCol) = "Advertiser/ Product"
        grdPost.TextMatrix(0, PLEDGEDAYSINDEX) = "Pledge Days"
        grdPost.TextMatrix(0, PLEDGETIMEINDEX) = "Pledge Times"
        grdPost.TextMatrix(0, AIRDATEINDEX) = "Air Date"
        grdPost.TextMatrix(0, AIRTIMEINDEX) = "Aired Time"
        grdPost.TextMatrix(0, LENINDEX) = "Len"
        grdPost.TextMatrix(0, STATUSINDEX) = "Status"
        gGrid_IntegralHeight grdPost
        mClearGrid
        grdPost.Row = 0
        For ilCol = VEHICLEINDEX To LENINDEX Step 1
            grdPost.Col = ilCol
            grdPost.CellBackColor = LIGHTBLUE
        Next ilCol
        gSetMousePointer grdPost, grdPost, vbHourglass
        mGetAst
        gSetMousePointer grdPost, grdPost, vbDefault
        imFirstTime = False
        imFieldChgd = False
    End If

End Sub

Private Sub Form_Click()
    mPostSetShow
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Dim iLoop As Integer
    Dim sStationZone As String
    
    Me.Width = Screen.Width / 1.15
    Me.Height = Screen.Height / 1.55
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    sStationZone = ""
    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(iLoop).iCode = tgCPPosting(0).iShttCode Then
            sStationZone = " for " & Trim$(tgStationInfo(iLoop).sCallLetters) & " (" & (tgStationInfo(iLoop).sZone) & ")"
            Exit For
        End If
    Next iLoop
    smPostSDate = Format$(gObtainPrevMonday(tgCPPosting(0).sDate), sgShowDateForm)
    If igTimes = 0 Then
        smPostEDate = Format$(gObtainEndStd(tgCPPosting(0).sDate), sgShowDateForm)
    Else
        smPostEDate = Format$(gObtainNextSunday(tgCPPosting(0).sDate), sgShowDateForm)
    End If
    If igTimes = 0 Then
        frmDateTimes.Caption = "Spots by Advertiser" & sStationZone & " " & smPostSDate & "-" & smPostEDate
    Else
        frmDateTimes.Caption = "Spots by Date" & sStationZone & " " & smPostSDate & "-" & smPostEDate
    End If
    txtKey.Text = ""
    For iLoop = 0 To UBound(tgStatusTypes) Step 1
        If txtKey.Text = "" Then
            txtKey.Text = Trim$(tgStatusTypes(iLoop).sName)
        Else
            txtKey.Text = txtKey.Text & sgCRLF & Trim$(tgStatusTypes(iLoop).sName)
        End If
    Next iLoop
    
End Sub

Private Sub Form_Load()
    Dim sAdvtName As String
    Dim sVefName As String
    Dim iLoop As Integer
    Dim sStatus As String
    Dim sAirDate As String
    Dim sAirTime As String
    Dim sPdDate As String
    Dim sPdDays As String
    Dim sPdTime As String
    Dim sPdSTime As String
    Dim sPdETime As String
    Dim iAst As Integer
    Dim iRet As Integer
    Dim ilCol As Integer
    Dim ilRet As Integer

    On Error GoTo ErrHand
    
    gSetMousePointer grdPost, grdPost, vbHourglass
    frmDateTimes.Caption = "Times by Date - " & sgClientName
    imFirstTime = True
    imIntegralSet = False
    imFirstDrop = True
    imMouseDown = False
    imIgnoreRowChg = False
    imFieldChgd = False
    imBSMode = False
    imShowGridBox = False
    lmTopRow = -1
    imFromArrow = False
    lmEnableRow = -1
    lmEnableCol = -1
    imcKey.Picture = frmDirectory!imcKey.Picture
    imLastRow = -1
    imArrowKeyPressed = 0
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    ReDim tmAstInfo(0 To 0) As ASTINFO
    For iLoop = 0 To UBound(tgStatusTypes) Step 1
        ''If tgStatusTypes(gGetAirStatus(iLoop)).iStatus < 20 Then
        'If tgStatusTypes(iLoop).iStatus < ASTEXTENDED_MG Then
        If tgStatusTypes(iLoop).iStatus < ASTEXTENDED_MG Or ((sgMissedMGBypass = "Y") And (tgStatusTypes(iLoop).iStatus = ASTAIR_MISSED_MG_BYPASS)) Then
            '3/11/11: Remove 7-Air Outside Pledge and 8-Air not pledged
            If (tgStatusTypes(iLoop).iStatus <> 6) And (tgStatusTypes(iLoop).iStatus <> 7) Then
                lbcStatus.AddItem Trim$(tgStatusTypes(iLoop).sName)
                lbcStatus.ItemData(lbcStatus.NewIndex) = iLoop
            End If
        End If
    Next iLoop
    
    
    'mClearGrid
    
    
    'mGetAst
    If igCPPostingStatus = 2 Then
        optStatus(0).Value = True
    ElseIf igCPPostingStatus = 1 Then
        optStatus(1).Value = True
    Else
        'D.S. 06/22/06
        optStatus(2).Value = True
    End If
    imFieldChgd = False
    If sgUstWin(7) <> "I" Then
        cmdMG.Enabled = False
        cmdBonus.Enabled = False
        frcPosting.Enabled = False
    End If
    gSetMousePointer grdPost, grdPost, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdPost, grdPost, vbDefault
    gHandleError "AffErrorLog.txt", "frmDateTime-Form Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    lst_rst.Close
    dat_rst.Close
    ast_rst.Close
    Erase tmCPDat
    Erase tmAstInfo
    Erase tmGameAstInfo
    Set frmDateTimes = Nothing
End Sub

Private Function mSave() As Integer
    Dim sStr As String
    Dim sDate As String
    Dim sTime As String
    Dim sAirDate As String
    Dim sAirTime As String
    Dim sIAirDate As String
    Dim sIAirTime As String
    Dim lCode As Long
    Dim iIndex As Integer
    Dim iStatus As Integer
    Dim sStatus As String
    Dim iRow As Integer
    Dim iChg As Integer
    Dim sFdDate As String
    Dim sFdTime As String
    Dim sPdDate As String
    Dim sPdSTime As String
    Dim sPdETime As String
    Dim iLoop As Integer
    Dim iRet As Integer
    
    On Error GoTo ErrHand
    
    mSave = True
    
    If UBound(tmAstInfo) <= LBound(tmAstInfo) Then
        Exit Function
    End If
    If sgUstWin(7) <> "I" Then
        Exit Function
    End If
    If Not mTestGridValues() Then
        mSave = False
        Exit Function
    End If
    grdPost.Redraw = False
    'For iRow = 0 To UBound(tmAstInfo) - 1 Step 1
    iRow = 0
    Do While iRow <= UBound(tmAstInfo) - 1
        grdPost.Row = grdPost.FixedRows + iRow
        If Not mSaveRow(False) Then
            mSave = False
            Exit Function
        End If
        iRow = iRow + 1
    'Next iRow
    Loop
    On Error GoTo 0
    grdPost.Redraw = True
    Exit Function
ErrHand:
    gSetMousePointer grdPost, grdPost, vbDefault
    gHandleError "AffErrorLog.txt", "frmDateTime-mSave"
    mSave = False
End Function


Public Sub mGetAst()
    Dim sAdvtName As String
    Dim sVefName As String
    Dim iLoop As Integer
    Dim sStatus As String
    Dim sAirDate As String
    Dim sAirTime As String
    Dim sPdDate As String
    Dim sPdDays As String
    Dim sPdTime As String
    Dim sPdSTime As String
    Dim sPdETime As String
    Dim iAst As Integer
    Dim iRet As Integer
    Dim sCart As String
    Dim slInfo As String
    Dim sISCI As String
    Dim sProd As String
    Dim lCrfCsfCode As Long
    Dim sRCart As String
    Dim sRISCI As String
    Dim sRCreative As String
    Dim sRProd As String
    Dim lRCrfCsfCode As Long
    Dim lRCrfCode As Long
    Dim rst_Lst As ADODB.Recordset
    Dim rst_Gsf As ADODB.Recordset
        
    Dim llTRow As Long
    Dim llRow As Long
    Dim llCol As Long
    Dim ilUpper As Integer
    Dim llSet As Long
    
    Dim ilLang As Integer
    Dim ilTeam As Integer
    Dim slStr As String
    
    
    imFieldChgd = False
    llTRow = grdPost.TopRow
    grdPost.Redraw = False
    mClearGrid
    llRow = grdPost.FixedRows
    
    'D.S. 11/21/05
'    iRet = gGetMaxAstCode()
'    If Not iRet Then
'        Exit Sub
'    End If
    
    bgTaskBlocked = False
    sgTaskBlockedName = "Affiliate Affidavit"
    
    lacGame.Caption = ""
    If lgSelGameGsfCode <= 0 Then
        '12/10/08:  Create ast because of the blackout
        'iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), -1, False, False, True)
        iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), -1, True, False, True, , , , , , True)
    Else
        iRet = gGetAstInfo(hmAst, tmCPDat(), tmGameAstInfo(), -1, True, False, True, , , , , , True)
        ReDim tmAstInfo(0 To UBound(tmGameAstInfo)) As ASTINFO
        ilUpper = 0
        For iAst = 0 To UBound(tmGameAstInfo) - 1 Step 1
            If tmGameAstInfo(iAst).lgsfCode = lgSelGameGsfCode Then
                tmAstInfo(ilUpper) = tmGameAstInfo(iAst)
                ilUpper = ilUpper + 1
            End If
        Next iAst
        ReDim Preserve tmAstInfo(0 To ilUpper) As ASTINFO
        
        SQLQuery = "SELECT * FROM GSF_Game_Schd WHERE (gsfCode = " & lgSelGameGsfCode & ")"
        Set rst_Gsf = gSQLSelectCall(SQLQuery)
        If Not rst_Gsf.EOF Then
            slStr = ""
            'Feed Source
            If ((Asc(sgSpfSportInfo) And USINGFEED) = USINGFEED) Then
                If rst_Gsf!gsfFeedSource = "V" Then
                    slStr = "Feed: Visting"
                Else
                    slStr = "Feed: Home"
                End If
            End If
            'Language
            If ((Asc(sgSpfSportInfo) And USINGLANG) = USINGLANG) Then
                For ilLang = LBound(tgLangInfo) To UBound(tgLangInfo) - 1 Step 1
                    If tgLangInfo(ilLang).iCode = rst_Gsf!gsfLangMnfCode Then
                        If slStr = "" Then
                            slStr = Trim$(tgLangInfo(ilLang).sName)
                        Else
                            slStr = slStr & " " & Trim$(tgLangInfo(ilLang).sName)
                        End If
                        Exit For
                    End If
                Next ilLang
            End If
            'Visiting Team
            For ilTeam = LBound(tgTeamInfo) To UBound(tgTeamInfo) - 1 Step 1
                If tgTeamInfo(ilTeam).iCode = rst_Gsf!gsfVisitMnfCode Then
                    If slStr = "" Then
                        slStr = Trim$(tgTeamInfo(ilTeam).sName)
                    Else
                        slStr = slStr & " " & Trim$(tgTeamInfo(ilTeam).sName)
                    End If
                    Exit For
                End If
            Next ilTeam
            'Home Team
            For ilTeam = LBound(tgTeamInfo) To UBound(tgTeamInfo) - 1 Step 1
                If tgTeamInfo(ilTeam).iCode = rst_Gsf!gsfHomeMnfCode Then
                    If slStr = "" Then
                        slStr = Trim$(tgTeamInfo(ilTeam).sName)
                    Else
                        slStr = slStr & " @ " & Trim$(tgTeamInfo(ilTeam).sName)
                    End If
                    Exit For
                End If
            Next ilTeam
            'Air Date
            slStr = slStr & " on " & Format$(rst_Gsf!gsfAirDate, sgShowDateForm)
            'Start Time
            slStr = slStr & " at " & Format$(rst_Gsf!gsfAirTime, sgShowTimeWSecForm)
            lacGame.Caption = slStr
        End If
    End If
    '2/5/18: add block message
    If bgTaskBlocked Then
        gMsgBox "*** Spots not obtained as blocked", vbCritical
    End If
    bgTaskBlocked = False
    sgTaskBlockedName = ""
    
    gCloseRegionSQLRst

    For iAst = 0 To UBound(tmAstInfo) - 1 Step 1
        'If tmAstInfo(iAst).iStatus = 20 Then
        If gIsAstStatus(tmAstInfo(iAst).iStatus, ASTEXTENDED_MG) Then
            sStatus = "MG"
        'ElseIf tmAstInfo(iAst).iStatus = 21 Then
        ElseIf gIsAstStatus(tmAstInfo(iAst).iStatus, ASTEXTENDED_BONUS) Then
            sStatus = "Bonus"
        ElseIf gIsAstStatus(tmAstInfo(iAst).iStatus, ASTEXTENDED_REPLACEMENT) Then
            sStatus = "Replacement"
        Else
            sStatus = Trim$(tgStatusTypes(gGetAirStatus(tmAstInfo(iAst).iStatus)).sName)
        End If
        sAirDate = tmAstInfo(iAst).sAirDate
        sAirTime = tmAstInfo(iAst).sAirTime
        sPdDays = tmAstInfo(iAst).sPdDays
        sPdSTime = Trim$(tmAstInfo(iAst).sPledgeStartTime)
        sPdETime = Trim$(tmAstInfo(iAst).sPledgeEndTime)
        '3/7/16: For MG's show pledge information
        ''If tmAstInfo(iAst).iStatus < 20 Then
        'If (gIsAstStatus(tmAstInfo(iAst).iStatus, ASTEXTENDED_MG) = False) And (gIsAstStatus(tmAstInfo(iAst).iStatus, ASTEXTENDED_BONUS) = False) And (gIsAstStatus(tmAstInfo(iAst).iStatus, ASTEXTENDED_REPLACEMENT) = False) Then
            Select Case tgStatusTypes(gGetAirStatus(tmAstInfo(iAst).iStatus)).iPledged
                Case 0  'Carried
                Case 1  'Delayed
                Case 2  'Off Air
                    sAirDate = ""
                    sAirTime = ""
                Case 3  'Not Pledged
            End Select
            Select Case tgStatusTypes(gGetAirStatus(tmAstInfo(iAst).iPledgeStatus)).iPledged
                Case 0  'Carried
                    'sPdTime = sPdSTime
                    If Trim$(sPdETime) <> "" Then
                        sPdTime = sPdSTime & "-" & sPdETime
                    Else
                        sPdTime = sPdSTime & "-" & sPdSTime
                    End If
                Case 1  'Delayed
                    'sPdTime = sPdSTime & "-" & sPdETime
                    If Trim$(sPdETime) <> "" Then
                        sPdTime = sPdSTime & "-" & sPdETime
                    Else
                        sPdTime = sPdSTime & "-" & sPdSTime
                    End If
                Case 2  'Off Air
                    sPdTime = ""
                    sPdDays = ""
                Case 3  'Not Pledged
                    sPdDays = ""
                    sPdTime = "None"
            End Select
        'Else
        '    sPdTime = ""
        '    sPdDays = ""
        'End If
        sVefName = ""
        For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            If tgVehicleInfo(iLoop).iCode = tmAstInfo(iAst).iVefCode Then
                sVefName = Trim$(tgVehicleInfo(iLoop).sVehicle)
                Exit For
            End If
        Next iLoop
        sAdvtName = ""
        For iLoop = 0 To UBound(tgAdvtInfo) - 1 Step 1
            If tgAdvtInfo(iLoop).iCode = tmAstInfo(iAst).iAdfCode Then
                sAdvtName = Trim$(tgAdvtInfo(iLoop).sAdvtName)
                Exit For
            End If
        Next iLoop
        'This is where the grid gets loaded for the cp screen
        If llRow + 1 > grdPost.Rows Then
            grdPost.AddItem ""
        End If
        grdPost.Row = llRow
        If sgUstWin(7) <> "I" Then
            For llCol = VEHICLEINDEX To LENINDEX Step 1
                grdPost.Row = llRow
                grdPost.Col = llCol
                grdPost.CellBackColor = LIGHTYELLOW
            Next llCol
        Else
            For llCol = VEHICLEINDEX To PLEDGETIMEINDEX Step 1
                grdPost.Row = llRow
                grdPost.Col = llCol
                grdPost.CellBackColor = LIGHTYELLOW
            Next llCol
            grdPost.Col = LENINDEX
            grdPost.CellBackColor = LIGHTYELLOW
        End If
        If tmAstInfo(iAst).iCPStatus <> 0 Then
            For llCol = VEHICLEINDEX To LENINDEX Step 1
                grdPost.Row = llRow
                grdPost.Col = llCol
                grdPost.CellForeColor = DARKGREEN   'vbGreen
            Next llCol
        End If
        smMGFeedDate = Trim$(tmAstInfo(iAst).sFeedDate)
        smMGFeedTime = Trim$(tmAstInfo(iAst).sFeedTime)
        If igTimes = 0 Then
            grdPost.TextMatrix(llRow, 0) = Trim$(sVefName)
            grdPost.TextMatrix(llRow, 1) = Trim$(sAdvtName)
            grdPost.TextMatrix(llRow, 2) = Trim$(tmAstInfo(iAst).sFeedDate)
            grdPost.TextMatrix(llRow, 3) = Trim$(tmAstInfo(iAst).sFeedTime)
        Else
            grdPost.TextMatrix(llRow, 0) = Trim$(sVefName)
            grdPost.TextMatrix(llRow, 1) = Trim$(tmAstInfo(iAst).sFeedDate)
            grdPost.TextMatrix(llRow, 2) = Trim$(tmAstInfo(iAst).sFeedTime)
            grdPost.TextMatrix(llRow, 3) = Trim$(sAdvtName)
        End If
        grdPost.TextMatrix(llRow, PLEDGEDAYSINDEX) = Trim$(sPdDays)
        grdPost.TextMatrix(llRow, PLEDGETIMEINDEX) = Trim$(sPdTime)
        grdPost.TextMatrix(llRow, AIRDATEINDEX) = Trim$(sAirDate)
        grdPost.TextMatrix(llRow, AIRTIMEINDEX) = Trim$(sAirTime)
        grdPost.TextMatrix(llRow, LENINDEX) = Trim$(tmAstInfo(iAst).iLen)
        grdPost.TextMatrix(llRow, STATUSINDEX) = Trim$(sStatus)
        sProd = ""
        sCart = ""
        sISCI = ""
        slInfo = ""
        lCrfCsfCode = 0
        'Dan 7639
        Dim blImportedISCI As Boolean
        blImportedISCI = gIsISCIChanged(tmAstInfo(iAst).iStatus)
        If blImportedISCI = False Then
            SQLQuery = "SELECT lstProd, lstCart, lstISCI, lstCrfCsfCode"
            SQLQuery = SQLQuery & " FROM LST"
            SQLQuery = SQLQuery & " WHERE lstCode = " & tmAstInfo(iAst).lLstCode
            Set rst_Lst = gSQLSelectCall(SQLQuery)
            If Not rst_Lst.EOF Then
                If IsNull(rst_Lst!lstProd) = True Then
                    sProd = ""
                Else
                    sProd = Trim$(rst_Lst!lstProd)
                End If
                
                If IsNull(rst_Lst!lstCart) Or Left$(rst_Lst!lstCart, 1) = Chr$(0) Then
                    sCart = ""
                Else
                    sCart = Trim$(rst_Lst!lstCart)
                End If
                
                If IsNull(rst_Lst!lstISCI) = True Then
                    sISCI = ""
                Else
                    sISCI = Trim$(rst_Lst!lstISCI)
                End If
                
                'llRet = gBinarySearchCpf(rst_lst!lstCpfCode)
                'If llRet <> -1 Then
                '    sCreative = Trim$(tgCpfInfo(llRet).sCreative)
                'Else
                '    sCreative = ""
                'End If
                lCrfCsfCode = rst_Lst!lstCrfCsfCode
            End If
            ''iRet = gGetRegionCopy(tmAstInfo(iAst).iShttCode, tmAstInfo(iAst).lSdfCode, tmAstInfo(iAst).iVefCode, sRCart, sRProd, sRISCI, sRCreative, lRCrfCsfCode, lRCrfCode)
            'iRet = gGetRegionCopy(tmAstInfo(iAst), sRCart, sRProd, sRISCI, sRCreative, lRCrfCsfCode, lRCrfCode)
            'If iRet Then
            If tmAstInfo(iAst).iRegionType > 0 Then
                sCart = Trim$(tmAstInfo(iAst).sRCart)  'sRCart
                sProd = Trim$(tmAstInfo(iAst).sRProduct)   'sRProd
                sISCI = Trim$(tmAstInfo(iAst).sRISCI)  'sRISCI
                lCrfCsfCode = tmAstInfo(iAst).lRCrfCsfCode  'lRCrfCsfCode
            End If
        Else
            sISCI = Trim$(tmAstInfo(iAst).sISCI)
            sProd = Trim$(tmAstInfo(iAst).sProd)
        End If
        grdPost.TextMatrix(llRow, CARTISCIINDEX) = sCart & " " & sISCI
        
        If igTimes = 0 Then
            grdPost.TextMatrix(llRow, 1) = Trim$(sAdvtName) & "/" & sProd
        Else
            grdPost.TextMatrix(llRow, 3) = Trim$(sAdvtName) & "/" & sProd
        End If
        
        'D.S. 03/12/12 build missed here  Mg date time reason, missed
        'tmAstInfo(iAst).
        
        grdPost.TextMatrix(llRow, INFOINDEX) = mGetMGAndMissedInfo(iAst)

        If gIsAstStatus(tmAstInfo(iAst).iStatus, ASTEXTENDED_MG) Or gIsAstStatus(tmAstInfo(iAst).iStatus, ASTEXTENDED_REPLACEMENT) Or gIsAstStatus(tmAstInfo(iAst).iStatus, ASTEXTENDED_BONUS) Then
            If igTimes = 0 Then
                grdPost.TextMatrix(llRow, 2) = smMGFeedDate
                grdPost.TextMatrix(llRow, 3) = smMGFeedTime
            Else
                grdPost.TextMatrix(llRow, 1) = smMGFeedDate
                grdPost.TextMatrix(llRow, 2) = smMGFeedTime
            End If
        End If

        grdPost.TextMatrix(llRow, ASTINDEX) = iAst
        llRow = llRow + 1
    Next iAst
    'Paint remaining rows if in view mode
    If sgUstWin(7) <> "I" Then
        For llSet = llRow To grdPost.Rows - 1 Step 1
            For llCol = VEHICLEINDEX To LENINDEX Step 1
                grdPost.Row = llSet
                grdPost.Col = llCol
                grdPost.CellBackColor = LIGHTYELLOW
            Next llCol
        Next llSet
    End If
    'Don't add extra row
'    If llRow >= grdPost.Rows Then
'        grdPost.AddItem ""
'    End If
    grdPost.Redraw = True
End Sub


Private Sub grdPost_Click()
    Dim llRow As Long
    Dim llCol As Long
    
    If sgUstWin(7) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If UBound(tmAstInfo) <= LBound(tmAstInfo) Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If (grdPost.Col < STATUSINDEX) Or (grdPost.Col = LENINDEX) Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
'    If Not mPostColAllowed(grdPost.Col) Then
'        pbcClickFocus.SetFocus
'        Exit Sub
'    End If
'    If grdPost.Col > 9 Then
'        Exit Sub
'    End If
'    lmTopRow = grdPost.TopRow
'    llRow = grdPost.Row
'    mPostEnableBox
End Sub

Private Sub grdPost_EnterCell()
    If UBound(tmAstInfo) <= LBound(tmAstInfo) Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    mPostSetShow
    If sgUstWin(7) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
End Sub

Private Sub grdPost_GotFocus()
    If UBound(tmAstInfo) <= LBound(tmAstInfo) Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdPost.Col >= grdPost.Cols - 1 Then
        Exit Sub
    End If
    'grdPost_Click
End Sub

Private Sub grdPost_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdPost.TopRow
    grdPost.Redraw = False
End Sub

Private Sub grdPost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    
    'grdPost.ToolTipText = ""
    If (Y > grdPost.RowHeight(0)) And (Y < grdPost.RowHeight(0)) Then
        grdPost.ToolTipText = ""
        Exit Sub
    End If
    'ilFound = gGrid_GetRowCol(grdPost, X, Y, llRow, llCol)
    'If ilFound Then
    '    grdPost.ToolTipText = Trim$(grdPost.TextMatrix(llRow, CARTISCIINDEX))
    'End If
        
    If grdPost.MouseCol = STATUSINDEX Then
        grdPost.ToolTipText = Trim$(grdPost.TextMatrix(grdPost.MouseRow, INFOINDEX))
    ElseIf (grdPost.MouseCol = AIRDATEINDEX) Or (grdPost.MouseCol = AIRTIMEINDEX) Then
        grdPost.ToolTipText = Trim$(grdPost.TextMatrix(grdPost.MouseRow, CARTISCIINDEX))
    Else
        grdPost.ToolTipText = ""
    End If
            
        
        

End Sub

Private Sub grdPost_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilFound As Integer
    
    If Y < grdPost.RowHeight(0) Then
        gSetMousePointer grdPost, grdPost, vbHourglass
        grdPost.Col = grdPost.MouseCol
        mPostSortCol grdPost.Col
        grdPost.Row = 0
        grdPost.Col = ASTINDEX
        gSetMousePointer grdPost, grdPost, vbDefault
        Exit Sub
    End If
    If sgUstWin(7) <> "I" Then
        grdPost.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If UBound(tmAstInfo) <= LBound(tmAstInfo) Then
        grdPost.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdPost, X, Y)
    If Not ilFound Then
        grdPost.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdPost.Row - 1 >= UBound(tmAstInfo) Then
        grdPost.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If (grdPost.Col < STATUSINDEX) Or (grdPost.Col = LENINDEX) Then
        grdPost.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If Not mPostColAllowed(grdPost.Col) Then
        grdPost.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdPost.Col > AIRTIMEINDEX Then
        grdPost.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    lmTopRow = grdPost.TopRow
    grdPost.Redraw = True
    mPostEnableBox
End Sub

Private Sub grdPost_Scroll()
'    If (lmTopRow <> -1) And (lmTopRow <> grdPost.TopRow) Then
'        grdPost.TopRow = lmTopRow
'        lmTopRow = -1
'    End If
    If grdPost.Redraw = False Then
        grdPost.Redraw = True
        grdPost.TopRow = lmTopRow
        grdPost.Refresh
        grdPost.Redraw = False
    End If
    If (imShowGridBox) And (grdPost.Row >= grdPost.FixedRows) And (grdPost.Col >= 0) And (grdPost.Col < grdPost.Cols - 1) Then
        If grdPost.RowIsVisible(grdPost.Row) Then
            pbcArrow.Move grdPost.Left - pbcArrow.Width, grdPost.Top + grdPost.RowPos(grdPost.Row) + (grdPost.RowHeight(grdPost.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
           If grdPost.Col = AIRDATEINDEX Then  'Date
                txtDropdown.Move grdPost.Left + grdPost.ColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - cmcDropDown.Width - 30, grdPost.RowHeight(grdPost.Row) - 15
                txtDropdown.Visible = True
                txtDropdown.SetFocus
           ElseIf grdPost.Col = AIRTIMEINDEX Then  'Time
                txtDropdown.Move grdPost.Left + grdPost.ColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - cmcDropDown.Width - 30, grdPost.RowHeight(grdPost.Row) - 15
                txtDropdown.Visible = True
                txtDropdown.SetFocus
            ElseIf grdPost.Col = STATUSINDEX Then
                txtDropdown.Move grdPost.Left + grdPost.ColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - cmcDropDown.Width - 30, grdPost.RowHeight(grdPost.Row) - 15
                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
                lbcStatus.Move txtDropdown.Left, txtDropdown.Top + txtDropdown.Height, txtDropdown.Width + (7 * txtDropdown.Width) \ 2
                txtDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcStatus.Visible = True
                txtDropdown.SetFocus
            End If
        Else
            pbcPostFocus.SetFocus
            txtDropdown.Visible = False
            lbcStatus.Visible = False
            cmcDropDown.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcPostFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
    End If
End Sub

Private Sub imcKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtKey.Visible = True
    txtKey.ZOrder
    DoEvents
End Sub

Private Sub imcKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtKey.Visible = False
End Sub

Private Sub lbcStatus_Click()
    txtDropdown.Text = lbcStatus.List(lbcStatus.ListIndex)
    If (txtDropdown.Visible) And (txtDropdown.Enabled) Then
        txtDropdown.SetFocus
        lbcStatus.Visible = False
    End If
End Sub

Private Sub optStatus_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub optStatus_GotFocus(Index As Integer)
    mPostSetShow
End Sub

Private Sub pbcPostSTab_GotFocus()
    Dim slStr As String
    Dim llRowIndex As Long
    Dim iIndex As Integer
    
    If GetFocus() <> pbcPostSTab.hwnd Then
        imArrowKeyPressed = 0
        Exit Sub
    End If
    If sgUstWin(7) <> "I" Then
        pbcClickFocus.SetFocus
        imArrowKeyPressed = 0
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mPostEnableBox
        imArrowKeyPressed = 0
        Exit Sub
    End If
    If txtDropdown.Visible Then
        mPostSetShow
        If grdPost.Col = AIRDATEINDEX Then
            slStr = Trim$(txtDropdown.Text)
            If slStr = "" Then
                slStr = grdPost.TextMatrix(grdPost.Row, STATUSINDEX)
                llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                If llRowIndex >= 0 Then
                    iIndex = lbcStatus.ItemData(llRowIndex)
                    If tgStatusTypes(iIndex).iPledged <> 2 Then
                        Beep
                        grdPost.Col = grdPost.Col
                        mPostEnableBox
                        imArrowKeyPressed = 0
                        Exit Sub
                    End If
                End If
            Else
                If (Not gIsDate(slStr)) Then
                    Beep
                    grdPost.Col = grdPost.Col
                    mPostEnableBox
                    imArrowKeyPressed = 0
                    Exit Sub
                End If
            End If
            If imArrowKeyPressed = KEYUP Then
                If grdPost.Row > grdPost.FixedRows Then
                    lmTopRow = -1
                    grdPost.Row = grdPost.Row - 1
                    If Not grdPost.RowIsVisible(grdPost.Row) Then
                        grdPost.TopRow = grdPost.TopRow - 1
                    End If
                    grdPost.Col = AIRDATEINDEX
                    slStr = grdPost.TextMatrix(grdPost.Row, STATUSINDEX)
                    llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                    If llRowIndex >= 0 Then
                        iIndex = lbcStatus.ItemData(llRowIndex)
                        If tgStatusTypes(iIndex).iPledged = 2 Then
                            grdPost.Col = STATUSINDEX
                        End If
                    End If
                    mPostEnableBox
                Else
                    pbcClickFocus.SetFocus
                End If
            Else
                grdPost.Col = grdPost.Col - 1
                mPostEnableBox
            End If
        ElseIf grdPost.Col = AIRTIMEINDEX Then
            slStr = Trim$(txtDropdown.Text)
            If slStr = "" Then
                slStr = grdPost.TextMatrix(grdPost.Row, STATUSINDEX)
                llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                If llRowIndex >= 0 Then
                    iIndex = lbcStatus.ItemData(llRowIndex)
                    If tgStatusTypes(iIndex).iPledged <> 2 Then
                        Beep
                        grdPost.Col = grdPost.Col
                        mPostEnableBox
                        imArrowKeyPressed = 0
                        Exit Sub
                    End If
                End If
            Else
                If (Not gIsTime(slStr)) Then
                    Beep
                    grdPost.Col = grdPost.Col
                    mPostEnableBox
                    imArrowKeyPressed = 0
                    Exit Sub
                End If
            End If
'            grdPost.Col = grdPost.Col - 1
'            mPostEnableBox
            If imArrowKeyPressed <> KEYLEFT Then
                If grdPost.Row > grdPost.FixedRows Then
                    lmTopRow = -1
                    grdPost.Row = grdPost.Row - 1
                    If Not grdPost.RowIsVisible(grdPost.Row) Then
                        grdPost.TopRow = grdPost.TopRow - 1
                    End If
                    grdPost.Col = AIRTIMEINDEX
                    slStr = grdPost.TextMatrix(grdPost.Row, STATUSINDEX)
                    llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                    If llRowIndex >= 0 Then
                        iIndex = lbcStatus.ItemData(llRowIndex)
                        If tgStatusTypes(iIndex).iPledged = 2 Then
                            grdPost.Col = STATUSINDEX
                        End If
                    End If
                    mPostEnableBox
                Else
                    pbcClickFocus.SetFocus
                End If
            Else
                grdPost.Col = grdPost.Col - 1
                mPostEnableBox
            End If
        ElseIf grdPost.Col = STATUSINDEX Then
            If grdPost.Row > grdPost.FixedRows Then
                lmTopRow = -1
                grdPost.Row = grdPost.Row - 1
                If Not grdPost.RowIsVisible(grdPost.Row) Then
                    grdPost.TopRow = grdPost.TopRow - 1
                End If
                grdPost.Col = AIRTIMEINDEX
                slStr = grdPost.TextMatrix(grdPost.Row, STATUSINDEX)
                llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                If llRowIndex >= 0 Then
                    iIndex = lbcStatus.ItemData(llRowIndex)
                    If tgStatusTypes(iIndex).iPledged = 2 Then
                        grdPost.Col = STATUSINDEX
                    End If
                End If
                mPostEnableBox
            Else
                pbcClickFocus.SetFocus
            End If
        Else
            grdPost.Col = grdPost.Col - 1
            mPostEnableBox
        End If
    Else
        lmTopRow = -1
        grdPost.TopRow = grdPost.FixedRows
        grdPost.Col = STATUSINDEX
        grdPost.Row = grdPost.FixedRows
        mPostEnableBox
    End If
    imArrowKeyPressed = 0
End Sub

Private Sub pbcPostTab_GotFocus()
    Dim slStr As String
    Dim llRowIndex As Long
    Dim iIndex As Integer
    Dim llRow As Long
    
    If GetFocus() <> pbcPostTab.hwnd Then
        imArrowKeyPressed = 0
        Exit Sub
    End If
    If sgUstWin(7) <> "I" Then
        pbcClickFocus.SetFocus
        imArrowKeyPressed = 0
        Exit Sub
    End If
    If txtDropdown.Visible Then
        mPostSetShow
        If grdPost.Col = STATUSINDEX Then
'            If grdPost.Row + 1 < grdPost.Rows Then
            slStr = grdPost.TextMatrix(grdPost.Row, STATUSINDEX)
            llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
            If llRowIndex >= 0 Then
                iIndex = lbcStatus.ItemData(llRowIndex)
                If tgStatusTypes(iIndex).iPledged = 2 Then
                    llRow = grdPost.Rows
                    Do
                        llRow = llRow - 1
                    Loop While grdPost.TextMatrix(llRow, STATUSINDEX) = ""
                    llRow = llRow + 1
                    If (grdPost.Row + 1 < llRow) Then
                        lmTopRow = -1
                        grdPost.Row = grdPost.Row + 1
                        If Not grdPost.RowIsVisible(grdPost.Row) Then
                            grdPost.TopRow = grdPost.TopRow + 1
                        End If
                        grdPost.Col = STATUSINDEX
                        mPostEnableBox
                        imArrowKeyPressed = 0
                        Exit Sub
                    Else
                        pbcClickFocus.SetFocus
                    End If
                End If
            End If
            grdPost.Col = grdPost.Col + 1
            mPostEnableBox
        ElseIf grdPost.Col = AIRDATEINDEX Then
            slStr = Trim$(txtDropdown.Text)
            If slStr = "" Then
                slStr = grdPost.TextMatrix(grdPost.Row, STATUSINDEX)
                llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                If llRowIndex >= 0 Then
                    iIndex = lbcStatus.ItemData(llRowIndex)
                    If tgStatusTypes(iIndex).iPledged <> 2 Then
                        Beep
                        grdPost.Col = grdPost.Col
                        mPostEnableBox
                        imArrowKeyPressed = 0
                        Exit Sub
                    End If
                End If
                llRow = grdPost.Rows
                Do
                    llRow = llRow - 1
                Loop While grdPost.TextMatrix(llRow, STATUSINDEX) = ""
                llRow = llRow + 1
                If (grdPost.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdPost.Row = grdPost.Row + 1
                    If Not grdPost.RowIsVisible(grdPost.Row) Then
                        grdPost.TopRow = grdPost.TopRow + 1
                    End If
                    grdPost.Col = STATUSINDEX
                    mPostEnableBox
                    imArrowKeyPressed = 0
                    Exit Sub
                Else
                    pbcClickFocus.SetFocus
                    imArrowKeyPressed = 0
                    Exit Sub
                End If
            Else
                If Not gIsDate(slStr) Then
                    Beep
                    grdPost.Col = grdPost.Col
                    mPostEnableBox
                    imArrowKeyPressed = 0
                    Exit Sub
                End If
            End If
            If imArrowKeyPressed = KEYDOWN Then
                llRow = grdPost.Rows
                Do
                    llRow = llRow - 1
                Loop While grdPost.TextMatrix(llRow, STATUSINDEX) = ""
                llRow = llRow + 1
                If (grdPost.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdPost.Row = grdPost.Row + 1
                    If Not grdPost.RowIsVisible(grdPost.Row) Then
                        grdPost.TopRow = grdPost.TopRow + 1
                    End If
                    grdPost.Col = STATUSINDEX
                    slStr = grdPost.TextMatrix(grdPost.Row, STATUSINDEX)
                    llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                    If llRowIndex >= 0 Then
                        iIndex = lbcStatus.ItemData(llRowIndex)
                        If tgStatusTypes(iIndex).iPledged <> 2 Then
                            grdPost.Col = AIRDATEINDEX
                        End If
                    End If
                    mPostEnableBox
                Else
                    pbcClickFocus.SetFocus
                End If
            Else
                grdPost.Col = grdPost.Col + 1
                mPostEnableBox
            End If
        ElseIf grdPost.Col = AIRTIMEINDEX Then
            slStr = Trim$(txtDropdown.Text)
            If slStr = "" Then
                slStr = grdPost.TextMatrix(grdPost.Row, STATUSINDEX)
                llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                If llRowIndex >= 0 Then
                    iIndex = lbcStatus.ItemData(llRowIndex)
                    If tgStatusTypes(iIndex).iPledged <> 2 Then
                        Beep
                        grdPost.Col = grdPost.Col
                        mPostEnableBox
                        imArrowKeyPressed = 0
                        Exit Sub
                    End If
                End If
            Else
                If Not gIsTime(slStr) Then
                    Beep
                    grdPost.Col = grdPost.Col
                    mPostEnableBox
                    imArrowKeyPressed = 0
                    Exit Sub
                End If
            End If
            llRow = grdPost.Rows
            Do
                llRow = llRow - 1
            Loop While grdPost.TextMatrix(llRow, STATUSINDEX) = ""
            llRow = llRow + 1
            If (grdPost.Row + 1 < llRow) Then
                lmTopRow = -1
                grdPost.Row = grdPost.Row + 1
                If Not grdPost.RowIsVisible(grdPost.Row) Then
                    grdPost.TopRow = grdPost.TopRow + 1
                End If
                grdPost.Col = STATUSINDEX
                If imArrowKeyPressed <> KEYRIGHT Then
                    slStr = grdPost.TextMatrix(grdPost.Row, STATUSINDEX)
                    llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                    If llRowIndex >= 0 Then
                        iIndex = lbcStatus.ItemData(llRowIndex)
                        If tgStatusTypes(iIndex).iPledged <> 2 Then
                            grdPost.Col = AIRTIMEINDEX
                        End If
                    End If
                End If
                mPostEnableBox
            Else
                pbcClickFocus.SetFocus
            End If
        Else
            grdPost.Col = grdPost.Col + 1
            mPostEnableBox
        End If
    Else
        lmTopRow = -1
        grdPost.TopRow = grdPost.FixedRows
        grdPost.Col = STATUSINDEX
        grdPost.Row = grdPost.FixedRows
        mPostEnableBox
    End If
    imArrowKeyPressed = 0
End Sub

Private Sub txtDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As Integer
    
    slStr = txtDropdown.Text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    Select Case grdPost.Col
        Case AIRDATEINDEX
            'Removed to mPostSetShow so that status can to set
            slStr = Trim$(txtDropdown.Text)
            If (gIsDate(slStr)) And (slStr <> "") Then
                If grdPost.CellForeColor <> DARKGREEN Then
                    grdPost.CellForeColor = vbBlack
                End If
'                If grdPost.Text <> slStr Then
'                    imFieldChgd = True
'                End If
'                grdPost.Text = slStr
            End If
        Case AIRTIMEINDEX
            'Moved to mPostSetShow so that status can be set
            slStr = Trim$(txtDropdown.Text)
            If (gIsTime(slStr)) And (slStr <> "") Then
                If grdPost.CellForeColor <> DARKGREEN Then
                    grdPost.CellForeColor = vbBlack
                End If
'                slStr = gConvertTime(slStr)
'                If Second(slStr) = 0 Then
'                    slStr = Format$(slStr, sgShowTimeWOSecForm)
'                Else
'                    slStr = Format$(slStr, sgShowTimeWSecForm)
'                End If
'                If grdPost.Text <> slStr Then
'                    imFieldChgd = True
'                End If
'                grdPost.Text = slStr
            End If
        Case STATUSINDEX
            llRow = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
            If llRow >= 0 Then
                lbcStatus.ListIndex = llRow
                txtDropdown.Text = lbcStatus.List(lbcStatus.ListIndex)
                txtDropdown.SelStart = ilLen
                txtDropdown.SelLength = Len(txtDropdown.Text)
            End If
    End Select
End Sub

Private Sub txtDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub txtDropdown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If txtDropdown.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub txtDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case grdPost.Col
            Case AIRDATEINDEX, AIRTIMEINDEX
                If (KeyCode = KEYUP) Then
                    imArrowKeyPressed = KEYUP
                    pbcPostSTab.SetFocus
                End If
                If (KeyCode = KEYDOWN) Then
                    imArrowKeyPressed = KEYDOWN
                    pbcPostTab.SetFocus
                End If
            Case STATUSINDEX
                gProcessArrowKey Shift, KeyCode, lbcStatus, True
        End Select
    End If
    If (KeyCode = KEYLEFT) Then
        imArrowKeyPressed = KEYLEFT
        pbcPostSTab.SetFocus
    End If
    If (KeyCode = KEYRIGHT) Then
        imArrowKeyPressed = KEYRIGHT
        pbcPostTab.SetFocus
    End If
End Sub

Private Sub mSetRowToUnposted()
    Dim llRow As Long
    Dim slStr As String
    Dim llCol As Long
    
    grdPost.Redraw = False
    For llRow = grdPost.FixedRows To grdPost.Rows - 1 Step 1
        slStr = Trim$(grdPost.TextMatrix(llRow, STATUSINDEX))
        If slStr <> "" Then
            For llCol = VEHICLEINDEX To LENINDEX Step 1
                grdPost.Row = llRow
                grdPost.Col = llCol
                grdPost.CellForeColor = vbBlack
            Next llCol
        End If
    Next llRow
    grdPost.Redraw = True
End Sub

Private Sub mSetASTAndCPTTStatus()
    
    Dim ilLoop As Integer
    Dim sFWkDate As String
    Dim sLWkDate As String
    Dim ilPostingStatus As Integer
    Dim ilStatus As Integer
    Dim ilAst As Integer
    Dim ilSpotsAired As Integer
    Dim rst_Ast As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If optStatus(0).Value Then
        ilPostingStatus = 2 'Completed
        ilStatus = 1
    ElseIf optStatus(1).Value Then
        ilPostingStatus = 1 'Partially posted
        ilStatus = 0
    Else
        ilPostingStatus = 0 'Not Posted
        ilStatus = 0
    End If
    '9/29/14: always reset compliant info
    'If (igCPStatus = ilStatus) And (igCPPostingStatus = ilPostingStatus) Then
    '    Exit Sub
    'End If
    ReDim tlCPPosting(0 To UBound(tgCPPosting)) As CPPOSTING
    For ilLoop = 0 To UBound(tgCPPosting) - 1 Step 1
        tlCPPosting(ilLoop) = tgCPPosting(ilLoop)
    Next ilLoop
    For ilLoop = 0 To UBound(tlCPPosting) - 1 Step 1
        lmAttCode = tlCPPosting(ilLoop).lAttCode
        sFWkDate = Format$(gObtainPrevMonday(tlCPPosting(ilLoop).sDate), sgShowDateForm)
        If igTimes = 0 Then
            sLWkDate = Format$(gObtainEndStd(tlCPPosting(ilLoop).sDate), sgShowDateForm)
        Else
            sLWkDate = Format$(gObtainNextSunday(tlCPPosting(ilLoop).sDate), sgShowDateForm)
        End If
        
        If lgSelGameGsfCode <= 0 Then
            If ilPostingStatus = 2 Then
                SQLQuery = "UPDATE ast SET "
                'If optStatus(0).Value Then
                    SQLQuery = SQLQuery + "astCPStatus = " & "1"
                'Else
                '    SQLQuery = SQLQuery + "astCPStatus = " & "2"
                'End If
                '10/19/18: added setting user
                SQLQuery = SQLQuery + ", " & "astUstCode = " & igUstCode
                SQLQuery = SQLQuery + " WHERE (astAtfCode = " & tlCPPosting(ilLoop).lAttCode
                SQLQuery = SQLQuery + " AND astCPStatus = 0"
                SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')" & ")"
                If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                    'gSetMousePointer grdPost, grdPost, vbDefault
                    'Unload frmDateTimes
                    'Exit Sub
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    gSetMousePointer grdPost, grdPost, vbDefault
                    gHandleError "AffErrorLog.txt", "CPDateTimes-mSetASTAndCPTTStatus"
                    Exit Sub
                End If
            ElseIf ilPostingStatus = 0 Then
                '12/13/13: PledgeStatus now in DAT
                'SQLQuery = "UPDATE ast SET "
                ''If optStatus(0).Value Then
                '    SQLQuery = SQLQuery + "astCPStatus = " & "0" & ", "
                ''Else
                ''    SQLQuery = SQLQuery + "astCPStatus = " & "2"
                ''End If
                'SQLQuery = SQLQuery + "astStatus = astPledgeStatus"
                'SQLQuery = SQLQuery + " WHERE (astAtfCode = " & tlCPPosting(ilLoop).lAttCode
                'SQLQuery = SQLQuery + " AND astCPStatus = 1"
                'SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')" & ")"
                'If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                '    'gSetMousePointer grdPost, grdPost, vbDefault
                '    'Unload frmDateTimes
                '    'Exit Sub
                '    GoSub ErrHand
                'End If
                mSetStatusToPledgeStatus tlCPPosting(ilLoop).lAttCode, 0, sFWkDate, sLWkDate
                mSetRowToUnposted
            End If
        Else
            'Change each game individually unless setup special link to lst Update ast set astCPStatus = 1 where field from (select ast, lst where lstcode = astlstcode and ........)
            If (ilPostingStatus = 2) Or (ilPostingStatus = 0) Then
                For ilAst = 0 To UBound(tmAstInfo) - 1 Step 1
                    SQLQuery = "UPDATE ast SET "
                    If ilPostingStatus = 2 Then
                        SQLQuery = SQLQuery + "astCPStatus = " & "1"
                    ElseIf ilPostingStatus = 0 Then
                        '10/19/18" added comma
                        'SQLQuery = SQLQuery + "astCPStatus = " & "0"
                        SQLQuery = SQLQuery + "astCPStatus = " & "0" & ", "
                        '12/13/13: Pledge now part of DAT
                        'SQLQuery = SQLQuery + "astStatus = astPledgeStatus"
                        SQLQuery = SQLQuery + "astStatus = " & tmAstInfo(ilAst).iPledgeStatus
                    End If
                    '10/19/18: added setting user
                    SQLQuery = SQLQuery + ", " & "astUstCode = " & igUstCode
                    SQLQuery = SQLQuery + " WHERE (astCode = " & tmAstInfo(ilAst).lCode & ")"
                    If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                        'gSetMousePointer grdPost, grdPost, vbDefault
                        'Unload frmDateTimes
                        'Exit Sub
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        gSetMousePointer grdPost, grdPost, vbDefault
                        gHandleError "AffErrorLog.txt", "CPDateTimes-mSetASTAndCPTTStatus"
                        Exit Sub
                    End If
                Next ilAst
                If ilPostingStatus = 0 Then
                    mSetRowToUnposted
                End If
            End If
        End If
        
        If lgSelGameGsfCode > 0 Then
            If DateValue(sLWkDate) < DateValue(Format$(gNow(), "m/d/yy")) Then
                If ilPostingStatus = 2 Then
                    SQLQuery = "Select astCode FROM ast WHERE astCPStatus = 0"
                    SQLQuery = SQLQuery + " AND astAtfCode = " & lmAttCode
                    SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')"
                    Set rst_Ast = gSQLSelectCall(SQLQuery)
                    If Not rst_Ast.EOF Then
                        ilPostingStatus = 1 'Partially posted
                        ilStatus = 0
                    End If
                ElseIf ilPostingStatus = 0 Then
                    SQLQuery = "Select astCode FROM ast WHERE astCPStatus = 1"
                    SQLQuery = SQLQuery + " AND astAtfCode = " & lmAttCode
                    SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')"
                    Set rst_Ast = gSQLSelectCall(SQLQuery)
                    If Not rst_Ast.EOF Then
                        ilPostingStatus = 1 'Partially posted
                        ilStatus = 0
                    End If
                End If
            Else
                If ilPostingStatus = 2 Then
                    ilPostingStatus = 1 'Partially posted
                    ilStatus = 0
                End If
            End If
        End If
        
        If ilStatus <> 0 Then
            ilSpotsAired = gDidAnySpotsAir(lmAttCode, sFWkDate, sLWkDate)
            If ilSpotsAired Then
                'We know at least one spot aired
               ilStatus = 1
            Else
                'no aired spots were found
                ilStatus = 2
            End If
        End If

        SQLQuery = "UPDATE cptt SET "
        SQLQuery = SQLQuery + "cpttStatus = " & ilStatus & ", "
        SQLQuery = SQLQuery + "cpttPostingStatus = " & ilPostingStatus
        '10/19/18: added setting user
        SQLQuery = SQLQuery + ", " & "cpttUsfCode = " & igUstCode
        SQLQuery = SQLQuery + " WHERE (cpttVefCode = " & tlCPPosting(ilLoop).iVefCode
        SQLQuery = SQLQuery + " AND cpttShfCode = " & tlCPPosting(ilLoop).iShttCode
        SQLQuery = SQLQuery + " AND (cpttStartDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')" & ")"
        If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
            'gSetMousePointer grdPost, grdPost, vbDefault
            'Unload frmDateTimes
            'Exit Sub
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            gSetMousePointer grdPost, grdPost, vbDefault
            gHandleError "AffErrorLog.txt", "CPDateTimes-mSetASTAndCPTTStatus"
            Exit Sub
        End If
        gSetCpttCount tlCPPosting(ilLoop).lAttCode, sFWkDate, sLWkDate
    Next ilLoop
    gFileChgdUpdate "cptt.mkd", True
    igCPStatus = ilStatus
    igCPPostingStatus = ilPostingStatus
    Exit Sub
ErrHand:
    gSetMousePointer grdPost, grdPost, vbDefault
    gHandleError "AffErrorLog.txt", "frmDateTime-mSetASTAndCPTTStatus"
    Unload frmDateTimes
End Sub

Private Sub mPostSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
    For llRow = grdPost.FixedRows To grdPost.Rows - 1 Step 1
        slStr = Trim$(grdPost.TextMatrix(llRow, STATUSINDEX))
        If slStr <> "" Then
            If ilCol = imDateCol Then
                slSort = Trim$(Str$(DateValue(grdPost.TextMatrix(llRow, imDateCol))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = imTimeCol) Then
                slSort = Trim$(Str$(gTimeToLong(grdPost.TextMatrix(llRow, imTimeCol), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = PLEDGETIMEINDEX) Then
                slStr = grdPost.TextMatrix(llRow, PLEDGETIMEINDEX)
                ilPos = InStr(1, slStr, "-", vbTextCompare)
                If ilPos > 0 Then
                    slStr = Left(slStr, ilPos - 1)
                    slSort = Trim$(Str$(gTimeToLong(slStr, False)))
                    Do While Len(slSort) < 6
                        slSort = "0" & slSort
                    Loop
                    slStr = grdPost.TextMatrix(llRow, PLEDGETIMEINDEX)
                    slStr = Mid(slStr, ilPos + 1)
                    slStr = Trim$(Str$(gTimeToLong(slStr, False)))
                    Do While Len(slStr) < 6
                        slStr = "0" & slStr
                    Loop
                    slSort = slSort & slStr
               Else
                    slSort = Trim$(Str$(gTimeToLong(grdPost.TextMatrix(llRow, PLEDGETIMEINDEX), False)))
                    Do While Len(slSort) < 6
                        slSort = "0" & slSort
                    Loop
                End If
            ElseIf (ilCol = PLEDGEDAYSINDEX) Then
                slStr = UCase$(Trim$(grdPost.TextMatrix(llRow, ilCol)))
                Select Case slStr
                    Case "MO"
                        slSort = "1"
                    Case "TU"
                        slSort = "2"
                    Case "WE"
                        slSort = "3"
                    Case "TH"
                        slSort = "4"
                    Case "FR"
                        slSort = "5"
                    Case "SA"
                        slSort = "6"
                    Case "SU"
                        slSort = "7"
                    Case Else
                        slSort = "0"
                End Select
            ElseIf ilCol = AIRDATEINDEX Then
                If Trim$(grdPost.TextMatrix(llRow, AIRDATEINDEX)) = "" Then
                    slSort = Trim$(Str$(DateValue(grdPost.TextMatrix(llRow, imDateCol))))
                    Do While Len(slSort) < 6
                        slSort = "0" & slSort
                    Loop
                Else
                    slSort = Trim$(Str$(DateValue(grdPost.TextMatrix(llRow, AIRDATEINDEX))))
                    Do While Len(slSort) < 6
                        slSort = "0" & slSort
                    Loop
                End If
            ElseIf (ilCol = AIRTIMEINDEX) Then
                If Trim$(grdPost.TextMatrix(llRow, AIRTIMEINDEX)) = "" Then
                    slSort = Trim$(Str$(gTimeToLong(grdPost.TextMatrix(llRow, imTimeCol), False)))
                    Do While Len(slSort) < 6
                        slSort = "0" & slSort
                    Loop
                Else
                    slSort = Trim$(Str$(gTimeToLong(grdPost.TextMatrix(llRow, AIRTIMEINDEX), False)))
                    Do While Len(slSort) < 6
                        slSort = "0" & slSort
                    Loop
                End If
            Else
                slSort = UCase$(Trim$(grdPost.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdPost.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastPostColSorted) Or ((ilCol = imLastPostColSorted) And (imLastPostSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdPost.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdPost.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastPostColSorted Then
        imLastPostColSorted = SORTINDEX
    Else
        imLastPostColSorted = -1
        imLastPostSort = -1
    End If
    gGrid_SortByCol grdPost, STATUSINDEX, SORTINDEX, imLastPostColSorted, imLastPostSort
    imLastPostColSorted = ilCol
End Sub

Private Function mGetMGAndMissedInfo(ilAst As Integer) As String
    Dim slDateTimeMsg As String
    Dim slMissedReason As String
    Dim ilMnfCode As Integer
    Dim slStr As String
    Dim llAdf As Long
    Dim slAdvtName As String
    Dim blCreateMG As Boolean
    Dim slPdDate As String
    Dim slPdTime As String
    
    On Error GoTo ErrHandle
    mGetMGAndMissedInfo = ""
    slDateTimeMsg = ""
    slMissedReason = ""
    ilMnfCode = 0
    If gIsAstStatus(tmAstInfo(ilAst).iStatus, ASTEXTENDED_MG) Or gIsAstStatus(tmAstInfo(ilAst).iStatus, ASTEXTENDED_REPLACEMENT) Then
        'SQLQuery = "SELECT altLinkToAstCode FROM alt WHERE altAstCode = " & tmAstInfo(ilAst).lCode
        'Set rst = gSQLSelectCall(SQLQuery)
        'If Not rst.EOF Then
            'If rst!altLinkToAstCode > 0 Then
                'Get Missed
                'SQLQuery = "SELECT astCode, astAirDate, astAirTime, astLsfCode FROM ast WHERE AstCode = " & rst!altLinkToAstCode
                SQLQuery = "SELECT astCode, astDatCode, astAirDate, astAirTime, astLsfCode, astFeedDate, astFeedTime FROM ast WHERE AstCode = " & tmAstInfo(ilAst).lLkAstCode
                Set rst = gSQLSelectCall(SQLQuery)
                If Not rst.EOF Then
                    smMGFeedDate = Format(rst!astFeedDate, sgShowDateForm)
                    smMGFeedTime = Format(rst!astFeedTime, sgShowTimeWOSecForm)
                    If gIsAstStatus(tmAstInfo(ilAst).iStatus, ASTEXTENDED_MG) Then
                        'Pledge date/time stored into aired date/time
                        'blCreateMG = gDeterminePledgeDateTime(dat_rst, rst!astDatCode, Format(rst!astFeedDate, sgShowDateForm), Format(rst!astFeedTime, sgShowTimeWSecForm), Format(rst!astAirDate, sgShowDateForm), slPdDate, slPdTime)
                        'slDateTimeMsg = "Missed: " & Format(slPdDate, sgShowDateForm) & " " & Format(slPdTime, sgShowTimeWOSecForm)
                        slDateTimeMsg = "Missed: " & Format(rst!astAirDate, sgShowDateForm) & " " & Format(rst!astAirTime, sgShowTimeWOSecForm)
                    ElseIf gIsAstStatus(tmAstInfo(ilAst).iStatus, ASTEXTENDED_REPLACEMENT) Then
                        slAdvtName = ""
                        SQLQuery = "SELECT lstAdfCode FROM lst where lstcode = " & rst!astLsfCode
                        Set lst_rst = gSQLSelectCall(SQLQuery)
                        If Not lst_rst.EOF Then
                            llAdf = gBinarySearchAdf(CLng(lst_rst!lstAdfCode))
                            If llAdf <> -1 Then
                                slAdvtName = " " & Trim$(tgAdvtInfo(llAdf).sAdvtName)
                            End If
                        End If
                        slDateTimeMsg = "Replaced:" & slAdvtName & " " & Format(rst!astAirDate, sgShowDateForm) & " " & Format(rst!astAirTime, sgShowTimeWOSecForm)
                    End If
                    'Get Missed reason reference
                    'SQLQuery = "SELECT altMnfMissed FROM alt where altAstcode = " & rst!astCode
                    'Set rst = gSQLSelectCall(SQLQuery)
                    'If Not rst.EOF Then
                    '    ilMnfCode = rst!altMnfMissed
                    'End If
                End If
            'End If
        'End If
    ElseIf tmAstInfo(ilAst).lLkAstCode > 0 Then
        SQLQuery = "SELECT astCode, astStatus, astAirDate, astAirTime, astLsfCode, astAdfCode FROM ast WHERE AstCode = " & tmAstInfo(ilAst).lLkAstCode
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            If gIsAstStatus(rst!astStatus, ASTEXTENDED_MG) Then
                slDateTimeMsg = "MG: " & Format(rst!astAirDate, sgShowDateForm) & " " & Format(rst!astAirTime, sgShowTimeWOSecForm)
            ElseIf gIsAstStatus(rst!astStatus, ASTEXTENDED_REPLACEMENT) Then
                slAdvtName = ""
                'SQLQuery = "SELECT lstAdfCode FROM lst where lstcode = " & rst!astLsfCode
                'Set lst_rst = gSQLSelectCall(SQLQuery)
                'If Not lst_rst.EOF Then
                    llAdf = gBinarySearchAdf(CLng(rst!astAdfCode))
                    If llAdf <> -1 Then
                        slAdvtName = " " & Trim$(tgAdvtInfo(llAdf).sAdvtName)
                    End If
                'End If
                slDateTimeMsg = "Replacement:" & slAdvtName & " " & Format(rst!astAirDate, sgShowDateForm) & " " & Format(rst!astAirTime, sgShowTimeWOSecForm)
            End If
            'Get Missed reason reference
            'SQLQuery = "SELECT altMnfMissed FROM alt where altAstcode = " & rst!astCode
            'Set rst = gSQLSelectCall(SQLQuery)
            'If Not rst.EOF Then
            '    ilMnfCode = rst!altMnfMissed
            'End If
        End If
    ElseIf tgStatusTypes(gGetAirStatus(tmAstInfo(ilAst).iStatus)).iPledged = 2 Then
        'SQLQuery = "SELECT altLinkToAstCode, altMnfMissed FROM alt WHERE altAstCode = " & tmAstInfo(ilAst).lCode
        'Set rst = gSQLSelectCall(SQLQuery)
        'If Not rst.EOF Then
            'ilMnfCode = rst!altMnfMissed
            'If rst!altLinkToAstCode > 0 Then
                'Get MG or Replacement
                'SQLQuery = "SELECT astCode, astStatus, astAirDate, astAirTime, astLsfCode FROM ast WHERE AstCode = " & rst!altLinkToAstCode
                SQLQuery = "SELECT astCode, astStatus, astAirDate, astAirTime, astLsfCode, astAdfCode FROM ast WHERE AstCode = " & tmAstInfo(ilAst).lLkAstCode
                Set rst = gSQLSelectCall(SQLQuery)
                If Not rst.EOF Then
                    If gIsAstStatus(rst!astStatus, ASTEXTENDED_MG) Then
                        slDateTimeMsg = "MG: " & Format(rst!astAirDate, sgShowDateForm) & " " & Format(rst!astAirTime, sgShowTimeWOSecForm)
                    ElseIf gIsAstStatus(rst!astStatus, ASTEXTENDED_REPLACEMENT) Then
                        slAdvtName = ""
                        'SQLQuery = "SELECT lstAdfCode FROM lst where lstcode = " & rst!astLsfCode
                        'Set lst_rst = gSQLSelectCall(SQLQuery)
                        'If Not lst_rst.EOF Then
                            llAdf = gBinarySearchAdf(CLng(rst!astAdfCode))
                            If llAdf <> -1 Then
                                slAdvtName = " " & Trim$(tgAdvtInfo(llAdf).sAdvtName)
                            End If
                        'End If
                        slDateTimeMsg = "Replacement:" & slAdvtName & " " & Format(rst!astAirDate, sgShowDateForm) & " at " & Format(rst!astAirTime, sgShowTimeWOSecForm)
                    End If
                End If
            'End If
        'End If
    Else
        mGetMGAndMissedInfo = ""
        Exit Function
    End If
    'Get missed reason
    ilMnfCode = tmAstInfo(ilAst).iMissedMnfCode
    If ilMnfCode > 0 Then
        SQLQuery = "select mnfName from MNF_Multi_Names where mnfCode = " & ilMnfCode
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            slMissedReason = rst!mnfName
        End If
    End If
    If slMissedReason = "" Then
        mGetMGAndMissedInfo = slDateTimeMsg
    Else
        If slDateTimeMsg = "" Then
            mGetMGAndMissedInfo = slMissedReason
        Else
            mGetMGAndMissedInfo = slDateTimeMsg & ", " & slMissedReason
        End If
    End If
    Exit Function
ErrHandle:
    gHandleError "AffErrorLog.txt", "CP Date/Time-mGetMGAndMissedInfo"
    'Exit Function
    Resume Next
End Function
Private Sub mSetStatusToPledgeStatus(llAttCode As Long, llGsfCode As Long, slStartDate As String, slEndDate As String)
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
            SQLQuery = SQLQuery + "astCPStatus = " & "0" & ", "
            SQLQuery = SQLQuery + "astStatus = " & tlDatPledgeInfo.iPledgeStatus
            '10/19/18: added setting user
            SQLQuery = SQLQuery + ", " & "astUstCode = " & igUstCode
            SQLQuery = SQLQuery + " WHERE (astCode = " & rst!astCode & ")"
            If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                gSetMousePointer grdPost, grdPost, vbDefault
                gHandleError "AffErrorLog.txt", "CPDateTimes-mSetStatusToPledgeStatus"
                Exit Sub
            End If
        End If
        rst.MoveNext
    Loop
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "Post CP-mSetStatus"
    Exit Sub
'ErrHand1:
'    gHandleError "AffErrorLog.txt", "Post CP-mSetStatus"
'    Return
End Sub

Private Sub mClearMGs(llAstCode As Long)
    Dim slSQLQuery As String
    slSQLQuery = "Select astLkAstCode, astStatus from ast where astCode = " & llAstCode
    Set ast_rst = gSQLSelectCall(slSQLQuery)
    If Not ast_rst.EOF Then
        If ast_rst!astLkAstCode <= 0 Then
            Exit Sub
        End If
        'If ast_rst!astStatus Mod 100 <= 10 Then
        If (ast_rst!astStatus Mod 100 < ASTEXTENDED_MG) Or (ast_rst!astStatus Mod 100 = ASTAIR_MISSED_MG_BYPASS) Then
            Exit Sub
        End If
        slSQLQuery = "DELETE FROM Ast WHERE (astCode = " & ast_rst!astLkAstCode & ")"
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                        gSetMousePointer grdPost, grdPost, vbDefault
                        gHandleError "AffErrorLog.txt", "DateTimes-mSetASTAndCPTTStatus"
        End If
        slSQLQuery = "UPDATE ast SET "
        slSQLQuery = slSQLQuery & "astLkAstCode = 0"
        '10/19/18: added setting user
        slSQLQuery = slSQLQuery + ", " & "astUstCode = " & igUstCode
        slSQLQuery = slSQLQuery + " WHERE (astCode = " & llAstCode & ")"
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            gSetMousePointer grdPost, grdPost, vbDefault
            gHandleError "AffErrorLog.txt", "CPDateTimes-mClearMGs"
            Exit Sub
        End If
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "ReImport Affiliate Spots-mClearMGs"
End Sub

Private Sub mCheckForMG(tlAstInfo As ASTINFO)
    Dim llMGLstCode As Long
    Dim llMGAstCode As Long
    Dim slPdDate As String
    Dim slPdTime As String
    Dim slAirDate As String
    Dim slAirTime As String
    Dim llLstCode As Long
    Dim llAstCode As Long
    Dim slAffidavitSource As String
    
    bmMGCreated = False
    'If (tlAstInfo.iStatus Mod 100 <> 4) And (tlAstInfo.iStatus Mod 100 <> 8) And (tlAstInfo.iStatus Mod 100 <= 10) Then
    If (tlAstInfo.iStatus Mod 100 <= 10) And (tlAstInfo.lLkAstCode <= 0) Then
        slPdDate = tlAstInfo.sPledgeDate
        slPdTime = tlAstInfo.sPledgeStartTime
        slAirDate = tlAstInfo.sAirDate
        slAirTime = tlAstInfo.sAirTime
        llLstCode = tlAstInfo.lLstCode
        llAstCode = tlAstInfo.lCode
        If gObtainPrevMonday(slPdDate) <> gObtainPrevMonday(slAirDate) Then
            'Create the MG spots
            llMGLstCode = mAddMGLst(llLstCode, slAirDate, slAirTime)
            If llMGLstCode > 0 Then
                slAffidavitSource = ""
                llMGAstCode = mAddAstMG(llMGLstCode, tlAstInfo, slAffidavitSource)
                If llMGAstCode > 0 Then
                    tlAstInfo.iStatus = 4
                    tlAstInfo.sAirDate = slPdDate
                    tlAstInfo.sAirTime = slPdTime
                    mChgAstToMissed llMGAstCode, tlAstInfo
                    bmMGCreated = True
                Else
                End If
            Else
            End If
        End If
    End If
End Sub

Private Function mAddMGLst(llLsfCode As Long, slAirDate As String, slAirTime As String) As Long
    Dim llLst As Long
    
    On Error GoTo ErrHand
    SQLQuery = "SELECT * From lst Where lstCode = " & llLsfCode
    Set lst_rst = gSQLSelectCall(SQLQuery)
    If lst_rst.EOF Then
        mAddMGLst = 0
        Exit Function
    End If
    
    SQLQuery = "Insert Into lst ( "
    SQLQuery = SQLQuery & "lstCode, "
    SQLQuery = SQLQuery & "lstType, "
    SQLQuery = SQLQuery & "lstSdfCode, "
    SQLQuery = SQLQuery & "lstCntrNo, "
    SQLQuery = SQLQuery & "lstAdfCode, "
    SQLQuery = SQLQuery & "lstAgfCode, "
    SQLQuery = SQLQuery & "lstProd, "
    SQLQuery = SQLQuery & "lstLineNo, "
    SQLQuery = SQLQuery & "lstLnVefCode, "
    SQLQuery = SQLQuery & "lstStartDate, "
    SQLQuery = SQLQuery & "lstEndDate, "
    SQLQuery = SQLQuery & "lstMon, "
    SQLQuery = SQLQuery & "lstTue, "
    SQLQuery = SQLQuery & "lstWed, "
    SQLQuery = SQLQuery & "lstThu, "
    SQLQuery = SQLQuery & "lstFri, "
    SQLQuery = SQLQuery & "lstSat, "
    SQLQuery = SQLQuery & "lstSun, "
    SQLQuery = SQLQuery & "lstSpotsWk, "
    SQLQuery = SQLQuery & "lstPriceType, "
    SQLQuery = SQLQuery & "lstPrice, "
    SQLQuery = SQLQuery & "lstSpotType, "
    SQLQuery = SQLQuery & "lstLogVefCode, "
    SQLQuery = SQLQuery & "lstLogDate, "
    SQLQuery = SQLQuery & "lstLogTime, "
    SQLQuery = SQLQuery & "lstDemo, "
    SQLQuery = SQLQuery & "lstAud, "
    SQLQuery = SQLQuery & "lstISCI, "
    SQLQuery = SQLQuery & "lstWkNo, "
    SQLQuery = SQLQuery & "lstBreakNo, "
    SQLQuery = SQLQuery & "lstPositionNo, "
    SQLQuery = SQLQuery & "lstSeqNo, "
    SQLQuery = SQLQuery & "lstZone, "
    SQLQuery = SQLQuery & "lstCart, "
    SQLQuery = SQLQuery & "lstCpfCode, "
    SQLQuery = SQLQuery & "lstCrfCsfCode, "
    SQLQuery = SQLQuery & "lstStatus, "
    SQLQuery = SQLQuery & "lstLen, "
    SQLQuery = SQLQuery & "lstUnits, "
    SQLQuery = SQLQuery & "lstCifCode, "
    SQLQuery = SQLQuery & "lstAnfCode, "
    SQLQuery = SQLQuery & "lstEvtIDCefCode, "
    SQLQuery = SQLQuery & "lstSplitNetwork, "
    SQLQuery = SQLQuery & "lstRafCode, "
    SQLQuery = SQLQuery & "lstFsfCode, "
    SQLQuery = SQLQuery & "lstGsfCode, "
    SQLQuery = SQLQuery & "lstImportedSpot, "
    SQLQuery = SQLQuery & "lstBkoutLstCode, "
    SQLQuery = SQLQuery & "lstLnStartTime, "
    SQLQuery = SQLQuery & "lstLnEndTime, "
    SQLQuery = SQLQuery & "lstUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & 2 & ", "    'lstType
    SQLQuery = SQLQuery & lst_rst!lstSdfCode & ", "
    SQLQuery = SQLQuery & lst_rst!lstCntrNo & ", "
    SQLQuery = SQLQuery & lst_rst!lstAdfCode & ", "
    SQLQuery = SQLQuery & lst_rst!lstAgfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(lst_rst!lstProd) & "', "
    SQLQuery = SQLQuery & lst_rst!lstLineNo & ", "
    SQLQuery = SQLQuery & lst_rst!lstLnVefCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(slAirDate, sgSQLDateForm) & "', "      'lstStartDate
    SQLQuery = SQLQuery & "'" & Format$(slAirDate, sgSQLDateForm) & "', "      'lstEndDate
    SQLQuery = SQLQuery & 0 & ", " 'lstMon
    SQLQuery = SQLQuery & 0 & ", " 'lstTue
    SQLQuery = SQLQuery & 0 & ", " 'lstWed
    SQLQuery = SQLQuery & 0 & ", " 'lstThu
    SQLQuery = SQLQuery & 0 & ", " 'lstFri
    SQLQuery = SQLQuery & 0 & ", " 'lstSat
    SQLQuery = SQLQuery & 0 & ", " 'lstSun
    SQLQuery = SQLQuery & 0 & ", " 'lstSpotsWk
    SQLQuery = SQLQuery & lst_rst!lstPriceType & ", "   'lstPriceType
    SQLQuery = SQLQuery & 0 & ", "   'lstPrice
    SQLQuery = SQLQuery & 5 & ", "    'lstSpotType
    SQLQuery = SQLQuery & lst_rst!lstLogVefCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(slAirDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(slAirTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote("0") & "', "  'lstDemo
    SQLQuery = SQLQuery & 0 & ", " 'lstAud
    SQLQuery = SQLQuery & "'" & gFixQuote(lst_rst!lstISCI) & "', "
    SQLQuery = SQLQuery & 0 & ", "    'lstWkNo
    SQLQuery = SQLQuery & 0 & ", " 'lstBreakNo
    SQLQuery = SQLQuery & 0 & ", "  'lstPositionNo
    SQLQuery = SQLQuery & 0 & ", "   'lstSeqNo
    SQLQuery = SQLQuery & "'" & gFixQuote(lst_rst!lstZone) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(lst_rst!lstCart) & "', "
    SQLQuery = SQLQuery & lst_rst!lstCpfCode & ", "
    SQLQuery = SQLQuery & lst_rst!lstCrfCsfCode & ", "
    SQLQuery = SQLQuery & ASTEXTENDED_MG & ", "  'lstStatus
    SQLQuery = SQLQuery & lst_rst!lstLen & ", "
    SQLQuery = SQLQuery & 0 & ", "   'lstUnit
    SQLQuery = SQLQuery & lst_rst!lstCifCode & ", "
    SQLQuery = SQLQuery & 0 & ", " 'lstAnfCode
    SQLQuery = SQLQuery & 0 & ", "    'lstEvtIDCefCode
    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'lstSplitNetwork
    SQLQuery = SQLQuery & 0 & ", "     'lstRafCode
    SQLQuery = SQLQuery & 0 & ", " 'lstFsfCode
    SQLQuery = SQLQuery & 0 & ", " 'lstGsfCode
    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'lstImportedSpot
    SQLQuery = SQLQuery & 0 & ", "    'lstBkoutLstCode
    SQLQuery = SQLQuery & "'" & Format$("12AM", sgSQLTimeForm) & "', "  'lstLnStartTime
    SQLQuery = SQLQuery & "'" & Format$("12AM", sgSQLTimeForm) & "', "    'lstLnEndTime
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    llLst = gInsertAndReturnCode(SQLQuery, "lst", "lstCode", "Replace")
    If llLst > 0 Then
        mAddMGLst = llLst
    Else
        mAddMGLst = 0
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Affiliate Affidavit-mChgAstToMissed"
End Function

Private Function mAddAstMG(llLstCode As Long, tlAstInfo As ASTINFO, slAffidavitSource As String) As Long
    Dim llAst As Long

    On Error GoTo ErrHand
    SQLQuery = "Insert Into ast ( "
    SQLQuery = SQLQuery & "astCode, "
    SQLQuery = SQLQuery & "astAtfCode, "
    SQLQuery = SQLQuery & "astShfCode, "
    SQLQuery = SQLQuery & "astVefCode, "
    SQLQuery = SQLQuery & "astSdfCode, "
    SQLQuery = SQLQuery & "astLsfCode, "
    SQLQuery = SQLQuery & "astAirDate, "
    SQLQuery = SQLQuery & "astAirTime, "
    SQLQuery = SQLQuery & "astStatus, "
    SQLQuery = SQLQuery & "astCPStatus, "
    SQLQuery = SQLQuery & "astFeedDate, "
    SQLQuery = SQLQuery & "astFeedTime, "
    SQLQuery = SQLQuery & "astAdfCode, "
    SQLQuery = SQLQuery & "astDatCode, "
    SQLQuery = SQLQuery & "astCpfCode, "
    SQLQuery = SQLQuery & "astRsfCode, "
    SQLQuery = SQLQuery & "astStationCompliant, "
    SQLQuery = SQLQuery & "astAgencyCompliant, "
    SQLQuery = SQLQuery & "astAffidavitSource, "
    SQLQuery = SQLQuery & "astCntrNo, "
    SQLQuery = SQLQuery & "astLen, "
    SQLQuery = SQLQuery & "astLkAstCode, "
    SQLQuery = SQLQuery & "astMissedMnfCode, "
    SQLQuery = SQLQuery & "astUstCode "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & tlAstInfo.lAttCode & ", "
    SQLQuery = SQLQuery & tlAstInfo.iShttCode & ", "
    SQLQuery = SQLQuery & tlAstInfo.iVefCode & ", "
    SQLQuery = SQLQuery & 0 & ", "         'astsdfCode
    SQLQuery = SQLQuery & llLstCode & ", "  'astlsfCode
    SQLQuery = SQLQuery & "'" & Format$(tlAstInfo.sAirDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlAstInfo.sAirTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & ASTEXTENDED_MG & ", " 'astStatus
    SQLQuery = SQLQuery & tlAstInfo.iCPStatus & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlAstInfo.sAirDate, sgSQLDateForm) & "', "  'astFeedDate
    SQLQuery = SQLQuery & "'" & Format$(tlAstInfo.sAirTime, sgSQLTimeForm) & "', "  'astFeedTime
    SQLQuery = SQLQuery & tlAstInfo.iAdfCode & ", "
    SQLQuery = SQLQuery & tlAstInfo.lDatCode & ", "
    SQLQuery = SQLQuery & tlAstInfo.lCpfCode & ", "
    SQLQuery = SQLQuery & tlAstInfo.lRRsfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'astStationCompliant
    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "   'astAgencyCompliant
    SQLQuery = SQLQuery & "'" & gFixQuote(slAffidavitSource) & "', "
    SQLQuery = SQLQuery & tlAstInfo.lCntrNo & ", "
    SQLQuery = SQLQuery & tlAstInfo.iLen & ", "
    SQLQuery = SQLQuery & tlAstInfo.lCode & ", "    'astLkAstCode
    SQLQuery = SQLQuery & 0 & ", "   'astMissedMnfCode
    SQLQuery = SQLQuery & igUstCode 'astUstCode
    SQLQuery = SQLQuery & ") "
    llAst = gInsertAndReturnCode(SQLQuery, "ast", "astCode", "Replace")
    If llAst > 0 Then
        mAddAstMG = llAst
    Else
        mAddAstMG = 0
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Affiliate Affidavit-mChgAstToMissed"
End Function


Private Sub mChgAstToMissed(llMGAstCode As Long, tlAstInfo As ASTINFO)
    On Error GoTo ErrHand
    SQLQuery = "UPDATE ast SET "
    SQLQuery = SQLQuery & "astAirDate = '" & Format$(tlAstInfo.sAirDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "astAirTime = '" & Format$(tlAstInfo.sAirTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery + "astLkAstCode = " & llMGAstCode & ", "
    SQLQuery = SQLQuery + "astAgencyCompliant = '" & "Y" & "',"
    SQLQuery = SQLQuery + "astStationCompliant = '" & "Y" & "',"
    SQLQuery = SQLQuery + "astStatus = " & tlAstInfo.iStatus
    '10/19/18: added setting user
    SQLQuery = SQLQuery + ", " & "astUstCode = " & igUstCode
    SQLQuery = SQLQuery + " WHERE (astCode = " & tlAstInfo.lCode & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        gSetMousePointer grdPost, grdPost, vbDefault
        gHandleError "AffErrorLog.txt", "CPDateTimes-mChgAstToMissed"
        Exit Sub
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Affiliate Affidavit-mChgAstToMissed"
'ErrHand1:
'    gHandleError "AffErrorLog.txt", "Affiliate Affidavit-mChgAstToMissed"
'    Return
End Sub


Private Sub mPopStatus(llRow As Long)
    Dim iIndex As Integer
    Dim iLoop As Integer
    
    iIndex = grdPost.TextMatrix(llRow, ASTINDEX)
    lbcStatus.Clear
    'If tmAstInfo(iIndex).iStatus = 20 Then
    If gIsAstStatus(tmAstInfo(iIndex).iStatus, ASTEXTENDED_MG) Then
        For iLoop = 0 To UBound(tgStatusTypes) Step 1
            If (tgStatusTypes(iLoop).iPledged = 2) Or (Trim$(tgStatusTypes(iLoop).sName) = "MG") Then
                lbcStatus.AddItem Trim$(tgStatusTypes(iLoop).sName)
                lbcStatus.ItemData(lbcStatus.NewIndex) = iLoop
            End If
        Next iLoop
    'ElseIf tmAstInfo(iIndex).iStatus = 21 Then
    ElseIf gIsAstStatus(tmAstInfo(iIndex).iStatus, ASTEXTENDED_BONUS) Then
        For iLoop = 0 To UBound(tgStatusTypes) Step 1
            If (tgStatusTypes(iLoop).iPledged = 2) Or (Trim$(tgStatusTypes(iLoop).sName) = "Bonus") Then
                lbcStatus.AddItem Trim$(tgStatusTypes(iLoop).sName)
                lbcStatus.ItemData(lbcStatus.NewIndex) = iLoop
            End If
        Next iLoop
    ElseIf gIsAstStatus(tmAstInfo(iIndex).iStatus, ASTEXTENDED_REPLACEMENT) Then
        For iLoop = 0 To UBound(tgStatusTypes) Step 1
            If (tgStatusTypes(iLoop).iPledged = 2) Or (Trim$(tgStatusTypes(iLoop).sName) = "Replacement") Then
                lbcStatus.AddItem Trim$(tgStatusTypes(iLoop).sName)
                lbcStatus.ItemData(lbcStatus.NewIndex) = iLoop
            End If
        Next iLoop
        
    Else
        For iLoop = 0 To UBound(tgStatusTypes) Step 1
            ''If tgStatusTypes(gGetAirStatus(iLoop)).iStatus < 20 Then
            'If tgStatusTypes(iLoop).iStatus < ASTEXTENDED_MG Then
            If tgStatusTypes(iLoop).iStatus < ASTEXTENDED_MG Or ((sgMissedMGBypass = "Y") And (tgStatusTypes(iLoop).iStatus = ASTAIR_MISSED_MG_BYPASS)) Then
                '3/11/11: Remove 7-Air Outside Pledge and 8-Air not pledged
                If (tgStatusTypes(iLoop).iStatus <> 6) And (tgStatusTypes(iLoop).iStatus <> 7) Then
                    lbcStatus.AddItem Trim$(tgStatusTypes(iLoop).sName)
                    lbcStatus.ItemData(lbcStatus.NewIndex) = iLoop
                End If
            End If
        Next iLoop
    End If
End Sub


