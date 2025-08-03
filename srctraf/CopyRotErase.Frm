VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form CopyRotErase 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5040
   ClientLeft      =   6420
   ClientTop       =   4755
   ClientWidth     =   9045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "CopyRotErase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lbcAvail 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "CopyRotErase.frx":08CA
      Left            =   6540
      List            =   "CopyRotErase.frx":08CC
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4215
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.CommandButton cmcErase 
      Caption         =   "Erase"
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
      Left            =   4965
      TabIndex        =   5
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
      TabIndex        =   3
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
      Picture         =   "CopyRotErase.frx":08CE
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   645
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Done"
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
      Left            =   2730
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCopyRot 
      Height          =   3810
      Left            =   165
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   345
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   6720
      _Version        =   393216
      Cols            =   14
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
      _Band(0).Cols   =   14
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label plcScreen 
      Caption         =   "Copy Rotation Erase"
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
      Left            =   285
      TabIndex        =   0
      Top             =   30
      Width           =   6930
   End
End
Attribute VB_Name = "CopyRotErase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  CopyRotErase - displays missed spots to be changed to Makegoods
'*
'*  Created January,1998 by Wade Bjerke
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
Private lmEnableRow As Long
Private lmEnableCol As Long
Private smNowDate As String
Private lmNowDate As Long

Dim smMMnfStamp As String
Dim tmMMnf() As MNF

Private imCopyColSorted As Integer
Private imCopySort As Integer

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Height adjustment factor

Dim tmAvailCode() As SORTCODE
Dim smAvailCodeTag As String

Private tmCifInfo() As CIFINFO
Private tmCrfDates() As CRFDATES

Private Type ASSIGNEDSPOTDATES
    iVefCode As Integer
    lStartDate As Long
    lEndDate As Long
End Type
Dim tmAssignedSpotDates() As ASSIGNEDSPOTDATES

'Inventory
Dim hmCif As Integer        'CIF Handle

Private rst_Crf As ADODB.Recordset
Private rst_Raf As ADODB.Recordset
Private rst_Smf As ADODB.Recordset
Private rst_Tzf As ADODB.Recordset
Private rst_Csf As ADODB.Recordset
Private rst_Caf As ADODB.Recordset
Private rst_Cnf As ADODB.Recordset
Private rst_Cuf As ADODB.Recordset
Private rst_Cvf As ADODB.Recordset
Private rst_Rsf As ADODB.Recordset
Private rst_Sdf As ADODB.Recordset

'Grid Name

Const CNTRINDEX = 0    'Contract control/field 870
Const LENINDEX = 1     'Length control/field 315
Const LIVECOPYINDEX = 2     'Length control/field 255
Const VEHICLEINDEX = 3  'Vehicle control/field 1185
Const MGASSIGNINDEX = 4 '180
Const ZONEINDEX = 5     'Zone control/field 360
Const RDATEINDEX = 6    'Start Date control/field 690
Const RTIMEINDEX = 7    'Start time control/field 600
Const DAYSINDEX = 8      'Monday control/field 210
Const AVAILINDEX = 9   'Avail control/field 885, 885
Const TYPEINDEX = 10    'Type control/field 345
Const SORTINDEX = 11
Const SELECTEDINDEX = 12
Const CRFCODEINDEX = 13






Private Sub mClearGrid()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llCol                                                                                 *
'******************************************************************************************

    Dim llRow As Long
    Dim llCol As Long

    'Blank rows within grid
'    gGrid_Clear grdCopyRot, True
    'Set color within cells
    grdCopyRot.rowHeight(0) = fgBoxGridH + 15
    For llRow = grdCopyRot.FixedRows To grdCopyRot.Rows - 1 Step 1
        For llCol = CNTRINDEX To CRFCODEINDEX Step 1
            grdCopyRot.Col = llCol
            grdCopyRot.CellBackColor = WHITE
            grdCopyRot.TextMatrix(llRow, llCol) = ""
        Next llCol
        grdCopyRot.rowHeight(llRow) = fgBoxGridH + 15
    Next llRow
End Sub



Private Sub cmcDone_Click()
    Dim llRow As Long
    Dim slStr As String

    mTerminate
End Sub

Private Sub cmcErase_Click()
    Dim llRow As Long
    Dim slStr As String
    Dim ilRet As Integer
    Dim llCrfCode As Long
    Dim slSQLQuery As String
    
    ilRet = MsgBox("This will Remove the Rotation and All Copy Assignment to Spots for selected Rotation", vbOKCancel + vbExclamation, "Erase Rotation")
    If ilRet = vbCancel Then
        Exit Sub
    End If
    For llRow = grdCopyRot.FixedRows To grdCopyRot.Rows - 1 Step 1
        slStr = Trim$(grdCopyRot.TextMatrix(llRow, CNTRINDEX))
        If (slStr <> "") Then
            If grdCopyRot.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                llCrfCode = grdCopyRot.TextMatrix(llRow, CRFCODEINDEX)
                slSQLQuery = "Select * From crf_Copy_Rot_Header"
                slSQLQuery = slSQLQuery & " Where crfCode = " & llCrfCode
                Set rst_Crf = gSQLSelectCall(slSQLQuery)
                If Not rst_Crf.EOF Then
                    ilRet = mEraseRot(llCrfCode)
                End If
            End If
        End If
    Next llRow
    mClearGrid
    mPopulate
End Sub

Private Sub Form_Activate()

    If imFirstTime Then
        imFirstTime = False
        gSetMousePointer grdCopyRot, grdCopyRot, vbDefault
    End If

End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    'Me.Width = (CLng(90) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    'Me.Height = (CLng(90) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        'fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        'Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        'fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        'Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
        fmAdjFactorW = ((80 * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = (80 * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.height) / (480 * 15 / Me.height))) / 100) / Me.height
        Me.height = (lgPercentAdjH * ((Screen.height) / (480 * 15 / Me.height))) / 100
    End If
    gCenterStdAlone CopyRotErase
    DoEvents
    mSetControls
End Sub

Private Sub Form_Load()

    gSetMousePointer grdCopyRot, grdCopyRot, vbHourglass

    mInit
    gSetMousePointer grdCopyRot, grdCopyRot, vbDefault
    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer

    On Error Resume Next
    
    Erase tmCifInfo
    Erase tmCrfDates
    Erase tmAvailCode
    Erase tmAssignedSpotDates
    
    rst_Crf.Close
    rst_Raf.Close
    rst_Smf.Close
    rst_Tzf.Close
    rst_Csf.Close
    rst_Caf.Close
    rst_Cnf.Close
    rst_Cuf.Close
    rst_Cvf.Close
    rst_Rsf.Close
    rst_Sdf.Close
    
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    
    Set CopyRotErase = Nothing
End Sub


Private Sub grdCopyRot_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub grdCopyRot_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
    imShiftKey = False
End Sub

Private Sub mInit()
    Dim ilRet As Integer
    Dim ilAdf As Integer

    gSetMousePointer grdCopyRot, grdCopyRot, vbHourglass
    ilAdf = gBinarySearchAdf(igGetCopyGridAdfCode)
    If ilAdf <> -1 Then
        plcScreen.Caption = "Copy Rotation Erase-" & Trim$(tgCommAdf(ilAdf).sName)
    Else
        plcScreen.Caption = "Copy Rotation Erase"
    End If
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmLastClickedRow = -1
    lmScrollTop = grdCopyRot.FixedRows
    imCopyColSorted = -1
    imCopySort = -1

    hmCif = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    
    smMMnfStamp = ""
    ilRet = gObtainMnfForType("H", smMMnfStamp, tmMMnf())

    mAvailsPop
    mClearGrid
    'ilRet = mReadRec()
    mPopulate

    Screen.MousePointer = vbDefault
    gSetMousePointer grdCopyRot, grdCopyRot, vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdCopyRot, grdCopyRot, vbDefault
    Exit Sub

End Sub

Private Sub mPopulate()
    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilValue As Integer
    Dim slRecCode As String
    Dim ilTest As Integer
    Dim slSQLQuery As String
    Dim slDay As String
    Dim slFormat As String
    Dim ilMnf As Integer
    
    On Error GoTo ErrHand:

    grdCopyRot.Redraw = False
    cmcErase.Enabled = False
    'grdCopyRot.Row = 0
    'For llCol = GRIDNAMEINDEX To ENTEREDDATEINDEX Step 1
    '    grdCopyRot.Col = llCol
    '    grdCopyRot.CellBackColor = vbBlack
    '    grdCopyRot.CellBackColor = LIGHTBLUE
    'Next llCol
    grdCopyRot.rowHeight(0) = fgBoxGridH + 15
    llRow = grdCopyRot.FixedRows
    slSQLQuery = "Select * From crf_Copy_Rot_Header Left Outer Join chf_Contract_Header On crfChfCode = chfCode"
    slSQLQuery = slSQLQuery & " Where crfAdfCode = " & igCopyRotEraseAdfCode
    slSQLQuery = slSQLQuery & " And crfStartDate > '" & Format(smNowDate, sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery & " Order By chfCntrNo Desc, crfRotNo Desc"
    Set rst_Crf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Crf.EOF
        If llRow >= grdCopyRot.Rows Then
            grdCopyRot.AddItem ""
        End If
        grdCopyRot.rowHeight(llRow) = fgBoxGridH + 15
        If lgCopyRotEraseCrfCode <> rst_Crf!crfCode Then
            grdCopyRot.TextMatrix(llRow, SELECTEDINDEX) = "0"
        Else
            grdCopyRot.TextMatrix(llRow, SELECTEDINDEX) = "1"
            cmcErase.Enabled = True
        End If
        grdCopyRot.TextMatrix(llRow, CNTRINDEX) = Trim$(str(rst_Crf!chfCntrNo)) & " i" & Trim$(str$(rst_Crf!crfRotNo)) & " R" & rst_Crf!chfCntRevNo & " " & Format(rst_Crf!chfStartDate, "m/d/yy") & "-" & Format(rst_Crf!chfEndDate, "m/d/yy") & " " & Trim$(rst_Crf!chfProduct)
        grdCopyRot.TextMatrix(llRow, LENINDEX) = Trim$(str$(rst_Crf!crfLen))
        'Live Copy
        If ((Asc(tgSpf.sUsingFeatures) And LIVECOPY) = LIVECOPY) Or ((Asc(tgSpf.sUsingFeatures3) And PROMOCOPY) = PROMOCOPY) Then 'Using Live Copy
            If rst_Crf!crfLiveCopy = "L" Then
                slStr = "LC"
            ElseIf rst_Crf!crfLiveCopy = "M" Then
                slStr = "LP"
            ElseIf rst_Crf!crfLiveCopy = "S" Then
                slStr = "RP"
            ElseIf rst_Crf!crfLiveCopy = "P" Then
                slStr = "PC"
            ElseIf rst_Crf!crfLiveCopy = "Q" Then
                slStr = "PP"
            Else
                slStr = "RC"
            End If
        Else
            slStr = "RC"
        End If
        grdCopyRot.TextMatrix(llRow, LIVECOPYINDEX) = slStr
        If rst_Crf!CRFVEFCODE = 0 And rst_Crf!crfCvfCode <> 0 Then
            slStr = "[Several]"
        Else
            ilVef = gBinarySearchVef(rst_Crf!CRFVEFCODE)
            If ilVef <> -1 Then
                slStr = Trim$(tgMVef(ilVef).sName)
                slFormat = ""
                If tgMVef(ilVef).iMnfVehGp4Fmt > 0 Then
                    For ilMnf = LBound(tmMMnf) To UBound(tmMMnf) - 1 Step 1
                        If tgMVef(ilVef).iMnfVehGp4Fmt = tmMMnf(ilMnf).iCode Then
                            slFormat = Trim$(tmMMnf(ilMnf).sName)
                            slStr = slStr & "/" & slFormat
                            Exit For
                        End If
                    Next ilMnf
                End If
                
            Else
                slStr = "Missing: " & rst_Crf!CRFVEFCODE
            End If
        End If
        grdCopyRot.TextMatrix(llRow, VEHICLEINDEX) = slStr
        If rst_Crf!crfMGCopyAssign = "O" Then
            slStr = "O"
        ElseIf rst_Crf!crfMGCopyAssign = "S" Then
            slStr = "S"
        ElseIf rst_Crf!crfMGCopyAssign = "B" Then
            slStr = "B"
        Else
            ilValue = Asc(tgSpf.sMOFCopyAssign)  'Mg Copy Assignment
            If (ilValue And MGORIGVEHONLY) = MGORIGVEHONLY Then
                slStr = "O"
            ElseIf (ilValue And MGSCHVEHONLY) = MGSCHVEHONLY Then
                slStr = "S"
            Else
                slStr = "B"
            End If
        End If
        grdCopyRot.TextMatrix(llRow, MGASSIGNINDEX) = slStr
        If (rst_Crf!crfRafCode > 0) And (Trim$(rst_Crf!crfZone) = "R") Then
            slSQLQuery = "Select rafName From raf_Region_Area Where rafCode = " & rst_Crf!crfRafCode
            Set rst_Raf = gSQLSelectCall(slSQLQuery)
            If Not rst_Raf.EOF Then
                If rst_Crf!crfBkoutInstAdfCode <= 0 Then
                    slStr = "r" & " " & rst_Raf!rafName
                Else
                    slStr = "b" & " " & rst_Raf!rafName
                End If
            Else
                slStr = "Generic"
            End If
        Else
            If Trim$(rst_Crf!crfZone) = "" Then
                slStr = "Generic"
            Else
                slStr = Trim$(rst_Crf!crfZone)
            End If
        End If
        grdCopyRot.TextMatrix(llRow, ZONEINDEX) = slStr
        slStr = Format(rst_Crf!crfStartDate, "m/d/yy") & "-" & Format(rst_Crf!crfEndDate, "m/d/yy")
        grdCopyRot.TextMatrix(llRow, RDATEINDEX) = slStr
        If rst_Crf!crfAirGameType = "G" Then
            slStr = "By Event"
            grdCopyRot.TextMatrix(llRow, RTIMEINDEX) = slStr
            grdCopyRot.TextMatrix(llRow, DAYSINDEX) = slStr
        ElseIf rst_Crf!crfAirGameType = "T" Then
            slStr = "By Team"
            grdCopyRot.TextMatrix(llRow, RTIMEINDEX) = slStr
            grdCopyRot.TextMatrix(llRow, DAYSINDEX) = slStr
        Else
            slStr = gFormatTime(Format(rst_Crf!crfStartTime, "ttttt"), "A", "1") & "-" & gFormatTime(Format(rst_Crf!crfEndTime, "ttttt"), "A", "1")
            grdCopyRot.TextMatrix(llRow, RTIMEINDEX) = slStr
            slStr = ""
            slDay = ""
            If rst_Crf!crfMo = "Y" Then
                slStr = slStr & "Mo"
                slDay = slDay & "Y"
            Else
                slDay = slDay & "N"
            End If
            If rst_Crf!crfTu = "Y" Then
                slStr = slStr & "Tu"
                slDay = slDay & "Y"
            Else
                slDay = slDay & "N"
            End If
            If rst_Crf!crfWe = "Y" Then
                slStr = slStr & "We"
                slDay = slDay & "Y"
            Else
                slDay = slDay & "N"
            End If
            If rst_Crf!crfTh = "Y" Then
                slStr = slStr & "Th"
                slDay = slDay & "Y"
            Else
                slDay = slDay & "N"
            End If
            If rst_Crf!crfFr = "Y" Then
                slStr = slStr & "Fr"
                slDay = slDay & "Y"
            Else
                slDay = slDay & "N"
            End If
            If rst_Crf!crfSa = "Y" Then
                slStr = slStr & "Sa"
                slDay = slDay & "Y"
            Else
                slDay = slDay & "N"
            End If
            If rst_Crf!crfSu = "Y" Then
                slStr = slStr & "Su"
                slDay = slDay & "Y"
            Else
                slDay = slDay & "N"
            End If
            slStr = gDayStrToNames(slDay, slStr, 2)
            grdCopyRot.TextMatrix(llRow, DAYSINDEX) = slStr
        End If
        If (rst_Crf!crfInOut = "I") Or (rst_Crf!crfInOut = "O") Then
            slStr = "Missing"
            slRecCode = Trim$(str$(rst_Crf!crfanfCode))
            For ilTest = 0 To UBound(tmAvailCode) - 1 Step 1 'lbcAvailCode.ListCount - 1 Step 1
                slNameCode = tmAvailCode(ilTest).sKey  'lbcAvailCode.List(ilTest)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If ilRet = CP_MSG_NONE Then
                    If slRecCode = slCode Then
                        ilRet = gParseItem(slNameCode, 1, "\", slStr)
                        If rst_Crf!crfInOut = "O" Then
                            slStr = "Ø" & slStr
                        End If
                        Exit For
                    End If
                End If
            Next ilTest
        Else
            slStr = "All Avails"
        End If
        grdCopyRot.TextMatrix(llRow, AVAILINDEX) = slStr
        
        Select Case rst_Crf!crfRotType
            Case "A"
                slStr = "C S"
            Case "O"
                If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
                    slStr = "BB"
                Else
                    slStr = "OBB"
                End If
            Case "C"
                slStr = "CBB"
            Case "E"
                slStr = "ABB"
        End Select
        grdCopyRot.TextMatrix(llRow, TYPEINDEX) = slStr
        grdCopyRot.TextMatrix(llRow, CRFCODEINDEX) = rst_Crf!crfCode
        
        mPaintRowColor llRow
        llRow = llRow + 1
        rst_Crf.MoveNext
    Loop
    
'    For llRow = grdCopyRot.FixedRows To grdCopyRot.Rows - 1 Step 1
'        slStr = Trim$(grdCopyRot.TextMatrix(llRow, GRIDNAMEINDEX))
'        If slStr <> "" Then
'            grdCopyRot.Row = llRow
'            grdCopyRot.Col = GRIDNAMEINDEX
'            grdCopyRot.CellBackColor = LIGHTYELLOW
'            grdCopyRot.Col = ENTEREDDATEINDEX
'            grdCopyRot.CellBackColor = LIGHTYELLOW
'        End If
'    Next llRow
'    'rst_Gsf.Close
'    gGrid_AlignAllColsLeft grdCopyRot
'    mCopyRotSortCol ENTEREDDATEINDEX
'    mCopyRotSortCol ENTEREDDATEINDEX
    grdCopyRot.Redraw = True
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    Resume Next
    On Error GoTo 0

End Sub

Private Sub mSetRotColumns()
    Dim ilCol As Integer
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim flAdj As Single
    Dim ilColCount As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer
    Dim llAdj As Long

    grdCopyRot.ColWidth(CRFCODEINDEX) = 0
    grdCopyRot.ColWidth(SORTINDEX) = 0
    grdCopyRot.ColWidth(SELECTEDINDEX) = 0
    grdCopyRot.ColWidth(CNTRINDEX) = 1560 * fmAdjFactorW
    grdCopyRot.ColWidth(LENINDEX) = 315 * fmAdjFactorW
    grdCopyRot.ColWidth(LIVECOPYINDEX) = 255 * fmAdjFactorW
    grdCopyRot.ColWidth(VEHICLEINDEX) = 1385 * fmAdjFactorW
    grdCopyRot.ColWidth(MGASSIGNINDEX) = 180 * fmAdjFactorW
    grdCopyRot.ColWidth(ZONEINDEX) = 960 * fmAdjFactorW
    grdCopyRot.ColWidth(RDATEINDEX) = 690 * fmAdjFactorW
    grdCopyRot.ColWidth(RTIMEINDEX) = 1000 * fmAdjFactorW
    grdCopyRot.ColWidth(DAYSINDEX) = 870 * fmAdjFactorW
    grdCopyRot.ColWidth(AVAILINDEX) = 885 * fmAdjFactorW
    grdCopyRot.ColWidth(TYPEINDEX) = 345 * fmAdjFactorW
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdCopyRot.Width
    For ilCol = 0 To grdCopyRot.Cols - 1 Step 1
        llWidth = llWidth + grdCopyRot.ColWidth(ilCol)
        If (grdCopyRot.ColWidth(ilCol) > 15) And (grdCopyRot.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdCopyRot.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdCopyRot.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdCopyRot.Width
            For ilCol = 0 To grdCopyRot.Cols - 1 Step 1
                If (grdCopyRot.ColWidth(ilCol) > 15) And (grdCopyRot.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdCopyRot.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdCopyRot.FixedCols To grdCopyRot.Cols - 1 Step 1
                If grdCopyRot.ColWidth(ilCol) > 15 Then
                    ilColInc = grdCopyRot.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdCopyRot.ColWidth(ilCol) = grdCopyRot.ColWidth(ilCol) + 15
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
    'Align columns to left
    gGrid_AlignAllColsLeft grdCopyRot
    
'    grdCopyRot.ColWidth(ENTEREDDATEINDEX) = grdCopyRot.Width * 0.2
'    grdCopyRot.ColWidth(GRIDNAMEINDEX) = grdCopyRot.Width - grdCopyRot.ColWidth(ENTEREDDATEINDEX) - GRIDSCROLLWIDTH - 15
    'Align columns to left
    'gGrid_AlignAllColsLeft grdCopyRot
End Sub

Private Sub mSetRotTitles()
    'Set column titles
'    grdCopyRot.TextMatrix(0, GRIDNAMEINDEX) = "Name"
'    grdCopyRot.TextMatrix(0, ENTEREDDATEINDEX) = "Entered Date"
    grdCopyRot.TextMatrix(0, CNTRINDEX) = "Contract #"
    grdCopyRot.TextMatrix(0, LENINDEX) = "Len"
    grdCopyRot.TextMatrix(0, LIVECOPYINDEX) = "R/L"
    grdCopyRot.TextMatrix(0, VEHICLEINDEX) = "Vehicle"
    grdCopyRot.TextMatrix(0, MGASSIGNINDEX) = "MG"
    grdCopyRot.TextMatrix(0, ZONEINDEX) = "Region"
    grdCopyRot.TextMatrix(0, RDATEINDEX) = "Date Range"
    grdCopyRot.TextMatrix(0, RTIMEINDEX) = "Time Range"
    grdCopyRot.TextMatrix(0, DAYSINDEX) = "Days"
    grdCopyRot.TextMatrix(0, AVAILINDEX) = "Avail"
    grdCopyRot.TextMatrix(0, TYPEINDEX) = "Type"
    'Set height of grid

End Sub

Private Sub mCopyRotSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdCopyRot.FixedRows To grdCopyRot.Rows - 1 Step 1
        slStr = Trim$(grdCopyRot.TextMatrix(llRow, CNTRINDEX))
        If slStr <> "" Then
            If ilCol = RDATEINDEX Then
                slStr = grdCopyRot.TextMatrix(llRow, RDATEINDEX)
                If slStr <> "" Then
                    slSort = Trim$(str$(gDateValue(slStr)))
                Else
                    slSort = "0"
                End If
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = RDATEINDEX Then
                slStr = grdCopyRot.TextMatrix(llRow, RDATEINDEX)
                If slStr <> "" Then
                    slSort = Trim$(str$(gDateValue(slStr)))
                Else
                    slSort = "0"
                End If
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdCopyRot.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdCopyRot.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imCopyColSorted) Or ((ilCol = imCopyColSorted) And (imCopySort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdCopyRot.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdCopyRot.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imCopyColSorted Then
        imCopyColSorted = SORTINDEX
    Else
        imCopyColSorted = -1
        imCopySort = -1
    End If
    gGrid_SortByCol grdCopyRot, CNTRINDEX, SORTINDEX, imCopyColSorted, imCopySort
    imCopyColSorted = ilCol
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
    gSetMousePointer grdCopyRot, grdCopyRot, vbDefault
    igManUnload = YES
    Unload CopyRotErase
    igManUnload = NO
End Sub


Private Sub mSetControls()

    Dim ilGap As Integer

    ilGap = cmcErase.Left - (cmcDone.Left + cmcDone.Width)
    cmcDone.Top = Me.height - cmcDone.height - 120
    cmcErase.Top = cmcDone.Top
    'ckcIncludeCopy.Top = cmcDone.Top + cmcDone.Height - ckcIncludeCopy.Height
    grdCopyRot.Move 180, 255, CopyRotErase.Width - 360, cmcDone.Top - 255 - 120
    'ckcIncludeCopy.Left = grdCopyRot.Left
    'cmcErase.Left = grdCopyRot.Left + grdCopyRot.Width - cmcErase.Width
    'cmcCancel.Left = cmcErase.Left - cmcCancel.Width - ilGap
    'cmcDone.Left = cmcCancel.Left - cmcDone.Width - ilGap
    cmcDone.Left = grdCopyRot.Width / 2 - ilGap / 2 - cmcDone.Width
    cmcErase.Left = grdCopyRot.Width / 2 + ilGap / 2
    mSetRotColumns
    mSetRotTitles
    gGrid_IntegralHeight grdCopyRot, fgBoxGridH + 15

End Sub


Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long

    grdCopyRot.Row = llRow
    For llCol = CNTRINDEX To TYPEINDEX Step 1
        grdCopyRot.Col = llCol
        If grdCopyRot.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
            grdCopyRot.CellBackColor = LIGHTYELLOW
        Else
            grdCopyRot.CellBackColor = GRAY    'vbBlue
        End If
    Next llCol

End Sub

Private Sub grdCopyRot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    'If Y < grdCopyRot.RowHeight(0) Then
    '    grdCopyRot.Col = grdCopyRot.MouseCol
    '    mCopyRotSortCol grdCopyRot.Col
    '    grdCopyRot.Row = 0
    '    grdCopyRot.Col = CRFCODEINDEX
    '    Exit Sub
    'End If
    ilFound = gGrid_GetRowCol(grdCopyRot, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdCopyRot.FixedRows Then
        Exit Sub
    End If
    gSetMousePointer grdCopyRot, grdCopyRot, vbHourglass
    grdCopyRot.Redraw = False
    If llCurrentRow >= grdCopyRot.FixedRows Then
        If grdCopyRot.TextMatrix(llCurrentRow, CNTRINDEX) <> "" Then
            grdCopyRot.TopRow = lmScrollTop
            llTopRow = grdCopyRot.TopRow
            If (Shift And CTRLMASK) > 0 Then
                If grdCopyRot.TextMatrix(grdCopyRot.Row, CRFCODEINDEX) <> "" Then
                    If grdCopyRot.TextMatrix(grdCopyRot.Row, SELECTEDINDEX) <> "1" Then
                        grdCopyRot.TextMatrix(grdCopyRot.Row, SELECTEDINDEX) = "1"
                    Else
                        grdCopyRot.TextMatrix(grdCopyRot.Row, SELECTEDINDEX) = "0"
                    End If
                    mPaintRowColor grdCopyRot.Row
                End If
            Else
                For llRow = grdCopyRot.FixedRows To grdCopyRot.Rows - 1 Step 1
                    If grdCopyRot.TextMatrix(llRow, CNTRINDEX) <> "" Then
                        grdCopyRot.TextMatrix(llRow, SELECTEDINDEX) = "0"
                        If grdCopyRot.TextMatrix(llRow, CRFCODEINDEX) <> "" Then
                            If (lmLastClickedRow = -1) Or ((Shift And SHIFTMASK) <= 0) Then
                                If llRow = llCurrentRow Then
                                    grdCopyRot.TextMatrix(llRow, SELECTEDINDEX) = "1"
                                Else
                                    grdCopyRot.TextMatrix(llRow, SELECTEDINDEX) = "0"
                                End If
                            ElseIf lmLastClickedRow < llCurrentRow Then
                                If (llRow >= lmLastClickedRow) And (llRow <= llCurrentRow) Then
                                    grdCopyRot.TextMatrix(llRow, SELECTEDINDEX) = "1"
                                End If
                            Else
                                If (llRow >= llCurrentRow) And (llRow <= lmLastClickedRow) Then
                                    grdCopyRot.TextMatrix(llRow, SELECTEDINDEX) = "1"
                                End If
                            End If
                            mPaintRowColor llRow
                        End If
                    End If
                Next llRow
                grdCopyRot.TopRow = llTopRow
                grdCopyRot.Row = llCurrentRow
            End If
            lmLastClickedRow = llCurrentRow
        End If
    End If
    mSetCommands
    grdCopyRot.Redraw = True
    gSetMousePointer grdCopyRot, grdCopyRot, vbDefault
    

End Sub

Private Function mEraseRot(llCrfCode As Long) As Integer
    Dim ilAdfCode As Integer
    Dim llChfCode As Long
    Dim ilRotNo As Integer
    Dim slSQLQuery As String
    Dim llRet As Long
    Dim llChgCount As Long
    Dim llTotalCount As Long
    Dim ilVef As Integer
    
    On Error GoTo Eraseerr:
    ReDim tmAssignedSpotDates(0 To 0) As ASSIGNEDSPOTDATES
    ReDim tmCifInfo(0 To 0) As CIFINFO
    ReDim tmCrfDates(0 To 0) As CRFDATES
    tmCrfDates(0).lNext = -1
    
    llChfCode = rst_Crf!crfChfCode
    ilAdfCode = rst_Crf!crfAdfCode
    ilRotNo = rst_Crf!crfRotNo
    
    'Gather vehicles for the contract
    slSQLQuery = "Select Distinct sdfVefCode From sdf_Spot_Detail"
    slSQLQuery = slSQLQuery & " Where sdfChfCode = " & llChfCode
    slSQLQuery = slSQLQuery & " And sdfRotNo = " & ilRotNo
    Set rst_Sdf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Sdf.EOF
        tmAssignedSpotDates(UBound(tmAssignedSpotDates)).iVefCode = rst_Sdf!sdfVefCode
        tmAssignedSpotDates(UBound(tmAssignedSpotDates)).lStartDate = -1
        tmAssignedSpotDates(UBound(tmAssignedSpotDates)).lEndDate = -1
        ReDim Preserve tmAssignedSpotDates(0 To UBound(tmAssignedSpotDates) + 1) As ASSIGNEDSPOTDATES
        rst_Sdf.MoveNext
    Loop
    If UBound(tmAssignedSpotDates) <= LBound(tmAssignedSpotDates) Then
        slSQLQuery = "Select Distinct sdfVefCode From RSF_Region_Schd_Copy"
        slSQLQuery = slSQLQuery & " Left Outer Join sdf_Spot_Detail On rsfSdfCode = sdfCode"
        slSQLQuery = slSQLQuery & " Where rsfRChfCode = " & llChfCode
        slSQLQuery = slSQLQuery & " And rsfRotNo = " & ilRotNo
        Set rst_Rsf = gSQLSelectCall(slSQLQuery)
        Do While Not rst_Rsf.EOF
            tmAssignedSpotDates(UBound(tmAssignedSpotDates)).iVefCode = rst_Rsf!sdfVefCode
            tmAssignedSpotDates(UBound(tmAssignedSpotDates)).lStartDate = -1
            tmAssignedSpotDates(UBound(tmAssignedSpotDates)).lEndDate = -1
            ReDim Preserve tmAssignedSpotDates(0 To UBound(tmAssignedSpotDates) + 1) As ASSIGNEDSPOTDATES
            rst_Rsf.MoveNext
        Loop
    End If
    
    'Determine earlest and latest date spots assigned copy, used to set alerts
    For ilVef = 0 To UBound(tmAssignedSpotDates) - 1 Step 1
        slSQLQuery = "Select Min(sdfDate) as StartDate, Max(sdfDate) as EndDate From sdf_Spot_Detail"
        slSQLQuery = slSQLQuery & " Left Outer Join rsf_Region_Schd_Copy on rsfsdfCode = sdfCode"
        slSQLQuery = slSQLQuery & " Where sdfChfCode = " & llChfCode
        slSQLQuery = slSQLQuery & " And sdfVefCode = " & tmAssignedSpotDates(ilVef).iVefCode
        slSQLQuery = slSQLQuery & " And rsfRotNo = " & ilRotNo
        Set rst_Sdf = gSQLSelectCall(slSQLQuery)
        If Not rst_Sdf.EOF Then
            If rst_Sdf!StartDate <> "" Then
                mSetUnassignedDate gDateValue(rst_Sdf!StartDate), gDateValue(rst_Sdf!EndDate), tmAssignedSpotDates(ilVef).lStartDate, tmAssignedSpotDates(ilVef).lEndDate
            End If
        End If
        slSQLQuery = "Select Min(sdfDate) as StartDate, Max(sdfDate) as EndDate From sdf_Spot_Detail"
        slSQLQuery = slSQLQuery & " Where sdfChfCode = " & llChfCode
        slSQLQuery = slSQLQuery & " And sdfVefCode = " & tmAssignedSpotDates(ilVef).iVefCode
        slSQLQuery = slSQLQuery & " And sdfRotNo = " & ilRotNo & " And sdfPointer = '3'"
        Set rst_Sdf = gSQLSelectCall(slSQLQuery)
        If Not rst_Sdf.EOF Then
            If rst_Sdf!StartDate <> "" Then
                mSetUnassignedDate gDateValue(rst_Sdf!StartDate), gDateValue(rst_Sdf!EndDate), tmAssignedSpotDates(ilVef).lStartDate, tmAssignedSpotDates(ilVef).lEndDate
            End If
        End If
        slSQLQuery = "Select Min(sdfDate) as StartDate, Max(sdfDate) as EndDate From sdf_Spot_Detail"
        slSQLQuery = slSQLQuery & " Where sdfChfCode = " & llChfCode
        slSQLQuery = slSQLQuery & " And sdfVefCode = " & tmAssignedSpotDates(ilVef).iVefCode
        slSQLQuery = slSQLQuery & " And sdfRotNo = " & ilRotNo
        Set rst_Sdf = gSQLSelectCall(slSQLQuery)
        If Not rst_Sdf.EOF Then
            If rst_Sdf!StartDate <> "" Then
                mSetUnassignedDate gDateValue(rst_Sdf!StartDate), gDateValue(rst_Sdf!EndDate), tmAssignedSpotDates(ilVef).lStartDate, tmAssignedSpotDates(ilVef).lEndDate
            End If
        End If
    Next ilVef
    
    'Remove time zone reference
    slSQLQuery = "Select * From smf_Spot_MG_Specs"
    slSQLQuery = slSQLQuery & " Where smfChfCode = " & llChfCode
    slSQLQuery = slSQLQuery & " And smfRotNo = " & ilRotNo & " And smfPointer = '3'"
    Set rst_Smf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Smf.EOF
        slSQLQuery = "Delete From tzf_Time_Zone_Copy Where tzfCode = " & rst_Smf!smfCopy
        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
        rst_Smf.MoveNext
    Loop
    
    'Clear retained copy with MG specification
    slSQLQuery = "Update smf_Spot_MG_Specs Set smfPointer = '0', smfRotNo = 0, smfCopy = 0"
    slSQLQuery = slSQLQuery & " Where smfChfCode = " & llChfCode & " And smfRotNo = " & ilRotNo
    llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            
    'Remove region assigned copy to spots
    If rst_Crf!crfRafCode > 0 Then
        
        slSQLQuery = "Select * From sdf_Spot_Detail"
        slSQLQuery = slSQLQuery & " Where sdfChfCode = " & llChfCode
        Set rst_Sdf = gSQLSelectCall(slSQLQuery)
        Do While Not rst_Sdf.EOF
            slSQLQuery = "Delete From rsf_Region_Schd_Copy Where rsfsdfCode = " & rst_Sdf!sdfCode
            slSQLQuery = slSQLQuery & " And rsfRotNo = " & ilRotNo
            slSQLQuery = slSQLQuery & " And rsfRafCode = " & rst_Crf!crfRafCode
            llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            'Remove blackout lst
            slSQLQuery = "Delete From lst Where lstsdfCode = " & rst_Sdf!sdfCode
            slSQLQuery = slSQLQuery & " And lstBkoutLstCode > " & 0
            slSQLQuery = slSQLQuery & " And lstRafCode = " & rst_Crf!crfRafCode
            llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            rst_Sdf.MoveNext
        Loop
    End If
    
    'Remove time zone reference
    slSQLQuery = "Select * From sdf_Spot_Detail"
    slSQLQuery = slSQLQuery & " Where sdfChfCode = " & llChfCode
    slSQLQuery = slSQLQuery & " And sdfRotNo = " & ilRotNo & " And sdfPointer = '3'"
    Set rst_Sdf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Sdf.EOF
        slSQLQuery = "Delete From tzf_Time_Zone_Copy Where tzfCode = " & rst_Sdf!sdfCopy
        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
        rst_Sdf.MoveNext
    Loop
    
    
    'Clear copy assigned to spots
    slSQLQuery = "Update sdf_Spot_Detail Set sdfPointer = '0', sdfRotNo = 0, sdfCopy = 0"
    slSQLQuery = slSQLQuery & " Where sdfChfCode = " & llChfCode & " And sdfRotNo = " & ilRotNo
    llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    
    'Remove copy comments (check if same comment with many rotation
    If (rst_Crf!crfCsfCode <> 0) And (tgSpf.sGUseAffSys <> "Y") Then
        slSQLQuery = "Delete From csf_Copy_Scripts Where csfCode = " & rst_Crf!crfCsfCode
        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    End If
    
    'Build Array of copy inventory so that Inventory date can be addjusted
    slSQLQuery = "Select Distinct cnfCifCode From cnf_Copy_Instruction"
    slSQLQuery = slSQLQuery & " Where cnfCrfCode = " & llCrfCode
    Set rst_Cnf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Cnf.EOF
        tmCifInfo(UBound(tmCifInfo)).lCifCode = rst_Cnf!cnfCifCode
        tmCifInfo(UBound(tmCifInfo)).lFirst = -1
        If rst_Crf!crfBkoutInstAdfCode <= 0 Then
            tmCifInfo(UBound(tmCifInfo)).iAdfCode = ilAdfCode
        Else
            tmCifInfo(UBound(tmCifInfo)).iAdfCode = rst_Crf!crfBkoutInstAdfCode
        End If
        ReDim Preserve tmCifInfo(0 To UBound(tmCifInfo) + 1) As CIFINFO
        rst_Cnf.MoveNext
    Loop
    
    For ilVef = 0 To UBound(tmAssignedSpotDates) - 1 Step 1
        mSetAlerts tmAssignedSpotDates(ilVef).iVefCode, tmAssignedSpotDates(ilVef).lStartDate, tmAssignedSpotDates(ilVef).lEndDate
    Next ilVef
    
    mBuildDates llCrfCode
    
    'Clear rotation reference from cuf
    mClearCuf llCrfCode
    
    'Remove cvf
    slSQLQuery = "Delete From cvf_Copy_Vehicles Where cvfCrfCode = " & llCrfCode
    llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    
    'Remove cnf
    
    slSQLQuery = "Delete From cnf_Copy_Instruction Where cnfCrfCode = " & llCrfCode
    llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    
    'Remove caf
    slSQLQuery = "Delete From caf_Copy_Air_Game Where cafCrfCode = " & llCrfCode
    llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    
    'Remove crf
    slSQLQuery = "Delete From crf_Copy_Rot_Header Where CrfCode = " & llCrfCode
    llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
        
    gUpdateCifDates hmCif, tmCifInfo(), tmCrfDates(), llChgCount, llTotalCount, False
        
    mEraseRot = True
    Exit Function
Eraseerr:
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAvailsPop                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mAvailsPop()
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilLp As Integer
    Dim slStr As String
    ilIndex = lbcAvail.ListIndex
    If ilIndex > 0 Then
        slName = lbcAvail.List(ilIndex)
    End If
'    ilFilter(0) = NOFILTER
'    slFilter(0) = ""
'    ilOffset(0) = 0
    ilFilter(0) = CHARFILTERNOT
    slFilter(0) = "F"
    ilOffSet(0) = gFieldOffset("Anf", "AnfBookLocalFeed") '2
    'ilRet = gIMoveListBox(Copy, lbcAvail, lbcAvailCode, "Anf.Btr", gFieldOffset("Anf", "AnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(CopyRotErase, lbcAvail, tmAvailCode(), smAvailCodeTag, "Anf.Btr", gFieldOffset("Anf", "AnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        'Remove "Post Log" avail name
        For ilLoop = 0 To lbcAvail.ListCount - 1 Step 1
            slStr = Trim$(lbcAvail.List(ilLoop))
            If StrComp(slStr, "Post Log", 1) = 0 Then
                lbcAvail.RemoveItem ilLoop
                For ilLp = ilLoop To UBound(tmAvailCode) - 1 Step 1
                    tmAvailCode(ilLp) = tmAvailCode(ilLp + 1)
                Next ilLp
                ReDim Preserve tmAvailCode(LBound(tmAvailCode) To UBound(tmAvailCode) - 1) As SORTCODE
                Exit For
            End If
        Next ilLoop
    End If
    Exit Sub
mAvailsErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub


Private Sub mClearCuf(llCrfCode As Long) 'L.Bianchi 04/30/2021
    Dim ilCuf As Integer
    Dim ilAdfCode As Integer
    Dim slSQLQuery As String
    'Dim llRet As Integer
    Dim llRet As Long 'TTP10589 JJB 2022-12-14  Note:  Function gSQLWaitNoMsgBox returns a long value and when it did caused an unhandled exception
    
    ilAdfCode = rst_Crf!crfAdfCode
    slSQLQuery = "Select * From cuf_Copy_Usage"
    slSQLQuery = slSQLQuery & " Where cufAdfCode = " & ilAdfCode
    Set rst_Cuf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Cuf.EOF
        For ilCuf = 1 To 25 Step 1
            Select Case ilCuf
                Case 1
                    If rst_Cuf!CufCrfCode1 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode1 = 0, cufRotStartDate1= null, cufRotEndDate1 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                    End If
                Case 2
                    If rst_Cuf!cufCrfCode2 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode2 = 0, cufRotStartDate2= null, cufRotEndDate2 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 3
                    If rst_Cuf!CufCrfCode3 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode3 = 0, cufRotStartDate3= null, cufRotEndDate3 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 4
                    If rst_Cuf!CufCrfCode4 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode4 = 0, cufRotStartDate4= null, cufRotEndDate4 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 5
                    If rst_Cuf!CufCrfCode5 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode5 = 0, cufRotStartDate5= null, cufRotEndDate5 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 6
                    If rst_Cuf!CufCrfCode6 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode6 = 0, cufRotStartDate6= null, cufRotEndDate6 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 7
                    If rst_Cuf!CufCrfCode7 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode7 = 0, cufRotStartDate7= null, cufRotEndDate7 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 8
                    If rst_Cuf!CufCrfCode8 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode8 = 0, cufRotStartDate8= null, cufRotEndDate8 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 9
                    If rst_Cuf!CufCrfCode9 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode9 = 0, cufRotStartDate9= null, cufRotEndDate9 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 10
                    If rst_Cuf!CufCrfCode10 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode10 = 0, cufRotStartDate10= null, cufRotEndDate10 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 11
                    If rst_Cuf!CufCrfCode11 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode11 = 0, cufRotStartDate11= null, cufRotEndDate11 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 12
                    If rst_Cuf!CufCrfCode12 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode12 = 0, cufRotStartDate12= null, cufRotEndDate12 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 13
                    If rst_Cuf!CufCrfCode13 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode13 = 0, cufRotStartDate13= null, cufRotEndDate13 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 14
                    If rst_Cuf!CufCrfCode14 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode14 = 0, cufRotStartDate14= null, cufRotEndDate14 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 15
                    If rst_Cuf!CufCrfCode15 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode15 = 0, cufRotStartDate15= null, cufRotEndDate15 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 16
                    If rst_Cuf!CufCrfCode16 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode16 = 0, cufRotStartDate16= null, cufRotEndDate16 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 17
                    If rst_Cuf!CufCrfCode17 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode17 = 0, cufRotStartDate17= null, cufRotEndDate17 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 18
                    If rst_Cuf!CufCrfCode18 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode18, cufRotStartDate18= null, cufRotEndDate18 = null = 0"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 19
                    If rst_Cuf!CufCrfCode19 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode19 = 0, cufRotStartDate19= null, cufRotEndDate19 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 20
                    If rst_Cuf!CufCrfCode20 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode20 = 0, cufRotStartDate20= null, cufRotEndDate20 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 21
                    If rst_Cuf!CufCrfCode21 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode21 = 0, cufRotStartDate21= null, cufRotEndDate21 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 22
                    If rst_Cuf!CufCrfCode22 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode22 = 0, cufRotStartDate22= null, cufRotEndDate22 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 23
                    If rst_Cuf!CufCrfCode23 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode23 = 0, cufRotStartDate23= null, cufRotEndDate23 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 24
                    If rst_Cuf!CufCrfCode24 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode24 = 0, cufRotStartDate24= null, cufRotEndDate24 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
                Case 25
                    If rst_Cuf!CufCrfCode25 = llCrfCode Then
                        slSQLQuery = "Update cuf_Copy_Usage Set cufCrfCode25 = 0, cufRotStartDate25= null, cufRotEndDate25 = null"
                        slSQLQuery = slSQLQuery & " Where cufCode = " & rst_Cuf!cufCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        Exit For
                    End If
            End Select
        Next ilCuf
        rst_Cuf.MoveNext
    Loop

End Sub

Private Sub mBuildDates(llCrfCode As Long)
    Dim ilCuf As Integer
    Dim ilAdfCode As Integer
    Dim slSQLQuery As String
    'Dim llRet As Integer
    Dim llRet As Long 'TTP10589 JJB 2022-12-14
    Dim ilCif As Integer
    
    For ilCif = 0 To UBound(tmCifInfo) - 1 Step 1
        slSQLQuery = "Select * From cuf_Copy_Usage"
        slSQLQuery = slSQLQuery & " Where cufCifCode = " & tmCifInfo(ilCif).lCifCode
        Set rst_Cuf = gSQLSelectCall(slSQLQuery)
        Do While Not rst_Cuf.EOF
            For ilCuf = 1 To 25 Step 1
                Select Case ilCuf
                    Case 1
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode1, gDateValue(Format(rst_Cuf!cufRotStartDate1, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate1, "m/dd/yy")), ilCif
                    Case 2
                        mAddToDateArray llCrfCode, rst_Cuf!cufCrfCode2, gDateValue(Format(rst_Cuf!cufRotStartDate2, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate2, "m/dd/yy")), ilCif
                    Case 3
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode3, gDateValue(Format(rst_Cuf!cufRotStartDate3, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate3, "m/dd/yy")), ilCif
                    Case 4
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode4, gDateValue(Format(rst_Cuf!cufRotStartDate4, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate4, "m/dd/yy")), ilCif
                    Case 5
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode5, gDateValue(Format(rst_Cuf!cufRotStartDate5, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate5, "m/dd/yy")), ilCif
                    Case 6
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode6, gDateValue(Format(rst_Cuf!cufRotStartDate6, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate6, "m/dd/yy")), ilCif
                    Case 7
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode7, gDateValue(Format(rst_Cuf!cufRotStartDate7, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate7, "m/dd/yy")), ilCif
                    Case 8
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode8, gDateValue(Format(rst_Cuf!cufRotStartDate8, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate8, "m/dd/yy")), ilCif
                    Case 9
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode9, gDateValue(Format(rst_Cuf!cufRotStartDate9, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate9, "m/dd/yy")), ilCif
                    Case 10
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode10, gDateValue(Format(rst_Cuf!cufRotStartDate10, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate10, "m/dd/yy")), ilCif
                    Case 11
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode11, gDateValue(Format(rst_Cuf!cufRotStartDate11, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate11, "m/dd/yy")), ilCif
                    Case 12
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode12, gDateValue(Format(rst_Cuf!cufRotStartDate12, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate12, "m/dd/yy")), ilCif
                    Case 13
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode13, gDateValue(Format(rst_Cuf!cufRotStartDate13, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate13, "m/dd/yy")), ilCif
                    Case 14
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode14, gDateValue(Format(rst_Cuf!cufRotStartDate14, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate14, "m/dd/yy")), ilCif
                    Case 15
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode15, gDateValue(Format(rst_Cuf!cufRotStartDate15, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate15, "m/dd/yy")), ilCif
                    Case 16
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode16, gDateValue(Format(rst_Cuf!cufRotStartDate16, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate16, "m/dd/yy")), ilCif
                    Case 17
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode17, gDateValue(Format(rst_Cuf!cufRotStartDate17, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate17, "m/dd/yy")), ilCif
                    Case 18
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode18, gDateValue(Format(rst_Cuf!cufRotStartDate18, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate18, "m/dd/yy")), ilCif
                    Case 19
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode19, gDateValue(Format(rst_Cuf!cufRotStartDate19, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate19, "m/dd/yy")), ilCif
                    Case 20
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode20, gDateValue(Format(rst_Cuf!cufRotStartDate20, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate20, "m/dd/yy")), ilCif
                    Case 21
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode21, gDateValue(Format(rst_Cuf!cufRotStartDate21, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate21, "m/dd/yy")), ilCif
                    Case 22
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode22, gDateValue(Format(rst_Cuf!cufRotStartDate22, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate22, "m/dd/yy")), ilCif
                    Case 23
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode23, gDateValue(Format(rst_Cuf!cufRotStartDate23, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate23, "m/dd/yy")), ilCif
                    Case 24
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode24, gDateValue(Format(rst_Cuf!cufRotStartDate24, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate24, "m/dd/yy")), ilCif
                    Case 25
                        mAddToDateArray llCrfCode, rst_Cuf!CufCrfCode25, gDateValue(Format(rst_Cuf!cufRotStartDate25, "m/dd/yy")), gDateValue(Format(rst_Cuf!cufRotEndDate25, "m/dd/yy")), ilCif
                End Select
            Next ilCuf
            rst_Cuf.MoveNext
        Loop
    Next ilCif
    
End Sub

Private Sub mAddToDateArray(llCrfCode As Long, llCufCrfCode As Long, llStartDate As Long, llEndDate As Long, ilCif As Integer)
    'Dim ilCif As Integer
    Dim llCrf As Long
    Dim llCrfNext As Long
    Dim slSQLQuery As String
    Dim rst_Temp As ADODB.Recordset
    
    If llCufCrfCode > 0 And llCufCrfCode <> llCrfCode Then
        slSQLQuery = "Select * From crf_Copy_Rot_Header"
        slSQLQuery = slSQLQuery & " Where crfCode = " & llCufCrfCode
        Set rst_Temp = gSQLSelectCall(slSQLQuery)
        If Not rst_Temp.EOF Then
            llCrf = UBound(tmCrfDates)
            llCrfNext = tmCifInfo(ilCif).lFirst
            If rst_Temp!crfBkoutInstAdfCode <= 0 Then
                tmCrfDates(llCrf).iAdfCode = rst_Temp!crfAdfCode
            Else
                tmCrfDates(llCrf).iAdfCode = rst_Temp!crfBkoutInstAdfCode
            End If
            tmCrfDates(llCrf).lCrfCode = llCufCrfCode
            tmCrfDates(llCrf).lStartDate = llStartDate
            tmCrfDates(llCrf).lEndDate = llEndDate
            tmCrfDates(llCrf).lNext = llCrfNext
            tmCifInfo(ilCif).lFirst = llCrf
            ReDim Preserve tmCrfDates(0 To llCrf + 1) As CRFDATES
        End If
    End If

End Sub

Private Sub mSetCommands()
    Dim llRow As Long
                    
    cmcErase.Enabled = False
    For llRow = grdCopyRot.FixedRows To grdCopyRot.Rows - 1 Step 1
        If grdCopyRot.TextMatrix(llRow, CNTRINDEX) <> "" Then
            If grdCopyRot.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                cmcErase.Enabled = True
                Exit For
            End If
        End If
    Next llRow
End Sub
Private Sub mSetUnassignedDate(llStartDate As Long, llEndDate As Long, llEarliestDateUnassigned As Long, llLatestDateUnassigned As Long)
    If llEarliestDateUnassigned = -1 Then
        llEarliestDateUnassigned = llStartDate
        llLatestDateUnassigned = llEndDate
    Else
        If llStartDate < llEarliestDateUnassigned Then
            llEarliestDateUnassigned = llStartDate
        End If
        If llEndDate > llLatestDateUnassigned Then
            llLatestDateUnassigned = llEndDate
        End If
    End If
End Sub

Private Sub mSetAlerts(ilVefCode As Integer, llStartDate As Long, llInEndDate As Long)
    Dim ilVpf As Integer
    Dim llLastLogDate As Long
    Dim llNowDate As Long
    Dim llEndDate As Long
    Dim slLogDate As String
    Dim ilRet As Integer
    Dim llDate As Long
    
    ilVpf = gBinarySearchVpfPlus(ilVefCode)
    If ilVpf > 0 Then
        If (tgVpf(ilVpf).iLLD(0) <> 0) Or (tgVpf(ilVpf).iLLD(1) <> 0) Then
            gUnpackDateLong tgVpf(ilVpf).iLLD(0), tgVpf(ilVpf).iLLD(1), llLastLogDate
        Else
            llLastLogDate = -1
        End If
        If llInEndDate = 0 Then
            llEndDate = llLastLogDate
        Else
            llEndDate = llInEndDate
        End If
        llNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
        If (llEndDate <> -1) And (llStartDate <= llLastLogDate) And (llLastLogDate <> -1) Then
            For llDate = llStartDate To llEndDate Step 7
                If (llDate > llNowDate) And (llDate <= llLastLogDate) Then
                    slLogDate = Format$(llDate, "m/d/yy")
                    'ilRet = gAlertAdd("L", "C", 0, tmSdf.iVefCode, slLogDate)
                    ilRet = gAlertAdd("L", "C", 0, ilVefCode, slLogDate)
                End If
            Next llDate
        End If
    End If
End Sub
