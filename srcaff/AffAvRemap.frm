VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "resize32.ocx"
Begin VB.Form frmAvRemap 
   Caption         =   "Time Remap"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   Icon            =   "AffAvRemap.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   9030
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo &Previous Remap"
      Height          =   375
      Left            =   6465
      TabIndex        =   15
      Top             =   5310
      Width           =   2295
   End
   Begin VB.Timer tmcFill 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   8385
      Top             =   5325
   End
   Begin VB.ListBox lbcAvail 
      Height          =   3375
      Index           =   2
      ItemData        =   "AffAvRemap.frx":08CA
      Left            =   5580
      List            =   "AffAvRemap.frx":08CC
      TabIndex        =   11
      Top             =   1530
      Width           =   3180
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Agreements"
      Height          =   375
      Left            =   225
      TabIndex        =   12
      Top             =   5310
      Width           =   2190
   End
   Begin VB.Frame Frame2 
      Caption         =   "Se&lect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   210
      TabIndex        =   0
      Top             =   -15
      Width           =   8520
      Begin VB.TextBox txtModel 
         Height          =   285
         Left            =   7110
         TabIndex        =   5
         Top             =   540
         Width           =   1110
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   7110
         TabIndex        =   3
         Top             =   180
         Width           =   1110
      End
      Begin VB.ComboBox cboSelect 
         Height          =   315
         ItemData        =   "AffAvRemap.frx":08CE
         Left            =   135
         List            =   "AffAvRemap.frx":08D0
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   255
         Width           =   3555
      End
      Begin VB.Label Label3 
         Caption         =   "&Existing Programming Date"
         Height          =   270
         Left            =   4335
         TabIndex        =   4
         Top             =   585
         Width           =   2820
      End
      Begin VB.Label Label2 
         Caption         =   "&Start Date of New Programming"
         Height          =   270
         Left            =   4335
         TabIndex        =   2
         Top             =   225
         Width           =   2820
      End
   End
   Begin VB.ListBox lbcAvail 
      Height          =   3375
      Index           =   1
      ItemData        =   "AffAvRemap.frx":08D2
      Left            =   2865
      List            =   "AffAvRemap.frx":08D4
      TabIndex        =   9
      Top             =   1530
      Width           =   2190
   End
   Begin VB.ListBox lbcAvail 
      Height          =   3375
      Index           =   0
      ItemData        =   "AffAvRemap.frx":08D6
      Left            =   210
      List            =   "AffAvRemap.frx":08D8
      TabIndex        =   7
      Top             =   1530
      Width           =   2190
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   45
      Top             =   5280
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5820
      FormDesignWidth =   9030
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "&Auto Match"
      Height          =   375
      Left            =   2610
      TabIndex        =   13
      Top             =   5310
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   5310
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Remap"
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   1050
      Width           =   3165
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "New Times:"
      Height          =   465
      Left            =   2895
      TabIndex        =   8
      Top             =   1050
      Width           =   2130
   End
   Begin VB.Label lacTitle1 
      Alignment       =   2  'Center
      Caption         =   "Existing Times:"
      Height          =   465
      Left            =   240
      TabIndex        =   6
      Top             =   1050
      Width           =   2130
   End
End
Attribute VB_Name = "frmAvRemap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmAvRemap - allows for Agreement avails to be remapped
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private imVefCode As Integer
Private imInChg As Integer
Private imBSMode As Integer
Private smDate As String     'Effective Date
Private smCurDate As String
Private smCurTime As String
Private attrst As ADODB.Recordset
Private attrstmod As ADODB.Recordset
Private DATRST As ADODB.Recordset
Private EPTRST As ADODB.Recordset
Private imAvailGenerated As Integer
Private hmMsg As Integer
Private imVefCombo As Integer
Private lmAttCode() As Long
Private tmNewDat() As DAT
Private tmOldDat() As DAT
Private tmAvInfo() As AVINFO
Private tmUndoAvTime() As UNDOAVTIME

'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Zone Change Info               *
'*                                                     *
'*******************************************************
Private Function mOpenMsgFile(slToFile As String) As Integer
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer
    
    'On Error GoTo mOpenMsgFileErr:
    ilRet = 0
    slToFile = sgExportDirectory & "RM" & Format$(gNow(), "mm") & Format$(gNow(), "dd") & Format$(gNow(), "yy") & ".txt"
    slNowDate = Format$(gNow(), sgShowDateForm)
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        On Error GoTo 0
        'ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Append As hmMsg
        ilRet = gFileOpen(slToFile, "Append", hmMsg)
        If ilRet <> 0 Then
            Close hmMsg
            hmMsg = -1
            gMsgBox "Open File " & slToFile & " error #" & Str$(Err.Number), vbOKOnly
            mOpenMsgFile = False
            Exit Function
        End If
    Else
        On Error GoTo 0
        'ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Close hmMsg
            hmMsg = -1
            gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, "** Time Re-Map Info: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = 1
'    Resume Next
End Function

Private Sub mFillVehicle()
    Dim iLoop As Integer
    cboSelect.Clear
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            cboSelect.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            cboSelect.ItemData(cboSelect.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
End Sub

Private Sub cboSelect_Change()
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
    lbcAvail(0).Clear
    lbcAvail(1).Clear
    lbcAvail(2).Clear
    imAvailGenerated = False
    sName = LTrim$(cboSelect.Text)
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    lRow = SendMessageByString(cboSelect.hwnd, CB_FINDSTRING, -1, sName)
    If lRow >= 0 Then
        cboSelect.ListIndex = lRow
        cboSelect.SelStart = iLen
        cboSelect.SelLength = Len(cboSelect.Text)
        If cboSelect.ListIndex < 0 Then
            imVefCode = 0
        Else
            imVefCode = CInt(cboSelect.ItemData(cboSelect.ListIndex))
            If gIsDate(txtDate.Text) Or gIsDate(txtModel.Text) Then
                tmcFill.Enabled = True
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
    imInChg = False
    Exit Sub
End Sub

Private Sub cboSelect_Click()
    cboSelect_Change
End Sub

Private Sub cboSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboSelect.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub


Private Sub cboSelect_LostFocus()
    Dim slDate As String
    
    'slDate = txtDate.Text
    'If gIsDate(slDate) And (imVefCode > 0) Then
    '    tmcFill.Enabled = True
    'End If
End Sub

Private Sub cmdAuto_Click()
    Dim ilHour0 As Integer
    Dim ilHour1 As Integer
    Dim ilAvail0 As Integer
    Dim ilAvail1 As Integer
    Dim ilIndex0 As Integer
    Dim ilIndex1 As Integer
    Dim ilHour As Integer
    Dim ilFound As Integer
    Dim ilDay As Integer
    
    ilAvail0 = 0
    ilHour = 0
    lbcAvail(0).ListIndex = -1
    lbcAvail(1).ListIndex = -1
    Do
        ilFound = False
        For ilAvail0 = 0 To lbcAvail(0).ListCount - 1 Step 1
            ilIndex0 = lbcAvail(0).ItemData(ilAvail0)
            ilHour0 = Hour(tmOldDat(ilIndex0).sFdSTime)
            If ilHour = ilHour0 Then
                For ilAvail1 = 0 To lbcAvail(1).ListCount - 1 Step 1
                    ilIndex1 = lbcAvail(1).ItemData(ilAvail1)
                    ilHour1 = Hour(tmNewDat(ilIndex1).sFdSTime)
                    If ilHour = ilHour1 Then
                        For ilDay = 0 To 6 Step 1
                            If (tmOldDat(ilIndex0).iFdDay(ilDay) = 1) And (tmNewDat(ilIndex1).iFdDay(ilDay) = 1) Then
                                ilFound = True
                                lbcAvail(0).ListIndex = ilAvail0
                                lbcAvail(1).ListIndex = ilAvail1
                                Exit For
                            End If
                        Next ilDay
                        If ilFound Then
                            Exit For
                        End If
                    End If
                Next ilAvail1
                If ilFound Then
                    Exit For
                End If
            End If
        Next ilAvail0
        If Not ilFound Then
            ilHour = ilHour + 1
        End If
    Loop While ilHour <= 23
End Sub

Private Sub cmdCancel_Click()
    tmcFill.Enabled = False
    'txtDate.Text = ""
    Unload frmAvRemap
End Sub


Private Sub cmdUndo_Click()
    Dim slSDate As String
    Dim slEDate As String
    Dim slName As String
    Dim llAttCode As Long
    Dim ilShttCode As Integer
    Dim llAtt As Long
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    
    On Error GoTo ErrHand
    slSDate = txtDate.Text
    If Trim$(slSDate) = "" Then
        gMsgBox "Date must be selected." & Chr$(13) & Chr$(10) & "Please enter.", vbOKOnly
        txtDate.SetFocus
        Exit Sub
    End If
    If Not gIsDate(slSDate) Then
        gMsgBox "Date format not correct (m/d/yyyy)." & Chr$(13) & Chr$(10) & "Please enter correctly.", vbOKOnly
        txtDate.SetFocus
        Exit Sub
    End If
    If imVefCode <= 0 Then
        gMsgBox "Vehicle not selected." & Chr$(13) & Chr$(10) & "Please select.", vbOKOnly
        cboSelect.SetFocus
        Exit Sub
    End If
    If gIsDate(slSDate) And (imVefCode > 0) Then
        slName = Trim$(cboSelect.Text)
        slSDate = Format$(gObtainPrevMonday(slSDate), sgShowDateForm)
        slEDate = Format$(DateValue(gAdjYear(slSDate)) - 1, sgShowDateForm)
        ilRet = gMsgBox("This will remove Agreements and Posting for " & slName & " starting as of " & slSDate & " and change previously terminated Agreements on " & slEDate & " to TFN, Proceed", vbYesNo)
        If ilRet = vbNo Then
            Exit Sub
        End If
        DoEvents
        Screen.MousePointer = vbHourglass
        ReDim tmUndoAvTime(0 To 0) As UNDOAVTIME
        smCurDate = Format(gNow(), sgShowDateForm)
        smCurTime = Format(gNow(), sgShowTimeWSecForm)
        SQLQuery = "SELECT *"
        SQLQuery = SQLQuery + " FROM att"
        SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode
        SQLQuery = SQLQuery + " AND attOnAir = '" & Format$(slSDate, sgSQLDateForm) & "'" & ")"
        Set attrst = gSQLSelectCall(SQLQuery)
        While Not attrst.EOF
            'lmAttCode(UBound(lmAttCode)) = attrst!attCode
            'ReDim Preserve lmAttCode(0 To UBound(lmAttCode) + 1) As Integer
            ilFound = False
            For ilLoop = 0 To UBound(tmUndoAvTime) - 1 Step 1
                If (tmUndoAvTime(ilLoop).iShfCode = attrst!attshfCode) And (tmUndoAvTime(ilLoop).iVefCode = attrst!attvefCode) Then
                    ilFound = True
                    If DateValue(gAdjYear(attrst!attOffAir)) > DateValue(gAdjYear(tmUndoAvTime(ilLoop).sOffDate)) Then
                        tmUndoAvTime(ilLoop).lAttCode = attrst!attCode
                        tmUndoAvTime(ilLoop).sOffDate = Format$(gAdjYear(attrst!attOffAir), sgShowDateForm)
                    End If
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                tmUndoAvTime(UBound(tmUndoAvTime)).lAttCode = attrst!attCode
                tmUndoAvTime(UBound(tmUndoAvTime)).iShfCode = attrst!attshfCode
                tmUndoAvTime(UBound(tmUndoAvTime)).iVefCode = attrst!attvefCode
                tmUndoAvTime(UBound(tmUndoAvTime)).sOffDate = Format$(attrst!attOffAir, sgShowDateForm)
                ReDim Preserve tmUndoAvTime(0 To UBound(tmUndoAvTime) + 1) As UNDOAVTIME
            End If
            attrst.MoveNext
        Wend
        For llAtt = 0 To UBound(tmUndoAvTime) - 1 Step 1
            DoEvents
            ' JD 12-18-2006 Added new function to properly remove an agreement.
            If Not gDeleteAgreement(tmUndoAvTime(llAtt).lAttCode, "AffAvRemapLog.Txt") Then
                gLogMsg "FAIL: cmdUndo_Click - Unable to delete att code " & tmUndoAvTime(llAtt).lAttCode, "AffErrorLog.Txt", False
            End If
'            cnn.BeginTrans
'            SQLQuery = "DELETE FROM Ast WHERE (astAtfCode = " & tmUndoAvTime(llAtt).lattCode & ")"
'            cnn.Execute SQLQuery, rdExecDirect
'            SQLQuery = "DELETE FROM Cptt WHERE (cpttAtfCode = " & tmUndoAvTime(llAtt).lattCode & ")"
'            cnn.Execute SQLQuery, rdExecDirect
'            SQLQuery = "DELETE FROM dat WHERE (datAtfCode = " & tmUndoAvTime(llAtt).lattCode & ")"
'            cnn.Execute SQLQuery, rdExecDirect
'            SQLQuery = "DELETE FROM Att WHERE (AttCode = " & tmUndoAvTime(llAtt).lattCode & ")"
'            cnn.Execute SQLQuery, rdExecDirect
'            cnn.CommitTrans
        Next llAtt
        DoEvents
        'Get records not deleted
        ReDim lmAttCode(0 To 0) As Long
        SQLQuery = "SELECT *"
        SQLQuery = SQLQuery + " FROM att"
        SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode
        SQLQuery = SQLQuery + " AND attOnAir = '" & Format$(slSDate, sgSQLDateForm) & "'" & ")"
        Set attrst = gSQLSelectCall(SQLQuery)
        While Not attrst.EOF
            lmAttCode(UBound(lmAttCode)) = attrst!attCode
            ReDim Preserve lmAttCode(0 To UBound(lmAttCode) + 1) As Long
            attrst.MoveNext
        Wend
        For llAtt = 0 To UBound(lmAttCode) - 1 Step 1
            DoEvents
            ' JD 12-18-2006 Added new function to properly remove an agreement.
            If Not gDeleteAgreement(lmAttCode(llAtt), "AffAvRemap.Txt") Then
                gLogMsg "FAIL: cmdUndo_Click - Unable to delete att code " & lmAttCode(llAtt), "AffErrorLog.Txt", False
            End If
'            cnn.BeginTrans
'            SQLQuery = "DELETE FROM Ast WHERE (astAtfCode = " & lmAttCode(llAtt) & ")"
'            cnn.Execute SQLQuery, rdExecDirect
'            SQLQuery = "DELETE FROM Cptt WHERE (cpttAtfCode = " & lmAttCode(llAtt) & ")"
'            cnn.Execute SQLQuery, rdExecDirect
'            SQLQuery = "DELETE FROM dat WHERE (datAtfCode = " & lmAttCode(llAtt) & ")"
'            cnn.Execute SQLQuery, rdExecDirect
'            SQLQuery = "DELETE FROM Att WHERE (AttCode = " & lmAttCode(llAtt) & ")"
'            cnn.Execute SQLQuery, rdExecDirect
'            cnn.CommitTrans
        Next llAtt
        DoEvents
        ReDim lmAttCode(0 To 0) As Long
        SQLQuery = "SELECT *"
        SQLQuery = SQLQuery + " FROM att"
        SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode
        SQLQuery = SQLQuery + " AND attOffAir = '" & Format$(slEDate, sgSQLDateForm) & "'" & ")"
        Set attrst = gSQLSelectCall(SQLQuery)
        While Not attrst.EOF
            lmAttCode(UBound(lmAttCode)) = attrst!attCode
            ReDim Preserve lmAttCode(0 To UBound(lmAttCode) + 1) As Long
            attrst.MoveNext
        Wend
        For llAtt = 0 To UBound(lmAttCode) - 1 Step 1
            DoEvents
            SQLQuery = "SELECT *"
            SQLQuery = SQLQuery + " FROM att"
            SQLQuery = SQLQuery + " WHERE (attCode = " & lmAttCode(llAtt) & ")"
            Set attrst = gSQLSelectCall(SQLQuery)
            If Not attrst.EOF Then
                SQLQuery = "UPDATE att SET "
                ilFound = False
                For ilLoop = 0 To UBound(tmUndoAvTime) - 1 Step 1
                    If (tmUndoAvTime(ilLoop).iShfCode = attrst!attshfCode) And (tmUndoAvTime(ilLoop).iVefCode = attrst!attvefCode) Then
                        ilFound = True
                        SQLQuery = SQLQuery & "attOffAir = '" & Format$(gAdjYear(tmUndoAvTime(ilLoop).sOffDate), sgSQLDateForm) & "', "
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    SQLQuery = SQLQuery & "attOffAir = '" & Format$("12/31/2069", sgSQLDateForm) & "', "
                End If
                SQLQuery = SQLQuery & "attSentToXDSStatus = '" & "M" & "'"
                'SQLQuery = SQLQuery & "attDropDate = '" & slEDate & "'"
                SQLQuery = SQLQuery & " WHERE attCode = " & lmAttCode(llAtt)
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "AvaulRemap-cmdUndo"
                    Exit Sub
                End If
            End If
        Next llAtt
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AvRemap-cmdUndo"
    Exit Sub
End Sub

Private Sub cmdUpdate_Click()
    Dim slSDateNewProg As String
    Dim slEDateOldProg As String
    Dim slSExDate As String
    Dim slEExDate As String
    Dim llOldAttCode As Long
    Dim llModAttCode As Long
    Dim llAttCode As Long
    Dim ilShttCode As Integer
    Dim ilOld As Integer
    Dim ilNew As Integer
    Dim ilLoop As Integer
    Dim ilIndex0 As Integer
    Dim ilIndex1 As Integer
    Dim ilFound As Integer
    Dim slFdStTime As String
    Dim slFdEdTime As String
    Dim slPdStTime As String
    Dim slPdEdTime As String
    Dim ilVef As Integer
    Dim ilZone As Integer
    Dim ilTimeAdj As Integer
    Dim llFdTime As Long
    Dim ilDayMatch As Integer
    Dim slTime As String
    Dim ilTDay As Integer
    Dim ilLDay As Integer
    Dim llAtt As Long
    Dim ilRet As Integer
    Dim slVehName As String
    Dim slToFile As String
    Dim slVefName As String
    Dim slShfName As String
    Dim llTemp As Long
    ReDim ilAvDay(0 To 6) As Integer
    ReDim ilDay(0 To 6) As Integer
    Dim slSvSQLQuery As String
    Dim ilAirPlayNo As Integer
    Dim ilNoAirPlays As Integer
    Dim llDATCode As Long
    Dim llEPTCode As Long
    Dim blPledgeByAvails As Boolean
    Dim ilPos As Integer
    Dim slToName As String
    Dim slDrop As String
    Dim ilUpper As Integer
    Dim ilIdx As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStr As String
    Dim ilAgreeType As Integer
    Dim blRet As Boolean
    
    On Error GoTo ErrHand
    
    'ReDim lmAttCode(0 To 0) As Long
    Screen.MousePointer = vbHourglass
    slSDateNewProg = txtDate.Text
    If Trim$(slSDateNewProg) = "" Then
        Screen.MousePointer = vbDefault
        gMsgBox "New Date must be specified." & Chr$(13) & Chr$(10) & "Please enter.", vbOKOnly
        txtDate.SetFocus
        Exit Sub
    End If
    If Not gIsDate(slSDateNewProg) Then
        Screen.MousePointer = vbDefault
        gMsgBox "New Date format not correct (m/d/yyyy)." & Chr$(13) & Chr$(10) & "Please enter correctly.", vbOKOnly
        txtDate.SetFocus
        Exit Sub
    End If
    If Weekday(slSDateNewProg, vbSunday) <> vbMonday Then
        Screen.MousePointer = vbDefault
        gMsgBox "Date Must be a Monday", vbOKOnly
        txtDate.SetFocus
        Exit Sub
    End If
    slSExDate = txtModel.Text
    If Trim$(slSExDate) = "" Then
        Screen.MousePointer = vbDefault
        gMsgBox "Existing Date must be specified." & Chr$(13) & Chr$(10) & "Please enter.", vbOKOnly
        txtModel.SetFocus
        Exit Sub
    End If
    If Not gIsDate(slSExDate) Then
        Screen.MousePointer = vbDefault
        gMsgBox "Existing Date format not correct (m/d/yyyy)." & Chr$(13) & Chr$(10) & "Please enter correctly.", vbOKOnly
        txtModel.SetFocus
        Exit Sub
    End If
    If Weekday(slSExDate, vbSunday) <> vbMonday Then
        Screen.MousePointer = vbDefault
        gMsgBox "Date Must be a Monday", vbOKOnly
        txtModel.SetFocus
        Exit Sub
    End If
    If imVefCode <= 0 Then
        Screen.MousePointer = vbDefault
        gMsgBox "Vehicle not selected." & Chr$(13) & Chr$(10) & "Please select.", vbOKOnly
        cboSelect.SetFocus
        Exit Sub
    End If
    If lbcAvail(2).ListCount <= 0 Then
        Screen.MousePointer = vbDefault
        gMsgBox "Avails not matched." & Chr$(13) & Chr$(10) & "Please specify.", vbOKOnly
        Exit Sub
    End If
    If gIsDate(slSDateNewProg) And (imVefCode > 0) And (lbcAvail(2).ListCount > 0) Then
        slVehName = Trim$(cboSelect.Text)
        slSDateNewProg = Format$(gObtainPrevMonday(slSDateNewProg), sgShowDateForm)
        slEDateOldProg = gAdjYear(Format$(DateValue(slSDateNewProg) - 1, sgShowDateForm))
        Screen.MousePointer = vbDefault
        If lbcAvail(1).ListCount <= 0 Then
            ilRet = gMsgBox("This will terminate selected Agreements for " & slVehName & " as of " & slEDateOldProg & " and create new ones starting " & slSDateNewProg & ", Proceed", vbYesNo)
        Else
            ilRet = gMsgBox("This will terminate selected Agreements for " & slVehName & " as of " & slEDateOldProg & " and create new ones starting " & slSDateNewProg & " Unmapped Avails in 'New Times' Status Will Be Set As Live, Proceed", vbYesNo)
        End If
        If ilRet = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        ilRet = mOpenMsgFile(slToFile)
        If Not ilRet Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        ilPos = InStrRev(slToFile, "\", -1, vbTextCompare)
        If ilPos > 0 Then
            slToName = Mid$(slToFile, ilPos + 1)
        Else
            slToName = slToFile
        End If
        If lbcAvail(1).ListCount <= 0 Then   'New Times
            Print #hmMsg, "   Add Agreements for " & slVehName & " as of " & slSDateNewProg
        Else
            Print #hmMsg, "   Add Agreements for " & slVehName & " as of " & slSDateNewProg & " Avails Not Mapped Set to a Status of Live"
        End If
        Print #hmMsg, "   Terminated Agreements for " & slVehName & " as of " & slEDateOldProg
        DoEvents
        Print #hmMsg, ""
        Print #hmMsg, "      Remap Times"
        For ilLoop = 0 To lbcAvail(2).ListCount - 1 Step 1  'Remaped Times
            Print #hmMsg, "      " & lbcAvail(2).List(ilLoop)
        Next ilLoop
        If lbcAvail(1).ListCount > 0 Then   'New Times
            Print #hmMsg, ""
            Print #hmMsg, "      Avail Set to Status of Live"
            For ilLoop = 0 To lbcAvail(1).ListCount - 1 Step 1
                Print #hmMsg, "      " & lbcAvail(1).List(ilLoop)
            Next ilLoop
        End If
        Print #hmMsg, ""
        Print #hmMsg, ""

        smCurDate = Format(gNow(), sgShowDateForm)
        smCurTime = Format(gNow(), sgShowTimeWSecForm)
        tgRemapInfo.iVefCode = imVefCode
        tgRemapInfo.sStartDate = slSDateNewProg
        igOkToRemap = True
        
        'Handle selective remapping
        If lbcAvail(1).ListCount > 0 Then
            'This might be in error, it is referencing frmSelRemap and
            'frmSelRemap references this form
            ilRet = gPopSelRemap(frmSelRemap, lbcAvail(1).ListCount, 0)
            Screen.MousePointer = vbDefault
            frmSelRemap.Show vbModal
            Screen.MousePointer = vbHourglass
        Else
            'This might be in error, it is referencing frmSelRemap and
            'frmSelRemap references this form
            ilRet = mPopSelRemap(frmSelRemap, lbcAvail(1).ListCount, 0)
        End If
        Screen.MousePointer = vbHourglass
        If ilRet = False Then
            Print #hmMsg, ""
            Print #hmMsg, "** Time Re-Map Failed ii gPopSelRemap: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
            Print #hmMsg, "** No Re-mapping occured."
            Print #hmMsg, ""
            Close #hmMsg
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        'User hit Cancel button in frmSelRemap
        If igOkToRemap = False Then
            Print #hmMsg, ""
            Print #hmMsg, "** Time Re-map was Canceled By the User: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
            Print #hmMsg, "** No Re-mapping occured."
            Print #hmMsg, ""
            Close #hmMsg
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        'SQLQuery = "SELECT *"
        SQLQuery = "SELECT attCode"
        SQLQuery = SQLQuery + " FROM att"
        SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode
        SQLQuery = SQLQuery + " AND attOffAir >= '" & Format$(slSDateNewProg, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery + " AND attDropDate >= '" & Format$(slSDateNewProg, sgSQLDateForm) & "'" & ")"
        Set attrst = gSQLSelectCall(SQLQuery)

        'Build array of current header agreements for the selected vehicle
        ReDim lmAttCode(0 To 0) As Long
        While Not attrst.EOF
            lmAttCode(UBound(lmAttCode)) = attrst!attCode
            ReDim Preserve lmAttCode(0 To UBound(lmAttCode) + 1) As Long
            attrst.MoveNext
        Wend

        'Process each agreement
        'For llAtt = 0 To UBound(lmAttCode) - 1 Step 1'
        For llAtt = 0 To UBound(tgAttInfo) - 1 Step 1
            DoEvents
            SQLQuery = "SELECT *"
            SQLQuery = SQLQuery + " FROM att  "
            'SQLQuery = SQLQuery + " WHERE (attCode = " & lmAttCode(llAtt) & ")"
            SQLQuery = SQLQuery + " WHERE (attCode = " & tgAttInfo(llAtt).lAttCode & ")"
            Set attrst = gSQLSelectCall(SQLQuery)

            If Not attrst.EOF Then
                'SQLQuery = "SELECT *"
                SQLQuery = "SELECT attCode, attPledgeType"
                SQLQuery = SQLQuery + " FROM att"
                SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode
                SQLQuery = SQLQuery + " AND attShfCode = " & attrst!attshfCode
                SQLQuery = SQLQuery + " AND attOnAir <= '" & Format$(slSExDate, sgSQLDateForm) & "'"
                SQLQuery = SQLQuery + " AND attOffAir > '" & Format$(slSExDate, sgSQLDateForm) & "'"
                SQLQuery = SQLQuery + " AND attDropDate > '" & Format$(slSExDate, sgSQLDateForm) & "'" & ")"
                Set attrstmod = gSQLSelectCall(SQLQuery)
                If Not attrstmod.EOF Then
                    slVefName = ""
                    slShfName = ""
                    llModAttCode = attrstmod!attCode
                    llOldAttCode = attrst!attCode
                    ilShttCode = attrst!attshfCode

                    'Determine station zone
                    For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                        If tgStationInfo(ilLoop).iCode = ilShttCode Then
                            slShfName = tgStationInfo(ilLoop).sCallLetters
                            ilTimeAdj = 0
                            For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
                                If tgVehicleInfo(ilVef).iCode = imVefCode Then
                                    slVefName = tgVehicleInfo(ilVef).sVehicle
                                    For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                                        If StrComp(tgStationInfo(ilLoop).sZone, tgVehicleInfo(ilVef).sZone(ilZone), 1) = 0 Then
                                            ilTimeAdj = tgVehicleInfo(ilVef).iVehLocalAdj(ilZone)
                                            Exit For
                                        End If
                                    Next ilZone
                                    Exit For
                                End If
                            Next ilVef
                            Exit For
                        End If
                    Next ilLoop

                    'Determine is Pledge by Avails
                    blPledgeByAvails = gIsPledgeByAvails(llModAttCode)
                    'Insert Pledges with remap
                    SQLQuery = "SELECT * "
                    SQLQuery = SQLQuery + " FROM dat"
                    SQLQuery = SQLQuery + " WHERE (datAtfCode= " & llModAttCode
                    SQLQuery = SQLQuery + " AND datShfCode= " & ilShttCode
                    SQLQuery = SQLQuery + " AND datVefCode = " & imVefCode & ")"
                    SQLQuery = SQLQuery & " ORDER BY datFdStTime"
                    Set DATRST = gSQLSelectCall(SQLQuery)
                    DoEvents
                    If Not DATRST.EOF Then
                        If blPledgeByAvails Then
                            'Terminate current agreement
                            '01/14/20 D.S. Start TTP 5670
                            blRet = gCleanUPFiles(slSDateNewProg, llOldAttCode, imVefCode)
                            '01/14/20 D.S. End TTP 5670
                            SQLQuery = "UPDATE att SET "
                            SQLQuery = SQLQuery & "attOffAir = '" & Format$(slEDateOldProg, sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & "attDropDate = '" & Format$(slEDateOldProg, sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & "attSentToXDSStatus = '" & "M" & "'"
                            SQLQuery = SQLQuery & " WHERE attCode = " & llOldAttCode
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError slToName, "AvailRemap-cmdUpdate_Click"
                                Print #hmMsg, ""
                                Print #hmMsg, "** Time Re-Map Incomplete: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
                                Print #hmMsg, ""
                                Close #hmMsg
                                gMsgBox "See Status File " & slToFile, vbOKOnly
                                Exit Sub
                            End If
                            'D.S. 8/2/05
                            llTemp = gFindAttHole()
                            If llTemp = -1 Then
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            ilNoAirPlays = attrst!attNoAirPlays
                            If ilNoAirPlays <= 0 Then
                                ilNoAirPlays = 1
                            End If
                            'Insert new agreement
                            SQLQuery = "INSERT INTO att(attCode, attShfCode, attVefCode, attAgreeStart, "
                            SQLQuery = SQLQuery & "attAgreeEnd, attOnAir, attOffAir, attSigned, attSignDate, "
                            SQLQuery = SQLQuery & "attLoad, attTimeType, attComp, attBarCode, attDropDate, "
                            SQLQuery = SQLQuery & "attUsfCode, attEnterDate, attEnterTime, attNotice, "
                            SQLQuery = SQLQuery & "attCarryCmml, attNoCDs, attSendTape, attACName, "
                            SQLQuery = SQLQuery & "attACPhone, attGenLog, attGenCP, attPostingType, attPrintCP, "
                            SQLQuery = SQLQuery & "attExportType, attLogType, attPostType, attWebPW, attWebEmail, "
                            SQLQuery = SQLQuery & "attSendLogEMail, attSuppressNotice, attLabelID, attLabelShipInfo, "
                            SQLQuery = SQLQuery & "attComments, attGenOther, attStartTime, attMulticast, "
                            SQLQuery = SQLQuery & "attRadarClearType, attArttCode, attNCR, attFormerNCR, attForbidSplitLive, "
                            SQLQuery = SQLQuery & "attXDReceiverID, attVoiceTracked, attWebInterface, "
                            SQLQuery = SQLQuery & "attContractPrinted, "
                            SQLQuery = SQLQuery & "attMktRepUstCode, "
                            SQLQuery = SQLQuery & "attServRepUstCode, "
                            SQLQuery = SQLQuery & "attVehProgStartTime, "
                            SQLQuery = SQLQuery & "attVehProgEndTime, "
                            SQLQuery = SQLQuery & "attExportToWeb, "
                            SQLQuery = SQLQuery & "attExportToUnivision, "
                            SQLQuery = SQLQuery & "attExportToMarketron, "
                            SQLQuery = SQLQuery & "attExportToCBS, "
                            SQLQuery = SQLQuery & "attExportToClearCh, "
                            SQLQuery = SQLQuery & "attPledgeType, "
                            SQLQuery = SQLQuery & "attNoAirPlays, "
                            SQLQuery = SQLQuery & "attDesignVersion, "
                            SQLQuery = SQLQuery & "attIDCReceiverID, "
                            SQLQuery = SQLQuery & "attSentToXDSStatus, "
                            SQLQuery = SQLQuery & "attAudioDelivery, "
                            SQLQuery = SQLQuery & "attExportToJelli, "
                            '3/23/15: Add Send Delays to XDS
                            SQLQuery = SQLQuery & "attSendDelayToXDS, "
                            SQLQuery = SQLQuery & "attServiceAgreement, "
                            '4/3/19
                            SQLQuery = SQLQuery & "attExcludeFillSpot, "
                            SQLQuery = SQLQuery & "attExcludeCntrTypeQ, "
                            SQLQuery = SQLQuery & "attExcludeCntrTypeR, "
                            SQLQuery = SQLQuery & "attExcludeCntrTypeT, "
                            SQLQuery = SQLQuery & "attExcludeCntrTypeM, "
                            SQLQuery = SQLQuery & "attExcludeCntrTypeS, "
                            SQLQuery = SQLQuery & "attExcludeCntrTypeV, "

                            SQLQuery = SQLQuery & "attUnused "
                            SQLQuery = SQLQuery & ")"

                            SQLQuery = SQLQuery & " VALUES"
                            SQLQuery = SQLQuery & "(" & llTemp & ", " & ilShttCode & ", " & imVefCode & ", '" & Format$(attrst!attAgreeStart, sgSQLDateForm) & "', '" & Format$(attrst!attAgreeEnd, sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & "'" & Format$(slSDateNewProg, sgSQLDateForm) & "', '" & Format$(attrst!attOffAir, sgSQLDateForm) & "', " & attrst!attSigned & ", "
                            SQLQuery = SQLQuery & "'" & Format$(attrst!attSignDate, sgSQLDateForm) & "', " & attrst!attLoad & ", " & attrst!attTimeType & ", "
                            SQLQuery = SQLQuery & attrst!attComp & ", " & attrst!attBarCode & ", '" & Format$(attrst!attDropDate, sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & igUstCode & ", '" & Format$(smCurDate, sgSQLDateForm) & "', '" & Format$(smCurTime, sgSQLTimeForm) & "', '" & gFixQuote(attrst!attNotice) & "', "
                            SQLQuery = SQLQuery & attrst!attCarryCmml & ", " & attrst!attNoCDs & ", " & attrst!attSendTape & ", '" & gFixQuote(attrst!attACName) & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attACPhone & "', '" & attrst!attGenLog & "', '" & attrst!attGenCP & "', " & attrst!attPostingType & ", " & attrst!attPrintCP & ", "
                            SQLQuery = SQLQuery & attrst!attExportType & ", " & attrst!attLogType & ", " & attrst!attPostType & ", '" & gFixQuote(attrst!attWebPW) & "', '" & gFixQuote(attrst!attWebEmail) & "', "
                            'SQLQuery = SQLQuery & attrst!attSendLogEmail & ", '" & attrst!attSuppressNotice & "', '" & attrst!attLabelID & "', '" & attrst!attLabelShipInfo & "', '" & attrst!attComments & "', '" & attrst!attGenOther & "', '" & Format$(attrst!attStartTime, sgSQLTimeForm) & "', '" & attrst!attMulticast & "', '" & attrst!attWebInterface & "', "
                            SQLQuery = SQLQuery & attrst!attSendLogEmail & ", '" & attrst!attSuppressNotice & "', '" & attrst!attLabelID & "', '" & attrst!attLabelShipInfo & "', '" & gFixQuote(attrst!attComments) & "', '" & attrst!attGenOther & "', '" & Format$(attrst!attStartTime, sgSQLTimeForm) & "', '" & attrst!attMulticast & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attRadarClearType & "', " & attrst!attArttCode & ", '" & attrst!attNCR & "' , '" & attrst!attFormerNCR & "',  '" & attrst!attForbidSplitLive & "', "
                            SQLQuery = SQLQuery & attrst!attXDReceiverId & ", '" & attrst!attVoiceTracked & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attWebInterface & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attContractPrinted & "', "
                            SQLQuery = SQLQuery & attrst!attMktRepUstCode & ", "
                            SQLQuery = SQLQuery & attrst!attServRepUstCode & ", "
                            SQLQuery = SQLQuery & "'" & Format$(attrst!attVehProgStartTime, sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & "'" & Format$(attrst!attVehProgEndTime, sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attExportToWeb & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attExportToUnivision & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attExportToMarketron & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attExportToCBS & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attExportToClearCh & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attPledgeType & "', "
                            SQLQuery = SQLQuery & ilNoAirPlays & ", "
                            SQLQuery = SQLQuery & attrst!attDesignVersion & ", "
                            SQLQuery = SQLQuery & "'" & attrst!attIDCReceiverID & "', "
                            SQLQuery = SQLQuery & "'" & "M" & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attAudioDelivery & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attExportToJelli & "', "
                            '3/23/15: Add Send Delays to XDS
                            SQLQuery = SQLQuery & "'" & attrst!attSendDelayToXDS & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attServiceAgreement & "', "
                            '4-3-19
                            SQLQuery = SQLQuery & "'" & attrst!attExcludeFillSpot & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attExcludeCntrTypeQ & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attExcludeCntrTypeR & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attExcludeCntrTypeT & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attExcludeCntrTypeM & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attExcludeCntrTypeS & "', "
                            SQLQuery = SQLQuery & "'" & attrst!attExcludeCntrTypeV & "', "

                            SQLQuery = SQLQuery & "'" & "" & "'"
                            SQLQuery = SQLQuery & ")"
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError slToName, "AvailRemap-cmdUpdate_Click"
                                'cnn.RollbackTrans
                                Print #hmMsg, ""
                                Print #hmMsg, "** Time Re-Map Incomplete: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
                                Print #hmMsg, ""
                                Close #hmMsg
                                gMsgBox "See Status File " & slToFile, vbOKOnly
                                Exit Sub
                            End If
                            If llTemp = 0 Then
                                SQLQuery = "Select MAX(attCode) from att"
                                Set rst = gSQLSelectCall(SQLQuery)
                                llAttCode = rst(0).Value
                            Else
                                llAttCode = llTemp
                            End If
                            '7701
                            SQLQuery = "insert into VAT_Vendor_Agreement (vatAttCode,vatWvtVendorId) ( select " & llAttCode & ", vatWvtVendorId from VAT_Vendor_Agreement where vatAttCode = " & llOldAttCode & ")"
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError slToName, "AvailRemap-cmdUpdate_Click"
                                'cnn.RollbackTrans
                                Print #hmMsg, ""
                                Print #hmMsg, "** Time Re-Map Incomplete: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
                                Print #hmMsg, ""
                                Close #hmMsg
                                gMsgBox "See Status File " & slToFile, vbOKOnly
                                Exit Sub
                            End If
                            'Handle the newly remapped times
                            While Not DATRST.EOF
                                For ilLoop = 0 To lbcAvail(2).ListCount - 1 Step 1
                                    DoEvents
                                    ilIndex0 = tmAvInfo(lbcAvail(2).ItemData(ilLoop)).iIndex0
                                    llFdTime = gTimeToLong(tmOldDat(ilIndex0).sFdSTime, False) + 3600 * ilTimeAdj
                                    For ilLDay = 0 To 6 Step 1
                                        ilAvDay(ilLDay) = tmAvInfo(lbcAvail(2).ItemData(ilLoop)).iDay(ilLDay)
                                    Next ilLDay
                                    If llFdTime < 0 Then
                                        llFdTime = llFdTime + 86400
                                        'If iDay = vbSunday Then
                                        '    iDay = vbSaturday
                                        'Else
                                        '    iDay = iDay - 1
                                        'End If
                                        ilTDay = ilAvDay(0)
                                        For ilLDay = 0 To 5 Step 1
                                            ilAvDay(ilLDay) = ilAvDay(ilLDay + 1)
                                        Next ilLDay
                                        ilAvDay(6) = ilTDay
                                    ElseIf llFdTime > 86400 Then
                                        llFdTime = llFdTime - 86400
                                        'If iDay = vbSaturday Then
                                        '    iDay = vbSunday
                                        'Else
                                        '    iDay = iDay + 1
                                        'End If
                                        ilTDay = ilAvDay(6)
                                        For ilLDay = 5 To 0 Step -1
                                            ilAvDay(ilLDay + 1) = ilAvDay(ilLDay)
                                        Next ilLDay
                                        ilAvDay(0) = ilTDay
                                    End If


                                    slTime = Format$(DATRST!datFdStTime, "h:m:ssam/pm")
                                    If gTimeToLong(slTime, False) = llFdTime Then
                                        ilIndex1 = tmAvInfo(lbcAvail(2).ItemData(ilLoop)).iIndex1
                                        ilDayMatch = False
                                        For ilLDay = 0 To 6 Step 1
                                            ilDay(ilLDay) = 0
                                        Next ilLDay
                                        If (DATRST!datFdMon = 1) And (ilAvDay(0) = 1) Then
                                            ilDayMatch = True
                                            ilDay(0) = 1
                                        End If
                                        If (DATRST!datFdTue = 1) And (ilAvDay(1) = 1) Then
                                            ilDayMatch = True
                                            ilDay(1) = 1
                                        End If
                                        If (DATRST!datFdWed = 1) And (ilAvDay(2) = 1) Then
                                            ilDayMatch = True
                                            ilDay(2) = 1
                                        End If
                                        If (DATRST!datFdThu = 1) And (ilAvDay(3) = 1) Then
                                            ilDayMatch = True
                                            ilDay(3) = 1
                                        End If
                                        If (DATRST!datFdFri = 1) And (ilAvDay(4) = 1) Then
                                            ilDayMatch = True
                                            ilDay(4) = 1
                                        End If
                                        If (DATRST!datFdSat = 1) And (ilAvDay(5) = 1) Then
                                            ilDayMatch = True
                                            ilDay(5) = 1
                                        End If
                                        If (DATRST!datFdSun = 1) And (ilAvDay(6) = 1) Then
                                            ilDayMatch = True
                                            ilDay(6) = 1
                                        End If
                                        If ilDayMatch Then
                                            llFdTime = gTimeToLong(tmNewDat(ilIndex1).sFdSTime, False) + 3600 * ilTimeAdj
                                            If llFdTime < 0 Then
                                                llFdTime = llFdTime + 86400
                                            ElseIf llFdTime > 86400 Then
                                                llFdTime = llFdTime - 86400
                                            End If
                                            slFdStTime = Format$(gLongToTime(llFdTime), "hh:mm:ss")
                                            llFdTime = gTimeToLong(tmNewDat(ilIndex1).sFdETime, False) + 3600 * ilTimeAdj
                                            If llFdTime < 0 Then
                                                llFdTime = llFdTime + 86400
                                            ElseIf llFdTime > 86400 Then
                                                llFdTime = llFdTime - 86400
                                            End If
                                            slFdEdTime = Format$(gLongToTime(llFdTime), "hh:mm:ss")
                                            If DATRST!datFdStatus = 0 Then
                                                slPdStTime = slFdStTime
                                                slPdEdTime = slFdEdTime
                                                'SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, datDACode, "
                                                SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, "
                                                SQLQuery = SQLQuery & "datFdMon, datFdTue, datFdWed, datFdThu, "
                                                SQLQuery = SQLQuery & "datFdFri, datFdSat, datFdSun, datFdStTime, datFdEdTime, datFdStatus, "
                                                SQLQuery = SQLQuery & "datPdMon, datPdTue, datPdWed, datPdThu, datPdFri, "
                                                SQLQuery = SQLQuery & "datPdSat, datPdSun, datPdDayFed, datPdStTime, datPdEdTime, datAirPlayNo, datEstimatedTime)"
                                                SQLQuery = SQLQuery & " VALUES (" & "Replace" & ", " & llAttCode & ", " & ilShttCode & ", " & imVefCode
                                                SQLQuery = SQLQuery & "," '& DATRST!datDACode & ","
                                                SQLQuery = SQLQuery & ilDay(0) & ", " & ilDay(1) & ","
                                                SQLQuery = SQLQuery & ilDay(2) & ", " & ilDay(3) & ","
                                                SQLQuery = SQLQuery & ilDay(4) & ", " & ilDay(5) & ","
                                                SQLQuery = SQLQuery & ilDay(6) & ", "
                                                SQLQuery = SQLQuery & "'" & Format$(slFdStTime, sgSQLTimeForm) & "','" & Format$(slFdEdTime, sgSQLTimeForm) & "',"

                                                SQLQuery = SQLQuery & DATRST!datFdStatus & ","

                                                SQLQuery = SQLQuery & ilDay(0) & ", " & ilDay(1) & ","
                                                SQLQuery = SQLQuery & ilDay(2) & ", " & ilDay(3) & ","
                                                SQLQuery = SQLQuery & ilDay(4) & ", " & ilDay(5) & ","
                                                SQLQuery = SQLQuery & ilDay(6) & ", "
                                                SQLQuery = SQLQuery & "'" & DATRST!datPdDayFed & "',"
                                                'SQLQuery = SQLQuery & "'" & Format$(slPdStTime, sgSQLTimeForm) & "','" & Format$(slPdEdTime, sgSQLTimeForm) & "')"
                                                SQLQuery = SQLQuery & "'" & Format$(slPdStTime, sgSQLTimeForm) & "','" & Format$(slPdEdTime, sgSQLTimeForm) & "',"
                                                SQLQuery = SQLQuery & DATRST!datAirPlayNo & ", "
                                                SQLQuery = SQLQuery & "'" & DATRST!datEstimatedTime & "')"
                                            Else
                                                slPdStTime = Format$(DATRST!datPdStTime, "hh:mm:ss")
                                                slPdEdTime = Format$(DATRST!datPdEdTime, "hh:mm:ss")
                                                'SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, datDACode, "
                                                SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, "
                                                SQLQuery = SQLQuery & "datFdMon, datFdTue, datFdWed, datFdThu, "
                                                SQLQuery = SQLQuery & "datFdFri, datFdSat, datFdSun, datFdStTime, datFdEdTime, datFdStatus, "
                                                SQLQuery = SQLQuery & "datPdMon, datPdTue, datPdWed, datPdThu, datPdFri, "
                                                SQLQuery = SQLQuery & "datPdSat, datPdSun, datPdDayFed, datPdStTime, datPdEdTime, datAirPlayNo, datEstimatedTime)"
                                                SQLQuery = SQLQuery & " VALUES (" & "Replace" & ", " & llAttCode & ", " & ilShttCode & ", " & imVefCode
                                                SQLQuery = SQLQuery & "," '& DATRST!datDACode & ","
                                                'SQLQuery = SQLQuery & datrst!datFdMon & ", " & datrst!datFdTue & ","
                                                'SQLQuery = SQLQuery & datrst!datFdWed & ", " & datrst!datFdThu & ","
                                                'SQLQuery = SQLQuery & datrst!datFdFri & ", " & datrst!datFdSat & ","
                                                'SQLQuery = SQLQuery & datrst!datFdSun & ", "
                                                SQLQuery = SQLQuery & ilDay(0) & ", " & ilDay(1) & ","
                                                SQLQuery = SQLQuery & ilDay(2) & ", " & ilDay(3) & ","
                                                SQLQuery = SQLQuery & ilDay(4) & ", " & ilDay(5) & ","
                                                SQLQuery = SQLQuery & ilDay(6) & ", "
                                                SQLQuery = SQLQuery & "'" & Format$(slFdStTime, sgSQLTimeForm) & "','" & Format$(slFdEdTime, sgSQLTimeForm) & "',"

                                                SQLQuery = SQLQuery & DATRST!datFdStatus & ","

                                                SQLQuery = SQLQuery & DATRST!datPdMon & ", " & DATRST!datPdTue & ","
                                                SQLQuery = SQLQuery & DATRST!datPdWed & ", " & DATRST!datPdThu & ","
                                                SQLQuery = SQLQuery & DATRST!datPdFri & ", " & DATRST!datPdSat & ","
                                                SQLQuery = SQLQuery & DATRST!datPdSun & ", "
                                                SQLQuery = SQLQuery & "'" & DATRST!datPdDayFed & "',"
                                                'SQLQuery = SQLQuery & "'" & Format$(slPdStTime, sgSQLTimeForm) & "','" & Format$(slPdEdTime, sgSQLTimeForm) & "')"
                                                SQLQuery = SQLQuery & "'" & Format$(slPdStTime, sgSQLTimeForm) & "','" & Format$(slPdEdTime, sgSQLTimeForm) & "',"
                                                SQLQuery = SQLQuery & DATRST!datAirPlayNo & ", "
                                                SQLQuery = SQLQuery & "'" & DATRST!datEstimatedTime & "')"
                                            End If
                                            'cnn.Execute SQLQuery, rdExecDirect
                                            'If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                            llDATCode = gInsertAndReturnCode(SQLQuery, "dat", "datCode", "Replace")
                                            If llDATCode > 0 Then
                                                'Create Estimates if required
                                                If DATRST!datEstimatedTime = "Y" Then
                                                    SQLQuery = "SELECT * FROM ept"
                                                    SQLQuery = SQLQuery + " WHERE (eptDatCode = " & DATRST!datCode & ")"
                                                    Set EPTRST = gSQLSelectCall(SQLQuery)
                                                    Do While Not EPTRST.EOF
                                                        SQLQuery = "Insert Into ept ( "
                                                        SQLQuery = SQLQuery & "eptCode, "
                                                        SQLQuery = SQLQuery & "eptDatCode, "
                                                        SQLQuery = SQLQuery & "eptSeqNo, "
                                                        SQLQuery = SQLQuery & "eptAttCode, "
                                                        SQLQuery = SQLQuery & "eptShttCode, "
                                                        SQLQuery = SQLQuery & "eptVefCode, "
                                                        SQLQuery = SQLQuery & "eptFdAvailDay, "
                                                        SQLQuery = SQLQuery & "eptFdAvailTime, "
                                                        SQLQuery = SQLQuery & "eptEstimatedDay, "
                                                        SQLQuery = SQLQuery & "eptEstimatedTime, "
                                                        SQLQuery = SQLQuery & "eptUnused "
                                                        SQLQuery = SQLQuery & ") "
                                                        SQLQuery = SQLQuery & "Values ( "
                                                        SQLQuery = SQLQuery & "Replace" & ", "
                                                        SQLQuery = SQLQuery & llDATCode & ", "
                                                        SQLQuery = SQLQuery & EPTRST!eptSeqNo & ", "
                                                        SQLQuery = SQLQuery & llAttCode & ", "
                                                        SQLQuery = SQLQuery & ilShttCode & ", "
                                                        SQLQuery = SQLQuery & imVefCode & ", "
                                                        SQLQuery = SQLQuery & "'" & gFixQuote(EPTRST!eptFdAvailDay) & "', "
                                                        SQLQuery = SQLQuery & "'" & Format$(slFdStTime, sgSQLTimeForm) & "', "
                                                        If Trim$(EPTRST!eptEstimatedDay) <> "" Then
                                                            SQLQuery = SQLQuery & "'" & gFixQuote(EPTRST!eptEstimatedDay) & "', "
                                                            SQLQuery = SQLQuery & "'" & Format$(EPTRST!eptEstimatedTime, sgSQLTimeForm) & "', "
                                                        Else
                                                            SQLQuery = SQLQuery & "'" & "" & "', "
                                                            SQLQuery = SQLQuery & "'" & Format$("12AM", sgSQLTimeForm) & "', "
                                                        End If
                                                        SQLQuery = SQLQuery & "'" & "" & "' "
                                                        SQLQuery = SQLQuery & ") "
                                                        llEPTCode = gInsertAndReturnCode(SQLQuery, "ept", "eptCode", "Replace")
                                                        If llEPTCode <= 0 Then
                                                            '6/10/16: Replaced GoSub
                                                            'GoSub ErrHand:
                                                            Screen.MousePointer = vbDefault
                                                            gHandleError slToName, "AvailRemap-cmdUpdate_Click"
                                                            'cnn.RollbackTrans
                                                            Print #hmMsg, ""
                                                            Print #hmMsg, "** Time Re-Map Incomplete: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
                                                            Print #hmMsg, ""
                                                            Close #hmMsg
                                                            gMsgBox "See Status File " & slToFile, vbOKOnly
                                                            Exit Sub
                                                        End If
                                                        EPTRST.MoveNext
                                                    Loop
                                                End If
                                            Else
                                                '6/10/16: Replaced GoSub
                                                'GoSub ErrHand:
                                                Screen.MousePointer = vbDefault
                                                gHandleError slToName, "AvailRemap-cmdUpdate_Click"
                                                'cnn.RollbackTrans
                                                Print #hmMsg, ""
                                                Print #hmMsg, "** Time Re-Map Incomplete: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
                                                Print #hmMsg, ""
                                                Close #hmMsg
                                                gMsgBox "See Status File " & slToFile, vbOKOnly
                                                Exit Sub
                                            End If
                                            Exit For
                                        End If
                                    End If
                                Next ilLoop
                                DATRST.MoveNext
                            Wend

                            'Add all remaining items in lbcAvail(1) as live
                            For ilLoop = 0 To lbcAvail(1).ListCount - 1 Step 1
                                DoEvents
                                ilIndex1 = lbcAvail(1).ItemData(ilLoop)
                                slFdStTime = Format$(tmNewDat(ilIndex1).sFdSTime, "hh:mm:ss")
                                slFdEdTime = Format$(tmNewDat(ilIndex1).sFdETime, "hh:mm:ss")
                                slPdStTime = slFdStTime
                                slPdEdTime = slFdEdTime
                                'SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, datDACode, "
                                SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, "
                                SQLQuery = SQLQuery & "datFdMon, datFdTue, datFdWed, datFdThu, "
                                SQLQuery = SQLQuery & "datFdFri, datFdSat, datFdSun, datFdStTime, datFdEdTime, datFdStatus, "
                                SQLQuery = SQLQuery & "datPdMon, datPdTue, datPdWed, datPdThu, datPdFri, "
                                SQLQuery = SQLQuery & "datPdSat, datPdSun, datPdDayFed, datPdStTime, datPdEdTime, datAirPlayNo, datEstimatedTime)"
                                SQLQuery = SQLQuery & " VALUES (" & 0 & ", " & llAttCode & ", " & ilShttCode & ", " & imVefCode
                                SQLQuery = SQLQuery & "," '& 1 & ","
                                SQLQuery = SQLQuery & tmNewDat(ilIndex1).iFdDay(0) & ", " & tmNewDat(ilIndex1).iFdDay(1) & ","
                                SQLQuery = SQLQuery & tmNewDat(ilIndex1).iFdDay(2) & ", " & tmNewDat(ilIndex1).iFdDay(3) & ","
                                SQLQuery = SQLQuery & tmNewDat(ilIndex1).iFdDay(4) & ", " & tmNewDat(ilIndex1).iFdDay(5) & ","
                                SQLQuery = SQLQuery & tmNewDat(ilIndex1).iFdDay(6) & ", "
                                SQLQuery = SQLQuery & "'" & Format$(slFdStTime, sgSQLTimeForm) & "','" & Format$(slFdEdTime, sgSQLTimeForm) & "',"
                                If tgAttInfo(llAtt).iSelected Then
                                    SQLQuery = SQLQuery & 0 & ","
                                Else
                                    SQLQuery = SQLQuery & 8 & ","
                                End If
                                SQLQuery = SQLQuery & tmNewDat(ilIndex1).iFdDay(0) & ", " & tmNewDat(ilIndex1).iFdDay(1) & ","
                                SQLQuery = SQLQuery & tmNewDat(ilIndex1).iFdDay(2) & ", " & tmNewDat(ilIndex1).iFdDay(3) & ","
                                SQLQuery = SQLQuery & tmNewDat(ilIndex1).iFdDay(4) & ", " & tmNewDat(ilIndex1).iFdDay(5) & ","
                                SQLQuery = SQLQuery & tmNewDat(ilIndex1).iFdDay(6) & ", "
                                'SQLQuery = SQLQuery & "'A', " & "'" & Format$(slPdStTime, sgSQLTimeForm) & "','" & Format$(slPdEdTime, sgSQLTimeForm) & "')"
                                SQLQuery = SQLQuery & "'A', " & "'" & Format$(slPdStTime, sgSQLTimeForm) & "','" & Format$(slPdEdTime, sgSQLTimeForm) & "',"
                                'cnn.Execute SQLQuery, rdExecDirect
                                slSvSQLQuery = SQLQuery
                                For ilAirPlayNo = 1 To ilNoAirPlays Step 1
                                    SQLQuery = slSvSQLQuery
                                    SQLQuery = SQLQuery & ilAirPlayNo & ", "
                                    SQLQuery = SQLQuery & "'" & "N" & "')"
                                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                        '6/10/16: Replaced GoSub
                                        'GoSub ErrHand:
                                        Screen.MousePointer = vbDefault
                                        gHandleError slToName, "AvailRemap-cmdUpdate_Click"
                                        'cnn.RollbackTrans
                                        Print #hmMsg, ""
                                        Print #hmMsg, "** Time Re-Map Incomplete: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
                                        Print #hmMsg, ""
                                        Close #hmMsg
                                        gMsgBox "See Status File " & slToFile, vbOKOnly
                                        Exit Sub
                                    End If
                                Next ilAirPlayNo
                            Next ilLoop
                            'cnn.CommitTrans
                            Print #hmMsg, "      " & slShfName & " For " & slVefName & " Re-Mapped"
                        End If
                    End If
                End If
            End If
        Next llAtt
        Print #hmMsg, ""
        Print #hmMsg, "** Time Re-Map Completed: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
        Print #hmMsg, ""
        Close #hmMsg
        'Below terminates agreements that have been totally supercedeed (off air < on air)
        gCleanUpAtt
        gMsgBox "See Status File " & slToFile, vbOKOnly

        cmdCancel.Caption = "&Done"
        cmdUpdate.Enabled = False
        cmdAuto.Enabled = False
        cmdCancel.SetFocus
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    'cnn.RollbackTrans
    gHandleError "AffErrorLog.txt", "AvRemap-cmdUpdate"
    Print #hmMsg, ""
'Resume Next
    Print #hmMsg, gMsg
    Print #hmMsg, "** Time Re-Map Incomplete: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    Close #hmMsg
    gMsgBox "See Status File " & slToFile, vbOKOnly
    Exit Sub
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.25
    Me.Height = Screen.Height / 1.35
    Me.Top = (Screen.Height - Me.Height) / 1.5
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    frmAvRemap.Caption = "Time Remap - " & sgClientName
    'Me.Width = Screen.Width / 1.5
    'Me.Height = Screen.Height / 1.7
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    smDate = ""
    imInChg = False
    imBSMode = False
    imAvailGenerated = False
    mFillVehicle
    Screen.MousePointer = vbDefault
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase lmAttCode
    Erase tmNewDat
    Erase tmOldDat
    Erase tgDat
    Erase tmAvInfo
    Erase tmUndoAvTime
    Erase tgAttInfo
    attrst.Close
    DATRST.Close
    EPTRST.Close
    Unload frmSelRemap
    Set frmAvRemap = Nothing
    Set frmSelRemap = Nothing
End Sub


Private Sub lbcAvail_Click(Index As Integer)
    Dim slTime0 As String
    Dim slTime1 As String
    Dim ilIndex0 As Integer
    Dim ilIndex1 As Integer
    Dim ilFound As Integer
    Dim ilDay As Integer
    Dim slStr As String
    Dim slIndex As String
    Dim ilRet As Integer
    Dim slDays As String
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    
    If lbcAvail(Index).ListIndex >= 0 Then
        If Index <= 1 Then
            If (lbcAvail(0).ListIndex >= 0) And (lbcAvail(1).ListIndex >= 0) Then
                ilIndex0 = lbcAvail(0).ItemData(lbcAvail(0).ListIndex)
                slTime0 = tmOldDat(ilIndex0).sFdSTime
                ilIndex1 = lbcAvail(1).ItemData(lbcAvail(1).ListIndex)
                slTime1 = tmNewDat(ilIndex1).sFdSTime
                ilFound = False
                For ilDay = 0 To 6 Step 1
                    If (tmOldDat(ilIndex0).iFdDay(ilDay) = 1) And (tmNewDat(ilIndex1).iFdDay(ilDay) = 1) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilDay
                If Not ilFound Then
                    Beep
                    lbcAvail(Index).ListIndex = -1
                Else
                    'Get Days and turn days off
                    ilUpper = UBound(tmAvInfo)
                    slStr = ""
                    For ilDay = 0 To 6 Step 1
                        tmAvInfo(ilUpper).iDay(ilDay) = 0
                        If (tmOldDat(ilIndex0).iFdDay(ilDay) = 1) And (tmNewDat(ilIndex1).iFdDay(ilDay) = 1) Then
                            tmAvInfo(ilUpper).iDay(ilDay) = 1
                            Select Case ilDay
                                Case 0
                                    slStr = slStr + "Mo"
                                Case 1
                                    slStr = slStr + "Tu"
                                Case 2
                                    slStr = slStr + "We"
                                Case 3
                                    slStr = slStr + "Th"
                                Case 4
                                    slStr = slStr + "Fr"
                                Case 5
                                    slStr = slStr + "Sa"
                                Case 6
                                    slStr = slStr + "Su"
                            End Select
                            tmOldDat(ilIndex0).iFdDay(ilDay) = 0
                            tmNewDat(ilIndex1).iFdDay(ilDay) = 0
                        End If
                    Next ilDay
                    slDays = gDayMap(slStr)
                    lbcAvail(2).AddItem slTime0 & "->" & slTime1 & " " & slDays
                    lbcAvail(2).ItemData(lbcAvail(2).NewIndex) = ilUpper
                    tmAvInfo(ilUpper).iIndex0 = ilIndex0
                    tmAvInfo(ilUpper).iIndex1 = ilIndex1
                    ReDim Preserve tmAvInfo(0 To ilUpper + 1) As AVINFO
                    slStr = ""
                    For ilDay = 0 To 6 Step 1
                        If tmOldDat(ilIndex0).iFdDay(ilDay) = 1 Then
                            Select Case ilDay
                                Case 0
                                    slStr = slStr + "Mo"
                                Case 1
                                    slStr = slStr + "Tu"
                                Case 2
                                    slStr = slStr + "We"
                                Case 3
                                    slStr = slStr + "Th"
                                Case 4
                                    slStr = slStr + "Fr"
                                Case 5
                                    slStr = slStr + "Sa"
                                Case 6
                                    slStr = slStr + "Su"
                            End Select
                        End If
                    Next ilDay
                    If Len(slStr) <> 0 Then
                        slDays = gDayMap(slStr)
                        lbcAvail(0).List(lbcAvail(0).ListIndex) = tmOldDat(ilIndex0).sFdSTime & " " & slDays
                    Else
                        lbcAvail(0).RemoveItem (lbcAvail(0).ListIndex)
                    End If
                    slStr = ""
                    For ilDay = 0 To 6 Step 1
                        If tmNewDat(ilIndex1).iFdDay(ilDay) = 1 Then
                            Select Case ilDay
                                Case 0
                                    slStr = slStr + "Mo"
                                Case 1
                                    slStr = slStr + "Tu"
                                Case 2
                                    slStr = slStr + "We"
                                Case 3
                                    slStr = slStr + "Th"
                                Case 4
                                    slStr = slStr + "Fr"
                                Case 5
                                    slStr = slStr + "Sa"
                                Case 6
                                    slStr = slStr + "Su"
                            End Select
                        End If
                    Next ilDay
                    If Len(slStr) <> 0 Then
                        slDays = gDayMap(slStr)
                        lbcAvail(1).List(lbcAvail(1).ListIndex) = tmNewDat(ilIndex1).sFdSTime & " " & slDays
                    Else
                        lbcAvail(1).RemoveItem (lbcAvail(1).ListIndex)
                    End If
                End If
                lbcAvail(0).ListIndex = -1
                lbcAvail(1).ListIndex = -1
            End If
        Else
            ilIndex0 = tmAvInfo(lbcAvail(2).ItemData(lbcAvail(2).ListIndex)).iIndex0
            ilIndex1 = tmAvInfo(lbcAvail(2).ItemData(lbcAvail(2).ListIndex)).iIndex1
            'Set Days back on
            For ilDay = 0 To 6 Step 1
                If tmAvInfo(lbcAvail(2).ItemData(lbcAvail(2).ListIndex)).iDay(ilDay) = 1 Then
                    tmOldDat(ilIndex0).iFdDay(ilDay) = 1
                    tmNewDat(ilIndex1).iFdDay(ilDay) = 1
                End If
            Next ilDay
            slStr = ""
            For ilDay = 0 To 6 Step 1
                If tmOldDat(ilIndex0).iFdDay(ilDay) = 1 Then
                    Select Case ilDay
                        Case 0
                            slStr = slStr + "Mo"
                        Case 1
                            slStr = slStr + "Tu"
                        Case 2
                            slStr = slStr + "We"
                        Case 3
                            slStr = slStr + "Th"
                        Case 4
                            slStr = slStr + "Fr"
                        Case 5
                            slStr = slStr + "Sa"
                        Case 6
                            slStr = slStr + "Su"
                    End Select
                End If
            Next ilDay
            slDays = gDayMap(slStr)
            ilFound = False
            For ilLoop = 0 To lbcAvail(0).ListCount - 1 Step 1
                If ilIndex0 = lbcAvail(0).ItemData(ilLoop) Then
                    ilFound = True
                    lbcAvail(0).List(ilLoop) = tmOldDat(ilIndex0).sFdSTime & " " & slDays
                    Exit For
                ElseIf ilIndex0 < lbcAvail(0).ItemData(ilLoop) Then
                    ilFound = True
                    lbcAvail(0).AddItem tmOldDat(ilIndex0).sFdSTime & " " & slDays, ilLoop
                    lbcAvail(0).ItemData(lbcAvail(0).NewIndex) = ilIndex0
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                lbcAvail(0).AddItem tmOldDat(ilIndex0).sFdSTime & " " & slDays
                lbcAvail(0).ItemData(lbcAvail(0).NewIndex) = ilIndex0
            End If
            slStr = ""
            For ilDay = 0 To 6 Step 1
                If tmNewDat(ilIndex1).iFdDay(ilDay) = 1 Then
                    Select Case ilDay
                        Case 0
                            slStr = slStr + "Mo"
                        Case 1
                            slStr = slStr + "Tu"
                        Case 2
                            slStr = slStr + "We"
                        Case 3
                            slStr = slStr + "Th"
                        Case 4
                            slStr = slStr + "Fr"
                        Case 5
                            slStr = slStr + "Sa"
                        Case 6
                            slStr = slStr + "Su"
                    End Select
                End If
            Next ilDay
            slDays = gDayMap(slStr)
            'lbcAvail(1).List(ilIndex1) = tmNewDat(ilIndex1).sFdSTime & " " & slDays
            ilFound = False
            For ilLoop = 0 To lbcAvail(1).ListCount - 1 Step 1
                If ilIndex1 = lbcAvail(1).ItemData(ilLoop) Then
                    ilFound = True
                    lbcAvail(1).List(ilLoop) = tmNewDat(ilIndex1).sFdSTime & " " & slDays
                    Exit For
                ElseIf ilIndex1 < lbcAvail(1).ItemData(ilLoop) Then
                    ilFound = True
                    lbcAvail(1).AddItem tmNewDat(ilIndex1).sFdSTime & " " & slDays, ilLoop
                    lbcAvail(1).ItemData(lbcAvail(1).NewIndex) = ilIndex1
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                lbcAvail(1).AddItem tmNewDat(ilIndex1).sFdSTime & " " & slDays
                lbcAvail(1).ItemData(lbcAvail(1).NewIndex) = ilIndex1
            End If
            lbcAvail(2).RemoveItem (lbcAvail(2).ListIndex)
        End If
    End If
End Sub

Private Sub tmcFill_Timer()
    tmcFill.Enabled = False
    If imAvailGenerated = True Then
        Exit Sub
    End If
    mGetAvails
End Sub

Private Sub txtDate_Change()
    Dim slDate As String
    
    tmcFill.Enabled = False
    'lbcAvail(0).Clear
    lbcAvail(1).Clear
    lbcAvail(2).Clear
    imAvailGenerated = False
    slDate = txtDate.Text
    If gIsDate(slDate) And (imVefCode > 0) Then
        tmcFill.Enabled = True
    End If
End Sub

Private Sub txtDate_GotFocus()
    'tmcFill.Enabled = False
    gCtrlGotFocus ActiveControl
End Sub


Private Sub txtDate_LostFocus()
    Dim slDate As String
    
    slDate = txtDate.Text
    If gIsDate(slDate) Then
        If Trim$(txtModel.Text) = "" Then
            txtModel.Text = gAdjYear(Format$(DateValue(slDate) - 7, sgShowDateForm))
        End If
    End If
    'If gIsDate(slDate) And (imVefCode > 0) Then
    '    tmcFill.Enabled = True
    'End If
End Sub

Private Sub mGetAvails()
    Dim slSDate As String
    Dim slEDate As String
    Dim ilLoop As Integer
    Dim ilDay As Integer
    Dim slStr As String
    Dim slDays As String
    Dim VehCombo_rst As ADODB.Recordset
    
    lbcAvail(0).Clear
    lbcAvail(1).Clear
    lbcAvail(2).Clear
    ReDim tmAvInfo(0 To 0) As AVINFO
    ReDim tgDat(0 To 0) As DAT
    slSDate = txtDate.Text
    If gIsDate(slSDate) Then
        If DateValue(gAdjYear(slSDate)) < DateValue(gAdjYear(Format$(gNow(), "m/d/yy"))) Then
            gMsgBox "Date Cannot be Prior To: " & Format$(gNow(), sgShowDateForm), vbOKOnly
            Exit Sub
        End If
    End If
    If gIsDate(slSDate) And gIsDate(txtModel.Text) Then
        If DateValue(gAdjYear(slSDate)) < DateValue(gAdjYear(txtModel.Text)) Then
            gMsgBox "New Date Cannot be Prior to Existing Date", vbOKOnly
            Exit Sub
        End If
    End If
    If gIsDate(slSDate) And (imVefCode > 0) And (lbcAvail(1).ListCount <= 0) Then
        Screen.MousePointer = vbHourglass
        imAvailGenerated = True
        slSDate = gObtainPrevMonday(slSDate)
        slEDate = gAdjYear(Format$(DateValue(slSDate) + 6, sgShowDateForm))
        lacTitle2.Caption = "New Times: " & sgCRLF & slSDate & "-" & slEDate
        ReDim tmOldDat(0 To 0) As DAT
        'Note:  The Agreement and Station references are zero.  The station reference being zero
        '       will result in not time adjustment made to the avails
        SQLQuery = "Select vefCombineVefCode from VEF_Vehicles Where vefCode = " & imVefCode
        Set VehCombo_rst = gSQLSelectCall(SQLQuery)
        If Not VehCombo_rst.EOF Then
            imVefCombo = VehCombo_rst!vefCombineVefCode
        End If
        gGetAvails 0, 0, imVefCode, imVefCombo, slSDate, False
        Screen.MousePointer = vbHourglass
        ReDim tmNewDat(0 To UBound(tgDat)) As DAT
        If UBound(tgDat) > LBound(tgDat) Then
            For ilLoop = 0 To UBound(tgDat) - 1 Step 1
                tmNewDat(ilLoop) = tgDat(ilLoop)
                slStr = ""
                For ilDay = 0 To 6 Step 1
                    If tmNewDat(ilLoop).iFdDay(ilDay) = 1 Then
                        Select Case ilDay
                            Case 0
                                slStr = slStr + "Mo"
                            Case 1
                                slStr = slStr + "Tu"
                            Case 2
                                slStr = slStr + "We"
                            Case 3
                                slStr = slStr + "Th"
                            Case 4
                                slStr = slStr + "Fr"
                            Case 5
                                slStr = slStr + "Sa"
                            Case 6
                                slStr = slStr + "Su"
                        End Select
                    End If
                Next ilDay
                slDays = gDayMap(slStr)
                lbcAvail(1).AddItem tmNewDat(ilLoop).sFdSTime & " " & slDays
                lbcAvail(1).ItemData(lbcAvail(1).NewIndex) = ilLoop
            Next ilLoop
        End If
    End If
    slSDate = txtModel.Text
    If gIsDate(slSDate) And (imVefCode > 0) And (lbcAvail(0).ListCount <= 0) Then
        Screen.MousePointer = vbHourglass
        imAvailGenerated = True
        'slSDate = Format$(DateValue(slSDate) - 7, sgShowDateForm)
        slSDate = gObtainPrevMonday(slSDate)
        slEDate = gAdjYear(Format$(DateValue(slSDate) + 6, sgShowDateForm))
        lacTitle1.Caption = "Existing Times: " & sgCRLF & slSDate & "-" & slEDate
        ReDim tgDat(0 To 0) As DAT
        SQLQuery = "Select vefCombineVefCode from VEF_Vehicles Where vefCode = " & imVefCode
        Set VehCombo_rst = gSQLSelectCall(SQLQuery)
        If Not VehCombo_rst.EOF Then
            imVefCombo = VehCombo_rst!vefCombineVefCode
        End If
        gGetAvails 0, 0, imVefCode, imVefCombo, slSDate, False
        ReDim tmOldDat(0 To UBound(tgDat)) As DAT
        If UBound(tgDat) > LBound(tgDat) Then
            For ilLoop = 0 To UBound(tgDat) - 1 Step 1
                tmOldDat(ilLoop) = tgDat(ilLoop)
                slStr = ""
                For ilDay = 0 To 6 Step 1
                    If tmOldDat(ilLoop).iFdDay(ilDay) = 1 Then
                        Select Case ilDay
                            Case 0
                                slStr = slStr + "Mo"
                            Case 1
                                slStr = slStr + "Tu"
                            Case 2
                                slStr = slStr + "We"
                            Case 3
                                slStr = slStr + "Th"
                            Case 4
                                slStr = slStr + "Fr"
                            Case 5
                                slStr = slStr + "Sa"
                            Case 6
                                slStr = slStr + "Su"
                        End Select
                    End If
                Next ilDay
                slDays = gDayMap(slStr)
                'grdDayparts.AddItem iFdDay(0) & "|" & iFdDay(1) & "|" & iFdDay(2) & "|" & iFdDay(3) & "|" & iFdDay(4) & "|" & iFdDay(5) & "|" & iFdDay(6) & "|" & tmOldDat(iLoop).sFdSTime & "|" & tmOldDat(iLoop).sFdETime & "|" & sStatus & "|" & iPdDay(0) & "|" & iPdDay(1) & "|" & iPdDay(2) & "|" & iPdDay(3) & "|" & iPdDay(4) & "|" & iPdDay(5) & "|" & iPdDay(6) & "|" & tmOldDat(iLoop).sPdSTime & "|" & tmOldDat(iLoop).sPdETime & "|" & tmOldDat(iLoop).lCode
                lbcAvail(0).AddItem tmOldDat(ilLoop).sFdSTime & " " & slDays
                lbcAvail(0).ItemData(lbcAvail(0).NewIndex) = ilLoop
            Next ilLoop
        End If
    End If
    Screen.MousePointer = vbDefault
    imAvailGenerated = False
    
End Sub

Private Sub txtModel_Change()
    Dim slDate As String
    
    tmcFill.Enabled = False
    lbcAvail(0).Clear
    'lbcAvail(1).Clear
    lbcAvail(2).Clear
    imAvailGenerated = False
    slDate = txtModel.Text
    If gIsDate(slDate) And (imVefCode > 0) Then
        tmcFill.Enabled = True
    End If
End Sub

Private Sub txtModel_GotFocus()
    'tmcFill.Enabled = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtModel_LostFocus()
    Dim slDate As String
    
    'slDate = txtDate.Text
    'If gIsDate(slDate) And (imVefCode > 0) Then
    '    tmcFill.Enabled = True
    'End If

End Sub

Public Function mPopSelRemap(frm As Form, ilListCount As Integer, iIndex As Integer) As Integer

    Dim att_rst As ADODB.Recordset
    Dim shtt_rst As ADODB.Recordset
    Dim ilIdx As Integer
    
    On Error GoTo ErrHand
    
    'Screen.MousePointer = vbHourglass
    SQLQuery = "SELECT attCode, attShfCode"
    SQLQuery = SQLQuery + " FROM att"
    SQLQuery = SQLQuery + " WHERE (attVefCode = " & tgRemapInfo.iVefCode
    SQLQuery = SQLQuery + " AND attOffAir >= '" & Format$(tgRemapInfo.sStartDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery + " AND attDropDate >= '" & Format$(tgRemapInfo.sStartDate, sgSQLDateForm) & "'" & ")"
    Set att_rst = gSQLSelectCall(SQLQuery)

    ReDim tgAttInfo(0 To 0) As ATTINFO
    While Not att_rst.EOF
        tgAttInfo(UBound(tgAttInfo)).lAttCode = att_rst!attCode
        'We only need stn code is we allow users to pick affiliates
        'If frmAvRemap!lbcAvail(1).ListCount > 0 Then
        If ilListCount > 0 Then
            tgAttInfo(UBound(tgAttInfo)).iStnCode = att_rst!attshfCode
        End If
        'default value in the case that the New Times column had no times left in it.
        'If New Times is empty we don't allow users to select which affiliates to remap.
        'We do them all.  If New Times was not empty we allow the users to select which
        'affiliates to remap and set the iSelected element at that time.
        'If frmAvRemap!lbcAvail(1).ListCount = 0 Then
        If ilListCount = 0 Then
            tgAttInfo(UBound(tgAttInfo)).iSelected = True
        End If
        ReDim Preserve tgAttInfo(0 To (UBound(tgAttInfo) + 1))
        att_rst.MoveNext
    Wend
    
    'We don't need this info if we are not going to allow users to select affiliates to remap
    'If frmAvRemap!lbcAvail(1).ListCount > 0 Then
    If ilListCount > 0 Then
        For ilIdx = 0 To UBound(tgAttInfo) - 1 Step 1
            'SQLQuery = "SELECT shttCallLetters, shttMarket"
            'SQLQuery = SQLQuery + " FROM shtt"
            SQLQuery = "SELECT shttCallLetters, mktName"
            SQLQuery = SQLQuery + " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode"
            SQLQuery = SQLQuery + " WHERE (shttCode = " & tgAttInfo(ilIdx).iStnCode & ")"
            Set shtt_rst = gSQLSelectCall(SQLQuery)
            'tgAttInfo(ilIdx).sStnName = Trim$(shtt_rst!shttCallLetters) & ", " & Trim$(shtt_rst!shttMarket)
            If IsNull(shtt_rst!mktName) = True Then
                tgAttInfo(ilIdx).sStnName = Trim$(shtt_rst!shttCallLetters)
            Else
                tgAttInfo(ilIdx).sStnName = Trim$(shtt_rst!shttCallLetters) & ", " & Trim$(shtt_rst!mktName)
            End If
        Next ilIdx
    End If
    
    'Index 0 populate the to be remapped side; Index 1 populate the Not remapped side
    If iIndex = 0 Then
        'frmSelRemap!lbcSelRemap(1).Clear
        frm!lbcSelRemap(1).Clear
    Else
        'frmSelRemap!lbcSelRemap(0).Clear
        frm!lbcSelRemap(0).Clear
    End If
    
    'Screen.MousePointer = vbDefault
    For ilIdx = 0 To UBound(tgAttInfo) - 1 Step 1
        'frmSelRemap!lbcSelRemap(iIndex).AddItem tgAttInfo(ilIdx).sStnName
        'frmSelRemap!lbcSelRemap(iIndex).ItemData(frmSelRemap!lbcSelRemap(iIndex).NewIndex) = tgAttInfo(ilIdx).lAttCode
        frm!lbcSelRemap(iIndex).AddItem tgAttInfo(ilIdx).sStnName
        frm!lbcSelRemap(iIndex).ItemData(frm!lbcSelRemap(iIndex).NewIndex) = tgAttInfo(ilIdx).lAttCode
    Next ilIdx
    att_rst.Close
    If ilListCount > 0 Then
        shtt_rst.Close
    End If
    
    mPopSelRemap = True
    Exit Function
    
ErrHand:
    gHandleError "AffErrorLog.txt", "modAgmnt-gPopSelRemap"
    Screen.MousePointer = vbDefault
End Function
