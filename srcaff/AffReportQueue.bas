Attribute VB_Name = "modReportQueue"
'******************************************************
'*  modReportQueue - various global declarations for Report Queue
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private rst_rqt As ADODB.Recordset
Private rst_Rct As ADODB.Recordset
Private rst_site As ADODB.Recordset




Public Function gCreateReportQueue(frmForm As Form) As Integer

    Dim llRqtCode As Long
    Dim llRctCode As Long
    Dim Ctrl As control
    Dim blVisible As Boolean
    Dim ilRet As Integer
    Dim slCtrlName As String
    Dim slCtrlIndex As String
    Dim slDataType As String
    Dim llLongData As Long
    Dim slStringData As String
    Dim slDescription As String
    Dim ilTabIndex As Integer
    Dim ilLoop As Integer
    Dim ilDestination As Integer
    Dim ilFileType As Integer
        
    On Error GoTo CreateReportQueueErr
    llRqtCode = -1
    
    'Get Description
    sgRQReportName = sgReportListName
    frmReportQueueDescription.Show vbModal
    If igRQReturnStatus = 0 Then
        MsgBox "Report not sent to the Report Queue", vbOKOnly, "Queue"
        gCreateReportQueue = False
        Exit Function
    End If
    slDescription = sgRQDescription
    
    'Add RQT
    ilDestination = -1
    ilFileType = -1
    For Each Ctrl In frmForm.Controls
        On Error GoTo CreateReportQueueErr
        slCtrlName = Trim$(Ctrl.Name)
        On Error GoTo TabIndexErr
        ilTabIndex = Ctrl.TabIndex
        On Error GoTo CreateReportQueueErr
        If slCtrlName = "optRptDest" Then
            On Error GoTo IndexErr
            slCtrlIndex = Ctrl.Index
            On Error GoTo CreateReportQueueErr
            If Val(slCtrlIndex) = 1 Then
                If Ctrl.Value = True Then
                    ilDestination = 1
                End If
            ElseIf Val(slCtrlIndex) = 2 Then
                If Ctrl.Value = True Then
                    ilDestination = 2
                End If
            End If
        ElseIf slCtrlName = "cboFileType" Then
            If Ctrl.ListIndex >= 0 Then
                ilFileType = Ctrl.ListIndex
            End If
        End If
    Next Ctrl
    If (ilDestination = -1) Or ((ilDestination = 2) And (ilFileType = -1)) Then
        MsgBox "Report not sent to the Report Queue", vbOKOnly, "Queue"
        gCreateReportQueue = False
        Exit Function
    End If
    llRqtCode = mAddRqt(sgReportListName, slDescription, ilFileType)
    If llRqtCode <= 0 Then
        MsgBox "Report not sent to the Report Queue", vbOKOnly, "Queue"
        gCreateReportQueue = False
        Exit Function
    End If
    
    For Each Ctrl In frmForm.Controls
        On Error GoTo VisibleErr
        blVisible = Ctrl.Visible
        On Error GoTo IndexErr
        slCtrlIndex = Ctrl.Index
        On Error GoTo CreateReportQueueErr
        slCtrlName = Trim$(Ctrl.Name)
        On Error GoTo TabIndexErr
        ilTabIndex = Ctrl.TabIndex
        On Error GoTo CreateReportQueueErr
        If blVisible Then
            If TypeOf Ctrl Is ListBox Then
                If Ctrl.MultiSelect = 0 Then
                    If Ctrl.ListIndex >= 0 Then
                        slDataType = "S"
                        llLongData = Ctrl.ItemData(Ctrl.ListIndex)
                        slStringData = Ctrl.List(Ctrl.ListIndex)
                        llRctCode = mAddRct(llRqtCode, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                        If llRctCode <= 0 Then
                            MsgBox "Report not sent to the Report Queue", vbOKOnly, "Queue"
                            gCreateReportQueue = False
                            Exit Function
                        End If
                    End If
                Else
                    For ilLoop = 0 To Ctrl.ListCount - 1 Step 1
                        If Ctrl.Selected(ilLoop) Then
                            slDataType = "S"
                            llLongData = Ctrl.ItemData(ilLoop)
                            slStringData = Ctrl.List(ilLoop)
                            llRctCode = mAddRct(llRqtCode, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                            If llRctCode <= 0 Then
                                MsgBox "Report not sent to the Report Queue", vbOKOnly, "Queue"
                                gCreateReportQueue = False
                                Exit Function
                            End If
                        End If
                    Next ilLoop
                End If
            ElseIf TypeOf Ctrl Is TextBox Then
                slDataType = "S"
                llLongData = 0
                slStringData = Ctrl.Text
                llRctCode = mAddRct(llRqtCode, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                If llRctCode <= 0 Then
                    MsgBox "Report not sent to the Report Queue", vbOKOnly, "Queue"
                    gCreateReportQueue = False
                    Exit Function
                End If
            ElseIf TypeOf Ctrl Is ComboBox Then
                If Ctrl.ListIndex >= 0 Then
                    slDataType = "S"
                    llLongData = Ctrl.ItemData(Ctrl.ListIndex)
                    slStringData = Ctrl.List(Ctrl.ListIndex)
                    llRctCode = mAddRct(llRqtCode, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                    If llRctCode <= 0 Then
                        MsgBox "Report not sent to the Report Queue", vbOKOnly, "Queue"
                        gCreateReportQueue = False
                        Exit Function
                    End If
                End If
            ElseIf TypeOf Ctrl Is OptionButton Then
                slDataType = "L"
                llLongData = Ctrl.Value
                slStringData = ""
                llRctCode = mAddRct(llRqtCode, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                If llRctCode <= 0 Then
                    MsgBox "Report not sent to the Report Queue", vbOKOnly, "Queue"
                    gCreateReportQueue = False
                    Exit Function
                End If
            ElseIf TypeOf Ctrl Is CheckBox Then
                slDataType = "L"
                llLongData = Ctrl.Value
                slStringData = ""
                llRctCode = mAddRct(llRqtCode, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                If llRctCode <= 0 Then
                    MsgBox "Report not sent to the Report Queue", vbOKOnly, "Queue"
                    gCreateReportQueue = False
                    Exit Function
                End If
            ElseIf TypeOf Ctrl Is Frame Then
                slDataType = "S"
                llLongData = 0
                slStringData = Ctrl.Caption
                llRctCode = mAddRct(llRqtCode, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                If llRctCode <= 0 Then
                    MsgBox "Report not sent to the Report Queue", vbOKOnly, "Queue"
                    gCreateReportQueue = False
                    Exit Function
                End If
            ElseIf TypeOf Ctrl Is CSI_Calendar Then
                slDataType = "S"
                llLongData = 0
                slStringData = Ctrl.Text
                llRctCode = mAddRct(llRqtCode, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                If llRctCode <= 0 Then
                    MsgBox "Report not sent to the Report Queue", vbOKOnly, "Queue"
                    gCreateReportQueue = False
                    Exit Function
                End If
            ElseIf TypeOf Ctrl Is CSI_Calendar_UP Then
                slDataType = "S"
                llLongData = 0
                slStringData = Ctrl.Text
                llRctCode = mAddRct(llRqtCode, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                If llRctCode <= 0 Then
                    MsgBox "Report not sent to the Report Queue", vbOKOnly, "Queue"
                    gCreateReportQueue = False
                    Exit Function
                End If
            End If
        Else
            'Include the not visible controls inacse visible to checked within the report
            slDataType = "N"
            llLongData = 0
            slStringData = ""
            llRctCode = mAddRct(llRqtCode, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
            If llRctCode <= 0 Then
                MsgBox "Report not sent to the Report Queue", vbOKOnly, "Queue"
                gCreateReportQueue = False
                Exit Function
            End If
        End If
    Next Ctrl
    MsgBox "Report sent to the Report Queue successfully", vbOKOnly, "Queue"
    gCreateReportQueue = True
    Exit Function
IndexErr:
    slCtrlIndex = ""
    Resume Next
VisibleErr:
    blVisible = False
    Resume Next
TabIndexErr:
    ilTabIndex = -1
    Resume Next
CreateReportQueueErr:
    If llRqtCode <> -1 Then
        ilRet = mClearQueue(llRqtCode)
    End If
    MsgBox "Report not sent to the Report Queue", vbOKOnly, "Queue"
    gCreateReportQueue = False
    Resume Next
    Exit Function
End Function

Private Function mAddRct(llRqtCode As Long, ilTabIndex As Integer, slCtrlName As String, slCtrlIndex As String, slDataType As String, llLongData As Long, slStringData As String) As Long
    Dim llRctCode As Long
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    SQLQuery = "Insert Into rct ( "
    SQLQuery = SQLQuery & "rctCode, "
    SQLQuery = SQLQuery & "rctRqtCode, "
    SQLQuery = SQLQuery & "rctTabIndex, "
    SQLQuery = SQLQuery & "rctCtrlName, "
    SQLQuery = SQLQuery & "rctCtrlIndex, "
    SQLQuery = SQLQuery & "rctDataType, "
    SQLQuery = SQLQuery & "rctLongData, "
    SQLQuery = SQLQuery & "rctStringData, "
    SQLQuery = SQLQuery & "rctUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & llRqtCode & ", "
    SQLQuery = SQLQuery & ilTabIndex & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(slCtrlName) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(slCtrlIndex) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(slDataType) & "', "
    SQLQuery = SQLQuery & llLongData & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(slStringData) & "', "
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    llRctCode = gInsertAndReturnCode(SQLQuery, "rct", "rctCode", "Replace")
    mAddRct = llRctCode
    Exit Function
ErrHand:
    ilRet = mClearQueue(llRqtCode)
    gHandleError "AffErrorLog.txt", "Report Export Specifications-mAddRct"
    mAddRct = -1
End Function

Private Function mAddRqt(slReportName As String, slDescription As String, ilFileType As Integer) As Long
    Dim llRqtCode As Long
    Dim ilPriority As Integer
    
    On Error GoTo ErrHand
    'Get next Priority number
    SQLQuery = "Select "
    SQLQuery = SQLQuery & "Max(rqtPriority) "
    SQLQuery = SQLQuery & "From rqt "
    SQLQuery = SQLQuery & "Where rqtStatus <> 'C'"
    SQLQuery = SQLQuery & " And rqtStatus <> 'E'"
    SQLQuery = SQLQuery & " And rqtDateEntered >= '" & Format(gNow(), sgSQLDateForm) & "'"
    Set rst_rqt = gSQLSelectCall(SQLQuery)
    If IsNull(rst_rqt(0).Value) Then
        ilPriority = 1
    Else
        If Not rst_rqt.EOF Then
            ilPriority = rst_rqt(0).Value + 1
        Else
            ilPriority = 1
        End If
    End If
    
    
    SQLQuery = "Insert Into rqt ( "
    SQLQuery = SQLQuery & "rqtCode, "
    SQLQuery = SQLQuery & "rqtReportName, "
    SQLQuery = SQLQuery & "rqtDescription, "
    SQLQuery = SQLQuery & "rqtViewed, "
    SQLQuery = SQLQuery & "rqtFileType, "
    SQLQuery = SQLQuery & "rqtPriority, "
    SQLQuery = SQLQuery & "rqtDateEntered, "
    SQLQuery = SQLQuery & "rqtTimeEntered, "
    SQLQuery = SQLQuery & "rqtStatus, "
    SQLQuery = SQLQuery & "rqtDateStarted, "
    SQLQuery = SQLQuery & "rqtTimeStarted, "
    SQLQuery = SQLQuery & "rqtDateCompleted, "
    SQLQuery = SQLQuery & "rqtTimeCompleted, "
    SQLQuery = SQLQuery & "rqtUstCode, "
    SQLQuery = SQLQuery & "rqtUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(slReportName) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(slDescription) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
    SQLQuery = SQLQuery & ilFileType & ", "
    SQLQuery = SQLQuery & ilPriority & ", "
    SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & "N" & "', "
    SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("12AM", sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("12AM", sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & igUstCode & ", "
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    llRqtCode = gInsertAndReturnCode(SQLQuery, "rqt", "rqtCode", "Replace")
    mAddRqt = llRqtCode
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Report Export Specifications-mAddRqt"
    mAddRqt = -1
End Function

Private Function mClearQueue(llRqtCode As Long) As Integer
    On Error GoTo ErrHand
    
    SQLQuery = "DELETE FROM rct WHERE rctRqtCode = " & llRqtCode
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "AffErrorLog.txt", "Report Export Specifications-mClearQueue"
        mClearQueue = False
        Exit Function
    End If

    SQLQuery = "DELETE FROM rqt WHERE rqtCode = " & llRqtCode
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "AffErrorLog.txt", "Report Export Specifications-mClearQueue"
        mClearQueue = False
        Exit Function
    End If

    mClearQueue = True
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Report Export Specifications-mClearQueue"
    mClearQueue = False
    Exit Function
'ErrHand1:
'    gHandleError "AffErrorLog.txt", "Report Export Specifications-mClearQueue"
'    mClearQueue = False
End Function

Public Function gSetReportCtrls(frmForm As Form, llRqtCode As Long) As Integer
    Dim Ctrl As control
    Dim slCtrlIndex As String
    Dim slCtrlName As String
    Dim ilTabIndex As Integer
    Dim blListFound As Boolean
    Dim ilLoop As Integer
    
    
    On Error GoTo ErrHand
    
    SQLQuery = "SELECT * FROM rct WHERE rctRqtCode = " & llRqtCode & " ORDER BY rctTabIndex"
    Set rst_Rct = gSQLSelectCall(SQLQuery)
    Do While (Not rst_Rct.EOF) And (Not rst_Rct.BOF)
        For Each Ctrl In frmForm.Controls
            On Error GoTo IndexErr
            slCtrlIndex = Ctrl.Index
            On Error GoTo SetReportCtrlsErr
            slCtrlName = Trim$(Ctrl.Name)
            If slCtrlName = Trim$(rst_Rct!rctCtrlName) Then
                If Val(slCtrlIndex) = Val(rst_Rct!rctCtrlIndex) Then
                    If rst_Rct!rctDataType <> "N" Then
                        On Error Resume Next
                        Ctrl.Visible = True
                        On Error GoTo SetReportCtrlsErr
                        If TypeOf Ctrl Is ListBox Then
                            blListFound = False
                            For ilLoop = 0 To Ctrl.ListCount - 1 Step 1
                                If Trim$(rst_Rct!rctStringData) = Trim$(Ctrl.List(ilLoop)) Then
                                    blListFound = True
                                    If Ctrl.MultiSelect = 0 Then
                                        Ctrl.ListIndex = ilLoop
                                    Else
                                        Ctrl.Selected(ilLoop) = True
                                    End If
                                    Exit For
                                End If
                            Next ilLoop
                            If Not blListFound Then
                                Ctrl.AddItem Trim$(rst_Rct!StringData)
                                Ctrl.ItemData(Ctrl.NewIndex) = rst_Rct!rctLongData
                                If Ctrl.MultiSelect = 0 Then
                                    Ctrl.ListIndex = Ctrl.NewIndex
                                Else
                                    Ctrl.Selected(Ctrl.NewIndex) = True
                                End If
                            End If
                        ElseIf TypeOf Ctrl Is TextBox Then
                            Ctrl.Text = Trim$(rst_Rct!rctStringData)
                        ElseIf TypeOf Ctrl Is ComboBox Then
                            blListFound = False
                            For ilLoop = 0 To Ctrl.ListCount - 1 Step 1
                                If Trim$(rst_Rct!rctStringData) = Trim$(Ctrl.List(ilLoop)) Then
                                    blListFound = True
                                    Ctrl.ListIndex = ilLoop
                                    Exit For
                                End If
                            Next ilLoop
                            If Not blListFound Then
                                Ctrl.AddItem Trim$(rst_Rct!StringData)
                                Ctrl.ItemData(Ctrl.NewIndex) = rst_Rct!rctLongData
                                Ctrl.ListIndex = Ctrl.NewIndex
                            End If
                        ElseIf TypeOf Ctrl Is OptionButton Then
                            Ctrl.Value = rst_Rct!rctLongData
                        ElseIf TypeOf Ctrl Is CheckBox Then
                            Ctrl.Value = rst_Rct!rctLongData
                        ElseIf TypeOf Ctrl Is Frame Then
                            Ctrl.Caption = Trim$(rst_Rct!rctStringData)
                        ElseIf TypeOf Ctrl Is CSI_Calendar Then
                            Ctrl.Text = Trim$(rst_Rct!rctStringData)
                        ElseIf TypeOf Ctrl Is CSI_Calendar_UP Then
                            Ctrl.Text = Trim$(rst_Rct!rctStringData)
                        End If
                    Else
                        On Error Resume Next
                        Ctrl.Visible = False
                        On Error GoTo SetReportCtrlsErr
                    End If
                End If
            End If
        Next Ctrl
        rst_Rct.MoveNext
    Loop
    
    gSetReportCtrls = True
    Exit Function
IndexErr:
    slCtrlIndex = ""
    Resume Next
ErrHand:
    gHandleError "AffErrorLog.txt", "Report Export Specifications-gSetReportCtrls"
    'Resume Next
    gSetReportCtrls = False
    Exit Function
ErrHand1:
    gHandleError "AffErrorLog.txt", "Report Export Specifications-gSetReportCtrls"
    'Return
    gSetReportCtrls = False
    Exit Function
SetReportCtrlsErr:
    gSetReportCtrls = False
    Exit Function
End Function

