Attribute VB_Name = "FORMCTRLSUBS"

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: FormCtrlSubs.BAS
'
' Release: 1.0
'
Option Explicit
Option Compare Text

Type LISTBOXSORT
    sKey As String * 20 'L or C plus Top + Left
    sCtrlName As String * 40
    lCtrlIndex As Long
    lTabIndex As Long
End Type
Dim tmListBoxSort() As LISTBOXSORT

Private rst_Fcf As ADODB.Recordset


Public Function gCreateFormControlFile(frmForm As Form, slFormName As String, slBypassCtrls() As String, Optional blChangeTabIndex As Boolean = True) As Integer

    Dim llRqtCode As Long
    Dim llFcfCode As Long
    Dim Ctrl As Control
    Dim blVisible As Boolean
    Dim ilRet As Integer
    Dim slCtrlName As String
    Dim slCtrlIndex As String
    Dim ilCtrlIndex As Integer
    Dim slDataType As String
    Dim llLongData As Long
    Dim slStringData As String
    Dim slDescription As String
    Dim ilTabIndex As Integer
    Dim ilLoop As Integer
    Dim ilDestination As Integer
    Dim slSQLQuery As String
    Dim ilFileType As Integer
    Dim blBypassCtrl As Boolean
    Dim ilBypass As Integer
    Dim llRet As Long
    Dim ilMaxTabIndex As Integer
    'Dim ilListLeft As Integer
    'Dim ilListTop As Integer
    Dim slListLeft As String
    Dim slListTop As String
    Dim ilNoListCtrls As Integer
    Dim ilNoCheckBoxCtrls As Integer
        
    On Error GoTo gCreateFormControlFileErr
    'Remove previously defined control settings for the form
    slSQLQuery = "DELETE FROM fcf_Form_Control WHERE fcfName = '" & slFormName & "'"
    llRet = gSQLWaitNoMsgBox(slSQLQuery, False)

    ilMaxTabIndex = -1
    ilNoListCtrls = 0
    ilNoCheckBoxCtrls = 0
    ReDim tmListBoxSort(0 To 0) As LISTBOXSORT
    If blChangeTabIndex Then
        For Each Ctrl In frmForm.Controls
            On Error GoTo VisibleErr
            blVisible = Ctrl.Visible
            On Error GoTo IndexErr
            slCtrlIndex = Ctrl.Index
            On Error GoTo gCreateFormControlFileErr
            slCtrlName = Trim$(Ctrl.Name)
            On Error GoTo TabIndexErr
            ilTabIndex = Ctrl.TabIndex
            On Error GoTo gCreateFormControlFileErr
            blBypassCtrl = False
            If blVisible Then
                If ilTabIndex > ilMaxTabIndex Then
                    ilMaxTabIndex = ilTabIndex
                End If
                If (TypeOf Ctrl Is ListBox) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CSI_ComboBoxList) Then
                    ilNoListCtrls = ilNoListCtrls + 1
                    slListLeft = Ctrl.Left
                    Do While Len(slListLeft) < 6
                        slListLeft = "0" & slListLeft
                    Loop
                    slListTop = Ctrl.Top
                    Do While Len(slListTop) < 6
                        slListTop = "0" & slListTop
                    Loop
                    If (TypeOf Ctrl Is ListBox) Then
                        tmListBoxSort(UBound(tmListBoxSort)).sKey = "L" & slListTop & slListLeft
                    Else
                        tmListBoxSort(UBound(tmListBoxSort)).sKey = "C" & slListTop & slListLeft
                    End If
                    tmListBoxSort(UBound(tmListBoxSort)).sCtrlName = slCtrlName
                    If slCtrlIndex <> "" Then
                        tmListBoxSort(UBound(tmListBoxSort)).lCtrlIndex = Val(slCtrlIndex)
                    Else
                        tmListBoxSort(UBound(tmListBoxSort)).lCtrlIndex = -1
                    End If
                    ReDim Preserve tmListBoxSort(0 To UBound(tmListBoxSort) + 1) As LISTBOXSORT
                End If
            End If
        Next Ctrl
        If UBound(tmListBoxSort) - 1 > 0 Then
            ArraySortTyp fnAV(tmListBoxSort(), 0), UBound(tmListBoxSort), 0, LenB(tmListBoxSort(0)), 0, LenB(tmListBoxSort(0).sKey), 0
        End If
        ilMaxTabIndex = ilMaxTabIndex + 1
        For ilLoop = 0 To UBound(tmListBoxSort) - 1 Step 1
            tmListBoxSort(ilLoop).lTabIndex = ilMaxTabIndex + ilLoop
        Next ilLoop
    End If
    
    For Each Ctrl In frmForm.Controls
        On Error GoTo VisibleErr
        blVisible = Ctrl.Visible
        On Error GoTo IndexErr
        slCtrlIndex = Ctrl.Index
        On Error GoTo gCreateFormControlFileErr
        slCtrlName = Trim$(Ctrl.Name)
        On Error GoTo TabIndexErr
        ilTabIndex = Ctrl.TabIndex
        On Error GoTo gCreateFormControlFileErr
        blBypassCtrl = False
        If UBound(slBypassCtrls) > LBound(slBypassCtrls) Then
            For ilBypass = LBound(slBypassCtrls) To UBound(slBypassCtrls) Step 1
                If slCtrlIndex = "" Then
                    If slCtrlName = slBypassCtrls(ilBypass) Then
                        blBypassCtrl = True
                        Exit For
                    End If
                Else
                    If InStr(1, slBypassCtrls(ilBypass), "(") > 0 Then
                        If slCtrlName & "(" & Trim$(slCtrlIndex) & ")" = slBypassCtrls(ilBypass) Then
                            blBypassCtrl = True
                            Exit For
                        End If
                    Else
                        If slCtrlName = slBypassCtrls(ilBypass) Then
                            blBypassCtrl = True
                            Exit For
                        End If
                    End If
                End If
            Next ilBypass
        End If
        If Not blBypassCtrl Then
            If blVisible Then
                If TypeOf Ctrl Is ListBox Then
                    If blChangeTabIndex Then
                        'ilTabIndex = ilMaxTabIndex
                        'ilMaxTabIndex = ilMaxTabIndex + 1
                        For ilLoop = 0 To UBound(tmListBoxSort) - 1 Step 1
                            If Left(tmListBoxSort(ilLoop).sKey, 1) = "L" Then
                                If slCtrlIndex <> "" Then
                                    ilCtrlIndex = Val(slCtrlIndex)
                                Else
                                    ilCtrlIndex = -1
                                End If
                                If (Trim$(tmListBoxSort(ilLoop).sCtrlName) = slCtrlName) And (tmListBoxSort(ilLoop).lCtrlIndex = ilCtrlIndex) Then
                                    ilTabIndex = tmListBoxSort(ilLoop).lTabIndex
                                    Exit For
                                End If
                            End If
                        Next ilLoop
                    End If
                    If Ctrl.MultiSelect = 0 Then
                        If Ctrl.ListIndex >= 0 Then
                            slDataType = "S"
                            llLongData = Ctrl.ItemData(Ctrl.ListIndex)
                            slStringData = Ctrl.List(Ctrl.ListIndex)
                            llFcfCode = mAddFcf(slFormName, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                            If llFcfCode <= 0 Then
                                gCreateFormControlFile = False
                                Exit Function
                            End If
                        End If
                    Else
                        If Ctrl.SelCount <> Ctrl.ListCount Then
                            For ilLoop = 0 To Ctrl.ListCount - 1 Step 1
                                If Ctrl.Selected(ilLoop) Then
                                    slDataType = "S"
                                    llLongData = Ctrl.ItemData(ilLoop)
                                    slStringData = Ctrl.List(ilLoop)
                                    llFcfCode = mAddFcf(slFormName, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                                    If llFcfCode <= 0 Then
                                        gCreateFormControlFile = False
                                        Exit Function
                                    End If
                                End If
                            Next ilLoop
                        Else
                            slDataType = "S"
                            llLongData = -1
                            slStringData = "ALL~SELECTED"
                            llFcfCode = mAddFcf(slFormName, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                            If llFcfCode <= 0 Then
                                gCreateFormControlFile = False
                                Exit Function
                            End If
                        End If
                    End If
                ElseIf TypeOf Ctrl Is TextBox Then
                    slDataType = "S"
                    llLongData = 0
                    slStringData = Ctrl.Text
                    llFcfCode = mAddFcf(slFormName, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                    If llFcfCode <= 0 Then
                        gCreateFormControlFile = False
                        Exit Function
                    End If
                ElseIf TypeOf Ctrl Is ComboBox Or TypeOf Ctrl Is CSI_ComboBoxList Then
                    If blChangeTabIndex Then
                        'ilTabIndex = ilMaxTabIndex
                        'ilMaxTabIndex = ilMaxTabIndex + 1
                        For ilLoop = 0 To UBound(tmListBoxSort) - 1 Step 1
                            If Left(tmListBoxSort(ilLoop).sKey, 1) = "C" Then
                                If slCtrlIndex <> "" Then
                                    ilCtrlIndex = Val(slCtrlIndex)
                                Else
                                    ilCtrlIndex = -1
                                End If
                                If Trim$(tmListBoxSort(ilLoop).sCtrlName) = slCtrlName And tmListBoxSort(ilLoop).lCtrlIndex = ilCtrlIndex Then
                                    ilTabIndex = tmListBoxSort(ilLoop).lTabIndex
                                    Exit For
                                End If
                            End If
                        Next ilLoop
                    End If
                    If TypeOf Ctrl Is CSI_ComboBoxList Then
                        If Ctrl.ListIndex >= 0 Then
                            slDataType = "S"
                            llLongData = Ctrl.GetItemData(Ctrl.ListIndex)
                            slStringData = Ctrl.GetName(Ctrl.ListIndex)
                            llFcfCode = mAddFcf(slFormName, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                            If llFcfCode <= 0 Then
                                gCreateFormControlFile = False
                                Exit Function
                            End If
                        ElseIf Trim$(Ctrl.Text) <> "" Then
                            For ilLoop = 0 To Ctrl.ListCount - 1 Step 1
                                If Trim$(Ctrl.Text) = Trim$(Ctrl.GetName(ilLoop)) Then
                                    slDataType = "S"
                                    llLongData = Ctrl.GettemData(ilLoop)
                                    slStringData = Ctrl.GetName(ilLoop)
                                    llFcfCode = mAddFcf(slFormName, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                                    If llFcfCode <= 0 Then
                                        gCreateFormControlFile = False
                                        Exit Function
                                    End If
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                    Else
                        If Ctrl.ListIndex >= 0 Then
                            slDataType = "S"
                            llLongData = Ctrl.ItemData(Ctrl.ListIndex)
                            slStringData = Ctrl.List(Ctrl.ListIndex)
                            llFcfCode = mAddFcf(slFormName, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                            If llFcfCode <= 0 Then
                                gCreateFormControlFile = False
                                Exit Function
                            End If
                        ElseIf Trim$(Ctrl.Text) <> "" Then
                            For ilLoop = 0 To Ctrl.ListCount - 1 Step 1
                                If Trim$(Ctrl.Text) = Trim$(Ctrl.List(ilLoop)) Then
                                    slDataType = "S"
                                    llLongData = Ctrl.ItemData(ilLoop)
                                    slStringData = Ctrl.List(ilLoop)
                                    llFcfCode = mAddFcf(slFormName, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                                    If llFcfCode <= 0 Then
                                        gCreateFormControlFile = False
                                        Exit Function
                                    End If
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                    End If
                ElseIf TypeOf Ctrl Is OptionButton Then
                    slDataType = "L"
                    llLongData = Ctrl.Value
                    slStringData = ""
                    llFcfCode = mAddFcf(slFormName, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                    If llFcfCode <= 0 Then
                        gCreateFormControlFile = False
                        Exit Function
                    End If
                ElseIf TypeOf Ctrl Is CheckBox Then
                    'If (blChangeTabIndex) And (Ctrl.Left >= ilListLeft - 5) Then
                    If (blChangeTabIndex) And (InStr(1, UCase(slCtrlName), "ALL") > 0) Then
                        ilTabIndex = ilMaxTabIndex + ilNoListCtrls + ilNoCheckBoxCtrls
                        ilNoCheckBoxCtrls = ilNoCheckBoxCtrls + 1
                    End If
                    slDataType = "L"
                    If Ctrl.Value = vbChecked Then
                        llLongData = 1
                    Else
                        llLongData = 0
                    End If
                    slStringData = ""
                    llFcfCode = mAddFcf(slFormName, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                    If llFcfCode <= 0 Then
                        gCreateFormControlFile = False
                        Exit Function
                    End If
                ElseIf TypeOf Ctrl Is Frame Then
                    slDataType = "S"
                    llLongData = 0
                    slStringData = Ctrl.Caption
                    llFcfCode = mAddFcf(slFormName, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                    If llFcfCode <= 0 Then
                        gCreateFormControlFile = False
                        Exit Function
                    End If
                ElseIf TypeOf Ctrl Is CSI_Calendar Then
                    slDataType = "S"
                    llLongData = 0
                    slStringData = Ctrl.Text
                    llFcfCode = mAddFcf(slFormName, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                    If llFcfCode <= 0 Then
                        gCreateFormControlFile = False
                        Exit Function
                    End If
                'Not part of traffic, if required obtain the user control from Affiliate
                'ElseIf TypeOf Ctrl Is CSI_Calendar_UP Then
                '    slDataType = "S"
                '    llLongData = 0
                '    slStringData = Ctrl.Text
                '    llFcfCode = mAddFcf(slFormName, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                '    If llFcfCode <= 0 Then
                '        MsgBox "Report not sent to the Report Queue", vbOKOnly, "Queue"
                '        gCreateFormControlFile = False
                '        Exit Function
                '    End If
                End If
            Else
                'Include the not visible controls inacse visible to checked within the report
                slDataType = "N"
                llLongData = 0
                slStringData = ""
                llFcfCode = mAddFcf(slFormName, ilTabIndex, slCtrlName, slCtrlIndex, slDataType, llLongData, slStringData)
                If llFcfCode <= 0 Then
                    gCreateFormControlFile = False
                    Exit Function
                End If
            End If
        End If
    Next Ctrl
    gCreateFormControlFile = True
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
gCreateFormControlFileErr:
    gCreateFormControlFile = False
    'Resume Next
    Exit Function
End Function
Public Function gSetFormCtrls(frmForm As Form, slFormName As String) As Integer
    Dim Ctrl As Control
    Dim slCtrlIndex As String
    Dim slCtrlName As String
    Dim slPrevListCtrlName As String
    Dim ilTabIndex As Integer
    Dim blListFound As Boolean
    Dim ilLoop As Integer
    Dim slSQLQuery As String
    
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    
    On Error GoTo ErrHand
    slPrevListCtrlName = ""
    
    slSQLQuery = "SELECT * FROM fcf_Form_Control WHERE fcfName = '" & slFormName & "' ORDER BY fcfTabIndex"
    Set rst_Fcf = gSQLSelectCall(slSQLQuery)
    Do While (Not rst_Fcf.EOF) And (Not rst_Fcf.BOF)
        For Each Ctrl In frmForm.Controls
            On Error GoTo IndexErr
            slCtrlIndex = Ctrl.Index
            On Error GoTo SetReportCtrlsErr
            slCtrlName = Trim$(Ctrl.Name)
            If slCtrlName = Trim$(rst_Fcf!fcfCtrlName) Then
                If Val(slCtrlIndex) = Val(rst_Fcf!fcfCtrlIndex) Then
                    If rst_Fcf!fcfDataType <> "N" Then
                        On Error Resume Next
                        Ctrl.Visible = True
                        On Error GoTo SetReportCtrlsErr
                        If TypeOf Ctrl Is ListBox Then
                            If slPrevListCtrlName <> slCtrlName Then
                                If Ctrl.MultiSelect <> 0 Then
                                    'For ilLoop = 0 To Ctrl.ListCount - 1 Step 1
                                    '    Ctrl.Selected(ilLoop) = False
                                    'Next ilLoop
                                    llRg = CLng(Ctrl.ListCount - 1) * &H10000 Or 0
                                    ilValue = False
                                    llRet = SendMessageByNum(Ctrl.hWnd, LB_SELITEMRANGE, ilValue, llRg)
                                End If
                            End If
                            slPrevListCtrlName = slCtrlName
                            If Trim$(rst_Fcf!FcfStringData) <> Trim$("ALL~SELECTED") Then
                                blListFound = False
                                For ilLoop = 0 To Ctrl.ListCount - 1 Step 1
                                    If Trim$(rst_Fcf!FcfStringData) = Trim$(Ctrl.List(ilLoop)) Then
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
                                    Ctrl.AddItem Trim$(rst_Fcf!FcfStringData)
                                    Ctrl.ItemData(Ctrl.NewIndex) = rst_Fcf!fcfLongData
                                    If Ctrl.MultiSelect = 0 Then
                                        Ctrl.ListIndex = Ctrl.NewIndex
                                    Else
                                        Ctrl.Selected(Ctrl.NewIndex) = True
                                    End If
                                End If
                            Else
                                llRg = CLng(Ctrl.ListCount - 1) * &H10000 Or 0
                                ilValue = True
                                llRet = SendMessageByNum(Ctrl.hWnd, LB_SELITEMRANGE, ilValue, llRg)
                            End If
                        ElseIf TypeOf Ctrl Is TextBox Then
                            If igReportButtonIndex = 1 Then
                                If Not gIsDate(Trim$(rst_Fcf!FcfStringData)) Then
                                    Ctrl.Text = Trim$(rst_Fcf!FcfStringData)
                                End If
                            Else
                                Ctrl.Text = Trim$(rst_Fcf!FcfStringData)
                            End If
                        ElseIf TypeOf Ctrl Is ComboBox Then
                            blListFound = False
                            For ilLoop = 0 To Ctrl.ListCount - 1 Step 1
                                If Trim$(rst_Fcf!FcfStringData) = Trim$(Ctrl.List(ilLoop)) Then
                                    blListFound = True
                                    Ctrl.ListIndex = ilLoop
                                    Exit For
                                End If
                            Next ilLoop
                            If Not blListFound Then
                                Ctrl.AddItem Trim$(rst_Fcf!FcfStringData)
                                Ctrl.ItemData(Ctrl.NewIndex) = rst_Fcf!fcfLongData
                                Ctrl.ListIndex = Ctrl.NewIndex
                            End If
                        ElseIf TypeOf Ctrl Is CSI_ComboBoxList Then
                            blListFound = False
                            For ilLoop = 0 To Ctrl.ListCount - 1 Step 1
                                If Trim$(rst_Fcf!FcfStringData) = Trim$(Ctrl.GetName(ilLoop)) Then
                                    blListFound = True
                                    Ctrl.SetListIndex = ilLoop
                                    Exit For
                                End If
                            Next ilLoop
                            If Not blListFound Then
                                Ctrl.AddItem Trim$(rst_Fcf!FcfStringData)
                                Ctrl.SetItemData = rst_Fcf!fcfLongData
                            End If
                        ElseIf TypeOf Ctrl Is OptionButton Then
                            Ctrl.Value = rst_Fcf!fcfLongData
                        ElseIf TypeOf Ctrl Is CheckBox Then
                            If rst_Fcf!fcfLongData = 1 Then
                                Ctrl.Value = vbChecked
                            Else
                                Ctrl.Value = vbUnchecked
                            End If
                        ElseIf TypeOf Ctrl Is Frame Then
                            Ctrl.Caption = Trim$(rst_Fcf!FcfStringData)
                        ElseIf TypeOf Ctrl Is CSI_Calendar Then
                            If igReportButtonIndex = 0 Then
                                Ctrl.Text = Trim$(rst_Fcf!FcfStringData)
                            End If
                        'Not part of traffic, if required obtain the user control from Affiliate
                        'ElseIf TypeOf Ctrl Is CSI_Calendar_UP Then
                        '    Ctrl.Text = Trim$(rst_Fcf!fcfStringData)
                        End If
                    Else
                        On Error Resume Next
                        Select Case UCase$(slCtrlName)
                            Case "CMCSETUP", "EDCCOPIES", "CBCFILETYPE", "EDCFILENAME", "CMCBROWSE"
                                Ctrl.Visible = True
                            Case Else
                                Ctrl.Visible = False
                        End Select
                        'Ctrl.Visible = False
                        On Error GoTo SetReportCtrlsErr
                    End If
                    Exit For
                End If
            End If
        Next Ctrl
        rst_Fcf.MoveNext
    Loop
    
    gSetFormCtrls = True
    Exit Function
IndexErr:
    slCtrlIndex = ""
    Resume Next
ErrHand:
    gHandleError "TrafficErrors.txt", "Form Control-gSetFormCtrls"
    'Resume Next
    gSetFormCtrls = False
    Exit Function
SetReportCtrlsErr:
    gSetFormCtrls = False
    Exit Function
End Function

Private Function mAddFcf(slFormName As String, ilTabIndex As Integer, slCtrlName As String, slCtrlIndex As String, slDataType As String, llLongData As Long, slStringData As String) As Long
    Dim llFcfCode As Long
    Dim ilRet As Integer
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand
    slSQLQuery = "Insert Into fcf_Form_Control ( "
    slSQLQuery = slSQLQuery & "fcfCode, "
    slSQLQuery = slSQLQuery & "fcfName, "
    slSQLQuery = slSQLQuery & "fcfTabIndex, "
    slSQLQuery = slSQLQuery & "fcfCtrlName, "
    slSQLQuery = slSQLQuery & "fcfCtrlIndex, "
    slSQLQuery = slSQLQuery & "fcfDataType, "
    slSQLQuery = slSQLQuery & "fcfLongData, "
    slSQLQuery = slSQLQuery & "fcfStringData, "
    slSQLQuery = slSQLQuery & "fcfUrfCode, "
    slSQLQuery = slSQLQuery & "fcfUnused "
    slSQLQuery = slSQLQuery & ") "
    slSQLQuery = slSQLQuery & "Values ( "
    slSQLQuery = slSQLQuery & "Replace-X" & ", "
    slSQLQuery = slSQLQuery & "'" & slFormName & "'" & ", "
    slSQLQuery = slSQLQuery & ilTabIndex & ", "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(slCtrlName) & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(slCtrlIndex) & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(slDataType) & "', "
    slSQLQuery = slSQLQuery & llLongData & ", "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(slStringData) & "', "
    slSQLQuery = slSQLQuery & tgUrf(0).iCode & ", "
    slSQLQuery = slSQLQuery & "'" & "" & "' "
    slSQLQuery = slSQLQuery & ") "
    llFcfCode = gInsertAndReturnCode(slSQLQuery, "fcf_Form_Control", "fcfCode", "Replace-X")
    mAddFcf = llFcfCode
    Exit Function
ErrHand:
    gHandleError "TrafficErrors.txt", "Report Export Specifications-mAddFcf"
    mAddFcf = -1
End Function

