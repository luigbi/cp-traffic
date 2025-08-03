Attribute VB_Name = "ContractNative2"

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ContractNative2.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the extra Contract routines so that code can be compiled as Native code
Option Explicit
Option Compare Text

Public imCPMCtrlVisible As Integer
Public lmCPMEnableRow As Long
Public lmCPMEnableCol As Long
Public lmCPMTopRow As Long
Dim imInitNoRows As Integer
Public imCPMIgnoreScroll As Integer
Public igCPMShift As Integer
Public igCPMKeyCode As Integer
Public imCPMFromArrow As Integer
Public smHiddenSpace As String

Public bgContractAttachments_Changed As Boolean

Public tmCopyTypeCode() As SORTCODE
Public smCopyTypeCodeTag As String
Public imCopyTypeChgMode As Integer

Public tmPodTargetCode() As SORTCODE
Public smPodTargetCodeTag As String
Public imPodTargetChgMode As Integer
Dim tmPodTarget() As POPLCODENAME
Public smCPMLLD As String
'10719 last booked date
Public smCPMLBD As String

Public lgDeletedPcfCode() As Long

Public tgCPMPkVeh() As SORTCODE

Dim imLastColSorted As Integer
Dim imLastSort As Integer
Public imCPMSortMethod As Integer   '=Std; 1= #
Public imPriorityAvfCode As Integer
Public avf_rst As ADODB.Recordset
Public imPositionAvfCode As Integer
Public mnf_rst As ADODB.Recordset

Public bmFlatRateExist As Boolean
'10910
Public tgContractDocuments() As ACFList
'10910-show moved here so can be public
'10967
Public tmDigitalVendorsToTest() As VendorInfo
Public Const ATTACHMENTINDEX = 34
Public Const VENDORINDEX = 26
Public Const CPMEXPAND = 0
Public Const CPMID = 1
Public Const CPMVEHICLE = 2
Public Const CPMDAYPART = 3
'Public Const CPMCOPYTYPE = 4
'10879 increment each to add Length
'Public Const CPMPRICETYPE = 4
'Public Const CPMPOSITION = 5
'Public Const CPMPRIORITY = 6
'Public Const CPMPODTARGET = 7
'Public Const CPMSTARTDATE = 8
'Public Const CPMENDDATE = 9
'Public Const CPMIMPRESSIONS = 10
'Public Const CPMIMPRESSIONCAP = 11
'Public Const CPMCPMVALUE = 12
'Public Const CPMCOST = 13
'Public Const CPMCOMMENT = 14
'Public Const CPMPCFCODE = 15
'Public Const CPMLNTYPE = 16
'Public Const CPMVEFCODE = 17
'Public Const CPMRDFCODE = 18
'Public Const CPMCHFCODE = 19
'Public Const CPMPACKAGEID = 20
'Public Const CPMSHOWROW = 21
'Public Const CPMORIGSTARTDATE = 22
'Public Const CPMORIGENDDATE = 23
'Public Const CPMCXFCODE = 24
'Public Const CPMCOMMSHOWON = 25
'Public Const CPMCOLOR = 26
'Public Const CPMBILLED = 27
'Public Const CPMSORT = 28
'10738
'Public Const CPMORIGIMPRESSIONS = 29

Public Const CPMLENGTH = 4
Public Const CPMPRICETYPE = 5
Public Const CPMPOSITION = 6
Public Const CPMPRIORITY = 7
Public Const CPMPODTARGET = 8
Public Const CPMSTARTDATE = 9
Public Const CPMENDDATE = 10
Public Const CPMIMPRESSIONS = 11
Public Const CPMIMPRESSIONCAP = 12
Public Const CPMCPMVALUE = 13
Public Const CPMCOST = 14
Public Const CPMCOMMENT = 15
Public Const CPMPCFCODE = 16
Public Const CPMLNTYPE = 17
Public Const CPMVEFCODE = 18
Public Const CPMRDFCODE = 19
Public Const CPMCHFCODE = 20
Public Const CPMPACKAGEID = 21
Public Const CPMSHOWROW = 22
Public Const CPMORIGSTARTDATE = 23
Public Const CPMORIGENDDATE = 24
Public Const CPMCXFCODE = 25
Public Const CPMCOMMSHOWON = 26
Public Const CPMCOLOR = 27
Public Const CPMBILLED = 28
Public Const CPMSORT = 29
Public Const CPMORIGIMPRESSIONS = 30
'10967
Public Const CPMORIGCAP = 31
Public Const CPMORIGPOSITION = 32
Public Const CPMORIGPRIORITY = 33
Public Const CPMORIGVEHICLE = 34
Public Const CPMORIGTOTAL = 35
Public Const CPMAVFCODE = 36
Public Const CPMDELIVERYSTATUS = 37
'10929 could remove if have to
Private Const HIDDENLENMATCH = 0
Private Const HIDDENLENCONSISTENTBUTNOMATCH = 1
Private Const HIDDENLENNOTCONSISTENT = 2
Private Const HIDDENLENNOTALLOWED = 3
Public Sub mGridCPMBuyColumns()
    Dim ilCol As Integer
    Dim llRow As Long
        
    Contract.grdCPMBuy.ColWidth(CPMSORT) = 0
    Contract.grdCPMBuy.ColWidth(CPMBILLED) = 0
    Contract.grdCPMBuy.ColWidth(CPMSHOWROW) = 0
    Contract.grdCPMBuy.ColWidth(CPMORIGENDDATE) = 0
    Contract.grdCPMBuy.ColWidth(CPMORIGSTARTDATE) = 0
    Contract.grdCPMBuy.ColWidth(CPMCXFCODE) = 0
    Contract.grdCPMBuy.ColWidth(CPMCOMMSHOWON) = 0
    Contract.grdCPMBuy.ColWidth(CPMCOLOR) = 0
    Contract.grdCPMBuy.ColWidth(CPMPACKAGEID) = 0
    Contract.grdCPMBuy.ColWidth(CPMCHFCODE) = 0
    Contract.grdCPMBuy.ColWidth(CPMRDFCODE) = 0
    Contract.grdCPMBuy.ColWidth(CPMVEFCODE) = 0
    Contract.grdCPMBuy.ColWidth(CPMLNTYPE) = 0
    Contract.grdCPMBuy.ColWidth(CPMPCFCODE) = 0
    '10738
    Contract.grdCPMBuy.ColWidth(CPMORIGIMPRESSIONS) = 0
    '10967
    Contract.grdCPMBuy.ColWidth(CPMORIGCAP) = 0
    Contract.grdCPMBuy.ColWidth(CPMORIGPOSITION) = 0
    Contract.grdCPMBuy.ColWidth(CPMORIGPRIORITY) = 0
    Contract.grdCPMBuy.ColWidth(CPMORIGVEHICLE) = 0
    Contract.grdCPMBuy.ColWidth(CPMORIGTOTAL) = 0
    Contract.grdCPMBuy.ColWidth(CPMDELIVERYSTATUS) = 0
    Contract.grdCPMBuy.ColWidth(CPMAVFCODE) = 0
    If Not bgHideHiddenLines Then
        Contract.grdCPMBuy.ColWidth(CPMEXPAND) = 0
    Else
        Contract.grdCPMBuy.ColWidth(CPMEXPAND) = Contract.grdCPMBuy.Width * 0.015
    End If
    Contract.grdCPMBuy.ColWidth(CPMID) = Contract.grdCPMBuy.Width * 0.03
    Contract.grdCPMBuy.ColWidth(CPMVEHICLE) = Contract.grdCPMBuy.Width * 0.15
    Contract.grdCPMBuy.ColWidth(CPMDAYPART) = Contract.grdCPMBuy.Width * 0.15
    'Contract.grdCPMBuy.ColWidth(CPMCOPYTYPE) = Contract.grdCPMBuy.Width * 0.05
    '10879
    Contract.grdCPMBuy.ColWidth(CPMLENGTH) = Contract.grdCPMBuy.Width * 0.05
    Contract.grdCPMBuy.ColWidth(CPMPRICETYPE) = Contract.grdCPMBuy.Width * 0.05
    Contract.grdCPMBuy.ColWidth(CPMPOSITION) = Contract.grdCPMBuy.Width * 0.06
    Contract.grdCPMBuy.ColWidth(CPMPRIORITY) = Contract.grdCPMBuy.Width * 0.04
    Contract.grdCPMBuy.ColWidth(CPMSTARTDATE) = Contract.grdCPMBuy.Width * 0.05
    Contract.grdCPMBuy.ColWidth(CPMENDDATE) = Contract.grdCPMBuy.Width * 0.05
    Contract.grdCPMBuy.ColWidth(CPMIMPRESSIONS) = Contract.grdCPMBuy.Width * 0.06
    Contract.grdCPMBuy.ColWidth(CPMIMPRESSIONCAP) = Contract.grdCPMBuy.Width * 0.06
    Contract.grdCPMBuy.ColWidth(CPMCPMVALUE) = Contract.grdCPMBuy.Width * 0.06
    Contract.grdCPMBuy.ColWidth(CPMCOST) = Contract.grdCPMBuy.Width * 0.06
    Contract.grdCPMBuy.ColWidth(CPMCOMMENT) = Contract.grdCPMBuy.Width * 0.015
    
    Contract.grdCPMBuy.ColWidth(CPMPODTARGET) = Contract.grdCPMBuy.Width - GRIDSCROLLWIDTH - 15
    For ilCol = CPMEXPAND To CPMSORT Step 1
        If ilCol <> CPMPODTARGET Then
            Contract.grdCPMBuy.ColWidth(CPMPODTARGET) = Contract.grdCPMBuy.ColWidth(CPMPODTARGET) - Contract.grdCPMBuy.ColWidth(ilCol)
        End If
    Next ilCol
    '10967 fails starting with my new CPMORIGCAP(31)
    'Align columns to left
    'gGrid_AlignAllColsLeft Contract.grdCPMBuy
    For ilCol = 0 To CPMORIGIMPRESSIONS Step 1
        Contract.grdCPMBuy.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
    'Contract.grdCPMBuy.ColAlignment(CPMPOSITION) = flexAlignRightCenter
    'Contract.grdCPMBuy.ColAlignment(CPMPRIORITY) = flexAlignRightCenter
    Contract.grdCPMBuy.ColAlignment(CPMIMPRESSIONS) = flexAlignRightCenter
    Contract.grdCPMBuy.ColAlignment(CPMIMPRESSIONCAP) = flexAlignRightCenter
    Contract.grdCPMBuy.ColAlignment(CPMCPMVALUE) = flexAlignRightCenter
    Contract.grdCPMBuy.ColAlignment(CPMCOST) = flexAlignRightCenter
End Sub
Public Sub mGridCPMBuyTitles()
    'Set column titles
    Dim ilCol As Integer
    
    Contract.grdCPMBuy.Row = 0
    For ilCol = CPMEXPAND To CPMCOMMENT Step 1
        Contract.grdCPMBuy.Col = ilCol
        Contract.grdCPMBuy.CellFontBold = False
        Contract.grdCPMBuy.CellFontName = "Arial"
        Contract.grdCPMBuy.CellFontSize = 6.75
        Contract.grdCPMBuy.CellForeColor = vbBlue
        If ilCol = CPMID Then
            Contract.grdCPMBuy.CellBackColor = LIGHTBLUE  'vbWhite
        End If
    Next ilCol
    Contract.grdCPMBuy.TextMatrix(0, CPMID) = "ID"
    Contract.grdCPMBuy.TextMatrix(0, CPMVEHICLE) = "Vehicle"
    Contract.grdCPMBuy.TextMatrix(0, CPMDAYPART) = "Ad Location"
    'Contract.grdCPMBuy.TextMatrix(0, CPMCOPYTYPE) = "Copy Type"
    Contract.grdCPMBuy.TextMatrix(0, CPMPRICETYPE) = "Price Type"
    Contract.grdCPMBuy.TextMatrix(0, CPMPOSITION) = "Position"
    Contract.grdCPMBuy.TextMatrix(0, CPMPRIORITY) = "Priority"
    Contract.grdCPMBuy.TextMatrix(0, CPMPODTARGET) = "Ad Server Targeting"
    Contract.grdCPMBuy.TextMatrix(0, CPMSTARTDATE) = "Start Date"
    Contract.grdCPMBuy.TextMatrix(0, CPMENDDATE) = "End Date"
    '10879
'    Contract.grdCPMBuy.TextMatrix(0, CPMIMPRESSIONS) = "Impression Goal"
'    Contract.grdCPMBuy.TextMatrix(0, CPMIMPRESSIONCAP) = "Impression Cap"
'    Contract.grdCPMBuy.TextMatrix(0, CPMCPMVALUE) = "CPM"
    Contract.grdCPMBuy.TextMatrix(0, CPMIMPRESSIONS) = "Quantity"
    Contract.grdCPMBuy.TextMatrix(0, CPMIMPRESSIONCAP) = "Cap"
    Contract.grdCPMBuy.TextMatrix(0, CPMCPMVALUE) = "Rate"
    Contract.grdCPMBuy.TextMatrix(0, CPMLENGTH) = "Length"
    Contract.grdCPMBuy.TextMatrix(0, CPMCOST) = "Cost"
    Contract.grdCPMBuy.TextMatrix(0, CPMCOMMENT) = "C"
    
End Sub
Public Sub mClearCPMBuyGrid()

    Dim llRow As Long
    Dim llCol As Long

    Contract.grdCPMBuy.Redraw = False
    
    Contract.edcLineNo(2).Text = ""
    Contract.edcLineNo(3).Text = ""
    Contract.edcLineNo(4).Text = ""
    
    gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbHourglass
    'Blank rows within grid
'    gGrid_Clear grdUsersLog, True
    'Set color within cells
    Contract.grdCPMBuy.rowHeight(0) = fgBoxGridH + 15 'fgBoxGridH + 15
    Contract.grdCPMBuy.Row = 0
    For llCol = CPMEXPAND To CPMCOMMENT Step 1
        Contract.grdCPMBuy.Col = llCol
        Contract.grdCPMBuy.CellBackColor = WHITE
        Contract.grdCPMBuy.CellForeColor = vbBlue
        If llCol = CPMID Then
            Contract.grdCPMBuy.CellBackColor = LIGHTBLUE  'vbWhite
        End If
    Next llCol
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        '10967
'        For llCol = CPMEXPAND To CPMSORT Step 1
         For llCol = CPMEXPAND To CPMDELIVERYSTATUS Step 1
            Contract.grdCPMBuy.TextMatrix(llRow, llCol) = ""
            If llCol = CPMEXPAND Then
                Contract.grdCPMBuy.Row = llRow
                Contract.grdCPMBuy.Col = llCol
                Contract.grdCPMBuy.CellFontName = "Monotype Sorts"
                Contract.grdCPMBuy.CellFontSize = 7
                Contract.grdCPMBuy.CellForeColor = BLUE
            End If
            'If llCol = CPMCOST Then
            '    Contract.grdCPMBuy.Row = llRow
            '    Contract.grdCPMBuy.Col = llCol
            '    Contract.grdCPMBuy.CellBackColor = LIGHTYELLOW
            '    Contract.grdCPMBuy.CellForeColor = vbBlack
            'End If
            If llCol = CPMEXPAND Then
                Contract.grdCPMBuy.Col = CPMVEHICLE
                Contract.grdCPMBuy.CellFontItalic = True
            End If
        Next llCol
        Contract.grdCPMBuy.rowHeight(llRow) = fgBoxGridH + 15 'fgBoxGridH + 15
    Next llRow
    imLastColSorted = -1
    imLastSort = -1
    smHiddenSpace = "    "
    imInitNoRows = (Contract.grdCPMBuy.height - Contract.grdCPMBuy.rowHeight(0)) \ fgBoxGridH
    '10677
    Contract.grdCPMBuy.rows = imInitNoRows - 2
    Contract.grdCPMBuy.Redraw = True
    gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbDefault
    
End Sub
Public Sub mCPMEnableBox()

'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilIDNo As Integer
    Dim llRow As Long
    Dim llRif As Long
    Dim ilRcfCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    
    If imTerminate Then
        Exit Sub
    End If
    If (Contract.grdCPMBuy.Row < Contract.grdCPMBuy.FixedRows) Or (Contract.grdCPMBuy.Row >= Contract.grdCPMBuy.rows) Or (Contract.grdCPMBuy.Col < Contract.grdCPMBuy.FixedCols) Or (Contract.grdCPMBuy.Col >= Contract.grdCPMBuy.cols - 1) Then
        Exit Sub
    End If
    If (Not imUpdateAllowed) Then
        Contract.pbcClickFocus.SetFocus
        Exit Sub
    End If
    '10857
    If (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row, CPMBILLED) = "Y") And (Contract.grdCPMBuy.Col <> CPMIMPRESSIONS And Contract.grdCPMBuy.Col <> CPMENDDATE) Then
        Contract.pbcClickFocus.SetFocus
        Exit Sub
    End If
    Contract.mAllSetShow 3   'clear header
    imCPMFromArrow = False
    
    lmCPMEnableRow = Contract.grdCPMBuy.Row
    lmCPMEnableCol = Contract.grdCPMBuy.Col
    
    Contract.pbcLineArrow.Visible = False
    Contract.pbcLineArrow.Move Contract.grdCPMBuy.Left - Contract.pbcLineArrow.Width - 30, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row) + (Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) - Contract.pbcLineArrow.height) / 2
    Contract.pbcLineArrow.Visible = True
    Select Case Contract.grdCPMBuy.Col
        Case CPMEXPAND
            If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID) = "" Then
                Exit Sub
            End If
            If Not bgHideHiddenLines Then
                Exit Sub
            End If
            If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) <> "P" Then
                Exit Sub
            End If
            'Dan-t or blank =compressed and s = expanded for the expand package
            If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMEXPAND) = "t" Then
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMEXPAND) = "s"
            Else
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMEXPAND) = "t"
            End If
            For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID) = Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) Then
                    'SHOWROW is tested in other...showing?
                    If Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = False Then
                        Contract.grdCPMBuy.rowHeight(llRow) = fgBoxGridH
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = True
                    Else
                        Contract.grdCPMBuy.rowHeight(llRow) = 0
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = False
                    End If
                End If
            Next llRow
        Case CPMVEHICLE
        
            If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                'If Contract.pbcLineArrow.Visible Then
                '    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                'Else
                    Contract.pbcClickFocus.SetFocus
                'End If
                Exit Sub
            End If
            
            If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID) = "" Then
                ilIDNo = 0
                For ilLoop = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                    If Val(Contract.grdCPMBuy.TextMatrix(ilLoop, CPMID)) > ilIDNo Then
                        ilIDNo = Val(Contract.grdCPMBuy.TextMatrix(ilLoop, CPMID))
                    End If
                Next ilLoop
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID) = ilIDNo + 1
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMEXPAND) = ""
            End If
            Contract.mVehPop False
            If imTerminate Then
                Exit Sub
            End If
            Contract.lbcLnVehicle(igTabMapIndex).height = gListBoxHeight(Contract.lbcLnVehicle(igTabMapIndex).ListCount, 5)
            Contract.lbcPkVehicle(igTabMapIndex).height = gListBoxHeight(Contract.lbcPkVehicle(igTabMapIndex).ListCount, 5)
            Contract.edcLNDropDown.Width = Contract.grdCPMBuy.ColWidth(CPMVEHICLE)
            If tgSpf.iVehLen <= 40 Then
                Contract.edcLNDropDown.MaxLength = tgSpf.iVehLen
            Else
                Contract.edcLNDropDown.MaxLength = 20
            End If
            slStr = Contract.grdCPMBuy.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If Contract.grdCPMBuy.Row > Contract.grdCPMBuy.FixedRows Then
                    slStr = Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row - 1, Contract.grdCPMBuy.Col)
                End If
            End If
            Contract.edcLNDropDown.Text = slStr
            If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) = "P" Then
                gFindMatch slStr, 0, Contract.lbcPkVehicle(igTabMapIndex)
                If gLastFound(Contract.lbcPkVehicle(igTabMapIndex)) >= 0 Then
                    Contract.lbcPkVehicle(igTabMapIndex).ListIndex = gLastFound(Contract.lbcPkVehicle(igTabMapIndex))
                    imComboBoxIndex = Contract.lbcPkVehicle(igTabMapIndex).ListIndex
                Else
                    If Contract.lbcPkVehicle(igTabMapIndex).ListCount > 0 Then
                        If lmCPMEnableRow <= Contract.grdCPMBuy.FixedRows Then
                            Contract.lbcPkVehicle(igTabMapIndex).ListIndex = 0
                            imComboBoxIndex = Contract.lbcPkVehicle(igTabMapIndex).ListIndex
                        Else
                            slStr = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow - 1, CPMVEHICLE)
                            gFindMatch slStr, 0, Contract.lbcPkVehicle(igTabMapIndex)
                            If gLastFound(Contract.lbcPkVehicle(igTabMapIndex)) >= 0 Then
                                Contract.lbcPkVehicle(igTabMapIndex).ListIndex = gLastFound(Contract.lbcPkVehicle(igTabMapIndex))
                                imComboBoxIndex = Contract.lbcPkVehicle(igTabMapIndex).ListIndex
                            Else
                                Contract.lbcPkVehicle(igTabMapIndex).ListIndex = 0
                                imComboBoxIndex = Contract.lbcPkVehicle(igTabMapIndex).ListIndex
                            End If
                        End If
                    End If
                End If
            Else
                gFindMatch slStr, 0, Contract.lbcLnVehicle(igTabMapIndex)
                If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                    Contract.lbcLnVehicle(igTabMapIndex).ListIndex = gLastFound(Contract.lbcLnVehicle(igTabMapIndex))
                    imComboBoxIndex = Contract.lbcLnVehicle(igTabMapIndex).ListIndex
                Else
                    If Contract.lbcLnVehicle(igTabMapIndex).ListCount > 0 Then
                        If lmCPMEnableRow <= Contract.grdCPMBuy.FixedRows Then
                            Contract.lbcLnVehicle(igTabMapIndex).ListIndex = 0
                            imComboBoxIndex = Contract.lbcLnVehicle(igTabMapIndex).ListIndex
                        Else
                            slStr = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow - 1, CPMVEHICLE)
                            gFindMatch slStr, 0, Contract.lbcLnVehicle(igTabMapIndex)
                            If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                                Contract.lbcLnVehicle(igTabMapIndex).ListIndex = gLastFound(Contract.lbcLnVehicle(igTabMapIndex))
                                imComboBoxIndex = Contract.lbcLnVehicle(igTabMapIndex).ListIndex
                            Else
                                Contract.lbcLnVehicle(igTabMapIndex).ListIndex = 0
                                imComboBoxIndex = Contract.lbcLnVehicle(igTabMapIndex).ListIndex
                            End If
                        End If
                    End If
                End If
            End If
        Case CPMDAYPART
        
            Contract.edcLNDropDown.MaxLength = 0
            Contract.mPrgPop imLnRowNo, "C"
            If imTerminate Then
                Exit Sub
            End If
            Contract.lbcLnProgram(igTabMapIndex).height = gListBoxHeight(Contract.lbcLnProgram(igTabMapIndex).ListCount, 5)
            Contract.lbcPkProgram(igTabMapIndex).height = gListBoxHeight(Contract.lbcPkProgram(igTabMapIndex).ListCount, 5)
            Contract.edcLNDropDown.Width = Contract.grdCPMBuy.ColWidth(CPMVEHICLE)
            slStr = Contract.grdCPMBuy.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If Contract.grdCPMBuy.Row > Contract.grdCPMBuy.FixedRows Then
                    If Trim$(Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row, CPMVEHICLE)) = Trim$(Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row - 1, CPMVEHICLE)) Then
                        slStr = Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row - 1, Contract.grdCPMBuy.Col)
                    End If
                End If
            End If
            Contract.edcLNDropDown.MaxLength = 20
            Contract.edcLNDropDown.Text = slStr
            If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) = "P" Then
                gFindMatch slStr, 0, Contract.lbcPkProgram(igTabMapIndex)
                If gLastFound(Contract.lbcPkProgram(igTabMapIndex)) >= 0 Then
                    Contract.lbcPkProgram(igTabMapIndex).ListIndex = gLastFound(Contract.lbcPkProgram(igTabMapIndex))
                    imComboBoxIndex = Contract.lbcPkProgram(igTabMapIndex).ListIndex
                Else
                    If Contract.lbcPkProgram(igTabMapIndex).ListCount > 0 Then
                        Contract.lbcPkProgram(igTabMapIndex).ListIndex = 0
                        imComboBoxIndex = Contract.lbcPkProgram(igTabMapIndex).ListIndex
                    Else
                        Contract.lbcPkProgram(igTabMapIndex).ListIndex = -1
                        imComboBoxIndex = Contract.lbcPkProgram(igTabMapIndex).ListIndex
                        Contract.edcLNDropDown.Text = ""
                    End If
                End If
            Else
                gFindMatch slStr, 0, Contract.lbcLnProgram(igTabMapIndex)
                If gLastFound(Contract.lbcLnProgram(igTabMapIndex)) >= 0 Then
                    Contract.lbcLnProgram(igTabMapIndex).ListIndex = gLastFound(Contract.lbcLnProgram(igTabMapIndex))
                    imComboBoxIndex = Contract.lbcLnProgram(igTabMapIndex).ListIndex
                Else
                    If Contract.lbcLnProgram(igTabMapIndex).ListCount > 0 Then
                        Contract.lbcLnProgram(igTabMapIndex).ListIndex = 0
                        imComboBoxIndex = Contract.lbcLnProgram(igTabMapIndex).ListIndex
                    Else
                        Contract.lbcLnProgram(igTabMapIndex).ListIndex = -1
                        imComboBoxIndex = Contract.lbcLnProgram(igTabMapIndex).ListIndex
                        Contract.edcLNDropDown.Text = ""
                    End If
                End If
            End If
        'Case CPMCOPYTYPE
        '    Contract.lbcPrg(2).height = gListBoxHeight(Contract.lbcPrg(2).ListCount, 5)
        '    Contract.edcLNDropDown.Width = Contract.grdCPMBuy.ColWidth(CPMCOPYTYPE)
        '    Contract.edcLNDropDown.MaxLength = 20
        '    slStr = Contract.grdCPMBuy.Text
        '    If slStr = "Missing" Then
        '        slStr = ""
        '    End If
        '    If (slStr = "") Then
        '        'If Contract.grdCPMBuy.Row > Contract.grdCPMBuy.FixedRows Then
        '        '    slStr = Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row - 1, Contract.grdCPMBuy.Col)
        '        'End If
        '        slStr = "[None]"
        '    End If
        '    Contract.edcLNDropDown.Text = Trim$(slStr)
        '10879
        Case CPMLENGTH
            'returns default length
            ilRet = mPopCPMLength(lmCPMEnableRow)
            Contract.lbcPrg(6).height = gListBoxHeight(Contract.lbcPrg(6).ListCount, 6)
            Contract.edcLNDropDown.Width = Contract.grdCPMBuy.ColWidth(CPMLENGTH)
            Contract.edcLNDropDown.MaxLength = 3
            slStr = Contract.grdCPMBuy.Text
            'does this do anything?
            If slStr = "Missing" Then
                slStr = ""
            End If
            '10879+ force to default length if new line
            If slStr = "" And Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLENGTH) = "" Then
                slStr = ilRet
            End If
            Contract.edcLNDropDown.Text = slStr
        Case CPMPRICETYPE
            If Contract.grdCPMBuy.Text = "" Then
                Contract.grdCPMBuy.Text = "CPM"
            End If
        Case CPMPOSITION
            mPopCPMPosition lmCPMEnableRow
            Contract.lbcPrg(4).height = gListBoxHeight(Contract.lbcPrg(4).ListCount, 6)
            Contract.edcLNDropDown.Width = Contract.grdCPMBuy.ColWidth(CPMPOSITION)
            Contract.edcLNDropDown.MaxLength = 0
            slStr = Contract.grdCPMBuy.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                'If Contract.grdCPMBuy.Row > Contract.grdCPMBuy.FixedRows Then
                '    slStr = Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row - 1, Contract.grdCPMBuy.Col)
                'End If
                slStr = "[No Position]"
            End If
            Contract.edcLNDropDown.Text = Trim$(slStr)
        Case CPMPRIORITY
            mPopCPMPriority
            Contract.lbcPrg(5).height = gListBoxHeight(Contract.lbcPrg(5).ListCount, 6)
            Contract.edcLNDropDown.Width = Contract.grdCPMBuy.ColWidth(CPMPRIORITY)
            Contract.edcLNDropDown.MaxLength = 0
            slStr = Contract.grdCPMBuy.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                'If Contract.grdCPMBuy.Row > Contract.grdCPMBuy.FixedRows Then
                '    slStr = Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row - 1, Contract.grdCPMBuy.Col)
                'End If
                slStr = "1"
            End If
            Contract.edcLNDropDown.Text = Trim$(slStr)
        Case CPMPODTARGET
            Contract.lbcPrg(3).height = gListBoxHeight(Contract.lbcPrg(3).ListCount, 5)
            Contract.edcLNDropDown.Width = Contract.grdCPMBuy.ColWidth(CPMPODTARGET)
            Contract.edcLNDropDown.MaxLength = 80
            slStr = Contract.grdCPMBuy.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                'If Contract.grdCPMBuy.Row > Contract.grdCPMBuy.FixedRows Then
                '    slStr = Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row - 1, Contract.grdCPMBuy.Col)
                'End If
                slStr = "[None]"
            End If
            Contract.edcLNDropDown.Text = Trim$(slStr)
        Case CPMSTARTDATE
            Contract.edcLNDropDown.MaxLength = 10
            slStr = Contract.grdCPMBuy.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            'If (slStr = "") Then
            '    If Contract.grdCPMBuy.Row > Contract.grdCPMBuy.FixedRows Then
            '        slStr = Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row - 1, Contract.grdCPMBuy.Col)
            '    End If
            'End If
            If slStr <> "" Then
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                Contract.pbcCalendar_Paint
            Else
                slStr = Format$(gNow(), "m/d/yy")
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                slStr = ""
            End If
            Contract.edcLNDropDown.Text = slStr
        Case CPMENDDATE
            Contract.edcLNDropDown.MaxLength = 10
            slStr = Contract.grdCPMBuy.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            'If (slStr = "") Then
            '    If Contract.grdCPMBuy.Row > Contract.grdCPMBuy.FixedRows Then
            '        slStr = Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row - 1, Contract.grdCPMBuy.Col)
            '    End If
            'End If
            If slStr <> "" Then
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                Contract.pbcCalendar_Paint
            Else
                slStr = Format$(gNow(), "m/d/yy")
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                slStr = ""
            End If
            Contract.edcLNDropDown.Text = slStr
        Case CPMIMPRESSIONS
            Contract.edcLNDropDown.MaxLength = 10
            slStr = Contract.grdCPMBuy.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                'If Contract.grdCPMBuy.Row > Contract.grdCPMBuy.FixedRows Then
                '    slStr = Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row - 1, Contract.grdCPMBuy.Col)
                'End If
            End If
            Contract.edcLNDropDown.Text = slStr
        Case CPMIMPRESSIONCAP
            Contract.edcLNDropDown.MaxLength = 10
            slStr = Contract.grdCPMBuy.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                'If Contract.grdCPMBuy.Row > Contract.grdCPMBuy.FixedRows Then
                '    slStr = Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row - 1, Contract.grdCPMBuy.Col)
                'End If
            End If
            Contract.edcLNDropDown.Text = slStr
        Case CPMCPMVALUE
            Contract.edcLNDropDown.MaxLength = 10
            slStr = Contract.grdCPMBuy.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If Contract.lbcRateCard.ListIndex >= 0 Then
                    slNameCode = tmCntRateCard(Contract.lbcRateCard.ListIndex).sKey   'Traffic!lbcRateCardCode.List(lbcRateCard.ListIndex)' - 1)    'lbcRateCard.ListCount - lbcRateCard.ListIndex - 1)
                    ilRet = gParseItem(slNameCode, 3, "\", slCode)
                    slCode = Trim$(slCode)
                    ilRcfCode = CInt(slCode)
                    For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                        If (tgMRif(llRif).iVefCode = Val(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMVEFCODE))) And (tgMRif(llRif).iRcfCode = ilRcfCode) And (tgMRif(llRif).iRdfCode = Val(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMRDFCODE))) Then
                            slStr = gLongToStrDec(tgMRif(llRif).lPodCPM, 2)
                            Exit For
                        End If
                    Next llRif
                End If
            End If
            Contract.edcLNDropDown.Text = slStr
        Case CPMCOST
            Contract.edcLNDropDown.MaxLength = 10
            slStr = Contract.grdCPMBuy.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                slStr = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS), Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE)), "1000.00")
            End If
            Contract.edcLNDropDown.Text = slStr
        Case CPMCOMMENT
            imInLnComment = True
            Contract.lbcComment(1).height = gListBoxHeight(Contract.lbcComment(1).ListCount, 2)
            Contract.edcLNDropDown.Width = (2 * Traffic.pbcArial.TextWidth("[Comment]")) + 180
            Contract.edcLNDropDown.MaxLength = 20
            'gMoveTableCtrl Contract.pbcLine, Contract.edcLNDropDown, tmLnCtrls(PRICEINDEX).fBoxX + Contract.cmcLnDropDown.Width, tmLnCtrls(ilBoxNo).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15)
            'Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
            'Contract.lbcComment(1).Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width + Contract.cmcLnDropDown.Width - Contract.lbcComment(1).Width, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
            
            If Contract.lbcComment(1).Width < Contract.edcLNDropDown.Width Then
                Contract.lbcComment(1).Width = Contract.edcLNDropDown.Width + Contract.cmcLnDropDown.Width
            End If
            imChgMode = True
            If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOMMENT) <> "" Then
                Contract.lbcComment(1).Clear
                Contract.lbcComment(1).AddItem Trim$(Left$(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOMMENT), 20)), 0
                Contract.lbcComment(1).AddItem "[None]", 0
                Contract.edcLNDropDown.Text = Contract.lbcComment(1).List(1)
            Else
                Contract.lbcComment(1).Clear
                Contract.lbcComment(1).AddItem "[Comment]", 0
                Contract.lbcComment(1).AddItem "[None]", 0
                Contract.lbcComment(1).ListIndex = 0 '[None]
                Contract.edcLNDropDown.Text = Contract.lbcComment(1).List(0)
            End If
            Contract.edcComment.Text = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOMMENT)
            imComboBoxIndex = Contract.lbcComment(1).ListIndex
            slStr = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOMMSHOWON)
            If slStr = "" Then
                slStr = "NNDND"
            End If
            For ilLoop = 1 To 5 Step 1
                Contract.ckcCommOn(ilLoop - 1).Enabled = True
                If Mid(slStr, ilLoop, 1) = "Y" Then
                    Contract.ckcCommOn(ilLoop - 1).Value = vbChecked
                ElseIf Mid(slStr, ilLoop, 1) = "N" Then
                    Contract.ckcCommOn(ilLoop - 1).Value = vbUnchecked
                ElseIf Mid(slStr, ilLoop, 1) = "D" Then
                    Contract.ckcCommOn(ilLoop - 1).Value = vbUnchecked
                    Contract.ckcCommOn(ilLoop - 1).Enabled = False
                Else
                    Contract.ckcCommOn(ilLoop - 1).Value = vbUnchecked
                End If
            Next ilLoop
            imChgMode = False
            Screen.MousePointer = vbDefault
            If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                If Trim$(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOMMENT)) <> "" Then
                    imDoubleClickName = True
                    ilRet = Contract.mLnCommentBranch()
                Else
                    Contract.pbcClickFocus.SetFocus
                End If
                Exit Sub
            End If
            
    End Select
    mSetBGColorGray lmCPMEnableRow
    mCPMSetFocus
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Public Function mCPMSetShow()
     '10738 made into a function so I could see if impressions failed
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilVpf As Integer
    Dim ilVef As Integer
    Dim ilVefCode As Integer
    Dim ilLoop As Integer
    Dim slShowOn As String
    
    '10738
    mCPMSetShow = True
    Contract.pbcLineArrow.Visible = False
    If (lmCPMEnableRow >= Contract.grdCPMBuy.FixedRows) And (lmCPMEnableRow < Contract.grdCPMBuy.rows) Then
        Select Case lmCPMEnableCol
            Case CPMVEHICLE
                Contract.edcLNDropDown.Visible = False  'Set visibility
                Contract.cmcLnDropDown.Visible = False  'Set visibility
                Contract.lbcLnVehicle(igTabMapIndex).Visible = False
                Contract.pbcLbcVehicle.Visible = False
                Contract.lbcPkVehicle(igTabMapIndex).Visible = False
                slStr = Contract.edcLNDropDown.Text
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) <> slStr Then
                    '10967
                    If mVehicleSentToVendorAndCannotAlter(lmCPMEnableRow) Then
                        'Beep
                        MsgBox "Vehicle has already been sent to digital vendor", 64, "Cannot change vehicle"
                        'I was getting loop because drop down still had 'new' vehicle in it
                        Contract.edcLNDropDown.Text = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol)
                        Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMVEHICLE) = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMORIGVEHICLE)
                        If Contract.pbcLineArrow.Visible Then
                            Contract.pbcLineArrow.SetFocus
                        Else
                            Contract.pbcClickFocus.SetFocus
                        End If
                        Exit Function
                    End If
                    imCPMChg = True
                End If
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) = "H" Then
                    gFindMatch slStr, 0, Contract.lbcLnVehicle(igTabMapIndex)
                    If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                        slNameCode = tmVehicleCode2(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(lbcLnVehicle(igTabMapIndex)))
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        ilVefCode = CInt(slCode)
                        ilVef = gBinarySearchVef(CInt(slCode))
                        If ilVef <> -1 Then
                            '10/20/15: Disallow a schedule line to be changed from standard vehicle to package vehicle
                            'If (tgMVef(ilVef).sType = "R") Then
                            If (tgMVef(ilVef).sType = "R") Or ((tgMVef(ilVef).sType = "P")) Then
                                Beep
                                If Contract.pbcLineArrow.Visible Then
                                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                                Else
                                    Contract.pbcClickFocus.SetFocus
                                End If
                                Exit Function
                            End If
                        End If
                    End If
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = smHiddenSpace & slStr
                Else
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
                End If
                'Contract.grdCPMBuy.CellFontItalic = False
                'Get Vehicle
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMVEFCODE) = ""
                
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) = "P" Then
                    gFindMatch slStr, 0, Contract.lbcPkVehicle(1)
                    If gLastFound(Contract.lbcPkVehicle(1)) >= 0 Then
                        slNameCode = tgCPMPkVeh(gLastFound(Contract.lbcPkVehicle(1))).sKey    'lbcVehicle.List(gLastFound(lbcLnVehicle(igTabMapIndex)))
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMVEFCODE) = CInt(slCode)
                        'ilVpf = gBinarySearchVpf(CInt(slCode))
                        'If ilVpf <> -1 Then
                        '    If tgVpf(ilVpf).sGMedium = "M" Then
                        '        Contract.grdCPMBuy.CellFontItalic = True
                        '    End If
                        'End If
                    End If
               Else
                    gFindMatch slStr, 0, Contract.lbcLnVehicle(1)
                    If gLastFound(Contract.lbcLnVehicle(1)) >= 0 Then
                        slNameCode = tmVehicleCode2(gLastFound(Contract.lbcLnVehicle(1))).sKey    'lbcVehicle.List(gLastFound(lbcLnVehicle(igTabMapIndex)))
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        ilVefCode = CInt(slCode)
                        Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMVEFCODE) = ilVefCode
                        ilVef = gBinarySearchVef(ilVefCode)
                        If ilVef <> -1 Then
                            If tgMVef(ilVef).sType = "P" Then
                                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMPRICETYPE) = "CPM"
                                'Contract.grdCPMBuy.Col = CPMPRICETYPE
                                'Contract.grdCPMBuy.CellBackColor = LIGHTYELLOW
                                'Contract.grdCPMBuy.Col = CPMVEHICLE
                            End If
                        End If
                        'ilVpf = gBinarySearchVpf(CInt(slCode))
                        'If ilVpf <> -1 Then
                        '    If tgVpf(ilVpf).sGMedium = "P" Then
                        '        Contract.grdCPMBuy.CellFontItalic = True
                        '    End If
                        'End If
                    End If
                End If
                mSetFGColorForRows
            Case CPMDAYPART
                Contract.edcLNDropDown.Visible = False  'Set visibility
                Contract.cmcLnDropDown.Visible = False  'Set visibility
                Contract.lbcLnProgram(igTabMapIndex).Visible = False
                Contract.lbcPkProgram(igTabMapIndex).Visible = False
                slStr = Contract.edcLNDropDown.Text
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) <> slStr Then
                    imCPMChg = True
                End If
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMRDFCODE) = ""
                For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                    If slStr = Trim$(tgMRdf(ilRdf).sName) Then
                        Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMRDFCODE) = tgMRdf(ilRdf).iCode
                        Exit For
                    End If
                Next ilRdf
            'Case CPMCOPYTYPE
            '    Contract.edcLNDropDown.Visible = False  'Set visibility
            '    Contract.cmcLnDropDown.Visible = False  'Set visibility
            '    Contract.lbcPrg(2).Visible = False
            '    slStr = Contract.edcLNDropDown.Text
            '    If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) <> slStr Then
            '        imCPMChg = True
            '    End If
            '    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
            '10879
            Case CPMLENGTH
                Contract.edcLNDropDown.Visible = False
                'list box stuff
                Contract.cmcLnDropDown.Visible = False
                Contract.lbcPrg(6).Visible = False
                slStr = Contract.edcLNDropDown.Text
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) <> slStr Then
                    '10929 test length--let's non package go through
                    If mDigitalPackageLengthTest(slStr) Then
                        imCPMChg = True
                        Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
                        gFormatStr Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol), FMTLEAVEBLANK, 0, Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol)
                    Else
                        If Contract.pbcLineArrow.Visible Then
                            Contract.pbcLineArrow.SetFocus
                        Else
                            Contract.pbcClickFocus.SetFocus
                        End If
                        Exit Function
                    End If
                End If
            Case CPMPRICETYPE
                Contract.pbcSumType(3).Visible = False
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = "Flat Rate" Then
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE) = ""
                End If
            Case CPMPOSITION
                Contract.edcLNDropDown.Visible = False  'Set visibility
                Contract.cmcLnDropDown.Visible = False  'Set visibility
                Contract.lbcPrg(4).Visible = False
                slStr = Contract.edcLNDropDown.Text
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) <> slStr Then
                    imCPMChg = True
                End If
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
            Case CPMPRIORITY
                Contract.edcLNDropDown.Visible = False  'Set visibility
                Contract.cmcLnDropDown.Visible = False  'Set visibility
                Contract.lbcPrg(5).Visible = False
                slStr = Contract.edcLNDropDown.Text
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) <> slStr Then
                    imCPMChg = True
                End If
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
            Case CPMPODTARGET
                Contract.edcLNDropDown.Visible = False  'Set visibility
                Contract.cmcLnDropDown.Visible = False  'Set visibility
                Contract.lbcPrg(3).Visible = False
                slStr = Contract.edcLNDropDown.Text
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) <> slStr Then
                    imCPMChg = True
                End If
                If slStr = "[None]" Then
                    slStr = ""
                End If
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
            Case CPMSTARTDATE
                Contract.edcLNDropDown.Visible = False  'Set visibility
                Contract.plcCalendar.Visible = False
                Contract.cmcLnDropDown.Visible = False
                slStr = Contract.edcLNDropDown.Text
                If gDateValue(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol)) <> gDateValue(slStr) Then
                    imCPMChg = True
                End If
                slStr = Format(gAdjYear(slStr), "m/d/yy")
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
            Case CPMENDDATE
                Contract.edcLNDropDown.Visible = False  'Set visibility
                Contract.plcCalendar.Visible = False
                Contract.cmcLnDropDown.Visible = False
                slStr = Contract.edcLNDropDown.Text
                If gDateValue(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol)) <> gDateValue(slStr) Then
                    imCPMChg = True
                End If
                slStr = Format(gAdjYear(slStr), "m/d/yy")
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
            Case CPMIMPRESSIONS
                '10738
                If (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMBILLED) = "Y") Then
                    If Not mDigitalTestImpressionChangeIfBilled(lmCPMEnableRow, Contract.edcLNDropDown.Text) Then
                        '10738-more
                        mCPMSetShow = False
                        mDigitalSetValueToOriginal lmCPMEnableRow, CPMIMPRESSIONS, CPMORIGIMPRESSIONS
                        mCalculateValues
                        If Contract.pbcLineArrow.Visible Then
                            Contract.pbcLineArrow.SetFocus
                        Else
                            Contract.pbcClickFocus.SetFocus
                        End If
                        Exit Function
                    End If
                End If
                Contract.edcLNDropDown.Visible = False  'Set visibility
                slStr = Contract.edcLNDropDown.Text
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) <> slStr Then
                    imCPMChg = True
                    'gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
                    mChkForStdCPMPkg
                    mCalculateValues
                    gFormatStr Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol), FMTLEAVEBLANK, 0, Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol)
                End If
                ''gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                ''Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
                ''If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS) <> "" And (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE) <> "") Then
                ''    slStr = gMulStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS), Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE))
                ''    slStr = gDivStr(slStr, "1000")
                ''    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST) = slStr
                ''Else
                ''    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST) = ""
                ''End If
                'mChkForStdCPMPkg
            Case CPMIMPRESSIONCAP
                Contract.edcLNDropDown.Visible = False  'Set visibility
                slStr = Contract.edcLNDropDown.Text
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) <> slStr Then
                    imCPMChg = True
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
                    gFormatStr Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol), FMTLEAVEBLANK, 0, Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol)
                End If
            Case CPMCPMVALUE
                Contract.edcLNDropDown.Visible = False  'Set visibility
                slStr = Contract.edcLNDropDown.Text
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) <> slStr Then
                    imCPMChg = True
                    'gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
                    mCalculateValues
                    gFormatStr Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol), FMTLEAVEBLANK, 2, Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol)
                End If
                'gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                'Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
                'If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS) <> "" And (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE) <> "") Then
                '    slStr = gMulStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS), Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE))
                '    slStr = gDivStr(slStr, "1000.00")
                '    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST) = slStr
                'Else
                '    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST) = ""
                'End If
                mChkForStdCPMPkg
            Case CPMCOST
                Contract.edcLNDropDown.Visible = False  'Set visibility
                slStr = Contract.edcLNDropDown.Text
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) <> slStr Then
                    imCPMChg = True
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
                    mCalculateValues
                    gFormatStr Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol), FMTLEAVEBLANK, 2, Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol)
                End If
                'gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                'Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
                'If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS) <> "" And (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE) <> "") Then
                '    slStr = gMulStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS), Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE))
                '    slStr = gDivStr(slStr, "1000.00")
                '    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST) = slStr
                'Else
                '    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST) = ""
                'End If
            Case CPMCOMMENT
                imInLnComment = False
                If Contract.plcComment.Visible Then
                    Contract.plcComment.Visible = False  'Set visibility
                    'If (imUpdateAllowed) And (Not igRepNetChgAllowed) Then
                    '    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = gRemoveLeadingCrCl(Contract.edcComment.Text)
                    'End If
                End If
                imInCommentBranch = False
                Contract.edcLNDropDown.Visible = False
                Contract.cmcLnDropDown.Visible = False
                Contract.lbcComment(1).Visible = False
                Contract.ckcCommOn(2).Enabled = True    'Spot/Copy
                Contract.ckcCommOn(4).Enabled = True    'Insertion
                If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                    '12/28/18: Make work similar to header, click in field a second time caused the pop-up information to disappear
                    'If Contract.pbcLineArrow.Visible Then
                    '    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                    'Else
                        Contract.pbcClickFocus.SetFocus
                    'End If
                    Exit Function
                End If
               '10843 remove dead code
'                If tgSpf.sCLnStdQt = "Y" Then
'                    slStr = gRemoveLeadingCrCl(Contract.edcComment.Text)
'                Else
                    If Contract.lbcComment(1).ListIndex < 0 Then
                        Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = ""
                        slStr = ""
                    ElseIf Contract.lbcComment(1).ListIndex = 0 Then
                        Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = ""
                        slStr = ""  '"[None]"
                    Else
                        slStr = gRemoveLeadingCrCl(Contract.edcComment.Text)
                    End If
               ' End If
                If Trim$(slStr) <> "" Then
                    slShowOn = ""
                    For ilLoop = 0 To 4 Step 1
                        If Contract.ckcCommOn(ilLoop).Enabled Then
                            If Contract.ckcCommOn(ilLoop).Value = vbChecked Then
                                slShowOn = slShowOn & "Y"
                            Else
                                slShowOn = slShowOn & "N"
                            End If
                        Else
                            slShowOn = slShowOn & "D"
                        End If
                    Next ilLoop
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOMMSHOWON) = slShowOn
                Else
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOMMSHOWON) = "NNDND"
                End If
                'gSetShow Contract.pbcLine, slStr, tmLnCtrls(ilBoxNo)
                'smLnShow(LINECOMMINDEX, imLnRowNo) = tmLnCtrls(ilBoxNo).sShow
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) <> slStr Then
                    imCPMChg = True
                End If
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, lmCPMEnableCol) = slStr
        End Select
    End If
    If lmCPMEnableCol >= CPMIMPRESSIONS Or lmCPMEnableCol = CPMPRICETYPE Then
        Contract.mCntrTotal
    End If
    lmCPMEnableRow = -1
    lmCPMEnableCol = -1
    imCPMCtrlVisible = False
    Contract.mSetCommands
End Function



'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Public Sub mCPMSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer
    Dim llColWidth As Long

    If (Contract.grdCPMBuy.Row < Contract.grdCPMBuy.FixedRows) Or (Contract.grdCPMBuy.Row >= Contract.grdCPMBuy.rows) Or (Contract.grdCPMBuy.Col < Contract.grdCPMBuy.FixedCols) Or (Contract.grdCPMBuy.Col >= Contract.grdCPMBuy.cols - 1) Then
        Exit Sub
    End If
    
    Contract.pbcLineArrow.Visible = False
    Contract.pbcLineArrow.Move Contract.grdCPMBuy.Left - Contract.pbcLineArrow.Width - 30, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row) + (Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) - Contract.pbcLineArrow.height) / 2
    Contract.pbcLineArrow.Visible = True
    
    imCPMCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To Contract.grdCPMBuy.Col - 1 Step 1
        llColPos = llColPos + Contract.grdCPMBuy.ColWidth(ilCol)
    Next ilCol
    Select Case Contract.grdCPMBuy.Col
        Case CPMVEHICLE
            Contract.edcLNDropDown.Move Contract.grdCPMBuy.Left + llColPos + 15, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row), Contract.grdCPMBuy.ColWidth(Contract.grdCPMBuy.Col) - 30, Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) + 15
            Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
            Contract.edcLNDropDown.SelStart = 0
            Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
            Contract.edcLNDropDown.Visible = True
            Contract.cmcLnDropDown.Visible = True
            Contract.edcLNDropDown.SetFocus
        
            If Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height + Contract.lbcLnVehicle(igTabMapIndex).height <= Contract.grdCPMBuy.Top + Contract.grdCPMBuy.height Then
                Contract.lbcLnVehicle(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
                Contract.lbcPkVehicle(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
            Else
                Contract.lbcLnVehicle(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.lbcLnVehicle(igTabMapIndex).height
                Contract.lbcPkVehicle(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.lbcPkVehicle(igTabMapIndex).height
            End If
            If Contract.lbcLnVehicle(igTabMapIndex).Width < Contract.edcLNDropDown.Width Then
                Contract.lbcLnVehicle(igTabMapIndex).Width = Contract.edcLNDropDown.Width + Contract.cmcLnDropDown.Width
                Contract.lbcPkVehicle(igTabMapIndex).Width = Contract.lbcLnVehicle(igTabMapIndex).Width
            End If
        
        Case CPMDAYPART
            Contract.edcLNDropDown.Move Contract.grdCPMBuy.Left + llColPos + 15, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row), Contract.grdCPMBuy.ColWidth(Contract.grdCPMBuy.Col) - 30, Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) + 15
            Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
        
            If Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height + Contract.lbcLnProgram(igTabMapIndex).height <= Contract.grdCPMBuy.Top + Contract.grdCPMBuy.height Then
                Contract.lbcLnProgram(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
                Contract.lbcPkProgram(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
            Else
                Contract.lbcLnProgram(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.lbcLnProgram(igTabMapIndex).height
                Contract.lbcPkProgram(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.lbcPkProgram(igTabMapIndex).height
            End If
            If Contract.lbcLnVehicle(igTabMapIndex).Width < Contract.edcLNDropDown.Width Then
                Contract.lbcLnProgram(igTabMapIndex).Width = Contract.edcLNDropDown.Width + Contract.cmcLnDropDown.Width
                Contract.lbcPkProgram(igTabMapIndex).Width = Contract.lbcLnProgram(igTabMapIndex).Width
            End If
            Contract.edcLNDropDown.SelStart = 0
            Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
            Contract.edcLNDropDown.Visible = True
            Contract.cmcLnDropDown.Visible = True
            Contract.edcLNDropDown.SetFocus
        'Case CPMCOPYTYPE
        '    Contract.edcLNDropDown.Move Contract.grdCPMBuy.Left + llColPos + 15, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row), Contract.grdCPMBuy.ColWidth(Contract.grdCPMBuy.Col) - 30, Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) + 15
        '    Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
        '
        '    If Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height + Contract.lbcPrg(2).height <= Contract.grdCPMBuy.Top + Contract.grdCPMBuy.height Then
        '        Contract.lbcPrg(2).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
        '    Else
        '        Contract.lbcPrg(2).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.lbcLnProgram(igTabMapIndex).height
        '    End If
        '    If Contract.lbcPrg(2).Width < Contract.edcLNDropDown.Width Then
        '        Contract.lbcPrg(2).Width = Contract.edcLNDropDown.Width + Contract.cmcLnDropDown.Width
        '    End If
        '    Contract.edcLNDropDown.SelStart = 0
        '    Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
        '    Contract.edcLNDropDown.Visible = True
        '    Contract.cmcLnDropDown.Visible = True
        '    Contract.edcLNDropDown.SetFocus
        '10879
        Case CPMLENGTH
            Contract.edcLNDropDown.Move Contract.grdCPMBuy.Left + llColPos + 15, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row), Contract.grdCPMBuy.ColWidth(Contract.grdCPMBuy.Col) - 30, Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) + 15
            'list box button
            Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
            'list box stuff
            If Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height + Contract.lbcPrg(6).height <= Contract.grdCPMBuy.Top + Contract.grdCPMBuy.height Then
                Contract.lbcPrg(6).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
            Else
                Contract.lbcPrg(6).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.lbcLnProgram(igTabMapIndex).height
            End If
            If Contract.lbcPrg(6).Width < Contract.edcLNDropDown.Width Then
                Contract.lbcPrg(6).Width = Contract.edcLNDropDown.Width + Contract.cmcLnDropDown.Width
            End If
            Contract.edcLNDropDown.SelStart = 0
            Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
            Contract.edcLNDropDown.Visible = True
            'list box button
            Contract.cmcLnDropDown.Visible = True
            Contract.edcLNDropDown.SetFocus
        Case CPMPRICETYPE
            Contract.pbcSumType_Paint 3
            Contract.pbcSumType(3).Move Contract.grdCPMBuy.Left + llColPos + 15, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row), Contract.grdCPMBuy.ColWidth(Contract.grdCPMBuy.Col) - 30, Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) + 15
            Contract.pbcSumType(3).Visible = True
            Contract.pbcSumType(3).SetFocus
            
        Case CPMPOSITION
            Contract.edcLNDropDown.Move Contract.grdCPMBuy.Left + llColPos + 15, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row), Contract.grdCPMBuy.ColWidth(Contract.grdCPMBuy.Col) - 30, Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) + 15
            Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
        
            If Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height + Contract.lbcPrg(4).height <= Contract.grdCPMBuy.Top + Contract.grdCPMBuy.height Then
                Contract.lbcPrg(4).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
            Else
                Contract.lbcPrg(4).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.lbcLnProgram(igTabMapIndex).height
            End If
            If Contract.lbcPrg(4).Width < Contract.edcLNDropDown.Width Then
                Contract.lbcPrg(4).Width = Contract.edcLNDropDown.Width + Contract.cmcLnDropDown.Width
            End If
            Contract.edcLNDropDown.SelStart = 0
            Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
            Contract.edcLNDropDown.Visible = True
            Contract.cmcLnDropDown.Visible = True
            Contract.edcLNDropDown.SetFocus
       
        Case CPMPRIORITY
            Contract.edcLNDropDown.Move Contract.grdCPMBuy.Left + llColPos + 15, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row), Contract.grdCPMBuy.ColWidth(Contract.grdCPMBuy.Col) - 30, Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) + 15
            Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
        
            If Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height + Contract.lbcPrg(5).height <= Contract.grdCPMBuy.Top + Contract.grdCPMBuy.height Then
                Contract.lbcPrg(5).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
            Else
                Contract.lbcPrg(5).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.lbcLnProgram(igTabMapIndex).height
            End If
            If Contract.lbcPrg(5).Width < Contract.edcLNDropDown.Width Then
                Contract.lbcPrg(5).Width = Contract.edcLNDropDown.Width + Contract.cmcLnDropDown.Width
            End If
            Contract.edcLNDropDown.SelStart = 0
            Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
            Contract.edcLNDropDown.Visible = True
            Contract.cmcLnDropDown.Visible = True
            Contract.edcLNDropDown.SetFocus
       
        Case CPMPODTARGET
            Contract.edcLNDropDown.Move Contract.grdCPMBuy.Left + llColPos + 15, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row), Contract.grdCPMBuy.ColWidth(Contract.grdCPMBuy.Col) - 30, Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) + 15
            Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
        
            If Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height + Contract.lbcPrg(3).height <= Contract.grdCPMBuy.Top + Contract.grdCPMBuy.height Then
                Contract.lbcPrg(3).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
            Else
                Contract.lbcPrg(3).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.lbcLnProgram(igTabMapIndex).height
            End If
            If Contract.lbcPrg(3).Width < Contract.edcLNDropDown.Width Then
                Contract.lbcPrg(3).Width = Contract.edcLNDropDown.Width + Contract.cmcLnDropDown.Width
            End If
            Contract.edcLNDropDown.SelStart = 0
            Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
            Contract.edcLNDropDown.Visible = True
            Contract.cmcLnDropDown.Visible = True
            Contract.edcLNDropDown.SetFocus
        Case CPMSTARTDATE
            Contract.edcLNDropDown.Move Contract.grdCPMBuy.Left + llColPos + 15, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row), Contract.grdCPMBuy.ColWidth(Contract.grdCPMBuy.Col) - 30, Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) + 15
            'If Contract.edcLNDropDown.Top + Contract.edcLNDropDown.Height + Contract.plcCalendar.Height < cmcDone.Top Then
                Contract.plcCalendar.Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
            'Else
            '    Contract.plcCalendar.Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.plcCalendar.Height
            'End If
            Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
            Contract.plcCalendar.ZOrder
            Contract.edcLNDropDown.Visible = True
            Contract.cmcLnDropDown.Visible = True
            Contract.edcLNDropDown.SetFocus
        Case CPMENDDATE
            Contract.edcLNDropDown.Move Contract.grdCPMBuy.Left + llColPos + 15, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row), Contract.grdCPMBuy.ColWidth(Contract.grdCPMBuy.Col) - 30, Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) + 15
            'If Contract.edcLNDropDown.Top + Contract.edcLNDropDown.Height + Contract.plcCalendar.Height < cmcDone.Top Then
                Contract.plcCalendar.Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
            'Else
            '    Contract.plcCalendar.Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.plcCalendar.Height
            'End If
            Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
            Contract.plcCalendar.ZOrder
            Contract.edcLNDropDown.Visible = True
            Contract.cmcLnDropDown.Visible = True
            Contract.edcLNDropDown.SetFocus
        Case CPMIMPRESSIONS
            Contract.edcLNDropDown.Move Contract.grdCPMBuy.Left + llColPos + 15, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row), Contract.grdCPMBuy.ColWidth(Contract.grdCPMBuy.Col) - 30, Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) + 15
            Contract.edcLNDropDown.SelStart = 0
            Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
            Contract.edcLNDropDown.Visible = True
            '10738 error here in odd situations
        On Error Resume Next
            Contract.edcLNDropDown.SetFocus
        Case CPMIMPRESSIONCAP
            Contract.edcLNDropDown.Move Contract.grdCPMBuy.Left + llColPos + 15, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row), Contract.grdCPMBuy.ColWidth(Contract.grdCPMBuy.Col) - 30, Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) + 15
            Contract.edcLNDropDown.SelStart = 0
            Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
            Contract.edcLNDropDown.Visible = True
            Contract.edcLNDropDown.SetFocus
        Case CPMCPMVALUE
            Contract.edcLNDropDown.Move Contract.grdCPMBuy.Left + llColPos + 15, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row), Contract.grdCPMBuy.ColWidth(Contract.grdCPMBuy.Col) - 30, Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) + 15
            Contract.edcLNDropDown.SelStart = 0
            Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
            Contract.edcLNDropDown.Visible = True
            Contract.edcLNDropDown.SetFocus
        Case CPMCOST
            Contract.edcLNDropDown.Move Contract.grdCPMBuy.Left + llColPos + 15, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row), Contract.grdCPMBuy.ColWidth(Contract.grdCPMBuy.Col) - 30, Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) + 15
            Contract.edcLNDropDown.SelStart = 0
            Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
            Contract.edcLNDropDown.Visible = True
            Contract.edcLNDropDown.SetFocus
        Case CPMCOMMENT
            'gMoveTableCtrl Contract.pbcLine, Contract.edcLNDropDown, tmLnCtrls(PRICEINDEX).fBoxX + Contract.cmcLnDropDown.Width, tmLnCtrls(ilBoxNo).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15)
            'Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
            'Contract.lbcComment(1).Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width + Contract.cmcLnDropDown.Width - Contract.lbcComment(1).Width, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
            Contract.cmcLnDropDown.Move Contract.grdCPMBuy.Left + Contract.grdCPMBuy.ColPos(CPMCOMMENT), Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row)
            Contract.edcLNDropDown.Move Contract.cmcLnDropDown.Left - Contract.edcLNDropDown.Width - 15, Contract.cmcLnDropDown.Top, Contract.edcLNDropDown.Width, Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) + 15
       
            If Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height + Contract.lbcComment(1).height <= Contract.grdCPMBuy.Top + Contract.grdCPMBuy.height Then
                Contract.lbcComment(1).Move Contract.cmcLnDropDown.Left - Contract.lbcComment(1).Width - 15, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
            Else
                Contract.lbcComment(1).Move Contract.cmcLnDropDown.Left - Contract.lbcComment(1).Width - 15, Contract.edcLNDropDown.Top - Contract.lbcComment(1).height
            End If
            Contract.edcLNDropDown.SelStart = 0
            Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
            Contract.edcLNDropDown.Visible = False   'True
            Contract.cmcLnDropDown.Visible = True
            Contract.cmcLnDropDown.SetFocus
    End Select
End Sub
Public Sub gCPMSTab()
    Dim ilPrev As Integer
    Dim llEnableRow As Long
    Dim ilVef As Integer
    
    If imCPMCtrlVisible Then
        'If Contract.grdCPMBuy.Col = CPMCOPYTYPE Then
        '    If mCopyTypeBranch() Then
        '        Exit Sub
        '    End If
        'End If
        If Contract.grdCPMBuy.Col = CPMPODTARGET Then
            If mPodTargetBranch() Then
                Exit Sub
            End If
        End If
        If Contract.grdCPMBuy.Col = CPMCOMMENT Then
            If Contract.mLnCommentBranch() Then
                Exit Sub
            End If
        End If
        '10738 test billed here to stay in field
        If Contract.grdCPMBuy.Col = CPMIMPRESSIONS And (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMBILLED) = "Y") Then
            If Not mDigitalTestImpressionChangeIfBilled(lmCPMEnableRow, Contract.edcLNDropDown.Text) Then
                mDigitalSetValueToOriginal lmCPMEnableRow, CPMIMPRESSIONS, CPMORIGIMPRESSIONS
                mCalculateValues
                Exit Sub
            End If
        End If
        llEnableRow = lmCPMEnableRow
        'this is done first because the loop may be moving forward
        mCPMSetShow
        Do
            ' ilPrev checking to see if allowed going to next
            ilPrev = False
            If Contract.grdCPMBuy.Col = CPMVEHICLE Then
                If Contract.grdCPMBuy.Row > Contract.grdCPMBuy.FixedRows Then
                    lmCPMTopRow = -1
                    Contract.grdCPMBuy.Row = Contract.grdCPMBuy.Row - 1
                    If Not Contract.grdCPMBuy.RowIsVisible(Contract.grdCPMBuy.Row) Then
                        Contract.grdCPMBuy.TopRow = Contract.grdCPMBuy.TopRow - 1
                    End If
                    '10857 test billing on previous line
                    If Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row, CPMBILLED) = "Y" Then
                        Contract.pbcClickFocus.SetFocus
                    Else
                        '10859 go to comments instead
                        'Contract.grdCPMBuy.Col = CPMCPMVALUE
                        Contract.grdCPMBuy.Col = CPMCOMMENT
                        mCPMEnableBox
                    End If
                Else
                    Contract.cmcCancel.SetFocus
                End If
            ElseIf Contract.grdCPMBuy.Col = CPMPOSITION Then
                If (Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMLNTYPE) = "P") Or (Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMLNTYPE) = "H") Then
                    Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col - 2
                Else
                    ilVef = gBinarySearchVef(Val(Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMVEFCODE)))
                    If ilVef <> -1 Then
                        If (tgMVef(ilVef).sType = "P") Or (Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMLNTYPE) = "H") Then
                            Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col - 2
                        Else
                            Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col - 1
                        End If
                    Else
                        Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col - 1
                    End If
                End If
                mCPMEnableBox
            ElseIf (Contract.grdCPMBuy.Col = CPMPODTARGET) Or (Contract.grdCPMBuy.Col = CPMPRIORITY) Then
                If (Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMLNTYPE) = "P") Then
                    Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col - 1
                    ilPrev = True
                Else
                    ilVef = gBinarySearchVef(Val(Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMVEFCODE)))
                    If ilVef <> -1 Then
                        If (tgMVef(ilVef).sType = "P") Then
                            Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col - 1
                            ilPrev = True
                        End If
                    End If
                End If
                If ilPrev = False Then
                    Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col - 1
                    mCPMEnableBox
                End If
             
             ElseIf Contract.grdCPMBuy.Col = CPMCOST Then
                If Trim$(Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMPRICETYPE)) = "Flat Rate" Then
                    Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col - 2
                Else
                    Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col - 1
                End If
                mCPMEnableBox
            ElseIf Contract.grdCPMBuy.Col = CPMENDDATE And (Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMBILLED) = "Y") Then
                Contract.pbcClickFocus.SetFocus
                Exit Sub
            Else
                Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col - 1
                mCPMEnableBox
            End If
        Loop While ilPrev
    Else
        '10857 from header/comment If billed, go to 'clickFocus'. I used 'fixed rows'
        If Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMBILLED) = "Y" Then
            Contract.pbcClickFocus.SetFocus
        Else
            lmCPMTopRow = -1
            Contract.grdCPMBuy.TopRow = Contract.grdCPMBuy.FixedRows
            Contract.grdCPMBuy.Col = CPMVEHICLE
            Contract.grdCPMBuy.Row = Contract.grdCPMBuy.FixedRows
            mCPMEnableBox
        End If
    End If

End Sub
Public Sub gCPMTab()
    Dim llRow As Long
    Dim ilCol As Integer
    Dim ilNext As Integer
    Dim llEnableRow As Long
    Dim slStr As String
    Dim llSvRow As Long
    Dim llSvCol As Long
    Dim ilVef As Integer

    If imCPMFromArrow Then
        imCPMFromArrow = False
        Contract.grdCPMBuy.Col = CPMVEHICLE
        mCPMEnableBox
        Exit Sub
    End If
    If imCPMCtrlVisible Then
        llSvRow = Contract.grdCPMBuy.Row
        llSvCol = Contract.grdCPMBuy.Col
        'If Contract.grdCPMBuy.Col = CPMCOPYTYPE Then
        '    If mCopyTypeBranch() Then
        '        Exit Sub
        '    End If
        'End If
        If Contract.grdCPMBuy.Col = CPMPODTARGET Then
            If mPodTargetBranch() Then
                Exit Sub
            End If
        End If
        If Contract.grdCPMBuy.Col = CPMCOMMENT Then
            If Contract.mLnCommentBranch() Then
                Exit Sub
            End If
        End If
        '10738 test billed here to stay in field
        If Contract.grdCPMBuy.Col = CPMIMPRESSIONS And (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMBILLED) = "Y") Then
            If Not mDigitalTestImpressionChangeIfBilled(lmCPMEnableRow, Contract.edcLNDropDown.Text) Then
                mDigitalSetValueToOriginal lmCPMEnableRow, CPMIMPRESSIONS, CPMORIGIMPRESSIONS
                mCalculateValues
                Exit Sub
            End If
        End If
        llEnableRow = lmCPMEnableRow
        mCPMSetShow
        Do
            ilNext = False
            If Contract.grdCPMBuy.Col = CPMCOMMENT Then 'CPMCOST Then
                '10857 billing test   make sure safe to test next line
                If Contract.grdCPMBuy.Row + 1 < Contract.grdCPMBuy.rows Then
                    If (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row + 1, CPMBILLED) = "Y") Then
                        Contract.pbcClickFocus.SetFocus
                    Else
                        '10859 and refactored CODE REMOVED
                        mDigitalTabToNextLine
                    End If
                Else
                    mDigitalTabToNextLine
                End If ' billed?
            ElseIf Contract.grdCPMBuy.Col = CPMVEHICLE Then
                If Trim$(Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMVEHICLE)) = "" Then
                    Contract.pbcClickFocus.SetFocus
                Else
                    Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 1
                    mCPMEnableBox
                End If
            '10879-2
'            ElseIf Contract.grdCPMBuy.Col = CPMDAYPART Then
            ElseIf Contract.grdCPMBuy.Col = CPMLENGTH Then
                If (Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMLNTYPE) = "P") Or (Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMLNTYPE) = "H") Then
                    If (Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMLNTYPE) = "P") Then
                        Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 4
                    Else
                        Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 2
                    End If
                Else
                    ilVef = gBinarySearchVef(Val(Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMVEFCODE)))
                    If ilVef <> -1 Then
                        If (tgMVef(ilVef).sType = "P") Or (Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMLNTYPE) = "H") Then
                            If (tgMVef(ilVef).sType = "P") Then
                                Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 4
                            Else
                                Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 2
                            End If
                        Else
                            Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 1
                        End If
                    Else
                        Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 1
                    End If
                End If
                mCPMEnableBox
            ElseIf (Contract.grdCPMBuy.Col = CPMPRICETYPE) Or (Contract.grdCPMBuy.Col = CPMPOSITION) Then
                If (Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMLNTYPE) = "P") Then
                    Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 1
                    ilNext = True
                Else
                    ilVef = gBinarySearchVef(Val(Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMVEFCODE)))
                    If ilVef <> -1 Then
                        If (tgMVef(ilVef).sType = "P") Then
                            Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 1
                            ilNext = True
                        End If
                    End If
                End If
                If ilNext = False Then
                    Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 1
                    mCPMEnableBox
                End If
            ElseIf Contract.grdCPMBuy.Col = CPMIMPRESSIONCAP Then
                If Trim$(Contract.grdCPMBuy.TextMatrix(llEnableRow, CPMPRICETYPE)) = "Flat Rate" Then
                    Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 2
                Else
                    Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 1
                End If
                mCPMEnableBox
            ElseIf Contract.grdCPMBuy.Col = CPMSTARTDATE Then
                slStr = Contract.edcLNDropDown.Text
                If slStr <> "" Then
                    If Not gValidDate(slStr) Then
                        '10738 refactoring
                        mDigitalSetValueToOriginal llSvRow, llSvCol, CPMORIGSTARTDATE
'                        Beep
'                        Contract.grdCPMBuy.Row = llSvRow
'                        Contract.grdCPMBuy.Col = llSvCol
'                        Contract.grdCPMBuy.TextMatrix(llSvRow, llSvCol) = Contract.grdCPMBuy.TextMatrix(llSvRow, CPMORIGSTARTDATE)
'                        mCPMEnableBox
                        Exit Sub
                    End If
                Else
                    '10738 refactoring
                    mDigitalSetValueToOriginal llSvRow, llSvCol, CPMORIGSTARTDATE
'                    Beep
'                    Contract.grdCPMBuy.Row = llSvRow
'                    Contract.grdCPMBuy.Col = llSvCol
'                    Contract.grdCPMBuy.TextMatrix(llSvRow, llSvCol) = Contract.grdCPMBuy.TextMatrix(llSvRow, CPMORIGSTARTDATE)
'                    mCPMEnableBox
                    Exit Sub
                End If
                Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 1
                mCPMEnableBox
            ElseIf Contract.grdCPMBuy.Col = CPMENDDATE Then
                slStr = Contract.edcLNDropDown.Text
                If slStr <> "" Then
                    If Not gValidDate(slStr) Then
                        '10738 refactoring
                        mDigitalSetValueToOriginal llSvRow, llSvCol, CPMORIGENDDATE
'                        Beep
'                        Contract.grdCPMBuy.Row = llSvRow
'                        Contract.grdCPMBuy.Col = llSvCol
'                        Contract.grdCPMBuy.TextMatrix(llSvRow, llSvCol) = Contract.grdCPMBuy.TextMatrix(llSvRow, CPMORIGENDDATE)
'                        mCPMEnableBox
                        Exit Sub
                    End If
                Else
                    '10738 refactoring
                    mDigitalSetValueToOriginal llSvRow, llSvCol, CPMORIGENDDATE
'                    Beep
'                    Contract.grdCPMBuy.Row = llSvRow
'                    Contract.grdCPMBuy.Col = llSvCol
'                    Contract.grdCPMBuy.TextMatrix(llSvRow, llSvCol) = Contract.grdCPMBuy.TextMatrix(llSvRow, CPMORIGENDDATE)
'                    mCPMEnableBox
                    Exit Sub
                End If
                Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 1
                mCPMEnableBox
            '10738 added
            ElseIf Contract.grdCPMBuy.Col = CPMIMPRESSIONS Then
                If (Contract.grdCPMBuy.TextMatrix(llSvRow, CPMBILLED) = "Y") Then
                    Contract.pbcClickFocus.SetFocus
                    Exit Sub
                Else
                    Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 1
                    mCPMEnableBox
                End If
            Else
                Contract.grdCPMBuy.Col = Contract.grdCPMBuy.Col + 1
                mCPMEnableBox
            End If
        Loop While ilNext
    Else
        '10857 test if billed and if is, pbcClickFocus.setfocus.  Did not test-couldn't hit this spot
        If (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row, CPMBILLED) <> "Y") Then
            lmCPMTopRow = -1
            Contract.grdCPMBuy.TopRow = Contract.grdCPMBuy.FixedRows
            Contract.grdCPMBuy.Col = CPMVEHICLE
            Contract.grdCPMBuy.Row = Contract.grdCPMBuy.FixedRows
            mCPMEnableBox
        Else
            Contract.pbcClickFocus.SetFocus
        End If
    End If

End Sub



Public Sub gCPMDropdown(ilSource As Integer, Optional ilValue As Integer = -1)
    Dim ilKey As Integer
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilRet As Integer
    Dim slMaxValue As String
    
    If lmCPMEnableRow < 0 Then
        Exit Sub
    End If
    Select Case ilSource
        Case 1  'cmcLNDropdown
            Select Case Contract.grdCPMBuy.Col
                Case CPMVEHICLE
                    If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) <> "P" Then
                        Contract.lbcLnVehicle(igTabMapIndex).Visible = Not Contract.lbcLnVehicle(igTabMapIndex).Visible
                        imINPBCPaint = False
                        If Contract.lbcLnVehicle(igTabMapIndex).Visible Then
                            Contract.pbcLbcVehicle.Visible = True
                        Else
                            Contract.pbcLbcVehicle.Visible = False
                        End If
                    Else
                        Contract.lbcPkVehicle(igTabMapIndex).Visible = Not Contract.lbcPkVehicle(igTabMapIndex).Visible
                        If Contract.lbcPkVehicle(igTabMapIndex).Visible Then
                            Contract.pbcLbcVehicle.Visible = True
                        Else
                            Contract.pbcLbcVehicle.Visible = False
                        End If
                    End If
                Case CPMDAYPART
                    If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) <> "P" Then
                        Contract.lbcLnProgram(igTabMapIndex).Visible = Not Contract.lbcLnProgram(igTabMapIndex).Visible
                    Else
                        Contract.lbcPkProgram(igTabMapIndex).Visible = Not Contract.lbcPkProgram(igTabMapIndex).Visible
                    End If
                'Case CPMCOPYTYPE
                '    Contract.lbcPrg(2).Visible = Not Contract.lbcPrg(2).Visible
                '10879
                Case CPMLENGTH
                    Contract.lbcPrg(6).Visible = Not Contract.lbcPrg(6).Visible
                Case CPMPRICETYPE
                Case CPMPOSITION
                    Contract.lbcPrg(4).Visible = Not Contract.lbcPrg(4).Visible
                Case CPMPRIORITY
                    Contract.lbcPrg(5).Visible = Not Contract.lbcPrg(5).Visible
                Case CPMPODTARGET
                    Contract.lbcPrg(3).Visible = Not Contract.lbcPrg(3).Visible
                Case CPMSTARTDATE
                    Contract.plcCalendar.Visible = Not Contract.plcCalendar.Visible
                Case CPMENDDATE
                    Contract.plcCalendar.Visible = Not Contract.plcCalendar.Visible
                Case CPMIMPRESSIONS
                Case CPMIMPRESSIONCAP
                Case CPMCPMVALUE
                Case CPMCOST
                Case CPMCOMMENT
                    Contract.lbcComment(1).Visible = Not Contract.lbcComment(1).Visible
                    If Contract.lbcComment(1).Visible Then
                        Contract.edcLNDropDown.Visible = True
                    End If
                    Contract.edcLNDropDown.SelStart = 0
                    Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
                    Exit Sub
            End Select
            
        Case 2  'edcLnDropDown_Change
            Select Case Contract.grdCPMBuy.Col
                Case CPMVEHICLE
                    If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) <> "P" Then
                        imLbcArrowSetting = True
                        gMatchLookAhead Contract.edcLNDropDown, Contract.lbcLnVehicle(igTabMapIndex), imBSMode, imComboBoxIndex
                    Else
                        imLbcArrowSetting = True
                        ilRet = gOptionalLookAhead(Contract.edcLNDropDown, Contract.lbcPkVehicle(igTabMapIndex), imBSMode, slStr)
                        imLbcArrowSetting = False
                    End If
                    Contract.pbcLbcVehicle_Paint
                Case CPMDAYPART
                    If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) <> "P" Then
                        imLbcArrowSetting = True
                        gMatchLookAhead Contract.edcLNDropDown, Contract.lbcLnProgram(igTabMapIndex), imBSMode, imComboBoxIndex
                    Else
                        'imLbcArrowSetting = True
                        'ilRet = gOptionalLookAhead(edcLnDropDown, lbcPkProgram(igTabMapIndex), imBSMode, slStr)
                        'imLbcArrowSetting = False
                        imLbcArrowSetting = True
                        gMatchLookAhead Contract.edcLNDropDown, Contract.lbcPkProgram(igTabMapIndex), imBSMode, imComboBoxIndex
                        imLbcArrowSetting = False
                    End If
                'Case CPMCOPYTYPE
                '    imLbcArrowSetting = True
                '    gMatchLookAhead Contract.edcLNDropDown, Contract.lbcPrg(2), imBSMode, imComboBoxIndex
                '10879
                Case CPMLENGTH
                    imLbcArrowSetting = True
                    gMatchLookAhead Contract.edcLNDropDown, Contract.lbcPrg(6), imBSMode, imComboBoxIndex
                Case CPMPRICETYPE
                Case CPMPOSITION
                    imLbcArrowSetting = True
                    gMatchLookAhead Contract.edcLNDropDown, Contract.lbcPrg(4), imBSMode, imComboBoxIndex
                Case CPMPRIORITY
                    imLbcArrowSetting = True
                    gMatchLookAhead Contract.edcLNDropDown, Contract.lbcPrg(5), imBSMode, imComboBoxIndex
                Case CPMPODTARGET
                    imLbcArrowSetting = True
                    gMatchLookAhead Contract.edcLNDropDown, Contract.lbcPrg(3), imBSMode, imComboBoxIndex
                Case CPMSTARTDATE
                    Contract.pbcCalendar_Paint   'mBoxCalDate called within paint
                Case CPMENDDATE
                    Contract.pbcCalendar_Paint   'mBoxCalDate called within paint
                Case CPMIMPRESSIONS
                Case CPMIMPRESSIONCAP
                Case CPMCPMVALUE
                Case CPMCOST
                Case CPMCOMMENT
                   '10843 remove dead code
                   ' If tgSpf.sCLnStdQt <> "Y" Then
                        imLbcArrowSetting = True
                        gMatchLookAhead Contract.edcLNDropDown, Contract.lbcComment(1), imBSMode, imComboBoxIndex
                   ' End If
            End Select
        Case 3  'edcLNDropDown_KeyPress
            Select Case Contract.grdCPMBuy.Col
                Case CPMVEHICLE
                    ilKey = ilValue
                    If Not gCheckKeyAscii(ilKey) Then
                        ilValue = 0
                        Exit Sub
                    End If
                Case CPMDAYPART
                     ilKey = ilValue
                    If Not gCheckKeyAscii(ilKey) Then
                        ilValue = 0
                        Exit Sub
                    End If
                'Case CPMCOPYTYPE
                '10879 note max is 3 digits
                Case CPMLENGTH
                     If (ilValue <> KEYBACKSPACE) And ((ilValue < KEY0) Or (ilValue > KEY9)) Then
                        Beep
                        ilValue = 0
                        Exit Sub
                    End If
                    slStr = Contract.edcLNDropDown.Text
                    slStr = Left$(slStr, Contract.edcLNDropDown.SelStart) & Chr$(ilValue) & right$(slStr, Len(slStr) - Contract.edcLNDropDown.SelStart - Contract.edcLNDropDown.SelLength)
                    slMaxValue = "999"
                    If gCompNumberStr(slStr, slMaxValue) > 0 Then
                        Beep
                        ilValue = 0
                        Exit Sub
                    End If
                Case CPMPRICETYPE
                Case CPMPOSITION
                Case CPMPRIORITY
                Case CPMPODTARGET
                Case CPMSTARTDATE
                Case CPMENDDATE
                Case CPMIMPRESSIONS
                    If (ilValue <> KEYBACKSPACE) And ((ilValue < KEY0) Or (ilValue > KEY9)) Then
                        Beep
                        ilValue = 0
                        Exit Sub
                    End If
                    slStr = Contract.edcLNDropDown.Text
                    slStr = Left$(slStr, Contract.edcLNDropDown.SelStart) & Chr$(ilValue) & right$(slStr, Len(slStr) - Contract.edcLNDropDown.SelStart - Contract.edcLNDropDown.SelLength)
                    slMaxValue = "999999999"
                    If gCompNumberStr(slStr, slMaxValue) > 0 Then
                        Beep
                        ilValue = 0
                        Exit Sub
                    End If
                Case CPMIMPRESSIONCAP
                    If (ilValue <> KEYBACKSPACE) And ((ilValue < KEY0) Or (ilValue > KEY9)) Then
                        Beep
                        ilValue = 0
                        Exit Sub
                    End If
                    slStr = Contract.edcLNDropDown.Text
                    slStr = Left$(slStr, Contract.edcLNDropDown.SelStart) & Chr$(ilValue) & right$(slStr, Len(slStr) - Contract.edcLNDropDown.SelStart - Contract.edcLNDropDown.SelLength)
                    slMaxValue = "999999999"
                    If gCompNumberStr(slStr, slMaxValue) > 0 Then
                        Beep
                        ilValue = 0
                        Exit Sub
                    End If
                Case CPMCPMVALUE
                    ilPos = InStr(Contract.edcLNDropDown.SelText, ".")
                    If ilPos = 0 Then
                        ilPos = InStr(Contract.edcLNDropDown.Text, ".")    'Disallow multi-decimal points
                        If ilPos > 0 Then
                            If ilValue = KEYDECPOINT Then
                                Beep
                                ilValue = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
                    If (ilValue <> KEYBACKSPACE) And ((ilValue < KEY0) Or (ilValue > KEY9)) And (ilValue <> KEYDECPOINT) Then
                        Beep
                        ilValue = 0
                        Exit Sub
                    End If
                    slStr = Contract.edcLNDropDown.Text
                    slStr = Left$(slStr, Contract.edcLNDropDown.SelStart) & Chr$(ilValue) & right$(slStr, Len(slStr) - Contract.edcLNDropDown.SelStart - Contract.edcLNDropDown.SelLength)
                    If gCompNumberStr(slStr, "999.99") > 0 Then
                        Beep
                        ilValue = 0
                        Exit Sub
                    End If
                Case CPMCOST
                    ilPos = InStr(Contract.edcLNDropDown.SelText, ".")
                    If ilPos = 0 Then
                        ilPos = InStr(Contract.edcLNDropDown.Text, ".")    'Disallow multi-decimal points
                        If ilPos > 0 Then
                            If ilValue = KEYDECPOINT Then
                                Beep
                                ilValue = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
                    If (ilValue <> KEYBACKSPACE) And ((ilValue < KEY0) Or (ilValue > KEY9)) And (ilValue <> KEYDECPOINT) Then
                        Beep
                        ilValue = 0
                        Exit Sub
                    End If
                    slStr = Contract.edcLNDropDown.Text
                    slStr = Left$(slStr, Contract.edcLNDropDown.SelStart) & Chr$(ilValue) & right$(slStr, Len(slStr) - Contract.edcLNDropDown.SelStart - Contract.edcLNDropDown.SelLength)
                    If gCompNumberStr(slStr, "9999999.99") > 0 Then
                        Beep
                        ilValue = 0
                        Exit Sub
                    End If
                Case CPMCOMMENT
            End Select
        Case 4  'edcLNDropDown_KeyUp
            Select Case Contract.grdCPMBuy.Col
                Case CPMVEHICLE
                    If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) <> "P" Then
                        gProcessArrowKey igCPMShift, igCPMKeyCode, Contract.lbcLnVehicle(igTabMapIndex), imLbcArrowSetting
                    Else
                        gProcessArrowKey igCPMShift, igCPMKeyCode, Contract.lbcPkVehicle(igTabMapIndex), imLbcArrowSetting
                    End If
                Case CPMDAYPART
                    If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) <> "P" Then
                        gProcessArrowKey igCPMShift, igCPMKeyCode, Contract.lbcLnProgram(igTabMapIndex), imLbcArrowSetting
                    Else
                        gProcessArrowKey igCPMShift, igCPMKeyCode, Contract.lbcPkProgram(igTabMapIndex), imLbcArrowSetting
                    End If
                'Case CPMCOPYTYPE
                '    gProcessArrowKey igCPMShift, igCPMKeyCode, Contract.lbcPrg(2), imLbcArrowSetting
                '10879
                Case CPMLENGTH
                    gProcessArrowKey igCPMShift, igCPMKeyCode, Contract.lbcPrg(6), imLbcArrowSetting
                Case CPMPRICETYPE
                Case CPMPOSITION
                    gProcessArrowKey igCPMShift, igCPMKeyCode, Contract.lbcPrg(4), imLbcArrowSetting
                Case CPMPRIORITY
                    gProcessArrowKey igCPMShift, igCPMKeyCode, Contract.lbcPrg(5), imLbcArrowSetting
                Case CPMPODTARGET
                    gProcessArrowKey igCPMShift, igCPMKeyCode, Contract.lbcPrg(3), imLbcArrowSetting
                Case CPMSTARTDATE
                Case CPMENDDATE
                Case CPMIMPRESSIONS
                Case CPMIMPRESSIONCAP
                Case CPMCPMVALUE
                Case CPMCOST
                Case CPMCOMMENT
            End Select
        Case 5  'Timer
            Select Case Contract.grdCPMBuy.Col
                Case CPMVEHICLE
                Case CPMDAYPART
                'Case CPMCOPYTYPE
                '    imLbcArrowSetting = False
                '    gProcessLbcClick Contract.lbcPrg(2), Contract.edcLNDropDown, imChgMode, imLbcArrowSetting
                Case CPMPRICETYPE
                Case CPMPOSITION
                Case CPMPRIORITY
                Case CPMPODTARGET
                    imLbcArrowSetting = False
                    gProcessLbcClick Contract.lbcPrg(3), Contract.edcLNDropDown, imChgMode, imLbcArrowSetting
                Case CPMSTARTDATE
                Case CPMENDDATE
                Case CPMIMPRESSIONS
                Case CPMIMPRESSIONCAP
                Case CPMCPMVALUE
                Case CPMCOST
                Case CPMCOMMENT
                    imLbcArrowSetting = False
                    If Contract.cmcLnDropDown.Visible Then
                        If Not Contract.edcLNDropDown.Visible Then
                            Contract.edcLNDropDown.Visible = True
                        End If
                        gProcessLbcClick Contract.lbcComment(1), Contract.edcLNDropDown, imChgMode, imLbcArrowSetting
                    End If
           End Select
    End Select
End Sub
Private Function mCopyTypeBranch() As Integer
'
'   ilRet = mCopyTypeBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(Contract.edcLNDropDown, Contract.lbcPrg(2), imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (Contract.edcLNDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mCopyTypeBranch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(INVOICESORTLIST)) Then
    '    imDoubleClickName = False
    '    mCopyTypeBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "4"
    igMNmCallSource = CALLSOURCEVEHICLE
    If Contract.edcLNDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    If igTestSystem Then
        slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName & "\1"
    Else
        slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName & "\1"
    End If
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mCopyTypeBranch = True
    imUpdateAllowed = ilUpdateAllowed
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        Contract.lbcPrg(2).Clear
        Contract.lbcPrg(2).Clear
        'smCopyTypeCodeTag = ""
        'mPopCPMCopyType
        If imTerminate Then
            mCopyTypeBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, Contract.lbcPrg(2)
        sgMNmName = ""
        If gLastFound(Contract.lbcPrg(2)) > 0 Then
            imChgMode = True
            Contract.lbcPrg(2).ListIndex = gLastFound(Contract.lbcPrg(2))
            Contract.edcLNDropDown.Text = Contract.lbcPrg(2).List(Contract.lbcPrg(2).ListIndex)
            imChgMode = False
            mCopyTypeBranch = False
            'mSetChg imBoxNo
        Else
            imChgMode = True
            Contract.lbcPrg(2).ListIndex = 1
            Contract.edcLNDropDown.Text = Contract.lbcPrg(2).List(1)
            imChgMode = False
            'mSetChg imBoxNo
            Contract.edcLNDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mCPMEnableBox
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mCPMEnableBox
        Exit Function
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyType                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Vehicle Group Code    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Public Sub mPopCPMCopyType()
'
'   mCopyType
'   Where:
'
    'ReDim ilFilter(0) As Integer
    'ReDim slFilter(0) As String
    'ReDim ilOffset(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = Contract.lbcPrg(2).ListIndex
    If ilIndex > 0 Then
        slName = Contract.lbcPrg(2).List(ilIndex)
    End If
    ilRet = gPopMnfPlusFieldsBox(Contract, Contract.lbcPrg(2), tmCopyTypeCode(), smCopyTypeCodeTag, "4")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mCopyTypeErr
        gCPErrorMsg ilRet, "mCopyType (gPopMnfPlusFieldsBox)", Contract
        On Error GoTo 0
        Contract.lbcPrg(2).AddItem "[None]", 0  'Force as first item on list
        Contract.lbcPrg(2).AddItem "[New]", 0  'Force as first item on list
        imCopyTypeChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, Contract.lbcPrg(2)
            If gLastFound(Contract.lbcPrg(2)) >= 1 Then
                Contract.lbcPrg(2).ListIndex = gLastFound(Contract.lbcPrg(2))
            Else
                Contract.lbcPrg(2).ListIndex = -1
            End If
        Else
            Contract.lbcPrg(2).ListIndex = ilIndex
        End If
        imCopyTypeChgMode = False
    End If
    Exit Sub
mCopyTypeErr:
    On Error GoTo 0
    imTerminate = True
End Sub

Private Function mPodTargetBranch() As Integer
'
'   ilRet = mPodTargetBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    
    ilRet = gOptionalLookAhead(Contract.edcLNDropDown, Contract.lbcPrg(3), imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (Contract.edcLNDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mPodTargetBranch = False
        Exit Function
    End If
    
    If Contract.edcLNDropDown.Text = "" Then
        imDoubleClickName = False
        mPodTargetBranch = False
        Exit Function
    End If
    igPodTargetCallSource = CALLSOURCECONTRACT
    If Contract.edcLNDropDown.Text = "[New]" Then
        sgPodTargetName = ""
    Else
        sgPodTargetName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    If igTestSystem Then
        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igPodTargetCallSource)) & "\" & sgPodTargetName
    Else
        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igPodTargetCallSource)) & "\" & sgPodTargetName
    End If
    If Contract.lbcAdvt.ListIndex >= 1 Then
        slNameCode = tmCntAdvertiser(Contract.lbcAdvt.ListIndex - 1).sKey
        ilParse = Trim$(gParseItem(slNameCode, 1, "\", slName))
        ilParse = Trim$(gParseItem(slNameCode, 2, "\", slCode))
        slStr = slStr & "\" & slName & "\" & slCode
    Else
        slStr = slStr & "\" & "" & "\" & "0"
    End If
    sgCommandStr = slStr
    PodTarget.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgPodTargetName)
    igPodTargetCallSource = Val(sgPodTargetName)
    ilParse = gParseItem(slStr, 2, "\", sgPodTargetName)
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mPodTargetBranch = True
    imUpdateAllowed = ilUpdateAllowed
    gShowBranner imUpdateAllowed
    If igPodTargetCallSource = CALLDONE Then  'Done
        igPodTargetCallSource = CALLNONE
'        gSetMenuState True
        Contract.lbcPrg(3).Clear
        smPodTargetCodeTag = ""
        mPopPodTarget
        If imTerminate Then
            mPodTargetBranch = False
            Exit Function
        End If
        gFindMatch sgPodTargetName, 1, Contract.lbcPrg(3)
        sgPodTargetName = ""
        If gLastFound(Contract.lbcPrg(3)) > 0 Then
            imChgMode = True
            Contract.lbcPrg(3).ListIndex = gLastFound(Contract.lbcPrg(3))
            Contract.edcLNDropDown.Text = Contract.lbcPrg(3).List(Contract.lbcPrg(3).ListIndex)
            imChgMode = False
            mPodTargetBranch = False
            'mSetChg imBoxNo
        Else
            imChgMode = True
            Contract.lbcPrg(3).ListIndex = 1
            Contract.edcLNDropDown.Text = Contract.lbcPrg(3).List(1)
            imChgMode = False
            'mSetChg imBoxNo
            Contract.edcLNDropDown.SetFocus
            Exit Function
        End If
    End If
    If igPodTargetCallSource = CALLCANCELLED Then  'Cancelled
        igPodTargetCallSource = CALLNONE
        sgPodTargetName = ""
        mCPMEnableBox
        Exit Function
    End If
    If igPodTargetCallSource = CALLTERMINATED Then
        igPodTargetCallSource = CALLNONE
        sgPodTargetName = ""
        mCPMEnableBox
        Exit Function
    End If
    Exit Function
End Function


Public Sub mPopPodTarget()
    Dim ilRet As Integer
    Dim ilRaf As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilUpper As Integer
    Dim ilParse As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slSQLQuery As String
    Dim raf_rst As ADODB.Recordset

    ilIndex = Contract.lbcPrg(3).ListIndex
    If ilIndex > 0 Then
        slName = Contract.lbcPrg(3).List(ilIndex)
    End If
    ReDim tmPodTarget(0 To 0) As POPLCODENAME
    If Contract.lbcAdvt.ListIndex >= 1 Then
        slNameCode = tmCntAdvertiser(Contract.lbcAdvt.ListIndex - 1).sKey
        ilParse = Trim$(gParseItem(slNameCode, 2, "\", slCode))
    Else
        slCode = "0"
    End If
    
    slSQLQuery = "Select * from raf_Region_Area Where rafAdfCode = " & slCode & " And rafType = 'P' And rafState <> 'D' Order by rafName"
    Set raf_rst = gSQLSelectCall(slSQLQuery)
    Do While Not raf_rst.EOF
        ilUpper = UBound(tmPodTarget)
        tmPodTarget(ilUpper).lCode = raf_rst!rafCode
        tmPodTarget(ilUpper).sChar = raf_rst!rafName
        ReDim Preserve tmPodTarget(0 To ilUpper + 1) As POPLCODENAME
        raf_rst.MoveNext
    Loop
    Contract.lbcPrg(3).Clear
    For ilRaf = 0 To UBound(tmPodTarget) - 1 Step 1
        Contract.lbcPrg(3).AddItem Trim$(tmPodTarget(ilRaf).sChar)
        Contract.lbcPrg(3).ItemData(Contract.lbcPrg(3).NewIndex) = tmPodTarget(ilRaf).lCode
    Next ilRaf
    Contract.lbcPrg(3).AddItem "[None]", 0  'Force as first item on list
    Contract.lbcPrg(3).AddItem "[New]", 0  'Force as first item on list
    imCopyTypeChgMode = True
    If ilIndex > 0 Then
        gFindMatch slName, 1, Contract.lbcPrg(3)
        If gLastFound(Contract.lbcPrg(3)) >= 1 Then
            Contract.lbcPrg(3).ListIndex = gLastFound(Contract.lbcPrg(3))
        Else
            Contract.lbcPrg(3).ListIndex = -1
        End If
    Else
        Contract.lbcPrg(3).ListIndex = ilIndex
    End If
    Contract.lbcPrg(3).height = gListBoxHeight(Contract.lbcPrg(3).ListCount, 5)
    imPodTargetChgMode = False
End Sub

Public Sub gAddPcfRec()
    Dim llRow As Long
    Dim slStr As String
    Dim ilRet As Integer
    Dim llRet As Long
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slEntryDate As String
    Dim slEntryTime As String
    Dim slSQLQuery As String
    Dim tlPcf As PCF
    
    Contract.mVehPop False, True
    tgChfCntr.sAdServerDefined = "N"  'Set in mSetHdVeh
    If lmCntrChfCode = tgChfCntr.lCode Then
        slStr = ""
        For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
                If (Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) <> "0") Then
                    If slStr = "" Then
                        slStr = Contract.grdCPMBuy.TextMatrix(llRow, CPMID)
                    Else
                        slStr = slStr & ", " & Contract.grdCPMBuy.TextMatrix(llRow, CPMID)
                    End If
                End If
            End If
        Next llRow
        '10531
'        If slStr <> "" Then
'            slSQLQuery = "Delete From pcf_Pod_CPM_Cntr"
'            slSQLQuery = slSQLQuery & " Where "
'            slSQLQuery = slSQLQuery & " pcfPodCPMID Not In (" & slStr & ")"
'            slSQLQuery = slSQLQuery & " And pcfChfCode = " & tgChfCntr.lCode
'            llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
'        End If
        slSQLQuery = "Delete From pcf_Pod_CPM_Cntr Where pcfChfCode = " & tgChfCntr.lCode
        If slStr <> "" Then
            slSQLQuery = slSQLQuery & " AND pcfPodCPMID Not In (" & slStr & ")"
        End If
        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)

    End If
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
            tlPcf.lChfCode = tgChfCntr.lCode
            tlPcf.iCntRevNo = tgChfCntr.iCntRevNo
            tlPcf.iPropVer = tgChfCntr.iPropVer
            tlPcf.iPodCPMID = Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMID))
            If Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "P" Then
                tlPcf.sType = "P"
                tlPcf.iPkCPMID = 0
            ElseIf Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H" Then
                tlPcf.sType = "H"
                tlPcf.iPkCPMID = Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID)
            Else
                tlPcf.sType = "S"
                tlPcf.iPkCPMID = 0
            End If
            'Get Vehicle
            'slStr = Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE)
            'If Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) <> "P" Then
            '    gFindMatch slStr, 0, Contract.lbcLnVehicle(1)
            '    If gLastFound(Contract.lbcLnVehicle(1)) >= 0 Then
            '        slNameCode = tmVehicleCode2(gLastFound(Contract.lbcLnVehicle(1))).sKey    'lbcVehicle.List(gLastFound(lbcLnVehicle(igTabMapIndex)))
            '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
            '        tlPcf.iVefCode = CInt(slCode)
            '    End If
            'Else
            '    gFindMatch slStr, 0, Contract.lbcPkVehicle(1)
            '    If gLastFound(Contract.lbcPkVehicle(1)) >= 0 Then
            '        slNameCode = tgCPMPkVeh(gLastFound(Contract.lbcPkVehicle(1))).sKey    'lbcVehicle.List(gLastFound(lbcLnVehicle(igTabMapIndex)))
            '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
            '        tlPcf.iVefCode = CInt(slCode)
            '    End If
            'End If
            tlPcf.iVefCode = CInt(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE))
            'Get daypart
            tlPcf.iRdfCode = 0
            slStr = Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART)
            For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                If slStr = Trim$(tgMRdf(ilRdf).sName) Then
                    tlPcf.iRdfCode = tgMRdf(ilRdf).iCode
                    Exit For
                End If
            Next ilRdf
            'Copy Type
            tlPcf.iCopyTypeMnfCode = 0
            'slStr = Contract.grdCPMBuy.TextMatrix(llRow, CPMCOPYTYPE)
            'gFindMatch slStr, 1, Contract.lbcPrg(2)
            'If gLastFound(Contract.lbcPrg(2)) > 1 Then
            '    slNameCode = tmCopyTypeCode(gLastFound(Contract.lbcPrg(2)) - 2).sKey 'lbcVehicle.List(gLastFound(lbcLnVehicle(igTabMapIndex)))
            '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
            '    tlPcf.iCopyTypeMnfCode = CInt(slCode)
            'End If
            'Price Type
            If Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMPRICETYPE)) = "" Then
                tlPcf.sPriceType = "C"
            Else
                tlPcf.sPriceType = Left$(Contract.grdCPMBuy.TextMatrix(llRow, CPMPRICETYPE), 1)
            End If
            'Position
            If Contract.grdCPMBuy.TextMatrix(llRow, CPMPOSITION) = "" Or Contract.grdCPMBuy.TextMatrix(llRow, CPMPOSITION) = "[No Position]" Then
                tlPcf.iPositionMnfCode = 0
            Else
                mPopCPMPosition llRow
                tlPcf.iPositionMnfCode = 0
                slStr = Contract.grdCPMBuy.TextMatrix(llRow, CPMPOSITION)
                gFindMatch slStr, 1, Contract.lbcPrg(4)
                If gLastFound(Contract.lbcPrg(4)) >= 1 Then
                    tlPcf.iPositionMnfCode = Contract.lbcPrg(4).ItemData(gLastFound(Contract.lbcPrg(4)))
                End If
            End If
            'Priority
            If Contract.grdCPMBuy.TextMatrix(llRow, CPMPRIORITY) = "" Or Contract.grdCPMBuy.TextMatrix(llRow, CPMPRIORITY) = "[None]" Or Contract.grdCPMBuy.TextMatrix(llRow, CPMPRIORITY) = "[No Priority" Then
                tlPcf.iPriority = 0
            Else
                tlPcf.iPriority = Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMPRIORITY))
            End If
            'Pod-Target
            tlPcf.lRafCode = 0
            slStr = Contract.grdCPMBuy.TextMatrix(llRow, CPMPODTARGET)
            gFindMatch slStr, 1, Contract.lbcPrg(3)
            If gLastFound(Contract.lbcPrg(3)) > 1 Then
                'slNameCode = tmPodTarget(gLastFound(Contract.lbcPrg(3)) - 2).sKey  'lbcVehicle.List(gLastFound(lbcLnVehicle(igTabMapIndex)))
                'ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tlPcf.lRafCode = Contract.lbcPrg(3).ItemData(gLastFound(Contract.lbcPrg(3)))
            End If
            'Delete
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) = "CBS") Or (Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) = "Deleted") Then
                tlPcf.sDelete = "Y"
                slStartDate = "12/31/2069"
                slEndDate = "12/30/2069"
                'Hide CBS
                If (slStartDate = "Deleted") Then
                    tlPcf.sHideCBS = "Y"
                Else
                    tlPcf.sHideCBS = "N"
                End If
            Else
                tlPcf.sDelete = "N"
                'Start date
                slStartDate = Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)
                'End Date
                slEndDate = Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)
                'Hide CBS
                tlPcf.sHideCBS = "N"
            End If
            'Total Impressions
            tlPcf.lImpressionGoal = Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS))
            'Impression Cap
            tlPcf.lImpressionCap = Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONCAP))
            'CPM
            tlPcf.lPodCPM = gStrDecToLong(Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE), 2)
            'Total Cost
            tlPcf.lTotalCost = gStrDecToLong(Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST), 2)
            'Comment
            tlPcf.lCxfCode = mSaveCPMComment(llRow)
            'User
            tlPcf.iUrfCode = tgUrf(0).iCode
            'Entry info
            slEntryDate = Format$(Now, "m/d/yy")
            slEntryTime = Format$(Now, "h:mm:ssAM/PM")
            '10879
            tlPcf.iLen = Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMLENGTH))
            '10967
            tlPcf.iAdServDlvyStatus = Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMDELIVERYSTATUS))
            tlPcf.iAvfCode = Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMAVFCODE))
            tlPcf.sUnused = ""
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = "") Or (Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = "0") Or (lmCntrChfCode <> tgChfCntr.lCode) Then
                'Insert
                slSQLQuery = "Insert Into pcf_Pod_CPM_Cntr ( "
                slSQLQuery = slSQLQuery & "pcfCode, "
                slSQLQuery = slSQLQuery & "pcfChfCode, "
                slSQLQuery = slSQLQuery & "pcfPodCPMID, "
                slSQLQuery = slSQLQuery & "pcfCntRevNo, "
                slSQLQuery = slSQLQuery & "pcfPropVer, "
                slSQLQuery = slSQLQuery & "pcfVefCode, "
                slSQLQuery = slSQLQuery & "pcfRdfCode, "
                slSQLQuery = slSQLQuery & "pcfType, "
                slSQLQuery = slSQLQuery & "pcfPkCPMID, "
                slSQLQuery = slSQLQuery & "pcfCopyTypeMnfCode, "
                slSQLQuery = slSQLQuery & "pcfPriceType, "
                slSQLQuery = slSQLQuery & "pcfPositionMnfCode, "
                slSQLQuery = slSQLQuery & "pcfPriority, "
                slSQLQuery = slSQLQuery & "pcfRafCode, "
                slSQLQuery = slSQLQuery & "pcfStartDate, "
                slSQLQuery = slSQLQuery & "pcfEndDate, "
                slSQLQuery = slSQLQuery & "pcfImpressionGoal, "
                slSQLQuery = slSQLQuery & "pcfImpressionCap, "
                slSQLQuery = slSQLQuery & "pcfPodCPM, "
                slSQLQuery = slSQLQuery & "pcfTotalCost, "
                slSQLQuery = slSQLQuery & "pcfCxfCode, "
                slSQLQuery = slSQLQuery & "pcfUrfCode, "
                slSQLQuery = slSQLQuery & "pcfDelete, "
                slSQLQuery = slSQLQuery & "pcfEntryDate, "
                slSQLQuery = slSQLQuery & "pcfEntryTime, "
                slSQLQuery = slSQLQuery & "pcfHideCBS, "
                '10879
                slSQLQuery = slSQLQuery & "pcfLen, "
                '10967
                slSQLQuery = slSQLQuery & "pcfavfCode,pcfAdServDlvyStatus, "
                slSQLQuery = slSQLQuery & "pcfUnused "
                slSQLQuery = slSQLQuery & ") "
                slSQLQuery = slSQLQuery & "Values ( "
                slSQLQuery = slSQLQuery & "Replace" & ", "
                slSQLQuery = slSQLQuery & tlPcf.lChfCode & ", "
                slSQLQuery = slSQLQuery & tlPcf.iPodCPMID & ", "
                slSQLQuery = slSQLQuery & tlPcf.iCntRevNo & ", "
                slSQLQuery = slSQLQuery & tlPcf.iPropVer & ", "
                slSQLQuery = slSQLQuery & tlPcf.iVefCode & ", "
                slSQLQuery = slSQLQuery & tlPcf.iRdfCode & ", "
                slSQLQuery = slSQLQuery & "'" & gFixQuote(tlPcf.sType) & "', "
                slSQLQuery = slSQLQuery & tlPcf.iPkCPMID & ", "
                slSQLQuery = slSQLQuery & 0 & ","   'tlPcf.iCopyTypeMnfCode & ", "
                slSQLQuery = slSQLQuery & "'" & tlPcf.sPriceType & "', "
                slSQLQuery = slSQLQuery & tlPcf.iPositionMnfCode & ","
                slSQLQuery = slSQLQuery & tlPcf.iPriority & ","
                slSQLQuery = slSQLQuery & tlPcf.lRafCode & ", "
                slSQLQuery = slSQLQuery & "'" & Format$(slStartDate, sgSQLDateForm) & "', "
                slSQLQuery = slSQLQuery & "'" & Format$(slEndDate, sgSQLDateForm) & "', "
                slSQLQuery = slSQLQuery & tlPcf.lImpressionGoal & ", "
                slSQLQuery = slSQLQuery & tlPcf.lImpressionCap & ", "
                slSQLQuery = slSQLQuery & tlPcf.lPodCPM & ", "
                slSQLQuery = slSQLQuery & tlPcf.lTotalCost & ", "
                slSQLQuery = slSQLQuery & tlPcf.lCxfCode & ", "
                slSQLQuery = slSQLQuery & tlPcf.iUrfCode & ", "
                slSQLQuery = slSQLQuery & "'" & gFixQuote(tlPcf.sDelete) & "', "
                slSQLQuery = slSQLQuery & "'" & Format$(slEntryDate, sgSQLDateForm) & "', "
                slSQLQuery = slSQLQuery & "'" & Format$(slEntryTime, sgSQLTimeForm) & "', "
                slSQLQuery = slSQLQuery & "'" & gFixQuote(tlPcf.sHideCBS) & "', "
                '10879
                slSQLQuery = slSQLQuery & tlPcf.iLen & ", "
'                '10967
                slSQLQuery = slSQLQuery & tlPcf.iAvfCode & ", "
                slSQLQuery = slSQLQuery & tlPcf.iAdServDlvyStatus & ","
                slSQLQuery = slSQLQuery & "'" & gFixQuote(tlPcf.sUnused) & "' "
                slSQLQuery = slSQLQuery & ") "
                llRet = gInsertAndReturnCode(slSQLQuery, "pcf_Pod_CPM_Cntr", "pcfCode", "Replace")
                Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = llRet
            Else
                'Update
                slSQLQuery = "Update pcf_Pod_CPM_Cntr Set "
                'slSQLQuery = slSQLQuery & "pcfCode = " & tlPcf.lCode & ", "
                slSQLQuery = slSQLQuery & "pcfChfCode = " & tlPcf.lChfCode & ", "
                slSQLQuery = slSQLQuery & "pcfPodCPMID = " & tlPcf.iPodCPMID & ", "
                slSQLQuery = slSQLQuery & "pcfCntRevNo = " & tlPcf.iCntRevNo & ", "
                slSQLQuery = slSQLQuery & "pcfPropVer = " & tlPcf.iPropVer & ", "
                slSQLQuery = slSQLQuery & "pcfVefCode = " & tlPcf.iVefCode & ", "
                slSQLQuery = slSQLQuery & "pcfRdfCode = " & tlPcf.iRdfCode & ", "
                slSQLQuery = slSQLQuery & "pcfType = '" & gFixQuote(tlPcf.sType) & "', "
                slSQLQuery = slSQLQuery & "pcfPkCPMID = " & tlPcf.iPkCPMID & ", "
                slSQLQuery = slSQLQuery & "pcfCopyTypeMnfCode = " & 0 & ", "    'tlPcf.iCopyTypeMnfCode & ", "
                slSQLQuery = slSQLQuery & "pcfPriceType = '" & tlPcf.sPriceType & "', "
                slSQLQuery = slSQLQuery & "pcfPositionMnfCode = " & tlPcf.iPositionMnfCode & ", "
                slSQLQuery = slSQLQuery & "pcfPriority = " & tlPcf.iPriority & ", "
                slSQLQuery = slSQLQuery & "pcfRafCode = " & tlPcf.lRafCode & ", "
                slSQLQuery = slSQLQuery & "pcfStartDate = '" & Format$(slStartDate, sgSQLDateForm) & "', "
                slSQLQuery = slSQLQuery & "pcfEndDate = '" & Format$(slEndDate, sgSQLDateForm) & "', "
                slSQLQuery = slSQLQuery & "pcfImpressionGoal = " & tlPcf.lImpressionGoal & ", "
                slSQLQuery = slSQLQuery & "pcfImpressionCap = " & tlPcf.lImpressionCap & ", "
                slSQLQuery = slSQLQuery & "pcfPodCPM = " & tlPcf.lPodCPM & ", "
                slSQLQuery = slSQLQuery & "pcfTotalCost = " & tlPcf.lTotalCost & ", "
                slSQLQuery = slSQLQuery & "pcfCxfCode = " & tlPcf.lCxfCode & ", "
                slSQLQuery = slSQLQuery & "pcfUrfCode = " & tlPcf.iUrfCode & ", "
                slSQLQuery = slSQLQuery & "pcfDelete = '" & gFixQuote(tlPcf.sDelete) & "', "
                slSQLQuery = slSQLQuery & "pcfEntryDate = '" & Format$(slEntryDate, sgSQLDateForm) & "', "
                slSQLQuery = slSQLQuery & "pcfEntryTime = '" & Format$(slEntryTime, sgSQLTimeForm) & "', "
                slSQLQuery = slSQLQuery & "pcfHideCBS = '" & gFixQuote(tlPcf.sHideCBS) & "', "
                '10879
                slSQLQuery = slSQLQuery & "pcfLen = " & tlPcf.iLen & ", "
                '10967
                slSQLQuery = slSQLQuery & "pcfAvfCode = " & tlPcf.iAvfCode & ", pcfAdServDlvyStatus = " & tlPcf.iAdServDlvyStatus & ", "
                slSQLQuery = slSQLQuery & "pcfUnused = '" & gFixQuote(tlPcf.sUnused) & "' "
                slSQLQuery = slSQLQuery & "Where pcfCode = " & Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE)
                llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            End If
            tgChfCntr.sAdServerDefined = "Y"  'Set in mSetHdVeh
        End If
    Next llRow
End Sub

Public Function gReadPcfAndLoadGrid() As Boolean
    Dim ilCol As Integer
    Dim llRow As Long
    Dim ilRdf As Integer
    Dim ilVef As Integer
    Dim slStr As String
    Dim ilNotShownCount As Integer
    Dim ilAdd As Integer
    Dim slSQLQuery As String
    Dim pcf_rst As ADODB.Recordset
    Dim mnf_rst As ADODB.Recordset
    Dim raf_rst As ADODB.Recordset
    Dim cxf_rst As ADODB.Recordset
    Dim ibf_rst As ADODB.Recordset
    
    gReadPcfAndLoadGrid = False
    Contract.grdCPMBuy.Redraw = False
    gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbHourglass
    mPopPodTarget
    mClearCPMBuyGrid
    ilNotShownCount = 0
    Contract.grdCPMBuy.Row = Contract.grdCPMBuy.FixedRows - 1
    slSQLQuery = "Select * From pcf_Pod_CPM_Cntr Where pcfChfCode = " & tgChfCntr.lCode
    Set pcf_rst = gSQLSelectCall(slSQLQuery)
    Do While Not pcf_rst.EOF
        If Contract.grdCPMBuy.Row + 1 >= Contract.grdCPMBuy.rows Then
            Contract.grdCPMBuy.AddItem ""
            Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row + 1) = fgBoxGridH
            Contract.grdCPMBuy.Row = Contract.grdCPMBuy.Row + 1
            For ilCol = 0 To Contract.grdCPMBuy.cols - 1 Step 1
                Contract.grdCPMBuy.ColAlignment(ilCol) = flexAlignLeftCenter
            Next ilCol
            Contract.grdCPMBuy.Col = CPMVEHICLE
            Contract.grdCPMBuy.CellFontItalic = True
            Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row, CPMPCFCODE) = "0"
            Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row, CPMSORT) = ""
            Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row, CPMCXFCODE) = "0"
            Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row, CPMCOMMSHOWON) = "NNDND"   'Proposal;Order;Spots/Copy;Invoice;Insertion
            Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.Row, CPMSHOWROW) = "True"
        Else
            Contract.grdCPMBuy.Row = Contract.grdCPMBuy.Row + 1
        End If
        Contract.grdCPMBuy.Col = CPMID
        Contract.grdCPMBuy.CellBackColor = LIGHTYELLOW
        llRow = Contract.grdCPMBuy.Row
        Contract.grdCPMBuy.Col = CPMEXPAND
        Contract.grdCPMBuy.CellFontName = "Monotype Sorts"
        Contract.grdCPMBuy.CellFontSize = 7
        Contract.grdCPMBuy.CellForeColor = BLUE
        Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = pcf_rst!pcfCode
        'Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = lChfCode = rst!pcfChfCode
        Contract.grdCPMBuy.TextMatrix(llRow, CPMID) = pcf_rst!pcfPodCPMID
        'Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = iCntRevNo = rst!pcfCntRevNo
        'Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = iPropVer = rst!pcfPropVer
        ilVef = gBinarySearchVef(pcf_rst!pcfVefCode)
        If ilVef <> -1 Then
            If pcf_rst!pcfType = "H" Then
                Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) = smHiddenSpace & Trim$(tgMVef(ilVef).sName)
            Else
                Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) = Trim$(tgMVef(ilVef).sName)
            End If
        End If
        Contract.grdCPMBuy.Col = CPMVEHICLE
        Contract.grdCPMBuy.CellFontItalic = True
        ilRdf = gBinarySearchRdf(pcf_rst!pcfRdfCode)
        If ilRdf <> -1 Then
            Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) = Trim$(tgMRdf(ilRdf).sName)
        End If
        Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = pcf_rst!pcfType
        'If pcf_rst!pcfType <> "P" Then
        '    Contract.grdCPMBuy.TextMatrix(llRow, CPMPKCPMID) = ""
        'Else
        '    Contract.grdCPMBuy.TextMatrix(llRow, CPMPKCPMID) = pcf_rst!pcfPkCPMID
        'End If
        ''Copy type
        'slSQLQuery = "Select mnfName From mnf_Multi_Names Where mnfCode = " & pcf_rst!pcfCopyTypeMnfCode
        'Set mnf_rst = gSQLSelectCall(slSQLQuery)
        'If Not mnf_rst.EOF Then
        '    Contract.grdCPMBuy.TextMatrix(llRow, CPMCOPYTYPE) = Trim$(mnf_rst!mnfname)
        'End If
        'Price Type
        If pcf_rst!pcfPriceType = "F" Then
            Contract.grdCPMBuy.TextMatrix(llRow, CPMPRICETYPE) = "Flat Rate"
        Else
            Contract.grdCPMBuy.TextMatrix(llRow, CPMPRICETYPE) = "CPM"
        End If
        'Position
        If pcf_rst!pcfPositionMnfCode = 0 Then
            Contract.grdCPMBuy.TextMatrix(llRow, CPMPOSITION) = "[No Position]"
        Else
            slSQLQuery = "Select mnfName, mnfCode, mnfType from MNF_Multi_Names"
            slSQLQuery = slSQLQuery & " Where mnfCode = " & pcf_rst!pcfPositionMnfCode
            slSQLQuery = slSQLQuery & " And mnfType = '6'"
            Set mnf_rst = gSQLSelectCall(slSQLQuery)
            If Not mnf_rst.EOF Then
                Contract.grdCPMBuy.TextMatrix(llRow, CPMPOSITION) = mnf_rst!mnfName
            End If
        End If
        'Position
        If pcf_rst!pcfPriority = 0 Then
            Contract.grdCPMBuy.TextMatrix(llRow, CPMPRIORITY) = ""
        Else
            Contract.grdCPMBuy.TextMatrix(llRow, CPMPRIORITY) = pcf_rst!pcfPriority
        End If
            
        'Podcast-Target
        slSQLQuery = "Select rafName From raf_Region_Area Where rafCode = " & pcf_rst!pcfRafCode
        Set raf_rst = gSQLSelectCall(slSQLQuery)
        If Not raf_rst.EOF Then
            Contract.grdCPMBuy.TextMatrix(llRow, CPMPODTARGET) = Trim$(raf_rst!rafName)
        End If
        If pcf_rst!pcfDelete = "Y" Then
            If pcf_rst!pcfHideCBS <> "Y" Then
                Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) = "CBS"
            Else
                Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) = "Deleted"
            End If
            Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE) = ""
            '10738
'            Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) = Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)
           ' Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGENDDATE) = Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)
            Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS) = ""
            Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONCAP) = ""
            Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE) = ""
            Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST) = ""
        Else
            Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) = Format$(pcf_rst!pcfStartDate, sgShowDateForm)
            Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE) = Format$(pcf_rst!pcfEndDate, sgShowDateForm)
           ' Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) = Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)
           ' Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGENDDATE) = Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)
            Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS) = pcf_rst!pcfImpressionGoal
            Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONCAP) = pcf_rst!pcfImpressionCap
            If Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMPRICETYPE)) <> "Flat Rate" Then
                Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE) = gLongToStrDec(pcf_rst!pcfPodCPM, 2)
            Else
                Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE) = ""
            End If
            Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST) = gLongToStrDec(pcf_rst!pcfTotalCost, 2)
        End If
        '10738  moved out of if statement
        Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) = Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)
        Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGENDDATE) = Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)
        Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGIMPRESSIONS) = Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS)
        '10967
        Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGCAP) = Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONCAP)
        Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGPOSITION) = Contract.grdCPMBuy.TextMatrix(llRow, CPMPOSITION)
        Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGPRIORITY) = Contract.grdCPMBuy.TextMatrix(llRow, CPMPRIORITY)
        Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGVEHICLE) = Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE)
        Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGTOTAL) = Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST)
        Contract.grdCPMBuy.TextMatrix(llRow, CPMAVFCODE) = pcf_rst!pcfavfcode
        Contract.grdCPMBuy.TextMatrix(llRow, CPMDELIVERYSTATUS) = pcf_rst!pcfAdServDlvyStatus
        Contract.grdCPMBuy.TextMatrix(llRow, CPMCOMMENT) = ""
        Contract.grdCPMBuy.TextMatrix(llRow, CPMCXFCODE) = "0"
        Contract.grdCPMBuy.TextMatrix(llRow, CPMCOMMSHOWON) = "NNDND"   'Proposal;Order;Spots/Copy;Invoice;Insertion
        
        If pcf_rst!pcfCxfCode > 0 Then
            slSQLQuery = "Select * From CXF_Hdr_Ln_Comments Where cxfCode = " & pcf_rst!pcfCxfCode
            Set cxf_rst = gSQLSelectCall(slSQLQuery)
            If Not cxf_rst.EOF Then
                Contract.grdCPMBuy.TextMatrix(llRow, CPMCOMMENT) = gStripChr0(cxf_rst!cxfComment)
                Contract.grdCPMBuy.TextMatrix(llRow, CPMCXFCODE) = cxf_rst!cxfCode
                slStr = cxf_rst!cxfShProp
                slStr = slStr & cxf_rst!cxfShOrder
                slStr = slStr & "D"
                slStr = slStr & cxf_rst!cxfShInv
                slStr = slStr & "D"
                Contract.grdCPMBuy.TextMatrix(llRow, CPMCOMMSHOWON) = slStr   'Proposal;Order;Spots/Copy;Invoice;Insertion
            End If
        End If
        Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE) = pcf_rst!pcfVefCode
        Contract.grdCPMBuy.TextMatrix(llRow, CPMRDFCODE) = pcf_rst!pcfRdfCode
        Contract.grdCPMBuy.TextMatrix(llRow, CPMCHFCODE) = tgChfCntr.lCode
        Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = pcf_rst!pcfPKCPMID
        '10879
        Contract.grdCPMBuy.TextMatrix(llRow, CPMLENGTH) = pcf_rst!pcfLen
        If Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "P" Then
            Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "True"
            If Not bgHideHiddenLines Then
                Contract.grdCPMBuy.TextMatrix(llRow, CPMEXPAND) = ""
            Else
                If Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) Then
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMEXPAND) = "t"
                Else
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMEXPAND) = "s"
                End If
            End If
        ElseIf Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H" Then
            If Not bgHideHiddenLines Then
                Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "True"
            Else
                Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "False"
                ilNotShownCount = ilNotShownCount + 1
            End If
            If Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "True" Then
                Contract.grdCPMBuy.rowHeight(llRow) = fgBoxGridH
            Else
                Contract.grdCPMBuy.rowHeight(llRow) = 0
            End If
        Else
            Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "True"
        End If
        'Check if billed
        Contract.grdCPMBuy.TextMatrix(llRow, CPMBILLED) = ""
        slSQLQuery = "Select Count(1) as BilledCount From ibf_Impression_Bill Where ibfCntrNo = " & tgChfCntr.lCntrNo
        slSQLQuery = slSQLQuery & " And ibfPodCPMID = " & pcf_rst!pcfPodCPMID
        slSQLQuery = slSQLQuery & " And ibfBilled = 'Y'"
        Set ibf_rst = gSQLSelectCall(slSQLQuery)
        If Not ibf_rst.EOF Then
            If IsNull(ibf_rst!BilledCount) = False Then
                If ibf_rst!BilledCount > 0 Then
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMBILLED) = "Y"
                End If
            End If
        End If
        Contract.grdCPMBuy.TextMatrix(llRow, CPMSORT) = ""
        pcf_rst.MoveNext
    Loop
    If Contract.grdCPMBuy.Row + 1 = Contract.grdCPMBuy.rows Then
        ilNotShownCount = ilNotShownCount + 1
    End If
    'If Contract.grdCPMBuy.Row + 1 = Contract.grdCPMBuy.Rows Then
    If ilNotShownCount > 0 Then
        For ilAdd = 1 To ilNotShownCount Step 1
            llRow = Contract.grdCPMBuy.rows
            Do
                llRow = llRow - 1
            Loop While Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) = ""
            llRow = llRow + 1
            Contract.grdCPMBuy.AddItem "", llRow
            Contract.grdCPMBuy.Row = llRow
            For ilCol = 0 To Contract.grdCPMBuy.cols - 1 Step 1
                Contract.grdCPMBuy.ColAlignment(ilCol) = flexAlignLeftCenter
            Next ilCol
            Contract.grdCPMBuy.rowHeight(llRow) = fgBoxGridH
            Contract.grdCPMBuy.Col = CPMID
            Contract.grdCPMBuy.CellBackColor = LIGHTYELLOW
            Contract.grdCPMBuy.Col = CPMEXPAND
            Contract.grdCPMBuy.CellFontName = "Monotype Sorts"
            Contract.grdCPMBuy.CellFontSize = 7
            Contract.grdCPMBuy.CellForeColor = BLUE
            Contract.grdCPMBuy.Col = CPMVEHICLE
            Contract.grdCPMBuy.CellFontItalic = True
            Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = "0"
            Contract.grdCPMBuy.TextMatrix(llRow, CPMSORT) = ""
            Contract.grdCPMBuy.TextMatrix(llRow, CPMCXFCODE) = "0"
            Contract.grdCPMBuy.TextMatrix(llRow, CPMCOMMSHOWON) = "NNDND"   'Proposal;Order;Spots/Copy;Invoice;Insertion
            Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "True"
        Next ilAdd
        llRow = Contract.grdCPMBuy.rows
        Do
            llRow = llRow - 1
        Loop While Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) = ""
        Contract.grdCPMBuy.Row = llRow + 1
    End If
    On Error Resume Next
    ibf_rst.Close
    mnf_rst.Close
    raf_rst.Close
    pcf_rst.Close
    gCPMSort False  'This routine also calls mSetFGColorForRows
    mComputeTotals
    Contract.grdCPMBuy.Redraw = True
    gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbDefault
    gReadPcfAndLoadGrid = True
    Exit Function
End Function

Private Sub mChkForStdCPMPkg()
    Dim ilCode As Integer
    Dim ilVef As Integer
    Dim ilRdf As Integer
    Dim llRif As Long
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilCol As Integer
    Dim llRow As Long
    Dim ilIDNo As Integer
    Dim ilID As Integer
    Dim slStr As String
    Dim slHiddenImpressions As String
    Dim llFirstHiddenLnRow As Long
    Dim ilRcfCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilNotShownCount As Integer
    Dim ilAdd As Integer
    
    Contract.grdCPMBuy.Redraw = False
    gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbHourglass
    imPvfRecLen = Len(tmPvf)
    ilNotShownCount = 0
    ilCode = CInt(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMVEFCODE))

    ilVef = gBinarySearchVef(ilCode)
    If ilVef <> -1 Then
        If (tgMVef(ilVef).sType = "P") And (tgMVef(ilVef).lPvfCode > 0) Then
            If lmCPMEnableRow + 1 < Contract.grdCPMBuy.rows Then
                If (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow + 1, CPMLNTYPE) = "H") And (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID) = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow + 1, CPMPACKAGEID)) Then
                    Contract.grdCPMBuy.Redraw = True
                    gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbDefault
                    Exit Sub
                End If
            End If
            llFirstHiddenLnRow = -1
            slHiddenImpressions = "0"
            If Contract.lbcRateCard.ListIndex >= 0 Then
                slNameCode = tmCntRateCard(Contract.lbcRateCard.ListIndex).sKey   'Traffic!lbcRateCardCode.List(lbcRateCard.ListIndex)' - 1)    'lbcRateCard.ListCount - lbcRateCard.ListIndex - 1)
                ilRet = gParseItem(slNameCode, 3, "\", slCode)
                slCode = Trim$(slCode)
                ilRcfCode = CInt(slCode)
            Else
                ilRcfCode = -1
            End If
            Contract.grdCPMBuy.Col = CPMID
            Contract.grdCPMBuy.CellBackColor = LIGHTYELLOW
            Contract.grdCPMBuy.Col = CPMEXPAND
            Contract.grdCPMBuy.CellFontName = "Monotype Sorts"
            Contract.grdCPMBuy.CellFontSize = 7
            Contract.grdCPMBuy.CellForeColor = BLUE
            Contract.grdCPMBuy.Col = CPMVEHICLE
            Contract.grdCPMBuy.CellFontItalic = True
            Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) = "P"
            Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMSHOWROW) = "True"
            If Not bgHideHiddenLines Then
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMEXPAND) = ""
            Else
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMSHOWROW) Then
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMEXPAND) = "t"
                Else
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMEXPAND) = "s"
                End If
            End If
            tmPvfSrchKey.lCode = tgMVef(ilVef).lPvfCode
            ilRet = btrGetEqual(hmPvf, tmPvf, imPvfRecLen, tmPvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
            Do While ilRet = BTRV_ERR_NONE
                'Create Lines
                For ilLoop = LBound(tmPvf.iVefCode) To UBound(tmPvf.iVefCode) Step 1
                    If (tmPvf.iVefCode(ilLoop) > 0) And (tmPvf.iRdfCode(ilLoop) > 0) Then
                        If Contract.grdCPMBuy.Row + 1 >= Contract.grdCPMBuy.rows Then
                            Contract.grdCPMBuy.AddItem ""
                            Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row + 1) = fgBoxGridH
                            Contract.grdCPMBuy.Row = Contract.grdCPMBuy.Row + 1
                            For ilCol = 0 To Contract.grdCPMBuy.cols - 1 Step 1
                                Contract.grdCPMBuy.ColAlignment(ilCol) = flexAlignLeftCenter
                            Next ilCol
                        Else
                            Contract.grdCPMBuy.Row = Contract.grdCPMBuy.Row + 1
                        End If
                        llRow = Contract.grdCPMBuy.Row
                        If llFirstHiddenLnRow = -1 Then
                            llFirstHiddenLnRow = llRow
                        End If
                        Contract.grdCPMBuy.Col = CPMID
                        Contract.grdCPMBuy.CellBackColor = LIGHTYELLOW
                        Contract.grdCPMBuy.Col = CPMEXPAND
                        Contract.grdCPMBuy.CellFontName = "Monotype Sorts"
                        Contract.grdCPMBuy.CellFontSize = 7
                        Contract.grdCPMBuy.CellForeColor = BLUE
                        Contract.grdCPMBuy.Col = CPMVEHICLE
                        Contract.grdCPMBuy.CellFontItalic = True
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = "0"
                        If Contract.grdCPMBuy.TextMatrix(llRow, CPMID) = "" Then
                            ilIDNo = 0
                            For ilID = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                                If Val(Contract.grdCPMBuy.TextMatrix(ilID, CPMID)) > ilIDNo Then
                                    ilIDNo = Val(Contract.grdCPMBuy.TextMatrix(ilID, CPMID))
                                End If
                            Next ilID
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMID) = ilIDNo + 1
                        End If
                        ilVef = gBinarySearchVef(tmPvf.iVefCode(ilLoop))
                        If ilVef <> -1 Then
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) = smHiddenSpace & Trim$(tgMVef(ilVef).sName)
                        End If
                        ilRdf = gBinarySearchRdf(tmPvf.iRdfCode(ilLoop))
                        If ilRdf <> -1 Then
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) = Trim$(tgMRdf(ilRdf).sName)
                        End If
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H"
                        'Contract.grdCPMBuy.TextMatrix(llRow, CPMCOPYTYPE) = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOPYTYPE)
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMPRICETYPE) = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMPRICETYPE)
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMPOSITION) = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMPOSITION)
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMPRIORITY) = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMPRIORITY)
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMPODTARGET) = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMPODTARGET)
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMSTARTDATE)
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE) = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMENDDATE)
                        '10943
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMLENGTH) = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLENGTH)
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS) = gRoundStr(gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS), gIntToStrDec(tmPvf.iPctRate(ilLoop), 2)), "100.00"), 1, 0)
                        slHiddenImpressions = gAddStr(slHiddenImpressions, Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS))
                        slStr = "0.00"
                        For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                            If (tgMRif(llRif).iVefCode = tmPvf.iVefCode(ilLoop)) And (tgMRif(llRif).iRcfCode = ilRcfCode) And (tgMRif(llRif).iRdfCode = tmPvf.iRdfCode(ilLoop)) Then
                                slStr = gLongToStrDec(tgMRif(llRif).lPodCPM, 2)
                                Exit For
                            End If
                        Next llRif
                        
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE) = slStr
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONCAP) = ""
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS), Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE)), "1000.00")
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMCOMMENT) = ""
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE) = tmPvf.iVefCode(ilLoop)
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMRDFCODE) = tmPvf.iRdfCode(ilLoop)
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMCHFCODE) = tgChfCntr.lCode
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID)
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMCXFCODE) = "0"
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMCOMMSHOWON) = "NNDND"   'Proposal;Order;Spots/Copy;Invoice;Insertion
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMSORT) = ""
                        If bgHideHiddenLines Then
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "False"
                            ilNotShownCount = ilNotShownCount + 1
                        Else
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "True"
                        End If

                        If Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "True" Then
                            Contract.grdCPMBuy.rowHeight(llRow) = fgBoxGridH
                        Else
                            Contract.grdCPMBuy.rowHeight(llRow) = 0
                        End If
                    End If
                Next ilLoop
                If tmPvf.lLkPvfCode <= 0 Then
                    Exit Do
                End If
                tmPvfSrchKey.lCode = tmPvf.lLkPvfCode
                ilRet = btrGetEqual(hmPvf, tmPvf, imPvfRecLen, tmPvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
            Loop
            If Contract.grdCPMBuy.Row + 1 = Contract.grdCPMBuy.rows Then
                ilNotShownCount = ilNotShownCount + 1
            End If
            'If Contract.grdCPMBuy.Row + 1 = Contract.grdCPMBuy.Rows Then
            If ilNotShownCount > 0 Then
                For ilAdd = 1 To ilNotShownCount Step 1
                    llRow = Contract.grdCPMBuy.rows
                    Do
                        llRow = llRow - 1
                    Loop While Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) = ""
                    llRow = llRow + 1
                    Contract.grdCPMBuy.AddItem "", llRow
                    Contract.grdCPMBuy.Row = llRow  'Contract.grdCPMBuy.Row + 1
                    For ilCol = 0 To Contract.grdCPMBuy.cols - 1 Step 1
                        Contract.grdCPMBuy.ColAlignment(ilCol) = flexAlignLeftCenter
                    Next ilCol
                    Contract.grdCPMBuy.rowHeight(llRow) = fgBoxGridH
                    Contract.grdCPMBuy.Col = CPMID
                    Contract.grdCPMBuy.CellBackColor = LIGHTYELLOW
                    Contract.grdCPMBuy.Col = CPMEXPAND
                    Contract.grdCPMBuy.CellFontName = "Monotype Sorts"
                    Contract.grdCPMBuy.CellFontSize = 7
                    Contract.grdCPMBuy.CellForeColor = BLUE
                    Contract.grdCPMBuy.Col = CPMVEHICLE
                    Contract.grdCPMBuy.CellFontItalic = True
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = "0"
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMSORT) = ""
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMCXFCODE) = "0"
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMCOMMSHOWON) = "NNDND"   'Proposal;Order;Spots/Copy;Invoice;Insertion
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "True"
                Next ilAdd
                llRow = Contract.grdCPMBuy.rows
                Do
                    llRow = llRow - 1
                Loop While Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) = ""
                Contract.grdCPMBuy.Row = llRow + 1
            End If
            
            'Rebalance
            If llFirstHiddenLnRow <> -1 Then
                If Val(slHiddenImpressions) <> Val(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS)) Then
                    Contract.grdCPMBuy.TextMatrix(llFirstHiddenLnRow, CPMIMPRESSIONS) = gAddStr(Contract.grdCPMBuy.TextMatrix(llFirstHiddenLnRow, CPMIMPRESSIONS), gSubStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS), slHiddenImpressions))
                    Contract.grdCPMBuy.TextMatrix(llFirstHiddenLnRow, CPMCOST) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(llFirstHiddenLnRow, CPMIMPRESSIONS), Contract.grdCPMBuy.TextMatrix(llFirstHiddenLnRow, CPMCPMVALUE)), "1000.00")
                End If
            End If
            mSetFGColorForRows
        End If
    End If
    Contract.grdCPMBuy.Redraw = True
    gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbDefault

End Sub

Public Sub mSetBGColorWhite(blSaveEnableRow As Boolean)
    Dim llRow As Long
    Dim llCol As Long
    Dim llSvRow As Long
    Dim llSvCol As Long
    
    Contract.grdCPMBuy.Redraw = False
    If blSaveEnableRow Then
        llSvRow = lmCPMEnableRow
        llSvCol = lmCPMEnableCol
    Else
        llSvRow = Contract.grdCPMBuy.Row
        llSvCol = Contract.grdCPMBuy.Col
    End If
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        If Contract.grdCPMBuy.TextMatrix(llRow, CPMCOLOR) = Trim$(str$(GRAY)) Then
            Contract.grdCPMBuy.Row = llRow
            For llCol = CPMVEHICLE To CPMCOMMENT Step 1 'CPMCOST Step 1
                Contract.grdCPMBuy.Col = llCol
                Contract.grdCPMBuy.CellBackColor = WHITE
            Next llCol
            Contract.grdCPMBuy.TextMatrix(llRow, CPMCOLOR) = Trim$(str$(WHITE))
            Exit For
        End If
    Next llRow
    If blSaveEnableRow Then
        lmCPMEnableRow = llSvRow
        If lmCPMEnableRow <> -1 Then
            Contract.grdCPMBuy.Row = lmCPMEnableRow
        End If
        lmCPMEnableCol = llSvCol
        If lmCPMEnableCol <> -1 Then
            Contract.grdCPMBuy.Col = lmCPMEnableCol
        End If
    Else
        If llSvRow <> -1 Then
            Contract.grdCPMBuy.Row = llSvRow
        End If
        If llSvCol <> -1 Then
            Contract.grdCPMBuy.Col = llSvCol
        End If
    End If
    Contract.grdCPMBuy.Redraw = True
End Sub
Public Sub mSetBGColorGray(llRow As Long)
    Dim llRowLoop As Long
    Dim llCol As Long
    Dim llSvRow As Long
    Dim llSvCol As Long
    
    'If lmCPMEnableRow < Contract.grdCPMBuy.FixedRows Then
    '    Exit Sub
    'End If
    Contract.grdCPMBuy.Redraw = False
    gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbHourglass
    llSvRow = lmCPMEnableRow
    llSvCol = lmCPMEnableCol
    For llRowLoop = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        If Contract.grdCPMBuy.TextMatrix(llRowLoop, CPMCOLOR) = Trim$(str(GRAY)) Then
            If llRowLoop = llRow Then
                gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbDefault
                Contract.grdCPMBuy.Redraw = True
                Exit Sub
            End If
            Contract.grdCPMBuy.Row = llRowLoop
            For llCol = CPMVEHICLE To CPMCOMMENT Step 1 'CPMCOST Step 1
                Contract.grdCPMBuy.Col = llCol
                Contract.grdCPMBuy.CellBackColor = WHITE
            Next llCol
            Contract.grdCPMBuy.TextMatrix(llRowLoop, CPMCOLOR) = Trim$(str$(WHITE))
            Exit For
        End If
    Next llRowLoop
    If Contract.grdCPMBuy.TextMatrix(llRow, CPMCOLOR) <> Trim$(str(GRAY)) Then
        If llRow >= Contract.grdCPMBuy.FixedRows Then
            Contract.grdCPMBuy.Row = llRow
            For llCol = CPMVEHICLE To CPMCOMMENT Step 1 'CPMCOST Step 1
                Contract.grdCPMBuy.Col = llCol
                Contract.grdCPMBuy.CellBackColor = GRAY
            Next llCol
            Contract.grdCPMBuy.TextMatrix(llRow, CPMCOLOR) = Trim$(str(GRAY))
            Contract.grdCPMBuy.Col = CPMID
            Contract.grdCPMBuy.CellBackColor = LIGHTYELLOW
        End If
    End If
    lmCPMEnableRow = llSvRow
    If lmCPMEnableRow <> -1 Then
        Contract.grdCPMBuy.Row = lmCPMEnableRow
    End If
    lmCPMEnableCol = llSvCol
    If lmCPMEnableCol <> -1 Then
        Contract.grdCPMBuy.Col = lmCPMEnableCol
    End If
    Contract.grdCPMBuy.Redraw = True
    gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbDefault
End Sub

Public Sub mSetFGColorForRows()
    Dim llRow As Long
    Dim llCol As Long
    Dim llSvRow As Long
    Dim llSvCol As Long
    Dim llColor As Long
    
    Contract.grdCPMBuy.Redraw = False
    llSvRow = lmCPMEnableRow
    llSvCol = lmCPMEnableCol
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        Contract.grdCPMBuy.Row = llRow
        If Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "P" Then
            llColor = BLUE
            'For llCol = CPMVEHICLE To CPMCOMMENT Step 1 'CPMCOST Step 1    'CPMCPMVALUE Step 1
            '    Contract.grdCPMBuy.Col = llCol
            '    Contract.grdCPMBuy.CellForeColor = BLUE
            'Next llCol
        ElseIf Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H" Then
            llColor = MIDDLEBLUE
            'For llCol = CPMVEHICLE To CPMCOMMENT Step 1 'CPMCOST Step 1    'CPMCPMVALUE Step 1
            '    Contract.grdCPMBuy.Col = llCol
            '    Contract.grdCPMBuy.CellForeColor = MIDDLEBLUE
            'Next llCol
        Else
            llColor = BLACK
            'For llCol = CPMVEHICLE To CPMCOMMENT Step 1 'CPMCOST Step 1    'CPMCPMVALUE Step 1
            '    Contract.grdCPMBuy.Col = llCol
            '    Contract.grdCPMBuy.CellForeColor = BLACK
            'Next llCol
        End If
        For llCol = CPMVEHICLE To CPMCOMMENT Step 1 'CPMCOST Step 1    'CPMCPMVALUE Step 1
            Contract.grdCPMBuy.Col = llCol
            '10738 added Imp
'            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMBILLED) <> "Y") Or (llCol = CPMENDDATE) Then
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMBILLED) <> "Y") Or (llCol = CPMENDDATE) Or llCol = CPMIMPRESSIONS Then
                Contract.grdCPMBuy.CellForeColor = llColor
            Else
                Contract.grdCPMBuy.CellForeColor = DARKGREEN
            End If
        Next llCol
        Contract.grdCPMBuy.Col = CPMID
        Contract.grdCPMBuy.CellBackColor = LIGHTYELLOW
    Next llRow
    lmCPMEnableRow = llSvRow
    lmCPMEnableCol = llSvCol
    If lmCPMEnableRow <> -1 Then
        Contract.grdCPMBuy.Row = lmCPMEnableRow
    End If
    If lmCPMEnableCol <> -1 Then
        Contract.grdCPMBuy.Col = lmCPMEnableCol
    End If
    Contract.grdCPMBuy.Redraw = True
End Sub

Public Sub gCPMSort(blSortByLineNo As Boolean)
    'Sort by Line number (Major sort by standard Line  # and Package Line #, Minor sort hidden lines within package line numbers
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slLine As String
    Dim ilVef As Integer
    Dim llRif As Long
    Dim slGpSort As String
    Dim slVehName As String
    Dim slVehSort As String
    Dim slDPSort As String
    Dim ilLoop As Integer
    Dim slKey As String

    Contract.grdCPMBuy.Redraw = False

    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        slStr = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMID))
        If slStr <> "" Then
            slVehName = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE))
            slVehSort = "999"
            ilVef = gBinarySearchVef(Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE)))
            If ilVef <> -1 Then
                slVehName = tgMVef(ilVef).sName
                slVehSort = Trim$(str$(tgMVef(ilVef).iSort))
                'tmMnfSrchKey.iCode = tgMVef(ilVef).iOwnerMnfCode
                'If tmMnf.iCode <> tmMnfSrchKey.iCode Then
                '    ilRet = btrGetEqual(hlMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                '    If ilRet <> BTRV_ERR_NONE Then
                '        tmMnf.iGroupNo = 999
                '        tmMnf.iCode = 0
                '    End If
                'End If
                'slGpSort = Trim$(str$(tmMnf.iGroupNo))
                slGpSort = ""
            End If
            slDPSort = "999"
            If (Asc(tgSpf.sUsingFeatures5) And CNTRINVSORTRC) = CNTRINVSORTRC Then     'use rate card items sort code
                For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                    If (tgMRif(llRif).iVefCode = Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE))) And (tgMRif(llRif).iRdfCode = Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMRDFCODE))) Then
                        If (tgMRif(llRif).iSort > 0) Then
                            slDPSort = Trim$(str$(tgMRif(llRif).iSort))
                        End If
                        Exit For
                    End If
                Next llRif
            End If
            If (Asc(tgSpf.sUsingFeatures5) And CNTRINVSORTLN) = CNTRINVSORTLN Then
                slDPSort = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMID))
            Else
                If (slDPSort = "999") Then
                    ilLoop = gBinarySearchRdf(Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMRDFCODE)))
                    If ilLoop <> -1 Then
                        If tgMRdf(ilLoop).iSortCode > 0 Then
                            slDPSort = Trim$(str$(tgMRdf(ilLoop).iSortCode))
                        End If
                    End If
                End If
                If (slDPSort = "999") Then
                    slDPSort = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMID))
                End If
            End If
            Do While Len(slGpSort) < 3
                slGpSort = "0" & slGpSort
            Loop
            Do While Len(slVehSort) < 3
                slVehSort = "0" & slVehSort
            Loop
            Do While Len(slDPSort) < 4
                slDPSort = "0" & slDPSort
            Loop
            slLine = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMID))
            Do While Len(slLine) < 4
                slLine = "0" & slLine
            Loop
            slKey = slGpSort & slVehSort & slVehName & slDPSort & slLine
        
            If Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE)) = "P" Then
                slSort = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMID))
                Do While Len(slSort) < 4
                    slSort = "0" & slSort
                Loop
                slLine = slSort
                slSort = slSort & "A" '& slLine
                If blSortByLineNo Then
                    slLine = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMID))
                    Do While Len(slLine) < 4
                        slLine = "0" & slLine
                    Loop
                End If
                slSort = slSort & slLine
            ElseIf Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE)) = "H" Then
                slSort = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID))
                Do While Len(slSort) < 4
                    slSort = "0" & slSort
                Loop
                slLine = slSort
                slSort = slSort & "C"
                If blSortByLineNo Then
                    slLine = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMID))
                    Do While Len(slLine) < 4
                        slLine = "0" & slLine
                    Loop
                End If
                slSort = slSort & slLine
            Else    'Find parent to get line number
                slSort = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMID))
                Do While Len(slSort) < 4
                    slSort = "0" & slSort
                Loop
                slLine = slSort
                slSort = slSort & "~" & slLine
            End If
            slSort = slSort & slKey
            slStr = Contract.grdCPMBuy.TextMatrix(llRow, CPMSORT)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            slRow = Trim$(str$(llRow))
            Do While Len(slRow) < 4
                slRow = "0" & slRow
            Loop
            Contract.grdCPMBuy.TextMatrix(llRow, CPMSORT) = slSort & slStr & "|" & slRow
        End If
    Next llRow
    imLastColSorted = -1    'ascending only
    gGrid_SortByCol Contract.grdCPMBuy, CPMID, CPMSORT, imLastColSorted, imLastSort
    mSetFGColorForRows
    Contract.grdCPMBuy.Redraw = True
End Sub



Public Sub gCPMLineNo(ilLineNo As Integer)
    Dim llRow As Long
    Dim llLnRow As Long
    Dim llPkgRow As Long
    Dim llFocusRow As Long
    Dim slStr As String
    
    Contract.grdCPMBuy.Redraw = False
    mCPMSetShow
    llFocusRow = -1
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        slStr = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMID))
        If slStr <> "" Then
            If Val(slStr) = ilLineNo Then
                llFocusRow = llRow
                If Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "P" Then
                    If bgHideHiddenLines Then
                        If Contract.grdCPMBuy.TextMatrix(llRow, CPMEXPAND) = "s" Then
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMEXPAND) = "t"
                            For llLnRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                                If Contract.grdCPMBuy.TextMatrix(llRow, CPMID) = Contract.grdCPMBuy.TextMatrix(llLnRow, CPMPACKAGEID) Then
                                    Contract.grdCPMBuy.rowHeight(llLnRow) = fgBoxGridH
                                    Contract.grdCPMBuy.TextMatrix(llLnRow, CPMSHOWROW) = True
                                End If
                            Next llLnRow
                        End If
                    End If
                ElseIf Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H" Then
                    If bgHideHiddenLines Then
                        If Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = False Then
                            'find package line
                            For llPkgRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                                If Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMID) Then
                                    Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMEXPAND) = "s"
                                    For llLnRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                                        If Contract.grdCPMBuy.TextMatrix(llPkgRow, CPMID) = Contract.grdCPMBuy.TextMatrix(llLnRow, CPMPACKAGEID) Then
                                            Contract.grdCPMBuy.rowHeight(llLnRow) = fgBoxGridH
                                            Contract.grdCPMBuy.TextMatrix(llLnRow, CPMSHOWROW) = True
                                        End If
                                    Next llLnRow
                                    Exit For
                                End If
                            Next llPkgRow
                        End If
                    End If
                End If
            End If
        End If
    Next llRow
    If llFocusRow >= Contract.grdCPMBuy.FixedRows Then
        If Not Contract.grdCPMBuy.RowIsVisible(llFocusRow) Then
            Contract.grdCPMBuy.TopRow = Contract.grdCPMBuy.FixedRows
            For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                If Not Contract.grdCPMBuy.RowIsVisible(llFocusRow) Then
                    imCPMIgnoreScroll = True
                    Contract.grdCPMBuy.TopRow = Contract.grdCPMBuy.TopRow + 1
                End If
                If llRow = llFocusRow Then
                    Exit For
                End If
            Next llRow
        End If
        mSetBGColorGray llFocusRow
        'Contract.pbcLineArrow.Visible = False
        'Contract.pbcLineArrow.Move Contract.grdCPMBuy.Left - Contract.pbcLineArrow.Width - 30, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(llFocusRow) + (Contract.grdCPMBuy.RowHeight(llFocusRow) - Contract.pbcLineArrow.Height) / 2
        'Contract.pbcLineArrow.Visible = True
    End If
    Contract.grdCPMBuy.Redraw = True

End Sub

Public Function gCreateCPMVef() As Integer
    Dim tlVef As VEF
    Dim llRow As Long
    Dim ilRet As Integer
    Dim llRet As Long
    Dim slSQLQuery As String
    Dim blDynamicCreated As Boolean
    
    gCreateCPMVef = True
    blDynamicCreated = False
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "P") Then
                If (Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = "0") Or (Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE)) = "") Then
                    If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE) = "0") Or (Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE)) = "") Then
                        blDynamicCreated = True
                        'Create dynamic package
                        gInitVef tlVef
                        tlVef.iCode = 0
                        tlVef.sName = Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE)
                        tlVef.sType = "P"
                        tlVef.sState = "A"
                        tlVef.sExportRAB = "N"
                        slSQLQuery = "Insert Into VEF_Vehicles ( "
                        slSQLQuery = slSQLQuery & "vefCode, "
                        slSQLQuery = slSQLQuery & "vefName, "
                        slSQLQuery = slSQLQuery & "vefAddr1, "
                        slSQLQuery = slSQLQuery & "vefAddr2, "
                        slSQLQuery = slSQLQuery & "vefAddr3, "
                        slSQLQuery = slSQLQuery & "vefPhone, "
                        slSQLQuery = slSQLQuery & "vefFax, "
                        slSQLQuery = slSQLQuery & "vefUnused1, "
                        slSQLQuery = slSQLQuery & "vefDialPos, "
                        slSQLQuery = slSQLQuery & "vefPvfCode, "
                        slSQLQuery = slSQLQuery & "vefReallDnfCode, "
                        slSQLQuery = slSQLQuery & "vefUpdateRvf1, "
                        slSQLQuery = slSQLQuery & "vefUpdateRvf2, "
                        slSQLQuery = slSQLQuery & "vefUpdateRvf3, "
                        slSQLQuery = slSQLQuery & "vefUpdateRvf4, "
                        slSQLQuery = slSQLQuery & "vefUpdateRvf5, "
                        slSQLQuery = slSQLQuery & "vefUpdateRvf6, "
                        slSQLQuery = slSQLQuery & "vefUpdateRvf7, "
                        slSQLQuery = slSQLQuery & "vefUpdateRvf8, "
                        slSQLQuery = slSQLQuery & "vefCombineVefCode, "
                        slSQLQuery = slSQLQuery & "vefMnfHubCode, "
                        slSQLQuery = slSQLQuery & "vefTrfCode, "
                        slSQLQuery = slSQLQuery & "vefType, "
                        slSQLQuery = slSQLQuery & "vefCodeStn, "
                        slSQLQuery = slSQLQuery & "vefvefCode, "
                        slSQLQuery = slSQLQuery & "vefOwnerMnfCode, "
                        slSQLQuery = slSQLQuery & "vefProdPct, "
                        slSQLQuery = slSQLQuery & "vefProdPct2, "
                        slSQLQuery = slSQLQuery & "vefProdPct3, "
                        slSQLQuery = slSQLQuery & "vefProdPct4, "
                        slSQLQuery = slSQLQuery & "vefProdPct5, "
                        slSQLQuery = slSQLQuery & "vefProdPct6, "
                        slSQLQuery = slSQLQuery & "vefProdPct7, "
                        slSQLQuery = slSQLQuery & "vefProdPct8, "
                        slSQLQuery = slSQLQuery & "vefState, "
                        slSQLQuery = slSQLQuery & "vefmnfGroup, "
                        slSQLQuery = slSQLQuery & "vefmnfGroup2, "
                        slSQLQuery = slSQLQuery & "vefmnfGroup3, "
                        slSQLQuery = slSQLQuery & "vefmnfGroup4, "
                        slSQLQuery = slSQLQuery & "vefmnfGroup5, "
                        slSQLQuery = slSQLQuery & "vefmnfGroup6, "
                        slSQLQuery = slSQLQuery & "vefmnfGroup7, "
                        slSQLQuery = slSQLQuery & "vefmnfGroup8, "
                        slSQLQuery = slSQLQuery & "vefSort, "
                        slSQLQuery = slSQLQuery & "vefdnfCode, "
                        slSQLQuery = slSQLQuery & "vefmnfDemo, "
                        slSQLQuery = slSQLQuery & "vefmnfSSCode1, "
                        slSQLQuery = slSQLQuery & "vefmnfSSCode2, "
                        slSQLQuery = slSQLQuery & "vefmnfSSCode3, "
                        slSQLQuery = slSQLQuery & "vefmnfSSCode4, "
                        slSQLQuery = slSQLQuery & "vefmnfSSCode5, "
                        slSQLQuery = slSQLQuery & "vefmnfSSCode6, "
                        slSQLQuery = slSQLQuery & "vefmnfSSCode7, "
                        slSQLQuery = slSQLQuery & "vefmnfSSCode8, "
                        slSQLQuery = slSQLQuery & "vefExportRAB, "
                        slSQLQuery = slSQLQuery & "vefVsfCode, "
                        slSQLQuery = slSQLQuery & "vefRateAud, "
                        slSQLQuery = slSQLQuery & "vefCPPCPM, "
                        slSQLQuery = slSQLQuery & "vefYearAvails, "
                        slSQLQuery = slSQLQuery & "vefPctSellout, "
                        slSQLQuery = slSQLQuery & "vefMnfVehGp2, "
                        slSQLQuery = slSQLQuery & "vefMnfVehGp3Mkt, "
                        slSQLQuery = slSQLQuery & "vefMnfVehGp4Fmt, "
                        slSQLQuery = slSQLQuery & "vefMnfVehGp5Rsch, "
                        slSQLQuery = slSQLQuery & "vefMnfVehGp6Sub, "
                        slSQLQuery = slSQLQuery & "vefNrfCode, "
                        slSQLQuery = slSQLQuery & "vefSSMnfCode, "
                        slSQLQuery = slSQLQuery & "vefStdPrice, "
                        slSQLQuery = slSQLQuery & "vefStdInvTime, "
                        slSQLQuery = slSQLQuery & "vefStdAlter, "
                        slSQLQuery = slSQLQuery & "vefStdIndex, "
                        slSQLQuery = slSQLQuery & "vefStdAlterName, "
                        slSQLQuery = slSQLQuery & "vefRemoteID, "
                        slSQLQuery = slSQLQuery & "vefAutoCode, "
                        slSQLQuery = slSQLQuery & "vefExtUpdateRvf1, "
                        slSQLQuery = slSQLQuery & "vefExtUpdateRvf2, "
                        slSQLQuery = slSQLQuery & "vefExtUpdateRvf3, "
                        slSQLQuery = slSQLQuery & "vefExtUpdateRvf4, "
                        slSQLQuery = slSQLQuery & "vefExtUpdateRvf5, "
                        slSQLQuery = slSQLQuery & "vefExtUpdateRvf6, "
                        slSQLQuery = slSQLQuery & "vefExtUpdateRvf7, "
                        slSQLQuery = slSQLQuery & "vefExtUpdateRvf8, "
                        slSQLQuery = slSQLQuery & "vefStdSelCriteria, "
                        slSQLQuery = slSQLQuery & "vefStdOverrideFlag, "
                        slSQLQuery = slSQLQuery & "vefContact "
                        slSQLQuery = slSQLQuery & ") "
                        slSQLQuery = slSQLQuery & "Values ( "
                        slSQLQuery = slSQLQuery & "Replace" & ", "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sName) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sAddr(0)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sAddr(1)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sAddr(2)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sPhone) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sFax) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sUnused1) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sDialPos) & "', "
                        slSQLQuery = slSQLQuery & tlVef.lPvfCode & ", "
                        slSQLQuery = slSQLQuery & tlVef.iReallDnfCode & ", "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sUpdateRVF(0)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sUpdateRVF(1)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sUpdateRVF(2)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sUpdateRVF(3)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sUpdateRVF(4)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sUpdateRVF(5)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sUpdateRVF(6)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sUpdateRVF(7)) & "', "
                        slSQLQuery = slSQLQuery & tlVef.iCombineVefCode & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfHubCode & ", "
                        slSQLQuery = slSQLQuery & tlVef.iTrfCode & ", "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sType) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sCodeStn) & "', "
                        slSQLQuery = slSQLQuery & tlVef.iVefCode & ", "
                        slSQLQuery = slSQLQuery & tlVef.iOwnerMnfCode & ", "
                        slSQLQuery = slSQLQuery & tlVef.iProdPct(0) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iProdPct(1) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iProdPct(2) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iProdPct(3) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iProdPct(4) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iProdPct(5) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iProdPct(6) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iProdPct(7) & ", "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sState) & "', "
                        slSQLQuery = slSQLQuery & tlVef.iMnfGroup(0) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfGroup(1) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfGroup(2) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfGroup(3) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfGroup(4) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfGroup(5) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfGroup(6) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfGroup(7) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iSort & ", "
                        slSQLQuery = slSQLQuery & tlVef.iDnfCode & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfDemo & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfSSCode(0) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfSSCode(1) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfSSCode(2) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfSSCode(3) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfSSCode(4) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfSSCode(5) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfSSCode(6) & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfSSCode(7) & ", "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sExportRAB) & "', "
                        slSQLQuery = slSQLQuery & tlVef.lVsfCode & ", "
                        slSQLQuery = slSQLQuery & tlVef.lRateAud & ", "
                        slSQLQuery = slSQLQuery & tlVef.lCPPCPM & ", "
                        slSQLQuery = slSQLQuery & tlVef.lYearAvails & ", "
                        slSQLQuery = slSQLQuery & tlVef.iPctSellout & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfVehGp2 & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfVehGp3Mkt & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfVehGp4Fmt & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfVehGp5Rsch & ", "
                        slSQLQuery = slSQLQuery & tlVef.iMnfVehGp6Sub & ", "
                        slSQLQuery = slSQLQuery & tlVef.iNrfCode & ", "
                        slSQLQuery = slSQLQuery & tlVef.iSSMnfCode & ", "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sStdPrice) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sStdInvTime) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sStdAlter) & "', "
                        slSQLQuery = slSQLQuery & tlVef.iStdIndex & ", "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sStdAlterName) & "', "
                        slSQLQuery = slSQLQuery & tlVef.iRemoteID & ", "
                        slSQLQuery = slSQLQuery & tlVef.iAutoCode & ", "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf(0)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf(1)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf(2)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf(3)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf(4)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf(5)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf(6)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf(7)) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sStdSelCriteria) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sStdOverrideFlag) & "', "
                        slSQLQuery = slSQLQuery & "'" & gFixQuote(tlVef.sContact) & "' "
                        slSQLQuery = slSQLQuery & ") "
                        llRet = gInsertAndReturnCode(slSQLQuery, "VEF_Vehicles", "vefCode", "Replace")
                        If llRet <= 0 Then
                            gCreateCPMVef = False
                            Exit Function
                        End If
                        Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE) = llRet
                        tlVef.iCode = llRet
                        slSQLQuery = "Update VEF_Vehicles Set "
                        slSQLQuery = slSQLQuery & "vefRemoteID = " & tgUrf(0).iRemoteUserID & ", "
                        slSQLQuery = slSQLQuery & "vefAutoCode = " & tlVef.iCode & " "
                        slSQLQuery = slSQLQuery & "Where vefCode = " & tlVef.iCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        If llRet <> 0 Then
                            gCreateCPMVef = False
                            Exit Function
                        End If
                        'Update tgMVef and create VPF, then set vpfGMedium to P (M removed)
                        sgMVefStamp = ""
                        ilRet = gObtainVef()
                        ilRet = gVpfFind(Contract, tlVef.iCode)
                        slSQLQuery = "Update VPF_Vehicle_Options Set "
                        slSQLQuery = slSQLQuery & "vpfGMedium = '" & "P" & "' "
                        slSQLQuery = slSQLQuery & "Where vpfvefKCode  = " & tlVef.iCode
                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        If llRet <> 0 Then
                            gCreateCPMVef = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next llRow
    If blDynamicCreated Then
        Contract.mVehPop False, True
    End If
End Function

Public Sub gCPMPkVehPop(lbcLocal As Control)
    Dim ilVef As Integer
    Dim ilVpf As Integer
    Dim ilRet As Integer
    Dim llRow As Long
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    
    ReDim tgCPMPkVeh(0 To 0) As SORTCODE
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        ilVpf = gBinarySearchVpf(tgMVef(ilVef).iCode)
        If ilVpf <> -1 Then
            'If tgVpf(ilVpf).sGMedium = "M" Then
            If tgVpf(ilVpf).sGMedium = "P" Then
                If tgMVef(ilVef).sType = "P" Then
                    tgCPMPkVeh(UBound(tgCPMPkVeh)).sKey = "A" & Trim$(tgMVef(ilVef).sName) & "\" & tgMVef(ilVef).iCode
                    ReDim Preserve tgCPMPkVeh(0 To UBound(tgCPMPkVeh) + 1) As SORTCODE
                Else
                    tgCPMPkVeh(UBound(tgCPMPkVeh)).sKey = "B" & Trim$(tgMVef(ilVef).sName) & "\" & tgMVef(ilVef).iCode
                    ReDim Preserve tgCPMPkVeh(0 To UBound(tgCPMPkVeh) + 1) As SORTCODE
                End If
            End If
        End If
    Next ilVef
    'Add dynamic names
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "P") Then
                If (Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = "0") Or (Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE)) = "") Then
                    If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE) = "0") Or (Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE)) = "") Then
                        tgCPMPkVeh(UBound(tgCPMPkVeh)).sKey = "A" & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE)) & "\" & "0"
                        ReDim Preserve tgCPMPkVeh(0 To UBound(tgCPMPkVeh) + 1) As SORTCODE
                    End If
                End If
            End If
        End If
    Next llRow
    If UBound(tgCPMPkVeh) - 1 > 0 Then
        ArraySortTyp fnAV(tgCPMPkVeh(), 0), UBound(tgCPMPkVeh), 0, LenB(tgCPMPkVeh(0)), 0, LenB(tgCPMPkVeh(0).sKey), 0
    End If
    
    lbcLocal.Clear
    For ilVef = 0 To UBound(tgCPMPkVeh) - 1 Step 1
        slNameCode = tgCPMPkVeh(ilVef).sKey
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        lbcLocal.AddItem Mid$(slName, 2)
        lbcLocal.ItemData(lbcLocal.NewIndex) = slCode
    Next ilVef
End Sub
Public Sub gAddMissingCPMVehicle()
    Dim llRow As Long
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim ilVef As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) <> "P") Then
                ilFound = False
                For ilTest = 0 To UBound(tmVehicleCode2) - 1 Step 1  'lbcVehicle.ListCount - 1 Step 1
                    slNameCode = tmVehicleCode2(ilTest).sKey    'lbcVehicle.List(ilTest)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE)) = Val(slCode) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    ilVef = gBinarySearchVef(Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE)))
                    If ilVef <> -1 Then
                        slName = "999" & "|" & "999" & "|" & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & "\" & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE))
                        tmVehicleCode2(UBound(tmVehicleCode2)).sKey = slName
                        ReDim Preserve tmVehicleCode2(0 To UBound(tmVehicleCode2) + 1) As SORTCODE
                        Contract.lbcLnVehicle(igTabMapIndex).AddItem Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE))
                    End If
                End If
            End If
        End If
    Next llRow
End Sub

Private Sub mCalculateValues()
    'Base formula: Cost = Impression * CPM /1000
    'Use Impressions to either compute CPM or Cost
    Dim llRow As Long
    Dim slImpressions As String
    Dim slCost As String
    Dim slCPM As String
    Dim slRatio As String
    Dim llLastRow As Long
    Dim slBalance As String
    Dim slTotalCost As String
    Dim slStr As String
    Dim ilCount As Integer
    Dim slAnyNonZero As Boolean
    Dim slTotalImpressions As String
    Dim slTotalCPM As String
    
    Contract.grdCPMBuy.Redraw = False
    If Trim$(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMPRICETYPE)) = "Flat Rate" Then
        mComputeTotals
        Contract.grdCPMBuy.Redraw = True
        gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbDefault
        Exit Sub
    End If
    gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbHourglass
    Select Case lmCPMEnableCol
        Case CPMIMPRESSIONS
            If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS) <> "" Then
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE) <> "" Then
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS), Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE)), "1000.00")
                ElseIf Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST) <> "" Then
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST), "1000.00"), Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS))
                End If
            Else
                gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbDefault
                Contract.grdCPMBuy.Redraw = True
                Exit Sub
            End If
        Case CPMCPMVALUE
            If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE) <> "" Then
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS) <> "" Then
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS), Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE)), "1000.00")
                ElseIf Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST) <> "" Then
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST), "1000.00"), Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE))
                End If
            Else
                gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbDefault
                Contract.grdCPMBuy.Redraw = True
                Exit Sub
            End If
        Case CPMCOST
            If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST) <> "" Then
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS) <> "" Then
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST), "1000.00"), Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS))
                ElseIf Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE) <> "" Then
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST), "1000.00"), Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE))
                End If
            Else
                gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbDefault
                Contract.grdCPMBuy.Redraw = True
                Exit Sub
            End If
    End Select
    If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) = "P" Then
        'Distribute to hidden lines, first compute current hidden total so that ratio can be computed
        Select Case lmCPMEnableCol
            Case CPMIMPRESSIONS
                slImpressions = "0"
                llLastRow = -1
                ilCount = 0
                For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                    If (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H") And (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID) = Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID)) Then
                        slImpressions = gAddStr(slImpressions, Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS))
                        llLastRow = llRow
                        ilCount = ilCount + 1
                    End If
                Next llRow
                If slImpressions = "0" Then
                    slImpressions = gDivStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS), str(ilCount))
                    slTotalImpressions = "0"
                    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H") And (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID) = Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID)) Then
                            If llRow <> llLastRow Then
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS) = slImpressions
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST), "1000.00"), Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS))
                                slBalance = gAddStr(slBalance, Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS))
                                slTotalCost = gAddStr(slTotalImpressions, Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST))
                            Else
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS) = gSubStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS), slBalance)
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST), "1000.00"), Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS))
                                slTotalCost = gAddStr(slTotalCost, Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST))
                                Exit For
                            End If
                        End If
                    Next llRow
                Else
                    slBalance = "0"
                    slTotalCost = "0.00"
                    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H") And (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID) = Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID)) Then
                            If llRow <> llLastRow Then
                                slRatio = gDivStr(Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS) & ".00", slImpressions)
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS) = gMulStr(slRatio, Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS))
                                'slStr = Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS)
                                'gFormatStr slStr, 0, 0, Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS)
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS) = Round(Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS))
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS), Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE)), "1000.00")
                                slBalance = gAddStr(slBalance, Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS))
                                slTotalCost = gAddStr(slTotalCost, Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST))
                            Else
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS) = gSubStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS), slBalance)
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS), Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE)), "1000.00")
                                slTotalCost = gAddStr(slTotalCost, Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST))
                                Exit For
                            End If
                        End If
                    Next llRow
                End If
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST) = slTotalCost
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST), "1000.00"), Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS))
            Case CPMCPMVALUE, CPMCOST
                If lmCPMEnableCol = CPMCPMVALUE Then
                    Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS), Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE)), "1000.00")
                End If
                slCPM = "0.00"
                llLastRow = -1
                ilCount = 0
                For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                    If (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H") And (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID) = Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID)) Then
                        slCost = gAddStr(slCost, Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST))
                        llLastRow = llRow
                        ilCount = ilCount + 1
                    End If
                Next llRow
                If slCost = "0.00" Then
                    slCost = gDivStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST), str(ilCount))
                    slTotalCost = "0.00"
                    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H") And (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID) = Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID)) Then
                            If llRow <> llLastRow Then
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST) = slCost
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST), "1000.00"), Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS))
                                slBalance = gAddStr(slBalance, Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST))
                                slTotalCost = gAddStr(slTotalCost, Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST))
                            Else
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST) = gSubStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST), slBalance)
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST), "1000.00"), Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS))
                                slTotalCost = gAddStr(slTotalCost, Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST))
                                Exit For
                            End If
                        End If
                    Next llRow
                Else
                    slBalance = "0"
                    slTotalCost = "0.00"
                    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H") And (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID) = Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID)) Then
                            If llRow <> llLastRow Then
                                slRatio = gDivStr(Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST) & ".00", slCost)
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST) = gMulStr(slRatio, Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST))
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST), "1000.00"), Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS))
                                slBalance = gAddStr(slBalance, Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST))
                                slTotalCost = gAddStr(slTotalCost, Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST))
                            Else
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST) = gSubStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST), slBalance)
                                Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST), "1000.00"), Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS))
                                slTotalCost = gAddStr(slTotalCost, Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST))
                                Exit For
                            End If
                        End If
                    Next llRow
                End If
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST) = slTotalCost
                Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCPMVALUE) = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMCOST), "1000.00"), Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMIMPRESSIONS))
        End Select
        
    ElseIf Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) = "H" Then
        'Reset hiddden, then recalculate package
        slImpressions = "0"
        slCost = "0.00"
        For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "H") And (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMPACKAGEID) = Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID)) Then
                slImpressions = gAddStr(slImpressions, Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS))
                slCost = gAddStr(slCost, gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS), Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE)), "1000.00"))
            End If
        Next llRow
        slCPM = gDivStr(gMulStr(slCost, "1000.00"), slImpressions)
        For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "P") And (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMPACKAGEID) = Contract.grdCPMBuy.TextMatrix(llRow, CPMID)) Then
                Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS) = slImpressions
                Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE) = slCPM
                Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST) = slCost
            End If
        Next llRow
    End If
'    slTotalImpressions = "0"
'    slTotalCost = "0.00"
'    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.Rows - 1 Step 1
'        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
'
'            slImpressions = Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS)
'            If Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) <> "P" Then
'                slTotalImpressions = gAddStr(slTotalImpressions, slImpressions)
'            End If
'            gFormatStr slImpressions, FMTLEAVEBLANK, 0, slImpressions
'            Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS) = slImpressions
'
'            slCPM = Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE)
'            gFormatStr slCPM, FMTLEAVEBLANK, 2, slCPM
'            Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE) = slCPM
'
'            slCost = Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST)
'            If Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) <> "P" Then
'                slTotalCost = gAddStr(slTotalCost, slCost)
'            End If
'            gFormatStr slCost, FMTLEAVEBLANK, 2, slCost
'            Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST) = slCost
'        End If
'    Next llRow
'    slTotalCPM = gDivStr(gMulStr(slTotalCost, "1000.00"), slTotalImpressions)
'    gFormatStr slTotalImpressions, FMTLEAVEBLANK, 0, slTotalImpressions
'    Contract.edcLineNo(2).Text = slTotalImpressions
'    gFormatStr slTotalCPM, FMTLEAVEBLANK, 2, slTotalCPM
'    Contract.edcLineNo(3).Text = slTotalCPM
'    gFormatStr slTotalCost, FMTLEAVEBLANK, 2, slTotalCost
'    Contract.edcLineNo(4).Text = slTotalCost
    mComputeTotals
    Contract.grdCPMBuy.Redraw = True
    gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbDefault

End Sub

Public Function gCPMTestFields() As Boolean
    Dim llRow As Long
    Dim ilCol As Integer
    Dim ilRet As Integer
    Dim blError As Boolean
    Dim slCPMAllowedDate As String
    Dim blDateChgd As Boolean
    Dim llHiddenRow As Long

    Contract.grdCPMBuy.Redraw = False
    gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbHourglass
    gCPMTestFields = True
    mObtainCPMAllowedStartDate
    blError = False
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) = "CBS") Or (Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) = "Deleted") Then
                For ilCol = CPMVEHICLE To CPMCOMMENT Step 1 'CPMCOST Step 1
                    Select Case ilCol
                        Case CPMVEHICLE
                        Case CPMDAYPART
                        'Case CPMCOPYTYPE
                        Case CPMPRICETYPE
                        Case CPMPOSITION
                        Case CPMPRIORITY
                        Case CPMPODTARGET
                        Case CPMSTARTDATE
                        Case CPMENDDATE
                        Case CPMIMPRESSIONS
                        Case CPMIMPRESSIONCAP
                        Case CPMCPMVALUE
                        Case CPMCOST
                    End Select
                Next ilCol
            Else
                slCPMAllowedDate = mCPMAllowedDate(llRow)
                For ilCol = CPMVEHICLE To CPMCOMMENT Step 1 'CPMCOST Step 1
                    Select Case ilCol
                        Case CPMVEHICLE
                        Case CPMDAYPART
                        'Case CPMCOPYTYPE
                        Case CPMPRICETYPE
                            If Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS)) = "" Then
                                ilRet = MsgBox("Price type " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) & " is missing", vbOKOnly + vbExclamation, "Incomplete")
                                gCPMTestFields = False
                                blError = True
                                Exit For
                            End If
                        Case CPMPOSITION
                        Case CPMPRIORITY
                        Case CPMPODTARGET
                        Case CPMSTARTDATE
                            If Not gIsDate(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) Then
                                '10731
                                ilRet = MsgBox("Invalid Start Date form for Digital " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART), vbOKOnly + vbExclamation, "Incomplete")
                                gCPMTestFields = False
                                blError = True
                                Exit For
                            Else
                                If (Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = "") Or (Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = "0") Then
                                    If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) < gDateValue(slCPMAllowedDate) Then
                                        '10731
                                        ilRet = MsgBox("Start Date for Digital " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) & " must be on or after " & slCPMAllowedDate, vbOKOnly + vbExclamation, "Incomplete")
                                        gCPMTestFields = False
                                        blError = True
                                        Exit For
                                    End If
                                Else
                                    'check if date changed
                                    blDateChgd = False
                                    If (Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) = "CBS") Or (Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) = "Deleted") Then
                                        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) <> "CBS") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) <> "Deleted") Then
                                            blDateChgd = True
                                        End If
                                    Else
                                        If Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) <> "" Then
                                            If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE)) <> gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) Then
                                                blDateChgd = True
                                            End If
                                        Else
                                            blDateChgd = True
                                        End If
                                    End If
                                    If blDateChgd Then
                                        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) = "CBS") Or (Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) = "Deleted") Then
                                            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) <> "CBS") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) <> "Deleted") Then
                                                If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) < gDateValue(slCPMAllowedDate) Then
                                                    '10731
                                                    ilRet = MsgBox("Start Date for Digital " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) & " must be on or after " & slCPMAllowedDate, vbOKOnly + vbExclamation, "Incomplete")
                                                    gCPMTestFields = False
                                                    blError = True
                                                    Exit For
                                                End If
                                            End If
                                        Else
                                            If Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) <> "" Then
                                                If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE)) <> gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) Then
                                                    If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE)) < gDateValue(slCPMAllowedDate) Then
                                                        '10731
                                                        ilRet = MsgBox("Start Date for Digital " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) & " can't be changed from " & Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE), vbOKOnly + vbExclamation, "Incomplete")
                                                        gCPMTestFields = False
                                                        blError = True
                                                        Exit For
                                                    Else
                                                        If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) < gDateValue(slCPMAllowedDate) Then
                                                            '10731
                                                            ilRet = MsgBox("Start Date for Digital " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) & " must be on or after " & slCPMAllowedDate, vbOKOnly + vbExclamation, "Incomplete")
                                                            gCPMTestFields = False
                                                            blError = True
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) < gDateValue(slCPMAllowedDate) Then
                                                    '10731
                                                    ilRet = MsgBox("Start Date for Digital " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) & " must be on or after " & slCPMAllowedDate, vbOKOnly + vbExclamation, "Incomplete")
                                                    gCPMTestFields = False
                                                    blError = True
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Case CPMENDDATE
                            '10719 we now want the end of the month if billed
                            slCPMAllowedDate = mCPMAllowedDate(llRow, True)
                            If Not gIsDate(Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)) Then
                                '10731
                                ilRet = MsgBox("Invalid End Date form for Digital " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART), vbOKOnly + vbExclamation, "Incomplete")
                                gCPMTestFields = False
                                blError = True
                                Exit For
                            Else
                                If gIsDate(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) Then
                                    If gDateValue(slCPMAllowedDate) < gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) Then
                                        slCPMAllowedDate = gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE))
                                    End If
                                End If
                                If (Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = "") Or (Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = "0") Then
                                    If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)) < gDateValue(slCPMAllowedDate) Then
                                        '10731
                                        ilRet = MsgBox("End Date for Digital " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) & " must be on or after " & slCPMAllowedDate, vbOKOnly + vbExclamation, "Incomplete")
                                        gCPMTestFields = False
                                        blError = True
                                        Exit For
                                    Else
                                        If gIsDate(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) Then
                                            If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)) < gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) Then
                                                '10731
                                                ilRet = MsgBox("End Date for Digital " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) & " must be on or after Start date " & Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE), vbOKOnly + vbExclamation, "Incomplete")
                                                gCPMTestFields = False
                                                blError = True
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Else
                                    'blDateChgd = False
                                    If (Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) <> "CBS") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) <> "Deleted") Then
                                        If Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGENDDATE) <> "" Then
                                            If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGENDDATE)) <> gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)) Then
                                                blDateChgd = True
                                            End If
                                        Else
                                            blDateChgd = True
                                        End If
                                    End If
                                    If blDateChgd Then
                                        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) <> "CBS") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) <> "Deleted") Then
                                            If Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGENDDATE) <> "" Then
                                                If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGENDDATE)) <> gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)) Then
                                                    If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGENDDATE)) < gDateValue(slCPMAllowedDate) Then
                                                        '10731
                                                        ilRet = MsgBox("End Date for Digital " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) & " can't be changed from " & Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGENDDATE), vbOKOnly + vbExclamation, "Incomplete")
                                                        gCPMTestFields = False
                                                        blError = True
                                                        Exit For
                                                    Else
                                                        If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)) < gDateValue(slCPMAllowedDate) Then
                                                            '10731
                                                            ilRet = MsgBox("End Date for Digital " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) & " must be on or after " & slCPMAllowedDate, vbOKOnly + vbExclamation, "Incomplete")
                                                            gCPMTestFields = False
                                                            blError = True
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)) < gDateValue(slCPMAllowedDate) Then
                                                    '10731
                                                    ilRet = MsgBox("End Date for Digital " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) & " must be on or after " & slCPMAllowedDate, vbOKOnly + vbExclamation, "Incomplete")
                                                    gCPMTestFields = False
                                                    blError = True
                                                    Exit For
                                                End If
                                            End If
                                            If gIsDate(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) Then
                                                If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)) < gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) Then
                                                    '10731
                                                    ilRet = MsgBox("End Date for Digital " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) & " must be on or after Start date " & Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE), vbOKOnly + vbExclamation, "Incomplete")
                                                    gCPMTestFields = False
                                                    blError = True
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Case CPMIMPRESSIONS
                            If Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMPRICETYPE)) <> "Flat Rate" Then
                                If Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS)) = "" Then
                                    '10731
                                    ilRet = MsgBox("Impressions for Digital " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) & " is missing", vbOKOnly + vbExclamation, "Incomplete")
                                    gCPMTestFields = False
                                    blError = True
                                    Exit For
                                End If

                            End If
                        Case CPMIMPRESSIONCAP
                        Case CPMCPMVALUE
                            If Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMPRICETYPE)) <> "Flat Rate" Then
                                If Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE)) = "" Then
                                    ilRet = MsgBox("CPM value for " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) & " is missing", vbOKOnly + vbExclamation, "Incomplete")
                                    gCPMTestFields = False
                                    blError = True
                                    Exit For
                                End If
                            End If
                        Case CPMCOST
                            If Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST)) = "" Then
                                '10731
                                ilRet = MsgBox("Cost for Digital " & Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) & " " & Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) & " is missing", vbOKOnly + vbExclamation, "Incomplete")
                                gCPMTestFields = False
                                blError = True
                                Exit For
                            End If
                        Case CPMCOMMENT
                    End Select
                                
                Next ilCol
                If blError Then
                    Exit For
                End If
                If Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "P" Then
                    For llHiddenRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                        If Contract.grdCPMBuy.TextMatrix(llRow, CPMID) = Contract.grdCPMBuy.TextMatrix(llHiddenRow, CPMPACKAGEID) Then
                            If (Contract.grdCPMBuy.TextMatrix(llHiddenRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llHiddenRow, CPMDAYPART) <> "") Then
                                If (Contract.grdCPMBuy.TextMatrix(llHiddenRow, CPMSTARTDATE) <> "CBS") And (Contract.grdCPMBuy.TextMatrix(llHiddenRow, CPMSTARTDATE) <> "Deleted") Then
                                    If gIsDate(Contract.grdCPMBuy.TextMatrix(llHiddenRow, CPMSTARTDATE)) And gIsDate(Contract.grdCPMBuy.TextMatrix(llHiddenRow, CPMENDDATE)) Then
                                        If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)) > gDateValue(Contract.grdCPMBuy.TextMatrix(llHiddenRow, CPMSTARTDATE)) Then
                                            ilRet = MsgBox("The Package line ID " & Contract.grdCPMBuy.TextMatrix(llRow, CPMID) & " dates must encompass the Hidden line dates", vbOKOnly + vbExclamation, "Incomplete")
                                            gCPMTestFields = False
                                            blError = True
                                            Exit For
                                        Else
                                            If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)) < gDateValue(Contract.grdCPMBuy.TextMatrix(llHiddenRow, CPMENDDATE)) Then
                                                ilRet = MsgBox("The Package line ID " & Contract.grdCPMBuy.TextMatrix(llRow, CPMID) & " dates must encompass the Hidden line dates", vbOKOnly + vbExclamation, "Incomplete")
                                                gCPMTestFields = False
                                                blError = True
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next llHiddenRow
                    If blError Then
                        Exit For
                    End If
                End If
            End If
        End If
    Next llRow
    Contract.grdCPMBuy.Redraw = True
    gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbDefault

End Function

Private Sub mObtainCPMAllowedStartDate()
    Dim slDate As String
    Dim llUnbilledDate As Long
    
    smCPMLLD = Format$(gNow(), "m/d/yy")
    '10719
    smCPMLBD = smCPMLLD
    If tgUrf(0).sChgCntr = "I" Then
        If (tgChfCntr.sBillCycle = "C") Then    'Or (tgChfCntr.sBillCycle = "D") Then
            gUnpackDate tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1), slDate
            slDate = gObtainEndCal(slDate)
        ElseIf (tgChfCntr.sBillCycle = "W") Then
            gUnpackDate tgSaf(0).iBLastWeeklyDate(0), tgSaf(0).iBLastWeeklyDate(1), slDate
            slDate = gObtainNextSunday(slDate)
        Else
            gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slDate
            slDate = gObtainEndStd(slDate)
        End If
        llUnbilledDate = gDateValue(slDate) + 1
        If (llUnbilledDate > 0) And (llUnbilledDate < gDateValue(smCPMLLD)) Then
            smCPMLLD = Format$(llUnbilledDate, "m/d/yy")
        End If
        '10719
        llUnbilledDate = gDateValue(slDate)
        If (llUnbilledDate > 0) And (llUnbilledDate < gDateValue(smCPMLBD)) Then
            smCPMLBD = Format$(llUnbilledDate, "m/d/yy")
        End If
    End If
End Sub

Public Sub gTrashCPMRow()
    Dim blRemove As Boolean
    Dim llRow As Long
    Dim ilCol As Integer
    Dim rst_Chf As ADODB.Recordset
    Dim slDate As String
    Dim slSQLQuery As String
    Dim ilCount As Integer
    Dim ilAdd As Integer
    '10569
    Dim blContinue As Boolean
    
    If (Contract.grdCPMBuy.Row < Contract.grdCPMBuy.FixedRows) Or (Contract.grdCPMBuy.Row >= Contract.grdCPMBuy.rows) Or (Contract.grdCPMBuy.Col < Contract.grdCPMBuy.FixedCols) Or (Contract.grdCPMBuy.Col >= Contract.grdCPMBuy.cols - 1) Then
        Exit Sub
    End If
    If (lmCPMEnableRow < Contract.grdCPMBuy.FixedRows) Or (lmCPMEnableRow >= Contract.grdCPMBuy.rows) Then
        Exit Sub
    End If
    blRemove = False
    If (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMPCFCODE) = "") Or (Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMPCFCODE) = "0") Then
        blRemove = True
    Else
        mObtainCPMAllowedStartDate
        slDate = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMORIGSTARTDATE)
        If (slDate <> "CBS") And (slDate <> "Deleted") Then
            If slDate <> "" Then
                If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) = "P" Then
                    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                        If Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID) Then
                            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) <> "CBS") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) <> "Deleted") Then
                                If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE)) < gDateValue(slDate) Then
                                    slDate = Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE)
                                End If
                            End If
                        End If
                    Next llRow
                End If
                If (slDate <> "CBS") And (slDate <> "Deleted") Then
                    'If gDateValue(slDate) >= gDateValue(smCPMLLD) Then
'                    If gDateValue(slDate) >= gDateValue(mCPMAllowedDate(lmCPMEnableRow)) Then
'                        blRemove = True
'                    Else
'                        If (tgChfCntr.sSchStatus = "F") Or (tgChfCntr.sSchStatus = "M") Or (tgChfCntr.sSchStatus = "I") Then
'                            MsgBox "CPM can't be deleted as it is scheduled", vbOKOnly + vbExclamation, "Warning"
'                        Else
'                            If tgChfCntr.iCntRevNo > 0 Then
'                                MsgBox "CPM can't be deleted as previous version exist", vbOKOnly + vbExclamation, "Warning"
'                            Else
'                                slSQLQuery = "Select Count(1) as CntrCount From chf_Contract_Header Where chfCntrNo = " & tgChfCntr.lCntrNo
'                                Set rst_Chf = gSQLSelectCall(slSQLQuery)
'                                If Not rst_Chf.EOF Then
'                                    If rst_Chf!CntrCount > 1 Then
'                                        MsgBox "CPM can't be deleted as previous version exist", vbOKOnly + vbExclamation, "Warning"
'                                    End If
'                                End If
'                                rst_Chf.Close
'                            End If
'                        End If
                    ' 10569 reversed order of first 2 tests. Test Scheduled first, as it matters for removal.
                    If (tgChfCntr.sSchStatus = "F") Or (tgChfCntr.sSchStatus = "M") Or (tgChfCntr.sSchStatus = "I") Then
                        MsgBox "CPM can't be deleted as it is scheduled", vbOKOnly + vbExclamation, "Warning"
                    Else
                        'swap this too as previous versions also affects if date matters.
                         blContinue = True
                        ' proposal of a previous version, which has been scheduled!
                        If tgChfCntr.iCntRevNo > 0 Then
                            MsgBox "CPM can't be deleted as previous version exist", vbOKOnly + vbExclamation, "Warning"
                             blContinue = False
                        Else
                            slSQLQuery = "Select Count(1) as CntrCount From chf_Contract_Header Where chfCntrNo = " & tgChfCntr.lCntrNo
                            Set rst_Chf = gSQLSelectCall(slSQLQuery)
                            If Not rst_Chf.EOF Then
                                If rst_Chf!CntrCount > 1 Then
                                    MsgBox "CPM can't be deleted as previous version exist", vbOKOnly + vbExclamation, "Warning"
                                    blContinue = False
                                End If
                            End If
                            rst_Chf.Close
                        End If
                        If blContinue And gDateValue(slDate) >= gDateValue(mCPMAllowedDate(lmCPMEnableRow)) Then
                            blRemove = True
                        End If
                    End If
                End If
            Else
                blRemove = True
            End If
        End If
    End If
    ilCount = 0
    If blRemove Then
        'New Row
        If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) = "P" Then
            sgGenMsg = "Remove Package Line or Package and Hidden Lines"
            sgCMCTitle(0) = "Pkg Only"
            sgCMCTitle(1) = "Pkg+Hidden"
            sgCMCTitle(2) = "Cancel"
            sgCMCTitle(3) = ""
            igDefCMC = 1
            igEditBox = 0
            GenMsg.Show vbModal
            If igAnsCMC = 0 Then
                Contract.grdCPMBuy.Redraw = False
                For llRow = Contract.grdCPMBuy.rows - 1 To Contract.grdCPMBuy.FixedRows Step -1
                    If llRow <> lmCPMEnableRow Then
                        If Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID) Then
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE))
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = ""
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) = "S"
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "True"
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMEXPAND) = ""
                        End If
                    Else
                        Contract.grdCPMBuy.RemoveItem llRow
                        ilCount = ilCount + 1
                    End If
                Next llRow
                imCPMChg = True
            ElseIf igAnsCMC = 1 Then
                Contract.grdCPMBuy.Redraw = False
                For llRow = Contract.grdCPMBuy.rows - 1 To Contract.grdCPMBuy.FixedRows Step -1
                    If llRow <> lmCPMEnableRow Then
                        If Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID) Then
                            Contract.grdCPMBuy.RemoveItem llRow
                            ilCount = ilCount + 1
                        End If
                    Else
                        Contract.grdCPMBuy.RemoveItem llRow
                        ilCount = ilCount + 1
                    End If
                Next llRow
                imCPMChg = True
            Else
                Exit Sub
            End If
        Else
            Contract.grdCPMBuy.RemoveItem lmCPMEnableRow
            ilCount = ilCount + 1
            imCPMChg = True
        End If
        For ilAdd = 0 To ilCount - 1 Step 1
            llRow = Contract.grdCPMBuy.rows
            Do
                llRow = llRow - 1
            Loop While Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) = ""
            llRow = llRow + 1
            Contract.grdCPMBuy.AddItem "", llRow
            Contract.grdCPMBuy.rowHeight(llRow) = fgBoxGridH
            Contract.grdCPMBuy.Row = llRow
            For ilCol = 0 To Contract.grdCPMBuy.cols - 1 Step 1
                Contract.grdCPMBuy.ColAlignment(ilCol) = flexAlignLeftCenter
            Next ilCol
            Contract.grdCPMBuy.Col = CPMID
            Contract.grdCPMBuy.CellBackColor = LIGHTYELLOW
            Contract.grdCPMBuy.Col = CPMEXPAND
            Contract.grdCPMBuy.CellFontName = "Monotype Sorts"
            Contract.grdCPMBuy.CellFontSize = 7
            Contract.grdCPMBuy.CellForeColor = BLUE
            Contract.grdCPMBuy.Col = CPMVEHICLE
            Contract.grdCPMBuy.CellFontItalic = True
            Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = "0"
            Contract.grdCPMBuy.TextMatrix(llRow, CPMSORT) = ""
            Contract.grdCPMBuy.TextMatrix(llRow, CPMCXFCODE) = "0"
            Contract.grdCPMBuy.TextMatrix(llRow, CPMCOMMSHOWON) = "NNDND"   'Proposal;Order;Spots/Copy;Invoice;Insertion
            Contract.grdCPMBuy.TextMatrix(llRow, CPMSHOWROW) = "True"
        Next ilAdd
        mSetFGColorForRows
        Contract.grdCPMBuy.Redraw = False
        mComputeTotals
        Contract.grdCPMBuy.Redraw = True
    End If
End Sub

Public Sub gResetPkgValues(llRow As Long)
    Dim llRowP As Long
    Dim llRowH As Long
    Dim slImpressions As String
    Dim slCost As String
    Dim slCPM As String
    Contract.grdCPMBuy.Redraw = False
    gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbHourglass
    
    For llRowP = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        If (llRow = llRowP) Or (llRow = -1) Then
            If (Contract.grdCPMBuy.TextMatrix(llRowP, CPMLNTYPE) = "P") Then
                slImpressions = "0"
                slCost = "0.00"
                For llRowH = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
                    If (Contract.grdCPMBuy.TextMatrix(llRowH, CPMPACKAGEID) = Contract.grdCPMBuy.TextMatrix(llRowP, CPMID)) Then
                        If (Contract.grdCPMBuy.TextMatrix(llRowH, CPMSTARTDATE) <> "CBS") And (Contract.grdCPMBuy.TextMatrix(llRowH, CPMSTARTDATE) <> "Deleted") Then
                            slImpressions = gAddStr(slImpressions, Contract.grdCPMBuy.TextMatrix(llRowH, CPMIMPRESSIONS))
                            slCost = gAddStr(slCost, Contract.grdCPMBuy.TextMatrix(llRowH, CPMCOST))
                        End If
                    End If
                Next llRowH
                Contract.grdCPMBuy.TextMatrix(llRowP, CPMIMPRESSIONS) = Round(slImpressions)
                gFormatStr slCost, FMTLEAVEBLANK, 2, slCost
                Contract.grdCPMBuy.TextMatrix(llRowP, CPMCOST) = slCost
                slCPM = gDivStr(gMulStr(Contract.grdCPMBuy.TextMatrix(llRowP, CPMCOST), "1000.00"), Contract.grdCPMBuy.TextMatrix(llRowP, CPMIMPRESSIONS))
                gFormatStr slCPM, FMTLEAVEBLANK, 2, slCPM
                Contract.grdCPMBuy.TextMatrix(llRowP, CPMCPMVALUE) = slCPM
                If llRow = llRowP Then
                    Exit For
                End If
            End If
        End If
    Next llRowP
    Contract.grdCPMBuy.Redraw = True
    gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbDefault
    
End Sub

Private Function mSaveCPMComment(llRow As Long) As Long
    Dim slStr As String
    Dim llRet As Long
    Dim slSQLQuery As String
    Dim cxf_rst As ADODB.Recordset
    Dim tlCxf As CXF
    Dim llCode As Long
    
    'Update comment if required
    tlCxf.sComType = "B"
    slStr = Contract.grdCPMBuy.TextMatrix(llRow, CPMCOMMSHOWON)
    If Mid$(slStr, 1, 1) = "Y" Then
        tlCxf.sShProp = "Y"
    Else
        tlCxf.sShProp = "N"
    End If
    If Mid$(slStr, 2, 1) = "Y" Then
        tlCxf.sShOrder = "Y"
    Else
        tlCxf.sShOrder = "N"
    End If
    tlCxf.sShSpot = "N"
    If Mid$(slStr, 4, 1) = "Y" Then
        tlCxf.sShInv = "Y"
    Else
        tlCxf.sShInv = "N"
    End If
    tlCxf.sShInsertion = "N"
    tlCxf.sUnused = ""
    tlCxf.sComment = Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMCOMMENT)) ' Chr$(0) ' & Chr$(0) 'sgTB
    tlCxf.iRemoteID = tgUrf(0).iRemoteUserID
    tlCxf.iSourceID = tgUrf(0).iRemoteUserID
    If (Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMCXFCODE)) = "") Or (Contract.grdCPMBuy.TextMatrix(llRow, CPMCXFCODE) = "0") Or (lmCntrChfCode <> tgChfCntr.lCode) Then  'New selected
        tlCxf.lCode = 0 'Autoincrement
    Else
        tlCxf.lCode = Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMCXFCODE))
    End If
    tlCxf.lAutoCode = tlCxf.lCode
    If (Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMCXFCODE)) = "") Or (Contract.grdCPMBuy.TextMatrix(llRow, CPMCXFCODE) = "0") Or lmCntrChfCode <> tgChfCntr.lCode Then  'New selected
        If Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMCOMMENT)) <> "" Then
            llRet = 0
            slSQLQuery = "Insert Into CXF_Hdr_Ln_Comments ( "
            slSQLQuery = slSQLQuery & "cxfCode, "
            slSQLQuery = slSQLQuery & "cxfComType, "
            slSQLQuery = slSQLQuery & "cxfShProp, "
            slSQLQuery = slSQLQuery & "cxfShOrder, "
            slSQLQuery = slSQLQuery & "cxfShSpot, "
            slSQLQuery = slSQLQuery & "cxfShInv, "
            slSQLQuery = slSQLQuery & "cxfRemoteID, "
            slSQLQuery = slSQLQuery & "cxfAutoCode, "
            slSQLQuery = slSQLQuery & "cxfSyncDate, "
            slSQLQuery = slSQLQuery & "cxfSyncTime, "
            slSQLQuery = slSQLQuery & "cxfSourceID, "
            slSQLQuery = slSQLQuery & "cxfShInsertion, "
            slSQLQuery = slSQLQuery & "cxfUnused, "
            slSQLQuery = slSQLQuery & "cxfComment "
            slSQLQuery = slSQLQuery & ") "
            slSQLQuery = slSQLQuery & "Values ( "
            slSQLQuery = slSQLQuery & "Replace" & ", "
            slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCxf.sComType) & "', "
            slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCxf.sShProp) & "', "
            slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCxf.sShOrder) & "', "
            slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCxf.sShSpot) & "', "
            slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCxf.sShInv) & "', "
            slSQLQuery = slSQLQuery & tlCxf.iRemoteID & ", "
            slSQLQuery = slSQLQuery & tlCxf.lAutoCode & ", "
            slSQLQuery = slSQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
            slSQLQuery = slSQLQuery & "'" & Format$(gNow(), sgSQLTimeForm) & "', "
            slSQLQuery = slSQLQuery & tlCxf.iSourceID & ", "
            slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCxf.sShInsertion) & "', "
            slSQLQuery = slSQLQuery & "'" & gFixQuote(Trim$(tlCxf.sUnused)) & "', "
            slSQLQuery = slSQLQuery & "'" & gFixQuote(Trim$(tlCxf.sComment)) & "'"
            slSQLQuery = slSQLQuery & ") "
            llRet = gInsertAndReturnCode(slSQLQuery, "CXF_Hdr_Ln_Comments", "cxfCode", "Replace")
            If llRet <= 0 Then
                mSaveCPMComment = 0
                Exit Function
            End If
            tlCxf.lCode = llRet
            mSaveCPMComment = tlCxf.lCode
            Contract.grdCPMBuy.TextMatrix(llRow, CPMCXFCODE) = tlCxf.lCode
            slSQLQuery = "Update CXF_Hdr_Ln_Comments Set "
            slSQLQuery = slSQLQuery & "cxfRemoteID = " & tgUrf(0).iRemoteUserID & ", "
            slSQLQuery = slSQLQuery & "cxfAutoCode = " & tlCxf.lCode & " "
            slSQLQuery = slSQLQuery & "Where cxfCode = " & tlCxf.lCode
            llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            If llRet <> 0 Then
                mSaveCPMComment = 0
                Exit Function
            End If
        Else
            mSaveCPMComment = 0
            Exit Function
        End If
    Else 'Old record-Update
        If Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMCOMMENT)) <> "" Then
            slSQLQuery = "Update CXF_Hdr_Ln_Comments Set "
            slSQLQuery = slSQLQuery & "cxfCode = " & tlCxf.lCode & ", "
            slSQLQuery = slSQLQuery & "cxfComType = '" & gFixQuote(tlCxf.sComType) & "', "
            slSQLQuery = slSQLQuery & "cxfShProp = '" & gFixQuote(tlCxf.sShProp) & "', "
            slSQLQuery = slSQLQuery & "cxfShOrder = '" & gFixQuote(tlCxf.sShOrder) & "', "
            slSQLQuery = slSQLQuery & "cxfShSpot = '" & gFixQuote(tlCxf.sShSpot) & "', "
            slSQLQuery = slSQLQuery & "cxfShInv = '" & gFixQuote(tlCxf.sShInv) & "', "
            slSQLQuery = slSQLQuery & "cxfRemoteID = " & tlCxf.iRemoteID & ", "
            slSQLQuery = slSQLQuery & "cxfAutoCode = " & tlCxf.lAutoCode & ", "
            slSQLQuery = slSQLQuery & "cxfSyncDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
            slSQLQuery = slSQLQuery & "cxfSyncTime = '" & Format$(gNow(), sgSQLTimeForm) & "', "
            slSQLQuery = slSQLQuery & "cxfSourceID = " & tlCxf.iSourceID & ", "
            slSQLQuery = slSQLQuery & "cxfShInsertion = '" & gFixQuote(tlCxf.sShInsertion) & "', "
            slSQLQuery = slSQLQuery & "cxfUnused = '" & gFixQuote(Trim$(tlCxf.sUnused)) & "', "
            slSQLQuery = slSQLQuery & "cxfComment = '" & gFixQuote(Trim$(tlCxf.sComment)) & "'"
            slSQLQuery = slSQLQuery & " Where cxfCode = " & tlCxf.lCode
            llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            If llRet <> 0 Then
                mSaveCPMComment = 0
                Exit Function
            End If
            mSaveCPMComment = tlCxf.lCode
        Else
            slSQLQuery = "Delete From CXF_Hdr_Ln_Comments Where cxfcode = " & tlCxf.lCode
            llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            mSaveCPMComment = 0
        End If
    End If
End Function

Public Sub mComputeTotals()
    Dim llRow As Long
    Dim slImpressions As String
    Dim slTotalCost As String
    Dim slTotalImpressions As String
    Dim slCPM As String
    Dim slCost As String
    Dim slTotalCPM As String
    Dim bmFlatRateExist As Boolean
    
    slTotalImpressions = "0"
    slTotalCost = "0.00"
    bmFlatRateExist = False
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
    
            slImpressions = Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS)
            If Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) <> "P" Then
                slTotalImpressions = gAddStr(slTotalImpressions, slImpressions)
            End If
            gFormatStr slImpressions, FMTLEAVEBLANK, 0, slImpressions
            Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS) = slImpressions
        
            slCPM = Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE)
            gFormatStr slCPM, FMTLEAVEBLANK, 2, slCPM
            Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE) = slCPM
        
            slCost = Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST)
            If Contract.grdCPMBuy.TextMatrix(llRow, CPMLNTYPE) <> "P" Then
                slTotalCost = gAddStr(slTotalCost, slCost)
            End If
            gFormatStr slCost, FMTLEAVEBLANK, 2, slCost
            Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST) = slCost
        End If
        If Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMPRICETYPE)) = "Flat Rate" Then
            bmFlatRateExist = True
        End If
    Next llRow
    slTotalCPM = gDivStr(gMulStr(slTotalCost, "1000.00"), slTotalImpressions)
    gFormatStr slTotalImpressions, FMTLEAVEBLANK, 0, slTotalImpressions
    Contract.edcLineNo(2).Text = slTotalImpressions
    If bmFlatRateExist Then
        Contract.edcLineNo(3).Visible = False
    Else
        gFormatStr slTotalCPM, FMTLEAVEBLANK, 2, slTotalCPM
        Contract.edcLineNo(3).Text = slTotalCPM
        If Contract.grdCPMBuy.Visible Then
            Contract.edcLineNo(3).Visible = True
        End If
    End If
    gFormatStr slTotalCost, FMTLEAVEBLANK, 2, slTotalCost
    Contract.edcLineNo(4).Text = slTotalCost
    'Contract.mSetCpmInfo
End Sub

Public Sub mPopCPMPosition(llRow As Long)
    Dim ilVff As Integer
    Dim slSQLQuery As String
    
    'Contract.lbcPrg(4).AddItem "1"
    'Contract.lbcPrg(4).AddItem "2"
    'Contract.lbcPrg(4).AddItem "3"
    'Contract.lbcPrg(4).AddItem "4"
    'Contract.lbcPrg(4).AddItem "5"
    ilVff = gBinarySearchVff(Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE)))
    If ilVff <> -1 Then
        If (imPositionAvfCode <> tgVff(ilVff).iAvfCode) Or (Contract.lbcPrg(4).ListCount <= 0) Then
            imPositionAvfCode = tgVff(ilVff).iAvfCode
            Contract.lbcPrg(4).Clear
            Contract.lbcPrg(4).AddItem "[No Position]"
            slSQLQuery = "Select mnfName, mnfCode from MNF_Multi_Names"
            slSQLQuery = slSQLQuery & " Where mnfType = '6'"
            slSQLQuery = slSQLQuery & " And mnfGroupNo = " & tgVff(ilVff).iAvfCode
            Set mnf_rst = gSQLSelectCall(slSQLQuery)
            Do While Not mnf_rst.EOF
                Contract.lbcPrg(4).AddItem Trim$(mnf_rst!mnfName)
                Contract.lbcPrg(4).ItemData(Contract.lbcPrg(4).NewIndex) = mnf_rst!mnfCode
                mnf_rst.MoveNext
            Loop
        End If
    Else
        Contract.lbcPrg(4).Clear
    End If

End Sub

Public Sub mPopCPMPriority()
    Dim ilVff As Integer
    Dim ilLoop As Integer
    Dim slSQLQuery As String
    

    'Contract.lbcPrg(5).Clear
    ''Contract.lbcPrg(5).AddItem "[None]"
    'Contract.lbcPrg(5).AddItem "1"
    'Contract.lbcPrg(5).AddItem "2"
    'Contract.lbcPrg(5).AddItem "3"
    'Contract.lbcPrg(5).AddItem "4"
    'Contract.lbcPrg(5).AddItem "5"
    ilVff = gBinarySearchVff(Val(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMVEFCODE)))
    If ilVff <> -1 Then
        If (imPriorityAvfCode <> tgVff(ilVff).iAvfCode) Or (Contract.lbcPrg(5).ListCount <= 0) Then
            imPriorityAvfCode = tgVff(ilVff).iAvfCode
            Contract.lbcPrg(5).Clear
            slSQLQuery = "Select avfPriorityStart, avfPriorityEnd from avf_AdVendor"
            slSQLQuery = slSQLQuery & " Where avfCode = " & tgVff(ilVff).iAvfCode
            Set avf_rst = gSQLSelectCall(slSQLQuery)
            If Not avf_rst.EOF Then
                For ilLoop = avf_rst!avfprioritystart To avf_rst!avfpriorityend Step 1
                    Contract.lbcPrg(5).AddItem ilLoop
                    Contract.lbcPrg(5).ItemData(Contract.lbcPrg(5).NewIndex) = ilLoop
                Next ilLoop
            End If
        End If
    Else
        Contract.lbcPrg(5).Clear
    End If
    
End Sub

Public Function mCPMAllowedDate(llRow As Long, Optional blUseBilling = False) As String
    '10719 added optional parameter
    Dim ibf_rst As ADODB.Recordset
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim slDate As String
    Dim slSQLQuery As String
    Dim slRet As String
    
    'mCPMAllowedDate = smCPMLLD
    '10719
    slRet = smCPMLLD
    If blUseBilling Then
        slRet = smCPMLBD
    End If
    ilYear = 0
    If Contract.grdCPMBuy.TextMatrix(llRow, CPMBILLED) = "Y" Then
        slSQLQuery = "Select * From ibf_Impression_Bill Where ibfCntrNo = " & tgChfCntr.lCntrNo
        slSQLQuery = slSQLQuery & " And ibfPodCPMID = " & Contract.grdCPMBuy.TextMatrix(llRow, CPMID)
        slSQLQuery = slSQLQuery & " And ibfBilled = 'Y'"
        Set ibf_rst = gSQLSelectCall(slSQLQuery)
        Do While Not ibf_rst.EOF
            If ilYear = 0 Then
                ilYear = ibf_rst!ibfBillYear
                ilMonth = ibf_rst!ibfBillMonth
            Else
                If (ilYear = ibf_rst!ibfBillYear) And (ilMonth < ibf_rst!ibfBillMonth) Then
                    ilMonth = ibf_rst!ibfBillMonth
                ElseIf (ilYear < ibf_rst!ibfBillYear) Then
                    ilYear = ibf_rst!ibfBillYear
                    ilMonth = ibf_rst!ibfBillMonth
                End If
            End If
            ibf_rst.MoveNext
        Loop
        If ilYear > 0 Then
            slDate = ilMonth & "/" & "15" & "/" & ilYear
            If tgChfCntr.sBillCycle = "C" Then
                slDate = gObtainEndCal(slDate)
            ElseIf tgChfCntr.sBillCycle = "S" Then
                slDate = gObtainEndStd(slDate)
            End If
            If gDateValue(slDate) > gDateValue(slRet) Then
                '10719 don't add date if blUseBilling
                If blUseBilling Then
                    slRet = slDate
                Else
                    slRet = DateAdd("d", 1, slDate)
                End If
            End If
        End If
        ibf_rst.Close
    End If
    mCPMAllowedDate = slRet
End Function

Public Sub mClearPodTartget()
    Dim llRow As Long
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        Contract.grdCPMBuy.TextMatrix(llRow, CPMPODTARGET) = ""
    Next llRow
    mPopPodTarget
End Sub
Private Function mDigitalLastRowWithVehicle() As Long
    '10859
    Dim llRow As Long
    
    llRow = Contract.grdCPMBuy.rows
    Do
        llRow = llRow - 1
    Loop While Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) = ""
    mDigitalLastRowWithVehicle = llRow
End Function
Private Sub mDigitalAddNewLine(llNewRow As Long)
    Dim llCurrentRow As Long
    Dim ilCol As Integer
    
    With Contract.grdCPMBuy
        llCurrentRow = .Row
        .AddItem ""
        .rowHeight(llNewRow) = fgBoxGridH
        .Row = llNewRow
        For ilCol = 0 To .cols - 1 Step 1
            .ColAlignment(ilCol) = flexAlignLeftCenter
        Next ilCol
        .Col = CPMID
        .CellBackColor = LIGHTYELLOW
        .Col = CPMEXPAND
        .CellFontName = "Monotype Sorts"
        .CellFontSize = 7
        .CellForeColor = BLUE
        .Col = CPMVEHICLE
        .CellFontItalic = True
        .TextMatrix(llNewRow, CPMPCFCODE) = "0"
        .TextMatrix(llNewRow, CPMSORT) = ""
        .TextMatrix(llNewRow, CPMCXFCODE) = "0"
        .TextMatrix(llNewRow, CPMCOMMSHOWON) = "NNDND"   'Proposal;Order;Spots/Copy;Invoice;Insertion
        .TextMatrix(llNewRow, CPMSHOWROW) = "True"
        .Row = llCurrentRow
    End With
End Sub
Private Sub mDigitalHideScrollAsNeeded()
    With Contract.grdCPMBuy
        If Not .RowIsVisible(.Row) Or (.Row - (.TopRow - .FixedRows) >= imInitNoRows) Then
            imCPMIgnoreScroll = True
            .TopRow = .TopRow + 1
        End If
    End With
End Sub
Private Sub mDigitalPrepNewLine()
    imCPMFromArrow = True
    mSetBGColorWhite False
    Contract.pbcLineArrow.Move Contract.grdCPMBuy.Left - Contract.pbcLineArrow.Width - 30, Contract.grdCPMBuy.Top + Contract.grdCPMBuy.RowPos(Contract.grdCPMBuy.Row) + (Contract.grdCPMBuy.rowHeight(Contract.grdCPMBuy.Row) - Contract.pbcLineArrow.height) / 2
    Contract.pbcLineArrow.Visible = True
    Contract.pbcLineArrow.SetFocus
End Sub
Private Sub mDigitalTabToNextLine()
    '10859
    Dim llLastRowWithVehicle As Long
    
    llLastRowWithVehicle = mDigitalLastRowWithVehicle()
    With Contract.grdCPMBuy
    If .Row < llLastRowWithVehicle Then
        lmCPMTopRow = -1
        .Row = .Row + 1
        mDigitalHideScrollAsNeeded
        .Col = CPMVEHICLE
        If Trim$(.TextMatrix(.Row, CPMVEHICLE)) <> "" Then
            mCPMEnableBox
        Else
            '? maybe this catches if a line had the vehicle erased?
             .Row = llLastRowWithVehicle
             mDigitalPrepNewLine
        End If
    Else
        'new line
        lmCPMTopRow = -1
        If .Row + 1 >= .rows Then
            mDigitalAddNewLine .Row + 1
            'add an additional row
            If .Row + 2 = .rows Then
                mDigitalAddNewLine .Row + 2
            End If
        End If
        .Row = .Row + 1
        mDigitalHideScrollAsNeeded
        mDigitalPrepNewLine
    End If
    End With
End Sub
Private Sub mDigitalSetValueToOriginal(llSvRow As Long, llSvCol As Long, OriginalIndex As Integer)
    Beep
    Contract.grdCPMBuy.Row = llSvRow
    Contract.grdCPMBuy.Col = llSvCol
    Contract.grdCPMBuy.TextMatrix(llSvRow, llSvCol) = Contract.grdCPMBuy.TextMatrix(llSvRow, OriginalIndex)
    mCPMEnableBox
End Sub
Private Function mDigitalTestImpressionChangeIfBilled(llRow As Long, slNewImpression As String) As Boolean
    Dim blRet As Boolean
    Dim llNewImp As Long
    Dim llOriginalImp As Long
    Dim dlInvoicedAmount As Double
    Dim dlNewCost As Double
    Dim slNewCost As String
  '  Dim slBilledDate As String

On Error GoTo ERRBOX
    blRet = True
    If Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMPRICETYPE)) <> "Flat Rate" Then
        'slBilledDate = mCPMAllowedDate(llRow, True)
        'end date is out in future.
        '10738 remove date test
       ' If gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)) >= gDateValue(slBilledDate) Then
            slNewImpression = Trim$(slNewImpression)
            llNewImp = CLng(slNewImpression)
            llOriginalImp = CLng(Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGIMPRESSIONS))
            'Impression has changed and is less.  This could be a problem, so keep testing
            If llNewImp < llOriginalImp Then
                slNewCost = gDivStr(gMulStr(slNewImpression, Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE)), "1000.00")
                dlNewCost = CDbl(slNewCost)
                dlInvoicedAmount = mDigitalInvoiced(tgChfCntr.lCntrNo, Contract.grdCPMBuy.TextMatrix(llRow, CPMID))
                If dlNewCost < dlInvoicedAmount Then
                    blRet = False
                End If
            End If
      '  Else
      '      blRet = False
      '  End If
    End If
    mDigitalTestImpressionChangeIfBilled = blRet
    Exit Function
ERRBOX:
    mDigitalTestImpressionChangeIfBilled = False
End Function
Private Function mDigitalInvoiced(llContractNo As Long, llLineID As Long) As Double
    Dim slSQLQuery As String
    Dim rst As ADODB.Recordset
    'or decimal? currency?
    Dim dlTotal As Double
    ' sum(rvfGross) As RvfTotalGross
    '.lRvfTotalGross = IIF(IsNull(rs("RvfTotalGross")), .lRvfTotalGross, rs!rvftotalgross)
    'gPDNToLong
    dlTotal = 0
    slSQLQuery = "Select sum(rvfGross) As TotalGross from RVF_Receivables WHERE rvfcntrno = " & llContractNo & " and rvfpcfcode in ( Select pcfcode from pcf_Pod_CPM_Cntr where pcfPodCPMID = " & llLineID & ")"
    Set rst = gSQLSelectCall(slSQLQuery)
    If Not rst.EOF Then
        If Not IsNull(rst!totalgross) Then
            dlTotal = rst!totalgross
        End If
    End If
    slSQLQuery = "Select sum(phfGross) As TotalGross from PHF_Payment_History WHERE phfcntrno = " & llContractNo & " and phfpcfcode in ( Select pcfcode from pcf_Pod_CPM_Cntr where pcfPodCPMID = " & llLineID & ")"
    Set rst = gSQLSelectCall(slSQLQuery)
    If Not rst.EOF Then
        If Not IsNull(rst!totalgross) Then
            dlTotal = dlTotal + rst!totalgross
        End If
    End If
    mDigitalInvoiced = dlTotal
End Function
Private Function mPopCPMLength(llCPMEnableRow As Long) As Integer
    '10879 and return default length vpfSDLen
    Dim ilVpf As Integer
    Dim ilLoop As Integer
    Dim ilVef As Integer
    Dim ilNewLen As Integer
    Dim ilCurrentLen As Integer
    Dim blFound As Boolean
    Dim ilRet As Integer
    
    ilRet = 0
    Contract.lbcPrg(6).Clear
    Contract.lbcPrg(6).AddItem "0"
    'Contract.lbcPrg(6).AddItem "[None]"
    blFound = False
    ilCurrentLen = Val(Contract.grdCPMBuy.TextMatrix(llCPMEnableRow, CPMLENGTH))
    ilVef = Val(Contract.grdCPMBuy.TextMatrix(llCPMEnableRow, CPMVEFCODE))
    If ilVef > 0 Then
        ilVpf = gBinarySearchVpf(ilVef)
        If ilVpf <> -1 Then
            For ilLoop = 0 To 9 Step 1
                ilNewLen = tgVpf(ilVpf).iSLen(ilLoop)
                If ilNewLen > 0 And ilNewLen <= 999 Then
                     Contract.lbcPrg(6).AddItem ilNewLen
                     If ilNewLen = ilCurrentLen Then
                        blFound = True
                     End If
                End If
            Next ilLoop
            ilRet = tgVpf(ilVpf).iSDLen
        End If
    End If
    If blFound = False And ilCurrentLen > 0 Then
        Contract.lbcPrg(6).AddItem ilCurrentLen
    End If
    mPopCPMLength = ilRet
 'Contract.lbcPrg(6).ItemData(Contract.lbcPrg(6).NewIndex) = ilLoop
End Function
Public Function mGetContractDocuments(llCntrNo As Long) As Integer
    '10910 return count
    Dim ilRet As Integer
    Dim slSQLQuery As String
    Dim rst As Recordset
    Dim ilUpper As Integer
    
    ilRet = 0
    If llCntrNo > 0 Then
        ReDim tgContractDocuments(0)
        slSQLQuery = "select * from ACF_Attach_Contract where acfcntrno = " & llCntrNo
        Set rst = gSQLSelectCall(slSQLQuery)
        Do While Not rst.EOF
            ilRet = ilRet + 1
            ilUpper = UBound(tgContractDocuments)
            tgContractDocuments(ilUpper).bChanged = False
            tgContractDocuments(ilUpper).lCntrNo = llCntrNo
            tgContractDocuments(ilUpper).lCode = rst!acfCode
            tgContractDocuments(ilUpper).sFileName = rst!acfFileName
            tgContractDocuments(ilUpper).sDescription = rst!acfDescription
            tgContractDocuments(ilUpper).iUrfCode = rst!acfUrfCode
            tgContractDocuments(ilUpper).bDeleted = False
            If rst!acfDeleted = "Y" Then
                tgContractDocuments(ilUpper).bDeleted = True
            End If
            tgContractDocuments(ilUpper).sEnteredDate = rst!acfEnteredDate
            tgContractDocuments(ilUpper).sEnteredTime = rst!acfEnteredTime
            tgContractDocuments(ilUpper).sDeletedDate = rst!acfDeletedDate
            tgContractDocuments(ilUpper).sDeletedTime = rst!acfDeletedTime
            ReDim Preserve tgContractDocuments(0 To ilUpper + 1) As ACFList
            rst.MoveNext
        Loop
    End If
    mGetContractDocuments = ilRet
End Function

Public Function mSaveContractDocuments(llCntrNo As Long) As Boolean
    '10910
    Dim blRet As Boolean
    Dim slSQLQuery As String
    Dim rst As Recordset
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim llCount As Long
    Dim slDelete As String
    Dim fs As FileSystemObject
    Dim blTestedMainFolder As Boolean
    Dim slWriteToPath As String
    Dim slWriteFromFull As String
    Dim slWriteToFull As String
    Dim slNameForDatabase As String
    
    blTestedMainFolder = False
    blRet = True
    Set fs = New FileSystemObject
    ilUpper = UBound(tgContractDocuments)
    For ilLoop = 0 To ilUpper - 1 Step 1
        With tgContractDocuments(ilLoop)
            If .bChanged Then
                'have a new document
                If Len(.sTempPath) > 0 Then
                    'Add to folder
                    If Not blTestedMainFolder Then
                        slWriteToPath = mContractDocumentFolderPath(llCntrNo, fs)
                        'issue!
                        If slWriteToPath = "" Then
                            blRet = False
                            gMsgBox "Attachments not saved-issue building contract folders"
                            mContractDocumentWriteAmount True, True
                            Exit Function
                        End If
                        blTestedMainFolder = True
                    End If
                    slWriteFromFull = .sTempPath 'fs.BuildPath(.sTempPath, .sFileName)
                    slWriteToFull = mContractDocumentFileName(slWriteToPath, .sFileName, fs)
                    slNameForDatabase = fs.GetFileName(slWriteToFull)
    On Error GoTo ERRFILE
                    fs.CopyFile slWriteFromFull, slWriteToFull, False
                    SetAttr slWriteToFull, vbReadOnly
    On Error Resume Next
                    'hard-coded 'N' for delete and '' for unused
                    slSQLQuery = "INSERT into ACF_Attach_Contract (acfCntrNo,acfFileName,acfUrfCode,acfDescription,acfEnteredDate,acfEnteredTime,acfDeleted,acfDeletedDate,acfDeletedTime,acfUnused) Values (" & llCntrNo & ",'" & slNameForDatabase & "'," & .iUrfCode & ",'" & .sDescription & "','" & Format(.sEnteredDate, sgSQLDateForm) & "','" & Format(.sEnteredTime, sgSQLTimeForm) & "','N','" & Format(.sDeletedDate, sgSQLDateForm) & " ','" & Format(.sDeletedTime, sgSQLTimeForm) & "','')"
                'do I need this test?
                ElseIf tgContractDocuments(ilLoop).lCode > 0 Then
                    'no change to contract # or code
                    If tgContractDocuments(ilLoop).bDeleted Then
                        slDelete = "Y"
                    Else
                        slDelete = "N"
                    End If
                    slSQLQuery = "UPDATE ACF_Attach_Contract set acfFileName='" & Trim$(.sFileName) & "',acfUrfCode=" & .iUrfCode & ", acfDescription='" & Trim$(.sDescription) & "',acfEnteredDate='" & Format(.sEnteredDate, sgSQLDateForm) & "',acfEnteredTime='" & Format(.sEnteredTime, sgSQLTimeForm) & "',acfDeletedDate='" & Format(.sDeletedDate, sgSQLDateForm) & "',acfDeletedTime='" & Format(.sDeletedTime, sgSQLTimeForm) & "',acfDeleted='" & slDelete & "' WHERE acfcode =" & .lCode
                Else
                    'no temp path and no code?
                    blRet = False
                    gMsgBox "Attachments not saved-missing acfCode"
                    Exit Function
                End If
                If gSQLAndReturn(slSQLQuery, False, llCount) <> 0 Then
                    blRet = False
                    gHandleError "TrafficErrors.txt", "Contracts-mSaveDocuments"
                    Exit Function
                End If
                'this way, if there is an error, I can write out the # that were saved
                .bChanged = False
            End If
       End With
    Next ilLoop
    mSaveContractDocuments = blRet
    Exit Function
ERRFILE:
    mSaveContractDocuments = False
    gMsgBox "Attachments not saved-couldn't move " & slWriteFromFull & " to " & slWriteToFull
    mContractDocumentWriteAmount True, True
End Function
Public Function mContractDocumentChanged() As Boolean
    Dim blRet As Boolean
    Dim ilLoop As Integer
    
    blRet = False
    For ilLoop = 0 To UBound(tgContractDocuments) - 1 Step 1
        If tgContractDocuments(ilLoop).bChanged Then
            blRet = True
            Exit For
        End If
    Next ilLoop
    mContractDocumentChanged = blRet
End Function
Private Function mContractDocumentFolderPath(llCntrNo As Long, fs As FileSystemObject) As String
    Dim slRet As String
 On Error GoTo PATHISSUE
    If Not fs.FolderExists(sgContractAttachmentPath) Then
        fs.CreateFolder (sgContractAttachmentPath)
    End If
    slRet = fs.BuildPath(sgContractAttachmentPath, llCntrNo)
    If Not fs.FolderExists(slRet) Then
        fs.CreateFolder (slRet)
    End If
    mContractDocumentFolderPath = slRet
    Exit Function
PATHISSUE:
    mContractDocumentFolderPath = ""
End Function
Private Function mContractDocumentFileName(slPath As String, slName As String, fs As FileSystemObject) As String
    Dim slRet As String
    Dim ilCounter As Integer
    Dim slTestFile As String
    Dim slExt As String
    Dim slPartName As String
    
    slRet = fs.BuildPath(slPath, slName)
    
    If fs.FILEEXISTS(slRet) Then
        slExt = fs.GetExtensionName(slName)
        slPartName = fs.GetBaseName(slName)
        For ilCounter = 1 To 99 Step 1
            slTestFile = fs.BuildPath(slPath, slPartName & "_" & ilCounter & "." & slExt)
            If Not fs.FILEEXISTS(slTestFile) Then
                slRet = slTestFile
                Exit For
            End If
        Next ilCounter
    End If
    mContractDocumentFileName = slRet
End Function

Public Sub mContractDocumentWriteAmount(Optional blPaint As Boolean = True, Optional blIsError As Boolean = False)
    '10910
    'blPaint? sometimes we only want to gSetShow
    'isError? We will only show those that we tried to change the bchanged flag to false while saving
    Dim ilNumberToShow As Integer
    Dim ilIndex As Integer
    
    ilNumberToShow = 0
    For ilIndex = 0 To UBound(tgContractDocuments) - 1 Step 1
        If tgContractDocuments(ilIndex).bDeleted = False Then
            If blIsError Then
                If tgContractDocuments(ilIndex).bChanged = False Then
                    ilNumberToShow = ilNumberToShow + 1
                End If
            Else
                ilNumberToShow = ilNumberToShow + 1
            End If
        End If
    Next
    gSetShow Contract.pbcHeader, str$(ilNumberToShow), tmHdCtrls(ATTACHMENTINDEX)
    If blPaint Then
        'cover so paint will write
        gPaintArea Contract.pbcHeader, tmHdCtrls(ATTACHMENTINDEX).fBoxX + 15, tmHdCtrls(ATTACHMENTINDEX).fBoxY + 15, tmHdCtrls(ATTACHMENTINDEX).fBoxW - 30, tmHdCtrls(ATTACHMENTINDEX).fBoxH - 30, LIGHTBLUE
        'now will write
        Contract.pbcHeader_Paint
    End If
End Sub
Public Sub mVendorWriteStatus()
    '10967-show
    Dim slStatus As String
    
    Select Case tgChfCntr.iAdServDlvyStatus
        Case DeliveryStatus_NA
            slStatus = "N/A"
        Case DeliveryStatus_NotPushed
            slStatus = "Not Pushed"
        Case DeliveryStatus_Pushed
            slStatus = "Pushed"
        Case DeliveryStatus_Partial
            slStatus = "Partial"
        Case DeliveryStatus_IssueEncountered
            slStatus = "Issue"
        Case DeliveryStatus_RequiresRepush
            slStatus = "Repush Rqd"
    End Select
    gSetShow Contract.pbcHeader, slStatus, tmHdCtrls(VENDORINDEX)
    'cover so paint will write
    gPaintArea Contract.pbcHeader, tmHdCtrls(VENDORINDEX).fBoxX + 15, tmHdCtrls(VENDORINDEX).fBoxY + 15, tmHdCtrls(VENDORINDEX).fBoxW - 30, tmHdCtrls(VENDORINDEX).fBoxH - 30, LIGHTBLUE
    'now will write
    Contract.pbcHeader_Paint
End Sub
Private Function mVehicleSentToVendorAndCannotAlter(llRow As Long) As Boolean
    Dim blRet As Boolean
    Dim llNow As Long
    Dim llStart As Long
    
    blRet = False
    'sent to vendor?
    If tgChfCntr.iAdServDlvyStatus > DeliveryStatus_NotPushed And Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMDELIVERYSTATUS)) > DeliveryStatus_NotPushed Then
        llNow = gDateValue(Format$(gNow(), "m/d/yy"))
        llStart = gDateValue(Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE))
        ' if today is after start date, user can't change the vehicle.
        If llNow >= llStart Then
            blRet = True
        End If
    End If
    mVehicleSentToVendorAndCannotAlter = blRet
End Function
'10929
Public Function mDigitalPackageLengthTest(slLen As String) As Boolean
    'modeled off of Native1 mLnSetShowSub
    Dim blRet As Boolean
    Dim ilHiddenLengths As Integer
    Dim ilNewlength As Integer
    Dim ilLineIDForPackage As Integer
    Dim ilOldLength As Integer
    
    blRet = True
    If IsNumeric(slLen) Then
        ilNewlength = Val(slLen)
    ' new line? CURRENTLY SKIPPING because didn't understand it
        ' package?
        If Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLNTYPE) = "P" Then
            ilLineIDForPackage = Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMID)
            'compare to hidden lengths
            ilHiddenLengths = mDigitalPackageHiddenLenMatch(ilNewlength, ilLineIDForPackage)
            If ilHiddenLengths = HIDDENLENMATCH Then
            ElseIf ilHiddenLengths = HIDDENLENNOTALLOWED Then
                blRet = False
                MsgBox "Package Length change not permitted as the length " & ilNewlength & " does not exist in all hidden lines", vbExclamation + vbOKOnly, "Change Disallowed"
            Else
                If Trim$(Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLENGTH)) <> "" Then
                    sgGenMsg = "Package Spot Length has been changed from " & Contract.grdCPMBuy.TextMatrix(lmCPMEnableRow, CPMLENGTH) & " to " & ilNewlength
                Else
                    sgGenMsg = "Spot Length set to " & ilNewlength
                End If
                sgCMCTitle(0) = "Change Hidden"
                sgCMCTitle(1) = "Cancel Package Change"
                sgCMCTitle(2) = ""
                sgCMCTitle(3) = ""
                If ilHiddenLengths = HIDDENLENCONSISTENTBUTNOMATCH Then
                    sgGenMsg = sgGenMsg & ".  If the spot prices are correct, then press Change Hidden.  If the spot prices are incorrect, then press Cancel and delete entire package and re-enter"
                Else 'HIDDENLENNOTCONSISTENT. They can choose to change or not
                    sgGenMsg = sgGenMsg & ".  If the spot prices are correct, then press either Change Hidden or Leave Hidden.  If the spot prices are incorrect, then press Cancel and delete entire package and re-enter"
                    sgCMCTitle(1) = "Leave Hidden"
                    sgCMCTitle(2) = "Cancel Package Change"
                End If
                GenMsg.Show vbModal
                'change hidden lines
                If igAnsCMC = 0 Then
                    mDigitalPackageLengthToHidden ilNewlength, ilLineIDForPackage
                'chose to cancel
                ElseIf (ilHiddenLengths = HIDDENLENCONSISTENTBUTNOMATCH And igAnsCMC = 1) Or (ilHiddenLengths = HIDDENLENNOTCONSISTENT And igAnsCMC = 2) Then
                    blRet = False
                End If
            End If
        End If 'package
    Else 'number
        blRet = False
    End If
    mDigitalPackageLengthTest = blRet
End Function
Public Function mDigitalPackageHiddenLenMatch(ilNewlength As Integer, ilPackageID As Integer) As Integer
    'matching Dick's mPackageHiddenLenMatch
        'Return: 0 = All lengths match and length exist in all hidden vehicles  HIDDENLENMATCH = 0 hidden lines already matched the new package length
        '        1 = Not all lengths match the package length                   HIDDENLENCONSISTENTBUTNOMATCH = 1 all hidden line lengths were the same, but that didn't match the new package length
        '        2 = Length varies across the hidden lines                      HIDDENLENNOTCONSISTENT = 2 hidden line lengths were not consistent.  Obviously, they don't all match the new package length
        '        3 = Length does not exist with the hidden vehicle              HIDDENLENNOTALLOWED = 3 not only did a hidden line not match, it's not allowed to be changed to the new value
    Dim ilRet As Integer
    Dim llRow As Long
    Dim ilVef As Integer
    Dim blMatch As Boolean
    Dim ilHiddenConsistentLength As Integer
    Dim ilHiddenLength As Integer
 
    ilRet = HIDDENLENMATCH
    ilHiddenConsistentLength = -1
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
            If ilPackageID = Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID)) Then
                ilHiddenLength = Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMLENGTH))
                If ilHiddenLength <> ilNewlength Then
                    ilVef = Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE)
                    If Not mDigitalCompareAllowedLength(ilNewlength, ilVef) Then
                        ilRet = HIDDENLENNOTALLOWED
                        Exit For
                    ElseIf ilRet = HIDDENLENMATCH Then
                        ilRet = HIDDENLENCONSISTENTBUTNOMATCH
                    End If
                End If
                'test constancy. 'not allowed' can't be here so doesn't need to be tested.  If already inconsistent we don't care
                If ilRet <> HIDDENLENNOTCONSISTENT Then
                    If ilHiddenConsistentLength = -1 Then
                        ilHiddenConsistentLength = ilHiddenLength
                    ElseIf ilHiddenConsistentLength <> ilHiddenLength Then
                        ilRet = HIDDENLENNOTCONSISTENT
                    End If
                End If
            End If
        Else
            Exit For
        End If
    Next llRow
    mDigitalPackageHiddenLenMatch = ilRet
End Function
Public Sub mDigitalPackageLengthToHidden(ilNewlength As Integer, ilPackageID As Integer)
    Dim llRow As Long
         
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
            If ilPackageID = Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID)) Then
                If Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMLENGTH)) <> ilNewlength Then
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMLENGTH) = ilNewlength
                End If
            End If
        Else
            Exit For
        End If
    Next llRow
End Sub
Public Function mDigitalCompareAllowedLength(ilLengthToCompare As Integer, ilVef As Integer) As Boolean
    Dim ilVpf As Integer
    Dim ilLoop As Integer
    Dim blRet As Boolean
    
    blRet = False
    If ilVef > 0 Then
        ilVpf = gBinarySearchVpf(ilVef)
        If ilVpf <> -1 Then
            For ilLoop = 0 To 9 Step 1
                If tgVpf(ilVpf).iSLen(ilLoop) = ilLengthToCompare Then
                   blRet = True
                   Exit For
                End If
            Next ilLoop
        End If
    End If
    mDigitalCompareAllowedLength = blRet
End Function

