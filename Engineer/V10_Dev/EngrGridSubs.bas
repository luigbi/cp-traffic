Attribute VB_Name = "EngrGridSubs"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: EngrGrid.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the API declarations
Option Explicit

Public Sub gGrid_AlignAllColsLeft(grdCtrl As MSHFlexGrid)
    Dim ilCol As Integer
    
    For ilCol = 0 To grdCtrl.Cols - 1 Step 1
        grdCtrl.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol

End Sub

Public Sub gGrid_FillWithRows(grdCtrl As MSHFlexGrid)
    Dim llRows As Long
    Dim llCols As Long
    Dim llFillNoRow As Long
    Dim llHeight As Long
    
    'llFillNoRow = grdCtrl.Height \ grdCtrl.RowHeight(grdCtrl.FixedRows) - grdCtrl.FixedRows
    llFillNoRow = 0
    llHeight = grdCtrl.FixedRows * grdCtrl.RowHeight(grdCtrl.FixedRows) + 15
    Do
        llFillNoRow = llFillNoRow + 1
        llHeight = llHeight + grdCtrl.RowHeight(grdCtrl.FixedRows)
    Loop While llHeight + grdCtrl.RowHeight(grdCtrl.FixedRows) <= grdCtrl.Height
    For llRows = grdCtrl.FixedRows + 1 To grdCtrl.FixedRows + llFillNoRow Step 1
        Do While llRows > grdCtrl.Rows
            grdCtrl.AddItem ""
            For llCols = 0 To grdCtrl.Cols - 1 Step 1
                grdCtrl.TextMatrix(llRows - 1, llCols) = ""
            Next llCols
        Loop
    Next llRows
End Sub

Public Function gGrid_DetermineRowCol(grdCtrl As MSHFlexGrid, x As Single, y As Single) As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim llColLeftPos As Long
    
    For llRow = grdCtrl.FixedRows To grdCtrl.Rows - 1 Step 1
        If grdCtrl.RowIsVisible(llRow) Then
            If (y >= grdCtrl.RowPos(llRow)) And (y <= grdCtrl.RowPos(llRow) + grdCtrl.RowHeight(llRow) - 15) Then
                llColLeftPos = grdCtrl.ColPos(0)
                For llCol = 0 To grdCtrl.Cols - 1 Step 1
                    If grdCtrl.ColWidth(llCol) > 0 Then
                        If (x >= llColLeftPos) And (x <= llColLeftPos + grdCtrl.ColWidth(llCol)) Then
                            grdCtrl.Row = llRow
                            grdCtrl.Col = llCol
                            gGrid_DetermineRowCol = True
                            Exit Function
                        End If
                        llColLeftPos = llColLeftPos + grdCtrl.ColWidth(llCol)
                    End If
                Next llCol
            End If
        End If
    Next llRow
    gGrid_DetermineRowCol = False
    Exit Function
End Function

Public Sub gGrid_IntegralHeight(grdCtrl As MSHFlexGrid)
'    If grdCtrl.Rows > 0 Then
'        If (grdCtrl.Height - 15) Mod grdCtrl.RowHeight(grdCtrl.FixedRows) <> 0 Then
'            'grdHistory.Height = ((grdHistory.Height \ grdHistory.RowHeight(1)) + 1) * grdHistory.RowHeight(1) + 15
'            grdCtrl.Height = (grdCtrl.Height \ grdCtrl.RowHeight(grdCtrl.FixedRows)) * grdCtrl.RowHeight(grdCtrl.FixedRows) + 15
'        End If
'    End If
    Dim llHeight As Long
    Dim llRow As Long
    
    llHeight = 0
    If grdCtrl.FixedRows > 0 Then
        For llRow = 1 To grdCtrl.FixedRows Step 1
            llHeight = llHeight + grdCtrl.RowHeight(llRow - 1)
        Next llRow
        llHeight = llHeight + 15
    End If
    Do While llHeight + grdCtrl.RowHeight(grdCtrl.FixedRows) <= grdCtrl.Height
        llHeight = llHeight + grdCtrl.RowHeight(grdCtrl.FixedRows)
    Loop
    If grdCtrl.ScrollBars <> 1 And grdCtrl.ScrollBars <> 3 Then
        grdCtrl.Height = llHeight ' + 15
    Else
        grdCtrl.Height = llHeight + GRIDSCROLLHEIGHT '+ 15 + GRIDSCROLLHEIGHT
    End If
End Sub

Public Sub gGrid_Clear(grdCtrl As MSHFlexGrid, ilFillRows As Integer)
    
'
'   grdCtrl (I)-  Grid Control name
'   ilFillRows (I)- True=Fill Grid with blank rows; False=Only have one blank row
'
    Dim llRows As Long
    Dim llCols As Long
    Dim llFillNoRow As Long
    Dim llHeight As Long
    
    If ilFillRows Then
        ''llFillNoRow = grdCtrl.Height \ grdCtrl.RowHeight(grdCtrl.FixedRows) - 2
        'llFillNoRow = grdCtrl.Height \ grdCtrl.RowHeight(grdCtrl.FixedRows) - grdCtrl.FixedRows
        llFillNoRow = 0
        llHeight = grdCtrl.FixedRows * grdCtrl.RowHeight(grdCtrl.FixedRows) + 15
        Do
            llFillNoRow = llFillNoRow + 1
            llHeight = llHeight + grdCtrl.RowHeight(grdCtrl.FixedRows)
        Loop While llHeight + grdCtrl.RowHeight(grdCtrl.FixedRows) <= grdCtrl.Height
    Else
        llFillNoRow = 0
    End If
    llRows = grdCtrl.Rows
    Do While llRows > grdCtrl.FixedRows + llFillNoRow + 1
        grdCtrl.RemoveItem llRows - 1
        llRows = llRows - 1
    Loop
    If ilFillRows Then
        gGrid_FillWithRows grdCtrl
    Else
        llRows = grdCtrl.FixedRows
        If llRows >= grdCtrl.Rows Then
            Do While llRows >= grdCtrl.Rows
                grdCtrl.AddItem ""
            Loop
'        Else
'            llRows = grdCtrl.Rows
'            For llCols = 0 To grdCtrl.Cols - 1 Step 1
'                grdCtrl.TextMatrix(llRows, llCols) = ""
'            Next llCols
        End If
    End If
    llRows = grdCtrl.FixedRows
    Do While llRows < grdCtrl.Rows
        For llCols = 0 To grdCtrl.Cols - 1 Step 1
            grdCtrl.TextMatrix(llRows, llCols) = ""
            grdCtrl.Row = llRows
            grdCtrl.Col = llCols
            grdCtrl.BackColor = vbWhite
            grdCtrl.ForeColor = vbBlack
            grdCtrl.CellBackColor = vbWhite
            grdCtrl.CellForeColor = vbBlack
        Next llCols
        llRows = llRows + 1
    Loop
End Sub



Public Sub gGrid_SortByCol(grdCtrl As MSHFlexGrid, ilTestCol As Integer, ilSortCol As Integer, ilPrevSortCol As Integer, ilPrevSortDirection As Integer)
    Dim llEndRow As Long
    
    grdCtrl.Redraw = False
    grdCtrl.Col = ilSortCol
    grdCtrl.Row = grdCtrl.FixedRows
    llEndRow = grdCtrl.Rows - 1
    If grdCtrl.TextMatrix(llEndRow, ilTestCol) = "" Then
        Do
            llEndRow = llEndRow - 1
            If llEndRow <= grdCtrl.FixedRows Then
                Exit Do
            End If
        Loop While grdCtrl.TextMatrix(llEndRow, ilTestCol) = ""
    End If
    If llEndRow > grdCtrl.FixedRows Then
        grdCtrl.RowSel = llEndRow
        If ilPrevSortCol = grdCtrl.Col Then
            If ilPrevSortDirection = flexSortStringNoCaseAscending Then
                grdCtrl.Sort = flexSortStringNoCaseDescending
                ilPrevSortDirection = flexSortStringNoCaseDescending
            Else
                grdCtrl.Sort = flexSortStringNoCaseAscending
                ilPrevSortDirection = flexSortStringNoCaseAscending
            End If
        Else
            grdCtrl.Sort = flexSortStringNoCaseAscending 'flexSortStringNoCaseAscending
            ilPrevSortDirection = flexSortStringNoCaseAscending
        End If
    End If
    ilPrevSortCol = grdCtrl.Col
    grdCtrl.Row = grdCtrl.FixedRows
    grdCtrl.RowSel = grdCtrl.Row
    grdCtrl.Redraw = True
End Sub

Public Function gGrid_Search(grdCtrl As MSHFlexGrid, ilSearchCol As Integer, slSearchValue As String) As Long
    Dim llRow As Long
    Dim slStr As String
    Dim ilPos As Integer
    
    For llRow = grdCtrl.FixedRows To grdCtrl.Rows - 1 Step 1
        slStr = Trim$(grdCtrl.TextMatrix(llRow, ilSearchCol))
        If (slStr <> "") Then
            If StrComp(slStr, slSearchValue, vbTextCompare) = 0 Then
                grdCtrl.TopRow = grdCtrl.FixedRows
                grdCtrl.Row = llRow
                igGridIgnoreScroll = True
                Do While Not grdCtrl.RowIsVisible(grdCtrl.Row)
                    grdCtrl.TopRow = grdCtrl.TopRow + 1
                Loop
                igGridIgnoreScroll = False
                grdCtrl.Col = ilSearchCol
                gGrid_Search = llRow
                DoEvents
                Exit Function
            End If
        End If
    Next llRow
    For llRow = grdCtrl.FixedRows To grdCtrl.Rows - 1 Step 1
        slStr = Trim$(grdCtrl.TextMatrix(llRow, ilSearchCol))
        If (slStr <> "") Then
            If InStr(1, slStr, slSearchValue, vbTextCompare) = 1 Then
                grdCtrl.TopRow = grdCtrl.FixedRows
                grdCtrl.Row = llRow
                igGridIgnoreScroll = True
                Do While Not grdCtrl.RowIsVisible(grdCtrl.Row)
                    grdCtrl.TopRow = grdCtrl.TopRow + 1
                Loop
                igGridIgnoreScroll = False
                grdCtrl.Col = ilSearchCol
                gGrid_Search = llRow
                DoEvents
                Exit Function
            End If
        End If
    Next llRow
    gGrid_Search = -1
End Function

Public Function gGrid_SearchByType(ilType As Integer, grdCtrl As MSHFlexGrid, ilSearchCol As Integer, slSearchValue As String) As Long
    'ilType(I): 0=String
    '           1=Multi-Name strings
    '           2=Date
    '           3=Time
    '           4=Time in Tenths
    '           5=Length of Time
    '           6=Length of Time in Tenths
    Dim llRow As Long
    Dim slStr As String
    Dim slStrCheck As String
    Dim slStrTest As String
    Dim ilPos As Integer
    Dim ilCheck As Integer
    Dim ilTest As Integer
    Dim ilMatch As Integer
    Dim ilFound As Integer
    Dim llSearchDate As Long
    Dim llSearchTime As Long
    Dim llDate As Long
    Dim llTime As Long
    ReDim slCheck(0 To 0) As String
    ReDim slTest(0 To 0) As String
    
    If ilType = 1 Then
        gParseCDFields slSearchValue, False, slCheck()
    End If
    Select Case ilType
        Case 2  'Date
            If Not gIsDate(slSearchValue) Then
                gGrid_SearchByType = -1
                Exit Function
            End If
            llSearchDate = gDateValue(slSearchValue)
        Case 3  'Time
            If Not gIsTime(slSearchValue) Then
                gGrid_SearchByType = -1
                Exit Function
            End If
            llSearchTime = gTimeToLong(slSearchValue, False)
        Case 4  'Time
            If Not gIsTimeTenths(slSearchValue) Then
                gGrid_SearchByType = -1
                Exit Function
            End If
            llSearchTime = gStrTimeInTenthToLong(slSearchValue, False)
        Case 5  'Length of Time
            If Not gIsLength(slSearchValue) Then
                gGrid_SearchByType = -1
                Exit Function
            End If
            llSearchTime = gLengthToLong(slSearchValue)
        Case 6  'Length of Time in Tenths
            If Not gIsLengthTenths(slSearchValue) Then
                gGrid_SearchByType = -1
                Exit Function
            End If
            llSearchTime = gStrLengthInTenthToLong(slSearchValue)
    End Select
    For llRow = grdCtrl.FixedRows To grdCtrl.Rows - 1 Step 1
        slStr = Trim$(grdCtrl.TextMatrix(llRow, ilSearchCol))
        If (slStr <> "") Then
            Select Case ilType
                Case 0  'String
                    If StrComp(slStr, slSearchValue, vbTextCompare) = 0 Then
                        grdCtrl.TopRow = grdCtrl.FixedRows
                        grdCtrl.Row = llRow
                        igGridIgnoreScroll = True
                        Do While Not grdCtrl.RowIsVisible(grdCtrl.Row)
                            grdCtrl.TopRow = grdCtrl.TopRow + 1
                        Loop
                        igGridIgnoreScroll = False
                        grdCtrl.Col = ilSearchCol
                        gGrid_SearchByType = llRow
                        DoEvents
                        Exit Function
                    End If
                Case 1  'Multi-select string
                    gParseCDFields slStr, False, slTest()
                    ilMatch = True
                    For ilCheck = LBound(slCheck) To UBound(slCheck) Step 1
                        slStrCheck = Trim$(slCheck(ilCheck))
                        If slStrCheck <> "" Then
                            ilFound = False
                            For ilTest = LBound(slTest) To UBound(slTest) Step 1
                                slStrTest = Trim$(slTest(ilTest))
                                If StrComp(slStrCheck, slStrTest, vbTextCompare) = 0 Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilTest
                            If Not ilFound Then
                                ilMatch = False
                                Exit For
                            End If
                        End If
                    Next ilCheck
                    If ilMatch Then
                        grdCtrl.TopRow = grdCtrl.FixedRows
                        grdCtrl.Row = llRow
                        igGridIgnoreScroll = True
                        Do While Not grdCtrl.RowIsVisible(grdCtrl.Row)
                            grdCtrl.TopRow = grdCtrl.TopRow + 1
                        Loop
                        igGridIgnoreScroll = False
                        grdCtrl.Col = ilSearchCol
                        gGrid_SearchByType = llRow
                        DoEvents
                        Exit Function
                    End If
                Case 2  'Date
                    If gIsDate(slStr) Then
                        llDate = gDateValue(slStr)
                        If llSearchDate = llDate Then
                            grdCtrl.TopRow = grdCtrl.FixedRows
                            grdCtrl.Row = llRow
                            igGridIgnoreScroll = True
                            Do While Not grdCtrl.RowIsVisible(grdCtrl.Row)
                                grdCtrl.TopRow = grdCtrl.TopRow + 1
                            Loop
                            igGridIgnoreScroll = False
                            grdCtrl.Col = ilSearchCol
                            gGrid_SearchByType = llRow
                            DoEvents
                            Exit Function
                        End If
                    End If
                Case 3  'Time
                    If gIsTime(slStr) Then
                        llTime = gTimeToLong(slStr, False)
                        If llSearchTime = llTime Then
                            grdCtrl.TopRow = grdCtrl.FixedRows
                            grdCtrl.Row = llRow
                            igGridIgnoreScroll = True
                            Do While Not grdCtrl.RowIsVisible(grdCtrl.Row)
                                grdCtrl.TopRow = grdCtrl.TopRow + 1
                            Loop
                            igGridIgnoreScroll = False
                            grdCtrl.Col = ilSearchCol
                            gGrid_SearchByType = llRow
                            DoEvents
                            Exit Function
                        End If
                    End If
                Case 4  'Time in Tenths
                    If gIsTimeTenths(slStr) Then
                        llTime = gStrTimeInTenthToLong(slStr, False)
                        If llSearchTime = llTime Then
                            grdCtrl.TopRow = grdCtrl.FixedRows
                            grdCtrl.Row = llRow
                            igGridIgnoreScroll = True
                            Do While Not grdCtrl.RowIsVisible(grdCtrl.Row)
                                grdCtrl.TopRow = grdCtrl.TopRow + 1
                            Loop
                            igGridIgnoreScroll = False
                            grdCtrl.Col = ilSearchCol
                            gGrid_SearchByType = llRow
                            DoEvents
                            Exit Function
                        End If
                    End If
                Case 5  'Length of Time
                    If gIsLength(slStr) Then
                        llTime = gLengthToLong(slStr)
                        If llSearchTime = llTime Then
                            grdCtrl.TopRow = grdCtrl.FixedRows
                            grdCtrl.Row = llRow
                            igGridIgnoreScroll = True
                            Do While Not grdCtrl.RowIsVisible(grdCtrl.Row)
                                grdCtrl.TopRow = grdCtrl.TopRow + 1
                            Loop
                            igGridIgnoreScroll = False
                            grdCtrl.Col = ilSearchCol
                            gGrid_SearchByType = llRow
                            DoEvents
                            Exit Function
                        End If
                    End If
                Case 6  'Length of Time in Tenths
                    If gIsLengthTenths(slStr) Then
                        llTime = gStrLengthInTenthToLong(slStr)
                        If llSearchTime = llTime Then
                            grdCtrl.TopRow = grdCtrl.FixedRows
                            grdCtrl.Row = llRow
                            igGridIgnoreScroll = True
                            Do While Not grdCtrl.RowIsVisible(grdCtrl.Row)
                                grdCtrl.TopRow = grdCtrl.TopRow + 1
                            Loop
                            igGridIgnoreScroll = False
                            grdCtrl.Col = ilSearchCol
                            gGrid_SearchByType = llRow
                            DoEvents
                            Exit Function
                        End If
                    End If
            End Select
        End If
    Next llRow
    gGrid_SearchByType = -1
End Function

Public Function gSetCtrlWidth(slCtrlName As String, llStdLetterWidth As Long, llCtrlWidth As Long, ilMaxStr As Integer) As Long
    Dim llWidth As Long
    Dim ilMaxChars As Integer
    
    gSetCtrlWidth = llCtrlWidth
    ilMaxChars = gGetMaxChars(slCtrlName)
    If ilMaxStr > ilMaxChars Then
        ilMaxChars = ilMaxStr
    End If
    If ilMaxChars > 0 Then
        llWidth = llStdLetterWidth * ilMaxChars
        If llWidth > llCtrlWidth Then
            gSetCtrlWidth = llWidth
        End If
    End If
    
End Function

Public Function gGetMaxChars(slCtrlName As String) As Integer
    gGetMaxChars = 0
    Select Case UCase(slCtrlName)
        Case "EVENTTYPE"
            gGetMaxChars = tgNoCharAFE.iEventType
        Case "BUSNAME"
            gGetMaxChars = tgNoCharAFE.iBus
        Case "BUSCTRL"
            gGetMaxChars = tgNoCharAFE.iBusControl
        Case "TIME"
            gGetMaxChars = tgNoCharAFE.iTime
        Case "STARTTYPE"
            gGetMaxChars = tgNoCharAFE.iStartType
        Case "FIXED"
            gGetMaxChars = tgNoCharAFE.iFixedTime
        Case "ENDTYPE"
            gGetMaxChars = tgNoCharAFE.iEndType
        Case "DURATION"
            gGetMaxChars = tgNoCharAFE.iDuration
        Case "MATERIAL"
            gGetMaxChars = tgNoCharAFE.iMaterialType
        Case "AUDIONAME"
            gGetMaxChars = tgNoCharAFE.iAudioName
        Case "AUDIOITEMID"
            gGetMaxChars = tgNoCharAFE.iAudioItemID
        Case "AUDIOISCI"
            gGetMaxChars = tgNoCharAFE.iAudioISCI
        Case "AUDIOCTRL"
            gGetMaxChars = tgNoCharAFE.iAudioControl
        Case "PROTNAME"
            gGetMaxChars = tgNoCharAFE.iProtAudioName
        Case "PROTITEMID"
            gGetMaxChars = tgNoCharAFE.iProtItemID
        Case "PROTISCI"
            gGetMaxChars = tgNoCharAFE.iProtISCI
        Case "PROTCTRL"
            gGetMaxChars = tgNoCharAFE.iProtAudioControl
        Case "BKUPNAME"
            gGetMaxChars = tgNoCharAFE.iBkupAudioName
        Case "BKUPCTRL"
            gGetMaxChars = tgNoCharAFE.iBkupAudioControl
        Case "RELAY1"
            gGetMaxChars = tgNoCharAFE.iRelay1
        Case "RELAY2"
            gGetMaxChars = tgNoCharAFE.iRelay2
        Case "FOLLOW"
            gGetMaxChars = tgNoCharAFE.iFollow
        Case "SILENCETIME"
            gGetMaxChars = tgNoCharAFE.iSilenceTime
        Case "SILENCE1"
            gGetMaxChars = tgNoCharAFE.iSilence1
        Case "SILENCE2"
            gGetMaxChars = tgNoCharAFE.iSilence2
        Case "SILENCE3"
            gGetMaxChars = tgNoCharAFE.iSilence3
        Case "SILENCE4"
            gGetMaxChars = tgNoCharAFE.iSilence4
        Case "NETCUE1"
            gGetMaxChars = tgNoCharAFE.iStartNetcue
        Case "NETCUE2"
            gGetMaxChars = tgNoCharAFE.iStopNetcue
        Case "TITLE1"
            gGetMaxChars = tgNoCharAFE.iTitle1
        Case "TITLE2"
            gGetMaxChars = tgNoCharAFE.iTitle2
        Case "ABCFORMAT"
            gGetMaxChars = tgNoCharAFE.iABCFormat
        Case "ABCPGMCODE"
            gGetMaxChars = tgNoCharAFE.iABCPgmCode
        Case "ABCXDSMODE"
            gGetMaxChars = tgNoCharAFE.iABCXDSMode
        Case "ABCRECORDITEM"
            gGetMaxChars = tgNoCharAFE.iABCRecordItem
    End Select

End Function

Public Function gSetMaxChars(slCtrlName As String, ilMaxStr As Integer)
    Dim ilMaxChars As Integer
    
    ilMaxChars = gGetMaxChars(slCtrlName)
    If ilMaxStr > ilMaxChars Then
        gSetMaxChars = ilMaxStr
    Else
        gSetMaxChars = ilMaxChars
    End If
End Function



Public Function gGrid_GetRowCol(grdCtrl As MSHFlexGrid, x As Single, y As Single, llOutRow As Long, llOutCol As Long) As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim llColLeftPos As Long
    
    For llRow = grdCtrl.FixedRows To grdCtrl.Rows - 1 Step 1
        If grdCtrl.RowIsVisible(llRow) Then
            If (y >= grdCtrl.RowPos(llRow)) And (y <= grdCtrl.RowPos(llRow) + grdCtrl.RowHeight(llRow) - 15) Then
                llColLeftPos = grdCtrl.ColPos(0)
                For llCol = 0 To grdCtrl.Cols - 1 Step 1
                    If grdCtrl.ColWidth(llCol) > 0 Then
                        If (x >= llColLeftPos) And (x <= llColLeftPos + grdCtrl.ColWidth(llCol)) Then
                            llOutRow = llRow
                            llOutCol = llCol
                            gGrid_GetRowCol = True
                            Exit Function
                        End If
                        llColLeftPos = llColLeftPos + grdCtrl.ColWidth(llCol)
                    End If
                Next llCol
            End If
        End If
    Next llRow
    gGrid_GetRowCol = False
    Exit Function
End Function

Public Sub gSetMousePointer(grdCtrl1 As MSHFlexGrid, grdCtrl2 As MSHFlexGrid, ilPointer As Integer)
    Screen.MousePointer = ilPointer
    grdCtrl1.MousePointer = ilPointer
    grdCtrl2.MousePointer = ilPointer
End Sub

Public Function gColOk(grdCtrl As MSHFlexGrid, llRow As Long, llCol As Long) As Integer
    Dim slStr As String
    
    gColOk = True
    If grdCtrl.ColWidth(llCol) <= 0 Then
        gColOk = False
        Exit Function
    End If
    grdCtrl.Row = llRow
    grdCtrl.Col = llCol
    If grdCtrl.CellBackColor = LIGHTYELLOW Then
        gColOk = False
        Exit Function
    End If
End Function

Public Function gGetAllowedChars(slCtrlName As String, ilRecordFieldSize As Integer)
    Dim ilMaxChars As Integer
    
    ilMaxChars = gGetMaxChars(slCtrlName)
    If ilRecordFieldSize > ilMaxChars Then
        gGetAllowedChars = ilMaxChars
    Else
        gGetAllowedChars = ilRecordFieldSize
    End If
End Function
