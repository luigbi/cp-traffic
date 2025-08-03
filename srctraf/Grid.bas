Attribute VB_Name = "GridSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Grid.bas on Wed 6/17/09 @ 12:56 PM **
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Grid.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Grid subs and functions
Option Explicit
Option Compare Text


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
    'grdCtrl.Row = grdCtrl.FixedRows
    'grdCtrl.RowSel = grdCtrl.Row
    If grdCtrl.FixedRows > 0 Then
        grdCtrl.Row = grdCtrl.FixedRows - 1
        grdCtrl.RowSel = grdCtrl.Row
    Else
        grdCtrl.Row = grdCtrl.FixedRows
        grdCtrl.RowSel = grdCtrl.Row
    End If
    grdCtrl.Redraw = True
End Sub

Public Function gGrid_GetRowCol(grdCtrl As MSHFlexGrid, X As Single, Y As Single, llOutRow As Long, llOutCol As Long) As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim llColLeftPos As Long

    For llRow = grdCtrl.FixedRows To grdCtrl.Rows - 1 Step 1
        If grdCtrl.RowIsVisible(llRow) Then
            If (Y >= grdCtrl.RowPos(llRow)) And (Y <= grdCtrl.RowPos(llRow) + grdCtrl.RowHeight(llRow) - 15) Then
                llColLeftPos = grdCtrl.ColPos(0)
                For llCol = 0 To grdCtrl.Cols - 1 Step 1
                    If grdCtrl.ColWidth(llCol) > 0 Then
                        If (X >= llColLeftPos) And (X <= llColLeftPos + grdCtrl.ColWidth(llCol)) Then
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

Public Sub gGrid_IntegralHeight(grdCtrl As MSHFlexGrid, ilStdRowHeight As Integer)
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

    Do While llHeight + ilStdRowHeight <= grdCtrl.Height
        llHeight = llHeight + ilStdRowHeight
    Loop
    If grdCtrl.ScrollBars <> 1 And grdCtrl.ScrollBars <> 3 Then
        grdCtrl.Height = llHeight ' + 15
    Else
        grdCtrl.Height = llHeight + GRIDSCROLLHEIGHT '+ 15 + GRIDSCROLLHEIGHT
    End If
End Sub

Public Sub gGrid_AlignAllColsLeft(grdCtrl As MSHFlexGrid)
    Dim ilCol As Integer

    For ilCol = 0 To grdCtrl.Cols - 1 Step 1
        grdCtrl.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol

End Sub

Public Sub gGrid_FillWithRows(grdCtrl As MSHFlexGrid, llRowHeight As Long)
    Dim llRows As Long
    Dim llCols As Long
    Dim llFillNoRow As Long

    'llFillNoRow = grdCtrl.Height \ grdCtrl.RowHeight(grdCtrl.FixedRows) - grdCtrl.FixedRows - 1
    llFillNoRow = grdCtrl.Height \ llRowHeight - grdCtrl.FixedRows - 1

    For llRows = grdCtrl.FixedRows To grdCtrl.FixedRows + llFillNoRow Step 1
        Do While llRows >= grdCtrl.Rows
            grdCtrl.AddItem ""
            For llCols = 0 To grdCtrl.Cols - 1 Step 1
                grdCtrl.TextMatrix(llRows, llCols) = ""
            Next llCols
        Loop
    Next llRows
End Sub

