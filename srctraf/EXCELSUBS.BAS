Attribute VB_Name = "EXCELSUBS"

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExcelSubs.BAS
'
' Release: 1.0
'

Public Enum XlFormatConditionType
    xlCellValue = 1
    xlExpression = 2
End Enum
Public Enum XlFormatConditionOperator
    xlBetween = 1
    xlNotBetween = 2
    xlEqual = 3
    xlNotEqual = 4
    xlGreater = 5
    xlLess = 6
    xlGreaterEqual = 7
    xlLessEqual = 8
End Enum
Public Enum XlAlign
    xlTop = -4160
    xlCenter = -4108
    xlBottom = -4107
    xlLeft = -4131
    xlRight = -4152
End Enum

Public Function gExcelOutputGeneration(slAction As String, Optional olBook As Object, Optional olSheet As Object, Optional ilSheetNo As Integer, Optional slvalue = "", Optional ilRow As Integer = -1, Optional ilColumn As Integer = -1, Optional slSplitDelimiter = "") As Boolean
    'slAction:
    '          "C" Create Ecel application: (Parameters: slAction)
    '          "O" Open (Book and Sheet): (Parameters: slAction, olBook, olSheet, ilSheetNo)
    '          "W"=Write slValue into row(ilRow), column(ilColumn): (Parameters: slAction, , olSheet, ilSheetNo, slValue, ilRow, ilColumn, slSplitDelimiter)
    '              If slDelimiter = "", then write value w/o split
    '              If slDelimiter defined, then split and start writing columns at ilColumn
    '          "S" Save into slValue: (Parameters: slAction, olBook, , , slValue), using Excel 2007 without Macros format (.xlsx format)
    '          "S8" Save into slValue: (Parameters: slAction, olBook, , , slValue), using Excel 8 format (.xls format)
    '          "MN" Make a Named Range  (Could be called again to make additional ranges) - see example below
    '               Parameters:
    '               slAction = "NM:
    '               slValue = Name of Range
    '               ilRow = Row Number of End of Range
    '               ilColumn = Column Number of End of Range
    '               slSplitDelimiter = optional Start Range (If left empty the $A$1 is used)
    '          "V" View generated Excel spread sheet: (Parameters: slAction)
    '          "Q" Quit: (Parameters: slAction)
    '                    (the quit option also will close the spread sheet if visible)
    '          "AF" AutoFit : (Parameters: slAction, , olSheet, , , ,ilColumn)
    '                         (ilColumn = -1: All Columns)
    '          "FS" font size: (Parameters: slAction, , olSheet, , slValue, ilRow, ilColumn)
    '                          (ilRow = -1: By column; ilColumn = -1: By Row)
    '          "FN" font name: (Parameters: slAction, , olSheet, , slValue, ilRow, ilColumn)
    '                          (ilRow = -1: By column; ilColumn = -1: By Row)
    '          "FB" font Bold: (Parameters: slAction, , olSheet, , slValue, ilRow, ilColumn)
    '                          (ilRow = -1: By column; ilColumn = -1: By Row)
    '          "FI" font italic: (Parameters: slAction, , olSheet, , slValue, ilRow, ilColumn)
    '                            (ilRow = -1: By column; ilColumn = -1: By Row)
    '          "FU" font underline: (Parameters: slAction, , olSheet, , slValue, ilRow, ilColumn)
    '                               (ilRow = -1: By column; ilColumn = -1: By Row)
    '          "FX" font strikethrough: (Parameters: slAction, , olSheet, , slValue, ilRow, ilColumn)
    '                               (ilRow = -1: By column; ilColumn = -1: By Row)
    '          "FC" Forground Color (font color): (Parameters: slAction, , olSheet, , slValue, ilRow, ilColumn)
    '          "BC" Background Color: (Parameters: slAction, , olSheet, , slValue, ilRow, ilColumn)
    '          "H" Hide column: (Parameters: slAction, , olSheet, , slValue, ilRow, ilColumn)
    '                           (ilRow = -1: By column; ilColumn = -1: By Row)
    '          "CW" Column Width: (Parameters: slAction, , olSheet, , slValue, ilRow, ilColumn)
    '          "NF" Number Format: (Parameters: slAction, , olSheet, , slValue, ilRow, ilColumn)
    '                              (ilRow = -1: By column; ilColumn = -1: By Row)
    '          "FV" Formula Value: (Parameters: slAction, , olSheet, , slValue, ilRow, ilColumn)
    '                              (ilRow = -1: By column; ilColumn = -1: By Row)
    '                              gExcelOutputGeneration("FV", omBook, omSheet, , "=" & gExcelColumnToLetter(18) & imExcelRow & "*" & gExcelColumnToLetter(19) & imExcelRow, imExcelRow, 23, slDelimiter)
    '          "HA" Horizontal Alighnment:(Parameters: slAction, , olSheet, , slValue, ilRow, ilColumn)
    '                                     slValue= Left; Center or Right
    '          "VA" Vertical Alighnment:(Parameters: slAction, , olSheet, , slValue, ilRow, ilColumn)
    '                                    slValue= Top; Center or Bottm
    '
    'Sequence of calls #1:
    '  Create Excel object  gGenExcel("C")
    '  Open object          gGenExcel("O", omBook, omSheet, 1)
    '  Write to object      gGenExcel("W",, omSheet, slValue, 1, 1, "")
    '  Save Object (this is after the write to object and must include extension xlsx)
    '                       gGenExcel("S", omBook,, slValue)
    '  Either Quit or view excel. (if quit, need to create object)
    '                       gGenExcel("V") or gGenExcel("Q")
    '
    'or Sequence of calls #2 (Sheet # used is 1):
    '  Write to object      gGenExcel("W",, omSheet, slValue, 1, 1, "")
    '  Save Object (this is after the write to object and must include extension xlsx)
    '                       gGenExcel("S", omBook,, slValue)
    '  Either Quit or view excel. (if quit, need to re-create object)
    '                       gGenExcel("V") or gGenExcel("Q")
    '
    'see: http://www.vbforums.com/showthread.php?391665-Tutorial-FAQ-How-to-Automate-Excel-from-VB6-(or-VB5-VBA)
    '     for examples
    '
    'In Form_Unload
    '    On Error Resume Next
    '    Set omSheet = Nothing
    '    Set omBook = Nothing
    '    ilRet = gExcelOutputGeneration("Q")
    '    Set ogExcel = Nothing
    '
    Dim slValues() As String
    Dim ilValue As Integer
    Dim slTmpValue As String
    
    On Error GoTo err:
    If slAction <> "C" Then
        If Not bgExcelCreated Then
            Set ogExcel = CreateObject("Excel.Application")
            bgExcelCreated = True
            ogExcel.Visible = False
            If slAction <> "O" Then
                Set olBook = ogExcel.Workbooks.Add
                ogExcel.Visible = False
                Set olSheet = olBook.Worksheets(1)
                'ogExcel.Visible = False
            End If
        End If
    End If
    If slAction = "C" Then
        If Not bgExcelCreated Then
            Set ogExcel = CreateObject("Excel.Application")
            ogExcel.Visible = False
            bgExcelCreated = True
        End If
    ElseIf slAction = "O" Then
        Set olBook = ogExcel.Workbooks.Add
        ogExcel.Visible = False
        Set olSheet = olBook.Worksheets(ilSheetNo)
        'ogExcel.Visible = False
    ElseIf slAction = "W" Then
        If slSplitDelimiter = "" Then
            olSheet.Cells(ilRow, ilColumn).Value = slvalue
        Else
            slValues = Split(slvalue, slSplitDelimiter)
            If Not IsArray(slValues) Then
                Exit Function
            End If
            For ilValue = LBound(slValues) To UBound(slValues) Step 1
                olSheet.Cells(ilRow, ilColumn + ilValue).Value = slValues(ilValue)
            Next ilValue
        End If
    ElseIf slAction = "S" Then
        olBook.SaveAs slvalue, 51   'xlOpenxmlworkbook    '"C:csi\Book1.xlsx"
        'These are the main file formats in Excel 2007:
        '   51 = xlOpenXMLWorkbook (without macro's in 2007, .xlsx)
        '   52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007, .xlsm)
        '   50 = xlExcel12 (Excel Binary Workbook in 2007 with or without macro's, .xlsb)
        '   56 = xlExcel8 (97-2003 format in Excel 2007, .xls)
    ElseIf slAction = "S8" Then 'Save in Excel 8 Compatibility (XLS format)
        Replace slvalue, ".xlsx", ".xls" 'just incase name contains xlsx, rename it to XLS
        olBook.SaveAs slvalue, 56   'xlOpenxmlworkbook    '"C:csi\Book1.xls"
        'These are the main file formats in Excel 2007:
        '   51 = xlOpenXMLWorkbook (without macro's in 2007, .xlsx)
        '   52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007, .xlsm)
        '   50 = xlExcel12 (Excel Binary Workbook in 2007 with or without macro's, .xlsb)
        '   56 = xlExcel8 (97-2003 format in Excel 2007, .xls)
    ElseIf slAction = "NM" Then
        'TTP 10260 - 7/28/21 - JW -  "Name Manager"
        'Example - Makes range out of entire range of data: ilRet = gExcelOutputGeneration("NM", , , , "RangeName")
        'Example - makes range B2 - Z3 titled 'SmallRange': ilRet = gExcelOutputGeneration("NM", , , , "SmallRange", 3, 26,"$B$2")
        If slSplitDelimiter <> "" Then
            slTmpValue = "=Sheet1!" & slSplitDelimiter & ":$"
        Else
            slTmpValue = "=Sheet1!$A$1:$"
        End If
        If ilColumn = -1 Then
            ilColumn = olSheet.UsedRange.Columns.Count
        End If
        If ilRow = -1 Then
            ilRow = olSheet.UsedRange.Rows.Count
        End If
        slTmpValue = slTmpValue & mConvertToLetter(ilColumn)
        slTmpValue = slTmpValue & "$"
        slTmpValue = slTmpValue & ilRow
        'Debug.Print slvalue & slTmpValue
        ogExcel.Names.Add slvalue, slTmpValue
    ElseIf slAction = "V" Then
        ogExcel.Visible = True
        'the activeexcel code appears to be working but the actual error was in ReRate after returning from here I set focus to cmcCancel
        'ActivateExcel
    ElseIf slAction = "Q" Then
        If bgExcelCreated Then
            ogExcel.Quit
            Set ogExcel = Nothing
        End If
        bgExcelCreated = False
    ElseIf slAction = "AF" Then
        If ilColumn = -1 Then
            'olSheet.Columns("A:ZZ").AutoFit
            olSheet.Columns().AutoFit
        Else
            olSheet.Columns(ilColumn).AutoFit
        End If
    ElseIf slAction = "FS" Then
        If ilColumn = -1 Then
            olSheet.Rows(ilRow).Font.Size = Val(slvalue)
        ElseIf ilRow = -1 Then
            olSheet.Columns(ilColumn).Font.Size = Val(slvalue)
        Else
            olSheet.Cells(ilRow, ilColumn).Font.Size = Val(slvalue)
        End If
    ElseIf slAction = "FN" Then
        If ilColumn = -1 Then
            olSheet.Rows(ilRow).Font.Name = slvalue
        ElseIf ilRow = -1 Then
            olSheet.Columns(ilColumn).Font.Name = slvalue
        Else
            olSheet.Cells(ilRow, ilColumn).Font.Name = slvalue
        End If
    ElseIf slAction = "FB" Then
        If ilColumn = -1 Then
            If slvalue = "True" Or Left(slvalue, 1) = "T" Then
                olSheet.Rows(ilRow).Font.Bold = True
            Else
                olSheet.Rows(ilRow).Font.Bold = False
            End If
        ElseIf ilRow = -1 Then
            If slvalue = "True" Or Left(slvalue, 1) = "T" Then
                olSheet.Columns(ilColumn).Font.Bold = True
            Else
                olSheet.Columns(ilColumn).Font.Bold = False
            End If
        Else
            If slvalue = "True" Or Left(slvalue, 1) = "T" Then
                olSheet.Cells(ilRow, ilColumn).Font.Bold = True
            Else
                olSheet.Cells(ilRow, ilColumn).Font.Bold = False
            End If
        End If
    ElseIf slAction = "FI" Then
        If ilColumn = -1 Then
            If slvalue = "True" Or Left(slvalue, 1) = "T" Then
                olSheet.Rows(ilRow).Font.Italic = True
            Else
                olSheet.Rows(ilRow).Font.Italic = False
            End If
        ElseIf ilRow = -1 Then
            If slvalue = "True" Or Left(slvalue, 1) = "T" Then
                olSheet.Columns(ilColumn).Font.Italic = True
            Else
                olSheet.Columns(ilColumn).Font.Italic = False
            End If
        Else
            If slvalue = "True" Or Left(slvalue, 1) = "T" Then
                olSheet.Cells(ilRow, ilColumn).Font.Italic = True
            Else
                olSheet.Cells(ilRow, ilColumn).Font.Italic = False
            End If
        End If
    ElseIf slAction = "FU" Then
        If ilColumn = -1 Then
            If slvalue = "True" Or Left(slvalue, 1) = "T" Then
                olSheet.Rows(ilRow).Font.Underline = xlUnderlineStyleSingle
            Else
                olSheet.Rows(ilRow).Font.Underline = xlUnderlineStyleNone
            End If
        ElseIf ilRow = -1 Then
            If slvalue = "True" Or Left(slvalue, 1) = "T" Then
                olSheet.Columns(ilColumn).Font.Underline = xlUnderlineStyleSingle
            Else
                olSheet.Columns(ilColumn).Font.Underline = xlUnderlineStyleNone
            End If
        Else
            If slvalue = "True" Or Left(slvalue, 1) = "T" Then
                olSheet.Cells(ilRow, ilColumn).Font.Underline = xlUnderlineStyleSingle
            Else
                olSheet.Cells(ilRow, ilColumn).Font.Underline = xlUnderlineStyleNone
            End If
        End If
    ElseIf slAction = "FX" Then 'Strikethrough
        If ilColumn = -1 Then
            If slvalue = "True" Or Left(slvalue, 1) = "T" Then
                olSheet.Rows(ilRow).Font.Strikethrough = True
            Else
                olSheet.Rows(ilRow).Font.Strikethrough = True
            End If
        ElseIf ilRow = -1 Then
            If slvalue = "True" Or Left(slvalue, 1) = "T" Then
                olSheet.Columns(ilColumn).Font.Strikethrough = True
            Else
                olSheet.Columns(ilColumn).Font.Strikethrough = True
            End If
        Else
            If slvalue = "True" Or Left(slvalue, 1) = "T" Then
                olSheet.Cells(ilRow, ilColumn).Font.Strikethrough = True
            Else
                olSheet.Cells(ilRow, ilColumn).Font.Strikethrough = True
            End If
        End If
    ElseIf slAction = "FC" Then
        'olSheet.Cells(ilRow, ilColumn).Font.Color = Val(slvalue)
        If ilColumn = -1 Then
            olSheet.Rows(ilRow).Font.Color = Val(slvalue)
        ElseIf ilRow = -1 Then
            olSheet.Columns(ilColumn).Font.Color = Val(slvalue)
        Else
            olSheet.Cells(ilRow, ilColumn).Font.Color = Val(slvalue)
        End If
    ElseIf slAction = "BC" Then
        'olSheet.Cells(ilRow, ilColumn).Interior.Color = Val(slvalue)
        If ilColumn = -1 Then
            olSheet.Rows(ilRow).Interior.Color = Val(slvalue)
        ElseIf ilRow = -1 Then
            olSheet.Columns(ilColumn).Interior.Color = Val(slvalue)
        Else
            olSheet.Cells(ilRow, ilColumn).Interior.Color = Val(slvalue)
        End If
    ElseIf slAction = "NF" Then
        If ilColumn = -1 Then
            olSheet.Rows(ilRow).NumberFormat = slvalue
        ElseIf ilRow = -1 Then
            olSheet.Columns(ilColumn).NumberFormat = slvalue
        Else
            olSheet.Cells(ilRow, ilColumn).NumberFormat = slvalue
        End If
    ElseIf slAction = "FV" Then
        If ilColumn = -1 Then
            olSheet.Rows(ilRow).Formula = slvalue
        ElseIf ilRow = -1 Then
            olSheet.Columns(ilColumn).Formula = slvalue
        Else
            olSheet.Cells(ilRow, ilColumn).Formula = slvalue
        End If
    ElseIf slAction = "HA" Then
        If IsNumeric(slvalue) Then
            If ilColumn = -1 Then
                olSheet.Rows(ilRow).HorizontalAlignment = Val(slvalue)
            ElseIf ilRow = -1 Then
                olSheet.Columns(ilColumn).HorizontalAlignment = Val(slvalue)
            Else
                olSheet.Cells(ilRow, ilColumn).HorizontalAlignment = Val(slvalue)
            End If
        Else
            If ilColumn = -1 Then
                If slvalue = "Left" Or Left(slvalue, 1) = "L" Then
                    olSheet.Rows(ilRow).HorizontalAlignment = -4131
                ElseIf slvalue = "Center" Or Left(slvalue, 1) = "C" Then
                    olSheet.Rows(ilRow).HorizontalAlignment = -4108
                ElseIf slvalue = "Right" Or Left(slvalue, 1) = "R" Then
                    olSheet.Rows(ilRow).HorizontalAlignment = 4152
                End If
            ElseIf ilRow = -1 Then
                If slvalue = "Left" Or Left(slvalue, 1) = "L" Then
                    olSheet.Columns(ilColumn).HorizontalAlignment = -4131
                ElseIf slvalue = "Center" Or Left(slvalue, 1) = "C" Then
                    olSheet.Columns(ilColumn).HorizontalAlignment = -4108
                ElseIf slvalue = "Right" Or Left(slvalue, 1) = "R" Then
                    olSheet.Columns(ilColumn).HorizontalAlignment = -4152
                End If
            Else
                If slvalue = "Left" Or Left(slvalue, 1) = "L" Then
                    olSheet.Cells(ilRow, ilColumn).HorizontalAlignment = -4131
                ElseIf slvalue = "Center" Or Left(slvalue, 1) = "C" Then
                    olSheet.Cells(ilRow, ilColumn).HorizontalAlignment = -4108
                ElseIf slvalue = "Right" Or Left(slvalue, 1) = "R" Then
                    olSheet.Cells(ilRow, ilColumn).HorizontalAlignment = 4152
                End If
            End If
        End If
    ElseIf slAction = "VA" Then
        If ilColumn = -1 Then
            If slvalue = "Top" Or Left(slvalue, 1) = "T" Then
                olSheet.Rows(ilRow).VerticalAlignment = -4160
            ElseIf slvalue = "Center" Or Left(slvalue, 1) = "C" Then
                olSheet.Rows(ilRow).VerticalAlignment = -4108
            ElseIf slvalue = "Bottom" Or Left(slvalue, 1) = "B" Then
                olSheet.Rows(ilRow).VerticalAlignment = 4107
            End If
        ElseIf ilRow = -1 Then
            If slvalue = "Top" Or Left(slvalue, 1) = "T" Then
                olSheet.Columns(ilColumn).VerticalAlignment = -4160
            ElseIf slvalue = "Center" Or Left(slvalue, 1) = "C" Then
                olSheet.Columns(ilColumn).VerticalAlignment = -4108
            ElseIf slvalue = "Bottom" Or Left(slvalue, 1) = "B" Then
                olSheet.Columns(ilColumn).VerticalAlignment = 4107
            End If
        Else
            If slvalue = "Top" Or Left(slvalue, 1) = "T" Then
                olSheet.Cells(ilRow, ilColumn).VerticalAlignment = -4160
            ElseIf slvalue = "Center" Or Left(slvalue, 1) = "C" Then
                olSheet.Cells(ilRow, ilColumn).VerticalAlignment = -4108
            ElseIf slvalue = "Bottom" Or Left(slvalue, 1) = "B" Then
                olSheet.Cells(ilRow, ilColumn).VerticalAlignment = 4107
            End If
        End If
    ElseIf slAction = "WT" Then 'Wrap Text
        olSheet.Cells(ilRow, ilColumn).WrapText = slvalue
    End If
    
    If slAction = "H" Then
        If ilColumn = -1 Then
            If slvalue = "True" Or Left(slvalue, 1) = "T" Then
                olSheet.Rows(ilRow).Hidden = True
            Else
                olSheet.Rows(ilRow).Hidden = False
            End If
        ElseIf ilRow = -1 Then
            If slvalue = "True" Or Left(slvalue, 1) = "T" Then
                olSheet.Columns(ilColumn).Hidden = True
            Else
                olSheet.Columns(ilColumn).Hidden = False
            End If
        Else
            If slvalue = "True" Or Left(slvalue, 1) = "T" Then
                olSheet.Cells(ilRow, ilColumn).Hidden = True
            Else
                olSheet.Cells(ilRow, ilColumn).Hidden = False
            End If
        End If
    End If
    If slAction = "CW" Then
        olSheet.Columns(ilColumn).ColumnWidth = Val(slvalue)
    End If
    
    gExcelOutputGeneration = True
    Exit Function
err:
    gExcelOutputGeneration = False
    Exit Function
    
End Function
Public Function gExcelColumnToLetter(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      gExcelColumnToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      gExcelColumnToLetter = gExcelColumnToLetter & Chr(iRemainder + 64)
   End If
End Function

Public Sub ActivateExcel()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ActivateExcel
' This procedure activates the main Excel application window,
' ("XLMAIN") moving it to the top of the Z-Order and sets keyboard
' focus to Excel.
'
' !!!!!!!!!!!!!!!!!!!!!!!!!
' NOTE: This will not work properly if a VBA Editor is open.
' If a VBA Editor window is open, the system will set focus
' to that window, rather than the XLMAIN window.
' !!!!!!!!!!!!!!!!!!!!!!!!!
'
' This code should work to activate any application. Change the
' value of C_MAIN_WINDOW_CLASS to the application's main window
' class (e.g., "OpusApp" for Word).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Res As Long     ' General purpose Result variable
    Dim XLHWnd As Long  ' Window handle of Excel
    Const C_MAIN_WINDOW_CLASS = "XLMAIN"
    '''''''''''''''''''''''''''''''''''''''''''
    ' Get the window handle of the main
    ' Excel application window ("XLMAIN"). If
    ' more than one instance of Excel is running,
    ' you have no control over which
    ' instance's HWnd will be retrieved.
    ' Related Note: You MUST use vbNullString
    ' not an empty string "" in the call to
    ' FindWindow. When calling API functions
    ' there is a difference between vbNullString
    ' and an empty string "".
    ''''''''''''''''''''''''''''''''''''''''''
    XLHWnd = FindWindow(lpClassName:=C_MAIN_WINDOW_CLASS, _
                    lpWindowName:=vbNullString)
        
        If XLHWnd > 0 Then
        '''''''''''''''''''''''''''''''''''''''''
        ' If HWnd is > 0, FindWindow successfully
        ' found the Excel main application window.
        ' Move XLMAIN to the top of the
        ' Z-Order.
        '''''''''''''''''''''''''''''''''''''''''
        Res = BringWindowToTop(HWnd:=XLHWnd)
         
        If Res = 0 Then
            Debug.Print "Error With BringWindowToTop:  " & _
                CStr(err.LastDllError)
          Else
            '''''''''''''''''''''''''''''''''
            ' No error.
            ' Set keyboard input focus XLMAIN
            '''''''''''''''''''''''''''''''''
            SetFocus HWnd:=XLHWnd
          End If
    Else
        '''''''''''''''''''''''''''''''''
        ' HWnd was 0. FindWindow couldn't
        ' find Excel.
        '''''''''''''''''''''''''''''''''
        Debug.Print "Can't find Excel"
    End If
End Sub

Function mConvertToLetter(iCol As Integer) As String
   Dim A As Long
   Dim b As Long
   A = iCol
   mConvertToLetter = ""
   Do While iCol > 0
      A = Int((iCol - 1) / 26)
      b = (iCol - 1) Mod 26
      mConvertToLetter = Chr(b + 65) & mConvertToLetter
      iCol = A
   Loop
End Function
