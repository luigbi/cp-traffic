VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExcelWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' to test if user has excel:
'    Set myExcel = New CExcelWriter
'    If myExcel.HasExcel = ExcelNotPresent Then
'        msgbox "User doesn't have excel!"
'    End If

Dim smFilePath As String
Dim myFile As FileSystemObject
Dim smErrorMessage As String
Dim bmIsError As Boolean
Dim bmRanStart As Boolean
Dim bmRanWrite As Boolean
Dim myExcel As Object ' excel.Application
Dim myWorkbook As Object ' excel.Workbook
Private Const EXT = "xls"
Dim imSheets As Integer
Dim imIsExcel As IsExcel
Enum IsExcel
    ExcelExists = 1
    ExcelNotPresent = 0
End Enum
Public Property Get HasExcel() As IsExcel
    HasExcel = imIsExcel
End Property
Public Property Get ExcelVersion() As String
    Dim slRet As String
    
    If bmRanStart = True Then
        If Not myExcel Is Nothing Then
            slRet = myExcel.Version & "." & myExcel.Build
        Else
            slRet = "0"
        End If
    Else
        slRet = "0"
    End If
    ExcelVersion = slRet
End Property
Public Property Get SheetsInWorkBook() As Integer
    SheetsInWorkBook = imSheets
End Property
Public Property Let SheetsInWorkBook(ByVal ilValue As Integer)
    If ilValue > 0 Then
        imSheets = ilValue
    Else
        imSheets = 1
    End If
End Property
Public Property Get ErrorMessage() As String
    ErrorMessage = smErrorMessage
End Property
Public Property Get FilePath() As String
    FilePath = smFilePath
End Property
Public Property Let FilePath(ByVal slvalue As String)
    Dim ilRet As Integer
    
    mClearError
    If myFile.FolderExists(slvalue) Then
        smFilePath = slvalue
    Else
        smFilePath = ""
        bmIsError = True
        smErrorMessage = "The filepath does not exist."
    End If
End Property
Private Sub Class_Initialize()
On Error GoTo errbox
    'set myExcel just to see if it will fail
    Set myExcel = CreateObject("excel.application")
    imIsExcel = ExcelExists
Cleanup:
    Clear True
    Exit Sub
errbox:
    smErrorMessage = " Could not create Excel Objects."
    bmIsError = True
    imIsExcel = ExcelNotPresent
    GoTo Cleanup
End Sub
Public Function Start(slOutPath As String, blOverwriteOld As Boolean) As Boolean
    Dim slPath As String
    Dim slName As String
    'refresh variables, fix name and path, deleting old file as needed
On Error GoTo errbox
    mClearError
    mCreateExcel
    'changes bmIsError as needed
    slName = mGetName(slOutPath)
    slPath = myFile.GetParentFolderName(slOutPath)
    If Not myFile.FolderExists(slPath) Then
        bmIsError = True
        smErrorMessage = "The folder to write the excel file does not exist."
    End If
    If Not bmIsError Then
        smFilePath = myFile.BuildPath(slPath, slName)
        If myFile.FileExists(smFilePath) Then
            If blOverwriteOld Then
                myFile.DeleteFile smFilePath, True
            Else
                smErrorMessage = "that file already exists"
                bmIsError = True
                smFilePath = ""
            End If
        End If
    End If
    bmRanStart = Not bmIsError
    Start = bmRanStart
    Exit Function
errbox:
    bmIsError = True
    smErrorMessage = smErrorMessage & " Problem in cExcelWriter-Start."
End Function
Public Sub WriteExcelForceToString(ilRow As Integer, ilColumn As Integer, slvalue As String)
    'a string that looks like a date will automatically be formatted as date unless below is set!
    myWorkbook.ActiveSheet.Cells(ilRow, ilColumn).NumberFormat = "@"
    WriteExcel ilRow, ilColumn, slvalue
End Sub
Public Sub WriteExcel(ilRow As Integer, ilColumn As Integer, slvalue As String)
    myWorkbook.ActiveSheet.Cells(ilRow, ilColumn).Value = slvalue
    bmRanWrite = True
End Sub
Public Function GoToSheet(ilSheetNumber As Integer, Optional slNameSheet As String) As Boolean
    Dim blRet As Boolean
    
    If Not bmIsError Then
        If bmRanStart Then
            If ilSheetNumber <= imSheets Then
                myWorkbook.Sheets(ilSheetNumber).Select
                If Len(slNameSheet) > 0 Then
                    myExcel.Worksheets(ilSheetNumber).Name = slNameSheet
                End If
            Else
                bmIsError = True
                smErrorMessage = " Attempted to write to sheet that did not exist."
            End If
        Else
            bmIsError = True
            smErrorMessage = " You must start an excel file before opening a new workbook."
        End If
    End If
    GoToSheet = Not bmIsError
End Function
Public Function finish() As Boolean
    Const WORKBOOKNORMAL As Integer = -4143
    If Not bmIsError Then
        mClearError
        If bmRanStart Then
            If bmRanWrite Then
                If Len(smFilePath) > 0 Then
                    myWorkbook.SaveAs smFilePath, WORKBOOKNORMAL ' xlWorkbookNormal
                    myExcel.Quit
                Else
                    bmIsError = True
                    smErrorMessage = "couldn't set excel file path in 'finish'."
                End If
            Else
                bmIsError = True
                smErrorMessage = "must write values to excel before saving."
            End If
        Else
            bmIsError = True
            smErrorMessage = "must run 'start' before running 'finish'"
        End If
    End If
    Clear True
    finish = Not bmIsError
End Function
Public Sub Clear(IncludeExcel As Boolean)
    Set myFile = New FileSystemObject
    smFilePath = ""
    bmRanStart = False
    bmRanWrite = False
    imSheets = 1
    mClearError
    If IncludeExcel Then
        mClearExcel
    End If
End Sub
Private Function mGetName(slPath As String) As String
    Dim slName As String
    Dim slExt As String
    
    slName = myFile.GetFileName(slPath)
    If Len(slName) > 0 Then
        slExt = myFile.GetExtensionName(slName)
        If slExt <> EXT Then
            If Len(slExt) = 0 Then
                slName = slName & "." & EXT
            Else
                bmIsError = True
                smErrorMessage = "The file must be have an '." & EXT & "' extension"
            End If
        End If
    Else
        bmIsError = True
        smErrorMessage = "Couldn't read filename."
    End If
    mGetName = slName
End Function

Private Sub mCreateExcel()
    Dim ilUsersSheets As Integer
    
On Error GoTo errbox
    Set myExcel = CreateObject("excel.application") 'New excel.Application
    'user's predefined # of sheets
    ilUsersSheets = myExcel.SheetsInNewWorkbook
    myExcel.SheetsInNewWorkbook = imSheets
    Set myWorkbook = myExcel.Workbooks.Add
    'give user back his default
    myExcel.SheetsInNewWorkbook = ilUsersSheets
    imIsExcel = ExcelExists
    Exit Sub
errbox:
    smErrorMessage = " Could not create Excel Objects."
    bmIsError = True
    imIsExcel = ExcelNotPresent
End Sub
Private Sub mClearError()
    smErrorMessage = ""
    bmIsError = False
End Sub
Private Sub mClearExcel()
    Set myExcel = Nothing
    Set myWorkbook = Nothing
End Sub
Private Sub Class_Terminate()
    Set myExcel = Nothing
    Set myWorkbook = Nothing
    Set myFile = Nothing
End Sub
