Attribute VB_Name = "Templates"
'********************************************************************************************************
'
'Doug Smith
'Created 08/28/09
'
'A handful of templates I created because I'm tired of writing them over and over or copying
'code that does the same thing, but has other code I still have to strip.
'
'Feel free to add any templates that you think are useful, but make sure that the are "completly" generic
'Don't include anything that won't compile as stands
'
'Template Listing:
'
'   mFunctionTemplate - create a new funtion with error trapping
'   mSubTemplate - create a new sub routine with error trapping
'   mReadDelimitedFile - reads in a comma delimited file and parses each field
'
'********************************************************************************************************


Private Function mFunctionTemplate() As Integer

    Dim ilRet As Integer
    
    On Error GoTo Err_Handler
    
    mFunctionTemplate = False


    mFunctionTemplate = True
    Exit Function
    
Err_Handler:
    gHandleError "AffErrorLog.txt", "modTemplates-mFunctionTemplate"
    mFunctionTemplate = False
    Exit Function
End Function

Private Sub mSubTemplate()

    Dim ilRet As Integer
    
    On Error GoTo Err_Handler
    
    Exit Sub
    
Err_Handler:
    gHandleError "AffErrorLog.txt", "modTemplates-mSubTemplate"
    Exit Sub
End Sub

Private Function mReadDelimitedFile(sFileName As String) As Boolean

    'Param
        'sFileName is the full path and filename with extension
    
    'Notes:
        'To change the delimiter change the Split call below frim "," to whatever you want
    
    'Local Vars
        'alRecordsArray contains all of the lines in the file -
        'The line's fileds are parsed into alRecordsArray(0), alRecordsArray(1), etc.
        'llLineCount is line the count
    
    Dim ilRet As Integer
    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject
    Dim slRetString As String
    Dim alRecordsArray() As String
    Dim llLineCount As Long
    
    On Error GoTo Err_Handler
    
    mReadDelimitedFile = False
    
    If fs.FileExists(sFileName) Then
        Set tlTxtStream = fs.OpenTextFile(sFileName, ForReading, False)
    Else
        MsgBox "** No Data Available **"
        Exit Function
    End If

    llLineCount = 0
    Do While tlTxtStream.AtEndOfStream <> True
        llLineCount = llLineCount + 1
        slRetString = tlTxtStream.ReadLine
        alRecordsArray = Split(slRetString, ",")
    Loop

    mReadDelimitedFile = True
    
    'Close up the files and exit!
    tlTxtStream.Close
    Exit Function
    
Err_Handler:
    gHandleError "AffErrorLog.txt", "modTemplates-mReadDelimitedFile"
    mReadDelimitedFile = False
    Exit Function
End Function


Private Function mReadDelimitedFileLookAhead() As Boolean

    'D.S. 07/07/09
    'This function opens a comma delimited, CSV file, that's passed in, then it uses 2 different file streams
    'reading the same file.  The records are automatically parsed into two different arrays,
    'alRecordsArray1 and alRecordsArray2.  Any element of the arrays can be tested against. array2
    'stays in front of the array1 by one record.  The purpose of the array2 is to act as a look ahead that
    'can be compared to array1.  This allow for easy groupings such as reading the file until a
    'given field changes, say the station or agreement code etc..  In this case the call letters are tested.
    'Could be handy in several other applications.

    'Assumes that there is a module level string, smFileName, that already has the path and filename

    'File IO
    Dim fs As New FileSystemObject
    Dim fs2 As New FileSystemObject
    Dim tlTxtStream As TextStream
    Dim tlTxtStream2 As TextStream

    Dim slRetString As String
    Dim slRetString2 As String
    Dim slTemp As String
    Dim alRecordsArray() As String
    Dim alRecordsArray2() As String
    Dim llIdx As Long
    Dim ilRet As Integer
    Dim ilStnCode As Integer
    Dim ilSameSta As Boolean

    On Error GoTo Err_Handler

    mReadDelimitedFileLookAhead = False

    'start debug
    'smFileName = "C:\sample data.csv"
    'End debug


    llIdx = 0
    If fs.FileExists(smFileName) Then
        Set tlTxtStream = fs.OpenTextFile(smFileName, ForReading, False)
    Else
        MsgBox "** No Data Available **"
        Exit Function
    End If

    If fs2.FileExists(smFileName) Then
        Set tlTxtStream2 = fs2.OpenTextFile(smFileName, ForReading, False)
    Else
        MsgBox "** No Data Available **"
        Exit Function
    End If

    'start off by reading the first line into the second array, the look ahead array2
    slRetString2 = tlTxtStream2.ReadLine
    alRecordsArray2 = Split(slRetString2, ",")
    'Loop while the first array1 has records
    Do While tlTxtStream.AtEndOfStream <> True
        llIdx = 0
        slRetString = tlTxtStream.ReadLine
        alRecordsArray = Split(slRetString, ",")
        ilSameSta = True
        'Get past any blank lines at the top as well as the header information
        If alRecordsArray(0) = "" Or alRecordsArray(0) = "CALL_LETTERS" Then
            ilSameSta = False
            slRetString2 = tlTxtStream2.ReadLine
        Else
            ilStnCode = gGetShttCodeFromCallLetters(alRecordsArray(0))
        End If

        If ilStnCode = 0 And ilSameSta = True Then
            'We don't have this station in SHTT so write it to a file and move array2 up one record
            Call gLogMsg("Station " & alRecordsArray(1) & " was not found.", "AffErrorLog.txt", False) 'station not found
            slRetString2 = tlTxtStream2.ReadLine
            alRecordsArray2 = Split(slRetString2, ",")
            ilSameSta = False
        End If

        If ilSameSta = True Then
            Do While ilSameSta
                'we now know the station exists
                'Debug
                'Call gLogMsg(cGreen, tlTxtStream.Line - 1 & " " & alRecordsArray(0) & " " & alRecordsArray(0) & " " & alRecordsArray(3), "AffErrorLog.txt", False, False, False)

                llIdx = llIdx + 1

                'test to see if we are at the end of the file for array2
                If tlTxtStream2.AtEndOfStream <> True Then
                    slRetString2 = tlTxtStream2.ReadLine
                    alRecordsArray2 = Split(slRetString2, ",")
                    'Get the line number - mainly for display or debug purposes
                    llLineNumber = tlTxtStream2.Line
                    DoEvents

                    'Debug
                    'If CLng(tlTxtStream2.Line) = 1000 Then
                    '    ilret = ilret
                    'End If

                    'are the stations still the same???
                    If (alRecordsArray(0) = alRecordsArray2(0)) Then
                        slRetString = tlTxtStream.ReadLine
                        alRecordsArray = Split(slRetString, ",")
                        ilStnCode = gGetShttCodeFromCallLetters(alRecordsArray(0))
                        ilSameSta = True
                    Else
                        'Call area to process the logical groups as they are found.
                        ilSameSta = False
                        Call gLogMsg("", "AffErrorLog.txt", False)
                    End If
                Else
                    'We reached the end of the file so get out, we're done
                    'Call area to process the logical groups as they are found.
                    ilSameSta = False
                End If

            Loop
            'Call area to process the logical groups as they are found.
        End If
    Loop

    'Close up the files and exit!
    tlTxtStream.Close
    tlTxtStream2.Close
    Exit Function

Err_Handler:
    'debug
    'Resume Next
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modTemplates-mReadDelimitedFileLookAhead"
End Function

