Attribute VB_Name = "modGenSubs"
'******************************************************
'*  modGenSubd - contains various general routines
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text
'5666
Public sgShortDate As String
Public bgIllegalCharsFound As Boolean
'6394
Public hgExportResult As Integer
Public Const MYSHORTDATE As String = "M/d/yy"
Const REG_SZ = 1 ' Unicode nul terminated string
Const REG_BINARY = 3
Const HKEY_CURRENT_USER = &H80000001
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long


Public gStopWatch As New StopWatch


'Spell Checker declares start

Private SpellCheck As Object
'9/15/11 for fCrExportViewer--like traffic
Public cgList As String    'Control associated with igFoundRow
Public igFoundRow As Integer 'Row found with gFndFirst or gFndNext


Public Const FRMNOMOVE = 2
Public Const FRMNOSIZE = 1
Public Const WNDNOTOPMOST = -2

Public Const SW_SHOWMINIMIZED = 2

Type RECT
    Left As Long
    Top As Long
    right As Long
    bottom As Long
End Type

Public Type POINTAPI
    X       As Long
    Y       As Long
End Type

Public Type WINDOWPLACEMENT
    LENGTH            As Long
    Flags             As Long
    showCmd           As Long
    ptMinPosition     As POINTAPI
    ptMaxPosition     As POINTAPI
    rcNormalPosition  As RECT
End Type

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function GetWindowPlacement Lib "user32" _
  (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Declare Function SetWindowPlacement Lib "user32" _
  (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

'End spell check declares

Private smDays() As String

Type DRIVEPATH
    sFolder As String
    iPos As Integer
End Type

Type CLFINFO
    lCode As Long
    lChfCode As Long
    sBillCycle As String * 1
    iLineNo As Integer
    iCntRevNo As Integer
    iPropVer As Integer
    sStartDate As String * 10
    sEndDate As String * 10
    sStartTime As String * 11
    sEndTime As String * 11
    iRdfCode As Integer
End Type
Public tmClfInfo As CLFINFO

Private dat_rst As ADODB.Recordset
'Private ast_rst As ADODB.Recordset
'Private lst_rst As ADODB.Recordset
Private sdf_rst As ADODB.Recordset
Private smf_rst As ADODB.Recordset
Private chf_rst As ADODB.Recordset
Private clf_rst As ADODB.Recordset
Private cff_rst As ADODB.Recordset
'Private rdf_rst As ADODB.Recordset
Private tmf_rst As ADODB.Recordset

Public bgDevEnv As Boolean
Private Declare Function GetModuleFileName Lib _
    "kernel32" Alias "GetModuleFileNameA" (ByVal _
    hModule As Long, ByVal lpFileName As String, _
    ByVal nSize As Long) As Long
Public sgArgs() As String

Public lgShellAndWaitID As Long

Public Sub gCommandArgs()
    sgArgs = Split(Command$, " ")
End Sub

Public Function gCmdLine(sCmd As String) As Boolean
    
    'D.S. 10/9/19 A more elegant and reliable method to check the command line parameters.
    'Not subject to ordering errors. i.e. autoimport and compelautoimport will return true with Instr calls.
    Dim ilLoop As Integer
    Dim slStr As String
    
    slStr = UCase(sCmd)
    gCmdLine = False
    If UBound(sgArgs) < 0 Then
        Exit Function
    End If
    For ilLoop = 0 To UBound(sgArgs)
        If UCase(sgArgs(ilLoop)) = Trim$(slStr) Then
            gCmdLine = True
            Exit For
        End If
   Next
End Function


Public Function gGetWebInterface(llAttCode As Long) As String
    
    'D.S. 07/08/10
  '  Dim att_rst As ADODB.Recordset
    '8583 restore 'cumulus/cbs' for v1 stations.  Note that this field was comandeered to handle v2 marketron.  cumulus/cbs for v1 will override setting for Marketron because agreement could be for both
    '7701-per Jim v70 and V81 only care about marketron
    gGetWebInterface = "NONE"
    If gIsVendorWithAgreement(llAttCode, Vendors.NetworkConnect) Then
        'D.S. 05/07/19 - TTP 9338 - Added if statement below
        If Not gIsWebVendor(22) Then
            gGetWebInterface = "LOGSONLY"
        End If
    End If
    '8842 don't care about version
   ' If gStationWebVersion(llAttCode) = 1 Then
        If gIsVendorWithAgreement(llAttCode, Vendors.stratus) Then
            gGetWebInterface = "Cumulus"
        ElseIf gIsVendorWithAgreement(llAttCode, Vendors.cBs) Then
            gGetWebInterface = "CBS"
        End If
    'End If
    
'    '6592 add CBS Dan M
'    'SQLQuery = "SELECT attWebInterface"
'    'D.S. 05-08-15 Added Marketron
'    SQLQuery = "SELECT attWebInterface,attExportToCBS,attExportToMarketron"
'    SQLQuery = SQLQuery + " FROM att"
'    SQLQuery = SQLQuery + " WHERE (attCode = " & llAttCode & ")"
'    Set att_rst = gSQLSelectCall(SQLQuery)
'
'    Select Case Trim$(att_rst!attWebInterface)
'        Case "C"
'            gGetWebInterface = "Cumulus"
'        Case "L"
'            gGetWebInterface = "LOGSONLY"
'        'Future Use
'        Case "??"
'            gGetWebInterface = "??"
'        Case Else
'            gGetWebInterface = "NONE"
'    End Select
'    '6592
'    If Trim$(att_rst!attExportToCBS) = "Y" Then
'        gGetWebInterface = "CBS"
'    End If
'    If Trim$(att_rst!attExportToMarketron) = "Y" Then
'        gGetWebInterface = "LOGSONLY"
'    End If
    
'    att_rst.Close

End Function
Public Function gStationWebVersion(llAttCode As Long) As Integer
    Dim ilRet As Integer
    Dim slSQLQuery As String
    Dim rst_Temp As ADODB.Recordset
    
    ilRet = 1
    slSQLQuery = "SELECT shttWebNumber from shtt inner join att on shttCode = attshfcode WHERE attcode = " & llAttCode
    Set rst_Temp = gSQLSelectCall(slSQLQuery)
    If Not rst_Temp.EOF Then
        If IsNumeric(rst_Temp!shttWebNumber) Then
            ilRet = rst_Temp!shttWebNumber
        End If
    End If
    gStationWebVersion = ilRet
End Function
Public Sub gChDrDir()
    If InStr(1, sgCurDir, ":") > 0 Then 'colon exists
        ChDrive Left$(sgCurDir, 2)  'Set the default drive
        ChDir sgCurDir
    End If
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:gCenterForm                     *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Center form within Traffic Form *
'*                                                     *
'*******************************************************
Sub gCenterForm(FrmName As Form)
'
'   gCenterForm FrmName
'   Where:
'       FrmName (I)- Name of modeless form to be centered within Traffic form
'
    Dim flLeft As Single
    Dim flTop As Single
    flLeft = frmMain.Left + (frmMain.Width - frmMain.ScaleWidth) / 2 + (frmMain.ScaleWidth - FrmName.Width) / 2
    flTop = frmMain.Top + (frmMain.Height - FrmName.Height) / 2 - 240
    FrmName.Move flLeft, flTop
End Sub


Public Sub gLogMsg(sMsg As String, sFileName As String, iKill As Integer)
    'D.S. 4/04
    'Purpose: A general file routine that shows: Date and Time followed by a message
    'so we can try to stop adding a separate file routine to every single module
    
    'Params
    'sMsg is the string to be written out
    'sFileName is the name of the file to be written to in the Messages directory
    'iKill = True then delete the file first, iKill = False then append to the file
    
    Dim slFullMsg As String
    Dim hlLogFile As Integer
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim slToFile As String
    
    slToFile = sgMsgDirectory & sFileName
    'On Error GoTo Error

    If iKill = True Then
        'ilRet = 0
        'slDateTime = FileDateTime(slToFile)
        ilRet = gFileExist(slToFile)
        If ilRet = 0 Then
            Kill slToFile
        End If
    End If
    
    'hlLogFile = FreeFile
    If sgUserName = "" Then
        sgUserName = "Unknown"
    End If
    'Open slToFile For Append As hlLogFile
    ilRet = gFileOpen(slToFile, "Append", hlLogFile)
    If ilRet = 0 Then
        slFullMsg = Format$(Now, "mm-dd-yyyy") & " " & Format$(Now, "hh:mm:ssam/pm") & " User: " & sgUserName & " - " & sMsg
        If sMsg = "" Then slFullMsg = "" 'A blank line
        Print #hlLogFile, slFullMsg
    End If
    Close hlLogFile
    
    slFullMsg = UCase(slFullMsg)
    If InStr(1, slFullMsg, "ERROR", vbTextCompare) > 0 Then
        sgTmfStatus = "E"
        gSaveStackTrace slToFile
    End If
    Exit Sub
    
'Error:
'    ilRet = 1
'    Resume Next
    
End Sub

Public Sub gLogMsgWODT(slAction As String, hlFileHandle As Integer, slMsg As String)

    'Add line to file without adding Date and time like gLogMsg
    
    'Params
    'slAction:  "ON" open as New (kill any previous version);
    '           "OD" open as New if not on todays date, otherwise append;
    '           "OA"=Open in append mode (retain previous version);
    '           "W"=Write message to file;
    '           "C" = Close handle
    'slMsg= If Open, Drive\Path\File name; If Write, message to write to file
    'hlFileHandle: If Open, return value; If Write or Close, handle of file to write to or close
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim blBypass As Boolean
    
    'On Error GoTo Error
    If slAction = "ON" Then 'Open as New
        ilRet = 0
        'slDateTime = FileDateTime(slMsg)
        ilRet = gFileExist(slMsg)
        If ilRet = 0 Then
            Kill slMsg
        End If
    End If
    '12/4/12: Add new option "OD", added for LogActivityFileName
    If slAction = "OD" Then 'Open as New if different date
        'ilRet = 0
        'slDateTime = FileDateTime(slMsg)
        ilRet = gFileExist(slMsg)
        If ilRet = 0 Then
            If gDateValue(Format(slDateTime, "m/d/yy")) <> gDateValue(Format(Now, "m/d/yy")) Then
                Kill slMsg
            End If
        End If
    End If
    '12/4/12: end of change
    
    Select Case Left$(slAction, 1)
        Case "O"
            'hlFileHandle = FreeFile
            'Open slMsg For Append Shared As hlFileHandle
            ilRet = gFileOpen(slMsg, "Append Shared", hlFileHandle)
        Case "W"
            If InStr(1, UCase$(slMsg), "ERROR", vbTextCompare) > 0 Then
                sgTmfStatus = "E"
            End If
            blBypass = False
            If hlFileHandle = hgSQLTrace Then
                If InStr(1, UCase$(slMsg), "RLF_RECORD", vbTextCompare) > 0 Then
                    blBypass = True
                End If
                If InStr(1, UCase$(slMsg), "FCT_FILE", vbTextCompare) > 0 Then
                    blBypass = True
                End If
                If InStr(1, UCase$(slMsg), "AUF_ALERT", vbTextCompare) > 0 Then
                    blBypass = True
                End If
                If InStr(1, UCase$(slMsg), "ABF_AST", vbTextCompare) > 0 Then
                    blBypass = True
                End If
            End If
            If Not blBypass Then
                Print #hlFileHandle, slMsg
            End If
        Case "C"
            Close hlFileHandle
            hlFileHandle = -1
    End Select
    Exit Sub
    
'Error:
'    ilRet = 1
'    Resume Next
    
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gListBoxHeight                  *
'*                                                     *
'*             Created:9/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute the height of a list    *
'*                     box                             *
'*                                                     *
'*******************************************************
Sub gSetListBoxHeight(lbcCtrl As ListBox, ilMaxRow As Integer)
'
'  flHeight = gListBoxHeight (ilNoRows, ilMaxRows)
'   Where:
'       ilNoRows (I) - current number of items within the list box
'       ilMaxRows (I) - max number of list box items to be displayed
'       flHeight (O) - height of list box in twips
'
    '+30 because of line above and below
    If lbcCtrl.ListCount > 0 Then
        'Determine standard height
        lbcCtrl.Height = 10
        
        If lbcCtrl.ListCount <= ilMaxRow Then
            lbcCtrl.Height = (lbcCtrl.Height - 30) * lbcCtrl.ListCount + 30 '375 + 255 * (ilNoRows - 1)
        Else
            lbcCtrl.Height = (lbcCtrl.Height - 30) * ilMaxRow + 30 '375 + 255 * (ilMaxRow - 1)
        End If
    End If
End Sub

Public Function gGetListBoxRow(lbcCtrl As ListBox, Y As Single) As Long
    Dim llRow As Long
    Dim slStr As String
    Dim llRowHeight As Long

    If lbcCtrl.ListCount <= 0 Then
        gGetListBoxRow = -1
        Exit Function
    End If
    llRowHeight = 15 * SendMessageByString(lbcCtrl.hwnd, LB_GETITEMHEIGHT, 0, slStr)
    llRow = (Y - 15) \ llRowHeight + lbcCtrl.TopIndex
    gGetListBoxRow = llRow
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gProcessArrowKey                *
'*                                                     *
'*             Created:9/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Process user arrow keys.       *
'*                      Used with combo boxes made up  *
'*                      from text control/button       *
'*                      control/list box control.      *
'*                                                     *
'*******************************************************
Sub gProcessArrowKey(ilShift As Integer, ilKeyCode As Integer, lbcCtrl As control, ilRetainState As Integer)
'
'   gProcessArrowKey Shift, KeyCode, lbcCtrl, imLbcArrowSetting
'   Where:
'       Shift (I)- Shift key state
'       KeyCode (I)- Key code
'       lbcCtrl (I)- list box control
'       ilLbcArrowSetting (I/O) - list box arrow setting flag
'                               True= make list box invisible (user click on item)
'                               False= retain list box visible state
'

    Dim ilState As Integer
    
    If (ilShift And ALTMASK) > 0 Then
        lbcCtrl.Visible = Not lbcCtrl.Visible
    ElseIf (ilShift And SHIFTMASK) > 0 Then
    Else
        ilState = lbcCtrl.Visible
        If ilKeyCode = KEYUP Then    'Up arrow
            If lbcCtrl.ListIndex > 0 Then
                lbcCtrl.ListIndex = lbcCtrl.ListIndex - 1
                If ilRetainState Then
                    lbcCtrl.Visible = ilState
                End If
            End If
        Else
            If lbcCtrl.ListIndex < lbcCtrl.ListCount - 1 Then
                lbcCtrl.ListIndex = lbcCtrl.ListIndex + 1
                If ilRetainState Then
                    lbcCtrl.Visible = ilState
                End If
            End If
        End If
    End If
End Sub

Public Function gRemoveChar(sInStr As String, sRemoveChar As String) As String

    Dim sOutStr As String
    Dim sChar As String
    Dim iLoop As Integer
    
    sOutStr = ""
    If IsNull(sInStr) <> True Then
        For iLoop = 1 To Len(sInStr) Step 1
            sChar = Mid$(sInStr, iLoop, 1)
            If sChar <> sRemoveChar Then
                sOutStr = sOutStr & sChar
            End If
        Next iLoop
    End If
    gRemoveChar = sOutStr
End Function

Public Function gTestForMultipleEmail(sEmailAddress As String, sRegEmailOrBCC As String) As Integer

    'D.S. 07/29/05
    'Create an array of email addresses if needed and provide a basic sanity check
    
    Dim ilPos As Integer
    Dim ilStart As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slEmailAddressArray() As String
    
    gTestForMultipleEmail = False
    
    'Don't force them to have a BCC
    If sEmailAddress = "" And sRegEmailOrBCC = "BCC" Then
        gTestForMultipleEmail = True
        Exit Function
    End If
    
    If sEmailAddress = "" And sRegEmailOrBCC <> "BCC" Then
'        gMsgBox "Warning: No email address has been defined for the Administrator."
'        sgErrorMsg = "Warning: No email address has been defined for the Administrator."
        gTestForMultipleEmail = True
        Exit Function
    End If
    
    'Check to see if the string has multiple email addresses
    ilPos = InStr(1, sEmailAddress, ",", vbTextCompare)
    If ilPos > 0 Then
        ilStart = 1
        ReDim slEmailAddressArray(0 To 0) As String
        slEmailAddressArray(0) = Trim$(Mid$(sEmailAddress, ilStart, ilPos - 1))
        ReDim Preserve slEmailAddressArray(0 To UBound(slEmailAddressArray) + 1)
        For ilLoop = ilPos To Len(sEmailAddress) - 1 Step 1
            ilStart = ilPos + 1
            ilPos = InStr(ilStart, sEmailAddress, ",", vbTextCompare)
            If ilPos > 0 Then
                slEmailAddressArray(UBound(slEmailAddressArray)) = Trim$(Mid$(sEmailAddress, ilStart, ilPos - ilStart))
                ReDim Preserve slEmailAddressArray(0 To UBound(slEmailAddressArray) + 1)
                ilLoop = ilPos
            Else
                slEmailAddressArray(UBound(slEmailAddressArray)) = Trim$(Mid$(sEmailAddress, ilStart, Len(sEmailAddress) - (ilStart - 1)))
                ReDim Preserve slEmailAddressArray(0 To UBound(slEmailAddressArray) + 1)
                Exit For
            End If
        Next ilLoop
    Else
        ReDim slEmailAddressArray(0 To 1) As String
        slEmailAddressArray(0) = Trim$(sEmailAddress)
    End If
    
    If Not mTestForValidEmailAddress(slEmailAddressArray()) Then
        gTestForMultipleEmail = False
        Exit Function
    End If
    
    gTestForMultipleEmail = True
    
End Function

Private Function mTestForValidEmailAddress(sEmailAddressArray() As String) As Integer
    '9938 made private
    'D.S. 07/29/05
    'Take an array of email addresses and provide a basic sanity check
    
    Dim ilLoop As Integer
    Dim ilResult As Integer
    Dim slTemp As String
    Dim ilIdx As Integer
    Dim slEMsg As String
    Dim slExt As String
    Dim ilMax As Integer
    Dim ilYesNo As Integer
    Dim slEmailAddress As String
    
    For ilLoop = 0 To UBound(sEmailAddressArray) - 1
        mTestForValidEmailAddress = False
        ilResult = Len(sEmailAddressArray(ilLoop))
        If ilResult < 6 Then
            sgErrorMsg = "ERROR: " & """" & sEmailAddressArray(ilLoop) & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Your email address is shorter than 6 characters which is impossible."
            gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
            Exit Function
        End If
        '9938 made reverse
        ilResult = InStrRev(sEmailAddressArray(ilLoop), ".")
        If ilResult = 0 Then
            sgErrorMsg = "ERROR: " & """" & sEmailAddressArray(ilLoop) & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address does not contain a period"
            gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
            Exit Function
        ElseIf ilResult = Len(sEmailAddressArray(ilLoop)) Then
            sgErrorMsg = "ERROR: " & """" & sEmailAddressArray(ilLoop) & """" & " is not a valid email address." & Chr(13) & Chr(10) & "A period cannot be the last character"
            gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
            Exit Function
        End If
        
        ilResult = InStr(1, sEmailAddressArray(ilLoop), "@", vbTextCompare)
        If ilResult = 0 Then
            sgErrorMsg = "ERROR: " & """" & sEmailAddressArray(ilLoop) & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address does not contain an @"
            gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
            Exit Function
        End If
        
        ilResult = InStr(1, sEmailAddressArray(ilLoop), ";", vbTextCompare)
        If ilResult <> 0 Then
            sgErrorMsg = "ERROR: " & """" & sEmailAddressArray(ilLoop) & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address contains a semicolon"
            gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
            Exit Function
        End If
        
        ilResult = InStr(1, sEmailAddressArray(ilLoop), "@", vbTextCompare)
        If ilResult = 1 Then
            sgErrorMsg = "ERROR: " & """" & sEmailAddressArray(ilLoop) & """" & " is not a valid email address." & Chr(13) & Chr(10) & "@ Cannot be the first character"
            gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
            Exit Function
        End If
        
        ilResult = InStr(1, sEmailAddressArray(ilLoop), "@", vbTextCompare)
        If ilResult = Len(sEmailAddressArray(ilLoop)) Then
            sgErrorMsg = "ERROR: " & """" & sEmailAddressArray(ilLoop) & """" & " is not a valid email address." & Chr(13) & Chr(10) & "@ Cannot be the last character"
            gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
            Exit Function
        End If
        
        ilResult = InStr(1, sEmailAddressArray(ilLoop), "..", vbTextCompare)
        If ilResult <> 0 Then
            sgErrorMsg = "ERROR: " & """" & sEmailAddressArray(ilLoop) & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address contains two or more periods together"
            gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
            Exit Function
        End If
        
        ilResult = InStr(1, sEmailAddressArray(ilLoop), " ", vbTextCompare)
        If ilResult <> 0 Then
            sgErrorMsg = "ERROR: " & """" & sEmailAddressArray(ilLoop) & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address contains a space"
            gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
            Exit Function
        End If
        
        ilResult = InStr(1, sEmailAddressArray(ilLoop), "[", vbTextCompare)
        If ilResult <> 0 Then
            sgErrorMsg = "ERROR: " & """" & sEmailAddressArray(ilLoop) & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address contains a [ character"
            gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
            Exit Function
        End If
        
        ilResult = InStr(1, sEmailAddressArray(ilLoop), "]", vbTextCompare)
        If ilResult <> 0 Then
            sgErrorMsg = "ERROR: " & """" & sEmailAddressArray(ilLoop) & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address contains a ] character"
            gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
            Exit Function
        End If
        
        slTemp = sEmailAddressArray(ilLoop)
        Do While InStr(1, slTemp, "@") <> 0
           ilIdx = 1
           slTemp = right(slTemp, Len(slTemp) - InStr(1, slTemp, "@"))
        Loop
        If ilIdx > 1 Then
            sgErrorMsg = "ERROR: " & """" & sEmailAddressArray(ilLoop) & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address contains more than one @ sign"
            gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
            Exit Function
        End If
        'what is this doing?
        slExt = sEmailAddressArray(ilLoop)
        Do While InStr(1, slExt, ".") <> 0
            slExt = right(slExt, Len(slExt) - InStr(1, slExt, "."))
        Loop
        '9938 test for screwy characters
        slEmailAddress = sEmailAddressArray(ilLoop)
        For ilIdx = Len(slEmailAddress) To 1 Step -1
            If Asc(Mid(slEmailAddress, ilIdx, 1)) > 126 Then
                sgErrorMsg = "ERROR: " & """" & sEmailAddressArray(ilLoop) & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address contains an ASCII character greater than 126"
                gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
                Exit Function
            End If
        Next ilIdx
        'D.S. 5/12/15 Extensions have been opened up to where almost anything is legal.  No longer checking fro extensions
'        If gEmailsExtIsOK(slExt) <> True Then
'            slEMsg = slEMsg & " Your email address does not appear to be carrying a valid ending!"
'            slEMsg = slEMsg & " It must be one of the following..."
'            slEMsg = slEMsg & " .info, .com, .net, .gov, .org, .edu, .biz, .coop, .tv Or included country's assigned country code"
'            sgErrorMsg = "ERROR: " & """" & sEmailAddressArray(ilLoop) & """" & slEMsg
'            gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
'            gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
'            ilYesNo = gMsgBox(sgErrorMsg & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Do you want continue saving this email (Y/N)?", vbYesNo)
'            If ilYesNo = vbYes Then
'                gLogMsg "User accepted the email.", "WebEmailLog.Txt", False
'            Else
'                gLogMsg "User did not accept the email.", "WebEmailLog.Txt", False
'            Exit Function
'           End If
'        End If
        
        mTestForValidEmailAddress = True
        
    Next ilLoop

End Function

Public Function gTestForSingleValidEmailAddress(sEmailAddress As String) As Integer

    'D.S. 11/20/08
    'Take an array of email addresses and provide a basic sanity check
    
    Dim ilResult As Integer
    Dim slTemp As String
    Dim ilIdx As Integer
    Dim slEMsg As String
    Dim slExt As String
    Dim ilYesNo As Integer
    
    gTestForSingleValidEmailAddress = False
    ilResult = Len(sEmailAddress)
    If ilResult < 6 Then
        sgErrorMsg = "ERROR: " & """" & sEmailAddress & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Your email address is shorter than 6 characters which is impossible."
        gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
        Exit Function
    End If
    
    ilResult = InStrRev(1, sEmailAddress, ".")
    If ilResult = 0 Then
        sgErrorMsg = "ERROR: " & """" & sEmailAddress & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address does not contain a period"
        gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
        Exit Function
    ElseIf ilResult = Len(sEmailAddress) Then
        sgErrorMsg = "ERROR: " & """" & sEmailAddress & """" & " is not a valid email address." & Chr(13) & Chr(10) & "A period cannot be the last character"
        gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
        Exit Function
    End If
    
    ilResult = InStr(1, sEmailAddress, "@", vbTextCompare)
    If ilResult = 0 Then
        sgErrorMsg = "ERROR: " & """" & sEmailAddress & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address does not contain an @"
        gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
        Exit Function
    End If
    
    ilResult = InStr(1, sEmailAddress, ";", vbTextCompare)
    If ilResult <> 0 Then
        sgErrorMsg = "ERROR: " & """" & sEmailAddress & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address contains a semicolon"
        gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
        Exit Function
    End If
    
    ilResult = InStr(1, sEmailAddress, "@", vbTextCompare)
    If ilResult = 1 Then
        sgErrorMsg = "ERROR: " & """" & sEmailAddress & """" & " is not a valid email address." & Chr(13) & Chr(10) & "@ Cannot be the first character"
        gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
        Exit Function
    End If
    
    ilResult = InStr(1, sEmailAddress, "@", vbTextCompare)
    If ilResult = Len(sEmailAddress) Then
        sgErrorMsg = "ERROR: " & """" & sEmailAddress & """" & " is not a valid email address." & Chr(13) & Chr(10) & "@ Cannot be the last character"
        gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
        Exit Function
    End If
    
    ilResult = InStr(1, sEmailAddress, "..", vbTextCompare)
    If ilResult <> 0 Then
        sgErrorMsg = "ERROR: " & """" & sEmailAddress & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address contains two or more periods together"
        gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
        Exit Function
    End If
    
    ilResult = InStr(1, sEmailAddress, " ", vbTextCompare)
    If ilResult <> 0 Then
        sgErrorMsg = "ERROR: " & """" & sEmailAddress & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address contains a space"
        gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
        Exit Function
    End If
    
    ilResult = InStr(1, sEmailAddress, "[", vbTextCompare)
    If ilResult <> 0 Then
        sgErrorMsg = "ERROR: " & """" & sEmailAddress & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address contains a [ character"
        gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
        Exit Function
    End If
    
    ilResult = InStr(1, sEmailAddress, "]", vbTextCompare)
    If ilResult <> 0 Then
        sgErrorMsg = "ERROR: " & """" & sEmailAddress & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address contains a ] character"
        gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
        Exit Function
    End If
    
    slTemp = sEmailAddress
    Do While InStr(1, slTemp, "@") <> 0
       ilIdx = 1
       slTemp = right(slTemp, Len(slTemp) - InStr(1, slTemp, "@"))
    Loop
    If ilIdx > 1 Then
        sgErrorMsg = "ERROR: " & """" & sEmailAddress & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address contains more than one @ sign"
        gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
        Exit Function
    End If
    '9938 test for screwy characters
    slTemp = sEmailAddress
    For ilIdx = Len(slTemp) To 1 Step -1
        If Asc(Mid(slTemp, ilIdx, 1)) > 126 Then
            sgErrorMsg = "ERROR: " & """" & sEmailAddress & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Address contains an ASCII character greater than 126"
            gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
            Exit Function
        End If
    Next ilIdx

    slExt = sEmailAddress
    Do While InStr(1, slExt, ".") <> 0
        slExt = right(slExt, Len(slExt) - InStr(1, slExt, "."))
    Loop


    If gEmailsExtIsOK(slExt) <> True Then
        slEMsg = slEMsg & " Your email address is not carrying a valid ending!"
        slEMsg = slEMsg & " It must be one of the following..."
        slEMsg = slEMsg & " .info, .com, .net, .gov, .org, .edu, .biz, .coop, .tv Or included country's assigned country code."
        sgErrorMsg = "ERROR: " & """" & sEmailAddress & """" & slEMsg
        gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
        ilYesNo = gMsgBox(sgErrorMsg & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Do you want continue saving this email (Y/N)?", vbYesNo)
        If ilYesNo = vbYes Then
            gLogMsg "User accepted the email.", "WebEmailLog.Txt", False
        Else
            gLogMsg "User did not accept the email.", "WebEmailLog.Txt", False
        Exit Function
    End If
    End If
    
    gTestForSingleValidEmailAddress = True
End Function


Public Function gEmailsExtIsOK(ByVal sEXT As String) As Boolean

    'D.S. 07/29/05
    'Check for valid email Extensions
    
    Dim slExt As String

    gEmailsExtIsOK = False
    If Left(sEXT, 1) <> "." Then
        sEXT = "." & sEXT
    End If
    sEXT = UCase(sEXT) 'just to avoid errors
    slExt = slExt & ".COM.COOP.EDU.GOV.NET.BIZ.ORG.TV.INFO"
    slExt = slExt & ".AF.AL.DZ.As.AD.AO.AI.AQ.AG.AP.AR.AM.AW.AU.AT.AZ.BS.BH.BD.BB.BY"
    slExt = slExt & ".BE.BZ.BJ.BM.BT.BO.BA.BW.BV.BR.IO.BN.BG.BF.MM.BI.KH.CM.CA.CV.KY"
    slExt = slExt & ".CF.TD.CL.CN.CX.CC.CO.KM.CG.CD.CK.CR.CI.HR.CU.CY.CZ.DK.DJ.DM.DO"
    slExt = slExt & ".TP.EC.EG.SV.GQ.ER.EE.ET.FK.FO.FJ.FI.CS.SU.FR.FX.GF.PF.TF.GA.GM.GE.DE"
    slExt = slExt & ".GH.GI.GB.GR.GL.GD.GP.GU.GT.GN.GW.GY.HT.HM.HN.HK.HU.IS.IN.ID.IR.IQ"
    slExt = slExt & ".IE.IL.IT.JM.JP.JO.KZ.KE.KI.KW.KG.LA.LV.LB.LS.LR.LY.LI.LT.LU.MO.MK.MG"
    slExt = slExt & ".MW.MY.MV.ML.MT.MH.MQ.MR.MU.YT.MX.FM.MD.MC.MN.MS.MA.MZ.NA"
    slExt = slExt & ".NR.NP.NL.AN.NT.NC.NZ.NI.NE.NG.NU.NF.KP.MP.NO.OM.PK.PW.PA.PG.PY"
    slExt = slExt & ".PE.PH.PN.PL.PT.PR.QA.RE.RO.RU.RW.GS.SH.KN.LC.PM.ST.VC.SM.SA.SN.SC"
    slExt = slExt & ".SL.SG.SK.SI.SB.SO.ZA.KR.ES.LK.SD.SR.SJ.SZ.SE.CH.SY.TJ.TW.TZ.TH.TG.TK"
    slExt = slExt & ".TO.TT.TN.TR.TM.TC.TV.UG.UA.AE.UK.US.UY.UM.UZ.VU.VA.VE.VN.VG.VI"
    slExt = slExt & ".WF.WS.EH.YE.YU.ZR.ZM.ZW"
    sEXT = UCase(sEXT) 'just to avoid errors
    If InStr(1, slExt, sEXT, vbBinaryCompare) <> 0 Then
        gEmailsExtIsOK = True
    End If

End Function

Public Sub gBuildCPYRotCom(lCSFCode As Long, sCSFComment As String)

    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilUpper As Integer
    
    ilUpper = UBound(tgCopyRotInfo)
    For ilLoop = 0 To ilUpper Step 1
        If tgCopyRotInfo(ilLoop).lCode = lCSFCode Then
            Exit Sub
        End If
    Next ilLoop
    tgCopyRotInfo(ilUpper).lCode = lCSFCode
    tgCopyRotInfo(ilUpper).sComment = sCSFComment
    ReDim Preserve tgCopyRotInfo(0 To ilUpper + 1)

    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occurred in frmGenSubs-gBuildCPYRotCom: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "WebExportLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If

End Sub
Public Function gGeneratePassword(iLen As Integer, iPwdStrength) As String
    
' gGeneratePassword returns a string of random letters equal to iLen
' Strength of password 1 - 4
' The string will contain at least
' 1 upper case letter A - Z
' 1 lower case letter a - z
' 1 number 0 - 9
' 1 special chacter of #, $, %, &
' The length of the password must be at least 3 for this to be true.
    
    On Error GoTo Err_gGeneratePassword
    Dim ilIdx, ilSeed As Integer
    Dim slPassword As String
    Dim lDateTime As Date
    Dim dlRNum As Double
    Dim blSpecialCharAssigned As Boolean
    Dim ThisChar As String * 1
 
    ' Seed the random # generator so a different password is generated each time.
    lDateTime = Now()
    ilSeed = Int(Second(lDateTime) * Minute(lDateTime) * Hour(lDateTime))
    Randomize (ilSeed)
 
    For ilIdx = 1 To iLen
        Select Case (ilSeed + ilIdx) Mod iPwdStrength  ' Cycle between the four different letter types, starting at random.
            Case 0
                ThisChar = Chr((Int(Rnd() * 26) + Asc("A")))
            Case 1
                ThisChar = Chr((Int(Rnd() * 26) + Asc("a")))
            Case 2
                ThisChar = Chr((Int(Rnd() * 10) + Asc("0")))
            Case 3
                If blSpecialCharAssigned Then
                    ThisChar = Chr((Int(Rnd() * 26) + Asc("A")))
                Else
                    ' Only allow 1 special character.
                    ThisChar = Chr((Int(Rnd() * 4) + Asc("#")))
                    blSpecialCharAssigned = True
                End If
        End Select
        slPassword = slPassword + ThisChar
    Next
    gGeneratePassword = slPassword
    Exit Function
 
Err_gGeneratePassword:
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetSyncDateTime                *
'*                                                     *
'*             Created:10/22/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get sync Date and Time          *
'*                                                     *
'*******************************************************
Sub gGetSyncDateTime(slSyncDate As String, slSyncTime As String)
    Dim hlFile As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim slInitStamp As String
    
    If sgSpfRemoteUsers = "Y" Then
        ilRet = 0
        On Error GoTo gGetSyncDateTimeErr:
        'If tgUrf(0).iRemoteUserID <= 0 Then
        '    'Get Central date and time (i.e. if Dallas, get it from NY)
        '    hlFile = FreeFile
        '    If Trim$(sgMDBPath) <> "" Then
        '    slInitStamp = FileDateTime(sgMDBPath & "RUStamp.Txt")
        '    Do
        '        ilRet = 0
        '        Open sgMDBPath & "RUStamp.Txt" For Output Shared As hlFile
        '        Print #hlFile, "Date Time Stamp File"
        '        Close #hlFile
        '        slStamp = FileDateTime(sgMDBPath & "RUStamp.Txt")
        '    Loop While (ilRet <> 0) Or (StrComp(slStamp, slInitStamp, 1) = 0)
        '    Else
        '    slInitStamp = FileDateTime(sgDBPath & "RUStamp.Txt")
        '    Do
        '        ilRet = 0
        '        Open sgDBPath & "RUStamp.Txt" For Output Shared As hlFile
        '        Print #hlFile, "Date Time Stamp File"
        '        Close #hlFile
        '        slStamp = FileDateTime(sgDBPath & "RUStamp.Txt")
        '    Loop While (ilRet <> 0) Or (StrComp(slStamp, slInitStamp, 1) = 0)
        '    End If
        '    slSyncDate = Format$(slStamp, sgShowDateForm)
        '    slSyncTime = Format$(slStamp, sgShowTimeWSecForm)
        'Else
            'Get Local date and time
            slSyncDate = Format$(gNow(), sgShowDateForm)
            slSyncTime = Format$(gNow(), sgShowTimeWSecForm)
        'End If
        On Error GoTo 0
    Else
        'slSyncDate = ""
        'slSyncTime = ""
        slSyncDate = "1/1/1970"
        slSyncTime = "12:00:00AM"
    End If
    Exit Sub
gGetSyncDateTimeErr:
    ilRet = Err
    Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gParseItem                      *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain a substring from a string*
'*                                                     *
'*******************************************************
Function gParseItem(ByVal slInputStr As String, ByVal ilItemNo As Integer, slDelimiter As String, slOutputStr As String) As Integer
'
'   iRet = gParseItem(slInputStr, ilItemNo, slDelimiter, slOutStr)
'   Where:
'       slInputStr (I)-string from which to obtain substring
'       ilItemNo (I)-substring number to obtain (first string is Item number 1)
'       slDelimiter (I)-delimiter string or character between strings
'       slOutStr (O)-substring
'       iRet =  TRUE if substring found, FALSE if substring not found
'

    Dim ilEndPos As Integer  'Enp position of substring within sInputStr
    Dim ilStartPos As Integer    'Start position of each substring
    Dim ilIndex As Integer   'For loop parameter
    Dim ilLen As Integer 'Length of string to be parsed
    Dim ilDelimiterLen As Integer    'Delimiter length

    ilLen = Len(slInputStr)   'Obtain length so start position will not exceed length
    ilDelimiterLen = Len(slDelimiter)
    ilStartPos = 1   'Initialize start position
    For ilIndex = 1 To ilItemNo - 1 Step 1    'Loop until at starting position of substring to be found
        ilStartPos = InStr(ilStartPos, slInputStr, slDelimiter, 1) + ilDelimiterLen
        If (ilStartPos = ilDelimiterLen) Or (ilStartPos > ilLen) Then
            slOutputStr = ""
            gParseItem = CSI_MSG_PARSE
            Exit Function
        End If
    Next ilIndex
    ilEndPos = InStr(ilStartPos, slInputStr, slDelimiter, 1)   'Position end to end of substring plus 1 (start of delimiter position)
    If (ilEndPos = 0) Then   'No end delimiter-copy reTrafficing string
        slOutputStr = Trim$(Mid$(slInputStr, ilStartPos))
        gParseItem = CSI_MSG_NONE
        Exit Function
    End If
    slOutputStr = Trim$(Mid$(slInputStr, ilStartPos, ilEndPos - ilStartPos))
    gParseItem = CSI_MSG_NONE
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:gParseCDFields                  *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Parse comma delimited fields    *
'*                     Note:including quotes that are  *
'*                     enclosed within quotes          *
'*                     ""xxxxxxxx"","xxxxx",           *
'*                                                     *
'*******************************************************
Sub gParseCDFields(slCDStr As String, ilLower As Integer, slFields() As String)
'
'   gParseCDFields slCDStr, ilLower, slFields()
'   Where:
'       slCDStr(I)- Comma delinited string
'       ilLower(I)- True=Convert string fields characters to lower case (preceding character is A-Z)
'       slFields() (O)- fields parsed from comma delimited string
'
    Dim ilFieldNo As Integer
    Dim ilFieldType As Integer  '0=String, 1=Number
    Dim slChar As String
    Dim ilIndex As Integer
    Dim ilAscChar As Integer
    Dim ilAddToStr As Integer
    Dim slNextChar As String

    For ilIndex = LBound(slFields) To UBound(slFields) Step 1
        slFields(ilIndex) = ""
    Next ilIndex
    'ilFieldNo = 1
    ilFieldNo = LBound(slFields)
    ilIndex = 1
    ilFieldType = -1
    Do While ilIndex <= Len(Trim$(slCDStr))
        slChar = Mid$(slCDStr, ilIndex, 1)
        If ilFieldType = -1 Then
            If slChar = "," Then    'Comma was followed by a comma-blank field
                ilFieldType = -1
                ilFieldNo = ilFieldNo + 1
                If ilFieldNo > UBound(slFields) Then
                    Exit Sub
                End If
            ElseIf slChar <> """" Then
                ilFieldType = 1
                slFields(ilFieldNo) = slChar
            Else
                ilFieldType = 0 'Quote field
            End If
        Else
            If ilFieldType = 0 Then 'Started with a Quote
                'Add to string unless "
                ilAddToStr = True
                If slChar = """" Then
                    If ilIndex = Len(Trim$(slCDStr)) Then
                        ilAddToStr = False
                    Else
                        slNextChar = Mid$(slCDStr, ilIndex + 1, 1)
                        If slNextChar = "," Then
                            ilAddToStr = False
                        End If
                    End If
                End If
                If ilAddToStr Then
                    If (slFields(ilFieldNo) <> "") And ilLower Then
                        ilAscChar = Asc(UCase(right$(slFields(ilFieldNo), 1)))
                        If ((ilAscChar >= Asc("A")) And (ilAscChar <= Asc("Z"))) Then
                            slChar = LCase$(slChar)
                        End If
                    End If
                    slFields(ilFieldNo) = slFields(ilFieldNo) & slChar
                Else
                    ilFieldType = -1
                    ilFieldNo = ilFieldNo + 1
                    If ilFieldNo > UBound(slFields) Then
                        Exit Sub
                    End If
                    ilIndex = ilIndex + 1   'bypass comma
                End If
            Else
                'Add to string unless ,
                If slChar <> "," Then
                    slFields(ilFieldNo) = slFields(ilFieldNo) & slChar
                Else
                    ilFieldType = -1
                    ilFieldNo = ilFieldNo + 1
                    If ilFieldNo > UBound(slFields) Then
                        Exit Sub
                    End If
                End If
            End If
        End If
        ilIndex = ilIndex + 1
    Loop
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name: gCtrlGotFocus                  *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Text in the control is          *
'*                     highlighted                     *
'*                                                     *
'*******************************************************
Public Sub gCtrlGotFocus(Ctrl As control)
'
'   gCtrlGotFocus Ctrl
'   Where:
'       Ctrl (I)- control for which text will be highlighted
'

'    Traffic.plcHelp.Caption = " " & Ctrl.Tag
    If TypeOf Ctrl Is TextBox Then
        Ctrl.SelStart = 0
        Ctrl.SelLength = Len(Ctrl.Text)
    'ElseIf TypeOf Ctrl Is MaskEdBox Then
    '    Ctrl.SelStart = 0
    '    Ctrl.SelLength = Len(Ctrl.Text)
    'ElseIf TypeOf Ctrl Is SSCombo Then
    '    Ctrl.SelStart = 0
    '    Ctrl.SelLength = Len(Ctrl.Text)
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gFadeForm                       *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Fade background color of a form *
'*                                                     *
'*******************************************************
Public Sub gFadeForm(frm As Form, ilRed As Integer, ilGreen As Integer, ilBlue As Integer)
    Dim ilSaveScale As Integer
    Dim ilSaveStyle As Integer
    Dim ilSaveRedraw As Integer

    Dim llI As Long
    Dim llJ As Long
    Dim llX As Long
    Dim llY As Long
    Dim ilPixels As Integer

    'Save current values
    ilSaveScale = frm.ScaleMode
    ilSaveStyle = frm.DrawStyle
    ilSaveRedraw = frm.AutoRedraw

    'Paint Screen
    frm.ScaleMode = 3
    ilPixels = Screen.Height / Screen.TwipsPerPixelY
    llX = ilPixels / 64# + 0.5
    frm.DrawStyle = 5
    frm.AutoRedraw = True
    For llJ = 0 To ilPixels Step llX
        llY = 240 - 245 * llJ \ ilPixels
        'can tweak this to preference.
        If llY < 0 Then
            llY = 0
        End If
        frm.Line (-2, llJ - 2)-(Screen.Width + 2, llJ + llX + 3), RGB(-ilRed * llY, -ilGreen * llY, -ilBlue * llY), BF
    Next llJ

    'Reset
    frm.ScaleMode = ilSaveScale
    frm.DrawStyle = ilSaveStyle
    frm.AutoRedraw = ilSaveRedraw
End Sub



Public Sub gShellAndWait(sExe As String)
    Dim lProcessId As Long
    Dim hProcess As Long
    Dim lExitCode As Long
    Dim lRet As Long
    Dim fStart As Single
    
    On Error GoTo ErrHand
    'Pause required for NT to work
    fStart = Timer
    Do While Timer < fStart + 5
        DoEvents
    Loop
    lgShellAndWaitID = Shell(sExe, vbNormalFocus)
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, lgShellAndWaitID)
    Do
        lRet = GetExitCodeProcess(hProcess, lExitCode)
        DoEvents
    Loop While (lExitCode = STILL_ACTIVE)
    lgShellAndWaitID = 0
    lRet = CloseHandle(hProcess)
    fStart = Timer
    Do While Timer < fStart + 5
        DoEvents
    Loop
    On Error GoTo 0
    Exit Sub
ErrHand:
    gMsgBox "Shell Error " & Str$(Err.Number) & Err.Description, vbOKOnly
    On Error GoTo 0
    Exit Sub
End Sub

Public Function gFixQuote(sInStr As String) As String

    Dim sOutStr As String
    Dim sChar As String
    Dim iLoop As Integer
    
    sOutStr = ""
    If IsNull(sInStr) <> True Then
        For iLoop = 1 To Len(sInStr) Step 1
            sChar = Mid$(sInStr, iLoop, 1)
            If sChar = "'" Then
                sOutStr = sOutStr & "''"
            Else
                sOutStr = sOutStr & sChar
            End If
        Next iLoop
    End If
    gFixQuote = sOutStr
End Function

Public Function gFixDoubleQuote(sInStr As String) As String

    Dim sOutStr As String
    Dim sChar As String
    Dim iLoop As Integer
    
    sOutStr = ""
    If IsNull(sInStr) <> True Then
        For iLoop = 1 To Len(sInStr) Step 1
            sChar = Mid$(sInStr, iLoop, 1)
            If sChar = """" Then
                sOutStr = sOutStr & "''"
            Else
                sOutStr = sOutStr & sChar
            End If
        Next iLoop
    End If
    gFixDoubleQuote = sOutStr
End Function
Public Function gFixDoubleQuoteWithSingle(sInStr As String) As String

    Dim sOutStr As String
    Dim sChar As String
    Dim iLoop As Integer
    
    sOutStr = ""
    If IsNull(sInStr) <> True Then
        For iLoop = 1 To Len(sInStr) Step 1
            sChar = Mid$(sInStr, iLoop, 1)
            If sChar = """" Then
                sOutStr = sOutStr & "'"
            Else
                sOutStr = sOutStr & sChar
            End If
        Next iLoop
    End If
    gFixDoubleQuoteWithSingle = sOutStr
End Function

Public Function gStripComma(sInStr As String) As String

    Dim sOutStr As String
    Dim sChar As String
    Dim iLoop As Integer
    
    sOutStr = ""
    If IsNull(sInStr) <> True Then
        For iLoop = 1 To Len(sInStr) Step 1
            sChar = Mid$(sInStr, iLoop, 1)
            If sChar = "," Then
                sOutStr = sOutStr & " "
            Else
                sOutStr = sOutStr & sChar
            End If
        Next iLoop
    End If
    gStripComma = sOutStr
End Function
Public Function gStripCommaDG(sInStr As String) As String

    Dim sOutStr As String
    Dim sChar As String
    Dim sChar2 As String
    Dim iLoop As Integer

    On Error GoTo ErrHand

    sOutStr = ""
    If IsNull(sInStr) <> True Then
        For iLoop = 1 To Len(sInStr) Step 1
            sChar = Mid$(sInStr, iLoop, 1)
            sChar2 = Mid$(sInStr, iLoop + 1, 1)
            If iLoop < Len(sInStr) - 1 Then
                If sChar = "," And Asc(sChar2) <> 34 Then
                        sOutStr = sOutStr & " "
                Else
                    sOutStr = sOutStr & sChar
                End If
            Else
                If sChar = "," Then
                        sOutStr = sOutStr & " "
                Else
                    sOutStr = sOutStr & sChar
                End If
            End If
        Next iLoop
    End If
    gStripCommaDG = sOutStr
    
Exit Function
    
ErrHand:

    
End Function



''*******************************************************
''*                                                     *
''*      Procedure Name:gSQLWait                        *
''*                                                     *
''*             Created:4/12/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments:Execute Insert or Update or     *
''*                     Delete SQL operation            *
''*                                                     *
''*     gSQLWaitNoMsgBox should be used instead of this *
''*     routine                                         *
''*******************************************************
'Public Function gSQLWait(sSQLQuery As String, iDoTrans As Integer) As Integer
'    Dim iRet As Integer
'    Dim fStart As Single
'    Dim iCount As Integer
'    Dim hlMsg As Integer
'
'    On Error GoTo ErrHand
'    iCount = 0
'    Do
'        iRet = 0
'        If iDoTrans Then
'            cnn.BeginTrans
'        End If
'        'cnn.Execute sSQLQuery, rdExecDirect
'        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'            '6/13/16: Replace GoSub
'            'GoSub ErrHand:
'            mErrHand iRet, iDoTrans
'        End If
'        If iRet = 0 Then
'            If iDoTrans Then
'                cnn.CommitTrans
'            End If
'        ElseIf (iRet = BTRV_ERR_REC_LOCKED) Or (iRet = BTRV_ERR_FILE_LOCKED) Or (iRet = BTRV_ERR_INCOM_LOCK) Or (iRet = BTRV_ERR_CONFLICT) Then
'            fStart = Timer
'            Do While Timer <= fStart
'                iRet = iRet
'            Loop
'            iCount = iCount + 1
'            If iCount > igWaitCount Then
'                gMsgBox "A SQL error has occurred: " & "Error # " & iRet, vbCritical
'                Exit Do
'            End If
'        End If
'    Loop While (iRet = BTRV_ERR_REC_LOCKED) Or (iRet = BTRV_ERR_FILE_LOCKED) Or (iRet = BTRV_ERR_INCOM_LOCK) Or (iRet = BTRV_ERR_CONFLICT)
'    gSQLWait = iRet
'    If iRet <> 0 Then
'        On Error GoTo mOpenFileErr:
'        hlMsg = FreeFile
'        Open sgMsgDirectory & "AffErrorLog.txt" For Append As hlMsg
'        Print #hlMsg, sSQLQuery
'        Print #hlMsg, "Error # " & iRet
'        Close #hlMsg
'    End If
'    On Error GoTo 0
'    Exit Function
'
'ErrHand:
'    For Each gErrSQL In cnn.Errors
'        iRet = gErrSQL.NativeError
'        If iRet < 0 Then
'            iRet = iRet + 4999
'        End If
'        'If (iRet = 84) And (iDoTrans) Then              'SQLSetConnectAttr vs. SQLSetOpenConnection
'        If (iRet = BTRV_ERR_REC_LOCKED) Or (iRet = BTRV_ERR_FILE_LOCKED) Or (iRet = BTRV_ERR_INCOM_LOCK) Or (iRet = BTRV_ERR_CONFLICT) Then
'            If iDoTrans Then
'                cnn.RollbackTrans
'            End If
'            cnn.Errors.Clear
'            Resume Next
'        End If
'        If iRet <> 0 Then              'SQLSetConnectAttr vs. SQLSetOpenConnection
'            gMsgBox "A SQL error has occurred: " & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, vbCritical
'        End If
'    Next gErrSQL
'    If iDoTrans Then
'        cnn.RollbackTrans
'    End If
'    cnn.Errors.Clear
'    Resume Next
'mOpenFileErr:
'    Resume Next
'End Function




'***************************************************************************************
'*
'* Procedure Name: gLoadOption
'*
'* Created: 8/22/03 - J. Dutschke
'*
'* Modified:              By:
'*
'* Comments: This function loads a string value from the ini file.
'*           It relies on the global variable sgIniPathFileName to
'*           contain the path and name of the ini file to use.
'*
'***************************************************************************************
Public Function gLoadOption(Section As String, Key As String, sValue As String) As Boolean
    On Error GoTo ERR_gLoadOption
    Dim BytesCopied As Integer
    Dim sBuffer As String * 128

    gLoadOption = False
    BytesCopied = GetPrivateProfileString(Section, Key, "Not Found", sBuffer, 128, sgIniPathFileName)
    If BytesCopied > 0 Then
        If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
            sValue = Left(sBuffer, BytesCopied)
            gLoadOption = True
        End If
    End If
    Exit Function

ERR_gLoadOption:
    ' return now if an error occurs
End Function

'***************************************************************************************
'*
'* Procedure Name: gSaveOption
'*
'* Created: 5/12/03 - J. Dutschke
'*
'* Modified:              By:
'*
'* Comments:
'*
'***************************************************************************************
Public Function gSaveOption(Section As String, Key As String, sValue As String) As Boolean
    On Error GoTo ERR_gSaveOption
    Dim BytesCopied As Integer

    gSaveOption = False
    If WritePrivateProfileString(Section, Key, sValue, sgIniPathFileName) Then
        gSaveOption = True
    End If
    Exit Function

ERR_gSaveOption:
    ' return now if an error occurs
End Function

'***************************************************************************************
'*
'* Procedure Name: gSetPathEndSlash
'*
'* Created: 10/02/03 - J. Dutschke
'*
'* Modified:              By:
'*
'* Comments: Install the final back slash if it does not already exist.
'*
'***************************************************************************************
Public Function gSetPathEndSlash(ByVal slInPath As String, ilAdjDrivePath As Integer) As String
    Dim slPath As String
    slPath = Trim$(slInPath)
    If right$(slPath, 1) <> "\" Then
        slPath = slPath + "\"
    End If
    If ilAdjDrivePath Then
        slPath = gAdjustDrivePath(slPath)
    End If
    gSetPathEndSlash = slPath
End Function

Public Function gSetURLPathEndSlash(ByVal slInPath As String, ilAdjDrivePath As Integer) As String
    Dim slPath As String
    slPath = Trim$(slInPath)
    If right$(slPath, 1) <> "/" Then
        slPath = slPath + "/"
    End If
    If ilAdjDrivePath Then
        slPath = gAdjustDrivePath(slPath)
    End If
    gSetURLPathEndSlash = slPath
End Function




Public Sub gGrid_Clear(grdCtrl As MSHFlexGrid, ilFillRows As Integer)
    
'
'   grdCtrl (I)-  Grid Control name
'   ilFillRows (I)- True=Fill Grid with blank rows; False=Only have one blank row
'
    Dim llRows As Long
    Dim llCols As Long
    Dim llFillNoRow As Long
    'TTP 10390 - Affiliate grids: lines drawn one at a time, seeming to appear more slowly than before
    'grdCtrl.Redraw = False
    If ilFillRows Then
        llFillNoRow = grdCtrl.Height \ grdCtrl.RowHeight(grdCtrl.FixedRows) - 2
    Else
        llFillNoRow = 0
    End If
    grdCtrl.Rows = grdCtrl.FixedRows + 1
    llRows = grdCtrl.Rows
    Do While llRows >= grdCtrl.FixedRows + llFillNoRow + 1
        If llRows <= grdCtrl.FixedRows + 1 Then
            Exit Do
        End If
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
        Next llCols
        llRows = llRows + 1
    Loop
    'TTP 10390 - Affiliate grids: lines drawn one at a time, seeming to appear more slowly than before
    'grdCtrl.Redraw = True
End Sub

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
    End If
    Do While llHeight + grdCtrl.RowHeight(grdCtrl.FixedRows) < grdCtrl.Height
        llHeight = llHeight + grdCtrl.RowHeight(grdCtrl.FixedRows)
    Loop
    grdCtrl.Height = llHeight + 15
End Sub

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
    
    llFillNoRow = grdCtrl.Height \ grdCtrl.RowHeight(grdCtrl.FixedRows) - grdCtrl.FixedRows - 1
    
    For llRows = grdCtrl.FixedRows To grdCtrl.FixedRows + llFillNoRow Step 1
        Do While llRows >= grdCtrl.Rows
            grdCtrl.AddItem ""
            For llCols = 0 To grdCtrl.Cols - 1 Step 1
                grdCtrl.TextMatrix(llRows, llCols) = ""
            Next llCols
        Loop
    Next llRows
End Sub

Public Function gGrid_DetermineRowCol(grdCtrl As MSHFlexGrid, X As Single, Y As Single) As Integer
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

Public Sub gSetFonts(frm As Form)
    Dim Ctrl As control
    Dim ilFontSize As Integer
    Dim ilColorFontSize As Integer
    Dim ilBold As Integer
    Dim ilChg As Integer
    Dim slStr As String
    Dim slFontName As String
    
    'On Error Resume Next
    ilFontSize = 14
    ilBold = True
    ilColorFontSize = 10
    slFontName = "Arial"
    If Screen.Height / Screen.TwipsPerPixelY <= 480 Then
        ilFontSize = 8
        ilBold = False
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 600 Then
        ilFontSize = 8
        ilBold = False
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 768 Then
        ilFontSize = 9  '10
        ilBold = False
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 800 Then
        ilFontSize = 9  '10
        ilBold = True
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 1024 Then
        ilFontSize = 11 '12
        ilBold = True
    End If
    For Each Ctrl In frm.Controls
        If TypeOf Ctrl Is MSHFlexGrid Then
            'If Ctrl.Name = "grdCounts" Then
            '    Ctrl.Font.Name = "Arial Narrow"
            '    Ctrl.FontFixed.Name = "Arial Narrow"
            'Else
            '    Ctrl.Font.Name = slFontName
            '    Ctrl.FontFixed.Name = slFontName
            'End If
            If Ctrl.Font.Name = "Arial Narrow" Then
                slFontName = Ctrl.Font.Name
            End If
            Ctrl.Font.Name = slFontName
            Ctrl.FontFixed.Name = slFontName
            Ctrl.Font.Size = ilFontSize
            Ctrl.FontFixed.Size = ilFontSize
            Ctrl.Font.Bold = ilBold
            Ctrl.FontFixed.Bold = ilBold
        ElseIf TypeOf Ctrl Is TabStrip Then
            Ctrl.Font.Name = slFontName
            Ctrl.Font.Size = ilFontSize
            Ctrl.Font.Bold = ilBold
        'ElseIf TypeOf Ctrl Is Resize Then
        'ElseIf TypeOf Ctrl Is Timer Then
        'ElseIf TypeOf Ctrl Is Image Then
        'ElseIf TypeOf Ctrl Is ImageList Then
        'ElseIf TypeOf Ctrl Is CommonDialog Then
        'ElseIf TypeOf Ctrl Is AffExportCriteria Then
        'ElseIf TypeOf Ctrl Is AffCommentGrid Then
        'ElseIf TypeOf Ctrl Is AffContactGrid Then
        ElseIf (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is ListBox) _
               Or (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is PictureBox) _
               Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Label) _
               Or (TypeOf Ctrl Is CSI_Calendar) Or (TypeOf Ctrl Is CSI_Calendar_UP) Or (TypeOf Ctrl Is CSI_ComboBoxList) Or (TypeOf Ctrl Is CSI_DayPicker) Then
            ilChg = 0
            If TypeOf Ctrl Is CommandButton Then
               ilChg = 1
            Else
                If (Ctrl.ForeColor = vbBlack) Or (Ctrl.ForeColor = &H80000008) Or (Ctrl.ForeColor = &H80000012) Or (Ctrl.ForeColor = &H8000000F) Then
                    ilChg = 1
                Else
                    ilChg = 2
                End If
            End If
            slStr = Ctrl.Name
            If (InStr(1, slStr, "Arrow", vbTextCompare) > 0) Or ((InStr(1, slStr, "Dropdown", vbTextCompare) > 0) And (TypeOf Ctrl Is CommandButton)) Then
                ilChg = 0
            End If
            If ilChg = 1 Then
                Ctrl.FontName = slFontName
                Ctrl.FontSize = ilFontSize
                Ctrl.FontBold = ilBold
            ElseIf ilChg = 2 Then
                Ctrl.FontName = slFontName
                Ctrl.FontSize = ilColorFontSize
                Ctrl.FontBold = False
            End If
        End If
    Next Ctrl
End Sub

Public Function gCheckIfSpotsHaveBeenExported(iVefCode As Integer, sWeek As String, iAgreeType As Integer) As Integer
    
    'D.S. 10/25/04
    'Test to see if the spots for either the Web or Univision have been exported.  If so return TRUE
    
    Dim temp_rst As ADODB.Recordset
    
    gCheckIfSpotsHaveBeenExported = False
    
    If iAgreeType = 1 Then
        'It's a web agreement
        SQLQuery = " Select aufCode FROM AUF_Alert_User"
        SQLQuery = SQLQuery & " Where aufMoWeekDate = '" & Format$(sWeek, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " And  aufVefCode = " & iVefCode
        SQLQuery = SQLQuery & " And  aufType = '" & "F" & "'"
        SQLQuery = SQLQuery & " And  aufSubType = '" & "S" & "'"
        SQLQuery = SQLQuery & " And  aufStatus = '" & "C" & "'"
               
        Set temp_rst = gSQLSelectCall(SQLQuery)
        If (Not temp_rst.EOF) Then
            gCheckIfSpotsHaveBeenExported = True
        End If
    End If
    
    If iAgreeType = 2 Then
        'It's a Univision agreement
        SQLQuery = " Select aetCode FROM Aet"
        SQLQuery = SQLQuery & " Where aetPledgeStartDate >= '" & Format$(sWeek, sgSQLDateForm) & "'"
        sWeek = DateAdd("d", 6, sWeek)
        SQLQuery = SQLQuery & " And  aetPledgeEndDate <= '" & Format$(sWeek, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " And  aetVefCode = " & iVefCode

        Set temp_rst = gSQLSelectCall(SQLQuery)
        If (Not temp_rst.EOF) Then
            gCheckIfSpotsHaveBeenExported = True
        End If
    End If
    
    Exit Function

End Function




Public Function gGetComputerName() As String

'D.S. 2/9/05 Returns the name of the users computer

   Dim strBuffer As String * 255

   If GetComputerName(strBuffer, 255&) <> 0 Then
      ' Name exist
      gGetComputerName = Left(strBuffer, InStr(strBuffer, vbNullChar) - 1)
   Else
      ' Name doesn't exist
      gGetComputerName = "N/A"
   End If
   
End Function


Public Function gFindVehStaFromAttCode(sATTCode As String) As String

    Dim att_rst As ADODB.Recordset
    Dim vef_rst As ADODB.Recordset
    Dim sta_rst As ADODB.Recordset
    
    Dim slStation As String
    Dim slVehicle As String
    
    'Find the vehicle and station name using the AttCode
    'Return a string in the form of vehicle name, statiion name,  Exp. "Billy Crystal, KAAA-FM,"
    'or Return "Unable to find Agreement Code "
    
    On Error GoTo ErrHand:
    
    SQLQuery = "SELECT *"
    SQLQuery = SQLQuery + " FROM att"
    SQLQuery = SQLQuery + " WHERE (attCode = " & sATTCode & ")"
    Set att_rst = gSQLSelectCall(SQLQuery)
    
    If Not att_rst.EOF Then
        SQLQuery = "SELECT shttCallLetters"
        SQLQuery = SQLQuery + " FROM shtt"
        SQLQuery = SQLQuery + " WHERE shttCode = " & att_rst!attshfcode
        Set sta_rst = gSQLSelectCall(SQLQuery)
        slStation = Trim$(sta_rst!shttCallLetters)

        SQLQuery = "SELECT VefName"
        SQLQuery = SQLQuery + " FROM VEF_Vehicles"
        SQLQuery = SQLQuery + " WHERE vefCode = " & att_rst!attvefCode
        Set vef_rst = gSQLSelectCall(SQLQuery)
        slVehicle = Trim$(vef_rst!vefName)
        gFindVehStaFromAttCode = slVehicle & ", " & slStation & ", "
    Else
        gFindVehStaFromAttCode = "Unable to find Agreement Code "
        Exit Function
    End If
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modGenSubs-gFindVehStaFromAttCode"
    gFindVehStaFromAttCode = "Failure in  gFindVehStaFromAttCode"
    Exit Function
End Function


Public Function gGetDataNoQuotes(sDataStr As String) As String
    Dim ilLen As Integer
    Dim ilLoop As Integer
    Dim slNewStr As String
    Dim clOneChar As String
    
    ilLen = Len(sDataStr)
    slNewStr = ""
    For ilLoop = 1 To ilLen
        clOneChar = Mid(sDataStr, ilLoop, 1)
        If clOneChar <> """" Then
            slNewStr = slNewStr + clOneChar
        End If
    Next
    gGetDataNoQuotes = Trim(slNewStr)
End Function

Public Function GetServersDateTime() As String
    On Error GoTo ErrHandler
    Dim objXMLHTTP
    Dim llReturn As Long
    Dim slISAPIExtensionDLL As String
    Dim slRootURL As String
    Dim slResponse As String
    Dim slCommand As String
    Dim slRegSection As String
    Dim alRecordsArray() As String
    Dim aDataArray() As String
    Dim WebCmds As New WebCommands

    GetServersDateTime = ""

    If Not gLoadOption(sgWebServerSection, "RootURL", slRootURL) Then
        gMsgBox "FAIL: gCheckWebSession: LoadOption RootURL Failed"
        Exit Function
    End If
    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    slCommand = "Select GetDate() as ServerDate"
    If gLoadOption(sgWebServerSection, "RegSection", slRegSection) Then
        slISAPIExtensionDLL = slRootURL & "ExecuteSQL.dll" & "?ExecSQL?PW=jfdl" & Now() & "&RK=" & Trim(slRegSection) & "&SQL=" & slCommand
    End If
    llReturn = 1
    If bgUsingSockets Then
        slResponse = WebCmds.ExecSQL(slCommand)
        If Not Left(slResponse, 5) = "ERROR" Then
            llReturn = 200
        End If
    Else
        Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
        objXMLHTTP.Open "GET", slISAPIExtensionDLL, False
        objXMLHTTP.Send
        llReturn = objXMLHTTP.Status
        slResponse = objXMLHTTP.responseText
        Set objXMLHTTP = Nothing
    End If
    If llReturn <> 200 Then
        Exit Function
    End If
    alRecordsArray = Split(slResponse, vbCrLf)
    If Not IsArray(alRecordsArray) Then
        Exit Function
    End If
    If UBound(alRecordsArray) < 1 Then
        Exit Function
    End If

    GetServersDateTime = gGetDataNoQuotes(alRecordsArray(1))
    Exit Function
    
ErrHandler:
    gMsg = "A general error has occurred in modGenSubs-GetServersDateTime: "
    gLogMsg gMsg & Err.Description & " Error #" & Err.Number & "; Line #" & Erl, "WebImportLog.Txt", False
    gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
End Function

Public Sub gTestAgreeAndStationEmails()

    Dim temp_rst As ADODB.Recordset
    Dim temp_shtt As ADODB.Recordset
    Dim temp_att As ADODB.Recordset
    Dim slTemp As String
    Dim ilRet As Integer
    Dim llTemp As Long
    Dim slTemp2 As String
    
        
    On Error GoTo ErrHand:
    SQLQuery = " Select attCode, attshfcode, attVefCode, attWebEmail FROM ATT "
    SQLQuery = SQLQuery & " Where attExportType  = 1 order by attCode"
           
    Set temp_rst = gSQLSelectCall(SQLQuery)
    While Not temp_rst.EOF
        slTemp = Trim(temp_rst!attWebEmail)
        ilRet = gTestForMultipleEmail(slTemp, "RegEmail")
        If ilRet = False Then
            llTemp = CLng(temp_rst!attshfcode)
            
            SQLQuery = " Select shttCallLetters FROM shtt"
            SQLQuery = SQLQuery & " Where (shttCode = " & CLng(temp_rst!attshfcode) & ")"
            Set temp_shtt = gSQLSelectCall(SQLQuery)
            
            SQLQuery = " Select vefName FROM VEF_Vehicles"
            SQLQuery = SQLQuery & " Where (vefCode = " & CLng(temp_rst!attvefCode) & ")"
            Set temp_att = gSQLSelectCall(SQLQuery)
            slTemp2 = Trim$(temp_shtt!shttCallLetters)
            
            'gMsgBox sgErrorMsg & " " & Trim$(temp_shtt!shttCallLetters)
            gLogMsg sgErrorMsg & " " & slTemp2 & " " & Trim$(temp_att!vefName), "BadEmail.txt", False
            gLogMsg "", "BadEmail.txt", False
        End If
        temp_rst.MoveNext
    Wend
    Exit Sub
        
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modGenSubs-gTestAgreeAndStationEmails"
    Exit Sub
End Sub

Public Function gFindAttHole() As Long

    'D.S. 8/2/05 Find a hole in Att to insert the next record

    Dim temp_rst As ADODB.Recordset
    Dim llIdx1 As Long
    Dim llIdx2 As Long
        
    On Error GoTo ErrHand:
    
    SQLQuery = "Select Max(attCode) FROM ATT"
    Set temp_rst = gSQLSelectCall(SQLQuery)
    
    'Its a new database let it find it's own
    If IsNull(temp_rst(0).Value) Then
        gFindAttHole = 0
        Exit Function
    End If
    
    'Its not out of space let it find it's own
    If CLng(temp_rst(0).Value) > 0 And CLng(temp_rst(0).Value) < 2147483646 Then
        gFindAttHole = 0
        Exit Function
    End If
    
    'The Att is maxed out and we need to look for a hole for the insert
    SQLQuery = "Select attcode FROM ATT Order By attCode"
    Set temp_rst = gSQLSelectCall(SQLQuery)
    
    While Not temp_rst.EOF
        llIdx1 = CLng(temp_rst!attCode)
        temp_rst.MoveNext
        llIdx2 = CLng(temp_rst!attCode)
        If (llIdx2 < 2147483646) Then
            If (llIdx2 <> llIdx1 + 1) Then
                gFindAttHole = llIdx1 + 1
                Exit Function
            End If
        Else
            gMsgBox "There is no more room in the Att file to insert an agreement." & Chr(13) & Chr(10) & "Call Counteropint", vbCritical
            gLogMsg "There is no more room in the Att file to insert an agreement", "AttInsertError.Txt", False
            gFindAttHole = -1
        End If
    Wend
Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modGenSubs-gFindAttHole"
    gFindAttHole = -1
    Exit Function
End Function

Public Sub gSleep(ilTotalSeconds As Long)
    Dim ilLoop As Integer
    
    For ilLoop = 0 To ilTotalSeconds
        DoEvents
        Sleep (1000)   ' Wait 1 second
    Next
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



Public Function gRemoteMaxPostDayResults(sFileName As String, sIniValue As String) As String
    
    'D.S. 6/04
    'Purpose:

    Dim slLocation As String
    Dim hlFrom As Integer
    Dim ilRet  As Integer
    Dim slPostDate As String
    
    If Not gHasWebAccess() Then
        'Doug- on 11/17/06 I changed this from True to blank
        gRemoteMaxPostDayResults = ""
        Exit Function
    End If
    
    On Error GoTo ErrHand
    
    gRemoteMaxPostDayResults = ""
    
    Call gLoadOption(sgWebServerSection, sIniValue, slLocation)
    'slLocation = gSetPathEndSlash(slLocation, True)
    If (StrComp(sIniValue, "WebImports", vbTextCompare) = 0) Or (StrComp(sIniValue, "WebExports", vbTextCompare) = 0) Then
        slLocation = gSetPathEndSlash(slLocation, True)
    Else
        slLocation = gSetPathEndSlash(slLocation, False)
    End If
    slLocation = slLocation & sFileName
    
    On Error GoTo FileErrHand:
    'hlFrom = FreeFile
    ilRet = 0
    'Open slLocation For Input Access Read As hlFrom
    ilRet = gFileOpen(slLocation, "Input Access Read", hlFrom)
    If ilRet <> 0 Then
        Close hlFrom
        gMsgBox "Error: modGenSubsg-RemoteProcessPWResults was unable to open the file."
        GoTo ErrHand
        Exit Function
    End If
    
    'Move past the header information
    Input #hlFrom, slPostDate
    If Not EOF(hlFrom) Then
        Input #hlFrom, slPostDate
        'Process the agreement passwords
        gRemoteMaxPostDayResults = Format$(slPostDate, sgShowDateForm)
    End If
    Close hlFrom
    Exit Function
    
FileErrHand:
    Close hlFrom
    gRemoteMaxPostDayResults = ""
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occurred in frmGenSubs-gRemoteMaxPostDayResults: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "FastAddVerbose.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gRemoteMaxPostDayResults = ""
End Function

Public Function gGetLastPostedDate(llAttCode As Long, ilExportType As Integer, slExportToWeb As String, slExportToUnivision As String, slExportToMarketron As String, slExportToCBS As String, slExportToClearCh As String) As String
    
    Dim slLastWebPostedDate As String
    Dim ilRet As Integer
    Dim slTempDate As String
    Dim rst_Aet As ADODB.Recordset
    Dim rst_Ast As ADODB.Recordset
    Dim rst_Webl As ADODB.Recordset
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    slLastWebPostedDate = "1/1/1970"
    
    'This covers the case where all spots have been deleted.
    'D.S. 10/23/14 Select Count was too slow. All we need to know if there was at least one record
    'SQLQuery = "Select COUNT(astCode) from AST"
    SQLQuery = "Select Top 1 astCode from AST"
    SQLQuery = SQLQuery + " WHERE"
    SQLQuery = SQLQuery + " astAtfCode = " & llAttCode
    Set rst = gSQLSelectCall(SQLQuery)
    
    If Not rst.EOF Then
        If rst(0).Value = 0 Then
            gGetLastPostedDate = slLastWebPostedDate
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    
        
    'If ilExportType = 2 Then    'Univision
    If slExportToUnivision = "Y" Then
        SQLQuery = " Select Max(aetFeedDate) FROM Aet"
        SQLQuery = SQLQuery & " Where aetatfCode = " & llAttCode
        Set rst_Aet = gSQLSelectCall(SQLQuery)
        If Not rst_Aet.EOF Then
            'D.S. 10/11/06
            'If rst_Aet(0).Value <> Null Then
            If IsDate(rst_Aet(0).Value) Then
                slLastWebPostedDate = Format$(rst_Aet(0).Value, sgShowDateForm)
            End If
        Else
            SQLQuery = "SELECT max(astFeedDate) FROM ast WHERE"
            SQLQuery = SQLQuery & " astAtfCode = " & rst!attCode
            SQLQuery = SQLQuery & " AND astCPStatus = 1"
            Set rst_Ast = gSQLSelectCall(SQLQuery)
            If Not rst_Ast.EOF Then
                'D.S. 10/11/06
                'If rst_Ast(0).Value <> Null Then
                If IsDate(rst_Ast(0).Value) Then
                    slLastWebPostedDate = Format$(rst_Ast(0).Value, sgShowDateForm)
                End If
            End If
        End If
    Else
        'If ilExportType = 1 Then    'Web
        If slExportToWeb = "Y" Then
            'D.S. 12/23/08
            'I no longer think that the webl is OK to go by
            'SQLQuery = "SELECT max(weblPostDay) FROM webl WHERE"
            'SQLQuery = SQLQuery & " weblType = 1 And weblAttCode = " & llAttCode
            'Set rst_Webl = gSQLSelectCall(SQLQuery)
            'If Not rst_Webl.EOF Then
            '    'D.S. 10/11/06
            '    'If rst_Webl(0).Value <> Null Then
            '    If IsDate(rst_Webl(0).Value) Then
            '        slLastWebPostedDate = Format$(rst_Webl(0).Value, sgShowDateForm)
            '    End If
            'End If
            'Doug (9/25/06)- Add test to see if submitted date is newer then smLastPostedDate
            '      If so, update smLastPostedDate with that date
            '      Don't show error message if unable to access web
            'ilRet = gRemoteExecSql("Select Max(PostDate) from spots where attCode = " & "'" & rst!attCode & "'", "MaxPostDate.txt", "WebImports", True, True)
            'ilRet = gRemoteExecSql("Select Max(PledgeStartDate) As MaxPostDate from spots  where postedFlag = 1 And attCode = " & "'" & llAttCode & "'", "MaxPostDate.txt", "WebImports", True, True, 30)
            ilRet = gRemoteExecSql("Select Max(PledgeStartDate) from spots WITH (INDEX(IX_Spots_Headers)) Where attCode = " & llAttCode & " And postedFlag = 1 ", "MaxPostDate.txt", "WebImports", True, True, 30)
            slTempDate = gRemoteMaxPostDayResults("MaxPostDate.txt", "WebImports")
            If slTempDate <> "" Then
                If DateValue(slTempDate) > DateValue(slLastWebPostedDate) Then
                    slLastWebPostedDate = slTempDate
                End If
            End If
        End If
        If (ilExportType = 1) Or (slExportToMarketron = "Y") Or (slExportToCBS = "Y") Or (slExportToClearCh = "Y") Then
            'Get lastest posted date as user will not be allowed to drop prior to that date
            'SQLQuery = "SELECT max(astFeedDate) FROM ast WHERE"
            'SQLQuery = SQLQuery & " astAtfCode = " & llAttCode
            'SQLQuery = SQLQuery & " AND astCPStatus = 1"
            'D.S. 10/23/14 SELECT max(astFeedDate) FROM ast WHERE"... above was too slow. Now we look at CPTT to get the last week posting occured instead of the exact date
            SQLQuery = "select top 1 cpttStartDate from cptt where cpttPostingstatus > 0 AND cpttatfcode = " & llAttCode & " order by cpttstartdate desc"
            Set rst_Ast = gSQLSelectCall(SQLQuery)
            If Not rst_Ast.EOF Then
                'D.S. 10/11/06
                'If rst_Ast(0).Value <> Null Then
                If IsDate(rst_Ast(0).Value) Then
                    If (DateValue(Format$(rst_Ast(0).Value, sgShowDateForm)) > DateValue(slLastWebPostedDate)) Then
                        slLastWebPostedDate = Format$(rst_Ast(0).Value, sgShowDateForm)
                    End If
                End If
            End If
        End If
    End If
    
    If slLastWebPostedDate <> "" Then
        slLastWebPostedDate = gAdjYear(slLastWebPostedDate)
    End If
    gGetLastPostedDate = slLastWebPostedDate
    Screen.MousePointer = vbDefault
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modGenSubs-gGetLastPostedDate"
    gGetLastPostedDate = "1/1/1970"
End Function

Public Function gTerminate(slFileName As String) As Integer

    Dim ilYesNo As Integer
    
    gTerminate = False
    ilYesNo = gMsgBox("Are you sure that you want to cancel the program?", vbYesNo)
    If ilYesNo = vbYes Then
        gLogMsg "User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **", slFileName, False
        gTerminate = True
    End If
End Function



'***********************************************************************************************
'
'  Created:12/15/2006    By:J. Dutschke
' Modified:              By:
'
' Comments:
' This function removes an agreement from the att table and ensures all other records from
' the CPTT, AET, DAT, AST tables also get there entries removed.
'
' This function also ensures the entire operation is completed using a begin and commit
' transaction.
'
' If it is a web agreement, they are also deleted from the web site.
'
'***********************************************************************************************
Public Function gDeleteAgreement(lAttCode As Long, sFileName As String) As Boolean
    Dim llTotalRecordsDeleted As Long
    Dim ilExportType As Integer
    Dim slExportToWeb As String
    '7701 removed
   ' Dim slExportToCumulus As String
    Dim attrst As ADODB.Recordset
    On Error GoTo ErrHand

    gDeleteAgreement = False
    ilExportType = 0
    ' first load this agreement to determine if it is web enabled or not.
    SQLQuery = "SELECT attExportType, attExportToWeb, attWebInterface From ATT Where attCode = " & lAttCode
    Set attrst = gSQLSelectCall(SQLQuery)
    If Not attrst.EOF Then
        ilExportType = attrst!attExportType
        slExportToWeb = attrst!attExportToWeb
        '7701 removed
'        If gIfNullInteger(attrst!vatWvtIdCodeLog) = Vendors.Cumulus Then
'            slExportToCumulus = "C"
'        End If
    End If

    llTotalRecordsDeleted = 0
    'cnn.BeginTrans

    ' ATT Table
    SQLQuery = "DELETE FROM Att WHERE AttCode = " & lAttCode
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError "AffErrorLog.txt", "modGenSubs-gDeleteAgreement"
        'cnn.RollbackTrans
        gDeleteAgreement = False
        Exit Function
    End If
    '7701 vat table
    SQLQuery = "DELETE FROM VAT_Vendor_Agreement WHERE vatAttCode = " & lAttCode
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError "AffErrorLog.txt", "modGenSubs-gDeleteAgreement"
        'cnn.RollbackTrans
        gDeleteAgreement = False
        Exit Function
    End If
    
    ' CPTT Table
    SQLQuery = "DELETE FROM Cptt WHERE cpttAtfCode = " & lAttCode
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError "AffErrorLog.txt", "modGenSubs-gDeleteAgreement"
        'cnn.RollbackTrans
        gDeleteAgreement = False
        Exit Function
    End If

    ' DAT Table
    SQLQuery = "DELETE FROM dat WHERE datAtfCode = " & lAttCode
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError "AffErrorLog.txt", "modGenSubs-gDeleteAgreement"
        'cnn.RollbackTrans
        gDeleteAgreement = False
        Exit Function
    End If

    ' EPT Table
    SQLQuery = "DELETE FROM ept WHERE eptAttCode = " & lAttCode
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError "AffErrorLog.txt", "modGenSubs-gDeleteAgreement"
        'cnn.RollbackTrans
        gDeleteAgreement = False
        Exit Function
    End If

    ' PET Table
    SQLQuery = "DELETE FROM pet WHERE petAttCode = " & lAttCode
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError "AffErrorLog.txt", "ModGenSubs-gDeleteAgreement"
        'cnn.RollbackTrans
        gDeleteAgreement = False
        Exit Function
    End If

    ' AET Table
    SQLQuery = "DELETE FROM aet WHERE aetAtfCode = " & lAttCode
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError "AffErrorLog.txt", "ModGenSubs-gDeleteAgreement"
        'cnn.RollbackTrans
        gDeleteAgreement = False
        Exit Function
    End If

    ' AST Table
    SQLQuery = "DELETE FROM Ast WHERE astAtfCode = " & lAttCode
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError "AffErrorLog.txt", "modGenSubs-gDeleteAgreement"
        'cnn.RollbackTrans
        gDeleteAgreement = False
        Exit Function
    End If
    'Dan 7701 12/3/15 no longer need to test cumulus
    'If ilExportType <> 1 Then ' Is this a web enabled agreement?
'    If (slExportToWeb <> "Y") And (slExportToCumulus <> "C") Then
'        cnn.CommitTrans       ' Nope. So go ahead and commit the changes and return now.
'        gDeleteAgreement = True
'        Exit Function
'    End If
    If (slExportToWeb <> "Y") Then
        'cnn.CommitTrans       ' Nope. So go ahead and commit the changes and return now.
        gDeleteAgreement = True
        Exit Function
    End If

    ' If we make it here, then it is a web enabled agreement. Now that all local records have been
    ' deleted, delete the agreement and all spots from the web site as well.

    ' First remove the agreement itself from the header table.
    SQLQuery = "Delete From Header Where attCode = " & lAttCode
    llTotalRecordsDeleted = gExecWebSQLWithRowsEffected(SQLQuery)
    If llTotalRecordsDeleted = -1 Then
        'cnn.RollbackTrans
        gLogMsg "Error trying to delete agreement from the web for attCode." & lAttCode, sFileName, False
        Exit Function
    End If
    gLogMsg "    " & CStr(llTotalRecordsDeleted) & " Web agreements were deleted for attCode." & lAttCode, sFileName, False

    ' Now remove the spots associate with this agreement.
    SQLQuery = "Delete From Spots Where attCode = " & lAttCode
    llTotalRecordsDeleted = gExecWebSQLWithRowsEffected(SQLQuery)
    If llTotalRecordsDeleted = -1 Then
        'cnn.RollbackTrans
        gLogMsg "Error trying to delete spots from the web for attCode." & lAttCode, sFileName, False
        Exit Function
    End If
    gLogMsg "    " & CStr(llTotalRecordsDeleted) & " Web spot records were Deleted for attCode: " & lAttCode, sFileName, False

    SQLQuery = "Delete From SpotRevisions Where attCode = " & lAttCode
    llTotalRecordsDeleted = gExecWebSQLWithRowsEffected(SQLQuery)
    If llTotalRecordsDeleted = -1 Then
        'cnn.RollbackTrans
        gLogMsg "Error trying to delete SpotsRevisions from the web for attCode." & lAttCode, sFileName, False
        Exit Function
    End If
    gLogMsg "    " & CStr(llTotalRecordsDeleted) & " Web spot revisions records were Deleted for attCode: " & lAttCode, sFileName, False

    'cnn.CommitTrans
    gDeleteAgreement = True
    Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "modGenSubs-gDeleteAgreement"
End Function

Public Sub gGrid_SortByCol(grdCtrl As MSHFlexGrid, ilTestCol As Integer, ilSortCol As Integer, ilPrevSortCol As Integer, ilPrevSortDirection As Integer, Optional blSetRedraw As Boolean = True)
    Dim llEndRow As Long
    
    If blSetRedraw Then
        grdCtrl.Redraw = False
    End If
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
    If blSetRedraw Then
        grdCtrl.Redraw = True
    End If
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mGetProdOrShtTitle              *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get either Product Name or     *
'*                      Short Title                    *
'*                                                     *
'*******************************************************
Function gGetProdOrShtTitle(llSdfCode As Long, llInSifCode As Long, Optional ilMaxAbbrChar = -1) As String
    Dim llSifCode As Long
    Dim rst_chf As ADODB.Recordset
    Dim rst_Sif As ADODB.Recordset

    On Error GoTo ErrHand:
    If (sgSpfUseProdSptScr <> "P") Then
        SQLQuery = "SELECT adfName, adfAbbr, chfProduct"
        SQLQuery = SQLQuery & " FROM SDF_Spot_Detail, ADF_Advertisers, CHF_Contract_Header"
        SQLQuery = SQLQuery & " WHERE sdfCode = " & llSdfCode
        SQLQuery = SQLQuery & " AND chfCode = sdfChfCode"
        SQLQuery = SQLQuery & " AND adfCode = sdfAdfCode"
        Set rst_chf = gSQLSelectCall(SQLQuery)
        If Not rst_chf.EOF Then
            If ilMaxAbbrChar = -1 Then
                gGetProdOrShtTitle = Trim$(rst_chf!adfAbbr) & "," & Trim$(rst_chf!chfProduct)
            Else
                gGetProdOrShtTitle = Trim$(Left(Trim$(rst_chf!adfAbbr), ilMaxAbbrChar)) & "," & Trim$(rst_chf!chfProduct)
            End If
        Else
            gGetProdOrShtTitle = ""
        End If
        rst_chf.Close
        Exit Function
    Else
        If llInSifCode <= 0 Then
            SQLQuery = "SELECT chfSifCode"
            SQLQuery = SQLQuery & " FROM SDF_Spot_Detail, CHF_Contract_Header"
            SQLQuery = SQLQuery & " WHERE sdfCode = " & llSdfCode
            SQLQuery = SQLQuery & " AND chfCode = sdfChfCode"
            Set rst_chf = gSQLSelectCall(SQLQuery)
            If Not rst_chf.EOF Then
                llSifCode = rst_chf!chfSifCode
            Else
                llSifCode = 0
            End If
            rst_chf.Close
            If llSifCode <= 0 Then
                gGetProdOrShtTitle = ""
                Exit Function
            End If
        Else
            llSifCode = llInSifCode
        End If
        
        SQLQuery = "SELECT sifName"
        SQLQuery = SQLQuery & " FROM SIF_Short_Title"
        SQLQuery = SQLQuery & " WHERE sifCode = " & llSifCode
        Set rst_Sif = gSQLSelectCall(SQLQuery)
        If Not rst_Sif.EOF Then
            gGetProdOrShtTitle = Trim$(rst_Sif!sifName)
        Else
            gGetProdOrShtTitle = ""
        End If
        rst_Sif.Close
    End If
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modGenSubs-gGetProdOrShtTitle"
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:gGetShortTitle                  *
'*                                                     *
'*             Created:10/22/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get associated short title for  *
'*                     a spot                          *
'*                     Note: First check Rotation if   *
'*                           defined, if not use       *
'*                           one from contract         *
'*            Same code in gBuildODFSpotCode           *
'*                                                     *
'*******************************************************
Function gGetShortTitle(llSdfCode As Long, Optional ilMaxAbbrChar = -1) As String
    Dim llVsfCode As Long
    Dim llSifCode As Long
    Dim ilSchVefCode As Integer
    Dim rst_chf As ADODB.Recordset
    Dim rst_Vsf As ADODB.Recordset

    llSifCode = 0
    On Error GoTo ErrHand:
    If sgSpfUseProdSptScr = "P" Then
        SQLQuery = "SELECT chfVefCode, sdfVefCode"
        SQLQuery = SQLQuery & " FROM SDF_Spot_Detail, CHF_Contract_Header"
        SQLQuery = SQLQuery & " WHERE sdfCode = " & llSdfCode
        SQLQuery = SQLQuery & " AND chfCode = sdfChfCode"
        Set rst_chf = gSQLSelectCall(SQLQuery)
        If Not rst_chf.EOF Then
            If rst_chf!chfVefCode < 0 Then
                ilSchVefCode = rst_chf!sdfVefCode
                llVsfCode = -rst_chf!chfVefCode
                Do While llVsfCode > 0
                    SQLQuery = "SELECT *"
                    SQLQuery = SQLQuery & " FROM VSF_Veh_Slsp_Combos"
                    SQLQuery = SQLQuery & " WHERE vsfCode = " & llVsfCode
                    Set rst_Vsf = gSQLSelectCall(SQLQuery)
                    If Not rst_Vsf.EOF Then
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode1, rst_Vsf!vsfFSComm1, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode2, rst_Vsf!vsfFSComm2, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode3, rst_Vsf!vsfFSComm3, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode4, rst_Vsf!vsfFSComm4, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode5, rst_Vsf!vsfFSComm5, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode6, rst_Vsf!vsfFSComm6, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode7, rst_Vsf!vsfFSComm7, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode8, rst_Vsf!vsfFSComm8, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode9, rst_Vsf!vsfFSComm9, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode10, rst_Vsf!vsfFSComm10, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode11, rst_Vsf!vsfFSComm11, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode12, rst_Vsf!vsfFSComm12, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode13, rst_Vsf!vsfFSComm13, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode14, rst_Vsf!vsfFSComm14, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode15, rst_Vsf!vsfFSComm15, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode16, rst_Vsf!vsfFSComm16, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode17, rst_Vsf!vsfFSComm17, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode18, rst_Vsf!vsfFSComm18, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode19, rst_Vsf!vsfFSComm19, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode20, rst_Vsf!vsfFSComm20, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode21, rst_Vsf!vsfFSComm21, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode22, rst_Vsf!vsfFSComm22, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode23, rst_Vsf!vsfFSComm23, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode24, rst_Vsf!vsfFSComm24, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode25, rst_Vsf!vsfFSComm25, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode26, rst_Vsf!vsfFSComm26, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode27, rst_Vsf!vsfFSComm27, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode28, rst_Vsf!vsfFSComm28, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode29, rst_Vsf!vsfFSComm29, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode30, rst_Vsf!vsfFSComm30, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode31, rst_Vsf!vsfFSComm31, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode32, rst_Vsf!vsfFSComm32, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode33, rst_Vsf!vsfFSComm33, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode34, rst_Vsf!vsfFSComm34, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode35, rst_Vsf!vsfFSComm35, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode36, rst_Vsf!vsfFSComm36, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode37, rst_Vsf!vsfFSComm37, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode38, rst_Vsf!vsfFSComm38, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode39, rst_Vsf!vsfFSComm39, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode40, rst_Vsf!vsfFSComm40, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode41, rst_Vsf!vsfFSComm41, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode42, rst_Vsf!vsfFSComm42, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode43, rst_Vsf!vsfFSComm43, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode44, rst_Vsf!vsfFSComm44, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode45, rst_Vsf!vsfFSComm45, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode46, rst_Vsf!vsfFSComm46, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode47, rst_Vsf!vsfFSComm47, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode48, rst_Vsf!vsfFSComm48, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode49, rst_Vsf!vsfFSComm49, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        If mGetSifCodeFromVsf(rst_Vsf!vsfFSCode50, rst_Vsf!vsfFSComm50, ilSchVefCode, llSifCode) Then
                            Exit Do
                        End If
                        llVsfCode = rst_Vsf!vsfLkVsfCode
                    Else
                        Exit Do
                    End If
                Loop
                rst_Vsf.Close
            End If
        End If
        rst_chf.Close
        gGetShortTitle = gGetProdOrShtTitle(llSdfCode, llSifCode, ilMaxAbbrChar)
    Else
        gGetShortTitle = gGetProdOrShtTitle(llSdfCode, llSifCode, ilMaxAbbrChar)
    End If
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modGenSubs-gGetShortTitle"
End Function

Function mGetSifCodeFromVsf(ilVsfFSCode As Integer, llVsfFSComm As Long, ilSchVefCode As Integer, llSifCode As Long) As Integer
    If ilVsfFSCode > 0 Then
        If ilVsfFSCode = ilSchVefCode Then
            llSifCode = llVsfFSComm
            mGetSifCodeFromVsf = True
            Exit Function
        End If
    End If
    mGetSifCodeFromVsf = False
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gLongToStrDec                   *
'*                                                     *
'*             Created:5/6/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Format as money or percentage   *
'*                                                     *
'*******************************************************
Function gLongToStrDec(llInNumber As Long, ilNoDecPlaces As Integer) As String
'
'   slOutStr = gLongToStrDec(llInNumber, ilNoDecPlaces)
'   Where:
'       llInNumber(I)- Number to be converted to string with/without dec point
'       ilNoDecPlaces(I)-Number of decimal places (can be zero)
'       slOutStr(O)- Output string format with a decimal point
'
    Dim slStr As String
    Dim ilNegSign As Integer

    If llInNumber < 0 Then
        ilNegSign = True
    Else
        ilNegSign = False
    End If
    slStr = Trim$(Str$(Abs(llInNumber)))
    If ilNoDecPlaces > 0 Then
        If Len(slStr) >= ilNoDecPlaces Then
            slStr = Left$(slStr, Len(slStr) - ilNoDecPlaces) & "." & right$(slStr, ilNoDecPlaces)
        Else
            Do While Len(slStr) < ilNoDecPlaces
                slStr = "0" & slStr
            Loop
            slStr = "." & slStr
        End If
    'Else
        'slStr = slStr & "."
    End If
    If ilNegSign Then
        slStr = "-" & slStr
    End If
    gLongToStrDec = slStr
    Exit Function
End Function


Public Function gGetLatestAttCode(shttCode As Integer, vefCode As Integer) As Long

    Dim rst As ADODB.Recordset
    Dim llAttCode As Long
    Dim llTempAttCode As Long
    Dim slTempDate As String
    Dim slLatestDate As String
    Dim ilCBS As Integer

    On Error GoTo ErrHand
    slLatestDate = "1/1/1969"
    gGetLatestAttCode = 0
    
    SQLQuery = " Select attCode, attOnAir, attDropDate, attOffAir FROM Att"
    SQLQuery = SQLQuery + " WHERE attshfCode = " & shttCode
    SQLQuery = SQLQuery + " AND attVefCode = " & vefCode
    Set rst = gSQLSelectCall(SQLQuery)
    
    If rst.EOF Then
        gGetLatestAttCode = 0
        Exit Function
    End If
    
    While Not rst.EOF
        'We want the lesser date of the attDropdate and the attOffAir
        If DateValue(rst!attDropDate) <= DateValue(rst!attOffAir) Then
            slTempDate = rst!attDropDate
        Else
            slTempDate = rst!attOffAir
        End If
        
        ilCBS = False  'default to cancel before start
        'Check to see if the agreement is a Cancel Before Start, if so, do nothing
        If DateValue(rst!attOnAir) >= DateValue(slTempDate) Then
            ilCBS = True
        Else
            ilCBS = False
        End If
        
        If Not ilCBS Then
            If DateValue(slTempDate) > DateValue(slLatestDate) Then
                slLatestDate = slTempDate
                llTempAttCode = rst!attCode
                gGetLatestAttCode = llTempAttCode
            End If
        End If
        rst.MoveNext
    Wend

Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modGenSubs-gGetLatestAgrnmt"
    Exit Function
End Function

Public Function gMsgBox(Optional sMsg As String = "Press OK to Continue", Optional vbIcon = vbInformation, Optional sTitle As String = "Affiliate") As Integer
    
    'D.S. 10/11/07
    'Warning: One thing to remember is that if you are expecting a return value from a gMsgBox
    'and the gMsgBox is turned OFF then you need to make sure that you handle that case in the code.
    'example:   ilRet = gMsgBox "xxxx"
    
    Dim ilMouse As Integer
    Dim ilRet As Integer
    
    If igShowMsgBox Then
        'Save the mouse pointer
        ilMouse = Screen.MousePointer
        Screen.MousePointer = vbDefault
        If InStr(1, sMsg, "Error #0") = 0 Then
            ilRet = MsgBox(sMsg, vbIcon, sTitle)
        Else
            ilRet = vbOK
        End If
        Screen.MousePointer = ilMouse
        gMsgBox = ilRet
    '6394
    ElseIf igExportSource = 2 And Len(sgExportResultName) > 0 Then
        gLogMsgWODT "W", hgExportResult, sMsg
        igExportReturn = 2
        igReportReturn = 2
    Else
        gLogMsg "gMsgBox: " & sMsg, "AffErrorLog.Txt", False
        igExportReturn = 2
        igReportReturn = 2
    End If
End Function

Function gGetLocalTZName() As String
    
    Dim tlTimeZone As TIME_ZONE_INFORMATION
    Dim llResult As Long
    'Dim ll As Long
    Dim slStr As String
    Dim slTZName As String
    slTZName = ""
    llResult = GetTimeZoneInformation&(tlTimeZone)
    Select Case llResult
        Case 0&, 1& 'use standard time
            'GetLocalTZ = -(objTimeZone.Bias + objTimeZone.StandardBias) * 60 'into minutes
            'For ll = 0 To 31
            '    If tlTimeZone.StandardName(ll) = 0 Then Exit For
            '    slTZName = slTZName & Chr(tlTimeZone.StandardName(ll))
            'Next
            slStr = tlTimeZone.StandardName
            slTZName = StrConv(slStr, vbFromUnicode)
        Case 2& 'use daylight savings time
            'GetLocalTZ = -(objTimeZone.Bias + objTimeZone.DaylightBias) * 60 'into minutes
            'For ll = 0 To 31
            '    If tlTimeZone.DaylightName(ll) = 0 Then Exit For
            '    slTZName = slTZName & Chr(tlTimeZone.DaylightName(ll))
            'Next
            slStr = tlTimeZone.StandardName
            slTZName = StrConv(slStr, vbFromUnicode)
    End Select
    gGetLocalTZName = slTZName
End Function


'-----------------------------------------------------------------------------------
' Get the system's MAC address(es) via GetAdaptersInfo API function (IPHLPAPI.DLL)
'
' Note: GetAdaptersInfo returns information about physical adapters
'-----------------------------------------------------------------------------------
Public Function gGetMACs_AdaptInfo() As String

    Dim AdapInfo As IP_ADAPTER_INFO, bufLen As Long, sts As Long
    Dim retStr As String, numStructs%, i%, IPinfoBuf() As Byte, srcPtr As Long
    
    
    ' Get size of buffer to allocate
    sts = GetAdaptersInfo(AdapInfo, bufLen)
    If (bufLen = 0) Then Exit Function
    numStructs = bufLen / Len(AdapInfo)
    retStr = numStructs & " Adapter(s):" & vbCrLf
    
    ' reserve byte buffer & get it filled with adapter information
    ' !!! Don't Redim AdapInfo array of IP_ADAPTER_INFO,
    ' !!! because VB doesn't allocate it contiguous (padding/alignment)
    ReDim IPinfoBuf(0 To bufLen - 1) As Byte
    sts = GetAdaptersInfo(IPinfoBuf(0), bufLen)
    If (sts <> 0) Then Exit Function
    
    ' Copy IP_ADAPTER_INFO slices into UDT structure
    srcPtr = VarPtr(IPinfoBuf(0))
    For i = 0 To numStructs - 1
        If (srcPtr = 0) Then Exit For
'        CopyMemory AdapInfo, srcPtr, Len(AdapInfo)
        CopyMemory AdapInfo, ByVal srcPtr, Len(AdapInfo)
        
        ' Extract Ethernet MAC address
        With AdapInfo
            If (.AdapterType = MIB_IF_TYPE_ETHERNET) Then
                'retStr = retStr & vbCrLf & "[" & i & "] " & sz2string(.Description) _
                '        & vbCrLf & vbTab & MAC2String(.MACaddress) & vbCrLf
                retStr = mMAC2String(.MACaddress)
                Exit For
            End If
        End With
        srcPtr = AdapInfo.Next
    Next i
    
    ' Return list of MAC address(es)
    gGetMACs_AdaptInfo = retStr
    
End Function

' Convert a byte array containing a MAC address to a hex string
Private Function mMAC2String(AdrArray() As Byte) As String
    Dim aStr As String, hexStr As String, i%
    
    For i = 0 To 5
        If (i > UBound(AdrArray)) Then
            hexStr = "00"
        Else
            hexStr = Hex$(AdrArray(i))
        End If
        
        If (Len(hexStr) < 2) Then hexStr = "0" & hexStr
        aStr = aStr & hexStr
        If (i < 5) Then aStr = aStr & "-"
    Next i
    
    mMAC2String = aStr
    
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gDecryptField                   *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Decrypt data field              *
'*                                                     *
'*******************************************************
Function gDecryptField(slEncryptName As String) As String
    Dim slDecrypt As String
    Dim slName As String
    Dim ilLen As Integer
    Dim ilLoop As Integer

    slName = slEncryptName
    ilLen = Len(slEncryptName)
    slDecrypt = ""
    For ilLoop = 1 To ilLen Step 1
        slDecrypt = slDecrypt & Chr(Asc(slName) - 128)
        slName = Mid$(slName, 2)
    Next ilLoop
    gDecryptField = slDecrypt
End Function



'*********************************************************************************
'
'*********************************************************************************
Public Sub gCheckForContFiles()
    Dim ilRet As Integer
    Dim slMsg As String
    Dim imSalesperson As Boolean
    Dim slLastBackupDateTime As String
    Dim slCurDateTime As String
    Dim llTotalHours As Long
    Dim ilValue As Integer
    Dim SvrRsp_FilesStuckInCntMode As CSISvr_Rsp_Answer

    On Error GoTo ErrHand

    If (StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0) Or (StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0) Then
        Exit Sub
    End If


    ilRet = csiCheckForFilesStuckInCntMode(sgDBPath, SvrRsp_FilesStuckInCntMode)
    If SvrRsp_FilesStuckInCntMode.iAnswer = 1 Then
        slMsg = ""
        slMsg = slMsg & "<<< WARNING >>>" & vbCrLf
        slMsg = slMsg & "<<< Your last database back failed. >>>" & vbCrLf & vbCrLf
        slMsg = slMsg & "Adding to or editing information while in this condition could result in data loss and/or data corruption." & vbCrLf & vbCrLf
        slMsg = slMsg & "Although you may continue to view information, it is imperative that you call Counterpoint or email Counterpoint at service@counterpoint.net ASAP to remedy this condition." & vbCrLf
        gMsgBox slMsg, vbCritical, "Backup Failure"
        gLogMsg "User was warned that files are stuck in continuous mode.", "AffErrorLog.Txt", False
        Exit Sub
    End If
    If Not gUsingCSIBackup Then
        ' CSI Backups are not turned on.
        Exit Sub
    End If

    slLastBackupDateTime = gGetLastBackupDateTime()
    If LenB(slLastBackupDateTime) > 0 Then
        slCurDateTime = gNow()
        llTotalHours = DateDiff("h", slLastBackupDateTime, slCurDateTime)
        If llTotalHours > 24 And igShowVersionNo <> 2 Then
            slMsg = ""
            slMsg = slMsg & "<<< WARNING >>>" & vbCrLf
            slMsg = slMsg & "<<< A database backup has not occurred in over 24 hours.  >>>" & vbCrLf & vbCrLf
            gMsgBox slMsg, vbExclamation, "Backup Notice"
            gLogMsg "User was warned that backup has not been performed within 24 hours.", "AffErrorLog.Txt", False
        End If
    Else
        slMsg = "<<< A database backup has never been performed.  >>>" & vbCrLf & vbCrLf
        gMsgBox slMsg, vbExclamation, "Backup Notice"
        gLogMsg "User was warned that backup has never been performed.", "AffErrorLog.Txt", False
    End If
    Exit Sub
'    slCurDateTime = gNow()
'    llTotalHours = DateDiff("h", slLastBackupDateTime, slCurDateTime)
'    If llTotalHours > 24 Then
'        slMsg = ""
'        slMsg = slMsg & "<<< WARNING >>>" & vbCrLf
'        slMsg = slMsg & "<<< A database backup has not occurred in over 24 hours.  >>>" & vbCrLf & vbCrLf
'        gMsgBox slMsg, vbExclamation, "Backup Notice"
'        gLogMsg "User was warned that backup has not been performed within 24 hours.", "AffErrorLog.Txt", False
'    End If
'    Exit Sub

ErrHand:
    gLogMsg "A general error occurred in gCheckForContFiles", "AffErrorLog.Txt", False
End Sub

'*********************************************************************************
'
'*********************************************************************************
Public Function gGetLastBackupDateTime() As String
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim SvrRsp_GetLastBackupDate As CSISvr_Rsp_GetLastBackupDate

    gGetLastBackupDateTime = ""
    ilRet = csiGetLastBackupDate(sgDBPath, SvrRsp_GetLastBackupDate)
    slDateTime = SvrRsp_GetLastBackupDate.sLastBackupDateTime
    gGetLastBackupDateTime = gRemoveIllegalChars(slDateTime)
    Exit Function
    
ErrHand:
    MsgBox "A general error has occurred in gGetLastBackupDateTime."
End Function

'*********************************************************************************
'
'*********************************************************************************
Public Function gGetLastCopyDateTime() As String
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim SvrRsp_GetLastCopyDate As CSISvr_Rsp_GetLastBackupDate

    gGetLastCopyDateTime = ""
    ilRet = csiGetLastCopyDate(sgDBPath, SvrRsp_GetLastCopyDate)
    slDateTime = SvrRsp_GetLastCopyDate.sLastBackupDateTime
    gGetLastCopyDateTime = gRemoveIllegalChars(slDateTime)
    Exit Function
    
ErrHand:
    MsgBox "A general error has occurred in gGetLastCopyDateTime."
End Function

'*********************************************************************************
'
'*********************************************************************************
Public Function gIsBackupRunning() As Boolean
    Dim ilRet As Integer
    Dim SvrRsp_IsBackupRunning As CSISvr_Rsp_Answer

    On Error GoTo ErrHand

    gIsBackupRunning = False
    ilRet = csiIsBackupRunning(sgDBPath, SvrRsp_IsBackupRunning)
    If ilRet <> 0 Then
        Exit Function
    End If
    If SvrRsp_IsBackupRunning.iAnswer = 1 Then
        gIsBackupRunning = True
    End If
    Exit Function

ErrHand:
    MsgBox "A general error has occurred in IsBackupRunning."
End Function

'*********************************************************************************
'
'*********************************************************************************
Public Function gLoadINIValue(sPathFileName As String, Section As String, Key As String, sValue As String) As Boolean
    On Error GoTo ErrHand
    Dim BytesCopied As Integer
    Dim sBuffer As String * 128

    gLoadINIValue = False
    BytesCopied = GetPrivateProfileString(Section, Key, "Not Found", sBuffer, 128, sPathFileName)
    If BytesCopied > 0 Then
        If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
            sValue = Left(sBuffer, BytesCopied)
            gLoadINIValue = True
        End If
    End If
    Exit Function

ErrHand:
    ' return now if an error occurs
End Function

'*********************************************************************************
'
'*********************************************************************************
Public Function gSaveINIValue(sPathFileName As String, Section As String, Key As String, sValue As String) As Boolean
    On Error GoTo ErrHand
    Dim BytesCopied As Integer

    gSaveINIValue = False
    If WritePrivateProfileString(Section, Key, sValue, sPathFileName) Then
        gSaveINIValue = True
    End If
    Exit Function

ErrHand:
    ' return now if an error occurs
End Function

Public Function gRemoveIllegalChars(slInStr As String) As String
    Dim slStr As String
    Dim ilLen As Integer
    Dim ilLoop As Integer
    Dim slOneChar As String
    
    ilLen = Len(slInStr)
    slStr = ""
    For ilLoop = 1 To ilLen Step 1
        slOneChar = Mid(slInStr, ilLoop, 1)
        If (Asc(slOneChar) >= Asc(" ")) And (Asc(slOneChar) <= Asc("~")) Then
            slStr = slStr & slOneChar
        End If
    Next
    gRemoveIllegalChars = Trim$(slStr)
End Function

Public Function gRemoveIllegalCharsAndLog(slInStr As String, sFileName As String, lLineNo As Long, bLeaveCR As Boolean) As String
    Dim slStr As String
    Dim ilLen As Integer
    Dim ilLoop As Integer
    Dim slOneChar As String
    Dim blFound As Boolean
    Dim slPos As String
    Dim slvalue As String
    Dim llline As Long
    
    ilLen = Len(slInStr)
    slStr = ""
    slPos = ""
    slvalue = ""
    blFound = False
    '11/20/19: Added below
    ''D.S. 11/15/19 Added two Replace statements below on Dick's request.  # and % were causing the web to choke when sending JAVA to IIS per Jeff
    'slInStr = Replace(slInStr, "%", "")
    'slInStr = Replace(slInStr, "#", "")
    For ilLoop = 1 To ilLen Step 1
        slOneChar = Mid(slInStr, ilLoop, 1)
        'bLeave indicates whether or not to remove carrage return and line feeds
        If bLeaveCR Then
            If (Asc(slOneChar) >= Asc(" ")) And (Asc(slOneChar) <= Asc("~")) And (Asc(slOneChar) <> Asc("#")) And (Asc(slOneChar) <> Asc("%")) And (Asc(slOneChar) <> Asc("+")) Then
                slStr = slStr & slOneChar
            Else
                blFound = True
                bgIllegalCharsFound = True
                slPos = slPos & ilLoop & ","
                slvalue = slvalue & Asc(slOneChar) & ","
            End If
        Else
            If (Asc(slOneChar) >= Asc(" ")) And (Asc(slOneChar) <= Asc("~")) And (Asc(slOneChar) <> 10) And (Asc(slOneChar) <> 13) And (Asc(slOneChar) <> Asc("#")) And (Asc(slOneChar) <> Asc("%")) And (Asc(slOneChar) <> Asc("+")) Then
                slStr = slStr & slOneChar
            Else
                blFound = True
                bgIllegalCharsFound = True
                slPos = slPos & ilLoop & ","
                slvalue = slvalue & Asc(slOneChar) & ","
            End If
        End If
    Next
    If blFound Then
        llline = lLineNo + 1
        gLogMsg "Illegal character(s) found: " & sFileName & ", Line: " & llline & ", Position(s): " & slPos & " Value(s): " & slvalue & vbCrLf & "[" & slInStr & "]", "AffBadCharLog.Txt", False
    End If
    gRemoveIllegalCharsAndLog = Trim$(slStr)
End Function



Public Function gAdjTimeToEasternZone(sDateTime As String) As String
    Dim slLocalTimeZone As String
    Dim slNewDateTime As String

    On Error GoTo ErrHand:
    gAdjTimeToEasternZone = sDateTime  ' return what was passed in if any errors occur.

    slLocalTimeZone = gGetLocalTZName()
    Select Case Left$(gGetLocalTZName(), 1)
        Case "E"
        Case "C"
            slNewDateTime = DateAdd("s", 3600, sDateTime)
        Case "M"
            slNewDateTime = DateAdd("s", 7200, sDateTime)
        Case "P"
            slNewDateTime = DateAdd("s", 10800, sDateTime)
    End Select
    gAdjTimeToEasternZone = slNewDateTime
    Exit Function
ErrHand:

End Function


Public Function gAddQuotes(slInStr As String) As String
    gAddQuotes = """" & slInStr & """"
End Function



Public Function gStripChr0(slInStr As String) As String
    Dim slStr As String
    Dim slChr As String
    Dim ilIndex As Integer
    
    slStr = ""
    If Len(slInStr) > 0 Then
        ilIndex = 1
        slChr = Mid(slInStr, ilIndex, 1)
        Do While slChr <> Chr$(0)
            slStr = slStr & slChr
            ilIndex = ilIndex + 1
            If ilIndex > Len(slInStr) Then
                Exit Do
            End If
            slChr = Mid(slInStr, ilIndex, 1)
        Loop
    End If
    gStripChr0 = Trim$(slStr)
End Function

Public Function gStripIllegalChr(slInStr As String) As String
    Dim slStr As String
    Dim slChr As String
    Dim ilIndex As Integer
    
    slStr = ""
    If Len(slInStr) > 0 Then
        ilIndex = 1
        Do While ilIndex <= Len(slInStr)
            slChr = Mid(slInStr, ilIndex, 1)
            If (Asc(slChr) >= Asc(" ")) And (Asc(slChr) <= Asc("~")) Then
                slStr = slStr & slChr
            End If
            ilIndex = ilIndex + 1
        Loop
    End If
    gStripIllegalChr = Trim$(slStr)
End Function
Public Function gStrongPassword(slPassword As String) As Integer
    Dim ilNoUCaseLetters As Integer
    Dim ilNoLCaseLetters As Integer
    Dim ilNoNumbers As Integer
    Dim ilNoSymbols As Integer
    Dim ilLoop As Integer
    Dim ilCheck As Integer
    Dim slStr As String
    
    slStr = Trim(slPassword)
    'dan get out if empty
     If Len(slStr) = 0 Then
        gStrongPassword = False
        Exit Function
    End If
    ilNoUCaseLetters = 0
    ilNoLCaseLetters = 0
    ilNoNumbers = 0
    ilNoSymbols = 0
    'ttp 5608. fixing login issues.  Use bgStrongPassword as it's set at load of program
    If bgStrongPassword Then
   ' If ((Asc(sgSpfUsingFeatures2) And STRONGPASSWORD) = STRONGPASSWORD) Then
        For ilLoop = 1 To Len(slStr) Step 1
            ilCheck = Asc(Mid$(slStr, ilLoop, 1))
            If (ilCheck >= Asc("A")) And (ilCheck <= Asc("Z")) Then
                ilNoUCaseLetters = ilNoUCaseLetters + 1
            Else
                If (ilCheck >= Asc("a")) And (ilCheck <= Asc("z")) Then
                    ilNoLCaseLetters = ilNoLCaseLetters + 1
                Else
                    If (ilCheck >= Asc("0")) And (ilCheck <= Asc("9")) Then
                        ilNoNumbers = ilNoNumbers + 1
                    Else
                        ilNoSymbols = ilNoSymbols + 1
                    End If
                End If
            End If
        Next ilLoop
        If ((ilNoUCaseLetters > 0) Or (ilNoLCaseLetters > 0)) And (ilNoNumbers > 0) And (ilNoSymbols > 0) And (Len(slStr) >= 8) Then
            gStrongPassword = True
        Else
            gStrongPassword = False
        End If
    Else
        gStrongPassword = True
    End If

End Function

Public Function gAdjustDrivePath(slInDrivePath As String) As String
    'This routine is used to check if the application is running on the server and the drive must be converted
    Dim slCurDir As String
    Dim slDrivePath As String
    Dim ilLoop As Integer
    Dim ilLoop2 As Integer
    Dim slINIPathParts() As String
    Dim slCurPathParts() As String
    
    slDrivePath = Trim$(slInDrivePath)
    gAdjustDrivePath = slDrivePath
    
    ' JD - 12-18-09
    ' Due to a random error we received at Reach Media, this function is temporarily disabled
    ' until we figure out what to do.
    Exit Function
    
    slCurDir = Trim$(sgCurDir)
    If InStr(1, slDrivePath, "//", vbTextCompare) > 0 Then
        Exit Function
    End If
    If InStr(1, slDrivePath, "C:", vbTextCompare) > 0 Then
        Exit Function
    End If
    If InStr(1, slDrivePath, "D:", vbTextCompare) > 0 Then
        Exit Function
    End If
    If (Len(slDrivePath) <= 1) Or (Len(slCurDir) <= 1) Then
        Exit Function
    End If
    If StrComp(Left$(slCurDir, 1), Left(slDrivePath, 1), vbTextCompare) = 0 Then
        Exit Function
    End If
    'Find common parts
    slINIPathParts = Split(slDrivePath, "\")
    slCurPathParts = Split(slCurDir, "\")

    ' Find the first part past the drive letter that matches.
    For ilLoop = 1 To UBound(slINIPathParts)
        If ilLoop > UBound(slCurPathParts) Then
            Exit Function
        End If
        If slINIPathParts(ilLoop) = slCurPathParts(1) Then
            ' Take the local drive and add the rest of this path to it.
            gAdjustDrivePath = ""
            For ilLoop2 = 0 To UBound(slCurPathParts)
                 gAdjustDrivePath = gAdjustDrivePath & slCurPathParts(ilLoop2) & "\"
            Next
            MsgBox "Path has been adjusted to " & gAdjustDrivePath
            Exit Function
        End If
    Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gAdjustScreenMessage            *
'*                                                     *
'*             Created:8/28/09       By:D. Michaelson  *
'*            Modified:              By:               *
'*                                                     *
'*Comments:'printing..' messages uniform look          *
'*                                                     *
'*******************************************************
Public Sub gAdjustScreenMessage(ByRef myForm As Form, ByRef plcMyBox As PictureBox)
Const IMAGESIZE = 650
Const IMAGEWEBPADHEIGHT = 260   'scrollbar correction
Const IMAGEWEBPADWIDTH = 190    'scrollbar correction
Const RIGHTMARGIN = 200         'lengthen main box additional amount
Const PADRIGHT = 100            'move away from edge.
Const IMAGENAME = "Wait.gif"
Dim picBox As PictureBox
Dim wbcAnimate As WebBrowser
Dim slGifPath As String
slGifPath = sgReportDirectory & IMAGENAME
    'modify message picture box
    plcMyBox.BackColor = vbButtonFace
    plcMyBox.Width = plcMyBox.Width + IMAGESIZE + IMAGEWEBPADWIDTH + RIGHTMARGIN
    'create picturebox to hold webbrowswer --need to constrict webbrowser to lose frame
    Set picBox = myForm.Controls.Add("VB.PictureBox", "picBox")
    With picBox
        .Height = IMAGESIZE + IMAGEWEBPADHEIGHT
        .Width = IMAGESIZE + IMAGEWEBPADHEIGHT
        Set .Container = plcMyBox
        ' reposition original picture box
        .Left = plcMyBox.Width - IMAGESIZE - IMAGEWEBPADWIDTH - PADRIGHT
        .BorderStyle = 0
        'manipulate to lose frame
        .ScaleHeight = IMAGESIZE
        .ScaleWidth = IMAGESIZE
        .Visible = True
    End With
    'create webbrowser
    Set wbcAnimate = myForm.Controls.Add("shell.explorer.2", "MyGif")
    With wbcAnimate
        Set .Container = picBox
        .Width = IMAGESIZE + IMAGEWEBPADWIDTH
        .Height = IMAGESIZE + IMAGEWEBPADHEIGHT
        .Move -myForm.ScaleX(2, vbPixels), -myForm.ScaleY(2, vbPixels)
        'change image source!
        ' .Navigate "about:<html><body scroll='no'><img src='D:\vbAnimatePic\wait.gif'></img></body></html>"
         .Navigate "about:<html><body scroll='no'><img src='" & slGifPath & "'></img></body></html>"
        .Document.bgcolor = "#ece9d8"
        .Visible = True
    End With
End Sub
'Public Function gSendEmail(tlMyEmailInfo As EmailInformation, Optional ByRef ctrResultBox As control, Optional ZipBox As dzactxctrl, Optional Mail As MailSender) As Boolean
'' In: tlMyEmailInfo-- to/from/subject/message/attachment.  bUserFromHasPriority is confusing: use the values here over default(site options), or rather use default first and these values if default blank?
''                  -- blank to/from and blank default will mean message not sent.
'' In: ctrResultBox control to display result  text or listbox
'' In: Mail object.  frmLogEmail sets the 'to' of the mail object on the form, but could pass CC:, reply, etc.
'Dim slErrorFact As String
'Dim slNonZipFact As String
'Dim blThrowError As Boolean
'Dim slAttachments() As String
'Dim slFailedToZip() As String
'Dim slZipFile As String
'Dim blDeleteZipFile As Boolean
'Dim slWord As String    'variable word in string for zip failure
''Dan M 9/7/10 added to keep in sync with affiliate
'Dim slNames() As String
'Dim c As Integer
'Const MAILREGKEY = "16374-29451-54460"
'Dim rstSite As ADODB.Recordset
'
'On Error GoTo ErrHand
'    SQLQuery = "SELECT * FROM Site Where siteCode = 1"
'    Set rstSite = gSQLSelectCall(SQLQuery)
'    If Not rstSite.EOF Then
'        If Mail Is Nothing Then
'            Set Mail = New ASPEMAILLib.MailSender
'        End If
'        Mail.RegKey = MAILREGKEY
'        'to allow verifying, passed mailsender will have host and other variables set.
'        If LenB(Mail.Host) = 0 Then
'            'smtp
'            If Trim$(rstSite!siteEmailHost) = "" Then
'                Mail.Host = ""
'                slErrorFact = "SMTP Host is undefined in Site Options."
'                blThrowError = True
'            Else
'                Mail.Host = Trim$(rstSite!siteEmailHost)
'            End If
'            'port
'            If rstSite!siteEmailPort = 0 Then
'                Mail.Port = 0
'                slErrorFact = slErrorFact & " Port Number is undefined in Site Options"
'                blThrowError = True
'            Else
'                Mail.Port = Trim$(rstSite!siteEmailPort)
'            End If
'            'username
'            If LenB(Trim(rstSite!siteEmailAcctName)) = 0 Then
'                Mail.UserName = ""
'            Else
'                Mail.UserName = Trim$(rstSite!siteEmailAcctName)
'            End If
'            'password
'            If LenB(Trim(rstSite!siteEmailPassword)) = 0 Then
'                Mail.Password = ""
'            Else
'                Mail.Password = Trim$(rstSite!siteEmailPassword)
'            End If
'        End If
'        If Len(Mail.FromName) = 0 Then
'            Mail.FromName = tlMyEmailInfo.sFromName
'        End If
'        ' from address must be filled or will fail
'         'email site options no longer exist.
'        If Len(Mail.From) = 0 Then
'            If (LenB(Trim(tlMyEmailInfo.sFromAddress)) = 0) Then
'                slErrorFact = slErrorFact & " No available 'from' address"
'                blThrowError = True
'            Else
'                Mail.From = tlMyEmailInfo.sFromAddress
'            End If
'        End If
'         If Mail.ValidateAddress(Mail.From) <> 0 Then
'             slErrorFact = slErrorFact & " 'from' address is not a valid address."
'             blThrowError = True
'         End If
'         ' receipt? add this code Dan M 9/17/10
'        ' Mail.AddCustomHeader "Return-Receipt-To: " & Mail.From
'        ' Mail.AddCustomHeader "Disposition-Notification-To: " & Mail.From
'         ' To address, To Name some mail objects have already set to address before coming to this procedure
'         If LenB(tlMyEmailInfo.sToAddress) > 0 Then
'             If Mail.ValidateAddress(tlMyEmailInfo.sToAddress) <> 0 Then
'                 slErrorFact = slErrorFact & " 'To' address is not a valid address."
'                 blThrowError = True
'             Else
'                 Mail.AddAddress tlMyEmailInfo.sToAddress, tlMyEmailInfo.sToName
'             End If
'        End If
'            'Dan added 9/7/10 for....changed type to include more To addresses, CC, Bcc
'         'To addresses
'         If Len(tlMyEmailInfo.sToMultiple) > 0 Then
'             slNames = Split(tlMyEmailInfo.sToMultiple, ",")
'             If Not mTestAddresses(slNames, Mail) Then
'                 slErrorFact = slErrorFact & " One of the multiple 'To' addresses is not a valid address."
'                 blThrowError = True
'             Else
'                 For c = 0 To UBound(slNames)
'                     Mail.AddAddress slNames(c)
'                 Next c
'             End If
'         End If
'         'CC
'         If Len(tlMyEmailInfo.sCCMultiple) > 0 Then
'             slNames = Split(tlMyEmailInfo.sCCMultiple, ",")
'             If Not mTestAddresses(slNames, Mail) Then
'                 slErrorFact = slErrorFact & " One of the multiple 'CC' addresses is not a valid address."
'                 blThrowError = True
'             Else
'                 For c = 0 To UBound(slNames)
'                     Mail.AddCC slNames(c)
'                 Next c
'             End If
'         End If
'         'bcc
'         If Len(tlMyEmailInfo.sBCCMulitple) > 0 Then
'             slNames = Split(tlMyEmailInfo.sBCCMulitple, ",")
'             If Not mTestAddresses(slNames, Mail) Then
'                 slErrorFact = slErrorFact & " One of the multiple 'BCC' addresses is not a valid address."
'                 blThrowError = True
'             Else
'                 For c = 0 To UBound(slNames)
'                     Mail.AddBcc slNames(c)
'                 Next c
'             End If
'         End If
'         'subject
'         Mail.Subject = tlMyEmailInfo.sSubject
'        'body
'          If LenB(Trim(tlMyEmailInfo.sMessage)) > 0 Then
'              Mail.Body = tlMyEmailInfo.sMessage
'          Else
'             Mail.Body = " ** No Message **"
'          End If
'          'attachments
'          If LenB(Trim(tlMyEmailInfo.sAttachment)) > 0 Then
'             slAttachments = Split(tlMyEmailInfo.sAttachment, ";")
'             If mAllFilesExist(slAttachments) Then
'                 If ZipBox Is Nothing Then  'don 't zip
'                 'Dan M 9/7/10 now defined at top
'                    ' Dim c As Integer
'                     For c = 0 To UBound(slAttachments)  'test multiple not zipped
'                         Mail.AddAttachment slAttachments(c)
'                     Next c
'                 Else
'                     ReDim slFailedToZip(0)
'                     slZipFile = mZipAllFiles(slAttachments, slFailedToZip, ZipBox)
'                     If (StrComp(slZipFile, "NoXne", vbBinaryCompare) <> 0) Then  ' not error zipping
'                         If (UBound(slAttachments) + 1 <> UBound(slFailedToZip)) Then  '  if all files can't be zipped, don't add attachment
'                             Mail.AddAttachment slZipFile
'                             blDeleteZipFile = True
'                         End If
'                         If UBound(slFailedToZip) > 0 Then   'nothing to do error from zipping: send with unzipped
'                             'code to write out message:
'                             If UBound(slFailedToZip) = UBound(slAttachments) + 1 Then   'all attachments failed
'                                 slWord = " attached file"
'                                 If UBound(slFailedToZip) > 1 Then 'more than one
'                                     slWord = slWord & "s"
'                                 End If
'                             Else
'                                 If UBound(slFailedToZip) > 1 Then
'                                     slWord = " some attached files"     'only some failed
'                                 Else
'                                     slWord = " an attached file"
'                                 End If
'                             End If
'                             slNonZipFact = " But" & slWord & " could not be zipped."
'                             For c = 0 To UBound(slFailedToZip) - 1 'test multiple not zipped
'                                 Mail.AddAttachment slFailedToZip(c)
'                             Next c
'                         End If
'                     Else 'couldn't zip stop email
'                         slErrorFact = slErrorFact & " Email not sent. Attached files could not be zipped."
'                         blThrowError = True
'                     End If 'error zipping
'                 End If  ' zip?
'             Else
'                 slErrorFact = slErrorFact & " Email not sent.  Some attached files do not exist."
'                 blThrowError = True
'             End If 'files exist?
'             Erase slAttachments
'        End If  'attachment?
'         'TLS
'         'dan added tls 6/28/11.  Verify passes value, don't grab from site.
'        If Not tlMyEmailInfo.bTLSSet Then
'            If tlMyEmailInfo.sFromName = "1" Then
'                Mail.TLS = True
'            Else
'                Mail.TLS = False
'            End If
'        End If
'    Else
'        slErrorFact = "Email information is undefined in Traffic's Site Options."
'        blThrowError = True
'    End If  ' found site options
'    On Error Resume Next
'    If Not blThrowError Then
'        gSendEmail = Mail.Send ' send message
'    Else    'no address, attachment file issue, zipping issue
'        Err.Raise 5555, , " "
'    End If
'    If Not ctrResultBox Is Nothing Then
'        With ctrResultBox
'            If TypeOf ctrResultBox Is ListBox Then
'                .Clear
'                If Err <> 0 Then ' error occurred
'                    .ForeColor = vbRed
'                    .AddItem Err.Description & "  " & slErrorFact
'                Else
'                    .ForeColor = vbGreen
'                    .AddItem "Mail sent." & slNonZipFact     'attachments sent but could not be zipped.
'                End If
'            ElseIf TypeOf ctrResultBox Is TextBox Then
'                .Text = ""
'                If Err <> 0 Then ' error occurred
'                    .ForeColor = vbRed
'                    .Text = Err.Description & "  " & slErrorFact
'                Else
'                    .ForeColor = vbGreen
'                    .Text = "Mail sent." & slNonZipFact
'                End If
'            End If  'list box/text box
'        End With
'    End If      'send control?
'    If blDeleteZipFile Then
'        Kill slZipFile
'    End If
'    Erase slNames
'    Set Mail = Nothing
'    If Not rstSite Is Nothing Then
'        If (rstSite.State And adStateOpen) <> 0 Then
'            rstSite.Close
'            Set rstSite = Nothing
'        End If
'    End If
'    Exit Function
'
'End Function
'
'Private Function mTestAddresses(slNames() As String, olMail As MailSender) As Boolean
'    Dim c As Integer
'
'    For c = 0 To UBound(slNames)
'        If olMail.ValidateAddress(slNames(c)) <> 0 Then
'            mTestAddresses = False
'            Exit Function
'        End If
'    Next c
'    mTestAddresses = True
'End Function
'
'
'Private Function mAllFilesExist(slFiles() As String) As Boolean
'    Dim olFile As FileSystemObject
'    Dim c As Integer
'
'    Set olFile = New FileSystemObject
'    For c = 0 To UBound(slFiles)
'        If Not olFile.FileExists(slFiles(c)) Then
'            mAllFilesExist = False
'            Set olFile = Nothing
'            Exit Function
'        End If
'    Next c
'    mAllFilesExist = True
'    Set olFile = Nothing
'End Function
'
''zipping procedures
'Private Function mZipAllFiles(ByRef slAttachments() As String, ByRef slFailure() As String, zpcDZip As dzactxctrl) As String
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilPos                         slStr                                                   *
''******************************************************************************************
'
'    Dim ilRet As Integer
'    Dim slData As String
'    Dim slDateTime As String
'    Dim ilLoop As Integer
'    Dim slZipPathName As String
'    Dim ilIndex As Integer
'    DoEvents
'    slDateTime = " " & Format$(gNow(), "ddmmyy")
'    'Dan traffic/affiliate  sgClientName
'    slZipPathName = sgDBPath & Trim$(sgClientName) & slDateTime & ".zip"  'BuildZipName
'    On Error Resume Next
'    Kill slZipPathName  'if errors zipping, file might exist from before
'    On Error GoTo 0
'    For ilLoop = 0 To UBound(slAttachments)
'        'slData = slData & slAttachments(ilLoop) & " "
'        ilRet = mAddFileToZip(slZipPathName, slAttachments(ilLoop), zpcDZip)
'        If ilRet > 0 Then
'            If ilRet = 12 Then  'nothing to zip
'                ilIndex = UBound(slFailure)
'                ReDim Preserve slFailure(0 To ilIndex + 1)
'                slFailure(ilIndex) = slAttachments(ilLoop)
'            Else    'error
'               mZipAllFiles = "NoXne"
'               gChDrDir
'               Exit Function
'            End If
'        End If
'    Next ilLoop
'    mZipAllFiles = slZipPathName
'    gChDrDir
'   DoEvents
'End Function
'
'Private Function mAddFileToZip(szZip As String, szFile As String, zpcDZip As dzactxctrl) As Integer
'
'    'Init the Zip control structure
'    Call minitZIPCmdStruct(zpcDZip)
'
'    zpcDZip.ZIPFile = szZip    'The ZIP file name
'    zpcDZip.ItemList = szFile  'The file list to be added
'    zpcDZip.BackgroundProcessFlag = True
'    zpcDZip.ActionDZ = ZIP_ADD   'ADD files to the ZIP file
'    'Returns the error code.  This code can be translated by the sub mTranslateErrors.
'    'It is not currently being used to log to a file.
'    mAddFileToZip = zpcDZip.ErrorCode
'
'End Function
'
'' **************************************************************************************
''
''  Procedure:  initZIPCmdStruct()
''
''  Purpose:  Set the ZIP control values
''
'' **************************************************************************************
'Private Sub minitZIPCmdStruct(zpcDZip As dzactxctrl)
'    zpcDZip.ActionDZ = 0 'NO_ACTION
'    zpcDZip.AddCommentFlag = False
'    zpcDZip.AfterDateFlag = False
'    zpcDZip.BackgroundProcessFlag = False
'    zpcDZip.Comment = ""
'    zpcDZip.CompressionFactor = 5
'    zpcDZip.ConvertLFtoCRLFFlag = False
'    zpcDZip.Date = ""
'    zpcDZip.DeleteOriginalFlag = False
'    zpcDZip.DiagnosticFlag = False
'    zpcDZip.DontCompressTheseSuffixesFlag = False
'    zpcDZip.DosifyFlag = False
'    zpcDZip.EncryptFlag = False
'    zpcDZip.FixFlag = False
'    zpcDZip.FixHarderFlag = False
'    zpcDZip.GrowExistingFlag = False
'    zpcDZip.IncludeFollowing = ""
'    zpcDZip.IncludeOnlyFollowingFlag = False
'    zpcDZip.IncludeSysandHiddenFlag = False
'    zpcDZip.IncludeVolumeFlag = False
'    zpcDZip.ItemList = ""
'    zpcDZip.MajorStatusFlag = True
'    zpcDZip.MessageCallbackFlag = True
'    zpcDZip.MinorStatusFlag = True
'    zpcDZip.MultiVolumeControl = 0
'    zpcDZip.NoDirectoryEntriesFlag = True
'    zpcDZip.NoDirectoryNamesFlag = True
'
'    zpcDZip.OldAsLatestFlag = False
'    zpcDZip.PathForTempFlag = False
'    zpcDZip.QuietFlag = False
'    zpcDZip.RecurseFlag = False
'    zpcDZip.StoreSuffixes = ""
'    zpcDZip.TempPath = ""
'    zpcDZip.ZIPFile = ""
'
'    'Write out a log file in the windows sub directory
'    zpcDZip.ZipSubOptions = 256
'
'    ' added for rev 3.00
'    zpcDZip.RenameCallbackFlag = False
'    zpcDZip.ExtProgTitle = ""
'    zpcDZip.ZIPString = ""
'    'Dan m 9/14/09 don't show error message
'    zpcDZip.AllQuiet = True
'End Sub
'
'Public Sub gInsertNewFromAddress(slAddress As String)
''Dan M no cef for email? update
'Dim olMail As ASPEMAILLib.MailSender
''Dim slAddress As String
'Dim llNewCefCode As Long
'Dim slQuery As String
'    'slAddress = txtFromEmail
'    Set olMail = New ASPEMAILLib.MailSender
'    If olMail.ValidateAddress(slAddress) = 0 Then
'        slQuery = "INSERT into cef_comments_events (cefCode, cefComment) values( Replace, '" & slAddress & "')"
'        llNewCefCode = gInsertAndReturnCode(slQuery, "cef_comments_events", "cefCode", "Replace")
'        If llNewCefCode > 0 Then
'            SQLQuery = "UPDATE ust SET ustEMailCefCode = " & llNewCefCode & " WHERE ustname = '" & sgUserName & "'"
'            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then  'fail to update, remove from cef. no message box
'            'Dan M 9/7/10 remove name of form...
'                'gMsg = "A general error has occurred in frmEmail-mInsertNewFromAddress: "
'                gMsg = "A general error has occurred in mInsertNewFromAddress(modGenSubs): "
'                gLogMsg "Error: " & gMsg & Err.Description & "; Error # " & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
'                SQLQuery = "delete FROM cef_comments_events WHERE cefCode = " & llNewCefCode
'                gSQLWaitNoMsgBox SQLQuery, False
'            End If  'failed to update Ust
'        End If  'inserted to CEF
'    End If  'valid address
'    Set olMail = Nothing
'
'End Sub
'Public Function gGetEMailAddress(ByRef blAddFromAddress As Boolean) As String
'' search ust for sgClientName; get ustmailcefcode.  Look at Cef and see if address;
'Dim rstUst As ADODB.Recordset
'Dim rstCef As ADODB.Recordset
'    SQLQuery = "SELECT ustemailcefcode FROM ust Where ustname = '" & sgUserName & "'"
'    Set rstUst = gSQLSelectCall(SQLQuery)
'    blAddFromAddress = True
'    If Not rstUst.EOF Then  'just in case no ust record
'        If rstUst!ustEMailCefCode > 0 Then
'            SQLQuery = "SELECT cefComment FROM cef_comments_events where cefCode = " & rstUst!ustEMailCefCode
'            Set rstCef = gSQLSelectCall(SQLQuery)
'            If Not rstCef.EOF Then
'                gGetEMailAddress = Trim(rstCef!cefComment)
'                If LenB(gGetEMailAddress) > 0 Then
'                    blAddFromAddress = False
'                End If
'            End If
'        End If
'    End If
'    Set rstUst = Nothing
'    Set rstCef = Nothing
'End Function

Public Sub gSetMousePointer(grdCtrl1 As MSHFlexGrid, grdCtrl2 As MSHFlexGrid, ilPointer As Integer)
    Screen.MousePointer = ilPointer
    grdCtrl1.MousePointer = ilPointer
    grdCtrl2.MousePointer = ilPointer
End Sub


Public Function gDidAnySpotsAir(lAttCode As Long, sMoDate As String, sSuDate As String) As Boolean

    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    'Test to see if any spots aired or were they all not aired
    SQLQuery = "Select astCode, astStatus FROM ast WHERE"
    SQLQuery = SQLQuery + " astAtfCode = " & lAttCode
    SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(sMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sSuDate, sgSQLDateForm) & "')"
    SQLQuery = SQLQuery + " AND (astCPStatus = 1)"
    'SQLQuery = SQLQuery + " AND (Mod(astStatus, 100) <= 1 Or Mod(astStatus, 100) >= 9))"
    'Set rst = gSQLSelectCall(SQLQuery)
    'If Not rst.EOF Then
    '    'We know at least one spot aired
    '   gDidAnySpotsAir = True
    'Else
    '    'no aired spots were found
    '    gDidAnySpotsAir = False
    'End If
    'SQLQuery = SQLQuery + ")"
    'Set rst = cnn.Execute(SQLQuery)
    Set rst = gSQLSelectCall(SQLQuery)
    gDidAnySpotsAir = False
    Do While Not rst.EOF
        If tgStatusTypes(gGetAirStatus(rst!astStatus)).iPledged <> 2 Then
            gDidAnySpotsAir = True
            Exit Do
        End If
        rst.MoveNext
    Loop
    
    rst.Close
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modGenSubs-gDidAnySpotsAir"
End Function

Public Function gGetStationPW(iShttCode As Integer) As String
    
    Dim rst As ADODB.Recordset
    
    SQLQuery = "SELECT shttWebPW"
    SQLQuery = SQLQuery & " FROM shtt"
    SQLQuery = SQLQuery + " WHERE (shttCode = " & iShttCode & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        gGetStationPW = Trim$(rst!shttWebPW)
    Else
        gGetStationPW = ""
    End If
    
    rst.Close
    
End Function

Public Sub gCenterStdAlone(frm As Form)
    frm.Move (Screen.Width - frm.Width) \ 2, (Screen.Height - frm.Height) \ 2 + 115 '+ Screen.Height \ 10
End Sub

Public Function gCreateDayStr(slDayName As String) As String
    Dim slStr As String
    Dim slDays As String
    Dim ilPos As Integer
    Dim Days As Integer
    Dim slStart As String
    Dim slEnd As String
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim ilSet As Integer
    Dim ilDays As Integer
    
    slStr = Trim$(slDayName)
    slDays = String(7, "N")
    gParseUnknownNumberCDFields slStr, False, smDays()
    For ilDays = LBound(smDays) To UBound(smDays) Step 1
        ilPos = InStr(1, smDays(ilDays), "-", vbTextCompare)
        If ilPos <= 0 Then
            slStart = smDays(ilDays)
            slEnd = slStart
        Else
            slStart = Left$(smDays(ilDays), ilPos - 1)
            slEnd = Mid$(smDays(ilDays), ilPos + 1)
        End If
        slStr = UCase(slStart)
        'Could use switch to get the index
        'ilStart = Switch(slStr = "M", 1, slStr = "MO", 1, slStr = "TU", 2, slStr = "W", 3, slStr = "WE", 3,...)
        If slStr = "M" Or slStr = "MO" Then
            ilStart = 1
        ElseIf (slStr = "TU") Then
            ilStart = 2
        ElseIf slStr = "W" Or slStr = "WE" Then
            ilStart = 3
        ElseIf (slStr = "TH") Then
            ilStart = 4
        ElseIf slStr = "F" Or slStr = "FR" Then
            ilStart = 5
        ElseIf (slStr = "SA") Then
            ilStart = 6
        ElseIf (slStr = "SU") Then
            ilStart = 7
        End If
        slStr = UCase(slEnd)
        If slStr = "M" Or slStr = "MO" Then
            ilEnd = 1
        ElseIf (slStr = "TU") Then
            ilEnd = 2
        ElseIf slStr = "W" Or slStr = "WE" Then
            ilEnd = 3
        ElseIf (slStr = "TH") Then
            ilEnd = 4
        ElseIf slStr = "F" Or slStr = "FR" Then
            ilEnd = 5
        ElseIf (slStr = "SA") Then
            ilEnd = 6
        ElseIf (slStr = "SU") Then
            ilEnd = 7
        End If
        If (ilStart < 1) Or (ilStart > 7) Or (ilEnd < 1) Or (ilEnd > 7) Or (ilEnd < ilStart) Then
            slDays = "" 'String(7, "N")
            Exit For
        Else
            For ilSet = ilStart To ilEnd Step 1
                Mid$(slDays, ilSet, 1) = "Y"
            Next ilSet
        End If
    Next ilDays
    gCreateDayStr = slDays
End Function

Public Function gDayConvert(slInDays As String) As String
    Dim slDays As String
    Dim ilDay As Integer
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim slStr As String
    
    slDays = Trim$(slInDays)
    If (InStr(1, slDays, "Y", vbTextCompare) > 0) Or (InStr(1, slDays, "N", vbTextCompare) > 0) Then
        slStr = ""
        ilDay = 1
        Do
            If Mid(slDays, ilDay, 1) = "Y" Then
                ilStart = ilDay
                ilEnd = ilStart
                ilDay = ilDay + 1
                Do
                    If ilDay > 7 Then
                        Exit Do
                    End If
                    If Mid(slDays, ilDay, 1) = "N" Then
                        Exit Do
                    Else
                        ilEnd = ilDay
                    End If
                    ilDay = ilDay + 1
                Loop
                If slStr = "" Then
                    If ilStart = ilEnd Then
                        slStr = Switch(ilStart = 1, "M", ilStart = 2, "Tu", ilStart = 3, "W", ilStart = 4, "Th", ilStart = 5, "F", ilStart = 6, "Sa", ilStart = 7, "Su")
                    Else
                        slStr = Switch(ilStart = 1, "M", ilStart = 2, "Tu", ilStart = 3, "W", ilStart = 4, "Th", ilStart = 5, "F", ilStart = 6, "Sa", ilStart = 7, "Su")
                        slStr = slStr & "-" & Switch(ilEnd = 1, "M", ilEnd = 2, "Tu", ilEnd = 3, "W", ilEnd = 4, "Th", ilEnd = 5, "F", ilEnd = 6, "Sa", ilEnd = 7, "Su")
                    End If
                Else
                    If ilStart = ilEnd Then
                        slStr = slStr & "," & Switch(ilStart = 1, "M", ilStart = 2, "Tu", ilStart = 3, "W", ilStart = 4, "Th", ilStart = 5, "F", ilStart = 6, "Sa", ilStart = 7, "Su")
                    Else
                        slStr = slStr & "," & Switch(ilStart = 1, "M", ilStart = 2, "Tu", ilStart = 3, "W", ilStart = 4, "Th", ilStart = 5, "F", ilStart = 6, "Sa", ilStart = 7, "Su")
                        slStr = slStr & "-" & Switch(ilEnd = 1, "M", ilEnd = 2, "Tu", ilEnd = 3, "W", ilEnd = 4, "Th", ilEnd = 5, "F", ilEnd = 6, "Sa", ilEnd = 7, "Su")
                    End If
                End If
            End If
            ilDay = ilDay + 1
        Loop While ilDay <= 7
        slDays = slStr
    Else
        If (slDays = "MoTuWeThFrSaSu") Then
            slDays = "M-Su"
        ElseIf (slDays = "MoTuWeThFrSa") Then
            slDays = "M-Sa"
        ElseIf (slDays = "MoTuWeThFr") Then
            slDays = "M-F"
        ElseIf (slDays = "MoTuWeTh") Then
            slDays = "M-Th"
        ElseIf (slDays = "MoTuWe") Then
            slDays = "M-W"
        ElseIf slDays = ("MoTu") Then
            slDays = "M-Tu"
        ElseIf (slDays = "TuWeThFrSaSu") Then
            slDays = "Tu-Su"
        ElseIf (slDays = "TuWeThFrSa") Then
            slDays = "Tu-Sa"
        ElseIf (slDays = "TuWeThFr") Then
            slDays = "Tu-F"
        ElseIf (slDays = "TuWeTh") Then
            slDays = "Tu-Th"
        ElseIf (slDays = "TuWe") Then
            slDays = "Tu-W"
        ElseIf (slDays = "WeThFrSaSu") Then
            slDays = "W-Su"
        ElseIf (slDays = "WeThFrSa") Then
            slDays = "W-Sa"
        ElseIf (slDays = "WeThFr") Then
            slDays = "W-F"
        ElseIf (slDays = "WeTh") Then
            slDays = "W-Th"
        ElseIf (slDays = "ThFrSaSu") Then
            slDays = "Th-Su"
        ElseIf (slDays = "ThFrSa") Then
            slDays = "Th-Sa"
        ElseIf (slDays = "ThFr") Then
            slDays = "Th-F"
        ElseIf slDays = "FrSaSu" Then
            slDays = "F-Su"
        ElseIf slDays = "FrSa" Then
            slDays = "F-Sa"
        ElseIf slDays = "SaSu" Then
            slDays = "S-S"
        End If
    End If
    gDayConvert = slDays
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gParseCDFields                  *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Parse comma delimited fields    *
'*                     Note:including quotes that are  *
'*                     enclosed within quotes          *
'*                     ""xxxxxxxx"","xxxxx",           *
'*                                                     *
'*******************************************************
Sub gParseUnknownNumberCDFields(slCDStr As String, ilLower As Integer, slFields() As String)
'
'   gParseCDFields slCDStr, ilLower, slFields()
'   Where:
'       slCDStr(I)- Comma delinited string
'       ilLower(I)- True=Convert string fields characters to lower case (preceding character is A-Z)
'       slFields() (O)- fields parsed from comma delimited string
'
    Dim ilFieldNo As Integer
    Dim ilFieldType As Integer  '0=String, 1=Number
    Dim slChar As String
    Dim ilIndex As Integer
    Dim ilAscChar As Integer
    Dim ilAddToStr As Integer
    Dim slNextChar As String

'    For ilIndex = LBound(slFields) To UBound(slFields) Step 1
'        slFields(ilIndex) = ""
'    Next ilIndex
    'ReDim slFields(1 To 1) As String
    ReDim slFields(0 To 0) As String
    slFields(UBound(slFields)) = ""
    'ilFieldNo = 1
    ilFieldNo = 0
    ilIndex = 1
    ilFieldType = -1
    Do While ilIndex <= Len(Trim$(slCDStr))
        slChar = Mid$(slCDStr, ilIndex, 1)
        If ilFieldType = -1 Then
            If slChar = "," Then    'Comma was followed by a comma-blank field
                ilFieldType = -1
                ilFieldNo = ilFieldNo + 1
                If ilFieldNo > UBound(slFields) Then
                    ReDim Preserve slFields(LBound(slFields) To UBound(slFields) + 1) As String
                    slFields(UBound(slFields)) = ""
                End If
            ElseIf slChar <> """" Then
                ilFieldType = 1
                slFields(ilFieldNo) = slChar
            Else
                ilFieldType = 0 'Quote field
            End If
        Else
            If ilFieldType = 0 Then 'Started with a Quote
                'Add to string unless "
                ilAddToStr = True
                If slChar = """" Then
                    If ilIndex = Len(Trim$(slCDStr)) Then
                        ilAddToStr = False
                    Else
                        slNextChar = Mid$(slCDStr, ilIndex + 1, 1)
                        If slNextChar = "," Then
                            ilAddToStr = False
                        End If
                    End If
                End If
                If ilAddToStr Then
                    If (slFields(ilFieldNo) <> "") And ilLower Then
                        ilAscChar = Asc(UCase(right$(slFields(ilFieldNo), 1)))
                        If ((ilAscChar >= Asc("A")) And (ilAscChar <= Asc("Z"))) Then
                            slChar = LCase$(slChar)
                        End If
                    End If
                    slFields(ilFieldNo) = slFields(ilFieldNo) & slChar
                Else
                    ilFieldType = -1
                    ilFieldNo = ilFieldNo + 1
                    If ilFieldNo > UBound(slFields) Then
                        ReDim Preserve slFields(LBound(slFields) To UBound(slFields) + 1) As String
                        slFields(UBound(slFields)) = ""
                    End If
                    ilIndex = ilIndex + 1   'bypass comma
                End If
            Else
                'Add to string unless ,
                If slChar <> "," Then
                    slFields(ilFieldNo) = slFields(ilFieldNo) & slChar
                Else
                    ilFieldType = -1
                    ilFieldNo = ilFieldNo + 1
                    If ilFieldNo > UBound(slFields) Then
                        ReDim Preserve slFields(LBound(slFields) To UBound(slFields) + 1) As String
                        slFields(UBound(slFields)) = ""
                    End If
                End If
            End If
        End If
        ilIndex = ilIndex + 1
    Loop
End Sub


Public Function gGetTitleByTntCode(ilCode As Integer) As String
    
    Dim rst_tnt As ADODB.Recordset

    gGetTitleByTntCode = ""
    SQLQuery = "Select tntCode, tntTitle From Tnt where tntCode = " & ilCode
    Set rst_tnt = gSQLSelectCall(SQLQuery)
    If Not rst_tnt.EOF Then
        gGetTitleByTntCode = Trim$(rst_tnt!tntTitle)
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modGenSubs-gGetTitleByTntCode"
End Function

Public Function gGetTntCodeByTitle(sTitle As String) As Integer
    
    'D.S. 02/23/11
    
    Dim rst_tnt As ADODB.Recordset

    gGetTntCodeByTitle = 0
    SQLQuery = "Select tntCode, tntTitle From Tnt where UPPER(tntTitle) = " & "'" & UCase(gFixQuote(sTitle)) & "'"
    Set rst_tnt = gSQLSelectCall(SQLQuery)
    If Not rst_tnt.EOF Then
        gGetTntCodeByTitle = Trim$(rst_tnt!tntCode)
    Else
        'Title was not found so add into the TNT table
        gGetTntCodeByTitle = gAddTitleName(Trim$(sTitle))
    End If
    
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modGenSubs-gGetTntCodeByTitle"
End Function


Public Function gAddTitleName(slTitle As String) As Long
    Dim ilCode As Integer
    Dim slSQLQuery As String
    
    On Error GoTo ErrHandler:
    gAddTitleName = -1
    slSQLQuery = "Insert Into tnt ( "
    slSQLQuery = slSQLQuery & "tntCode, "
    slSQLQuery = slSQLQuery & "tntTitle, "
    slSQLQuery = slSQLQuery & "tntUsfCode, "
    slSQLQuery = slSQLQuery & "tntUnused "
    slSQLQuery = slSQLQuery & ") "
    slSQLQuery = slSQLQuery & "Values ( "
    slSQLQuery = slSQLQuery & "Replace" & ", "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(slTitle) & "', "
    slSQLQuery = slSQLQuery & igUstCode & ", "
    slSQLQuery = slSQLQuery & "'" & "" & "' "
    slSQLQuery = slSQLQuery & ") "
    ilCode = CInt(gInsertAndReturnCode(slSQLQuery, "tnt", "tntCode", "Replace"))
    If ilCode > 0 Then
        gAddTitleName = ilCode
    End If
    
    Exit Function

ErrHandler:
    gHandleError "AffErrorLog.txt", "modGenSubs-gAddTitleName"
    Exit Function
End Function

Public Function gGetTimeZoneOffset(iShttCode As Integer, lVefCode As Long) As Integer
    
    'D.S. 03/22/11
    'Copied old that keeps getting written over and over and made a function out of it.
    'Get the time offset between the vehicle and the local station's time zones
    
    'Returns the offset time or -1 if an error occurred
    
    Dim slZone As String
    Dim ilVefArrayInx As Integer
    Dim ilLocalAdj As Integer
    Dim ilZoneFound As Integer
    Dim ilNumberAsterisk As Integer
    Dim ilZone As Integer
    Dim ilRet As Integer
    Dim ilShttRet As Integer
    
    On Error GoTo ErrHandler
    
    'init error condition
    gGetTimeZoneOffset = -1
    
    '2/21/15: Get index
    'slZone = UCase$(Trim$(tgShttInfo1(iShttCode - 1).shttTimeZone))
    ilShttRet = gBinarySearchShtt(iShttCode)
    If ilShttRet = -1 Then
        Exit Function
    End If
    slZone = UCase$(Trim$(tgShttInfo1(ilShttRet).shttTimeZone))
    ilVefArrayInx = gBinarySearchVef(lVefCode)
    ilLocalAdj = 0
    ilZoneFound = False
    ilNumberAsterisk = 0
    ' Adjust time zone properly.
    If Len(slZone) <> 0 Then
        'Get zone
        For ilZone = LBound(tgVehicleInfo(ilVefArrayInx).sZone) To UBound(tgVehicleInfo(ilVefArrayInx).sZone) Step 1
            If Trim$(tgVehicleInfo(ilVefArrayInx).sZone(ilZone)) = slZone Then
                '2/21/15: Change if base zone defined
                'If tgVehicleInfo(ilVefArrayInx).sFed(ilZone) <> "*" Then
                If (tgVehicleInfo(ilVefArrayInx).sFed(ilZone) <> "*") And (tgVehicleInfo(ilVefArrayInx).iBaseZone(ilZone) <> -1) Then
                    slZone = tgVehicleInfo(ilVefArrayInx).sZone(tgVehicleInfo(ilVefArrayInx).iBaseZone(ilZone))
                    ilLocalAdj = tgVehicleInfo(ilVefArrayInx).iLocalAdj(ilZone)
                    ilZoneFound = True
                End If
                Exit For
            End If
        Next ilZone
        For ilZone = LBound(tgVehicleInfo(ilVefArrayInx).sZone) To UBound(tgVehicleInfo(ilVefArrayInx).sZone) Step 1
            If tgVehicleInfo(ilVefArrayInx).sFed(ilZone) = "*" Then
                ilNumberAsterisk = ilNumberAsterisk + 1
            End If
        Next ilZone
    End If
    If (Not ilZoneFound) And (ilNumberAsterisk <= 1) Then
        slZone = ""
    End If
    ilLocalAdj = -1 * ilLocalAdj
    
    gGetTimeZoneOffset = ilLocalAdj
    
    Exit Function
    
ErrHandler:
    gHandleError "AffErrorLog.txt", "modGenSubs-gGetTimeZoneOffset"
    gGetTimeZoneOffset = -1
    Exit Function
End Function


Public Function gGetAgreementDateRange(sATTCode As String) As String

    'D.S. 03/22/11
    'Get Start and End dates for the given agreement
    
    Dim slEndDate As String
    Dim slRange As String
    
    gGetAgreementDateRange = ""
    slRange = ""
    
    SQLQuery = "SELECT attOnAir, attOffAir, attDropDate"
    SQLQuery = SQLQuery & " FROM att"
    SQLQuery = SQLQuery + " WHERE (attCode = " & sATTCode & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    If rst.EOF = False Then
        If DateValue(gAdjYear(rst!attDropDate)) < DateValue(gAdjYear(rst!attOffAir)) Then
            slEndDate = Format$(rst!attDropDate, sgShowDateForm)
        Else
            slEndDate = Format$(rst!attOffAir, sgShowDateForm)
        End If
        If (DateValue(gAdjYear(rst!attOnAir)) = DateValue("1/1/1970")) Then
            slRange = ""
        Else
            slRange = Format$(Trim$(rst!attOnAir), sgShowDateForm)
        End If
        If (DateValue(gAdjYear(slEndDate)) = DateValue("12/31/2069") Or DateValue(gAdjYear(slEndDate)) = DateValue("12/31/69")) Then  'Or (rst!attOffAir = "12/31/69") Then
            If slRange <> "" Then
                slRange = slRange & "-TFN"
            End If
        Else
            If slRange <> "" Then
                slRange = slRange & "-" & slEndDate    'rst!attOffAir
            Else
                slRange = "Thru " & slEndDate 'rst!attOffAir
            End If
        End If
    End If
    
    gGetAgreementDateRange = slRange
    
    Exit Function

ErrHandler:
    gHandleError "AffErrorLog.txt", "modGenSubs-gGetAgreementDateRange"
    gGetAgreementDateRange = slRange
    Exit Function
End Function

'****************************************************************************
'
'****************************************************************************
Public Sub gSpellCheckUsingMSWord(edcCtrl As TextBox)
    On Error GoTo Err_SpellCheckUsingMSWord
    Dim slText As String

    Screen.MousePointer = vbHourglass
    'RaiseEvent SpellCheckerStarting
    slText = edcCtrl.Text

    App.OleRequestPendingTimeout = 999999   ' Prevent the "Switch To" dialog from appearing.
    'App.OleServerBusyMsgText = "Press Alt-Esc to see the spell checking results"
    'App.OleRequestPendingMsgText = "Press Alt-Esc to see the spell checking results"
'    DoEvents
'    App.OleServerBusyTimeout = 1000
'    App.OleServerBusyRaiseError = True
    Set SpellCheck = CreateObject("Word.Application")
    SpellCheck.Visible = False
    Call mMinimizeWordIfOpen
    SpellCheck.Documents.Add                              'Open New Document (Hidden)
    Clipboard.Clear
    Clipboard.SetText slText, vbCFText                    'Copy Text To Clipboard
    SpellCheck.Selection.Paste                            'Paste Text Into WORD
    Call mBringWindowToTopMost
    SpellCheck.Visible = False
    SpellCheck.ActiveDocument.CheckSpelling               'Activate The Spell Checker
    'SpellCheck.ActiveDocument.CheckGrammar                ' Does both spelling and grammer.
    SpellCheck.Visible = False                            'Hide WORD From User
    SpellCheck.ActiveDocument.Select                      'Select The Corrected Text
    SpellCheck.Selection.Cut                              'Cut The Text To Clipboard
    edcCtrl.Text = Clipboard.GetText(vbCFText)  'Assign Text To SPELLCHECKER Function
    SpellCheck.ActiveDocument.Close False
    SpellCheck.Quit
    Set SpellCheck = Nothing
'    'RaiseEvent SpellCheckerCompleted
'    Screen.MousePointer = vbNormal
'    MsgBox "Spell Checking is Complete"
'    rtfRichTextBox1.SetFocus
    edcCtrl.SetFocus
    edcCtrl.SelStart = Len(edcCtrl.Text)
    Screen.MousePointer = vbDefault
    Exit Sub

Err_SpellCheckUsingMSWord:
    Screen.MousePointer = vbDefault
    SpellCheck.ActiveDocument.Close False
    SpellCheck.Quit
    Set SpellCheck = Nothing
    MsgBox "Error: " & Err.Number & ", " & Err.Description & vbCrLf & vbCrLf & "Please note you must have Microsoft Word installed to utilize the spell check feature.", vbExclamation, "Spell Check Problem"
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub mBringWindowToTopMost()
    Dim hwnd As Long
    Dim ilResult As Long

    'hWnd = FindWindow(vbNullString, "Spelling: English (U.S.)")
    hwnd = FindWindow(vbNullString, "Document1 - Microsoft Word")

    If hwnd <> 0 Then
        ilResult = SetWindowPos(hwnd, WNDNOTOPMOST, 0, 0, 0, 0, FRMNOMOVE Or FRMNOSIZE)
    End If
End Sub

'****************************************************************************
' This function will look for a word doc that is currently open with the
' title of "Document1 - Microsoft Word", indicating a new blank word doc.
' If this is found, we need to minimize it to avoid having it become the
' top most visible window.
'
'****************************************************************************
Private Sub mMinimizeWordIfOpen()
    Dim hwnd As Long
    Dim wp As WINDOWPLACEMENT

    ' Const WM_COMMAND = &H111
    hwnd = FindWindow(vbNullString, "Document1 - Microsoft Word")

    If hwnd <> 0 Then
        If GetWindowPlacement(hwnd, wp) > 0 Then
            wp.LENGTH = Len(wp)
            wp.Flags = 0&
            wp.showCmd = SW_SHOWMINIMIZED
            SetWindowPlacement hwnd, wp
        End If
        ' SendMessage hWnd, &H111, MIN_ALL, ByVal 0&
    End If
End Sub

Public Sub gClearListScrollBar(myLbc As ListBox)
    Dim llRet As Long
    'clear the horz. scroll bar if its there
    llRet = SendMessageByNum(myLbc.hwnd, LB_SETHORIZONTALEXTENT, 0, 0)
End Sub



Public Function gIsAstStatus(ilAstStatus As Integer, ilIsStatus As Integer) As Boolean
    'ilAstStatus (I): astStatus to be checked for a give status
    'ilIsStatus (I): Status to be checked
    '
    'Example call
    'If gIsAstStatus(tmAst.iStatus, ASTEXTENDEDMG) Then
    '     'Spot is a MG
    'Else
    '     'Spot is not a MG
    'End If
    '
    Dim ilRemainder As Integer
    Dim llScaler As Long
    
    If ilIsStatus = ilAstStatus Then
        gIsAstStatus = True
        Exit Function
    End If
    If ilIsStatus > ilAstStatus Then
        gIsAstStatus = False
        Exit Function
    End If
    If ilIsStatus < 10 Then
        llScaler = 10
    ElseIf ilIsStatus < 100 Then
        llScaler = 100
    ElseIf ilIsStatus < 1000 Then
        llScaler = 1000
    ElseIf ilIsStatus < 10000 Then
        llScaler = 10000
    Else
        llScaler = 100000
    End If
    ilRemainder = ilAstStatus Mod llScaler
    If llScaler <= 100 Then
        If ilRemainder = ilIsStatus Then
            gIsAstStatus = True
        Else
            gIsAstStatus = False
        End If
    Else
        If ilRemainder - (ilRemainder Mod ilIsStatus) = ilIsStatus Then
            gIsAstStatus = True
        Else
            gIsAstStatus = False
        End If
    End If
End Function

Public Function gGetAirStatus(ilAstStatus) As Integer
    gGetAirStatus = ilAstStatus Mod 100
End Function
'Dan changed name 7/22/15
'Public Function gGetISCIChgdStataus(ilAstStatus) As Boolean
Public Function gIsISCIChanged(ilAstStatus) As Boolean
    gIsISCIChanged = False
    If (ilAstStatus \ ASTEXTENDED_ISCICHGD) = 1 Then
        gIsISCIChanged = True
    End If
End Function
Public Function gIsMissedReasonDefined(ilAstStatus) As Boolean
    gIsMissedReasonDefined = False
    If (ilAstStatus Mod ASTEXTENDED_ISCICHGD) \ ASTEXTENDED_MISSREASON = 1 Then
        gIsMissedReasonDefined = True
    End If
End Function
' look ahead for combo boxes 9/15/11 Dan M taken from traffic for fCrViewerExport

'*******************************************************
'*                                                     *
'*      Procedure Name:gManLookAhead                   *
'*                                                     *
'*             Created:5/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Validates combo box input and,  *
'*                     if invalid, resets the input to *
'*                     the specified last valid input. *
'*                     It validates input by testing   *
'*                     to see if the combo box contains*
'*                     any matching string to the input*
'*                     string.  This is used for       *
'*                     ComboBox controls.               *
'*                                                     *
'*******************************************************
Public Sub gManLookAhead(cbcComboBox As ComboBox, ilBSMode As Integer, ilErrHighLightIndex As Integer)
'
'   gManLookAhead cbcCtrl, ilBSMode, ilHighlightIndex
'   Where:
'       cbcCtrl (I)- Combo box Control containing values
'       ilBSMode (I/O)- Backspace flag (True = backspace key was pressed, False =                       '       backspace key was not pressed)
'       ilHighlightIndex (I)- Selection to be highlighted if input is invalid
'

    Dim ilLen As Integer    'Length of current enter text
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    Dim ilBracket As Integer
    Dim ilSearch As Integer
    Dim ilSvLastFound As Integer
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim ilSelStart As Integer

    ilSelStart = cbcComboBox.SelStart
    slStr = LTrim$(cbcComboBox.Text)    'Remove leading blanks only
    ilLen = Len(cbcComboBox.Text)
    ilIndex = cbcComboBox.ListIndex
    If slStr = "" Then  'If space bar selected, text will be blank- ListIndex will contain a value
        cbcComboBox.ListIndex = -1
        Exit Sub
        If cbcComboBox.ListIndex >= 0 Then
            slStr = cbcComboBox.List(cbcComboBox.ListIndex)
            ilLen = 0
            ilIndex = -1    'Force dispaly of selected item by space bar
        Else
            Beep
            If ilErrHighLightIndex >= 0 Then
                cbcComboBox.ListIndex = ilErrHighLightIndex
            End If
            Exit Sub
        End If
    End If
    If ilBSMode Then    'If backspace mode- reduce string by one character
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        ilBSMode = False
        If ilSelStart > 0 Then
            ilSelStart = ilSelStart - 1
        End If
    End If
    If Left$(slStr, 1) = "[" Then   'Search does not work when starting with [
        ilSvLastFound = -1
        For ilLoop = 0 To cbcComboBox.ListCount - 1 Step 1
            ilPos = InStr(1, cbcComboBox.List(ilLoop), slStr, 1)
            If ilPos = 1 Then
                ilSvLastFound = ilLoop
                Exit For
            Else
                If Left$(cbcComboBox.List(ilLoop), 1) <> "[" Then
                    Exit For
                End If
            End If
        Next ilLoop
    Else
        'Test if matching string found in the combo box- if so display it (look ahead typing)
    '    cbcComboBox.ListIndex = 0
        gFndFirst cbcComboBox, slStr
        ilBracket = False
        Do
            If gLastFound(cbcComboBox) >= 0 Then
                If (Left$(cbcComboBox.List(gLastFound(cbcComboBox)), 1) = "[") And (Left$(slStr, 1) <> "[") Then
                    gFndNext cbcComboBox, slStr
                    ilBracket = True
                Else
                    ilBracket = False
                End If
            Else
                ilBracket = False
            End If
        Loop While ilBracket
        'Test if another name matches encase names are not in sorted order
        ilSearch = True
        ilSvLastFound = gLastFound(cbcComboBox)
        Do
            If gLastFound(cbcComboBox) >= 0 Then
                If StrComp(slStr, cbcComboBox.List(gLastFound(cbcComboBox)), 1) = 0 Then
                    ilSvLastFound = gLastFound(cbcComboBox)
                    Exit Do
                End If
                gFndNext cbcComboBox, slStr
            Else
                Exit Do
            End If
        Loop While ilSearch
    End If
    If ilSvLastFound >= 0 Then
        If (ilIndex <> ilSvLastFound) Or ((ilIndex = ilSvLastFound) And Not ilBSMode) Then
            cbcComboBox.ListIndex = ilSvLastFound 'This will cause a change event (reason for imChgMode)
        End If
        ''If item found not same as current selected- change current
        'If (ilIndex <> ilSvLastFound) And (ilIndex >= 0) Then
        '    'Test if same name- and slStr contain whole name- if so leave index
        '    If (cbcComboBox.List(ilIndex) <> cbcComboBox.List(ilSvLastFound)) Or (slStr <> cbcComboBox.List(ilIndex)) Then
        '        cbcComboBox.ListIndex = ilSvLastFound 'This will cause a change event (reason for imChgMode)
        '    Else
        '        cbcComboBox.ListIndex = ilIndex
        '    End If
        'ElseIf (ilIndex <> ilSvLastFound) Or ((ilIndex = ilSvLastFound) And Not ilBSMode) Then
        '    cbcComboBox.ListIndex = ilSvLastFound 'This will cause a change event (reason for imChgMode)
        'Else
        'End If
        ilErrHighLightIndex = cbcComboBox.ListIndex

'        If (ilIndex <> ilSvLastFound) Or ((ilIndex = ilSvLastFound) And Not ilBSMode) Then 'If indices not equal- highlight look ahead text
'            cbcComboBox.SelStart = ilLen
'            cbcComboBox.SelLength = Len(cbcComboBox.Text)
'        End If
    Else
        Beep
        If (ilErrHighLightIndex >= 0) And (ilErrHighLightIndex < cbcComboBox.ListCount) Then
            cbcComboBox.ListIndex = ilErrHighLightIndex
            ilSelStart = 0
        End If
    End If
    If cbcComboBox.ListIndex >= 0 Then
        cbcComboBox.Text = cbcComboBox.List(cbcComboBox.ListIndex)
    End If
    If ilSelStart <= Len(cbcComboBox.Text) Then
        cbcComboBox.SelStart = ilSelStart
        cbcComboBox.SelLength = Len(cbcComboBox.Text)
    Else
        cbcComboBox.SelStart = 0
        cbcComboBox.SelLength = Len(cbcComboBox.Text)
    End If
End Sub
Private Sub gFndFirst(lbcList As control, slInMatch As String)
    cgList = lbcList
    If TypeOf lbcList Is ComboBox Then
        igFoundRow = SendMessageByString(lbcList.hwnd, CB_FINDSTRING, -1, slInMatch)
    Else
        igFoundRow = SendMessageByString(lbcList.hwnd, LB_FINDSTRING, -1, slInMatch)
    End If
End Sub
Private Sub gFndNext(lbcList As control, slInMatch As String)
    Dim slNext As String
    Dim ilTestRow As Integer

    If cgList <> lbcList Then
        igFoundRow = -1
        Exit Sub
    End If
    ilTestRow = igFoundRow + 1
    Do While ilTestRow < lbcList.ListCount
        slNext = lbcList.List(ilTestRow)
        If InStr(1, slNext, slInMatch, vbTextCompare) = 1 Then
            igFoundRow = ilTestRow
            Exit Sub
        End If
        ilTestRow = ilTestRow + 1
    Loop
    igFoundRow = -1

End Sub
Private Function gLastFound(lbcList As control) As Integer
    If cgList <> lbcList Then
        gLastFound = -1
    Else
        gLastFound = igFoundRow
    End If
End Function

Public Function gGetTrueRecCount(sPervFile As String) As Long

    'D.S. 11/10/11
    'Returns a long with the true count of records rather than the max increment code
    'Return value of -1 is an error condition
    'Input a file name:  ast.mkd, sdf.btr etc.
    'Example call:  llRet = mGetTrueRecCount(ast.mkd)
    'Dependents:  TrueRecCount.bat in the exe folder
    'If TrueRecCount.bat not found, it will create it in the exe folder
    
    Dim ilRet As Integer
    Dim lProcessId As Long
    Dim hProcess As Long
    Dim lExitCode As Long
    Dim lRet As Long
    Dim fStart As Single
    Dim slCommand As String
    Dim llCount As Long
    Dim hlFrom As Integer   'String
    Dim slLocation As String
    Dim slReadLine As String
    Dim ilLen As String
    Dim ilPos As Integer
    Dim hmDetail As Integer
        
    gGetTrueRecCount = -1
    
    'On Error GoTo FileErrHand
    
    'Create TrueRecCount.bat if it's not present
    ilRet = 0
    'slReadLine = FileDateTime(sgExeDirectory & "TrueRecCount.bat")
    ilRet = gFileExist(sgExeDirectory & "TrueRecCount.bat")
    If ilRet <> 0 Then
        ilRet = 0
        slReadLine = sgExeDirectory & "TrueRecCount.bat"
        'hmDetail = FreeFile
        'Open slReadLine For Output Lock Write As hmDetail
        ilRet = gFileOpen(slReadLine, "Output Lock Write", hmDetail)
        If ilRet <> 0 Then
            gMsgBox "Open File " & slReadLine & " error#" & Str$(Err.Number), vbOKOnly
            Exit Function
        End If
        Print #hmDetail, "butil -stat %1 > %2"
        Close hmDetail
    End If
    
    On Error GoTo ErrHand
    'Execute the batch file to create the Pervasive results file
    slCommand = sgExeDirectory & "TrueRecCount.bat" & " " & sgDBPath & sPervFile & " " & sgDBPath & "TrueRecCount.txt"
    lProcessId = Shell(slCommand, vbHide)
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, lProcessId)
    Do
        lRet = GetExitCodeProcess(hProcess, lExitCode)
        DoEvents
    Loop While (lExitCode = STILL_ACTIVE)
    
    lRet = CloseHandle(hProcess)
    Sleep 500
    
    'On Error GoTo FileErrHand
    
    'hlFrom = FreeFile
    ilRet = 0
    slLocation = sgDBPath & "TrueRecCount.txt"
    'Open slLocation For Input Access Read As hlFrom
    ilRet = gFileOpen(slLocation, "Input Access Read", hlFrom)
    If ilRet <> 0 Then
        Close hlFrom
        gMsgBox "Open File " & slLocation & " error#" & Str$(Err.Number), vbOKOnly
        Exit Function
    End If

    Do While Not EOF(hlFrom)
        ilRet = 0
        Input #hlFrom, slReadLine
        If ilRet <> 0 Then
            gLogMsg "Error: frmDataDoubler-mGetTrueRecCount was unable read/input statement.", "WebExportLog.Txt", False
            Exit Function
        End If
        
        ilLen = Len(slReadLine)
        ilPos = InStr(slReadLine, "Total Number of Records =")
        
        If ilPos > 0 Then
            'Found the total records line
            llCount = CLng(Mid$(slReadLine, 27, ilLen))
            Exit Do
        End If
    Loop
    
    Close hlFrom

    On Error GoTo ErrHand
    gGetTrueRecCount = llCount
    Exit Function
    
'FileErrHand:
'    ilRet = -1
'    Resume Next

ErrHand:
    gMsgBox "Shell Error " & Str$(Err.Number) & Err.Description, vbOKOnly
    gGetTrueRecCount = -1
    On Error GoTo 0
    Exit Function
End Function


Public Function gIsPledgeByAvails(llAttCode As Long) As Boolean
    Dim blPledgeByAvails As Boolean
    On Error GoTo gIsPledgeByAvailsErr:
    blPledgeByAvails = True
    SQLQuery = "SELECT * "
    SQLQuery = SQLQuery + " FROM dat"
    SQLQuery = SQLQuery + " WHERE (datAtfCode= " & llAttCode & ")"
    Set dat_rst = gSQLSelectCall(SQLQuery)
    Do While Not dat_rst.EOF
        DoEvents
        If gTimeToLong(Format$(dat_rst!datFdEdTime, "h:m:ssam/pm"), True) - gTimeToLong(Format$(dat_rst!datFdStTime, "h:m:ssam/pm"), False) > AVAIL_OR_DP_TIME Then
            blPledgeByAvails = False
            Exit Do
        End If
        dat_rst.MoveNext
    Loop
    gIsPledgeByAvails = blPledgeByAvails
    Exit Function
gIsPledgeByAvailsErr:
    gIsPledgeByAvails = False
    Exit Function
End Function

'
Public Sub gAddMsgToListBox(frm As Form, llCurrentMaxWidth As Long, slMsg As String, lbcMe As ListBox)
    ' add vertical scroll bar as needed.  llCurrentMaxWidth is previous max, so don't shrink because line smaller than a previous line.
    'Out- llCurrentMaxWidth.
    ' uses pbcRedAlert on frmDirectory
    Dim llValue As Long
    Dim llRg As Long
    Dim llRet As Long
    
    lbcMe.AddItem slMsg
    If (frm.pbcTextWidth.TextWidth(slMsg)) > llCurrentMaxWidth Then
        If (Len(slMsg) * 155) > llCurrentMaxWidth Then
            llCurrentMaxWidth = Len(slMsg) * 155
            If llCurrentMaxWidth > lbcMe.Width Then
                llValue = llCurrentMaxWidth / 15 + 120
                llRg = 0
                llRet = SendMessageByNum(lbcMe.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
            End If
        End If
    End If
End Sub

Public Function gIsTrueGuide() As Boolean
    Dim blRet As Boolean
    
    If (StrComp(sgUserName, "Guide", 1) = 0) Then
        If bgLimitedGuide Then
            blRet = True
        Else
            blRet = False
        End If
    Else
        blRet = False
    End If
    gIsTrueGuide = blRet
End Function

Public Function gIsInternalGuide() As Boolean
    Dim blRet As Boolean
    
    If (StrComp(sgUserName, "Guide", 1) = 0) Then
        If bgLimitedGuide Then
            blRet = False
        Else
            blRet = True
        End If
    Else
        blRet = False
    End If
    gIsInternalGuide = blRet
End Function

Public Function gSearchFile(slPathAndFileName As String, slInSearchStr As String, blIgnoreCase As Boolean, ilInPos As Integer, slSearchResult() As String) As Integer
    '
    '   slPathAndFileName(I)- Drive plus Path Plus file Name to search
    '   slInSearchStr(I)- String to search for match
    '   blIgnoreCase(I)- True = Ignore Case, False = Match case
    '   ilInPos(I)- -1= Any position; > 0 = Match exact position
    '   slSearchResult(O)- Lines within file that match the search string
    '
    Dim ilRet As Integer
    Dim hlFrom As Integer
    Dim slLine As String
    Dim slTest As String
    Dim slSearchStr As String
    Dim ilPos As Integer
    
    gSearchFile = False
    ReDim slSearchResult(0 To 0) As String
    slSearchStr = Trim$(slInSearchStr)
    If blIgnoreCase Then
        slSearchStr = UCase(slInSearchStr)
    End If
    ilRet = 0
    On Error GoTo gSearchFileErr:
    'hlFrom = FreeFile
    'Open slPathAndFileName For Input Access Read As hlFrom
    ilRet = gFileOpen(slPathAndFileName, "Input Access Read", hlFrom)
    If ilRet <> 0 Then
        Close hlFrom
        Exit Function
    End If

    Do While Not EOF(hlFrom)
        ilRet = 0
        On Error GoTo gSearchFileErr:
        Line Input #hlFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                slTest = slLine
                If blIgnoreCase Then
                    slTest = UCase(slLine)
                End If
                ilPos = InStr(1, slTest, slSearchStr, vbBinaryCompare)
                If ((ilPos > 0) And (ilInPos = -1)) Or ((ilPos > 0) And (ilPos = ilInPos)) Then
                    slSearchResult(UBound(slSearchResult)) = slLine
                    ReDim Preserve slSearchResult(0 To UBound(slSearchResult) + 1) As String
                End If
            End If
        End If
    Loop
    Close hlFrom
    gSearchFile = True
    Exit Function
gSearchFileErr:
    ilRet = Err.Number
    Resume Next


End Function

Public Function gUpdateLastExportDate(ilVefCode As Integer, slDate As String)

    SQLQuery = "update VFF_Vehicle_Features set vffLastAffExptDate = '" & Format(slDate, sgSQLDateForm) & "' WHERE vffVefCode = " & ilVefCode & " AND (vffLastAffExptDate < '" & Format(slDate, sgSQLDateForm) & "' OR vffLastAffExptDate Is Null)"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError "AffErrorLog.txt", "modGenSubs-gUpdateLastExportDate"
        gUpdateLastExportDate = False
        Exit Function
    End If
    gUpdateLastExportDate = True
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "gUpdateLastExportDate"
    gUpdateLastExportDate = False
    Exit Function
End Function
'5666
Public Function gRegistryGetShortDate() As String
    Dim slRet As String
    Dim slSubKey As String
    Dim slName As String
    
    slSubKey = "Control Panel\International"
    slName = "sShortDate"
    slRet = mGetString(HKEY_CURRENT_USER, slSubKey, slName)
    gRegistryGetShortDate = slRet
    
End Function
Public Function gRegistrySetShortDate(slvalue As String) As Boolean
    Dim slSubKey As String
    Dim slName As String
    
    slSubKey = "Control Panel\International"
    slName = "sShortDate"
    mSaveString HKEY_CURRENT_USER, slSubKey, slName, slvalue
End Function
'private
Private Function mGetString(llKey As Long, slPath As String, slvalue As String)
    Dim llRet As Long
    
    RegOpenKey llKey, slPath, llRet
    mGetString = mRegQueryStringValue(llRet, slvalue)
    RegCloseKey llRet
End Function
Private Sub mSaveString(llKey As Long, slPath As String, slvalue As String, slData As String)
    Dim llRet As Long
    
    RegCreateKey llKey, slPath, llRet
    RegSetValueEx llRet, slvalue, 0, REG_SZ, ByVal slData, Len(slData)
    RegCloseKey llRet
End Sub
Private Function mRegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim llRet As Long
    Dim llValueType As Long
    Dim slBuffer As String
    Dim llBufferSize As Long
    'retrieve information about the key
    llRet = RegQueryValueEx(hKey, strValueName, 0, llValueType, ByVal 0, llBufferSize)
    If llRet = 0 Then
        If llValueType = REG_SZ Then
            'Create a buffer
            slBuffer = String(llBufferSize, Chr$(0))
            'retrieve the key's content
            llRet = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal slBuffer, llBufferSize)
            If llRet = 0 Then
                'Remove the unnecessary chr$(0)'s
                mRegQueryStringValue = Left$(slBuffer, InStr(1, slBuffer, Chr$(0)) - 1)
            End If
        ElseIf llValueType = REG_BINARY Then
            Dim strData As Integer
            'retrieve the key's value
            llRet = RegQueryValueEx(hKey, strValueName, 0, 0, strData, llBufferSize)
            If llRet = 0 Then
                mRegQueryStringValue = strData
            End If
        End If
    End If
End Function
'end 5666

Public Function gGetLineParameters(blGetCompliant As Boolean, tlAstInfo As ASTINFO, slAllowedStartDate As String, slAllowedEndDate As String, slAllowedStartTime As String, slAllowedEndTime As String, ilAllowedDays() As Integer, Optional blMarketronExport As Boolean = False) As Integer  '8/1/14: Compliant not required: , ilCompliant As Integer) As Integer
    'Input Values
    '   blGetCompliant: False or True
    '   tlAstInfo: AST Reference obtained from gGetAst
    '
    'Return values
    '   slAllowedStartDate: Start Date with week that spot can air (Obtained from Flight)
    '   slAllowedEndDate: End Date with week that spot can air (Obtained from Flight)
    '   slAllowedStartTime: Start Time with week that spot can air (Obtained from Line and Daypart)
    '   slAllowedEndTime: End Time with week that spot can air (Obtained from Line and Daypart)
    '   ilAllowedDays: Allowed days that the spot can air within (Obtained from Flight)
    '   ilCompliant:  Set from Aired Date and time compared to Allowed Date/Time/Day
    '   gGetLineParameters:
    '      0=Ok
    '      1=Unable to read AST, returning Pledge date/time as Allowed Date/Time, Compliant = False
    '      2=Unable to read LST, returning Pledge date/time as Allowed Date/Time, Compliant = False
    '      3=Unable to read SDF, returning Pledge date/time as Allowed Date/Time/Day, Compliant determined from Allowed Date/Time
    '      4=Unable to read CLF, returning Pledge date/time as Allowed Date/Time/Day, Compliant determined from Allowed Date/Time
    '      5=Unable to read RDF, returning Pledge date/time as Allowed Date/Time/Day, Compliant determined from Allowed Date/Time
    '      6=Blackout, returning Pledge date/time as Allowed Date/Time/Day, Compliant = False
    '      7=Line and Booked vehicles don't match, returning Pledge date/time as Allowed Date/Time/Day, Compliant determined from Allowed Date/Time
    '      8=SQL Error Logged to file
    '
    Dim slStr As String
    Dim slAirDate As String
    Dim slAirTime As String
    Dim llLstCode As Long
    Dim llChfCode As Long
    Dim ilLineNo As Integer
    Dim llSdfCode As Long
    Dim slSdfDate As String
    Dim slSdfTime As String
    Dim slSdfSchStatus As String
    Dim slSdfSpotType As String
    Dim slMissedDate As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slStart As String
    Dim slEnd As String
    Dim ilLoop As Integer
    Dim slLstStartDate As String
    Dim slLstEndDate As String
    Dim slClfStartDate As String
    Dim slClfEndDate As String
    Dim slCffStartDate As String
    Dim slCffEndDate As String
    Dim ilDay As Integer
    Dim ilRdf As Integer
    Dim slPledgeDate As String
    Dim slPledgeStartTime As String
    Dim slPledgeEndTime As String
    Dim ilSchdCount As Integer
    Dim ilAiredCount As Integer
    Dim ilPledgeCompliantCount As Integer
    Dim ilAgyCompliantCount As Integer
    Dim ilCffDay(0 To 6) As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand:
    gGetLineParameters = 0
    For ilLoop = 0 To 6 Step 1
        ilAllowedDays(ilLoop) = 0
    Next ilLoop
    lgSTime1 = timeGetTime
    slAirDate = Trim$(tlAstInfo.sAirDate)
    slAirTime = Trim$(tlAstInfo.sAirTime)
    llLstCode = tlAstInfo.lLstCode
    slPledgeDate = Trim$(tlAstInfo.sPledgeDate)
    slPledgeStartTime = Trim$(tlAstInfo.sPledgeStartTime)
    slPledgeEndTime = Trim$(tlAstInfo.sPledgeEndTime)
    'End time will be blank if Pledge is live
    If (Trim$(slPledgeEndTime) = "") Or (Asc(slPledgeEndTime) = 0) Then
        slPledgeEndTime = Trim$(tlAstInfo.sTruePledgeEndTime)
    End If
    
    '8/1/14: Not used with v7.0
    'If sgMarketronCompliant <> "A" Then
    '    slAllowedStartDate = slPledgeDate
    '    slAllowedEndDate = slPledgeDate
    '    slAllowedStartTime = slPledgeStartTime
    '    slAllowedEndTime = slPledgeEndTime
    '    If Second(slAllowedStartTime) = 0 Then
    '        slAllowedStartTime = Format(slAllowedStartTime, sgShowTimeWOSecForm)
    '    End If
    '    If Second(slAllowedEndTime) = 0 Then
    '        slAllowedEndTime = Format(slAllowedEndTime, sgShowTimeWOSecForm)
    '    End If
    '    ilAllowedDays(gWeekDayLong(gDateValue(slAllowedStartDate))) = 1
    '    If blGetCompliant Then
    '        ilSchdCount = 0
    '        ilAiredCount = 0
    '        ilPledgeCompliantCount = 0
    '        ilAgyCompliantCount = 0
    '        gIncSpotCounts tlAstInfo, ilSchdCount, ilAiredCount, ilPledgeCompliantCount, ilAgyCompliantCount
    '        If ilPledgeCompliantCount = 0 Then
    '            ilCompliant = False
    '        Else
    '            ilCompliant = True
    '        End If
    '    End If
    '    gGetLineParameters = 0
    '    Exit Function
    'End If
    
    '4/2/15: Restore marketron values when exporting to marketron
    If blMarketronExport Then
        slAllowedStartDate = slPledgeDate
        slAllowedEndDate = slPledgeDate
        slAllowedStartTime = slPledgeStartTime
        slAllowedEndTime = slPledgeEndTime
        If Second(slAllowedStartTime) = 0 Then
            slAllowedStartTime = Format(slAllowedStartTime, sgShowTimeWOSecForm)
        End If
        If Second(slAllowedEndTime) = 0 Then
            slAllowedEndTime = Format(slAllowedEndTime, sgShowTimeWOSecForm)
        End If
        ilAllowedDays(gWeekDayLong(gDateValue(slAllowedStartDate))) = 1
        gGetLineParameters = 0
        Exit Function
    End If
    
    lgETime1 = timeGetTime
    lgTtlTime1 = lgTtlTime1 + (lgETime1 - lgSTime1)
    lgSTime2 = timeGetTime
    
    'SQLQuery = "SELECT *"
    'SQLQuery = SQLQuery + " FROM lst"
    'SQLQuery = SQLQuery + " WHERE (lstCode = " & llLstCode & ")"
    'Set lst_rst = gSQLSelectCall(SQLQuery)
    
    'If lst_rst.EOF Then
    '    'Use ast values as output as unable to read lst
    '    ilCompliant = False
    '    slAllowedStartDate = slPledgeDate
    '    slAllowedEndDate = slPledgeDate
    '    slAllowedStartTime = slPledgeStartTime
    '    slAllowedEndTime = slPledgeEndTime
    '    If Second(slAllowedStartTime) = 0 Then
    '        slAllowedStartTime = Format(slAllowedStartTime, sgShowTimeWOSecForm)
    '    End If
    '    If Second(slAllowedEndTime) = 0 Then
    '        slAllowedEndTime = Format(slAllowedEndTime, sgShowTimeWOSecForm)
    '    End If
    '    ilAllowedDays(gWeekDayLong(gDateValue(slAllowedStartDate))) = 1
    '    gGetLineParameters = 2
    '    Exit Function
    'End If
    'If lst_rst!lstBkoutLstCode > 0 Then
    If tlAstInfo.lLstBkoutLstCode > 0 Then
        '8/1/14: Compliant not required in v7.0
        'ilCompliant = False
        slAllowedStartDate = slPledgeDate
        slAllowedEndDate = slPledgeDate
        slAllowedStartTime = slPledgeStartTime
        slAllowedEndTime = slPledgeEndTime
        If Second(slAllowedStartTime) = 0 Then
            slAllowedStartTime = Format(slAllowedStartTime, sgShowTimeWOSecForm)
        End If
        If Second(slAllowedEndTime) = 0 Then
            slAllowedEndTime = Format(slAllowedEndTime, sgShowTimeWOSecForm)
        End If
        ilAllowedDays(gWeekDayLong(gDateValue(slAllowedStartDate))) = 1
        gGetLineParameters = 6
        Exit Function
    End If
    'llSdfCode = CLng(lst_rst!lstSdfCode)
    llSdfCode = tlAstInfo.lSdfCode
    lgETime2 = timeGetTime
    lgTtlTime2 = lgTtlTime2 + (lgETime2 - lgSTime2)
    lgSTime3 = timeGetTime
    'slLstStartDate = Format(lst_rst!lstStartDate, sgShowDateForm)
    slLstStartDate = Trim$(tlAstInfo.sLstStartDate)
    'slLstEndDate = Format(lst_rst!lstEndDate, sgShowDateForm)
    slLstEndDate = Trim$(tlAstInfo.sLstEndDate)
    SQLQuery = "SELECT *"
    SQLQuery = SQLQuery + " FROM sdf_Spot_Detail"
    SQLQuery = SQLQuery + " WHERE (sdfCode = " & llSdfCode & ")"
    Set sdf_rst = gSQLSelectCall(SQLQuery)
    

    If sdf_rst.EOF Then
        'Use ast values
        slAllowedStartDate = slPledgeDate
        slAllowedEndDate = slPledgeDate
        slAllowedStartTime = slPledgeStartTime
        slAllowedEndTime = slPledgeEndTime
        '8/1/14: Compliant not required in v7.0
        'ilCompliant = True
        'If (gTimeToLong(slAirTime, False) < gTimeToLong(slAllowedStartTime, False)) Or (gTimeToLong(slAirTime, True) > gTimeToLong(slAllowedEndTime, True)) Then
        '    ilCompliant = False
        'End If
        'If (gDateValue(slAirDate) < gDateValue(slAllowedStartDate)) Or (gDateValue(slAirDate) > gDateValue(slAllowedEndDate)) Then
        '    ilCompliant = False
        'End If
        If Second(slAllowedStartTime) = 0 Then
            slAllowedStartTime = Format(slAllowedStartTime, sgShowTimeWOSecForm)
        End If
        If Second(slAllowedEndTime) = 0 Then
            slAllowedEndTime = Format(slAllowedEndTime, sgShowTimeWOSecForm)
        End If
        ilAllowedDays(gWeekDayLong(gDateValue(slAllowedStartDate))) = 1
        gGetLineParameters = 3
        Exit Function
    End If
    'If sdf_rst!sdfVefCode <> lst_rst!lstLnVefCode Then
'MG, Outside and Files are all ways Compliant
'    If sdf_rst!sdfVefCode <> tlAstInfo.iLstLnVefCode Then
'        'Use ast values
'        slAllowedStartDate = slPledgeDate
'        slAllowedEndDate = slPledgeDate
'        slAllowedStartTime = slPledgeStartTime
'        slAllowedEndTime = slPledgeEndTime
'        ilCompliant = True
'        If (gTimeToLong(slAirTime, False) < gTimeToLong(slAllowedStartTime, False)) Or (gTimeToLong(slAirTime, True) > gTimeToLong(slAllowedEndTime, True)) Then
'            ilCompliant = False
'        End If
'        If (gDateValue(slAirDate) < gDateValue(slAllowedStartDate)) Or (gDateValue(slAirDate) > gDateValue(slAllowedEndDate)) Then
'            ilCompliant = False
'        End If
'        If Second(slAllowedStartTime) = 0 Then
'            slAllowedStartTime = Format(slAllowedStartTime, sgShowTimeWOSecForm)
'        End If
'        If Second(slAllowedEndTime) = 0 Then
'            slAllowedEndTime = Format(slAllowedEndTime, sgShowTimeWOSecForm)
'        End If
'        ilAllowedDays(gWeekDayLong(gDateValue(slAllowedStartDate))) = 1
'        gGetLineParameters = 7
'        Exit Function
'    End If
    llChfCode = sdf_rst!sdfChfCode
    ilLineNo = sdf_rst!sdfLineNo

lgETime3 = timeGetTime
lgTtlTime3 = lgTtlTime3 + (lgETime3 - lgSTime3)

        
lgSTime5 = timeGetTime


    slSdfDate = Format(sdf_rst!sdfDate, sgShowDateForm)
    slSdfTime = Format(sdf_rst!sdfTime, sgShowTimeWSecForm)
    slSdfSchStatus = sdf_rst!sdfSchStatus
    slSdfSpotType = sdf_rst!sdfSpotType
    
    If (tmClfInfo.lChfCode <> llChfCode) Or (tmClfInfo.iLineNo <> ilLineNo) Then
        If (tmClfInfo.lChfCode <> llChfCode) Then
            SQLQuery = "SELECT chfCode, chfBillCycle"
            SQLQuery = SQLQuery + " FROM CHF_Contract_Header"
            SQLQuery = SQLQuery + " WHERE (chfCode = " & llChfCode & ")"
            Set chf_rst = gSQLSelectCall(SQLQuery)
            If chf_rst.EOF Then
                slAllowedStartDate = slPledgeDate
                slAllowedEndDate = slPledgeDate
                slAllowedStartTime = slPledgeStartTime
                slAllowedEndTime = slPledgeEndTime
                '8/1/14: Compliant not required in v7.0
                'ilCompliant = True
                'If (gTimeToLong(slAirTime, False) < gTimeToLong(slAllowedStartTime, False)) Or (gTimeToLong(slAirTime, True) > gTimeToLong(slAllowedEndTime, True)) Then
                '    ilCompliant = False
                'End If
                'If (gDateValue(slAirDate) < gDateValue(slAllowedStartDate)) Or (gDateValue(slAirDate) > gDateValue(slAllowedEndDate)) Then
                '    ilCompliant = False
                'End If
                If Second(slAllowedStartTime) = 0 Then
                    slAllowedStartTime = Format(slAllowedStartTime, sgShowTimeWOSecForm)
                End If
                If Second(slAllowedEndTime) = 0 Then
                    slAllowedEndTime = Format(slAllowedEndTime, sgShowTimeWOSecForm)
                End If
                ilAllowedDays(gWeekDayLong(gDateValue(slAllowedStartDate))) = 1
                gGetLineParameters = 4
                Exit Function
            End If
            tmClfInfo.lChfCode = llChfCode
            tmClfInfo.sBillCycle = chf_rst!chfBillCycle
        End If

        SQLQuery = "SELECT *"
        SQLQuery = SQLQuery + " FROM clf_Contract_Line"
        SQLQuery = SQLQuery + " WHERE (clfChfCode = " & llChfCode
        SQLQuery = SQLQuery + " AND clfLine = " & ilLineNo
        SQLQuery = SQLQuery + " AND clfSchStatus = 'F' AND clfDelete = 'N'" & ")"
        Set clf_rst = gSQLSelectCall(SQLQuery)
        
    
        If clf_rst.EOF Then
            slAllowedStartDate = slPledgeDate
            slAllowedEndDate = slPledgeDate
            slAllowedStartTime = slPledgeStartTime
            slAllowedEndTime = slPledgeEndTime
            '8/1/14: Compliant not required in v7.0
            'ilCompliant = True
            'If (gTimeToLong(slAirTime, False) < gTimeToLong(slAllowedStartTime, False)) Or (gTimeToLong(slAirTime, True) > gTimeToLong(slAllowedEndTime, True)) Then
            '    ilCompliant = False
            'End If
            'If (gDateValue(slAirDate) < gDateValue(slAllowedStartDate)) Or (gDateValue(slAirDate) > gDateValue(slAllowedEndDate)) Then
            '    ilCompliant = False
            'End If
            If Second(slAllowedStartTime) = 0 Then
                slAllowedStartTime = Format(slAllowedStartTime, sgShowTimeWOSecForm)
            End If
            If Second(slAllowedEndTime) = 0 Then
                slAllowedEndTime = Format(slAllowedEndTime, sgShowTimeWOSecForm)
            End If
            ilAllowedDays(gWeekDayLong(gDateValue(slAllowedStartDate))) = 1
            gGetLineParameters = 4
            Exit Function
        End If
        tmClfInfo.lCode = clf_rst!clfCode
        tmClfInfo.lChfCode = llChfCode
        tmClfInfo.iLineNo = clf_rst!clfLine
        tmClfInfo.iCntRevNo = clf_rst!clfCntRevNo
        tmClfInfo.iPropVer = clf_rst!clfPropVer
        tmClfInfo.sStartDate = Format(clf_rst!clfStartDate, sgShowDateForm)
        tmClfInfo.sEndDate = Format(clf_rst!clfEndDate, sgShowDateForm)
        If (InStr(1, clf_rst!clfStartTime, " ", vbBinaryCompare) = 0) And (InStr(1, clf_rst!clfEndTime, " ", vbBinaryCompare) = 0) Then
            tmClfInfo.sStartTime = ""
            tmClfInfo.sEndTime = ""
        Else
            tmClfInfo.sStartTime = Format(clf_rst!clfStartTime, sgShowTimeWSecForm)
            tmClfInfo.sEndTime = Format(clf_rst!clfEndTime, sgShowTimeWSecForm)
        End If
        tmClfInfo.iRdfCode = clf_rst!clfRdfCode
    End If
lgETime5 = timeGetTime
lgTtlTime5 = lgTtlTime5 + (lgETime5 - lgSTime5)

    slClfStartDate = tmClfInfo.sStartDate   'Format(clf_rst!clfStartDate, sgShowDateForm)
    slClfEndDate = tmClfInfo.sEndDate   'Format(clf_rst!clfEndDate, sgShowDateForm)
    'Problem:  Override time of 12m-12m will not be uncovered as the clfStartTime will only have the date
    'Note: 12m returns on the date as does a illegal value
lgSTime6 = timeGetTime
    'If (InStr(1, clf_rst!clfStartTime, " ", vbBinaryCompare) = 0) And (InStr(1, clf_rst!clfEndTime, " ", vbBinaryCompare) = 0) Then
    If Trim$(tmClfInfo.sStartTime) = "" Then

        'No override times- get times from rdf
        '7/11/13: Check that daypart array have been populated
        On Error GoTo PopDaypart
        ilRet = UBound(tgDaypartInfo)
        On Error GoTo ErrHand:
        ilRdf = gBinarySearchRdf(tmClfInfo.iRdfCode)
        If ilRdf = -1 Then
            'Use ast values
            '8/1/14: Compliant not required in v7.0
            'ilCompliant = False
            slAllowedStartDate = slPledgeDate
            slAllowedEndDate = slPledgeDate
            slAllowedStartTime = slPledgeStartTime
            slAllowedEndTime = slPledgeEndTime
            '8/1/14: Compliant not required in v7.0
            'ilCompliant = True
            'If (gTimeToLong(slAirTime, False) < gTimeToLong(slPledgeStartTime, False)) Or (gTimeToLong(slAirTime, True) > gTimeToLong(slAllowedEndTime, True)) Then
            '    ilCompliant = False
            'End If
            'If (gDateValue(slAirDate) < gDateValue(slStartDate)) Or (gDateValue(slAirDate) > gDateValue(slEndDate)) Then
            '    ilCompliant = False
            'End If
            If Second(slAllowedStartTime) = 0 Then
                slAllowedStartTime = Format(slAllowedStartTime, sgShowTimeWOSecForm)
            End If
            If Second(slAllowedEndTime) = 0 Then
                slAllowedEndTime = Format(slAllowedEndTime, sgShowTimeWOSecForm)
            End If
            ilAllowedDays(gWeekDayLong(gDateValue(slAllowedStartDate))) = 1
            gGetLineParameters = 5
            Exit Function
        End If
        slStartTime = tgDaypartInfo(ilRdf).sStartTime
        slEndTime = tgDaypartInfo(ilRdf).sEndTime

    Else
        'Override times defined
        slStartTime = tmClfInfo.sStartTime  'Format(clf_rst!clfStartTime, sgShowTimeWSecForm)
        slEndTime = tmClfInfo.sEndTime    'Format(clf_rst!clfStartTime, sgShowTimeWSecForm)
    End If
        
lgETime6 = timeGetTime
lgTtlTime6 = lgTtlTime6 + (lgETime6 - lgSTime6)
        

    'Determine date range
    slStartDate = slSdfDate
    slEndDate = slSdfDate
    ilDay = gWeekDayLong(gDateValue(slSdfDate))
    For ilLoop = 0 To 6 Step 1
        ilCffDay(ilLoop) = 0
    Next ilLoop
    If (gDateValue(slSdfDate) >= gDateValue(slLstStartDate)) And (gDateValue(slSdfDate) <= gDateValue(slLstEndDate)) Then
        'If lst_rst!lstSpotsWk > 0 Then
        If tlAstInfo.iLstSpotsWk > 0 Then
            ilCffDay(0) = tlAstInfo.iLstMon  'lst_rst!lstMon
            ilCffDay(1) = tlAstInfo.iLstTue  'lst_rst!lstTue
            ilCffDay(2) = tlAstInfo.iLstWed  'lst_rst!lstWed
            ilCffDay(3) = tlAstInfo.iLstThu  'lst_rst!lstThu
            ilCffDay(4) = tlAstInfo.iLstFri  'lst_rst!lstFri
            ilCffDay(5) = tlAstInfo.iLstSat  'lst_rst!lstSat
            ilCffDay(6) = tlAstInfo.iLstSun  'lst_rst!lstSun
        Else
            Select Case ilDay
                Case 0
                    ilCffDay(0) = tlAstInfo.iLstMon  'lst_rst!lstMon
                Case 1
                    ilCffDay(1) = tlAstInfo.iLstTue  'lst_rst!lstTue
                Case 2
                    ilCffDay(2) = tlAstInfo.iLstWed  'lst_rst!lstWed
                Case 3
                    ilCffDay(3) = tlAstInfo.iLstThu  'lst_rst!lstThu
                Case 4
                    ilCffDay(4) = tlAstInfo.iLstFri  'lst_rst!lstFri
                Case 5
                    ilCffDay(5) = tlAstInfo.iLstSat  'lst_rst!lstSat
                Case 6
                    ilCffDay(6) = tlAstInfo.iLstSun  'lst_rst!lstSun
            End Select
        End If
    Else
lgSTime7 = timeGetTime

        SQLQuery = "SELECT *"
        SQLQuery = SQLQuery + " FROM cff_Contract_Flight"
        SQLQuery = SQLQuery + " WHERE (cffChfCode = " & llChfCode
        SQLQuery = SQLQuery + " AND cffClfLine = " & ilLineNo
        SQLQuery = SQLQuery + " AND cffCntRevNo = " & tmClfInfo.iCntRevNo 'clf_rst!clfCntRevNo
        SQLQuery = SQLQuery + " AND cffPropVer = " & tmClfInfo.iPropVer   'clf_rst!clfPropVer
        'SQLQuery = SQLQuery + " AND cffStartDate >= '" & Format(slSdfDate, sgSQLDateForm) & "')"
        SQLQuery = SQLQuery + " AND cffStartDate >= '" & Format("12/28/2009", sgSQLDateForm) & "')"
        Set cff_rst = gSQLSelectCall(SQLQuery)

        If Not cff_rst.EOF Then
            Do While Not cff_rst.EOF
                slCffStartDate = Format(cff_rst!cffStartDate, sgShowDateForm)
                slCffEndDate = Format(cff_rst!cffEndDate, sgShowDateForm)
                If (gDateValue(slSdfDate) >= gDateValue(slCffStartDate)) And (gDateValue(slSdfDate) <= gDateValue(slCffEndDate)) Then
                    If cff_rst!cffSpotsWk > 0 Then
                        ilCffDay(0) = cff_rst!lstMo
                        ilCffDay(1) = cff_rst!lstTu
                        ilCffDay(2) = cff_rst!lstWe
                        ilCffDay(3) = cff_rst!lstTh
                        ilCffDay(4) = cff_rst!lstFr
                        ilCffDay(5) = cff_rst!lstSa
                        ilCffDay(6) = cff_rst!lstSu
                    Else
                        Select Case ilDay
                            Case 0
                                ilCffDay(0) = cff_rst!lstMo
                            Case 1
                                ilCffDay(1) = cff_rst!lstTu
                            Case 2
                                ilCffDay(2) = cff_rst!lstWe
                            Case 3
                                ilCffDay(3) = cff_rst!lstTh
                            Case 4
                                ilCffDay(4) = cff_rst!lstFr
                            Case 5
                                ilCffDay(5) = cff_rst!lstSa
                            Case 6
                                ilCffDay(6) = cff_rst!lstSu
                        End Select
                    End If
                    Exit Do
                End If
                cff_rst.MoveNext
            Loop
        End If
lgETime7 = timeGetTime
lgTtlTime7 = lgTtlTime7 + (lgETime7 - lgSTime7)

    End If
    If ilCffDay(ilDay) > 0 Then
        For ilLoop = ilDay - 1 To 0 Step -1
            If ilCffDay(ilLoop) > 0 Then
                slStartDate = DateAdd("d", -1, slStartDate)
            Else
                Exit For
            End If
        Next ilLoop
        ilDay = gWeekDayLong(gDateValue(slSdfDate))
        For ilLoop = ilDay + 1 To 6 Step 1
            If ilCffDay(ilLoop) > 0 Then
                slEndDate = DateAdd("d", 1, slEndDate)
            Else
                Exit For
            End If
        Next ilLoop
        If gDateValue(slStartDate) < gDateValue(slClfStartDate) Then
            slStartDate = slClfStartDate
        End If
        If gDateValue(slEndDate) > gDateValue(slClfEndDate) Then
            slEndDate = slClfEndDate
        End If
    Else
        slStartDate = slPledgeDate  'Format(ast_rst!astPledgeDate, sgShowDateForm)
        slEndDate = slStartDate
    End If
    slAllowedStartDate = slStartDate
    slAllowedEndDate = slEndDate
    If Second(slStartTime) = 0 Then
        slAllowedStartTime = Format(slStartTime, sgShowTimeWOSecForm)
    Else
        slAllowedStartTime = slStartTime
    End If
    If Second(slEndTime) = 0 Then
        slAllowedEndTime = Format(slEndTime, sgShowTimeWOSecForm)
    Else
        slAllowedEndTime = slEndTime
    End If
    For ilLoop = 0 To 6 Step 1
        ilAllowedDays(ilLoop) = ilCffDay(ilLoop)
    Next ilLoop
    If blGetCompliant Then
        'If Not Aired, then compliant is false
        'If tgStatusTypes(gGetAirStatus(ast_rst!astStatus)).iPledged = 2 Then
        If tgStatusTypes(gGetAirStatus(tlAstInfo.iStatus)).iPledged = 2 Then
            '8/1/14: Compliant not required in v7.0
            'ilCompliant = False
        Else
            '8/1/14: Compliant not required in v7.0
            'ilCompliant = True
            For ilLoop = 0 To 6 Step 1
                ilCffDay(ilLoop) = 0
            Next ilLoop
            ilDay = gWeekDayLong(gDateValue(slAirDate))
            If (gDateValue(slAirDate) >= gDateValue(slLstStartDate)) And (gDateValue(slAirDate) <= gDateValue(slLstEndDate)) Then
                'If lst_rst!lstSpotsWk > 0 Then
                If tlAstInfo.iLstSpotsWk > 0 Then
                    ilCffDay(0) = tlAstInfo.iLstMon  'lst_rst!lstMon
                    ilCffDay(1) = tlAstInfo.iLstTue  'lst_rst!lstTue
                    ilCffDay(2) = tlAstInfo.iLstWed  'lst_rst!lstWed
                    ilCffDay(3) = tlAstInfo.iLstThu  'lst_rst!lstThu
                    ilCffDay(4) = tlAstInfo.iLstFri  'lst_rst!lstFri
                    ilCffDay(5) = tlAstInfo.iLstSat  'lst_rst!lstSat
                    ilCffDay(6) = tlAstInfo.iLstSun  'lst_rst!lstSun
                Else
                    Select Case ilDay
                        Case 0
                            ilCffDay(0) = tlAstInfo.iLstMon  'lst_rst!lstMon
                        Case 1
                            ilCffDay(1) = tlAstInfo.iLstTue  'lst_rst!lstTue
                        Case 2
                            ilCffDay(2) = tlAstInfo.iLstWed  'lst_rst!lstWed
                        Case 3
                            ilCffDay(3) = tlAstInfo.iLstThu  'lst_rst!lstThu
                        Case 4
                            ilCffDay(4) = tlAstInfo.iLstFri  'lst_rst!lstFri
                        Case 5
                            ilCffDay(5) = tlAstInfo.iLstSat  'lst_rst!lstSat
                        Case 6
                            ilCffDay(6) = tlAstInfo.iLstSun  'lst_rst!lstSun
                    End Select
                End If
            Else
            
    lgSTime7 = timeGetTime
    
                SQLQuery = "SELECT *"
                SQLQuery = SQLQuery + " FROM cff_Contract_Flight"
                SQLQuery = SQLQuery + " WHERE (cffChfCode = " & llChfCode
                SQLQuery = SQLQuery + " AND cffClfLine = " & ilLineNo
                SQLQuery = SQLQuery + " AND cffCntRevNo = " & tmClfInfo.iCntRevNo 'clf_rst!clfCntRevNo
                SQLQuery = SQLQuery + " AND cffPropVer = " & tmClfInfo.iPropVer   'clf_rst!clfPropVer
                'SQLQuery = SQLQuery + " AND cffStartDate >= '" & Format(slAirDate, sgSQLDateForm) & "')"
                SQLQuery = SQLQuery + " AND cffStartDate >= '" & Format("12/28/2009", sgSQLDateForm) & "')"
                Set cff_rst = gSQLSelectCall(SQLQuery)
                
    
                If Not cff_rst.EOF Then
                    Do While Not cff_rst.EOF
                        slCffStartDate = Format(cff_rst!cffStartDate, sgShowDateForm)
                        slCffEndDate = Format(cff_rst!cffEndDate, sgShowDateForm)
                        If (gDateValue(slAirDate) >= gDateValue(slCffStartDate)) And (gDateValue(slAirDate) <= gDateValue(slCffEndDate)) Then
                            If cff_rst!cffSpotsWk > 0 Then
                                ilCffDay(0) = cff_rst!lstMo
                                ilCffDay(1) = cff_rst!lstTu
                                ilCffDay(2) = cff_rst!lstWe
                                ilCffDay(3) = cff_rst!lstTh
                                ilCffDay(4) = cff_rst!lstFr
                                ilCffDay(5) = cff_rst!lstSa
                                ilCffDay(6) = cff_rst!lstSu
                            Else
                                Select Case ilDay
                                    Case 0
                                        ilCffDay(0) = cff_rst!lstMo
                                    Case 1
                                        ilCffDay(1) = cff_rst!lstTu
                                    Case 2
                                        ilCffDay(2) = cff_rst!lstWe
                                    Case 3
                                        ilCffDay(3) = cff_rst!lstTh
                                    Case 4
                                        ilCffDay(4) = cff_rst!lstFr
                                    Case 5
                                        ilCffDay(5) = cff_rst!lstSa
                                    Case 6
                                        ilCffDay(6) = cff_rst!lstSu
                                End Select
                            End If
                            Exit Do
                        End If
                        cff_rst.MoveNext
                    Loop
                End If
    lgETime7 = timeGetTime
    lgTtlTime7 = lgTtlTime7 + (lgETime7 - lgSTime7)
    
            End If
            'MG, Outside and fill treated compliant
            'If (slSdfSchStatus <> "G") And (slSdfSchStatus <> "O") And (slSdfSpotType <> "X") Then
            '8/1/14: Compliant not required in v7.0
            'If (slSdfSpotType <> "X") Then
            '    If ilCffDay(ilDay) <= 0 Then
            '        ilCompliant = False
            '    End If
            '    If (gTimeToLong(slAirTime, False) < gTimeToLong(slStartTime, False)) Or (gTimeToLong(slAirTime, True) > gTimeToLong(slEndTime, True)) Then
            '        ilCompliant = False
            '    End If
            'End If
        End If
        '8/1/14: Compliant not required in v7.0
        'If (ilCompliant) And (slSdfSpotType <> "X") Then
        '    If ((slSdfSchStatus = "G") Or (slSdfSchStatus = "O")) Then
        '        'Test air week against sell week
        '        SQLQuery = "SELECT smfMissedDate"
        '        SQLQuery = SQLQuery + " FROM smf_Spot_MG_Specs"
        '        SQLQuery = SQLQuery + " WHERE (smfSdfCode = " & llSdfCode & ")"
        '        Set smf_rst = gSQLSelectCall(SQLQuery)
        '        If Not smf_rst.EOF Then
        '            slMissedDate = Format(smf_rst!smfMissedDate, sgShowDateForm)
        '            If gDateValue(gObtainPrevMonday(slMissedDate)) <> gDateValue(gObtainPrevMonday(slAirDate)) Then
        '                If (tmClfInfo.sBillCycle <> "W") And (tmClfInfo.sBillCycle <> "C") Then
        '                    If gDateValue(gObtainEndStd(slMissedDate)) <> gDateValue(gObtainEndStd(slAirDate)) Then
        '                        ilCompliant = False
        '                    End If
        '                Else
        '                    ilCompliant = False
        '                End If
        '            End If
        '        End If
        '    Else
        '        If (ilCompliant) Then
        '            If gDateValue(gObtainPrevMonday(slSdfDate)) <> gDateValue(gObtainPrevMonday(slAirDate)) Then
        '                If (tmClfInfo.sBillCycle <> "W") And (tmClfInfo.sBillCycle <> "C") Then
        '                    If gDateValue(gObtainEndStd(slSdfDate)) <> gDateValue(gObtainEndStd(slAirDate)) Then
        '                        ilCompliant = False
        '                    End If
        '                Else
        '                    ilCompliant = False
        '                End If
        '            End If
        '        End If
        '    End If
        'End If
    End If
    On Error Resume Next
    'ast_rst.Close
    'lst_rst.Close
    sdf_rst.Close
    smf_rst.Close
    chf_rst.Close
    clf_rst.Close
    cff_rst.Close
    'rdf_rst.Close
    gGetLineParameters = 0
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "gGetLineParameters"
    gGetLineParameters = 8
    Exit Function
PopDaypart:
    gPopDaypart
    Resume Next
End Function

Public Function gGetAgyCompliant(tlAstInfo As ASTINFO, slAllowedStartDate As String, slAllowedEndDate As String, slAllowedStartTime As String, slAllowedEndTime As String, ilAllowedDays() As Integer, ilCompliant As Integer) As Integer
    'Input Values
    '   tlAstInfo: AST Reference obtained from gGetAst
    '
    'Return values
    '   slAllowedStartDate: Start Date with week that spot can air (Obtained from Flight)
    '   slAllowedEndDate: End Date with week that spot can air (Obtained from Flight)
    '   slAllowedStartTime: Start Time with week that spot can air (Obtained from Line and Daypart)
    '   slAllowedEndTime: End Time with week that spot can air (Obtained from Line and Daypart)
    '   ilAllowedDays: Allowed days that the spot can air within (Obtained from Flight)
    '   ilCompliant:  Set from Aired Date and time compared to Allowed Date/Time/Day
    '   gGetAgyCompliant:
    '      0=Ok
    '      1=Unable to read AST, returning Pledge date/time as Allowed Date/Time, Compliant = False
    '      2=Unable to read LST, returning Pledge date/time as Allowed Date/Time, Compliant = False
    '      3=Unable to read SDF, returning Pledge date/time as Allowed Date/Time/Day, Compliant determined from Allowed Date/Time
    '      4=Unable to read CLF, returning Pledge date/time as Allowed Date/Time/Day, Compliant determined from Allowed Date/Time
    '      5=Unable to read RDF, returning Pledge date/time as Allowed Date/Time/Day, Compliant determined from Allowed Date/Time
    '      6=Blackout, returning Pledge date/time as Allowed Date/Time/Day, Compliant = False
    '      7=Line and Booked vehicles don't match, returning Pledge date/time as Allowed Date/Time/Day, Compliant determined from Allowed Date/Time
    '      8=SQL Error Logged to file
    '
    Dim slStr As String
    Dim slAirDate As String
    Dim slAirTime As String
    Dim llLstCode As Long
    Dim llChfCode As Long
    Dim ilLineNo As Integer
    Dim llSdfCode As Long
    Dim slSdfDate As String
    Dim slSdfTime As String
    Dim slSdfSchStatus As String
    Dim slSdfSpotType As String
    Dim slMissedDate As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slStart As String
    Dim slEnd As String
    Dim ilLoop As Integer
    Dim slLstStartDate As String
    Dim slLstEndDate As String
    Dim slClfStartDate As String
    Dim slClfEndDate As String
    Dim slCffStartDate As String
    Dim slCffEndDate As String
    Dim ilDay As Integer
    Dim ilRdf As Integer
    Dim slPledgeDate As String
    Dim slPledgeStartTime As String
    Dim slPledgeEndTime As String
    Dim ilSchdCount As Integer
    Dim ilAiredCount As Integer
    Dim ilPledgeCompliantCount As Integer
    Dim ilAgyCompliantCount As Integer
    Dim ilCffDay(0 To 6) As Integer
    Dim ilRet As Integer
        
    Dim ilVef As Integer
    Dim ilShtt As Integer
    Dim ilZone As Integer
    Dim slZone As String
    Dim ilLocalAdj As Integer
    Dim llAirDate As Long
    Dim llAirTime As Long
    
    On Error GoTo ErrHand:
    gGetAgyCompliant = 0
    For ilLoop = 0 To 6 Step 1
        ilAllowedDays(ilLoop) = 0
    Next ilLoop
    slAirDate = Trim$(tlAstInfo.sAirDate)
    slAirTime = Trim$(tlAstInfo.sAirTime)

    '10/30/14: Remap air time to sold time
    ilLocalAdj = 0
    On Error GoTo VefPop
    ilRet = UBound(tgVehicleInfo)
    On Error GoTo ErrHand:
    ilVef = gBinarySearchVef(CLng(tlAstInfo.iVefCode))
    On Error GoTo ShttPop:
    ilRet = UBound(tgShttInfo1)
    On Error GoTo ErrHand:
    ilShtt = gBinarySearchShtt(tlAstInfo.iShttCode)
    If ilShtt <> -1 Then
        slZone = Trim$(tgShttInfo1(ilShtt).shttTimeZone)
        If ilVef <> -1 Then
            For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                If Trim$(tgVehicleInfo(ilVef).sZone(ilZone)) = slZone Then
                    ilLocalAdj = tgVehicleInfo(ilVef).iLocalAdj(ilZone)
                    Exit For
                End If
            Next ilZone
        End If
    End If
    If ilLocalAdj <> 0 Then
        ilLocalAdj = -ilLocalAdj
    End If
    llAirDate = DateValue(gAdjYear(slAirDate))
    llAirTime = gTimeToLong(slAirTime, False)
    
    llAirTime = llAirTime + 3600 * ilLocalAdj
    If llAirTime < 0 Then
        llAirTime = llAirTime + 86400
        llAirDate = llAirDate - 1
    ElseIf llAirTime > 86400 Then
        llAirTime = llAirTime - 86400
        llAirDate = llAirDate + 1
    End If
    slAirTime = Format$(gLongToTime(llAirTime), "h:mm:ssAM/PM")
    slAirDate = Format$(llAirDate, "m/d/yyyy")
    
    
    llLstCode = tlAstInfo.lLstCode
    slPledgeDate = Trim$(tlAstInfo.sPledgeDate)
    slPledgeStartTime = Trim$(tlAstInfo.sPledgeStartTime)
    slPledgeEndTime = Trim$(tlAstInfo.sPledgeEndTime)
    'End time will be blank if Pledge is live
    If (Trim$(slPledgeEndTime) = "") Then
        slPledgeEndTime = Trim$(tlAstInfo.sTruePledgeEndTime)
    ElseIf (Asc(slPledgeEndTime) = 0) Then
        slPledgeEndTime = Trim$(tlAstInfo.sTruePledgeEndTime)
    End If
    
    
    If tlAstInfo.lLstBkoutLstCode > 0 Then
        ilCompliant = False
        slAllowedStartDate = slPledgeDate
        slAllowedEndDate = slPledgeDate
        slAllowedStartTime = slPledgeStartTime
        slAllowedEndTime = slPledgeEndTime
        If Second(slAllowedStartTime) = 0 Then
            slAllowedStartTime = Format(slAllowedStartTime, sgShowTimeWOSecForm)
        End If
        If Second(slAllowedEndTime) = 0 Then
            slAllowedEndTime = Format(slAllowedEndTime, sgShowTimeWOSecForm)
        End If
        ilAllowedDays(gWeekDayLong(gDateValue(slAllowedStartDate))) = 1
        gGetAgyCompliant = 6
        Exit Function
    End If
    llSdfCode = tlAstInfo.lSdfCode
    
    slLstStartDate = Trim$(tlAstInfo.sLstStartDate)
    slLstEndDate = Trim$(tlAstInfo.sLstEndDate)
    SQLQuery = "SELECT *"
    SQLQuery = SQLQuery + " FROM sdf_Spot_Detail"
    SQLQuery = SQLQuery + " WHERE (sdfCode = " & llSdfCode & ")"
    Set sdf_rst = gSQLSelectCall(SQLQuery)
    

    If sdf_rst.EOF Then
        'Use ast values
        slAllowedStartDate = slPledgeDate
        slAllowedEndDate = slPledgeDate
        slAllowedStartTime = slPledgeStartTime
        slAllowedEndTime = slPledgeEndTime
        '10/28/18: Since spot is invalid, it can't be compliant
        'ilCompliant = True
        'If (gTimeToLong(slAirTime, False) < gTimeToLong(slAllowedStartTime, False)) Or (gTimeToLong(slAirTime, True) > gTimeToLong(slAllowedEndTime, True)) Then
        '    ilCompliant = False
        'End If
        'If (gDateValue(slAirDate) < gDateValue(slAllowedStartDate)) Or (gDateValue(slAirDate) > gDateValue(slAllowedEndDate)) Then
        '    ilCompliant = False
        'End If
        ilCompliant = False
        If Second(slAllowedStartTime) = 0 Then
            slAllowedStartTime = Format(slAllowedStartTime, sgShowTimeWOSecForm)
        End If
        If Second(slAllowedEndTime) = 0 Then
            slAllowedEndTime = Format(slAllowedEndTime, sgShowTimeWOSecForm)
        End If
        ilAllowedDays(gWeekDayLong(gDateValue(slAllowedStartDate))) = 1
        gGetAgyCompliant = 3
        Exit Function
    End If
    llChfCode = sdf_rst!sdfChfCode
    ilLineNo = sdf_rst!sdfLineNo
    slSdfDate = Format(sdf_rst!sdfDate, sgShowDateForm)
    slSdfTime = Format(sdf_rst!sdfTime, sgShowTimeWSecForm)
    slSdfSchStatus = sdf_rst!sdfSchStatus
    slSdfSpotType = sdf_rst!sdfSpotType
    
    If (tmClfInfo.lChfCode <> llChfCode) Or (tmClfInfo.iLineNo <> ilLineNo) Then
        If (tmClfInfo.lChfCode <> llChfCode) Then
            SQLQuery = "SELECT chfCode, chfBillCycle"
            SQLQuery = SQLQuery + " FROM CHF_Contract_Header"
            SQLQuery = SQLQuery + " WHERE (chfCode = " & llChfCode & ")"
            Set chf_rst = gSQLSelectCall(SQLQuery)
            If chf_rst.EOF Then
                slAllowedStartDate = slPledgeDate
                slAllowedEndDate = slPledgeDate
                slAllowedStartTime = slPledgeStartTime
                slAllowedEndTime = slPledgeEndTime
                '10/28/18: Since contract can't be found, it can't be compliant
                'ilCompliant = True
                'If (gTimeToLong(slAirTime, False) < gTimeToLong(slAllowedStartTime, False)) Or (gTimeToLong(slAirTime, True) > gTimeToLong(slAllowedEndTime, True)) Then
                '    ilCompliant = False
                'End If
                'If (gDateValue(slAirDate) < gDateValue(slAllowedStartDate)) Or (gDateValue(slAirDate) > gDateValue(slAllowedEndDate)) Then
                '    ilCompliant = False
                'End If
                ilCompliant = False
                If Second(slAllowedStartTime) = 0 Then
                    slAllowedStartTime = Format(slAllowedStartTime, sgShowTimeWOSecForm)
                End If
                If Second(slAllowedEndTime) = 0 Then
                    slAllowedEndTime = Format(slAllowedEndTime, sgShowTimeWOSecForm)
                End If
                ilAllowedDays(gWeekDayLong(gDateValue(slAllowedStartDate))) = 1
                gGetAgyCompliant = 4
                Exit Function
            End If
            tmClfInfo.lChfCode = llChfCode
            tmClfInfo.sBillCycle = chf_rst!chfBillCycle
        End If

        SQLQuery = "SELECT *"
        SQLQuery = SQLQuery + " FROM clf_Contract_Line"
        SQLQuery = SQLQuery + " WHERE clfChfCode = " & llChfCode
        SQLQuery = SQLQuery + " AND clfLine = " & ilLineNo
        SQLQuery = SQLQuery + " AND clfDelete = 'N'"
        SQLQuery = SQLQuery + " AND (clfSchStatus = 'F' OR clfSchStatus = 'M'" & ")"
        Set clf_rst = gSQLSelectCall(SQLQuery)
        
    
        If clf_rst.EOF Then
            slAllowedStartDate = slPledgeDate
            slAllowedEndDate = slPledgeDate
            slAllowedStartTime = slPledgeStartTime
            slAllowedEndTime = slPledgeEndTime
            '10/28/18: Since contract line can't be found, it can't be compliant
            'ilCompliant = True
            'If (gTimeToLong(slAirTime, False) < gTimeToLong(slAllowedStartTime, False)) Or (gTimeToLong(slAirTime, True) > gTimeToLong(slAllowedEndTime, True)) Then
            '    ilCompliant = False
            'End If
            'If (gDateValue(slAirDate) < gDateValue(slAllowedStartDate)) Or (gDateValue(slAirDate) > gDateValue(slAllowedEndDate)) Then
            '    ilCompliant = False
            'End If
            ilCompliant = False
            If Second(slAllowedStartTime) = 0 Then
                slAllowedStartTime = Format(slAllowedStartTime, sgShowTimeWOSecForm)
            End If
            If Second(slAllowedEndTime) = 0 Then
                slAllowedEndTime = Format(slAllowedEndTime, sgShowTimeWOSecForm)
            End If
            ilAllowedDays(gWeekDayLong(gDateValue(slAllowedStartDate))) = 1
            gGetAgyCompliant = 4
            Exit Function
        End If
        tmClfInfo.lCode = clf_rst!clfCode
        tmClfInfo.lChfCode = llChfCode
        tmClfInfo.iLineNo = clf_rst!clfLine
        tmClfInfo.iCntRevNo = clf_rst!clfCntRevNo
        tmClfInfo.iPropVer = clf_rst!clfPropVer
        tmClfInfo.sStartDate = Format(clf_rst!clfStartDate, sgShowDateForm)
        tmClfInfo.sEndDate = Format(clf_rst!clfEndDate, sgShowDateForm)
        If (InStr(1, clf_rst!clfStartTime, " ", vbBinaryCompare) = 0) And (InStr(1, clf_rst!clfEndTime, " ", vbBinaryCompare) = 0) Then
            tmClfInfo.sStartTime = ""
            tmClfInfo.sEndTime = ""
        Else
            tmClfInfo.sStartTime = Format(clf_rst!clfStartTime, sgShowTimeWSecForm)
            tmClfInfo.sEndTime = Format(clf_rst!clfEndTime, sgShowTimeWSecForm)
        End If
        tmClfInfo.iRdfCode = clf_rst!clfRdfCode
    End If

    slClfStartDate = tmClfInfo.sStartDate   'Format(clf_rst!clfStartDate, sgShowDateForm)
    slClfEndDate = tmClfInfo.sEndDate   'Format(clf_rst!clfEndDate, sgShowDateForm)
    'Problem:  Override time of 12m-12m will not be uncovered as the clfStartTime will only have the date
    'Note: 12m returns on the date as does a illegal value
    If Trim$(tmClfInfo.sStartTime) = "" Then

        'No override times- get times from rdf
        '7/11/13: Check that daypart array have been populated
        On Error GoTo PopDaypart
        ilRet = UBound(tgDaypartInfo)
        On Error GoTo ErrHand:
        ilRdf = gBinarySearchRdf(tmClfInfo.iRdfCode)
        If ilRdf = -1 Then
            'Use ast values
            ilCompliant = False
            slAllowedStartDate = slPledgeDate
            slAllowedEndDate = slPledgeDate
            slAllowedStartTime = slPledgeStartTime
            slAllowedEndTime = slPledgeEndTime
            '10/28/18: Since daypart can't be found, it can't be compliant
            'ilCompliant = True
            'If (gTimeToLong(slAirTime, False) < gTimeToLong(slPledgeStartTime, False)) Or (gTimeToLong(slAirTime, True) > gTimeToLong(slAllowedEndTime, True)) Then
            '    ilCompliant = False
            'End If
            'If (gDateValue(slAirDate) < gDateValue(slStartDate)) Or (gDateValue(slAirDate) > gDateValue(slEndDate)) Then
            '    ilCompliant = False
            'End If
            ilCompliant = False
            If Second(slAllowedStartTime) = 0 Then
                slAllowedStartTime = Format(slAllowedStartTime, sgShowTimeWOSecForm)
            End If
            If Second(slAllowedEndTime) = 0 Then
                slAllowedEndTime = Format(slAllowedEndTime, sgShowTimeWOSecForm)
            End If
            ilAllowedDays(gWeekDayLong(gDateValue(slAllowedStartDate))) = 1
            gGetAgyCompliant = 5
            Exit Function
        End If
        slStartTime = tgDaypartInfo(ilRdf).sStartTime
        slEndTime = tgDaypartInfo(ilRdf).sEndTime

    Else
        'Override times defined
        slStartTime = tmClfInfo.sStartTime  'Format(clf_rst!clfStartTime, sgShowTimeWSecForm)
        slEndTime = tmClfInfo.sEndTime    'Format(clf_rst!clfStartTime, sgShowTimeWSecForm)
    End If
    'Determine date range
    slStartDate = slSdfDate
    slEndDate = slSdfDate
    ilDay = gWeekDayLong(gDateValue(slSdfDate))
    For ilLoop = 0 To 6 Step 1
        ilCffDay(ilLoop) = 0
    Next ilLoop
    If (gDateValue(slSdfDate) >= gDateValue(slLstStartDate)) And (gDateValue(slSdfDate) <= gDateValue(slLstEndDate)) Then
        'If lst_rst!lstSpotsWk > 0 Then
        If tlAstInfo.iLstSpotsWk > 0 Then
            ilCffDay(0) = tlAstInfo.iLstMon  'lst_rst!lstMon
            ilCffDay(1) = tlAstInfo.iLstTue  'lst_rst!lstTue
            ilCffDay(2) = tlAstInfo.iLstWed  'lst_rst!lstWed
            ilCffDay(3) = tlAstInfo.iLstThu  'lst_rst!lstThu
            ilCffDay(4) = tlAstInfo.iLstFri  'lst_rst!lstFri
            ilCffDay(5) = tlAstInfo.iLstSat  'lst_rst!lstSat
            ilCffDay(6) = tlAstInfo.iLstSun  'lst_rst!lstSun
        Else
            Select Case ilDay
                Case 0
                    ilCffDay(0) = tlAstInfo.iLstMon  'lst_rst!lstMon
                Case 1
                    ilCffDay(1) = tlAstInfo.iLstTue  'lst_rst!lstTue
                Case 2
                    ilCffDay(2) = tlAstInfo.iLstWed  'lst_rst!lstWed
                Case 3
                    ilCffDay(3) = tlAstInfo.iLstThu  'lst_rst!lstThu
                Case 4
                    ilCffDay(4) = tlAstInfo.iLstFri  'lst_rst!lstFri
                Case 5
                    ilCffDay(5) = tlAstInfo.iLstSat  'lst_rst!lstSat
                Case 6
                    ilCffDay(6) = tlAstInfo.iLstSun  'lst_rst!lstSun
            End Select
        End If
    Else

        SQLQuery = "SELECT *"
        SQLQuery = SQLQuery + " FROM cff_Contract_Flight"
        SQLQuery = SQLQuery + " WHERE (cffChfCode = " & llChfCode
        SQLQuery = SQLQuery + " AND cffClfLine = " & ilLineNo
        SQLQuery = SQLQuery + " AND cffCntRevNo = " & tmClfInfo.iCntRevNo 'clf_rst!clfCntRevNo
        SQLQuery = SQLQuery + " AND cffPropVer = " & tmClfInfo.iPropVer   'clf_rst!clfPropVer
        'SQLQuery = SQLQuery + " AND cffStartDate >= '" & Format(slSdfDate, sgSQLDateForm) & "')"
        SQLQuery = SQLQuery + " AND cffStartDate >= '" & Format("12/28/2009", sgSQLDateForm) & "')"
        Set cff_rst = gSQLSelectCall(SQLQuery)

        If Not cff_rst.EOF Then
            Do While Not cff_rst.EOF
                slCffStartDate = Format(cff_rst!cffStartDate, sgShowDateForm)
                slCffEndDate = Format(cff_rst!cffEndDate, sgShowDateForm)
                If (gDateValue(slSdfDate) >= gDateValue(slCffStartDate)) And (gDateValue(slSdfDate) <= gDateValue(slCffEndDate)) Then
                    If cff_rst!cffSpotsWk > 0 Then
                        ilCffDay(0) = cff_rst!cffMo
                        ilCffDay(1) = cff_rst!cffTu
                        ilCffDay(2) = cff_rst!cffWe
                        ilCffDay(3) = cff_rst!cffTh
                        ilCffDay(4) = cff_rst!cffFr
                        ilCffDay(5) = cff_rst!cffSa
                        ilCffDay(6) = cff_rst!cffSu
                    Else
                        Select Case ilDay
                            Case 0
                                ilCffDay(0) = cff_rst!cffMo
                            Case 1
                                ilCffDay(1) = cff_rst!cffTu
                            Case 2
                                ilCffDay(2) = cff_rst!cffWe
                            Case 3
                                ilCffDay(3) = cff_rst!cffTh
                            Case 4
                                ilCffDay(4) = cff_rst!cffFr
                            Case 5
                                ilCffDay(5) = cff_rst!cffSa
                            Case 6
                                ilCffDay(6) = cff_rst!cffSu
                        End Select
                    End If
                    Exit Do
                End If
                cff_rst.MoveNext
            Loop
        End If
    End If
    If ilCffDay(ilDay) > 0 Then
        For ilLoop = ilDay - 1 To 0 Step -1
            If ilCffDay(ilLoop) > 0 Then
                slStartDate = DateAdd("d", -1, slStartDate)
            Else
                Exit For
            End If
        Next ilLoop
        ilDay = gWeekDayLong(gDateValue(slSdfDate))
        For ilLoop = ilDay + 1 To 6 Step 1
            If ilCffDay(ilLoop) > 0 Then
                slEndDate = DateAdd("d", 1, slEndDate)
            Else
                Exit For
            End If
        Next ilLoop
        If gDateValue(slStartDate) < gDateValue(slClfStartDate) Then
            slStartDate = slClfStartDate
        End If
        If gDateValue(slEndDate) > gDateValue(slClfEndDate) Then
            slEndDate = slClfEndDate
        End If
    Else
        slStartDate = slPledgeDate  'Format(ast_rst!astPledgeDate, sgShowDateForm)
        slEndDate = slStartDate
    End If
    slAllowedStartDate = slStartDate
    slAllowedEndDate = slEndDate
    If Second(slStartTime) = 0 Then
        slAllowedStartTime = Format(slStartTime, sgShowTimeWOSecForm)
    Else
        slAllowedStartTime = slStartTime
    End If
    If Second(slEndTime) = 0 Then
        slAllowedEndTime = Format(slEndTime, sgShowTimeWOSecForm)
    Else
        slAllowedEndTime = slEndTime
    End If
    For ilLoop = 0 To 6 Step 1
        ilAllowedDays(ilLoop) = ilCffDay(ilLoop)
    Next ilLoop
    'If Not Aired, then compliant is false
    'If tgStatusTypes(gGetAirStatus(ast_rst!astStatus)).iPledged = 2 Then
    If tgStatusTypes(gGetAirStatus(tlAstInfo.iStatus)).iPledged = 2 Then
        ilCompliant = False
    Else
        ilCompliant = True
        For ilLoop = 0 To 6 Step 1
            ilCffDay(ilLoop) = 0
        Next ilLoop
        ilDay = gWeekDayLong(gDateValue(slAirDate))
        If (gDateValue(slAirDate) >= gDateValue(slLstStartDate)) And (gDateValue(slAirDate) <= gDateValue(slLstEndDate)) Then
            'If lst_rst!lstSpotsWk > 0 Then
            If tlAstInfo.iLstSpotsWk > 0 Then
                ilCffDay(0) = tlAstInfo.iLstMon  'lst_rst!lstMon
                ilCffDay(1) = tlAstInfo.iLstTue  'lst_rst!lstTue
                ilCffDay(2) = tlAstInfo.iLstWed  'lst_rst!lstWed
                ilCffDay(3) = tlAstInfo.iLstThu  'lst_rst!lstThu
                ilCffDay(4) = tlAstInfo.iLstFri  'lst_rst!lstFri
                ilCffDay(5) = tlAstInfo.iLstSat  'lst_rst!lstSat
                ilCffDay(6) = tlAstInfo.iLstSun  'lst_rst!lstSun
            Else
                Select Case ilDay
                    Case 0
                        ilCffDay(0) = tlAstInfo.iLstMon  'lst_rst!lstMon
                    Case 1
                        ilCffDay(1) = tlAstInfo.iLstTue  'lst_rst!lstTue
                    Case 2
                        ilCffDay(2) = tlAstInfo.iLstWed  'lst_rst!lstWed
                    Case 3
                        ilCffDay(3) = tlAstInfo.iLstThu  'lst_rst!lstThu
                    Case 4
                        ilCffDay(4) = tlAstInfo.iLstFri  'lst_rst!lstFri
                    Case 5
                        ilCffDay(5) = tlAstInfo.iLstSat  'lst_rst!lstSat
                    Case 6
                        ilCffDay(6) = tlAstInfo.iLstSun  'lst_rst!lstSun
                End Select
            End If
        Else
            SQLQuery = "SELECT *"
            SQLQuery = SQLQuery + " FROM cff_Contract_Flight"
            SQLQuery = SQLQuery + " WHERE (cffChfCode = " & llChfCode
            SQLQuery = SQLQuery + " AND cffClfLine = " & ilLineNo
            SQLQuery = SQLQuery + " AND cffCntRevNo = " & tmClfInfo.iCntRevNo 'clf_rst!clfCntRevNo
            SQLQuery = SQLQuery + " AND cffPropVer = " & tmClfInfo.iPropVer   'clf_rst!clfPropVer
            'SQLQuery = SQLQuery + " AND cffStartDate >= '" & Format(slAirDate, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery + " AND cffStartDate >= '" & Format("12/28/2009", sgSQLDateForm) & "')"
            Set cff_rst = gSQLSelectCall(SQLQuery)
            If Not cff_rst.EOF Then
                Do While Not cff_rst.EOF
                    slCffStartDate = Format(cff_rst!cffStartDate, sgShowDateForm)
                    slCffEndDate = Format(cff_rst!cffEndDate, sgShowDateForm)
                    If (gDateValue(slAirDate) >= gDateValue(slCffStartDate)) And (gDateValue(slAirDate) <= gDateValue(slCffEndDate)) Then
                        If cff_rst!cffSpotsWk > 0 Then
                            ilCffDay(0) = cff_rst!cffMo
                            ilCffDay(1) = cff_rst!cffTu
                            ilCffDay(2) = cff_rst!cffWe
                            ilCffDay(3) = cff_rst!cffTh
                            ilCffDay(4) = cff_rst!cffFr
                            ilCffDay(5) = cff_rst!cffSa
                            ilCffDay(6) = cff_rst!cffSu
                        Else
                            Select Case ilDay
                                Case 0
                                    ilCffDay(0) = cff_rst!cffMo
                                Case 1
                                    ilCffDay(1) = cff_rst!cffTu
                                Case 2
                                    ilCffDay(2) = cff_rst!cffWe
                                Case 3
                                    ilCffDay(3) = cff_rst!cffTh
                                Case 4
                                    ilCffDay(4) = cff_rst!cffFr
                                Case 5
                                    ilCffDay(5) = cff_rst!cffSa
                                Case 6
                                    ilCffDay(6) = cff_rst!cffSu
                            End Select
                        End If
                        Exit Do
                    End If
                    cff_rst.MoveNext
                Loop
            End If
        End If
        'MG, Outside and fill treated compliant
        'If (slSdfSchStatus <> "G") And (slSdfSchStatus <> "O") And (slSdfSpotType <> "X") Then
        If (slSdfSpotType <> "X") Then
            If ilCffDay(ilDay) <= 0 Then
                ilCompliant = False
            End If
            If (gTimeToLong(slAirTime, False) < gTimeToLong(slStartTime, False)) Or (gTimeToLong(slAirTime, True) > gTimeToLong(slEndTime, True)) Then
                ilCompliant = False
            End If
        End If
    End If
    If (ilCompliant) And (slSdfSpotType <> "X") Then
        If ((slSdfSchStatus = "G") Or (slSdfSchStatus = "O")) Then
            'Test air week against sell week
            SQLQuery = "SELECT smfMissedDate"
            SQLQuery = SQLQuery + " FROM smf_Spot_MG_Specs"
            SQLQuery = SQLQuery + " WHERE (smfSdfCode = " & llSdfCode & ")"
            Set smf_rst = gSQLSelectCall(SQLQuery)
            If Not smf_rst.EOF Then
                slMissedDate = Format(smf_rst!smfMissedDate, sgShowDateForm)
                If gDateValue(gObtainPrevMonday(slMissedDate)) <> gDateValue(gObtainPrevMonday(slAirDate)) Then
                    If (tmClfInfo.sBillCycle <> "W") And (tmClfInfo.sBillCycle <> "C") Then
                        If gDateValue(gObtainEndStd(slMissedDate)) <> gDateValue(gObtainEndStd(slAirDate)) Then
                            ilCompliant = False
                        End If
                    Else
                        ilCompliant = False
                    End If
                End If
            End If
        Else
            If (ilCompliant) Then
                If gDateValue(gObtainPrevMonday(slSdfDate)) <> gDateValue(gObtainPrevMonday(slAirDate)) Then
                    If (tmClfInfo.sBillCycle <> "W") And (tmClfInfo.sBillCycle <> "C") Then
                        If gDateValue(gObtainEndStd(slSdfDate)) <> gDateValue(gObtainEndStd(slAirDate)) Then
                            ilCompliant = False
                        End If
                    Else
                        ilCompliant = False
                    End If
                End If
            End If
        End If
    End If
    On Error Resume Next
    'ast_rst.Close
    'lst_rst.Close
    sdf_rst.Close
    smf_rst.Close
    chf_rst.Close
    clf_rst.Close
    cff_rst.Close
    'rdf_rst.Close
    gGetAgyCompliant = 0
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "gGetAgyCompliant"
    gGetAgyCompliant = 8
    Resume Next
    Exit Function
PopDaypart:
    gPopDaypart
    Resume Next
VefPop:
    gPopVehicles
    Resume Next
ShttPop:
    gPopShttInfo
    Resume Next
End Function

Public Function gPopAll() As Integer

    Dim ilRet As Integer

    ilRet = gPopCpf()
    ilRet = gPopSalesPeopleInfo()
    'Note:  Market and Territory populates must be prior to Station Populate
    ilRet = gPopMarkets()
    ilRet = gPopMSAMarkets()         'MSA markets
    ilRet = gPopMntInfo("T", tgTerritoryInfo())
    ilRet = gPopMntInfo("C", tgCityInfo())
    ilRet = gPopOwnerNames()
    ilRet = gPopFormats() 'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness - Moved above gPopStations, as tgStation now contains the owner name
    ilRet = gPopStations()
    ilRet = gPopVehicleOptions()
    ilRet = gPopVehicles()
    ilRet = gPopSellingVehicles()
    ilRet = gPopAdvertisers()
    ilRet = gPopReportNames()
    ilRet = gGetLatestRatecard()
    ilRet = gPopTimeZones()
    ilRet = gPopStates()
    'ilRet = gPopFormats()
    ilRet = gPopAvailNames()
    ilRet = gPopMediaCodes()
    ilRet = gPopVff()
    '6191 add agency
    ilRet = gPopAgencies()
End Function

Public Function gHashString(slHashString) As Long
    Dim llHash As Long
    Dim blHash() As Byte

    blHash() = slHashString
    llHash = mCalcCRC32(blHash)
    gHashString = llHash And &H7FFFFFFF

End Function

Private Function mCalcCRC32(ByteArray() As Byte) As Long
    Dim i As Long
    Dim j As Long
    Dim Limit As Long
    Dim CRC As Long
    Dim Temp1 As Long
    Dim Temp2 As Long
    Dim CRCTable(0 To 255) As Long
  
    Limit = &HEDB88320
    For i = 0 To 255
        CRC = i
        For j = 8 To 1 Step -1
            If CRC < 0 Then
                Temp1 = CRC And &H7FFFFFFF
                Temp1 = Temp1 \ 2
                Temp1 = Temp1 Or &H40000000
            Else
                Temp1 = CRC \ 2
            End If
            If CRC And 1 Then
                CRC = Temp1 Xor Limit
            Else
                CRC = Temp1
            End If
        Next j
        CRCTable(i) = CRC
    Next i
    Limit = UBound(ByteArray)
    CRC = -1
    For i = 0 To Limit
        If CRC < 0 Then
            Temp1 = CRC And &H7FFFFFFF
            Temp1 = Temp1 \ 256
            Temp1 = (Temp1 Or &H800000) And &HFFFFFF
        Else
            Temp1 = (CRC \ 256) And &HFFFFFF
        End If
        Temp2 = ByteArray(i)   ' get the byte
        Temp2 = CRCTable((CRC Xor Temp2) And &HFF)
        CRC = Temp1 Xor Temp2
    Next i
    CRC = CRC Xor &HFFFFFFFF
    mCalcCRC32 = CRC
End Function

Public Sub gGetEventTitles(ilVefCode As Integer, slEventTitle1 As String, slEventTitle2 As String)
    Dim ilRet As Integer 'btrieve status
    Dim vtf_rst As ADODB.Recordset
    Dim saf_rst As ADODB.Recordset

    On Error GoTo ErrHand
    slEventTitle1 = ""
    slEventTitle2 = ""
    SQLQuery = "Select * from VTF_Vehicle_Text Where vtfVefCode = " & ilVefCode & " And " & "vtfType = 1"
    Set vtf_rst = gSQLSelectCall(SQLQuery)
    If Not vtf_rst.EOF Then
        slEventTitle1 = gStripChr0(vtf_rst!vtfText)
    End If
    SQLQuery = "Select * from VTF_Vehicle_Text Where vtfVefCode = " & ilVefCode & " And " & "vtfType = 2"
    Set vtf_rst = gSQLSelectCall(SQLQuery)
    If Not vtf_rst.EOF Then
        slEventTitle2 = gStripChr0(vtf_rst!vtfText)
    End If
    
    If slEventTitle1 = "" Then
        SQLQuery = "Select safEventTitle1 From SAF_Schd_Attributes WHERE safVefCode = 0"
        Set saf_rst = gSQLSelectCall(SQLQuery)
        If Not saf_rst.EOF Then
            If Trim$(saf_rst!safEventTitle1) <> "" Then
                slEventTitle1 = Trim$(saf_rst!safEventTitle1)
            Else
                slEventTitle1 = "Visiting Team"
            End If
        Else
            slEventTitle1 = "Visiting Team"
        End If
    End If
    If slEventTitle2 = "" Then
        SQLQuery = "Select safEventTitle2 From SAF_Schd_Attributes WHERE safVefCode = 0"
        Set saf_rst = gSQLSelectCall(SQLQuery)
        If Not saf_rst.EOF Then
            If Trim$(saf_rst!safEventTitle2) <> "" Then
                slEventTitle2 = Trim$(saf_rst!safEventTitle2)
            Else
                slEventTitle2 = "Home Team"
            End If
        Else
            slEventTitle2 = "Home Team"
        End If
    End If
    On Error Resume Next
    vtf_rst.Close
    saf_rst.Close
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "gGetEventTitles"
    slEventTitle1 = "Visiting Team"
    slEventTitle2 = "Home Team"
    On Error Resume Next
    vtf_rst.Close
    saf_rst.Close
    Exit Sub
End Sub

Public Function gMapDays(slInDays As String) As String
    Dim slDays As String
    Dim ilDay As Integer
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim slStr As String
    
    slDays = Trim$(slInDays)
    If (InStr(1, slDays, "Y", vbTextCompare) > 0) Or (InStr(1, slDays, "N", vbTextCompare) > 0) Then
        slStr = ""
        ilDay = 1
        Do
            If Mid(slDays, ilDay, 1) = "Y" Then
                ilStart = ilDay
                ilEnd = ilStart
                ilDay = ilDay + 1
                Do
                    If ilDay > 7 Then
                        Exit Do
                    End If
                    If Mid(slDays, ilDay, 1) = "N" Then
                        Exit Do
                    Else
                        ilEnd = ilDay
                    End If
                    ilDay = ilDay + 1
                Loop
                If slStr = "" Then
                    If ilStart = ilEnd Then
                        slStr = Switch(ilStart = 1, "M", ilStart = 2, "Tu", ilStart = 3, "W", ilStart = 4, "Th", ilStart = 5, "F", ilStart = 6, "Sa", ilStart = 7, "Su")
                    Else
                        slStr = Switch(ilStart = 1, "M", ilStart = 2, "Tu", ilStart = 3, "W", ilStart = 4, "Th", ilStart = 5, "F", ilStart = 6, "Sa", ilStart = 7, "Su")
                        slStr = slStr & "-" & Switch(ilEnd = 1, "M", ilEnd = 2, "Tu", ilEnd = 3, "W", ilEnd = 4, "Th", ilEnd = 5, "F", ilEnd = 6, "Sa", ilEnd = 7, "Su")
                    End If
                Else
                    If ilStart = ilEnd Then
                        slStr = slStr & "," & Switch(ilStart = 1, "M", ilStart = 2, "Tu", ilStart = 3, "W", ilStart = 4, "Th", ilStart = 5, "F", ilStart = 6, "Sa", ilStart = 7, "Su")
                    Else
                        slStr = slStr & "," & Switch(ilStart = 1, "M", ilStart = 2, "Tu", ilStart = 3, "W", ilStart = 4, "Th", ilStart = 5, "F", ilStart = 6, "Sa", ilStart = 7, "Su")
                        slStr = slStr & "-" & Switch(ilEnd = 1, "M", ilEnd = 2, "Tu", ilEnd = 3, "W", ilEnd = 4, "Th", ilEnd = 5, "F", ilEnd = 6, "Sa", ilEnd = 7, "Su")
                    End If
                End If
            End If
            ilDay = ilDay + 1
        Loop While ilDay <= 7
        slDays = slStr
    Else
        If (slDays = "MoTuWeThFrSaSu") Then
            slDays = "M-Su"
        ElseIf (slDays = "MoTuWeThFrSa") Then
            slDays = "M-Sa"
        ElseIf (slDays = "MoTuWeThFr") Then
            slDays = "M-F"
        ElseIf (slDays = "MoTuWeTh") Then
            slDays = "M-Th"
        ElseIf (slDays = "MoTuWe") Then
            slDays = "M-W"
        ElseIf slDays = ("MoTu") Then
            slDays = "M-Tu"
        ElseIf (slDays = "TuWeThFrSaSu") Then
            slDays = "Tu-Su"
        ElseIf (slDays = "TuWeThFrSa") Then
            slDays = "Tu-Sa"
        ElseIf (slDays = "TuWeThFr") Then
            slDays = "Tu-F"
        ElseIf (slDays = "TuWeTh") Then
            slDays = "Tu-Th"
        ElseIf (slDays = "TuWe") Then
            slDays = "Tu-W"
        ElseIf (slDays = "WeThFrSaSu") Then
            slDays = "W-Su"
        ElseIf (slDays = "WeThFrSa") Then
            slDays = "W-Sa"
        ElseIf (slDays = "WeThFr") Then
            slDays = "W-F"
        ElseIf (slDays = "WeTh") Then
            slDays = "W-Th"
        ElseIf (slDays = "ThFrSaSu") Then
            slDays = "Th-Su"
        ElseIf (slDays = "ThFrSa") Then
            slDays = "Th-Sa"
        ElseIf (slDays = "ThFr") Then
            slDays = "Th-F"
        ElseIf slDays = "FrSaSu" Then
            slDays = "F-Su"
        ElseIf slDays = "FrSa" Then
            slDays = "F-Sa"
        ElseIf slDays = "SaSu" Then
            slDays = "S-S"
        End If
    End If
    gMapDays = slDays
End Function

Public Sub gUpdateTaskMonitor(ilWhichTime As Integer, slTaskCode As String)
    'ilWhichTime: 0=Running; 1=Start, 2= End
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim slSQLQuery As String
    Dim slStr As String
    Dim slZone As String
    Dim llDate As Long
    
    On Error GoTo ErrHand
    If sgTimeZone = "" Then
        sgTimeZone = Left$(gGetLocalTZName(), 1)
    End If
    slDateTime = Now
    slNowDate = Format$(slDateTime, "m/d/yy")
    slNowTime = Format$(slDateTime, "h:mm:ssAM/PM")
    Select Case sgTimeZone
        Case "E"
            slStr = slNowDate & " " & slNowTime
        Case "C"
            slStr = DateAdd("h", 1, slNowDate & " " & slNowTime)
        Case "M"
            slStr = DateAdd("h", 2, slNowDate & " " & slNowTime)
        Case "P"
            slStr = DateAdd("h", 3, slNowDate & " " & slNowTime)
        Case Else
            slStr = slNowDate & " " & slNowTime
    End Select
    slNowDate = Format$(slStr, "m/d/yy")
    slNowTime = Format$(slStr, "h:mm:ssAM/PM")
    '12/9/15: Handle case where the TMF not setup
    slSQLQuery = "SELECT tmf1stStartRunDate, tmf1stEndRunDate FROM TMF_Task_Monitor WHERE (tmfTaskCode = '" & slTaskCode & "'" & ")"
    Set tmf_rst = gSQLSelectCall(slSQLQuery)
    If tmf_rst.EOF Then
        Exit Sub
    End If
    If ilWhichTime = 1 Then
        sgTmfStatus = "S"
        slSQLQuery = "UPDATE tmf_Task_Monitor SET "
        slSQLQuery = slSQLQuery & "tmfRunningDate = '" & Format$(slNowDate, sgSQLDateForm) & "', "
        slSQLQuery = slSQLQuery & "tmfRunningTime = '" & Format$(slNowTime, sgSQLTimeForm) & "', "
        If IsNull(tmf_rst!tmf1stStartRunDate) Then
            llDate = -1
        Else
            llDate = gDateValue(gAdjYear(tmf_rst!tmf1stStartRunDate))
        End If
        If gDateValue(gAdjYear(slNowDate)) <> llDate Then
            slSQLQuery = slSQLQuery & "tmf1stStartRunDate = '" & Format$(slNowDate, sgSQLDateForm) & "', "
            slSQLQuery = slSQLQuery & "tmf1stStartRunTime = '" & Format$(slNowTime, sgSQLTimeForm) & "', "
        End If
        slSQLQuery = slSQLQuery & "tmfStartRunDate = '" & Format$(slNowDate, sgSQLDateForm) & "', "
        slSQLQuery = slSQLQuery & "tmfStartRunTime = '" & Format$(slNowTime, sgSQLTimeForm) & "', "
        slSQLQuery = slSQLQuery & "tmfStatus = '" & "S" & "' "
        slSQLQuery = slSQLQuery & "WHERE tmfTaskCode = '" & slTaskCode & "'"
    ElseIf ilWhichTime = 2 Then
        If sgTmfStatus = "S" Then
            sgTmfStatus = "C"
        End If
        slSQLQuery = "UPDATE tmf_Task_Monitor SET "
        slSQLQuery = slSQLQuery & "tmfRunningDate = '" & Format$(slNowDate, sgSQLDateForm) & "', "
        slSQLQuery = slSQLQuery & "tmfRunningTime = '" & Format$(slNowTime, sgSQLTimeForm) & "', "
        If IsNull(tmf_rst!tmf1stEndRunDate) Then
            llDate = -1
        Else
            llDate = gDateValue(gAdjYear(tmf_rst!tmf1stEndRunDate))
        End If
        If gDateValue(gAdjYear(slNowDate)) <> llDate Then
            slSQLQuery = slSQLQuery & "tmf1stEndRunDate = '" & Format$(slNowDate, sgSQLDateForm) & "', "
            slSQLQuery = slSQLQuery & "tmf1stEndRunTime = '" & Format$(slNowTime, sgSQLTimeForm) & "', "
        End If
        slSQLQuery = slSQLQuery & "tmfEndRunDate = '" & Format$(slNowDate, sgSQLDateForm) & "', "
        slSQLQuery = slSQLQuery & "tmfEndRunTime = '" & Format$(slNowTime, sgSQLTimeForm) & "', "
        slSQLQuery = slSQLQuery & "tmfStatus = '" & sgTmfStatus & "' "
        slSQLQuery = slSQLQuery & "WHERE tmfTaskCode = '" & slTaskCode & "'"
    Else
        slSQLQuery = "UPDATE tmf_Task_Monitor SET "
        slSQLQuery = slSQLQuery & "tmfRunningDate = '" & Format$(slNowDate, sgSQLDateForm) & "', "
        slSQLQuery = slSQLQuery & "tmfRunningTime = '" & Format$(slNowTime, sgSQLTimeForm) & "' "
        slSQLQuery = slSQLQuery & "WHERE tmfTaskCode = '" & slTaskCode & "'"
    End If
    'cnn.Execute slSQL_AlertClear, rdExecDirect
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "AffErrorLog.txt", "modGenSubs-gUpdateTaskMonitor"
        Exit Sub
    End If
    On Error GoTo ErrHand
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "gUpdateTaskMonitor"
    Exit Sub
ErrHand1:
    gHandleError "AffErrorLog.txt", "gUpdateTaskMonitor"
    Return
End Sub
'7458 dan 3/29/17 added 'long' to llReturnEntCode
Public Function gENTAddNew(tlENT As ENT, Optional slErrorLog As String = "AffErrorLog.txt", Optional llReturnEntCode As Long = -1) As Boolean
'Output:  returns whether successful or not.  if llReturnCode <> -1, then return the new EntCode
    
    SQLQuery = "Insert Into ent ( "
    If llReturnEntCode <> -1 Then
        SQLQuery = SQLQuery & "entCode, "
    End If
    SQLQuery = SQLQuery & "entType, "
    SQLQuery = SQLQuery & "ent3rdParty, "
    SQLQuery = SQLQuery & "entAttCode, "
    SQLQuery = SQLQuery & "entShttCode, "
    SQLQuery = SQLQuery & "entVefCode, "
    SQLQuery = SQLQuery & "entFeedDate, "
    SQLQuery = SQLQuery & "entGsfCode, "
    SQLQuery = SQLQuery & "entAstCount, "
    SQLQuery = SQLQuery & "entSpotCount, "
    SQLQuery = SQLQuery & "entMGCount, "
    SQLQuery = SQLQuery & "entReplaceCount, "
    SQLQuery = SQLQuery & "entBonusCount, "
    SQLQuery = SQLQuery & "entIngestedCount, "
    SQLQuery = SQLQuery & "entEnteredDate, "
    SQLQuery = SQLQuery & "entEnteredTime, "
    SQLQuery = SQLQuery & "entFileName, "
    SQLQuery = SQLQuery & "entStatus, "
    '7509 & 7510 2 new fields-delete and errorMsg
    SQLQuery = SQLQuery & "entUstCode, "
    SQLQuery = SQLQuery & "entDeleteCount, entErrorMsg"
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    If llReturnEntCode <> -1 Then
        SQLQuery = SQLQuery & "ReplaceMe, "
    End If
    SQLQuery = SQLQuery & "'" & gFixQuote(tlENT.sType) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlENT.s3rdParty) & "', "
    SQLQuery = SQLQuery & tlENT.lAttCode & ", "
    SQLQuery = SQLQuery & tlENT.iShttCode & ", "
    SQLQuery = SQLQuery & tlENT.iVefCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlENT.sFeedDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & tlENT.lgsfCode & ", "
    SQLQuery = SQLQuery & tlENT.iAstCount & ", "
    SQLQuery = SQLQuery & tlENT.iSpotCount & ", "
    SQLQuery = SQLQuery & tlENT.iMGCount & ", "
    SQLQuery = SQLQuery & tlENT.iReplaceCount & ", "
    SQLQuery = SQLQuery & tlENT.iBonusCount & ", "
    SQLQuery = SQLQuery & tlENT.iIngestedCount & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlENT.sEnteredDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlENT.sEnteredTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlENT.sFileName) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlENT.sStatus) & "', "
    '7509 & 7510 2 new fields-delete and errorMsg
    SQLQuery = SQLQuery & tlENT.iUstCode & ", " & tlENT.iDeleteCount & " , '" & gFixQuote(tlENT.sErrorMsg) & "'"
    SQLQuery = SQLQuery & ") "

    If llReturnEntCode <> -1 Then
        llReturnEntCode = gInsertAndReturnCode(SQLQuery, "ENT", "EntCode", "ReplaceMe")
    Else
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/13/16: Replaced GoSub
            'GoSub ErrHand:
            gHandleError slErrorLog, "gENTAddNew"
            gENTAddNew = False
            Exit Function
        End If
    End If
    gENTAddNew = True
    Exit Function
ErrHand:
    gHandleError slErrorLog, "gENTAddNew"
    gENTAddNew = False
End Function
Public Function gSiteISCIAndOrBreak() As Integer
'0 is problem. 1 is isci, 2 is break and 3 is both
    Dim ilRet As Integer
    Dim ilValue7 As Integer
    Dim ilValue8 As Integer
    
    SQLQuery = "Select spfUsingFeatures7,spfUsingFeatures8 From SPF_Site_Options"
On Error GoTo ErrHand
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        ilValue7 = Asc(rst!spfUsingFeatures7)
        If (ilValue7 And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT Then
            ilRet = 1
        End If
        ilValue8 = Asc(rst!spfUsingFeatures8)
        If (ilValue8 And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT Then
            ilRet = ilRet + 2
        End If
    End If
    gSiteISCIAndOrBreak = ilRet
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "gSiteISCIAndOrBreak"
    gSiteISCIAndOrBreak = 0
End Function

Public Sub gSetRptGGFlag(safName As String)
    Dim slStr As String
    Dim llDate As Long
    Dim ilRptDays As Integer
    Dim llNow As Long
    
    If igGGFlag = 0 Then
        igRptGGFlag = 0
        If Len(Trim$(safName)) > 16 Then
            slStr = Mid$(safName, 2, 5)
            llDate = Val(slStr)
            slStr = Mid$(safName, 17)
            ilRptDays = Val(slStr)
            llNow = gDateValue(Format$(gNow(), "m/d/yy"))
            If llDate + ilRptDays > llNow Then
                igRptGGFlag = 1
            End If
        End If
    Else
        igRptGGFlag = 1
    End If
End Sub

Public Function IsDevEnv() As Boolean

    Dim strFileName$
    Dim lngCount&
    
    On Error Resume Next
    strFileName = String(255, 0)
    lngCount = GetModuleFileName(App.hInstance, strFileName, 255&)
    strFileName = Left(strFileName, lngCount)
    
    IsDevEnv = (UCase(right(strFileName, 7)) Like "VB?.EXE")

End Function


'Private Sub mErrHand(iRet As Integer, iDoTrans As Integer)
'
'    For Each gErrSQL In cnn.Errors
'        iRet = gErrSQL.NativeError
'        If iRet < 0 Then
'            iRet = iRet + 4999
'        End If
'        'If (iRet = 84) And (iDoTrans) Then              'SQLSetConnectAttr vs. SQLSetOpenConnection
'        If (iRet = BTRV_ERR_REC_LOCKED) Or (iRet = BTRV_ERR_FILE_LOCKED) Or (iRet = BTRV_ERR_INCOM_LOCK) Or (iRet = BTRV_ERR_CONFLICT) Then
'            If iDoTrans Then
'                cnn.RollbackTrans
'            End If
'            cnn.Errors.Clear
'            'ResumeNext
'            Exit Sub
'        End If
'        If iRet <> 0 Then              'SQLSetConnectAttr vs. SQLSetOpenConnection
'            gMsgBox "A SQL error has occurred: " & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, vbCritical
'        End If
'    Next gErrSQL
'    If iDoTrans Then
'        cnn.RollbackTrans
'    End If
'    cnn.Errors.Clear
'    'ResumeNext
'    Exit Sub
'End Sub

Public Function gGrid_RowSearch(grdCtrl As MSHFlexGrid, ilColumnNumber As Integer, slSearchValue As String) As Long
        
          
            Dim ilRowIndex As Integer
            Dim llRow As Long
            Dim slStr As String
            
            slStr = UCase(slSearchValue)
            
            For llRow = grdCtrl.FixedRows To grdCtrl.Rows - 1 Step 1
                If InStr(1, UCase(grdCtrl.TextMatrix(llRow, ilColumnNumber)), slStr, vbBinaryCompare) = 1 Then
                   gGrid_RowSearch = llRow
                   Exit Function
                End If
                
            Next llRow
            gGrid_RowSearch = -1
End Function


Public Sub gRemoveFiles(slFolder As String, slInFileName As String, iNumberOfDaysToRetained As Integer)
    Dim objFSO As New FileSystemObject
    Dim objFolder As Folder
    Dim objFile As file
    Dim slCreatedDate As String
    Dim slFileName As String
    On Error Resume Next
    Set objFolder = objFSO.GetFolder(slFolder)
    slFileName = UCase$(slInFileName)
    'Iterate through subfolders.
    For Each objFile In objFolder.Files
        If InStr(1, UCase$(objFile.Name), slFileName) = 1 Then
            slCreatedDate = objFile.DateCreated
            If DateDiff("d", slCreatedDate, Now()) > iNumberOfDaysToRetained Then
                objFile.Delete
            End If
        End If
    Next objFile
End Sub
Public Function gBinarySearchListCtrl(ListCtrl As control, slInMatchString As String) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim llResult As Long
    Dim slMatchString As String
    
    slMatchString = UCase(Trim$(slInMatchString))
    llMin = 0
    llMax = ListCtrl.ListCount - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        llResult = StrComp(UCase(Trim(ListCtrl.List(llMiddle))), slMatchString, vbTextCompare)
        Select Case llResult
            Case 0:
                gBinarySearchListCtrl = llMiddle  ' Found it !
                Exit Function
            Case 1:
                llMax = llMiddle - 1
            Case -1:
                llMin = llMiddle + 1
        End Select
    Loop
    gBinarySearchListCtrl = -1
    Exit Function
    
End Function


Public Function gCleanUpFiles(slSDateNewProg As String, llAttCode As Long, ilVefCode As Integer) As Boolean

    Dim slEDateOldProg As String
    Dim slDrop As String
    Dim llOldAttCode As Long
    ReDim tlCpttArray(0 To 0) As CPTTARRAY
    Dim rst_Cptt As ADODB.Recordset
    Dim ilUpper As Integer
    Dim ilIdx As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStr As String
    Dim temp_rst As ADODB.Recordset
    Dim ilRet As Integer
    Dim ilAgreeType As Integer
    
    '01/14/20 TTP 5670
    gCleanUpFiles = False
    slSDateNewProg = Format$(gObtainPrevMonday(slSDateNewProg), sgShowDateForm)
    slEDateOldProg = gAdjYear(Format$(DateValue(slSDateNewProg) - 1, sgShowDateForm))
    llOldAttCode = llAttCode

    slDrop = DateAdd("d", 1, slEDateOldProg)
    SQLQuery = "SELECT cpttStartDate, cpttCode FROM cptt WHERE (cpttAtfCode = " & llOldAttCode & " And "
    SQLQuery = SQLQuery & "cpttStartDate >= '" & Format$(slDrop, sgSQLDateForm) & "'" & ")"
    Set rst_Cptt = gSQLSelectCall(SQLQuery)
    ilUpper = 0
    While Not rst_Cptt.EOF
        tlCpttArray(ilUpper).lCpttCode = rst_Cptt!cpttCode
        tlCpttArray(ilUpper).sCpttStartDate = rst_Cptt!CpttStartDate
        ilUpper = ilUpper + 1
        ReDim Preserve tlCpttArray(0 To ilUpper)
        rst_Cptt.MoveNext
    Wend
    For ilIdx = 0 To ilUpper - 1 Step 1
        SQLQuery = "DELETE FROM Cptt WHERE (cpttCode = " & tlCpttArray(ilIdx).lCpttCode & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjOverlapAgmnts"
            Exit Function
        End If
        slStartDate = tlCpttArray(ilIdx).sCpttStartDate
        slEndDate = DateAdd("d", 6, tlCpttArray(ilIdx).sCpttStartDate)
        'Delete the spots from the web first
        
'D.S. Add code for spot history and spot retrival/revision/archive

        slStr = "Select attExportType, attExportToWeb, attWebInterface from att where attCode = " & llOldAttCode
        Set temp_rst = gSQLSelectCall(slStr)
        If temp_rst.EOF = False Then
            If temp_rst!attExportType = 1 And gHasWebAccess Then
                ilRet = gWebDeleteSpots(llOldAttCode, Format$(slStartDate, sgSQLDateForm), Format$(slEndDate, sgSQLDateForm))
            End If
        End If
        SQLQuery = "DELETE FROM Ast WHERE (astAtfCode = " & llOldAttCode
        SQLQuery = SQLQuery & " AND astFeedDate >= '" & Format$(slStartDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND astFeedDate <= '" & Format$(slEndDate, sgSQLDateForm) & "')"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "AffAgmnt-mAdjOverlapAgmnts"
            Exit Function
        End If
        ilAgreeType = CInt(temp_rst!attExportType)
        '0 = manual, 1 = export
        If (ilAgreeType > 0) Then
            ilRet = gAlertAdd("R", "S", ilVefCode, tlCpttArray(ilIdx).sCpttStartDate)
        End If
    Next ilIdx
    If ilUpper > 0 Then
        rst_Cptt.Close
        temp_rst.Close
        Erase tlCpttArray
    End If
    gCleanUpFiles = True
    Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "gUpdateTaskMonitor"
    Exit Function
End Function


Public Sub gFormResize(frmForm As Form, llOldWidth As Long, llOldHeight As Long, llNewWidth As Long, llNewHeight As Long)
    Dim Ctrl As control
    Dim slName As String
    
    'Example of adding Resize call to forms
    'Remove the resize control from the form
    'Add a timer control. Initially set to enable = false and interval to 1000
    'Add definition:
    'Private lmOldWidth as Long
    'Private lmOldHeight as Long
    'The old width and height manually set because using Me.Width is res1zing the form because of the bounder style set to 2-sizable
    'Change
    'Private Sub Form_Initialize()
        'imFirstTime = True                 'Added
        'lmOldWidth = 9420   'Me.Width      'Added: Obtain setting from the form
        'lmOldHeight = 6840  'Me.Height     'Added: Obtain setting from the form
        'Me.Width = Screen.Width / 1.05
        'Me.Height = Screen.Height / 1.25
        'Me.Top = (Screen.Height - Me.Height) / 1.8
        'Me.Left = (Screen.Width - Me.Width) / 2
        'gFormResize frmPostLog, lmOldWidth, lmOldHeight, Me.Width, Me.Height   'Added
        'gSetFonts frmPostLog
        'gCenterForm frmPostLog
        'lmOldWidth = Me.Width              'Added
        'lmOldHeight = Me.Height            'Added
    'end sub
    
    'Add
    'Private Sub Form_Resize()
    '    If imFirstTime Then Exit Sub
    '    tmcResize.Enabled = False
    '    DoEvents
    '    tmcResize.Enabled = True
    'End Sub
    
    'Add
    'Private Sub tmcResize_Timer()
    '    tmcResize.Enabled = False
    '    If Me.WindowState = 1 Then Exit Sub
    '    gFormResize frmPostLog, lmOldWidth, lmOldHeight, Me.Width, Me.Height
    '    gSetFonts frmPostLog
    '    If Me.WindowState <> 2 Then gCenterForm frmPostLog
    '    lmOldWidth = Me.Width
    '    lmOldHeight = Me.Height
    '    'Reset grid column widths
    '    mSetColumnWidths True, True, True, True
    'End Sub
    
    On Error GoTo IgnoreErr
    For Each Ctrl In frmForm.Controls
        slName = Ctrl.Name
        Ctrl.Left = (Ctrl.Left * llNewWidth) / llOldWidth
        Ctrl.Width = (Ctrl.Width * llNewWidth) / llOldWidth
        Ctrl.Top = (Ctrl.Top * llNewHeight) / llOldHeight
        Ctrl.Height = (Ctrl.Height * llNewHeight) / llOldHeight
    Next Ctrl
    Exit Sub
IgnoreErr:
    Resume Next
End Sub


Public Function gGetUniqueFileName() As String
    
    Dim sUKey As String
    
    'Resolution appears to be about 3 microseconds in a tight loop
    sUKey = Replace(CStr(gStopWatch.ElapsedSeconds), ".", "_")
    gGetUniqueFileName = CurDir & "\TempFile_" & sgUserName & "_" & sUKey & ".txt"

End Function
