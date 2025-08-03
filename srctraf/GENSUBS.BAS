Attribute VB_Name = "GENSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of GENSUBS.BAS on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  gChangeCase                   gFixQuote                                               *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: GenSubs.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the general purpose subs and functions
Option Explicit
Option Compare Text

Private Const MIB_IF_TYPE_ETHERNET                As Long = 6
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH      As Long = 128
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH_p    As Long = MAX_ADAPTER_DESCRIPTION_LENGTH + 4
Private Const MAX_ADAPTER_NAME_LENGTH             As Long = 256
Private Const MAX_ADAPTER_NAME_LENGTH_p           As Long = MAX_ADAPTER_NAME_LENGTH + 4
Private Const MAX_ADAPTER_ADDRESS_LENGTH          As Long = 8

Private Type TIME_t
    aTime As Long
End Type

Private Type IP_ADDRESS_STRING
    IPadrString     As String * 16
End Type

Private Type IP_ADDR_STRING
    AdrNext         As Long
    IpAddress       As IP_ADDRESS_STRING
    IpMask          As IP_ADDRESS_STRING
    NTEcontext      As Long
End Type
' Information structure returned by GetIfEntry/GetIfTable
Private Type IP_ADAPTER_INFO
    Next As Long
    ComboIndex As Long
    AdapterName         As String * MAX_ADAPTER_NAME_LENGTH_p
    Description         As String * MAX_ADAPTER_DESCRIPTION_LENGTH_p
    MACadrLength        As Long
    MACaddress(0 To MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
    AdapterIndex        As Long
    AdapterType         As Long             ' MSDN Docs say "UInt", but is 4 bytes
    DhcpEnabled         As Long             ' MSDN Docs say "UInt", but is 4 bytes
    CurrentIpAddress    As Long
    IpAddressList       As IP_ADDR_STRING
    GatewayList         As IP_ADDR_STRING
    DhcpServer          As IP_ADDR_STRING
    HaveWins            As Long             ' MSDN Docs say "Bool", but is 4 bytes
    PrimaryWinsServer   As IP_ADDR_STRING
    SecondaryWinsServer As IP_ADDR_STRING
    LeaseObtained       As TIME_t
    LeaseExpires        As TIME_t
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any, ByRef Source As Any, ByVal numbytes As Long)
Public Declare Function GetAdaptersInfo Lib "iphlpapi.dll" (ByRef pAdapterInfo As Any, ByRef pOutBufLen As Long) As Long
'5666
Public sgShortDate As String
Public Const MYSHORTDATE As String = "M/d/yy"
Const REG_SZ = 1 ' Unicode nul terminated string
Const REG_BINARY = 3
Const HKEY_CURRENT_USER = &H80000001
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public bgDevEnv As Boolean
Private Declare Function GetModuleFileName Lib _
    "kernel32" Alias "GetModuleFileNameA" (ByVal _
    hModule As Long, ByVal lpFileName As String, _
    ByVal nSize As Long) As Long
    
Private Const xlUnderlineStyleDouble = -4119            'Double thick underline
Private Const xlUnderlineStyleDoubleAccounting = 5      'Two thin underlines placed close together.
Private Const xlUnderlineStyleNone = -4142              'No underlining
Private Const xlUnderlineStyleSingle = 2                'Single underlining.
Private Const xlUnderlineStyleSingleAccounting = 4      'Not supported.
Public ogExcel As Object
Public bgExcelCreated As Boolean

Public lgShellAndWaitID As Long

'*******************************************************
'*                                                     *
'*      Procedure Name:gDbl ToStrDec                   *
'*                                                     *
'*             Created:7/8/08        By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Format Double data type as      *
'*                     money or percentage             *
'*    Duplicate of gLongToSTrDec                       *
'*******************************************************
Function gDblToStrDec(dlInNumber As Double, ilNoDecPlaces As Integer) As String
'
'   slOutStr = gLongToStrDec(dlInNumber, ilNoDecPlaces)
'   Where:
'       dlInNumber(I)- Number to be converted to string with/without dec point
'       ilNoDecPlaces(I)-Number of decimal places (can be zero)
'       slOutStr(O)- Output string format with a decimal point
'
    Dim slStr As String
    Dim ilNegSign As Integer

    If dlInNumber < 0 Then
        ilNegSign = True
    Else
        ilNegSign = False
    End If
    slStr = Trim$(str$(Abs(dlInNumber)))
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
    gDblToStrDec = slStr
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gParseItemFields                *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain a substring from a string*
'*                                                     *
'*******************************************************
Sub gParseItemFields(ByVal slInputStr As String, slDelimiter As String, slFields() As String)
'
'   gParseItemToArray slInputStr, slDelimiter, slFields()
'   Where:
'       slInputStr (I)-string from which to obtain substring
'       slFields() (O)- fields parsed from comma delimited string
'

    Dim ilFieldNo As Integer
    Dim ilFieldType As Integer  '0=String, 1=Number
    Dim slChar As String
    Dim ilIndex As Integer
    Dim ilAddToStr As Integer
    Dim slNextChar As String

    For ilIndex = LBound(slFields) To UBound(slFields) Step 1
        slFields(ilIndex) = ""
    Next ilIndex
    ilFieldNo = LBound(slFields)    '1
    ilIndex = 1
    ilFieldType = -1
    Do While ilIndex <= Len(Trim$(slInputStr))
        slChar = Mid$(slInputStr, ilIndex, 1)
        If ilFieldType = -1 Then
            If slChar = slDelimiter Then
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
                    If ilIndex = Len(Trim$(slInputStr)) Then
                        ilAddToStr = False
                    Else
                        slNextChar = Mid$(slInputStr, ilIndex + 1, 1)
                        If slNextChar = slDelimiter Then
                            ilAddToStr = False
                        End If
                    End If
                End If
                If ilAddToStr Then
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
                If slChar <> slDelimiter Then
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
'*      Procedure Name:mResizeCtrls                    *
'*                                                     *
'*             Created:5/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Resize all controls on a form   *
'*                                                     *
'*******************************************************
'Sub gResizeCtrls (Frm As Form)

    ''Const DESIGN_WIDTH = 640
    ''Const DESIGN_HEIGHT = 480
    '
    'Dim ilWidthStart As Integer
    'Dim ilHeightStart As Integer
    'Dim flWidthFactor  As Single
    'Dim flHeightFactor As Single
    'Dim ilLoop As Integer
    'Dim aControl As Control
    '
    'ilWidthStart = Screen.Width / Screen.TwipsPerPixelX
    'ilHeightStart = Screen.Height / Screen.TwipsPerPixelY
    'flWidthFactor = ilWidthStart / 640  'DESIGN_WIDTH
    'flHeightFactor = ilHeightStart / 480    'DESIGN_HEIGHT
    '
    'For ilLoop = 0 To Frm.Controls.Count - 1 Step 1
    '    Set aControl = Frm.Controls(ilLoop)
    '    aControl.Width = flWidthFactor * aControl.Width
    '    aControl.Left = flWidthFactor * aControl.Left
    '    aControl.Height = flHeightFactor * aControl.Height
    '    aControl.Top = flHeightFactor * aControl.Top
    'Next ilLoop
'End Sub

Sub gCenterStdAlone(frm As Form)
    gSetPictureBoxFontSize frm, 9
    If igTaskLoc = 1 Then
        frm.Move (Screen.Width - frm.Width) \ 2, (Screen.height - frm.height) \ 2 + 115 '+ Screen.Height \ 10
    Else
        frm.Move (Screen.Width - frm.Width) \ 2, (Screen.height - frm.height) \ 2 + 450 '+ Screen.Height \ 10
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gChangeCase                     *
'*                                                     *
'*             Created:5/6/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Convert string to various styles*
'*                                                     *
'*******************************************************
Function gChangeCase(slInputStr As String, ilStyle As Integer) As String 'VBC NR
'
'   slRet = gChangeCase (slInputStr, ilStyle)
'   Where:
'       slInputStr(I)- String to be converted
'       slStyle(I)- 0=convert to upper case; 1= convert to lower case;
'                   2=First letter upper case, remaining letters lower case
'
    Dim ilLoop As Integer 'VBC NR
    Dim slChar As String 'VBC NR
    Dim slStr As String 'VBC NR
    Dim ilAscChar As Integer 'VBC NR

    Select Case ilStyle 'VBC NR
        Case 0  'Upper case 'VBC NR
            gChangeCase = UCase$(slInputStr) 'VBC NR
        Case 1  'Lower case 'VBC NR
            gChangeCase = LCase$(slInputStr) 'VBC NR
        Case 2 'VBC NR
            ilAscChar = Asc(UCase$(" ")) 'VBC NR
            For ilLoop = 1 To Len(slInputStr) Step 1 'VBC NR
                slChar = Mid$(slInputStr, ilLoop, 1) 'VBC NR
                If ((ilAscChar >= Asc("A")) And (ilAscChar <= Asc("Z"))) Then 'VBC NR
                    slStr = slStr & LCase$(slChar) 'VBC NR
                Else 'VBC NR
                    slStr = slStr & UCase$(slChar) 'VBC NR
                End If 'VBC NR
                ilAscChar = Asc(UCase$(slChar)) 'VBC NR
            Next ilLoop 'VBC NR
            gChangeCase = slStr 'VBC NR
    End Select 'VBC NR
    Exit Function 'VBC NR
End Function 'VBC NR

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
Sub gFadeForm(frm As Form, ilRed As Integer, ilGreen As Integer, ilBlue As Integer)
    Dim ilSaveScale As Integer
    Dim ilSaveStyle As Integer
    Dim ilSaveRedraw As Integer

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
    ilPixels = Screen.height / Screen.TwipsPerPixelY
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

'*******************************************************
'*                                                     *
'*      Procedure Name:gFormatStr                      *
'*                                                     *
'*             Created:5/6/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Format as money or percentage   *
'*                                                     *
'*******************************************************
Sub gFormatStr(slInStr As String, ilFmtSpec As Integer, ilNoDecPlaces As Integer, slOutStr As String)
'
'   gFormatStr slStr, ilFmt, ilNoDec, slOut
'   Where:
'       slStr (I)- String to be formatted
'       ilFmt (I)- Format specifications
'                  FMTDEFAULT = 0     'Only adjust for decimal place
'                  FMTCOMMA = 1       'Include commas
'                  FMTDOLLARSIGN = 2  'Include dollar sign
'                  FMTPERCENTSIGN = 4 'Include percent sign
'                  FMTNEGASBRACKETS = 8'Show negative numbers with brackets
'                  FMTNEGATBACK = 16  'Show negative sign as last character
'                  FMTLEAVEBLANK = 32   'Leave blank as blank
'       ilNoDec (I)- Number of decimal places (0 allowed)
'       slOut (O) - String to contain formatted string
'

    Dim ilPos As Integer    'Position of character within string
    Dim slLeft As String    'String to left of decimal
    Dim slRight As String   'String to right of decimal
    Dim ilNoDigits As Integer   'Number of digits after decimal point
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilSignPos As Integer    'Position of sign

    slOutStr = Trim$(slInStr)  'Initialize string
    If (slOutStr = "") And ((ilFmtSpec And FMTLEAVEBLANK) = FMTLEAVEBLANK) Then
        Exit Sub
    End If
    'Adjust decimal places
    If ilNoDecPlaces > 0 Then
        ilPos = InStr(slOutStr, ".")
        If ilPos = 0 Then   'Decimal place missing- Add
            slOutStr = slOutStr & "."
            ilPos = InStr(slOutStr, ".")
        End If
        ilNoDigits = ilNoDecPlaces - (Len(slOutStr) - ilPos)
        If ilNoDigits > 0 Then   'Add zeros
            For ilLoop = 1 To ilNoDigits Step 1
                slOutStr = slOutStr + "0"
            Next ilLoop
        ElseIf ilNoDigits < 0 Then   'Remove digits
            slOutStr = Left$(slOutStr, Len(slOutStr) + ilNoDigits)  'ilNoDigits is a negative number
        End If
    Else
        ilPos = InStr(slOutStr, ".")
        If ilPos > 0 Then   'Remove decimal place
            slOutStr = Left$(slOutStr, ilPos - 1)
        End If
    End If
    'Format negative sign
    ilSignPos = InStr(slOutStr, "-")
    If ilSignPos > 0 Then
        slOutStr = Mid$(slOutStr, 2, Len(slOutStr))
    End If
    'Format whole part of number
    If (ilFmtSpec And FMTCOMMA) = FMTCOMMA Then
        ilPos = InStr(slOutStr, ".")
        If ilPos = 0 Then   'Not fractional part
            slLeft = slOutStr
            slRight = ""
        Else
            slLeft = Left$(slOutStr, ilPos - 1)
            slRight = Mid$(slOutStr, ilPos, Len(slOutStr))
        End If
        If slLeft = "0" Then
            slOutStr = slLeft
        Else
            slOutStr = Format$(slLeft, "##,###,###")
        End If
        slOutStr = slOutStr & slRight
    End If
    If (ilFmtSpec And FMTDOLLARSIGN) = FMTDOLLARSIGN Then
        slOutStr = "$" & slOutStr
    End If
    If (ilFmtSpec And FMTPERCENTSIGN) = FMTPERCENTSIGN Then
        slOutStr = slOutStr & "%"
    End If
    If ilSignPos > 0 Then
        If (ilFmtSpec And FMTNEGASBRACKETS) = FMTNEGASBRACKETS Then
            slOutStr = "[" & slOutStr & "]"
        ElseIf (ilFmtSpec And FMTNEGATBACK) = FMTNEGATBACK Then
            slOutStr = slOutStr & "-"
        Else
            slOutStr = "-" & slOutStr
        End If
    End If
End Sub

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
    Dim ilRet As Integer
    Dim slStamp As String
    Dim slInitStamp As String
    Dim ilCount As Integer
    Dim hlUrf As Integer        'User Option file handle
    Dim tlUrf As URF
    Dim ilUrfRecLen As Integer
    Dim llTime As Long

    If tgSpf.sRemoteUsers = "Y" Then
        ilRet = 0
        On Error GoTo gGetSyncDateTimeErr:
        If tgUrf(0).iRemoteUserID <= 0 Then
            'Get Central date and time (i.e. if Dallas, get it from NY)
            ilCount = 0
            'hlFile = FreeFile
            If Trim$(sgMDBPath) <> "" Then
                'slInitStamp = FileDateTime(sgMDBPath & "RUStamp.Txt")
                slInitStamp = gFileDateTime(sgMDBPath & "URF.BTR")
                Do
                    'DoEvents
                    'ilRet = 0
                    'Open sgMDBPath & "RUStamp.Txt" For Output Shared As hlFile
                    'Print #hlFile, "Date Time Stamp File"
                    'Close #hlFile
                    'slStamp = FileDateTime(sgMDBPath & "RUStamp.Txt")
                    If ilCount <> 0 Then
                        llTime = Timer
                        Do
                            DoEvents
                        Loop While llTime = Timer
                    End If
                    hlUrf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
                    ilRet = btrOpen(hlUrf, "", sgDBPath & "URF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                    ilUrfRecLen = Len(tlUrf)
                    If ilRet = BTRV_ERR_NONE Then
                        ilRet = btrGetFirst(hlUrf, tlUrf, ilUrfRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            ilRet = btrUpdate(hlUrf, tlUrf, ilUrfRecLen)
                        End If
                    End If
                    btrDestroy hlUrf
                    slStamp = gFileDateTime(sgMDBPath & "URF.Btr")
                    ilCount = ilCount + 1
                    If ilCount > 10 Then
                        Exit Do
                    End If
                'Loop While (ilRet <> 0) Or (StrComp(slStamp, slInitStamp, 1) = 0)
                Loop While (ilRet <> BTRV_ERR_NONE) Or (StrComp(slStamp, slInitStamp, 1) = 0)
            Else
                'slInitStamp = FileDateTime(sgDBPath & "RUStamp.Txt")
                slInitStamp = gFileDateTime(sgDBPath & "URF.BTR")
                Do
                    'DoEvents
                    'ilRet = 0
                    'Open sgDBPath & "RUStamp.Txt" For Output Shared As hlFile
                    'Print #hlFile, "Date Time Stamp File"
                    'Close #hlFile
                    'slStamp = FileDateTime(sgDBPath & "RUStamp.Txt")
                    If ilCount <> 0 Then
                        llTime = Timer
                        Do
                            DoEvents
                        Loop While llTime = Timer
                    End If
                    hlUrf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
                    ilRet = btrOpen(hlUrf, "", sgDBPath & "URF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                    ilUrfRecLen = Len(tlUrf)
                    If ilRet = BTRV_ERR_NONE Then
                        ilRet = btrGetFirst(hlUrf, tlUrf, ilUrfRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            ilRet = btrUpdate(hlUrf, tlUrf, ilUrfRecLen)
                        End If
                    End If
                    btrDestroy hlUrf
                    slStamp = gFileDateTime(sgDBPath & "URF.BTR")
                    ilCount = ilCount + 1
                    If ilCount > 10 Then
                        Exit Do
                    End If
                'Loop While (ilRet <> 0) Or (StrComp(slStamp, slInitStamp, 1) = 0)
                Loop While (ilRet <> BTRV_ERR_NONE) Or (StrComp(slStamp, slInitStamp, 1) = 0)
            End If
            slSyncDate = Format$(slStamp, "m/d/yy")
            slSyncTime = Format$(slStamp, "h:mm:ssAM/PM")
        Else
            'Get Local date and time
            slSyncDate = Format$(gNow(), "m/d/yy")
            slSyncTime = Format$(gNow(), "h:mm:ssAM/PM")
        End If
        On Error GoTo 0
    Else
        slSyncDate = ""
        slSyncTime = ""
    End If
    Exit Sub
gGetSyncDateTimeErr:
    ilRet = err.Number
    Resume Next
End Sub




'*******************************************************
'*                                                     *
'*      Procedure Name:gIntToStrDec                    *
'*                                                     *
'*             Created:5/6/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Format as Integer with a decimal*
'*                                                     *
'*******************************************************
Function gIntToStrDec(ilInNumber As Integer, ilNoDecPlaces As Integer) As String
'
'   slOutStr = gIntToStrDec(ilInNumber, ilNoDecPlaces)
'   Where:
'       ilInNumber(I)- Number to be converted to string with/without dec point
'       ilNoDecPlaces(I)-Number of decimal places (can be zero)
'       slOutStr(O)- Output string format with a decimal point
'
    Dim slStr As String
    Dim ilNegSign As Integer

    If ilInNumber < 0 Then
        ilNegSign = True
    Else
        ilNegSign = False
    End If
    slStr = Trim$(str$(Abs(ilInNumber)))
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
    gIntToStrDec = slStr
    Exit Function
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
    slStr = Trim$(str$(Abs(llInNumber)))
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
    ilFieldNo = LBound(slFields)    '1
    ilIndex = 1
    ilFieldType = -1
    Do While ilIndex <= Len(Trim$(slCDStr))
        slChar = Mid$(slCDStr, ilIndex, 1)
        If (slChar <> " ") Or (slFields(ilFieldNo) <> "") Then
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
        End If
        ilIndex = ilIndex + 1
    Loop
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
'Removing calling the DLL as argument type problem when converting to C#
'    Dim ilStrLenNeeded As Integer
'    Dim slStr As String
'    Dim ilPos As Integer
'
'    slStr = slInputStr
'    Do While InStr(1, slStr, Chr(0), vbTextCompare) > 0
'        ilPos = InStr(1, slStr, Chr(0), vbTextCompare)
'        Mid$(slStr, ilPos, 1) = " "
'    Loop
'    ilStrLenNeeded = gParseGetStringSize(slStr, ilItemNo, slDelimiter)
'    If ilStrLenNeeded > 0 Then
'        slOutputStr = String(ilStrLenNeeded, " ")
'        gParseItem = gParseGetString(slOutputStr)
'        slOutputStr = Trim$(slOutputStr)
'        gParseItem = CP_MSG_NONE
'    ElseIf ilStrLenNeeded = 0 Then
'        slOutputStr = ""
'        gParseItem = CP_MSG_NONE
'    Else
'        slOutputStr = ""
'        gParseItem = CP_MSG_PARSE
'    End If
'    Exit Function

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
            gParseItem = CP_MSG_PARSE
            Exit Function
        End If
    Next ilIndex
    ilEndPos = InStr(ilStartPos, slInputStr, slDelimiter, 1)   'Position end to end of substring plus 1 (start of delimiter position)
    If (ilEndPos = 0) Then   'No end delimiter-copy reTrafficing string
        slOutputStr = Trim$(Mid$(slInputStr, ilStartPos))
        gParseItem = CP_MSG_NONE
        Exit Function
    End If
    slOutputStr = Trim$(Mid$(slInputStr, ilStartPos, ilEndPos - ilStartPos))
    gParseItem = CP_MSG_NONE
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gParseItemNoTrim                *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain a substring from a string*
'*                                                     *
'*******************************************************
Function gParseItemNoTrim(ByVal slInputStr As String, ByVal ilItemNo As Integer, slDelimiter As String, slOutputStr As String) As Integer
'
'   iRet = gParseItemNoTrim(slInputStr, ilItemNo, slDelimiter, slOutStr)
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
            gParseItemNoTrim = CP_MSG_PARSE
            Exit Function
        End If
    Next ilIndex
    ilEndPos = InStr(ilStartPos, slInputStr, slDelimiter, 1)   'Position end to end of substring plus 1 (start of delimiter position)
    If (ilEndPos = 0) Then   'No end delimiter-copy reTrafficing string
        slOutputStr = Mid$(slInputStr, ilStartPos)
        gParseItemNoTrim = CP_MSG_NONE
        Exit Function
    End If
    slOutputStr = Mid$(slInputStr, ilStartPos, ilEndPos - ilStartPos)
    gParseItemNoTrim = CP_MSG_NONE
End Function




'*******************************************************
'*                                                     *
'*      Procedure Name:gPDNToLong                      *
'*                                                     *
'*             Created:5/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Packed Decimal Number to Long   *
'*                                                     *
'*******************************************************
Sub gPDNToLong(slInPDN As String, llOutValue As Long)
'
'   gPDNToLong slInStr, llValue
'   Where:
'       slInStr (I)- Pack number to be converted
'       llValue (O)- Long value
'

    Dim slNumber As String  'String to be converted (left to right)
    Dim ilByte As Integer   'Extracted byte from string to be converted
    Dim ilLoop As Integer
    Dim slPDN As String

    '7/24/19: Trim in error as 201,442,602.38 results in 1,442,602.38
    'slPDN = Trim$(slInPDN)
    slPDN = slInPDN
    slNumber = slPDN
    llOutValue = 0
    For ilLoop = 1 To Len(slPDN) Step 1
        On Error GoTo gPDNToLongErr
        ilByte = Asc(slNumber)  'Obtain left most byte
        On Error GoTo 0
        slNumber = Mid$(slNumber, 2)    'Remove the left most byte
        If ilLoop = Len(slPDN) Then    'Last byte
            llOutValue = 10 * llOutValue + ilByte \ 16
            If (ilByte And 15) = 13 Then    'Negative number
                llOutValue = -llOutValue
            End If
            Exit Sub
        Else
            llOutValue = 100 * llOutValue + (ilByte \ 16) * 10 + (ilByte And 15)
        End If
    Next ilLoop
gPDNToLongErr:
    On Error GoTo 0
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gPDNToStr                       *
'*                                                     *
'*             Created:5/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Packed Decimal Number to String *
'*                                                     *
'*******************************************************
Sub gPDNToStr(slInPDN As String, ilNoDecPlaces As Integer, slOutValue As String)
'
'   gPDNToStr slPDN, ilDP, slStr
'   Where:
'       ilPDN (I)- Packed string to be converted
'       ilDP (I)- Number of decimal places
'       slStr (O) - Unpacked String
'

    Dim slNumber As String  'String to be converted (left to right)
    Dim ilByte As Integer   'Extracted byte from string to be converted
    Dim ilLoop As Integer
    Dim slPDN As String
    Dim slOut As String
    Dim slSign As String
    Dim slLeft As String
    Dim slRight As String
    Dim ilLen As Integer

'    Dim tlFix As FIXSTR
'    slOut = Space$(30)
'    ilByte = Len(slInPDN)
'    tlFix.sStr = slInPDN
'    ilLen = PDNToStr(tlFix, ilNoDecPlaces, ilByte, slOut)
'    If ilLen > 0 Then
'        slOutValue = Left$(slOut, ilLen)
'    Else
'        slOutValue = ""
'    End If
'    Exit Sub
    'Code below moved to DLL

    '7/24/19: Trim in error as 201,442,602.38 results in 1,442,602.38
    'slPDN = Trim$(slInPDN)
    slPDN = slInPDN
    slNumber = slPDN
    slSign = ""
    slOut = ""
    For ilLoop = 1 To Len(slPDN) Step 1
        On Error GoTo gPDNToStrErr
        ilByte = Asc(slNumber)  'Obtain left most byte
        On Error GoTo 0
        slNumber = Mid$(slNumber, 2)    'Remove the left most byte
        If ilLoop = Len(slPDN) Then    'Last byte
            slOut = slOut & Trim$(str$(ilByte \ 16))
            If (ilByte And 15) = 13 Then    'Negative number
                slSign = "-"
            End If
            Exit For
        Else
            slOut = slOut + Trim$(str$(ilByte \ 16)) + Trim$(str$(ilByte And 15))
        End If
    Next ilLoop
    If ilNoDecPlaces > 0 Then
        slLeft = Left$(slOut, Len(slOut) - ilNoDecPlaces)
        slRight = right$(slOut, ilNoDecPlaces)
    Else
        slLeft = slOut
    End If
    ilLen = Len(slLeft) - 1
    For ilLoop = 1 To ilLen Step 1
        If Left$(slLeft, 1) <> "0" Then
            Exit For
        End If
        slLeft = right$(slLeft, Len(slLeft) - 1)
    Next ilLoop
    If ilNoDecPlaces > 0 Then
        slOut = slLeft & "." & slRight
    Else
        slOut = slLeft
    End If
    slOutValue = Trim$(slSign & slOut)
    Exit Sub
gPDNToStrErr:
    On Error GoTo 0
    slOutValue = ""
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gStrDecToInt                    *
'*                                                     *
'*             Created:5/6/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Convert a String with decimal   *
'*                     to integer (round up)           *
'*                                                     *
'*******************************************************
Function gStrDecToInt(slInNumber As String, ilNoDecPlaces As Integer) As Integer
'
'   ilOut = gStrDecToInt(slInNumber, ilNoDecPlaces)
'   Where:
'       slInNumber(I)- Number to be converted to long
'       ilNoDecPlaces(I)-Number of decimal places (can be zero)
'       ilOut(O)- Output long
'                 slInNumber    # Dec Places   llOut
'                 234.56        2               23456
'                 234.56        0               235
'                 234.56        1               2346
'
'
    Dim ilPos As Integer
    Dim slStr As String
    Dim slLeft As String
    Dim slRight As String
    Dim ilAddOne As Integer
    Dim llNum As Long
    Dim ilLoop As Integer

    slStr = Trim$(slInNumber)
    ilPos = InStr(slStr, ".")
    If ilPos = 0 Then
        slStr = slStr & "."
    End If
    ilPos = InStr(slStr, ".")
    Do While Len(slStr) - ilPos < ilNoDecPlaces
        slStr = slStr & "0"
    Loop
    ilPos = InStr(slStr, ".")
    'Check Rounding
    ilAddOne = False
    If (Len(slStr) - ilPos) > ilNoDecPlaces Then
        slRight = right$(slStr, Len(slStr) - ilPos - ilNoDecPlaces)
        llNum = 5
        For ilLoop = 2 To Len(slRight) Step 1
            llNum = 10 * llNum
        Next ilLoop
        If Val(slRight) >= llNum Then
            ilAddOne = True
        End If
        slStr = Left(slStr, ilPos + ilNoDecPlaces)
    End If
    ilPos = InStr(slStr, ".")
    If ilPos > 0 Then
        slRight = Mid$(slStr, ilPos + 1)
        slLeft = Left$(slStr, ilPos - 1)
        slStr = slLeft & slRight
    End If
    If Val(slStr) > 32766! Then
        gStrDecToInt = 0
    Else
        If ilAddOne Then
            gStrDecToInt = Val(slStr) + 1
        Else
            gStrDecToInt = Val(slStr)
        End If
    End If
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gStrDecToLong                   *
'*                                                     *
'*             Created:5/6/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Convert a String with decimal   *
'*                     to long integer (round up)      *
'*                                                     *
'*******************************************************
Function gStrDecToLong(slInNumber As String, ilNoDecPlaces As Integer) As Long
'
'   llOut = gStrDecToLong(slInNumber, ilNoDecPlaces)
'   Where:
'       slInNumber(I)- Number to be converted to long
'       ilNoDecPlaces(I)-Number of decimal places (can be zero)
'       llOut(O)- Output long
'                 slInNumber     # Dec Places    llOut
'                 1234.56        2               123456
'                 1234.56        0               1235
'                 1234.56        1               12345
'                 1234.56        3               1234560
'
'
    Dim ilPos As Integer
    Dim slStr As String
    Dim slLeft As String
    Dim slRight As String
    Dim ilAddOne As Integer
    Dim llNum As Long
    Dim ilLoop As Integer

    slStr = Trim$(slInNumber)
    ilPos = InStr(slStr, ".")
    If ilPos = 0 Then
        slStr = slStr & "."
    End If
    ilPos = InStr(slStr, ".")
    Do While Len(slStr) - ilPos < ilNoDecPlaces
        slStr = slStr & "0"
    Loop
    ilPos = InStr(slStr, ".")
    'Check Rounding
    ilAddOne = False
    If (Len(slStr) - ilPos) > ilNoDecPlaces Then
        slRight = right$(slStr, Len(slStr) - ilPos - ilNoDecPlaces)
        llNum = 5
        For ilLoop = 2 To Len(slRight) Step 1
            llNum = 10 * llNum
        Next ilLoop
        If Val(slRight) >= llNum Then
            ilAddOne = True
        End If
        slStr = Left(slStr, ilPos + ilNoDecPlaces)
    End If
    ilPos = InStr(slStr, ".")
    If ilPos > 0 Then
        slRight = Mid$(slStr, ilPos + 1)
        slLeft = Left$(slStr, ilPos - 1)
        slStr = slLeft & slRight
    End If
    If Val(slStr) > 2.147482E+09! Then
        gStrDecToLong = 0
    Else
        If ilAddOne Then
            gStrDecToLong = Val(slStr) + 1
        Else
            gStrDecToLong = Val(slStr)
        End If
    End If
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gStrToPDN                       *
'*                                                     *
'*             Created:5/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:String to Packed Decimal Number *
'*                                                     *
'*******************************************************
Sub gStrToPDN(slInValue As String, ilNoDecPlaces As Integer, ilPDNLen As Integer, slOutPDN As String)
'
'   gStrToPDN slValue, iNoDP, ilLen, slStr
'   Where:
'       slValue (I)- String value to be packed (it can contain , $ - + .)
'       iNoDP (I)- Number of decimal places
'       ilLen (I)- Length of output string (number of bytes-two digits per byte, last byte contains one digit and sign)
'       slStr (O) - String to contain pack number
'

    Dim slNumExtra As String      'Number with extra characters
    Dim slNumStrip As String    'Number with characters removed
    Dim slLeftChar As String * 1
    Dim ilLoop As Integer
    Dim ilNoDp As Integer   'Number of decimal places found
    Dim ilCount As Integer  'Count decimal places
    Dim ilPosSign As Integer
    Dim ilLen As Integer

    slNumExtra = Trim$(slInValue)
    ilLen = Len(slNumExtra)
    'Remove all characters and count number of places after decimal
    ilNoDp = 0
    ilCount = False
    ilPosSign = True
    slNumStrip = ""
    For ilLoop = 1 To ilLen Step 1
        slLeftChar = Left$(slNumExtra, 1)
        slNumExtra = Mid$(slNumExtra, 2)
        If slLeftChar >= "0" And slLeftChar <= "9" Then
            slNumStrip = slNumStrip & slLeftChar
            If ilCount Then
                ilNoDp = ilNoDp + 1
            End If
        Else
            If slLeftChar = "." Then
                ilCount = True
            End If
            If slLeftChar = "-" Then
                ilPosSign = False
            End If
        End If
        If slNumExtra = "" Then
            Exit For
        End If
    Next ilLoop
    If ilNoDp < ilNoDecPlaces Then  'Add zeros
        For ilLoop = ilNoDp To ilNoDecPlaces - 1 Step 1
            slNumStrip = slNumStrip & "0"
        Next ilLoop
    End If
    If ilNoDp > ilNoDecPlaces Then  'Remove extra digits
        slNumStrip = Left$(slNumStrip, Len(slNumStrip) - (ilNoDp - ilNoDecPlaces))
    End If
    If Not ilPosSign Then
        slNumStrip = "-" & slNumStrip
    End If
    mMakePDN slNumStrip, ilPDNLen, slOutPDN
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:gUnformatStr                    *
'*                                                     *
'*             Created:5/29/93        By:D. LeVine     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Unformat money or percentage    *
'*                     (remove commas, dollar signs,..)*
'*                                                     *
'*******************************************************
Sub gUnformatStr(slInStr As String, ilFmtSpec As Integer, slOutStr As String)
'
'   gUnformatStr slStr, ilUnFmt, slOut
'   Where:
'       slStr (I)- String to be unformatted
'       ilUnFmt (I)- Format specifications
'                  UNFMTDEFAULT = 0     'Only remove special characters
'                  UNFMTREMOVEDECZERO = 1     'remove decimal value if zero
'       slOut (O) - String to contain unformatted string
'

    Dim slStr As String
    Dim slSvStr As String
    Dim ilLoop As Integer
    Dim ilZeroOnly As Integer
    Dim ilNegSign As Integer
    Dim ilPos As Integer

    slSvStr = slInStr
    ilZeroOnly = True
    slOutStr = ""
    ilPos = InStr(slSvStr, "-")
    If ilPos > 0 Then
        ilNegSign = True
    Else
        ilPos = InStr(slSvStr, "[")
        If ilPos > 0 Then
            ilNegSign = True
        Else
            ilNegSign = False
        End If
    End If
    For ilLoop = 1 To Len(slSvStr) Step 1
        slStr = Mid$(slSvStr, ilLoop, 1)
        If ((slStr >= "0") And (slStr <= "9")) Or slStr = "." Then
            slOutStr = slOutStr & slStr
        End If
    Next ilLoop
    If ilNegSign Then
        slOutStr = "-" & slOutStr
    End If
    If (ilFmtSpec And UNFMTREMOVEDECZERO) = UNFMTREMOVEDECZERO Then
        ilPos = InStr(slSvStr, ".")
        If ilPos > 0 Then
            For ilLoop = ilPos + 1 To Len(slSvStr) Step 1
                slStr = Mid$(slSvStr, ilLoop, 1)
                If slStr <> "0" Then
                    Exit Sub
                End If
            Next ilLoop
            slOutStr = Left$(slOutStr, ilPos - 1)
        End If
    End If
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mMakePDN                        *
'*                                                     *
'*             Created:5/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:String to Packed Decimal Number *
'*                                                     *
'*******************************************************
Private Sub mMakePDN(slInStr As String, ilPDNLen As Integer, slOutPDN As String)
'
'   mMakeToPDN slInValue, ilLen, slOutStr
'   Where:
'       slInValue (I)- string value to be packed
'       ilLen (I)- Length of output string (number of bytes-two digits per byte, last byte contains one digit and sign)
'       slOutStr (O) - String to contain pack number
'

    Dim ilRem As Integer    'Remainder (value to be packed)
    Dim ilHRem As Integer   'High part of nibble to be packed
    Dim ilLRem As Integer   'Low part of nibble to be packed
    Dim ilNibble As Integer 'Nibble (1=High part, 2 = Low part)
    Dim slvalue As String   'Value as string to be converted
    Dim ilPos As Integer    'Character position to be converted
    Dim ilZero As Integer   'ANSI code of zero: Asc("0")
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilIndex As Integer  'Location within string to store packed value
    Dim slRem As String 'Packed value for byte

    If Left$(slInStr, 1) = "-" Then
        slvalue = Mid$(slInStr, 2)
        ilLRem = 13 '"D"=Negative number
    Else
        If Left$(slInStr, 1) = "+" Then
            slvalue = Mid$(slInStr, 2)
        Else
            slvalue = slInStr
        End If
        ilLRem = 15 '"F"=Positive number
    End If
    'Initialize string to all zero's
    slOutPDN = ""
    For ilLoop = 1 To ilPDNLen Step 1
        slOutPDN = slOutPDN + Chr$(0)
    Next ilLoop
    'Pack number
    ilNibble = 1
    ilIndex = ilPDNLen
    ilPos = Len(slvalue)
    If ilPos <= 0 Then
        ilHRem = 0
        slRem = Chr$(ilHRem * 16 + ilLRem)
        Mid$(slOutPDN, ilIndex, 1) = slRem
        Exit Sub
    End If
    ilZero = Asc("0")
    For ilLoop = 1 To 2 * ilPDNLen - 1 Step 1
        ilRem = Asc(Mid$(slvalue, ilPos, 1)) - ilZero
        ilPos = ilPos - 1
        If ilNibble = 1 Then
            ilHRem = ilRem
        Else
            ilLRem = ilRem
            ilHRem = 0
            ilNibble = 0
            ilIndex = ilIndex - 1
        End If
        slRem = Chr$(ilHRem * 16 + ilLRem)
        Mid$(slOutPDN, ilIndex, 1) = slRem
        ilNibble = ilNibble + 1
        If ilPos <= 0 Then
            Exit For
        End If
    Next ilLoop
End Sub


Public Sub gShellAndWait(frm As Form, ByVal slFilePath As String, ilWinStyle As Integer, blEnableForm As Boolean)
    Dim llProcess As Long
    Dim llReturn As Long
    lgShellAndWaitID = Shell(slFilePath, ilWinStyle)
    llProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, lgShellAndWaitID)
    If llProcess <> 0 Then
        'Frm.WindowState = vbMinimized
        If blEnableForm Then
            frm.Enabled = False
        End If
        'Frm.Visible = False
        Do
            GetExitCodeProcess llProcess, llReturn
            Sleep 50
            DoEvents
        Loop While llReturn = STILL_ACTIVE
        frm.Enabled = True
        'If Frm = "Traffic" Then
        '    FrmWindowState = vbMaximized
        'Else
        '    Frm.WindowState = vbNormal
        'End If
        'Frm.Visible = True
        lgShellAndWaitID = 0
        llReturn = CloseHandle(llProcess)
    Else
        MsgBox "Unable to Shell to " & slFilePath, vbOKOnly, "Shell Error"
    End If
End Sub

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

Public Sub gMsgBox(Optional slErrMsg As String = "Press OK to Continue", Optional llMsgBoxStyle = vbInformation, Optional slTitle As String = "Traffic")
'Public Sub gMsgBox(slErrMsg As String, llMsgBoxStyle As Long, slTitle As String)
    If igBkgdProg = 0 Then
        If llMsgBoxStyle <> -1 Then
            If slTitle <> "" Then
                MsgBox slErrMsg, llMsgBoxStyle, slTitle
            Else
                MsgBox slErrMsg, llMsgBoxStyle
            End If
        End If
        '3-21-07 always log the error
            gLogMsg slErrMsg, "TrafficErrors.Txt", False
        'End If
    ElseIf igBkgdProg = 1 Then
        gLogMsg slErrMsg, "Bkgd_Schd.Txt", False
    ElseIf igBkgdProg = 2 Then
        gLogMsg slErrMsg, "Set_Credit.Txt", False
    ElseIf igBkgdProg = 3 Then
        gLogMsg slErrMsg, "ExportSQL.Txt", False
    ElseIf igBkgdProg = 4 Then
        gLogMsg slErrMsg, "ExportGetPaid.Txt", False
    ElseIf igBkgdProg = 5 Then
        gLogMsg slErrMsg, "GenASF.Txt", False
    ElseIf igBkgdProg = 6 Then
        gLogMsg slErrMsg, "RepNetLink.Txt", False
    ElseIf igBkgdProg = 7 Then
        gLogMsg slErrMsg, "DB_Archive.Txt", False
    ElseIf igBkgdProg = 8 Then
        gLogMsg slErrMsg, "ExportProj.Txt", False
    ElseIf igBkgdProg = 10 Then
        gLogMsg slErrMsg, "EngrService.Txt", False
    ElseIf igBkgdProg = 11 Then
        gLogMsg slErrMsg, "DB_Extractor_Log.Txt", False
    ElseIf igBkgdProg = 12 Then
        gLogMsg slErrMsg, "ExportInvSpots.Txt", False
    ElseIf igBkgdProg = 13 Then
        gLogMsg slErrMsg, "ResetMGSpec.Txt", False
    ElseIf igBkgdProg = 14 Then
        gLogMsg slErrMsg, "ExportSalesForce.Txt", False
    ElseIf igBkgdProg = 15 Then
        gLogMsg slErrMsg, "ExpMatrix.Txt", False
    ElseIf igBkgdProg = 16 Then
        gLogMsg slErrMsg, "EfficioExport.Txt", False
    ElseIf igBkgdProg = 17 Then
        gLogMsg slErrMsg, "ImportResearch.Txt", False
    ElseIf igBkgdProg = 18 Then                             '7-8-15 Tableau export added (same as matrix format)
        gLogMsg slErrMsg, "ExpTableau.Txt", False
    ElseIf igBkgdProg = 20 Then                             '7-8-15 Tableau export added (same as matrix format)
        gLogMsg slErrMsg, "ProgrammaticBuy.Txt", False
    ElseIf igBkgdProg = 21 Then                             '1-29-20 RAB
        gLogMsg slErrMsg, "ExpRAB.txt", False
    ElseIf igBkgdProg = 22 Then
        gLogMsg slErrMsg, "PodcastCntrImp.txt", False
    ElseIf igBkgdProg = 23 Then
        gLogMsg slErrMsg, "PodcastInvImpt.txt", False
    ElseIf igBkgdProg = 24 Then                             'TTP 9992 - Custom Revenue Export - logging
        gLogMsg slErrMsg, "ExpCustomRevenueExport.Txt", False
    Else
        gLogMsg slErrMsg, "TrafficErrors.Txt", False
    End If
End Sub

Public Function gFixQuote(sInStr As String) As String 'VBC NR

    Dim sOutStr As String 'VBC NR
    Dim sChar As String 'VBC NR
    Dim iLoop As Integer 'VBC NR

    sOutStr = "" 'VBC NR
    If IsNull(sInStr) <> True Then 'VBC NR
        For iLoop = 1 To Len(sInStr) Step 1 'VBC NR
            sChar = Mid$(sInStr, iLoop, 1) 'VBC NR
            If sChar = "'" Then 'VBC NR
                sOutStr = sOutStr & "''" 'VBC NR
            Else 'VBC NR
                sOutStr = sOutStr & sChar 'VBC NR
            End If 'VBC NR
        Next iLoop 'VBC NR
    End If 'VBC NR
    gFixQuote = sOutStr 'VBC NR
End Function 'VBC NR

Public Function gStrongPassword(slPassword As String) As Integer
    Dim ilNoUCaseLetters As Integer
    Dim ilNoLCaseLetters As Integer
    Dim ilNoNumbers As Integer
    Dim ilNoSymbols As Integer
    Dim ilLoop As Integer
    Dim ilCheck As Integer
    Dim slStr As String

    slStr = Trim(slPassword)
    ilNoUCaseLetters = 0
    ilNoLCaseLetters = 0
    ilNoNumbers = 0
    ilNoSymbols = 0
    If ((Asc(tgSpf.sUsingFeatures2) And STRONGPASSWORD) = STRONGPASSWORD) Then
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


Public Sub gSetPictureBoxFontSize(frm As Form, flFontSize As Single)
    Dim Ctrl As Control
    Dim ilRet As Integer

    If Screen.Width * 15 = 640 Then
        Exit Sub
    End If
    If InStr(1, frm.Name, "Rpt", vbTextCompare) = 1 Then
        Exit Sub
    End If
    For Each Ctrl In frm.Controls
        If TypeOf Ctrl Is PictureBox Then
            If Ctrl.height >= 360 Then
                If Ctrl.BackColor <> BLUE Then
                    If InStr(1, Ctrl.Name, "Plc", vbTextCompare) = 1 Then
                        ilRet = ilRet
                    Else
                        If (Ctrl.Font.Name <> "Arial") Or ((Ctrl.Font.Name = "Arial") And (Ctrl.Font.Size > 8)) Then
                            Ctrl.Font.Size = flFontSize
                        End If
                    End If
                End If
            End If
        End If
    Next Ctrl
End Sub

Public Sub gSendKeys(edcCtrl As TextBox, slSendKey As String)
    Dim slStr As String
    Dim ilSelStart As Integer
    slStr = edcCtrl.Text
    ilSelStart = edcCtrl.SelStart
    edcCtrl.Text = Mid(slStr, 1, edcCtrl.SelStart) & slSendKey & Mid(slStr, edcCtrl.SelStart + edcCtrl.SelLength + 1)
    edcCtrl.SelStart = ilSelStart + Len(slSendKey)
End Sub

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
slGifPath = sgRptPath & IMAGENAME
    'modify message picture box
    plcMyBox.BackColor = vbButtonFace
    plcMyBox.Width = plcMyBox.Width + IMAGESIZE + IMAGEWEBPADWIDTH + RIGHTMARGIN
    'create picturebox to hold webbrowswer --need to constrict webbrowser to lose frame
    Set picBox = myForm.Controls.Add("VB.PictureBox", "picBox")
    With picBox
        .height = IMAGESIZE + IMAGEWEBPADHEIGHT
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
        .height = IMAGESIZE + IMAGEWEBPADHEIGHT
        .Move -myForm.ScaleX(2, vbPixels), -myForm.ScaleY(2, vbPixels)
        'change image source!
        ' .Navigate "about:<html><body scroll='no'><img src='D:\vbAnimatePic\wait.gif'></img></body></html>"
         .Navigate "about:<html><body scroll='no'><img src='" & slGifPath & "'></img></body></html>"
        .Document.bgColor = "#ece9d8"
        .Visible = True
    End With
End Sub
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
Public Sub btrStopAppl()
    btrStopApplReal
    If Len(sgShortDate) > 0 And sgShortDate <> MYSHORTDATE Then
        gRegistrySetShortDate (sgShortDate)
    End If
End Sub
'end 5666

Public Function gReplaceIllegalCharacters(sInStr As String) As String

    Dim sOutStr As String
    Dim sChar As String
    Dim iLoop As Integer
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer

    sOutStr = ""
    If IsNull(sInStr) <> True Then
        ilPos1 = InStr(1, sInStr, "[", vbBinaryCompare)
        ilPos2 = InStr(ilPos1 + 1, sInStr, "]", vbBinaryCompare)
        If (ilPos1 <= 0) Or (ilPos2 <= ilPos1 + 1) Then
            For iLoop = 1 To Len(sInStr) Step 1
                sChar = Mid$(sInStr, iLoop, 1)
                If gCheckEnteredChar(Asc(sChar)) = 0 Then
                    sOutStr = sOutStr & " "
                Else
                    sOutStr = sOutStr & sChar
                End If
            Next iLoop
        Else
            sOutStr = sInStr
        End If
    End If
    gReplaceIllegalCharacters = sOutStr
End Function

Public Function gGetListBoxRow(lbcCtrl As ListBox, Y As Single) As Long
    Dim llRow As Long
    Dim slStr As String
    Dim llRowHeight As Long

    If lbcCtrl.ListCount <= 0 Then
        gGetListBoxRow = -1
        Exit Function
    End If
    llRowHeight = 15 * SendMessageByString(lbcCtrl.HWnd, LB_GETITEMHEIGHT, 0, slStr)
    llRow = (Y - 15) \ llRowHeight + lbcCtrl.TopIndex
    gGetListBoxRow = llRow
End Function



Public Function gGetCallLettersByShttCode(shttCode As Integer) As String

    'D.S. 12/9/05
    'Get a station's call letters from it's Shtt Code
    Dim tmp_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    SQLQuery = "Select shttCallLetters FROM shtt where shttCode = " & shttCode
    'Set tmp_rst = cnn.Execute(SQLQuery)
    Set tmp_rst = gSQLSelectCall(SQLQuery)
    If Not tmp_rst.EOF Then
        If IsNull(tmp_rst!shttCallLetters) = True Then
            gGetCallLettersByShttCode = ""
        Else
            gGetCallLettersByShttCode = Trim$(tmp_rst!shttCallLetters)
        End If
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function
Public Sub gGetContractParameters(llChfCode As Long, slType As String, ilPctTrade As Integer)

    Dim tmp_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    slType = ""
    ilPctTrade = 0
    SQLQuery = "Select chfType, chfPctTrade FROM chf_Contract_Header where chfCode = " & llChfCode
    'Set tmp_rst = cnn.Execute(SQLQuery)
    Set tmp_rst = gSQLSelectCall(SQLQuery)
    If Not tmp_rst.EOF Then
        slType = tmp_rst!chftype
        ilPctTrade = tmp_rst!chfPctTrade
    End If
    Exit Sub
ErrHand:
    On Error GoTo 0
End Sub


Public Sub gSleep(ilTotalSeconds As Long)
    
    Dim ilLoop As Integer
    For ilLoop = 0 To ilTotalSeconds
        DoEvents
        Sleep (1000)   ' Wait 1 second
    Next
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
Public Function gAddQuotes(slInStr As String) As String
    gAddQuotes = """" & slInStr & """"
End Function
Public Function gXmlIniPath(Optional blTestExistence As Boolean = False) As String
'dan M 11/01/2010
'search data folder, then exe folder; if not there, assume it's in sgDirectory folder. Use blTestExistence if don't want to assume, and return "" if doesn't exist.
    Dim oMyFileObj As FileSystemObject
    Dim slIniPath As String
    Set oMyFileObj = New FileSystemObject
    '4/14/12: Changed order of search to: Startup; database; exe From: database; exe; startup
    '         the reason for the change to to handle the case were two xml.ini definded and one placed
    '         in the database and the other placed into an ini folder.
'    slIniPath = oMyFileObj.BuildPath(sgDBPath, "xml.ini")
'    If oMyFileObj.FileExists(slIniPath) Then
'        gXmlIniPath = slIniPath
'    Else
'        slIniPath = oMyFileObj.BuildPath(sgExeDirectory, "xml.ini")
'        If oMyFileObj.FileExists(slIniPath) Then
'            gXmlIniPath = slIniPath
'        Else
'            slIniPath = oMyFileObj.BuildPath(sgStartupDirectory, "xml.ini")
'            If blTestExistence Then
'                If Not oMyFileObj.FileExists(slIniPath) Then
'                    slIniPath = vbNullString
'                End If
'            End If
'            gXmlIniPath = slIniPath
'        End If
'    End If
    slIniPath = oMyFileObj.BuildPath(App.Path, "xml.ini")
    If oMyFileObj.FILEEXISTS(slIniPath) Then
        gXmlIniPath = slIniPath
    Else
        slIniPath = oMyFileObj.BuildPath(sgDBPath, "xml.ini")
        If oMyFileObj.FILEEXISTS(slIniPath) Then
            gXmlIniPath = slIniPath
        Else
            slIniPath = oMyFileObj.BuildPath(sgExePath, "xml.ini")
            If Not oMyFileObj.FILEEXISTS(slIniPath) Then
                slIniPath = vbNullString
            End If
            gXmlIniPath = slIniPath
        End If
    End If
    Set oMyFileObj = Nothing
End Function

'***************************************************************************************
'*
'* Procedure Name: gLoadOption
'*
'* Created: 8/22/03 - J. Dutschke
'*
'* Modified:              By:
'*
'* Comments: This function loads a string value from the ini file.
'*
'***************************************************************************************
Public Function gLoadOption(Section As String, Key As String, sValue As String) As Boolean
    On Error GoTo ERR_gLoadOption
    Dim BytesCopied As Integer
    Dim sBuffer As String * 128
    Dim slFileName As String

    If igDirectCall = -1 Then
        slFileName = sgIniPath & "Traffic.Ini"
    Else
        slFileName = CurDir$ & "\Traffic.Ini"
    End If

    gLoadOption = False
    BytesCopied = GetPrivateProfileString(Section, Key, "Not Found", sBuffer, 128, slFileName)
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
Public Function gAddQuotesRemoveChars(slInStr As String, Optional blRelaceDoubleQuotes As Boolean = False) As String
    Dim slStr As String
    slStr = gRemoveIllegalChars(slInStr, blRelaceDoubleQuotes)
    gAddQuotesRemoveChars = """" & Trim$(slStr) & """"
End Function
Public Function gRemoveIllegalChars(slInStr As String, Optional blRelaceDoubleQuotes As Boolean = False) As String
    Dim slStr As String
    Dim ilLen As Integer
    Dim ilLoop As Integer
    Dim slOneChar As String
    
    ilLen = Len(slInStr)
    slStr = ""
    For ilLoop = 1 To ilLen Step 1
        slOneChar = Mid(slInStr, ilLoop, 1)
        If (Asc(slOneChar) >= Asc(" ")) And (Asc(slOneChar) <= Asc("~")) Then
            If (Asc(slOneChar) = 34) And (blRelaceDoubleQuotes) Then
                slStr = slStr & "'"
            Else
                slStr = slStr & slOneChar
            End If
        End If
    Next
    gRemoveIllegalChars = Trim$(slStr)
End Function
Public Function gRemoveLeadingCrCl(slInStr As String) As String
    Dim slOutStr As String
    Dim llPos As Long
    
    slOutStr = slInStr
    llPos = InStr(slOutStr, sgCR & sgLF)
    If llPos = 1 Then
        slOutStr = Mid$(slOutStr, 3)
    End If
    gRemoveLeadingCrCl = slOutStr
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gFileNameFilter                 *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove illegal characters from *
'*                      name                           *
'*                                                     *
'*******************************************************
Function gFileNameFilterPlusBlank(slInName As String) As String
    'Taken from Affiliate
    Dim slName As String
    Dim ilPos As Integer
    Dim ilFound As Integer
    slName = slInName
    'Remove " and '
    Do
        ilFound = False
        ilPos = InStr(1, slName, "'", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, """", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, "&", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "/", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "\", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "*", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ":", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "?", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "%", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        'ilPos = InStr(1, slName, """", 1)
        'If ilPos > 0 Then
        '    Mid$(slName, ilPos, 1) = "'"
        '    ilFound = True
        'End If
        ilPos = InStr(1, slName, "=", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "+", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "<", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ">", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "|", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ";", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "@", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "[", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "]", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "{", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "}", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "^", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ".", 1)    'If period, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ",", 1)    'If comma, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
        ilPos = InStr(1, slName, " ", 1)    'If space, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
    Loop While ilFound
    gFileNameFilterPlusBlank = slName
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gFileNameFilterNotPath          *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove illegal characters from *
'*                      name                           *
'*                                                     *
'*******************************************************
Function gFileNameFilterNotPath(slInName As String) As String
    'Taken from Affiliate
    Dim slName As String
    Dim ilPos As Integer
    Dim ilFound As Integer
    Dim slPath As String
    
    slPath = ""
    slName = slInName
    ilPos = InStrRev(slName, "\")
    If ilPos > 0 Then
        slPath = Left$(slName, ilPos)
        slName = Mid$(slName, ilPos + 1)
    End If
    'Remove " and '
    Do
        ilFound = False
        ilPos = InStr(1, slName, "'", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, """", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, "&", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "/", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "\", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "*", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ":", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "?", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "%", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        'ilPos = InStr(1, slName, """", 1)
        'If ilPos > 0 Then
        '    Mid$(slName, ilPos, 1) = "'"
        '    ilFound = True
        'End If
        ilPos = InStr(1, slName, "=", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "+", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "<", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ">", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "|", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ";", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "@", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "[", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "]", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "{", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "}", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "^", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        'ilPos = InStr(1, slName, ".", 1)    'If period, use underscore
        'If ilPos > 0 Then
        '    Mid$(slName, ilPos, 1) = "_"
        '    ilFound = True
        'End If
        ilPos = InStr(1, slName, ",", 1)    'If comma, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
        ilPos = InStr(1, slName, " ", 1)    'If space, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
    Loop While ilFound
    gFileNameFilterNotPath = slPath & slName
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gFileNameFilter2                *
'*            ReWritten:              By:J White       *
'*                                                     *
'*            Comments: Remove illegal characters from *
'*                      name, copied from Exports      *
'*                      gFileNameFilter                *
'*******************************************************
Function gFileNameFilter2(slInName As String) As String
    Dim slName As String
    Dim ilPos As Integer
    Dim ilFound As Integer
    slName = slInName
    slName = Replace(slName, "'", "")
    slName = Replace(slName, """", "")
    slName = Replace(slName, "&", "-")
    slName = Replace(slName, "/", "-")
    slName = Replace(slName, "\", "-")
    slName = Replace(slName, "*", "-")
    slName = Replace(slName, ":", "-")
    slName = Replace(slName, "?", "-")
    slName = Replace(slName, "%", "-")
    slName = Replace(slName, "=", "-")
    slName = Replace(slName, "+", "-")
    slName = Replace(slName, "<", "-")
    slName = Replace(slName, ">", "-")
    slName = Replace(slName, "|", "-")
    slName = Replace(slName, ";", "-")
    slName = Replace(slName, "@", "-")
    slName = Replace(slName, "[", "-")
    slName = Replace(slName, "]", "-")
    slName = Replace(slName, "{", "-")
    slName = Replace(slName, "}", "-")
    slName = Replace(slName, "^", "-")
    slName = Replace(slName, ".", "-")
    slName = Replace(slName, ",", "-")
    slName = Replace(slName, " ", "_")
    gFileNameFilter2 = slName
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

'*******************************************************
'*                                                     *
'*      Procedure Name:gExistLtf                       *
'*                                                     *
'*             Created:2/8/21        By:J. White       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Bool. Test if LTF contains      *
'*                     any programming defined for a   *
'*                     given Vehicle                   *
'*                                                     *
'*******************************************************
Function gExistLtf(ilVefCode As Integer) As Boolean
    Dim Ltf_rst As ADODB.Recordset
    Dim slSQLQuery As String
    
    gExistLtf = False
    slSQLQuery = "SELECT Count(1) as TestCount FROM LTF_Lbrary_Title WHERE LtfVefCode = " & ilVefCode
    Set Ltf_rst = gSQLSelectCall(slSQLQuery)
    If Not Ltf_rst.EOF Then
        If Ltf_rst!TestCount > 0 Then
            gExistLtf = True
        End If
    End If
    Ltf_rst.Close
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gExistClf                       *
'*                                                     *
'*             Created:1/22/21       By:J. White       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if Contract contains any   *
'*                     clf Records                     *
'*                                                     *
'*******************************************************
Function gExistClf(lChfCode As Long) As Boolean
    gExistClf = False
    
    Dim clf_rst As ADODB.Recordset
    Dim slSQLQuery As String
    slSQLQuery = "SELECT Count(1) as TestCount FROM CLF_Contract_Line WHERE clfChfCode = " & lChfCode
    Set clf_rst = gSQLSelectCall(slSQLQuery)
    If Not clf_rst.EOF Then
        If clf_rst!TestCount > 0 Then
            gExistClf = True
        End If
    End If
    clf_rst.Close
End Function

'TTP 10504 - Amazon web bucket upload cleanup
Public Function AmazonBucketFolder(slBucketName, slSubfolder) As String
    Dim slAmazonBucketFolder As String

    slBucketName = Trim(slBucketName)
    slBucketName = Replace(slBucketName, "\", "/")
    slBucketName = Replace(slBucketName, "|", "/")
    slBucketName = Replace(slBucketName, "//", "/")
    If right(slBucketName, 1) = "/" Then slBucketName = Mid(slBucketName, 1, Len(slBucketName) - 1) 'remove trailing /
    If Left(slBucketName, 1) = "/" Then slBucketName = Mid(slBucketName, 2) 'remove begining /
    
    slSubfolder = Trim(slSubfolder)
    slSubfolder = Replace(slSubfolder, "\", "/")
    slSubfolder = Replace(slSubfolder, "|", "/")
    slSubfolder = Replace(slSubfolder, "//", "/")
    If right(slSubfolder, 1) = "/" Then slSubfolder = Mid(slSubfolder, 1, Len(slSubfolder) - 1) 'remove trailing /
    If Left(slSubfolder, 1) = "/" Then slSubfolder = Mid(slSubfolder, 2) 'remove begining /
    
    If slSubfolder <> "" Then
        slAmazonBucketFolder = slBucketName + "/" + slSubfolder
    Else
        slAmazonBucketFolder = slBucketName
    End If
    
    AmazonBucketFolder = slAmazonBucketFolder
End Function

