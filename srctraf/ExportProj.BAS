Attribute VB_Name = "ExportProj1"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExportProj.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Global Variables for Set Advertiser/Agency Credit
Option Explicit
Option Compare Text

'Public tgStations() As SHTTINFO
'Public sgStationsStamp As String


Type PROJINFO
    lDate As Long
    iLen As Integer
    iInventory(0 To 24) As Integer  'inventory defined for this spot length, Index zero ignored
    lMinRate(0 To 24) As Long       'lowest spot rate (zero included), Index zero ignored
    lMaxRate(0 To 24) As Long       'highest spot rate, Index zero ignored
    lSchedRev(0 To 24) As Long      'total $ scheduled (missed included), Index zero ignored
    lMissedRev(0 To 24) As Long     'total $ missed, Index zero ignored
    iSchedUnits(0 To 24) As Integer 'total scheduled  units (missed included), Index zero ignored
    iNCUnits(0 To 24) As Integer    'total no charge (zero)  units, Index zero ignored
    iMissedUnits(0 To 24) As Integer    'total missed units, Index zero ignored
End Type


Public Function gAddQuotes(slInStr As String) As String
    Dim slStr As String
    slStr = gRemoveIllegalChars(slInStr)
    gAddQuotes = """" & Trim$(slStr) & """"
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



'*******************************************************
'*                                                     *
'*      Procedure Name:gFindMatch                      *
'*                                                     *
'*             Created:10/20/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Find matching match within list*
'*                      box                            *
'*                                                     *
'*******************************************************
Sub gFindMatch(slInMatch As String, ilFirstAllowedIndex As Integer, lbcList As Control)
'
'   gFindMatch slInMatch, ilFirstAllowedIndex, lbcList
'   Where:
'       slInMatch(I)- String to be found
'       ilFirstAllowedIndex(I)- First legal index
'       lbcList(I)- list box containing strings
'
'       lbcList.LastFound(O)- Found index
'
    Dim ilFound As Integer
    Dim slMatch As String
    Dim ilPass As Integer   'Pass: 0= Trim string; 1=Leave leading blanks
    Dim ilRow As Integer
    Dim slNext As String

    For ilPass = 0 To 1 Step 1
        ilFound = False
        If ilPass = 0 Then
            slMatch = Trim$(slInMatch)
        Else
            slMatch = RTrim$(slInMatch)
        End If
        gFndFirst lbcList, slMatch
        Do
            If gLastFound(lbcList) >= ilFirstAllowedIndex Then
                If StrComp(lbcList.List(gLastFound(lbcList)), slMatch, 1) = 0 Then
                    Exit Sub
                Else
                    ilFound = False
                    gFndNext lbcList, slMatch
                End If
            Else
                If ilPass = 1 Then
                    Exit Sub
                Else
                    Exit Do
                End If
            End If
        Loop While Not ilFound
    Next ilPass
End Sub


Public Sub gFndFirst(lbcList As Control, slInMatch As String)
    cgList = lbcList
    If TypeOf lbcList Is ComboBox Then
        igFoundRow = SendMessageByString(lbcList.hwnd, CB_FINDSTRING, -1, slInMatch)
    Else
        igFoundRow = SendMessageByString(lbcList.hwnd, LB_FINDSTRING, -1, slInMatch)
    End If
End Sub

Public Sub gFndNext(lbcList As Control, slInMatch As String)
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


Public Function gLastFound(lbcList As Control) As Integer
    If cgList <> lbcList Then
        gLastFound = -1
    Else
        gLastFound = igFoundRow
    End If
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:gSetPhoneNo                     *
'*                                                     *
'*             Created:4/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Insert phone number into mask   *
'*                     field.  Works with and without  *
'*                     extension                       *
'*                                                     *
'*******************************************************
Sub gSetPhoneNo(slPhone As String, mkcCtrl As Control)
'
'   gSetPhoneNo mkcCtrl, sPhone
'   Where:
'       sPhone (I)- Phone number to be inserted into the mask control
'       mkcCtrl (O)- Mask edit control to receive the phone number
'

    'Clear field as blanks will not place old values
    If Len(slPhone) <= 10 Then
        Traffic!mkcPhone.Mask = "(AAA) AAA-AAAA"
        Traffic!mkcPhone.MaxLength = 14
    Else
        Traffic!mkcPhone.Mask = "(AAA) AAA-AAAA Ext(AAAA)"
        Traffic!mkcPhone.MaxLength = 24
    End If
    Traffic!mkcPhone.SelStart = 0
    Traffic!mkcPhone.SelLength = Traffic!mkcPhone.MaxLength
    Traffic!mkcPhone.SelText = ""
    Traffic!mkcPhone.SelStart = 1
    Traffic!mkcPhone.SelLength = 3
    Traffic!mkcPhone.SelText = Mid$(slPhone, 1, 3)
    Traffic!mkcPhone.SelStart = 6
    Traffic!mkcPhone.SelLength = 3
    Traffic!mkcPhone.SelText = Mid$(slPhone, 4, 3)
    Traffic!mkcPhone.SelStart = 10
    Traffic!mkcPhone.SelLength = 4
    Traffic!mkcPhone.SelText = Mid$(slPhone, 7, 4)
    If Len(slPhone) <= 10 Then
        mkcCtrl.Text = Traffic!mkcPhone.Text
        Exit Sub
    End If
    Traffic!mkcPhone.SelStart = 19
    Traffic!mkcPhone.SelLength = 4
    Traffic!mkcPhone.SelText = Mid$(slPhone, 11, 4)
    mkcCtrl.Text = Traffic!mkcPhone.Text
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gFormatPhoneNo                  *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Remove _ characters from number *
'*                                                     *
'*******************************************************
Function gFormatPhoneNo(slPhone As String) As String
    Dim slStr As String
    Dim ilPos As Integer
    Dim slTmp As String
    Dim slPhoneImage As String

    If Len(slPhone) <= 10 Then
        slPhoneImage = "(___) ___-____"
        gSetPhoneNo slPhone, Traffic!mkcTempFaxNo
        slStr = Traffic!mkcTempFaxNo.Text
    Else
        slPhoneImage = "(___) ___-____ Ext(____)"
        gSetPhoneNo slPhone, Traffic!mkcTempPhoneNo
        slStr = Traffic!mkcTempPhoneNo.Text
    End If
    If slStr = slPhoneImage Then
        slStr = ""
    Else
        If InStr(slStr, "(____)") <> 0 Then 'Test for missing extension
            ilPos = InStr(slStr, "Ext(")
            slStr = Left$(slStr, ilPos - 1)
        End If
        If InStr(slStr, "(___)") <> 0 Then  'Test for missing area code
            slStr = right$(slStr, Len(slStr) - 5)
        End If
        ilPos = InStr(slStr, "_")
        Do While ilPos > 0
            slTmp = Left$(slStr, ilPos - 1) & Mid$(slStr, ilPos + 1)
            slStr = slTmp
            ilPos = InStr(slStr, "_")
        Loop
    End If
    gFormatPhoneNo = slStr
End Function

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
Sub gCtrlGotFocus(Ctrl As Control)
'
'   gCtrlGotFocus Ctrl
'   Where:
'       Ctrl (I)- control for which text will be highlighted
'

'    Traffic.plcHelp.Caption = " " & Ctrl.Tag
    On Error GoTo gCtrlGotFocusErr:
    If TypeOf Ctrl Is TextBox Then
        Ctrl.SelStart = 0
        Ctrl.SelLength = Len(Ctrl.Text)
    ElseIf TypeOf Ctrl Is MaskEdBox Then
        Ctrl.SelStart = 0
        Ctrl.SelLength = Len(Ctrl.Text)
    ElseIf TypeOf Ctrl Is ComboBox Then
        Ctrl.SelStart = 0
        Ctrl.SelLength = Len(Ctrl.Text)
    End If
    Exit Sub
gCtrlGotFocusErr:
    Exit Sub
End Sub

'***************************************************************************************
'*
'*      Procedure Name:gGetSpotsByVefDate - obtain spots from SDF by Key1:
'*                                          Vehicle, Date, Time, SchStatus
'*      <input> hlSdf - SDF handle
'*              ilVefCode - vehicle code
'*              slStartDate - earliest date to retrieve spots
'*              slEndDate - latest date to retrieve spots
'*              tlSdf() - array of spots for the requested span
'*      <return>
'*
'*     Created:8/19/05       By:D. Hosaka
'*
'***************************************************************************************
Public Function gGetSpotsbyVefDate(hlSdf As Integer, ilVefCode As Integer, slStartDate As String, slEndDate As String, tlSdf() As SDF) As Integer
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffset As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim llChfCode As Long
    Dim tlSdfSrchKey1 As SDFKEY1
    Dim ilSdfRecLen As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    
    ReDim tlSdf(0 To 0) As SDF
    Dim ilUpper As Integer

    gGetSpotsbyVefDate = True
    btrExtClear hlSdf   'Clear any previous extend operation
    ilExtLen = Len(tlSdf(0))  'Extract operation record size
    tlSdfSrchKey1.iVefCode = ilVefCode
    gPackDate slStartDate, tlSdfSrchKey1.iDate(0), tlSdfSrchKey1.iDate(1)
    tlSdfSrchKey1.iTime(0) = 0
    tlSdfSrchKey1.iTime(1) = 0
    tlSdfSrchKey1.sSchStatus = ""   'slType
    ilSdfRecLen = Len(tlSdf(0))
    ilUpper = 0
    ilRet = btrGetGreaterOrEqual(hlSdf, tlSdf(0), ilSdfRecLen, tlSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
    If (tlSdf(ilUpper).iVefCode = ilVefCode) And (ilRet <> BTRV_ERR_END_OF_FILE) Then
    
        ' Prepare to execute an extended operation.
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hlSdf, llNoRec, -1, "UC", "SDF", "") '"EG") 'Set extract limits (all records)

        ' We only the records for the passed in vehicle code.
        ilOffset = gFieldOffset("Sdf", "SdfVefCode")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, ilVefCode, 2)
        
        
        ' And only records where the ChfCode = 0
        llChfCode = 0
        ilOffset = gFieldOffset("Sdf", "SdfChfCode")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, llChfCode, 4)
        
        gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffset = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

        
        ' And on the records where the date is between the passed  date
        gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffset = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        
        ilRet = btrExtAddField(hlSdf, 0, ilExtLen) 'Extract the whole record
        
        ilRet = btrExtGetNext(hlSdf, tlSdf(ilUpper), ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            ilExtLen = Len(tlSdf(0))  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSdf, tlSdf(ilUpper), ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                
                'tlSDF(ilUpper) = tlSDF
                ReDim Preserve tlSdf(0 To UBound(tlSdf) + 1) As SDF
                ilUpper = ilUpper + 1
                ilExtLen = Len(tlSdf(0))
                ilRet = btrExtGetNext(hlSdf, tlSdf(ilUpper), ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSdf, tlSdf(ilUpper), ilExtLen, llRecPos)
                Loop
                DoEvents
            Loop
        End If
    End If

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gCheckKeyasc                    *
'*                                                     *
'*             Created:4/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Check for illegal key           *
'*                                                     *
'*******************************************************
Function gCheckEnteredChar(ilKeyAscii As Integer) As Integer
    '92=Backslash (\);47=Forwardslash (/); 94=Caret (^); 124=Verical Bar (|); 91=Sq Bracket ([); 93=Sq Bracket (]); 59=Semi-colon(;)
    'If (ilKeyAscii < 32) Or (ilKeyAscii > 126) Or (ilKeyAscii = 92) Or (ilKeyAscii = 94) Or (ilKeyAscii = 124) Or (ilKeyAscii = 91) Or (ilKeyAscii = 93) Or (ilKeyAscii = 59) Then
    'Allow LF, CR and Backspace.
    '                 LF                     CR                     Backspace
    If (ilKeyAscii <> 10) And (ilKeyAscii <> 13) And (ilKeyAscii <> 8) Then
    
        '                                      "                    \                    /                    ^                    |                     [                    ]                    ;
        If (ilKeyAscii < 32) Or (ilKeyAscii > 126) Or (ilKeyAscii = 34) Or (ilKeyAscii = 92) Or (ilKeyAscii = 47) Or (ilKeyAscii = 94) Or (ilKeyAscii = 124) Or (ilKeyAscii = 91) Or (ilKeyAscii = 93) Or (ilKeyAscii = 59) Then
            Beep
            gCheckEnteredChar = False
            Exit Function
        End If
    End If
    gCheckEnteredChar = True
End Function
