Attribute VB_Name = "FORMSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of FORMSUBS.BAS on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  gNumOutLine                                                                           *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: FormSubs.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Form subs and functions
Option Explicit
Option Compare Text

Private Const CB_GETITEMHEIGHT = &H154
Private Declare Function MoveWindow Lib "user32" _
  (ByVal HWnd As Long, _
   ByVal X As Long, ByVal Y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal bRepaint As Long) As Long

Private Declare Function GetWindowRect Lib "user32" _
  (ByVal HWnd As Long, _
   lpRect As RECT) As Long

Private Declare Function ScreenToClient Lib "user32" _
  (ByVal HWnd As Long, _
   lpPoint As POINTAPI) As Long

Dim imCalMonth As Integer
Dim imCalYear As Integer
Public igDOE As Integer 'Used as a counter to perform a DOEVENTs every so often.  Doing a DOEVENTS too often causes performance issues.

'*******************************************************
'*                                                     *
'*      Procedure Name:gAdjShowLen                     *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Adjust length of string to fit *
'*                 within specified area minus margins *
'*                                                     *
'*******************************************************
Sub gAdjShowLen(Ctrl As Control, slShow As String, flWidth As Single)
'
'   gAdjShowLen Ctrl, sShow, BoxWidth
'   Where:
'       Ctrl (I)- Control where string will be painted
'       sShow (I/O)- String to be shown
'       ilBoxWidth (I)- width of control area
'

    Dim flPicWidth
    flPicWidth = flWidth - 2 * fgBoxInsetX
    Do While (Ctrl.TextWidth(slShow) > flPicWidth)
        slShow = Left$(slShow, Len(slShow) - 1)
    Loop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gAnyFieldChgd                   *
'*                                                     *
'*             Created:5/11/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test fields for change         *
'*                                                     *
'*******************************************************
Function gAnyFieldChgd(tmCtrls() As FIELDAREA, ilCtrlNo As Integer) As Integer
'
'   iTest = TESTALLCTRLS
'   iRet = gAnyFieldChgd(tmCtrls, iTest)
'   Where:
'       tmCtrls (I)- FieldArea control array (change flag set in the ---gSetShow subs)
'       iTest (I)- Test all controls or control number specified
'       iRet (O)- YES if any field changed, NO if no fields changed
'

    Dim ilLoop As Integer   'For control parameter
    Dim ilLBCtrls As Integer
    
    ilLBCtrls = LBound(tmCtrls)
    If ilLBCtrls = 0 Then
        ilLBCtrls = 1
    End If
    For ilLoop = ilLBCtrls To UBound(tmCtrls) Step 1
        If (ilCtrlNo = ilLoop) Or (ilCtrlNo = TESTALLCTRLS) Then
            If tmCtrls(ilLoop).iChg Then
                gAnyFieldChgd = YES
                Exit Function
            End If
        End If
    Next ilLoop
    gAnyFieldChgd = NO
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gCenterShowStr                  *
'*                                                     *
'*             Created:4/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine the CurrentX position *
'*                     so value is centered            *
'*                                                     *
'*******************************************************
Function gCenterShowStr(pbcCtrl As Control, slShowStr As String, tlCtrlArray As FIELDAREA) As Single
'
'   flX = gCenterShowStr(pbcCtrl, slShow, CtrlArray)
'   Where:
'       pbcCtrl (I) Control that will contain the picture
'       slShow (I)- String to be shown
'       CtrlArray (O)- field control array
'       flX (O)- CurrentX value determine to center value
'
    Dim flPicWidth As Single
    Dim flStrWidth As Single
    If Len(slShowStr) = 0 Then
        Exit Function
    End If
    flPicWidth = tlCtrlArray.fBoxW - fgBoxInsetX
    flStrWidth = pbcCtrl.TextWidth(slShowStr)
    If flStrWidth > flPicWidth Then
        gCenterShowStr = tlCtrlArray.fBoxX + fgBoxInsetX
    Else
        gCenterShowStr = tlCtrlArray.fBoxX + fgBoxInsetX + (flPicWidth - flStrWidth) / 2
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
Function gCheckKeyAscii(ilKeyAscii As Integer) As Integer
    '92=Backslash (\);47=Forwardslash (/); 94=Caret (^); 124=Verical Bar (|); 91=Sq Bracket ([); 93=Sq Bracket (]); 59=Semi-colon(;)
    'If (ilKeyAscii < 32) Or (ilKeyAscii > 126) Or (ilKeyAscii = 92) Or (ilKeyAscii = 94) Or (ilKeyAscii = 124) Or (ilKeyAscii = 91) Or (ilKeyAscii = 93) Or (ilKeyAscii = 59) Then
    'Allow LF, CR and Backspace.  Just remove test for < 32
    '                                      "                    \                    ^                    |                     [                    ]                    ;
    If (ilKeyAscii > 126) Or (ilKeyAscii = 34) Or (ilKeyAscii = 92) Or (ilKeyAscii = 94) Or (ilKeyAscii = 124) Or (ilKeyAscii = 91) Or (ilKeyAscii = 93) Or (ilKeyAscii = 59) Then
        Beep
        gCheckKeyAscii = False
        Exit Function
    End If
    gCheckKeyAscii = True
End Function
Function gCheckKeyAsciiIncludeSlash(ilKeyAscii As Integer) As Integer
    '92=Backslash (\);47=Forwardslash (/); 94=Caret (^); 124=Verical Bar (|); 91=Sq Bracket ([); 93=Sq Bracket (]); 59=Semi-colon(;); 35=pound (#); 37=Percent (%)
    'If (ilKeyAscii < 32) Or (ilKeyAscii > 126) Or (ilKeyAscii = 92) Or (ilKeyAscii = 94) Or (ilKeyAscii = 124) Or (ilKeyAscii = 91) Or (ilKeyAscii = 93) Or (ilKeyAscii = 59) Then
    'Allow LF, CR and Backspace.  Just remove test for < 32
    '                                      "                    \                    /                    ^                    |                     [                    ]                    ;                     #                   %                    +
    If (ilKeyAscii > 126) Or (ilKeyAscii = 34) Or (ilKeyAscii = 92) Or (ilKeyAscii = 47) Or (ilKeyAscii = 94) Or (ilKeyAscii = 124) Or (ilKeyAscii = 91) Or (ilKeyAscii = 93) Or (ilKeyAscii = 59) Or (ilKeyAscii = 35) Or (ilKeyAscii = 37) Or (ilKeyAscii = 43) Then
        Beep
        gCheckKeyAsciiIncludeSlash = False
        Exit Function
    End If
    gCheckKeyAsciiIncludeSlash = True
End Function

'10137
Public Function gReplaceIllegalCharactersAllowSlashes(sInStr As String) As String
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
                If gCheckKeyAsciiAllowSlashes(Asc(sChar)) = 0 Then
                    sOutStr = sOutStr & " "
                Else
                    sOutStr = sOutStr & sChar
                End If
            Next iLoop
        Else
            sOutStr = sInStr
        End If
    End If
    gReplaceIllegalCharactersAllowSlashes = sOutStr
End Function
'10137
Function gCheckKeyAsciiAllowSlashes(ilKeyAscii As Integer) As Integer
    ' 94=Caret (^); 124=Verical Bar (|); 91=Sq Bracket ([); 93=Sq Bracket (]); 59=Semi-colon(;); 35=pound (#); 37=Percent (%)
    'allow 92=Backslash (\);47=Forwardslash (/);
    If (ilKeyAscii > 126) Or (ilKeyAscii = 34) Or (ilKeyAscii = 94) Or (ilKeyAscii = 124) Or (ilKeyAscii = 91) Or (ilKeyAscii = 93) Or (ilKeyAscii = 59) Or (ilKeyAscii = 35) Or (ilKeyAscii = 37) Or (ilKeyAscii = 43) Then
        Beep
        gCheckKeyAsciiAllowSlashes = False
        Exit Function
    End If
    gCheckKeyAsciiAllowSlashes = True
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

Public Function gFilterNameMatchingKeyPressCheck(slOldName As String, Optional slExclude As String = "") As String
    'Code taken from affiliate mSafeFileName
    '1/9/19: Changed to match the function gCheckKeyAsciiIncludeSlash
    Dim slTempName As String
    slTempName = slOldName
    'If InStr(1, slExclude, "?") <= 0 Then slTempName = Replace(slTempName, "?", "-")
    If InStr(1, slExclude, "/") <= 0 Then slTempName = Replace(slTempName, "/", "-")
    If InStr(1, slExclude, "\") <= 0 Then slTempName = Replace(slTempName, "\", "-")
    'If InStr(1, slExclude, "%") <= 0 Then slTempName = Replace(slTempName, "%", "-")
    'If InStr(1, slExclude, "*") <= 0 Then slTempName = Replace(slTempName, "*", "-")
    'If InStr(1, slExclude, ":") <= 0 Then slTempName = Replace(slTempName, ":", "-")
    If InStr(1, slExclude, "|") <= 0 Then slTempName = Replace(slTempName, "|", "-")
    If InStr(1, slExclude, """") <= 0 Then slTempName = Replace(slTempName, """", "-")
    'If InStr(1, slExclude, ".") <= 0 Then slTempName = Replace(slTempName, ".", "-")
    'If InStr(1, slExclude, "<") <= 0 Then slTempName = Replace(slTempName, "<", "-")
    'If InStr(1, slExclude, ">") <= 0 Then slTempName = Replace(slTempName, ">", "-")
    If InStr(1, slExclude, "^") <= 0 Then slTempName = Replace(slTempName, "^", "-")
    If InStr(1, slExclude, "[") <= 0 Then slTempName = Replace(slTempName, "[", "-")
    If InStr(1, slExclude, "]") <= 0 Then slTempName = Replace(slTempName, "]", "-")
    If InStr(1, slExclude, ";") <= 0 Then slTempName = Replace(slTempName, ";", "-")
    gFilterNameMatchingKeyPressCheck = slTempName
End Function
Public Function gRemoveIllegalPastedChar(slOldName As String, Optional slExclude As String = "") As String
    'Code taken from affiliate mSafeFileName
    '1/9/19: Changed to match the function gCheckKeyAsciiIncludeSlash
    Dim slTempName As String
    slTempName = slOldName
    If Trim$(slTempName) = "[New]" Or Trim$(slTempName) = "[None]" Or Trim$(slTempName) = "N/A" Or Trim$(slTempName) = "[N/A]" Then
        gRemoveIllegalPastedChar = slTempName
        Exit Function
    End If
    'If InStr(1, slExclude, "?") <= 0 Then slTempName = Replace(slTempName, "?", "-")
    If InStr(1, slExclude, "/") <= 0 Then slTempName = Replace(slTempName, "/", "")
    If InStr(1, slExclude, "\") <= 0 Then slTempName = Replace(slTempName, "\", "")
    'If InStr(1, slExclude, "%") <= 0 Then slTempName = Replace(slTempName, "%", "-")
    'If InStr(1, slExclude, "*") <= 0 Then slTempName = Replace(slTempName, "*", "-")
    'If InStr(1, slExclude, ":") <= 0 Then slTempName = Replace(slTempName, ":", "-")
    If InStr(1, slExclude, "|") <= 0 Then slTempName = Replace(slTempName, "|", "")
    If InStr(1, slExclude, """") <= 0 Then slTempName = Replace(slTempName, """", "")
    'If InStr(1, slExclude, ".") <= 0 Then slTempName = Replace(slTempName, ".", "-")
    'If InStr(1, slExclude, "<") <= 0 Then slTempName = Replace(slTempName, "<", "-")
    'If InStr(1, slExclude, ">") <= 0 Then slTempName = Replace(slTempName, ">", "-")
    If InStr(1, slExclude, "^") <= 0 Then slTempName = Replace(slTempName, "^", "")
    If InStr(1, slExclude, "[") <= 0 Then slTempName = Replace(slTempName, "[", "")
    If InStr(1, slExclude, "]") <= 0 Then slTempName = Replace(slTempName, "]", "")
    If InStr(1, slExclude, ";") <= 0 Then slTempName = Replace(slTempName, ";", "")
    If InStr(1, slExclude, "#") <= 0 Then slTempName = Replace(slTempName, "#", "")
    If InStr(1, slExclude, "%") <= 0 Then slTempName = Replace(slTempName, "%", "")
    If InStr(1, slExclude, "+") <= 0 Then slTempName = Replace(slTempName, "+", "")
    gRemoveIllegalPastedChar = slTempName
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

'*******************************************************
'*                                                     *
'*      Procedure Name:gFieldDefinedCtrl               *
'*                                                     *
'*             Created:4/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if specified control  *
'*                     contains input (not blank or    *
'*                     equal mask)                     *
'*                                                     *
'*******************************************************
Function gFieldDefinedCtrl(lCtrl As Control, slStr As String, slMess As String, ilReq As Integer, ilState As Integer) As Integer
'
'   ilRet = gFieldDefinedCtrl (lCtrl, slStr, sMess, iReq, iState)
'   Where:
'       lCtrl (I)- control to be checked
'       slStr (I)- Blank or mask string with blanks
'       sMess (I)- Display message if required field is not defined
'       iReq (I)- True = input required; False = input not required
'       iState (I)- Type of testing:
'                  ALLBLANK=All fields blank
'                  ALLMANBLANK=All mandatory field blank
'                  ALLMANDEFINED=All mandatory fields have data
'                  Plus
'                  NOMSG=No error message shown
'                  SHOWMSG=show error message
'       ilRet (O) - True if the field meets criteria specified above
'                    False if field does not meet specified criteria
'

    Dim slCtrl As String
    slCtrl = lCtrl.Text
    gFieldDefinedCtrl = gFieldDefinedStr(slCtrl, slStr, slMess, ilReq, ilState)
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gFieldDefinedStr                *
'*                                                     *
'*             Created:4/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if specified string   *
'*                     contains input (not blank or    *
'*                     equal mask)                     *
'*                                                     *
'*******************************************************
Function gFieldDefinedStr(slCtrl As String, slStr As String, slMess As String, ilReq As Integer, ilState As Integer) As Integer
'
'   ilRet = gFieldDefinedStr (sStrCtrl, slStr, sMess, iReq, iState)
'   Where:
'       sStrCtrl (I)- string from control to be checked
'       slStr (I)- Blank or mask string with blanks
'       sMess (I)- Display message if required field is not defined
'       iReq (I)- True = input required; False = input not required
'       iState (I)- Type of testing:
'                  ALLBLANK=All fields blank
'                  ALLMANBLANK=All mandatory field blank
'                  ALLMANDEFINED=All mandatory fields have data
'                  Plus
'                  NOMSG=No error message shown
'                  SHOWMSG=show error message
'        ilRet (O) - True if the field meets criteria specified above
'                    False if field does not meet specified criteria
'

    Dim ilRes As Integer
    If (Trim$(slCtrl) = slStr) Or (InStr(1, slCtrl, "[New]", vbTextCompare) > 0) Then
        If ilReq Then
            If (ilState And SHOWMSG) = SHOWMSG Then
                ilRes = MsgBox(slMess, vbOKOnly + vbExclamation, "Incomplete")
            End If
            If (ilState And ALLMANDEFINED) = ALLMANDEFINED Then
                gFieldDefinedStr = NO
                Exit Function
            End If
        End If
    Else
        If (ilState And ALLBLANK) = ALLBLANK Then
            gFieldDefinedStr = NO
            Exit Function
        End If
        If ilReq And ((ilState And ALLMANBLANK) = ALLMANBLANK) Then
            gFieldDefinedStr = NO
            Exit Function
        End If
    End If
    gFieldDefinedStr = YES

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
                If ilPass = 0 Then
                    If StrComp(Trim$(lbcList.List(gLastFound(lbcList))), slMatch, 1) = 0 Then
                        Exit Sub
                    Else
                        ilFound = False
                        gFndNext lbcList, slMatch
                    End If
                Else
                    If StrComp(RTrim$(lbcList.List(gLastFound(lbcList))), slMatch, 1) = 0 Then
                        Exit Sub
                    Else
                        ilFound = False
                        gFndNext lbcList, slMatch
                    End If
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

'*******************************************************
'*                                                     *
'*      Procedure Name:gFindPartialMatch               *
'*                                                     *
'*             Created:10/20/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Find matching match within list*
'*                      box                            *
'*                                                     *
'*******************************************************
Sub gFindPartialMatch(slInMatch As String, ilFirstAllowedIndex As Integer, ilMatchLen As Integer, lbcList As Control)
'
'   gFindPartialMatch slInMatch, ilFirstAllowedIndex, ilMatchLen,  lbcList
'   Where:
'       slInMatch(I)- String to be found
'       ilFirstAllowedIndex(I)- First legal index
'       ilMatchLen(I)- Length of string to be matched
'       lbcList(I)- list box containing strings
'
'       lbcList.LastFound(O)- Found index
'
    Dim ilFound As Integer
    Dim slMatch As String
    Dim ilPass As Integer   'Pass: 0= Trim string; 1=Leave leading blanks

    For ilPass = 0 To 1 Step 1
        ilFound = False
        If ilPass = 0 Then
            slMatch = Left$(Trim$(slInMatch), ilMatchLen)
        Else
            slMatch = Left$(RTrim$(slInMatch), ilMatchLen)
        End If
        gFndFirst lbcList, slMatch
        Do
            If gLastFound(lbcList) >= ilFirstAllowedIndex Then
                If ilPass = 0 Then
                    If StrComp(Trim$(Left$(lbcList.List(gLastFound(lbcList)), ilMatchLen)), slMatch, 1) = 0 Then
                        Exit Sub
                    Else
                        ilFound = False
                        gFndNext lbcList, slMatch
                    End If
                Else
                    If StrComp(RTrim$(Left$(lbcList.List(gLastFound(lbcList)), ilMatchLen)), slMatch, 1) = 0 Then
                        Exit Sub
                    Else
                        ilFound = False
                        gFndNext lbcList, slMatch
                    End If
                End If
            Else
                If ilPass = 1 Then
                    gFndNext lbcList, "|~|~|~|~|~|~|~|~|"   'Invaild value so LastFound is set to -1
                    Exit Sub
                Else
                    Exit Do
                End If
            End If
        Loop While Not ilFound
    Next ilPass
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
'*      Procedure Name:gGetPhoneNo                     *
'*                                                     *
'*             Created:4/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Remove mask from phone number   *
'*                     (Works with or without          *
'*                     extension)                      *
'*                                                     *
'*******************************************************
Sub gGetPhoneNo(mkcCtrl As Control, slPhone As String)
'
'   gGetPhoneNo mkcCtrl, sPhone
'   Where:
'       mkcCtrl (I)- Mask edit control containing phone number
'       sPhone (O)- Phone number without literals
'

    Dim ilSelStart As Integer
    Dim ilSelLength As Integer

    If Len(mkcCtrl.Text) <= 14 Then
        Traffic!mkcPhone.mask = "(AAA) AAA-AAAA"
        Traffic!mkcPhone.MaxLength = 14
    Else
        Traffic!mkcPhone.mask = "(AAA) AAA-AAAA Ext(AAAA)"
        Traffic!mkcPhone.MaxLength = 24
    End If
    Traffic!mkcPhone.Text = mkcCtrl.Text
    ilSelStart = mkcCtrl.SelStart
    ilSelLength = mkcCtrl.SelLength
    Traffic!mkcPhone.SelStart = 1
    Traffic!mkcPhone.SelLength = 3
    slPhone = Traffic!mkcPhone.SelText
    Traffic!mkcPhone.SelStart = 6
    Traffic!mkcPhone.SelLength = 3
    slPhone = slPhone & Traffic!mkcPhone.SelText
    Traffic!mkcPhone.SelStart = 10
    Traffic!mkcPhone.SelLength = 4
    slPhone = slPhone & Traffic!mkcPhone.SelText
    If Len(mkcCtrl.Text) <= 14 Then
        mkcCtrl.SelStart = ilSelStart
        mkcCtrl.SelLength = ilSelLength
        Exit Sub
    End If
    Traffic!mkcPhone.SelStart = 19
    Traffic!mkcPhone.SelLength = 4
    slPhone = slPhone & Traffic!mkcPhone.SelText
    mkcCtrl.SelStart = ilSelStart
    mkcCtrl.SelLength = ilSelLength
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetSocialSecurityNo            *
'*                                                     *
'*             Created:4/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Remove mask from social security*
'*                                                     *
'*******************************************************
Sub gGetSocialSecurityNo(mkcCtrl As Control, slSSNo As String)
'
'   gGetSocialSecurityNo mkcCtrl, sSSNo
'   Where:
'       mkcCtrl (I)- Mask edit control containing phone number
'       sSSNo (O)- Social security number without literals
'

    Dim ilSelStart As Integer
    Dim ilSelLength As Integer

    ilSelStart = mkcCtrl.SelStart
    ilSelLength = mkcCtrl.SelLength
    mkcCtrl.SelStart = 0
    mkcCtrl.SelLength = 3
    slSSNo = mkcCtrl.SelText
    mkcCtrl.SelStart = 4
    mkcCtrl.SelLength = 2
    slSSNo = slSSNo & mkcCtrl.SelText
    mkcCtrl.SelStart = 7
    mkcCtrl.SelLength = 4
    slSSNo = slSSNo & mkcCtrl.SelText
    mkcCtrl.SelStart = ilSelStart
    mkcCtrl.SelLength = ilSelLength
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:gInvalidateArea                 *
'*                                                     *
'*             Created:7/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Invalidates an area.            *
'*                     Note: a paint event is not      *
'*                     generated (a call to            *
'*                     pbcCtrl_Paint is required to    *
'*                     get the area repainted)         *
'*                                                     *
'*******************************************************
Sub gInvalidateArea(pbcCtrl As Control, ilCurrX As Integer, ilCurrY As Integer, ilWidth As Integer, ilHeight As Integer)
'
'   gInvalidateArea pbcCtrl, lX, lY, lW, lH
'   Where:
'       pbcCtrl (I)- Control of area to invaildate
'       lX (I)- Left position of invalid area
'       lY (I)- Top position of invalid area
'       lH (I)- Height of invalid area
'       lW (I)- Width of invalid area
'

    Dim tlRect As RECT
    Dim llRet As Long

    llRet = SetRect(tlRect, ilCurrX, ilCurrY, ilCurrX + ilWidth, ilCurrY + ilHeight)
    llRet = InvalidateRect(pbcCtrl.HWnd, tlRect, True)
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gInvertArea                     *
'*                                                     *
'*             Created:7/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Invert an area.                 *
'*                     Note: invert the value of each  *
'*                     pixel                           *
'*                                                     *
'*******************************************************
Sub gInvertArea(pbcCtrl As Control, ilCurrX As Integer, ilCurrY As Integer, ilWidth As Integer, ilHeight As Integer)
'
'   gInvertArea pbcCtrl, lX, lY, lW, lH
'   Where:
'       pbcCtrl (I)- Control of area to inverted
'       lX (I)- Left position of invalid area
'       lY (I)- Top position of invalid area
'       lW (I)- Width of invalid area
'       lH (I)- Height of invalid area
'

    Dim tlRect As RECT
    Dim llRet As Long

    llRet = SetRect(tlRect, ilCurrX, ilCurrY, ilCurrX + ilWidth, ilCurrY + ilHeight)
    llRet = InvertRect(pbcCtrl.hdc, tlRect)
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gLengthOutLine                  *
'*                                                     *
'*             Created:10/20/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Position outline box around    *
'*                      key pressed by the user in     *
'*                      the length input pad           *
'*                                                     *
'*******************************************************
Sub gLengthOutLine(ilKeyAscii As Integer, imcCtrl As Control)
'
'  gLengthOutLine ilKeyAscii, imcCtrl
'   Where:
'       ilKeyAscii (I) - standard numericANSI keycode obtained from the                                 '               KeyPress event
'       imcCtrl (I) - name of the outline control
'

    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    If ilKeyAscii <> KEYBACKSPACE Then
        Select Case UCase(Chr(ilKeyAscii))
            Case "H"
                ilRowNo = 1
                ilColNo = 1
            Case "M"
                ilRowNo = 1
                ilColNo = 2
            Case "S"
                ilRowNo = 1
                ilColNo = 3
            Case "7"
                ilRowNo = 2
                ilColNo = 1
            Case "8"
                ilRowNo = 2
                ilColNo = 2
            Case "9"
                ilRowNo = 2
                ilColNo = 3
            Case "4"
                ilRowNo = 3
                ilColNo = 1
            Case "5"
                ilRowNo = 3
                ilColNo = 2
            Case "6"
                ilRowNo = 3
                ilColNo = 3
            Case "1"
                ilRowNo = 4
                ilColNo = 1
            Case "2"
                ilRowNo = 4
                ilColNo = 2
            Case "3"
                ilRowNo = 4
                ilColNo = 3
            Case "0"
                ilRowNo = 5
                ilColNo = 1
            Case "00"   'Not possible
                ilRowNo = 5
                ilColNo = 2
        End Select
        flX = fgPadMinX + (ilColNo - 1) * fgPadDeltaX
        flY = fgPadMinY + (ilRowNo - 1) * fgPadDeltaY
        imcCtrl.Move flX - 15, flY - 15
        imcCtrl.Visible = True
    Else
        imcCtrl.Visible = False
    End If
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
Function gListBoxHeight(ByVal ilNoRows As Integer, ilMaxRow As Integer) As Single
'
'  flHeight = gListBoxHeight (ilNoRows, ilMaxRows)
'   Where:
'       ilNoRows (I) - current number of items within the list box
'       ilMaxRows (I) - max number of list box items to be displayed
'       flHeight (O) - height of list box in twips
'
    '+30 because of line above and below
    If ilNoRows <= 0 Then
        gListBoxHeight = (fgListHtArial825) * ilMaxRow + 30 '375 + 255 * (ilMaxRow - 1)
    Else
        If ilNoRows <= ilMaxRow Then
            gListBoxHeight = (fgListHtArial825) * ilNoRows + 30 '375 + 255 * (ilNoRows - 1)
        Else
            gListBoxHeight = (fgListHtArial825) * ilMaxRow + 30 '375 + 255 * (ilMaxRow - 1)
        End If
    End If
End Function

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
Sub gManLookAhead(cbcComboBox As ComboBox, ilBSMode As Integer, ilErrHighLightIndex As Integer)
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
'*                     to see if the list box contains *
'*                     any matching string to the input*
'*                     string.  This is used for combo *
'*                     boxes made up of text control/  *
'*                     button control/list box control.*
'*                                                     *
'*******************************************************
Sub gMatchLookAhead(edcTextBox As TextBox, lbcListBox As ListBox, ilBSMode As Integer, ilErrHighLightIndex As Integer)
'
'   gMatchLookAhead edcText, lbcCtrl, ilBSMode, ilHighlightIndex
'   Where:
'       edcText (I)- Text box control (containing input to be validated)
'       lbcCtrl (I)- List box control containing values to be matched
'       ilBSMode (I/O)- Backspace flag(True = backspace key was pressed, False =                        '       backspace key was not pressed)
'       ilHighlightIndex (I)- Selection to be highlighted if input is invalid
'

    Dim ilLen As Integer    'Length of current enter text
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    Dim ilSelStart As Integer
    Dim ilBracket As Integer
    Dim ilSearch As Integer
    Dim ilSvLastFound As Integer
    Dim ilLoop As Integer
    Dim ilPos As Integer

    ilSelStart = edcTextBox.SelStart
    slStr = LTrim$(edcTextBox.Text)    'Remove leading blanks only
    ilLen = Len(edcTextBox.Text)
    ilIndex = lbcListBox.ListIndex
    If slStr = "" Then  'If space bar selected, text will be blank- ListIndex will contain a value
        lbcListBox.ListIndex = -1
        Exit Sub
        If lbcListBox.ListIndex >= 0 Then
            slStr = lbcListBox.List(lbcListBox.ListIndex)
            ilLen = 0
            ilIndex = -1    'Force dispaly of selected item by space bar
        Else
            Beep
            If ilErrHighLightIndex >= 0 Then
                lbcListBox.ListIndex = ilErrHighLightIndex
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
        For ilLoop = 0 To lbcListBox.ListCount - 1 Step 1
            ilPos = InStr(1, lbcListBox.List(ilLoop), slStr, 1)
            If ilPos = 1 Then
                ilSvLastFound = ilLoop
                Exit For
            Else
                If Left$(lbcListBox.List(ilLoop), 1) <> "[" Then
                    Exit For
                End If
            End If
        Next ilLoop
    Else
        'Test if matching string found in the combo box- if so display it (look ahead typing)
        'lbcListBox.ListIndex = 0
        gFndFirst lbcListBox, slStr
        ilBracket = False
        Do
            If gLastFound(lbcListBox) >= 0 Then
                If (Left$(lbcListBox.List(gLastFound(lbcListBox)), 1) = "[") And (Left$(slStr, 1) <> "[") Then
                    gFndNext lbcListBox, slStr
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
        ilSvLastFound = gLastFound(lbcListBox)
        Do
            If gLastFound(lbcListBox) >= 0 Then
                If StrComp(slStr, lbcListBox.List(gLastFound(lbcListBox)), 1) = 0 Then
                    ilSvLastFound = gLastFound(lbcListBox)   'lbcListBox.LastFound
                    Exit Do
                End If
                gFndNext lbcListBox, slStr
            Else
                Exit Do
            End If
        Loop While ilSearch
    End If
    If ilSvLastFound >= 0 Then
        'If item found not same as current selected- change current
        If (ilIndex <> ilSvLastFound) Or ((ilIndex = ilSvLastFound) And Not ilBSMode) Then
            lbcListBox.ListIndex = ilSvLastFound 'This will cause a change event (reason for imChgMode)
        End If
'        If (ilIndex <> lbcListBox.LastFound) Or ((ilIndex = lbcListBox.LastFound) And Not ilBSMode) Then 'If indices not equal- highlight look ahead text
'            lbcListBox.SelStart = ilLen
'            lbcListBox.SelLength = Len(lbcListBox.Text)
'        End If
        ilErrHighLightIndex = ilSvLastFound
    Else
        Beep
        If (ilErrHighLightIndex >= 0) And (ilErrHighLightIndex < lbcListBox.ListCount) Then
            lbcListBox.ListIndex = ilErrHighLightIndex
            ilSelStart = 0
        End If
    End If
    If lbcListBox.ListIndex >= 0 Then
        edcTextBox.Text = lbcListBox.List(lbcListBox.ListIndex)
    Else
        edcTextBox.Text = lbcListBox.Text
    End If
    If ilSelStart <= Len(edcTextBox.Text) Then
        edcTextBox.SelStart = ilSelStart
        edcTextBox.SelLength = Len(edcTextBox.Text)
    Else
        edcTextBox.SelStart = 0
        edcTextBox.SelLength = Len(edcTextBox.Text)
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gMoveCtrl                       *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Move control within picture     *
'*                     control. Contol is measured from*
'*                     leftmost corner of the form,    *
'*                     not the picture box.            *
'*                                                     *
'*******************************************************
Sub gMoveCtrl(pbcCtrl As Control, Ctrl As Control, flBoxX As Single, flBoxY As Single)
'
'   gMoveCtrl pbcCtrl, InputCtrl, fBoxX, fBoxY
'   Where:
'       pbcCtrl (I)- Picture box control that will contain the input control (InputCtrl)
'       InputCtrl (I)- Input control that is to be moved
'       fBoxX (I)- x offset within pbcCtrl to place Contorl (InputCtrl)
'       fBoxY (I)- y offset within pbcCtrl to place Contorl (InputCtrl)
'

    If flBoxX + Ctrl.Width > pbcCtrl.Width Then
        Ctrl.Move pbcCtrl.Left + pbcCtrl.Width - Ctrl.Width - 5, fgOffset + pbcCtrl.Top + flBoxY
    Else
        Ctrl.Move pbcCtrl.Left + flBoxX, fgOffset + pbcCtrl.Top + flBoxY
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gMoveCtrl                       *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Move control within picture     *
'*                     control that represents a form  *
'*                     input area.                     *
'*                                                     *
'*******************************************************
Sub gMoveFormCtrl(pbcCtrl As Control, Ctrl As Control, flBoxX As Single, flBoxY As Single)
'
'   gMoveFormCtrl pbcCtrl, InputCtrl, fBoxX, fBoxY              '   Where:
'       pbcCtrl (I) Control that will contain the input control (InputCtrl)
'       InputCtrl (I)- Input control that is to be moved
'       fBoxX (I)- x offset within pbcCtrl (x coordinate of mouse area)
'       fBoxY (I)- y offset within pbcCtrl (y coordinate of mouse area)
'

    If flBoxX + Ctrl.Width > pbcCtrl.Width Then
        Ctrl.Move pbcCtrl.Left + pbcCtrl.Width - Ctrl.Width - 5, pbcCtrl.Top + flBoxY + fgBoxStH - Ctrl.height
    Else
        Ctrl.Move pbcCtrl.Left + flBoxX, pbcCtrl.Top + flBoxY + fgBoxStH - Ctrl.height
    End If
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:gMoveGridCtrl                   *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Move control within picture     *
'*                     control that represents a table *
'*                     input area (grid)               *
'*                                                     *
'*******************************************************
Sub gMoveTableCtrl(pbcCtrl As Control, Ctrl As Control, flBoxX As Single, flBoxY As Single)
'
'   gMoveTableCtrl pbcCtrl, InputCtrl, fBoxX, fBoxY
'   Where:
'       pbcCtrl (I) Control that will contain the input control
'       InputCtrl (I)- Input control that is to be moved
'       fBoxX (I)- x offset within pbcCtrl (x coordinate of mouse area)
'       fBoxY (I)- y offset within pbcCtrl (y coordinate of mouse area)
'

    Dim flAdjY As Single
    Dim flAdjX As Single

    If pbcCtrl.BorderStyle = 0 Then
        flAdjY = -15
        flAdjX = 0
    Else
        flAdjY = 0
        flAdjX = 15
    End If
    If flBoxX + Ctrl.Width > pbcCtrl.Width Then
        Ctrl.Move pbcCtrl.Left + pbcCtrl.Width - Ctrl.Width, pbcCtrl.Top + flBoxY + flAdjY
    Else
        Ctrl.Move pbcCtrl.Left + flBoxX + flAdjX, pbcCtrl.Top + flBoxY + flAdjY
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gNumOutLine                     *
'*                                                     *
'*             Created:10/20/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Position outline box around    *
'*                      key pressed by the user in the *
'*                      number input pad               *
'*                                                     *
'*******************************************************
Sub gNumOutLine(ilKeyAscii As Integer, imcCtrl As Control) 'VBC NR
'
'  gNumOutLine ilKeyAscii, imcCtrl
'   Where:
'       ilKeyAscii (I) - standard numericANSI keycode obtained from the                                 '               KeyPress event
'       imcCtrl (I) - name of the outline control
'
    Dim ilRowNo As Integer 'VBC NR
    Dim ilColNo As Integer 'VBC NR
    Dim flX As Single 'VBC NR
    Dim flY As Single 'VBC NR
    If ilKeyAscii <> KEYBACKSPACE Then 'VBC NR
        Select Case Chr(ilKeyAscii) 'VBC NR
            Case "7" 'VBC NR
                ilRowNo = 1 'VBC NR
                ilColNo = 1 'VBC NR
            Case "8" 'VBC NR
                ilRowNo = 1 'VBC NR
                ilColNo = 2 'VBC NR
            Case "9" 'VBC NR
                ilRowNo = 1 'VBC NR
                ilColNo = 3 'VBC NR
            Case "4" 'VBC NR
                ilRowNo = 2 'VBC NR
                ilColNo = 1 'VBC NR
            Case "5" 'VBC NR
                ilRowNo = 2 'VBC NR
                ilColNo = 2 'VBC NR
            Case "6" 'VBC NR
                ilRowNo = 2 'VBC NR
                ilColNo = 3 'VBC NR
            Case "1" 'VBC NR
                ilRowNo = 3 'VBC NR
                ilColNo = 1 'VBC NR
            Case "2" 'VBC NR
                ilRowNo = 3 'VBC NR
                ilColNo = 2 'VBC NR
            Case "3" 'VBC NR
                ilRowNo = 3 'VBC NR
                ilColNo = 3 'VBC NR
            Case "0" 'VBC NR
                ilRowNo = 4 'VBC NR
                ilColNo = 1 'VBC NR
            Case "00"   'Not possible 'VBC NR
                ilRowNo = 4 'VBC NR
                ilColNo = 2 'VBC NR
            Case "." 'VBC NR
                ilRowNo = 4 'VBC NR
                ilColNo = 3 'VBC NR
        End Select 'VBC NR
        flX = fgPadMinX + (ilColNo - 1) * fgPadDeltaX 'VBC NR
        flY = fgPadMinY + (ilRowNo - 1) * fgPadDeltaY 'VBC NR
        imcCtrl.Move flX - 15, flY - 15 'VBC NR
        imcCtrl.Visible = True 'VBC NR
    Else 'VBC NR
        imcCtrl.Visible = False 'VBC NR
    End If 'VBC NR
End Sub 'VBC NR

'*******************************************************
'*                                                     *
'*      Procedure Name:gOptionalLookAhead              *
'*                                                     *
'*             Created:9/9/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Validates combo box input and,  *
'*                     if invalid, retains input.      *
'*                     This is used for combo boxes    *
'*                     made up of text control/button  *
'*                     control/list box control.       *
'*                                                     *
'*******************************************************
Function gOptionalLookAhead(edcTextBox As TextBox, lbcListBox As ListBox, ilBSMode As Integer, slShown As String) As Integer
'
'   ilRet = gOptionalLookAhead (edcText, lbcCtrl, ilBSMode, slShown)
'   Where:
'       edcText (I)- Text box control (containing input to be validated)
'       lbcCtrl (I)- List box control containing values to be matched
'       ilBSMode (I/O)- Backspace flag(True = backspace key was pressed, False = backspace key was not pressed)
'       slStr (O)- String shown
'       ilRet (O)- 0=Selection found within list and lbcListBox.ListIndex set to the found selection
'                  1=No match found and search string was null ("")
'                  2=No match found and search string was not null
'

    Dim ilLen As Integer    'Length of current enter text
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    Dim ilSelStart As Integer
    Dim ilReturn As Integer
    Dim ilBracket As Integer
    Dim ilSearch As Integer
    Dim ilSvLastFound As Integer
    Dim ilLoop As Integer
    Dim ilPos As Integer




    ilSelStart = edcTextBox.SelStart
    slStr = LTrim$(edcTextBox.Text)    'Remove leading blanks only
    ilLen = Len(edcTextBox.Text)
    ilIndex = lbcListBox.ListIndex
    If slStr = "" Then  'If space bar selected, text will be blank- ListIndex will contain a value
'Remove test to see if behavior is OK-user removes text
'        If lbcListBox.ListIndex >= 0 Then
'            slStr = lbcListBox.List(lbcListBox.ListIndex)
'            ilLen = Len(lbcListBox.Text)
'            ilIndex = -1
'        Else
            gOptionalLookAhead = 1
            Exit Function
'        End If
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
        For ilLoop = 0 To lbcListBox.ListCount - 1 Step 1
            ilPos = InStr(1, lbcListBox.List(ilLoop), slStr, 1)
            If ilPos = 1 Then
                ilSvLastFound = ilLoop
                If StrComp(lbcListBox.List(ilLoop), "[New]", 1) = 0 Then
                    slShown = slStr
                    gOptionalLookAhead = 2
                    Exit Function
                End If
                Exit For
            Else
                If Left$(lbcListBox.List(ilLoop), 1) <> "[" Then
                    Exit For
                End If
            End If
        Next ilLoop
    Else
        'Test if matching string found in the combo box- if so display it (look ahead typing)
        gFndFirst lbcListBox, slStr
        ilBracket = False
        Do
            If gLastFound(lbcListBox) >= 0 Then
                If (Left$(lbcListBox.List(gLastFound(lbcListBox)), 1) = "[") And (Left$(slStr, 1) <> "[") Then
                    gFndNext lbcListBox, slStr
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
        ilSvLastFound = gLastFound(lbcListBox)
        Do
            If gLastFound(lbcListBox) >= 0 Then
                If StrComp(slStr, lbcListBox.List(gLastFound(lbcListBox)), 1) = 0 Then
                    ilSvLastFound = gLastFound(lbcListBox)   'lbcListBox.LastFound
                    Exit Do
                End If
                gFndNext lbcListBox, slStr
            Else
                Exit Do
            End If
        Loop While ilSearch
    End If
    If ilSvLastFound >= 0 Then
        'If item found not same as current selected- change current
        If (ilIndex <> ilSvLastFound) Or ((ilIndex = ilSvLastFound) And Not ilBSMode) Then
            lbcListBox.ListIndex = ilSvLastFound 'This will cause a change event (reason for imChgMode)
        End If
'        If (ilIndex <> lbcListBox.LastFound) Or ((ilIndex = lbcListBox.LastFound) And Not ilBSMode) Then 'If indices not equal- highlight look ahead text
'            lbcListBox.SelStart = ilLen
'            lbcListBox.SelLength = Len(lbcListBox.Text)
'        End If
        slShown = lbcListBox.Text
        ilReturn = 0
    Else
        slShown = slStr
        If slStr = "" Then
            ilReturn = 1
        Else
            ilReturn = 2
        End If
    End If
'    If lbcListBox.ListIndex >= 0 Then
'        edcTextBox.Text = lbcListBox.List(lbcListBox.ListIndex)
'    Else
    If StrComp(edcTextBox.Text, slShown, vbTextCompare) <> 0 Then
        edcTextBox.Text = slShown
    End If
'    End If
    If ilReturn = 0 Then
        If ilSelStart <= Len(edcTextBox.Text) Then
            edcTextBox.SelStart = ilSelStart
            edcTextBox.SelLength = Len(edcTextBox.Text)
        Else
            edcTextBox.SelStart = 0
            edcTextBox.SelLength = Len(edcTextBox.Text)
        End If
    End If
    gOptionalLookAhead = ilReturn
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gOptionLookAhead                *
'*                                                     *
'*             Created:5/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Validates combo box input and,  *
'*                     if invalid, retains input.      *
'*                     This is used for ComboBox .      *
'*                                                     *
'*******************************************************
Function gOptionLookAhead(cbcComboBox As ComboBox, ilBSMode As Integer, slShown As String) As Integer
'
'   ilRet = gOptionLookAhead (cbcCtrl, ilBSMode, slStr)
'   Where:
'       cbcCtrl (I)- Combo box Control containing values
'       ilBSMode (I/O)- Backspace flag (True = backspace key was pressed, False =                       '       backspace key was not pressed)
'       slStr (O)- String shown
'       ilRet (O)- 0=Selection found within list and cbcComboBox.ListIndex set to                       '               the found selection
'                  1=No match found and search string was null ("")
'                  2=No match found and search string was not null
'


    Dim ilLen As Integer    'Length of current enter text
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    Dim ilBracket As Integer
    Dim ilSearch As Integer
    Dim ilSvLastFound As Integer
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim ilReturn As Integer
    Dim ilSelStart As Integer


    ilSelStart = cbcComboBox.SelStart
    slStr = LTrim$(cbcComboBox.Text)    'Remove only leading blanks
    ilLen = Len(cbcComboBox.Text)
    ilIndex = cbcComboBox.ListIndex
    If slStr = "" Then  'If space bar selected, text will be blank- ListIndex will contain a value
'Remove test to see if behavior is OK-user removes text
'        If cbcComboBox.ListIndex >= 0 Then
'            slStr = cbcComboBox.List(cbcComboBox.ListIndex)
'            ilLen = Len(cbcComboBox.Text)
'            ilIndex = -1
'        Else
            gOptionLookAhead = 1
            Exit Function
'        End If
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
                If StrComp(cbcComboBox.List(ilLoop), "[New]", 1) = 0 Then
                    slShown = slStr
                    gOptionLookAhead = 2
                    Exit Function
                End If
                Exit For
            Else
                If Left$(cbcComboBox.List(ilLoop), 1) <> "[" Then
                    Exit For
                End If
            End If
        Next ilLoop
    Else

        'Test if matching string found in the combo box- if so display it (look ahead typing)
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
        'If item found not same as current selected- change current
        If (ilIndex <> ilSvLastFound) And (ilIndex >= 0) Then
            'Test if same name- and slStr contain whole name- if so leave index
            If (cbcComboBox.List(ilIndex) <> cbcComboBox.List(ilSvLastFound)) Or (slStr <> cbcComboBox.List(ilIndex)) Then
                cbcComboBox.ListIndex = ilSvLastFound 'This will cause a change event (reason for imChgMode)
            Else
                cbcComboBox.ListIndex = ilIndex
            End If
        ElseIf (ilIndex <> ilSvLastFound) Or ((ilIndex = ilSvLastFound) And Not ilBSMode) Then
            cbcComboBox.ListIndex = ilSvLastFound 'This will cause a change event (reason for imChgMode)
        End If
'        If (ilIndex <> ilSvLastFound) Or ((ilIndex = ilSvLastFound) And Not ilBSMode) Then 'If indices not equal- highlight look ahead text
'            cbcComboBox.SelStart = ilLen
'            cbcComboBox.SelLength = Len(cbcComboBox.Text)
'        End If
        slShown = cbcComboBox.Text
        ilReturn = 0
    Else
        slShown = slStr
        If slStr = "" Then
            ilReturn = 1
        Else
            ilReturn = 2
        End If
    End If
        cbcComboBox.Text = slShown
'    End If
    If ilReturn = 0 Then
        If ilSelStart <= Len(cbcComboBox.Text) Then
            cbcComboBox.SelStart = ilSelStart
            cbcComboBox.SelLength = Len(cbcComboBox.Text)
        Else
            cbcComboBox.SelStart = 0
            cbcComboBox.SelLength = Len(cbcComboBox.Text)
        End If
    Else    'This is not required with gOptionalLookAhead
        If cbcComboBox.SelStart = 0 Then
            cbcComboBox.SelStart = ilSelStart 'Len(cbcComboBox.Text)    'Set insert at end
        End If
    End If
    gOptionLookAhead = ilReturn
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPaintArea                      *
'*                                                     *
'*             Created:7/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear area                      *
'*                                                     *
'*******************************************************
Sub gPaintArea(pbcCtrl As Control, flCurrX As Single, flCurrY As Single, flWidth As Single, flHeight As Single, llPaintColor As Long)
'
'   gPaintArea pbcCtrl, lX, lY, lW, lH, lC
'   Where:
'       pbcCtrl (I)- Picture box control containing area to be painted
'       lX (I)- Left position of area to be painted
'       lY (I)- Top position of area to be painted
'       lH (I)- Height of area to be painted
'       lW (I)- Width of area to be painted
'       lC (I)- Color of area to be painted
'

    pbcCtrl.CurrentX = flCurrX
    pbcCtrl.CurrentY = flCurrY
    pbcCtrl.Line Step(0, 0)-Step(flWidth, flHeight), llPaintColor, BF
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gPaintCalendar                  *
'*                                                     *
'*             Created:8/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Paint standard or regular       *
'*                     calendar                        *
'*                                                     *
'*******************************************************
Sub gPaintCalendar(ilMonth As Integer, ilYear As Integer, ilType As Integer, pbcCtrl As PictureBox, tlCtrls() As FIELDAREA, llStartDate As Long, llEndDate As Long)
'
'   gPaintCalendar ilMonth, ilYear, ilType, pbcCtrl, tmCtrl, llStart, llEnd
'   Where:
'       ilMonth (I) - Calendar month to be painted (1 thru 12)
'       ilYear (I) - Calendar year to be painted (101 thru 9998) (100 & 9999 not                        '       allowed because of standard)
'       ilType (I)- 0=Paint standard month; 1= Paint regular month; 2=Julian +;
'                   3=Julian -; 4=Paint corporate on Jan-Dec; 5=Paint Corporate by Fiscal (Oct-Sept)
'       pbcCtrl (I)- Picture area control to be painted
'       tlCtrl() (I)- Array of control information about paint area
'       llStart (O)- Calendar start date
'       llEnd (O)- Calendar end date
'

    Dim llDate As Long
    Dim llStdDate As Long
    Dim slDate As String
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim slJulian As String
    Dim ilRowNo As Integer
    Dim llLastDate As Long
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim flBoxInsetX As Single
    Dim flBoxInsetY As Single
    Dim ilAdjYear As Integer
    Dim ilAdjMonth As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilFound As Integer
    Dim ilMnth As Integer

    If (ilYear < 101) Or (ilYear > 9998) Then
        pbcCtrl.Cls
        Exit Sub
    End If
    If (ilMonth < 1) Or (ilMonth > 12) Then
        pbcCtrl.Cls
        Exit Sub
    End If
    ilAdjMonth = ilMonth
    If ilType = 0 Then  'Standard month
        slDate = Trim$(str$(ilAdjMonth)) & "/1/" & Trim$(str$(ilYear))
        If ilYear < 100 Then
            slDate = gAdjYear(slDate)
            ilAdjYear = year(slDate)
            llDate = gDateValue(slDate)
            llStdDate = llDate
        Else
            ilAdjYear = year(slDate)
            llDate = DateValue(slDate)
            llStdDate = llDate
        End If
        Do While gWeekDayLong(llDate) <> 0   '0=monday
            llDate = llDate - 1
        Loop
        Do
            If gWeekDayLong(llStdDate) = 6 Then  'Save last sunday
                llLastDate = llStdDate
            End If
            llStdDate = llStdDate + 1
        Loop Until month(llStdDate) <> ilAdjMonth
    ElseIf ilType = 4 Then  'Corporate by Jan-Dec
        If (ilYear >= 0) And (ilYear <= 69) Then
            ilAdjYear = 2000 + ilYear
        ElseIf (ilYear >= 70) And (ilYear <= 99) Then
            ilAdjYear = 1900 + ilYear
        Else
            ilAdjYear = ilYear
        End If
        ilFound = False
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            If ilMonth >= tgMCof(ilLoop).iStartMnthNo Then
                If tgMCof(ilLoop).iYear = ilAdjYear + 1 Then
                    ilMnth = tgMCof(ilLoop).iStartMnthNo
                    ilIndex = 1
                    Do While ilMnth <> ilMonth
                        ilMnth = ilMnth + 1
                        If ilMnth > 12 Then
                            ilMnth = 1
                        End If
                        ilIndex = ilIndex + 1
                    Loop
                    gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex - 1), tgMCof(ilLoop).iStartDate(1, ilIndex - 1), llDate
                    gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex - 1), tgMCof(ilLoop).iEndDate(1, ilIndex - 1), llLastDate
                    ilFound = True
                    Exit For
                End If
            Else
                If tgMCof(ilLoop).iYear = ilAdjYear Then
                    ilMnth = tgMCof(ilLoop).iStartMnthNo
                    ilIndex = 1
                    Do While ilMnth <> ilMonth
                        ilMnth = ilMnth + 1
                        If ilMnth > 12 Then
                            ilMnth = 1
                        End If
                        ilIndex = ilIndex + 1
                    Loop
                    gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex - 1), tgMCof(ilLoop).iStartDate(1, ilIndex - 1), llDate
                    gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex - 1), tgMCof(ilLoop).iEndDate(1, ilIndex - 1), llLastDate
                    ilFound = True
                    Exit For
                End If
            End If
        Next ilLoop
        If Not ilFound Then
            pbcCtrl.Cls
            Exit Sub
        End If
        'slDate = "1/1/" & Trim$(Str$(ilYear))
        'If ilYear < 100 Then
        '    slDate = gAdjYear(slDate)
        '    ilAdjYear = Year(slDate)
        '    llDate = gDateValue(slDate)
        'Else
        '    ilAdjYear = Year(slDate)
        '    llDate = DateValue(slDate)
        'End If
        'If tgSpf.sRYEnd = "L" Then
        '    Do While gWeekDayLong(llDate) <> 0   '0=monday
        '        llDate = llDate - 1
        '    Loop
        'Else
        '    'Count the number of weeks from the first monday
        '    llTempDate = llDate
        '    Do While gWeekDayLong(llTempDate) <> 0   '0=monday
        '        llTempDate = llTempDate + 1
        '    Loop
        'End If
        'ilIndex = 0
        'For ilLoop = 2 To ilMonth Step 1
        '    If (ilLoop = 2) And (tgSpf.sRYEnd <> "L") Then
        '        llDate = llTempDate + 7 * tgSpf.iRCorp(ilIndex)
        '    Else
        '        llDate = llDate + 7 * tgSpf.iRCorp(ilIndex)
        '    End If
        '    ilIndex = ilIndex + 1
        '    If ilIndex > 2 Then
        '        ilIndex = 0
        '    End If
        'Next ilLoop
        'If ilMonth <> 12 Then
        '    If (ilMonth = 1) And (tgSpf.sRYEnd <> "L") Then
        '        llLastDate = llTempDate + 7 * tgSpf.iRCorp(ilIndex) - 1
        '    Else
        '        llLastDate = llDate + 7 * tgSpf.iRCorp(ilIndex) - 1
        '    End If
        'Else
        '    slLastDate = "1/1/" & Trim$(Str$(ilYear + 1))
        '    If ilYear < 100 Then
        '        slLastDate = gAdjYear(slLastDate)
        '        llLastDate = gDateValue(slLastDate)
        '    Else
        '        llLastDate = DateValue(slLastDate)
        '    End If
        '    If tgSpf.sRYEnd = "L" Then
        '        Do While gWeekDayLong(llLastDate) <> 0   '0=monday
        '            llLastDate = llLastDate - 1
        '        Loop
        '    End If
        '    llLastDate = llLastDate - 1
        'End If
    ElseIf ilType = 5 Then  'Corporate by Fiscal (Oct-Sept)
        If (ilYear >= 0) And (ilYear <= 69) Then
            ilAdjYear = 2000 + ilYear
        ElseIf (ilYear >= 70) And (ilYear <= 99) Then
            ilAdjYear = 1900 + ilYear
        Else
            ilAdjYear = ilYear
        End If
        ilFound = False
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            If tgMCof(ilLoop).iYear = ilAdjYear Then
                For ilIndex = 1 To 12 Step 1
                    ilAdjMonth = tgMCof(ilLoop).iStartMnthNo + ilIndex - 1
                    If ilAdjMonth > 12 Then
                        ilAdjMonth = ilAdjMonth - 12
                    End If
                    If ilAdjMonth = ilMonth Then
                        gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex - 1), tgMCof(ilLoop).iStartDate(1, ilIndex - 1), llDate
                        gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex - 1), tgMCof(ilLoop).iEndDate(1, ilIndex - 1), llLastDate
                        ilFound = True
                        Exit For
                    End If
                Next ilIndex
                If ilFound Then
                    Exit For
                End If
            End If
        Next ilLoop
        If Not ilFound Then
            pbcCtrl.Cls
            Exit Sub
        End If
    Else
        slDate = Trim$(str$(ilAdjMonth)) & "/1/" & Trim$(str$(ilYear))
        If ilYear < 100 Then
            slDate = gAdjYear(slDate)
            ilAdjYear = year(slDate)
            llDate = gDateValue(slDate)
        Else
            ilAdjYear = year(slDate)
            llDate = DateValue(slDate)
        End If
    End If
'    If gValidDate(slDate) = No Then
'        pbcCtrl.Cls
'        Exit Sub
'    End If
    If (ilAdjMonth <> imCalMonth) Or (ilAdjYear <> imCalYear) Then
        pbcCtrl.Cls
        imCalMonth = ilAdjMonth
        imCalYear = ilAdjYear
    End If
    ilRowNo = 0
    llStartDate = 0
    llEndDate = 0
    slFontName = pbcCtrl.FontName
    flFontSize = pbcCtrl.FontSize
    If (ilType <= 1) Or (ilType = 4) Or (ilType = 5) Then
        pbcCtrl.FontBold = True
        flBoxInsetX = fgBoxInsetX
        flBoxInsetY = -15
    Else
        pbcCtrl.FontBold = False
        pbcCtrl.FontSize = 7
        pbcCtrl.FontName = "Arial"
        pbcCtrl.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        flBoxInsetX = 5
        flBoxInsetY = 5
    End If
    llColor = pbcCtrl.ForeColor
    Do
        ilWkDay = gWeekDayLong(llDate)
        If llStartDate = 0 Then
            llStartDate = llDate
        End If
        If (ilType <= 1) Or (ilType = 4) Or (ilType = 5) Then
            slDay = Trim$(str$(Day(llDate)))
            If Len(slDay) <= 1 Then
                slDay = " " & slDay
            End If
            If ilAdjMonth <> month(llDate) Then
                pbcCtrl.ForeColor = MAGENTA 'GREEN
            Else
                pbcCtrl.ForeColor = llColor
            End If
        ElseIf ilType = 2 Then  'Julian +
            slDay = Format$(llDate, "y")
            If Len(slDay) <= 1 Then
                slDay = "   " & slDay
            ElseIf Len(slDay) = 2 Then
                If Val(slDay) < 20 Then
                    slDay = "  " & slDay
                Else
                    slDay = " " & slDay
                End If
            End If
        ElseIf ilType = 3 Then  'Julian -
            slJulian = "12/31/" & Format$(llDate, "yyyy")
            slDay = Trim$(str$(Val(Format$(slJulian, "y")) - Val(Format$(llDate, "y"))))
            If Len(slDay) <= 1 Then
                slDay = "   " & slDay
            ElseIf Len(slDay) = 2 Then
                If Val(slDay) < 20 Then
                    slDay = "  " & slDay
                Else
                    slDay = " " & slDay
                End If
            End If
        End If
'        gSetShow pbcCtrl, slDay, tlCtrls(ilWkDay + 1)
        pbcCtrl.CurrentX = tlCtrls(ilWkDay + 1).fBoxX + flBoxInsetX
        pbcCtrl.CurrentY = tlCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tlCtrls(ilWkDay + 1).fBoxH + 15) + flBoxInsetY '(fgBoxGridH + 15) -  30'+ fgBoxInsetY
        pbcCtrl.Print slDay 'tlCtrls(ilWkDay + 1).sShow
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
'    Loop Until (Month(llDate) <> ilMonth) Or ((ilType <> 1) And (llDate > llLastDate)) 'julian by std
    Loop Until (((ilType <> 0) And (ilType <> 4) And (ilType <> 5)) And (month(llDate) <> ilAdjMonth)) Or (((ilType = 0) Or (ilType = 4) Or (ilType = 5)) And (llDate > llLastDate))
    pbcCtrl.ForeColor = llColor
    llEndDate = llDate - 1
    pbcCtrl.FontSize = flFontSize
    pbcCtrl.FontName = slFontName
    pbcCtrl.FontSize = flFontSize
    pbcCtrl.FontBold = True
End Sub

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
Sub gProcessArrowKey(ilShift As Integer, ilKeyCode As Integer, lbcCtrl As Control, ilLbcArrowSetting As Integer)
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

    Dim ilLoop As Integer

    If (ilShift And ALTMASK) > 0 Then
        lbcCtrl.Visible = Not lbcCtrl.Visible
    ElseIf (ilShift And SHIFTMASK) > 0 Then
    Else
        ilLbcArrowSetting = True
        If ilKeyCode = KEYUP Then    'Up arrow
            If lbcCtrl.ListIndex > 0 Then
                lbcCtrl.ListIndex = lbcCtrl.ListIndex - 1
                If lbcCtrl.MultiSelect Then
                    For ilLoop = 0 To lbcCtrl.ListCount - 1 Step 1
                        If ilLoop = lbcCtrl.ListIndex Then
                            lbcCtrl.Selected(lbcCtrl.ListIndex) = True
                        Else
                            lbcCtrl.Selected(ilLoop) = False
                        End If
                    Next ilLoop
                End If
            End If
        Else
            If lbcCtrl.ListIndex < lbcCtrl.ListCount - 1 Then
                lbcCtrl.ListIndex = lbcCtrl.ListIndex + 1
                If lbcCtrl.MultiSelect Then
                    For ilLoop = 0 To lbcCtrl.ListCount - 1 Step 1
                        If ilLoop = lbcCtrl.ListIndex Then
                            lbcCtrl.Selected(lbcCtrl.ListIndex) = True
                        Else
                            lbcCtrl.Selected(ilLoop) = False
                        End If
                    Next ilLoop
                End If
            End If
        End If
        ilLbcArrowSetting = False
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gProcessLbcClick                *
'*                                                     *
'*             Created:9/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Process list click.  Used with *
'*                      combo boxes made up from text  *
'*                      control/button control/list    *
'*                      box control.                   *
'*                                                     *
'*******************************************************
Sub gProcessLbcClick(lbcCtrl As Control, edcCtrl As Control, ilChgMode As Integer, ilLbcArrowSetting As Integer)
'
'   gProcessLbcClick lbcSPerson, edcDropDown, imChgMode,                                                imLbcArrowSetting
'   Where:
'       lbcSPerson (I)- list box control
'       edcDropDown (I)- User edit control
'       imChgMode (I/O)- Change mode flag
'       imLbcArrowSetting (I/O) - list box arrow setting flag
'                               True= make list box invisible (user click on item)
'                               False= retain list box visible state
'

    Dim ilLbcArrowSet As Integer

    If ilChgMode = False Then
        ilChgMode = True
        ilLbcArrowSet = ilLbcArrowSetting
        edcCtrl.Text = lbcCtrl.List(lbcCtrl.ListIndex)
        If (edcCtrl.Visible) And (edcCtrl.Enabled) Then
            edcCtrl.SetFocus
        End If
        If Not ilLbcArrowSet Then
            lbcCtrl.Visible = False
        End If
        ilChgMode = False
    End If
    ilLbcArrowSetting = False
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:gRightJustifyShowStr            *
'*                                                     *
'*             Created:4/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine the CurrentX position *
'*                     so value is right justified     *
'*                                                     *
'*******************************************************
Function gRightJustifyShowStr(pbcCtrl As Control, slShowStr As String, tlCtrlArray As FIELDAREA) As Single
'
'   flX = gRightJustifyShowStr(pbcCtrl, slShow, CtrlArray)
'   Where:
'       pbcCtrl (I) Control that will contain the picture
'       slShow (I)- String to be shown
'       CtrlArray (O)- field control array
'       flX (O)- CurrentX value determine to right justify value
'
    Dim flPicWidth As Single
    Dim flStrWidth As Single
    If Len(slShowStr) = 0 Then
        Exit Function
    End If
    flPicWidth = tlCtrlArray.fBoxW - fgBoxInsetX
    flStrWidth = pbcCtrl.TextWidth(slShowStr)
    If flStrWidth > flPicWidth Then
        gRightJustifyShowStr = tlCtrlArray.fBoxX + fgBoxInsetX
    Else
        gRightJustifyShowStr = tlCtrlArray.fBoxX + fgBoxInsetX + (flPicWidth - flStrWidth)
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gSetChgFlag                     *
'*                                                     *
'*             Created:4/21/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set the change flag within the  *
'*                     control array                   *
'*                                                     *
'*******************************************************
Sub gSetChgFlag(slField As String, Ctrl As Control, tlCtrlArray As FIELDAREA)
'
'   gSetChgFlag slRecField, edcCtrl, CtrlArray
'   Where:
'       slRecField (I)- original record field
'       edcCtrl (I)- User control
'       CtrlArray (O)- field control array
'

    Dim slStr As String
    Dim slSpecStr As String

    If TypeOf Ctrl Is MaskEdBox Then    'If mask assume phone number
        slStr = Ctrl.Text
        If Left$(slStr, 1) = "(" Then
            gGetPhoneNo Ctrl, slStr
        Else
            gGetSocialSecurityNo Ctrl, slStr
        End If
        If StrComp(Trim$(slField), Trim$(slStr), vbTextCompare) <> 0 Then
            tlCtrlArray.iChg = True
        Else
            tlCtrlArray.iChg = False
        End If
        Exit Sub
    End If
    If TypeOf Ctrl Is ComboBox Then    'If Combo- test for blank and [none]
        slSpecStr = Trim$(Ctrl.List(0))
        If StrComp(slSpecStr, "[New]", vbTextCompare) = 0 Then
            If Ctrl.ListCount >= 2 Then
                slSpecStr = Trim$(Ctrl.List(1))
                If Left$(slStr, 1) <> "[" Then
                    slSpecStr = "[None]"
                End If
            Else
                slSpecStr = "[None]"
            End If
        Else
            If Left$(slStr, 1) <> "[" Then
                slSpecStr = "[None]"
            End If
        End If
        If StrComp(Trim$(Ctrl.Text), slSpecStr, vbTextCompare) = 0 Then
            If (slField = "") Or (StrComp(slField, slSpecStr, vbTextCompare) = 0) Then
                tlCtrlArray.iChg = False
            Else
                tlCtrlArray.iChg = True
            End If
            Exit Sub
        End If
        If StrComp(slField, slSpecStr, vbTextCompare) = 0 Then
            If (Trim$(Ctrl.Text) = "") Or (StrComp(Trim$(Ctrl.Text), slSpecStr, vbTextCompare) = 0) Then
                tlCtrlArray.iChg = False
            Else
                tlCtrlArray.iChg = True
            End If
            Exit Sub
        End If
        If StrComp(slField, Trim$(Ctrl.Text), vbTextCompare) = 0 Then
            tlCtrlArray.iChg = False
        Else
            tlCtrlArray.iChg = True
        End If
        Exit Sub
    End If
    If TypeOf Ctrl Is CheckBox Then    'If check box- test for on/off
        If slField = "Y" Then
            If Ctrl.Value Then
                tlCtrlArray.iChg = False
            Else
                tlCtrlArray.iChg = True
            End If
        Else
            If Ctrl.Value Then
                tlCtrlArray.iChg = True
            Else
                tlCtrlArray.iChg = False
            End If
        End If
        Exit Sub
    End If
    slStr = Ctrl.Text
    gSetChgFlagStr slField, slStr, tlCtrlArray
'    If StrComp(Trim$(slField), Trim$(Ctrl.Text), 0) <> 0 Then
'        If (Trim$(slField) = "") And (Ctrl.Text = "[None]") Then
'            tlCtrlArray.iChg = False
'        ElseIf (Trim$(slField) = "[None]") And (Ctrl.Text = "") Then
'            tlCtrlArray.iChg = False
'        Else
'            tlCtrlArray.iChg = True
'        End If
'    Else
'        tlCtrlArray.iChg = False
'    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gSetChgFlagStr                  *
'*                                                     *
'*             Created:4/21/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set the change flag within the  *
'*                     control array                   *
'*                                                     *
'*******************************************************
Sub gSetChgFlagStr(slField As String, slCtrl As String, tlCtrlArray As FIELDAREA)
'
'   gSetChgFlagStr slRecField, slCtrl, CtrlArray
'   Where:
'       slRecField (I)- original record field
'       slCtrl (I)- User input string
'       CtrlArray (O)- field control array
'

    Dim ilStr1Num As Integer    'True=Number
    Dim ilStr2Num As Integer    'True=Number
    Dim dlNum1 As Double        'Value os string 1
    Dim dlNum2 As Double        'Value of string 2

    If StrComp(Trim$(slField), Trim$(slCtrl), vbBinaryCompare) <> 0 Then
        If (Trim$(slField) = "") And (slCtrl = "[None]") Then
            tlCtrlArray.iChg = False
        ElseIf (Trim$(slField) = "[None]") And (slCtrl = "") Then
            tlCtrlArray.iChg = False
        ElseIf (Trim$(slField) = "") And (slCtrl = "[N/A]") Then
            tlCtrlArray.iChg = False
        ElseIf (Trim$(slField) = "[N/A]") And (slCtrl = "") Then
            tlCtrlArray.iChg = False
        ElseIf (Trim$(slField) = "") And (slCtrl = "N/A") Then
            tlCtrlArray.iChg = False
        ElseIf (Trim$(slField) = "N/A") And (slCtrl = "") Then
            tlCtrlArray.iChg = False
        ElseIf (slCtrl = "[New]") Then
            tlCtrlArray.iChg = False
        Else
            ilStr1Num = True
            ilStr2Num = True
            On Error GoTo gSetChgFlagStr1Err
            dlNum1 = CDbl(slField)
            On Error GoTo gSetChgFlagStr2Err
            dlNum2 = CDbl(slCtrl)
            On Error GoTo 0
            If ilStr1Num And ilStr2Num Then
                If (dlNum1 = dlNum2) And (Len(Trim$(slField)) = Len(Trim$(slCtrl))) Then
                    tlCtrlArray.iChg = False
                Else
                    tlCtrlArray.iChg = True
                End If
            Else
                If gValidDate(slField) And gValidDate(slCtrl) Then
                    If gDateValue(slField) = gDateValue(slCtrl) Then
                        tlCtrlArray.iChg = False
                    Else
                        tlCtrlArray.iChg = True
                    End If
                Else
                    If gValidTime(slField) And gValidTime(slCtrl) Then
                        If (Len(Trim$(slField)) = 0) And (Len(Trim$(slCtrl)) <> 0) Then
                            tlCtrlArray.iChg = True
                        ElseIf (Len(Trim$(slField)) <> 0) And (Len(Trim$(slCtrl)) = 0) Then
                            tlCtrlArray.iChg = True
                        Else
                            If gTimeToCurrency(slField, False) = gTimeToCurrency(slCtrl, False) Then
                                tlCtrlArray.iChg = False
                            Else
                                tlCtrlArray.iChg = True
                            End If
                        End If
                    Else
                        If gValidLength(slField) And gValidLength(slCtrl) Then
                            If (Len(Trim$(slField)) = 0) And (Len(Trim$(slCtrl)) <> 0) Then
                                tlCtrlArray.iChg = True
                            ElseIf (Len(Trim$(slField)) <> 0) And (Len(Trim$(slCtrl)) = 0) Then
                                tlCtrlArray.iChg = True
                            Else
                                If gLengthToCurrency(slField) = gLengthToCurrency(slCtrl) Then
                                    tlCtrlArray.iChg = False
                                Else
                                    tlCtrlArray.iChg = True
                                End If
                            End If
                        Else
                            tlCtrlArray.iChg = True
                        End If
                    End If
                End If
            End If
        End If
    Else
        tlCtrlArray.iChg = False
    End If
    Exit Sub
gSetChgFlagStr1Err:
    ilStr1Num = False
    Resume Next
gSetChgFlagStr2Err:
    ilStr2Num = False
    Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gSetCtrl                        *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set mouse area within the       *
'*                     control area.                   *
'*                                                     *
'*******************************************************
Sub gSetCtrl(tlCtrlArray As FIELDAREA, flBoxX As Single, flBoxY As Single, flBoxW As Single, flBoxH As Single)
'
'   gSetCtrl CtrlArray(1), fBoxX, fBoxY, fBoxW, fBoxH
'   Where
'       CtrlArray (I/O)- field control array
'       fBoxX (I)- x offset of mouse area within picture control
'       fBoxY (I)- y offset of mouse area within picture control
'       fBoxW (I)- Width of mouse area within picture
'       fBoxH (I)- Height of mouse area within picture
'

    tlCtrlArray.fBoxX = flBoxX
    tlCtrlArray.fBoxY = flBoxY
    tlCtrlArray.fBoxW = flBoxW
    tlCtrlArray.fBoxH = flBoxH
    tlCtrlArray.iReq = True
    tlCtrlArray.iChg = False
    tlCtrlArray.iAlign = 0
    tlCtrlArray.sShow = ""
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gSetCtrlReadOnly                *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set control as read only        *
'*                                                     *
'*******************************************************
Sub gSetCtrlReadOnly(edcCtrl As Control)
    Dim llRet As Long

    llRet = SendMessageByNum(edcCtrl.HWnd, EM_SETREADONLY, 1, 0)

End Sub


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
        Traffic!mkcPhone.mask = "(AAA) AAA-AAAA"
        Traffic!mkcPhone.MaxLength = 14
    Else
        Traffic!mkcPhone.mask = "(AAA) AAA-AAAA Ext(AAAA)"
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
'*      Procedure Name:gSetShow                        *
'*                                                     *
'*             Created:4/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set the Show field from Text    *
'*                     property of control.  This sub  *
'*                     adjusts the string length so    *
'*                     the string fits within the field*
'*                     control array area (mouse area).*
'*******************************************************
Sub gSetShow(pbcCtrl As Control, slShowStr As String, tlCtrlArray As FIELDAREA)
'
'   gSetShow pbcCtrl, slShow, CtrlArray
'   Where:
'       pbcCtrl (I) Control that will contain the input control
'       slShow (I)- String to be shown
'       CtrlArray (O)- field control array
'
    

    Dim flPicWidth As Single
    Dim flSingleChar As Single
    Dim flDeltaWidth As Single

    tlCtrlArray.sShow = Trim$(slShowStr)
    If Len(tlCtrlArray.sShow) = 0 Then
        Exit Sub
    End If
    
    On Error GoTo 0
    On Error Resume Next 'prevent error 16
    
'    flPicWidth = tlCtrlArray.fBoxW - 2 * fgBoxInsetX
    flPicWidth = tlCtrlArray.fBoxW - fgBoxInsetX
    If flPicWidth < 0 Then
        tlCtrlArray.sShow = ""
        Exit Sub
    End If
    flDeltaWidth = pbcCtrl.TextWidth(tlCtrlArray.sShow) - flPicWidth
    If flDeltaWidth < 0 Then
        Exit Sub
    End If
    flSingleChar = pbcCtrl.TextWidth("M")
    If Len(tlCtrlArray.sShow) - flDeltaWidth / flSingleChar < 0 Then
        Exit Sub
    End If
    tlCtrlArray.sShow = Left$(tlCtrlArray.sShow, Len(tlCtrlArray.sShow) - flDeltaWidth / flSingleChar)
    Do While (pbcCtrl.TextWidth(tlCtrlArray.sShow) > flPicWidth)
        tlCtrlArray.sShow = Left$(tlCtrlArray.sShow, Len(tlCtrlArray.sShow) - 1)
    Loop
    On Error GoTo 0
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gTimeOutLine                    *
'*                                                     *
'*             Created:10/20/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Position outline box around    *
'*                      key pressed by the user in the *
'*                      time input pad                 *
'*                                                     *
'*******************************************************
Sub gTimeOutLine(ilKeyAscii As Integer, imcCtrl As Control)
'
'  gTimeOutLine ilKeyAscii, imcCtrl
'   Where:
'       ilKeyAscii (I) - standard numericANSI keycode obtained from the                                 '               KeyPress event
'       imcCtrl (I) - name of the outline control
'
   Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    If ilKeyAscii <> KEYBACKSPACE Then
        Select Case UCase(Chr(ilKeyAscii))
            Case "7"
                ilRowNo = 1
                ilColNo = 1
            Case "8"
                ilRowNo = 1
                ilColNo = 2
            Case "9"
                ilRowNo = 1
                ilColNo = 3
            Case "4"
                ilRowNo = 2
                ilColNo = 1
            Case "5"
                ilRowNo = 2
                ilColNo = 2
            Case "6"
                ilRowNo = 2
                ilColNo = 3
            Case "1"
                ilRowNo = 3
                ilColNo = 1
            Case "2"
                ilRowNo = 3
                ilColNo = 2
            Case "3"
                ilRowNo = 3
                ilColNo = 3
            Case "0"
                ilRowNo = 4
                ilColNo = 1
            Case "00"   'Not possible
                ilRowNo = 4
                ilColNo = 2
            Case ":"
                ilRowNo = 5
                ilColNo = 1
            Case "A"   'Not possible
                ilRowNo = 5
                ilColNo = 2
            Case "P"
                ilRowNo = 5
                ilColNo = 3
        End Select
        flX = fgPadMinX + (ilColNo - 1) * fgPadDeltaX
        flY = fgPadMinY + (ilRowNo - 1) * fgPadDeltaY
        imcCtrl.Move flX - 15, flY - 15
        imcCtrl.Visible = True
    Else
        imcCtrl.Visible = False
    End If
End Sub


Public Sub gFndFirst(lbcList As Control, slInMatch As String)
    cgList = lbcList
    If TypeOf lbcList Is ComboBox Then
        igFoundRow = SendMessageByString(lbcList.HWnd, CB_FINDSTRING, -1, slInMatch)
    Else
        igFoundRow = SendMessageByString(lbcList.HWnd, LB_FINDSTRING, -1, slInMatch)
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

Public Function gSetCheck(lCntrl As Integer) As Integer
    'Returns False if the control equals vbUnChecked.
    'Returns True if the control equals vbChecked or vbGray.

    If lCntrl = vbUnchecked Then
        gSetCheck = False
    Else
        gSetCheck = True
    End If

End Function

Public Sub gFunctionKeyBranch(ilKeyCode As Integer)
    Dim slStr As String
    Dim ilParse As Integer

    If igBkgdProg <> 0 Then
        Exit Sub
    End If

    Select Case ilKeyCode
        Case KEYF1
            'Control is returned without waiting which causes prolem
            'in the calling application focus point (focus set to previous field)
            Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
            Traffic!cdcSetup.HelpCommand = cdlHelpIndex
            Traffic!cdcSetup.Action = 6

        Case KEYF4
            Calendar.Show vbModal
        Case KEYF5
            If igTestSystem Then
                slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CInt(fgCalcLeft))) & "\" & Trim$(str$(CInt(fgCalcTop)))
            Else
                slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CInt(fgCalcLeft))) & "\" & Trim$(str$(CInt(fgCalcTop)))
            End If
            sgCommandStr = slStr
            MathCalc.Show vbModal
            ilParse = gParseItem(sgDoneMsg, 2, "|", slStr)
            fgCalcLeft = Val(slStr)
            ilParse = gParseItem(sgDoneMsg, 3, "|", slStr)
            fgCalcTop = Val(slStr)
    End Select
End Sub


Public Sub gShellToBugdet(frm As Form)
    Dim slCommandStr
    Dim ilPos As Integer
    Dim slDate As String
    Dim ilRet As Integer
    Dim dlShellRet As Double
    
    If igTestSystem Then
        slCommandStr = "Traffic^Test\" & sgUserName & "\" & "-1"    'Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    Else
        slCommandStr = "Traffic^Prod\" & sgUserName & "\" & "-1"    'Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    End If
    If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
        ilPos = InStr(1, slCommandStr, "Guide", vbTextCompare)
        If ilPos > 0 Then
            slCommandStr = Left(slCommandStr, ilPos - 1) & "CSI" & Mid(slCommandStr, ilPos + 5)
        End If
    End If
    'Dan M 9/20/10 problems in v57 reports.exe running GetCsiName
    'slDate = Trim$(gGetCSIName("SYSDate"))
    slDate = gCSIGetName()
    If slDate <> "" Then
        'use slDate when writing to file later
        slDate = " /D:" & slDate
        'slCommandStr = slCommandStr & " /D:" & slDate
        slCommandStr = slCommandStr & slDate
    End If
    slCommandStr = slCommandStr & " /ULF:" & lgUlfCode
    On Error GoTo LoadErr
    ilRet = 0
    'AppActivate "Budget"
    'If ilRet = 1 Then
    '    dlShellRet = Shell(sgExePath & "Budget.Exe " & slCommandStr, vbNormalFocus)
    'End If
    Screen.MousePointer = vbDefault
    Traffic.WindowState = vbMinimized
    gShellAndWait frm, sgExePath & "Budget.exe " & slCommandStr, vbNormalFocus, True    'vbTrue  'vbFalse
    Traffic.WindowState = vbMaximized
    Exit Sub
LoadErr:
    ilRet = 1
    Resume Next
    Exit Sub

End Sub
'Public Sub gShellToInvoice(frm As Form)
Public Sub gShellToInvoice()
    Dim slCommandStr
    Dim ilPos As Integer
    Dim slDate As String
    Dim ilRet As Integer
    Dim dlShellRet As Double
    
    If igTestSystem Then
        slCommandStr = "Traffic^Test\" & sgUserName & "\" & "-1"    'Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    Else
        slCommandStr = "Traffic^Prod\" & sgUserName & "\" & "-1"    'Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    End If
    If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
        ilPos = InStr(1, slCommandStr, "Guide", vbTextCompare)
        If ilPos > 0 Then
            slCommandStr = Left(slCommandStr, ilPos - 1) & "CSI" & Mid(slCommandStr, ilPos + 5)
        End If
    End If
    'Dan M 9/20/10 problems in v57 reports.exe running GetCsiName
    'slDate = Trim$(gGetCSIName("SYSDate"))
    slDate = gCSIGetName()
    If slDate <> "" Then
        'use slDate when writing to file later
        slDate = " /D:" & slDate
        'slCommandStr = slCommandStr & " /D:" & slDate
        slCommandStr = slCommandStr & slDate
    End If
    slCommandStr = slCommandStr & " /ULF:" & lgUlfCode
    On Error GoTo LoadErr
    ilRet = 0
    'Screen.MousePointer = vbDefault
    'Traffic.WindowState = vbMinimized
    'gShellAndWait frm, sgExePath & "Invoice.exe " & slCommandStr, vbNormalFocus, True    'vbTrue  'vbFalse
    'Traffic.WindowState = vbMaximized
    AppActivate "CSI Invoices"
    If ilRet = 1 Then
        dlShellRet = Shell(sgExePath & "Invoice.Exe " & slCommandStr, vbNormalFocus)
    End If
    Exit Sub
LoadErr:
    ilRet = 1
    Resume Next
    Exit Sub
End Sub

Public Sub gShellToRateCard()
    Dim slCommandStr
    Dim ilPos As Integer
    Dim slDate As String
    Dim ilRet As Integer
    Dim dlShellRet As Double
    
    If igTestSystem Then
        slCommandStr = "Traffic^Test\" & sgUserName & "\" & "-1"    'Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    Else
        slCommandStr = "Traffic^Prod\" & sgUserName & "\" & "-1"    'Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    End If
    If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
        ilPos = InStr(1, slCommandStr, "Guide", vbTextCompare)
        If ilPos > 0 Then
            slCommandStr = Left(slCommandStr, ilPos - 1) & "CSI" & Mid(slCommandStr, ilPos + 5)
        End If
    End If
    slDate = gCSIGetName()
    If slDate <> "" Then
        'use slDate when writing to file later
        slDate = " /D:" & slDate
        slCommandStr = slCommandStr & slDate
    End If
    slCommandStr = slCommandStr & " /ULF:" & lgUlfCode
    err = 0
    On Error GoTo LoadErr
    ilRet = 0
    AppActivate ("Rate Card")
    If ilRet = 1 Then
        dlShellRet = Shell(sgExePath & "RateCard.Exe " & slCommandStr, vbNormalFocus)
    End If
    Exit Sub
LoadErr:
    ilRet = 1
    Resume Next
    Exit Sub
End Sub

Public Sub gSetComboboxDropdownHeight(Form1 As Form, Combo1 As ComboBox, ilNumberOfRows As Integer)
  Dim pt As POINTAPI
   Dim rc As RECT
   Dim cWidth As Long
   Dim newHeight As Long
   Dim oldScaleMode As Long
   Dim numItemsToDisplay As Long
   Dim itemHeight As Long
   
 'Save the current form scalemode, then
  'switch to pixels
   oldScaleMode = Form1.ScaleMode
   Form1.ScaleMode = vbPixels
   
  'the width of the combo, used below
   cWidth = Combo1.Width
  
  'get the system height of a single
  'combo box list item
   itemHeight = SendMessage(Combo1.HWnd, CB_GETITEMHEIGHT, 0, ByVal 0)
   
  'Calculate the new height of the combo box. This
  'is the number of items times the item height
  'plus two. The 'plus two' is required to allow
  'the calculations to take into account the size
  'of the edit portion of the combo as it relates
  'to item height. In other words, even if the
  'combo is only 21 px high (315 twips), if the
  'item height is 13 px per item (as it is with
  'small fonts), we need to use two items to
  'achieve this height.
   newHeight = itemHeight * (ilNumberOfRows + 2)
   
  'get the co-ordinates of the combo box
  'relative to the screen
   Call GetWindowRect(Combo1.HWnd, rc)
   pt.X = rc.Left
   pt.Y = rc.Top

  'then translate into co-ordinates
  'relative to the form.
   Call ScreenToClient(Form1.HWnd, pt)

  'using the values returned and set above,
  'call MoveWindow to reposition the combo box
   Call MoveWindow(Combo1.HWnd, pt.X, pt.Y, Combo1.Width, newHeight, True)
   
  ''it's done, so show the new combo height
   'Call SendMessage(Combo1.hwnd, CB_SHOWDROPDOWN, True, ByVal 0)
   
  'restore the original form scalemode
  'before leaving
   Form1.ScaleMode = oldScaleMode
End Sub


Public Sub gShellToPostManual(frm As Form, Optional slExtraCommand = "")
    Dim ilShell As Integer
    Dim dlShellRet As Double
    Dim slCommandStr As String
    Dim ilPos As Integer
    Dim slDate As String
    

    'dlShellRet = Shell(sgExePath & "ExportProj.Exe /IniLoc:" & CurDir$ & " /UserInput", vbMinimizedNoFocus)
    If igTestSystem Then
        slCommandStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    Else
        slCommandStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    End If
    If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
        ilPos = InStr(1, slCommandStr, "Guide", vbTextCompare)
        If ilPos > 0 Then
            slCommandStr = Left(slCommandStr, ilPos - 1) & "CSI" & Mid(slCommandStr, ilPos + 5)
        End If
    End If
    'Dan M 9/20/10 problems in v57 reports.exe running GetCsiName
    'slDate = Trim$(gGetCSIName("SYSDate"))
    slDate = gCSIGetName()
    If slDate <> "" Then
        'use slDate when writing to file later
        slDate = " /D:" & slDate
        'slCommandStr = slCommandStr & " /D:" & slDate
        slCommandStr = slCommandStr & slDate
    End If
    slCommandStr = slCommandStr & " /ULF:" & lgUlfCode
    slCommandStr = slCommandStr & " " & slExtraCommand
    Screen.MousePointer = vbDefault
    Traffic.WindowState = vbMinimized
    gShellAndWait frm, sgExePath & "PostManual.exe " & slCommandStr, vbNormalFocus, False    'vbFalse
    
    Traffic.WindowState = vbMaximized
End Sub

'TTP 10271: Insertion Orders report: add option to create a separate PDF file for each vehicle
'*******************************************************
'*                                                     *
'*      Procedure Name:gGetListSelectedCount           *
'*                                                     *
'*             Created:09/20/21      By:J.White        *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get a count of how many items  *
'*                      are selected in the provided   *
'*                      multiselect list box           *
'*                                                     *
'*******************************************************
Function gGetListSelectedCount(olList As ListBox) As Integer
    Dim ilLoop As Integer
    Dim ilCount As Integer
    For ilLoop = 0 To olList.ListCount - 1
        If olList.Selected(ilLoop) = True Then ilCount = ilCount + 1
    Next ilLoop
    gGetListSelectedCount = ilCount
End Function
