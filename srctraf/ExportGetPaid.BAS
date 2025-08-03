Attribute VB_Name = "EXPORTGETPAIDSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of ExportGetPaid.BAS on Wed 6/17/09 @ 12
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  gCreateFixedRec               gUnpackDateEightChar          gUnpackDateByFormat       *
'*  IsReallyNumeric               NoComma                       GetState                  *
'*                                                                                        *
'******************************************************************************************

Option Explicit
Option Compare Text
Public Enum enumJustification
    Justify_Left = 0
    Justify_Right = 1
End Enum
Public Enum enumDataTypes
    StrVal = 0
    DateVal = 1
    TimeVal = 2
    IntVal = 3
    LongVal = 4
    BoolVal = 5
    DecVal = 6
End Enum
Type FixedRec
    sFieldValue As String
    iStartPos As Integer
    iLength As Integer
    bLeadingZeroes As Boolean
    sJustification As enumJustification
    sDataType As enumDataTypes
End Type
Type ComDelRec
    sFieldValue As String
    sFieldNum As String
End Type
Public Function gCreateFixedRec(ilRecLen As Integer, tlFixedRec() As FixedRec) As String 'VBC NR

       Dim ilLoop As Integer 'VBC NR
       Dim slRecord As String 'VBC NR
       Dim slStr As String 'VBC NR
       slRecord = String$(ilRecLen, " ") 'VBC NR
       For ilLoop = LBound(tlFixedRec) To UBound(tlFixedRec) Step 1 'VBC NR
           slStr = tlFixedRec(ilLoop).sFieldValue 'VBC NR
               If tlFixedRec(ilLoop).sJustification Then 'VBC NR
                   slStr = right$(String$(tlFixedRec(ilLoop).iLength, " ") + slStr, tlFixedRec(ilLoop).iLength) 'VBC NR
               End If 'VBC NR
               If tlFixedRec(ilLoop).bLeadingZeroes Then 'VBC NR
                   slStr = right$(String$(tlFixedRec(ilLoop).iLength, "0") + slStr, tlFixedRec(ilLoop).iLength) 'VBC NR
               End If 'VBC NR
       If Len(slStr) > tlFixedRec(ilLoop).iLength Then 'VBC NR
           slStr = Left$(slStr, tlFixedRec(ilLoop).iLength) 'VBC NR
       End If 'VBC NR
       If tlFixedRec(ilLoop).iStartPos > 0 And tlFixedRec(ilLoop).iLength > 0 Then          '7-18-13 avoid error when a start position or length doesnt exist
            Mid$(slRecord, tlFixedRec(ilLoop).iStartPos, tlFixedRec(ilLoop).iLength) = slStr 'VBC NR
       End If
       Next ilLoop 'VBC NR
      gCreateFixedRec = slRecord 'VBC NR

End Function 'VBC NR

Public Function gCreateComDelRec(ilRecLen As Integer, tlComDelRec() As ComDelRec) As String

       Dim ilLoop As Integer
       Dim slRecord As String
       Dim slStr As String
       Dim slInt As String
       Dim slQuote As String
       Dim ilFlag As Integer
       slQuote = """"
       ilFlag = 0
        For ilLoop = LBound(tlComDelRec) To UBound(tlComDelRec) Step 1
            slStr = tlComDelRec(ilLoop).sFieldValue
            slInt = tlComDelRec(ilLoop).sFieldNum
            If slStr <> "" Then
            ilFlag = 1
                slRecord = slRecord & slQuote & slStr & slQuote & ","

            End If
            If slInt <> "" Then
            ilFlag = 1
                slRecord = slRecord & slInt & ","
            End If
            If ilFlag = 0 Then
                slRecord = slRecord & slQuote & slQuote & ","
            End If
            ilFlag = 0
        Next ilLoop

        slRecord = Left$(slRecord, Len(slRecord) - 1)
        gCreateComDelRec = slRecord

End Function


Sub gUnpackDateEightChar(ilDyMn As Integer, ilYear As Integer, slDate As String) 'VBC NR

       Dim ilDy As Integer 'Day # 'VBC NR
       Dim ilMn As Integer 'Month # 'VBC NR
       Dim dlDateSerial As Double 'VBC NR
       Dim ilAdjYear As Integer 'VBC NR
       Dim slSortDate As String 'VBC NR
       If (ilDyMn = 0) And (ilYear = 0) Then 'VBC NR
           slDate = "" 'VBC NR
           Exit Sub 'VBC NR
       End If 'VBC NR
       ilDy = ilDyMn And &HFF 'Obtain day # 'VBC NR
       ilMn = ilDyMn \ 256  'Obtain month 'VBC NR
       ilAdjYear = ilYear 'VBC NR
       If (ilAdjYear >= 0) And (ilAdjYear <= 69) Then 'VBC NR
           ilAdjYear = 2000 + ilAdjYear 'VBC NR
       ElseIf (ilAdjYear >= 70) And (ilAdjYear <= 99) Then 'VBC NR
          ilAdjYear = 1900 + ilAdjYear 'VBC NR
       End If 'VBC NR
      dlDateSerial = DateSerial(ilAdjYear, ilMn, ilDy) 'VBC NR
      slSortDate = Format$(dlDateSerial, "mmddyyyy") 'VBC NR
      slDate = slSortDate 'VBC NR

End Sub 'VBC NR
Public Sub gUnpackDateByFormat(ilDyMn As Integer, ilYear As Integer, slFormatDate As String, slOutDate As String) 'VBC NR
       Dim slDate As String 'VBC NR
       gUnpackDate ilDyMn, ilYear, slDate 'VBC NR
       slOutDate = Format$(slDate, slFormatDate) 'VBC NR

End Sub 'VBC NR
Function DoesFileExist(strFileSpec As String) As Boolean
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Constants (Removed)                                                              *
'*  INVALID_ARGUMENT                                                                      *
'******************************************************************************************


       On Error GoTo DoesfileExist_Err
       If (GetAttr(strFileSpec) And vbDirectory) <> vbDirectory Then
           DoesFileExist = CBool(Len(Dir(strFileSpec)) > 0)
       Else
           DoesFileExist = False
       End If
DoesfileExist_End:
       Exit Function
DoesfileExist_Err:
       DoesFileExist = False
       Resume DoesfileExist_End

End Function
Function IsReallyNumeric(str As String) As Boolean 'VBC NR

       Dim i As Integer 'VBC NR
       Dim d As String 'VBC NR
       IsReallyNumeric = True 'VBC NR
       For i = 1 To Len(str) 'VBC NR
           d = Mid(str, i, 1) 'VBC NR
           If Asc(d) < 48 Or Asc(d) > 57 Then 'VBC NR
               IsReallyNumeric = False 'VBC NR
               Exit For 'VBC NR
           End If 'VBC NR
       Next 'VBC NR

End Function 'VBC NR
Function NoComma(str As String) As Boolean 'VBC NR

       Dim i As Integer 'counter 'VBC NR
       Dim d As String ' one character 'VBC NR
       NoComma = True 'VBC NR
       For i = 1 To Len(str) 'VBC NR
           d = Mid(str, i, 1) 'VBC NR
           If Asc(d) = 44 Then 'VBC NR
               NoComma = False 'VBC NR
               Exit For 'VBC NR
           End If 'VBC NR
       Next 'VBC NR

End Function 'VBC NR
Function GetState(str As String) As String 'VBC NR

       Dim i As Integer 'counter 'VBC NR
       Dim d As String ' one character at a time 'VBC NR
       GetState = "" 'VBC NR
       For i = Len(str) To 1 Step -1 'VBC NR
           d = Mid(str, i, 1) 'VBC NR
           If Asc(d) = 32 Then 'VBC NR
               i = i - 1 ' check to see if there is another space 'VBC NR
               d = Mid(str, i, 1) 'VBC NR
               If Asc(d) = 32 Then 'VBC NR
                   i = i - 1 ' check to see if there is another space 'VBC NR
                   d = Mid(str, i, 1) 'VBC NR
                  If Asc(d) = 32 Then 'VBC NR
                      i = i - 1 ' check to see if there is another space 'VBC NR
                      d = Mid(str, i, 1) 'VBC NR
                      If Asc(d) = 32 Then 'VBC NR
                       Else 'VBC NR
                          GetState = Mid$(str, i - 1, 2) 'VBC NR
                           Exit For 'VBC NR
                       End If 'VBC NR
                   Else 'VBC NR
                      GetState = Mid$(str, i - 1, 2) 'VBC NR
                       Exit For 'VBC NR
                   End If 'VBC NR
               Else 'VBC NR
                  GetState = Mid$(str, i - 1, 2) 'VBC NR
                   Exit For 'VBC NR
               End If 'VBC NR
           Else 'VBC NR
           End If 'VBC NR
      Next 'VBC NR

End Function 'VBC NR
