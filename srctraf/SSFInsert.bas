Attribute VB_Name = "SSFInsert"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SSFInsert.bas on Wed 6/17/09 @ 12:56
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
Public Function gSSFInsert(hlSsf As Integer, tlSsf As SSF, ilSsfRecLen As Integer, ilKeyNo As Integer) As Integer

    'imSsfRecLen = Len(tmSsf)
    'ReDim bgByteArray(LenB(tmSsf))
    'ilRet = btrGetFirst(hmSsf, bgByteArray(0), imSsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    tlSsf.lCode = 0
    gSSFInsert = btrInsert(hlSsf, tlSsf, ilSsfRecLen, ilKeyNo)   'Get first record as starting point of extend operation
End Function


