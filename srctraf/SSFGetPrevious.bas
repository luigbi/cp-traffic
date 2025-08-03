Attribute VB_Name = "SSFGetPrevious"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SSFGetPrevious.bas on Wed 6/17/09 @ 1
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
Public Function gSSFGetPrevious(hlSsf As Integer, tlSsf As SSF, ilSsfRecLen As Integer, ilLock As Integer, ilForUpdate As Integer) As Integer

    'imSsfRecLen = Len(tmSsf)
    'ReDim bgByteArray(LenB(tmSsf))
    'ilRet = btrGetFirst(hmSsf, bgByteArray(0), imSsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    ilSsfRecLen = Len(tlSsf)
    gSSFGetPrevious = btrGetPrevious(hlSsf, tlSsf, ilSsfRecLen, ilLock, ilForUpdate)   'Get first record as starting point of extend operation
End Function


