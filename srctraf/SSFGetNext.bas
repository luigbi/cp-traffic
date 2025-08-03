Attribute VB_Name = "SSFGetNext"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SSFGetNext.bas on Wed 6/17/09 @ 12:56
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
Public Function gSSFGetNext(hlSsf As Integer, tlSsf As SSF, ilSsfRecLen As Integer, ilLock As Integer, ilReadOnly As Integer) As Integer

    'imSsfRecLen = Len(tmSsf) 'Max size of variable length record
    'ReDim bgByteArray(LenB(tmSsf))
    'HMemCpy bgByteArray(0), tmSsf, LenB(tmSsf)
    'ilRet = btrGetNext(hmSsf, bgByteArray(0), imSsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    ilSsfRecLen = Len(tlSsf)
    gSSFGetNext = btrGetNext(hlSsf, tlSsf, ilSsfRecLen, ilLock, ilReadOnly)
End Function

