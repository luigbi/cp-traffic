Attribute VB_Name = "SSFGetPosition"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SSFGetPosition.bas on Wed 6/17/09 @ 1
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
Public Function gSSFGetPosition(hlSsf As Integer, llRecPos As Long) As Integer

    'imSsfRecLen = Len(tmSsf) 'Max size of variable length record
    'ReDim bgByteArray(LenB(tmSsf))
    'HMemCpy bgByteArray(0), tmSsf, LenB(tmSsf)
    'ilRet = btrGetDirect(hmSsf, bgByteArray(0), imSsfRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    gSSFGetPosition = btrGetPosition(hlSsf, llRecPos)
End Function

