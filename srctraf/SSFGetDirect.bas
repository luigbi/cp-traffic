Attribute VB_Name = "SSFGetDirect"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SSFGetDirect.bas on Wed 6/17/09 @ 12:
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text

Public Function gSSFGetDirect(hlSsf As Integer, tlSsf As SSF, ilSsfRecLen As Integer, llRecPos As Long, ilKeyNo As Integer, ilLock As Integer) As Integer

    'imSsfRecLen = Len(tmSsf) 'Max size of variable length record
    'ReDim bgByteArray(LenB(tmSsf))
    'HMemCpy bgByteArray(0), tmSsf, LenB(tmSsf)
    'ilRet = btrGetDirect(hmSsf, bgByteArray(0), imSsfRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    ilSsfRecLen = Len(tlSsf)
    gSSFGetDirect = btrGetDirect(hlSsf, tlSsf, ilSsfRecLen, llRecPos, ilKeyNo, ilLock)
End Function
