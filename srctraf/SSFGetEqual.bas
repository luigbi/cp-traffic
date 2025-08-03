Attribute VB_Name = "SSFGetEqual"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SSFGetEqual.bas on Wed 6/17/09 @ 12:5
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
Public Function gSSFGetEqual(hlSsf As Integer, tlSsf As SSF, ilSsfRecLen As Integer, tlSsfSrchKey As SSFKEY0, ilKeyNo As Integer, ilLock As Integer, ilForUpdate As Integer) As Integer

    'imSsfRecLen = Len(tmSsf)
    'ReDim bgByteArray(LenB(tmSsf))
    'ilRet = btrGetFirst(hmSsf, bgByteArray(0), imSsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    ilSsfRecLen = Len(tlSsf)
    gSSFGetEqual = btrGetEqual(hlSsf, tlSsf, ilSsfRecLen, tlSsfSrchKey, ilKeyNo, ilLock, ilForUpdate)    'Get first record as starting point of extend operation
End Function


Public Function gSSFGetEqualKey1(hlSsf As Integer, tlSsf As SSF, ilSsfRecLen As Integer, tlSsfSrchKey1 As SSFKEY1, ilKeyNo As Integer, ilLock As Integer, ilForUpdate As Integer) As Integer

    'imSsfRecLen = Len(tmSsf)
    'ReDim bgByteArray(LenB(tmSsf))
    'ilRet = btrGetFirst(hmSsf, bgByteArray(0), imSsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    ilSsfRecLen = Len(tlSsf)
    gSSFGetEqualKey1 = btrGetEqual(hlSsf, tlSsf, ilSsfRecLen, tlSsfSrchKey1, ilKeyNo, ilLock, ilForUpdate)    'Get first record as starting point of extend operation
End Function

Public Function gSSFGetEqualKey2(hlSsf As Integer, tlSsf As SSF, ilSsfRecLen As Integer, tlSsfSrchKey2 As SSFKEY2, ilKeyNo As Integer, ilLock As Integer, ilForUpdate As Integer) As Integer

    'imSsfRecLen = Len(tmSsf)
    'ReDim bgByteArray(LenB(tmSsf))
    'ilRet = btrGetFirst(hmSsf, bgByteArray(0), imSsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    ilSsfRecLen = Len(tlSsf)
    gSSFGetEqualKey2 = btrGetEqual(hlSsf, tlSsf, ilSsfRecLen, tlSsfSrchKey2, ilKeyNo, ilLock, ilForUpdate)    'Get first record as starting point of extend operation
End Function

Public Function gSSFGetEqualKey3(hlSsf As Integer, tlSsf As SSF, ilSsfRecLen As Integer, tlSsfSrchKey3 As SSFKEY3, ilKeyNo As Integer, ilLock As Integer, ilForUpdate As Integer) As Integer
    ilSsfRecLen = Len(tlSsf)
    gSSFGetEqualKey3 = btrGetEqual(hlSsf, tlSsf, ilSsfRecLen, tlSsfSrchKey3, ilKeyNo, ilLock, ilForUpdate)    'Get first record as starting point of extend operation
End Function

