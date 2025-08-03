Attribute VB_Name = "SSFGetGreaterOrEqual"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SSFGetGreaterOrEqual.bas on Wed 6/17/
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  gSSFGetGreaterOrEqualKey1                                                             *
'******************************************************************************************

Option Explicit
Option Compare Text
Public Function gSSFGetGreaterOrEqual(hlSsf As Integer, tlSsf As SSF, ilSsfRecLen As Integer, tlSsfSrchKey As SSFKEY0, ilKeyNo As Integer, ilLock As Integer) As Integer

    'imSsfRecLen = Len(tmSsf)
    'ReDim bgByteArray(LenB(tmSsf))
    'ilRet = btrGetFirst(hmSsf, bgByteArray(0), imSsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    ilSsfRecLen = Len(tlSsf)
    gSSFGetGreaterOrEqual = btrGetGreaterOrEqual(hlSsf, tlSsf, ilSsfRecLen, tlSsfSrchKey, ilKeyNo, ilLock)   'Get first record as starting point of extend operation
End Function

Public Function gSSFGetGreaterOrEqualKey1(hlSsf As Integer, tlSsf As SSF, ilSsfRecLen As Integer, tlSsfSrchKey1 As SSFKEY1, ilKeyNo As Integer, ilLock As Integer) As Integer 'VBC NR

    'imSsfRecLen = Len(tmSsf)
    'ReDim bgByteArray(LenB(tmSsf))
    'ilRet = btrGetFirst(hmSsf, bgByteArray(0), imSsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    ilSsfRecLen = Len(tlSsf)
    gSSFGetGreaterOrEqualKey1 = btrGetGreaterOrEqual(hlSsf, tlSsf, ilSsfRecLen, tlSsfSrchKey1, ilKeyNo, ilLock)   'Get first record as starting point of extend operation 'VBC NR
End Function 'VBC NR

Public Function gSSFGetGreaterOrEqualKey2(hlSsf As Integer, tlSsf As SSF, ilSsfRecLen As Integer, tlSsfSrchKey2 As SSFKEY2, ilKeyNo As Integer, ilLock As Integer) As Integer

    'imSsfRecLen = Len(tmSsf)
    'ReDim bgByteArray(LenB(tmSsf))
    'ilRet = btrGetFirst(hmSsf, bgByteArray(0), imSsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    ilSsfRecLen = Len(tlSsf)
    gSSFGetGreaterOrEqualKey2 = btrGetGreaterOrEqual(hlSsf, tlSsf, ilSsfRecLen, tlSsfSrchKey2, ilKeyNo, ilLock)   'Get first record as starting point of extend operation
End Function


