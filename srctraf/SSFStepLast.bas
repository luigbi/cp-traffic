Attribute VB_Name = "SSFStepLast"
Public Function gSSFStepLast(hlSsf As Integer, tlSsf As SSF, ilSsfRecLen As Integer, ilLock As Integer) As Integer
    Dim ilRet As Integer
    
    'imSsfRecLen = Len(tmSsf)
    'ReDim bgByteArray(LenB(tmSsf))
    'ilRet = btrGetFirst(hmSsf, bgByteArray(0), imSsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    gSSFStepLast = btrStepLast(hlSsf, tlSsf, ilSsfRecLen, ilLock)   'Get first record as starting point of extend operation
End Function

