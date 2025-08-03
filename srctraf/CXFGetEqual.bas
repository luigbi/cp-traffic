Attribute VB_Name = "CXFGetEqual"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CXFGetEqual.bas on Wed 6/17/09 @ 12:5
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit

Public Function gCXFGetEqual(hlCxf As Integer, tlCxf As CXF, ilCxfRecLen As Integer, tlCxfSrchKey As LONGKEY0, ilKeyNo As Integer, ilLock As Integer, ilForUpdate As Integer) As Integer

    'imSsfRecLen = Len(tmSsf)
    'ReDim bgByteArray(LenB(tmSsf))
    'ilRet = btrGetFirst(hmSsf, bgByteArray(0), imSsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    gCXFGetEqual = btrGetEqual(hlCxf, tlCxf, ilCxfRecLen, tlCxfSrchKey, ilKeyNo, ilLock, ilForUpdate)    'Get first record as starting point of extend operation
End Function

