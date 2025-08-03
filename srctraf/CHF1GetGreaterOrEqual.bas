Attribute VB_Name = "CHF1GetGreaterOrEqual"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CHF1GetGreaterOrEqual.bas on Wed 6/17
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit

Public Function gCHF1GetGreaterOrEqual(hlChf As Integer, tlChf As CHF, ilChfRecLen As Integer, tlChfSrchKey As CHFKEY1, ilKeyNo As Integer, ilLock As Integer) As Integer

    'imSsfRecLen = Len(tmSsf)
    'ReDim bgByteArray(LenB(tmSsf))
    'ilRet = btrGetFirst(hmSsf, bgByteArray(0), imSsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    gCHF1GetGreaterOrEqual = btrGetGreaterOrEqual(hlChf, tlChf, ilChfRecLen, tlChfSrchKey, ilKeyNo, ilLock)   'Get first record as starting point of extend operation
End Function

