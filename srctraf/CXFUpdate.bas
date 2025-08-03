Attribute VB_Name = "CXFUpdate"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CXFUpdate.bas on Wed 6/17/09 @ 12:56 
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit

Public Function gCXFUpdate(hlCxf As Integer, tlCxf As CXF, ilCxfRecLen As Integer) As Integer

    'ReDim bgByteArray(LenB(tmSsf))
    'HMemCpy bgByteArray(0), tmSsf, LenB(tmSsf)
    'ilRet = btrUpdate(hmSsf, bgByteArray(0), imSsfRecLen)
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    gCXFUpdate = btrUpdate(hlCxf, tlCxf, ilCxfRecLen)
End Function

