Attribute VB_Name = "SSFUpdate"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SSFUpdate.bas on Wed 6/17/09 @ 12:56
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS

Public Function gSSFUpdate(hlSsf As Integer, tlSsf As SSF, ilSsfRecLen As Integer) As Integer

    'ReDim bgByteArray(LenB(tmSsf))
    'HMemCpy bgByteArray(0), tmSsf, LenB(tmSsf)
    'ilRet = btrUpdate(hmSsf, bgByteArray(0), imSsfRecLen)
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    gSetFilledBreak tlSsf
    gSSFUpdate = btrUpdate(hlSsf, tlSsf, ilSsfRecLen)
End Function


Private Sub gSetFilledBreak(tlSsf As SSF)
    Dim ilDiff As Integer
    Dim slDate As String
    Dim ilEvt As Integer
    Dim ilTempEvt As Integer
    Dim ilSpot As Integer
    Dim ilUnits As Integer
    Dim ilLens As Integer
    Dim ilVpf As Integer
    Dim ilAnf As Integer
    Dim blUnitsOnly As Boolean
    Dim llTime1 As Long
    Dim llTime2 As Long
    
    tlSsf.iFillRequired = 0
    gUnpackDate tlSsf.iDate(0), tlSsf.iDate(1), slDate
    ilDiff = DateDiff("d", gNow(), slDate)  'slDate - gNow()
    If (ilDiff < 0) Or (ilDiff > 18) Then
        Exit Sub
    End If
    ilVpf = gBinarySearchVpf(tlSsf.iVefCode)
    If ilVpf = -1 Then
        Exit Sub
    End If
    blUnitsOnly = False
    If tgVpf(ilVpf).sSSellOut <> "B" Then
        blUnitsOnly = True
    End If
    ilEvt = 1
    Do While ilEvt <= tlSsf.iCount
       LSet tmAvail = tlSsf.tPas(ADJSSFPASBZ + ilEvt)
        If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 2) Then
            ilAnf = gBinarySearchAnf(tmAvail.ianfCode, tgAvailAnf())
            If ilAnf <> -1 Then
                ilLens = 0
                ilUnits = 0
                For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                   LSet tmSpot = tlSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)
                    If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                        ilLens = ilLens + tmSpot.iPosLen And &HFFF
                        ilUnits = ilUnits + 1
                    End If
                Next ilSpot
                ilEvt = ilEvt + tmAvail.iNoSpotsThis
                If (tgAvailAnf(ilAnf).sFillRequired = "Y") Or ((tgAvailAnf(ilAnf).sFillRequired = "B") And (ilUnits > 0)) Then
                    If blUnitsOnly Then
                        If (ilUnits < (tmAvail.iAvInfo And &H1F)) Then
                            tlSsf.iFillRequired = 1
                            Exit Sub
                        End If
                    Else
                        If (ilLens < (tmAvail.iLen)) And (ilUnits < (tmAvail.iAvInfo And &H1F)) Then
                            tlSsf.iFillRequired = 1
                            Exit Sub
                        End If
                    End If
                ElseIf (tgAvailAnf(ilAnf).sFillRequired = "A") Then
                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime1
                    llTime1 = llTime1 + tmAvail.iLen
                    ilTempEvt = ilEvt + 1
                    Do While ilTempEvt <= tlSsf.iCount
                       LSet tmAvail = tlSsf.tPas(ADJSSFPASBZ + ilEvt)
                        If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 2) Then
                            gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime2
                            If llTime1 = llTime2 Then
                                If (ilUnits > 0) And (tmAvail.iNoSpotsThis = 0) Then
                                    tlSsf.iFillRequired = 1
                                    Exit Sub
                                End If
                                If (ilUnits = 0) And (tmAvail.iNoSpotsThis > 0) Then
                                    tlSsf.iFillRequired = 1
                                    Exit Sub
                                End If
                            End If
                        End If
                    Loop
                End If
            End If
        End If
        
        ilEvt = ilEvt + 1
    Loop
End Sub
