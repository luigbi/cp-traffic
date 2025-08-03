Attribute VB_Name = "ExpGen"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Expsubs.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Type SPOTTIMES
    iVefCode As Integer
    sCodeStn As String * 5
    lAirDate As Long   'Air Date
    lAvailTime As Long  'Avail Time
    lNextSpotTime As Long   'Next spot time
End Type

Type EXPRECIMAGE
    sKey As String * 25 'Date (1-6), time (7-12) , sorttype (A B or C) (13), LLCIndex (14-17), subNumber for Audio Vault multi-comments (18-20)
                        'LLC index and LLC sub-index for AudioVault RPS only
    'sRecord As String * 140     '4-7-06 chg from 118 to 140 for prophet nextgen extra fields
    sRecord As String * 255     '4-7-06 chg from 118 to 140 for prophet nextgen extra fields
End Type

Type AIRSELLLINK
    iAirVefCode As Integer
    iSellVefCode As Integer
End Type

Type EXPTMP2INFO
    iVefCode As Integer 'Selected vehicle
    lCrfCode As Long
End Type

Type PROGTIMERANGE
    iVefCode As Integer
    sProgID As String * 8
    lDate As Long
    iGameNo As Integer
    lStartTime As Long
    lEndTime As Long
End Type

Type BREAKBYPROG
    iVefCode As Integer
    iBreakNo As Integer
    iPositionNo As Integer
End Type

'Copy/Product
Dim tmCpf As CPF
Dim tmCpfSrchKey As LONGKEY0 'CPEF key record image
'Media code record information
Dim tmMcfSrchKey As INTKEY0 'MCF key record image
Dim tmMcf As MCF            'MCF record image
'Copy inventory record information
Dim tmCifSrchKey As LONGKEY0 'CIF key record image
Dim tmCif As CIF            'CIF record image
Dim tmTzfSrchKey As LONGKEY0 'TZF key record image
Dim tmTzf As TZF            'TZF record image
Dim tmLcf As LCF
Dim tmLcfSrchKey2 As LCFKEY2
Dim imLcfRecLen As Integer
Dim tmLvf As LVF
Dim tmLvfSrchKey As LONGKEY0
Dim imLvfRecLen As Integer

'8/17/21 - JW - TTP 10233 - Audacy: line summary export
Type EXPWOINVLN
   lChfCode As Long
   lContract_Number As Long
   lExternal_Version_Number As Long
   iProposal_Version_Number As Integer
   sContract_Type As String
   sSalesperson As String
   sSalesperson_email As String
   iSalesperson_ID As Integer
   sSales_Office As String
   iSales_Office_ID As Integer
   sAgency_Name As String
   iAgency_ID As Integer 'JW 6/13/23 incorrect type found
   sExternal_Agency_ID As String
   sAdvertiser_Name As String
   iAdvertiser_ID As Integer 'JW 6/13/23 incorrect type found
   sExternal_Advertiser_ID As String 'JW 6/13/23 incorrect type found
   sProduct_Name As String
   sCashTrade As String
   iTrade_Percentage As Integer
   sAir_TimeNTR As String
   sDemo As String
   sStatus As String
   sRevenue_Set_1 As String
   sRevenue_Set_2 As String
   sRevenue_Set_3 As String
   sVehicle_Name As String
   iVehicle_ID As Integer
   sMarket As String
   sResearch As String
   sSubCompany As String
   sFormat  As String
   sSubtotals As String
   iSpot_Length As Integer
   sDaypart As String
   sLine_Type As String
   sLine_Start_Date As String
   sLine_End_Date As String
   iTotal_Units As Integer
   sPrice_Type As String
   'lTotal_Gross As Long
   dTotal_Gross As Double
   iRating As Integer
   iLine_CPP As Long
   lLine_GRPs As Long
   lCPM As Long
   dAverage_Audience As Long
   iLine_Gross_Impressions As Long
   sNTR_Billing_Date As String
   sNTR_Description As String
   sNTR_Type As String
   iAmount_Per_NTR_Item As Long
   iNumber_of_NTR_Items As Long
   iPkLineNo As Integer
   iDnfCode As Integer
   sCBS As String
   lPop As Long
   'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
   sSpotAirDate As String
   sSpotAirTime As String
   sSpotAudioType As String
   sISCIcode As String
   sDay As String
   lLineCxfComment As Long
   sFormulaComment As String
End Type

'*****************************************************************
'*
'*      Procedure Name:mObtainCopy
'*
'*             Created:3/01/94       By:D. LeVine
'*            Modified:              By:
'*
'*            Comments: Obtain Copy
'*
'*      7/5/01 Return creative title for prophet export
'       10-12-06 Return length of the media code definition
'           In Audio Vault, there is an option to include/
'           exclude the media definition.  Need to know
'           how many characters to exclude from copy string
'           if option is to exclude media code
'******************************************************************
Function gObtainCopy(slZone As String, tmSdf As SDF, hmMcf As Integer, hmTzf As Integer, hmCif As Integer, hmCpf As Integer, ilWegenerOLA As Integer, slCifName As String, slCreativeTitle As String, ilMediaCodeLen As Integer, slISCI As String, slMcfPrefix As String, slMediaCodeSuppressSpot As String, slMcfCode As String) As Integer
'
'   gObtainCopy
'       Where:
'           tmSdf(I)- Spot record
'           tmCif(O)- Inventory
'           tmCpf(O)- Product/ISCI record
'           ilMediaCodeLen(O) - # of char for the media code used to format the copy string
'                       when option used to include media definition or not
'   1-19-12 Return Media code Prefix
'
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim ilCifFound As Integer
    Dim ilMcf As Integer
    Dim blMcfFound As Boolean
    
    slCifName = ""                  'init in case non found
    slCreativeTitle = ""            'init in case none found
    slISCI = ""
    slMcfPrefix = ""                '1-19-12 media code prefix
    slMcfCode = ""                  '11-5-20 media code
    ilCifFound = False
    tmCpf.sISCI = ""
    tmCpf.sName = ""
    tmCpf.sCreative = ""
    tmMcf.sName = ""
    tmMcf.sPrefix = ""
    ilMediaCodeLen = 0
    slMediaCodeSuppressSpot = "N"
    If tmSdf.sPtType = "1" Then  '  Single Copy
        ' Read CIF using lCopyCode from SDF
        tmCifSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmCif, tmCif, Len(tmCif), tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilCifFound = True
        End If
    ElseIf tmSdf.sPtType = "2" Then  '  Combo Copy
    ElseIf tmSdf.sPtType = "3" Then  '  Time Zone Copy
        ' Read TZF using lCopyCode from SDF
        tmTzfSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmTzf, tmTzf, Len(tmTzf), tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        ' Look for the first positive lZone value
        For ilIndex = 1 To 6 Step 1
            If tmTzf.lCifZone(ilIndex - 1) > 0 Then ' Process just the first positive Zone
                If StrComp(tmTzf.sZone(ilIndex - 1), slZone, 1) = 0 Then
                    ' Read CIF using lCopyCode from SDF
                    tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex - 1)
                    ilRet = btrGetEqual(hmCif, tmCif, Len(tmCif), tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        ilCifFound = True
                    End If
                    Exit For
                End If
            End If
        Next ilIndex
        If Not ilCifFound Then
            For ilIndex = 1 To 6 Step 1
                If tmTzf.lCifZone(ilIndex - 1) > 0 Then ' Process just the first positive Zone
                    If StrComp(tmTzf.sZone(ilIndex - 1), "Oth", 1) = 0 Then
                        ' Read CIF using lCopyCode from SDF
                        tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex - 1)
                        ilRet = btrGetEqual(hmCif, tmCif, Len(tmCif), tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            ilCifFound = True
                        End If
                        Exit For
                    End If
                End If
            Next ilIndex
        End If
    End If
    If ilCifFound Then
        If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
            slCifName = tmCif.sName           'return to caller
        Else
            '12/29/08:  Use reel number as cart number
            If ilWegenerOLA Then
                slCifName = Trim$(tmCif.sReel) '""
            End If
        End If
        ' Read CPF using lCpfCode from CIF
        If tmCif.lcpfCode > 0 Then  'see if ISCI exists
            tmCpfSrchKey.lCode = tmCif.lcpfCode
            ilRet = btrGetEqual(hmCpf, tmCpf, Len(tmCpf), tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                tmCpf.sISCI = ""
                tmCpf.sName = ""
                tmCpf.sCreative = ""
            Else
                If (tgSpf.sUseCartNo = "N") Or (tmCif.iMcfCode = 0) Then
                    If Not ilWegenerOLA Then
                        slCifName = tmCpf.sISCI
                    End If
                End If
                slCreativeTitle = Left$(Trim$(tmCpf.sCreative), 30)     '6-25-12 chg from 20 to 30
                slISCI = tmCpf.sISCI
            End If
        Else
            tmCpf.sISCI = ""
            tmCpf.sName = ""
            tmCpf.sCreative = ""
        End If
        blMcfFound = False
        If tmCif.iMcfCode <> 0 Then
            For ilMcf = LBound(tgMCF) To UBound(tgMCF) - 1 Step 1
                If tmCif.iMcfCode = tgMCF(ilMcf).iCode Then
                    If tgMCF(ilMcf).sSuppressOnExport = "Y" Then
                        slMediaCodeSuppressSpot = "Y"
                    End If
                    tmMcf = tgMCF(ilMcf)
                    blMcfFound = True
                    Exit For
                End If
            Next ilMcf
        End If
        If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
            'tmMcfSrchKey.iCode = tmCif.iMcfCode
            'ilRet = btrGetEqual(hmMcf, tmMcf, Len(tmMcf), tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            'If ilRet <> BTRV_ERR_NONE Then
            If Not blMcfFound Then
                tmMcf.sName = "C"
                tmMcf.sPrefix = "C"
                gObtainCopy = False
                Exit Function
            End If
            slCifName = Trim$(tmMcf.sName) & slCifName
            ilMediaCodeLen = Len(Trim$(tmMcf.sName))        'len of the media definition.  may need to exclude it from copystring
            slMcfPrefix = Trim$(tmMcf.sPrefix)              '1-19-12
            slMcfCode = Trim$(tmMcf.sName)                '12-05-20
            gObtainCopy = True
            Exit Function
        Else
            gObtainCopy = True
            Exit Function
        End If
    End If
    
    
    gObtainCopy = False
    Exit Function
End Function

Public Function gGetProgramTimes(hlLcf As Integer, hlLvf As Integer, ilInVefCode As Integer, slStartDate As String, slEndDate As String, tlProgTimeRange() As PROGTIMERANGE, Optional blTestMergeOption = True)
    Dim ilVefCode As Integer
    Dim ilVpf As Integer
    Dim llDate As Long
    Dim ilDay As Integer
    Dim llStartTime As Long
    Dim ilGameNo As Integer
    Dim ilLcf As Integer
    Dim ilVff As Integer
    Dim slProgCodeID As String
    Dim blFound As Boolean
    
    On Error GoTo ErrHand
    gGetProgramTimes = False
    'ReDim tlProgTimeRange(0 To 0) As PROGTIMERANGE
    If blTestMergeOption Then
        ilVefCode = ilInVefCode
        ilVpf = gBinarySearchVpf(CLng(ilInVefCode))
        If ilVpf <> -1 Then
            If (Asc(tgVpf(ilVpf).sUsingFeatures2) And XDSAPPLYMERGE) <> XDSAPPLYMERGE Then
                gGetProgramTimes = True
                Exit Function
            End If
        End If
    Else
        ilVefCode = ilInVefCode
    End If
    slProgCodeID = ""
    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
        If ilVefCode = tgVff(ilVff).iVefCode Then
            slProgCodeID = Trim$(tgVff(ilVff).sXDProgCodeID)
        End If
    Next ilVff
    If (slProgCodeID = "") Or (UCase(slProgCodeID) = "MERGE") Then
        gGetProgramTimes = True
        Exit Function
    End If
    blFound = False
    imLcfRecLen = Len(tmLcf)
    tmLcfSrchKey2.iVefCode = ilVefCode
    gPackDate slStartDate, tmLcfSrchKey2.iLogDate(0), tmLcfSrchKey2.iLogDate(1)
    ilRet = btrGetGreaterOrEqual(hlLcf, tmLcf, imLcfRecLen, tmLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = ilVefCode)
        gUnpackDateLong tmLcf.iLogDate(0), tmLcf.iLogDate(1), llDate
        If gDateValue(slEndDate) > llDate Then
            Exit Do
        End If
        If (tmLcf.sStatus = "C") Then
            blFound = True
            ilGameNo = tmLcf.iType
            For ilLcf = LBound(tmLcf.lLvfCode) To UBound(tmLcf.lLvfCode) Step 1
                If tmLcf.lLvfCode(ilLcf) > 0 Then
                    gUnpackTimeLong tmLcf.iTime(0, ilLcf), tmLcf.iTime(1, ilLcf), False, llStartTime
                    mSetEventTimes hlLvf, ilVefCode, llDate, ilGameNo, tmLcf.lLvfCode(ilLcf), llStartTime, slProgCodeID, tlProgTimeRange()
                End If
            Next ilLcf
        End If
        ilRet = btrGetNext(hlLcf, tmLcf, imLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If Not blFound Then
        ilVef = gBinarySearchVef(ilVefCode)
        If ilVef <> -1 Then
            If tgMVef(ilVef).sType = "A" Then
                llLDate = gGetLatestLCFDate(hlLcf, "C", ilVefCode)
                If gDateValue(slStartDate) > llLDate Then
                    ilDay = gWeekDayStr(slStartDate)
                    tmLcfSrchKey2.iVefCode = ilVefCode
                    tmLcfSrchKey2.iLogDate(0) = ilDay + 1
                    tmLcfSrchKey2.iLogDate(1) = 0
                    ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get current record
                    If (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = ilVefCode) Then
                        If (tmLcf.iLogDate(0) = ilDay + 1) And (tmLcf.iLogDate(1) = 0) Then
                            If ilDay + 1 = tmLcf.iLogDate(0) Then
                                ilGameNo = tmLcf.iType
                                For ilLcf = LBound(tmLcf.lLvfCode) To UBound(tmLcf.lLvfCode) Step 1
                                    If tmLcf.lLvfCode(ilLcf) > 0 Then
                                        gUnpackTimeLong tmLcf.iTime(0, ilLcf), tmLcf.iTime(1, ilLcf), False, llStartTime
                                        mSetEventTimes hlLvf, ilVefCode, llDate, ilGameNo, tmLcf.lLvfCode(ilLcf), llStartTime, slProgCodeID, tlProgTimeRange()
                                    End If
                                Next ilLcf
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    gGetProgramTimes = True
    Exit Function
    
ErrHand:
    gMsg = ""
    Screen.MousePointer = vbDefault
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then
            gMsg = "A SQL error has occured in modAgmnt-gGetProgramTimes: "
            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in modAgmnt-gGetProgramTimes: "
        gMsgBox gMsg & err.Description & "; Error #" & err.Number, vbCritical
    End If
End Function

Private Sub mSetEventTimes(hlLvf As Integer, ilVefCode As Integer, llDate As Long, ilGameNo As Integer, llLvfCode As Long, llLcfStartTime As Long, slProgCodeID As String, tlProgTimeRange() As PROGTIMERANGE)
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llLength As Long
    Dim slLength As String
    Dim ilUpper As Integer
    
    On Error GoTo mGetEventTimeErr
    If llLvfCode <= 0 Then
        Exit Sub
    End If
    
    llStartTime = llLcfStartTime
    
    imLvfRecLen = Len(tmLvf)
    tmLvfSrchKey.lCode = llLvfCode
    ilRet = btrGetEqual(hlLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
    If ilRet = BTRV_ERR_NONE Then
        gUnpackTimeLong tmLvf.iLen(0), tmLvf.iLen(1), False, llLength
        llEndTime = llStartTime + llLength
        ilUpper = UBound(tlProgTimeRange)
        tlProgTimeRange(ilUpper).iVefCode = ilVefCode
        tlProgTimeRange(ilUpper).sProgID = ""
        tlProgTimeRange(ilUpper).sProgID = slProgCodeID
        tlProgTimeRange(ilUpper).lDate = llDate
        tlProgTimeRange(ilUpper).iGameNo = ilGameNo
        tlProgTimeRange(ilUpper).lStartTime = llStartTime
        tlProgTimeRange(ilUpper).lEndTime = llEndTime
        ReDim Preserve tlProgTimeRange(0 To ilUpper + 1) As PROGTIMERANGE
    End If
    Exit Sub
mGetEventTimeErr:
    Exit Sub
End Sub
