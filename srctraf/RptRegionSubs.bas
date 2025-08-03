Attribute VB_Name = "RptRegionSubs"
Option Explicit

Dim tmSef As SEF
Dim tmSefSrchKey1 As SEFKEY1  'SEF key record image
Dim imSefRecLen As Integer
Dim tmClf As CLF
Dim imRafRecLen As Integer
Dim tmRaf As RAF
Dim tmRafSrchKey As LONGKEY0



Dim tmSplitNetList() As SHTTLIST

Dim tmTxr As TXR

Type SHTTLIST           'list contains unique vehicle and split network codes
    iVefCode As Integer
    sClfType As String * 1  'if package line, need to obtain the hidden lines
    iPkgLine As Integer     'package line reference (stored in hidden lines)
    sStartDate As String    'start date of sch line
    iLineInx As Integer
    lRafCode As Long
End Type


'Each Unique Region is formed by ANDing one Include Format Category with one Include Other Category and with ALL Exclude Categories
'           Region examples Two Include Format Categories; Three Include Other Categories and two Exclude Categories
'           6 unique regions will be formed
'           Format1^Other1^Exclude1^Exclude2 or Format1^Other2^Exclude1^Exclude2 or Format1^Other3^Exclude1^Exclude2
'           Format2^Other1^Exclude1^Exclude2 or Format2^Other2^Exclude1^Exclude2 or Format2^Other3^Exclude1^Exclude2
Type REGIONDEFINITION
    lRotNo As Long
    lRafCode As Long
    sRegionName As String * 80  'Region Name
    sCategory As String * 1
    sInclExcl As String * 1
    lFormatFirst As Long   'Reference SplitInclude for Format or SplitInclude for All Categories except Format
    lOtherFirst As Long    'Reference Split Includes for all Categories except Format
                                'If the first category is not format, then no link is required with other INCLUDES
                                '   Region:  Each INCLUDE is AND with each EXCLUDE
                                '           Other1^Exclude1^Exclude2 or Other2^Exclude1^Exclude2 or Other3^Exclude1^Exclude2
    lExcludeFirst As Long    'References Excludes that are to be AND with INCLUDES
    sPtType As String * 1
    lCopyCode As Long
    lCrfCode As Long
    lRsfCode As Long
End Type

Type SPLITCATEGORYINFO
    sCategory As String * 1
    sName As String * 40
    iIntCode As Integer
    lLongCode As Long
    lNext As Long
End Type

Public Function mTestCategorybyStation(slInclExcl As String, ilMktCode As Integer, ilMSAMktCode As Integer, slState As String, ilFmtCode As Integer, ilTztCode As Integer, ilShttCode As Integer, tlSplitCategoryInfo As SPLITCATEGORYINFO, slTypeCode As String) As Integer
    slTypeCode = slInclExcl & tlSplitCategoryInfo.sCategory
    Select Case tlSplitCategoryInfo.sCategory
        Case "M"    'DMA Market
            slTypeCode = slTypeCode & Trim$(str$(ilMktCode))
            If tlSplitCategoryInfo.iIntCode = ilMktCode Then
                If slInclExcl <> "E" Then
                    mTestCategorybyStation = True
                Else
                    mTestCategorybyStation = False
                End If
                Exit Function
            End If
        Case "A"    'MSA Market
            slTypeCode = slTypeCode & Trim$(str$(ilMSAMktCode))
            If tlSplitCategoryInfo.iIntCode = ilMSAMktCode Then
                If slInclExcl <> "E" Then
                    mTestCategorybyStation = True
                Else
                    mTestCategorybyStation = False
                End If
                Exit Function
            End If
        Case "N"    'State Name
            slTypeCode = slTypeCode & Trim$(slState)
            If StrComp(Trim$(tlSplitCategoryInfo.sName), Trim$(slState), vbBinaryCompare) = 0 Then

                If slInclExcl <> "E" Then
                    mTestCategorybyStation = True
                Else
                    mTestCategorybyStation = False
                End If
                Exit Function
            End If
        Case "F"    'Format
            slTypeCode = slTypeCode & Trim$(str$(ilFmtCode))
            If tlSplitCategoryInfo.iIntCode = ilFmtCode Then
                If slInclExcl <> "E" Then
                    mTestCategorybyStation = True
                Else
                    mTestCategorybyStation = False
                End If
                Exit Function
            End If
        Case "T"    'Time zone
            slTypeCode = slTypeCode & Trim$(str$(ilTztCode))
            If tlSplitCategoryInfo.iIntCode = ilTztCode Then
                If slInclExcl <> "E" Then
                    mTestCategorybyStation = True
                Else
                    mTestCategorybyStation = False
                End If
                Exit Function
            End If
        Case "S"    'Station
            slTypeCode = slTypeCode & Trim$(str$(ilShttCode))
            If tlSplitCategoryInfo.iIntCode = ilShttCode Then
                If slInclExcl <> "E" Then
                    mTestCategorybyStation = True
                Else
                    mTestCategorybyStation = False
                End If
                Exit Function
            End If
    End Select
    If slInclExcl <> "E" Then
        mTestCategorybyStation = False
    Else
        mTestCategorybyStation = True
        Select Case tlSplitCategoryInfo.sCategory
            Case "M"    'DMA Market
                slTypeCode = slInclExcl & tlSplitCategoryInfo.sCategory & Trim$(str$(tlSplitCategoryInfo.iIntCode))
            Case "A"    'MSA arket
                slTypeCode = slInclExcl & tlSplitCategoryInfo.sCategory & Trim$(str$(tlSplitCategoryInfo.iIntCode))
            Case "N"    'State Name
                slTypeCode = slInclExcl & tlSplitCategoryInfo.sCategory & Trim$(tlSplitCategoryInfo.sName)
            Case "F"    'Format
                slTypeCode = slInclExcl & tlSplitCategoryInfo.sCategory & Trim$(str$(tlSplitCategoryInfo.iIntCode))
            Case "T"    'Time zone
                slTypeCode = slInclExcl & tlSplitCategoryInfo.sCategory & Trim$(str$(tlSplitCategoryInfo.iIntCode))
            Case "S"    'Station
                slTypeCode = slInclExcl & tlSplitCategoryInfo.sCategory & Trim$(str$(tlSplitCategoryInfo.iIntCode))
        End Select
    End If
End Function
'
'           Search the table of Stations to find its record in memory
'           Return:  -1 if not found; otherwise record index
Public Function gBinarySearchStation(ilCode As Integer) As Integer
    
    'D.S. 01/16/06
    'Returns the index number of tgStations that matches the ilCode that was passed in
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    llMin = LBound(tgStations)
    llMax = UBound(tgStations) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tgStations(llMiddle).iCode Then
            'found the match
            gBinarySearchStation = llMiddle
            Exit Function
        ElseIf ilCode < tgStations(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchStation = -1
    Exit Function

End Function



Public Sub gBuildSplitCategoryInfo(hlSef As Integer, llUpperSplit As Long, tlRegionDefinition As REGIONDEFINITION, tlSplitCategoryInfo() As SPLITCATEGORYINFO, frm As Form)
    Dim llPreviousOther As Long
    Dim llPreviousFormat As Long
    Dim llPreviousExclude As Long
    Dim slCategory As String
    Dim slInclExcl As String
    Dim ilAddExclude As Integer
    Dim ilShtt As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    
    llPreviousFormat = -1
    llPreviousOther = -1
    llPreviousExclude = -1
    
    tmSefSrchKey1.lRafCode = tlRegionDefinition.lRafCode
    tmSefSrchKey1.iSeqNo = 0
    ilRet = btrGetGreaterOrEqual(hlSef, tmSef, Len(tmSef), tmSefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmSef.lRafCode = tlRegionDefinition.lRafCode)
        slCategory = tlRegionDefinition.sCategory
        slInclExcl = tlRegionDefinition.sInclExcl
        If Trim$(tmSef.sCategory) <> "" Then
            slCategory = tmSef.sCategory
            slInclExcl = tmSef.sInclExcl
        End If
        If slInclExcl <> "E" Then
            ilAddExclude = False
            If slCategory = "F" Then
                'Add to Format table
                If tlRegionDefinition.lFormatFirst = -1 Then
                    tlRegionDefinition.lFormatFirst = llUpperSplit
                End If
                tlSplitCategoryInfo(llUpperSplit).sCategory = slCategory
                tlSplitCategoryInfo(llUpperSplit).sName = tmSef.sName
                tlSplitCategoryInfo(llUpperSplit).iIntCode = tmSef.iIntCode
                tlSplitCategoryInfo(llUpperSplit).lLongCode = tmSef.lLongCode
                tlSplitCategoryInfo(llUpperSplit).lNext = -1
                If llPreviousFormat <> -1 Then
                    tlSplitCategoryInfo(llPreviousFormat).lNext = llUpperSplit
                End If
                llPreviousFormat = llUpperSplit
                'ReDim Preserve tlSplitCategoryInfo(0 To UBound(tlSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                llUpperSplit = llUpperSplit + 1
                If llUpperSplit >= UBound(tlSplitCategoryInfo) Then
                    ReDim Preserve tlSplitCategoryInfo(0 To UBound(tlSplitCategoryInfo) + 100) As SPLITCATEGORYINFO
                End If
            Else
                'Add to All Category except Format table
                If tlRegionDefinition.lOtherFirst = -1 Then
                    tlRegionDefinition.lOtherFirst = llUpperSplit
                End If
                tlSplitCategoryInfo(llUpperSplit).sCategory = slCategory
                tlSplitCategoryInfo(llUpperSplit).sName = tmSef.sName
                tlSplitCategoryInfo(llUpperSplit).iIntCode = tmSef.iIntCode
                tlSplitCategoryInfo(llUpperSplit).lLongCode = tmSef.lLongCode
                tlSplitCategoryInfo(llUpperSplit).lNext = -1
                If llPreviousOther <> -1 Then
                    tlSplitCategoryInfo(llPreviousOther).lNext = llUpperSplit
                End If
                llPreviousOther = llUpperSplit
                'ReDim Preserve tlSplitCategoryInfo(0 To UBound(tlSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                llUpperSplit = llUpperSplit + 1
                If llUpperSplit >= UBound(tlSplitCategoryInfo) Then
                    ReDim Preserve tlSplitCategoryInfo(0 To UBound(tlSplitCategoryInfo) + 100) As SPLITCATEGORYINFO
                End If
            End If
        Else
            'Add to Exclude table
            '7-2-10 always include regardless of wegner or OLA
            ilAddExclude = True
'            ilShtt = gBinarySearchStationInfoByCode(tmSef.iIntCode)
'            If ilShtt <> -1 Then
'                If slTarget = "W" Then
'                    If tgStationInfoByCode(ilShtt).sUsedForWegener <> "Y" Then
'                        ilAddExclude = False
'                    End If
'                ElseIf slTarget = "O" Then
'                    If tgStationInfoByCode(ilShtt).sUsedForOLA <> "Y" Then
'                        ilAddExclude = False
'                    End If
'                End If
'                ilAddExclude = ilAddExclude
'            Else
'                ilAddExclude = False
'            End If
        End If
        If ilAddExclude Then
            If tlRegionDefinition.lExcludeFirst = -1 Then
                tlRegionDefinition.lExcludeFirst = llUpperSplit
            End If
            tlSplitCategoryInfo(llUpperSplit).sCategory = slCategory
            tlSplitCategoryInfo(llUpperSplit).sName = tmSef.sName
            tlSplitCategoryInfo(llUpperSplit).iIntCode = tmSef.iIntCode
            tlSplitCategoryInfo(llUpperSplit).lLongCode = tmSef.lLongCode
            tlSplitCategoryInfo(llUpperSplit).lNext = -1
            If llPreviousExclude <> -1 Then
                tlSplitCategoryInfo(llPreviousExclude).lNext = llUpperSplit
            End If
            llPreviousExclude = llUpperSplit
            'ReDim Preserve tlSplitCategoryInfo(0 To UBound(tlSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
            llUpperSplit = llUpperSplit + 1
            If llUpperSplit >= UBound(tlSplitCategoryInfo) Then
                ReDim Preserve tlSplitCategoryInfo(0 To UBound(tlSplitCategoryInfo) + 100) As SPLITCATEGORYINFO
            End If
        End If
        ilRet = btrGetNext(hlSef, tmSef, Len(tmSef), BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Exit Sub
ErrHand:
    On Error GoTo 0
    'gCPErrorMsg ilRet, "gRegionTestDefinition - SEF ", RptSelSR
    gCPErrorMsg ilRet, "gBuildSplitCategoryInfo-SEF ", frm

    Exit Sub
End Sub
Public Sub gSeparateRegions(tlInRegionDefinition() As REGIONDEFINITION, tlInSplitCategoryInfo() As SPLITCATEGORYINFO, tlOutRegionDefinition() As REGIONDEFINITION, tlOutSplitCategoryInfo() As SPLITCATEGORYINFO)
    'If a region is defined as:
    '(Fmt1 or Fmt2 or Fmt3) and (St1 or St2) and (Not K1111 and Not K222)
    'Convert to:
    'Region 1: Fmt1 and St1 And Not K111 and Not K222
    'Region 2: Fmt1 and St2 And Not K111 and Not K222
    'Region 3: Fmt2 and St1 And Not K111 and Not K222
    'Region 4: Fmt2 and St2 And Not K111 and Not K222
    'Region 5: Fmt3 and St1 And Not K111 and Not K222
    'Region 6: Fmt3 and St2 And Not K111 and Not K222
    Dim llFormatIndex As Long
    Dim llRegion As Long
    Dim llOtherIndex As Long
    Dim llExcludeIndex As Long
    
    For llRegion = 0 To UBound(tlInRegionDefinition) - 1 Step 1
        llFormatIndex = tlInRegionDefinition(llRegion).lFormatFirst
            
        If tlInRegionDefinition(llRegion).lFormatFirst <> -1 Then
            'Test Format
            llFormatIndex = tlInRegionDefinition(llRegion).lFormatFirst
            Do
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)) = tlInRegionDefinition(llRegion)
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lFormatFirst = UBound(tlOutSplitCategoryInfo)
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lOtherFirst = -1
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lExcludeFirst = -1
                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llFormatIndex)
                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                ReDim Preserve tlOutSplitCategoryInfo(0 + UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                If tlInRegionDefinition(llRegion).lOtherFirst <> -1 Then
                    llOtherIndex = tlInRegionDefinition(llRegion).lOtherFirst
                    Do
                        tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lOtherFirst = UBound(tlOutSplitCategoryInfo)
                        tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llOtherIndex)
                        tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                        ReDim Preserve tlOutSplitCategoryInfo(0 + UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        llExcludeIndex = tlInRegionDefinition(llRegion).lExcludeFirst
                        If llExcludeIndex <> -1 Then
                            tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lExcludeFirst = UBound(tlOutSplitCategoryInfo)
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                            ReDim Preserve tlOutSplitCategoryInfo(0 + UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                            llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                            Do While llExcludeIndex <> -1
                                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo) - 1).lNext = UBound(tlOutSplitCategoryInfo)
                                ReDim Preserve tlOutSplitCategoryInfo(0 + UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                                llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                            Loop
                        End If
                        ReDim Preserve tlOutRegionDefinition(0 To UBound(tlOutRegionDefinition) + 1) As REGIONDEFINITION
                        llOtherIndex = tlInSplitCategoryInfo(llOtherIndex).lNext
                        If llOtherIndex <> -1 Then
                            tlOutRegionDefinition(UBound(tlOutRegionDefinition)) = tlInRegionDefinition(llRegion)
                            tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lFormatFirst = UBound(tlOutSplitCategoryInfo)
                            tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lOtherFirst = -1
                            tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lExcludeFirst = -1
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llFormatIndex)
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                            ReDim Preserve tlOutSplitCategoryInfo(0 + UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        End If
                    Loop While llOtherIndex <> -1
                ElseIf tlInRegionDefinition(llRegion).lExcludeFirst <> -1 Then
                    llExcludeIndex = tlInRegionDefinition(llRegion).lExcludeFirst
                    If llExcludeIndex <> -1 Then
                        tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lExcludeFirst = UBound(tlOutSplitCategoryInfo)
                        tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                        tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                        ReDim Preserve tlOutSplitCategoryInfo(0 + UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                        Do While llExcludeIndex <> -1
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo) - 1).lNext = UBound(tlOutSplitCategoryInfo)
                            ReDim Preserve tlOutSplitCategoryInfo(0 + UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                            llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                        Loop
                        ReDim Preserve tlOutRegionDefinition(0 To UBound(tlOutRegionDefinition) + 1) As REGIONDEFINITION
                    End If
                Else
                    ReDim Preserve tlOutRegionDefinition(0 To UBound(tlOutRegionDefinition) + 1) As REGIONDEFINITION
                End If
                llFormatIndex = tlInSplitCategoryInfo(llFormatIndex).lNext
            Loop While llFormatIndex <> -1
        ElseIf tlInRegionDefinition(llRegion).lOtherFirst <> -1 Then
            llOtherIndex = tlInRegionDefinition(llRegion).lOtherFirst
            Do
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)) = tlInRegionDefinition(llRegion)
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lFormatFirst = -1
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lOtherFirst = UBound(tlOutSplitCategoryInfo)
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lExcludeFirst = -1
                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llOtherIndex)
                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                ReDim Preserve tlOutSplitCategoryInfo(0 + UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                llExcludeIndex = tlInRegionDefinition(llRegion).lExcludeFirst
                If llExcludeIndex <> -1 Then
                    tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lExcludeFirst = UBound(tlOutSplitCategoryInfo)
                    tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                    tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                    ReDim Preserve tlOutSplitCategoryInfo(0 + UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                    llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                    Do While llExcludeIndex <> -1
                        tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                        tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                        tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo) - 1).lNext = UBound(tlOutSplitCategoryInfo)
                        ReDim Preserve tlOutSplitCategoryInfo(0 + UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                    Loop
                End If
                ReDim Preserve tlOutRegionDefinition(0 To UBound(tlOutRegionDefinition) + 1) As REGIONDEFINITION
                llOtherIndex = tlInSplitCategoryInfo(llOtherIndex).lNext
            Loop While llOtherIndex <> -1
        Else
            'Exclude only
            llExcludeIndex = tlInRegionDefinition(llRegion).lExcludeFirst
            If llExcludeIndex <> -1 Then
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)) = tlInRegionDefinition(llRegion)
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lFormatFirst = -1
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lOtherFirst = -1
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lExcludeFirst = UBound(tlOutSplitCategoryInfo)
                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                ReDim Preserve tlOutSplitCategoryInfo(0 + UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                Do While llExcludeIndex <> -1
                    tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                    tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                    tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo) - 1).lNext = UBound(tlOutSplitCategoryInfo)
                    ReDim Preserve tlOutSplitCategoryInfo(0 + UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                    llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                Loop
                ReDim Preserve tlOutRegionDefinition(0 To UBound(tlOutRegionDefinition) + 1) As REGIONDEFINITION
            End If
        End If
    Next llRegion

End Sub

Public Function gRegionTestDefinition(ilShttCode As Integer, ilMktCode As Integer, ilMSAMktCode As Integer, slState As String, ilFmtCode As Integer, ilTztCode As Integer, tlRegionDefinition() As REGIONDEFINITION, tlSplitCategoryInfo() As SPLITCATEGORYINFO, llRegionIndex As Long, slGroupInfo As String, frm As Form) As Integer
    Dim ilRet As Integer
    Dim llFormatIndex As Long
    Dim llOtherIndex As Long
    Dim llExcludeIndex As Long
    Dim ilExcludeOk As Integer
    Dim llRegion As Long
    Dim slTotalTypeCode As String
    Dim slTypeCode As String
    Dim ilExitDo As Integer
    
    On Error GoTo ErrHand
    
    gRegionTestDefinition = False
    For llRegion = 0 To UBound(tlRegionDefinition) - 1 Step 1
        slTotalTypeCode = ""
        llFormatIndex = tlRegionDefinition(llRegion).lFormatFirst
            
        If tlRegionDefinition(llRegion).lFormatFirst <> -1 Then
            'Test Format
            ilExitDo = False
            llFormatIndex = tlRegionDefinition(llRegion).lFormatFirst
            Do
                ilRet = mTestCategorybyStation("I", ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, ilShttCode, tlSplitCategoryInfo(llFormatIndex), slTypeCode)
                If ilRet Then
                    If slTotalTypeCode = "" Then
                        slTotalTypeCode = slTypeCode
                    Else
                        slTotalTypeCode = slTotalTypeCode & "|" & slTypeCode
                    End If
                    If tlRegionDefinition(llRegion).lOtherFirst <> -1 Then
                        'Test Other
                        llOtherIndex = tlRegionDefinition(llRegion).lOtherFirst
                        Do
                            ilRet = mTestCategorybyStation("I", ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, ilShttCode, tlSplitCategoryInfo(llOtherIndex), slTypeCode)
                            If ilRet Then
                                slTotalTypeCode = slTotalTypeCode & "|" & slTypeCode
                            Else
                                ilExitDo = True
                                Exit Do
                            End If
                            llOtherIndex = tlSplitCategoryInfo(llOtherIndex).lNext
                        Loop While llOtherIndex <> -1
                    End If
                    'Exclude
                    If Not ilExitDo Then
                        ilExcludeOk = True
                        llExcludeIndex = tlRegionDefinition(llRegion).lExcludeFirst
                        Do While llExcludeIndex <> -1
                            ilRet = mTestCategorybyStation("E", ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, ilShttCode, tlSplitCategoryInfo(llExcludeIndex), slTypeCode)
                            If Not ilRet Then
                                ilExcludeOk = False
                                Exit Do
                            Else
                                slTotalTypeCode = slTotalTypeCode & "|" & slTypeCode
                            End If
                            llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                        Loop
                        If ilExcludeOk Then
                            'ilRet = gGetCopy(tlRegionDefinition(llRegion).sPtType, tlRegionDefinition(llRegion).lCopyCode, tlRegionDefinition(llRegion).lCrfCode, slCartNo, slProduct, slISCI, slCreativeTitle, llCrfCsfCode, llCpfCode)
                            'gRegionTestDefinition = ilRet
                            llRegionIndex = llRegion
                            slGroupInfo = slTotalTypeCode
                            gRegionTestDefinition = True
                            Exit Function
                        Else
                            ilExitDo = True
                        End If
                    End If
                Else
                    ilExitDo = True
                End If
                If ilExitDo Then
                    Exit Do
                End If
                llFormatIndex = tlSplitCategoryInfo(llFormatIndex).lNext
                'Can't have two formats connected
                If llFormatIndex <> -1 Then
                    Exit Do
                End If
            Loop While llFormatIndex <> -1
        ElseIf tlRegionDefinition(llRegion).lOtherFirst <> -1 Then
            ilExitDo = False
            llOtherIndex = tlRegionDefinition(llRegion).lOtherFirst
            Do
                ilRet = mTestCategorybyStation("I", ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, ilShttCode, tlSplitCategoryInfo(llOtherIndex), slTypeCode)
                If ilRet Then
                    If slTotalTypeCode = "" Then
                        slTotalTypeCode = slTypeCode
                    Else
                        slTotalTypeCode = slTotalTypeCode & "|" & slTypeCode
                    End If
                Else
                    ilExitDo = True
                    Exit Do
                End If
                llOtherIndex = tlSplitCategoryInfo(llOtherIndex).lNext
            Loop While llOtherIndex <> -1
            If Not ilExitDo Then
                ilExcludeOk = True
                llExcludeIndex = tlRegionDefinition(llRegion).lExcludeFirst
                Do While llExcludeIndex <> -1
                    ilRet = mTestCategorybyStation("E", ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, ilShttCode, tlSplitCategoryInfo(llExcludeIndex), slTypeCode)
                    If Not ilRet Then
                        ilExcludeOk = False
                        Exit Do
                    Else
                        slTotalTypeCode = slTotalTypeCode & "|" & slTypeCode
                    End If
                    llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                Loop
                If ilExcludeOk Then
                    'ilRet = gGetCopy(tlRegionDefinition(llRegion).sPtType, tlRegionDefinition(llRegion).lCopyCode, tlRegionDefinition(llRegion).lCrfCode, slCartNo, slProduct, slISCI, slCreativeTitle, llCrfCsfCode, llCpfCode)
                    'gRegionTestDefinition = ilRet
                    llRegionIndex = llRegion
                    slGroupInfo = slTotalTypeCode
                    gRegionTestDefinition = True
                    Exit Function
                End If
            End If
        ElseIf tlRegionDefinition(llRegion).lExcludeFirst <> -1 Then
            'Exclude only
            ilExcludeOk = True
            llExcludeIndex = tlRegionDefinition(llRegion).lExcludeFirst
            Do While llExcludeIndex <> -1
                ilRet = mTestCategorybyStation("E", ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, ilShttCode, tlSplitCategoryInfo(llExcludeIndex), slTypeCode)
                If Not ilRet Then
                    ilExcludeOk = False
                    Exit Do
                Else
                    If slTotalTypeCode = "" Then
                        slTotalTypeCode = slTypeCode
                    Else
                        slTotalTypeCode = slTotalTypeCode & "|" & slTypeCode
                    End If
                End If
                llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
            Loop
            If ilExcludeOk Then
                'ilRet = gGetCopy(tlRegionDefinition(llRegion).sPtType, tlRegionDefinition(llRegion).lCopyCode, tlRegionDefinition(llRegion).lCrfCode, slCartNo, slProduct, slISCI, slCreativeTitle, llCrfCsfCode, llCpfCode)
                'gRegionTestDefinition = ilRet
                llRegionIndex = llRegion
                slGroupInfo = slTotalTypeCode
                gRegionTestDefinition = True
                Exit Function
            End If
        End If
    Next llRegion
    slGroupInfo = ""
    Exit Function
ErrHand:
    On Error GoTo 0
    'gCPErrorMsg ilRet, "gRegionTestDefinition - SEF ", RptSelSR
    gCPErrorMsg ilRet, "gRegionTestDefinition - SEF ", frm

   Exit Function
    
End Function

Public Function gBinarySearchMkt(llCode As Long) As Long
    
    'D.S. 01/06/06
    'Returns the index number of tgMarketInfo that matches the mktCode that was passed in
    'Note: for this to work tgMarketInfo was previously be sorted by mktCode
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long

    
    llMin = LBound(tgMarkets)
    llMax = UBound(tgMarkets) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tgMarkets(llMiddle).iCode Then
            'found the match
            gBinarySearchMkt = llMiddle
            Exit Function
        ElseIf llCode < tgMarkets(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    
    gBinarySearchMkt = -1
    Exit Function
    
End Function

Public Function gBinarySearchMSAMkt(llCode As Long) As Long
    
    'D.S. 01/06/06
    'Returns the index number of tgMSAMarketInfo that matches the mktCode that was passed in
    'Note: for this to work tgMSAMarketInfo was previously be sorted by mktCode
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long

    
    llMin = LBound(tgMSAMarkets)
    llMax = UBound(tgMSAMarkets) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tgMSAMarkets(llMiddle).iCode Then
            'found the match
            gBinarySearchMSAMkt = llMiddle
            Exit Function
        ElseIf llCode < tgMSAMarkets(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    
    gBinarySearchMSAMkt = -1
    Exit Function
End Function
Public Function gBinarySearchTzt(ilCode As Integer) As Integer
    
    'Returns the index number of tgTimeZoneInfo that matches the tztCode that was passed in
    'Note: for this to work tgTimeZoneInfo was previously be sorted by tztCode
    
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer

    
    ilMin = LBound(tgTimeZones)
    ilMax = UBound(tgTimeZones) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tgTimeZones(ilMiddle).iCode Then
            'found the match
            gBinarySearchTzt = ilMiddle
            Exit Function
        ElseIf ilCode < tgTimeZones(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    
    gBinarySearchTzt = -1
    Exit Function

End Function

Public Function gBinarySearchSnt(ilCode As Integer) As Integer
    
    'Returns the index number of tgStateInfo that matches the SntCode that was passed in
    'Note: for this to work tgStateInfo was previously be sorted by SntCode
    
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer

    
    ilMin = LBound(tgStates)
    ilMax = UBound(tgStates) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tgStates(ilMiddle).iCode Then
            'found the match
            gBinarySearchSnt = ilMiddle
            Exit Function
        ElseIf ilCode < tgStates(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    
    gBinarySearchSnt = -1
    Exit Function
End Function

Public Function gBinarySearchFmt(ilCode As Integer) As Integer
    
    'Returns the index number of tgFormatInfo that matches the FmtCode that was passed in
    'Note: for this to work tgFormatInfo was previously be sorted by FmtCode
    
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer

    
    ilMin = LBound(tgFormats)
    ilMax = UBound(tgFormats) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tgFormats(ilMiddle).iCode Then
            'found the match
            gBinarySearchFmt = ilMiddle
            Exit Function
        ElseIf ilCode < tgFormats(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    
    gBinarySearchFmt = -1
    Exit Function
End Function
'       gBuildSplitNetStations
'       loop thru the schedule lines and build the list of
'       station/markets included/excluded in the split network
'       <input> hlRaf - Region file handle
'               hlSef - Split entity file handle
'               hlTxr - text file handle
'               hlShf - station file handle
'               hlMkt - market file handle
'               hlVlf - Vehicle links file handle
'               hlAtt - agreement file handle: must have valid agreement for the vehicle/station
'               tlClfList() - array of sched lines
'       <output> TXR records containing list of stations and markets
'
Public Sub gBuildSplitNetStations(hlRaf As Integer, hlSef As Integer, hlTxr As Integer, hlShf As Integer, hlMkt As Integer, hlVlf As Integer, hlAtt As Integer, tlClfList() As CLFLIST, tlChf As CHF)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  tlStationsClf                 ilLast                                                  *
'******************************************************************************************

Dim ilClf As Integer
Dim slClfInclExcl As String
Dim ilFindNetList As Integer
Dim ilFound As Integer
Dim ilUpper As Integer
Dim ilVefIndex As Integer
Dim slVefAndRegion As String
Dim ilRemChar As Integer
Dim ilEnoughRoom As Integer
Dim ilStation As Integer
Dim ilRet As Integer
Dim slStationAndMkt As String
Dim slText As String
'ReDim tlStationClf(1 To 1) As REGIONSTATIONINFO
ReDim tlStationClf(0 To 0) As REGIONSTATIONINFO
ReDim tmSplitNetList(0 To 0) As SHTTLIST
ReDim ilVehForRegion(0 To 0) As Integer
Dim ilLoopForHidden As Integer

    'loop thru the lines for a split network assignment
    For ilClf = LBound(tlClfList) To UBound(tlClfList) - 1
        tmClf = tlClfList(ilClf).ClfRec
          If tmClf.lRafCode > 0 And tmClf.sType <> "H" Then          'a split network exists that is not a hidden line
            ilFound = False
            'build an array of the unique vehicles and split network defined.  Show the list
            'of stations only once for the same vehicle/region
            For ilFindNetList = LBound(tmSplitNetList) To UBound(tmSplitNetList) - 1
                If tmSplitNetList(ilFindNetList).iVefCode = tmClf.iVefCode And tmSplitNetList(ilFindNetList).lRafCode = tmClf.lRafCode Then
                    ilFound = True
                    Exit For
                End If
            Next ilFindNetList
            If Not ilFound Then         'entry not found, add unique entry to list
                ilUpper = UBound(tmSplitNetList)
                tmSplitNetList(ilUpper).lRafCode = tmClf.lRafCode
                tmSplitNetList(ilUpper).iVefCode = tmClf.iVefCode
                tmSplitNetList(ilUpper).sClfType = tmClf.sType
                tmSplitNetList(ilUpper).iPkgLine = tmClf.iLine      'should be non-zro for all hidden lines of a pkg
                gUnpackDate tmClf.iStartDate(0), tmClf.iStartDate(1), tmSplitNetList(ilUpper).sStartDate
                tmSplitNetList(ilUpper).iLineInx = ilClf            'index into list of sch lines
                ReDim Preserve tmSplitNetList(0 To ilUpper + 1) As SHTTLIST
            End If
        End If
    Next ilClf

    tmTxr.lGenTime = lgNowTime
    tmTxr.iGenDate(0) = igNowDate(0)
    tmTxr.iGenDate(1) = igNowDate(1)
    tmTxr.lSeqNo = 0
    tmTxr.iType = 0         'keep this 0 to avoid getting it intermixed with Book Names record for package line (extracted from all hidden lines of a pkg)
    'loop thru the list of unique vehicles/regions
    For ilClf = LBound(tmSplitNetList) To UBound(tmSplitNetList) - 1
        tmClf = tlClfList(tmSplitNetList(ilClf).iLineInx).ClfRec
        slVefAndRegion = ""
        ilRemChar = 130            'max 130 char per record
        'read the Region to see if it should be shown on contract
        tmTxr.lCsfCode = tmClf.lChfCode     'contract code
        slText = ""
        imRafRecLen = Len(tmRaf)
        tmRafSrchKey.lCode = tmClf.lRafCode
        ilRet = btrGetEqual(hlRaf, tmRaf, imRafRecLen, tmRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ((tmRaf.sShowOnOrder = "Y") And (tlChf.sStatus = "H" Or tlChf.sStatus = "O" Or tlChf.sStatus = "G" Or tlChf.sStatus = "N")) Or ((tmRaf.sShowNoProposal = "Y") And (tlChf.sStatus = "W" Or tlChf.sStatus = "C" Or tlChf.sStatus = "I")) Then
            ReDim ilVehForRegion(0 To 0) As Integer
            If tmSplitNetList(ilClf).sClfType = "O" Or tmSplitNetList(ilClf).sClfType = "A" Then            'package,get the associated hidden vehicles
                For ilLoopForHidden = LBound(tlClfList) To UBound(tlClfList) - 1
                    tmClf = tlClfList(ilLoopForHidden).ClfRec
                    If tmClf.sType = "H" And tmClf.iPkLineNo = tmSplitNetList(ilClf).iPkgLine Then
                        ilVehForRegion(UBound(ilVehForRegion)) = tmClf.iVefCode
                        ReDim Preserve ilVehForRegion(0 To UBound(ilVehForRegion) + 1) As Integer
                    End If
                Next ilLoopForHidden
            Else                            'non-package, pass just the vehicle of the sch line
                ilVehForRegion(0) = tmSplitNetList(ilClf).iVefCode
                ReDim Preserve ilVehForRegion(0 To 1) As Integer
            End If            'tlStationsClf() array contains list of station codes
            gBuildStationsFromRAFByCallLetters hlRaf, hlSef, hlVlf, hlAtt, tmSplitNetList(ilClf).lRafCode, tmSplitNetList(ilClf).sStartDate, ilVehForRegion(), slClfInclExcl, tlStationClf(), True
            'tlStationsClf() array contains list of station codes
            'form up the Region Name and Vehicle before generating the list of stations
            ilVefIndex = gBinarySearchVef(tmClf.iVefCode)
            slVefAndRegion = Trim$(tmRaf.sName)
            If ilVefIndex <> -1 Then
                slVefAndRegion = slVefAndRegion & "/" & Trim$(tgMVef(ilVefIndex).sName)
            Else
                slVefAndRegion = slVefAndRegion & "/" & "Vehicle Missing: " & tmClf.iVefCode
            End If
            'first time thru for the region, show the region and vehicle names
            slText = Trim$(slVefAndRegion) & " includes: "

            ilRemChar = ilRemChar - (Len(slText))
            For ilStation = LBound(tlStationClf) To UBound(tlStationClf) - 1
                'get the station call letters and market
                slStationAndMkt = Trim$(tlStationClf(ilStation).sCallLetters) & " " & Trim$(tlStationClf(ilStation).sMarket)
                'determine if enough room remaining in this record to fit the next entry.
                'if not, write out the record and create new one
                ilEnoughRoom = True
                Do While ilEnoughRoom
                    If ilRemChar > Len(Trim$(slStationAndMkt)) Then      'have enough room for this entry
                        slText = slText & Trim$(slStationAndMkt)
                        If ilStation < UBound(tlStationClf) - 1 Then        'not at the end of array, continue with ","
                            slText = slText & ", "
                        Else            'write out the last record of this region
                            tmTxr.sText = slText
                            ilRet = btrInsert(hlTxr, tmTxr, Len(tmTxr), INDEXKEY0)
                            ilRemChar = 130             'max to print on a single line
                            slText = ""
                            tmTxr.lSeqNo = tmTxr.lSeqNo + 1
                        End If
                        ilRemChar = ilRemChar - (Len(Trim$(slStationAndMkt)) + 1) 'Add 1 extra for comma
                        ilEnoughRoom = False
                    Else                'not enough room for current entry, write out the record
                        tmTxr.sText = slText
                        ilRet = btrInsert(hlTxr, tmTxr, Len(tmTxr), INDEXKEY0)
                        ilRemChar = 130             'max to print on a single line
                        slText = ""
                        tmTxr.lSeqNo = tmTxr.lSeqNo + 1
                    End If
                Loop
            Next ilStation

        End If
    Next ilClf

    Erase tlStationClf
    Exit Sub
End Sub
