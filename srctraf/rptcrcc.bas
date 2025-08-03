Attribute VB_Name = "RptCrCC"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptcrcc.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer

Dim tmVpf As VPF                'VPF record image


Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim tmGrf As GRF
Dim imTerminate As Integer  'True = terminating task, False= OK


Sub mObtainCodes(ilListIndex As Integer, lbcListBox() As SORTCODE, ilIncludeCodes, ilUseCodes() As Integer)
Dim ilHowManyDefined As Integer
Dim ilHowMany As Integer
Dim slNameCode As String
Dim ilLoop As Integer
Dim slCode As String
Dim ilRet As Integer
    ilHowManyDefined = RptSelCC!lbcSelection(ilListIndex).ListCount

    ilHowMany = RptSelCC!lbcSelection(ilListIndex).SelCount
    If ilHowMany > ilHowManyDefined / 2 Then    'more than half selected
        ilIncludeCodes = False
    Else
        ilIncludeCodes = True
    End If
    For ilLoop = 0 To RptSelCC!lbcSelection(ilListIndex).ListCount - 1 Step 1
        slNameCode = lbcListBox(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If RptSelCC!lbcSelection(ilListIndex).Selected(ilLoop) And ilIncludeCodes Then               'selected ?
            ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
            ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
        Else        'exclude these
            If (Not RptSelCC!lbcSelection(ilListIndex).Selected(ilLoop)) And (Not ilIncludeCodes) Then
                ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
                ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
            End If
        End If
    Next ilLoop
End Sub



'**************************************************************************************
'
'       gCreateProdProv
'       Create a prepass for a list of producers or content providers.
'       Drive thru the vehicle table to determine if the vehicle has a producer or providers.
'       If a producer or producer exists, report that vehicle
'       1-15-04
'
'**************************************************************************************
Sub gCreateProdProv()

    Dim ilRet As Integer    'Return Status
    Dim ilVefLoop As Integer
    Dim ilFound As Integer
    Dim slCode As String
    Dim slNameCode As String
                                            'false = exclude codes store din ilusecode array
                                            'or advt, agy or vehicles codes not to process
    Dim ilListIndex As Integer
    Dim ilVpfIndex As Integer
    Dim ilCommercialProvider As Integer     'true if by Provider and commercial option (vs program)
    Dim ilCkcAll As Integer
    Dim ilSelection As Integer

    Screen.MousePointer = vbHourglass

    hmGrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreateProdProvErr
    gBtrvErrorMsg ilRet, "gCreateProdProv (btrOpen: Grf.Btr)", RptSelCC
    On Error GoTo 0
    imGrfRecLen = Len(tmGrf)

    ilListIndex = RptSelCC!lbcRptType.ListIndex     'which option, producer or provider
    If ilListIndex = LIST_PROVIDER Then
        ilCommercialProvider = True
        tmGrf.sBktType = "C"
        If RptSelCC!rbcSelC4(1).Value = True Then       'test for pgm audio
            ilCommercialProvider = False
            tmGrf.sBktType = "A"        'pgm audio
        End If
    Else
        tmGrf.sBktType = "P"                'producer list
    End If

    ilCkcAll = True
    If RptSel!ckcAll.Value = vbUnchecked Then
        ilCkcAll = False
    End If

    imTerminate = False
    'obtain vehicle and advertiser lists
    ilRet = gObtainVef()
    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    tmGrf.lGenTime = lgNowTime


    For ilVefLoop = LBound(tgMVef) To UBound(tgMVef) - 1
        If tgMVef(ilVefLoop).sState = "A" Then      'must be active vehicle
            ilVpfIndex = gVpfFind(RptSelCC, tgMVef(ilVefLoop).iCode)
            tmVpf = tgVpf(ilVpfIndex)

            ilFound = False
            tmGrf.iSlfCode = 0
            tmGrf.iVefCode = tgMVef(ilVefLoop).iCode
            If ilListIndex = LIST_PRODUCER Then
                If tmVpf.iProducerArfCode > 0 Then       'producer exists for this vehicle
                    If ilCkcAll Then
                        ilFound = True
                    Else
                        'check if this records producer should be included
                        For ilSelection = 0 To RptSelCC!lbcSelection(0).ListCount - 1 Step 1
                            If (RptSelCC!lbcSelection(0).Selected(ilSelection)) Then
                                slNameCode = tgMultiCntrCodeAP(ilSelection).sKey
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                If Val(slCode) = tmVpf.iProducerArfCode Then
                                    ilFound = True
                                    Exit For
                                End If
                            End If
                        Next ilSelection
                    End If
                    tmGrf.iSlfCode = tmVpf.iProducerArfCode
                End If
            Else           'content provider
                If ilCommercialProvider Then        'provider for coml
                    If tmVpf.iCommProvArfCode > 0 Then
                        tmGrf.iSlfCode = tmVpf.iCommProvArfCode
                    End If
                Else        'not coml, program audio provider
                    If tmVpf.iProgProvArfCode > 0 Then
                        tmGrf.iSlfCode = tmVpf.iProgProvArfCode
                    End If
                End If

                If tmGrf.iSlfCode > 0 Then          'theres something that may need to be reported
                    'check if this records provider should be included
                    If ilCkcAll = True Then
                        ilFound = True
                    Else
                        For ilSelection = 0 To RptSelCC!lbcSelection(1).ListCount - 1 Step 1
                            If (RptSelCC!lbcSelection(1).Selected(ilSelection)) Then
                                slNameCode = tgMultiCntrCodeCB(ilSelection).sKey
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                If Val(slCode) = tmGrf.iSlfCode Then
                                    ilFound = True
                                    Exit For
                                End If
                            End If
                        Next ilSelection
                    End If
                End If

            End If
            If ilFound Then
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            End If
        End If
    Next ilVefLoop



    ilRet = btrClose(hmGrf)
    btrDestroy hmGrf

    Exit Sub
gCreateProdProvErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub


