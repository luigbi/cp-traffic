Attribute VB_Name = "modPopSubs"
Option Explicit
Private imLowLimit As Integer
'11/26/17: Add parameters used to determine if populate is required
Type FCTCHGDINFO
    lLastDateChgd As Long
    lLastTimeChgd As Long
End Type
Public tgFctChgdInfo(0 To 4) As FCTCHGDINFO
Const SHTTINDEX = 0
Const VEFINDEX = 1
Const VPFINDEX = 2
Const CPTTINDEX = 3
Const ADFINDEX = 4
Public fct_rst As ADODB.Recordset
'8886
Public Const FILEEXISTS = 0
Public Const FILEEXISTSNOT = 1
Public Function gPopSalesPeopleInfo() As Integer

    'D.S. 06/11/14
    Dim rstSalesPeople As ADODB.Recordset

    gPopSalesPeopleInfo = False
    SQLQuery = "SELECT slfCode, slfFirstName, slfLastName, sofName, mnfName, mnfCode"
    SQLQuery = SQLQuery + " FROM SLF_Salespeople, SOF_Sales_Offices, MNF_Multi_Names"
    SQLQuery = SQLQuery + " WHERE slfSofCode = sofCode And sofMnfSSCode = mnfCode"
    SQLQuery = SQLQuery + " ORDER BY slfCode"
    Set rstSalesPeople = gSQLSelectCall(SQLQuery)
    ReDim tgSalesPeopleInfo(0 To 0) As SALESPPLINFO
    While Not rstSalesPeople.EOF
        tgSalesPeopleInfo(UBound(tgSalesPeopleInfo)).iSlfCode = rstSalesPeople!slfCode
        tgSalesPeopleInfo(UBound(tgSalesPeopleInfo)).sFirstName = Trim$(rstSalesPeople!slfFirstName)
        tgSalesPeopleInfo(UBound(tgSalesPeopleInfo)).sLastName = Trim$(rstSalesPeople!slfLastName)
        tgSalesPeopleInfo(UBound(tgSalesPeopleInfo)).sOffice = Trim$(rstSalesPeople!sofName)
        tgSalesPeopleInfo(UBound(tgSalesPeopleInfo)).sSource = Trim$(rstSalesPeople!mnfName)
        tgSalesPeopleInfo(UBound(tgSalesPeopleInfo)).iSSMnfCode = rstSalesPeople!mnfCode
        ReDim Preserve tgSalesPeopleInfo(0 To UBound(tgSalesPeopleInfo) + 1) As SALESPPLINFO
        rstSalesPeople.MoveNext
    Wend
    rstSalesPeople.Close
    
   gPopSalesPeopleInfo = True
   Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopSalesPeopleInfo"
    gPopSalesPeopleInfo = False
    Exit Function
End Function
    
Public Function gPopMSAMarkets() As Integer
    Dim MSAmkt_rst As ADODB.Recordset
    Dim ilUpper As Integer
    Dim llMax As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(metCode) from met"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsNull(rst(0).Value) Then
        ReDim tgMSAMarketInfo(0 To 0) As MARKETINFO
        gPopMSAMarkets = True
        Exit Function
    End If
    
    llMax = rst(0).Value
    ReDim tgMSAMarketInfo(0 To llMax) As MARKETINFO
    
    SQLQuery = "Select metCode, metName, metRank,  metGroupName from met "
    Set MSAmkt_rst = gSQLSelectCall(SQLQuery)
    ilUpper = 0
    While Not MSAmkt_rst.EOF
        tgMSAMarketInfo(ilUpper).lCode = MSAmkt_rst!metCode
        tgMSAMarketInfo(ilUpper).sName = MSAmkt_rst!metName
        tgMSAMarketInfo(ilUpper).iRank = MSAmkt_rst!metRank
        tgMSAMarketInfo(ilUpper).sGroupName = MSAmkt_rst!metGroupName
        ilUpper = ilUpper + 1
        MSAmkt_rst.MoveNext
    Wend

    ReDim Preserve tgMSAMarketInfo(0 To ilUpper) As MARKETINFO

    'Now sort them by the mktCode
    If UBound(tgMSAMarketInfo) > 1 Then
        ArraySortTyp fnAV(tgMSAMarketInfo(), 0), UBound(tgMSAMarketInfo), 0, LenB(tgMSAMarketInfo(1)), 0, -2, 0
    End If
   
   gPopMSAMarkets = True
   MSAmkt_rst.Close
   rst.Close
   Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopMSAMarkets"
    gPopMSAMarkets = False
    Exit Function
End Function
'6191 xdigital transparency file uses agency info
Public Function gPopAgencies() As Integer

    Dim iUpper As Integer
    Dim llMax As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(agfCode) from AGF_Agencies"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsNull(rst(0).Value) Then
        ReDim tgAgencyInfo(0 To 0)
        gPopAgencies = True
        Exit Function
    End If
    llMax = rst(0).Value
    iUpper = 0
    ReDim tgAgencyInfo(0 To llMax) As AGENCYINFO
    SQLQuery = "SELECT agfName, agfAbbr, agfCode"
    SQLQuery = SQLQuery + " FROM AGF_Agencies"
    SQLQuery = SQLQuery + " ORDER BY agfName"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        tgAgencyInfo(iUpper).iCode = rst!agfCode
        tgAgencyInfo(iUpper).sAgencyName = rst!agfName
        tgAgencyInfo(iUpper).sAgencyAbbr = rst!agfAbbr
        iUpper = iUpper + 1
        rst.MoveNext
    Wend
    
    ReDim Preserve tgAgencyInfo(0 To iUpper) As AGENCYINFO
    
    'Now sort them by the vefCode (Dan:?)
    If UBound(tgAgencyInfo) > 1 Then
        ArraySortTyp fnAV(tgAgencyInfo(), 0), UBound(tgAgencyInfo), 0, LenB(tgAgencyInfo(1)), 0, -1, 0
    End If
    
    rst.Close
    gPopAgencies = True
    Exit Function
ErrHand:
    gMsg = ""
    gHandleError "AffErrorLog.txt", "modPopSubs-gPopAgencies"
    gPopAgencies = False
    Exit Function
End Function
Public Function gPopAdvertisers() As Integer

    Dim iUpper As Integer
    Dim llMax As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(adfCode) from ADF_Advertisers"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsNull(rst(0).Value) Then
        ReDim tgAdvtInfo(0 To 0) As ADVTINFO
        gPopAdvertisers = True
        Exit Function
    End If
    llMax = rst(0).Value
    
    iUpper = 0
    ReDim tgAdvtInfo(0 To llMax) As ADVTINFO
    SQLQuery = "SELECT adfName, adfAbbr, adfCode"
    SQLQuery = SQLQuery + " FROM ADF_Advertisers"
    SQLQuery = SQLQuery + " ORDER BY adfName"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        tgAdvtInfo(iUpper).iCode = rst!adfCode
        tgAdvtInfo(iUpper).sAdvtName = rst!adfName
        tgAdvtInfo(iUpper).sAdvtAbbr = rst!adfAbbr
        iUpper = iUpper + 1
        rst.MoveNext
    Wend
    
    ReDim Preserve tgAdvtInfo(0 To iUpper) As ADVTINFO
    
    'Now sort them by the vefCode
    If UBound(tgAdvtInfo) > 1 Then
        ArraySortTyp fnAV(tgAdvtInfo(), 0), UBound(tgAdvtInfo), 0, LenB(tgAdvtInfo(1)), 0, -1, 0
    End If
    
    rst.Close
    gPopAdvertisers = True
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopAdvertisers"
    gPopAdvertisers = False
    Exit Function
End Function

Public Function gPopVehicles() As Integer
    Dim iUpper As Integer
    Dim iFound As Integer
    Dim iLoop As Integer
    Dim iTest As Integer
    Dim iSet As Integer
    Dim iZone As Integer
    Dim sLetter As String
    Dim iAdd As Integer
    Dim llMax As Long
    Dim rstATT As ADODB.Recordset
    Dim rst As ADODB.Recordset
    Dim ilIdx As Integer
    
    On Error GoTo ErrHand
    
    '11/26/17: Check Changed date/time
    If Not gFileChgd("vef.btr") Then
        gPopVehicles = True
        Exit Function
    End If
    
    SQLQuery = "Select MAX(vefCode) from VEF_Vehicles"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsNull(rst(0).Value) Then
        ReDim tgVehicleInfo(0 To 0) As VEHICLEINFO
        gPopVehicles = True
        Exit Function
    End If
    llMax = rst(0).Value
    
    iUpper = 0
    sLetter = "A"
    ReDim tgVehicleInfo(0 To llMax) As VEHICLEINFO
    'TTP 10244 - 7/8/21 - JW - Affiliate system vehicle lists (fast add, agreement screen, network log, etc.): vehicles with multiple libraries are listed multiple times, instead of only once
    'SQLQuery = "SELECT vefName, vefCodeStn, vefCode, vefType, vefMnfVehGp2, vefOwnerMnfCode, vefmnfVehGp3Mkt, vefmnfVehgp4Fmt,vefmnfVehgp5Rsch, vefmnfVehgp6Sub, vefState, vefVefCode"
    SQLQuery = "SELECT DISTINCT vefName, vefCodeStn, vefCode, vefType, vefMnfVehGp2, vefOwnerMnfCode, vefmnfVehGp3Mkt, vefmnfVehgp4Fmt,vefmnfVehgp5Rsch, vefmnfVehgp6Sub, vefState, vefVefCode"
    SQLQuery = SQLQuery + " FROM VEF_Vehicles Left Outer Join VPF_Vehicle_Options on vefCode = vpfvefKcode"
    SQLQuery = SQLQuery + " Left Outer Join LTF_Lbrary_Title on vefCode = ltfVefCode"
    ''Changed when using the Conventional Vehicles instead of the Log Vehicle
    ''11/20/03
    ''SQLQuery = SQLQuery + " WHERE ((vefvefCode = 0 AND vefType = 'C') OR vefType = 'L' OR vefType = 'A' OR vefType = 'S')"
    'SQLQuery = SQLQuery + " WHERE ((vefType = 'C') OR (vefType = 'A') OR (vefType = 'S') OR (vefType = 'G') OR (vefType = 'I'))"
    '11/4/09-  Show Log and Conventional vehicle.  Let client pick which they want agreements to be used for
    'Temporarily include only for Special user until testing is complete
    'If (Len(sgSpecialPassword) = 4) Then
        'SQLQuery = SQLQuery + " WHERE ((vefType = 'C') OR (vefType = 'A') OR (vefType = 'S') OR (vefType = 'G') OR (vefType = 'I')  OR (vefType = 'L')) and (vpfGMedium <> 'M')"
        SQLQuery = SQLQuery + " WHERE ((vefType = 'C') OR (vefType = 'A') OR (vefType = 'S') OR (vefType = 'G') OR (vefType = 'I')  OR (vefType = 'L')) and (ltfCode > 0)"
    'Else
    '    SQLQuery = SQLQuery + " WHERE ((vefType = 'C') OR (vefType = 'A') OR (vefType = 'S') OR (vefType = 'G') OR (vefType = 'I'))"
    'End If
    SQLQuery = SQLQuery + " ORDER BY vefName"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        If rst!vefType = "S" Then
            SQLQuery = "Select MAX(attVefCode) from att where attVefCode =" & Str$(rst!vefCode)
            Set rstATT = gSQLSelectCall(SQLQuery)
            If rstATT(0).Value = rst!vefCode Then
                iAdd = True
            Else
                iAdd = False
            End If
        Else
            iAdd = True
        End If
        
        If iAdd Then
            tgVehicleInfo(iUpper).sState = rst!vefState
            tgVehicleInfo(iUpper).iCode = rst!vefCode
            tgVehicleInfo(iUpper).sVehType = rst!vefType
            If sgShowByVehType = "Y" Then
                tgVehicleInfo(iUpper).sVehicle = Trim$(tgVehicleInfo(iUpper).sVehType) & ":" & rst!vefName
            Else
                tgVehicleInfo(iUpper).sVehicle = rst!vefName
            End If
            tgVehicleInfo(iUpper).sVehicleName = rst!vefName
            If IsNull(rst!vefCodeStn) Then
                tgVehicleInfo(iUpper).sCodeStn = Left$(rst!vefName, 2) & sLetter
                If sLetter = "Z" Then
                    sLetter = "A"
                Else
                    sLetter = Chr$(Asc(sLetter) + 1)
                End If
            Else
                tgVehicleInfo(iUpper).sCodeStn = rst!vefCodeStn
            End If
            tgVehicleInfo(iUpper).iMnfVehGp2 = rst!vefMnfVehGp2
            tgVehicleInfo(iUpper).iOwnerMnfCode = rst!vefOwnermnfCode       'added all vehicle groups 5-31-16
            tgVehicleInfo(iUpper).iMnfVehGp3Mkt = rst!vefmnfVehGp3Mkt
            tgVehicleInfo(iUpper).iMnfVehGp4Fmt = rst!vefmnfVehGp4Fmt
            tgVehicleInfo(iUpper).iMnfVehGp5Rsch = rst!vefmnfVehGp5Rsch
            tgVehicleInfo(iUpper).iMnfVehGp6Sub = rst!vefmnfVehGp6Sub
           
            tgVehicleInfo(iUpper).iVefCode = rst!vefVefCode
            iUpper = iUpper + 1
            'ReDim Preserve tgVehicleInfo(0 To iUpper) As VEHICLEINFO
        End If
        rst.MoveNext
    Wend
    ReDim Preserve tgVehicleInfo(0 To iUpper) As VEHICLEINFO
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            ilIdx = gBinarySearchVpf(CLng(tgVehicleInfo(iLoop).iCode))
            If ilIdx <> -1 Then
            'If tgVehicleInfo(iLoop).iCode = tgVpfOptions(ilIdx).ivefKCode Then
                tgVehicleInfo(iLoop).iVpfSAGroupNo = tgVpfOptions(ilIdx).iSAGroupNo
                tgVehicleInfo(iLoop).iNoDaysCycle = tgVpfOptions(ilIdx).iLNoDaysCycle
                tgVehicleInfo(iLoop).sPrimaryZone = ""
                For iSet = LBound(tgVehicleInfo(iLoop).sZone) To UBound(tgVehicleInfo(iLoop).sZone) Step 1
                    If iSet = LBound(tgVehicleInfo(iLoop).sZone) Then
                        tgVehicleInfo(iLoop).sZone(iSet) = "   "
                        tgVehicleInfo(iLoop).sFed(iSet) = "*"
                    Else
                        tgVehicleInfo(iLoop).sZone(iSet) = "~~~"
                        tgVehicleInfo(iLoop).sFed(iSet) = ""
                    End If
                    tgVehicleInfo(iLoop).iLocalAdj(iSet) = 0
                    tgVehicleInfo(iLoop).iVehLocalAdj(iSet) = 0
                    tgVehicleInfo(iLoop).iBaseZone(iSet) = -1
                Next iSet
                iZone = 0
                If IsNull(tgVpfOptions(ilIdx).sGZone1) <> True Then
                    If Trim$(tgVpfOptions(ilIdx).sGZone1) <> "" Then
                        tgVehicleInfo(iLoop).sZone(iZone) = tgVpfOptions(ilIdx).sGZone1
                        tgVehicleInfo(iLoop).iLocalAdj(iZone) = tgVpfOptions(ilIdx).iGLocalAdj1
                        tgVehicleInfo(iLoop).iVehLocalAdj(iZone) = tgVpfOptions(ilIdx).iGLocalAdj1
                        If IsNull(tgVpfOptions(ilIdx).sFedZ1) <> True Then
                            tgVehicleInfo(iLoop).sFed(iZone) = tgVpfOptions(ilIdx).sFedZ1
                        Else
                            tgVehicleInfo(iLoop).sFed(iZone) = ""
                        End If
                        If (Trim$(tgVehicleInfo(iLoop).sPrimaryZone) <> "") And (tgVpfOptions(ilIdx).iGLocalAdj1 = 0) Then
                            tgVehicleInfo(iLoop).sPrimaryZone = UCase$(tgVpfOptions(ilIdx).sGZone1)
                        End If
                        iZone = iZone + 1
                    End If
                End If
                If IsNull(tgVpfOptions(ilIdx).sGZone2) <> True Then
                    If Trim$(tgVpfOptions(ilIdx).sGZone2) <> "" Then
                        tgVehicleInfo(iLoop).sZone(iZone) = tgVpfOptions(ilIdx).sGZone2
                        tgVehicleInfo(iLoop).iLocalAdj(iZone) = tgVpfOptions(ilIdx).iGLocalAdj2
                        tgVehicleInfo(iLoop).iVehLocalAdj(iZone) = tgVpfOptions(ilIdx).iGLocalAdj2
                        If IsNull(tgVpfOptions(ilIdx).sFedZ2) <> True Then
                            tgVehicleInfo(iLoop).sFed(iZone) = tgVpfOptions(ilIdx).sFedZ2
                        Else
                            tgVehicleInfo(iLoop).sFed(iZone) = ""
                        End If
                        If (Trim$(tgVehicleInfo(iLoop).sPrimaryZone) <> "") And (tgVpfOptions(ilIdx).iGLocalAdj2 = 0) Then
                            tgVehicleInfo(iLoop).sPrimaryZone = UCase$(tgVpfOptions(ilIdx).sGZone2)
                        End If
                        iZone = iZone + 1
                    End If
                End If
                If IsNull(tgVpfOptions(ilIdx).sGZone3) <> True Then
                    If Trim$(tgVpfOptions(ilIdx).sGZone3) <> "" Then
                        tgVehicleInfo(iLoop).sZone(iZone) = tgVpfOptions(ilIdx).sGZone3
                        tgVehicleInfo(iLoop).iLocalAdj(iZone) = tgVpfOptions(ilIdx).iGLocalAdj3
                        tgVehicleInfo(iLoop).iVehLocalAdj(iZone) = tgVpfOptions(ilIdx).iGLocalAdj3
                        If IsNull(tgVpfOptions(ilIdx).sFedZ3) <> True Then
                            tgVehicleInfo(iLoop).sFed(iZone) = tgVpfOptions(ilIdx).sFedZ3
                        Else
                            tgVehicleInfo(iLoop).sFed(iZone) = ""
                        End If
                        If (Trim$(tgVehicleInfo(iLoop).sPrimaryZone) <> "") And (tgVpfOptions(ilIdx).iGLocalAdj3 = 0) Then
                            tgVehicleInfo(iLoop).sPrimaryZone = UCase$(tgVpfOptions(ilIdx).sGZone3)
                        End If
                        iZone = iZone + 1
                    End If
                End If
                If IsNull(tgVpfOptions(ilIdx).sGZone4) <> True Then
                    If Trim$(tgVpfOptions(ilIdx).sGZone4) <> "" Then
                        tgVehicleInfo(iLoop).sZone(iZone) = tgVpfOptions(ilIdx).sGZone4
                        tgVehicleInfo(iLoop).iLocalAdj(iZone) = tgVpfOptions(ilIdx).iGLocalAdj4
                        tgVehicleInfo(iLoop).iVehLocalAdj(iZone) = tgVpfOptions(ilIdx).iGLocalAdj4
                        If IsNull(tgVpfOptions(ilIdx).sFedZ4) <> True Then
                            tgVehicleInfo(iLoop).sFed(iZone) = tgVpfOptions(ilIdx).sFedZ4
                        Else
                            tgVehicleInfo(iLoop).sFed(iZone) = ""
                        End If
                        If (Trim$(tgVehicleInfo(iLoop).sPrimaryZone) <> "") And (tgVpfOptions(ilIdx).iGLocalAdj3 = 0) Then
                            tgVehicleInfo(iLoop).sPrimaryZone = UCase$(tgVpfOptions(ilIdx).sGZone3)
                        End If
                        iZone = iZone + 1
                    End If
                End If
                tgVehicleInfo(iLoop).iNoZones = iZone
                'Adjust the adjustment
                For iZone = LBound(tgVehicleInfo(iLoop).sZone) To UBound(tgVehicleInfo(iLoop).sZone) Step 1
                    If (Len(Trim$(tgVehicleInfo(iLoop).sZone(iZone))) <> 0) And (Trim$(tgVehicleInfo(iLoop).sZone(iZone)) <> "~~~") And (tgVehicleInfo(iLoop).sFed(iZone) <> "*") Then
                        For iTest = LBound(tgVehicleInfo(iLoop).sZone) To UBound(tgVehicleInfo(iLoop).sZone) Step 1
                            If (Len(Trim$(tgVehicleInfo(iLoop).sZone(iTest))) <> 0) And (Trim$(tgVehicleInfo(iLoop).sZone(iTest)) <> "~~~") And (tgVehicleInfo(iLoop).sFed(iTest) = "*") And (Left$(tgVehicleInfo(iLoop).sZone(iTest), 1) = tgVehicleInfo(iLoop).sFed(iZone)) Then
                                tgVehicleInfo(iLoop).iLocalAdj(iZone) = tgVehicleInfo(iLoop).iLocalAdj(iZone) - tgVehicleInfo(iLoop).iLocalAdj(iTest)
                                tgVehicleInfo(iLoop).iBaseZone(iZone) = iTest
                                Exit For
                            End If
                        Next iTest
                    End If
                Next iZone
                For iZone = LBound(tgVehicleInfo(iLoop).sZone) To UBound(tgVehicleInfo(iLoop).sZone) Step 1
                    If (Len(Trim$(tgVehicleInfo(iLoop).sZone(iZone))) <> 0) And (Trim$(tgVehicleInfo(iLoop).sZone(iZone)) <> "~~~") And (tgVehicleInfo(iLoop).sFed(iZone) = "*") Then
                        tgVehicleInfo(iLoop).iLocalAdj(iZone) = 0
                        tgVehicleInfo(iLoop).iBaseZone(iZone) = iZone
                    End If
                Next iZone
                'tgVehicleInfo(iLoop).iESTEndTime(1) = tgVpfOptions(ilIdx).iESTEndTime1
                'tgVehicleInfo(iLoop).iESTEndTime(2) = tgVpfOptions(ilIdx).iESTEndTime2
                'tgVehicleInfo(iLoop).iESTEndTime(3) = tgVpfOptions(ilIdx).iESTEndTime3
                'tgVehicleInfo(iLoop).iESTEndTime(4) = tgVpfOptions(ilIdx).iESTEndTime4
                'tgVehicleInfo(iLoop).iESTEndTime(5) = tgVpfOptions(ilIdx).iESTEndTime5
                tgVehicleInfo(iLoop).iESTEndTime(0) = tgVpfOptions(ilIdx).iESTEndTime1
                tgVehicleInfo(iLoop).iESTEndTime(1) = tgVpfOptions(ilIdx).iESTEndTime2
                tgVehicleInfo(iLoop).iESTEndTime(2) = tgVpfOptions(ilIdx).iESTEndTime3
                tgVehicleInfo(iLoop).iESTEndTime(3) = tgVpfOptions(ilIdx).iESTEndTime4
                tgVehicleInfo(iLoop).iESTEndTime(4) = tgVpfOptions(ilIdx).iESTEndTime5
                'tgVehicleInfo(iLoop).iMSTEndTime(1) = tgVpfOptions(ilIdx).iMSTEndTime1
                'tgVehicleInfo(iLoop).iMSTEndTime(2) = tgVpfOptions(ilIdx).iESTEndTime2
                'tgVehicleInfo(iLoop).iMSTEndTime(3) = tgVpfOptions(ilIdx).iESTEndTime3
                'tgVehicleInfo(iLoop).iMSTEndTime(4) = tgVpfOptions(ilIdx).iESTEndTime4
                'tgVehicleInfo(iLoop).iMSTEndTime(5) = tgVpfOptions(ilIdx).iESTEndTime5
                tgVehicleInfo(iLoop).iMSTEndTime(0) = tgVpfOptions(ilIdx).iMSTEndTime1
                tgVehicleInfo(iLoop).iMSTEndTime(1) = tgVpfOptions(ilIdx).iESTEndTime2
                tgVehicleInfo(iLoop).iMSTEndTime(2) = tgVpfOptions(ilIdx).iESTEndTime3
                tgVehicleInfo(iLoop).iMSTEndTime(3) = tgVpfOptions(ilIdx).iESTEndTime4
                tgVehicleInfo(iLoop).iMSTEndTime(4) = tgVpfOptions(ilIdx).iESTEndTime5
                'tgVehicleInfo(iLoop).iCSTEndTime(1) = tgVpfOptions(ilIdx).iCSTEndTime1
                'tgVehicleInfo(iLoop).iCSTEndTime(2) = tgVpfOptions(ilIdx).iCSTEndTime2
                'tgVehicleInfo(iLoop).iCSTEndTime(3) = tgVpfOptions(ilIdx).iCSTEndTime3
                'tgVehicleInfo(iLoop).iCSTEndTime(4) = tgVpfOptions(ilIdx).iCSTEndTime4
                'tgVehicleInfo(iLoop).iCSTEndTime(5) = tgVpfOptions(ilIdx).iCSTEndTime5
                tgVehicleInfo(iLoop).iCSTEndTime(0) = tgVpfOptions(ilIdx).iCSTEndTime1
                tgVehicleInfo(iLoop).iCSTEndTime(1) = tgVpfOptions(ilIdx).iCSTEndTime2
                tgVehicleInfo(iLoop).iCSTEndTime(2) = tgVpfOptions(ilIdx).iCSTEndTime3
                tgVehicleInfo(iLoop).iCSTEndTime(3) = tgVpfOptions(ilIdx).iCSTEndTime4
                tgVehicleInfo(iLoop).iCSTEndTime(4) = tgVpfOptions(ilIdx).iCSTEndTime5
                'tgVehicleInfo(iLoop).iPSTEndTime(1) = tgVpfOptions(ilIdx).iPSTEndTime1
                'tgVehicleInfo(iLoop).iPSTEndTime(2) = tgVpfOptions(ilIdx).iPSTEndTime2
                'tgVehicleInfo(iLoop).iPSTEndTime(3) = tgVpfOptions(ilIdx).iPSTEndTime3
                'tgVehicleInfo(iLoop).iPSTEndTime(4) = tgVpfOptions(ilIdx).iPSTEndTime4
                'tgVehicleInfo(iLoop).iPSTEndTime(5) = tgVpfOptions(ilIdx).iPSTEndTime5
                tgVehicleInfo(iLoop).iPSTEndTime(0) = tgVpfOptions(ilIdx).iPSTEndTime1
                tgVehicleInfo(iLoop).iPSTEndTime(1) = tgVpfOptions(ilIdx).iPSTEndTime2
                tgVehicleInfo(iLoop).iPSTEndTime(2) = tgVpfOptions(ilIdx).iPSTEndTime3
                tgVehicleInfo(iLoop).iPSTEndTime(3) = tgVpfOptions(ilIdx).iPSTEndTime4
                tgVehicleInfo(iLoop).iPSTEndTime(4) = tgVpfOptions(ilIdx).iPSTEndTime5
                tgVehicleInfo(iLoop).lHd1CefCode = tgVpfOptions(ilIdx).lLgHd1CefCode
                tgVehicleInfo(iLoop).lLgNmCefCode = tgVpfOptions(ilIdx).lLgNmCefCode
                tgVehicleInfo(iLoop).lFt1CefCode = tgVpfOptions(ilIdx).lLgFt1CefCode
                tgVehicleInfo(iLoop).lFt2CefCode = tgVpfOptions(ilIdx).lLgFt2CefCode
                tgVehicleInfo(iLoop).iProducerArfCode = tgVpfOptions(ilIdx).iProducerArfCode
                tgVehicleInfo(iLoop).iProgProvArfCode = tgVpfOptions(ilIdx).iProgProvArfCode
                tgVehicleInfo(iLoop).iCommProvArfCode = tgVpfOptions(ilIdx).iCommProvArfCode
                tgVehicleInfo(iLoop).sEmbeddedComm = tgVpfOptions(ilIdx).sEmbeddedComm
                tgVehicleInfo(iLoop).iInterfaceID = tgVpfOptions(ilIdx).iInterfaceID
                tgVehicleInfo(iLoop).sWegenerExport = tgVpfOptions(ilIdx).sWegenerExport
                tgVehicleInfo(iLoop).sOLAExport = tgVpfOptions(ilIdx).sOLAExport
            End If
        Next iLoop
    
    'Now sort them by the vefCode
    If UBound(tgVehicleInfo) > 1 Then
        ArraySortTyp fnAV(tgVehicleInfo(), 0), UBound(tgVehicleInfo), 0, LenB(tgVehicleInfo(1)), 0, -1, 0
    End If
    
    gPopVehicles = True
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopVehicles"
    gPopVehicles = False
    Exit Function
End Function

Public Function gPopStations() As Integer
    Dim llUpper As Long
    Dim llMax As Long
    Dim ilCode As Integer
    Dim rst As ADODB.Recordset
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim ilUsedForATTSet As Integer
    Dim llCode As Long
    Dim slStamp As String
    Dim llIndex As Integer 'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness
    On Error GoTo ErrHandStamp

    slStamp = gFileDateTime(sgDBPath & "Shtt.Mkd")
    'ilRet = 0
    'imLowLimit = LBound(tgStationInfo)
    'If ilRet <> 0 Then
    '    sgShttTimeStamp = ""
    'End If
    'On Error GoTo 0
    
    '11/26/17: Check Changed date/time
    If Not gFileChgd("shtt.mkd") Then
        gPopStations = True
        Exit Function
    End If
    
    If PeekArray(tgStationInfo).Ptr <> 0 Then
        imLowLimit = LBound(tgStationInfo)
    Else
        sgShttTimeStamp = ""
        imLowLimit = 0
    End If

    If sgShttTimeStamp <> "" Then
        If StrComp(slStamp, sgShttTimeStamp, 1) = 0 Then
            gPopStations = True
            Exit Function
        End If
    End If
    sgShttTimeStamp = slStamp
    On Error GoTo ErrHand
    bgMarketRepDefinedByStation = False
    bgServiceRepDefinedByStation = False
    ilUsedForATTSet = False
'    SQLQuery = "Select MAX(shttCode) from shtt"
'    Set rst = gSQLSelectCall(SQLQuery)
'    If IsNull(rst(0).Value) Then
'        ReDim tgStationInfo(0 To 0) As STATIONINFO
'        gPopStations = True
'        Exit Function
'    End If
'    llMax = rst(0).Value

    llMax = 20000
    
    ''SQLQuery = "SELECT shttCallLetters, shttMktCode, shttFmtCode, shttOwnerArttCode, shttTimeZone, shttCode, shttType , shttStationID FROM shtt Order by shttCallLetters, shttMarket"
    ''SQLQuery = "SELECT shttCallLetters, shttMktCode, shttFmtCode, shttOwnerArttCode, shttTimeZone, shttCode, shttType , shttStationID FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode Order by shttCallLetters, mktName"
    'SQLQuery = "SELECT shttCallLetters, shttMktCode, shttMetCode, shttMntCode, shttFmtCode, shttOwnerArttCode, shttTimeZone, shttCode, shttType , shttStationID, shttState, shttTztcode, shttUsedForATT, shttUsedForXDigital, shttUsedForWegener, shttUsedForOLA, shttSerialNo1, shttSerialNo2, shttPort, shttPermStationID, shttAckDaylight, shttZip, shttWebAddress FROM shtt Order by UCase(shttCallLetters)"
    SQLQuery = "SELECT * FROM shtt Order by UCase(shttCallLetters)"
    Set rst = gSQLSelectCall(SQLQuery)
    llUpper = 0

    ReDim tgStationInfo(0 To llMax) As STATIONINFO
    
    While Not rst.EOF
        tgStationInfo(llUpper).iCode = rst!shttCode
        tgStationInfo(llUpper).sCallLetters = Trim$(rst!shttCallLetters)
        tgStationInfo(llUpper).lOwnerCode = rst!shttOwnerArttCode
        tgStationInfo(llUpper).iFormatCode = rst!shttFmtCode
        ilCode = rst!shttMktCode
        If ilCode <> 0 Then
            ilRet = gBinarySearchMkt(CLng(ilCode))
            If ilRet <> -1 Then
                tgStationInfo(llUpper).sMarket = tgMarketInfo(ilRet).sName
                tgStationInfo(llUpper).iMktCode = tgMarketInfo(ilRet).lCode
            Else
                tgStationInfo(llUpper).sMarket = ""
                tgStationInfo(llUpper).iMktCode = 0
            End If
        Else
            tgStationInfo(llUpper).sMarket = ""
            tgStationInfo(llUpper).iMktCode = 0
        End If
        tgStationInfo(llUpper).iMSAMktCode = rst!shttMetCode
        tgStationInfo(llUpper).iType = rst!shttType
        tgStationInfo(llUpper).lID = rst!shttStationId
        If IsNull(rst!shttTimeZone) = False Then
            tgStationInfo(llUpper).sZone = rst!shttTimeZone
        Else
            tgStationInfo(llUpper).sZone = ""
        End If
        tgStationInfo(llUpper).iTztCode = rst!shttTztCode
        llCode = rst!shttMntCode
        tgStationInfo(llUpper).lMntCode = llCode
        If llCode <> 0 Then
            ilRet = gBinarySearchMnt(llCode, tgTerritoryInfo())
            If ilRet <> -1 Then
                tgStationInfo(llUpper).sTerritory = tgTerritoryInfo(ilRet).sName
            Else
                tgStationInfo(llUpper).sTerritory = ""
            End If
        Else
            tgStationInfo(llUpper).sTerritory = ""
        End If
        tgStationInfo(llUpper).sPostalName = rst!shttState
        tgStationInfo(llUpper).sUsedForATT = rst!shttUsedForAtt
        If Trim$(tgStationInfo(llUpper).sUsedForATT) <> "" Then
            ilUsedForATTSet = True
        End If
        tgStationInfo(llUpper).lAreaMntCode = rst!shttAreaMntCode
        tgStationInfo(llUpper).sUsedForXDigital = rst!shttUsedForXDigital
        tgStationInfo(llUpper).sUsedForWegener = rst!shttUsedForWegener
        tgStationInfo(llUpper).sUsedForOLA = rst!shttUsedForOLA
        tgStationInfo(llUpper).sUsedForPledgeVsAir = rst!shttPledgeVsAir
        tgStationInfo(llUpper).sSerialNo1 = rst!shttSerialNo1
        tgStationInfo(llUpper).sSerialNo2 = rst!shttSerialNo2
        tgStationInfo(llUpper).sPort = rst!shttPort
        tgStationInfo(llUpper).lPermStationID = rst!shttPermStationID
        tgStationInfo(llUpper).iAckDaylight = rst!shttAckDaylight
        tgStationInfo(llUpper).sZip = rst!shttZip
        tgStationInfo(llUpper).sWebAddress = rst!shttWebAddress
        tgStationInfo(llUpper).sWebPW = rst!shttWebPW
        tgStationInfo(llUpper).sFrequency = rst!shttFrequency
        tgStationInfo(llUpper).lMonikerMntCode = rst!shttMonikerMntCode
        tgStationInfo(llUpper).lMultiCastGroupID = rst!shttMultiCastGroupID
        tgStationInfo(llUpper).lMarketClusterGroupID = rst!shttclustergroupId
        tgStationInfo(llUpper).sAgreementExist = rst!shttAgreementExist
        tgStationInfo(llUpper).sCommentExist = rst!shttCommentExist
        tgStationInfo(llUpper).iMktRepUstCode = rst!shttMktRepUstCode
        If rst!shttMktRepUstCode > 0 Then
            bgMarketRepDefinedByStation = True
        End If
        tgStationInfo(llUpper).iServRepUstCode = rst!shttServRepUstCode
        If rst!shttServRepUstCode > 0 Then
            bgServiceRepDefinedByStation = True
        End If
        tgStationInfo(llUpper).lCityLicMntCode = rst!shttCityLicMntCode
        tgStationInfo(llUpper).lHistStartDate = gDateValue(Format(rst!shttHistStartDate, sgShowDateForm))
        tgStationInfo(llUpper).sStationType = rst!shttStationType
        tgStationInfo(llUpper).lCountyLicMntCode = rst!shttCountyLicMntCode
        tgStationInfo(llUpper).sMailAddress1 = rst!shttAddress1
        tgStationInfo(llUpper).sMailAddress2 = rst!shttAddress2
        tgStationInfo(llUpper).lMailCityMntCode = rst!shttCityMntCode
        tgStationInfo(llUpper).sMailState = rst!shttState
        tgStationInfo(llUpper).sOnAir = rst!shttOnAir
        tgStationInfo(llUpper).lOperatorMntCode = rst!shttOperatorMntCode
        tgStationInfo(llUpper).lAudP12Plus = rst!shttAudP12Plus
        tgStationInfo(llUpper).sPhone = rst!shttPhone
        tgStationInfo(llUpper).sFax = rst!shttFax
        tgStationInfo(llUpper).sPhyAddress1 = rst!shttONAddress1
        tgStationInfo(llUpper).sPhyAddress2 = rst!shttONAddress2
        tgStationInfo(llUpper).lPhyCityMntCode = rst!shttONCityMntCode
        tgStationInfo(llUpper).sPhyState = rst!shttONState
        tgStationInfo(llUpper).sPhyZip = rst!shttOnZip
        tgStationInfo(llUpper).sStateLic = rst!shttStateLic
        tgStationInfo(llUpper).lXDSStationID = rst!shttStationId
        tgStationInfo(llUpper).lWatts = rst!shttWatts
        tgStationInfo(llUpper).lClusterGrougID = rst!shttclustergroupId
        tgStationInfo(llUpper).sMasterCluster = rst!shttMasterCluster
        '8418
        tgStationInfo(llUpper).sWebNumber = rst!shttWebNumber
        
        'TTP 10051 - JW - 6/30/21 - Fast Add grid slowness
        llIndex = gBinarySearchMkt(CLng(tgStationInfo(llUpper).iMktCode))
        If llIndex >= 0 Then tgStationInfo(llUpper).sRank = tgMarketInfo(llIndex).iRank
        
        llIndex = gBinarySearchFmt(CLng(tgStationInfo(llUpper).iFormatCode))
        If llIndex >= 0 Then tgStationInfo(llUpper).sFormat = Trim$(tgFormatInfo(llIndex).sName)
        
        llIndex = gBinarySearchOwner(CLng(tgStationInfo(llUpper).lOwnerCode))
        If llIndex >= 0 Then tgStationInfo(llUpper).sOwner = Trim(tgOwnerInfo(llIndex).sName)
        
        
        llUpper = llUpper + 1
        If llUpper >= llMax Then
            llMax = llMax + 10000
            ReDim Preserve tgStationInfo(0 To llMax) As STATIONINFO
        End If
        rst.MoveNext
    Wend
    ReDim Preserve tgStationInfo(0 To llUpper) As STATIONINFO
    ReDim tgStationInfoByCode(0 To llUpper) As STATIONINFO
    For llLoop = 0 To llUpper Step 1
        If (Not ilUsedForATTSet) And (llLoop < llUpper) Then
            SQLQuery = "UPDATE shtt"
            SQLQuery = SQLQuery & " SET shttUsedForAtt = 'Y'"
            SQLQuery = SQLQuery & " WHERE shttCode = " & tgStationInfo(llLoop).iCode
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand:
                gHandleError "AffErrorLog.txt", "modPopSubs-gPopStations"
                gPopStations = False
                On Error Resume Next
                rst.Close
                Exit Function
            End If
            tgStationInfo(llLoop).sUsedForATT = "Y"
        End If
        tgStationInfoByCode(llLoop) = tgStationInfo(llLoop)
    Next llLoop
    If UBound(tgStationInfoByCode) > 1 Then
        ArraySortTyp fnAV(tgStationInfoByCode(), 0), UBound(tgStationInfoByCode), 0, LenB(tgStationInfoByCode(0)), 0, -1, 0
    End If
    '11/26/17
    gFileChgdUpdate "shtt.mkd", False
    
    gPopStations = True
    rst.Close
    Exit Function
ErrHandStamp:
    ilRet = 1
    Resume Next
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopStations"
    gPopStations = False
    Exit Function
End Function

Public Function gPopReportNames() As Integer
    Dim iUpper As Integer
    Dim sChar  As String * 1
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    iUpper = 0
    ReDim tgRnfInfo(0 To 0) As RNFINFO
    SQLQuery = "SELECT rnfName, rnfRptExe, rnfCode"
    SQLQuery = SQLQuery + " FROM RNF_Report_Name"
    'SQLQuery = SQLQuery & " WHERE (rnfType = 'R') And (rnfName BEGINS WITH 'L' OR rnfName BEGINS WITH 'C')"
    SQLQuery = SQLQuery & " WHERE (rnfType = 'R') And (rnfName Like 'L%' OR rnfName Like 'C%')"
    SQLQuery = SQLQuery + " ORDER BY rnfName"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        sChar = Mid$(rst!rnfName, 2, 1)
        If (sChar >= "0") And (sChar <= "9") Then
            tgRnfInfo(iUpper).iCode = rst!rnfCode
            tgRnfInfo(iUpper).sName = rst!rnfName
            tgRnfInfo(iUpper).sRptExe = rst!rnfRptExe
            iUpper = iUpper + 1
            ReDim Preserve tgRnfInfo(0 To iUpper) As RNFINFO
        End If
        rst.MoveNext
    Wend
   
    gPopReportNames = True
    rst.Close
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopReportNames"
    gPopReportNames = False
    Exit Function
End Function


Public Function gPopSellingVehicles() As Integer
    
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim ilSet As Integer
    Dim ilZone As Integer
    Dim slLetter As String
    Dim rst_PopSelling As ADODB.Recordset
    Dim ilIdx As Integer
    Dim llMax As Long
     
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(vefCode) from VEF_Vehicles WHERE vefType = 'S'"
    Set rst_PopSelling = gSQLSelectCall(SQLQuery)
    If IsNull(rst_PopSelling(0).Value) Then
        ReDim tgSellingVehicleInfo(0 To 0) As SELLINGVEHICLEINFO
        gPopSellingVehicles = True
        Exit Function
    End If
    llMax = rst_PopSelling(0).Value
    
    ilUpper = 0
    slLetter = "A"
    ReDim tgSellingVehicleInfo(0 To llMax) As SELLINGVEHICLEINFO
    SQLQuery = "SELECT vefName, vefCodeStn, vefCode, vefType"
    SQLQuery = SQLQuery + " FROM VEF_Vehicles"
    SQLQuery = SQLQuery + " WHERE vefType = 'S'"
    'SQLQuery = SQLQuery + " ORDER BY vefName"
    Set rst_PopSelling = gSQLSelectCall(SQLQuery)
    While Not rst_PopSelling.EOF
        tgSellingVehicleInfo(ilUpper).iCode = rst_PopSelling!vefCode
        tgSellingVehicleInfo(ilUpper).sVehicle = rst_PopSelling!vefName
        tgSellingVehicleInfo(ilUpper).sVehType = rst_PopSelling!vefType
        If IsNull(rst_PopSelling!vefCodeStn) Then
            tgSellingVehicleInfo(ilUpper).sCodeStn = Left$(rst_PopSelling!vefName, 2) & slLetter
            If slLetter = "Z" Then
                slLetter = "A"
            Else
                slLetter = Chr$(Asc(slLetter) + 1)
            End If
        Else
            tgSellingVehicleInfo(ilUpper).sCodeStn = rst_PopSelling!vefCodeStn
        End If
        ilUpper = ilUpper + 1
        rst_PopSelling.MoveNext
    Wend
    ReDim Preserve tgSellingVehicleInfo(0 To ilUpper) As SELLINGVEHICLEINFO
    'Get Primary zones
    'D.S. 07/19/02
'    SQLQuery = "SELECT vpfvefKCode, VpfSAGroupNo, vpfLNoDaysCycle, "
'    SQLQuery = SQLQuery + "vpfGZone1, vpfGLocalAdj1, vpfFedZ1, "
'    SQLQuery = SQLQuery + "vpfGZone2, vpfGLocalAdj2, vpfFedZ2, "
'    SQLQuery = SQLQuery + "vpfGZone3, vpfGLocalAdj3, vpfFedZ3, "
'    SQLQuery = SQLQuery + "vpfGZone4, vpfGLocalAdj4, vpfFedZ4, "
'    SQLQuery = SQLQuery + "vpfESTEndTime1, vpfESTEndTime2, vpfESTEndTime3, vpfESTEndTime4, vpfESTEndTime5, "
'    SQLQuery = SQLQuery + "vpfMSTEndTime1, vpfMSTEndTime2, vpfMSTEndTime3, vpfMSTEndTime4, vpfMSTEndTime5, "
'    SQLQuery = SQLQuery + "vpfCSTEndTime1, vpfCSTEndTime2, vpfCSTEndTime3, vpfCSTEndTime4, vpfCSTEndTime5, "
'    SQLQuery = SQLQuery + "vpfPSTEndTime1, vpfPSTEndTime2, vpfPSTEndTime3, vpfPSTEndTime4, vpfPSTEndTime5, "
'    SQLQuery = SQLQuery + "vpfLgHd1CefCode, vpfLgNmCefCode, vpfLgFt1CefCode, vpfLgFt2CefCode "
'    SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
'
'    Set rst_PopSelling = gSQLSelectCall(SQLQuery)
'    While Not rst_PopSelling.EOF
        For ilLoop = 0 To UBound(tgSellingVehicleInfo) - 1 Step 1
            'ilIdx = gBinarySearchVpf(tgVehicleInfo(ilLoop).iCode)
            
            ilIdx = gBinarySearchVpf(CLng(tgSellingVehicleInfo(ilLoop).iCode))
            If ilIdx <> -1 Then
            'If tgSellingVehicleInfo(ilLoop).iCode = tgVpfOptions(ilIdx).ivefKCode Then
                tgSellingVehicleInfo(ilLoop).iNoDaysCycle = tgVpfOptions(ilIdx).iLNoDaysCycle
                tgSellingVehicleInfo(ilLoop).iVpfSAGroupNo = tgVpfOptions(ilIdx).iSAGroupNo
                tgSellingVehicleInfo(ilLoop).sPrimaryZone = ""
                For ilSet = LBound(tgSellingVehicleInfo(ilLoop).sZone) To UBound(tgSellingVehicleInfo(ilLoop).sZone) Step 1
                    If ilSet = LBound(tgSellingVehicleInfo(ilLoop).sZone) Then
                        tgSellingVehicleInfo(ilLoop).sZone(ilSet) = "   "
                        tgSellingVehicleInfo(ilLoop).sFed(ilSet) = "*"
                    Else
                        tgSellingVehicleInfo(ilLoop).sZone(ilSet) = "~~~"
                        tgSellingVehicleInfo(ilLoop).sFed(ilSet) = ""
                    End If
                    tgSellingVehicleInfo(ilLoop).iLocalAdj(ilSet) = 0
                    tgSellingVehicleInfo(ilLoop).iVehLocalAdj(ilSet) = 0
                    tgSellingVehicleInfo(ilLoop).iBaseZone(ilSet) = -1
                Next ilSet
                ilZone = 0
                If IsNull(tgVpfOptions(ilIdx).sGZone1) <> True Then
                    If Trim$(tgVpfOptions(ilIdx).sGZone1) <> "" Then
                        tgSellingVehicleInfo(ilLoop).sZone(ilZone) = tgVpfOptions(ilIdx).sGZone1
                        tgSellingVehicleInfo(ilLoop).iLocalAdj(ilZone) = tgVpfOptions(ilIdx).iGLocalAdj1
                        tgSellingVehicleInfo(ilLoop).iVehLocalAdj(ilZone) = tgVpfOptions(ilIdx).iGLocalAdj1
                        If IsNull(tgVpfOptions(ilIdx).sFedZ1) <> True Then
                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = tgVpfOptions(ilIdx).sFedZ1
                        Else
                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = ""
                        End If
                        If (Trim$(tgSellingVehicleInfo(ilLoop).sPrimaryZone) <> "") And (tgVpfOptions(ilIdx).iGLocalAdj1 = 0) Then
                            tgSellingVehicleInfo(ilLoop).sPrimaryZone = UCase$(tgVpfOptions(ilIdx).sGZone1)
                        End If
                        ilZone = ilZone + 1
                    End If
                End If
                If IsNull(tgVpfOptions(ilIdx).sGZone2) <> True Then
                    If Trim$(tgVpfOptions(ilIdx).sGZone2) <> "" Then
                        tgSellingVehicleInfo(ilLoop).sZone(ilZone) = tgVpfOptions(ilIdx).sGZone2
                        tgSellingVehicleInfo(ilLoop).iLocalAdj(ilZone) = tgVpfOptions(ilIdx).iGLocalAdj2
                        tgSellingVehicleInfo(ilLoop).iVehLocalAdj(ilZone) = tgVpfOptions(ilIdx).iGLocalAdj2
                        If IsNull(tgVpfOptions(ilIdx).sFedZ2) <> True Then
                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = tgVpfOptions(ilIdx).sFedZ2
                        Else
                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = ""
                        End If
                        If (Trim$(tgSellingVehicleInfo(ilLoop).sPrimaryZone) <> "") And (tgVpfOptions(ilIdx).iGLocalAdj2 = 0) Then
                            tgSellingVehicleInfo(ilLoop).sPrimaryZone = UCase$(tgVpfOptions(ilIdx).sGZone2)
                        End If
                        ilZone = ilZone + 1
                    End If
                End If
                If IsNull(tgVpfOptions(ilIdx).sGZone3) <> True Then
                    If Trim$(tgVpfOptions(ilIdx).sGZone3) <> "" Then
                        tgSellingVehicleInfo(ilLoop).sZone(ilZone) = tgVpfOptions(ilIdx).sGZone3
                        tgSellingVehicleInfo(ilLoop).iLocalAdj(ilZone) = tgVpfOptions(ilIdx).iGLocalAdj3
                        tgSellingVehicleInfo(ilLoop).iVehLocalAdj(ilZone) = tgVpfOptions(ilIdx).iGLocalAdj3
                        If IsNull(tgVpfOptions(ilIdx).sFedZ3) <> True Then
                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = tgVpfOptions(ilIdx).sFedZ3
                        Else
                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = ""
                        End If
                        If (Trim$(tgSellingVehicleInfo(ilLoop).sPrimaryZone) <> "") And (tgVpfOptions(ilIdx).iGLocalAdj3 = 0) Then
                            tgSellingVehicleInfo(ilLoop).sPrimaryZone = UCase$(tgVpfOptions(ilIdx).sGZone3)
                        End If
                        ilZone = ilZone + 1
                    End If
                End If
                If IsNull(tgVpfOptions(ilIdx).sGZone4) <> True Then
                    If Trim$(tgVpfOptions(ilIdx).sGZone4) <> "" Then
                        tgSellingVehicleInfo(ilLoop).sZone(ilZone) = tgVpfOptions(ilIdx).sGZone4
                        tgSellingVehicleInfo(ilLoop).iLocalAdj(ilZone) = tgVpfOptions(ilIdx).iGLocalAdj4
                        tgSellingVehicleInfo(ilLoop).iVehLocalAdj(ilZone) = tgVpfOptions(ilIdx).iGLocalAdj4
                        If IsNull(tgVpfOptions(ilIdx).sFedZ4) <> True Then
                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = tgVpfOptions(ilIdx).sFedZ4
                        Else
                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = ""
                        End If
                        If (Trim$(tgSellingVehicleInfo(ilLoop).sPrimaryZone) <> "") And (tgVpfOptions(ilIdx).iGLocalAdj3 = 0) Then
                            tgSellingVehicleInfo(ilLoop).sPrimaryZone = UCase$(tgVpfOptions(ilIdx).sGZone3)
                        End If
                        ilZone = ilZone + 1
                    End If
                End If
                tgSellingVehicleInfo(ilLoop).iNoZones = ilZone
                'Adjust the adjustment
                For ilZone = LBound(tgSellingVehicleInfo(ilLoop).sZone) To UBound(tgSellingVehicleInfo(ilLoop).sZone) Step 1
                    If (Len(Trim$(tgSellingVehicleInfo(ilLoop).sZone(ilZone))) <> 0) And (tgSellingVehicleInfo(ilLoop).sFed(ilZone) <> "*") Then
                        For ilTest = LBound(tgSellingVehicleInfo(ilLoop).sZone) To UBound(tgSellingVehicleInfo(ilLoop).sZone) Step 1
                            If (Len(Trim$(tgSellingVehicleInfo(ilLoop).sZone(ilTest))) <> 0) And (tgSellingVehicleInfo(ilLoop).sFed(ilTest) = "*") And (Left$(tgSellingVehicleInfo(ilLoop).sZone(ilTest), 1) = tgSellingVehicleInfo(ilLoop).sFed(ilZone)) Then
                                tgSellingVehicleInfo(ilLoop).iLocalAdj(ilZone) = tgSellingVehicleInfo(ilLoop).iLocalAdj(ilZone) - tgSellingVehicleInfo(ilLoop).iLocalAdj(ilTest)
                                tgSellingVehicleInfo(ilLoop).iBaseZone(ilZone) = ilTest
                                Exit For
                            End If
                        Next ilTest
                    End If
                Next ilZone
                For ilZone = LBound(tgSellingVehicleInfo(ilLoop).sZone) To UBound(tgSellingVehicleInfo(ilLoop).sZone) Step 1
                    If (Len(Trim$(tgSellingVehicleInfo(ilLoop).sZone(ilZone))) <> 0) And (tgSellingVehicleInfo(ilLoop).sFed(ilZone) = "*") Then
                        tgSellingVehicleInfo(ilLoop).iLocalAdj(ilZone) = 0
                        tgSellingVehicleInfo(ilLoop).iBaseZone(ilZone) = ilZone
                    End If
                Next ilZone
'                tgSellingVehicleInfo(ilLoop).iESTEndTime(1) = tgVpfOptions(ilIdx).iESTEndTime1
'                tgSellingVehicleInfo(ilLoop).iESTEndTime(2) = tgVpfOptions(ilIdx).iESTEndTime2
'                tgSellingVehicleInfo(ilLoop).iESTEndTime(3) = tgVpfOptions(ilIdx).iESTEndTime3
'                tgSellingVehicleInfo(ilLoop).iESTEndTime(4) = tgVpfOptions(ilIdx).iESTEndTime4
'                tgSellingVehicleInfo(ilLoop).iESTEndTime(5) = tgVpfOptions(ilIdx).iESTEndTime5
'                tgSellingVehicleInfo(ilLoop).iMSTEndTime(1) = tgVpfOptions(ilIdx).iMSTEndTime1
'                tgSellingVehicleInfo(ilLoop).iMSTEndTime(2) = tgVpfOptions(ilIdx).iMSTEndTime2
'                tgSellingVehicleInfo(ilLoop).iMSTEndTime(3) = tgVpfOptions(ilIdx).iMSTEndTime3
'                tgSellingVehicleInfo(ilLoop).iMSTEndTime(4) = tgVpfOptions(ilIdx).iMSTEndTime4
'                tgSellingVehicleInfo(ilLoop).iMSTEndTime(5) = tgVpfOptions(ilIdx).iMSTEndTime5
'                tgSellingVehicleInfo(ilLoop).iCSTEndTime(1) = tgVpfOptions(ilIdx).iCSTEndTime1
'                tgSellingVehicleInfo(ilLoop).iCSTEndTime(2) = tgVpfOptions(ilIdx).iCSTEndTime2
'                tgSellingVehicleInfo(ilLoop).iCSTEndTime(3) = tgVpfOptions(ilIdx).iCSTEndTime3
'                tgSellingVehicleInfo(ilLoop).iCSTEndTime(4) = tgVpfOptions(ilIdx).iCSTEndTime4
'                tgSellingVehicleInfo(ilLoop).iCSTEndTime(5) = tgVpfOptions(ilIdx).iCSTEndTime5
'                tgSellingVehicleInfo(ilLoop).iPSTEndTime(1) = tgVpfOptions(ilIdx).iPSTEndTime1
'                tgSellingVehicleInfo(ilLoop).iPSTEndTime(2) = tgVpfOptions(ilIdx).iPSTEndTime2
'                tgSellingVehicleInfo(ilLoop).iPSTEndTime(3) = tgVpfOptions(ilIdx).iPSTEndTime3
'                tgSellingVehicleInfo(ilLoop).iPSTEndTime(4) = tgVpfOptions(ilIdx).iPSTEndTime4
'                tgSellingVehicleInfo(ilLoop).iPSTEndTime(5) = tgVpfOptions(ilIdx).iPSTEndTime5
'                tgSellingVehicleInfo(ilLoop).lHd1CefCode = tgVpfOptions(ilIdx).lLgHd1CefCode
'                tgSellingVehicleInfo(ilLoop).lLgNmCefCode = tgVpfOptions(ilIdx).lLgNmCefCode
'                tgSellingVehicleInfo(ilLoop).lFt1CefCode = tgVpfOptions(ilIdx).lLgFt1CefCode
'                tgSellingVehicleInfo(ilLoop).lFt2CefCode = tgVpfOptions(ilIdx).lLgFt2CefCode
                tgSellingVehicleInfo(ilLoop).iESTEndTime(0) = tgVpfOptions(ilIdx).iESTEndTime1
                tgSellingVehicleInfo(ilLoop).iESTEndTime(1) = tgVpfOptions(ilIdx).iESTEndTime2
                tgSellingVehicleInfo(ilLoop).iESTEndTime(2) = tgVpfOptions(ilIdx).iESTEndTime3
                tgSellingVehicleInfo(ilLoop).iESTEndTime(3) = tgVpfOptions(ilIdx).iESTEndTime4
                tgSellingVehicleInfo(ilLoop).iESTEndTime(4) = tgVpfOptions(ilIdx).iESTEndTime5
                tgSellingVehicleInfo(ilLoop).iMSTEndTime(0) = tgVpfOptions(ilIdx).iMSTEndTime1
                tgSellingVehicleInfo(ilLoop).iMSTEndTime(1) = tgVpfOptions(ilIdx).iMSTEndTime2
                tgSellingVehicleInfo(ilLoop).iMSTEndTime(2) = tgVpfOptions(ilIdx).iMSTEndTime3
                tgSellingVehicleInfo(ilLoop).iMSTEndTime(3) = tgVpfOptions(ilIdx).iMSTEndTime4
                tgSellingVehicleInfo(ilLoop).iMSTEndTime(4) = tgVpfOptions(ilIdx).iMSTEndTime5
                tgSellingVehicleInfo(ilLoop).iCSTEndTime(0) = tgVpfOptions(ilIdx).iCSTEndTime1
                tgSellingVehicleInfo(ilLoop).iCSTEndTime(1) = tgVpfOptions(ilIdx).iCSTEndTime2
                tgSellingVehicleInfo(ilLoop).iCSTEndTime(2) = tgVpfOptions(ilIdx).iCSTEndTime3
                tgSellingVehicleInfo(ilLoop).iCSTEndTime(3) = tgVpfOptions(ilIdx).iCSTEndTime4
                tgSellingVehicleInfo(ilLoop).iCSTEndTime(4) = tgVpfOptions(ilIdx).iCSTEndTime5
                tgSellingVehicleInfo(ilLoop).iPSTEndTime(0) = tgVpfOptions(ilIdx).iPSTEndTime1
                tgSellingVehicleInfo(ilLoop).iPSTEndTime(1) = tgVpfOptions(ilIdx).iPSTEndTime2
                tgSellingVehicleInfo(ilLoop).iPSTEndTime(2) = tgVpfOptions(ilIdx).iPSTEndTime3
                tgSellingVehicleInfo(ilLoop).iPSTEndTime(3) = tgVpfOptions(ilIdx).iPSTEndTime4
                tgSellingVehicleInfo(ilLoop).iPSTEndTime(4) = tgVpfOptions(ilIdx).iPSTEndTime5
                tgSellingVehicleInfo(ilLoop).lHd1CefCode = tgVpfOptions(ilIdx).lLgHd1CefCode
                tgSellingVehicleInfo(ilLoop).lLgNmCefCode = tgVpfOptions(ilIdx).lLgNmCefCode
                tgSellingVehicleInfo(ilLoop).lFt1CefCode = tgVpfOptions(ilIdx).lLgFt1CefCode
                tgSellingVehicleInfo(ilLoop).lFt2CefCode = tgVpfOptions(ilIdx).lLgFt2CefCode

            End If
        Next ilLoop

'    Wend
    
    
'    SQLQuery = "SELECT vpfvefKCode, VpfSAGroupNo, vpfLNoDaysCycle, "
'    SQLQuery = SQLQuery + "vpfGZone1, vpfGLocalAdj1, vpfFedZ1, "
'    SQLQuery = SQLQuery + "vpfGZone2, vpfGLocalAdj2, vpfFedZ2, "
'    SQLQuery = SQLQuery + "vpfGZone3, vpfGLocalAdj3, vpfFedZ3, "
'    SQLQuery = SQLQuery + "vpfGZone4, vpfGLocalAdj4, vpfFedZ4, "
'    SQLQuery = SQLQuery + "vpfESTEndTime1, vpfESTEndTime2, vpfESTEndTime3, vpfESTEndTime4, vpfESTEndTime5, "
'    SQLQuery = SQLQuery + "vpfMSTEndTime1, vpfMSTEndTime2, vpfMSTEndTime3, vpfMSTEndTime4, vpfMSTEndTime5, "
'    SQLQuery = SQLQuery + "vpfCSTEndTime1, vpfCSTEndTime2, vpfCSTEndTime3, vpfCSTEndTime4, vpfCSTEndTime5, "
'    SQLQuery = SQLQuery + "vpfPSTEndTime1, vpfPSTEndTime2, vpfPSTEndTime3, vpfPSTEndTime4, vpfPSTEndTime5, "
'    SQLQuery = SQLQuery + "vpfLgHd1CefCode, vpfLgNmCefCode, vpfLgFt1CefCode, vpfLgFt2CefCode "
'    SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
'
'    Set rst_PopSelling = gSQLSelectCall(SQLQuery)
'    While Not rst_PopSelling.EOF
'        For ilLoop = 0 To UBound(tgSellingVehicleInfo) - 1 Step 1
'            If tgSellingVehicleInfo(ilLoop).iCode = rst_PopSelling!vpfvefKCode Then
'                tgSellingVehicleInfo(ilLoop).iNoDaysCycle = rst_PopSelling!vpfLNoDaysCycle
'                tgSellingVehicleInfo(ilLoop).iVpfSAGroupNo = rst_PopSelling!VpfSAGroupNo
'                tgSellingVehicleInfo(ilLoop).sPrimaryZone = ""
'                For i= LBound(tgSellingVehicleInfo(ilLoop).sZone) To UBound(tgSellingVehicleInfo(ilLoop).sZone) Step 1
'                    If i= LBound(tgSellingVehicleInfo(ilLoop).sZone) Then
'                        tgSellingVehicleInfo(ilLoop).sZone(ilSet) = "   "
'                        tgSellingVehicleInfo(ilLoop).sFed(ilSet) = "*"
'                    Else
'                        tgSellingVehicleInfo(ilLoop).sZone(ilSet) = "~~~"
'                        tgSellingVehicleInfo(ilLoop).sFed(ilSet) = ""
'                    End If
'                    tgSellingVehicleInfo(ilLoop).iLocalAdj(ilSet) = 0
'                    tgSellingVehicleInfo(ilLoop).iVehLocalAdj(ilSet) = 0
'                    tgSellingVehicleInfo(ilLoop).iBaseZone(ilSet) = -1
'                Next ilSet
'                ilZone = 0
'                If IsNull(rst_PopSelling!vpfGZone1) <> True Then
'                    If Trim$(rst_PopSelling!vpfGZone1) <> "" Then
'                        tgSellingVehicleInfo(ilLoop).sZone(ilZone) = rst_PopSelling!vpfGZone1
'                        tgSellingVehicleInfo(ilLoop).iLocalAdj(ilZone) = rst_PopSelling!vpfGLocalAdj1
'                        tgSellingVehicleInfo(ilLoop).iVehLocalAdj(ilZone) = rst_PopSelling!vpfGLocalAdj1
'                        If IsNull(rst_PopSelling!vpfFedZ1) <> True Then
'                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = rst_PopSelling!vpfFedZ1
'                        Else
'                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = ""
'                        End If
'                        If (Trim$(tgSellingVehicleInfo(ilLoop).sPrimaryZone) <> "") And (rst_PopSelling!vpfGLocalAdj1 = 0) Then
'                            tgSellingVehicleInfo(ilLoop).sPrimaryZone = UCase$(rst_PopSelling!vpfGZone1)
'                        End If
'                        ilZone = ilZone + 1
'                    End If
'                End If
'                If IsNull(rst_PopSelling!vpfGZone2) <> True Then
'                    If Trim$(rst_PopSelling!vpfGZone2) <> "" Then
'                        tgSellingVehicleInfo(ilLoop).sZone(ilZone) = rst_PopSelling!vpfGZone2
'                        tgSellingVehicleInfo(ilLoop).iLocalAdj(ilZone) = rst_PopSelling!vpfGLocalAdj2
'                        tgSellingVehicleInfo(ilLoop).iVehLocalAdj(ilZone) = rst_PopSelling!vpfGLocalAdj2
'                        If IsNull(rst_PopSelling!vpfFedZ2) <> True Then
'                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = rst_PopSelling!vpfFedZ2
'                        Else
'                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = ""
'                        End If
'                        If (Trim$(tgSellingVehicleInfo(ilLoop).sPrimaryZone) <> "") And (rst_PopSelling!vpfGLocalAdj2 = 0) Then
'                            tgSellingVehicleInfo(ilLoop).sPrimaryZone = UCase$(rst_PopSelling!vpfGZone2)
'                        End If
'                        ilZone = ilZone + 1
'                    End If
'                End If
'                If IsNull(rst_PopSelling!vpfGZone3) <> True Then
'                    If Trim$(rst_PopSelling!vpfGZone3) <> "" Then
'                        tgSellingVehicleInfo(ilLoop).sZone(ilZone) = rst_PopSelling!vpfGZone3
'                        tgSellingVehicleInfo(ilLoop).iLocalAdj(ilZone) = rst_PopSelling!vpfGLocalAdj3
'                        tgSellingVehicleInfo(ilLoop).iVehLocalAdj(ilZone) = rst_PopSelling!vpfGLocalAdj3
'                        If IsNull(rst_PopSelling!vpfFedZ3) <> True Then
'                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = rst_PopSelling!vpfFedZ3
'                        Else
'                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = ""
'                        End If
'                        If (Trim$(tgSellingVehicleInfo(ilLoop).sPrimaryZone) <> "") And (rst_PopSelling!vpfGLocalAdj3 = 0) Then
'                            tgSellingVehicleInfo(ilLoop).sPrimaryZone = UCase$(rst_PopSelling!vpfGZone3)
'                        End If
'                        ilZone = ilZone + 1
'                    End If
'                End If
'                If IsNull(rst_PopSelling!vpfGZone4) <> True Then
'                    If Trim$(rst_PopSelling!vpfGZone4) <> "" Then
'                        tgSellingVehicleInfo(ilLoop).sZone(ilZone) = rst_PopSelling!vpfGZone4
'                        tgSellingVehicleInfo(ilLoop).iLocalAdj(ilZone) = rst_PopSelling!vpfGLocalAdj4
'                        tgSellingVehicleInfo(ilLoop).iVehLocalAdj(ilZone) = rst_PopSelling!vpfGLocalAdj4
'                        If IsNull(rst_PopSelling!vpfFedZ4) <> True Then
'                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = rst_PopSelling!vpfFedZ4
'                        Else
'                            tgSellingVehicleInfo(ilLoop).sFed(ilZone) = ""
'                        End If
'                        If (Trim$(tgSellingVehicleInfo(ilLoop).sPrimaryZone) <> "") And (rst_PopSelling!vpfGLocalAdj3 = 0) Then
'                            tgSellingVehicleInfo(ilLoop).sPrimaryZone = UCase$(rst_PopSelling!vpfGZone3)
'                        End If
'                        ilZone = ilZone + 1
'                    End If
'                End If
'                tgSellingVehicleInfo(ilLoop).iNoZones = ilZone
'                'Adjust the adjustment
'                For ilZone = LBound(tgSellingVehicleInfo(ilLoop).sZone) To UBound(tgSellingVehicleInfo(ilLoop).sZone) Step 1
'                    If (Len(Trim$(tgSellingVehicleInfo(ilLoop).sZone(ilZone))) <> 0) And (tgSellingVehicleInfo(ilLoop).sFed(ilZone) <> "*") Then
'                        For ilTest = LBound(tgSellingVehicleInfo(ilLoop).sZone) To UBound(tgSellingVehicleInfo(ilLoop).sZone) Step 1
'                            If (Len(Trim$(tgSellingVehicleInfo(ilLoop).sZone(ilTest))) <> 0) And (tgSellingVehicleInfo(ilLoop).sFed(ilTest) = "*") And (Left$(tgSellingVehicleInfo(ilLoop).sZone(ilTest), 1) = tgSellingVehicleInfo(ilLoop).sFed(ilZone)) Then
'                                tgSellingVehicleInfo(ilLoop).iLocalAdj(ilZone) = tgSellingVehicleInfo(ilLoop).iLocalAdj(ilZone) - tgSellingVehicleInfo(ilLoop).iLocalAdj(ilTest)
'                                tgSellingVehicleInfo(ilLoop).iBaseZone(ilZone) = ilTest
'                                Exit For
'                            End If
'                        Next ilTest
'                    End If
'                Next ilZone
'                For ilZone = LBound(tgSellingVehicleInfo(ilLoop).sZone) To UBound(tgSellingVehicleInfo(ilLoop).sZone) Step 1
'                    If (Len(Trim$(tgSellingVehicleInfo(ilLoop).sZone(ilZone))) <> 0) And (tgSellingVehicleInfo(ilLoop).sFed(ilZone) = "*") Then
'                        tgSellingVehicleInfo(ilLoop).iLocalAdj(ilZone) = 0
'                        tgSellingVehicleInfo(ilLoop).iBaseZone(ilZone) = ilZone
'                    End If
'                Next ilZone
'                tgSellingVehicleInfo(ilLoop).iESTEndTime(1) = rst_PopSelling!vpfESTEndTime1
'                tgSellingVehicleInfo(ilLoop).iESTEndTime(2) = rst_PopSelling!vpfESTEndTime2
'                tgSellingVehicleInfo(ilLoop).iESTEndTime(3) = rst_PopSelling!vpfESTEndTime3
'                tgSellingVehicleInfo(ilLoop).iESTEndTime(4) = rst_PopSelling!vpfESTEndTime4
'                tgSellingVehicleInfo(ilLoop).iESTEndTime(5) = rst_PopSelling!vpfESTEndTime5
'                tgSellingVehicleInfo(ilLoop).iMSTEndTime(1) = rst_PopSelling!vpfMSTEndTime1
'                tgSellingVehicleInfo(ilLoop).iMSTEndTime(2) = rst_PopSelling!vpfMSTEndTime2
'                tgSellingVehicleInfo(ilLoop).iMSTEndTime(3) = rst_PopSelling!vpfMSTEndTime3
'                tgSellingVehicleInfo(ilLoop).iMSTEndTime(4) = rst_PopSelling!vpfMSTEndTime4
'                tgSellingVehicleInfo(ilLoop).iMSTEndTime(5) = rst_PopSelling!vpfMSTEndTime5
'                tgSellingVehicleInfo(ilLoop).iCSTEndTime(1) = rst_PopSelling!vpfCSTEndTime1
'                tgSellingVehicleInfo(ilLoop).iCSTEndTime(2) = rst_PopSelling!vpfCSTEndTime2
'                tgSellingVehicleInfo(ilLoop).iCSTEndTime(3) = rst_PopSelling!vpfCSTEndTime3
'                tgSellingVehicleInfo(ilLoop).iCSTEndTime(4) = rst_PopSelling!vpfCSTEndTime4
'                tgSellingVehicleInfo(ilLoop).iCSTEndTime(5) = rst_PopSelling!vpfCSTEndTime5
'                tgSellingVehicleInfo(ilLoop).iPSTEndTime(1) = rst_PopSelling!vpfPSTEndTime1
'                tgSellingVehicleInfo(ilLoop).iPSTEndTime(2) = rst_PopSelling!vpfPSTEndTime2
'                tgSellingVehicleInfo(ilLoop).iPSTEndTime(3) = rst_PopSelling!vpfPSTEndTime3
'                tgSellingVehicleInfo(ilLoop).iPSTEndTime(4) = rst_PopSelling!vpfPSTEndTime4
'                tgSellingVehicleInfo(ilLoop).iPSTEndTime(5) = rst_PopSelling!vpfPSTEndTime5
'                tgSellingVehicleInfo(ilLoop).lHd1CefCode = rst_PopSelling!vpfLgHd1CefCode
'                tgSellingVehicleInfo(ilLoop).lLgNmCefCode = rst_PopSelling!vpfLgNmCefCode
'                tgSellingVehicleInfo(ilLoop).lFt1CefCode = rst_PopSelling!vpfLgFt1CefCode
'                tgSellingVehicleInfo(ilLoop).lFt2CefCode = rst_PopSelling!vpfLgFt2CefCode
'                Exit For
'            End If
'        Next ilLoop
'        rst_PopSelling.MoveNext
'    Wend
    
    'rst_PopSelling.Close
    gPopSellingVehicles = True
    
    Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopSellingVehicles"
    gPopSellingVehicles = False
    Exit Function
End Function

Public Function gPopAvailNames() As Integer
    
    Dim iUpper As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    iUpper = 0
    ReDim tgAvailNamesInfo(0 To 0) As AVAILNAMESINFO
    SQLQuery = "SELECT *"
    SQLQuery = SQLQuery + " FROM ANF_AVAIL_NAMES "
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        tgAvailNamesInfo(iUpper).iCode = rst!AnfCode
        tgAvailNamesInfo(iUpper).sName = Trim$(rst!anfName)
        tgAvailNamesInfo(iUpper).sTrafToAff = Trim$(rst!anfTrafToAff)
        tgAvailNamesInfo(iUpper).sISCIExport = Trim$(rst!anfISCIExport)
        tgAvailNamesInfo(iUpper).sAudioExport = Trim$(rst!anfAudioExport)
        tgAvailNamesInfo(iUpper).sAutomationExport = Trim$(rst!anfAutomationExport)
        iUpper = iUpper + 1
        ReDim Preserve tgAvailNamesInfo(0 To iUpper) As AVAILNAMESINFO
        rst.MoveNext
    Wend
    'Now sort them by the anfCode
    If UBound(tgAvailNamesInfo) > 1 Then
        ArraySortTyp fnAV(tgAvailNamesInfo(), 0), UBound(tgAvailNamesInfo), 0, LenB(tgAvailNamesInfo(0)), 0, -1, 0
    End If
    gPopAvailNames = True
    rst.Close
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopAvailNames"
    gPopAvailNames = False
    Exit Function
End Function
'
'           Populate title (contact) names
'
Public Function gPopTitleNames() As Integer

    Dim iUpper As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    iUpper = 0
    ReDim tgTitleInfo(0 To 0) As TITLEINFO
    SQLQuery = "SELECT tntTitle, tntCode"
    SQLQuery = SQLQuery + " FROM TNT"
    SQLQuery = SQLQuery + " ORDER BY tntTitle"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        tgTitleInfo(iUpper).iCode = rst!tntCode
        tgTitleInfo(iUpper).sTitle = rst!tntTitle
        iUpper = iUpper + 1
        ReDim Preserve tgTitleInfo(0 To iUpper) As TITLEINFO
        rst.MoveNext
    Wend
    gPopTitleNames = True
    rst.Close
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopTitleNames"
    gPopTitleNames = False
    Exit Function
End Function

Public Function gBinarySalesPeopleInfo(iCode As Long) As Long
    
    'D.S. 06/11/14
    'Returns the index number of tgSalesPeopleInfo that matches the slfCode that was passed in
    'Note: for this to work tgSalesPeopleInfo was previously be sorted by slfCode
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    On Error GoTo ErrHand
    
    llMin = LBound(tgSalesPeopleInfo)
    llMax = UBound(tgSalesPeopleInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If iCode = tgSalesPeopleInfo(llMiddle).iSlfCode Then
            'found the match
            gBinarySalesPeopleInfo = llMiddle
            Exit Function
        ElseIf iCode < tgSalesPeopleInfo(llMiddle).iSlfCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    
    gBinarySalesPeopleInfo = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySalesPeopleInfo: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySalesPeopleInfo = -1
    Exit Function
    
End Function


Public Function gBinarySearchMkt(llCode As Long) As Long
    
    'D.S. 01/06/06
    'Returns the index number of tgMarketInfo that matches the mktCode that was passed in
    'Note: for this to work tgMarketInfo was previously be sorted by mktCode
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    On Error GoTo ErrHand
    
    llMin = LBound(tgMarketInfo)
    llMax = UBound(tgMarketInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tgMarketInfo(llMiddle).lCode Then
            'found the match
            gBinarySearchMkt = llMiddle
            Exit Function
        ElseIf llCode < tgMarketInfo(llMiddle).lCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    
    gBinarySearchMkt = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchMkt: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
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
    
    On Error GoTo ErrHand
    
    llMin = LBound(tgMSAMarketInfo)
    llMax = UBound(tgMSAMarketInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tgMSAMarketInfo(llMiddle).lCode Then
            'found the match
            gBinarySearchMSAMkt = llMiddle
            Exit Function
        ElseIf llCode < tgMSAMarketInfo(llMiddle).lCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    
    gBinarySearchMSAMkt = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchMSAMkt: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchMSAMkt = -1
    Exit Function
    
End Function
Public Function gBinarySearchStation(sCallLetters As String) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim ilResult As Integer
    Dim slCallLetters As String
    
    On Error GoTo ErrHand
    
    slCallLetters = UCase$(Trim$(sCallLetters))
    gBinarySearchStation = -1    ' Start out as not found.
    llMin = LBound(tgStationInfo)
    llMax = UBound(tgStationInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        'Need to use Binary so that StrComp(KEX-AM,KEXA-FM,...) yields correct results (-1 not 1 which is what vbTextCompare yields)
        ilResult = StrComp(UCase(Trim(tgStationInfo(llMiddle).sCallLetters)), slCallLetters, vbBinaryCompare)
        Select Case ilResult
            Case 0:
                gBinarySearchStation = llMiddle  ' Found it !
                Exit Function
            Case 1:
                llMax = llMiddle - 1
            Case -1:
                llMin = llMiddle + 1
        End Select
    Loop
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchStation: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "BIAImportLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    Exit Function
    
End Function


Public Function gBinarySearchVef(llCode As Long) As Long
    
    'D.S. 01/09/06
    'Returns the index number of tgVehicleInfo that matches the VefCode that was passed in
    'Note: for this to work tgVehicleInfo was previously be sorted by vefKCode
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    On Error GoTo ErrHand
    llMin = LBound(tgVehicleInfo)
    llMax = UBound(tgVehicleInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tgVehicleInfo(llMiddle).iCode Then
            'found the match
            gBinarySearchVef = llMiddle
            Exit Function
        ElseIf llCode < tgVehicleInfo(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchVef = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchVef: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchVef = -1
    Exit Function
    
End Function

Public Function gBinarySearchVpf(llCode As Long) As Long
    
    'D.S. 01/09/06
    'Returns the index number of tgVpfOptions that matches the VefCode that was passed in
    'Note: for this to work tgVpfOptions was previously be sorted by vefKCode
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    On Error GoTo ErrHand
    
    llMin = LBound(tgVpfOptions)
    llMax = UBound(tgVpfOptions) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tgVpfOptions(llMiddle).ivefKCode Then
            'found the match
            gBinarySearchVpf = llMiddle
            Exit Function
        ElseIf llCode < tgVpfOptions(llMiddle).ivefKCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchVpf = -1
    Exit Function
    
ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchVpf: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchVpf = -1
    Exit Function
    
End Function

Public Function gPopMarkets() As Integer

    'D.S. 01/06/06

    Dim mkt_rst As ADODB.Recordset
    Dim ilUpper As Integer
    Dim llMax As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(mktCode) from mkt"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsNull(rst(0).Value) Then
        ReDim tgMarketInfo(0 To 0) As MARKETINFO
        gPopMarkets = True
        Exit Function
    End If
    
    llMax = rst(0).Value
    ReDim tgMarketInfo(0 To llMax) As MARKETINFO
    
    SQLQuery = "Select mktCode, mktName, mktRank, mktBIA, mktARB, mktGroupName from mkt "
    Set mkt_rst = gSQLSelectCall(SQLQuery)
    ilUpper = 0
    While Not mkt_rst.EOF
        tgMarketInfo(ilUpper).lCode = mkt_rst!mktCode
        tgMarketInfo(ilUpper).sName = mkt_rst!mktName
        tgMarketInfo(ilUpper).iRank = mkt_rst!mktRank
        tgMarketInfo(ilUpper).sBIA = mkt_rst!mktBIA
        tgMarketInfo(ilUpper).sARB = mkt_rst!mktARB
        tgMarketInfo(ilUpper).sGroupName = mkt_rst!mktGroupName
        ilUpper = ilUpper + 1
        mkt_rst.MoveNext
    Wend

    ReDim Preserve tgMarketInfo(0 To ilUpper) As MARKETINFO

    'Now sort them by the mktCode
    If UBound(tgMarketInfo) > 1 Then
        ArraySortTyp fnAV(tgMarketInfo(), 0), UBound(tgMarketInfo), 0, LenB(tgMarketInfo(1)), 0, -2, 0
    End If
   
   gPopMarkets = True
   mkt_rst.Close
   rst.Close
   Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopMarkets"
    gPopMarkets = False
    Exit Function
End Function

Public Function gPopVehicleOptions() As Integer

    Dim rst As ADODB.Recordset
    Dim llMax As Long
    Dim ilUpper As Integer
    
    On Error GoTo ErrHand
    
    '11/26/17: Check Changed date/time
    If Not gFileChgd("vpf.btr") Then
        gPopVehicleOptions = True
        Exit Function
    End If
        
    SQLQuery = "Select MAX(vpfVefKCode) from VPF_Vehicle_Options"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsNull(rst(0).Value) Then
        ReDim tgVpfOptions(0 To 0) As VPFOPTIONS
        gPopVehicleOptions = True
        Exit Function
    End If
    
    llMax = rst(0).Value
    
    ReDim tgVpfOptions(0 To llMax) As VPFOPTIONS
    
    SQLQuery = "SELECT vpfvefKCode, VpfSAGroupNo, vpfLNoDaysCycle, "
    SQLQuery = SQLQuery + "vpfGZone1, vpfGLocalAdj1, vpfFedZ1, "
    SQLQuery = SQLQuery + "vpfGZone2, vpfGLocalAdj2, vpfFedZ2, "
    SQLQuery = SQLQuery + "vpfGZone3, vpfGLocalAdj3, vpfFedZ3, "
    SQLQuery = SQLQuery + "vpfGZone4, vpfGLocalAdj4, vpfFedZ4, "
    SQLQuery = SQLQuery + "vpfESTEndTime1, vpfESTEndTime2, vpfESTEndTime3, vpfESTEndTime4, vpfESTEndTime5, "
    SQLQuery = SQLQuery + "vpfMSTEndTime1, vpfMSTEndTime2, vpfMSTEndTime3, vpfMSTEndTime4, vpfMSTEndTime5, "
    SQLQuery = SQLQuery + "vpfCSTEndTime1, vpfCSTEndTime2, vpfCSTEndTime3, vpfCSTEndTime4, vpfCSTEndTime5, "
    SQLQuery = SQLQuery + "vpfPSTEndTime1, vpfPSTEndTime2, vpfPSTEndTime3, vpfPSTEndTime4, vpfPSTEndTime5, "
    SQLQuery = SQLQuery + "vpfLgHd1CefCode, vpfLgNmCefCode, vpfLgFt1CefCode, vpfLgFt2CefCode, "
    SQLQuery = SQLQuery + "vpfProducerArfCode, vpfProgProvArfCode, "
    SQLQuery = SQLQuery + "vpfCommProvArfCode, vpfEmbeddedComm, vpfAvailNameOnWeb, vpfUsingFeatures1, "
    SQLQuery = SQLQuery + "vpfWebLogSummary, vpfWebLogFeedTime, vpfEDASWindow, vpfStnFdXRef, vpfAllowSplitCopy, "
    SQLQuery = SQLQuery + "vpfInterfaceID, vpfWegenerExport, vpfOLAExport, vpfUsingFeatures2, vpfEmbeddedOrROS, vpfLLD"
    SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
    'SQLQuery = SQLQuery + " WHERE vpfGMedium <> 'M'"
    
    Set rst = gSQLSelectCall(SQLQuery)
    
    ilUpper = 0
    While Not rst.EOF
        tgVpfOptions(ilUpper).ivefKCode = rst!vpfvefKCode
        tgVpfOptions(ilUpper).iSAGroupNo = rst!VpfSAGroupNo
        tgVpfOptions(ilUpper).iLNoDaysCycle = rst!vpfLNoDaysCycle
        tgVpfOptions(ilUpper).sGZone1 = rst!vpfGZone1
        tgVpfOptions(ilUpper).iGLocalAdj1 = rst!vpfGLocalAdj1
        tgVpfOptions(ilUpper).sFedZ1 = rst!vpfFedZ1
        tgVpfOptions(ilUpper).sGZone2 = rst!vpfGZone2
        tgVpfOptions(ilUpper).iGLocalAdj2 = rst!vpfGLocalAdj2
        tgVpfOptions(ilUpper).sFedZ2 = rst!vpfFedZ2
        tgVpfOptions(ilUpper).sGZone3 = rst!vpfGZone3
        tgVpfOptions(ilUpper).iGLocalAdj3 = rst!vpfGLocalAdj3
        tgVpfOptions(ilUpper).sFedZ3 = rst!vpfFedZ3
        tgVpfOptions(ilUpper).sGZone4 = rst!vpfGZone4
        tgVpfOptions(ilUpper).iGLocalAdj4 = rst!vpfGLocalAdj4
        tgVpfOptions(ilUpper).sFedZ4 = rst!vpfFedZ4
        tgVpfOptions(ilUpper).iESTEndTime1 = rst!vpfESTEndTime1
        tgVpfOptions(ilUpper).iESTEndTime2 = rst!vpfESTEndTime2
        tgVpfOptions(ilUpper).iESTEndTime3 = rst!vpfESTEndTime3
        tgVpfOptions(ilUpper).iESTEndTime4 = rst!vpfESTEndTime4
        tgVpfOptions(ilUpper).iESTEndTime5 = rst!vpfESTEndTime5
        tgVpfOptions(ilUpper).iMSTEndTime1 = rst!vpfMSTEndTime1
        tgVpfOptions(ilUpper).iMSTEndTime2 = rst!vpfMSTEndTime2
        tgVpfOptions(ilUpper).iMSTEndTime3 = rst!vpfMSTEndTime3
        tgVpfOptions(ilUpper).iMSTEndTime4 = rst!vpfMSTEndTime4
        tgVpfOptions(ilUpper).iMSTEndTime5 = rst!vpfMSTEndTime5
        tgVpfOptions(ilUpper).iCSTEndTime1 = rst!vpfCSTEndTime1
        tgVpfOptions(ilUpper).iCSTEndTime2 = rst!vpfCSTEndTime2
        tgVpfOptions(ilUpper).iCSTEndTime3 = rst!vpfCSTEndTime3
        tgVpfOptions(ilUpper).iCSTEndTime4 = rst!vpfCSTEndTime4
        tgVpfOptions(ilUpper).iCSTEndTime5 = rst!vpfCSTEndTime5
        tgVpfOptions(ilUpper).iPSTEndTime1 = rst!vpfPSTEndTime1
        tgVpfOptions(ilUpper).iPSTEndTime2 = rst!vpfPSTEndTime2
        tgVpfOptions(ilUpper).iPSTEndTime3 = rst!vpfPSTEndTime3
        tgVpfOptions(ilUpper).iPSTEndTime4 = rst!vpfPSTEndTime4
        tgVpfOptions(ilUpper).iPSTEndTime5 = rst!vpfPSTEndTime5
        tgVpfOptions(ilUpper).lLgHd1CefCode = rst!vpfLgHd1CefCode
        tgVpfOptions(ilUpper).lLgNmCefCode = rst!vpfLgNmCefCode
        tgVpfOptions(ilUpper).lLgFt1CefCode = rst!vpfLgFt1CefCode
        tgVpfOptions(ilUpper).lLgFt2CefCode = rst!vpfLgFt2CefCode
        tgVpfOptions(ilUpper).iProducerArfCode = rst!vpfProducerArfCode
        tgVpfOptions(ilUpper).iProgProvArfCode = rst!vpfProgProvArfCode
        tgVpfOptions(ilUpper).iCommProvArfCode = rst!vpfCommProvArfCode
        tgVpfOptions(ilUpper).sEmbeddedComm = rst!vpfEmbeddedComm
        tgVpfOptions(ilUpper).sAvailNameOnWeb = rst!vpfAvailNameOnWeb
        tgVpfOptions(ilUpper).sUsingFeatures1 = rst!vpfUsingFeatures1
        If IsNull(rst!vpfUsingFeatures1) Or (Len(rst!vpfUsingFeatures1) = 0) Then
            tgVpfOptions(ilUpper).sUsingFeatures1 = Chr$(0)
        Else
            tgVpfOptions(ilUpper).sUsingFeatures1 = rst!vpfUsingFeatures1
        End If
        tgVpfOptions(ilUpper).sWebLogFeedTime = rst!vpfWebLogFeedTime
        tgVpfOptions(ilUpper).sWebLogSummary = rst!vpfWebLogSummary
        tgVpfOptions(ilUpper).lEDASWindow = rst!vpfEDASWindow
        tgVpfOptions(ilUpper).sStnFdXRef = rst!vpfStnFdXRef
        tgVpfOptions(ilUpper).sAllowSplitCopy = rst!vpfAllowSplitCopy
        tgVpfOptions(ilUpper).iInterfaceID = rst!vpfInterfaceID
        tgVpfOptions(ilUpper).sWegenerExport = rst!vpfWegenerExport
        tgVpfOptions(ilUpper).sOLAExport = rst!vpfOLAExport
        If IsNull(rst!vpfUsingFeatures2) Or (Len(rst!vpfUsingFeatures2) = 0) Then
            tgVpfOptions(ilUpper).sUsingFeatures2 = Chr$(0)
        Else
            tgVpfOptions(ilUpper).sUsingFeatures2 = rst!vpfUsingFeatures2
        End If
        tgVpfOptions(ilUpper).sEmbeddedOrROS = rst!vpfEmbeddedOrROS
        tgVpfOptions(ilUpper).sLLD = Format(rst!vpfLLD, sgShowDateForm)
        ilUpper = ilUpper + 1
        rst.MoveNext
    Wend

    ReDim Preserve tgVpfOptions(0 To ilUpper) As VPFOPTIONS
    'Now sort them by the vpfCode
    If UBound(tgVpfOptions) > 1 Then
        ArraySortTyp fnAV(tgVpfOptions(), 0), UBound(tgVpfOptions), 0, LenB(tgVpfOptions(1)), 0, -1, 0
    End If
    
    gPopVehicleOptions = True
   Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopVehicleOptions"
    gPopVehicleOptions = False
    Exit Function
End Function

Public Function gBinarySearchLst(llCode As Long) As Long
    
    'D.S. 01/16/06
    'Returns the index number of tgLstInfo that matches the lstCode that was passed in
    'Note: for this to work tglsttInfo was previously sorted
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    llMin = LBound(tgLstInfo)
    llMax = UBound(tgLstInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tgLstInfo(llMiddle).lstCode Then
            'found the match
            gBinarySearchLst = llMiddle
            Exit Function
        ElseIf llCode < tgLstInfo(llMiddle).lstCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchLst = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchLst: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchLst = -1
    Exit Function
    
End Function


Public Function gBinarySearchAtt(llCode As Long) As Long
    
    'D.S. 01/16/06
    'Returns the index number of tgAttInfo1 that matches the lstCode that was passed in
    'Note: for this to work tglsttInfo was previously sorted
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    llMin = LBound(tgAttInfo1)
    llMax = UBound(tgAttInfo1) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tgAttInfo1(llMiddle).attCode Then
            'found the match
            gBinarySearchAtt = llMiddle
            Exit Function
        ElseIf llCode < tgAttInfo1(llMiddle).attCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchAtt = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchAtt: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchAtt = -1
    Exit Function
    
End Function
Public Function gBinarySearchShtt(ilCode As Integer) As Integer
    
    'D.S. 01/16/06
    'Returns the index number of tgShttInfo1 that matches the lstCode that was passed in
    'Note: for this to work tglsttInfo was previously sorted
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    llMin = LBound(tgShttInfo1)
    llMax = UBound(tgShttInfo1) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tgShttInfo1(llMiddle).shttCode Then
            'found the match
            gBinarySearchShtt = llMiddle
            Exit Function
        ElseIf ilCode < tgShttInfo1(llMiddle).shttCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchShtt = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchShtt: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchShtt = -1
    Exit Function
    
End Function

Public Function gBinarySearchStationInfoByCode(ilCode As Integer) As Integer
    
    'D.S. 01/16/06
    'Returns the index number of tgStationInfoByCode that matches the ilCode that was passed in
    
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    
    ilMin = LBound(tgStationInfoByCode)
    ilMax = UBound(tgStationInfoByCode) - 1
    Do While ilMin <= ilMax
        ilMiddle = (CLng(ilMin) + ilMax) \ 2
        If ilCode = tgStationInfoByCode(ilMiddle).iCode Then
            'found the match
            gBinarySearchStationInfoByCode = ilMiddle
            Exit Function
        ElseIf ilCode < tgStationInfoByCode(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gBinarySearchStationInfoByCode = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchStationInfoByCode: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchStationInfoByCode = -1
    Exit Function
    
End Function

Public Function gBinarySearchAdf(llCode As Long) As Long
    
    'D.S. 01/09/06
    'Returns the index number of tgAgencyInfo that matches the adfCode that was passed in
    'Note: for this to work tgadvtInfo was previously be sorted by adfKCode
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    llMin = LBound(tgAdvtInfo)
    llMax = UBound(tgAdvtInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tgAdvtInfo(llMiddle).iCode Then
            'found the match
            gBinarySearchAdf = llMiddle
            Exit Function
        ElseIf llCode < tgAdvtInfo(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchAdf = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchAdf: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchAdf = -1
    Exit Function
    
End Function
Public Function gBinarySearchAgency(llCode As Long) As Long
    
    '6191 dan
    'Returns the index number of tgAgencyInfo that matches the AgencyCode that was passed in
    'Note: for this to work tgAgencyInfo was previously be sorted by AgencyCode
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    llMin = LBound(tgAgencyInfo)
    llMax = UBound(tgAgencyInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tgAgencyInfo(llMiddle).iCode Then
            'found the match
            gBinarySearchAgency = llMiddle
            Exit Function
        ElseIf llCode < tgAgencyInfo(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchAgency = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchAgency: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchAgency = -1
    Exit Function
    
End Function
Public Function gBinarySearchCpf(llCode As Long) As Long
    
    'D.S. 01/09/06
    'Returns the index number of tgCpfInfo that matches the cpfCode that was passed in
    'Note: for this to work tgCpfInfo was previously be sorted by cpfCode
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    On Error GoTo ErrHand
    
    llMin = LBound(tgCpfInfo)
    llMax = UBound(tgCpfInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tgCpfInfo(llMiddle).lCode Then
            'found the match
            gBinarySearchCpf = llMiddle
            Exit Function
        ElseIf llCode < tgCpfInfo(llMiddle).lCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchCpf = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchCpf: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchCpf = -1
    Exit Function
    
End Function

Public Function gPopCpf() As Integer
    
    Dim llUpper As Long
    Dim llMax As Long
    Dim rst As ADODB.Recordset
    Dim ilStop As Integer
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(cpfCode) from CPF_Copy_Prodct_ISCI"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsNull(rst(0).Value) Then
        ReDim tgCpfInfo(0 To 0) As CPFINFO
        gPopCpf = True
        Exit Function
    End If
    llMax = rst(0).Value
    
    llUpper = 0
    ReDim tgCpfInfo(0 To llMax) As CPFINFO
    SQLQuery = "Select "
    SQLQuery = SQLQuery & "cpfCode, "
    SQLQuery = SQLQuery & "cpfName, "
    SQLQuery = SQLQuery & "cpfISCI, "
    SQLQuery = SQLQuery & "cpfCreative "
    SQLQuery = SQLQuery & "From CPF_Copy_Prodct_ISCI"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        If Trim(rst!cpfISCI) = "SRBA 2343" Then
            '
            ilStop = 1
        End If
        tgCpfInfo(llUpper).lCode = rst!cpfCode
        tgCpfInfo(llUpper).sName = rst!cpfName
        tgCpfInfo(llUpper).sISCI = rst!cpfISCI
        tgCpfInfo(llUpper).sCreative = rst!cpfCreative
        llUpper = llUpper + 1
        rst.MoveNext
    Wend
    
    ReDim Preserve tgCpfInfo(0 To llUpper) As CPFINFO
    
    'Now sort them by the vefCode
    If UBound(tgCpfInfo) > 1 Then
        ArraySortTyp fnAV(tgCpfInfo(), 0), UBound(tgCpfInfo), 0, LenB(tgCpfInfo(1)), 0, -2, 0
    End If
    
    rst.Close
    gPopCpf = True
    Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopCpf"
    gPopCpf = False
    Exit Function
End Function
Public Function gPopOwnerNames() As Integer
    
    Dim llUpper As Long
    Dim rst As ADODB.Recordset
        
    On Error GoTo ErrHand
    
    llUpper = 0
    ReDim tgOwnerInfo(0 To 0) As OWNERINFO
    SQLQuery = "SELECT *"
    SQLQuery = SQLQuery + " FROM ARTT"
    SQLQuery = SQLQuery + " where arttType = 'O'"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        tgOwnerInfo(llUpper).lCode = rst!arttCode
        tgOwnerInfo(llUpper).sName = rst!arttLastName
        tgOwnerInfo(llUpper).sPhone = rst!arttPhone
        tgOwnerInfo(llUpper).sFax = rst!arttFax
        tgOwnerInfo(llUpper).sEmail = rst!arttEmail
        llUpper = llUpper + 1
        ReDim Preserve tgOwnerInfo(0 To llUpper) As OWNERINFO
        rst.MoveNext
    Wend
    
    rst.Close
    gPopOwnerNames = True
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopOwnerNames"
    gPopOwnerNames = False
    Exit Function
End Function

Public Function gPopFormats() As Integer
    
    Dim ilUpper As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    gPopFormats = False
    ilUpper = 0
    ReDim tgFormatInfo(0 To 0) As FORMATINFO
    SQLQuery = "SELECT * FROM FMT_Station_Format"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        tgFormatInfo(ilUpper).lCode = rst!FmtCode
        tgFormatInfo(ilUpper).sName = rst!FmtName
        tgFormatInfo(ilUpper).sGroupName = rst!fmtGroupName
        tgFormatInfo(ilUpper).iUstCode = rst!fmtustCode
        ilUpper = ilUpper + 1
        ReDim Preserve tgFormatInfo(0 To ilUpper) As FORMATINFO
        rst.MoveNext
    Wend
    
    'Now sort them by the mktCode
    If UBound(tgFormatInfo) > 1 Then
        ArraySortTyp fnAV(tgFormatInfo(), 0), UBound(tgFormatInfo), 0, LenB(tgFormatInfo(0)), 0, -2, 0
    End If
    
    gPopFormats = True
    rst.Close
    Exit Function
    
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopFormats"
    gPopFormats = False
    Exit Function
End Function
'
'           gPopSpotStatusCodes
'               <input> lbcStatus: list box containing status code selection
'                       ilHideNotCarried - true to deselected Not Carried status
'
Public Sub gPopSpotStatusCodes(lbcStatus As control, ilHideNotCarried As Integer)
Dim lRg As Long
Dim lRet As Long
    lbcStatus.Clear
    lbcStatus.AddItem "1-Aired Live"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 0
    lbcStatus.AddItem "2-Aired Delay Bcast"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 1
    lbcStatus.AddItem "3-Not Aired Tech Diff"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 2
    lbcStatus.AddItem "4-Not Aired Blackout"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 3
    lbcStatus.AddItem "5-Not Aired Other"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 4
    lbcStatus.AddItem "6-Not Aired Product"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 5
    lbcStatus.AddItem "7-Aired Outside Pledge"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 6
    lbcStatus.AddItem "8-Aired Not Pledged"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 7
    lbcStatus.AddItem "9-Not Carried"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 8
    lbcStatus.AddItem "10-Delay Cmml/Prg"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 9
    lbcStatus.AddItem "11-Air Cmml Only"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 10
   
    lRg = CLng(lbcStatus.ListCount - 1) * &H10000 Or 0
    lRet = SendMessageByNum(lbcStatus.hwnd, LB_SELITEMRANGE, True, lRg)
    If ilHideNotCarried Then            '9-18-08 default to deselected Not Carried?
        lbcStatus.Selected(8) = False
    End If

End Sub

Public Sub gCreateUDTForDat(rst As ADODB.Recordset, tlDat As DATRST)
    Dim blDayDefined As Boolean
    Dim ilDay As Integer
    Dim ilPledged As Integer
    
    tlDat.lCode = rst!datCode
    tlDat.lAtfCode = rst!datAtfCode
    tlDat.iShfCode = rst!datShfCode
    tlDat.iVefCode = rst!datVefCode
    'tlDat.iDACode = rst!datDACode
    tlDat.iFdMon = rst!datFdMon
    tlDat.iFdTue = rst!datFdTue
    tlDat.iFdWed = rst!datFdWed
    tlDat.iFdThu = rst!datFdThu
    tlDat.iFdFri = rst!datFdFri
    tlDat.iFdSat = rst!datFdSat
    tlDat.iFdSun = rst!datFdSun
    tlDat.sFdStTime = Format$(rst!datFdStTime, sgShowTimeWSecForm)
    tlDat.sFdEdTime = Format$(rst!datFdEdTime, sgShowTimeWSecForm)
    tlDat.iFdStatus = rst!datFdStatus
    tlDat.iPdMon = rst!datPdMon
    tlDat.iPdTue = rst!datPdTue
    tlDat.iPdWed = rst!datPdWed
    tlDat.iPdThu = rst!datPdThu
    tlDat.iPdFri = rst!datPdFri
    tlDat.iPdSat = rst!datPdSat
    tlDat.iPdSun = rst!datPdSun
    tlDat.sPdDayFed = rst!datPdDayFed
    tlDat.sPdStTime = Format$(rst!datPdStTime, sgShowTimeWSecForm)
    If (Not IsNull(rst!datPdEdTime)) And (Trim$(rst!datPdEdTime) <> "") And (Asc(rst!datPdEdTime) <> 0) Then
        tlDat.sPdEdTime = Format$(rst!datPdEdTime, sgShowTimeWSecForm)
    Else
        tlDat.sPdEdTime = ""
    End If
    tlDat.iAirPlayNo = rst!datAirPlayNo
    tlDat.sEmbeddedOrROS = rst!datEmbeddedOrROS
    tlDat.sUnused = rst!datUnused
    ''6/4/18: handle case where pledge days not defined
    'If (tlDat.iPdMon = 0) And (tlDat.iPdTue = 0) And (tlDat.iPdWed = 0) And (tlDat.iPdThu = 0) And (tlDat.iPdFri = 0) And (tlDat.iPdSat = 0) And (tlDat.iPdSun = 0) Then
    '    If tlDat.iFdStatus = 0 Or tlDat.iFdStatus = 1 Then
    '        tlDat.iPdMon = tlDat.iFdMon
    '        tlDat.iPdTue = tlDat.iFdTue
    '        tlDat.iPdWed = tlDat.iFdWed
    '        tlDat.iPdThu = tlDat.iFdThu
    '        tlDat.iPdFri = tlDat.iFdFri
    '        tlDat.iPdSat = tlDat.iFdSat
    '        tlDat.iPdSun = tlDat.iFdSun
    '    End If
    'End If
    'Match what is in AffCPReturn.bas mGetDat
    ilPledged = tgStatusTypes(tlDat.iFdStatus).iPledged
    If ilPledged = 0 Then
        tlDat.iPdMon = tlDat.iFdMon
        tlDat.iPdTue = tlDat.iFdTue
        tlDat.iPdWed = tlDat.iFdWed
        tlDat.iPdThu = tlDat.iFdThu
        tlDat.iPdFri = tlDat.iFdFri
        tlDat.iPdSat = tlDat.iFdSat
        tlDat.iPdSun = tlDat.iFdSun
    ElseIf ilPledged = 1 Then
        If (tlDat.iPdMon = 0) And (tlDat.iPdTue = 0) And (tlDat.iPdWed = 0) And (tlDat.iPdThu = 0) And (tlDat.iPdFri = 0) And (tlDat.iPdSat = 0) And (tlDat.iPdSun = 0) Then
            blDayDefined = False
        Else
            blDayDefined = True
        End If
        If Not blDayDefined Then
            tlDat.iPdMon = tlDat.iFdMon
            tlDat.iPdTue = tlDat.iFdTue
            tlDat.iPdWed = tlDat.iFdWed
            tlDat.iPdThu = tlDat.iFdThu
            tlDat.iPdFri = tlDat.iFdFri
            tlDat.iPdSat = tlDat.iFdSat
            tlDat.iPdSun = tlDat.iFdSun
        End If
    ElseIf ilPledged = 2 Then
        tlDat.iPdMon = 0
        tlDat.iPdTue = 0
        tlDat.iPdWed = 0
        tlDat.iPdThu = 0
        tlDat.iPdFri = 0
        tlDat.iPdSat = 0
        tlDat.iPdSun = 0
    ElseIf ilPledged = 3 Then
        tlDat.iPdMon = 0
        tlDat.iPdTue = 0
        tlDat.iPdWed = 0
        tlDat.iPdThu = 0
        tlDat.iPdFri = 0
        tlDat.iPdSat = 0
        tlDat.iPdSun = 0
    End If

End Sub

Public Function gPopAuf() As Integer

    Dim iUpper As Integer
    Dim llMax As Long
    Dim slStr As String
    Dim slDate As String
    Dim llDate As Long
    Dim llIdx As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(aufCode) from AUF_Alert_User"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsNull(rst(0).Value) Then
        ReDim sgAufsKey(0 To 0)
        gPopAuf = True
        Exit Function
    End If
    llMax = rst(0).Value
    
    iUpper = 0
    ReDim sgAufsKey(0 To llMax)
    SQLQuery = " Select DISTINCT aufVefCode, aufMoWeekDate FROM AUF_Alert_User"
    SQLQuery = SQLQuery & " Where  aufType = '" & "F" & "'"
    SQLQuery = SQLQuery & " And  aufStatus = '" & "C" & "'"
    SQLQuery = SQLQuery & " And  aufSubType = '" & "S" & "'"
    Set rst = gSQLSelectCall(SQLQuery)
    
    llIdx = 0
    While Not rst.EOF
        'pad out the vefcode with spaces
        slStr = Trim$(Str$(rst!aufVefCode))
        Do While Len(slStr) < 5
             slStr = "0" & slStr
        Loop
        
        'pad out the date with spaces
        
        llDate = DateValue(gAdjYear(rst!aufMoWeekDate))
        slDate = CStr(llDate)
        Do While Len(slDate) < 6
          slDate = "0" & slDate
        Loop

        'concatenate the vefcode and date to form the key value
        sgAufsKey(llIdx) = slStr & slDate
        llIdx = llIdx + 1
        rst.MoveNext
    Wend
    
    ReDim Preserve sgAufsKey(0 To llIdx)
    
'ArraySortTyp avStart, NumEls&, Direction&, ElSize&, MemberOffset&, MemberSize& CaseSensitive&

'Where

'avStart is the ArrayVector of the first element in the sort.
'NumEls& is the number of elements in the sort.
'Direction& is zero for an ascending sort, or nonzero to sort descending.
'ElSize& is the size in bytes of each element in the Type.
'MemberOffset& is the offset into the user defined type of the key member for the sort.
'MemberSize& -1 = Integer, -2 = Long, Positive length of string to sort LenB(...sKey)
'CaseSensitive determines the type of data considered. CaseSensitive& is Zero to ignore, or nonzero to honor capitalization.



'Copyright « 1991-1996 Crescent Software, Inc.
    
    
    
    'Now sort them
    If UBound(sgAufsKey) - 1 > 1 Then
        'ArraySortTyp     avStart,              NumEls&,        Direction&,     ElSize&,         MemberOffset&,      MemberSize&      CaseSensitive&
        'ArraySortTyp fnAV(sgAufsKey(), 0), UBound(sgAufsKey),      0,      LenB(sgAufsKey(0)),        0,         LenB(sgAufsKey(0)),        0
    End If
    
    rst.Close
    gPopAuf = True
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopAuf"
    gPopAuf = False
    Exit Function
End Function

Public Sub gCreateUDTforBOF(rst As ADODB.Recordset, tlBof As BOF)
    tlBof.lCode = rst!bofCode
    tlBof.sType = rst!bofType
    tlBof.iAdfCode = rst!bofadfCode
    tlBof.lsifCode = rst!bofsifCode
    tlBof.iVefCode = rst!bofvefCode
    tlBof.lCifCode = rst!bofcifCode
    tlBof.imnfComp1 = rst!bofmnfComp1
    tlBof.imnfComp2 = rst!bofmnfComp2
    tlBof.sStartDate = Format$(rst!bofStartDate, sgShowDateForm)
    tlBof.sEndDate = Format$(rst!bofEndDate, sgShowDateForm)
    tlBof.sMo = rst!bofMo
    tlBof.sTu = rst!bofTu
    tlBof.sWe = rst!bofWe
    tlBof.sTh = rst!bofTh
    tlBof.sFr = rst!bofFr
    tlBof.sSa = rst!bofSa
    tlBof.sSu = rst!bofSu
    tlBof.sStartTime = Format$(rst!bofStartTime, sgShowTimeWSecForm)
    tlBof.sEndTime = Format$(rst!bofEndTime, sgShowTimeWSecForm)
    tlBof.iurfCode = rst!bofurfCode
    tlBof.lSChfCode = rst!bofSChfCode
    tlBof.iRAdfCode = rst!bofRAdfCode
    tlBof.lRChfCode = rst!bofRChfCode
    tlBof.iLen = rst!bofLen
    tlBof.sSource = rst!bofSource
    tlBof.sUnused = rst!bofUnused
End Sub

Public Sub gCreateUDTforLST(rst As ADODB.Recordset, tlLST As LST)
    tlLST.lCode = rst!lstCode
    tlLST.iType = rst!lstType
    tlLST.lSdfCode = rst!lstSdfCode
    tlLST.lCntrNo = rst!lstCntrNo
    tlLST.iAdfCode = rst!lstAdfCode
    tlLST.iAgfCode = rst!lstAgfCode
    If IsNull(rst!lstProd) Then
        tlLST.sProd = ""
    Else
        If Len(rst!lstProd) > 0 Then
        If Asc(rst!lstProd) <> 0 Then
            tlLST.sProd = rst!lstProd
        Else
            tlLST.sProd = ""
        End If
        Else
            tlLST.sProd = ""
    End If
    End If
    tlLST.iLineNo = rst!lstLineNo
    tlLST.iLnVefCode = rst!lstLnVefCode
    tlLST.sStartDate = Format$(rst!lstStartDate, sgShowDateForm)
    tlLST.sEndDate = Format$(rst!lstEndDate, sgShowDateForm)
    tlLST.iMon = rst!lstMon
    tlLST.iTue = rst!lstTue
    tlLST.iWed = rst!lstWed
    tlLST.iThu = rst!lstThu
    tlLST.iFri = rst!lstFri
    tlLST.iSat = rst!lstSat
    tlLST.iSun = rst!lstSun
    tlLST.iSpotsWk = rst!lstSpotsWk
    tlLST.iPriceType = rst!lstPriceType
    tlLST.lPrice = rst!lstPrice
    tlLST.iSpotType = rst!lstSpotType
    tlLST.iLogVefCode = rst!lstLogVefCode
    tlLST.sLogDate = Format$(rst!lstLogDate, sgShowDateForm)
    tlLST.sLogTime = Format$(rst!lstLogTime, sgShowTimeWSecForm)
    tlLST.sDemo = rst!lstDemo
    tlLST.lAud = rst!lstAud
    'tlLST.sISCI = rst!lstISCI
    
    If IsNull(rst!lstISCI) Then
        tlLST.sISCI = ""
    Else
        If Len(rst!lstISCI) > 0 Then
        If Asc(rst!lstISCI) <> 0 Then
            tlLST.sISCI = rst!lstISCI
        Else
            tlLST.sISCI = ""
        End If
        Else
            tlLST.sISCI = ""
    End If
    End If
    tlLST.iWkNo = rst!lstWkNo
    tlLST.iBreakNo = rst!lstBreakNo
    tlLST.iPositionNo = rst!lstPositionNo
    tlLST.iSeqNo = rst!lstSeqNo
    tlLST.sZone = rst!lstZone
    'tlLST.sCart = rst!lstCart
    
    If IsNull(rst!lstCart) Then
        tlLST.sCart = ""
    Else
        If Len(rst!lstCart) > 0 Then
        If Asc(rst!lstCart) <> 0 Then
            tlLST.sCart = rst!lstCart
        Else
            tlLST.sCart = ""
        End If
        Else
            tlLST.sCart = ""
    End If
    End If
    tlLST.lCpfCode = rst!lstCpfCode
    tlLST.lCrfCsfCode = rst!lstCrfCsfCode
    tlLST.iStatus = rst!lstStatus
    tlLST.iLen = rst!lstLen
    tlLST.iUnits = rst!lstUnits
    tlLST.lCifCode = rst!lstCifCode
    tlLST.iAnfCode = rst!lstAnfCode
    tlLST.lEvtIDCefCode = rst!lstEvtIDCefCode
    tlLST.sSplitNetwork = rst!lstsplitnetwork
    tlLST.lRafCode = rst!lstRafCode
    tlLST.lFsfCode = rst!lstFsfCode
    tlLST.lgsfCode = rst!lstGsfCode
    tlLST.sImportedSpot = rst!lstImportedSpot
    tlLST.lBkoutLstCode = rst!lstBkoutLstCode
    tlLST.sLnStartTime = Format$(rst!lstLnStartTime, sgShowTimeWSecForm)
    tlLST.sLnEndTime = Format$(rst!lstLnEndTime, sgShowTimeWSecForm)
    tlLST.sUnused = rst!lstUnused
End Sub

Public Sub gCreateUDTforLSTPlusDT(rst As ADODB.Recordset, tlLSTPlusDT As LSTPLUSDT)
    tlLSTPlusDT.tLST.lCode = rst!lstCode
    tlLSTPlusDT.tLST.iType = rst!lstType
    tlLSTPlusDT.tLST.lSdfCode = rst!lstSdfCode
    tlLSTPlusDT.tLST.lCntrNo = rst!lstCntrNo
    tlLSTPlusDT.tLST.iAdfCode = rst!lstAdfCode
    tlLSTPlusDT.tLST.iAgfCode = rst!lstAgfCode
    If IsNull(rst!lstProd) Then
        tlLSTPlusDT.tLST.sProd = ""
    Else
        If Len(rst!lstProd) > 0 Then
        If Asc(rst!lstProd) <> 0 Then
            tlLSTPlusDT.tLST.sProd = rst!lstProd
        Else
            tlLSTPlusDT.tLST.sProd = ""
        End If
        Else
            tlLSTPlusDT.tLST.sProd = ""
    End If
    End If
    tlLSTPlusDT.tLST.iLineNo = rst!lstLineNo
    tlLSTPlusDT.tLST.iLnVefCode = rst!lstLnVefCode
    tlLSTPlusDT.tLST.sStartDate = Format$(rst!lstStartDate, sgShowDateForm)
    tlLSTPlusDT.tLST.sEndDate = Format$(rst!lstEndDate, sgShowDateForm)
    tlLSTPlusDT.tLST.iMon = rst!lstMon
    tlLSTPlusDT.tLST.iTue = rst!lstTue
    tlLSTPlusDT.tLST.iWed = rst!lstWed
    tlLSTPlusDT.tLST.iThu = rst!lstThu
    tlLSTPlusDT.tLST.iFri = rst!lstFri
    tlLSTPlusDT.tLST.iSat = rst!lstSat
    tlLSTPlusDT.tLST.iSun = rst!lstSun
    tlLSTPlusDT.tLST.iSpotsWk = rst!lstSpotsWk
    tlLSTPlusDT.tLST.iPriceType = rst!lstPriceType
    tlLSTPlusDT.tLST.lPrice = rst!lstPrice
    tlLSTPlusDT.tLST.iSpotType = rst!lstSpotType
    tlLSTPlusDT.tLST.iLogVefCode = rst!lstLogVefCode
    tlLSTPlusDT.tLST.sLogDate = Format$(rst!lstLogDate, sgShowDateForm)
    tlLSTPlusDT.tLST.sLogTime = Format$(rst!lstLogTime, sgShowTimeWSecForm)
    tlLSTPlusDT.tLST.sDemo = rst!lstDemo
    tlLSTPlusDT.tLST.lAud = rst!lstAud
    'tlLSTPlusDT.tLst.sISCI = rst!lstISCI
    
    If IsNull(rst!lstISCI) Then
        tlLSTPlusDT.tLST.sISCI = ""
    Else
        If Len(rst!lstISCI) > 0 Then
        If Asc(rst!lstISCI) <> 0 Then
            tlLSTPlusDT.tLST.sISCI = rst!lstISCI
        Else
            tlLSTPlusDT.tLST.sISCI = ""
        End If
        Else
            tlLSTPlusDT.tLST.sISCI = ""
    End If
    End If
    tlLSTPlusDT.tLST.iWkNo = rst!lstWkNo
    tlLSTPlusDT.tLST.iBreakNo = rst!lstBreakNo
    tlLSTPlusDT.tLST.iPositionNo = rst!lstPositionNo
    tlLSTPlusDT.tLST.iSeqNo = rst!lstSeqNo
    tlLSTPlusDT.tLST.sZone = rst!lstZone
    'tlLSTPlusDT.tLst.sCart = rst!lstCart
    
    If IsNull(rst!lstCart) Then
        tlLSTPlusDT.tLST.sCart = ""
    Else
        If Len(rst!lstCart) > 0 Then
        If Asc(rst!lstCart) <> 0 Then
            tlLSTPlusDT.tLST.sCart = rst!lstCart
        Else
            tlLSTPlusDT.tLST.sCart = ""
        End If
        Else
            tlLSTPlusDT.tLST.sCart = ""
    End If
    End If
    tlLSTPlusDT.tLST.lCpfCode = rst!lstCpfCode
    tlLSTPlusDT.tLST.lCrfCsfCode = rst!lstCrfCsfCode
    tlLSTPlusDT.tLST.iStatus = rst!lstStatus
    tlLSTPlusDT.tLST.iLen = rst!lstLen
    tlLSTPlusDT.tLST.iUnits = rst!lstUnits
    tlLSTPlusDT.tLST.lCifCode = rst!lstCifCode
    tlLSTPlusDT.tLST.iAnfCode = rst!lstAnfCode
    tlLSTPlusDT.tLST.lEvtIDCefCode = rst!lstEvtIDCefCode
    tlLSTPlusDT.tLST.sSplitNetwork = rst!lstsplitnetwork
    tlLSTPlusDT.tLST.lRafCode = rst!lstRafCode
    tlLSTPlusDT.tLST.lFsfCode = rst!lstFsfCode
    tlLSTPlusDT.tLST.lgsfCode = rst!lstGsfCode
    tlLSTPlusDT.tLST.sImportedSpot = rst!lstImportedSpot
    tlLSTPlusDT.tLST.lBkoutLstCode = rst!lstBkoutLstCode
    tlLSTPlusDT.tLST.sLnStartTime = Format$(rst!lstLnStartTime, sgShowTimeWSecForm)
    tlLSTPlusDT.tLST.sLnEndTime = Format$(rst!lstLnEndTime, sgShowTimeWSecForm)
    tlLSTPlusDT.tLST.sUnused = rst!lstUnused
    tlLSTPlusDT.lDate = gDateValue(tlLSTPlusDT.tLST.sLogDate)
    tlLSTPlusDT.lTime = gTimeToLong(tlLSTPlusDT.tLST.sLogTime, False)
End Sub
Public Function gBinarySearchAuf(sCode As String) As Long
    
    'D.S. 10/25/06
    
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim slCode As String
    Dim llResult As Long
    
    gBinarySearchAuf = -1
    
    On Error GoTo ErrHand
    llMin = LBound(sgAufsKey)
    llMax = UBound(sgAufsKey) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        llResult = StrComp(Trim(sgAufsKey(llMiddle)), sCode, vbTextCompare)
        Select Case llResult
            Case 0:
                gBinarySearchAuf = CLng(llMiddle)  ' Found it !
                Exit Function
            Case 1:
                llMax = llMiddle - 1
            Case -1:
                llMin = llMiddle + 1
        End Select
    Loop

    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchAuf: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchAuf = -1
    Exit Function
    
End Function

Public Function gObtainReplacments() As Integer
    Dim ilNum As Integer
    Dim slNum As String
    Dim slStr As String
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilLast As Integer
    Dim tlBof As BOF
    Dim bof_rst As ADODB.Recordset
    Dim cif_rst As ADODB.Recordset
    Dim slStamp As String

    On Error GoTo ErrHand
    
    slStamp = gFileDateTime(sgDBPath & "BOF.btr")
    If StrComp(slStamp, sgReplacementStamp, 1) = 0 Then
        gObtainReplacments = 1
        Exit Function
    End If
    sgReplacementStamp = slStamp
    
    ReDim tgRBofRec(0 To 0) As BOFREC
    ReDim tgSplitNetLastFill(0 To 0) As SPLITNETLASTFILL
    ilUpper = LBound(tgRBofRec)
    Randomize
    SQLQuery = "SELECT *"
    SQLQuery = SQLQuery + " FROM BOF_Blackout"
    SQLQuery = SQLQuery + " WHERE (bofType = 'R'" & ")"
    Set bof_rst = gSQLSelectCall(SQLQuery)
    While Not bof_rst.EOF
        gCreateUDTforBOF bof_rst, tlBof
        If tlBof.lCifCode > 0 Then
            SQLQuery = "SELECT *"
            SQLQuery = SQLQuery + " FROM CIF_Copy_Inventory"
            SQLQuery = SQLQuery + " WHERE (cifCode =  " & tlBof.lCifCode & ")"
            Set cif_rst = gSQLSelectCall(SQLQuery)
            slStr = Trim$(Str$(cif_rst!cifLen))
            Do While Len(slStr) < 3
                slStr = "0" & slStr
            Loop
            ilNum = Int(10000 * Rnd + 1)
            slNum = Trim$(Str$(ilNum))
            Do While Len(slNum) < 5
                slNum = "0" & slNum
            Loop
            tgRBofRec(ilUpper).sKey = slStr & slNum
            tgRBofRec(ilUpper).tBof = tlBof
            tgRBofRec(ilUpper).iLen = cif_rst!cifLen
            ilUpper = ilUpper + 1
            ReDim Preserve tgRBofRec(0 To ilUpper) As BOFREC
        End If
        bof_rst.MoveNext
    Wend
    ilUpper = UBound(tgRBofRec)
    If ilUpper > 1 Then
        ArraySortTyp fnAV(tgRBofRec(), 0), UBound(tgRBofRec), 0, LenB(tgRBofRec(0)), 0, LenB(tgRBofRec(0).sKey), 0
    End If
    For ilLoop = 0 To UBound(tgRBofRec) - 1 Step 1
        ilFound = False
        For ilLast = 0 To UBound(tgSplitNetLastFill) - 1 Step 1
            If tgSplitNetLastFill(ilLast).iFillLen = tgRBofRec(ilLoop).iLen Then
                ilFound = True
                Exit For
            End If
        Next ilLast
        If Not ilFound Then
            tgSplitNetLastFill(UBound(tgSplitNetLastFill)).iBofIndex = -1
            tgSplitNetLastFill(UBound(tgSplitNetLastFill)).iFillLen = tgRBofRec(ilLoop).iLen
            ReDim Preserve tgSplitNetLastFill(0 To UBound(tgSplitNetLastFill) + 1) As SPLITNETLASTFILL
        End If
    Next ilLoop
    gObtainReplacments = 2
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmPopSubs-gObtainReplacments"
    gObtainReplacments = 0
End Function

Public Function gPopTeams() As Integer

    'D.S. 01/06/06

    Dim team_rst As ADODB.Recordset
    Dim ilUpper As Integer
    Dim llMax As Long

    On Error GoTo ErrHand
    
    ReDim tgTeamInfo(0 To 0) As TEAMINFO
    
    SQLQuery = "Select mnfCode, mnfName, mnfUnitType from " & "MNF_Multi_Names" & " Where mnfType = 'Z'"
    Set team_rst = gSQLSelectCall(SQLQuery)
    ilUpper = 0
    While Not team_rst.EOF
        tgTeamInfo(ilUpper).iCode = team_rst!mnfCode
        tgTeamInfo(ilUpper).sName = team_rst!mnfName
        tgTeamInfo(ilUpper).sShortForm = team_rst!mnfUnitType
        ilUpper = ilUpper + 1
        ReDim Preserve tgTeamInfo(0 To ilUpper) As TEAMINFO
        team_rst.MoveNext
    Wend

    'Now sort them by the mktCode
    If UBound(tgTeamInfo) > 1 Then
        ArraySortTyp fnAV(tgTeamInfo(), 0), UBound(tgTeamInfo), 0, LenB(tgTeamInfo(1)), 0, -1, 0
    End If
   
   team_rst.Close
   gPopTeams = True
   Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopTeams"
    gPopTeams = False
    Exit Function
End Function

Public Function gPopLangs() As Integer

    'D.S. 01/06/06

    Dim Lang_rst As ADODB.Recordset
    Dim ilUpper As Integer
    Dim llMax As Long

    On Error GoTo ErrHand
    
    ReDim tgLangInfo(0 To 0) As LANGINFO
    
    SQLQuery = "Select mnfCode, mnfName, mnfUnitType from " & "MNF_Multi_Names" & " Where mnfType = 'L'"
    Set Lang_rst = gSQLSelectCall(SQLQuery)
    ilUpper = 0
    While Not Lang_rst.EOF
        tgLangInfo(ilUpper).iCode = Lang_rst!mnfCode
        tgLangInfo(ilUpper).sName = Lang_rst!mnfName
        tgLangInfo(ilUpper).sEnglish = Lang_rst!mnfUnitType
        ilUpper = ilUpper + 1
        ReDim Preserve tgLangInfo(0 To ilUpper) As LANGINFO
        Lang_rst.MoveNext
    Wend

    'Now sort them by the mktCode
    If UBound(tgLangInfo) > 1 Then
        ArraySortTyp fnAV(tgLangInfo(), 0), UBound(tgLangInfo), 0, LenB(tgLangInfo(1)), 0, -1, 0
    End If
   
   Lang_rst.Close
   gPopLangs = True
   Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopLangs"
    gPopLangs = False
    Exit Function
End Function

Public Function gPopLstInfo() As Integer

    Dim llUpper As Long
    Dim llMax As Long
    Dim rst_Lst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(lstCode) from LST"
    Set rst_Lst = gSQLSelectCall(SQLQuery)
    If IsNull(rst_Lst(0).Value) Then
        ReDim tgLstInfo(0 To 0) As LSTINFO
        gPopLstInfo = True
        Exit Function
    End If
    llMax = rst_Lst(0).Value
    
    llUpper = 0
    ReDim tgLstInfo(0 To llMax) As LSTINFO
    'SQLQuery = "SELECT lstCode, lstLogDate, lstLogTime, lstLogVefCode"
    SQLQuery = "SELECT *"
    SQLQuery = SQLQuery + " FROM LST"
    Set rst_Lst = gSQLSelectCall(SQLQuery)
    While Not rst_Lst.EOF
        tgLstInfo(llUpper).lstCode = rst_Lst!lstCode
        tgLstInfo(llUpper).lstLogVefCode = rst_Lst!lstLogVefCode
        tgLstInfo(llUpper).lstLogDate = rst_Lst!lstLogDate
        tgLstInfo(llUpper).lstLogTime = rst_Lst!lstLogTime
        llUpper = llUpper + 1
        rst_Lst.MoveNext
    Wend
    
    ReDim Preserve tgLstInfo(0 To llUpper) As LSTINFO
    
    'Now sort them by the lstCode
    If UBound(tgLstInfo) > 1 Then
        ArraySortTyp fnAV(tgLstInfo(), 0), UBound(tgLstInfo), 0, LenB(tgLstInfo(0)), 0, -2, 0
    End If
    
    gPopLstInfo = True
    rst_Lst.Close
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopLstInfo"
    gPopLstInfo = False
    Exit Function
End Function

Public Function gPopCpttInfo() As Integer

    Dim llUpper As Long
    Dim llMax As Long
    Dim rst_Cptt As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select Count(cpttCode) from Cptt"
    Set rst_Cptt = gSQLSelectCall(SQLQuery)
    If IsNull(rst_Cptt(0).Value) Then
        ReDim tgCpttInfo(0 To 0) As CPTTINFO
        gPopCpttInfo = True
        Exit Function
    End If
    llMax = rst_Cptt(0).Value
    
    llUpper = 0
    ReDim tgCpttInfo(0 To llMax) As CPTTINFO
    'SQLQuery = "SELECT lstCode, lstLogDate, lstLogTime, lstLogVefCode"
    SQLQuery = "SELECT cpttPostingStatus, cpttStatus, cpttatfCode, cpttCode, CpttStartDate"
    SQLQuery = SQLQuery + " FROM CPTT"
    Set rst_Cptt = gSQLSelectCall(SQLQuery)
    While Not rst_Cptt.EOF
        tgCpttInfo(llUpper).cpttCode = rst_Cptt!cpttCode
        tgCpttInfo(llUpper).cpttPostingStatus = rst_Cptt!cpttPostingStatus
        tgCpttInfo(llUpper).cpttStatus = rst_Cptt!cpttStatus
        tgCpttInfo(llUpper).cpttatfCode = rst_Cptt!cpttatfCode
        tgCpttInfo(llUpper).CpttStartDate = rst_Cptt!CpttStartDate
        llUpper = llUpper + 1
        rst_Cptt.MoveNext
    Wend
    
    lgCpttCount = llUpper
    ReDim Preserve tgCpttInfo(0 To llUpper) As CPTTINFO
    
    'Now sort them by the lstCode
    If UBound(tgCpttInfo) > 1 Then
        ArraySortTyp fnAV(tgCpttInfo(), 0), UBound(tgCpttInfo), 0, LenB(tgCpttInfo(0)), 0, -2, 0
    End If
    
    gPopCpttInfo = True
    rst_Cptt.Close
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopCpttInfo"
    gPopCpttInfo = False
    Exit Function
End Function
Public Function gPopAttByDate(sDate As String) As Integer

    Dim llUpper As Long
    Dim llMax As Long
    Dim llStaCode As Long
    Dim llVefCode As Long
    Dim iLoop As Integer
    Dim rst_att As ADODB.Recordset
    
    On Error GoTo ErrHand
  
    'Set rst_att = gSQLSelectCall(SQLQuery)

    llUpper = 0
    llMax = 10000
    ReDim tgAttExpMon(0 To llMax) As ATTEXPMON
        
    SQLQuery = "Select attCode, attvefCode, shttCallLetters  "
    SQLQuery = SQLQuery + " FROM att INNER JOIN shtt ON att.attShfCode = shtt.shttCode "
    SQLQuery = SQLQuery + " Where attExportType = 1 AND attOnAir <= '" & sDate & "' AND attOffAir >= '" & sDate & "' AND attDropDate >= '" & sDate & "'"

    Set rst_att = gSQLSelectCall(SQLQuery)
    
    While Not rst_att.EOF
        tgAttExpMon(llUpper).lCode = rst_att!attCode
        tgAttExpMon(llUpper).sCallLetters = Trim$(rst_att!shttCallLetters)
        llStaCode = gBinarySearchStation(Trim$(rst_att!shttCallLetters))
        If llStaCode <> -1 Then
            tgAttExpMon(llUpper).sMarket = Trim$(tgStationInfo(llStaCode).sMarket)
        Else
            tgAttExpMon(llUpper).sMarket = ""
        End If
        llVefCode = gBinarySearchVef(CLng(rst_att!attvefCode))
        If llVefCode <> -1 Then
            tgAttExpMon(llUpper).sVehName = tgVehicleInfo(llVefCode).sVehicle
        End If
        llUpper = llUpper + 1
        If llUpper = llMax Then
            llMax = llMax + 10000
            ReDim Preserve tgAttExpMon(0 To llMax) As ATTEXPMON
        End If
        rst_att.MoveNext
    Wend

    ReDim Preserve tgAttExpMon(0 To llUpper) As ATTEXPMON

    lgAttExpMonCount = llUpper
    
    'Now sort them by the attCode
    If UBound(tgAttExpMon) > 1 Then
        ArraySortTyp fnAV(tgAttExpMon(), 0), UBound(tgAttExpMon), 0, LenB(tgAttExpMon(0)), 0, -2, 0
    End If
    
    gPopAttByDate = True
    rst_att.Close
    Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopAttByDate"
    gPopAttByDate = False
    Exit Function
End Function

Public Function gPopAttInfo() As Integer

    Dim llUpper As Long
    Dim llMax As Long
    Dim rst_att As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(attCode) from ATT"
    Set rst_att = gSQLSelectCall(SQLQuery)
    If IsNull(rst_att(0).Value) Then
        ReDim tgAttInfo1(0 To 0) As ATTINFO1
        gPopAttInfo = True
        Exit Function
    End If
    llMax = rst_att(0).Value
    
    llUpper = 0
    ReDim tgAttInfo1(0 To llMax) As ATTINFO1
    'SQLQuery = "SELECT lstCode, lstLogDate, lstLogTime, lstLogVefCode"
    SQLQuery = "SELECT *"
    SQLQuery = SQLQuery + " FROM ATT"
    Set rst_att = gSQLSelectCall(SQLQuery)
    
    While Not rst_att.EOF
        tgAttInfo1(llUpper).attCode = rst_att!attCode
        tgAttInfo1(llUpper).attvefCode = rst_att!attvefCode
        tgAttInfo1(llUpper).attShttCode = rst_att!attshfcode
        tgAttInfo1(llUpper).attExportType = rst_att!attExportType
        tgAttInfo1(llUpper).attPledgeType = rst_att!attPledgeType
        llUpper = llUpper + 1
        rst_att.MoveNext
    Wend
    
    ReDim Preserve tgAttInfo1(0 To llUpper) As ATTINFO1
    lgAttCount = llUpper
    
    'Now sort them by the lstCode
    If UBound(tgAttInfo1) > 1 Then
        ArraySortTyp fnAV(tgAttInfo1(), 0), UBound(tgAttInfo1), 0, LenB(tgAttInfo1(0)), 0, -2, 0
    End If
    
    gPopAttInfo = True
    rst_att.Close
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubsgPopAttInfo"
    gPopAttInfo = False
    Exit Function
End Function

Public Function gPopShttInfo() As Integer

    Dim llUpper As Long
    Dim llMax As Long
    Dim rst_Shtt As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select Count(shttCode) from SHTT"
    Set rst_Shtt = gSQLSelectCall(SQLQuery)
    If IsNull(rst_Shtt(0).Value) Then
        ReDim tgShttInfo1(0 To 0) As SHTTINFO1
        gPopShttInfo = True
        Exit Function
    End If
    llMax = rst_Shtt(0).Value
    
    llUpper = 0
    ReDim tgShttInfo1(0 To llMax) As SHTTINFO1
    SQLQuery = "SELECT shttCode, shttTimeZone"
    SQLQuery = SQLQuery + " FROM SHTT"
    Set rst_Shtt = gSQLSelectCall(SQLQuery)
    
    While Not rst_Shtt.EOF
        tgShttInfo1(llUpper).shttCode = rst_Shtt!shttCode
        tgShttInfo1(llUpper).shttTimeZone = rst_Shtt!shttTimeZone
        llUpper = llUpper + 1
        rst_Shtt.MoveNext
    Wend
    
    ReDim Preserve tgShttInfo1(0 To llUpper) As SHTTINFO1
    
    'Now sort them by the lstCode
    If UBound(tgShttInfo1) > 1 Then
        ArraySortTyp fnAV(tgShttInfo1(), 0), UBound(tgShttInfo1), 0, LenB(tgShttInfo1(0)), 0, -1, 0
    End If
    
    gPopShttInfo = True
    rst_Shtt.Close
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopShttInfo"
    gPopShttInfo = False
    Exit Function
End Function


Public Function gBinarySearchMnt(llCode As Long, tlMultiNameInfo() As MNTINFO) As Integer
    
    'D.S. 01/06/06
    'Returns the index number of tlMultiNameInfo that matches the mktCode that was passed in
    'Note: for this to work tlMultiNameInfo was previously be sorted by MntCode
    
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    
    On Error GoTo ErrHand
    
    ilMin = LBound(tlMultiNameInfo)
    ilMax = UBound(tlMultiNameInfo) - 1
    Do While ilMin <= ilMax
        ilMiddle = (CLng(ilMin) + ilMax) \ 2
        If llCode = tlMultiNameInfo(ilMiddle).lCode Then
            'found the match
            gBinarySearchMnt = ilMiddle
            Exit Function
        ElseIf llCode < tlMultiNameInfo(ilMiddle).lCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    
    gBinarySearchMnt = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchMnt: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchMnt = -1
    Exit Function
    
End Function

Public Function gPopAffAE() As Integer

    'D.S. 01/06/06

    Dim artt_rst As ADODB.Recordset
    Dim ilUpper As Integer
    Dim llMax As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(arttCode) from artt"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsNull(rst(0).Value) Then
        ReDim tgAffAEInfo(0 To 0) As AFFAEINFO
        gPopAffAE = True
        Exit Function
    End If
    
    llMax = rst(0).Value
    ReDim tgAffAEInfo(0 To llMax) As AFFAEINFO
    
    SQLQuery = "SELECT arttFirstName, arttLastName, arttCode FROM artt Where arttType = 'R'"
    Set artt_rst = gSQLSelectCall(SQLQuery)
    ilUpper = 0
    While Not artt_rst.EOF
        tgAffAEInfo(ilUpper).lCode = artt_rst!arttCode
        tgAffAEInfo(ilUpper).sFirstName = artt_rst!arttFirstName
        tgAffAEInfo(ilUpper).sLastName = artt_rst!arttLastName
        tgAffAEInfo(ilUpper).sName = Trim$(artt_rst!arttFirstName) & " " & Trim$(artt_rst!arttLastName)
        ilUpper = ilUpper + 1
        artt_rst.MoveNext
    Wend

    ReDim Preserve tgAffAEInfo(0 To ilUpper) As AFFAEINFO

    'Now sort them by the arttCode
    If UBound(tgAffAEInfo) > 1 Then
        ArraySortTyp fnAV(tgAffAEInfo(), 0), UBound(tgAffAEInfo), 0, LenB(tgAffAEInfo(1)), 0, -2, 0
    End If
   
   gPopAffAE = True
   artt_rst.Close
   rst.Close
   Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopAffAE"
    gPopAffAE = False
    Exit Function
End Function

Public Function gPopSubTotalGroups() As Integer

    'D.S. 01/06/06

    Dim Subtotal_rst As ADODB.Recordset
    Dim ilUpper As Integer
    Dim llMax As Long

    On Error GoTo ErrHand
    
    ReDim tgSubtotalGroupInfo(0 To 0) As SUBTOTALGROUPINFO
    
    SQLQuery = "Select mnfCode, mnfName from " & "MNF_Multi_Names" & " Where mnfType = 'H' AND mnfUnitType = 2"
    Set Subtotal_rst = gSQLSelectCall(SQLQuery)
    ilUpper = 0
    While Not Subtotal_rst.EOF
        tgSubtotalGroupInfo(ilUpper).iCode = Subtotal_rst!mnfCode
        tgSubtotalGroupInfo(ilUpper).sName = Subtotal_rst!mnfName
        ilUpper = ilUpper + 1
        ReDim Preserve tgSubtotalGroupInfo(0 To ilUpper) As SUBTOTALGROUPINFO
        Subtotal_rst.MoveNext
    Wend

    'Now sort them by the mktCode
    If UBound(tgSubtotalGroupInfo) > 1 Then
        ArraySortTyp fnAV(tgSubtotalGroupInfo(), 0), UBound(tgSubtotalGroupInfo), 0, LenB(tgSubtotalGroupInfo(1)), 0, -1, 0
    End If
   
   Subtotal_rst.Close
   gPopSubTotalGroups = True
   Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopSubtotalGroups"
    gPopSubTotalGroups = False
    Exit Function
End Function

Public Function gPopTimeZones() As Integer

    'D.S. 01/06/06

    Dim tzt_rst As ADODB.Recordset
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim blRemoveExtraNames As Boolean
    Dim ilRemove As Integer
    Dim blAdd As Boolean
    Dim ilIndex As Integer
    '11/26/17
    Dim blShttChgd As Boolean

    On Error GoTo ErrHand
    
    blRemoveExtraNames = False
    '11/26/17
    blShttChgd = False
    ReDim tgTimeZoneInfo(0 To 0) As TIMEZONEINFO
    ReDim tlTimeZoneInfo(0 To 0) As TIMEZONEINFO
    
    Do
        SQLQuery = "Select * from " & "TZT"
        Set tzt_rst = gSQLSelectCall(SQLQuery)
        ilUpper = 0
        ilRemove = 0
        While Not tzt_rst.EOF
            If (Trim$(tzt_rst!tztName) = "Atlantic") Or (Trim$(tzt_rst!tztName) = "Samoa") Or (Trim$(tzt_rst!tztName) = "Palau") Or (Trim$(tzt_rst!tztName) = "Micronesia") Or (Trim$(tzt_rst!tztName) = "Marshall and Wake Islands") Then
                blRemoveExtraNames = True
                tlTimeZoneInfo(ilRemove).iCode = tzt_rst!tztCode
                tlTimeZoneInfo(ilRemove).sName = tzt_rst!tztName
                tlTimeZoneInfo(ilRemove).sCSIName = tzt_rst!tztCSIName
                tlTimeZoneInfo(ilRemove).sGroupName = tzt_rst!tztGroupName
                ilRemove = ilRemove + 1
                ReDim Preserve tlTimeZoneInfo(0 To ilRemove) As TIMEZONEINFO
            Else
                tgTimeZoneInfo(ilUpper).iCode = tzt_rst!tztCode
                tgTimeZoneInfo(ilUpper).sName = tzt_rst!tztName
                tgTimeZoneInfo(ilUpper).sCSIName = tzt_rst!tztCSIName
                tgTimeZoneInfo(ilUpper).sGroupName = tzt_rst!tztGroupName
                ilUpper = ilUpper + 1
                ReDim Preserve tgTimeZoneInfo(0 To ilUpper) As TIMEZONEINFO
            End If
            tzt_rst.MoveNext
        Wend
        If blRemoveExtraNames Then
            For ilLoop = 0 To UBound(tlTimeZoneInfo) - 1 Step 1
                If (Trim$(tlTimeZoneInfo(ilLoop).sName) = "Atlantic") Then
                    For ilIndex = 0 To UBound(tgTimeZoneInfo) - 1 Step 1
                        If (Trim$(tgTimeZoneInfo(ilIndex).sName) = "Eastern") Then
                            SQLQuery = "UPDATE shtt SET shttTztCode = " & tgTimeZoneInfo(ilIndex).iCode & " WHERE shttTztCode = " & tlTimeZoneInfo(ilLoop).iCode
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/13/16: Replaced GoSub
                                'GoSub ErrHand:
                                gHandleError "AffErrorLog.txt", "modPopSubs-gPopTimeZones"
                                gPopTimeZones = False
                                On Error Resume Next
                                tzt_rst.Close
                                Exit Function
                            End If
                            blShttChgd = True
                            SQLQuery = "DELETE FROM tzt WHERE tztCode = " & tlTimeZoneInfo(ilLoop).iCode
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/13/16: Replaced GoSub
                                'GoSub ErrHand:
                                gHandleError "AffErrorLog.txt", "modPopSubs-gPopTimeZones"
                                gPopTimeZones = False
                                On Error Resume Next
                                tzt_rst.Close
                                Exit Function
                            End If
                            Exit For
                        End If
                    Next ilIndex
                End If
                If (Trim$(tlTimeZoneInfo(ilLoop).sName) = "Samoa") Or (Trim$(tlTimeZoneInfo(ilLoop).sName) = "Palau") Or (Trim$(tlTimeZoneInfo(ilLoop).sName) = "Micronesia") Or (Trim$(tlTimeZoneInfo(ilLoop).sName) = "Marshall and Wake Islands") Then
                    For ilIndex = 0 To UBound(tgTimeZoneInfo) - 1 Step 1
                        If (Trim$(tgTimeZoneInfo(ilIndex).sName) = "Pacific") Then
                            SQLQuery = "UPDATE shtt SET shttTztCode = " & tgTimeZoneInfo(ilIndex).iCode & " WHERE shttTztCode = " & tlTimeZoneInfo(ilLoop).iCode
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/13/16: Replaced GoSub
                                'GoSub ErrHand:
                                gHandleError "AffErrorLog.txt", "modPopSubs-gPopTimeZones"
                                gPopTimeZones = False
                                On Error Resume Next
                                tzt_rst.Close
                                Exit Function
                            End If
                            blShttChgd = True
                            SQLQuery = "DELETE FROM tzt WHERE tztCode = " & tlTimeZoneInfo(ilLoop).iCode
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/13/16: Replaced GoSub
                                'GoSub ErrHand:
                                gHandleError "AffErrorLog.txt", "modPopSubs-gPopTimeZones"
                                gPopTimeZones = False
                                On Error Resume Next
                                tzt_rst.Close
                                Exit Function
                            End If
                            Exit For
                        End If
                    Next ilIndex
                End If
            Next ilLoop
            For ilIndex = 0 To UBound(tgTimeZoneInfo) - 1 Step 1
                If (Trim$(tgTimeZoneInfo(ilIndex).sName) = "Alaska") Then
                    tgTimeZoneInfo(ilIndex).sCSIName = "AST"
                    SQLQuery = "UPDATE tzt SET tztCSIName = '" & tgTimeZoneInfo(ilIndex).sCSIName & "' WHERE tztCode = " & tgTimeZoneInfo(ilIndex).iCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub ErrHand:
                        gHandleError "AffErrorLog.txt", "modPopSubs-gPopTimeZones"
                        gPopTimeZones = False
                        On Error Resume Next
                        tzt_rst.Close
                        Exit Function
                    End If
                End If
                If (Trim$(tgTimeZoneInfo(ilIndex).sName) = "Hawaii") Then
                    tgTimeZoneInfo(ilIndex).sCSIName = "HST"
                    SQLQuery = "UPDATE tzt SET tztCSIName = '" & tgTimeZoneInfo(ilIndex).sCSIName & "' WHERE tztCode = " & tgTimeZoneInfo(ilIndex).iCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub ErrHand:
                        gHandleError "AffErrorLog.txt", "modPopSubs-gPopTimeZones"
                        gPopTimeZones = False
                        On Error Resume Next
                        tzt_rst.Close
                        Exit Function
                    End If
                End If
            Next ilIndex
        End If
        If ilUpper = 0 Then
            For ilLoop = 1 To 11 Step 1
                blAdd = True
                SQLQuery = "INSERT INTO tzt (tztName, tztGroupName, tztCSIName, tztDisplaySeqNo, tztUstCode, tztUnused)"
                Select Case ilLoop
                    Case 1
                    '    SQLQuery = SQLQuery & " VALUES ('Atlantic', '', 'EST'," & 0 & "," & igUstCode & ", ''" & ")"
                        blAdd = False
                    Case 2
                        SQLQuery = SQLQuery & " VALUES ('Eastern', '', 'EST'," & 0 & "," & igUstCode & ", ''" & ")"
                    Case 3
                        SQLQuery = SQLQuery & " VALUES ('Central', '', 'CST'," & 0 & "," & igUstCode & ", ''" & ")"
                    Case 4
                        SQLQuery = SQLQuery & " VALUES ('Mountain', '', 'MST'," & 0 & "," & igUstCode & ", ''" & ")"
                    Case 5
                        SQLQuery = SQLQuery & " VALUES ('Pacific', '', 'PST'," & 0 & "," & igUstCode & ", ''" & ")"
                    Case 6
                        SQLQuery = SQLQuery & " VALUES ('Alaska', '', 'AST'," & 0 & "," & igUstCode & ", ''" & ")"
                    Case 7
                        SQLQuery = SQLQuery & " VALUES ('Hawaii', '', 'HST'," & 0 & "," & igUstCode & ", ''" & ")"
                    Case 8
                    '    SQLQuery = SQLQuery & " VALUES ('Samoa', '', 'PST'," & 0 & "," & igUstCode & ", ''" & ")"
                        blAdd = False
                    Case 9
                    '    SQLQuery = SQLQuery & " VALUES ('Palau', '', 'PST'," & 0 & "," & igUstCode & ", ''" & ")"
                        blAdd = False
                    Case 10
                    '    SQLQuery = SQLQuery & " VALUES ('Micronesia', '', 'PST'," & 0 & "," & igUstCode & ", ''" & ")"
                        blAdd = False
                    Case 11
                    '    SQLQuery = SQLQuery & " VALUES ('Marshall and Wake Islands', '', 'PST'," & 0 & "," & igUstCode & ", ''" & ")"
                        blAdd = False
                End Select
                If blAdd Then
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub ErrHand:
                        gHandleError "AffErrorLog.txt", "modPopSubs-gPopTimeZones"
                        gPopTimeZones = False
                        On Error Resume Next
                        tzt_rst.Close
                        Exit Function
                    End If
                End If
            Next ilLoop
            'Set Time zones
            ilRet = gSetStationTimeZones()
            blShttChgd = True
        End If
    Loop While ilUpper = 0

    'Now sort them by the mktCode
    If UBound(tgTimeZoneInfo) > 1 Then
        ArraySortTyp fnAV(tgTimeZoneInfo(), 0), UBound(tgTimeZoneInfo), 0, LenB(tgTimeZoneInfo(0)), 0, -1, 0
    End If
    '11/26/17
    If blShttChgd Then
        gFileChgdUpdate "shtt.mkd", True
    End If
    tzt_rst.Close
    gPopTimeZones = True
    Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopTimeZones"
    tzt_rst.Close
    gPopTimeZones = False
    Exit Function
End Function

Public Function gPopStates() As Integer

    'D.S. 01/06/06

    Dim snt_rst As ADODB.Recordset
    Dim ilUpper As Integer
    Dim ilLoop As Integer

    On Error GoTo ErrHand
    
    ReDim tgStateInfo(0 To 0) As STATEINFO
    
    Do
        SQLQuery = "Select * from " & "snt"
        Set snt_rst = gSQLSelectCall(SQLQuery)
        ilUpper = 0
        While Not snt_rst.EOF
            tgStateInfo(ilUpper).iCode = snt_rst!SntCode
            tgStateInfo(ilUpper).sName = snt_rst!sntName
            tgStateInfo(ilUpper).sPostalName = snt_rst!sntPostalName
            tgStateInfo(ilUpper).sGroupName = snt_rst!sntGroupName
            ilUpper = ilUpper + 1
            ReDim Preserve tgStateInfo(0 To ilUpper) As STATEINFO
            snt_rst.MoveNext
        Wend
        If ilUpper = 0 Then
            For ilLoop = 1 To 72 Step 1
                SQLQuery = "INSERT INTO snt (sntName, sntPostalName, sntGroupName, sntUstCode, sntUnused)"
                Select Case ilLoop
                    Case 1
                        SQLQuery = SQLQuery & " VALUES ('Alabama', 'AL', ''," & igUstCode & ", ''" & ")"
                    Case 2
                        SQLQuery = SQLQuery & " VALUES ('Alaska', 'AK', ''," & igUstCode & ", ''" & ")"
                    Case 3
                        SQLQuery = SQLQuery & " VALUES ('Arizona', 'AZ', ''," & igUstCode & ", ''" & ")"
                    Case 4
                        SQLQuery = SQLQuery & " VALUES ('Arkansas', 'AR', ''," & igUstCode & ", ''" & ")"
                    Case 5
                        SQLQuery = SQLQuery & " VALUES ('California', 'CA', ''," & igUstCode & ", ''" & ")"
                    Case 6
                        SQLQuery = SQLQuery & " VALUES ('Colorado', 'CO', ''," & igUstCode & ", ''" & ")"
                    Case 7
                        SQLQuery = SQLQuery & " VALUES ('Connecticut', 'CT', ''," & igUstCode & ", ''" & ")"
                    Case 8
                        SQLQuery = SQLQuery & " VALUES ('Delaware', 'DE', ''," & igUstCode & ", ''" & ")"
                    Case 9
                        SQLQuery = SQLQuery & " VALUES ('Florida', 'FL', ''," & igUstCode & ", ''" & ")"
                    Case 10
                        SQLQuery = SQLQuery & " VALUES ('Georgia', 'GA', ''," & igUstCode & ", ''" & ")"
                    Case 11
                        SQLQuery = SQLQuery & " VALUES ('Hawaii', 'HI', ''," & igUstCode & ", ''" & ")"
                    Case 12
                        SQLQuery = SQLQuery & " VALUES ('Idaho', 'ID', ''," & igUstCode & ", ''" & ")"
                    Case 13
                        SQLQuery = SQLQuery & " VALUES ('Illinois', 'IL', ''," & igUstCode & ", ''" & ")"
                    Case 14
                        SQLQuery = SQLQuery & " VALUES ('Indiana', 'IN', ''," & igUstCode & ", ''" & ")"
                    Case 15
                        SQLQuery = SQLQuery & " VALUES ('Iowa', 'IA', ''," & igUstCode & ", ''" & ")"
                    Case 16
                        SQLQuery = SQLQuery & " VALUES ('Kansas', 'KS', ''," & igUstCode & ", ''" & ")"
                    Case 17
                        SQLQuery = SQLQuery & " VALUES ('Kentucky', 'KY', ''," & igUstCode & ", ''" & ")"
                    Case 18
                        SQLQuery = SQLQuery & " VALUES ('Louisiana', 'LA', ''," & igUstCode & ", ''" & ")"
                    Case 19
                        SQLQuery = SQLQuery & " VALUES ('Maine', 'ME', ''," & igUstCode & ", ''" & ")"
                    Case 20
                        SQLQuery = SQLQuery & " VALUES ('Maryland', 'MD', ''," & igUstCode & ", ''" & ")"
                    Case 21
                        SQLQuery = SQLQuery & " VALUES ('Massachusetts', 'MA', ''," & igUstCode & ", ''" & ")"
                    Case 22
                        SQLQuery = SQLQuery & " VALUES ('Michigan', 'MI', ''," & igUstCode & ", ''" & ")"
                    Case 23
                        SQLQuery = SQLQuery & " VALUES ('Minnesota', 'MN', ''," & igUstCode & ", ''" & ")"
                    Case 24
                        SQLQuery = SQLQuery & " VALUES ('Mississippi', 'MS', ''," & igUstCode & ", ''" & ")"
                    Case 25
                        SQLQuery = SQLQuery & " VALUES ('Missouri', 'MO', ''," & igUstCode & ", ''" & ")"
                    Case 26
                        SQLQuery = SQLQuery & " VALUES ('Montana', 'MT', ''," & igUstCode & ", ''" & ")"
                    Case 27
                        SQLQuery = SQLQuery & " VALUES ('Nebraska', 'NE', ''," & igUstCode & ", ''" & ")"
                    Case 28
                        SQLQuery = SQLQuery & " VALUES ('Nevada', 'NV', ''," & igUstCode & ", ''" & ")"
                    Case 29
                        SQLQuery = SQLQuery & " VALUES ('New Hampshire', 'NH', ''," & igUstCode & ", ''" & ")"
                    Case 30
                        SQLQuery = SQLQuery & " VALUES ('New Jersey', 'NJ', ''," & igUstCode & ", ''" & ")"
                    Case 31
                        SQLQuery = SQLQuery & " VALUES ('New Mexico', 'NM', ''," & igUstCode & ", ''" & ")"
                    Case 32
                        SQLQuery = SQLQuery & " VALUES ('New York', 'NY', ''," & igUstCode & ", ''" & ")"
                    Case 33
                        SQLQuery = SQLQuery & " VALUES ('North Carolina', 'NC', ''," & igUstCode & ", ''" & ")"
                    Case 34
                        SQLQuery = SQLQuery & " VALUES ('North Dakota', 'ND', ''," & igUstCode & ", ''" & ")"
                    Case 35
                        SQLQuery = SQLQuery & " VALUES ('Ohio', 'OH', ''," & igUstCode & ", ''" & ")"
                    Case 36
                        SQLQuery = SQLQuery & " VALUES ('Oklahoma', 'OK', ''," & igUstCode & ", ''" & ")"
                    Case 37
                        SQLQuery = SQLQuery & " VALUES ('Oregon', 'OR', ''," & igUstCode & ", ''" & ")"
                    Case 38
                        SQLQuery = SQLQuery & " VALUES ('Pennsylvania', 'PA', ''," & igUstCode & ", ''" & ")"
                    Case 39
                        SQLQuery = SQLQuery & " VALUES ('Rhode Island', 'RI', ''," & igUstCode & ", ''" & ")"
                    Case 40
                        SQLQuery = SQLQuery & " VALUES ('South Carolina', 'SC', ''," & igUstCode & ", ''" & ")"
                    Case 41
                        SQLQuery = SQLQuery & " VALUES ('South Dakota', 'SD', ''," & igUstCode & ", ''" & ")"
                    Case 42
                        SQLQuery = SQLQuery & " VALUES ('Tennessee', 'TN', ''," & igUstCode & ", ''" & ")"
                    Case 43
                        SQLQuery = SQLQuery & " VALUES ('Texas', 'TX', ''," & igUstCode & ", ''" & ")"
                    Case 44
                        SQLQuery = SQLQuery & " VALUES ('Utah', 'UT', ''," & igUstCode & ", ''" & ")"
                    Case 45
                        SQLQuery = SQLQuery & " VALUES ('Vermont', 'VT', ''," & igUstCode & ", ''" & ")"
                    Case 46
                        SQLQuery = SQLQuery & " VALUES ('Virginia', 'VA', ''," & igUstCode & ", ''" & ")"
                    Case 47
                        SQLQuery = SQLQuery & " VALUES ('Washington', 'WA', ''," & igUstCode & ", ''" & ")"
                    Case 48
                        SQLQuery = SQLQuery & " VALUES ('West Virginia', 'WV', ''," & igUstCode & ", ''" & ")"
                    Case 49
                        SQLQuery = SQLQuery & " VALUES ('Wisconsin', 'WI', ''," & igUstCode & ", ''" & ")"
                    Case 50
                        SQLQuery = SQLQuery & " VALUES ('Wyoming', 'WY', ''," & igUstCode & ", ''" & ")"
                    Case 51
                        SQLQuery = SQLQuery & " VALUES ('American Somoa', 'AS', ''," & igUstCode & ", ''" & ")"
                    Case 52
                        SQLQuery = SQLQuery & " VALUES ('District of Columbia', 'DC', ''," & igUstCode & ", ''" & ")"
                    Case 53
                        SQLQuery = SQLQuery & " VALUES ('Federated Micronesia', 'FM', ''," & igUstCode & ", ''" & ")"
                    Case 54
                        SQLQuery = SQLQuery & " VALUES ('Guam', 'GU', ''," & igUstCode & ", ''" & ")"
                    Case 55
                        SQLQuery = SQLQuery & " VALUES ('Marshall Islands', 'MH', ''," & igUstCode & ", ''" & ")"
                    Case 56
                        SQLQuery = SQLQuery & " VALUES ('Northern Mariana Islands', 'MP', ''," & igUstCode & ", ''" & ")"
                    Case 57
                        SQLQuery = SQLQuery & " VALUES ('Palau', 'PW', ''," & igUstCode & ", ''" & ")"
                    Case 58
                        SQLQuery = SQLQuery & " VALUES ('Puerto Rico', 'PR', ''," & igUstCode & ", ''" & ")"
                    Case 59
                        SQLQuery = SQLQuery & " VALUES ('Virgin Islands', 'VI', ''," & igUstCode & ", ''" & ")"
                    Case 60
                        SQLQuery = SQLQuery & " VALUES ('Alberta', 'AB', ''," & igUstCode & ", ''" & ")"
                    Case 61
                        SQLQuery = SQLQuery & " VALUES ('British Columbia', 'BC', ''," & igUstCode & ", ''" & ")"
                    Case 62
                        SQLQuery = SQLQuery & " VALUES ('Manitoba', 'MB', ''," & igUstCode & ", ''" & ")"
                    Case 63
                        SQLQuery = SQLQuery & " VALUES ('New Brunswick', 'NB', ''," & igUstCode & ", ''" & ")"
                    Case 64
                        SQLQuery = SQLQuery & " VALUES ('Newfoundland and Labrador', 'NL', ''," & igUstCode & ", ''" & ")"
                    Case 65
                        SQLQuery = SQLQuery & " VALUES ('Northwest Territories', 'NT', ''," & igUstCode & ", ''" & ")"
                    Case 66
                        SQLQuery = SQLQuery & " VALUES ('Nova Scotia', 'NS', ''," & igUstCode & ", ''" & ")"
                    Case 67
                        SQLQuery = SQLQuery & " VALUES ('Nunavut', 'NU', ''," & igUstCode & ", ''" & ")"
                    Case 68
                        SQLQuery = SQLQuery & " VALUES ('Ontario', 'ON', ''," & igUstCode & ", ''" & ")"
                    Case 69
                        SQLQuery = SQLQuery & " VALUES ('Prince Edward Island', 'PE', ''," & igUstCode & ", ''" & ")"
                    Case 70
                        SQLQuery = SQLQuery & " VALUES ('Quebec', 'QC', ''," & igUstCode & ", ''" & ")"
                    Case 71
                        SQLQuery = SQLQuery & " VALUES ('Saskatchewan', 'SK', ''," & igUstCode & ", ''" & ")"
                    Case 72
                        SQLQuery = SQLQuery & " VALUES ('Yukon', 'YT', ''," & igUstCode & ", ''" & ")"
                End Select
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand:
                    gHandleError "AffErrorLog.txt", "modPopSubs-gPopStates"
                    gPopStates = False
                    On Error Resume Next
                    snt_rst.Close
                    Exit Function
                End If
            Next ilLoop
        End If
    Loop While ilUpper = 0

    'Now sort them by the mktCode
    If UBound(tgStateInfo) > 1 Then
        ArraySortTyp fnAV(tgStateInfo(), 0), UBound(tgStateInfo), 0, LenB(tgStateInfo(0)), 0, -1, 0
    End If
   
    snt_rst.Close
    gPopStates = True
    Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopStates"
    snt_rst.Close
    gPopStates = False
    Exit Function
End Function

Public Function gSetStationTimeZones() As Integer
    Dim tzt_rst As ADODB.Recordset
    Dim shtt_rst As ADODB.Recordset
    Dim ilESTCode As Integer
    Dim ilCSTCode As Integer
    Dim ilMSTCode As Integer
    Dim ilPSTCode As Integer
    Dim ilTztCode As Integer

    On Error GoTo ErrHand
    
    SQLQuery = "Select * from " & "TZT"
    Set tzt_rst = gSQLSelectCall(SQLQuery)
    While Not tzt_rst.EOF
        If StrComp(Trim$(tzt_rst!tztName), "Eastern", vbTextCompare) = 0 Then
            ilESTCode = tzt_rst!tztCode
        End If
        If StrComp(Trim$(tzt_rst!tztName), "Central", vbTextCompare) = 0 Then
            ilCSTCode = tzt_rst!tztCode
        End If
        If StrComp(Trim$(tzt_rst!tztName), "Mountain", vbTextCompare) = 0 Then
            ilMSTCode = tzt_rst!tztCode
        End If
        If StrComp(Trim$(tzt_rst!tztName), "Pacific", vbTextCompare) = 0 Then
            ilPSTCode = tzt_rst!tztCode
        End If
        tzt_rst.MoveNext
    Wend
    'Set Stations
    SQLQuery = "SELECT shttTimeZone, shttCode, shttType FROM shtt "
    Set shtt_rst = gSQLSelectCall(SQLQuery)
    While Not shtt_rst.EOF
        ilTztCode = 0
        If StrComp(Trim$(shtt_rst!shttTimeZone), "EST", vbTextCompare) = 0 Then
            ilTztCode = ilESTCode
        End If
        If StrComp(Trim$(shtt_rst!shttTimeZone), "CST", vbTextCompare) = 0 Then
            ilTztCode = ilCSTCode
        End If
        If StrComp(Trim$(shtt_rst!shttTimeZone), "MST", vbTextCompare) = 0 Then
            ilTztCode = ilMSTCode
        End If
        If StrComp(Trim$(shtt_rst!shttTimeZone), "PST", vbTextCompare) = 0 Then
            ilTztCode = ilPSTCode
        End If
        
        If ilTztCode > 0 Then
            SQLQuery = "UPDATE shtt"
            SQLQuery = SQLQuery & " SET shttTztCode = " & ilTztCode
            SQLQuery = SQLQuery & " WHERE shttCode = " & shtt_rst!shttCode
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand:
                gHandleError "AffErrorLog.txt", "modPopSubs-gSetStationTimeZones"
                gSetStationTimeZones = False
                On Error Resume Next
                tzt_rst.Close
                shtt_rst.Close
                Exit Function
            End If
        End If
        shtt_rst.MoveNext
    Wend
    tzt_rst.Close
    shtt_rst.Close
    gSetStationTimeZones = True
    Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gSetStationTimeZones"
    tzt_rst.Close
    shtt_rst.Close
    gSetStationTimeZones = False
    Exit Function
End Function
Public Function gBinarySearchTzt(ilCode As Integer) As Integer
    
    'Returns the index number of tgTimeZoneInfo that matches the tztCode that was passed in
    'Note: for this to work tgTimeZoneInfo was previously be sorted by tztCode
    
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    
    On Error GoTo ErrHand
    
    ilMin = LBound(tgTimeZoneInfo)
    ilMax = UBound(tgTimeZoneInfo) - 1
    Do While ilMin <= ilMax
        ilMiddle = (CLng(ilMin) + ilMax) \ 2
        If ilCode = tgTimeZoneInfo(ilMiddle).iCode Then
            'found the match
            gBinarySearchTzt = ilMiddle
            Exit Function
        ElseIf ilCode < tgTimeZoneInfo(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    
    gBinarySearchTzt = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchTzt: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchTzt = -1
    Exit Function
    
End Function

Public Function gBinarySearchSnt(ilCode As Integer) As Integer
    
    'Returns the index number of tgStateInfo that matches the SntCode that was passed in
    'Note: for this to work tgStateInfo was previously be sorted by SntCode
    
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    
    On Error GoTo ErrHand
    
    ilMin = LBound(tgStateInfo)
    ilMax = UBound(tgStateInfo) - 1
    Do While ilMin <= ilMax
        ilMiddle = (CLng(ilMin) + ilMax) \ 2
        If ilCode = tgStateInfo(ilMiddle).iCode Then
            'found the match
            gBinarySearchSnt = ilMiddle
            Exit Function
        ElseIf ilCode < tgStateInfo(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    
    gBinarySearchSnt = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchSnt: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchSnt = -1
    Exit Function
    
End Function

Public Function gBinarySearchVff(ilCode As Integer) As Integer
    
    'Returns the index number of tgVffInfo that matches the vefCode that was passed in
    'Note: for this to work tgVffInfo was previously be sorted by vefCode
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    On Error GoTo ErrHand
    
    llMin = LBound(tgVffInfo)
    llMax = UBound(tgVffInfo) - 1
    Do While llMin <= llMax
        llMiddle = (CLng(llMin) + llMax) \ 2
        If ilCode = tgVffInfo(llMiddle).iVefCode Then
            'found the match
            gBinarySearchVff = llMiddle
            Exit Function
        ElseIf ilCode < tgVffInfo(llMiddle).iVefCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    
    gBinarySearchVff = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchVff: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchVff = -1
    Exit Function
    
End Function

Public Function gBinarySearchFmt(llCode As Long) As Integer
    
    'Returns the index number of tgFormatInfo that matches the FmtCode that was passed in
    'Note: for this to work tgFormatInfo was previously be sorted by FmtCode
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    On Error GoTo ErrHand
    
    llMin = LBound(tgFormatInfo)
    llMax = UBound(tgFormatInfo) - 1
    Do While llMin <= llMax
        llMiddle = (CLng(llMin) + llMax) \ 2
        If llCode = tgFormatInfo(llMiddle).lCode Then
            'found the match
            gBinarySearchFmt = llMiddle
            Exit Function
        ElseIf llCode < tgFormatInfo(llMiddle).lCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    
    gBinarySearchFmt = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchFmt: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchFmt = -1
    Exit Function
    
End Function

Public Function gPopVff() As Integer

    'D.S. 01/06/06

    Dim vff_rst As ADODB.Recordset
    Dim ilUpper As Integer
    Dim llMax As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(vffCode) from VFF_Vehicle_Features"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsNull(rst(0).Value) Then
        ReDim tgVffInfo(0 To 0) As VEHICLEFEATURESINFO
        gPopVff = True
        Exit Function
    End If
    
    llMax = rst(0).Value
    ReDim tgVffInfo(0 To llMax) As VEHICLEFEATURESINFO
    
    SQLQuery = "Select * from VFF_Vehicle_Features"
    Set vff_rst = gSQLSelectCall(SQLQuery)
    ilUpper = 0
    While Not vff_rst.EOF
        tgVffInfo(ilUpper).iVefCode = vff_rst!vffvefCode
        tgVffInfo(ilUpper).sGroupName = vff_rst!VffGroupName
        tgVffInfo(ilUpper).sWegenerExportID = vff_rst!VffWegenerExportID
        tgVffInfo(ilUpper).sOLAExportID = vff_rst!VffOLAExportID
        tgVffInfo(ilUpper).iLiveCompliantAdj = vff_rst!VffLiveCompliantAdj
        tgVffInfo(ilUpper).sXDXMLForm = vff_rst!vffXDXMLForm
        tgVffInfo(ilUpper).sXDISCIPrefix = vff_rst!vffXDISCIPrefix
        tgVffInfo(ilUpper).sXDSaveCF = vff_rst!vffXDSaveCF
        tgVffInfo(ilUpper).sXDSaveHDD = vff_rst!vffXDSaveHDD
        tgVffInfo(ilUpper).sXDSaveNAS = vff_rst!vffXDSaveNAS
        tgVffInfo(ilUpper).sXDProgCodeID = vff_rst!vffxdprogcodeid
        tgVffInfo(ilUpper).sPledgeVsAir = vff_rst!vffPledgeVsAir
        tgVffInfo(ilUpper).sMergeAffiliate = vff_rst!vffMergeAffiliate
        tgVffInfo(ilUpper).sMergeTraffic = vff_rst!vffMergeTraffic
        tgVffInfo(ilUpper).sMergeWeb = vff_rst!vffMergeWeb
        tgVffInfo(ilUpper).sWebName = vff_rst!vffWebName
        tgVffInfo(ilUpper).sPledgeByEvent = vff_rst!vffPledgeByEvent
        tgVffInfo(ilUpper).sIPumpEventTypeOV = vff_rst!vffIPumpEventTypeOV
        tgVffInfo(ilUpper).sExportIPump = vff_rst!vffExportIPump
        tgVffInfo(ilUpper).sXDSISCIPrefix = vff_rst!vffXDSISCIPrefix
        tgVffInfo(ilUpper).sXDSSaveCF = vff_rst!vffXDSSaveCF
        tgVffInfo(ilUpper).sXDSSaveHDD = vff_rst!vffXDSSaveHDD
        tgVffInfo(ilUpper).sXDSSaveNAS = vff_rst!vffXDSSaveNAS
        tgVffInfo(ilUpper).sSentToXDS = vff_rst!vffSentToXDSStatus
        tgVffInfo(ilUpper).sExportJelli = vff_rst!vffExportJelli
        tgVffInfo(ilUpper).sMGsOnWeb = vff_rst!vffMGsOnWeb
        tgVffInfo(ilUpper).sReplacementOnWeb = vff_rst!vffReplacementOnWeb
        tgVffInfo(ilUpper).sStationComp = vff_rst!vffStationComp
        tgVffInfo(ilUpper).sHonorZeroUnits = vff_rst!vffHonorZeroUnits
        tgVffInfo(ilUpper).sHideCommOnWeb = vff_rst!vffHideCommOnWeb
            '10933
        tgVffInfo(ilUpper).sXDEventZone = vff_rst!vffXDEventZone
        ilUpper = ilUpper + 1
        vff_rst.MoveNext
    Wend

    ReDim Preserve tgVffInfo(0 To ilUpper) As VEHICLEFEATURESINFO

    'Now sort them by the mktCode
    If UBound(tgVffInfo) > 1 Then
        ArraySortTyp fnAV(tgVffInfo(), 0), UBound(tgVffInfo), 0, LenB(tgVffInfo(1)), 0, -1, 0
    End If
   
   gPopVff = True
   vff_rst.Close
   rst.Close
   Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopVff"
    gPopVff = False
    Exit Function
End Function

Public Function gPopMediaCodes() As Integer
    
    Dim iUpper As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    iUpper = 0
    ReDim tgMediaCodesInfo(0 To 0) As MEDIACODESINFO
    SQLQuery = "SELECT mcfName, mcfCode"
    SQLQuery = SQLQuery + " FROM MCF_MEDIA_CODE "
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        tgMediaCodesInfo(iUpper).iCode = rst!mcfCode
        tgMediaCodesInfo(iUpper).sName = Trim$(rst!mcfName)
        iUpper = iUpper + 1
        ReDim Preserve tgMediaCodesInfo(0 To iUpper) As MEDIACODESINFO
        rst.MoveNext
    Wend
    'Now sort them by the mcfCode
    If UBound(tgMediaCodesInfo) > 1 Then
        ArraySortTyp fnAV(tgMediaCodesInfo(), 0), UBound(tgMediaCodesInfo), 0, LenB(tgMediaCodesInfo(0)), 0, -1, 0
    End If
    gPopMediaCodes = True
    rst.Close
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopMediaCodes"
    gPopMediaCodes = False
    Exit Function
End Function

Public Function gBinarySearchAnf(ilCode As Integer) As Integer
    
    'Returns the index number of tgFormatInfo that matches the FmtCode that was passed in
    'Note: for this to work tgFormatInfo was previously be sorted by FmtCode
    
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    
    On Error GoTo ErrHand
    
    ilMin = LBound(tgAvailNamesInfo)
    ilMax = UBound(tgAvailNamesInfo) - 1
    Do While ilMin <= ilMax
        ilMiddle = (CLng(ilMin) + ilMax) \ 2
        If ilCode = tgAvailNamesInfo(ilMiddle).iCode Then
            'found the match
            gBinarySearchAnf = ilMiddle
            Exit Function
        ElseIf ilCode < tgAvailNamesInfo(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    
    gBinarySearchAnf = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchAnf: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchAnf = -1
    Exit Function
    
End Function

Public Function gBinarySearchMcf(ilCode As Integer) As Integer
    
    'Returns the index number of tgFormatInfo that matches the FmtCode that was passed in
    'Note: for this to work tgFormatInfo was previously be sorted by FmtCode
    
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    
    On Error GoTo ErrHand
    
    ilMin = LBound(tgMediaCodesInfo)
    ilMax = UBound(tgMediaCodesInfo) - 1
    Do While ilMin <= ilMax
        ilMiddle = (CLng(ilMin) + ilMax) \ 2
        If ilCode = tgMediaCodesInfo(ilMiddle).iCode Then
            'found the match
            gBinarySearchMcf = ilMiddle
            Exit Function
        ElseIf ilCode < tgMediaCodesInfo(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    
    gBinarySearchMcf = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchMcf: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchMcf = -1
    Exit Function
    
End Function

Public Function gPopMntInfo(slType As String, tlMultiNameInfo() As MNTINFO) As Integer

    'D.S. 01/06/06

    Dim mnt_rst As ADODB.Recordset
    Dim ilUpper As Integer
    Dim llMax As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(mntCode) from mnt"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsNull(rst(0).Value) Then
        ReDim tlMultiNameInfo(0 To 0) As MNTINFO
        gPopMntInfo = True
        Exit Function
    End If
    
    llMax = rst(0).Value
    ReDim tlMultiNameInfo(0 To llMax) As MNTINFO
    
    SQLQuery = "SELECT mntCode, mntName, mntState FROM mnt WHERE mntType = '" & slType & "'" & " ORDER BY mntCode"
    Set mnt_rst = gSQLSelectCall(SQLQuery)
    ilUpper = 0
    While Not mnt_rst.EOF
        tlMultiNameInfo(ilUpper).lCode = mnt_rst!mntCode
        tlMultiNameInfo(ilUpper).sName = mnt_rst!mntName
        tlMultiNameInfo(ilUpper).sState = mnt_rst!mntState
        ilUpper = ilUpper + 1
        mnt_rst.MoveNext
    Wend

    ReDim Preserve tlMultiNameInfo(0 To ilUpper) As MNTINFO

    'Now sort them by the mntCode
    If UBound(tlMultiNameInfo) > 1 Then
        ArraySortTyp fnAV(tlMultiNameInfo(), 0), UBound(tlMultiNameInfo), 0, LenB(tlMultiNameInfo(0)), 0, -2, 0
    End If
   
   gPopMntInfo = True
   mnt_rst.Close
   rst.Close
   Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopMntInfo"
    gPopMntInfo = False
    Exit Function
End Function

Public Function gPopRepInfo(slType As String, tlRepInfo() As REPINFO) As Integer

    'slType: M=Market Rep; S=Service Rep

    Dim rst_rep As ADODB.Recordset
    Dim ilUpper As Integer
    Dim llMax As Long
    
    On Error GoTo ErrHand
    
    
    ReDim tlRepInfo(0 To 100) As REPINFO
    ilUpper = 0
    
    SQLQuery = "SELECT UstName, ustReportName, ustCode FROM Ust INNER JOIN Dnt ON ustDntCode = dntCode WHERE dntType = '" & slType & "'"
    Set rst_rep = gSQLSelectCall(SQLQuery)
    Do While Not rst_rep.EOF
        If Trim$(rst_rep!ustReportName) <> "" Then
            tlRepInfo(ilUpper).sName = Trim$(rst_rep!ustReportName)
        Else
            tlRepInfo(ilUpper).sName = Trim$(rst_rep!ustname)
        End If
        tlRepInfo(ilUpper).sLogInName = Trim$(rst_rep!ustname)
        tlRepInfo(ilUpper).sReportName = Trim$(rst_rep!ustReportName)
        tlRepInfo(ilUpper).iUstCode = rst_rep!ustCode   'sort field
        If ilUpper = UBound(tlRepInfo) Then
            ReDim Preserve tlRepInfo(0 To UBound(tlRepInfo) + 100) As REPINFO
        End If
        ilUpper = ilUpper + 1
        rst_rep.MoveNext
    Loop

    ReDim Preserve tlRepInfo(0 To ilUpper) As REPINFO

    'Now sort them by the mntCode
    If UBound(tlRepInfo) > 1 Then
        ArraySortTyp fnAV(tlRepInfo(), 0), UBound(tlRepInfo), 0, LenB(tlRepInfo(0)), 0, -1, 0
    End If
   
   gPopRepInfo = True
   rst_rep.Close
   Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopRepIInfo"
    gPopRepInfo = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gFileDateTime                   *
'*                                                     *
'*             Created:10/22/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain file time stamp          *
'*                                                     *
'*******************************************************
Function gFileDateTime(slPathFile As String) As String
    Dim ilRet As Integer

    ilRet = 0
    'On Error GoTo gFileDateTimeErr
    'gFileDateTime = FileDateTime(slPathFile)
    ilRet = gFileExist(slPathFile)
    If ilRet = 0 Then
        gFileDateTime = FileDateTime(slPathFile)
    Else
        gFileDateTime = Format$(Now, "m/d/yy") & " " & Format$(Now, "h:mm:ssAM/PM")
    End If
    On Error GoTo 0
    Exit Function
gFileDateTimeErr:
    ilRet = Err.Number
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gFileExist                      *
'*                                                     *
'*             Created:1/4/16        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determina if file exist         *
'*                                                     *
'*******************************************************
Function gFileExist(slPathFile As String) As Integer
    Dim fs As New FileSystemObject

    If fs.FILEEXISTS(slPathFile) Then
        gFileExist = 0
    Else
        gFileExist = 1
    End If
End Function
Function gFolderExist(slPathFolder As String) As Boolean
    '8886
    Dim fs As New FileSystemObject

    If fs.FolderExists(slPathFolder) Then
        gFolderExist = True
    Else
        gFolderExist = False
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gFileOpen                       *
'*                                                     *
'*             Created:1/4/16      By:D. LeVine        *
'*            Modified:            By:D. Smith         *
'*                                                     *
'*            Comments:Open a File                     *
'*                                                     *
'*******************************************************
Function gFileOpen(slPathFile As String, slForClause As String, hlHandle As Integer) As Integer
    Dim slCase As String

    gFileOpen = 0
    slForClause = UCase$(slForClause)
    On Error GoTo gFileOpenErr
    hlHandle = FreeFile
    slCase = ""
    
    If InStr(1, slForClause, "APPEND", vbTextCompare) > 0 Then
        slCase = "A"
    ElseIf InStr(1, slForClause, "BINARY", vbTextCompare) > 0 Then
        slCase = "B"
    ElseIf InStr(1, slForClause, "INPUT", vbTextCompare) > 0 Then
        slCase = "I"
    ElseIf InStr(1, slForClause, "OUTPUT", vbTextCompare) > 0 Then
        slCase = "O"
    ElseIf InStr(1, slForClause, "RANDOM", vbTextCompare) > 0 Then
        slCase = "R"
    Else
        slCase = "R"
    End If
    
    If InStr(1, slForClause, "ACCESS", vbTextCompare) > 0 Then
        If InStr(1, slForClause, "READ WRITE", vbTextCompare) > 0 Then
            slCase = slCase & "-B"
        ElseIf InStr(1, slForClause, "READ", vbTextCompare) > 0 Then
            slCase = slCase & "-R"
        ElseIf InStr(1, slForClause, "WRITE", vbTextCompare) > 0 Then
            slCase = slCase & "-W"
        End If
    End If
    
    If InStr(1, slForClause, "SHARED", vbTextCompare) > 0 Then
        slCase = slCase & ":S"
    ElseIf InStr(1, slForClause, "LOCK READ WRITE", vbTextCompare) > 0 Then
        slCase = slCase & ":B"
    ElseIf InStr(1, slForClause, "LOCK READ", vbTextCompare) > 0 Then
        slCase = slCase & ":R"
    ElseIf InStr(1, slForClause, "LOCK WRITE", vbTextCompare) > 0 Then
        slCase = slCase & ":W"
    End If
    
    Select Case slCase
        Case "A"
            Open slPathFile For Append As hlHandle
        Case "B"
            Open slPathFile For Binary As hlHandle
        Case "I"
            Open slPathFile For Input As hlHandle
        Case "O"
            Open slPathFile For Output As hlHandle
        Case "R"
            Open slPathFile For Random As hlHandle
           
        Case "A-B"
            Open slPathFile For Append Access Read Write As hlHandle
        Case "A-W"
            Open slPathFile For Append Access Write As hlHandle
            
        Case "B-B"
            Open slPathFile For Binary Access Read Write As hlHandle
        Case "B-R"
            Open slPathFile For Binary Access Read As hlHandle
        Case "B-W"
            Open slPathFile For Binary Access Write As hlHandle
            
        Case "I-R"
            Open slPathFile For Input Access Read As hlHandle
            
        Case "O-W"
            Open slPathFile For Output Access Write As hlHandle
            
        Case "R-B"
            Open slPathFile For Random Access Read Write As hlHandle
        Case "R-R"
            Open slPathFile For Random Access Read As hlHandle
        Case "R-W"
            Open slPathFile For Random Access Write As hlHandle


        Case "A:S"
            Open slPathFile For Append Shared As hlHandle
        Case "A:B"
            Open slPathFile For Append Lock Read Write As hlHandle
        Case "A:R"
            Open slPathFile For Append Lock Read As hlHandle
        Case "A:W"
            Open slPathFile For Append Lock Write As hlHandle

        Case "B:S"
            Open slPathFile For Binary Shared As hlHandle
        Case "B:B"
            Open slPathFile For Binary Lock Read Write As hlHandle
        Case "B:R"
            Open slPathFile For Binary Lock Read As hlHandle
        Case "B:W"
            Open slPathFile For Binary Lock Write As hlHandle

        Case "I:S"
            Open slPathFile For Input Shared As hlHandle
        Case "I:B"
            Open slPathFile For Input Lock Read Write As hlHandle
        Case "I:R"
            Open slPathFile For Input Lock Read As hlHandle
        Case "I:W"
            Open slPathFile For Input Lock Write As hlHandle

        Case "O:S"
            Open slPathFile For Output Shared As hlHandle
        Case "O:B"
            Open slPathFile For Output Lock Read Write As hlHandle
        Case "O:R"
            Open slPathFile For Output Lock Read As hlHandle
        Case "O:W"
            Open slPathFile For Output Lock Write As hlHandle

        Case "R:S"
            Open slPathFile For Random Shared As hlHandle
        Case "R:B"
            Open slPathFile For Random Lock Read Write As hlHandle
        Case "R:R"
            Open slPathFile For Random Lock Read As hlHandle
        Case "R:W"
            Open slPathFile For Random Lock Write As hlHandle
                        
            
        Case "A-B:S"
            Open slPathFile For Append Access Read Write Shared As hlHandle
        Case "A-B:B"
            Open slPathFile For Append Access Read Write Lock Read Write As hlHandle
        Case "A-B:R"
            Open slPathFile For Append Access Read Write Lock Read As hlHandle
        Case "A-B:W"
            Open slPathFile For Append Access Read Write Lock Write As hlHandle
            
        Case "A-W:S"
            Open slPathFile For Append Access Write Shared As hlHandle
        Case "A-W:B"
            Open slPathFile For Append Access Write Lock Read Write As hlHandle
        Case "A-W:R"
            Open slPathFile For Append Access Write Lock Read As hlHandle
        Case "A-W:W"
            Open slPathFile For Append Access Write Lock Write As hlHandle
            
        Case "B-B:S"
            Open slPathFile For Binary Access Read Write Shared As hlHandle
        Case "B-B:B"
            Open slPathFile For Binary Access Read Write Lock Read Write As hlHandle
        Case "B-B:R"
            Open slPathFile For Binary Access Read Write Lock Read As hlHandle
        Case "B-B:W"
            Open slPathFile For Binary Access Read Write Lock Write As hlHandle
            
        Case "B-R:S"
            Open slPathFile For Binary Access Read Shared As hlHandle
        Case "B-R:B"
            Open slPathFile For Binary Access Read Lock Read Write As hlHandle
        Case "B-R:R"
            Open slPathFile For Binary Access Read Lock Read As hlHandle
        Case "B-R:W"
            Open slPathFile For Binary Access Read Lock Write As hlHandle
            
        Case "B-W:S"
            Open slPathFile For Binary Access Write Shared As hlHandle
        Case "B-W:B"
            Open slPathFile For Binary Access Write Lock Read Write As hlHandle
        Case "B-W:R"
            Open slPathFile For Binary Access Write Lock Read As hlHandle
        Case "B-W:W"
            Open slPathFile For Binary Access Write Lock Write As hlHandle
            
        Case "I-R:S"
            Open slPathFile For Input Access Read Shared As hlHandle
        Case "I-R:B"
            Open slPathFile For Input Access Read Lock Read Write As hlHandle
        Case "I-R:R"
            Open slPathFile For Input Access Read Lock Read As hlHandle
        Case "I-R:W"
            Open slPathFile For Input Access Read Lock Write As hlHandle
            
        Case "O-W:S"
            Open slPathFile For Output Access Write Shared As hlHandle
        Case "O-W:B"
            Open slPathFile For Output Access Write Lock Read Write As hlHandle
        Case "O-W:R"
            Open slPathFile For Output Access Write Lock Read As hlHandle
        Case "O-W:W"
            Open slPathFile For Output Access Write Lock Write As hlHandle
            
        Case "R-B:S"
            Open slPathFile For Random Access Read Write Shared As hlHandle
        Case "R-B:B"
            Open slPathFile For Random Access Read Write Lock Read Write As hlHandle
        Case "R-B:R"
            Open slPathFile For Random Access Read Write Lock Read As hlHandle
        Case "R-B:W"
            Open slPathFile For Random Access Read Write Lock Write As hlHandle
            
        Case "R-R:S"
            Open slPathFile For Random Access Read Shared As hlHandle
        Case "R-R:B"
            Open slPathFile For Random Access Read Lock Read Write As hlHandle
        Case "R-R:R"
            Open slPathFile For Random Access Read Lock Read As hlHandle
        Case "R-R:W"
            Open slPathFile For Random Access Read Lock Write As hlHandle
            
        Case "R-W:S"
            Open slPathFile For Random Access Write Shared As hlHandle
        Case "R-W:B"
            Open slPathFile For Random Access Write Lock Read Write As hlHandle
        Case "R-W:R"
            Open slPathFile For Random Access Write Lock Read As hlHandle
        Case "R-W:W"
            Open slPathFile For Random Access Write Lock Write As hlHandle
            
        Case Else
            gLogMsg "Unknown file open type: " & slForClause, "AffErrorLog.Txt", False
            gFileOpen = 2
    End Select
    On Error GoTo 0
    Exit Function
gFileOpenErr:
    On Error GoTo 0
    gFileOpen = 1
End Function
Public Function gBinarySearchCifCpf(llCode As Long) As Long
    
    'D.S. 10/7/11
            
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    llMin = LBound(tgCifCpfInfo1)
    llMax = UBound(tgCifCpfInfo1) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tgCifCpfInfo1(llMiddle).cifCode Then
            'found the match
            gBinarySearchCifCpf = llMiddle
            Exit Function
        ElseIf llCode < tgCifCpfInfo1(llMiddle).cifCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchCifCpf = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchCIF: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchCifCpf = -1
    Exit Function
    
End Function

Public Function gPopCifCpfInfo(sInStartDate As String) As Integer   'TTP 9923 changed sStartDate to sInStartDate

    'D.S. 10/7/11

    Dim llUpper As Long
    Dim llMax As Long
    Dim rst_Cif As ADODB.Recordset
    'TTP 9923 below
    Dim sStartDate As String
    
    On Error GoTo ErrHand
    
    'TTP 9923
    sStartDate = sInStartDate
    If Trim(sStartDate) = "" Then
        sStartDate = "1970-12-31"
    End If
    
    SQLQuery = "Select MAX(cifCode) from CIF_Copy_Inventory"
    Set rst_Cif = gSQLSelectCall(SQLQuery)
    If IsNull(rst_Cif(0).Value) Then
        ReDim tgCifCpfInfo1(0 To 0) As CIFCPFINFO1
        gPopCifCpfInfo = True
        Exit Function
    End If
    llMax = rst_Cif(0).Value
    
    llUpper = 0
    ReDim tgCifCpfInfo1(0 To llMax) As CIFCPFINFO1
    'SQLQuery = "SELECT cifCode, cifRotEndDate FROM cif_Copy_Inventory WHERE cifRotEndDate >= " & "'" & Format$(sStartDate, sgSQLDateForm) & "'" & " ORDER BY cifCode"
    SQLQuery = "Select cifCode, cifRotEndDate, cifAdfCode, cifName, cifMcfCode, cifReel, cifCpfCode, cpfISCI, cpfCreative, cpfName from CIF_Copy_Inventory, CPF_Copy_Prodct_ISCI"
    '10200 rotEndDate could be null
'    SQLQuery = SQLQuery & " WHERE cifCpfCode = cpfCode And cifRotEndDate >= " & "'" & Format$(sStartDate, sgSQLDateForm) & "'" & " ORDER BY cifCode"
    SQLQuery = SQLQuery & " WHERE cifCpfCode = cpfCode And (cifRotEndDate >= " & "'" & Format$(sStartDate, sgSQLDateForm) & "' OR cifRotEndDate is null)" & " ORDER BY cifCode"

    Set rst_Cif = gSQLSelectCall(SQLQuery)
    
    While Not rst_Cif.EOF
        tgCifCpfInfo1(llUpper).cifCode = rst_Cif!cifCode
        tgCifCpfInfo1(llUpper).cifAdfCode = rst_Cif!cifAdfCode
        tgCifCpfInfo1(llUpper).cifCode = rst_Cif!cifCode
        tgCifCpfInfo1(llUpper).cifCpfCode = rst_Cif!cifCpfCode
        tgCifCpfInfo1(llUpper).cifMcfCode = rst_Cif!cifMcfCode
        tgCifCpfInfo1(llUpper).cifName = rst_Cif!cifName
        tgCifCpfInfo1(llUpper).cifReel = rst_Cif!cifReel
        '10200 rotEndDate could be null
       ' tgCifCpfInfo1(llUpper).cifRotEndDate = rst_Cif!cifRotEndDate
        If IsNull(rst_Cif!cifRotEndDate) Then
            tgCifCpfInfo1(llUpper).cifRotEndDate = ""
        Else
            tgCifCpfInfo1(llUpper).cifRotEndDate = rst_Cif!cifRotEndDate
        End If
        tgCifCpfInfo1(llUpper).cpfCreative = rst_Cif!cpfCreative
        tgCifCpfInfo1(llUpper).cpfISCI = rst_Cif!cpfISCI
        tgCifCpfInfo1(llUpper).cpfName = rst_Cif!cpfName
        llUpper = llUpper + 1
        rst_Cif.MoveNext
    Wend
    
    ReDim Preserve tgCifCpfInfo1(0 To llUpper) As CIFCPFINFO1
    lgAttCount = llUpper
    
    gPopCifCpfInfo = True
    rst_Cif.Close
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopCifCpfInfo"
    gPopCifCpfInfo = False
    Exit Function
End Function


Public Function gPopCrfInfo(sStartDate As String) As Integer

    'D.S. 10/12/11

    Dim llUpper As Long
    Dim llMax As Long
    Dim rst_crf As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "select count(crfcode) from crf_copy_rot_header where crfrottype ='A' and crfbkoutinstadfcode = 0 and crfenddate >= " & "'" & Format(sStartDate, sgSQLDateForm) & "'" & " and crfcsfCode > 0"
    Set rst_crf = gSQLSelectCall(SQLQuery)
    If IsNull(rst_crf(0).Value) Then
        ReDim tgCrfInfo1(0 To 0) As CRFINFO1
        gPopCrfInfo = True
        Exit Function
    End If
    llMax = rst_crf(0).Value
    
    SQLQuery = "select crfcode, crfcsfCode from crf_copy_rot_header where crfrottype ='A' and crfbkoutinstadfcode = 0 and crfenddate >= " & "'" & Format(sStartDate, sgSQLDateForm) & "'" & " and crfcsfCode > 0 order by crfCode"
    Set rst_crf = gSQLSelectCall(SQLQuery)
    
    llUpper = 0
    ReDim tgCrfInfo1(0 To llMax) As CRFINFO1
    
    While Not rst_crf.EOF
        tgCrfInfo1(llUpper).crfCode = rst_crf!crfCode
        tgCrfInfo1(llUpper).crfCsfCode = rst_crf!crfCsfCode
        llUpper = llUpper + 1
        rst_crf.MoveNext
    Wend
    
    ReDim Preserve tgCrfInfo1(0 To llUpper) As CRFINFO1
    lgAttCount = llUpper
    
    gPopCrfInfo = True
    rst_crf.Close
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopCrfInfo"
    gPopCrfInfo = False
    Exit Function
End Function

Public Function gBinarySearchCrf(llCode As Long) As Long
    
    'D.S. 10/12/11
            
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    llMin = LBound(tgCrfInfo1)
    llMax = UBound(tgCrfInfo1) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tgCrfInfo1(llMiddle).crfCode Then
            'found the match
            gBinarySearchCrf = llMiddle
            Exit Function
        ElseIf llCode < tgCrfInfo1(llMiddle).crfCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchCrf = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchCrf: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchCrf = -1
    Exit Function
    
End Function

Public Function gPopCopy(slDate As String, slMsg As String) As Integer
    Dim ilRet As Integer
    Dim ilLimit As Integer
    
    'On Error GoTo gPopCopyErr1:
    'ilLimit = LBound(tgCifCpfInfo1)
    'On Error GoTo gPopCopyErr2:
    'ilLimit = LBound(tgCrfInfo1)
    'On Error GoTo 0
    If PeekArray(tgCifCpfInfo1).Ptr <> 0 Then
        ilLimit = LBound(tgCifCpfInfo1)
    Else
        ReDim tgCifCpfInfo1(0 To 0) As CIFCPFINFO1
        ilLimit = 0
    End If
    If PeekArray(tgCrfInfo1).Ptr <> 0 Then
        ilLimit = LBound(tgCrfInfo1)
    Else
        ReDim tgCrfInfo1(0 To 0) As CRFINFO1
        ilLimit = 0
    End If
    
    ilRet = gPopCifCpfInfo(slDate)
    If ilRet Then
        ilRet = gPopCrfInfo(slDate)
        If Not ilRet Then
            Screen.MousePointer = vbDefault
            Beep
            gMsgBox "gPopCrfInfo failed in " & slMsg
            gPopCopy = False
            Exit Function
        End If
    Else
        Screen.MousePointer = vbDefault
        Beep
        gMsgBox "gPopCifCpfInfo failed in " & slMsg
        gPopCopy = False
        Exit Function
    End If
    gPopCopy = True
    Exit Function
gPopCopyErr1:
    ReDim tgCifCpfInfo1(0 To 0) As CIFCPFINFO1
    Resume Next
gPopCopyErr2:
    ReDim tgCrfInfo1(0 To 0) As CRFINFO1
    Resume Next
End Function
'
'           gPopSpotStatusCodesExt (extended version with new status codes for MG feature
'               <input> lbcStatus: list box containing status code selection
'                       ilHideNotCarried - true to deselected Not Carried status
'
Public Sub gPopSpotStatusCodesExt(lbcStatus As control, ilHideNotCarried As Integer)
Dim lRg As Long
Dim lRet As Long
    lbcStatus.Clear
    lbcStatus.AddItem "1-Aired Live"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 0
    lbcStatus.AddItem "2-Aired Delay Bcast"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 1
    lbcStatus.AddItem "3-Not Aired Tech Diff"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 2
    lbcStatus.AddItem "4-Not Aired Blackout"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 3
    lbcStatus.AddItem "5-Not Aired Other"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 4
    lbcStatus.AddItem "6-Not Aired Product"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 5
    lbcStatus.AddItem "7-Aired Outside Pledge"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 6
    lbcStatus.AddItem "8-Aired Not Pledged"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 7
    lbcStatus.AddItem "9-Not Carried"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 8
    lbcStatus.AddItem "10-Delay Cmml/Prg"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 9
    lbcStatus.AddItem "11-Air Cmml Only"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 10
    lbcStatus.AddItem "12-MG"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 11
    lbcStatus.AddItem "13-Bonus"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 12
    lbcStatus.AddItem "14-Replacement"
    lbcStatus.ItemData(lbcStatus.NewIndex) = 13
    If sgMissedMGBypass = "Y" Then                 '4-12-17 feature needs to be set in site to allow missed spot to bypass a mg
        lbcStatus.AddItem "15-Missed-MG Bypass"
        lbcStatus.ItemData(lbcStatus.NewIndex) = 14
    End If

 
    lRg = CLng(lbcStatus.ListCount - 1) * &H10000 Or 0
    lRet = SendMessageByNum(lbcStatus.hwnd, LB_SELITEMRANGE, True, lRg)
    If ilHideNotCarried Then            '9-18-08 default to deselected Not Carried?
        lbcStatus.Selected(8) = False
    End If
End Sub

Public Function gPopPet(llAttCode As Long, tlPetInfo() As PETINFO) As Integer
    Dim iUpper As Integer
    Dim sChar  As String * 1
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    iUpper = 0
    ReDim tlPetInfo(0 To 0) As PETINFO
    SQLQuery = "SELECT petCode, petGsfCode, petDeclaredStatus, petClearStatus, gsfghfCode"
    SQLQuery = SQLQuery + " FROM pet LEFT OUTER JOIN gsf_Game_Schd ON petGsfCode = gsfCode"
    SQLQuery = SQLQuery & " WHERE (petAttCode = " & llAttCode & ")"
    SQLQuery = SQLQuery + " ORDER BY gsfGhfCode, petGsfCode"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        tlPetInfo(iUpper).lCode = rst!petCode
        tlPetInfo(iUpper).lGhfCode = rst!gsfGhfCode
        tlPetInfo(iUpper).lgsfCode = rst!petGsfCode
        tlPetInfo(iUpper).sDeclaredStatus = rst!petDeclaredStatus
        tlPetInfo(iUpper).sClearStatus = rst!petClearStatus
        tlPetInfo(iUpper).sChanged = "N"
        iUpper = iUpper + 1
        ReDim Preserve tlPetInfo(0 To iUpper) As PETINFO
        rst.MoveNext
    Wend
   
    gPopPet = True
    rst.Close
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmPopSubs-gPopPet"
    gPopPet = False
    Exit Function
End Function

Public Function gBinarySearchPet(llGsfCode As Long) As Long

    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    On Error GoTo ErrHand
    
    llMin = LBound(tgPetInfo)
    llMax = UBound(tgPetInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llGsfCode = tgPetInfo(llMiddle).lgsfCode Then
            'found the match
            gBinarySearchPet = llMiddle
            Exit Function
        ElseIf llGsfCode < tgPetInfo(llMiddle).lgsfCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchPet = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchPet: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchPet = -1
    Exit Function
End Function

Public Function gPopDaypart() As Integer

    Dim ilUpper As Integer
    Dim ilMax As Integer
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slStart As String
    Dim slEnd As String
    Dim ilLimit As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim rdf_rst As ADODB.Recordset
    
    
    'On Error GoTo gPopDaypartErr:
    'ilRet = 0
    'ilLimit = LBound(tgDaypartInfo)
    'If ilRet = 0 Then
    '    gPopDaypart = True
    '    Exit Function
    'End If
    If PeekArray(tgDaypartInfo).Ptr <> 0 Then
        ilLimit = LBound(tgDaypartInfo)
        gPopDaypart = True
        Exit Function
    Else
        ilLimit = 0
    End If
    
    On Error GoTo ErrHand
    SQLQuery = "Select MAX(rdfCode) from rdf_Standard_Daypart"
    Set rdf_rst = gSQLSelectCall(SQLQuery)
    If IsNull(rdf_rst(0).Value) Then
        ReDim tgDaypartInfo(0 To 0) As DAYPARTINFO
        gPopDaypart = True
        Exit Function
    End If
    ilMax = rdf_rst(0).Value
    
    ilUpper = 0
    ReDim tgDaypartInfo(0 To ilMax) As DAYPARTINFO
    SQLQuery = "SELECT *"
    SQLQuery = SQLQuery + " FROM rdf_Standard_Daypart"
    Set rdf_rst = gSQLSelectCall(SQLQuery)
    While Not rdf_rst.EOF
        slStartTime = ""
        For ilLoop = 7 To 1 Step -1
            slStart = ""
            Select Case ilLoop
                Case 7
                    'If InStr(1, rdf_rst!rdfStartTime7, " ", vbBinaryCompare) > 1 Then
                    If (rdf_rst!rdfMo7 = "Y") Or (rdf_rst!rdfTu7 = "Y") Or (rdf_rst!rdfWe7 = "Y") Or (rdf_rst!rdfTh7 = "Y") Or (rdf_rst!rdfFr7 = "Y") Or (rdf_rst!rdfSa7 = "Y") Or (rdf_rst!rdfSu7 = "Y") Then
                        slStart = Format(rdf_rst!rdfStartTime7, sgShowTimeWSecForm)
                        slEnd = Format(rdf_rst!rdfEndTime7, sgShowTimeWSecForm)
                    End If
                Case 6
                    'If InStr(1, rdf_rst!rdfStartTime6, " ", vbBinaryCompare) > 1 Then
                    If (rdf_rst!rdfMo6 = "Y") Or (rdf_rst!rdfTu6 = "Y") Or (rdf_rst!rdfWe6 = "Y") Or (rdf_rst!rdfTh6 = "Y") Or (rdf_rst!rdfFr6 = "Y") Or (rdf_rst!rdfSa6 = "Y") Or (rdf_rst!rdfSu6 = "Y") Then
                        slStart = Format(rdf_rst!rdfStartTime6, sgShowTimeWSecForm)
                        slEnd = Format(rdf_rst!rdfEndTime6, sgShowTimeWSecForm)
                    End If
                Case 5
                    'If InStr(1, rdf_rst!rdfStartTime5, " ", vbBinaryCompare) > 1 Then
                    If (rdf_rst!rdfMo5 = "Y") Or (rdf_rst!rdfTu5 = "Y") Or (rdf_rst!rdfWe5 = "Y") Or (rdf_rst!rdfTh5 = "Y") Or (rdf_rst!rdfFr5 = "Y") Or (rdf_rst!rdfSa5 = "Y") Or (rdf_rst!rdfSu5 = "Y") Then
                        slStart = Format(rdf_rst!rdfStartTime5, sgShowTimeWSecForm)
                        slEnd = Format(rdf_rst!rdfEndTime5, sgShowTimeWSecForm)
                    End If
                Case 4
                    'If InStr(1, rdf_rst!rdfStartTime4, " ", vbBinaryCompare) > 1 Then
                    If (rdf_rst!rdfMo4 = "Y") Or (rdf_rst!rdfTu4 = "Y") Or (rdf_rst!rdfWe4 = "Y") Or (rdf_rst!rdfTh4 = "Y") Or (rdf_rst!rdfFr4 = "Y") Or (rdf_rst!rdfSa4 = "Y") Or (rdf_rst!rdfSu4 = "Y") Then
                        slStart = Format(rdf_rst!rdfStartTime4, sgShowTimeWSecForm)
                        slEnd = Format(rdf_rst!rdfEndTime4, sgShowTimeWSecForm)
                    End If
                Case 3
                    'If InStr(1, rdf_rst!rdfStartTime3, " ", vbBinaryCompare) > 1 Then
                    If (rdf_rst!rdfMo3 = "Y") Or (rdf_rst!rdfTu3 = "Y") Or (rdf_rst!rdfWe3 = "Y") Or (rdf_rst!rdfTh3 = "Y") Or (rdf_rst!rdfFr3 = "Y") Or (rdf_rst!rdfSa3 = "Y") Or (rdf_rst!rdfSu3 = "Y") Then
                        slStart = Format(rdf_rst!rdfStartTime3, sgShowTimeWSecForm)
                        slEnd = Format(rdf_rst!rdfEndTime3, sgShowTimeWSecForm)
                    End If
                Case 2
                    'If InStr(1, rdf_rst!rdfStartTime2, " ", vbBinaryCompare) > 1 Then
                    If (rdf_rst!rdfMo2 = "Y") Or (rdf_rst!rdfTu2 = "Y") Or (rdf_rst!rdfWe2 = "Y") Or (rdf_rst!rdfTh2 = "Y") Or (rdf_rst!rdfFr2 = "Y") Or (rdf_rst!rdfSa2 = "Y") Or (rdf_rst!rdfSu2 = "Y") Then
                        slStart = Format(rdf_rst!rdfStartTime2, sgShowTimeWSecForm)
                        slEnd = Format(rdf_rst!rdfEndTime2, sgShowTimeWSecForm)
                    End If
                Case 1
                    'If InStr(1, rdf_rst!rdfStartTime1, " ", vbBinaryCompare) > 1 Then
                    If (rdf_rst!rdfMo1 = "Y") Or (rdf_rst!rdfTu1 = "Y") Or (rdf_rst!rdfWe1 = "Y") Or (rdf_rst!rdfTh1 = "Y") Or (rdf_rst!rdfFr1 = "Y") Or (rdf_rst!rdfSa1 = "Y") Or (rdf_rst!rdfSu1 = "Y") Then
                        slStart = Format(rdf_rst!rdfStartTime1, sgShowTimeWSecForm)
                        slEnd = Format(rdf_rst!rdfEndTime1, sgShowTimeWSecForm)
                    End If
            End Select
            If slStart <> "" Then
                If slStartTime = "" Then
                    slStartTime = slStart
                    slEndTime = slEnd
                Else
                    If gTimeToLong(slStart, False) < gTimeToLong(slStartTime, False) Then
                        slStartTime = slStart
                    End If
                    If gTimeToLong(slEnd, True) > gTimeToLong(slEndTime, True) Then
                        slEndTime = slEnd
                    End If
                End If
            End If
        Next ilLoop
        If slStartTime = "" Then
            slStartTime = "12:00:00am"
            slEndTime = slStartTime
        End If
        tgDaypartInfo(ilUpper).iCode = rdf_rst!rdfCode
        tgDaypartInfo(ilUpper).sStartTime = slStartTime
        tgDaypartInfo(ilUpper).sEndTime = slEndTime
        ilUpper = ilUpper + 1
        rdf_rst.MoveNext
    Wend
    
    ReDim Preserve tgDaypartInfo(0 To ilUpper) As DAYPARTINFO
    
    'Now sort them by the vefCode
    If UBound(tgAdvtInfo) > 1 Then
        ArraySortTyp fnAV(tgDaypartInfo(), 0), UBound(tgDaypartInfo), 0, LenB(tgDaypartInfo(0)), 0, -1, 0
    End If
    
    rdf_rst.Close
    gPopDaypart = True
    Exit Function
gPopDaypartErr:
    ilRet = 1
    Resume Next
ErrHand:
    gHandleError "AffErrorLog.txt", "gPopDaypart"
    gPopDaypart = False
    Exit Function
End Function

Public Function gBinarySearchRdf(ilCode As Integer) As Integer
    
    'Returns the index number of tgDaypartInfo that matches the SntCode that was passed in
    'Note: for this to work tgDaypartInfo was previously be sorted by SntCode
    
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    
    On Error GoTo ErrHand
    
    ilMin = LBound(tgDaypartInfo)
    ilMax = UBound(tgDaypartInfo) - 1
    Do While ilMin <= ilMax
        ilMiddle = (CLng(ilMin) + ilMax) \ 2
        If ilCode = tgDaypartInfo(ilMiddle).iCode Then
            'found the match
            gBinarySearchRdf = ilMiddle
            Exit Function
        ElseIf ilCode < tgDaypartInfo(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    
    gBinarySearchRdf = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchRdf: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchRdf = -1
    Exit Function
    
End Function
'  This list must match popsubs.bas in traffic
Public Sub gInitTaskInfo()
    tgTaskInfo(0).sTaskCode = "CSS"
    tgTaskInfo(0).sTaskName = "Contract Spot Scheduler"
    tgTaskInfo(0).sSortCode = "A"
    tgTaskInfo(0).iMenuIndex = 0
    tgTaskInfo(1).sTaskCode = "SSB"
    tgTaskInfo(1).sTaskName = "Station Spot Builder"
    tgTaskInfo(1).sSortCode = "B"
    tgTaskInfo(1).iMenuIndex = 0
    tgTaskInfo(2).sTaskCode = "AEQ"
    tgTaskInfo(2).sTaskName = "Affiliate Export Queue"
    tgTaskInfo(2).sSortCode = "C"
    tgTaskInfo(2).iMenuIndex = 0
    tgTaskInfo(3).sTaskCode = "ASI"
    tgTaskInfo(3).sTaskName = "Affiliate Spot Import"
    tgTaskInfo(3).sSortCode = "D"
    tgTaskInfo(3).iMenuIndex = 0
    tgTaskInfo(4).sTaskCode = "AMB"
    tgTaskInfo(4).sTaskName = "Affiliate Measurement Builder"
    tgTaskInfo(4).sSortCode = "E"
    tgTaskInfo(4).iMenuIndex = 0
    tgTaskInfo(5).sTaskCode = "ARQ"
    tgTaskInfo(5).sTaskName = "Affiliate Report Queue"
    tgTaskInfo(5).sSortCode = "F"
    tgTaskInfo(5).iMenuIndex = 0
    tgTaskInfo(6).sTaskCode = "ASG"
    tgTaskInfo(6).sTaskName = "Avail Summary Generation"
    tgTaskInfo(6).sSortCode = "G"
    tgTaskInfo(6).iMenuIndex = 0
    tgTaskInfo(7).sTaskCode = "SC"
    tgTaskInfo(7).sTaskName = "Set Credit"
    tgTaskInfo(7).sSortCode = "H"
    tgTaskInfo(7).iMenuIndex = 0
    tgTaskInfo(8).sTaskCode = "CE"
    tgTaskInfo(8).sTaskName = "Corporate Export"
    tgTaskInfo(8).sSortCode = "I"
    tgTaskInfo(8).iMenuIndex = 0
    tgTaskInfo(9).sTaskCode = "SFE"
    tgTaskInfo(9).sTaskName = "Sales Force Export"
    tgTaskInfo(9).sSortCode = "J"
    tgTaskInfo(9).iMenuIndex = 0
    tgTaskInfo(10).sTaskCode = "ME"
    tgTaskInfo(10).sTaskName = "Matrix Export"
    tgTaskInfo(10).sSortCode = "K"
    tgTaskInfo(10).iMenuIndex = 0
    tgTaskInfo(11).sTaskCode = "EPE"
    tgTaskInfo(11).sTaskName = "Efficio Projection Export"
    tgTaskInfo(11).sSortCode = "L"
    tgTaskInfo(11).iMenuIndex = 0
    tgTaskInfo(12).sTaskCode = "ERE"
    tgTaskInfo(12).sTaskName = "Efficio Revenue Export"
    tgTaskInfo(12).sSortCode = "M"
    tgTaskInfo(12).iMenuIndex = 0
    tgTaskInfo(13).sTaskCode = "GPE"
    tgTaskInfo(13).sTaskName = "Get Paid Export"
    tgTaskInfo(13).sSortCode = "N"
    tgTaskInfo(13).iMenuIndex = 0
    tgTaskInfo(14).sTaskCode = "BD"
    tgTaskInfo(14).sTaskName = "Backup Data"
    tgTaskInfo(14).sSortCode = "O"
    tgTaskInfo(14).iMenuIndex = 0
    tgTaskInfo(15).sTaskCode = "TE"
    tgTaskInfo(15).sTaskName = "Tableau Export"
    tgTaskInfo(15).sSortCode = "P"
    tgTaskInfo(15).iMenuIndex = 0
'    '7967 on/off here
    tgTaskInfo(16).sTaskCode = "WVM"
    tgTaskInfo(16).sTaskName = "Web Vendor Manager"
    tgTaskInfo(16).sSortCode = "Q"
    tgTaskInfo(16).iMenuIndex = 0
    'D.S. 03/21/18
    tgTaskInfo(17).sTaskCode = "PB"
    tgTaskInfo(17).sTaskName = "Programmatic Buy"
    tgTaskInfo(17).sSortCode = "R"
    tgTaskInfo(17).iMenuIndex = 0
    'D.S. 10/9/19
    tgTaskInfo(18).sTaskCode = "CAI"
    tgTaskInfo(18).sTaskName = "Compel Auto Import"
    tgTaskInfo(18).sSortCode = "S"
    tgTaskInfo(18).iMenuIndex = 0

    '1-29-20 Add RAB CRM export
    tgTaskInfo(19).sTaskCode = "RE"
    tgTaskInfo(19).sTaskName = "RAB Export"
    tgTaskInfo(19).sSortCode = "T"
    tgTaskInfo(19).iMenuIndex = 0

    'TTP 9992 Add Custom Revenue Export
    tgTaskInfo(20).sTaskCode = "CRE"
    tgTaskInfo(20).sTaskName = "Custom Revenue Export"
    tgTaskInfo(20).sSortCode = "U"
    tgTaskInfo(20).iMenuIndex = 0
End Sub
Public Function gBinarySearchRepInfo(ilCode As Integer, tlRepInfo() As REPINFO) As Integer
   'Returns the index number of tgMarketRepInfo or tgServiceRepInfo that matches the shttmktrepustcode/shttservustcode that was passed in
    
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    
    On Error GoTo ErrHand
    
    ilMin = LBound(tlRepInfo)
    ilMax = UBound(tlRepInfo) - 1
    Do While ilMin <= ilMax
        ilMiddle = (CLng(ilMin) + ilMax) \ 2
        If ilCode = tlRepInfo(ilMiddle).iUstCode Then
            'found the match
            gBinarySearchRepInfo = ilMiddle
            Exit Function
        ElseIf ilCode < tlRepInfo(ilMiddle).iUstCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    
    gBinarySearchRepInfo = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchRepInfo: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchRepInfo = -1
    Exit Function
End Function
Public Function gBinarySearchOwner(llCode As Long) As Long
    'Returns the index number of tgOwnerInfo that matches the shttarttcode that was passed in
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    On Error GoTo ErrHand
    
    llMin = LBound(tgOwnerInfo)
    llMax = UBound(tgOwnerInfo) - 1
    Do While llMin <= llMax
        llMiddle = (CLng(llMin) + llMax) \ 2
        If llCode = tgOwnerInfo(llMiddle).lCode Then
            'found the match
            gBinarySearchOwner = llMiddle
            Exit Function
        ElseIf llCode < tgOwnerInfo(llMiddle).lCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    
    gBinarySearchOwner = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchOwner: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    gBinarySearchOwner = -1
    Exit Function
End Function

Public Function gFileChgd(slFileName As String) As Boolean
    Dim slSQLQuery As String
    Dim ilIndex As Integer
    
    ilIndex = -1
    gFileChgd = True
    slSQLQuery = "Select * From fct_File_Chg_Table Where fctFileName = '" & slFileName & "'"
    Set fct_rst = gSQLSelectCall(slSQLQuery)
    If Not fct_rst.EOF Then
        Select Case UCase(slFileName)
            Case "SHTT.MKD"
                ilIndex = SHTTINDEX
            Case "VEF.BTR"
                ilIndex = VEFINDEX
            Case "VPF.BTR"
                ilIndex = VPFINDEX
            Case "CPTT.MKD"
                ilIndex = CPTTINDEX
            Case "ADF.BTR"
                ilIndex = ADFINDEX
        End Select
        If ilIndex <> -1 Then
            If (tgFctChgdInfo(ilIndex).lLastDateChgd <> fct_rst!fctDateChgd) Or (tgFctChgdInfo(ilIndex).lLastTimeChgd <> fct_rst!fctTimeChgd) Then
                gFileChgdSetInternal slFileName
            Else
                gFileChgd = False
            End If
        End If
    Else
        gFileChgdUpdate slFileName, False
    End If
    
End Function
Public Sub gFileChgdSetInternal(slFileName As String)
    Dim slSQLQuery As String
    Dim ilIndex As Integer
    
    ilIndex = -1
    slSQLQuery = "Select * From fct_File_Chg_Table Where  fctFileName = '" & slFileName & "'"
    Set fct_rst = gSQLSelectCall(slSQLQuery)
    If Not fct_rst.EOF Then
        Select Case UCase(slFileName)
            Case "SHTT.MKD"
                ilIndex = SHTTINDEX
            Case "VEF.BTR"
                ilIndex = VEFINDEX
            Case "VPF.BTR"
                ilIndex = VPFINDEX
            Case "CPTT.MKD"
                ilIndex = CPTTINDEX
             Case "ADF.BTR"
                ilIndex = ADFINDEX
        End Select
        If ilIndex <> -1 Then
            tgFctChgdInfo(ilIndex).lLastDateChgd = fct_rst!fctDateChgd
            tgFctChgdInfo(ilIndex).lLastTimeChgd = fct_rst!fctTimeChgd
        End If
    End If
End Sub

Public Sub gFileChgdUpdate(slFileName As String, blRepopRequired As Boolean)
    Dim slNow As String
    Dim llRet As Long
    Dim slSQLQuery As String
    
    slNow = Now
    slSQLQuery = "Select * From fct_File_Chg_Table Where  fctFileName = '" & slFileName & "'"
    Set fct_rst = gSQLSelectCall(slSQLQuery)
    If Not fct_rst.EOF Then
        slSQLQuery = "UPDATE fct_File_Chg_Table SET "
        slSQLQuery = slSQLQuery & "fctDate = '" & Format$(slNow, sgSQLDateForm) & "', "
        slSQLQuery = slSQLQuery & "fctDateChgd = " & gDateValue(slNow) & ", "
        slSQLQuery = slSQLQuery & "fctTime = '" & Format$(slNow, sgSQLTimeForm) & "', "
        slSQLQuery = slSQLQuery & "fctTimeChgd = " & gTimeToLong(slNow, False) & ", "
        slSQLQuery = slSQLQuery & "fctSystemType = '" & "A" & "', "
        slSQLQuery = slSQLQuery & "fctUstCode = " & igUstCode
        slSQLQuery = slSQLQuery & " WHERE (fctFileName = '" & slFileName & "')"
        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    Else
        slSQLQuery = "Insert Into FCT_File_Chg_Table ( "
        slSQLQuery = slSQLQuery & "fctCode, "
        slSQLQuery = slSQLQuery & "fctFileName, "
        slSQLQuery = slSQLQuery & "fctDate, "
        slSQLQuery = slSQLQuery & "fctDateChgd, "
        slSQLQuery = slSQLQuery & "fctTime, "
        slSQLQuery = slSQLQuery & "fctTimeChgd, "
        slSQLQuery = slSQLQuery & "fctSystemType, "
        slSQLQuery = slSQLQuery & "fctUrfCode, "
        slSQLQuery = slSQLQuery & "fctUstCode, "
        slSQLQuery = slSQLQuery & "fctUnused "
        slSQLQuery = slSQLQuery & ") "
        slSQLQuery = slSQLQuery & "Values ( "
        slSQLQuery = slSQLQuery & 0 & ", "
        slSQLQuery = slSQLQuery & "'" & gFixQuote(slFileName) & "', "
        slSQLQuery = slSQLQuery & "'" & Format$(slNow, sgSQLDateForm) & "', "
        slSQLQuery = slSQLQuery & gDateValue(slNow) & ", "
        slSQLQuery = slSQLQuery & "'" & Format$(slNow, sgSQLTimeForm) & "', "
        slSQLQuery = slSQLQuery & gTimeToLong(slNow, False) & ", "
        slSQLQuery = slSQLQuery & "'" & "A" & "', "
        slSQLQuery = slSQLQuery & 0 & ", "
        slSQLQuery = slSQLQuery & igUstCode & ", "
        slSQLQuery = slSQLQuery & "'" & "" & "' "
        slSQLQuery = slSQLQuery & ") "
        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    End If
    If llRet = 0 And Not blRepopRequired Then
        gFileChgdSetInternal slFileName
    End If
End Sub

