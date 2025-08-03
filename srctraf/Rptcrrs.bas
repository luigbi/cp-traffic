Attribute VB_Name = "RPTCRRS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrrs.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Variables (Removed)                                                             *
'*  smBNCodeTag                                                                           *
'******************************************************************************************
Option Explicit
Option Compare Text
' Dan M 1/20/11 rewrote research reports so that custom headings will be alphabetized.
' mnf heading goes with drf.llDemo(mnf.groupNumber).  There are 18 llDemos per drf record. Drf.DataType is the 'set number': "A" is first set of 18, "B" is 2nd, "d" is 3rd.
' Here is how I alphabetized the headings stored in mnf--for this example, the sets will be of 2 units and not 18.
' example: drf              1       2
'          drfType          "B"     "d"
'          Values           41 60   72 93
'          drfDemo          1  2    1  2
'          mnf              "C "B"  "D""A"
'          mnfGroupNumber   1  2    3  4
' Place mnf Names in listbox to be sorted: store the GroupNumber with Names
'          List Box
'           A       4
'           B       2
'           C       1
'           D       3
' Dimension llResults to the same size as the list box.
' I grab drf, use the type and demo index to stuff data into llResults.
' drf 'type' as number + drfDemo - 1 = llResult index
        'drf.llDemo(1) of type "B" goes to llResults(0); drf.llDemo(1) of type "d" goes to llResults(3)
' Then, when the drf's are all written in, I write to Rsr by going through listbox, getting itemdata to find where the data is stored in llResults.
'               listbox.item(3) = "C"  listbox.itemData(3) = 1 llResult uses this number -1, so grab: llResult(0) =  41
' each llResult has to be created independently-- for each vehicle, I have to grab "Vehicle" results in a different loop from "Daypart" results; "Dayparts" also
' must be run independently for each rdfCode.

Dim hmVef As Integer
Dim tmVef As VEF
Dim imVefRecLen As Integer
Dim hmDrf As Integer           'Demo Research Data file handle
Dim imDrfRecLen As Integer     'DRF record length
Dim tmDrfForBook() As DRF
Dim hmDnf As Integer           'Demo name file handle
Dim imDnfRecLen As Integer     'DRF record length
Dim tmDnf() As DNF
Dim tmDef As DEF                'demo estimates
Dim tmDefSrchKey1 As DEFKEY1    'Def key:  dnfcode & start date
Dim hmDef As Integer
Dim imDefRecLen As Integer
Dim tmDpf As DPF                'demo estimates
Dim hmDpf As Integer
Dim imDpfRecLen As Integer
Dim tmDpfSrchKey1 As DPFKEY1
Dim hmRsr As Integer           'Research Data file handle
Dim imRsrRecLen As Integer     'RSR record length
Dim tmRsr As RSR
Dim tmGrf As GRF                'prepass temporary file for Special Research summary report
Dim hmGrf As Integer
Dim imGrfRecLen As Integer
Dim tmMnf As MNF
Dim hmMnf As Integer
Dim imMnfRecLen As Integer
Dim tmRaf As RAF
Dim hmRaf As Integer
Dim imRafRecLen As Integer
Dim imListIndex As Integer
Dim bmIsItImpressions As Boolean
Dim tmDemoRankInfo() As DEMORANK_INFO
Type DEMORANK_INFO
    sKey As String * 10             'aud value key to sort
    iVefCode As Integer
    iMnfDemo(0 To 4) As Integer     'mnf demo reference- max 5 demos allowed
    lAvgAud(0 To 4) As Long         'avg aud
    iRank(0 To 4) As Integer
End Type
Dim tmDemoRank() As DEMORANK_INFO

'*******************************************************
'*                                                     *
'*      Procedure Name:mCloseBtrFiles                  *
'*                                                     *
'*            Created:12/5/00        By:D. Smith       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Close the BTR Files            *
'*                                                     *
'*******************************************************
Sub mCloseBtrFiles(ilListIndex As Integer)
    Dim ilRet As Integer

    ilRet = btrClose(hmDrf)
    btrDestroy hmDrf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    'Dan 01/24/2011 add close to special research report
     If ilListIndex = RS_SUMMARY Then
        ilRet = btrClose(hmRsr)
        btrDestroy hmRsr
        ilRet = btrClose(hmDnf)
        btrDestroy hmDnf
        ilRet = btrClose(hmDpf)
        btrDestroy hmDpf
     Else            'special research summary
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        ilRet = btrClose(hmDef)
        btrDestroy hmDef
     End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mLoadCustomDemos                *
'*                                                     *
'*            Created:12/5/00        By:D. Smith       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Load an array of 16 buckets    *
'*            with whatever custom demos that may be   *
'*            defined.                                 *
'*                                                     *
'*******************************************************
Sub mLoadCustomDemos(smCustDemos As ListBox)
    Dim ilIndex As Integer
    Dim ilRet As Integer
    
    'Load custom day parts into listbox to sort
    ilRet = btrGetFirst(hmMnf, tmMnf, imMnfRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record
    Do While ilRet = BTRV_ERR_NONE
        If tmMnf.iGroupNo <> 0 And tmMnf.sType = "D" Then
           smCustDemos.AddItem tmMnf.sName
           smCustDemos.ItemData(smCustDemos.NewIndex) = tmMnf.iGroupNo
        End If
        ilRet = btrGetNext(hmMnf, tmMnf, imMnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
    Loop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mOKtoSaveRec                    *
'*                                                     *
'*            Created:12/18/00       By:D. Smith       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if the research rec  *
'*            needs to saved based off user selections *
'*                                                     *
'*******************************************************
Function mOKtoSaveRec(tlDrf As DRF) As Integer
    'If the record is a Socio Eco rec and Socio Eco not selected then we are out of here
    If Not (RptSelRS!ckcSelC3(0).Value = vbChecked) And tlDrf.iMnfSocEco > 0 Then
        Exit Function
    End If
    
    'sDemoDataType "P" = Population Demo, D" = Imported Demos, "M" = Manually Defined Demos
    If tlDrf.sDemoDataType = "P" And tlDrf.iMnfSocEco <> 0 Then
        mOKtoSaveRec = False
        Exit Function
    End If
    mOKtoSaveRec = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenBtrFiles                   *
'*
'*      <input> ilListIndex - index of report selected
'*            Created:12/5/00        By:D. Smith       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Open the BTR Files             *
'*                                                     *
'*******************************************************
Sub mOpenBtrFiles(ilListIndex As Integer)
    Dim ilRet As Integer
    Dim tlDrf As DRF
    Dim tlRsr As RSR
    
    'Demo Research Data File
    hmDrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmDrf)
        btrDestroy hmDrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imDrfRecLen = Len(tlDrf)

    'Vehicle File
    hmVef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmDrf)
        btrDestroy hmDrf
        btrDestroy hmVef
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)

    'Multi-Name File
    hmMnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmDrf)
        btrDestroy hmDrf
        btrDestroy hmVef
        btrDestroy hmMnf
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)

    If ilListIndex = RS_SUMMARY Then            'research summary
        'Research Data File - Temporary file that's passed to Crystal Reports
        hmRsr = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmRsr, "", sgDBPath & "Rsr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmDrf)
            ilRet = btrClose(hmRsr)
            ilRet = btrClose(hmMnf)
            ilRet = btrClose(hmVef)
            btrDestroy hmDrf
            btrDestroy hmVef
            btrDestroy hmMnf
            btrDestroy hmRsr
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        imRsrRecLen = Len(tlRsr)
        
        ReDim tmDnf(0 To 0) As DNF
        hmDnf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmDnf)
            ilRet = btrClose(hmDrf)
            ilRet = btrClose(hmRsr)
            ilRet = btrClose(hmMnf)
            ilRet = btrClose(hmVef)
            btrDestroy hmDnf
            btrDestroy hmDrf
            btrDestroy hmVef
            btrDestroy hmMnf
            btrDestroy hmRsr
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        imDnfRecLen = Len(tmDnf(0))
        
        hmDpf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmDpf)
            ilRet = btrClose(hmDnf)
            ilRet = btrClose(hmDrf)
            ilRet = btrClose(hmRsr)
            ilRet = btrClose(hmMnf)
            ilRet = btrClose(hmVef)
            btrDestroy hmDpf
            btrDestroy hmDnf
            btrDestroy hmDrf
            btrDestroy hmVef
            btrDestroy hmMnf
            btrDestroy hmRsr
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        imDpfRecLen = Len(tmDpf)
    Else
        'Research Data File - Temporary file that's passed to Crystal Reports
        hmGrf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmDrf)
            ilRet = btrClose(hmMnf)
            ilRet = btrClose(hmVef)
            btrDestroy hmGrf
            btrDestroy hmDrf
            btrDestroy hmVef
            btrDestroy hmMnf
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        imGrfRecLen = Len(tmGrf)

        'Demo Estimates File - Temporary file that's passed to Crystal Reports
        hmDef = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmDef)
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmDrf)
            ilRet = btrClose(hmMnf)
            ilRet = btrClose(hmVef)
            btrDestroy hmDef
            btrDestroy hmGrf
            btrDestroy hmDrf
            btrDestroy hmVef
            btrDestroy hmMnf
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        imDefRecLen = Len(tmDef)
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gResearchReport_new             *
'*                                                     *
'*            Created:12/5/00        By:J. White       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Pre-pass for Reseach Report    *
'*                                                     *
'*******************************************************
Sub gResearchReport_New(LbcCustDemos As ListBox)
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilSortCode As Integer
    Dim ilBookName As Integer
    Dim blBookName As Boolean
    Dim ilDnfCode As Integer
    Dim c As Integer
    Dim ilVefCode As Integer
    
    imListIndex = RptSelRS!lbcRptType.ListIndex
    mOpenBtrFiles imListIndex     'Open all necessary BTR files
    
    If (RptSelRS!ckcSelC3(13).Value = vbChecked) Then
        mLoadCustomDemos LbcCustDemos 'Load any custom demos that may have been defined
    End If
    
    'this will go through the book names
    If RptSelRS!rbcSelC11(0).Value = True Then
        'Selected Book(s) from UI
         ilBookName = RptSelRS!lbcSelection(1).ListCount - 1
    End If
    If RptSelRS!rbcSelC11(1).Value = True Then
        'Vehicle Default
        ilBookName = 0
    End If
    RptSelRS.MousePointer = vbHourglass
    'TTP 10667 - Research Report: vehicle by default book option showing duplicate records
    'For c = 0 To RptSelRS!lbcSelection(1).ListCount - 1 Step 1
    For c = 0 To ilBookName Step 1
        'get Selected Book ID.  If 'Vehicle by Default Book' option then Book ID = 0
        ilDnfCode = -1
        If RptSelRS!rbcSelC11(0).Value = True Then 'use the selected Book Name(s)
            If RptSelRS!lbcSelection(1).Selected(c) Then
                slNameCode = tgBookNameCode(c).sKey
                gParseItem slNameCode, 2, "\", slCode
                ilDnfCode = Val(slCode)
            End If
        End If
        If RptSelRS!rbcSelC11(1).Value = True Then 'use the Vehicle Default book
            ilDnfCode = 0
        End If
        
        'go through vehicles
        If ilDnfCode >= 0 Then
            For illoop = 0 To RptSelRS!lbcSelection(0).ListCount - 1 Step 1
                If RptSelRS!lbcSelection(0).Selected(illoop) Then
                    slNameCode = tgVehicle(illoop).sKey
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilVefCode = Val(slCode)
                    If RptSelRS!rbcSelC11(1).Value = True Then 'use the Vehicle Default book
                        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, ilVefCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        ilDnfCode = tmVef.iDnfCode
                    End If
'Debug.Print "gResearchReport_New; vef:" & ilVefCode & ", DnfCode=" & ilDnfCode
                    
                    If ilDnfCode > 0 Then
                        bmIsItImpressions = mIsItImpressions(ilDnfCode)
                        '10/6/22 - JW - Fix per Jason email: v81 Research report - TTP 10556 (Issue 1)
                        If bmIsItImpressions Then
                            If RptSelRS!ckcSelC5(0).Value = vbChecked Then      'Sold Daypart
                                mGetDpfRecords ilDnfCode, ilVefCode             'Get Impressions research
                            End If
                        Else
                            mGetDrfRecords ilDnfCode, 0, "POP"                  'Get Population
                            If RptSelRS!ckcSelC5(0).Value = vbChecked Then      'Sold Daypart
                                mGetDrfRecords ilDnfCode, ilVefCode, "DAYPART"  'Get Research by Daypart
                            End If
                            If RptSelRS!ckcSelC5(1).Value = vbChecked Then      'Extra Daypart
                                mGetDrfRecords ilDnfCode, ilVefCode, "EXTRA"    'Get Research (Extra: data without a matching daypart)
                            End If
                            If RptSelRS!ckcSelC5(2).Value = vbChecked Then      'Time
                                mGetDrfRecords ilDnfCode, ilVefCode, "TIME"     'Get Research by Time
                            End If
                            If RptSelRS!ckcSelC5(3).Value = vbChecked Then      'Vehicle
                                mGetDrfRecords ilDnfCode, ilVefCode, "VEHICLE"  'Get Research by Vehicle
                            End If
                        End If
                    End If
                End If
            Next illoop
        End If
    Next c
    RptSelRS.MousePointer = vbDefault
    LbcCustDemos.Clear
    mCloseBtrFiles imListIndex
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gResearchReport                 *
'*                                                     *
'*            Created:12/5/00        By:D. Smith       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Pre-pass for Reseach Report    *
'*                                                     *
'*******************************************************
Sub gResearchReport(LbcCustDemos As ListBox)
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilSortCode As Integer
    Dim ilBookName As Integer
    Dim blBookName As Boolean
    Dim ilDnfCode As Integer
    Dim c As Integer
    Dim ilVefCode As Integer

    imListIndex = RptSelRS!lbcRptType.ListIndex
    mOpenBtrFiles imListIndex     'Open all necessary BTR files
    mLoadCustomDemos LbcCustDemos 'Load any custom demos that may have been defined

    'Sort by Book Name
    If RptSelRS!rbcSelC11(0).Value Then
        ilBookName = RptSelRS!lbcSelection(1).ListCount - 1
        blBookName = True
    Else
        ilBookName = 0
    End If

    'this will go through one time for default book names, essentially skipping the book names
    For c = 0 To ilBookName Step 1
        If blBookName Then
            'book names list box
            If RptSelRS!lbcSelection(1).Selected(c) Then
                slNameCode = tgBookNameCode(c).sKey
                gParseItem slNameCode, 2, "\", slCode
                ilDnfCode = Val(slCode)
            Else
                ilDnfCode = -1
            End If
        'get default book
        Else
            ilDnfCode = 0
        End If
        ' go through vehicles, only skipping if wanted book name and it wasn't selected
        If ilDnfCode >= 0 Then
            For illoop = 0 To RptSelRS!lbcSelection(0).ListCount - 1 Step 1
                If RptSelRS!lbcSelection(0).Selected(illoop) Then
                    slNameCode = tgVehicle(illoop).sKey
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilVefCode = Val(slCode)
                    If Not blBookName Then
                        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, ilVefCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        ilDnfCode = tmVef.iDnfCode
                    End If
                    'ilDnfCode from bookName or default, as needed
                    If ilDnfCode > 0 Then
                        bmIsItImpressions = mIsItImpressions(ilDnfCode)
                        mDrfToRsr ilDnfCode, ilVefCode, LbcCustDemos
                    End If
                End If
            Next illoop
        End If
    Next c
    LbcCustDemos.Clear
    mCloseBtrFiles imListIndex
End Sub

Private Sub mGetDrfRecords(ilDnfCode As Integer, ilVefCode As Integer, sResearchType As String)
    'DNF = "DNF_Demo_Rsrch_Names"
    Dim slSQLQuery As String
    Dim rst_Temp As ADODB.Recordset
    Dim tlRsr As RSR
    Dim slForm As String
    Dim illoop As Integer
    Dim ilRet As Integer
    Dim llDemoVal As Long
    Debug.Print " - Get DrfRecords - Book:" & ilDnfCode & " , Type:" & sResearchType & ", Vehicle:" & ilVefCode
    
    '---------------------------------
    'Build DRF Query
    slSQLQuery = "SELECT "
    If sResearchType = "POP" Then
        slSQLQuery = slSQLQuery & " TOP 1 "
    End If
    slSQLQuery = slSQLQuery & "    drfCode, drfdnfCode, drfDemoDataType, drfmnfSocEco, drfvefCode, drfInfotype, drfrdfCode, drfStartTime, drfEndTime, "
    slSQLQuery = slSQLQuery & "    drfDemo1, drfDemo2, drfDemo3, drfDemo4, drfDemo5, drfDemo6, drfDemo7, drfDemo8, drfDemo9, "
    slSQLQuery = slSQLQuery & "    drfDemo10, drfDemo11, drfDemo12, drfDemo13, drfDemo14, drfDemo15, drfDemo16, drfDemo17, drfDemo18, drfForm "
    slSQLQuery = slSQLQuery & " FROM DRF_Demo_Rsrch_Data "
    slSQLQuery = slSQLQuery & " WHERE "
    slSQLQuery = slSQLQuery & "     drfDnfCode = " & ilDnfCode 'Book
    
    If RptSelRS.ckcSelC3(0).Value = False Or sResearchType = "POP" Then 'Include Qualitative?
        slSQLQuery = slSQLQuery & "     and drfmnfSocEco = 0"
    End If
    
    Select Case sResearchType
        Case "POP"
            slSQLQuery = slSQLQuery & "     and drfDemoDataType = 'P'"  'P=Population Data
            slSQLQuery = slSQLQuery & "     and drfrdfCode = 0"
            slSQLQuery = slSQLQuery & "     and drfvefcode = 0"
            slSQLQuery = slSQLQuery & "     and drfInfoType = ''"
        
        Case "DAYPART"
            slSQLQuery = slSQLQuery & "     and drfDemoDataType = 'D'"  'D=Imported Demos
            slSQLQuery = slSQLQuery & "     and drfInfoType = 'D'"      'D=Daypart Data
            slSQLQuery = slSQLQuery & "     and drfrdfCode > 0"
            slSQLQuery = slSQLQuery & "     and drfvefcode = " & ilVefCode

        Case "EXTRA"
            slSQLQuery = slSQLQuery & "     and drfDemoDataType = 'D'"  'D=Imported Demos
            slSQLQuery = slSQLQuery & "     and drfInfoType = 'D'"      'D=Daypart Data
            slSQLQuery = slSQLQuery & "     and drfrdfCode = 0"
            slSQLQuery = slSQLQuery & "     and drfvefcode = " & ilVefCode
        
        Case "TIME"
            slSQLQuery = slSQLQuery & "     and drfDemoDataType = 'D'"  'D=Imported Demos
            slSQLQuery = slSQLQuery & "     and drfInfoType = 'T'"      'T=Time Data
            slSQLQuery = slSQLQuery & "     and drfrdfCode = 0"
            slSQLQuery = slSQLQuery & "     and drfvefcode = " & ilVefCode
        
        Case "VEHICLE"
            slSQLQuery = slSQLQuery & "     and drfDemoDataType = 'D'"  'D=Imported Demos
            slSQLQuery = slSQLQuery & "     and drfInfoType = 'V'"      'V=Vehicle Data
            slSQLQuery = slSQLQuery & "     and drfrdfCode = 0"
            slSQLQuery = slSQLQuery & "     and drfvefcode = " & ilVefCode
    End Select
    If (RptSelRS!ckcSelC3(13).Value = vbChecked) And (RptSelRS!ckcSelC3(14).Value = vbUnchecked) Then   'Custom Demos
        slSQLQuery = slSQLQuery & "     and drfDataType <> 'A' "
        'drfDataType ; custom and not A  OR  not custom and A
    Else
        slSQLQuery = slSQLQuery & "     and drfDataType = 'A' "
        'drfDataType ; custom and not A  OR  not custom and A
    End If
    
    slSQLQuery = slSQLQuery & " ORDER BY "
    slSQLQuery = slSQLQuery & "     case when drfDay1='Y' then '0' else '1' end +"
    slSQLQuery = slSQLQuery & "     case when drfDay2='Y' then '0' else '1' end +"
    slSQLQuery = slSQLQuery & "     case when drfDay3='Y' then '0' else '1' end +"
    slSQLQuery = slSQLQuery & "     case when drfDay4='Y' then '0' else '1' end +"
    slSQLQuery = slSQLQuery & "     case when drfDay5='Y' then '0' else '1' end +"
    slSQLQuery = slSQLQuery & "     case when drfDay6='Y' then '0' else '1' end +"
    slSQLQuery = slSQLQuery & "     case when drfDay7='Y' then '0' else '1' end ,"
    slSQLQuery = slSQLQuery & "     hour(drfStartTime), hour(drfEndTime), drfrdfCode, drfmnfSocEco "
    
    Set rst_Temp = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Temp.EOF
        '---------------------------------
        'Write RSR
        tlRsr.iGenDate(0) = igNowDate(0)    'Date stamp for Crystal to Key off of
        tlRsr.iGenDate(1) = igNowDate(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tlRsr.lGenTime = lgNowTime          'Time stamp for Crystal to Key off of
        tlRsr.lDrfCode = rst_Temp!drfCode   'DRF Code
        
        '10/6/22 - JW - Fix per Jason email: v81 Research report - TTP 10556 (Issue 2)
        'If tlRsr.sForm = "8" Then
        '10/13/22 - JW - Fix per Jason email:RE: v81 Research report - TTP 10556
        If rst_Temp!drfForm = "8" Then
            slForm = "8"
        Else
            slForm = ""
        End If
        Select Case sResearchType
            Case "POP"
                tlRsr.sDataType = "P"   'D = Sold DP, D = Extra DP, T = Time = T, V = Vehicle
                tlRsr.iPopAud = 0       'P = Population, Otherwise Audience
            Case "DAYPART"
                tlRsr.iPopAud = 1       'P = Population, Otherwise Audience
                tlRsr.sDataType = "D"   'D = Sold DP, D = Extra DP, T = Time = T, V = Vehicle
            Case "EXTRA"
                tlRsr.iPopAud = 1       'P = Population, Otherwise Audience
                tlRsr.sDataType = "D"   'D = Sold DP, D = Extra DP, T = Time = T, V = Vehicle
            Case "TIME"
                tlRsr.iPopAud = 1       'P = Population, Otherwise Audience
                tlRsr.sDataType = "T"   'D = Sold DP, D = Extra DP, T = Time = T, V = Vehicle
            Case "VEHICLE"
                tlRsr.iPopAud = 1       'P = Population, Otherwise Audience
                tlRsr.sDataType = "V"   'D = Sold DP, D = Extra DP, T = Time = T, V = Vehicle
        End Select
        
        
        '---------------------------------
        'Report Headers
        If (RptSelRS!ckcSelC3(13).Value = vbChecked) And (RptSelRS!ckcSelC3(14).Value = vbUnchecked) Then   'Custom Demos
            tlRsr.sDemoType = "B"                   'A = Standard Demo B = Custom Demo - for sorting
            For illoop = 0 To 17
                If RptSelRS!LbcCustDemos.ListCount >= illoop Then
                    tlRsr.sDemoDesc(illoop) = Trim(RptSelRS!LbcCustDemos.List(illoop))
                Else
                    tlRsr.sDemoDesc(illoop) = ""
                End If
            Next illoop
        Else
            tlRsr.sDemoType = "A"                   'A = Standard Demo B = Custom Demo - for sorting
            If slForm <> "8" Then                           '16 bucket Research
                tlRsr.sDemoDesc(0) = "M12-17"
                tlRsr.sDemoDesc(1) = "M18-24"
                tlRsr.sDemoDesc(2) = "M25-34"
                tlRsr.sDemoDesc(3) = "M35-44"
                tlRsr.sDemoDesc(4) = "M45-49"
                tlRsr.sDemoDesc(5) = "M50-54"
                tlRsr.sDemoDesc(6) = "M55-64"
                tlRsr.sDemoDesc(7) = "M65+"
                tlRsr.sDemoDesc(8) = ""
                tlRsr.sDemoDesc(9) = "W12-17"
                tlRsr.sDemoDesc(10) = "W18-24"
                tlRsr.sDemoDesc(11) = "W25-34"
                tlRsr.sDemoDesc(12) = "W35-44"
                tlRsr.sDemoDesc(13) = "W45-49"
                tlRsr.sDemoDesc(14) = "W50-54"
                tlRsr.sDemoDesc(15) = "W55-64"
                tlRsr.sDemoDesc(16) = "W65+"
                tlRsr.sDemoDesc(17) = ""
            Else
                If bmIsItImpressions Then                   'Impressions?
                    tlRsr.sDemoDesc(0) = "P12+"
                    For illoop = 1 To 17
                        tlRsr.sDemoDesc(illoop) = ""
                    Next illoop
                Else
                    tlRsr.sDemoDesc(0) = "M12-17"           '18 bucket Research
                    tlRsr.sDemoDesc(1) = "M18-20"
                    tlRsr.sDemoDesc(2) = "M21-24"
                    tlRsr.sDemoDesc(3) = "M25-34"
                    tlRsr.sDemoDesc(4) = "M35-44"
                    tlRsr.sDemoDesc(5) = "M45-49"
                    tlRsr.sDemoDesc(6) = "M50-54"
                    tlRsr.sDemoDesc(7) = "M55-64"
                    tlRsr.sDemoDesc(8) = "M65+"
                    tlRsr.sDemoDesc(9) = "W12-17"
                    tlRsr.sDemoDesc(10) = "W18-20"
                    tlRsr.sDemoDesc(11) = "W21-24"
                    tlRsr.sDemoDesc(12) = "W25-34"
                    tlRsr.sDemoDesc(13) = "W35-44"
                    tlRsr.sDemoDesc(14) = "W45-49"
                    tlRsr.sDemoDesc(15) = "W50-54"
                    tlRsr.sDemoDesc(16) = "W55-64"
                    tlRsr.sDemoDesc(17) = "W65+"
                End If
            End If
        End If
        '---------------------------------
        'Research Data
        '10/6/22 - JW - Fix per Jason email: v81 Research report - TTP 10556 (Issue 3)
        'Custom Demo uses mnfGroup# to determine which column the values belong in
        If (RptSelRS!ckcSelC3(13).Value = vbChecked) And (RptSelRS!ckcSelC3(14).Value = vbUnchecked) Then   'Custom Demos
            tlRsr.sDemoType = "B"                   'A = Standard Demo B = Custom Demo - for sorting
            For illoop = 0 To RptSelRS!LbcCustDemos.ListCount - 1
                Select Case illoop
                    Case 0: llDemoVal = rst_Temp!drfDemo1
                    Case 1: llDemoVal = rst_Temp!drfDemo2
                    Case 2: llDemoVal = rst_Temp!drfDemo3
                    Case 3: llDemoVal = rst_Temp!drfDemo4
                    Case 4: llDemoVal = rst_Temp!drfDemo5
                    Case 5: llDemoVal = rst_Temp!drfDemo6
                    Case 6: llDemoVal = rst_Temp!drfDemo7
                    Case 7: llDemoVal = rst_Temp!drfDemo8
                    Case 8: llDemoVal = rst_Temp!drfDemo9
                    Case 9: llDemoVal = rst_Temp!drfDemo10
                    Case 10: llDemoVal = rst_Temp!drfDemo11
                    Case 11: llDemoVal = rst_Temp!drfDemo12
                    Case 12: llDemoVal = rst_Temp!drfDemo13
                    Case 13: llDemoVal = rst_Temp!drfDemo14
                    Case 14: llDemoVal = rst_Temp!drfDemo15
                    Case 15: llDemoVal = rst_Temp!drfDemo16
                    Case 16: llDemoVal = rst_Temp!drfDemo17
                    Case 17: llDemoVal = rst_Temp!drfDemo18
                End Select
                'place value in demo column - based on the mnfGroupNo that is loaded into the LbcCustDemos list itemdata
                tlRsr.lDemo(RptSelRS!LbcCustDemos.ItemData(illoop) - 1) = llDemoVal
            Next illoop
        Else
            '10/6/22 - JW - Fix per Jason email: v81 Research report - TTP 10556 (Issue 2)
            If slForm <> "8" Then                           '16 bucket Research
                tlRsr.lDemo(0) = rst_Temp!drfDemo1
                tlRsr.lDemo(1) = rst_Temp!drfDemo2
                tlRsr.lDemo(2) = rst_Temp!drfDemo3
                tlRsr.lDemo(3) = rst_Temp!drfDemo4
                tlRsr.lDemo(4) = rst_Temp!drfDemo5
                tlRsr.lDemo(5) = rst_Temp!drfDemo6
                tlRsr.lDemo(6) = rst_Temp!drfDemo7
                tlRsr.lDemo(7) = rst_Temp!drfDemo8
                tlRsr.lDemo(8) = 0
                tlRsr.lDemo(9) = rst_Temp!drfDemo9
                tlRsr.lDemo(10) = rst_Temp!drfDemo10
                tlRsr.lDemo(11) = rst_Temp!drfDemo11
                tlRsr.lDemo(12) = rst_Temp!drfDemo12
                tlRsr.lDemo(13) = rst_Temp!drfDemo13
                tlRsr.lDemo(14) = rst_Temp!drfDemo14
                tlRsr.lDemo(15) = rst_Temp!drfDemo15
                tlRsr.lDemo(16) = rst_Temp!drfDemo16
                tlRsr.lDemo(17) = 0
            Else                                            '18 bucket Research
                tlRsr.lDemo(0) = rst_Temp!drfDemo1
                tlRsr.lDemo(1) = rst_Temp!drfDemo2
                tlRsr.lDemo(2) = rst_Temp!drfDemo3
                tlRsr.lDemo(3) = rst_Temp!drfDemo4
                tlRsr.lDemo(4) = rst_Temp!drfDemo5
                tlRsr.lDemo(5) = rst_Temp!drfDemo6
                tlRsr.lDemo(6) = rst_Temp!drfDemo7
                tlRsr.lDemo(7) = rst_Temp!drfDemo8
                tlRsr.lDemo(8) = rst_Temp!drfDemo9
                tlRsr.lDemo(9) = rst_Temp!drfDemo10
                tlRsr.lDemo(10) = rst_Temp!drfDemo11
                tlRsr.lDemo(11) = rst_Temp!drfDemo12
                tlRsr.lDemo(12) = rst_Temp!drfDemo13
                tlRsr.lDemo(13) = rst_Temp!drfDemo14
                tlRsr.lDemo(14) = rst_Temp!drfDemo15
                tlRsr.lDemo(15) = rst_Temp!drfDemo16
                tlRsr.lDemo(16) = rst_Temp!drfDemo17
                tlRsr.lDemo(17) = rst_Temp!drfDemo18
            End If
        End If
        ilRet = btrInsert(hmRsr, tlRsr, imRsrRecLen, INDEXKEY0)
        rst_Temp.MoveNext
    Loop
End Sub

'10/6/22 - JW - Fix per Jason email: v81 Research report - TTP 10556 (Issue 1)
Private Sub mGetDpfRecords(ilDnfCode As Integer, ilVefCode As Integer)
    'DNF = "DNF_Demo_Rsrch_Names"
    Dim slSQLQuery As String
    Dim rst_Temp As ADODB.Recordset
    Dim tlRsr As RSR
    Dim slForm As String
    Dim illoop As Integer
    Dim ilRet As Integer
    slForm = "8" 'TODO: Check this was correct, to hardcode slForm as was previously done
    Debug.Print " - Get DpfRecords - Book:" & ilDnfCode
    
    '---------------------------------
    'Build DPF Query
    slSQLQuery = ""
    slSQLQuery = slSQLQuery & "SELECT "
    slSQLQuery = slSQLQuery & "     dpfCode, dpfdrfCode, dpfMnfDemo, dpfDnfCode, dpfDemo "
    slSQLQuery = slSQLQuery & "FROM DPF_Demo_Plus_Data "
    slSQLQuery = slSQLQuery & "     JOIN DRF_Demo_Rsrch_Data on dpfDrfCode = drfCode "
    slSQLQuery = slSQLQuery & "WHERE "
    slSQLQuery = slSQLQuery & "     dpfDnfCode = " & ilDnfCode 'Book
    slSQLQuery = slSQLQuery & "     AND drfVefCode = " & ilVefCode 'Vehicle
    slSQLQuery = slSQLQuery & "ORDER BY "
    slSQLQuery = slSQLQuery & "     dpfmnfDemo, dpfDnfCode "
    
    Set rst_Temp = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Temp.EOF
        '---------------------------------
        'Write RSR
        tlRsr.iGenDate(0) = igNowDate(0)    'Date stamp for Crystal to Key off of
        tlRsr.iGenDate(1) = igNowDate(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tlRsr.lGenTime = lgNowTime          'Time stamp for Crystal to Key off of
        tlRsr.lDrfCode = rst_Temp!dpfdrfCode   'DRF Code
        tlRsr.iPopAud = 1   'P = Population (0), Otherwise Audience (1)
        '---------------------------------
        'Report Headers
        tlRsr.sDemoDesc(0) = "P12+"
        For illoop = 1 To 17
            tlRsr.sDemoDesc(illoop) = ""
        Next illoop
        '---------------------------------
        'Research Data
        tlRsr.lDemo(0) = rst_Temp!dpfDemo
        
        ilRet = btrInsert(hmRsr, tlRsr, imRsrRecLen, INDEXKEY0)
        rst_Temp.MoveNext
    Loop
End Sub


Private Sub mDrfToRsr(ilDnfCode As Integer, ilVefCode As Integer, LbcCustDemos As ListBox)
    Dim llresults() As Long
    Dim c As Integer
    Dim slInfoType As String
    Dim llDrfCode As Long
    Dim blRegularSelected As Boolean
    Dim blCustomSelected As Boolean
    Dim illoop As Integer
    Dim ilRedim As Integer
    Dim blCurrentCustom As Boolean
    Dim ilRdfCodes() As Integer
    Dim j As Integer
    Dim tlDrfSrchKey As DRFKEY0
    Dim tlDrf As DRF
    Dim slForm As String

    If (RptSelRS!ckcSelC3(14).Value = vbChecked) Then
        blRegularSelected = True
    End If
    If (RptSelRS!ckcSelC3(13).Value = vbChecked) Then
        blCustomSelected = True
    End If
    tlDrfSrchKey.iDnfCode = ilDnfCode
    tlDrfSrchKey.iMnfSocEco = 0

    'Note:  llResults is always 1 larger than is needed -- 18, 16 or count of listBox(custom headings)
    'perform similar action for regular(1), and then custom(2)
    For illoop = 1 To 2
        If (illoop = 1 And blRegularSelected) Or (illoop = 2 And blCustomSelected) Then
            'regular or custom?
            If illoop = 1 Then
                ilRedim = 18
                blCurrentCustom = False
            Else
                ilRedim = LbcCustDemos.ListCount
                blCurrentCustom = True
            End If
            ReDim llresults(ilRedim)
            'pop
            tlDrfSrchKey.sDemoDataType = "P"            'population
            tlDrfSrchKey.sInfoType = " "
            tlDrfSrchKey.iRdfCode = 0
            tlDrfSrchKey.iVefCode = 0
            llDrfCode = mReadDrfs(tlDrfSrchKey, blCurrentCustom, LbcCustDemos, llresults, slForm)
            If llDrfCode > 0 Then
                mFillRsr blCurrentCustom, True, " ", llresults, LbcCustDemos, llDrfCode, slForm
            End If
            ' separate audience by type-read and send for each type selected
            If (Not bmIsItImpressions) Or ((bmIsItImpressions) And (illoop = 1)) Then
                For c = 0 To 3
                    'if not using dayparts,  rdfCode(0) = 0
                    ReDim ilRdfCodes(1)
                    ilRdfCodes(0) = 0
                    tlDrfSrchKey.sDemoDataType = "D"
                    tlDrfSrchKey.iVefCode = ilVefCode
                    tlDrfSrchKey.iRdfCode = 0
                    '4940
                    If Not (RptSelRS!ckcSelC3(0).Value = vbChecked) Then
                        tlDrfSrchKey.iMnfSocEco = 0
                    End If
                    If RptSelRS!ckcSelC5(c).Value = vbChecked Then
                     'D = Sold DP, Extra DP = D, Time = T, Vehicle = V
                        Select Case c
                            Case 0
                                tlDrfSrchKey.sInfoType = "D"
                                slInfoType = "D"
                                mGetRdfCodes tlDrfSrchKey, ilRdfCodes
                            Case 1
                                'send X so can know extra daypart; will set to D before using
                                tlDrfSrchKey.sInfoType = "X"
                                slInfoType = "D"
                            Case 2
                                tlDrfSrchKey.sInfoType = "T"
                                slInfoType = "T"
                            Case 3
                                tlDrfSrchKey.sInfoType = "V"
                                slInfoType = "V"
                        End Select
                        'clear llResults
                        Select Case c
                            Case 0
                            ReDim llresults(ilRedim)
                            'aud, one for each rdfCode--only once if other than daypart
                            For j = 0 To UBound(ilRdfCodes) - 1 Step 1
                                tlDrfSrchKey.iRdfCode = ilRdfCodes(j)
                                llDrfCode = mReadDrfs(tlDrfSrchKey, blCurrentCustom, LbcCustDemos, llresults, slForm)
                                If llDrfCode > 0 Then
                                    mFillRsr blCurrentCustom, False, slInfoType, llresults, LbcCustDemos, llDrfCode, slForm
                                End If
                            Next j
                        Case 1, 2, 3
                            tlDrfSrchKey.iRdfCode = 0
                            tlDrfSrchKey.sInfoType = slInfoType
                            mReadDrfsForNonDP tlDrfSrchKey, slInfoType, blCurrentCustom, LbcCustDemos, llresults

                        End Select
                    End If
                Next c
            End If
        End If
    Next illoop
    Erase llresults
End Sub

Private Sub mGetRdfCodes(tlDrfSrchKey As DRFKEY0, ilRdfs() As Integer)   '
    'O- ilRdfs()
    Dim ilRet As Integer
    Dim tlDrf As DRF
    Dim ilUpperBound As Integer
    ReDim ilRdfs(0)

    ilRet = btrGetGreaterOrEqual(hmDrf, tlDrf, imDrfRecLen, tlDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        Do While (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tlDrfSrchKey.iDnfCode) And (tlDrf.sDemoDataType = "D") And tlDrf.iVefCode = tlDrfSrchKey.iVefCode
            If (tlDrf.sInfoType = "D") Then
                '4940  socio not checked, so iMnfSocEco must be 0
                If (RptSelRS!ckcSelC3(0).Value <> vbChecked And tlDrf.iMnfSocEco = 0) Or RptSelRS!ckcSelC3(0).Value = vbChecked Then
                    If mIsUniqueRdf(tlDrf.iRdfCode, ilRdfs) Then
                        ilUpperBound = UBound(ilRdfs)
                        ReDim Preserve ilRdfs(ilUpperBound + 1)
                        ilRdfs(ilUpperBound) = tlDrf.iRdfCode
                    End If
                End If
            End If
            ilRet = btrGetNext(hmDrf, tlDrf, imDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    If UBound(ilRdfs) = 1 Then
        ReDim ilRdfs(1)
        ilRdfs(0) = 0
    End If
End Sub

Private Function mIsUniqueRdf(ilRdfCode As Integer, ilRdfs() As Integer) As Boolean
    Dim c As Integer
    Dim blRet As Boolean
    blRet = True
    If ilRdfCode = 0 Then
        blRet = False
    Else
        For c = 0 To UBound(ilRdfs)
            If ilRdfCode = ilRdfs(c) Then
                blRet = False
                Exit For
            End If
        Next c
    End If
    mIsUniqueRdf = blRet
End Function

Private Function mReadDrfs(tlDrfSrchKey As DRFKEY0, blCustom As Boolean, smCustDemos As ListBox, llresults() As Long, slForm As String) As Long
    'out -- llResults()
    ' for custom, collect all relevent data into array, and then send out.  If there are 36 custom headings, get 36 values(if they exist) before writing to rsr.
    Dim slDataType As String
    Dim c As Integer
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim blExtraDaypart As Boolean
    Dim blContinue As Boolean
    Dim tlDrf As DRF

    If tlDrfSrchKey.sInfoType = "X" Then
        tlDrfSrchKey.sInfoType = "D"
        blExtraDaypart = True
    End If
    ilRet = btrGetGreaterOrEqual(hmDrf, tlDrf, imDrfRecLen, tlDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE Then
        Do While (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tlDrfSrchKey.iDnfCode) And (tlDrf.sDemoDataType = tlDrfSrchKey.sDemoDataType) And tlDrf.iVefCode = tlDrfSrchKey.iVefCode
            If (tlDrf.sInfoType = tlDrfSrchKey.sInfoType) And (tlDrfSrchKey.iRdfCode = 0 Or tlDrf.iRdfCode = tlDrfSrchKey.iRdfCode) Then
                blContinue = True
                If tlDrf.sInfoType = "D" Then
                    'TF or FT , quit
                    blContinue = Not (blExtraDaypart Xor tlDrf.iRdfCode = 0)
                End If
                If blContinue Then
                    ' custom and not A  OR  not custom and A
                    'FF or TT is okay
                    If Not (blCustom Xor tlDrf.sDataType <> "A") Then
                            If mOKtoSaveRec(tlDrf) Then
                            ''6 or blank = 16 buckets--will read dimension to determine names to add for regular
                                slForm = "8"
                                If Not blCustom And tlDrf.sForm <> "8" Then
                                    'ReDim llresults(16)
                                    ReDim llresults(18)
                                End If
                                If tlDrf.sForm <> "8" Then
                                    slForm = ""
                                End If
                                mFillArrayFromDrf blCustom, tlDrf, smCustDemos, llresults, slForm
                                mReadDrfs = tlDrf.lCode
                            Else
                            End If
                    End If  'custom?
                End If 'blcontinue
            End If 'right infotype
            ilRet = btrGetNext(hmDrf, tlDrf, imDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
End Function

Private Sub mFillArrayFromDrf(blCustom As Boolean, tlDrf As DRF, smCustDemos As ListBox, llresults() As Long, slForm As String)
    ' 6/28/19 determine if impressions book;  if so, need to get the dpf table to retieve aud values.  this type won't be defined as custom demos
    Dim c As Integer
    Dim ilUpper As Integer
    Dim ilBase As Integer
    Dim ilBucket As Integer
    Dim ilMax As Integer
    Dim tlSrchKey0 As INTKEY0
    Dim tlSrchKey1 As DPFKEY1
    Dim ilRet As Integer

    ilUpper = UBound(llresults) - 1
    If Not blCustom Then
        If bmIsItImpressions Then
            tlSrchKey1.lDrfCode = tlDrf.lCode
            tlSrchKey1.iMnfDemo = 0
            ilRet = btrGetGreaterOrEqual(hmDpf, tmDpf, imDpfRecLen, tlSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching estimates recd
            If ilRet = BTRV_ERR_NONE And tmDpf.lDrfCode = tlDrf.lCode Then
                For c = 0 To ilUpper Step 1
                    'for P12+ Impressions, get the impreesions from dpf and place into drf for common code
                    If c = 0 Then
                        llresults(c) = tmDpf.lDemo
                    Else
                        llresults(c) = 0
                    End If
                Next c
            End If
        Else
            If slForm = "8" Then
                For c = 0 To ilUpper Step 1
                    llresults(c) = tlDrf.lDemo(c)
                Next c
            Else            '16 demos, adjust to place into correct columns (1-8, not 1-9)
                For c = 0 To 7 Step 1
                    llresults(c) = tlDrf.lDemo(c)
                Next c
                llresults(8) = 0
                For c = 9 To 16 Step 1
                    llresults(c) = tlDrf.lDemo(c - 1)
                Next c
                llresults(17) = 0
            End If
        End If
    Else
        'gReconvert returns first number of set: 1,19,37.  ilBase becomes 0,18,36
        ilBase = gReconvertCustomGroup(tlDrf.sDataType) - 1
        ' if custom headings end before multiple of 18 (34 instead of 36), change ilMax to that value ( 15 instead of 17)
        ilMax = ilBase + 17
        If ilUpper < ilMax Then
            ilMax = (ilUpper - ilBase)
        Else
            ilMax = 17
        End If
        'llResult index =  listbox itemdata -1.  example(listbox.item(3) = 'Boys'  listbox.itemData(3) = 24  llResult(24) = 60..the value that goes with 'Boys'
        For ilBucket = 0 To ilMax Step 1
            llresults(ilBase + ilBucket) = tlDrf.lDemo(ilBucket)
        Next ilBucket
    End If
    Exit Sub
End Sub

Private Sub mFillRsr(blCustom As Boolean, blPopNotDemo As Boolean, slInfoType As String, llresults() As Long, smCustom As ListBox, llDrfCode As Long, slForm As String)
    Dim ilRet As Integer
    Dim tlRsr As RSR
    Dim c As Integer
    Dim ilMaxSet As Integer
    Dim ilBase As Integer
    Dim ilMax As Integer
    Dim ilCurrentNumber As Integer
    Dim slDemoType As String
    Dim ilNull As Integer
    Dim illoop As Integer
    
    tlRsr.iGenDate(0) = igNowDate(0)    'Date stamp for Crystal to Key off of
    tlRsr.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tlRsr.lGenTime = lgNowTime
    tlRsr.lDrfCode = llDrfCode
    tlRsr.sDataType = slInfoType        'D = Sold DP, Extra DP = D, Time = T, Vehicle = V
    If blPopNotDemo Then                'P = Population (0), Otherwise Audience (1)
        tlRsr.iPopAud = 0
    Else
        tlRsr.iPopAud = 1
    End If
    If Not blCustom Then
        tlRsr.sDemoType = "A"           'A = Standard Demo B = Custom Demo - for sorting
        If slForm <> "8" Then
            tlRsr.sDemoDesc(0) = "M12-17"
            tlRsr.sDemoDesc(1) = "M18-24"
            tlRsr.sDemoDesc(2) = "M25-34"
            tlRsr.sDemoDesc(3) = "M35-44"
            tlRsr.sDemoDesc(4) = "M45-49"
            tlRsr.sDemoDesc(5) = "M50-54"
            tlRsr.sDemoDesc(6) = "M55-64"
            tlRsr.sDemoDesc(7) = "M65+"
            tlRsr.sDemoDesc(8) = ""
            tlRsr.sDemoDesc(9) = "W12-17"
            tlRsr.sDemoDesc(10) = "W18-24"
            tlRsr.sDemoDesc(11) = "W25-34"
            tlRsr.sDemoDesc(12) = "W35-44"
            tlRsr.sDemoDesc(13) = "W45-49"
            tlRsr.sDemoDesc(14) = "W50-54"
            tlRsr.sDemoDesc(15) = "W55-64"
            tlRsr.sDemoDesc(16) = "W65+"
            tlRsr.sDemoDesc(17) = ""

            For c = 0 To 17
                tlRsr.lDemo(c) = llresults(c)
            Next c
        Else                            '8 = 18 buckets
            If bmIsItImpressions Then
                tlRsr.sDemoDesc(0) = "P12+"
                For illoop = 1 To 17
                    tlRsr.sDemoDesc(illoop) = ""
                Next illoop
                For c = 0 To 17
                    tlRsr.lDemo(c) = llresults(c)
                Next c
            Else
                tlRsr.sDemoDesc(0) = "M12-17"             '16 buckets that get filled with the standard demo names
                tlRsr.sDemoDesc(1) = "M18-20"             'new category
                tlRsr.sDemoDesc(2) = "M21-24"             'new category
                tlRsr.sDemoDesc(3) = "M25-34"
                tlRsr.sDemoDesc(4) = "M35-44"
                tlRsr.sDemoDesc(5) = "M45-49"
                tlRsr.sDemoDesc(6) = "M50-54"
                tlRsr.sDemoDesc(7) = "M55-64"
                tlRsr.sDemoDesc(8) = "M65+"
                tlRsr.sDemoDesc(9) = "W12-17"
                tlRsr.sDemoDesc(10) = "W18-20"
                tlRsr.sDemoDesc(11) = "W21-24"
                tlRsr.sDemoDesc(12) = "W25-34"
                tlRsr.sDemoDesc(13) = "W35-44"
                tlRsr.sDemoDesc(14) = "W45-49"
                tlRsr.sDemoDesc(15) = "W50-54"
                tlRsr.sDemoDesc(16) = "W55-64"
                tlRsr.sDemoDesc(17) = "W65+"
                For c = 0 To 17
                    tlRsr.lDemo(c) = llresults(c)
                Next c
            End If
        End If
        ilRet = btrInsert(hmRsr, tlRsr, imRsrRecLen, INDEXKEY0)
    Else
        'custom.
        'llResults is 1 larger than is used
        'llResult index =  listbox itemdata -1.  example(listbox.item(3) = 'Boys'  listbox.itemData(3) = 24  llResult(23) = 60..the value that goes with 'Boys'
        ilMax = UBound(llresults) - 1
        '36 custom headings? 2 sets of 18
        ilMaxSet = (ilMax \ 18)
        'ilMax is 'array-safe' which is 0 to x-1, not 1 to x
        For ilBase = 0 To ilMaxSet Step 1
            For c = 0 To 17 Step 1
                ' if 2nd set, 0 = 18, 1 = 19, etc.
                If c = 5 Then
                    c = c
                End If
                ilCurrentNumber = (ilBase * 18) + c
                If ilCurrentNumber <= ilMax Then
                    tlRsr.sDemoDesc(c) = smCustom.List(ilCurrentNumber)
                    tlRsr.lDemo(c) = llresults(smCustom.ItemData(ilCurrentNumber) - 1)
                Else
                    tlRsr.sDemoDesc(c) = ""
                    tlRsr.lDemo(c) = 0
                End If
            Next c
            'determine this groups 'B', 'd', etc.
            gConvCustomGroup ilBase * 18 + c, slDemoType, ilNull
            tlRsr.sDemoType = slDemoType
            ilRet = btrInsert(hmRsr, tlRsr, imRsrRecLen, INDEXKEY0)
        Next ilBase
    End If
End Sub

Public Sub gSpecialResearchReport()
    '******************************************************************************************
    '* Note: VBC id'd the following unreferenced items and handled them as described:         *
    '*                                                                                        *
    '* Local Variables (Removed)                                                              *
    '*  ilLoopVef                                                                             *
    '******************************************************************************************
    Dim ilLoopBook As Integer
    Dim slNameCode As String
    Dim slDate As String
    Dim ilSortCode As Integer
    Dim slCode As String
    Dim ilRet As Integer
    Dim llBookDate As Long
    Dim ilBookDate(0 To 1) As Integer
    Dim ilDnfCode As Integer
    Dim ilUseInclVefCodes As Integer
    ReDim ilUseVefCodes(0 To 0) As Integer

    imListIndex = RptSelRS!lbcRptType.ListIndex
    mOpenBtrFiles imListIndex    'Open all necessary BTR files
    ilSortCode = 0

    gObtainCodesForMultipleLists 0, tgVehicle(), ilUseInclVefCodes, ilUseVefCodes(), RptSelRS

    'generated date and time of prepass for selection
    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    tmGrf.lGenTime = lgNowTime

    'Loop on all selected books
    For ilLoopBook = 0 To RptSelRS!lbcSelection(1).ListCount - 1 Step 1
        If RptSelRS!lbcSelection(1).Selected(ilLoopBook) Then
            slNameCode = tgBookNameCode(ilLoopBook).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilDnfCode = Val(slCode)
            ilRet = gParseItem(slNameCode, 1, "|", slDate)
            llBookDate = 99999 - Val(slDate)
            gPackDateLong llBookDate, ilBookDate(0), ilBookDate(1)
            'process selected BookNames Estimates
            mObtainEstimates ilDnfCode, ilBookDate()
            'process all the demos for the selected book
            mObtainDemoDP ilDnfCode, ilBookDate(), ilUseInclVefCodes, ilUseVefCodes()

        End If
    Next ilLoopBook
    '1/24/2011 Dan M
    mCloseBtrFiles imListIndex
End Sub

'*******************************************************
'           Get all Demo Estimates for the selected book
'           and create a prepass record
'
'           2-10-06
Private Sub mObtainEstimates(ilDnfCode As Integer, ilBookDate() As Integer)
    Dim ilRet As Integer
    Dim ilCreatedOne As Integer
    Dim ilDrfLoop As Integer
    Dim ilTemp As Integer
    ReDim tmDrfForBook(0 To 0) As DRF

    'Build the demos in memory first so the population record can be retrieved
    ilRet = gObtainDRFByCode(RptSelRS, hmDrf, tmDrfForBook(), ilDnfCode)    'build array of all the matching demos for the selected book
    
    tmGrf.lDollars(1) = 0           'init subscriber base or listener count
    'Look for the Population record first, to get base subscriber or listener count
    For ilDrfLoop = LBound(tmDrfForBook) To UBound(tmDrfForBook) - 1
        If tmDrfForBook(ilDrfLoop).sDemoDataType = "P" Then
            For ilTemp = 1 To 18
                tmGrf.lDollars(1) = tmGrf.lDollars(1) + tmDrfForBook(ilDrfLoop).lDemo(ilTemp - 1)     'accum the aud demo categories
            Next ilTemp
            Exit For
        End If
    Next ilDrfLoop

    tmDefSrchKey1.iDnfCode = ilDnfCode
    tmDefSrchKey1.iStartDate(0) = ilBookDate(0)
    tmDefSrchKey1.iStartDate(1) = ilBookDate(1)

    ilRet = btrGetGreaterOrEqual(hmDef, tmDef, imDefRecLen, tmDefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching estimates recd
    ilCreatedOne = False
    Do While (ilRet = BTRV_ERR_NONE) And (tmDef.iDnfCode = ilDnfCode)
        ilCreatedOne = True
        'create prepass record
        tmGrf.iSofCode = ilDnfCode          'book name
        tmGrf.iStartDate(0) = ilBookDate(0)     'book date
        tmGrf.iStartDate(1) = ilBookDate(1)
        tmGrf.iDate(0) = tmDef.iStartDate(0)    'date estimate
        tmGrf.iDate(1) = tmDef.iStartDate(1)
        tmGrf.iCode2 = 0                        'type for sorting indicating header
        tmGrf.iVefCode = 0                      'headers dont have vehicle reference
        tmGrf.sGenDesc = ""                     'headers dont have DP description
        tmGrf.lCode4 = tmDef.lCode          'estimates record code
        tmGrf.lChfCode = 0                      'no demo (drf) reference for estimates info
        tmGrf.sBktType = ""                     'no DP type for header (sold or extra)
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        ilRet = btrGetNext(hmDef, tmDef, imDefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop

    'if no estimates created, create one to show message in header
    If ilCreatedOne = False Then
        tmGrf.iSofCode = ilDnfCode          'book name
        tmGrf.iStartDate(0) = ilBookDate(0)     'book date
        tmGrf.iStartDate(1) = ilBookDate(1)
        tmGrf.iDate(0) = 0
        tmGrf.iDate(1) = 0
        tmGrf.iCode2 = 0                        'type for sorting indicating header
        tmGrf.iVefCode = 0                      'headers dont have vehicle reference
        tmGrf.sGenDesc = "* No estimates exist *"    'message for header
        tmGrf.lCode4 = tmDef.lCode          'estimates record code
        tmGrf.lChfCode = 0                      'no demo (drf) reference for estimates info
        tmGrf.sBktType = ""                     'no DP type for header (sold or extra)
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        ilRet = btrGetNext(hmDef, tmDef, imDefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    End If
    Exit Sub
End Sub

'*******************************************************
'           Obtain all the demos belonging to the requested book
'           Build in array tmDrfForBook
'
'           <input> Book Name code
'
'           dh:  2-10-06
Private Sub mObtainDemoDP(ilDnfCode As Integer, ilBookDate() As Integer, ilIncludeCodes As Integer, ilUseCodes() As Integer)
    Dim ilTemp As Integer
    Dim ilFoundOption As Integer
    Dim ilDrfLoop As Integer
    Dim ilRet As Integer
    Dim slDaysOfWk(0 To 6) As String * 1
    Dim ilDays(0 To 6) As Integer
    Dim slTempDays As String
    Dim slStr As String
    Dim slTime As String
    
    For ilDrfLoop = LBound(tmDrfForBook) To UBound(tmDrfForBook) - 1
        ilFoundOption = False
        'ignore population records
        If tmDrfForBook(ilDrfLoop).sDemoDataType = "D" Or tmDrfForBook(ilDrfLoop).sDemoDataType = "M" Then      'imported or manually entered data (vs population record)
            If ilIncludeCodes Then
                For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                    If ilUseCodes(ilTemp) = tmDrfForBook(ilDrfLoop).iVefCode Then
                        ilFoundOption = True
                        Exit For
                    End If
                Next ilTemp
            Else
                ilFoundOption = True
                For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                    If ilUseCodes(ilTemp) = tmDrfForBook(ilDrfLoop).iVefCode Then
                        ilFoundOption = False
                        Exit For
                    End If
                Next ilTemp
            End If
        End If

        If ilFoundOption Then           'valid vehicle selected
            'create prepass record
            tmGrf.iSofCode = ilDnfCode          'book name
            tmGrf.iStartDate(0) = ilBookDate(0)     'book date
            tmGrf.iStartDate(1) = ilBookDate(1)
            tmGrf.iDate(0) = 0                      'no start date for detail record
            tmGrf.iDate(1) = 0
            tmGrf.iCode2 = 1                        'type for sorting indicating header
            tmGrf.iVefCode = tmDrfForBook(ilDrfLoop).iVefCode                      'headers dont have vehicle reference
            tmGrf.sGenDesc = ""                     'headers dont have DP description
            tmGrf.lCode4 = 0          'estimates record code
            tmGrf.lChfCode = tmDrfForBook(ilDrfLoop).lCode           'Demo Code

            'get the total audience for all categories
            tmGrf.lDollars(0) = 0
            For ilTemp = 1 To 18
                tmGrf.lDollars(0) = tmGrf.lDollars(0) + tmDrfForBook(ilDrfLoop).lDemo(ilTemp - 1)     'accum the aud demo categories
            Next ilTemp

            'if Sold Daypart indicated by flag "S", then use the DP name for description;
            'otherwise concatenate the days and times
            tmGrf.sGenDesc = ""
            If tmDrfForBook(ilDrfLoop).sDemoDataType = "D" And tmDrfForBook(ilDrfLoop).iRdfCode > 0 Then
                tmGrf.iRdfCode = tmDrfForBook(ilDrfLoop).iRdfCode       'get DP desc from DP record
                tmGrf.sBktType = "S"                'sold dp
            ElseIf (tmDrfForBook(ilDrfLoop).sDemoDataType = "D" And tmDrfForBook(ilDrfLoop).iRdfCode = 0) Or tmDrfForBook(ilDrfLoop).sDemoDataType = "T" Then
                tmGrf.iRdfCode = 0                              'get descr from tmgrf.sGenDesc
                tmGrf.sBktType = "E"                'extra DP

                For ilTemp = 0 To 6
                    slDaysOfWk(ilTemp) = ""     'unused for now, reqd for gDaynames parameter
                    If tmDrfForBook(ilDrfLoop).sDay(ilTemp) = "Y" Then
                        ilDays(ilTemp) = 1
                    Else
                        ilDays(ilTemp) = 0
                    End If
                Next ilTemp

                slTempDays = gDayNames(ilDays(), slDaysOfWk(), 2, slStr)            'slstr not needed when returned
                gUnpackTime tmDrfForBook(ilDrfLoop).iStartTime(0), tmDrfForBook(ilDrfLoop).iStartTime(1), "A", "1", slTime

                slTempDays = slTempDays & " " & Trim$(slTime) & "-"
                gUnpackTime tmDrfForBook(ilDrfLoop).iEndTime(0), tmDrfForBook(ilDrfLoop).iEndTime(1), "A", "1", slTime
                slTempDays = slTempDays & Trim$(slTime)
                tmGrf.sGenDesc = Trim$(slTempDays)
                If tmDrfForBook(ilDrfLoop).iStartTime2(0) <> 0 Or tmDrfForBook(ilDrfLoop).iStartTime2(1) <> 0 Then      'a second set of start/end times exists
                    gUnpackTime tmDrfForBook(ilDrfLoop).iStartTime(0), tmDrfForBook(ilDrfLoop).iStartTime(1), "A", "1", slTime
                    slTempDays = slTempDays & ", " & Trim$(slTime) & "-"
                    gUnpackTime tmDrfForBook(ilDrfLoop).iEndTime(0), tmDrfForBook(ilDrfLoop).iEndTime(1), "A", "1", slTime
                    slTempDays = slTempDays & Trim$(slTime)

                End If
            End If

            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        End If
    Next ilDrfLoop
    Exit Sub
End Sub

Private Sub mReadDrfsForNonDP(tlDrfSrchKey As DRFKEY0, slInfoType As String, blCustom As Boolean, smCustDemos As ListBox, llresults() As Long)
    'out -- llResults()
    ' for custom, collect all relevent data into array, and then send out.  If there are 36 custom headings, get 36 values(if they exist) before writing to rsr.
    Dim slDataType As String
    Dim c As Integer
    Dim tlDrf As DRF
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim blExtraDaypart As Boolean
    Dim blContinue As Boolean
    Dim slForm As String
    
    blExtraDaypart = True
    ilRet = btrGetGreaterOrEqual(hmDrf, tlDrf, imDrfRecLen, tlDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE Then
        Do While (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tlDrfSrchKey.iDnfCode) And (tlDrf.sDemoDataType = tlDrfSrchKey.sDemoDataType) And tlDrf.iVefCode = tlDrfSrchKey.iVefCode
            If (tlDrf.sInfoType = tlDrfSrchKey.sInfoType) And (tlDrf.iRdfCode = 0) Then
                blContinue = True
                If blContinue Then
                    ' custom and not A  OR  not custom and A
                    'FF or TT is okay
                    If Not (blCustom Xor tlDrf.sDataType <> "A") Then
                        If mOKtoSaveRec(tlDrf) Then
                            '6 or blank = 16 buckets--will read dimension to determine names to add for regular
                            slForm = "8"
                            If Not blCustom And tlDrf.sForm <> "8" Then
                                'ReDim llresults(16)
                                ReDim llresults(18)
                            End If
                            If tlDrf.sForm <> "8" Then
                                slForm = ""
                            End If
                            mFillArrayFromDrf blCustom, tlDrf, smCustDemos, llresults, slForm
                            mFillRsr blCustom, False, slInfoType, llresults, smCustDemos, tlDrf.lCode, tlDrf.sForm
                        End If
                    End If  'custom?
                End If 'blcontinue
            End If 'right infotype
            ilRet = btrGetNext(hmDrf, tlDrf, imDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
End Sub

'*******************************************************
'           Create report that shows vehicle ranks in specified research categories, based on the
'           default book inside each vehicle
Public Sub gDemoRankReport()
    Dim ilLoopOnVef As Integer
    Dim ilDemoLoop As Integer
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim ilVefCode As Integer
    Dim ilVefInx As Integer
    Dim ilDefaultDnfCode As Integer
    Dim slCode As String
    Dim llPop As Long
    Dim ilSocEcoMnfCode As Integer
    Dim ilDemoList() As Integer
    Dim llStartDateForEst As Long
    Dim llEndDateForEst As Long
    Dim ilRdfCode As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim llAvgAud As Long
    Dim llPopEst As Long
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long
    Dim llRafCode As Long
    Dim ilUpperDemo As Integer
    Dim ilLoopSortedDemo As Integer
    Dim ilValidDays(0 To 6) As Integer
    Dim illoop As Integer
    Dim ilUpper As Integer
    Dim tmDnf(0 To 0) As DNF
    Dim tmDrfForBook(0 To 0) As DRF
    Dim slDemoNames(0 To 4) As String
    Dim ilPrimaryIndex As Integer
    Dim slPrimaryDemo As String

    If Not mOpenDemoRankFiles(RptSelRS) Then
        Exit Sub
    End If

    'init variables for call to avg aud
    ilSocEcoMnfCode = 0
    llStartDateForEst = 0
    llEndDateForEst = 0
    ilRdfCode = 0
    llOvStartTime = -2       '-2 override times indicate Vehicle times only
    llOvEndTime = -2
    llRafCode = 0
    ilRdfCode = 0
    For illoop = 0 To 6                         'days of the week
        ilValidDays(illoop) = False              'force alldays as not valid
    Next illoop
    
    'init the desc field when running report without exiting selection screen
    For illoop = LBound(tmRsr.sDemoDesc) To UBound(tmRsr.sDemoDesc) - 1
        tmRsr.sDemoDesc(illoop) = ""
    Next illoop
    
    'build list of selectede demos
    ReDim ilDemoList(0 To 0) As Integer
    ReDim slDemoName(0 To 0) As String
    For ilDemoLoop = 0 To RptSelRS!lbcSelection(2).ListCount - 1 Step 1
        If (RptSelRS!lbcSelection(2).Selected(ilDemoLoop)) Then
            slNameCode = tgRptSelDemoCodeCT(ilDemoLoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilDemoList(UBound(ilDemoList)) = Val(slCode)                     'mnf code to Demo name
            slDemoName(UBound(slDemoName)) = Trim$(RptSelRS!lbcSelection(2).List(ilDemoLoop))
            ReDim Preserve ilDemoList(0 To UBound(ilDemoList) + 1) As Integer
            ReDim Preserve slDemoName(0 To UBound(slDemoName) + 1) As String
        End If
    Next ilDemoLoop
    
    ilPrimaryIndex = RptSelRS!cbcPrimaryDemo.ListIndex
    slPrimaryDemo = RptSelRS!cbcPrimaryDemo.List(ilPrimaryIndex)
    
    ReDim tmDemoRank(0 To 0) As DEMORANK_INFO
    For ilLoopOnVef = 0 To RptSelRS!lbcSelection(0).ListCount - 1
         If RptSelRS!lbcSelection(0).Selected(ilLoopOnVef) Then
            slNameCode = tgVehicle(ilLoopOnVef).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            ilVefInx = gBinarySearchVef(ilVefCode)
            If ilVefInx >= 0 Then
                ilDefaultDnfCode = tgMVef(ilVefInx).iDnfCode
                If ilDefaultDnfCode >= 0 Then            'does book exist?
                    ilUpperDemo = UBound(tmDemoRank)
                    tmDemoRank(ilUpperDemo).iVefCode = ilVefCode
                    For ilDemoLoop = 0 To UBound(ilDemoList) - 1
                        tmDemoRank(ilUpperDemo).iMnfDemo(ilDemoLoop) = ilDemoList(ilDemoLoop)
                        ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDefaultDnfCode, ilSocEcoMnfCode, ilDemoList(ilDemoLoop), llPop)
                        ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDefaultDnfCode, ilVefCode, ilSocEcoMnfCode, ilDemoList(ilDemoLoop), llStartDateForEst, llEndDateForEst, ilRdfCode, llOvStartTime, llOvEndTime, ilValidDays(), "S", llRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                        tmDemoRank(ilUpperDemo).lAvgAud(ilDemoLoop) = llAvgAud
                    Next ilDemoLoop
                    ReDim Preserve tmDemoRank(0 To ilUpperDemo + 1) As DEMORANK_INFO
                    ilUpperDemo = ilUpperDemo + 1
                End If
            End If
        End If
    Next ilLoopOnVef
    
    For ilDemoLoop = 0 To UBound(ilDemoList) - 1
        For ilLoopOnVef = LBound(tmDemoRank) To UBound(tmDemoRank) - 1      'loop on # of vehicles
            llAvgAud = tmDemoRank(ilLoopOnVef).lAvgAud(ilDemoLoop)
            slCode = Trim$(str(llAvgAud))
            Do While Len(slCode) < 10
                slCode = "0" & slCode
            Loop
            tmDemoRank(ilLoopOnVef).sKey = Trim$(slCode)
        Next ilLoopOnVef
        
        'sort each of the demo categories and rank them
        ilUpper = UBound(tmDemoRank)
        If ilUpper > 0 Then    'sort descending
            ArraySortTyp fnAV(tmDemoRank(), 0), ilUpper, 1, LenB(tmDemoRank(0)), 0, LenB(tmDemoRank(0).sKey), 0
            
            'set up the ranks
            For ilLoopSortedDemo = 0 To UBound(tmDemoRank) - 1
                tmDemoRank(ilLoopSortedDemo).iRank(ilDemoLoop) = ilLoopSortedDemo + 1
            Next ilLoopSortedDemo
        End If
    Next ilDemoLoop
    
    tmRsr.iGenDate(0) = igNowDate(0)
    tmRsr.iGenDate(1) = igNowDate(1)
    tmRsr.lGenTime = lgNowTime
    
    'Write all vehicles along with demos selected to temporary file
    For ilLoopOnVef = LBound(tmDemoRank) To UBound(tmDemoRank) - 1
        tmRsr.lDrfCode = tmDemoRank(ilLoopOnVef).iVefCode
        For ilDemoLoop = 0 To UBound(ilDemoList) - 1
            tmRsr.lDemo(ilDemoLoop) = tmDemoRank(ilLoopOnVef).lAvgAud(ilDemoLoop)
            tmRsr.lDemo(ilDemoLoop + 6) = tmDemoRank(ilLoopOnVef).iRank(ilDemoLoop)
            tmRsr.sDemoDesc(ilDemoLoop) = slDemoName(ilDemoLoop)
        Next ilDemoLoop
        ilRet = btrInsert(hmRsr, tmRsr, imRsrRecLen, INDEXKEY0)
    Next ilLoopOnVef
    
    Erase ilDemoList
    Erase tmDemoRank
    Erase tmDnf
    Erase tmDrfForBook
    
    ilRet = btrClose(hmRsr)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmDnf)
    ilRet = btrClose(hmDrf)
    ilRet = btrClose(hmDpf)
    ilRet = btrClose(hmDef)
    ilRet = btrClose(hmRaf)
    btrDestroy hmRsr
    btrDestroy hmMnf
    btrDestroy hmDnf
    btrDestroy hmDrf
    btrDestroy hmDpf
    btrDestroy hmDef
    btrDestroy hmRaf
    Exit Sub
End Sub

'*******************************************************
'           Open files for Demo Rank report
'           mOpenDemoRankFiles
'           <return> true of OK, else false
'
Private Function mOpenDemoRankFiles(frm As Form) As Boolean
    Dim ilRet As Integer

    mOpenDemoRankFiles = True
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mOpenDemoRankFilesErr
        gBtrvErrorMsg ilRet, "mOpenDemoRankFiles (btrOpen):" & "Mnf.Btr", frm
        On Error GoTo 0
        If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        Exit Function
    End If
    imMnfRecLen = Len(tmMnf)
    
    hmRsr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRsr, "", sgDBPath & "Rsr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    On Error GoTo mOpenDemoRankFilesErr
        gBtrvErrorMsg ilRet, "mOpenDemoRankFiles (btrOpen):" & "rsr.Btr", frm
        On Error GoTo 0
        ilRet = btrClose(hmRsr)
        ilRet = btrClose(hmMnf)
        btrDestroy hmRsr
        btrDestroy hmMnf
        Exit Function
    End If
    imRsrRecLen = Len(tmRsr)
    
    hmDnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        On Error GoTo mOpenDemoRankFilesErr
        gBtrvErrorMsg ilRet, "mOpenDemoRankFiles (btrOpen):" & "dnf.Btr", frm
        On Error GoTo 0
        ilRet = btrClose(hmDnf)
        ilRet = btrClose(hmRsr)
        ilRet = btrClose(hmMnf)
        btrDestroy hmDnf
        btrDestroy hmRsr
        btrDestroy hmMnf
        Exit Function
    End If
    imDnfRecLen = Len(tmDnf(0))
    
    hmDrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        On Error GoTo mOpenDemoRankFilesErr
        gBtrvErrorMsg ilRet, "mOpenDemoRankFiles (btrOpen):" & "drf.Btr", frm
        On Error GoTo 0
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmDnf)
        ilRet = btrClose(hmRsr)
        ilRet = btrClose(hmMnf)
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmRsr
        btrDestroy hmMnf
        Exit Function
    End If
    imDrfRecLen = Len(tmDrfForBook(0))
    
    hmDpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        On Error GoTo mOpenDemoRankFilesErr
        gBtrvErrorMsg ilRet, "mOpenDemoRankFiles (btrOpen):" & "dpf.Btr", frm
        On Error GoTo 0
        ilRet = btrClose(hmDpf)
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmDnf)
        ilRet = btrClose(hmRsr)
        ilRet = btrClose(hmMnf)
        btrDestroy hmDpf
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmRsr
        btrDestroy hmMnf
        Exit Function
    End If
    imDpfRecLen = Len(tmDpf)
    
    hmDef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        On Error GoTo mOpenDemoRankFilesErr
        gBtrvErrorMsg ilRet, "mOpenDemoRankFiles (btrOpen):" & "def.Btr", frm
        On Error GoTo 0
        ilRet = btrClose(hmDef)
        ilRet = btrClose(hmDpf)
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmDnf)
        ilRet = btrClose(hmRsr)
        ilRet = btrClose(hmMnf)
        btrDestroy hmDef
        btrDestroy hmDpf
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmRsr
        btrDestroy hmMnf
        Exit Function
    End If
    imDefRecLen = Len(tmDef)

    hmRaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRaf, "", sgDBPath & "raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        On Error GoTo mOpenDemoRankFilesErr
        gBtrvErrorMsg ilRet, "mOpenDemoRankFiles (btrOpen):" & "raf.Btr", frm
        On Error GoTo 0
        ilRet = btrClose(hmRaf)
        ilRet = btrClose(hmDef)
        ilRet = btrClose(hmDpf)
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmDnf)
        ilRet = btrClose(hmRsr)
        ilRet = btrClose(hmMnf)
        btrDestroy hmRaf
        btrDestroy hmDef
        btrDestroy hmDpf
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmRsr
        btrDestroy hmMnf
        Exit Function
    End If
    imRafRecLen = Len(tmRaf)
    Exit Function
        
mOpenDemoRankFilesErr:
    mOpenDemoRankFiles = False
    gDbg_HandleError "mOpenDemoRankFiles: btrOpen"
    Exit Function
End Function

'*******************************************************
'      Read the Demo book to detrmine if its an Impressions book
'      base on type = "I"
Public Function mIsItImpressions(ilDnfCode As Integer) As Boolean
    Dim tlSrchKey0 As INTKEY0
    Dim ilRet As Integer
    
    mIsItImpressions = False
    tlSrchKey0.iCode = ilDnfCode
    ilRet = btrGetEqual(hmDnf, tmDnf(0), imDnfRecLen, tlSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        If tmDnf(0).sSource = "I" Then          'Impressions book
            mIsItImpressions = True
        End If
    End If
End Function
