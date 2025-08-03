Attribute VB_Name = "COPYSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Copy.bas on Wed 6/17/09 @ 12:56 PM **
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Copy.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Copy subs and functions
Option Explicit
Option Compare Text

Public tgActiveCode() As SORTCODE
Public lgSetNextCvfCode As Long
Public igFLValue As Integer

Public igSetCopyTerminated As Integer

Public igCopyRotEraseAdfCode As Integer
Public lgCopyRotEraseCrfCode As Long

'******************************************************************************
' GNF_Grid_Name Record Definition
'
'******************************************************************************
Type GNF
    lCode                 As Long            ' Grid Name Reference code
    iAdfCode              As Integer         ' Advertiser reference code
    sGridName             As String * 60     ' Grid Name
    iEnteredDate(0 To 1)  As Integer         ' Entered Date
    iEnteredTime(0 To 1)  As Integer         ' Entered Time
    iUrfCode              As Integer         ' User reference code
    sUnused               As String * 10     ' Unused
End Type


'Type GNFKEY0
'    lCode                 As Long
'End Type

Type GNFKEY1
    iAdfCode              As Integer
End Type

Type GNFKEY2
    iAdfCode              As Integer         ' Advertiser reference code
    sGridName             As String * 60
End Type


'******************************************************************************
' GPF_Grid_Position Record Definition
'
'******************************************************************************
Type GPF
    lCode                 As Long            ' Grid Position (Row) reference
                                             ' code. This file and gcf are used
                                             ' to retain the Copy grid
                                             ' information
    lGnfCode              As Long            ' Grid Name reference code
    iRowNo                As Integer         ' Row number so that the
                                             ' information can be restored to
                                             ' the exact row
    lCsfCode              As Long            ' Comment reference code
    sUnused               As String * 14
End Type

'Type GPFKEY0
'    lCode                 As Long
'End Type

Type GPFKEY1
    lGnfCode              As Long            ' Grid Name reference code
    iRowNo                As Integer
End Type



'******************************************************************************
' GCF_Grid_Copy Record Definition
'
'******************************************************************************
Type GCF
    lCode                 As Long            ' Grid copy reference autoincrement
                                             ' code.  This file and gaf are used
                                             ' to retain the copy grid
                                             ' information
    lGpfCode              As Long            ' Grid Position reference
    iSeqNo                As Integer         ' Sequence number to order Type
                                             ' within Sort Order No
    sType                 As String * 1      ' Type (G=Generic, M=DMA, A=MSA,
                                             ' N=State, F=Format, T=Time zone,
                                             ' S=Station, I=Incl/Excl,
                                             ' B=Blackout, V=Vehicle, C=Copy)
    sInclExcl             As String * 1      ' Include or Exlude Format (I or E)
    iIntCode              As Integer         ' Reference code as a function of
                                             ' the selected Category
    lLongCode             As Long            ' Reference code as a function of
                                             ' the selected category
    iCopyCount            As Integer         ' Percentage or ratio value
                                             ' assigned to copy (Category C).
    sUnused               As String * 10
End Type


'Type GCFKEY0
'    lCode                 As Long
'End Type

Type GCFKEY1
    lGpfCode              As Long
    iSeqNo                As Integer
End Type

'******************************************************************************
' GTF_Grid_Type Record Definition
'
'******************************************************************************
Type GTF
    lCode                 As Long            ' Grid Name auto-increment
                                             ' reference code
    lGpfCode              As Long            ' Grid Position (Row) reference
                                             ' code
    sType                 As String * 1      ' Type (N=New, C=Change, M=Model)
    lRafCode              As Long            ' Region Name reference code
                                             ' (gnfType = M)
    sRegionName           As String * 80     ' Region Name (gnfType = N)
    sUnused               As String * 10
End Type


'Type GTFKEY0
'    lCode                 As Long
'End Type

Type GTFKEY1
    lGpfCode              As Long
End Type

'******************************************************************************
' CUF_Copy_Usage Record Definition
'
'******************************************************************************
Type CUF
    lCode                 As Long            ' Copy Inventory Usage reference
                                             ' code
    lCifCode              As Long            ' Copy Inventory reference code
    iAdfCode              As Integer         ' Advertiser reference code
    iMcfCode              As Integer         ' Media Code reference code
    sCifName              As String * 5      ' Copy Inventory name
    lCrfCode(0 To 24)     As Long            ' Copy Rotation Header reference
                                             ' code
    iRotStartDate(0 To 1, 0 To 24) As Integer
    iRotEndDate(0 To 1, 0 To 24)  As Integer
    sUnused               As String * 20
End Type


Type CUFKEY0
    lCode                 As Long
End Type

Type CUFKEY1
    lCifCode              As Long
End Type

Type CUFKEY2
    iAdfCode              As Integer
End Type

Type CUFKEY3
    iMcfCode              As Integer
    sCifName              As String * 5
End Type

'******************************************************************************
' CVF_Copy_Vehicles Record Definition
'
'******************************************************************************
Type CVF
    lCode                 As Long            ' Copy Vehicle reference code
    lCrfCode              As Long            ' Copy rotation reference
    lLkCvfCode            As Long            ' Link to next CVF Reference code
    iVefCode(0 To 99)     As Integer         ' Vehicle reference
    iNextFinal(0 To 99)   As Integer         ' Next copy instruction to assign within final log when vffSyncCopy = "Y"
    iNextPrelim(0 To 99)  As Integer         ' Next copy instruction to assign within preliminary log when vffSyncCopy = "Y"
    sUnused               As String * 20
End Type


'Type CVFKEY0
'    lCode                 As Long
'End Type

'Type CVFKEY1
'    lCrfCode              As Long
'End Type

'Type CVFKEY2
'    lLkCvfCode            As Long
'End Type

Type CIFINFO
    lCifCode As Long
    lFirst As Long
    iAdfCode As Integer
End Type
Type CRFDATES
    lCrfCode As Long
    lStartDate As Long
    lEndDate As Long
    iAdfCode As Integer
    lNext As Long
End Type
Type DATESORT
    lStartDate As Long
    lEndDate As Long
    iAdfCode As Integer
End Type

Dim tmDateSort() As DATESORT

'Copy Inventory
'Dim hmCif As Integer        'Contract file handle
Dim imCifRecLen As Integer  'CHF record length
Dim tmCif As CIF            'CHF record image
Dim tmCifSrchKey0 As LONGKEY0
Dim tmCifSrchKey1 As CIFKEY1

'Copy rotation
'Dim hmCrf As Integer        'Copy rotation file handle
Dim tmCrf As CRF            'CRF record image
Dim tmCrfSrchKey1 As CRFKEY1 'CRF key record image
Dim imCrfRecLen As Integer     'CRF record length

'Instruction
Dim tmCnf As CNF            'CNF record image
Dim tmCnfSrchKey As CNFKEY0  'CNF key record image
'Dim hmCnf As Integer        'CNF Handle
Dim imCnfRecLen As Integer      'CNF record length

'Copy Usage
'Dim hmCif As Integer        'Contract file handle
Dim imCufRecLen As Integer  'CHF record length
Dim tmCuf As CUF            'CHF record image
Dim tmCufSrchKey0 As LONGKEY0
Dim tmCufSrchKey1 As CUFKEY1


Dim tmVsf As VSF
'Comment record-Header/Line
Dim tmCxf As CXF            'CXF record image
Dim tmCxfSrchKey As LONGKEY0  'CXF key record image
Dim hmCxf As Integer        'CXF Handle
Dim imCxfRecLen As Integer      'CXF record length
Public tgCopySdfExtSort() As SDFEXTSORT
Public tgCopySdfExt() As SDFEXT
Public tgCVehicle() As SORTCODE
Public sgCVehicleTag As String

Public tgCopyAdvertiser() As SORTCODE
Public sgCopyAdvertiserTag As String

Type SETINVROTDATE
    lCifCode As Long
    lStartDate As Long
    lEndDate As Long
End Type

'7-6-15 moved from rptgen.bas so some modules can be removed from traffic
Type COPYROTNO
    iRotNo As Integer
    sZone As String * 3
End Type


'*******************************************************
'*                                                     *
'*      Procedure Name:gPopCntrForAASWithRotBox        *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     contract number given adfcode,  *
'*                     or agency or user               *
'*            Note: Code copied from gPopCntrAASBox    *
'*                                                     *
'*******************************************************
Function gPopCntrForAASWithRotBox(frm As Form, ilAAS As Integer, ilAASCode As Integer, slStatus As String, slCntrType As String, ilCurrent As Integer, ilHOType As Integer, ilShow As Integer, lbcRot As control, lbcLocal As control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopCntrForAASWithRotBox (MainForm, ilAAS, ilAASCode, slStatus, slCntrType, ilCurrent, ilHOType, ilShow, lbcRot, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilAAS(I)=0=Obtain Contracts for Specified Advertiser Code (ilAASCode) and User Salesperson (tgUrf(0).iSlfCode) defined as one of the contract salespersons
'                1=Obtain Contracts for Specified Agency Code (ilAASCode) and User Salesperson (tgUrf(0).iSlfCode) defined as one of the contract salespersons
'                2=Obtain Contracts for Specified Salesperson Code (ilAASCode) that matches one of the salespersons defined for the contract
'                3=Obtain Contracts for Specified Vehicle Code (ilAASCode)  and User Salesperson (tgUrf(0).iSlfCode) defined as one of the contract salespersons
'               -1=No selection by advertiser or agency or salesperson
'                Note if tgUrf(0).iSlfCode not specified, then salesperson test is bypassed
'       ilAASCode(I)- Advertiser or Agency code to obtain contracts for (-1 for all advertiser or agency)
'       slStatus (I)- chfStatus value or blank
'                         W=Working; D=Rejected; C=Completed; I=Unapproved; H=Hold; O=Order; G=Approved Hold; N=Approved Order
'                         Multiple status can be specified (WDI)
'                         If H or O or G or N, then only latest shown (Delete <> "Y")
'       slCntrType (I)- chfType value or blank
'                       C=Standard; V=Reservation; T=Remnant; R=DR; Q=PI; S=PSA; M=Promo
'       ilCurrent (I)- 0=Current (Active) (chfDelete <> y); 1=Past and Current (chfDelete <> y); 2=Current(Active) plus all cancel before start (chfDelete <> y); 3=All plus history (any value for chfDelete)
'       ilHOType (I)-  1=H or O only; 2=H or O or G or N (if G or N exists show it over H or O);
'                      3=H or O or G or N or W or C or I (if G or N or W or C or I exists show it over H or O)
'                        Note: G or N can't exist at the same time as W or C or I for an order
'                              G or N or W or C or I CntrRev > 0
'       ilShow(I)- 0=Only show numbers, 1= Show Number and advertiser (test site value) and product and internal comment,....
'                   2=Show Number, Dates, Product and vehicle
'                   3=Show Number, Advertiser, Dates
'                   4=Show Number, Dates
'                   5=Show Number, Advertiser
'       lbcRot(I)- List box of rotation that is to be check to determine if past contracts should be retained
'       lbcLocal (O)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'CHF date/time stamp
    Dim hlChf As Integer        'CHF handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlChf As CHF
    Dim llCntrNo As Long
    Dim ilRevNo As Integer
    Dim ilVerNo As Integer
    Dim slShow As String
    Dim slName As String
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilLoop1 As Integer
    Dim slNameCode As String
    Dim llTodayDate As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slCode As String    'Sales source code number
    Dim hlVef As Integer        'Vef handle
    Dim tlVef As VEF
    Dim ilVefRecLen As Integer     'Record length
    Dim tlVefSrchKey As INTKEY0
    Dim hlVsf As Integer        'Vsf handle
    'Dim tlVsf As VSF
    Dim ilVsfReclen As Integer     'Record length
    Dim tlSrchKey As LONGKEY0
    Dim hlAdf As Integer        'Adf handle
    Dim tlAdf As ADF
    Dim ilAdfRecLen As Integer     'Record length
    Dim tlAdfSrchKey As INTKEY0
    Dim hlSif As Integer        'Vef handle
    Dim tlSif As SIF
    Dim ilSifRecLen As Integer     'Record length
    Dim tlSifSrchKey As LONGKEY0
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim ilOffSet As Integer
    Dim tlVsf As VSF
    Dim llLen As Long
    Dim ilOper As Integer
    Dim slStr As String
    Dim ilSlfCode As Integer
    Dim ilVefCode As Integer
    Dim ilTestCntrNo As Integer
    Dim tlChfAdvtExt As CHFADVTEXT
    Dim slCntrStatus As String
    Dim ilSortCode As Integer
    Dim slHOStatus As String
    Dim blAdServerOnly As Boolean
    
    If slStatus = "" Then
        slCntrStatus = "WCIDHO"
    Else
        slCntrStatus = slStatus
    End If
    slHOStatus = ""
    If ilHOType = 1 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "H"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "O"
        End If
    ElseIf ilHOType = 2 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "GH"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "NO"
        End If
    ElseIf ilHOType = 3 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "GH"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "NO"
        End If
        If (InStr(1, slCntrStatus, "H", 1) <> 0) Or (InStr(1, slCntrStatus, "O", 1) <> 0) Then
            slHOStatus = slHOStatus & "WCI"
        End If
    End If
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Chf.Btr") & Trim$(str$(ilAASCode)) & Trim$(slCntrStatus) & Trim$(slCntrType) & Trim$(str$(ilCurrent)) & Trim$(str$(ilHOType)) & Trim$(str$(ilShow))
    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopCntrForAASWithRotBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
        End If
    End If
    gPopCntrForAASWithRotBox = CP_MSG_POPREQ
    llTodayDate = gDateValue(gNow())
    'gObtainVehComboList
    hlChf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrOpen):" & "Chf.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tlChf) 'btrRecordLength(hlChf)  'Get and save record length
    hlVsf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrOpen):" & "Vsf.Btr", frm
    On Error GoTo 0
    ilVsfReclen = Len(tmVsf) 'btrRecordLength(hlSlf)  'Get and save record length
    hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrOpen):" & "Vef.Btr", frm
    On Error GoTo 0
    ilVefRecLen = Len(tlVef) 'btrRecordLength(hlSlf)  'Get and save record length
    hlSif = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrOpen):" & "Sif.Btr", frm
    On Error GoTo 0
    ilSifRecLen = Len(tlSif) 'btrRecordLength(hlSif)  'Get and save record length
    hlAdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrOpen):" & "Adf.Btr", frm
    On Error GoTo 0
    hmCxf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrOpen):" & "Adf.Btr", frm
    On Error GoTo 0
    ilAdfRecLen = Len(tlAdf) 'btrRecordLength(hlSlf)  'Get and save record length
    tlAdf.iCode = 0
    ilSortCode = 0
    ReDim tlSortCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
    lbcLocal.Clear
    slSortCodeTag = slStamp
    ilExtLen = Len(tlChfAdvtExt)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlChf) 'Obtain number of records
    btrExtClear hlChf   'Clear any previous extend operation
    If (ilAAS = 0) And (ilAASCode > 0) Then
        tlIntTypeBuff.iType = ilAASCode
        ilRet = btrGetEqual(hlChf, tlChf, ilRecLen, tlIntTypeBuff, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Else
        ilRet = btrGetFirst(hlChf, tlChf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    End If
    If (ilRet = BTRV_ERR_END_OF_FILE) Or (ilRet = BTRV_ERR_KEY_NOT_FOUND) Then
        ilRet = btrClose(hmCxf)
        btrDestroy hmCxf
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        ilRet = btrClose(hlVsf)
        btrDestroy hlVsf
        ilRet = btrClose(hlSif)
        btrDestroy hlSif
        ilRet = btrClose(hlVef)
        btrDestroy hlVef
        ilRet = btrClose(hlChf)
        btrDestroy hlChf
        Exit Function
    Else
        On Error GoTo gPopCntrForAASWithRotBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrGetFirst):" & "Chf.Btr", frm
        On Error GoTo 0
    End If
    Call btrExtSetBounds(hlChf, llNoRec, -1, "UC", "CHFADVTEXTPK", CHFADVTEXTPK) 'Set extract limits (all records)
    ilSlfCode = tgUrf(0).iSlfCode
    If ilAAS = 1 Then
        If ilAASCode > 0 Then
            tlIntTypeBuff.iType = ilAASCode
            ilOffSet = gFieldOffset("Chf", "ChfAgfCode")
            If (slCntrType = "") And (ilCurrent = 3) Then
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
            Else
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
            End If
        End If
    ElseIf ilAAS = 2 Then
        If ilAASCode > 0 Then
            ilSlfCode = ilAASCode
        End If
    ElseIf ilAAS = 3 Then
        If ilAASCode > 0 Then
            ilVefCode = ilAASCode
        End If
   ElseIf ilAAS = 0 Then
        If ilAASCode > 0 Then
            tlIntTypeBuff.iType = ilAASCode
            ilOffSet = gFieldOffset("Chf", "ChfAdfCode")
            If ilCurrent <> 3 Then
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
            Else
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
            End If
        End If
    End If
    tlCharTypeBuff.sType = "Y"
    ilOffSet = gFieldOffset("Chf", "ChfDelete")
    'If selecting by advertiser- bypass slCntrType and slCntrStatus Test until get contract for speed
    'If ((slCntrStatus = "") And (slCntrType = "")) Or ((ilAAS = 0) And (ilAASCode > 0)) Then
    If (slCntrType = "") Or ((ilAAS = 0) And (ilAASCode > 0)) Then
        If ilCurrent <> 3 Then
            ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        End If
    Else
        If ilCurrent <> 3 Then
            ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
        End If
        'If slCntrStatus <> "" Then
        '    ilOper = BTRV_EXT_OR
        '    slStr = slCntrStatus
        '    Do While slStr <> ""
        '        If Len(slStr) = 1 Then
        '            If slCntrType <> "" Then
        '                ilOper = BTRV_EXT_AND
        '            Else
        '                ilOper = BTRV_EXT_LAST_TERM
        '            End If
        '        End If
        '        tlCharTypeBuff.sType = Left$(slStr, 1)
        '        ilOffset = gFieldOffset("Chf", "ChfStatus")
        '        ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, ilOper, tlCharTypeBuff, 1)
        '        slStr = Mid$(slStr, 2)
        '    Loop
        'End If
        If slCntrType <> "" Then
            ilOper = BTRV_EXT_OR
            slStr = slCntrType
            Do While slStr <> ""
                If Len(slStr) = 1 Then
                    ilOper = BTRV_EXT_LAST_TERM
                End If
                tlCharTypeBuff.sType = Left$(slStr, 1)
                ilOffSet = gFieldOffset("Chf", "ChfType")
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, ilOper, tlCharTypeBuff, 1)
                slStr = Mid$(slStr, 2)
            Loop
        End If
    End If
    ilOffSet = gFieldOffset("Chf", "ChfCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract iCode field
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfCntrNo")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract Contract number
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfExtRevNo")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfCntRevNo")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfType")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfAdfCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfProduct")
    ilRet = btrExtAddField(hlChf, ilOffSet, 35) 'Extract Product
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfAgfCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfSlfCode1")
    ilRet = btrExtAddField(hlChf, ilOffSet, 20) 'Extract salesperson code
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfMnfDemo1")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract salesperson code
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfCxfInt")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfPropVer")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract end date
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfStatus")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfMnfPotnType")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract SellNet
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfStartDate")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfEndDate")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract end date
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfVefCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfSifCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0

     '8-21-05 add pct of trade to array
    ilOffSet = gFieldOffset("Chf", "ChfPctTrade")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'pct trade
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    
    '7/12/10
    ilOffSet = gFieldOffset("Chf", "ChfCBSOrder")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0

    '2/24/12
    ilOffSet = gFieldOffset("Chf", "ChfBillCycle")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    
    '1/22/21 - was missing
    ilOffSet = gFieldOffset("Chf", "ChfSource")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1)   'Extract ChfSource
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    
    '1/22/21 - Copy: Exclude CPM only contracts
    ilOffSet = gFieldOffset("Chf", "ChfAdServerDefined")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1)   'Extract AdServerDefined
    On Error GoTo gPopCntrForAASWithRotBoxErr
    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0

    'ilRet = btrExtGetNextExt(hlChf)    'Extract record
    ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        On Error GoTo gPopCntrForAASWithRotBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrExtGetNextExt):" & "Chf.Btr", frm
        On Error GoTo 0
        ilExtLen = Len(tlChfAdvtExt)  'Extract operation record size
        'ilRet = btrExtGetFirst(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilFound = True
            '10660
            blAdServerOnly = False
            '1/22/21: Bypass Podcast CPM only Contracts; Dick: CPM Only Contract: chfPodCPMDefined = Y and no clf records specified for the contract.
            '2/10/21: Bypass "Ad Server Only" Contracts:  Where AdServerDefined="Y" and 0 Clf records
            If tlChfAdvtExt.sAdServerDefined = "Y" Then
                ilFound = gExistClf(tlChfAdvtExt.lCode)          'False if No CLF records found
                blAdServerOnly = Not ilFound
            End If

            If (ilAAS = 3) And (ilVefCode <> -1) Then   '-1 = All Vehicles
                ilFound = False
                If tlChfAdvtExt.lVefCode > 0 Then
                    If tlChfAdvtExt.lVefCode = ilVefCode Then
                        ilFound = True
                    End If
                ElseIf tlChfAdvtExt.lVefCode < 0 Then
                    tlSrchKey.lCode = -tlChfAdvtExt.lVefCode
                    ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    On Error GoTo gPopCntrForAASWithRotBoxErr
                    gBtrvErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrGetEqual): Vsf.Btr", frm
                    On Error GoTo 0
                    For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                        If tmVsf.iFSCode(ilLoop) > 0 Then
                            If tmVsf.iFSCode(ilLoop) = ilVefCode Then
                                ilFound = True
                                Exit For
                            End If
                            If tlVef.iCode <> tmVsf.iFSCode(ilLoop) Then
                                tlVefSrchKey.iCode = tmVsf.iFSCode(ilLoop)
                                ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo gPopCntrForAASWithRotBoxErr
                                gCPErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrGetEqual: Vef)", frm
                                On Error GoTo 0
                            End If
                            If tlVef.sType = "V" Then
                                If tlVef.iCode <> ilVefCode Then
                                    tlSrchKey.lCode = tlVef.lVsfCode
                                    ilRet = btrGetEqual(hlVsf, tlVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    For ilLoop1 = LBound(tlVsf.iFSCode) To UBound(tlVsf.iFSCode) Step 1
                                        If tlVsf.iFSCode(ilLoop1) > 0 Then
                                            If tlVsf.iFSCode(ilLoop1) = ilVefCode Then
                                                ilFound = True
                                                Exit For
                                            End If
                                        End If
                                    Next ilLoop1
                                Else
                                    ilFound = True
                                End If
                            End If
                        End If
                    Next ilLoop
                Else
                    ilFound = True  'all vehicles
                End If
            End If
            ilTestCntrNo = False
            'For Proposal CntRevNo = 0; For Orders CntRevNo >= 0 (for W, C, I CntRevNo > 0)
            If (tlChfAdvtExt.iCntRevNo = 0) And ((tlChfAdvtExt.sStatus <> "H") And (tlChfAdvtExt.sStatus <> "O") And (tlChfAdvtExt.sStatus <> "G") And (tlChfAdvtExt.sStatus <> "N")) Then  'Proposal
                If (InStr(1, slCntrStatus, tlChfAdvtExt.sStatus) = 0) Then
                    ilFound = False
                End If
            Else    'Order
                If (InStr(1, slHOStatus, tlChfAdvtExt.sStatus) <> 0) Then
                    If (ilHOType = 2) Or (ilHOType = 3) Then
                        ilTestCntrNo = True
                    End If
                Else
                    ilFound = False
                End If
            End If
            If slCntrType <> "" Then
                If InStr(1, slCntrType, tlChfAdvtExt.sType) = 0 Then
                    ilFound = False
                End If
            End If
            '7/12/10: Bypas CBS contracts
            If ilFound Then
                If tlChfAdvtExt.sCBSOrder = "Y" Then
                    ilFound = False
                End If
            End If
            If ilFound Then
                ilFound = mTestChfAdvtExt(frm, ilSlfCode, tlChfAdvtExt, hlVsf, ilCurrent, lbcRot)
            End If
            If ilFound Then
                slStr = Trim$(str$(99999999 - tlChfAdvtExt.lCntrNo))
                Do While Len(slStr) < 8
                    slStr = "0" & slStr
                Loop
                slName = slStr
                slStr = Trim$(str$(999 - tlChfAdvtExt.iCntRevNo))
                Do While Len(slStr) < 3
                    slStr = "0" & slStr
                Loop
                slName = slName & "|" & slStr & "|"
                If (tlChfAdvtExt.sStatus = "W") Or (tlChfAdvtExt.sStatus = "C") Or (tlChfAdvtExt.sStatus = "I") Then
                    'Add Potential
                    If tlChfAdvtExt.iMnfPotnType > 0 Then
                        slStr = " "
                    Else
                        slStr = "~"
                    End If
                    slName = slName & slStr & "|"
                Else
                    slName = slName & " |"
                End If
                slStr = Trim$(str$(999 - tlChfAdvtExt.iPropVer))
                Do While Len(slStr) < 3
                    slStr = "0" & slStr
                Loop
                slName = slName & slStr & "|"
                slName = slName & tlChfAdvtExt.sStatus & "|"
                If ilShow = 0 Then
                    'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo)) & " V" & Trim$(Str$(tlChfAdvtExt.iPropVer))
                ElseIf ilShow = 2 Then
                    'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                    gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                    gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slEndDate
                    slName = slName & ": " & slStartDate & "-" & slEndDate
                    If tgSpf.sUseProdSptScr = "P" Then  'Short Title
                        tlSif.sName = ""
                        If (tlChfAdvtExt.lSifCode > 0) Then
                            tlSifSrchKey.lCode = tlChfAdvtExt.lSifCode
                            ilRet = btrGetEqual(hlSif, tlSif, ilSifRecLen, tlSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet <> BTRV_ERR_NONE Then
                                tlSif.sName = ""
                            End If
                        End If
                        slName = slName & " " & Trim$(tlSif.sName)
                    Else
                        slName = slName & " " & Trim$(tlChfAdvtExt.sProduct)
                    End If
                    If tlChfAdvtExt.lVefCode > 0 Then
                        If tlChfAdvtExt.lVefCode <> tlVef.iCode Then
                            tlVefSrchKey.iCode = tlChfAdvtExt.lVefCode
                            ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            On Error GoTo gPopCntrForAASWithRotBoxErr
                            gCPErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrGetEqual: Adf)", frm
                            On Error GoTo 0
                        End If
                        slName = slName & " " & Trim$(tlVef.sName)
                    End If
                ElseIf ilShow = 3 Then
                    'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                    If tlChfAdvtExt.iAdfCode <> tlAdf.iCode Then
                        tlAdfSrchKey.iCode = tlChfAdvtExt.iAdfCode
                        ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        On Error GoTo gPopCntrForAASWithRotBoxErr
                        gCPErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrGetEqual: Adf)", frm
                        On Error GoTo 0
                    End If
                    If (tlAdf.sBillAgyDir = "D") And (Trim$(tlAdf.sAddrID) <> "") Then
                        slName = slName & " " & Trim$(tlAdf.sName) & ", " & Trim$(tlAdf.sAddrID) '& "/" & Trim$(tlChfAdvtExt.sProduct)
                    Else
                        slName = slName & " " & Trim$(tlAdf.sName) '& "/" & Trim$(tlChfAdvtExt.sProduct)
                    End If
                    gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                    gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slEndDate
                    slName = slName & " " & slStartDate & "-" & slEndDate
                ElseIf ilShow = 4 Then
                    'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                    gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                    gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slEndDate
                    slName = slName & " " & slStartDate & "-" & slEndDate
                ElseIf ilShow = 5 Then
                    'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                    If tlChfAdvtExt.iAdfCode <> tlAdf.iCode Then
                        tlAdfSrchKey.iCode = tlChfAdvtExt.iAdfCode
                        ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        On Error GoTo gPopCntrForAASWithRotBoxErr
                        gCPErrorMsg ilRet, "gPopCntrForAASWithRotBox (btrGetEqual: Adf)", frm
                        On Error GoTo 0
                    End If
                    If (tlAdf.sBillAgyDir = "D") And (Trim$(tlAdf.sAddrID) <> "") Then
                        slName = slName & " " & Trim$(tlAdf.sName) & ", " & Trim$(tlAdf.sAddrID) '& "/" & Trim$(tlChfAdvtExt.sProduct)
                    Else
                        slName = slName & " " & Trim$(tlAdf.sName) '& "/" & Trim$(tlChfAdvtExt.sProduct)
                    End If
                Else
                    'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo)) & " V" & Trim$(Str$(tlChfAdvtExt.iPropVer))
                    Select Case tlChfAdvtExt.sStatus
                        Case "W"
                            If tlChfAdvtExt.iCntRevNo > 0 Then
                                slStr = "Rev Working"
                            Else
                                slStr = "Working"
                            End If
                        Case "D"
                            slStr = "Rejected"
                        Case "C"
                            If tlChfAdvtExt.iCntRevNo > 0 Then
                                slStr = "Rev Completed"
                            Else
                                slStr = "Completed"
                            End If
                        Case "I"
                            If tlChfAdvtExt.iCntRevNo > 0 Then
                                slStr = "Rev Unapproved"
                            Else
                                slStr = "Unapproved"
                            End If
                        Case "G"
                            slStr = "Approved Hold"
                        Case "N"
                            slStr = "Approved Order"
                        Case "H"
                            slStr = "Hold"
                        Case "O"
                            slStr = "Order"
                    End Select
                    slName = slName & " " & slStr
                    slName = slName & " " & Trim$(tlChfAdvtExt.sProduct)
                    tmCxfSrchKey.lCode = tlChfAdvtExt.lCxfInt
                    If tmCxfSrchKey.lCode <> 0 Then
                        tmCxf.sComment = ""
                        imCxfRecLen = Len(tmCxf) '5027
                        ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            'If tmCxf.iStrLen > 0 Then
                            '    If tmCxf.iStrLen < 40 Then
                            '        slName = slName & " " & Trim$(Left$(tmCxf.sComment, tmCxf.iStrLen))
                            '    Else
                                    slName = slName & " " & gStripChr0(Left$(tmCxf.sComment, 40))
                            '    End If
                            'End If
                        End If
                    End If
                End If
                ilFound = False
                If ilTestCntrNo Then
                    For ilLoop = 0 To ilSortCode - 1 Step 1 'lbcMster.ListCount - 1 Step 1
                        slNameCode = tlSortCode(ilLoop).sKey    'lbcMster.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 1, "\", slNameCode)
                        ilRet = gParseItem(slNameCode, 1, "|", slCode)
                        llCntrNo = 99999999 - CLng(slCode)
                        If llCntrNo = tlChfAdvtExt.lCntrNo Then
                            ilRet = gParseItem(slNameCode, 2, "|", slCode)
                            ilRevNo = 999 - CLng(slCode)
                            If tlChfAdvtExt.iCntRevNo > ilRevNo Then
                                'Replace
                                'lbcMster.RemoveItem ilLoop
                                'llLen = llLen - Len(slNameCode)
                                tlSortCode(ilLoop).sKey = slNameCode
                                ilFound = True
                            Else
                                'Leave
                                ilFound = True
                            End If
                            Exit For
                        End If
                    Next ilLoop
                End If
                '1/22/21: Dont show Ad Server ONLY Contracts
                If Not ilFound And Not blAdServerOnly Then
                    slName = slName & "\" & Trim$(str$(tlChfAdvtExt.lCode))
                    'If Not gOkAddStrToListBox(slName, llLen, True) Then
                    '    Exit Do
                    'End If
                    'lbcMster.AddItem slName
                    tlSortCode(ilSortCode).sKey = slName
                    If ilSortCode >= UBound(tlSortCode) Then
                        ReDim Preserve tlSortCode(0 To UBound(tlSortCode) + 100) As SORTCODE
                    End If
                    ilSortCode = ilSortCode + 1
                End If
            End If
            ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
            Loop
        Loop
        'Sort then output new headers and lines
        ReDim Preserve tlSortCode(0 To ilSortCode) As SORTCODE
        If UBound(tlSortCode) - 1 > 0 Then
            ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
        End If
        llLen = 0
        For ilLoop = 0 To UBound(tlSortCode) - 1 Step 1 'lbcMster.ListCount - 1 Step 1
            slNameCode = tlSortCode(ilLoop).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 1, "|", slCode)
            llCntrNo = 99999999 - CLng(slCode)
            slShow = Trim$(str$(llCntrNo))
            ilRet = gParseItem(slName, 2, "|", slCode)
            ilRevNo = 999 - CLng(slCode)
            ilRet = gParseItem(slName, 4, "|", slCode)
            ilVerNo = 999 - CLng(slCode)
            ilRet = gParseItem(slName, 5, "|", slCode)
            If (slCode = "W") Or (slCode = "C") Or (slCode = "I") Or (slCode = "D") Then
                If ilRevNo > 0 Then
                    slShow = slShow & " R" & Trim$(str$(ilRevNo))
                Else
                    slShow = slShow & " V" & Trim$(str$(ilVerNo))
                End If
            Else
                slShow = slShow & " R" & Trim$(str$(ilRevNo))
            End If
            If ilShow = 0 Then      'Number only
            ElseIf ilShow = 2 Then  'Number, Dates, Product, Vehicle
                'Other fields
                ilRet = gParseItem(slName, 6, "|", slCode)
                slShow = slShow & " " & slCode
            ElseIf ilShow = 3 Then  'Number, Advertiser, Dates
                'Other fields
                ilRet = gParseItem(slName, 6, "|", slCode)
                slShow = slShow & " " & slCode
            ElseIf ilShow = 4 Then  'Number, Dates
                'Other fields
                ilRet = gParseItem(slName, 6, "|", slCode)
                slShow = slShow & " " & slCode
            ElseIf ilShow = 5 Then  'Number, Advertiser
                'Other fields
                ilRet = gParseItem(slName, 6, "|", slCode)
                slShow = slShow & " " & slCode
            Else                    'Number, Product, Internal comment
                'Potential
                ilRet = gParseItem(slName, 3, "|", slCode)
                If (Trim$(slCode) <> "") And (slCode <> "~") Then
                    slShow = slShow & " " & slCode
                End If
                'Other fields
                ilRet = gParseItem(slName, 6, "|", slCode)
                slShow = slShow & " " & slCode
            End If
            If Not gOkAddStrToListBox(slShow, llLen, True) Then
                Exit For
            End If
            lbcLocal.AddItem slShow  'Add ID to list box
        Next ilLoop
    End If
    ilRet = btrClose(hmCxf)
    btrDestroy hmCxf
    ilRet = btrClose(hlAdf)
    btrDestroy hlAdf
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    ilRet = btrClose(hlSif)
    btrDestroy hlSif
    ilRet = btrClose(hlVef)
    btrDestroy hlVef
    ilRet = btrClose(hlChf)
    btrDestroy hlChf
    Exit Function
gPopCntrForAASWithRotBoxErr:
    ilRet = btrClose(hmCxf)
    btrDestroy hmCxf
    ilRet = btrClose(hlAdf)
    btrDestroy hlAdf
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    ilRet = btrClose(hlSif)
    btrDestroy hlSif
    ilRet = btrClose(hlVef)
    btrDestroy hlVef
    ilRet = btrClose(hlChf)
    btrDestroy hlChf
    gDbg_HandleError "CopySubs: gPopCntrForAASWithRotBox"
'    gPopCntrForAASWithRotBox = CP_MSG_NOSHOW
'    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestChfAdvtExt                 *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if Contract is OK to be    *
'*                     viewed by the user              *
'*                                                     *
'*******************************************************
Private Function mTestChfAdvtExt(frm As Form, ilSlfCode As Integer, tlChfAdvtExt As CHFADVTEXT, hlVsf As Integer, ilCurrent As Integer, lbcRot As control) As Integer

    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim slStartDate As String
    'Dim tlVsf As VSF
    Dim ilVsfReclen As Integer     'Record length
    Dim tlSrchKey As LONGKEY0
    Dim llTodayDate As Long
    Dim ilUser As Integer
    Dim ilRet As Integer
    Dim ilSlf As Integer
    Dim ilRot As Integer
    Dim slNameCode As String
    Dim slCode As String
    llTodayDate = gDateValue(gNow())
    ilVsfReclen = Len(tmVsf) 'btrRecordLength(hlSlf)  'Get and save record length
    If (tgUrf(0).iCode = 1) Or (tgUrf(0).iCode = 2) Then
        ilFound = True
    Else
        ilFound = False
        For ilLoop = LBound(tgUrf) To UBound(tgUrf) Step 1
            If (tgUrf(ilLoop).iVefCode = 0) Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
    End If
    If Not ilFound Then
        If tlChfAdvtExt.lVefCode > 0 Then
            For ilLoop = LBound(tgUrf) To UBound(tgUrf) Step 1
                If (tgUrf(ilLoop).iVefCode = tlChfAdvtExt.lVefCode) Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
        ElseIf tlChfAdvtExt.lVefCode < 0 Then
            tlSrchKey.lCode = -tlChfAdvtExt.lVefCode
            ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            On Error GoTo mTestChfAdvtExtErr
            gBtrvErrorMsg ilRet, "mPopCntrBoxRec (btrGetEqual): Vsf.Btr", frm
            On Error GoTo 0
            For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                If tmVsf.iFSCode(ilLoop) > 0 Then
                    For ilUser = LBound(tgUrf) To UBound(tgUrf) Step 1
                        If (tgUrf(ilUser).iVefCode = tmVsf.iFSCode(ilLoop)) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilUser
                    If ilFound Then
                        Exit For
                    End If
                End If
            Next ilLoop
        Else    'All vehicles
            ilFound = True
            If igUserByVeh Then 'Test lines as user defined by vehicle after contract added
            End If
        End If
    End If
    If ilFound Then
        If ilSlfCode > 0 Then
            ilFound = False
            For ilSlf = LBound(tlChfAdvtExt.iSlfCode) To UBound(tlChfAdvtExt.iSlfCode) Step 1
                If tlChfAdvtExt.iSlfCode(ilSlf) <> 0 Then
                    If ilSlfCode = tlChfAdvtExt.iSlfCode(ilSlf) Then
                        ilFound = True
                        Exit For
                    End If
                End If
            Next ilSlf
        End If
    End If
    If ilFound Then
        If ilCurrent = 0 Then   'Current
            gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slDate
            If gDateValue(slDate) < llTodayDate Then
                ilFound = False
                For ilRot = 0 To lbcRot.ListCount - 1 Step 1
                    slNameCode = lbcRot.List(ilRot)
                    ilRet = gParseItem(slNameCode, 1, "|", slCode)
                    If Val(slCode) = tlChfAdvtExt.lCntrNo Then
                        ilFound = True
                        Exit For
                    End If
                Next ilRot
            End If
        ElseIf ilCurrent = 2 Then
            If (tlChfAdvtExt.iStartDate(0) <> 0) Or (tlChfAdvtExt.iStartDate(1) <> 0) Or (tlChfAdvtExt.iEndDate(0) <> 0) Or (tlChfAdvtExt.iEndDate(1) <> 0) Then
                gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slDate
                If gDateValue(slStartDate) <= gDateValue(slDate) Then
                    If gDateValue(slDate) < llTodayDate Then
                        ilFound = False
                    End If
                End If
            End If
        End If
    End If
    mTestChfAdvtExt = ilFound
    Exit Function
mTestChfAdvtExtErr:
    mTestChfAdvtExt = False
    Exit Function
End Function

Public Sub gBuildCifArray(hlCif As Integer, ilMcfCode As Integer, tlCifInfo() As CIFINFO)
    Dim ilRet As Integer
    Dim llCif As Long
    
    ReDim tlCifInfo(0 To 0) As CIFINFO
    tlCifInfo(0).lFirst = -1
    imCifRecLen = Len(tmCif)
    tmCifSrchKey1.iMcfCode = ilMcfCode
    tmCifSrchKey1.sName = ""
    tmCifSrchKey1.sCut = ""
    ilRet = btrGetGreaterOrEqual(hlCif, tmCif, imCifRecLen, tmCifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmCif.iMcfCode = ilMcfCode)
        If igSetCopyTerminated Then
            Exit Sub
        End If
        If (tmCif.sPurged = "A") Then
            llCif = UBound(tlCifInfo)
            tlCifInfo(llCif).lCifCode = tmCif.lCode
            tlCifInfo(llCif).iAdfCode = tmCif.iAdfCode
            tlCifInfo(llCif).lFirst = -1
            ReDim Preserve tlCifInfo(0 To llCif + 1) As CIFINFO
        End If
        ilRet = btrGetNext(hlCif, tmCif, imCifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    llCif = UBound(tlCifInfo)
    If llCif > 0 Then
        ArraySortTyp fnAV(tlCifInfo(), 0), UBound(tlCifInfo), 0, LenB(tlCifInfo(0)), 0, -2, 0
    End If

End Sub

Sub gBuildRotDates(hlCrf As Integer, hlCnf As Integer, slActiveDate As String, tlCifInfo() As CIFINFO, tlCrfDates() As CRFDATES)
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim slType As String
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim llRecPos As Long
    Dim llCrf As Long
    Dim llNext As Long
    Dim llCif As Long
    Dim blFound As Boolean
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record

    imCrfRecLen = Len(tmCrf)
    imCnfRecLen = Len(tmCnf)
    btrExtClear hlCrf   'Clear any previous extend operation
    ilExtLen = Len(tmCrf)  'Extract operation record size
    slType = ""
    tmCrfSrchKey1.sRotType = slType
    tmCrfSrchKey1.iEtfCode = 0
    tmCrfSrchKey1.iEnfCode = 0
    tmCrfSrchKey1.iAdfCode = 0
    tmCrfSrchKey1.lChfCode = 0
    tmCrfSrchKey1.lFsfCode = 0
    tmCrfSrchKey1.iVefCode = 0
    tmCrfSrchKey1.iRotNo = 32000
    ilRet = btrGetGreaterOrEqual(hlCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If (ilRet <> BTRV_ERR_END_OF_FILE) Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hlCrf, llNoRec, -1, "UC", "CRF", "") 'Set extract limits (all records)
        gPackDate slActiveDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Crf", "CrfEndDate")
        ilRet = btrExtAddLogicConst(hlCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        tlCharTypeBuff.sType = "D"
        ilOffSet = gFieldOffset("Crf", "CrfState")
        ilRet = btrExtAddLogicConst(hlCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        ilOffSet = 0
        ilRet = btrExtAddField(hlCrf, ilOffSet, ilExtLen)  'Extract start/end time, and days
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        'ilRet = btrExtGetNextExt(hmClf)    'Extract record
        ilRet = btrExtGetNext(hlCrf, tmCrf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            'ilRet = btrExtGetFirst(hmClf, tlClfExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlCrf, tmCrf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                If igSetCopyTerminated Then
                    Exit Sub
                End If
                tmCnfSrchKey.lCrfCode = tmCrf.lCode
                tmCnfSrchKey.iInstrNo = 0
                ilRet = btrGetGreaterOrEqual(hlCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
                    If igSetCopyTerminated Then
                        Exit Sub
                    End If
                    llCif = mBinarySearchCifInfo(tmCnf.lCifCode, tlCifInfo())
                    If llCif >= 0 Then
                        blFound = False
                        llNext = tlCifInfo(llCif).lFirst
                        Do While llNext >= 0
                            If tlCrfDates(llNext).lCrfCode = tmCrf.lCode Then
                                blFound = True
                                Exit Do
                            End If
                            llNext = tlCrfDates(llNext).lNext
                        Loop
                        If Not blFound Then
                            llCrf = UBound(tlCrfDates)
                            If tlCifInfo(llCif).lFirst = -1 Then
                                llNext = -1
                            Else
                                llNext = tlCifInfo(llCif).lFirst
                            End If
                            tlCifInfo(llCif).lFirst = llCrf
                            tlCrfDates(llCrf).lCrfCode = tmCrf.lCode
                            gUnpackDateLong tmCrf.iStartDate(0), tmCrf.iStartDate(1), tlCrfDates(llCrf).lStartDate
                            gUnpackDateLong tmCrf.iEndDate(0), tmCrf.iEndDate(1), tlCrfDates(llCrf).lEndDate
                            If tmCrf.iBkoutInstAdfCode <= 0 Then
                                tlCrfDates(llCrf).iAdfCode = tmCrf.iAdfCode
                            Else
                                tlCrfDates(llCrf).iAdfCode = tmCrf.iBkoutInstAdfCode
                            End If
                            tlCrfDates(llCrf).lNext = llNext
                            ReDim Preserve tlCrfDates(0 To llCrf + 1) As CRFDATES
                            tlCrfDates(UBound(tlCrfDates)).lNext = -1
                        End If
                    End If
                    ilRet = btrGetNext(hlCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                ilRet = btrExtGetNext(hlCrf, tmCrf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlCrf, tmCrf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
        btrExtClear hlCrf   'Clear any previous extend operation
    End If
End Sub

Private Function mBinarySearchCifInfo(llCode As Long, tlCifInfo() As CIFINFO) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    llMin = LBound(tlCifInfo)
    llMax = UBound(tlCifInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tlCifInfo(llMiddle).lCifCode Then
            'found the match
            mBinarySearchCifInfo = llMiddle
            Exit Function
        ElseIf llCode < tlCifInfo(llMiddle).lCifCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    mBinarySearchCifInfo = -1
End Function

Sub gCreateCufDates(hlCif As Integer, hlCuf As Integer, tlCifInfo() As CIFINFO, tlCrfDates() As CRFDATES)
    Dim llCif As Long
    Dim ilRet As Integer
    Dim llNext As Long
    Dim ilIndex As Integer
    
    imCufRecLen = Len(tmCuf)
    imCifRecLen = Len(tmCif)
    For llCif = 0 To UBound(tlCifInfo) - 1 Step 1
        tmCufSrchKey1.lCifCode = tlCifInfo(llCif).lCifCode
        ilRet = btrGetEqual(hlCuf, tmCuf, imCufRecLen, tmCufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmCuf.lCifCode = tlCifInfo(llCif).lCifCode)
            ilRet = btrDelete(hlCuf)
            tmCufSrchKey1.lCifCode = tlCifInfo(llCif).lCifCode
            ilRet = btrGetEqual(hlCuf, tmCuf, imCufRecLen, tmCufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Loop
        tmCifSrchKey0.lCode = tlCifInfo(llCif).lCifCode
        ilRet = btrGetEqual(hlCif, tmCif, imCifRecLen, tmCifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            tmCuf.lCode = 0
            tmCuf.lCifCode = tlCifInfo(llCif).lCifCode
            tmCuf.iAdfCode = tmCif.iAdfCode
            tmCuf.iMcfCode = tmCif.iMcfCode
            tmCuf.sCifName = tmCif.sName
            For ilIndex = LBound(tmCuf.lCrfCode) To UBound(tmCuf.lCrfCode) Step 1
                tmCuf.lCrfCode(ilIndex) = 0
                gPackDate "", tmCuf.iRotStartDate(0, ilIndex), tmCuf.iRotStartDate(1, ilIndex)
                gPackDate "", tmCuf.iRotEndDate(0, ilIndex), tmCuf.iRotEndDate(1, ilIndex)
            Next ilIndex
            ilIndex = LBound(tmCuf.lCrfCode)
            llNext = tlCifInfo(llCif).lFirst
            Do While llNext >= 0
                tmCuf.lCrfCode(ilIndex) = tlCrfDates(llNext).lCrfCode
                gPackDateLong tlCrfDates(llNext).lStartDate, tmCuf.iRotStartDate(0, ilIndex), tmCuf.iRotStartDate(1, ilIndex)
                gPackDateLong tlCrfDates(llNext).lEndDate, tmCuf.iRotEndDate(0, ilIndex), tmCuf.iRotEndDate(1, ilIndex)
                If ilIndex = UBound(tmCuf.lCrfCode) Then
                    ilRet = btrInsert(hlCuf, tmCuf, imCufRecLen, INDEXKEY0)
                    tmCuf.lCode = 0
                    For ilIndex = LBound(tmCuf.lCrfCode) To UBound(tmCuf.lCrfCode) Step 1
                        tmCuf.lCrfCode(ilIndex) = 0
                        gPackDate "", tmCuf.iRotStartDate(0, ilIndex), tmCuf.iRotStartDate(1, ilIndex)
                        gPackDate "", tmCuf.iRotEndDate(0, ilIndex), tmCuf.iRotEndDate(1, ilIndex)
                    Next ilIndex
                    ilIndex = 0
                Else
                    ilIndex = ilIndex + 1
                End If
                llNext = tlCrfDates(llNext).lNext
            Loop
            If ilIndex > 0 Then
                ilRet = btrInsert(hlCuf, tmCuf, imCufRecLen, INDEXKEY0)
            End If
        End If
    Next llCif
    
End Sub


Sub gUpdateCifDates(hlCif As Integer, tlCifInfo() As CIFINFO, tlCrfDates() As CRFDATES, llChgCount As Long, llTotalCount As Long, ilLogChg As Boolean)
    Dim llCif As Long
    Dim ilRet As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llNext As Long
    Dim llDate As Long
    Dim llRotStartDate As Long
    Dim llRotEndDate As Long
    Dim ilAdfCode As Integer
    Dim blSameAdfCode As Boolean
    Dim slNowDate As String
    
    
    llChgCount = 0
    llTotalCount = 0
    imCifRecLen = Len(tmCif)
    For llCif = 0 To UBound(tlCifInfo) - 1 Step 1
        ReDim tmDateSort(0 To 0) As DATESORT
        ilAdfCode = -1
        blSameAdfCode = True
        llNext = tlCifInfo(llCif).lFirst
        Do While llNext >= 0
            llDate = UBound(tmDateSort)
            tmDateSort(llDate).lStartDate = tlCrfDates(llNext).lStartDate
            tmDateSort(llDate).lEndDate = tlCrfDates(llNext).lEndDate
            tmDateSort(llDate).iAdfCode = tlCrfDates(llNext).iAdfCode
            If ilAdfCode = -1 Then
                ilAdfCode = tlCrfDates(llNext).iAdfCode
            Else
                If ilAdfCode <> tlCrfDates(llNext).iAdfCode Then
                    blSameAdfCode = False
                End If
            End If
            ReDim Preserve tmDateSort(0 To llDate + 1) As DATESORT
            llNext = tlCrfDates(llNext).lNext
        Loop
        llDate = UBound(tmDateSort)
        If llDate > 0 Then
            ArraySortTyp fnAV(tmDateSort(), 0), UBound(tmDateSort), 0, LenB(tmDateSort(0)), 0, -2, 0
        End If
        llStartDate = 0
        llEndDate = 0
        For llDate = 0 To UBound(tmDateSort) - 1 Step 1
            If tlCifInfo(llCif).iAdfCode = tmDateSort(llDate).iAdfCode Then
                If llStartDate = 0 Then
                    llStartDate = tmDateSort(llDate).lStartDate
                    llEndDate = tmDateSort(llDate).lEndDate
                Else
                    'Test for gap
                    If (llEndDate + 1 < tmDateSort(llDate).lStartDate) And (Not blSameAdfCode) Then
                        llStartDate = tmDateSort(llDate).lStartDate
                        llEndDate = tmDateSort(llDate).lEndDate
                    Else
                        If llEndDate < tmDateSort(llDate).lEndDate Then
                            llEndDate = tmDateSort(llDate).lEndDate
                        End If
                    End If
                End If
            End If
        Next llDate
        'Update CIF
        If llStartDate > 0 Then
            llTotalCount = llTotalCount + 1
            tmCifSrchKey0.lCode = tlCifInfo(llCif).lCifCode
            ilRet = btrGetEqual(hlCif, tmCif, imCifRecLen, tmCifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                gUnpackDateLong tmCif.iRotStartDate(0), tmCif.iRotStartDate(1), llRotStartDate
                gUnpackDateLong tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), llRotEndDate
                If (llRotStartDate <> llStartDate) Or (llRotEndDate <> llEndDate) Then
                    llChgCount = llChgCount + 1
                    If ilLogChg Then
                        gLogMsg "Copy Item " & Trim$(tmCif.sName) & " From Date Range " & Format(llRotStartDate, "m/d/yy") & "-" & Format(llRotEndDate, "m/d/yy") & " to " & Format(llStartDate, "m/d/yy") & "-" & Format(llEndDate, "m/d/yy"), "SetCopyInventoryDates.txt", False
                    End If
                    gPackDateLong llStartDate, tmCif.iRotStartDate(0), tmCif.iRotStartDate(1)
                    gPackDateLong llEndDate, tmCif.iRotEndDate(0), tmCif.iRotEndDate(1)
                    
                    'D.S. 06/10/13
                    slNowDate = Format$(gNow(), "m/d/yy")
                    If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
                            tmCif.sCleared = "N"
                            gPackDate slNowDate, tmCif.iInvSentDate(0), tmCif.iInvSentDate(1)
                    End If
                    
                    ilRet = btrUpdate(hlCif, tmCif, imCifRecLen)
                End If
            End If
        Else
            'Only adjust if End Date in future.  The rotation must have been erased
            tmCifSrchKey0.lCode = tlCifInfo(llCif).lCifCode
            ilRet = btrGetEqual(hlCif, tmCif, imCifRecLen, tmCifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                gUnpackDateLong tmCif.iRotStartDate(0), tmCif.iRotStartDate(1), llRotStartDate
                gUnpackDateLong tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), llRotEndDate
                If (llRotEndDate > gDateValue(Format(gNow(), "m/d/yy"))) Then
                    llChgCount = llChgCount + 1
                    If ilLogChg Then
                        gLogMsg "Copy Item " & Trim$(tmCif.sName) & " From Date Range " & Format(llRotStartDate, "m/d/yy") & "-" & Format(llRotEndDate, "m/d/yy") & " to " & Format(llStartDate, "m/d/yy") & "-" & Format(llEndDate, "m/d/yy"), "SetCopyInventoryDates.txt", False
                    End If
                    gPackDateLong llStartDate, tmCif.iRotStartDate(0), tmCif.iRotStartDate(1)
                    gPackDateLong llEndDate, tmCif.iRotEndDate(0), tmCif.iRotEndDate(1)
                    
                    'D.S. 06/10/13
                    slNowDate = Format$(gNow(), "m/d/yy")
                    If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
                            tmCif.sCleared = "N"
                            gPackDate slNowDate, tmCif.iInvSentDate(0), tmCif.iInvSentDate(1)
                    End If
                    
                    ilRet = btrUpdate(hlCif, tmCif, imCifRecLen)
                End If
            End If
        End If
    Next llCif
End Sub
