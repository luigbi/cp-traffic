Attribute VB_Name = "INVEXPORTSUBS"
'******************************************************************************************
'                       Invoice Export Subroutines (Site feature)
'******************************************************************************************

' Proprietary Software, Do not copy
'
' File Name: InvExportSubs.Bas
'
' Release: 7.0
'
' Description:
'   This file contains the Invoice support functions for the Invoice Export Site feature
Option Explicit
Option Compare Text

Private tmInvExport_Header As INVEXPORT_HEADER
Private tmInvExport_Spot As INVEXPORT_SPOT
Private tmInvExport_NTR As INVEXPORT_NTR

'
'           All the invoices header information has been obtained and built for the printed output
'           Build into array for the invoice export feature
'           5-17-17 create separate module due to out of room, will not compile
Public Sub gInvExportHeaderBuildInfo(tlIvr As IVR, tlAdf As ADF, tlagf As AGF, tlSlf As SLF, tlSof As SOF, slTerms As String)

    Dim llDate As Long
    Dim ilRemainder As Integer              '5-31-17
    Dim slStripCents As String
    Dim slStr As String

    If ((Asc(tgSpf.sUsingFeatures6) And INVEXPORTPARAMETERS) = INVEXPORTPARAMETERS) And ((Invoice!rbcType(INVGEN_Final).Value) Or (Invoice!rbcType(INVGEN_Reprint).Value)) Then           '5-11-17, create header only if site feature to export Inv set, and finals or reprint
        tmInvExport_Header.sInvNo = str$(tlIvr.lInvNo)
        gUnpackDateLong tlIvr.iInvDate(0), tlIvr.iInvDate(1), llDate
        tmInvExport_Header.sInvStartDate = Format$(llDate, "ddddd")     '5-24-17 this was changed to end date of billing period
        
        tmInvExport_Header.sCntrNo = str$(tgChfInv.lCntrNo)
        
        gUnpackDateLong tgChfInv.iStartDate(0), tgChfInv.iStartDate(1), llDate
        tmInvExport_Header.sCntStartDate = Format$(llDate, "ddddd")
        
        gUnpackDateLong tgChfInv.iEndDate(0), tgChfInv.iEndDate(1), llDate
        tmInvExport_Header.sCntEndDate = Format$(llDate, "ddddd")
        
        tmInvExport_Header.sCashTrade = tlIvr.sCashTrade
        tmInvExport_Header.sAdvName = tlAdf.sName
        tmInvExport_Header.sProduct = tgChfInv.sProduct
        tmInvExport_Header.sSlspName = Trim$(tlSlf.sFirstName) + " " + Trim$(tlSlf.sLastName)
        tmInvExport_Header.sSlspOffice = Trim$(tlSof.sName)
        tmInvExport_Header.sAdfCode = str$(tlAdf.iCode)
        tmInvExport_Header.sSlfCode = str$(tlSlf.iCode)
        tmInvExport_Header.sTerms = slTerms

        If tgChfInv.iAgfCode > 0 Then               'agy commissionable
            '5-31-17 store commission % as whole number without decimal if its a whole percentage
            ilRemainder = tlagf.iComm Mod 100
            If ilRemainder = 0 Then         'strip off the pennies if whole number
                slStripCents = Trim$(gIntToStrDec(tlagf.iComm, 2))
                slStr = slStr & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3))
            Else
                slStr = slStr & Trim$(gIntToStrDec(tlagf.iComm, 2))
            End If

            'tmInvExport_Header.sAgyComm = str$(tlagf.iComm)
            tmInvExport_Header.sAgyComm = Trim$(slStr)
            tmInvExport_Header.sPayee = tlagf.sName
            tmInvExport_Header.sAgfCode = str$(tlagf.iCode)
        Else                                        'direct
            tmInvExport_Header.sAgyComm = "0"
            tmInvExport_Header.sPayee = tlAdf.sName
            tmInvExport_Header.sAgfCode = "0"
        End If
    Else
        Exit Sub
    End If
    Exit Sub
End Sub

'               For Invoice Export Feature, create 2 files in the export folder:
'               One for spot data; one for NTR data
'               '5-11-17
Public Sub gInvExportFileNameCreate(hlInvExportSpots As Integer, hlInvExportNTR As Integer, hlMsg As Integer, llGenDate As Long, llGenTime As Long)
    '5-11-17 fields for Inv Export Parameter feature
    Dim slExportSpotName As String      'csv file for spots
    Dim slExportNTRName As String       'csv file for ntr
    Dim slClientName As String          'abbrev client name
    Dim ilRet As Integer
    Dim ilGenDate(0 To 1) As Integer    'gen date to be used in filename
    Dim slGenDate As String
    Dim slGenTime As String             'gen time to be used in filename
    Dim slTemp As String
    Dim slTemp2 As String
    Dim slOneChar As String * 1
    Dim slTempStr As String
    Dim tlMnf As MNF            'MNF record image
    Dim tlMnfSrchKey As INTKEY0 'MNF key record image
    Dim hlMnf As Integer        'MNF Handle
    Dim slDelimStr As String
    Dim slDelimiter As String

   'determine the selection options:  Airtime ntr, etc and the types of invoicing (prelims, reprints, finals)
    If Invoice!ckcType(INVTYPE_Commercial).Value = vbChecked Then
        slTemp = "Air Time"
    End If
    If Invoice!ckcType(INVTYPE_PrintRep).Value = vbChecked Then
        If Trim$(slTemp) = "" Then
            slTemp = "Print Rep"
        Else
            slTemp = slTemp & ",Print Rep"
        End If
    End If
    If Invoice!ckcType(INVTYPE_Installment).Value = vbChecked Then
        If Trim$(slTemp) = "" Then
            slTemp = "Installment"
        Else
            slTemp = slTemp & ",Installment"
        End If
    End If
    If Invoice!ckcType(INVTYPE_NTR).Value = vbChecked Then
        If Trim$(slTemp) = "" Then
            slTemp = "NTR"
        Else
            slTemp = slTemp & ",NTR"
        End If
    End If
    If Invoice!ckcType(INVTYPE_GenRepAR).Value = vbChecked Then
        If Trim$(slTemp) = "" Then
            slTemp = "Gen Rep"
        Else
            slTemp = slTemp & ",Gen Rep"
        End If
    End If
    If Invoice!rbcType(INVGEN_Preliminary).Value Then
        slTemp = slTemp & ",Prelim"
    End If
    If Invoice!rbcType(INVGEN_Final).Value Then
        slTemp = slTemp & ",Final"
    End If
    If Invoice!rbcType(INVGEN_Reprint).Value Then
        slTemp = slTemp & ",Reprint"
    End If
    If Invoice!rbcType(INVGEN_Archive).Value Then
        slTemp = slTemp & ",Archive"
    End If
    If ((Asc(tgSpf.sUsingFeatures6) And INVEXPORTPARAMETERS) = INVEXPORTPARAMETERS) Then
        If ((Invoice!rbcType(INVGEN_Final).Value) Or (Invoice!rbcType(INVGEN_Reprint).Value)) Then
            gLogMsg "Generating Invoice Export for option: " & Trim$(slTemp), "InvoiceExport.Txt", False
        Else
            gLogMsg "Invoice Export Not created for this option:  " & Trim$(slTemp) & sgCR & sgLF, "InvoiceExport.Txt", False
        End If
    End If
    '5-11-17 if exporting invoices by spots and ntr csv files, open file handle
    If ((Asc(tgSpf.sUsingFeatures6) And INVEXPORTPARAMETERS) = INVEXPORTPARAMETERS) And ((Invoice!rbcType(INVGEN_Final).Value) Or (Invoice!rbcType(INVGEN_Reprint).Value)) Then           'create file only site feature to export Inv set, and finals or reprint
        'determine filename
        'use short client name vs long client name
        slClientName = Trim$(tgSpf.sGClient)
        If tgSpf.iMnfClientAbbr > 0 Then
            hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hlMnf)
                btrDestroy hlMnf
            Else
                tlMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
                ilRet = btrGetEqual(hlMnf, tlMnf, Len(tlMnf), tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    slClientName = Trim$(tlMnf.sName)
                End If
            End If
        End If
        
        gPackDateLong llGenDate, ilGenDate(0), ilGenDate(1)
        mFormatGenDateTimeToStr ilGenDate(), slGenDate, llGenTime, slGenTime        'remove slashes and colons from date & time fields
        If Invoice!rbcType(INVGEN_Reprint).Value Then                'if reprint, designate that in the filename
            slTemp = "R"
        Else
            slTemp = ""
        End If
        'month year has commas and blanks, filter out
        slTemp2 = ""
        For ilRet = 1 To Len(sgInvMonthYear)
            slOneChar = Mid(sgInvMonthYear, ilRet, 1)
            If (Asc(slOneChar) <> Asc(" ")) And (Asc(slOneChar) <> Asc(",")) Then
                slTemp2 = slTemp2 & slOneChar
            End If
        Next
        
        slTempStr = Trim$(sgExportPath)                'insure path name correct, may need a slash at end
        If right(slTempStr, 1) <> "\" Then
            slTempStr = slTempStr & "\"
        End If

        slExportSpotName = Trim$(slTempStr) & slClientName & "_" & slTemp2 & Trim$(slTemp) & "_" & "Spots" & "_" & slGenDate & "_" & slGenTime & ".csv"
        slExportNTRName = Trim$(slTempStr) & slClientName & "_" & slTemp2 & Trim$(slTemp) & "_" & "NTR" & "_" & slGenDate & "_" & slGenTime & ".csv"
        
        slTempStr = "Record Type,Invoice #,Invoice Date,Contract #,Contract Start Date,Contract End Date,Cash/Trade,Agency Commission %,Payee,Advertiser,"
        slTempStr = slTempStr & "Product,Salesperson,Sales Office,Agency Internal Code,Advertiser Internal Code,Salesperson Internal Code,Invoice Terms,"
        
        'Date: 03/17/2020 used the selected delimiter from SAF table; default is comma delimited
        slDelimiter = IIF(Trim$(tgSaf(0).sInvExpDelimiter) = "", Chr(44), Trim(tgSaf(0).sInvExpDelimiter))
        'replace comma delimiter with selected delimiter from SAD table (e.g. 2 pipe characters "||")
        If slDelimiter <> "," Then
            slTempStr = Replace(slTempStr, Chr(44), slDelimiter)
        End If
        
        ilRet = 0
        On Error GoTo mOpenInvExpFileErr:
        hlInvExportSpots = FreeFile
        Open slExportSpotName For Output As hlInvExportSpots
        If ilRet <> 0 Then
            Print #hlMsg, "Unable to open " & slExportSpotName
            Close #hlInvExportSpots
            Screen.MousePointer = vbDefault
            MsgBox "Open Error #" & str$(err.Numner) & slExportSpotName, vbOKOnly, "Open Error"
            gLogMsg "Invoice Export File Open Error: mInvExportFileNameCreate for " & Trim$(slExportSpotName) & sgCR & sgLF, "InvoiceExport.Txt", False
            Exit Sub
        Else
            'Date: 03/17/2020 use selected delimiter from SAF table, comma is default delimiter
            If slDelimiter <> "," Then
                slDelimStr = Replace("Reconciliation Amount,Week of,Vehicle,Length,Ordered Days,# Spots Per Week,Line #,Date Aired,Time Aired,Aired Status,MG/Bonus Status,MG Missed Date,Price,Copy", Chr(44), slDelimiter)
                'slDelimStr = slTempStr & Replace(slDelimStr, Chr(44), slDelimiter)
                Print #hlInvExportSpots, slTempStr & slDelimStr '"Reconciliation Amount,Week of,Vehicle,Length,Ordered Days,# Spots Per Week,Line #,Date Aired,Time Aired,Aired Status,MG/Bonus Status,MG Missed Date,Price,Copy"
            Else
                Print #hlInvExportSpots, slTempStr & "Reconciliation Amount,Week of,Vehicle,Length,Ordered Days,# Spots Per Week,Line #,Date Aired,Time Aired,Aired Status,MG/Bonus Status,MG Missed Date,Price,Copy"
            End If
            If ilRet <> 0 Then
                Close #hlInvExportSpots
                Screen.MousePointer = vbDefault
                MsgBox "Open Error #" & str$(err.Numner) & slExportNTRName, vbOKOnly, "Writing Spot Header Error"
                gLogMsg "Invoice Export File Write Error: mInvExportFileNameCreate for " & Trim$(slExportSpotName) & sgCR & sgLF, "InvoiceExport.Txt", False
                Exit Sub
            End If
            gLogMsg "Invoice Export Spot FileName: " & Trim$(slExportSpotName), "InvoiceExport.Txt", False
        End If
        ilRet = 0
        On Error GoTo mOpenInvExpFileErr:
        
        hlInvExportNTR = FreeFile
        Open slExportNTRName For Output As hlInvExportNTR
        If ilRet <> 0 Then
            Print #hlMsg, "Unable to open " & slExportNTRName
            Close #hlInvExportNTR
            Screen.MousePointer = vbDefault
            MsgBox "Open Error #" & str$(err.Numner) & slExportNTRName, vbOKOnly, "Open Error"
            gLogMsg "Invoice Export File Open Error: mInvExportFileNameCreate for " & Trim$(slExportNTRName) & sgCR & sgLF, "InvoiceExport.Txt", False
            Exit Sub
        Else
            'Date: 03/17/2020 use selected delimiter from SAF table, comma is default delimiter
            If slDelimiter <> "," Then
                slDelimStr = "NTR Date,Vehicle,Description,Total Gross,Total Net"
                slTempStr = slTempStr & Replace(slDelimStr, Chr(44), slDelimiter)
                Print #hlInvExportNTR, slTempStr  '"Reconciliation Amount,Week of,Vehicle,Length,Ordered Days,# Spots Per Week,Line #,Date Aired,Time Aired,Aired Status,MG/Bonus Status,MG Missed Date,Price,Copy"
            Else
                Print #hlInvExportNTR, slTempStr & "NTR Date,Vehicle,Description,Total Gross,Total Net"
            End If
            
            If ilRet <> 0 Then
                Close #hlInvExportNTR
                Screen.MousePointer = vbDefault
                MsgBox "Open Error #" & str$(err.Numner) & slExportNTRName, vbOKOnly, "Writing NTR Header Error"
                gLogMsg "Invoice Export File Write Error: mInvExportFileNameCreate for " & Trim$(slExportNTRName) & sgCR & sgLF, "InvoiceExport.Txt", False
               Exit Sub
            End If
            gLogMsg "Invoice Export NTR FileName: " & Trim$(slExportNTRName), "InvoiceExport.Txt", False
        End If
        On Error GoTo 0
    End If
    Exit Sub
mOpenInvExpFileErr:
    ilRet = 1
    Resume Next
End Sub

'
'           Gather info to create  spot or NTR record for the Invoice Export feature
'           '5-12-17
Public Sub gInvExport_GatherDetail(ilCombineAirAndNTR As Integer, hlInvExportSpots As Integer, hlInvExportNTR As Integer, hlSbf As Integer, blItsREP As Boolean, tlIvr As IVR, tlSmf As SMF, tlSbf As SBF)
    Dim llDate As Long
    Dim ilDay As Integer
    Dim ilTemp As Integer

    If ((Asc(tgSpf.sUsingFeatures6) And INVEXPORTPARAMETERS) = INVEXPORTPARAMETERS) And ((Invoice!rbcType(INVGEN_Final).Value) Or (Invoice!rbcType(INVGEN_Reprint).Value)) Then           'only process if site feature to export Inv set, and finals or reprint
        If tlIvr.iType = IVRTYPE_Spot Then             'spot type
            'Reconciliation
            tmInvExport_Spot.sReconciliationAmt = ""
            If InStr(tlIvr.sRAmount, "Aired Past Mid") > 0 Or InStr(tlIvr.sRAmount, "BB") = 0 Then  'determine if spot aired past midnight and need to show comment in footer
              
               tmInvExport_Spot.sReconciliationAmt = ""                                                                    'look for presence of BB.  if found, do nothing; otherwise its a money field
            Else
                If (InStr(tlIvr.sRAmount, ".") <> 0) Then        'found reconciliation amt
                    'is it a .00?
                    If InStr(tlIvr.sRAmount, ".00") = 0 Then       '.00 not found, either bonus, mg, etc or has pennies
                        'use rate as is; leave decimal and pennies
                        tmInvExport_Spot.sReconciliationAmt = tlIvr.sRAmount
                    Else
                        'strip pennies
                        ilTemp = InStr(tlIvr.sRAmount, ".00")         'look for starting position of decimal point
                        If ilTemp = 1 Then                  '.00 starts at position 1
                            tmInvExport_Spot.sReconciliationAmt = ""
                        Else
                            tmInvExport_Spot.sReconciliationAmt = Trim$(Mid(tlIvr.sRAmount, 1, ilTemp - 1))
                        End If
                    End If
                Else
                    tmInvExport_Spot.sReconciliationAmt = Trim$(tlIvr.sRAmount)         'its a money field
                End If
            End If
            tmInvExport_Spot.sLen = str$(tlIvr.iLen)
            tmInvExport_Spot.sOrderedDays = tlIvr.sODays
            tmInvExport_Spot.sSpotsPerWk = str$(tlIvr.lONoSpots)
            tmInvExport_Spot.sLine = str$(tlIvr.iLineNo)
            
            tmInvExport_Spot.sDateAired = Mid$(tlIvr.sADayDate, 4)
            'find the start of week based on air date for week of
            llDate = gDateValue(tmInvExport_Spot.sDateAired)
            ilDay = gWeekDayLong(llDate)
            Do While ilDay <> 0           'backup MF to monday
                llDate = llDate - 1
                ilDay = gWeekDayLong(llDate)
            Loop
            tmInvExport_Spot.sWeekOf = Format$(llDate, "ddddd")
            tmInvExport_Spot.sTimeAired = tlIvr.sATime
            
             If (InStr(tlIvr.sARate, ".") <> 0) Then        'found spot cost (vs rate of bonus, mg, n/c, etc)
                'is it a .00?
                If InStr(tlIvr.sARate, ".00") = 0 Then       '.00 not found, either bonus, mg, etc or has pennies
                    'use rate as is; leave decimal and pennies
                    tmInvExport_Spot.sSpotPrice = tlIvr.sARate
                Else
                    'strip pennies
                    ilTemp = InStr(tlIvr.sARate, ".00")         'look for starting position of decimal point
                    tmInvExport_Spot.sSpotPrice = Trim$(Mid(tlIvr.sARate, 1, ilTemp - 1))
                End If
            Else
                tmInvExport_Spot.sSpotPrice = Trim$(tlIvr.sARate)
            End If
            'tmInvExport_Spot.sSpotPrice = tlIvr.sARate
            tmInvExport_Spot.sCopy = tlIvr.sACopy(1)
            tmInvExport_Spot.sMGBonus = ""
            tmInvExport_Spot.sMGMissedDate = ""
            tmInvExport_Spot.sAirStatus = "A"               'default to aired,unless Bonus (extra), makegood on  or missed from
            If Trim$(tlIvr.sAVehName) = "" Then
                tmInvExport_Spot.sVehicle = Trim$(tlIvr.sOVehName)
            Else
                tmInvExport_Spot.sVehicle = Trim$(tlIvr.sAVehName)
            End If
            
            If (InStr(tlIvr.sARate, "Bonus") > 0 And InStr(tlIvr.sORate, "Bonus") = 0) Then          'aired rate is bonus, ordered rate is not bonus, using $ spot to fill
                tmInvExport_Spot.sMGBonus = "B"                     'bonus vs mg
                tmInvExport_Spot.sVehicle = tlIvr.sRRemark          'vehicle bonus spot aired on
           ElseIf InStr(tlIvr.sARate, "Bonus") > 0 And Trim$(tlIvr.sRRemark) <> "" Then      'aired rate is Bonus and remark field has something in it; must be a fill spot
                tmInvExport_Spot.sMGBonus = "B"
                tmInvExport_Spot.sVehicle = Trim$(tlIvr.sRRemark)       'bonus defined line spot,used  as a fill
           ElseIf InStr(tlIvr.sRRemark, "MG for") > 0 Then
                tmInvExport_Spot.sMGBonus = "M"                     'mg vs bonus
                tmInvExport_Spot.sReconciliationAmt = tlIvr.sRAmount
                gUnpackDateLong tlSmf.iMissedDate(0), tlSmf.iMissedDate(1), llDate
                tmInvExport_Spot.sMGMissedDate = Format$(llDate, "ddddd")
            End If
            If InStr(tlIvr.sRRemark, "Missed, MG") > 0 Then
                tmInvExport_Spot.sAirStatus = "M"
                tmInvExport_Spot.sReconciliationAmt = tlIvr.sRAmount
            ElseIf InStr(tlIvr.sRRemark, "Cancel") > 0 Then
                tmInvExport_Spot.sAirStatus = "M"
                tmInvExport_Spot.sReconciliationAmt = tlIvr.sRAmount
            ElseIf InStr(tlIvr.sRRemark, "Missed") > 0 Then
                tmInvExport_Spot.sAirStatus = "M"
                tmInvExport_Spot.sReconciliationAmt = tlIvr.sRAmount
            ElseIf Trim$(tlIvr.sRRemark) = "" Then
                tmInvExport_Spot.sAirStatus = "A"
            End If
            mInvExport_WriteType 1, hlInvExportSpots, hlInvExportNTR
        ElseIf (ilCombineAirAndNTR) Then                ' combined air time and ntr
            If Invoice!ckcType(INVTYPE_Commercial).Value = vbChecked And Invoice!ckcType(INVTYPE_NTR).Value = vbChecked Then        'processing both air time and ntr together
                'ntr detail = 6
                'TTP 11051 - Invoice Export: NTR records not being included on NTR output
                'If tlIvr.iType = IVRTYPE_TotalNTR Then
                If tlIvr.iType = IVRTYPE_NTR Then
                    gInvExport_GatherNTRDetail hlInvExportSpots, hlInvExportNTR, hlSbf, tlIvr, tlSbf
                End If
            ElseIf Invoice!ckcType(INVTYPE_Commercial).Value = vbUnchecked And Invoice!ckcType(INVTYPE_NTR).Value = vbChecked Then     'air time turned off, doing ntr only
                'TTP 10517 - Invoices: if "ad server" option is not checked on, and "commercial and NTR" invoices are set to be separate, the air time portion of the invoice does not print
                'ntr detail = 6
                If tlIvr.iType = IVRTYPE_NTR Then
                    gInvExport_GatherNTRDetail hlInvExportSpots, hlInvExportNTR, hlSbf, tlIvr, tlSbf
                End If
            End If
        Else                                'not combined air time and ntr
            '5-31-17 rep comes thru here and is flagged as type 2; ignore them
            If blItsREP Then
                blItsREP = blItsREP
            Else
                'TTP 10517 - Invoices: if "ad server" option is not checked on, and "commercial and NTR" invoices are set to be separate, the air time portion of the invoice does not print
                'NTR detail = 6
                If tlIvr.iType = IVRTYPE_NTR Then
                    gInvExport_GatherNTRDetail hlInvExportSpots, hlInvExportNTR, hlSbf, tlIvr, tlSbf
                End If
            End If
        End If
    End If
End Sub

'               write Inv export spot record  or NTR record
'               5-12-17
'               mInvExport_writeType
'               <input> ilWhichType : 1 = spot, 2 = NTR
Public Sub mInvExport_WriteType(ilWhichType As Integer, hlInvExportSpots As Integer, hlInvExportNTR As Integer)
    Dim slStr As String
    Dim slDelimiter As String

    'Date: 03/14/2020 used the selected delimiter from SAF table; default is comma delimited
    slDelimiter = IIF(Trim$(tgSaf(0).sInvExpDelimiter) = "", Chr(44), Trim(tgSaf(0).sInvExpDelimiter))

    If slDelimiter = "," Then
        slStr = Trim$(tmInvExport_Header.sInvNo)
        slStr = slStr & "," & Trim$(tmInvExport_Header.sInvStartDate)
        slStr = slStr & "," & Trim$(tmInvExport_Header.sCntrNo)
        slStr = slStr & "," & Trim$(tmInvExport_Header.sCntStartDate)
        slStr = slStr & "," & Trim$(tmInvExport_Header.sCntEndDate)
        slStr = slStr & "," & """" & Trim$(tmInvExport_Header.sCashTrade) & """"
        slStr = slStr & "," & Trim$(tmInvExport_Header.sAgyComm)
        slStr = slStr & "," & """" & Trim$(tmInvExport_Header.sPayee) & """"
        slStr = slStr & "," & """" & Trim$(tmInvExport_Header.sAdvName) & """"
        slStr = slStr & "," & """" & Trim$(tmInvExport_Header.sProduct) & """"
        slStr = slStr & "," & """" & Trim$(tmInvExport_Header.sSlspName) & """"
        slStr = slStr & "," & """" & Trim$(tmInvExport_Header.sSlspOffice) & """"
        slStr = slStr & "," & Trim$(tmInvExport_Header.sAgfCode)
        slStr = slStr & "," & Trim$(tmInvExport_Header.sAdfCode)
        slStr = slStr & "," & Trim$(tmInvExport_Header.sSlfCode)
        slStr = slStr & "," & """" & Trim$(tmInvExport_Header.sTerms) & """"
        
        If ilWhichType = 1 Then         'spots
            slStr = """" & "Spot" & """" & "," & slStr
            slStr = slStr & "," & Trim$(tmInvExport_Spot.sReconciliationAmt)
            slStr = slStr & "," & Trim$(tmInvExport_Spot.sWeekOf)
            slStr = slStr & "," & """" & Trim$(tmInvExport_Spot.sVehicle) & """"
            slStr = slStr & "," & Trim$(tmInvExport_Spot.sLen)
            slStr = slStr & "," & """" & Trim$(tmInvExport_Spot.sOrderedDays) & """"
            slStr = slStr & "," & Trim$(tmInvExport_Spot.sSpotsPerWk)
            slStr = slStr & "," & Trim$(tmInvExport_Spot.sLine)
            slStr = slStr & "," & Trim$(tmInvExport_Spot.sDateAired)
            slStr = slStr & "," & Trim$(tmInvExport_Spot.sTimeAired)
            slStr = slStr & "," & """" & Trim$(tmInvExport_Spot.sAirStatus) & """"
            slStr = slStr & "," & """" & Trim$(tmInvExport_Spot.sMGBonus) & """"
            slStr = slStr & "," & Trim$(tmInvExport_Spot.sMGMissedDate)
            slStr = slStr & "," & Trim$(tmInvExport_Spot.sSpotPrice)
            slStr = slStr & "," & """" & Trim$(tmInvExport_Spot.sCopy) & """"
            
            'replace comma delimiter with selected delimiter from SAD table (e.g. 2 pipe characters "||")
'                If slDelimiter <> "," Then
'                    slStr = Replace(slStr, Chr(44), slDelimiter)
'                End If
            Print #hlInvExportSpots, slStr
            
        Else
            slStr = """" & "NTR" & """" & "," & slStr
            slStr = slStr & "," & Trim$(tmInvExport_NTR.sNTRDate)
            slStr = slStr & "," & """" & Trim$(tmInvExport_NTR.sVehicle) & """"
            slStr = slStr & "," & """" & Trim$(tmInvExport_NTR.sDescription) & """"
            slStr = slStr & "," & Trim$(tmInvExport_NTR.sGross)
            slStr = slStr & "," & Trim$(tmInvExport_NTR.sNet)
            
            'replace comma delimiter with selected delimiter from SAD table (e.g. 2 pipe characters "||")
'                If slDelimiter <> "," Then
'                    slStr = Replace(slStr, Chr(44), slDelimiter)
'                End If
            Print #hlInvExportNTR, slStr
        End If
    Else
        slStr = Trim$(tmInvExport_Header.sInvNo)
        slStr = slStr & slDelimiter & Trim$(tmInvExport_Header.sInvStartDate)
        slStr = slStr & slDelimiter & Trim$(tmInvExport_Header.sCntrNo)
        slStr = slStr & slDelimiter & Trim$(tmInvExport_Header.sCntStartDate)
        slStr = slStr & slDelimiter & Trim$(tmInvExport_Header.sCntEndDate)
        slStr = slStr & slDelimiter & """" & Trim$(tmInvExport_Header.sCashTrade) & """"
        slStr = slStr & slDelimiter & Trim$(tmInvExport_Header.sAgyComm)
        slStr = slStr & slDelimiter & """" & Trim$(tmInvExport_Header.sPayee) & """"
        slStr = slStr & slDelimiter & """" & Trim$(tmInvExport_Header.sAdvName) & """"
        slStr = slStr & slDelimiter & """" & Trim$(tmInvExport_Header.sProduct) & """"
        slStr = slStr & slDelimiter & """" & Trim$(tmInvExport_Header.sSlspName) & """"
        slStr = slStr & slDelimiter & """" & Trim$(tmInvExport_Header.sSlspOffice) & """"
        slStr = slStr & slDelimiter & Trim$(tmInvExport_Header.sAgfCode)
        slStr = slStr & slDelimiter & Trim$(tmInvExport_Header.sAdfCode)
        slStr = slStr & slDelimiter & Trim$(tmInvExport_Header.sSlfCode)
        slStr = slStr & slDelimiter & """" & Trim$(tmInvExport_Header.sTerms) & """"
        
        If ilWhichType = 1 Then         'spots
            slStr = """" & "Spot" & """" & slDelimiter & slStr
            slStr = slStr & slDelimiter & Trim$(tmInvExport_Spot.sReconciliationAmt)
            slStr = slStr & slDelimiter & Trim$(tmInvExport_Spot.sWeekOf)
            slStr = slStr & slDelimiter & """" & Trim$(tmInvExport_Spot.sVehicle) & """"
            slStr = slStr & slDelimiter & Trim$(tmInvExport_Spot.sLen)
            slStr = slStr & slDelimiter & """" & Trim$(tmInvExport_Spot.sOrderedDays) & """"
            slStr = slStr & slDelimiter & Trim$(tmInvExport_Spot.sSpotsPerWk)
            slStr = slStr & slDelimiter & Trim$(tmInvExport_Spot.sLine)
            slStr = slStr & slDelimiter & Trim$(tmInvExport_Spot.sDateAired)
            slStr = slStr & slDelimiter & Trim$(tmInvExport_Spot.sTimeAired)
            slStr = slStr & slDelimiter & """" & Trim$(tmInvExport_Spot.sAirStatus) & """"
            slStr = slStr & slDelimiter & """" & Trim$(tmInvExport_Spot.sMGBonus) & """"
            slStr = slStr & slDelimiter & Trim$(tmInvExport_Spot.sMGMissedDate)
            slStr = slStr & slDelimiter & Trim$(tmInvExport_Spot.sSpotPrice)
            slStr = slStr & slDelimiter & """" & Trim$(tmInvExport_Spot.sCopy) & """"
            
            Print #hlInvExportSpots, slStr
        Else
            slStr = """" & "NTR" & """" & slDelimiter & slStr
            slStr = slStr & slDelimiter & Trim$(tmInvExport_NTR.sNTRDate)
            slStr = slStr & slDelimiter & """" & Trim$(tmInvExport_NTR.sVehicle) & """"
            slStr = slStr & slDelimiter & """" & Trim$(tmInvExport_NTR.sDescription) & """"
            slStr = slStr & slDelimiter & Trim$(tmInvExport_NTR.sGross)
            slStr = slStr & slDelimiter & Trim$(tmInvExport_NTR.sNet)
            
            Print #hlInvExportNTR, slStr
        End If
    End If
End Sub

'
'               mInvExport_GatherNTRDetail - obtain all the data required for
'               an NTR export record
'
Public Sub gInvExport_GatherNTRDetail(hlInvExportSpots As Integer, hlInvExportNTR As Integer, hlSbf As Integer, tlIvr As IVR, tlSbf As SBF)
    Dim ilRet As Integer
    Dim llDate As Long
    Dim llNTRGross As Long
    Dim llNTRNet As Long
    Dim tlSbfSrchKey1 As LONGKEY0
    Dim ilRemainder As Integer          '5-31-17
    Dim slStripCents As String
    Dim slStr As String
    
    If ((Asc(tgSpf.sUsingFeatures6) And INVEXPORTPARAMETERS) = INVEXPORTPARAMETERS) Then
        'obtain the sbf record for description and dates
        tlSbfSrchKey1.lCode = tlIvr.lSpotKeyNo
        ilRet = btrGetEqual(hlSbf, tlSbf, Len(tlSbf), tlSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            gUnpackDateLong tlSbf.iDate(0), tlSbf.iDate(1), llDate
            tmInvExport_NTR.sNTRDate = Format$(llDate, "ddddd")
            tmInvExport_NTR.sVehicle = tlIvr.sOVehName
            tmInvExport_NTR.sDescription = tlSbf.sDescr
            
            llNTRGross = tlSbf.lGross * tlSbf.iNoItems
            'calc agy comm
            If tlSbf.sAgyComm = "Y" Then          'agy commissionable
                llNTRNet = (llNTRGross * CDbl(10000 - tlIvr.iPctComm)) / 10000      'round for proper decimals due to multiplication of agy comm
                '7-21-17 air time may have had commission; then some NTR has commission, some not.  re-establish agy comm just in case there are differences
                ilRemainder = tlIvr.iPctComm Mod 100
                If ilRemainder = 0 Then         'strip off the pennies if whole number
                    slStripCents = Trim$(gIntToStrDec(tlIvr.iPctComm, 2))
                    slStr = slStr & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3))
                Else
                    slStr = slStr & Trim$(gIntToStrDec(tlIvr.iPctComm, 2))
                End If
                tmInvExport_Header.sAgyComm = Trim$(slStr)
            Else
                llNTRNet = llNTRGross
                tmInvExport_Header.sAgyComm = "0"               '7-21-17 air time may have had commission, this ntr doesnt
            End If
            '5-31-17 store rate as whole number without decimal if no cents
            ilRemainder = llNTRGross Mod 100
            If ilRemainder = 0 Then         'strip off the pennies if whole number
                slStripCents = Trim$(gLongToStrDec(llNTRGross, 2))
                slStr = Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3))
            Else
                slStr = Trim$(gLongToStrDec(llNTRGross, 2))
            End If
            'tmInvExport_NTR.sGross = gLongToStrDec(llNTRGross, 2)
            tmInvExport_NTR.sGross = Trim$(slStr)
            
            ilRemainder = llNTRNet Mod 100
            If ilRemainder = 0 Then         'strip off the pennies if whole number
                slStripCents = Trim$(gLongToStrDec(llNTRNet, 2))
                slStr = Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3))
            Else
                slStr = Trim$(gLongToStrDec(llNTRNet, 2))
            End If
            'tmInvExport_NTR.sNet = gLongToStrDec(llNTRNet, 2)
            tmInvExport_NTR.sNet = Trim$(slStr)
            mInvExport_WriteType 2, hlInvExportSpots, hlInvExportNTR                     'write NTR
        End If
    End If
    Exit Sub
End Sub

'
'               convert generation date and time to string for Email PDF filename
'               Remove slash and replace with nothing in date, remove colon in time
'           gFormatGenDateTime()
'           <input>  ilInputdate(0 to 1)
'                    llInputTime
'           <output>  slDate
'                     slTime
Private Sub mFormatGenDateTimeToStr(ilInputDate() As Integer, slDate As String, llInputTime As Long, slTime As String)
    Dim slTemp As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llDate As Long
    
    gUnpackDateLong ilInputDate(0), ilInputDate(1), llDate
     slStr = Format$(llDate, "ddddd")               'Now date as string
    'replace slash with nothing in date
     slDate = ""
     For ilLoop = 1 To Len(slStr) Step 1
         slTemp = Mid$(slStr, ilLoop, 1)
         If slTemp <> "/" Then
             slDate = Trim$(slDate) & Trim$(slTemp)
         End If
     Next ilLoop
     
     slStr = gFormatTimeLong(llInputTime, "A", "1")
     'remove colons from time
     slTime = ""
     For ilLoop = 1 To Len(slStr) Step 1
         slTemp = Mid$(slStr, ilLoop, 1)
         If slTemp <> ":" Then
             slTime = Trim$(slTime) & Trim$(slTemp)
         End If
     Next ilLoop
     Do While Len(slTime) < 8
         slTime = "0" & slTime
     Loop
     Exit Sub
End Sub

