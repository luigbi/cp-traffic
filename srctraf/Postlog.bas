Attribute VB_Name = "POSTLOGSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Postlog.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Variables (Removed)                                                             *
'*  sgRepMissedDate               igRepAirVefCode               sgRepSpotType             *
'*  igRepReturnStatus                                                                     *
'*                                                                                        *
'* Public Type Defs (Marked)                                                              *
'*  PLCNTSPOT                                                                             *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PostLog.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the constant definitions for all the status codes returned by
'   functions within the traffic system.
Option Explicit
Option Compare Text

'Rep Spot Posting
Public lgRepSpotChfCode As Long
Public igRepSpotLineNo As Integer
Public lgRepSpotStartStd As Long
Public lgRepSpotEndStd As Long
Public lgRepSpotStartCal As Long
Public lgRepSpotEndCal As Long
Public igRepSpotBilled As Integer
Type REPSPOTLINEINFO
    iLineNo As Integer
    'iCarry As Integer
    'iCalCarry As Integer
End Type
Public tgRepSpotLineInfo() As REPSPOTLINEINFO

'Manual posting times from invoices
Public sgCntrLineInfo As String
Public igPostManualTimes As Integer '0=Cancel; 1= Done
Public igPostManualAdfCode As Integer
Public igPostManualLength As Integer
Type POSTMANUALCNTRINFO
    lRow As Long
    sDaypartTimes As String * 25
    sDaypartDays As String * 21
    sLength As String * 3
End Type
Type POSTMANUALTIMESINFO
    lRow As Long
    sDate As String * 10
    sTime As String * 11
    sISCI As String * 20
    sSchStatus As String * 1
    lCifCode As Long
    sWkDate As String * 10
    iWeek As Integer
    lIndex As Long
End Type
Public tgPostManualCntrInfo() As POSTMANUALCNTRINFO
Public tgPostManualTimesInfo() As POSTMANUALTIMESINFO

'Air Time
Public igAvailVefCode As Integer
Public sgAvailDate As String
Public sgAvailTime As String
Public igAvailAnfCode As Integer
Public igAvailGameNo As Integer
Public sgPLogAirTime As String
Public igPLogSource As Integer  '0=Vehicle; 1=Package
Dim imCalMonth As Integer
Dim imCalYear As Integer
Type DATES
    sDate As String
    lDate As Long
    iStatus As Integer  '0=Not Complete; 1= Incomplete; 2=Complete
    iGameNo As Integer
    sLiveLogMerge As String * 1
End Type
Public tgDates() As DATES   'All Post Log Dates
Public tgWkDates() As DATES 'Only dates within week selected
'Taken from SpotFill.Bas
Type PLCNTSPOT 'VBC NR
    sKey As String * 90     'Sort by type, advertiser, Contract #, Vehicle 'VBC NR
    sType As String * 1     'ChfType: C=Commercial Spot, S=PSA; M=Promo 'VBC NR
    iAdfCode As Integer     'Advertiser Code 'VBC NR
    lChfCode As Long        'Contract Code 'VBC NR
    iVefCode As Integer     'Spot Vehicle Code 'VBC NR
    iLineNo As Integer      'Line number 'VBC NR
    iLnVefCode As Integer   'Line vehicle (used to bypass allow day test) 'VBC NR
    iNoSSpots As Integer     'Number of spots scheduled 'VBC NR
    iNoGSpots As Integer     'Number of MG or Outside spots 'VBC NR
    iNoMSpots As Integer     'Number of missed spots 'VBC NR
    iNoESpots As Integer    'Number of Extra spots 'VBC NR
    lSdfRecPos As Long        'Record image to be used when creating Fills 'VBC NR
    sProduct As String * 35 'VBC NR
    sLen As String * 3 'VBC NR
    sDate As String * 17    'Date range of line 'VBC NR
    'iAllowedDays(0 To 6) As Integer
    'sPriceType As String * 1
    'lPrice As Long
    lAllowedSTime As Long 'VBC NR
    lAllowedETime As Long 'VBC NR
    iNoTimesUsed As Integer 'VBC NR
    iMnfComp0 As Integer 'VBC NR
    iMnfComp1 As Integer 'VBC NR
    'iSpotLkIndex As Integer
    'iRdfCode As Integer
    'lAud As Long            'Audience
End Type 'VBC NR
'Public tgTPLCntSpot() As PLCNTSPOT
Type SAVEINFO
    sKey As String * 62 'Key (Advertiser; Date; Time; Count or Date;Time; Count
    iType As Integer    '0=Avail; 1=Spot
    sAirDate As String * 10    'Air date
    sAirTime As String * 10    'Air Time
    iCount As Integer       'Spot count within avail * 10
    lCntrNo As Long
    iLineNo As Integer
    iLen As Integer
    iUnits As Integer
    sTZone As String * 1
    sSpotType As String * 1
    sSchStatus As String * 1
    sAffChg As String * 1
    sCopy As String * 10        'Media Code + Number + Cut if using Cart #'s otherwise blank
    sISCI As String * 20        'ISCI
    sCopyProduct As String * 35 'Product from Cpf
    sPrice As String * 14       'Price
    sAdvtName As String * 42    'Advertiser Name
    lSdfRecPos As Long
    sProd As String * 35        'Contract Product
    sSchTime As String * 10
    sSchDate As String * 10
    iPrice As Integer       'Price (0=Charge; 1=N/C; -1=Can't Alter)
    iSvPrice As Integer     'Price (for chg test)
    iISCIReq As Integer     'True=ISCI Required; False=ISCI not required
    iISCI As Integer        'True=ISCI missing; False=ISCI defined
    iBilled As Integer      'True = billed; False=Not billed
    iSimulCast As Integer   'True=SimulCast; False=Not SimulCast
    ianfCode As Integer
    sXMid As String * 1
    iShowInfoIndex As Integer
End Type
Type SHOWINFO
    iType As Integer
    sShow(0 To 13) As String * 40   'Index zero ignored
    iChk As Integer         'Row checked
    iSaveInfoIndex As Integer
End Type
Public tgSave() As SAVEINFO
Public tgShow() As SHOWINFO
Type MDSDFREC
    lSdfCode As Long
    lSdfRecPos As Long
    lMissedDate As Long
    sSchStatus As String * 1
    lPrice As Long
    iNextIndex As Integer
End Type
Type MDSAVEINFO
    lChfCode As Long
    lFsfCode As Long
    iAdfCode As Integer
    lCntrNo As Long
    lEndDate As Long    'End Date of contract
    iVefCode As Integer
    iLineNo As Integer  'Used for creation of Bonus spots
    iLen As Integer
    lWkMissed As Long
    iRdfCode As Integer
    iStartTime(0 To 1) As Integer 'Override Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
                            'If not defined-> use rate card times (not defined is hund sec = 1 all other times = 0)
    iEndTime(0 To 1) As Integer 'Override End Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iDay(0 To 6)  As Integer    'Spot per day if daily or flag if weekly
                                'For weekly 1=Air day; 0=not air day
                                'Index 0=Mo; 1=tu,...6=Su
    iMissedCount As Integer     'Number of missed spots within week
    iCancelCount As Integer
    iHiddenCount As Integer
    sBill As String * 1
    iFirstIndex As Integer
End Type
Type MDSHOWINFO
    sKey As String * 82
    sShow(0 To 9) As String * 42    'Index zero ignored
    iType As Integer                '0=Contract; 1=Missed; 2=Hidden; 3=Cancelled
    sBill As String * 1
    iMdSaveInfoIndex As Integer
End Type
Public tgMdSdfRec() As MDSDFREC
Public tgMdSaveInfo() As MDSAVEINFO
Public tgMdShowInfo() As MDSHOWINFO

'******************************************************************************
' LRF_Log_Remote_Stn Record Definition
'
'******************************************************************************
Type LRF
    lCode                 As Long            ' Internal Reference Code
    sUserName             As String * 40     ' Remote User Name when accepting
                                             ' Citation
    iInvStartDate(0 To 1) As Integer         ' Invoice month start date
    iPostVefCode          As Integer         ' Vehicle reference that is being
                                             ' posted
    iPostStartDate(0 To 1) As Integer        ' Start date of the posting (now
                                             ' date)
    iPostStartTime(0 To 1) As Integer        ' Start time of the posting (now
                                             ' time)
    iPostEndDate(0 To 1)  As Integer         ' End date of the posting (now
                                             ' date)
    iPostEndTime(0 To 1)  As Integer         ' End time of the posting (now
                                             ' time)
    sUnused               As String * 10     ' Unused
End Type


Type LRFKEY0
    lCode                 As Long
End Type

Type LRFKEY1
    iPostVefCode          As Integer
    iInvStartDate(0 To 1) As Integer
End Type
Public sgRemoteUserName As String
Public tgLrf As LRF
Private hmLrf As Integer            'Contract BR file handle
Private tmLrfSrchKey0 As LONGKEY0     'LVF key record image
Private imLrfRecLen As Integer      'CBF record length


'*******************************************************
'*                                                     *
'*      Procedure Name:gPLPaintCalendar                *
'*                                                     *
'*             Created:8/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Paint standard or regular       *
'*                     calendar                        *
'*                     Taken from gPaintCalendar       *
'*                                                     *
'*******************************************************
Sub gPLPaintCalendar(ilMonth As Integer, ilYear As Integer, ilType As Integer, pbcCtrl As PictureBox, tlCtrls() As FIELDAREA, llStartDate As Long, llEndDate As Long, tlDates() As DATES)
'
'   gPLPaintCalendar ilMonth, ilYear, ilType, pbcCtrl, tmCtrl, llStart, llEnd
'   Where:
'       ilMonth (I) - Calendar month to be painted (1 thru 12)
'       ilYear (I) - Calendar year to be painted (101 thru 9998) (100 & 9999 not                        '       allowed because of standard)
'       ilType (I)- 0=Paint standard month; 1= Paint regular month; 2=Julian +;
'                   3=Julian -; 4=Paint corporate on Jan-Dec; 5=Paint Corporate by Fiscal (Oct-Sept)
'       pbcCtrl (I)- Picture area control to be painted
'       tlCtrl() (I)- Array of control information about paint area
'       llStart (O)- Calendar start date
'       llEnd (O)- Calendar end date
'       tlDaters()(I)- Legal Dates
'
    Dim llDate As Long
    Dim llStdDate As Long
    Dim slDate As String
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim slJulian As String
    Dim ilRowNo As Integer
    Dim llLastDate As Long
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim flBoxInsetX As Single
    Dim flBoxInsetY As Single
    Dim ilAdjYear As Integer
    Dim ilAdjMonth As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilFound As Integer
    Dim ilMnth As Integer
    If (ilYear < 101) Or (ilYear > 9998) Then
        pbcCtrl.Cls
        Exit Sub
    End If
    If (ilMonth < 1) Or (ilMonth > 12) Then
        pbcCtrl.Cls
        Exit Sub
    End If
    ilAdjMonth = ilMonth
    If ilType = 0 Then  'Standard month
        slDate = Trim$(str$(ilAdjMonth)) & "/1/" & Trim$(str$(ilYear))
        If ilYear < 100 Then
            slDate = gAdjYear(slDate)
            ilAdjYear = Year(slDate)
            llDate = gDateValue(slDate)
            llStdDate = llDate
        Else
            ilAdjYear = Year(slDate)
            llDate = DateValue(slDate)
            llStdDate = llDate
        End If
        Do While gWeekDayLong(llDate) <> 0   '0=monday
            llDate = llDate - 1
        Loop
        Do
            If gWeekDayLong(llStdDate) = 6 Then  'Save last sunday
                llLastDate = llStdDate
            End If
            llStdDate = llStdDate + 1
        Loop Until Month(llStdDate) <> ilAdjMonth
    ElseIf ilType = 4 Then  'Corporate by Jan-Dec
        If (ilYear >= 0) And (ilYear <= 69) Then
            ilAdjYear = 2000 + ilYear
        ElseIf (ilYear >= 70) And (ilYear <= 99) Then
            ilAdjYear = 1900 + ilYear
        Else
            ilAdjYear = ilYear
        End If
        ilFound = False
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            If ilMonth >= tgMCof(ilLoop).iStartMnthNo Then
                If tgMCof(ilLoop).iYear = ilAdjYear + 1 Then
                    ilMnth = tgMCof(ilLoop).iStartMnthNo
                    ilIndex = 1
                    Do While ilMnth <> ilMonth
                        ilMnth = ilMnth + 1
                        If ilMnth > 12 Then
                            ilMnth = 1
                        End If
                        ilIndex = ilIndex + 1
                    Loop
                    gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex - 1), tgMCof(ilLoop).iStartDate(1, ilIndex - 1), llDate
                    gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex - 1), tgMCof(ilLoop).iEndDate(1, ilIndex - 1), llLastDate
                    ilFound = True
                    Exit For
                End If
            Else
                If tgMCof(ilLoop).iYear = ilAdjYear Then
                    ilMnth = tgMCof(ilLoop).iStartMnthNo
                    ilIndex = 1
                    Do While ilMnth <> ilMonth
                        ilMnth = ilMnth + 1
                        If ilMnth > 12 Then
                            ilMnth = 1
                        End If
                        ilIndex = ilIndex + 1
                    Loop
                    gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex - 1), tgMCof(ilLoop).iStartDate(1, ilIndex - 1), llDate
                    gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex - 1), tgMCof(ilLoop).iEndDate(1, ilIndex - 1), llLastDate
                    ilFound = True
                    Exit For
                End If
            End If
        Next ilLoop
        If Not ilFound Then
            pbcCtrl.Cls
            Exit Sub
        End If
    ElseIf ilType = 5 Then  'Corporate by Fiscal (Oct-Sept)
        If (ilYear >= 0) And (ilYear <= 69) Then
            ilAdjYear = 2000 + ilYear
        ElseIf (ilYear >= 70) And (ilYear <= 99) Then
            ilAdjYear = 1900 + ilYear
        Else
            ilAdjYear = ilYear
        End If
        ilFound = False
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            If tgMCof(ilLoop).iYear = ilAdjYear Then
                For ilIndex = 1 To 12 Step 1
                    ilAdjMonth = tgMCof(ilLoop).iStartMnthNo + ilIndex - 1
                    If ilAdjMonth > 12 Then
                        ilAdjMonth = ilAdjMonth - 12
                    End If
                    If ilAdjMonth = ilMonth Then
                        gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex - 1), tgMCof(ilLoop).iStartDate(1, ilIndex - 1), llDate
                        gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex - 1), tgMCof(ilLoop).iEndDate(1, ilIndex - 1), llLastDate
                        ilFound = True
                        Exit For
                    End If
                Next ilIndex
                If ilFound Then
                    Exit For
                End If
            End If
        Next ilLoop
        If Not ilFound Then
            pbcCtrl.Cls
            Exit Sub
        End If
    Else
        slDate = Trim$(str$(ilAdjMonth)) & "/1/" & Trim$(str$(ilYear))
        If ilYear < 100 Then
            slDate = gAdjYear(slDate)
            ilAdjYear = Year(slDate)
            llDate = gDateValue(slDate)
        Else
            ilAdjYear = Year(slDate)
            llDate = DateValue(slDate)
        End If
    End If
    If (ilAdjMonth <> imCalMonth) Or (ilAdjYear <> imCalYear) Then
        pbcCtrl.Cls
        imCalMonth = ilAdjMonth
        imCalYear = ilAdjYear
    End If
    ilRowNo = 0
    llStartDate = 0
    llEndDate = 0
    slFontName = pbcCtrl.FontName
    flFontSize = pbcCtrl.FontSize
    If (ilType <= 1) Or (ilType = 4) Or (ilType = 5) Then
        pbcCtrl.FontBold = True
        flBoxInsetX = fgBoxInsetX
        flBoxInsetY = -15
    Else
        pbcCtrl.FontBold = False
        pbcCtrl.FontSize = 7
        pbcCtrl.FontName = "Arial"
        pbcCtrl.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        flBoxInsetX = 5
        flBoxInsetY = 5
    End If
    llColor = pbcCtrl.ForeColor
    Do
        ilWkDay = gWeekDayLong(llDate)
        If llStartDate = 0 Then
            llStartDate = llDate
        End If
        If (ilType <= 1) Or (ilType = 4) Or (ilType = 5) Then
            slDay = Trim$(str$(Day(llDate)))
            If Len(slDay) <= 1 Then
                slDay = " " & slDay
            End If
            'If ilAdjMonth <> Month(llDate) Then
            '    pbcCtrl.ForeColor = MAGENTA 'GREEN
            'Else
            '    pbcCtrl.ForeColor = llColor
            'End If
            pbcCtrl.ForeColor = RED
            For ilLoop = LBound(tlDates) To UBound(tlDates) - 1 Step 1
                If tlDates(ilLoop).lDate = llDate Then
                    If tlDates(ilLoop).iStatus = 2 Then
                        pbcCtrl.ForeColor = DARKGREEN
                    Else
                        pbcCtrl.ForeColor = BLUE
                    End If
                    Exit For
                End If
            Next ilLoop
        ElseIf ilType = 2 Then  'Julian +
            slDay = Format$(llDate, "y")
            If Len(slDay) <= 1 Then
                slDay = "   " & slDay
            ElseIf Len(slDay) = 2 Then
                If Val(slDay) < 20 Then
                    slDay = "  " & slDay
                Else
                    slDay = " " & slDay
                End If
            End If
        ElseIf ilType = 3 Then  'Julian -
            slJulian = "12/31/" & Format$(llDate, "yyyy")
            slDay = Trim$(str$(Val(Format$(slJulian, "y")) - Val(Format$(llDate, "y"))))
            If Len(slDay) <= 1 Then
                slDay = "   " & slDay
            ElseIf Len(slDay) = 2 Then
                If Val(slDay) < 20 Then
                    slDay = "  " & slDay
                Else
                    slDay = " " & slDay
                End If
            End If
        End If
'        gSetShow pbcCtrl, slDay, tlCtrls(ilWkDay + 1)
        pbcCtrl.CurrentX = tlCtrls(ilWkDay + 1).fBoxX + flBoxInsetX
        pbcCtrl.CurrentY = tlCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tlCtrls(ilWkDay + 1).fBoxH + 15) + flBoxInsetY '(fgBoxGridH + 15) -  30'+ fgBoxInsetY
        pbcCtrl.Print slDay 'tlCtrls(ilWkDay + 1).sShow
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
'    Loop Until (Month(llDate) <> ilMonth) Or ((ilType <> 1) And (llDate > llLastDate)) 'julian by std
    Loop Until (((ilType <> 0) And (ilType <> 4) And (ilType <> 5)) And (Month(llDate) <> ilAdjMonth)) Or (((ilType = 0) Or (ilType = 4) Or (ilType = 5)) And (llDate > llLastDate))
    pbcCtrl.ForeColor = llColor
    llEndDate = llDate - 1
    pbcCtrl.FontSize = flFontSize
    pbcCtrl.FontName = slFontName
    pbcCtrl.FontSize = flFontSize
    pbcCtrl.FontBold = True
End Sub

Public Sub gAddOrUpdateLrf(slOperType As String, ilRemoteVefCode As Integer, slInvStartDate As String)
    'slOperType(I): A=Add; U=Update
    Dim ilRet As Integer
    
    If igSportsSystem <> 3 Then
        Exit Sub
    End If
    If (sgRemoteUserName = "") And (slOperType = "U") Then
        Exit Sub
    End If
    If (tgLrf.lCode <= 0) And (slOperType = "U") Then
        Exit Sub
    End If
    hmLrf = CBtrvTable(TWOHANDLES) 'CBtrvTable()
    ilRet = btrOpen(hmLrf, "", sgDBPath & "Lrf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    imLrfRecLen = Len(tgLrf)
    If slOperType = "A" Then
        tgLrf.lCode = 0
        tgLrf.sUserName = sgRemoteUserName
        tgLrf.iPostVefCode = ilRemoteVefCode
        gPackDate Format(Now, "m/d/yy"), tgLrf.iPostStartDate(0), tgLrf.iPostStartDate(1)
        gPackTime Format(Now, "h:mm:ssAM/PM"), tgLrf.iPostStartTime(0), tgLrf.iPostStartTime(1)
        gPackDate Format("12/31/2069", "m/d/yy"), tgLrf.iPostEndDate(0), tgLrf.iPostEndDate(1)
        gPackTime Format("12am", "h:mm:ssAM/PM"), tgLrf.iPostEndTime(0), tgLrf.iPostEndTime(1)
        gPackDate Format(slInvStartDate, "m/d/yy"), tgLrf.iInvStartDate(0), tgLrf.iInvStartDate(1)
        ilRet = btrInsert(hmLrf, tgLrf, imLrfRecLen, INDEXKEY0)
        If ilRet <> BTRV_ERR_NONE Then
            tgLrf.lCode = -1
        End If
    Else
        tmLrfSrchKey0.lCode = tgLrf.lCode
        ilRet = btrGetEqual(hmLrf, tgLrf, imLrfRecLen, tmLrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            gPackDate Format(Now, "m/d/yy"), tgLrf.iPostEndDate(0), tgLrf.iPostEndDate(1)
            gPackTime Format(Now, "h:mm:ssAM/PM"), tgLrf.iPostEndTime(0), tgLrf.iPostEndTime(1)
            ilRet = btrUpdate(hmLrf, tgLrf, imLrfRecLen)
            tgLrf.lCode = -1
        End If
    End If
    ilRet = btrClose(hmLrf)
    btrDestroy hmLrf
End Sub
