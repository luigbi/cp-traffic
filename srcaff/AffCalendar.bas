Attribute VB_Name = "modCalendar"
'******************************************************
'*  modCalendar - various global declarations for importing
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit


Type FIELDAREA
    fBoxX As Single  'Box X position
    fBoxY As Single  'Box Y position
    fBoxW As Single  'Box Width
    fBoxH As Single  'Box Height
    iReq As Integer  'Input required
    iChg As Integer  'Field changed flag (if so, show as bold)
    iAlign As Integer 'Align: 0= left (LEFTJUSTIFY), 1= Right(RIGHTJUSTIFY), 2= Center(CENTER)
    sShow As String  'String to be shown
    'iUpArrowIndex As Integer    'ID Index to go to for up arrow
    'iDnArrowIndex As Integer    'ID Index to go to for down arrow
    'sCaption As String  'Title to be shown in top left of box
End Type

Dim imCalMonth As Integer
Dim imCalYear As Integer
Dim fmBoxInsetX As Single


'*******************************************************
'*                                                     *
'*      Procedure Name:gSetCtrl                        *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set mouse area within the       *
'*                     control area.                   *
'*                                                     *
'*******************************************************
Sub gSetCtrl(tlCtrlArray As FIELDAREA, flBoxX As Single, flBoxY As Single, flBoxW As Single, flBoxH As Single)
'
'   gSetCtrl CtrlArray(1), fBoxX, fBoxY, fBoxW, fBoxH
'   Where
'       CtrlArray (I/O)- field control array
'       fBoxX (I)- x offset of mouse area within picture control
'       fBoxY (I)- y offset of mouse area within picture control
'       fBoxW (I)- Width of mouse area within picture
'       fBoxH (I)- Height of mouse area within picture
'

    tlCtrlArray.fBoxX = flBoxX
    tlCtrlArray.fBoxY = flBoxY
    tlCtrlArray.fBoxW = flBoxW
    tlCtrlArray.fBoxH = flBoxH
    tlCtrlArray.iReq = True
    tlCtrlArray.iChg = False
    tlCtrlArray.iAlign = 0
    tlCtrlArray.sShow = ""
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gPaintCalendar                  *
'*                                                     *
'*             Created:8/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Paint standard or regular       *
'*                     calendar                        *
'*                                                     *
'*******************************************************
Sub gPaintCalendar(ilMonth As Integer, ilYear As Integer, ilType As Integer, pbcCtrl As PictureBox, tlCtrls() As FIELDAREA, llStartDate As Long, llEndDate As Long)
'
'   gPaintCalendar ilMonth, ilYear, ilType, pbcCtrl, tmCtrl, llStart, llEnd
'   Where:
'       ilMonth (I) - Calendar month to be painted (1 thru 12)
'       ilYear (I) - Calendar year to be painted (101 thru 9998) (100 & 9999 not                        '       allowed because of standard)
'       ilType (I)- 0=Paint standard month; 1= Paint regular month; 2=Julian +;
'                   3=Julian -; 4=Paint corporate on Jan-Dec; 5=Paint Corporate by Fiscal (Oct-Sept)
'       pbcCtrl (I)- Picture area control to be painted
'       tlCtrl() (I)- Array of control information about paint area
'       llStart (O)- Calendar start date
'       llEnd (O)- Calendar end date
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

    fmBoxInsetX = 30    'X Margin from box outline to text
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
        slDate = Trim$(Str$(ilAdjMonth)) & "/1/" & Trim$(Str$(ilYear))
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
'    ElseIf ilType = 4 Then  'Corporate by Jan-Dec
'        If (ilYear >= 0) And (ilYear <= 69) Then
'            ilAdjYear = 2000 + ilYear
'        ElseIf (ilYear >= 70) And (ilYear <= 99) Then
'            ilAdjYear = 1900 + ilYear
'        Else
'            ilAdjYear = ilYear
'        End If
'        ilFound = False
'        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
'            If ilMonth >= tgMCof(ilLoop).iStartMnthNo Then
'                If tgMCof(ilLoop).iYear = ilAdjYear + 1 Then
'                    ilMnth = tgMCof(ilLoop).iStartMnthNo
'                    ilIndex = 1
'                    Do While ilMnth <> ilMonth
'                        ilMnth = ilMnth + 1
'                        If ilMnth > 12 Then
'                            ilMnth = 1
'                        End If
'                        ilIndex = ilIndex + 1
'                    Loop
'                    gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex), tgMCof(ilLoop).iStartDate(1, ilIndex), llDate
'                    gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex), tgMCof(ilLoop).iEndDate(1, ilIndex), llLastDate
'                    ilFound = True
'                    Exit For
'                End If
'            Else
'                If tgMCof(ilLoop).iYear = ilAdjYear Then
'                    ilMnth = tgMCof(ilLoop).iStartMnthNo
'                    ilIndex = 1
'                    Do While ilMnth <> ilMonth
'                        ilMnth = ilMnth + 1
'                        If ilMnth > 12 Then
'                            ilMnth = 1
'                        End If
'                        ilIndex = ilIndex + 1
'                    Loop
'                    gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex), tgMCof(ilLoop).iStartDate(1, ilIndex), llDate
'                    gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex), tgMCof(ilLoop).iEndDate(1, ilIndex), llLastDate
'                    ilFound = True
'                    Exit For
'                End If
'            End If
'        Next ilLoop
'        If Not ilFound Then
'            pbcCtrl.Cls
'            Exit Sub
'        End If
'    ElseIf ilType = 5 Then  'Corporate by Fiscal (Oct-Sept)
'        If (ilYear >= 0) And (ilYear <= 69) Then
'            ilAdjYear = 2000 + ilYear
'        ElseIf (ilYear >= 70) And (ilYear <= 99) Then
'            ilAdjYear = 1900 + ilYear
'        Else
'            ilAdjYear = ilYear
'        End If
'        ilFound = False
'        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
'            If tgMCof(ilLoop).iYear = ilAdjYear Then
'                For ilIndex = 1 To 12 Step 1
'                    ilAdjMonth = tgMCof(ilLoop).iStartMnthNo + ilIndex - 1
'                    If ilAdjMonth > 12 Then
'                        ilAdjMonth = ilAdjMonth - 12
'                    End If
'                    If ilAdjMonth = ilMonth Then
'                        gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex), tgMCof(ilLoop).iStartDate(1, ilIndex), llDate
'                        gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex), tgMCof(ilLoop).iEndDate(1, ilIndex), llLastDate
'                        ilFound = True
'                        Exit For
'                    End If
'                Next ilIndex
'                If ilFound Then
'                    Exit For
'                End If
'            End If
'        Next ilLoop
'        If Not ilFound Then
'            pbcCtrl.Cls
'            Exit Sub
'        End If
    Else
        slDate = Trim$(Str$(ilAdjMonth)) & "/1/" & Trim$(Str$(ilYear))
        If ilYear < 100 Then
            slDate = gAdjYear(slDate)
            ilAdjYear = Year(slDate)
            llDate = gDateValue(slDate)
        Else
            ilAdjYear = Year(slDate)
            llDate = DateValue(slDate)
        End If
    End If
'    If gValidDate(slDate) = No Then
'        pbcCtrl.Cls
'        Exit Sub
'    End If
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
        flBoxInsetX = fmBoxInsetX
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
            slDay = Trim$(Str$(Day(llDate)))
            If Len(slDay) <= 1 Then
                slDay = " " & slDay
            End If
            If ilAdjMonth <> Month(llDate) Then
                pbcCtrl.ForeColor = vbMagenta 'GREEN
            Else
                pbcCtrl.ForeColor = llColor
            End If
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
            slDay = Trim$(Str$(Val(Format$(slJulian, "y")) - Val(Format$(llDate, "y"))))
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
''        gSetShow pbcCtrl, slDay, tlCtrls(ilWkDay + 1)
        'pbcCtrl.CurrentX = tlCtrls(ilWkDay + 1).fBoxX + flBoxInsetX
        'pbcCtrl.CurrentY = tlCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tlCtrls(ilWkDay + 1).fBoxH + 15) + flBoxInsetY '(fgBoxGridH + 15) -  30'+ fgBoxInsetY
        pbcCtrl.CurrentX = tlCtrls(ilWkDay).fBoxX + flBoxInsetX
        pbcCtrl.CurrentY = tlCtrls(ilWkDay).fBoxY + ilRowNo * (tlCtrls(ilWkDay).fBoxH + 15) + flBoxInsetY '(fgBoxGridH + 15) -  30'+ fgBoxInsetY
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


