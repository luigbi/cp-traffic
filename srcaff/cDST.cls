VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'for exports that need to alter info from station time zone to the export's time zone (which I call site)  Note the 'omEarliestZone': because we don't support anything other than EMCP, change A or H to P
' 1- Start (getSite-doesSqlCall  setSite-you already know the site's zone) 2- isDSTActive(export range) this could be Y N or M (if export range crosses in or out of DST, or if start date or end date or export in on start or end of DST)
' 2 (more) some methods don't have the time, so set bmApproximateNoTime = true  We'll look at each day and set the day as right as possible (a little off on start and end dst days)
' 3- set station time zone 4- set acknowlege DST Do this 2nd, as changing station clears old values 5-ZoneByMethod
'depends on
'       gObtainNextSunday, gConvertRegHourToMilitaryHour
'Enums & Constants
Public Enum ZoneMethod
    NOTDEFINED
    Ipump
    isci_export
    XDS
End Enum
Public Enum TimeZone
    None = 0
    EASTERN = 8
    CENTRAL = 7
    MOUNTAIN = 6
    PACIFIC = 5
    ALASKAN = 4
    HAWAIIAN = 3
End Enum
Public Enum DSTActive
    DSTYes
    DSTNo
    DSTMaybe
    DSTUnknown
End Enum
' Private Properties
Private omZoneMethod As ZoneMethod
Private omSiteZone As TimeZone
Private omStationZone As TimeZone
Private omStationHonorsDaylight As Boolean
Private omEarliestZone As TimeZone
Private omDSTActive As DSTActive
'doesn't really do much, but I used it for testing
Private imAlterAmount As Integer
Private bmNextDay As Boolean
Private bmPreviousDay As Boolean
Private smErrorMessage As String
Private dmDaylightSavingsStart As Date
Private dmDaylightSavingsEnd As Date
'currently not used
Private bmApproximateNoTime As Boolean

' Public Properties
Public Property Get ZoneMethod() As ZoneMethod
    ZoneMethod = omZoneMethod
End Property
Public Property Let ZoneMethod(olMethod As ZoneMethod)
    omZoneMethod = olMethod
End Property
Public Property Get SiteZone() As String
    SiteZone = Left$(omSiteZone, 1)
End Property
Public Property Let SiteZone(slZoneChar As String)
    omSiteZone = mConvertToTimeZone(slZoneChar)
End Property
Public Property Get StationZone() As String
    StationZone = Left$(omStationZone, 1)
End Property
Public Property Let StationZone(slZoneChar As String)
    mClear False
    If Len(slZoneChar) > 1 Then
        slZoneChar = Left$(slZoneChar, 1)
    End If
    omStationZone = mConvertToTimeZone(slZoneChar)
    omStationZone = mAlterZoneIfTooEarly(omStationZone)
End Property
Public Property Get StationHonorsDaylight() As Boolean
    StationHonorsDaylight = omStationHonorsDaylight
End Property
Public Property Let StationHonorsDaylight(blHonor As Boolean)
    omStationHonorsDaylight = blHonor
End Property
Public Property Get EarliestZoneAllowed() As String
    EarliestZoneAllowed = Left$(omEarliestZone, 1)
End Property
Public Property Let EarliestZoneAllowed(slZoneChar As String)
    omEarliestZone = mConvertToTimeZone(slZoneChar)
End Property
Public Property Get DaylighSavingsTimeIsActive() As String
    Dim slRet As String
    Select Case omDSTActive
        Case DSTYes
            slRet = "Y"
        Case DSTNo
            slRet = "N"
        Case DSTMaybe
            slRet = "M"
        Case Else
            slRet = "U"
    End Select
    DaylighSavingsTimeIsActive = slRet
End Property
Public Property Get DSTStarts() As Date
    DSTStarts = dmDaylightSavingsStart
End Property
Public Property Get DSTEnds() As Date
    DSTEnds = dmDaylightSavingsEnd
End Property
Public Property Get isNextDay() As Boolean
    isNextDay = bmNextDay
End Property
Public Property Get isPreviousDay() As Boolean
    isPreviousDay = bmPreviousDay
End Property
Public Property Get ErrorMessage() As String
    ErrorMessage = Trim$(smErrorMessage)
End Property
'public methods
Public Function StartGetSite(oMethod As ZoneMethod) As Boolean
    Dim blRet As Boolean
    
    blRet = True
    mClear True
    Select Case oMethod
        Case Ipump
            omZoneMethod = Ipump
            blRet = mZoneGetSiteForIPump()
            omEarliestZone = PACIFIC
        Case isci_export
            omZoneMethod = isci_export
            omSiteZone = EASTERN
           ' omEarliestZone = PACIFIC
        Case XDS
            omZoneMethod = XDS
            blRet = mZoneGetSiteForXDS()
            omEarliestZone = PACIFIC
            bmApproximateNoTime = True
    End Select
    omSiteZone = mAlterZoneIfTooEarly(omSiteZone)
    StartGetSite = blRet
End Function
Public Function StartSetSite(oMethod As ZoneMethod, slSiteChar As String) As Boolean
    Dim blRet As Boolean
    
    blRet = True
    mClear True
    omZoneMethod = oMethod
    omSiteZone = mConvertToTimeZone(slSiteChar)
    'omEarliestZone = PACIFIC
    omSiteZone = mAlterZoneIfTooEarly(omSiteZone)
    StartSetSite = blRet
End Function
Public Function isDSTActive(slExportStart As String, slExportEnd As String) As DSTActive
    Dim olRet As DSTActive
    
    olRet = DSTNo
    'takes the year of the export start and create dmDaylightSavingsStart and dmDaylightSavingsEnd
    mSetDaylightSavingsDates slExportStart
    'export dates entirely within daylight savings time
    If DateDiff("d", dmDaylightSavingsStart, slExportStart) > 0 And DateDiff("d", slExportEnd, dmDaylightSavingsEnd) > 0 Then
        olRet = DSTYes
    ' this export crosses into daylight savings time! must do more later for when this happens as each date/time in export is important
    'start date after daylight starts and before daylight ends...or end date is
    ElseIf (DateDiff("d", dmDaylightSavingsStart, slExportStart) >= 0 And DateDiff("d", slExportStart, dmDaylightSavingsEnd) >= 0) Or (DateDiff("d", slExportEnd, dmDaylightSavingsEnd) >= 0 And DateDiff("d", dmDaylightSavingsStart, slExportEnd) >= 0) Then
        olRet = DSTMaybe
    End If
    omDSTActive = olRet
    isDSTActive = olRet
End Function
Public Sub StationHonorDaylight(ilHonor As Integer)
    '0 = yes 1 = no. Yipes!
    If ilHonor = 0 Then
        omStationHonorsDaylight = True
    Else
        omStationHonorsDaylight = False
    End If
End Sub
Public Function FindZoneDifference() As Integer
    'doesn't handle if station doesn't acknowledge daylight
    imAlterAmount = omSiteZone - omStationZone
    FindZoneDifference = imAlterAmount
End Function
Public Function ZoneByMethod(slDate As String, slTime As String) As String
    'I need date to alter if omDSTActive = Maybe
    'if start or end of export falls on the date change - (approximates don't use times):
        'DST is 3/10/2024 to 11/5/2024
'            if export starts on 3/10 treat as YES (after 2:00 am)
'            if export ends on 3/10 treat as YES (after 2:00 am)
'            if export starts on 11/5 treat as NO (after 2:00 am)
'            if export ends on 11/5 treat as NO (after 2:00 am)
    Dim slRet As String
    Dim ilMilitaryHour As Integer
    Dim olTempDSTActive As DSTActive
    Dim ilTempAlterAmount As Integer
    Dim slMilitaryHour As String
    
    slRet = gConvertRegHourToMilitaryHour(slTime)
    'if DST = 'M', we may need to test the time.  I also need this for Isci export.
    ilMilitaryHour = Val(slRet)
    imAlterAmount = omSiteZone - omStationZone
    'in case station doesn't acknowlege, we have to add one here
    ilTempAlterAmount = imAlterAmount
    If omStationHonorsDaylight = False Then
        olTempDSTActive = omDSTActive
        If olTempDSTActive = DSTMaybe Then
            'date completely between DST
            If DateDiff("d", dmDaylightSavingsStart, slDate) > 0 And DateDiff("d", slDate, dmDaylightSavingsEnd) > 0 Then
                olTempDSTActive = DSTYes
            Else
                'start as if date completely out of DST
                olTempDSTActive = DSTNo
                'this date matches start of DST. Need to test time
                If DateDiff("d", dmDaylightSavingsStart, slDate) = 0 Then
                    'if time on or after 2:00 am, it's a Y
                    If bmApproximateNoTime Or ilMilitaryHour >= 2 Then
                        olTempDSTActive = DSTYes
                        'note that the only time there is a repetition is when time = 2 Both 1 and 2 on end of dst show same value!
                    End If
                'insstead, this date matches end of dst. So opposite of above
                ElseIf DateDiff("d", dmDaylightSavingsEnd, slDate) = 0 Then
                    'if time before 2:00 am, it's a Y
                    If Not bmApproximateNoTime And ilMilitaryHour < 2 Then
                        olTempDSTActive = DSTYes
                    End If
                End If
            End If
        End If 'change maybe to yes or no
        'all the above to this for Arizona!
        If olTempDSTActive = DSTYes Then
            ilTempAlterAmount = ilTempAlterAmount + 1
        End If
    End If
    ilMilitaryHour = ilMilitaryHour + ilTempAlterAmount
    slMilitaryHour = mTestIfDayChanged(ilMilitaryHour)
    Select Case omZoneMethod
        Case isci_export
            slRet = slMilitaryHour
        Case XDS
            slRet = mXDSArrayAlter(ilTempAlterAmount)
'            'make sure not bigger than size of array
'            If ilTempAlterAmount < 0 Then
'                ' -1 in central means eastern. becomes 0.  -1 in mountain = central by adding 2, we get 1, which is Central.
'                ilTempAlterAmount = ilTempAlterAmount + omSiteZone
'            End If
'            'change here if allowing more zones. Right now, a pacific station that doesn't acknowledge will remain pacific. Remember, this is passing into an array so make sure not bigger than allowed
'            If omEarliestZone = PACIFIC And ilTempAlterAmount > 3 Then
'                ilTempAlterAmount = 3
'            End If
'            slRet = ilTempAlterAmount
        Case Else
            slRet = slMilitaryHour
    End Select
    ZoneByMethod = slRet
End Function
Public Sub TestIsDaylight()
    '3/8/2015 to 11/1/2015 completely in dst
    Me.isDSTActive "5/01/2015", "10/30/2015"
        Debug.Print "For 5/01/2015 to 10/30/2015 should be Y: isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    Me.isDSTActive "11/01/2016", "11/08/2016"
        Debug.Print "For 11/01/2016 to 11/08/2016 should M(cross end): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    Me.isDSTActive "3/07/2016", "3/15/2016"
        Debug.Print "For 3/07/2016 to 3/15/2016 should be M(cross start): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    '3/10/2024 to 11/3/2024
    Me.isDSTActive "11/06/2024", "11/30/2024"
        Debug.Print "For 11/06/2024 to 11/30/2024 should be N: isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    Me.isDSTActive "3/11/2024", "11/02/2024"
        Debug.Print "For 3/11/2024 to 11/02/2024 should be Y: isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
        'start on start; end on end
    Me.isDSTActive "3/10/2024", "3/11/2024"
        Debug.Print "For 3/10/2024 to 3/11/2024 should be M(start on start): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    Me.isDSTActive "3/09/2024", "3/10/2024"
        Debug.Print "For 3/09/2024 to 3/10/2024 should be M(end on end): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    '3/12/2023 to 11/5/2023 start on end, end on start
    Me.isDSTActive "11/05/2023", "11/06/2023"
        Debug.Print "For 11/05/2023 to 11/06/2023 should be M(start on dst end): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    Me.isDSTActive "3/02/2023", "3/12/2023"
        Debug.Print "For 3/02/2023 to 3/12/2023 should be M (end on dst start): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
End Sub
Public Sub TestZoneByMethod()
    Dim slRet As String
    Dim olRestoreStationZone As TimeZone
    Dim blRestoreHonor As Boolean
    
    'dst = 3/10/2024 to 11/3/2024
    
    olRestoreStationZone = Me.StationZone
    blRestoreHonor = Me.StationHonorsDaylight
    
    Debug.Print "ApproximateTime = " & bmApproximateNoTime
    Debug.Print "PACIFIC BASIC OK"
    Me.StationZone = "P"
    Me.StationHonorsDaylight = True
        Debug.Print "Station's Zone = " & Me.StationZone & " Acknowledge DST = " & Me.StationHonorsDaylight
    Me.isDSTActive "3/10/2024", "3/11/2024"
        Debug.Print "For 3/10/2024 to 3/11/2024 should be M(start on start): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    slRet = Me.ZoneByMethod("3/10/2024", "2:00a")
        Debug.Print "For 3/10/2024 2:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet

    Debug.Print "CENTRAL BASIC OK"
    Me.StationZone = "M"
    Me.StationHonorsDaylight = True
        Debug.Print "Station's Zone = " & Me.StationZone & " Acknowledge DST = " & Me.StationHonorsDaylight
    Me.isDSTActive "3/10/2024", "3/11/2024"
        Debug.Print "For 3/10/2024 to 3/11/2024 should be M(start on start): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    slRet = Me.ZoneByMethod("3/10/2024", "2:00a")
        Debug.Print "For 3/10/2024 2:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
    ' in dst
    Debug.Print "CENTRAL ALTER+ = YES"
    Me.StationHonorsDaylight = False
        Debug.Print "Station's Zone = " & Me.StationZone & " Acknowledge DST = " & Me.StationHonorsDaylight
    Me.isDSTActive "3/11/2024", "3/11/2024"
        Debug.Print "For 3/11/2024 to 3/11/2024 should be Y: isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    slRet = Me.ZoneByMethod("3/11/2024", "2:00a")
        Debug.Print "For 3/11/2024 2:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
    ' out dst
    Debug.Print "CENTRAL ALTER+ = NO"
        Debug.Print "Station's Zone = " & Me.StationZone & " Acknowledge DST = " & Me.StationHonorsDaylight
    Me.isDSTActive "2/11/2024", "2/17/2024"
        Debug.Print "For 2/11/2024 to 2/17/2024 should be N: isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    slRet = Me.ZoneByMethod("2/11/2024", "2:00a")
        Debug.Print "For 2/11/2024 2:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
    ' time = in dst
    Debug.Print "CENTRAL ALTER+ = YES"
        Debug.Print "Station's Zone = " & Me.StationZone & " Acknowledge DST = " & Me.StationHonorsDaylight
    Me.isDSTActive "3/10/2024", "3/11/2024"
        Debug.Print "For 3/10/2024 to 3/11/2024 should be M(start on start): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    slRet = Me.ZoneByMethod("3/10/2024", "2:00a")
        Debug.Print "For 3/10/2024 2:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
    ' time = in dst
    Debug.Print "CENTRAL ALTER+ = YES"
        Debug.Print "Station's Zone = " & Me.StationZone & " Acknowledge DST = " & Me.StationHonorsDaylight
    Me.isDSTActive "3/10/2024", "3/11/2024"
        Debug.Print "For 3/10/2024 to 3/11/2024 should be M(start on start): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    slRet = Me.ZoneByMethod("3/10/2024", "3:00a")
        Debug.Print "For 3/10/2024 3:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
    ' time = out of dst
    Debug.Print "CENTRAL ALTER+ = NO APPROX DIFFERENT!"
        Debug.Print "Station's Zone = " & Me.StationZone & " Acknowledge DST = " & Me.StationHonorsDaylight
    Me.isDSTActive "3/10/2024", "3/11/2024"
        Debug.Print "For 3/10/2024 to 3/11/2024 should be M(start on start): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    slRet = Me.ZoneByMethod("3/10/2024", "1:00a")
        Debug.Print "For 3/10/2024 1:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
    ' date = in dst
    Debug.Print "CENTRAL ALTER+ = YES"
        Debug.Print "Station's Zone = " & Me.StationZone & " Acknowledge DST = " & Me.StationHonorsDaylight
    Me.isDSTActive "3/10/2024", "3/11/2024"
        Debug.Print "For 3/10/2024 to 3/11/2024 should be M(start on start): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    slRet = Me.ZoneByMethod("3/11/2024", "2:00a")
        Debug.Print "For 3/11/2024 2:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
    ' time = out of dst
    Debug.Print "CENTRAL ALTER+ = NO"
        Debug.Print "Station's Zone = " & Me.StationZone & " Acknowledge DST = " & Me.StationHonorsDaylight
    Me.isDSTActive "11/03/2024", "11/11/2024"
        Debug.Print "For 11/03/2024 to 11/11/2024 should be M(start on dst end): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    slRet = Me.ZoneByMethod("11/03/2024", "2:00a")
        Debug.Print "For 11/03/2024 2:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
    ' time = in dst
    Debug.Print "CENTRAL ALTER+ = YES APPROX DIFFERENT"
        Debug.Print "Station's Zone = " & Me.StationZone & " Acknowledge DST = " & Me.StationHonorsDaylight
    Me.isDSTActive "11/03/2024", "11/11/2024"
        Debug.Print "For 11/03/2024 to 11/11/2024 should be M(start on dst end): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    slRet = Me.ZoneByMethod("11/03/2024", "12:00a")
        Debug.Print "For 11/03/2024 12:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
    ' date = out of dst
    Debug.Print "CENTRAL ALTER+ = NO"
        Debug.Print "Station's Zone = " & Me.StationZone & " Acknowledge DST = " & Me.StationHonorsDaylight
    Me.isDSTActive "11/03/2024", "11/11/2024"
        Debug.Print "For 11/03/2024 to 11/11/2024 should be M(start on dst end): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    slRet = Me.ZoneByMethod("11/04/2024", "12:00a")
        Debug.Print "For 11/04/2024 12:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
Debug.Print "TEST SKIPS AND REPEATS"
    ' time = out to in dst
    Debug.Print "ENTERS DAYLIGHT"
        Debug.Print "Station's Zone = " & Me.StationZone & " Acknowledge DST = " & Me.StationHonorsDaylight
    Me.isDSTActive "3/10/2024", "3/11/2024"
        Debug.Print "For 3/10/2024 to 3/11/2024 should be M(start on dst start): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    slRet = Me.ZoneByMethod("3/10/2024", "12:00a")
        Debug.Print "For 3/10/2024 12:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
    slRet = Me.ZoneByMethod("3/10/2024", "1:00a")
        Debug.Print "For 3/10/2024 1:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
    slRet = Me.ZoneByMethod("3/10/2024", "2:00a")
        Debug.Print "For 3/10/2024 2:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
    'time = in to out dst
     Debug.Print "LEAVES DAYLIGHT"
        Debug.Print "Station's Zone = " & Me.StationZone & " Acknowledge DST = " & Me.StationHonorsDaylight
    Me.isDSTActive "11/3/2024", "11/4/2024"
        Debug.Print "For 11/3/2024 to 11/4/2024 should be M(start on dst start): isDaylight: " & Me.DaylighSavingsTimeIsActive & "  Starts: " & Me.DSTStarts & "  Ends: " & Me.DSTEnds
    slRet = Me.ZoneByMethod("11/3/2024", "12:00a")
        Debug.Print "For 11/3/2024 12:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
    slRet = Me.ZoneByMethod("11/3/2024", "1:00a")
        Debug.Print "For 11/3/2024 1:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
    slRet = Me.ZoneByMethod("11/3/2024", "2:00a")
        Debug.Print "For 3/10/2024 2:00 am with basic alteration of " & imAlterAmount & " time becomes " & slRet
   
    Me.StationZone = olRestoreStationZone
    Me.StationHonorsDaylight = blRestoreHonor

End Sub
'private methods
Private Sub mClear(bIsAll As Boolean)
    If bIsAll Then
        omSiteZone = None
        omZoneMethod = NOTDEFINED
        omEarliestZone = None
        omDSTActive = DSTUnknown
        dmDaylightSavingsStart = "1/1/1970"
        dmDaylightSavingsEnd = "1/1/1970"
        bmApproximateNoTime = False
    End If
    omStationZone = None
    omStationHonorsDaylight = False
    imAlterAmount = 0
    bmNextDay = False
    bmPreviousDay = False
    smErrorMessage = ""
End Sub
Private Function mConvertToTimeZone(slZoneChar As String) As TimeZone
    Dim olZone As TimeZone
    Select Case slZoneChar
        Case "E"
            olZone = EASTERN
        Case "C"
            olZone = CENTRAL
        Case "M"
            olZone = MOUNTAIN
        Case "P"
            olZone = PACIFIC
        Case "A"
            olZone = ALASKAN
        Case "H"
            olZone = HAWAIIAN
        Case Else
            olZone = None
    End Select
    mConvertToTimeZone = olZone
End Function
Private Function mZoneGetSiteForIPump() As Boolean
    Dim blRet As Boolean
    Dim slZone As String
    Dim rsZone As ADODB.Recordset
    Dim slSql As String
    
On Error GoTo ERRORBOX
    blRet = True
    omSiteZone = None
    slZone = ""
    slSql = "SELECT safIPumpZone as SiteZone From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set rsZone = gSQLSelectCall(slSql)
    If Not rsZone.EOF Then
        slZone = UCase(rsZone!SiteZone)
        omSiteZone = mConvertToTimeZone(slZone)
    End If
Cleanup:
    If Not rsZone Is Nothing Then
        If (rsZone.State And adStateOpen) <> 0 Then
            rsZone.Close
        End If
        Set rsZone = Nothing
    End If
    mZoneGetSiteForIPump = blRet
    Exit Function
ERRORBOX:
    blRet = False
    smErrorMessage = smErrorMessage & " Issue in mZoneGetSiteForIPump"
    GoTo Cleanup
End Function
Private Function mZoneGetSiteForXDS() As Boolean
    Dim blRet As Boolean
    Dim slSQLQuery As String
    Dim rsZone As ADODB.Recordset
    Dim slHeadEndZone As String
    
On Error GoTo ERRORBOX
    blRet = True
    omSiteZone = None
    slHeadEndZone = ""
    slSQLQuery = "Select safXDSHeadEndZone From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set rsZone = gSQLSelectCall(slSQLQuery)
    If Not rsZone.EOF Then
        slHeadEndZone = rsZone!safXDSHeadEndZone
    End If
    If (Trim$(slHeadEndZone) = "") Then
        slHeadEndZone = "E"
    End If
    omSiteZone = mConvertToTimeZone(slHeadEndZone)
Cleanup:
    If Not rsZone Is Nothing Then
        If (rsZone.State And adStateOpen) <> 0 Then
            rsZone.Close
        End If
        Set rsZone = Nothing
    End If
        mZoneGetSiteForXDS = blRet
    Exit Function
ERRORBOX:
    blRet = False
    smErrorMessage = smErrorMessage & " Issue in mZoneGetSiteForIPump"
    GoTo Cleanup
End Function
Private Function mAlterZoneIfTooEarly(olZone As TimeZone) As TimeZone
    'our system only handles EACP!
    Dim olRet As TimeZone
    olRet = olZone
    If omEarliestZone > None Then
        If olZone < omEarliestZone Then
            olRet = omEarliestZone
        End If
    End If
    mAlterZoneIfTooEarly = olRet
End Function
Private Function mTestIfDayChanged(ilValue As Integer) As String
    Dim slRet As String
    
    bmNextDay = False
    bmPreviousDay = False
    If ilValue >= 24 Then
        ilValue = ilValue - 24
        bmNextDay = True
    ElseIf ilValue < 0 Then
        '-1 becomes 23
        ilValue = 24 + ilValue
        bmPreviousDay = True
    End If
    slRet = ilValue
    If Len(slRet) = 1 Then
        slRet = "0" & slRet
    End If
    mTestIfDayChanged = slRet
End Function
Private Sub mSetDaylightSavingsDates(slStartOfExport As String)
    ' start: 2nd sunday of march end: 1st sunday in November  both 2:00 am
    Dim slDate As String
    Dim ilDay As Integer
  
     slDate = "3/01/" & DatePart("yyyy", slStartOfExport)
     ilDay = DatePart("w", slDate)
     If ilDay <> 1 Then
         slDate = gObtainNextSunday(slDate)
     End If
     slDate = DateAdd("ww", 1, slDate)
    dmDaylightSavingsStart = CDate(slDate & " 2:00:00 am")

     slDate = "11/01/" & DatePart("yyyy", slStartOfExport)
     ilDay = DatePart("w", slDate)
     If ilDay <> 1 Then
         slDate = gObtainNextSunday(slDate)
     End If
    dmDaylightSavingsEnd = CDate(slDate & " 2:00:00 am")
End Sub
Private Function mXDSArrayAlter(ilTempAlterAmount As Integer) As Integer
    'array is 0Eastern 1Central 2Mtn 3Pacific.  Assumes head end.  Add adjustment if not Eastern, so Pacific to Mtn (which is -1) becomes 2
    Dim ilSiteNotEasternAdjuster As Integer
    Dim ilEastern As TimeZone
    Dim ilRet As Integer
    
    ilEastern = EASTERN
    ilSiteNotEasternAdjuster = ilEastern - omSiteZone
    ilRet = ilTempAlterAmount + ilSiteNotEasternAdjuster
'    'make sure not smaller, but shouldn't happen
    If ilRet < 0 Then
        ilRet = 0
    End If
    'change here if allowing more zones. Right now, a pacific station that doesn't acknowledge will remain pacific. Remember, this is passing into an array so make sure not bigger than allowed
    If omEarliestZone = PACIFIC And ilRet > 3 Then
        ilRet = 3
    End If
    mXDSArrayAlter = ilRet
End Function

