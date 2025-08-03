VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CcsiToIdcRegionalInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

   Private smRegionalISCIs As String
   Private smStartDate As String
   Private smEndDate As String
   Private smStartTime As String
   Private smEndTime As String
   Private smSites As String
   Private smSplitMajor As String
   Private smSplitMinor As String
   Private smErrorMessage As String
   Public Rotation As Integer
   Public Definition As String
    Public Property Get SplitMajor() As String
        SplitMajor = smSplitMajor
    End Property
    Public Property Get SplitMinor() As String
        SplitMinor = smSplitMinor
    End Property
    Public Property Get RegionalISCIs() As String
        RegionalISCIs = smRegionalISCIs
    End Property
    Public Property Let RegionalISCIs(slvalue As String)
        smRegionalISCIs = slvalue
    End Property
    Public Property Get endDate() As String
        endDate = smEndDate
    End Property
    Public Property Let endDate(slvalue As String)
        smEndDate = slvalue
    End Property
    Public Property Get startDate() As String
        startDate = smStartDate
    End Property
    Public Property Let startDate(slvalue As String)
        smStartDate = slvalue
    End Property
    Public Property Get endTime() As String
        endTime = smEndTime
    End Property
    Public Property Let endTime(slvalue As String)
        smEndTime = slvalue
    End Property
    Public Property Get startTime() As String
        startTime = smStartTime
    End Property
    Public Property Let startTime(slvalue As String)
        smStartTime = slvalue
    End Property
    Public Property Get Sites() As String
        Sites = smSites
    End Property
    Public Property Let Sites(slvalue As String)
        smSites = slvalue
    End Property
    Public Property Get ErrorMessage() As String
        ErrorMessage = smErrorMessage
    End Property
    Public Sub Clear(blAll As Boolean)
        smEndDate = ""
        smRegionalISCIs = ""
        smSites = ""
        smStartDate = ""
        Definition = ""
        smStartDate = "00:00:00"
        smEndDate = "00:00:00"
        If blAll Then
           smSplitMajor = ";"
           smSplitMinor = "~"
           smErrorMessage = ""
        End If
    End Sub
    Public Function isSafe(slvalue As String) As Boolean
        If InStr(1, slvalue, smSplitMajor, vbBinaryCompare) > 0 And InStr(1, slvalue, smSplitMinor, vbBinaryCompare) > 0 Then
            isSafe = False
        Else
            isSafe = True
        End If
    End Function
    
Private Sub Class_Initialize()
    Clear (True)
End Sub
