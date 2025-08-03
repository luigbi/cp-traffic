VERSION 5.00
Begin VB.Form EDSLink 
   Caption         =   "Link To EDS"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8445
   Icon            =   "EDSLink.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ckcClearAllMasterLinks 
      Caption         =   "Clear All Master Links"
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdLinktoEDS 
      Caption         =   "Link to EDS"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton cmdTestConnection 
      Caption         =   "Test Connection"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CheckBox ckcTrafficUser 
      Caption         =   "Traffic Users"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   3255
   End
   Begin VB.CheckBox ckcAffPersonnel 
      Caption         =   "Station Personnel"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CheckBox ckcTrafficSite 
      Caption         =   "Traffic Site"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.CheckBox ckcAffStation 
      Caption         =   "Affiliate Station"
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   2040
      Width           =   3375
   End
End
Attribute VB_Name = "EDSLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'EDS Support Sub Routines
'Doug Smith 7/16/15
'Copyright 2015 Counterpoint Software, Inc. All rights reserved.
'Proprietary Software, Do not copy

Option Explicit
Option Compare Text

Private Sub cmdCancel_Click()
    Unload EDSLink
End Sub

Private Sub cmdLinktoEDS_Click()

    Dim blRet As Boolean

    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    If ckcTrafficSite.Enabled = True Then
        blRet = mAddNetwork()
        blRet = mAddNetworkUsers()
        blRet = mAddStations()
        'blRet = gLinkstations()
        blRet = mAddStationUsers()
    Else
        'If ckcClearAllMasterLinks.Value = vbChecked Then
        '    blRet = gClearAllMasterLinks()
        'End If
        If ckcTrafficUser.Value = vbChecked Then
            blRet = mAddNetworkUsers()
        End If
        If ckcAffStation.Value = vbChecked Then
            blRet = mAddStations()
            'blRet = gLinkstations()
        End If
        If ckcAffPersonnel.Value = vbChecked Then
            blRet = mAddStationUsers()
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    On Error GoTo 0

End Sub

Private Sub cmdTestConnection_Click()

    Dim blRet As Boolean

    On Error GoTo ErrHand
    blRet = gGetEDSAutorization()
    Screen.MousePointer = vbHourglass
    If blRet Then
        MsgBox "Connection Was Successful", vbOKOnly
        Screen.MousePointer = vbDefault
    Else
        MsgBox "Connection Failed", vbOKOnly
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
ErrHand:
    On Error Resume Next
    Screen.MousePointer = vbDefault
    On Error GoTo 0

End Sub

Private Sub Form_Load()

    Dim blRet As Boolean
    Dim slStr As String
    Dim Saf_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    gCenterStdAlone EDSLink
    SQLQuery = "Select safClientSentToEds from Saf_Schd_Attributes"
    'Set Saf_rst = cnn.Execute(SQLQuery)
    Set Saf_rst = gSQLSelectCall(SQLQuery)
    If Saf_rst!safClientSentToEds = "Y" Then
        ckcTrafficSite.Enabled = False
    End If
    
    If Saf_rst!safClientSentToEds <> "Y" Then
        ckcTrafficSite.Enabled = True
        ckcTrafficUser.Enabled = True
        ckcAffPersonnel.Enabled = True
        ckcAffPersonnel.Enabled = True
    End If
    
    Screen.MousePointer = vbHourglass
    blRet = gGetEDSAutorization()
    If blRet Then
        Screen.MousePointer = vbDefault
        'gCenterForm EDSLink
    Else
        MsgBox "Could Not Obtain Autorization"
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    On Error GoTo 0

End Sub
Private Function mAddNetwork() As Boolean
    
    Dim slBody As String
    Dim slStr As String
    Dim blRet As Boolean
    Dim Saf_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    SQLQuery = "Select safClientSentToEds from Saf_Schd_Attributes"
    'Set Saf_rst = cnn.Execute(SQLQuery)
    Set Saf_rst = gSQLSelectCall(SQLQuery)
    If Saf_rst!safClientSentToEds = "Y" Then
        mAddNetwork = True
       ' Exit Function
    End If
    mAddNetwork = False
    slStr = Trim$(tgSpf.sGClient)
    'debug
    slBody = "{" & """" & "Name" & """" & " : " & """" & slStr & """" & "}"
    blRet = gSend_Post_APIs(slBody, "AddNetwork")
    If blRet Then
        SQLQuery = "Update Saf_Schd_Attributes Set safClientSentToEds = " & "'" & "Y" & "'" & " WHERE safVefCode = 0"
        'Set Saf_rst = cnn.Execute(SQLQuery)
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            gHandleError "TrafficErrors.txt", "EDSLink: mAddNetwork"
        End If
        mAddNetwork = True
    End If
    Saf_rst.Close
    Exit Function
ErrHand:
    On Error Resume Next
    Saf_rst.Close
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function

Private Function mAddNetworkUsers() As Boolean

    Dim ilIdx As Integer
    Dim ilTemp As Integer
    Dim ilRet As Integer
    Dim blRet As Boolean
    Dim slSQLQuery As String
    Dim slBody As String
    Dim rst_Temp As ADODB.Recordset
    'Items passed to EDS web
    Dim slUserNameEmail  As String 'username (same as their email address in this project)
    Dim slUserRights As String  'not implemented yet, pass an empty list
    Dim blIsActive As Boolean
    Dim slFullName As String
    Dim slPassword As String
    Dim slNetworkName As String
    Dim ilMinOffset As Integer
    Dim slID As String

        
    On Error GoTo ErrHand
    mAddNetworkUsers = False
    ilRet = gObtainUrf()
    On Error GoTo ErrHand
    ilRet = gObtainUrf()
    ilTemp = 0
    'For ilIdx = 1 To UBound(tgPopUrf) - 1
    For ilIdx = LBound(tgPopUrf) To UBound(tgPopUrf) - 1
        If tgPopUrf(ilIdx).sDelete <> "Y" And tgPopUrf(ilIdx).lEMailCefCode > 0 Then
            slSQLQuery = "Select cefComment from CEF_Comments_Events where cefCode = " & tgPopUrf(ilIdx).lEMailCefCode
            'Set rst_Temp = cnn.Execute(slSQLQuery)
            Set rst_Temp = gSQLSelectCall(slSQLQuery)
            If Not rst_Temp.EOF Then
                slUserNameEmail = Trim$(rst_Temp!cefComment)
            Else
                slUserNameEmail = "UnDefined"
            End If
            slUserRights = "V"
            blIsActive = 1
            slFullName = Trim$(tgPopUrf(ilIdx).sName)
            slPassword = Trim$(tgPopUrf(ilIdx).sPassword)
            slNetworkName = Trim$(tgSpf.sGClient)
            ilMinOffset = mGetUTCMinutesOffset("")
            slID = "URF" & Format(tgPopUrf(ilIdx).iAutoCode, String(5, "0"))
            slBody = "{" & """" & "UsernameEmail" & """" & ":" & """" & slUserNameEmail & """" & ","
            slBody = slBody & """" & "UserRights" & """" & ":[" & "4" & "],"
            slBody = slBody & """" & "IsActive" & """" & ":" & "true" & ","
            slBody = slBody & """" & "FullName" & """" & ":" & """" & slFullName & """" & ","
            slBody = slBody & """" & "Password" & """" & ":" & """" & slPassword & """" & ","
            slBody = slBody & """" & "UTCMinutesOffset" & """" & ":" & ilMinOffset & ","
            'slBody = slBody & """" & "ID" & """" & ":" & slID & ","
            slBody = slBody & """" & "NetworkName" & """" & ":" & """" & slNetworkName & """" & "}"
            blRet = gSend_Post_APIs(slBody, "AddOrUpdateNetworkUser")
        End If
    Next ilIdx
    rst_Temp.Close
    mAddNetworkUsers = True
    Screen.MousePointer = vbDefault
    Exit Function
ErrHand:
    rst_Temp.Close
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function

Private Function mAddStationUsers() As Boolean

    Dim rst_vef As ADODB.Recordset
    Dim rst_Shtt As ADODB.Recordset
    Dim rst_artt As ADODB.Recordset
    Dim slTemp As String
    Dim ilRet As Integer
    Dim ilIdx As Integer
    Dim blRet As Boolean
    Dim slUserNameEmail  As String 'username (same as their email address in this project)
    Dim slUserRights() As String 'not implemented yet, pass an empty list
    Dim blIsActive As Boolean
    Dim slFullName As String
    Dim slBody As String
    Dim slPassword As String
    Dim slStationName As String
    Dim slNetworkName As String
    Dim ilMinOffset As Integer
    Dim slID As String
        
    On Error GoTo ErrHand:
    Screen.MousePointer = vbHourglass
    mAddStationUsers = False
    ilRet = gObtainStations()
    ReDim slUserRights(0 To 0)
    ilIdx = 0
    SQLQuery = "Select vefCode, vefName From VEF_Vehicles where vefType in ('C', 'A', 'G')"
    'Set rst_vef = cnn.Execute(SQLQuery)
    Set rst_vef = gSQLSelectCall(SQLQuery)
    ilIdx = 0
    'Loop through all of the vehicles
    While Not rst_vef.EOF
        'Do we have a station with a name that matches the vehicle name?
        slTemp = "Select * from shtt where shttCallLetters = " & "'" & Trim$(rst_vef!VEFNAME) & "'" & "And shttType = 0"
        'Set rst_Shtt = cnn.Execute(slTemp)
        Set rst_Shtt = gSQLSelectCall(slTemp)
        If Not rst_Shtt.EOF Then
            'If rst_Shtt!shttClusterGroupID = 0 Or (rst_Shtt!shttClusterGroupID <> 0 And rst_Shtt!shttMasterCluster = "Y") Then
                'We have a station with a name that matches the vehicle name. Get the Artt personnel info
                SQLQuery = "Select * from artt where arttShttCode = " & rst_Shtt!shttCode & " And ArttType = " & "'" & "P" & "'" & " And ArttState = 0 "
                SQLQuery = SQLQuery & " and arttEmailRights In ('M', 'A', 'V')"
                'Set rst_artt = cnn.Execute(SQLQuery)
                'Set rst_artt = cnn.Execute(SQLQuery)
                Set rst_artt = gSQLSelectCall(SQLQuery)
                'If Not rst_Artt.EOF Then
                While Not rst_artt.EOF
                    If Len(rst_artt!arttEmail) > 0 Then
                        slUserNameEmail = rst_artt!arttEmail
                        ilMinOffset = mGetUTCMinutesOffset("")
                        slID = "ARTT" & Format(rst_artt!arttCode, String(5, "0"))
                        'M = Master Accept/Reject; A = Alternate Accept/Reject; V = View; N or Blank = No
                        'M = 2, A = 5, V = 4
                        If rst_artt!arttEmailRights = "M" Then
                            slUserRights(ilIdx) = 2
                        ElseIf rst_artt!arttEmailRights = "A" Then
                            slUserRights(ilIdx) = 5
                        ElseIf rst_artt!arttEmailRights = "V" Or rst_artt!arttEmailRights = " " Then
                            slUserRights(ilIdx) = 4
                        End If
                        'A=Active; D=Dormant
                        If rst_artt!ArttState = "A" Then
                            blIsActive = True
                        Else
                            blIsActive = False
                        End If
                        slFullName = Trim$(rst_artt!arttFirstName) & " " & Trim$(rst_artt!arttLastName)
                        'Don't supply a password at this time
                        slPassword = ""
                        'Future use, we don't have multiple user rights at this time
                        'ilIdx = ilIdx + 1
                        slNetworkName = Trim$(tgSpf.sGClient)
                        slStationName = Trim$(gGetCallLettersByShttCode(rst_Shtt!shttCode))
                        slBody = "{" & """" & "UsernameEmail" & """" & ":" & """" & Trim$(slUserNameEmail) & """" & ","
                        slBody = slBody & """" & "UserRights" & """" & ":[" & slUserRights(ilIdx) & "],"
                        slBody = slBody & """" & "IsActive" & """" & ":" & "true" & ","
                        slBody = slBody & """" & "FullName" & """" & ":" & """" & slFullName & """" & ","
                        slBody = slBody & """" & "Password" & """" & ":" & """" & slPassword & """" & ","
                        slBody = slBody & """" & "StationName" & """" & ":" & """" & slStationName & """" & ","
                        slBody = slBody & """" & "UTCMinutesOffset" & """" & ":" & ilMinOffset & ","
                        'slBody = slBody & """" & "ID" & """" & ":" & slID & ","
                        slBody = slBody & """" & "NetworkName" & """" & ":" & """" & slNetworkName & """" & "}"
                        blRet = gSend_Post_APIs(slBody, "AddOrUpdateStationUser")
                        rst_artt.MoveNext
                    End If
                Wend
            'End If
        End If
        rst_vef.MoveNext
    Wend
    mAddStationUsers = True
    Resume Next
    rst_vef.Close
    rst_Shtt.Close
    rst_artt.Close
    Screen.MousePointer = vbDefault
    Exit Function
ErrHand:
    Resume Next
    rst_vef.Close
    rst_Shtt.Close
    rst_artt.Close
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function


Private Function mAddStations()

    Dim slBody As String
    Dim blRet As Boolean
    Dim rst_Shtt As ADODB.Recordset
    Dim rst_vef As ADODB.Recordset
    Dim slTemp As String
    
    On Error GoTo ErrHand
    SQLQuery = "Select vefCode, vefName From VEF_Vehicles where vefType in ('C', 'A', 'G')"
    'Set rst_vef = cnn.Execute(SQLQuery)
    Set rst_vef = gSQLSelectCall(SQLQuery)
    'Loop through all of the vehicles
    While Not rst_vef.EOF
        'Do we have a station with a name that matches the vehicle name?
        slTemp = "Select * from shtt where shttCallLetters = " & "'" & Trim$(rst_vef!VEFNAME) & "'" & "And shttType = 0"
        'Set rst_Shtt = cnn.Execute(slTemp)
        Set rst_Shtt = gSQLSelectCall(slTemp)
        While Not rst_Shtt.EOF
            slBody = "{" & """" & "Name" & """" & ":" & """" & Trim$(rst_Shtt!shttCallLetters) & """" & "}"
            blRet = gSend_Post_APIs(slBody, "AddOrUpdateStation")
            sgNetworkName = Trim$(tgSpf.sGClient)
            blRet = gLinkStationToNetwork(rst_Shtt!shttCallLetters, sgNetworkName)
            rst_Shtt.MoveNext
        Wend
        rst_vef.MoveNext
    Wend
    On Error Resume Next
    rst_Shtt.Close
    Screen.MousePointer = vbDefault
    Exit Function
ErrHand:
    Resume Next
    rst_Shtt.Close
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function


Private Sub Form_Unload(Cancel As Integer)
    Unload EDSLink
    Set EDSLink = Nothing
End Sub


