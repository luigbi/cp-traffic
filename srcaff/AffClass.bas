Attribute VB_Name = "modClass"
'**************************************************************************
' Copyright: Counterpoint Software, Inc. 2002
' Created by: Doug Smith
' Date: August 2002
' Name: modCrystal
'**************************************************************************
Option Explicit
'Dan M 12/9/14
'Public Const CSISITE = "smtpauth.hosting.earthlink.net"
'Public Const CSIPORT = 587
'Public Const CSIUSERNAME = "emailSender@counterpoint.net"
'Public Const CSIPASSWORD = "Csi44Sic"


'TTP 10837,10564,10564    2023-10-30 JJB
Public Const CSISITE As String = "smtp.office365.com"
Public Const CSIPORT As Integer = 587
Public Const CSIUSERNAME As String = "noreply@counterpoint.net" '
Public Const CSIPASSWORD As String = "csi#8x21"

'Dan M 9/22/11 not currently using
Public ogEmailer As CEmail
'Dan M 12/18/09 test before closing csiNetReporter to save time.
Public bgReportModuleRunning As Boolean

'added global report object  Dan
Public ogReport As CReportHelper

'Dan M 9/14/09
Type EmailInformation
    sFromName As String
    sToName As String
    sFromAddress As String
    sToAddress As String
    sSubject As String
    sMessage As String
    sAttachment As String
'Dan M 9/7/10
    sToMultiple As String
    sCCMultiple As String   '"XX@counterpoint.net,YY@counterpoint.net,etc..."
    sBCCMulitple As String
'dan M 6/28/11 really for traffic. if setting host info, need to know tls has been set too.
    bTLSSet As Boolean
    'Dan M 11/05/09 not needed now that site options doesn't have email tab
   ' bUserFromHasPriority As Boolean  'user's from name and from address used before default in site options?
End Type
Public Sub gSendServiceEmail(slSubject As String, slBody As String)
    'Dan M 12/9/14 for monitor program
    
    If slSubject = "" Then
        slSubject = "Automated message from a client"
    End If
    Set ogEmailer = New CEmail
    
    With ogEmailer
        .FromAddress = "AClient@Counterpoint.net"
        .FromName = Trim$(sgClientName)
        .AddTOAddress "Service@counterpoint.net", "Service"
        .Subject = slSubject
        .Message = slBody
        .SetHost CSISITE, CSIPORT, CSIUSERNAME, CSIPASSWORD, False
        If Not .Send() Then
            gLogMsg "Email could not be sent from modClass-gSendServiceEmail: ", "AffErrorLog.Txt", False
        End If
    End With
    Set ogEmailer = Nothing
End Sub
