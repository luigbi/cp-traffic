VERSION 5.00
Begin VB.Form frmCrystal 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCrystal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sgReportFormExe As String
Public igReportRnfCode As Integer
Public sgReportCtrlSaveName As String
Public fgReportForm As Form
Public igReportButtonIndex As Integer '0=Restore All; 1= Restore All except Dates; 2=Use default

Dim smReportCaption As String
'**************************************************************************
' Copyright: Counterpoint Software, Inc. 2002
' Created by: Doug Smith/ revised Dan Michaelson
' Date: August 2002/ Nov 2008
' Name: gCrystlReports
'
' The main entry point for all Crystal Reports. From here various subs get
' called to process the report.  Revised for Cr2008 and CsiNetReporter2 7-29-09 Dan M
'**************************************************************************
Public Sub gCrystlReports(sSQLString As String, iExportType As Integer, iRptDest As Integer, sRptName As String, sExpName As String, Optional olConditionChoice As IConditionalLogo)
    
    'gCrystlReports(sSQLString, iExportType, iRptDest, sRptName, sExpName)
    Debug.Print "sql:" & sSQLString
    'sSQLString - SQL string created by report modules
    'iExportType - Current range is 1-10; Creates Exports to PDF, CSV, DIF, XLS etc.
    'iRptDest - Current Range 0-2; 0 = Display, 1 = Print, 2 = Export
    'sRptName - Name of the report - aflabels.rpt, afdelqvh.rpt etc.
    'sExpName - File Name used when creating disk file w/o extension
    'olConditionChoice- Optional.  Class to use to choose logo.  Standard is default. Must set object before passing: dim olConditionChoice as New LogoExample.
    Dim ilRet As Integer
    Dim ilSvReportReturn As Integer
    
On Error GoTo ErrHand

    '5/26/19: Save Control settings
    gSaveReportCtrlsSetting

    If ogReport Is Nothing Then
        Set ogReport = New CReportHelper
    ElseIf ogReport.iLastPrintJob = 1 Then  'in case mulitple report
        Set ogReport = New CReportHelper
    End If
        
    If Not olConditionChoice Is Nothing Then
        Set ogReport.Alternatelogo = olConditionChoice
    End If
    ilRet = ogReport.OpenReport(sRptName)
            'Pass the SQL String to Crystal
    ogReport.Reports(sRptName).SQLQueryString = sSQLString
    mSetFormulas sRptName
     
    Screen.MousePointer = vbDefault
    If ogReport.iLastPrintJob = 1 Or ogReport.TreatAsLastReport Then    'multi reports: don't allow printing until last one
        DoEvents
            'Handle the report destination - Display, Print, Export
        If iRptDest = 0 Then                   'Display Option
            Report.Caption = smReportCaption
            Report.Show vbModal
'            frmViewReport.Caption = smReportCaption
'            frmViewReport.Show vbModal     'Affiliate is currently not vbmodal, but I lost logo because of ogreport = nothing below.
        ElseIf iRptDest = 1 Then               'Print Option.  Because affiliate doesn't currently have multiple reports, optional set to true
             gUserActivityLog "S", sgReportListName & ": Printing"
            ogReport.PrintOut False, True
            gUserActivityLog "E", sgReportListName & ": Printing"
        Else                                    'Export to File Options.  Optional set to true as above
            gUserActivityLog "S", sgReportListName & ": Exporting"
            ilRet = ogReport.Export(sExpName, iExportType, True)
            gUserActivityLog "E", sgReportListName & ": Exporting"
    '        'Don't force the users to click OK on Logs or CPs
    'dan todo !!!  also, make sure works right. changed from V81
            If Trim$(sRptName) <> "L31a.rpt" And Trim$(sRptName) <> "C17.rpt" And Trim$(sRptName) <> "L32a.rpt" And Trim$(sRptName) <> "af01.rpt" Then
                ilSvReportReturn = igReportReturn
                gMsgBox "Output Sent To: " & ogReport.Reports(ogReport.CurrentReportName).ExportOptions.DiskFileName, vbInformation
                If igReportSource = 2 Then
                    igReportReturn = ilSvReportReturn
                End If
            End If
        End If
        Set ogReport = Nothing
    End If 'multi print job and done
' dan m 9/15/11 moved above. note that display handled in report
'    Select Case iRptDest
'        Case 1  'Print
'            gUserActivityLog "E", sgReportListName & ": Printing"
'        Case 2  'Export
'            gUserActivityLog "E", sgReportListName & ": Exporting"
'        'handled in display form
''        Case Else
''            gUserActivityLog "E", sgReportListName & ": Display"
'    End Select
    
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmCrystal - gCrystlReports: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
        Set ogReport = Nothing
    End If
   
End Sub
Private Sub mSetFormulas(sRptName As String)
On Error GoTo ErrHand
    '2/28/22 - JW - Fix TTP 10403 - Affiliate Spot MGMT report showing extra vehicles
    Dim sSelectionFormula As String
    Dim ilRet As Integer
   ' DoEvents
    Select Case sRptName        'report names are case sensitive to calling rtn
        Case "aflabels.rpt"
           smReportCaption = "Mailing Labels"
           ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula1)
           ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula2)
           ilRet = ogReport.SetFormula("Contact", sgCrystlFormula3)
           ilRet = ogReport.SetFormula("WhichMethod", sgCrystlFormula4)

        Case "aflabel3.rpt"
            smReportCaption = "Mailing Labels"
           ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula1)
           ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula2)
           ilRet = ogReport.SetFormula("Contact", sgCrystlFormula3)
           ilRet = ogReport.SetFormula("WhichMethod", sgCrystlFormula4)

        Case "aflabship.rpt"
            smReportCaption = "Mailing Labels"
           ilRet = ogReport.SetFormula("Contact", sgCrystlFormula3)
           ilRet = ogReport.SetFormula("WhichMethod", sgCrystlFormula4)

        Case "afStatin.rpt"
            smReportCaption = "Station Information"
           ilRet = ogReport.SetFormula("SortBy", sgCrystlFormula1)
           ilRet = ogReport.SetFormula("Subsort", sgCrystlFormula2)
        Case "afVhStVh.rpt", "afAgreeOwnerVh.rpt"
            smReportCaption = "Agreements"
            ilRet = ogReport.SetFormula("ShowPhone", "'" & sgCrystlFormula1 & "'")
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)
            ilRet = ogReport.SetFormula("DatesEntered", "'" & sgCrystlFormula4 & "'")
            ilRet = ogReport.SetFormula("DatesBeginning", "'" & sgCrystlFormula5 & "'")
            ilRet = ogReport.SetFormula("DatesEnding", "'" & sgCrystlFormula6 & "'")
            ilRet = ogReport.SetFormula("MulticastOnly", "'" & sgCrystlFormula8 & "'")
            ilRet = ogReport.SetFormula("ShowPledgeOrPgm", "'" & sgCrystlFormula9 & "'")            '1-26-12 show program or avail times (for start/end time of pgm)
            ilRet = ogReport.SetFormula("Service", sgCrystlFormula11)
            ilRet = ogReport.SetFormula("ShowContactInfo", "'" & sgCrystlFormula12 & "'")            '2-20-15 Show contract info
            ilRet = ogReport.SetFormula("ShowComments", "'" & sgCrystlFormula13 & "'")               '2-20-15 Show agreement comments

        Case "AfAgreeExpCodes.rpt"             '7-25-12  affiliate agreement for export codes
            smReportCaption = "Affiliate Agreements"
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)
            ilRet = ogReport.SetFormula("DatesEntered", "'" & sgCrystlFormula4 & "'")
            ilRet = ogReport.SetFormula("DatesBeginning", "'" & sgCrystlFormula5 & "'")
            ilRet = ogReport.SetFormula("DatesEnding", "'" & sgCrystlFormula6 & "'")
            ilRet = ogReport.SetFormula("MulticastOnly", "'" & sgCrystlFormula8 & "'")
            ilRet = ogReport.SetFormula("SortBy", "'" & sgCrystlFormula10 & "'")
        
        Case "afVhStSt.rpt", "afAgreeOther.rpt", "AfAgreeStnFmt.rpt"                                'added SHOW option: FORMAT  Date: 8/13/2018 FYM
            smReportCaption = "Affiliate Agreements"
            ilRet = ogReport.SetFormula("ShowPhone", "'" & sgCrystlFormula1 & "'")
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)
            ilRet = ogReport.SetFormula("ShowStationInfo", "'" & sgCrystlFormula7 & "'")
            ilRet = ogReport.SetFormula("DatesEntered", "'" & sgCrystlFormula4 & "'")
            ilRet = ogReport.SetFormula("DatesBeginning", "'" & sgCrystlFormula5 & "'")
            ilRet = ogReport.SetFormula("DatesEnding", "'" & sgCrystlFormula6 & "'")
            ilRet = ogReport.SetFormula("MulticastOnly", "'" & sgCrystlFormula8 & "'")
            ilRet = ogReport.SetFormula("ShowPledgeOrPgm", "'" & sgCrystlFormula9 & "'")            '1-26-12 show program or avail times (for start/end time of pgm)
            ilRet = ogReport.SetFormula("Service", sgCrystlFormula11)
            ilRet = ogReport.SetFormula("ShowContactInfo", "'" & sgCrystlFormula12 & "'")            '2-20-15 Show contract info
            ilRet = ogReport.SetFormula("ShowComments", "'" & sgCrystlFormula13 & "'")               '2-20-15 Show agreement comments
            If sRptName = "afAgreeOther.rpt" Then
                ilRet = ogReport.SetFormula("SortBy", "'" & sgCrystlFormula10 & "'")
            End If
        Case "afdelqvh.rpt", "afdelqPro.rpt", "afNCRvh.rpt", "afNCRPro.rpt"
            smReportCaption = "Overdue Affidavits"
'            ogReport.Reports(sRptName).FormulaFields(2).text = sgCrystlFormula1 'StartDate
'            ogReport.Reports(sRptName).FormulaFields(3).text = sgCrystlFormula2 'EndDate
'            ogReport.Reports(sRptName).FormulaFields(6).text = "'" & sgCrystlFormula3 & "'" 'NewPage
'            ogReport.Reports(sRptName).FormulaFields(26).text = "'" & sgCrystlFormula4 & "'" 'Include action comments
'            ogReport.Reports(sRptName).FormulaFields(27).text = sgCrystlFormula5  'effective comment date
'            ogReport.Reports(sRptName).FormulaFields(4).text = "'" & sgCrystlFormula6 & "'"  ' Sort By. "V" or "A"  (vehicle/Affiliate ae)
'            ogReport.Reports(sRptName).FormulaFields(32).text = "'" & sgCrystlFormula7 & "'"  ' Show "F" = fax or "P" = password
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula1)
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula2)
            ilRet = ogReport.SetFormula("NewPage", "'" & sgCrystlFormula3 & "'") '6
            ilRet = ogReport.SetFormula("ShowActionComment", "'" & sgCrystlFormula4 & "'") '26
            ilRet = ogReport.SetFormula("SortBy", "'" & sgCrystlFormula6 & "'")     '4
            ilRet = ogReport.SetFormula("FaxOrPasword", "'" & sgCrystlFormula7 & "'") '32
            ilRet = ogReport.SetFormula("EffecActionDate", sgCrystlFormula5)    '27

            If sRptName = "afNCRPro.rpt" Or sRptName = "afNCRvh.rpt" Then
                smReportCaption = "Critically Overdue"
                ilRet = ogReport.SetFormula("UpdateFlag", "'" & sgCrystlFormula8 & "'")
                ilRet = ogReport.SetFormula("HonorSuppressNotice", "'" & sgCrystlFormula9 & "'")
                ilRet = ogReport.SetFormula("OldestNCRDate", sgCrystlFormula10)             'Oldest delinquent date to include
            Else
                ilRet = ogReport.SetFormula("UnPostedTypes", "'" & sgCrystlFormula11 & "'")      'post type:  partial, unposted or both
            End If
        Case "afdelqst.rpt", "afNCRst.rpt", "afdelqAud.rpt"
            smReportCaption = "Overdue Affidavits"
'            ogReport.Reports(sRptName).FormulaFields(2).text = sgCrystlFormula1 'StartDate
'            ogReport.Reports(sRptName).FormulaFields(3).text = sgCrystlFormula2 'EndDate
'            ogReport.Reports(sRptName).FormulaFields(6).text = "'" & sgCrystlFormula3 & "'" 'NewPage
'            ogReport.Reports(sRptName).FormulaFields(28).text = "'" & sgCrystlFormula4 & "'" 'Include action comments
'            ogReport.Reports(sRptName).FormulaFields(27).text = sgCrystlFormula5  'effective comment date
'            ogReport.Reports(sRptName).FormulaFields(4).text = "'" & sgCrystlFormula6 & "'"  ' Sort By. "S", "M" or "R" (Station, Market or Rank)
'            ogReport.Reports(sRptName).FormulaFields(37).text = "'" & sgCrystlFormula7 & "'"  ' Show "F" = fax or "P" = password
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula1)
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula2)
            ilRet = ogReport.SetFormula("NewPage", "'" & sgCrystlFormula3 & "'")
            ilRet = ogReport.SetFormula("ShowActionComment", "'" & sgCrystlFormula4 & "'") '28
            ilRet = ogReport.SetFormula("SortBy", "'" & sgCrystlFormula6 & "'") '4
            ilRet = ogReport.SetFormula("FaxOrPasword", "'" & sgCrystlFormula7 & "'") '37
            ilRet = ogReport.SetFormula("EffecActionDate", sgCrystlFormula5)     '27
            If sRptName = "afNCRst.rpt" Then
                smReportCaption = "Critically Overdue"
                ilRet = ogReport.SetFormula("UpdateFlag", "'" & sgCrystlFormula8 & "'") '40
                ilRet = ogReport.SetFormula("HonorSuppressNotice", "'" & sgCrystlFormula9 & "'")    '41
               ilRet = ogReport.SetFormula("OldestNCRDate", sgCrystlFormula10)             'Oldest delinquent date to include
            Else
                ilRet = ogReport.SetFormula("UnPostedTypes", "'" & sgCrystlFormula11 & "'")      'post type:  partial, unposted or both
            End If

        Case "AfMissWks.rpt", "AfMissWksPro.rpt", "AfMissWksVh.rpt", "AfMissWksAud.rpt"
            smReportCaption = "Affiliate Missing Weeks"
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula1)
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula2)
            ilRet = ogReport.SetFormula("NewPage", "'" & sgCrystlFormula3 & "'")
            ilRet = ogReport.SetFormula("ShowActionComment", "'N'") '28
            ilRet = ogReport.SetFormula("SortBy", "'" & sgCrystlFormula6 & "'") '4
            'Default to Not to skip to new page in Crystal report
            'Default to sort by Station in Crystal report
        Case "afAdvClr.Rpt"
            smReportCaption = "Advertiser Clearances"
            ogReport.Reports(sRptName).FormulaFields(7).Text = sgCrystlFormula1 'BaseDate
'            ogReport.Reports(sRptName).FormulaFields(31).text = "'" & sgCrystlFormula2 & "'" 'Sortby
'            ogReport.Reports(sRptName).FormulaFields(34).text = "'" & sgCrystlFormula3 & "'" 'AirInDPOption
            ilRet = ogReport.SetFormula("BaseDate", sgCrystlFormula1)     '7
            ilRet = ogReport.SetFormula("SortBy", "'" & sgCrystlFormula2 & "'") '31
            ilRet = ogReport.SetFormula("AirInDPOption", "'" & sgCrystlFormula3 & "'") '34

        Case "AfPledge.rpt"
            smReportCaption = "Pledges"
'            ogReport.Reports(sRptName).FormulaFields(3).text = sgCrystlFormula2 'StartDate
'            ogReport.Reports(sRptName).FormulaFields(4).text = sgCrystlFormula3 'EndDate
'            ogReport.Reports(sRptName).FormulaFields(6).text = "'" & sgCrystlFormula1 & "'" 'ExceptFlag
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)     '3
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)     '4
            ilRet = ogReport.SetFormula("ExceptFlag", "'" & sgCrystlFormula1 & "'") '6
            ilRet = ogReport.SetFormula("PageSkip", "'" & sgCrystlFormula4 & "'")
            
        Case "L31a.rpt", "L32a.rpt"
            smReportCaption = "Affiliate Log Report"
'            ogReport.Reports(sRptName).FormulaFields(9).text = sgCrystlFormula1 'StdYear
'            ogReport.Reports(sRptName).FormulaFields(12).text = "'" & sgCrystlFormula2 & "'" 'Week
'            ogReport.Reports(sRptName).FormulaFields(10).text = sgCrystlFormula3 'InputDate
'            ogReport.Reports(sRptName).FormulaFields(11).text = "'" & sgCrystlFormula4 & "'" 'NumberDays
            ilRet = ogReport.SetFormula("StdYear", sgCrystlFormula1)     '9
            ilRet = ogReport.SetFormula("Week", "'" & sgCrystlFormula2 & "'") '12
            ilRet = ogReport.SetFormula("InputDate", sgCrystlFormula3)     '10
            ilRet = ogReport.SetFormula("NumberDays", "'" & sgCrystlFormula4 & "'") '11

'        Case "L32a.rpt"
'            smReportCaption = "Affiliate Log Report"
'            ogReport.Reports(sRptName).FormulaFields(3).text = sgCrystlFormula1 'StdYear
'            ogReport.Reports(sRptName).FormulaFields(4).text = "'" & sgCrystlFormula2 & "'" 'Week
'            ogReport.Reports(sRptName).FormulaFields(5).text = sgCrystlFormula3 'InputDate
'            '8-25-03 invalid formula index changed from 12 to 6
'            ogReport.Reports(sRptName).FormulaFields(6).text = sgCrystlFormula4   'NumberDays

        Case "C17.rpt"
            smReportCaption = "Affiliate CP Report"
'            ogReport.Reports(sRptName).FormulaFields(9).text = sgCrystlFormula1 'StartDate
'            ogReport.Reports(sRptName).FormulaFields(10).text = sgCrystlFormula2 'EndDate
'            ogReport.Reports(sRptName).FormulaFields(22).text = "'" & sgCrystlFormula3 & "'" 'CoverPageOnly
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula1)     '9
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula2)     '10
            ilRet = ogReport.SetFormula("CoverPageOnly", "'" & sgCrystlFormula3 & "'") '22

        Case "AfStnClr.rpt"
            smReportCaption = "Spot Clearance"
'            ogReport.Reports(sRptName).FormulaFields(3).text = sgCrystlFormula1    'sort by
'            ogReport.Reports(sRptName).FormulaFields(6).text = sgCrystlFormula2    'Start Date
'            ogReport.Reports(sRptName).FormulaFields(7).text = sgCrystlFormula3    'End Date
'            ogReport.Reports(sRptName).FormulaFields(22).text = sgCrystlFormula4    'New Page
'            ogReport.Reports(sRptName).FormulaFields(23).text = sgCrystlFormula5    'Show certification
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)     '6
            ilRet = ogReport.SetFormula("SortBy", sgCrystlFormula1)     '3
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)     '7
            ilRet = ogReport.SetFormula("NewPage", sgCrystlFormula4)     '22
            ilRet = ogReport.SetFormula("ShowCertification", sgCrystlFormula5)     '23
            ilRet = ogReport.SetFormula("UserTimes", sgCrystlFormula11)     '27
            ilRet = ogReport.SetFormula("ExcludeMissedIfMG", sgCrystlFormula12)        '9-12-16
            'TTP 10067 - Spot Clearance report - date/time filter stopped working
            'the issue is, AfStnClr.rpt was converted from ODBC to use Pervasive datasource (Maybe for performance?).  Which now; using Pervasive datasource Driver, Crystal doesnt support the SQL Query, and all the records aren't filtered.
            'Changes made to AfStnClr.rpt to filter the generated records based on the @StartDate, @EndDate and @UserTimes formula's
            ilRet = ogReport.SetFormula("UserAfrGenDate", "'" & Format$(sgGenDate, sgSQLDateForm) & "'")
            ilRet = ogReport.SetFormula("UserAfrGenTime", Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))))
            
    

'        Case "AfPostStaAct.rpt", "AfPostVehAct.rpt", "AfLogStaAct.rpt", "AfLogStaInAct.rpt", "AfLogVehAct.rpt", "AfLogVehInAct.rpt"
'            ilRet = ogReport.SetFormula("SortBy", sgCrystlFormula1)     '1
'            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)     '4
'            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)     '5
            
        Case "AfPostStaAct.rpt"
            'smReportCaption = "Web Posting Activity"
            smReportCaption = "Affiliate Affidavit Posting Activity"
            ilRet = ogReport.SetFormula("SortBy", sgCrystlFormula1)     '1
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)     '4
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)     '5

'            ogReport.Reports(sRptName).FormulaFields(1).Text = sgCrystlFormula1    'sort by
'            ogReport.Reports(sRptName).FormulaFields(4).Text = sgCrystlFormula2    'Start Date
'            ogReport.Reports(sRptName).FormulaFields(5).Text = sgCrystlFormula3    'End Date
        Case "AfPostVehAct.rpt"
            smReportCaption = "Affiliate Affidavit Posting Activity"
            ilRet = ogReport.SetFormula("SortBy", sgCrystlFormula1)     '1
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)     '4
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)     '5
            'smReportCaption = "Web Posting Activity"
            smReportCaption = "Affiliate Affidavit Posting Activity"
'            ogReport.Reports(sRptName).FormulaFields(1).Text = sgCrystlFormula1    'sort by
'            ogReport.Reports(sRptName).FormulaFields(4).Text = sgCrystlFormula2    'Start Date
'            ogReport.Reports(sRptName).FormulaFields(5).Text = sgCrystlFormula3    'End Date
        Case "AfLogStaAct.rpt"
            smReportCaption = "Web Log Activity"
            ilRet = ogReport.SetFormula("SortBy", sgCrystlFormula1)     '1
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)     '4
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)     '5
'            ogReport.Reports(sRptName).FormulaFields(1).Text = sgCrystlFormula1    'sort by
'            ogReport.Reports(sRptName).FormulaFields(4).Text = sgCrystlFormula2    'Start Date
'            ogReport.Reports(sRptName).FormulaFields(5).Text = sgCrystlFormula3    'End Date
        Case "AfLogStaInAct.rpt"
            smReportCaption = "Web Log Inactivity"
            ilRet = ogReport.SetFormula("SortBy", sgCrystlFormula1)     '1
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)     '4
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)     '5
'            ogReport.Reports(sRptName).FormulaFields(1).Text = sgCrystlFormula1    'sort by
'            ogReport.Reports(sRptName).FormulaFields(4).Text = sgCrystlFormula2    'Start Date
'            ogReport.Reports(sRptName).FormulaFields(5).Text = sgCrystlFormula3    'End Date
        Case "AfLogVefAct.rpt"
            smReportCaption = "Web Log Activity"
            ilRet = ogReport.SetFormula("SortBy", sgCrystlFormula1)     '1
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)     '4
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)     '5
'            ogReport.Reports(sRptName).FormulaFields(1).Text = sgCrystlFormula1    'sort by
'            ogReport.Reports(sRptName).FormulaFields(4).Text = sgCrystlFormula2    'Start Date
'            ogReport.Reports(sRptName).FormulaFields(5).Text = sgCrystlFormula3    'End Date
        Case "AfLogVehInAct.rpt"
            smReportCaption = "Web Log Inactivity"
            ilRet = ogReport.SetFormula("SortBy", sgCrystlFormula1)     '1
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)     '4
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)     '5
'            ogReport.Reports(sRptName).FormulaFields(1).Text = sgCrystlFormula1    'sort by
'            ogReport.Reports(sRptName).FormulaFields(4).Text = sgCrystlFormula2    'Start Date
'            ogReport.Reports(sRptName).FormulaFields(5).Text = sgCrystlFormula3    'End Date


        Case "AlertStatusSQL.rpt"
            smReportCaption = "Alert Status"
'            ogReport.Reports(sRptName).FormulaFields(3).text = sgCrystlFormula1    'sort by
'            ogReport.Reports(sRptName).FormulaFields(14).text = sgCrystlFormula2    'Start Date
            ilRet = ogReport.SetFormula("SortBy", sgCrystlFormula1)     '3
            ilRet = ogReport.SetFormula("EffClearDate", sgCrystlFormula2)     '14

        Case "AfCounts.rpt"
            smReportCaption = "Affiliate Clearance Counts"
'            ogReport.Reports(sRptName).FormulaFields(16).text = sgCrystlFormula1    'Start Date
'            ogReport.Reports(sRptName).FormulaFields(17).text = sgCrystlFormula2     'end date
'            ogReport.Reports(sRptName).FormulaFields(1).text = sgCrystlFormula3     'Sort group #1
'            ogReport.Reports(sRptName).FormulaFields(26).text = sgCrystlFormula4     'Sort group #2
'            ogReport.Reports(sRptName).FormulaFields(27).text = sgCrystlFormula5     'Sort group #3
'            ogReport.Reports(sRptName).FormulaFields(28).text = sgCrystlFormula6     'Sort group #4
'            ogReport.Reports(sRptName).FormulaFields(57).text = sgCrystlFormula7     'Page Skip group #1
'            ogReport.Reports(sRptName).FormulaFields(58).text = sgCrystlFormula8     'Page Skip group #2
'            ogReport.Reports(sRptName).FormulaFields(59).text = sgCrystlFormula9     'Page Skip group #3
'            ogReport.Reports(sRptName).FormulaFields(60).text = sgCrystlFormula10     'Page Skip group #4
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula1)     '16
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula2)     '17
            ilRet = ogReport.SetFormula("SortBy1", sgCrystlFormula3)     '1
            ilRet = ogReport.SetFormula("SortBy2", sgCrystlFormula4)     '26
            ilRet = ogReport.SetFormula("SortBy3", sgCrystlFormula5)     '27
            ilRet = ogReport.SetFormula("SortBy4", sgCrystlFormula6)     '28
            ilRet = ogReport.SetFormula("PageSkipSort1", sgCrystlFormula7)     '57
            ilRet = ogReport.SetFormula("PageSkipSort2", sgCrystlFormula8)     '58
            ilRet = ogReport.SetFormula("PageSkipSort3", sgCrystlFormula9)     '59
            ilRet = ogReport.SetFormula("PageSkipSort4", sgCrystlFormula10)     '60

        Case "AfPgmClr.rpt", "AfPgmClrMin.rpt"
            smReportCaption = "Program Clearance"
'            ogReport.Reports(sRptName).FormulaFields(1).text = sgCrystlFormula1    'dates & times
'            ogReport.Reports(sRptName).FormulaFields(3).text = sgCrystlFormula4    'inclusions/exclusions on named avails
'            If sRptName = "AfPgmClr.rpt" Then                           'units option
'                ogReport.Reports(sRptName).FormulaFields(84).text = "'" & sgCrystlFormula5 & "'"    'inclusions/exclusions on named avails
'            Else                                                        'minutes option
'                ogReport.Reports(sRptName).FormulaFields(87).text = "'" & sgCrystlFormula5 & "'"    'inclusions/exclusions on named avails
'            End If
            ilRet = ogReport.SetFormula("DatesTimes", sgCrystlFormula1)     '1
            ilRet = ogReport.SetFormula("NamedAvails", sgCrystlFormula4)     '3
            ilRet = ogReport.SetFormula("StatusSelections", "'" & sgCrystlFormula5 & "'") '87 (afpgmclrmin) or 84 (afpgmclr)

        Case "AfPldgAir.rpt"
            smReportCaption = "Pledged vs Aired Clearance"
'            ogReport.Reports(sRptName).FormulaFields(3).text = sgCrystlFormula1    'sort option
'            ogReport.Reports(sRptName).FormulaFields(6).text = sgCrystlFormula2     'start date
'            ogReport.Reports(sRptName).FormulaFields(7).text = sgCrystlFormula3     'end date
'            ogReport.Reports(sRptName).FormulaFields(19).text = sgCrystlFormula4    'discrep only (Y/N)
'            ogReport.Reports(sRptName).FormulaFields(20).text = sgCrystlFormula5    'Separate Statuses (Y/N)
'            ogReport.Reports(sRptName).FormulaFields(24).text = "'" & sgCrystlFormula6 & "'"   'status inclusion/exlusions
'            ogReport.Reports(sRptName).FormulaFields(25).text = sgCrystlFormula7    'Subsort (minor sort) by pledge end date & start time or Air Time
'            ogReport.Reports(sRptName).FormulaFields(28).text = sgCrystlFormula8    'show the status code description on report
'            ogReport.Reports(sRptName).FormulaFields(31).text = sgCrystlFormula9    'using fed or air dates for selectivity
            ilRet = ogReport.SetFormula("SortBy", sgCrystlFormula1)     '3
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)     '6
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)     '7
            ilRet = ogReport.SetFormula("DiscrepOnly", sgCrystlFormula4)     '19
            ilRet = ogReport.SetFormula("SeparateStatus", sgCrystlFormula5)     '20
            ilRet = ogReport.SetFormula("StatusSelections", "'" & sgCrystlFormula6 & "'") '24
            ilRet = ogReport.SetFormula("TimeSort", sgCrystlFormula7)     '25
            ilRet = ogReport.SetFormula("ShowStatusCodes", sgCrystlFormula8)     '28
            ilRet = ogReport.SetFormula("UsingWhichDates", sgCrystlFormula9)     '31
            ilRet = ogReport.SetFormula("UserTimes", "'" & sgCrystlFormula11 & "'")     '37
            ilRet = ogReport.SetFormula("SuppressCounts", sgCrystlFormula12)      'suppress spot counts
            ilRet = ogReport.SetFormula("AvailName", sgCrystlFormula13)      'avail named selected

        Case "AfFedAir.rpt", "AfFedAirStatusDiscrp.rpt"         '12-11-13 STatus Discrepancy option
            If sRptName = "AfFedAirStatusDiscrp.rpt" Then
                smReportCaption = "Fed vs Aired Status Discrepancy"
            Else
                smReportCaption = "Fed vs Aired Clearance"
            End If
'            ogReport.Reports(sRptName).FormulaFields(3).text = sgCrystlFormula1    'sort option
'            ogReport.Reports(sRptName).FormulaFields(6).text = sgCrystlFormula2     'start date
'            ogReport.Reports(sRptName).FormulaFields(7).text = sgCrystlFormula3     'end date
'            ogReport.Reports(sRptName).FormulaFields(19).text = sgCrystlFormula4    'discrep only (Y/N)
'            ogReport.Reports(sRptName).FormulaFields(20).text = sgCrystlFormula5    'Separate Statuses (Y/N)
'            ogReport.Reports(sRptName).FormulaFields(24).text = "'" & sgCrystlFormula6 & "'"   'status inclusion/exlusions
'            ogReport.Reports(sRptName).FormulaFields(25).text = sgCrystlFormula7    'Subsort (minor sort) by pledge end date & start time or Air Time
'            ogReport.Reports(sRptName).FormulaFields(28).text = sgCrystlFormula8    'show the status code description on report
'            ogReport.Reports(sRptName).FormulaFields(31).text = sgCrystlFormula9    'using fed or air dates for selectivity
'            ogReport.Reports(sRptName).FormulaFields(34).text = sgCrystlFormula10    'show exact times aired (for feed vs aired report)
            ilRet = ogReport.SetFormula("SortBy", sgCrystlFormula1)     '3
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)     '6
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)     '7
            ilRet = ogReport.SetFormula("DiscrepOnly", sgCrystlFormula4)     '19
            ilRet = ogReport.SetFormula("SeparateStatus", sgCrystlFormula5)     '20
            ilRet = ogReport.SetFormula("StatusSelections", "'" & sgCrystlFormula6 & "'") '24
            ilRet = ogReport.SetFormula("TimeSort", sgCrystlFormula7)     '25
            ilRet = ogReport.SetFormula("ShowStatusCodes", sgCrystlFormula8)     '28
            ilRet = ogReport.SetFormula("UsingWhichDates", sgCrystlFormula9)     '31
            ilRet = ogReport.SetFormula("ShowExactFeed", sgCrystlFormula10)     '34
            ilRet = ogReport.SetFormula("UserTimes", "'" & sgCrystlFormula11 & "'")     '39
            'sgCrystalFormula12 used in afPldgAir.rpt which comes from the same routines as Fed vs aired
            ilRet = ogReport.SetFormula("AvailName", sgCrystlFormula13)      'avail named selected

        Case "AfVerify.rpt", "AfVerifySum.rpt"    ', "afMonExportsTTx.rpt"
            smReportCaption = "Feed Verification"
'            ogReport.Reports(sRptName).FormulaFields(2).text = "'" & sgCrystlFormula1 & "'"   'dates requested
            ilRet = ogReport.SetFormula("DatesRequested", "'" & sgCrystlFormula1 & "'") '2
            ilRet = ogReport.SetFormula("ShowBy", sgCrystlFormula2)

        Case "AfJournal.rpt"             '4-26-07
            smReportCaption = "Export Journal"
'            ogReport.Reports(sRptName).FormulaFields(9).text = "'" & sgCrystlFormula1 & "'"    'Activity dates header
'            ogReport.Reports(sRptName).FormulaFields(10).text = "'" & sgCrystlFormula2 & "'"    'Log dates header
'            ogReport.Reports(sRptName).FormulaFields(1).text = sgCrystlFormula3     'Include Stations
'            ogReport.Reports(sRptName).FormulaFields(11).text = sgCrystlFormula4   'discrep only (Y/N)
'            ogReport.Reports(sRptName).FormulaFields(4).text = sgCrystlFormula5    'Sort option
            ilRet = ogReport.SetFormula("ActivityDatesHdr", "'" & sgCrystlFormula1 & "'") '9
            ilRet = ogReport.SetFormula("LogDatesHdr", "'" & sgCrystlFormula2 & "'") '10
            ilRet = ogReport.SetFormula("IncludeStation", sgCrystlFormula3)     '1
            ilRet = ogReport.SetFormula("DiscrepOnly", sgCrystlFormula4)     '11
            ilRet = ogReport.SetFormula("SortBy", sgCrystlFormula5)     '4

        Case "afMonExports.rpt"  ' see afVerify.rpt
            smReportCaption = "Export Monitoring"
            ilRet = ogReport.SetFormula("DatesRequested", "'" & sgCrystlFormula1 & "'") '2
'            ogReport.Reports(sRptName).FormulaFields(2).text = "'" & sgCrystlFormula1 & "'"   'dates requested

        Case "AFMarkAssign.rpt"
            smReportCaption = "Market Assignment"
'            ogReport.Reports(sRptName).FormulaFields(3).text = sgCrystlFormula1    'Sort By
            ilRet = ogReport.SetFormula("Sort", sgCrystlFormula1)     '3
        Case "afUserOptions.rpt"
            smReportCaption = "User Options"
            ilRet = ogReport.SetFormula("User", sgCrystlFormula1)
            'ust state: 0 for active, 1 for dormant.  sgcrystlformula2 = 0 for don't include dormant(false), 1 for include dormant (true)
            ilRet = ogReport.SetSelection("{ust.ustState} <= " & sgCrystlFormula2)
        Case "afSiteOptions.rpt"
            smReportCaption = "Site Options"
        Case "AfRegCopy.rpt"
            smReportCaption = "Regional Affiliate Copy Assignment"
            'fNewForm.Caption = "Regional Copy Assignment"
            ilRet = ogReport.SetFormula("DatesRequested", sgCrystlFormula1)       'DatesRequested xx/xx/xx-xx/xx/xx
            'fNewForm.Report.FormulaFields(13).text = sgCrystlFormula1    'dates
            ilRet = ogReport.SetFormula("AffiliateSort", sgCrystlFormula2)        'AffiliateSort (v = vehicle, s = station)
            'fNewForm.Report.FormulaFields(3).text = sgCrystlFormula2        'which sort (V = vehicle, s = station)
            ilRet = ogReport.SetSelection("{GRF_Generic_Report.grfGenTime} =" & sgCrystlFormula3 & " and {GRF_Generic_Report.grfGenDate} = " & sgCrystlFormula4)
            ilRet = ogReport.SetFormula("ShowOnlyRegAsgn", sgCrystlFormula5)        'Show only regional copy assigned (Y/N)
        Case "AfRegCopyTrace.rpt"       '2-23-10
            smReportCaption = "Regional Affiliate Copy Tracing"
            ilRet = ogReport.SetFormula("DatesRequested", sgCrystlFormula1)       'DatesRequested xx/xx/xx-xx/xx/xx
            ilRet = ogReport.SetFormula("AffiliateSort", sgCrystlFormula2)        'AffiliateSort (v = vehicle, s = station)
            ilRet = ogReport.SetSelection("{GRF_Generic_Report.grfGenTime} =" & sgCrystlFormula3 & " and {GRF_Generic_Report.grfGenDate} = " & sgCrystlFormula4)
            ilRet = ogReport.SetFormula("ShowOnlyRegAsgn", sgCrystlFormula5)        'Show only regional copy assigned (Y/N)
        Case "AfDMAMkt.rpt", "AfMSAMkt.rpt", "AfFormat.rpt", "AfState.rpt", "AfTimeZone.rpt", "AfVehicle.rpt", "AfOwner.rpt"
            smReportCaption = "Groups"
            'fNewForm.Caption = "Groups"
        Case "AfAdvFulFill.rpt", "AfAdvPlace.rpt"
            If sRptName = "AfAdvFulFill.rpt" Then
                smReportCaption = "Advertiser Fulfillment"
            Else
                smReportCaption = "Advertiser Placement"
            End If
            'fNewForm.Caption = "Advertiser Fulfillment"
            'fNewForm.Report.FormulaFields(27).Text = sgCrystlFormula1    'sort option (0 - 8)
            ilRet = ogReport.SetFormula("SortByFromAff", sgCrystlFormula1)
            'fNewForm.Report.FormulaFields(3).Text = sgCrystlFormula2     'start date
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)
            'fNewForm.Report.FormulaFields(4).Text = sgCrystlFormula3     'end date
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)
            'fNewForm.Report.FormulaFields(14).Text = sgCrystlFormula4    'using fed (F) or air (A) dates for selectivity
            ilRet = ogReport.SetFormula("UsingWhichDates", sgCrystlFormula4)
            'fNewForm.Report.FormulaFields(16).Text = "'" & sgCrystlFormula5 & "'" 'User times
            ilRet = ogReport.SetFormula("UserTimes", "'" & sgCrystlFormula5 & "'")
           'fNewForm.Report.FormulaFields(7).Text = "'" & sgCrystlFormula6 & "'"   'status inclusion/exlusion list
            ilRet = ogReport.SetFormula("StatusSelections", "'" & sgCrystlFormula6 & "'")
           'fNewForm.Report.FormulaFields(6).Text = sgCrystlFormula7    'show the status code description on report (Y/N)
            ilRet = ogReport.SetFormula("ShowStatusCodes", sgCrystlFormula7)
            'fNewForm.Report.FormulaFields(5).Text = "'" & sgCrystlFormula8 & "'"    'text description of extra field to show
            ilRet = ogReport.SetFormula("ExtraField", "'" & sgCrystlFormula8 & "'")
            'fNewForm.Report.FormulaFields(26).Text = sgCrystlFormula9    'Include (I)/Exclude (E) station subtotal spot counts
            ilRet = ogReport.SetFormula("InclSubTotals", sgCrystlFormula9)
            ilRet = ogReport.SetFormula("MarkRegional", sgCrystlFormula10)  'highlight regional spots
            ilRet = ogReport.SetFormula("AvailName", sgCrystlFormula11)      'avail named selected
            ilRet = ogReport.SetFormula("NonCompliant", sgCrystlFormula12)      'non-compliance indicator

        Case "AfStnMgmt.rpt"
            smReportCaption = "Station Management"
        Case "afComments.rpt"
            smReportCaption = "Contact Comments"
            ilRet = ogReport.SetFormula("EnterDates", "'" & sgCrystlFormula1 & "'")
            ilRet = ogReport.SetFormula("FollowupDates", "'" & sgCrystlFormula2 & "'")

            ilRet = ogReport.SetFormula("SortBy1", sgCrystlFormula3)
            ilRet = ogReport.SetFormula("SortBy2", sgCrystlFormula4)
            ilRet = ogReport.SetFormula("SortBy3", sgCrystlFormula5)
            ilRet = ogReport.SetFormula("TotalsBy", sgCrystlFormula6)

            ilRet = ogReport.SetFormula("PageSkipSort1", sgCrystlFormula7)
            ilRet = ogReport.SetFormula("PageSkipSort2", sgCrystlFormula8)
            ilRet = ogReport.SetFormula("PageSkipSort3", sgCrystlFormula9)
            ilRet = ogReport.SetFormula("DoneUndone", sgCrystlFormula10)
            ilRet = ogReport.SetFormula("InternalID", sgCrystlFormula11)
        Case "afStnFilter.rpt"
            smReportCaption = "Station Filter"
            ilRet = ogReport.SetFormula("IncludePersonnel", sgCrystlFormula1)
            ilRet = ogReport.SetFormula("IncludeVehicleForStation", sgCrystlFormula2)
            
        Case "AfWebImportLog.rpt"
            smReportCaption = "Web Import Log"
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula1)
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula2)
            ilRet = ogReport.SetFormula("FileDoesNotExist", sgCrystlFormula3)
        Case "AfLogDeliveryVh.rpt"
            smReportCaption = "Affiliate Log Delivery"
            ilRet = ogReport.SetFormula("SortBy", "'" & sgCrystlFormula1 & "'")
        Case "AfAudioDeliveryVh.rpt"
            smReportCaption = "Affiliate Audio Delivery"
            ilRet = ogReport.SetFormula("SortBy", "'" & sgCrystlFormula1 & "'")
'        Case "AfDelVendorDet.rpt", "AfDelDetail.rpt"                                              '5-24-18
        Case "AfDelDetail.rpt"                                              '5-24-18, chged to use one .rpt
            smReportCaption = "Affiliate Delivery Detail"
            ilRet = ogReport.SetFormula("Sort1", sgCrystlFormula1)
            ilRet = ogReport.SetFormula("Sort2", sgCrystlFormula2)
            ilRet = ogReport.SetFormula("Sort3", sgCrystlFormula3)
            ilRet = ogReport.SetFormula("ActiveDates", sgCrystlFormula4)
            ilRet = ogReport.SetFormula("SkipPage", sgCrystlFormula5)
        Case "AfSpotMgmt.rpt"                                                  '3-9-12
            smReportCaption = "Affiliate Spot Management"
            ilRet = ogReport.SetFormula("UserTimes", "'" & sgCrystlFormula1 & "'")
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)
            ilRet = ogReport.SetFormula("NewPage", sgCrystlFormula4)
            ilRet = ogReport.SetFormula("StatusSelection", sgCrystlFormula5)
            
            '2/28/22 - JW - Fix TTP 10403 - Affiliate Spot MGMT report showing extra vehicles
            '2/28/22 - JW - Fix Date format: DATE(Year,Month,Day)
            sSelectionFormula = "{ast.astStatus} > 1 AND {afr.afrgenDate} = DATE(" & Year(sgGenDate) & "," & Month(sgGenDate) & "," & Day(sgGenDate) & ") AND {afr.afrGenTime} = " & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False)))))
            ilRet = ogReport.SetSelection(sSelectionFormula)
            
        Case "AfExpHistory.rpt"
            smReportCaption = "Affiliate Export History"
        Case "AfSportDeclare.rpt"
            smReportCaption = "Station Sports Declaration"
            ilRet = ogReport.SetFormula("SuppressDeclaration", sgCrystlFormula1)
            ilRet = ogReport.SetFormula("VisitorText", sgCrystlFormula4)
            ilRet = ogReport.SetFormula("HomeText", sgCrystlFormula5)
            
        Case "AfSportClear.rpt"                 '10-16-12
            smReportCaption = "Sports Clearance"
            ilRet = ogReport.SetFormula("AllDeclarationsOption", sgCrystlFormula1)
            ilRet = ogReport.SetFormula("DelinqDate", sgCrystlFormula2)
            ilRet = ogReport.SetFormula("DelinquentOrPotentialOption", sgCrystlFormula3)
        
        Case "AfRenewal.rpt"                    '11-8-12        Affiliate Renewal status
            smReportCaption = "Agreement Renewal Status"
            ilRet = ogReport.SetFormula("SortBy", sgCrystlFormula1)         'D = sort by date, V = sort by vehicle
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)      'User entered start/end date span
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)
        
        Case "AfAdvComply.rpt"
            smReportCaption = "Advertiser Compliance"
            ilRet = ogReport.SetFormula("SortBy", sgCrystlFormula1)     '3
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)     '6
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)     '7
            ilRet = ogReport.SetFormula("DiscrepOnly", sgCrystlFormula4)     '19
            ilRet = ogReport.SetFormula("UserTimes", "'" & sgCrystlFormula5 & "'")     '37
            ilRet = ogReport.SetFormula("NewPage", sgCrystlFormula6)
            ilRet = ogReport.SetFormula("ShowDiscrepCode", sgCrystlFormula7)
            ilRet = ogReport.SetFormula("SiteCompliantFlag", sgCrystlFormula8)
        Case "AfRadarClr.rpt"
            smReportCaption = "Radar Clearance"
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula1)
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula2)
            ilRet = ogReport.SetFormula("Excluded", "'" & sgCrystlFormula3 & "'")
        Case "AfMeasure.rpt", "AfMeasureChartIt.rpt"
            smReportCaption = "Affiliate Measurement"
            ilRet = ogReport.SetFormula("WeekOf", sgCrystlFormula1)
            ilRet = ogReport.SetFormula("MajorSort", sgCrystlFormula2)
            ilRet = ogReport.SetFormula("MinorSort", sgCrystlFormula3)
            ilRet = ogReport.SetFormula("ZtoA", "'" & sgCrystlFormula4 & "'")       'blank or A indicates Ascending field, Z = descending
            ilRet = ogReport.SetFormula("ShowDetail", "'" & sgCrystlFormula5 & "'")     ' D = show station detail; else S = summary for totals by sort option
            ilRet = ogReport.SetFormula("InclNetworkNC", "'" & sgCrystlFormula6 & "'")
            ilRet = ogReport.SetFormula("InclResponse", "'" & sgCrystlFormula7 & "'")
            ilRet = ogReport.SetFormula("Debug", "'" & sgCrystlFormula8 & "'")          'show additional fields (weeks aired & # spots posted for debugging)
            ilRet = ogReport.SetFormula("DefaultDate", sgCrystlFormula9)          'Default week generated in smt
            ilRet = ogReport.SetFormula("CountsOrPct", "'" & sgCrystlFormula10 & "'")        'C = show by counts, P = show by Pcts
            ilRet = ogReport.SetFormula("AiredOrYear", "'" & sgCrystlFormula11 & "'")        'A = used aired weeks, Y = use 52
            ilRet = ogReport.SetFormula("ClientName", "'" & sgCrystlFormula12 & "'")        'Client name from site so a sub report doesnt have to be used to retrieve it
            ilRet = ogReport.SetFormula("MajorPageSkip", "'" & sgCrystlFormula13 & "'")        'Skip to new page each major change
        Case "AfVehicleVisual.rpt"
            smReportCaption = "Vehicle Visual Summary"
            ilRet = ogReport.SetFormula("DatesRequested", "'" & sgCrystlFormula1 & "'") '2
            ilRet = ogReport.SetFormula("VGSortBy", sgCrystlFormula3)
        Case "AfWebVendor.rpt"
            smReportCaption = "Web Vendor Export/Import"
            ilRet = ogReport.SetFormula("DatesRequested", "'" & sgCrystlFormula2 & "'")
            ilRet = ogReport.SetFormula("ExportOrImport", "'" & sgCrystlFormula1 & "'")         ' E = export, I = import
            ilRet = ogReport.SetFormula("SortBy", "'" & sgCrystlFormula3 & "'")         'sort by Station, Vehicle or VEndor
         Case "AfStationPersonRpt.rpt"
            '8/28/2018 New report for Personnel Station     FYM
            smReportCaption = "Personnel Station"
            ilRet = ogReport.SetFormula("MissingPersonnelOnly", "'" & sgCrystlFormula1 & "'")
        Case "AfCluster.rpt"
            ilRet = ogReport.SetFormula("SortBy", "'" & sgCrystlFormula1 & "'")     '
            ilRet = ogReport.SetFormula("StartDate", sgCrystlFormula2)
            ilRet = ogReport.SetFormula("EndDate", sgCrystlFormula3)
         End Select
    Exit Sub
ErrHand:
        Screen.MousePointer = vbDefault
        gMsg = ""
        If (Err.Number <> 0) And (gMsg = "") Then
            gMsg = "A general error has occured in frmCrystal - mSetFormulas: "
            gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
        End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload frmCrystal
End Sub

Public Sub gActiveCrystalReports(iExportType As Integer, iRptDest As Integer, slRptName As String, sExpName As String, rstActive As ADODB.Recordset)
Dim ilRet As Integer

'    Set fNewForm.report = Appl.OpenReport(sgReportDirectory + slRptName)
'    fNewForm.report.Database.Tables(1).SetDataSource rstActive, 3
'Dan changes for crxi
    Set ogReport = New CReportHelper
    ilRet = ogReport.OpenAndSetDataSource(slRptName, rstActive)
    'Without this do event you get errors when chosing destination printer
     DoEvents

    'Set up the formulas for a given report
    '''''''''''''''''''''''FIX THIS''''''''''''''''''mSetFormulas fNewForm
    'Set up the formulas for a given report
   ' mSetFormulas fNewForm, slRptName   'Dan changed for crxi
    mSetFormulas slRptName

    'gMsgBox fNewForm.Report.Database.Tables(1).DllName      'check which database driver is used
    Screen.MousePointer = vbDefault
    DoEvents    'Dan added this since using in gCrystalReports
         'Handle the report destination - Display, Print, Export  Dan changed for new crxi 12-10-08
    If iRptDest = 0 Then                   'Display Option
       ' Set frmViewReport.Report = ogReport.CurrentReportObject
            Report.Caption = smReportCaption
            Report.Show vbModal
'            frmViewReport.Caption = smReportCaption
'            frmViewReport.Show vbModal     'Affiliate is currently not vbmodal, but I lost logo because of ogreport = nothing below.
    ElseIf iRptDest = 1 Then               'Print Option.  Because affiliate doesn't currently have multiple reports, optional set to true
        gUserActivityLog "S", sgReportListName & ": Printing"
        ogReport.PrintOut False, True
        gUserActivityLog "E", sgReportListName & ": Printing"
    Else                                    'Export to File Options.  Optional set to true as above
        gUserActivityLog "S", sgReportListName & ": Exporting"
        ilRet = ogReport.Export(sExpName, iExportType, True)
        gUserActivityLog "E", sgReportListName & ": Exporting"
    End If

 Set ogReport = Nothing

End Sub
'
'
'           mFindDDFTableName - find the matching table name from DDFs
'           so that a database location can be set
'
'           <input> full table name from crystal
'           <output> none
'           return - true if a table name has been found
' 12-10-08  Dan  not tested as not being used
'Private Function mFindDDFTableName(slTableName As String) As Integer
'Dim ilLoopOnFile As Integer
'Dim ilFound As Integer
'
'    'look for the full name in valid array of filenames--if it exists from the DDF file, then
'    'it isnt an alias table in the report.
'    ilFound = False
'    For ilLoopOnFile = LBound(tgDDFFileNames) To UBound(tgDDFFileNames) - 1
'        If Trim$(slTableName) = Trim$(tgDDFFileNames(ilLoopOnFile).sLongName) Then
'            ilFound = True
'            Exit For
'        End If
'    Next ilLoopOnFile
'        mFindDDFTableName = ilFound
'End Function
''
''
''       find the base table name (i.e. for alias file naming:  uie.mkd
''           has been aliased to uie_to.  By finding the base table name,
''           the location can be set for the table
''
''           <Input> slTableName - Full alias name
''           <output> ilIndexToTableLoc - index to the base table name
''           Return - true  if the base table name found
'' 12-10-08  Dan  not tested as not being used
'Private Function mFindAliasTableName(slTableName As String, ilIndexToTableLoc As Integer) As Integer
'Dim ilLoopOnFile As Integer
'Dim slTempStr As String
'Dim ilRet As Integer
'Dim ilFound As Integer
'
'    ilFound = False
'    'valid file name not found.  This must be an alias table defined.
'    'look for valid 1st 3or 4 character name, then pick up the associated full filename for the location definition
'    For ilLoopOnFile = LBound(tgDDFFileNames) To UBound(tgDDFFileNames) - 1
'        slTempStr = RTrim$(tgDDFFileNames(ilLoopOnFile).sShortName)
'        ilRet = InStr(1, slTableName, Trim$(slTempStr))
'        If ilRet > 0 Then
'            'fNewForm.Report.Database.Tables(ilLoop).Location = sgDatabaseName & "." & Trim$(tgDDFFileNames(ilLoopOnFile).sLongName)
'            ilFound = True
'            ilIndexToTableLoc = ilLoopOnFile
'            Exit For
'        End If
'    Next ilLoopOnFile
'    mFindAliasTableName = ilFound
'End Function
Private Sub mFindFormulaIndexNums()

    'mFindFormulaIndexNums(fNewForm)'Dan-took out parameter
    
    Dim FormulaField As CRAXDRT.FormulaFieldDefinition
    Dim FormulaFields As CRAXDRT.FormulaFieldDefinitions
    Dim ilLoop As Integer

    On Error GoTo ErrHand
    'changed to new way of accessing report
    Set FormulaFields = ogReport.CurrentReportObject.FormulaFields
    'Set FormulaFields = fNewForm.report.FormulaFields
    ilLoop = 1
    For Each FormulaField In FormulaFields
       With FormulaField
           'gMsgBox "Crystal Formula Name: " & .FormulaFieldName & "  Equals Index = " & ilLoop, vbOKOnly
           Debug.Print "Crystal Formula Name: " & .FormulaFieldName & "  Equals Index = " & ilLoop
       End With
       ilLoop = ilLoop + 1
    Next FormulaField
finish:
    Set FormulaField = Nothing
    Set FormulaFields = Nothing
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmCrystal - mFindFormulaIndexNums: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    GoTo finish
End Sub
Public Sub gSaveReportCtrlsSetting()
    Dim ilRet As Integer
    ReDim slBypassCtrls(0 To 0) As String
    
    ilRet = gCreateFormControlFile(fgReportForm, sgReportCtrlSaveName, slBypassCtrls())
End Sub

Public Sub gSetReportCtrlsSetting()
    Dim ilRet As Integer
    
    On Error Resume Next
    
    ilRet = gSetFormCtrls(fgReportForm, sgReportCtrlSaveName)

End Sub
