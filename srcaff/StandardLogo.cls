VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLogoStandard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IConditionalLogo
Private WithEvents omSection As CRAXDRT.Section
Attribute omSection.VB_VarHelpID = -1
Private omOle As CRAXDRT.OLEObject
'9-22-11 not changing or setting logo.
'Private Sub Class_Terminate()
'Set omSection = Nothing
'Set omOle = Nothing
'End Sub

Private Sub IConditionalLogo_Condition(ReportName As String)
'Dim ilReportObject As Integer
'Dim olSection As CRAXDRT.Section
'Dim blFoundLogo As Boolean
'    For Each olSection In ogReport.Reports.Item(ReportName).Areas("PH").Sections
'        For ilReportObject = 1 To olSection.ReportObjects.Count
'            If olSection.ReportObjects(ilReportObject).Kind = CRAXDRT.crOleObject Then
'                Set omOle = olSection.ReportObjects(ilReportObject)
'                Set omSection = olSection
'                blFoundLogo = True
'                Exit For
'            End If
'        Next ilReportObject
'        If blFoundLogo Then
'            Exit For
'        End If
'    Next olSection
'    Set olSection = Nothing
End Sub
'Private Sub omSection_format(ByVal pFormattingInfo As Object)
'   On Error Resume Next
'    Set omOle.FormattedPicture = LoadPicture(ogReport.LogoPath)
'End Sub
