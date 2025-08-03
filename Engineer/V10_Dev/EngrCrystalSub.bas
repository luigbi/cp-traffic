Attribute VB_Name = "EngrCrystalSub"
'
' Release: 1.0
'
' Description:
'   This file contains the declarations

Option Explicit

'Crystal Vars
Public Appl As New CRAXDRT.Application

Public igRptIndex As Integer            'report selected
Public igRptSource As Integer           'vbModal if coming from Snapshot icon, else vbModeless

'Generic Storage areas for formulas passed to Crystal reports
Public sgCrystlFormula1 As String
Public sgCrystlFormula2 As String
Public sgCrystlFormula3 As String
Public sgCrystlFormula4 As String
Public sgCrystlFormula5 As String
Public sgCrystlFormula6 As String
Public sgCrystlFormula7 As String
Public sgCrystlFormula8 As String
Public sgCrystlFormula9 As String
Public sgCrystlFormula10 As String

