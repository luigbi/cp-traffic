Attribute VB_Name = "ModContact"
'******************************************************
'*  modContact - various global declarations
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit


Type CONTACTINFO
    sStation As String * 40
    sVehicle As String * 42
    sACName As String * 40
    sACPhone As String * 20
    iShttCode As Integer
    iVefCode As Integer
    iCDateIndex As Integer
End Type

Type CONTACTDATE
    iCol As Integer
    sDate As String * 10
    iPostingStatus As Integer
End Type

