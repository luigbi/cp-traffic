VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Custom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'
'   Decrypt User fields
'
'   <input> EncryptName - encrypted name
'   <return> Decrypted name
'   '10-16-01
'
Public Function DecryptString(EncryptString As String) As String
    Dim slDecrypt As String
    Dim slName As String
    Dim ilLen As Integer
    Dim ilLoop As Integer

    slName = EncryptString
    ilLen = Len(EncryptString)
    slDecrypt = ""
    For ilLoop = 1 To ilLen Step 1
        slDecrypt = slDecrypt & Chr(Asc(slName) - 128)
        slName = Mid$(slName, 2)
    Next ilLoop
    DecryptString = slDecrypt
End Function

Public Function TimeIn24HrTenths(llInputTime As Long) As String
Dim llTenths As Long
Dim slTenths As String
Dim llTime As Long
Dim llMinute As Long
Dim llSecond As Long
Dim llHour As Long

    TimeIn24HrTenths = Trim$(Str$(llInputTime))
    llTenths = llInputTime Mod 10
    slTenths = Trim$(Str$(llTenths))
    llTime = llInputTime \ 10           'take off the 10ths
    llHour = llTime \ 3600              'get hour index
    llMinute = (llTime - (llHour * 3600)) \ 60    'get minutes
    llSecond = llTime - ((llHour * 3600) + (llMinute * 60)) 'get seconds
    TimeIn24HrTenths = Trim$(Str$(llHour)) & ":" & Trim$(Str$(llMinute)) & ":" & Trim$(Str$(llSecond)) & "." & slTenths
    
End Function
