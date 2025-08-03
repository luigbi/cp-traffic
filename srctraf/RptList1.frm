VERSION 5.00
Begin VB.Form RptList 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   210
   ClientLeft      =   2505
   ClientTop       =   2685
   ClientWidth     =   225
   ControlBox      =   0   'False
   Enabled         =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   210
   ScaleWidth      =   225
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "RptList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptList1.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract revision number increment screen code
Option Explicit
Option Compare Text


Private Sub Form_Load()
    Dim ilRet As Integer
    Dim ilShell As Integer
    Dim slCommandStr As String
    Dim ilPos As String
    Dim slDate As String
    Dim ilLoop As Integer
    Dim oMyFileObj As FileSystemObject
    Dim MyFile As TextStream
    Dim slName As String
    Dim slFullPath As String
    Dim dlShellRet As Double
    
    RptList.Left = -1000
    RptList.Top = -1000
    slCommandStr = sgCommandStr
    If ((Len(Trim$(sgSpecialPassword)) = 4) And (Val(sgSpecialPassword) >= 1) And (Val(sgSpecialPassword) < 10000)) Then
        ilPos = InStr(1, slCommandStr, "Guide", vbTextCompare)
        If ilPos > 0 Then
            slCommandStr = Left(slCommandStr, ilPos - 1) & "CSI" & Mid(slCommandStr, ilPos + 5)
        End If
    End If
    'Dan M 9/20/10 problems in v57 reports.exe running GetCsiName
    'slDate = Trim$(gGetCSIName("SYSDate"))
    slDate = gCSIGetName()
    If slDate <> "" Then
        'use slDate when writing to file later
        slDate = " /D:" & slDate
        'slCommandStr = slCommandStr & " /D:" & slDate
        slCommandStr = slCommandStr & slDate
    End If
    slCommandStr = slCommandStr & " /ULF:" & lgUlfCode
    On Error GoTo LoadErr
    ilRet = 0
    AppActivate "CSI Reports"
    If ilRet = 1 Then
        'ilShell = Shell(sgExePath & "Reports.Exe " & slCommandStr, vbNormalFocus)
        dlShellRet = Shell(sgExePath & "Reports.Exe " & slCommandStr, vbNormalFocus)
    'create file of date only if there is a date.
    ElseIf LenB(slDate) > 0 Then
'        SendKeys "{~}", True
'        For ilLoop = 1 To Len(slCommandStr) Step 1
'            SendKeys Mid(slCommandStr, ilLoop, 1)
'        Next ilLoop
'        SendKeys "#", True
'      Dan M 9/21/10 replace with writing to text file
        '5676 c not hardcoded
        'slFullPath = "C:\csi\"
        slFullPath = sgRootDrive & "csi\"
        Set oMyFileObj = New FileSystemObject
        If oMyFileObj.FolderExists(slFullPath) Then
            '8903
           ' slFullPath = slFullPath & "ReportPasser.txt"
            slFullPath = slFullPath & "ReportPasser-" & tgUrf(0).iCode & ".txt"
            ilRet = 0
            On Error GoTo LoadErr
            Set MyFile = oMyFileObj.OpenTextFile(slFullPath, ForWriting, True)
            If ilRet = 0 Then
                MyFile.WriteLine slDate
                MyFile.Close
            End If
            Set MyFile = Nothing
        End If
        Set oMyFileObj = Nothing
    End If
    tmcTerminate.Enabled = True
    Exit Sub
LoadErr:
    ilRet = 1
    Resume Next
    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RptList = Nothing
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload RptList
End Sub
