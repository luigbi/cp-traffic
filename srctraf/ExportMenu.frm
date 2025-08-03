VERSION 5.00
Begin VB.Form ExportMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Menu"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4410
   Begin VB.Timer tmcExport 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1920
      Top             =   1665
   End
End
Attribute VB_Name = "ExportMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mInit()
    Me.Left = -Me.Width - 100
End Sub



Private Sub Form_Load()
    mInit
    tmcExport.Enabled = True
End Sub

Private Sub tmcExport_Timer()
    tmcExport.Enabled = False
    mExport
    igManUnload = YES
    Unload ExportMenu
    Set ExportMenu = Nothing   'Remove data segment
    igManUnload = NO
    Screen.MousePointer = vbDefault
End Sub

Private Sub mExport()
    Dim ilShell As Integer
    Dim slCommandStr As String
    Dim ilPos As Integer
    Dim slDate As String
    Dim blStart As Boolean
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    
    If igTestSystem Then
        slCommandStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    Else
        slCommandStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    End If
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
    slCommandStr = slCommandStr & "/Export:" & sgExportMenuItem & "~"
'    ilShell = Shell(sgExePath & "Exports.Exe " & slCommandStr, vbNormalFocus)
'    DoEvents
'    blStart = True
'    Do
'        On Error GoTo LoadErr
'        ilRet = 0
'        AppActivate "CSI Exports"
'        If ilRet = 1 Then
'            If blStart Then
'                Screen.MousePointer = vbDefault
'                blStart = False
'            Else
'                Exit Sub
'            End If
'        Else
'            blStart = False
'            Screen.MousePointer = vbDefault
'        End If
'        Sleep 1000
'    Loop
    Screen.MousePointer = vbDefault
    Traffic.WindowState = vbMinimized
    gShellAndWait ExportMenu, sgExePath & "Exports.exe " & slCommandStr, vbNormalFocus, False   'vbFalse
    
    sgVpfStamp = "~"
    ilRet = gVpfRead()          '3-26-13 force table to be refreshed
    Traffic.WindowState = vbMaximized

    Exit Sub
LoadErr:
    ilRet = 1
    Resume Next
End Sub
