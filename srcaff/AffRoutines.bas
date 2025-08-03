Attribute VB_Name = "modAffRoutines"
'**************************************************************************
' Copyright: Counterpoint Software, Inc. 2002
' Created by: Doug Smith
' Date: August 2002
' Name: modCrystal
'**************************************************************************
Option Explicit

Public Sub gWriteBkgd()
    Dim slVersion As String
    Dim llWidth As Long
    Dim ilFlip As Integer
    Dim ilIndex As Integer
    Dim ilCycle As Integer
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim slCountdown As String
    Dim ilCountdown As Integer
    Dim llForeColor As Long
    Dim slDateTime1 As String
    Dim slAppName As String
    Dim ilRet As Integer
    Dim iTileStyle As Integer
    'igTestSystem = True
    'sgWallpaper = "Wallpaper INI Text"
    'igShowVersionNo = 0
    
    'Defaults
    iTileStyle = 2
    llForeColor = &H8000000F
    frmMstPict.ForeColor = &H8000000F
    slAppName = App.EXEName
    ilPos = InStr(1, slAppName, ".", 1)
    If ilPos > 0 Then
        slAppName = Left$(slAppName, ilPos - 1)
    End If
    slAppName = slAppName & ".exe"
    ilRet = 0
    slVersion = "Version " & App.Major & "." & App.Minor
    ilRet = gFileExist(sgExeDirectory & slAppName)
    If ilRet = 0 Then
        slDateTime1 = gFileDateTime(sgExeDirectory & slAppName)
        slVersion = slVersion & " created " & Format$(slDateTime1, "m/d/yy") & " at " & Format$(slDateTime1, "h:mm:ssAM/PM")
    End If

    frmMstPict.ZOrder 1
    If (igTestSystem) And (igShowVersionNo <> -1) Then
        'TTP 10281: Jobs screen refresh - in Test mode, keep Tile - make RED Text
        iTileStyle = 0
        frmMstPict.ForeColor = vbRed
        llForeColor = vbRed
        slVersion = "Test System"
    ElseIf (igShowVersionNo = 1) Or (igShowVersionNo = 2) Then
        'TTP 10281: Jobs screen refresh - When production, use grey color text
        iTileStyle = 2
        llForeColor = &H8000000F
        frmMstPict.ForeColor = &H8000000F
        slVersion = "Version " & App.Major & "." & App.Minor
        If igShowVersionNo = 2 Then
            slVersion = slVersion & " Debug"
        End If
        If Trim$(sgWallpaper) <> "" Then
            slVersion = slVersion & " " & sgWallpaper
        ElseIf igShowVersionNo = 1 Then
            slAppName = App.EXEName
            ilPos = InStr(1, slAppName, ".", 1)
            If ilPos > 0 Then
                slAppName = Left$(slAppName, ilPos - 1)
            End If
            slAppName = slAppName & ".exe"
            ilRet = 0
            'On Error GoTo gWriteBkgdErr:
            'slDateTime1 = FileDateTime(sgExeDirectory & slAppName)
            ilRet = gFileExist(sgExeDirectory & slAppName)
            If ilRet = 0 Then
                slDateTime1 = gFileDateTime(sgExeDirectory & slAppName)
                slVersion = slVersion & " created " & Format$(slDateTime1, "m/d/yy") & " at " & Format$(slDateTime1, "h:mm:ssAM/PM")
            End If
        End If
    ElseIf (Trim$(sgWallpaper) <> "") Then
        llForeColor = vbBlue
        iTileStyle = 1
        If igShowVersionNo = -1 Then
            ilPos = InStr(1, sgWallpaper, "Shutdown:", vbTextCompare)
            If ilPos > 0 Then
                slCountdown = Mid$(sgWallpaper, ilPos + 9)
                If right$(slCountdown, 1) = "m" Then
                    slCountdown = Left$(slCountdown, Len(slCountdown) - 1)
                End If
                ilCountdown = Val(slCountdown)
                If ilCountdown > 5 Then
                    llForeColor = vbYellow
                ElseIf ilCountdown > 2 Then
                    llForeColor = ORANGE
                Else
                    llForeColor = vbRed
                End If
            End If
        End If
        slVersion = sgWallpaper
    End If
    slVersion = slVersion & "          "
    llWidth = frmMstPict.TextWidth(slVersion)
    slVersion = Trim$(slVersion)
    ilCycle = frmMstPict.Width \ llWidth
    frmMstPict!lacMsg(0).Width = frmMstPict.Width
    frmMstPict!lacMsg(0).Top = frmMstPict.TextHeight(slVersion)
    ilFlip = 0
    ilIndex = 0
    Select Case iTileStyle
        Case 0: 'lots of tiles
            Do
                If ilFlip = 0 Then
                    slVersion = slVersion & "          "
                    ilFlip = 1
                Else
                    slVersion = "          " & slVersion
                    ilFlip = 0
                End If
                frmMstPict!lacMsg(ilIndex).Caption = ""
                frmMstPict!lacMsg(ilIndex).ForeColor = llForeColor
                For ilLoop = 1 To ilCycle Step 1
                    frmMstPict!lacMsg(ilIndex).Caption = frmMstPict!lacMsg(ilIndex).Caption & slVersion
                Next ilLoop
                frmMstPict!lacMsg(ilIndex).Visible = True
                slVersion = Trim$(slVersion)
                If frmMstPict!lacMsg(ilIndex).Top + (3 * frmMstPict.TextHeight(slVersion)) > frmMstPict.Height Then
                    Exit Do
                End If
                ilIndex = ilIndex + 1
                On Error Resume Next
                Load frmMstPict!lacMsg(ilIndex)
                frmMstPict!lacMsg(ilIndex).Top = frmMstPict!lacMsg(ilIndex - 1).Top + (3 * frmMstPict.TextHeight(slVersion))
            Loop
        Case 1: 'reduced tiles
            Do
                If ilFlip = 0 Then
                    slVersion = slVersion & "          " & "        "
                    ilFlip = 1
                Else
                    slVersion = "        " & "          " & slVersion
                    ilFlip = 0
                End If
                llWidth = frmMstPict.TextWidth(slVersion)
                ilCycle = frmMstPict.Width \ llWidth
                
                frmMstPict!lacMsg(ilIndex).Caption = ""
                frmMstPict!lacMsg(ilIndex).ForeColor = llForeColor
                For ilLoop = 1 To ilCycle Step 1
                    frmMstPict!lacMsg(ilIndex).Caption = frmMstPict!lacMsg(ilIndex).Caption & slVersion
                Next ilLoop
                frmMstPict!lacMsg(ilIndex).Visible = True
                slVersion = Trim$(slVersion)
                If frmMstPict!lacMsg(ilIndex).Top + 6 * frmMstPict.TextHeight(slVersion) > frmMstPict.Height Then
                    Exit Do
                End If
                ilIndex = ilIndex + 1
                On Error Resume Next
                Load frmMstPict!lacMsg(ilIndex)
                frmMstPict!lacMsg(ilIndex).Top = frmMstPict!lacMsg(ilIndex - 1).Top + 5.6 * frmMstPict.TextHeight(slVersion)
            Loop
        Case 2: 'Production view - centered below Menu
            On Error Resume Next
            Load frmMstPict!lacMsg(1)
            frmMstPict!lacMsg(1).Visible = True
            frmMstPict!lacMsg(1).Left = (frmMain.Width - frmMstPict.TextWidth(frmMstPict!lacMsg(ilIndex).Caption & slVersion)) / 2
            frmMstPict!lacMsg(1).Top = (frmMain.Height / 2) + 4500 + frmMstPict.TextHeight("Tj")
            frmMstPict!lacMsg(1).ForeColor = llForeColor
            frmMstPict!lacMsg(1).Caption = frmMstPict!lacMsg(ilIndex).Caption & slVersion
            
    End Select
    Exit Sub
'gWriteBkgdErr:
'    ilRet = Err.Number
'    Resume Next
End Sub

