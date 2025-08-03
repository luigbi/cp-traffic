Attribute VB_Name = "MiscSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Misc.Bas on Wed 6/17/09 @ 12:56 PM **
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Misc.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Misc subs and functions
'   Modules not required in stand alone projects
Option Explicit
Option Compare Text


Public Sub gWriteBkgd()
    Dim ilFlip As Integer
    Dim slVersion As String
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    Dim slCountdown As String
    Dim ilCountdown As Integer

    'sgWallpaper = "Testing INI text"
    'igShowVersionNo = -1
    'igTestSystem = True
    
    MstPict.ForeColor = &H8000000F
    slVersion = "Version " & App.Major & "." & App.Minor
    
    MstPict.ZOrder 1
    If (igTestSystem) And (igShowVersionNo <> -1) Then
        'TTP 10281: Jobs screen refresh - in Test mode, keep Tile - make RED Text
        MstPict.ForeColor = vbRed
        ilFlip = 0
        MstPict.CurrentX = 0
        MstPict.CurrentY = MstPict.TextHeight("Tj")
        Do
            Do
                If MstPict.CurrentX + MstPict.TextWidth("Test System") > MstPict.Width Then
                    Exit Do
                End If
                MstPict.Print "Test System";
                MstPict.CurrentX = MstPict.CurrentX + MstPict.TextWidth("Test System") \ 3
            Loop
            If ilFlip = 0 Then
                MstPict.CurrentX = MstPict.TextWidth("Test System") \ 2
                ilFlip = 1
            Else
                MstPict.CurrentX = 0
                ilFlip = 0
            End If
            MstPict.CurrentY = MstPict.CurrentY + 3 * MstPict.TextHeight("Tj")
            If MstPict.CurrentY > MstPict.height Then
                Exit Do
            End If
        Loop
        Exit Sub
    End If
    If (igShowVersionNo = 1) Or (igShowVersionNo = 2) Or (igShowVersionNo = 0 And sgWallpaper = "") Then
        slVersion = sgCSVersion
        ilPos1 = InStr(1, slVersion, " B", vbTextCompare)
        ilPos2 = InStr(1, slVersion, " Created", vbTextCompare)
        If (ilPos1 > 0) And (ilPos2 > 0) Then
            slVersion = Left(slVersion, ilPos1 - 1) & Mid(slVersion, ilPos2 - 1)
        ElseIf (ilPos1 > 0) And (ilPos2 = 0) Then
            slVersion = Left(slVersion, ilPos1 - 1)
        End If
        If igShowVersionNo = 2 Then
            ilPos2 = InStr(1, slVersion, " Created", vbTextCompare)
            If ilPos2 > 0 Then
                slVersion = Left$(slVersion, ilPos2) & " Debug"
            End If
        End If
        ilFlip = 0
        'MstPict.CurrentX = 0
        'MstPict.CurrentY = MstPict.TextHeight("Tj")
        'TTP 10281: Jobs screen refresh - When production, use grey color text and show text Centered under menu - not tiled
        MstPict.FontSize = 16
        MstPict.ForeColor = &H8000000F
        MstPict.CurrentY = (MstPict.height / 2) + 4500 + MstPict.TextHeight("Tj")
        MstPict.CurrentX = (MstPict.Width - MstPict.TextWidth(slVersion)) / 2
        'Do
            'Do
            '    If MstPict.CurrentX + MstPict.TextWidth(slVersion) > MstPict.Width Then
            '        Exit Do
            '    End If
                MstPict.Print slVersion;
                MstPict.CurrentX = MstPict.CurrentX + MstPict.TextWidth(slVersion) \ 3
            'Loop
            'If ilFlip = 0 Then
            '    MstPict.CurrentX = MstPict.TextWidth(slVersion) \ 2
            '    ilFlip = 1
            'Else
            '    MstPict.CurrentX = 0
            '    ilFlip = 0
            'End If
            'MstPict.CurrentY = MstPict.CurrentY + 3 * MstPict.TextHeight("Tj")
            'If MstPict.CurrentY > MstPict.height Then
            '    Exit Do
            'End If
        'Loop
        Exit Sub
    End If
    If (sgWallpaper <> "") Then
        'TTP 10281: Jobs screen refresh - If using a INI wallpaper, then tile in blue, just less repetitions...
        MstPict.ForeColor = vbBlue
        If igShowVersionNo = -1 Then
            ilPos1 = InStr(1, sgWallpaper, "Shutdown:", vbTextCompare)
            If ilPos1 > 0 Then
                slCountdown = Mid$(sgWallpaper, ilPos1 + 9)
                If right$(slCountdown, 1) = "m" Then
                    slCountdown = Left$(slCountdown, Len(slCountdown) - 1)
                End If
                ilCountdown = Val(slCountdown)
                If ilCountdown > 5 Then
                    MstPict.ForeColor = vbYellow
                ElseIf ilCountdown > 2 Then
                    MstPict.ForeColor = ORANGE
                Else
                    MstPict.ForeColor = vbRed
                End If
            End If
        End If
        slVersion = Trim$(sgWallpaper)
        ilFlip = 0
        MstPict.CurrentX = 0
        MstPict.CurrentY = MstPict.TextHeight("Tj")
        Do
            Do
                If MstPict.CurrentX + MstPict.TextWidth(slVersion) > MstPict.Width Then
                    Exit Do
                End If
                MstPict.Print slVersion;
                MstPict.CurrentX = (MstPict.TextWidth(slVersion) / 2) + (MstPict.CurrentX + MstPict.TextWidth(slVersion) * 3)
            Loop
            If ilFlip = 0 Then
                MstPict.CurrentX = (MstPict.TextWidth(slVersion) / 2) + (MstPict.TextWidth(slVersion) * 3)
                ilFlip = 1
            Else
                MstPict.CurrentX = (MstPict.TextWidth(slVersion) / 2)
                ilFlip = 0
            End If
            MstPict.CurrentY = MstPict.CurrentY + 3 * (MstPict.TextHeight("Tj") * 4.5)
            If MstPict.CurrentY > MstPict.height Then
                Exit Do
            End If
        Loop
        Exit Sub
    End If
End Sub

