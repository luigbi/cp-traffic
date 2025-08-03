Attribute VB_Name = "EngrRoutines"
'
' Release: 1.0
'
' Description:
'   This file contains the General declarations
Option Explicit


Public Sub gSetFonts(Frm As Form)
    Dim Ctrl As Control
    Dim ilFontSize As Integer
    Dim ilColorFontSize As Integer
    Dim ilBold As Integer
    Dim ilChg As Integer
    Dim slStr As String
    Dim slFontName As String
    Dim llHeight As Long
    
    
    On Error Resume Next
    ilFontSize = 10 '12
    ilBold = True
    ilColorFontSize = 10
    slFontName = "Arial"
    If Screen.Height / Screen.TwipsPerPixelY <= 480 Then
        ilFontSize = 7  '8
        ilBold = False
        ilColorFontSize = 7 '8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 600 Then
        ilFontSize = 7  '8
        ilBold = False
        ilColorFontSize = 7 '8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 768 Then
        ilFontSize = 9  '10
        ilBold = False
        ilColorFontSize = 7 '8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 800 Then
        ilFontSize = 9  '10
        ilBold = True
        ilColorFontSize = 7 '8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 1024 Then
        ilFontSize = 10 '12
        ilBold = True
    End If
    ilBold = False
    For Each Ctrl In Frm.Controls
        If TypeOf Ctrl Is MSHFlexGrid Then
            Ctrl.Font.Name = slFontName
            Ctrl.FontFixed.Name = slFontName
            Ctrl.Font.Size = ilFontSize
            Ctrl.FontFixed.Size = ilFontSize
            Ctrl.Font.Bold = ilBold
            Ctrl.FontFixed.Bold = ilBold
        ElseIf TypeOf Ctrl Is TabStrip Then
            Ctrl.Font.Name = slFontName
            Ctrl.Font.Size = ilFontSize
            Ctrl.Font.Bold = ilBold
        ElseIf TypeOf Ctrl Is ListView Then
            Ctrl.Font.Name = slFontName
            Ctrl.Font.Size = ilFontSize
            Ctrl.Font.Bold = ilBold
        Else
            ilChg = 0
            If (Ctrl.ForeColor = vbBlack) Or (Ctrl.ForeColor = &H80000008) Or (Ctrl.ForeColor = &H80000012) Or (Ctrl.ForeColor = &H8000000F) Then
                ilChg = 1
            Else
                ilChg = 2
            End If
            slStr = Ctrl.Name
            If (InStr(1, slStr, "Arrow", vbTextCompare) > 0) Or ((InStr(1, slStr, "Dropdown", vbTextCompare) > 0) And (TypeOf Ctrl Is CommandButton)) Then
                ilChg = 0
            End If
            If (InStr(1, slStr, "Search", vbTextCompare) > 0) Then
                ilChg = 2
            End If
            If ilChg = 1 Then
                Ctrl.FontName = slFontName
                Ctrl.FontSize = ilFontSize
                Ctrl.FontBold = ilBold
            ElseIf ilChg = 2 Then
                Ctrl.FontName = slFontName
                Ctrl.FontSize = ilColorFontSize
                Ctrl.FontBold = False
            End If
        End If
    Next Ctrl
    llHeight = -1
    For Each Ctrl In Frm.Controls
        If StrComp(slStr, "edcSearch", vbTextCompare) = 0 Then
            Ctrl.Height = 90
            llHeight = Ctrl.Height
            Exit For
        End If
    Next Ctrl
    If llHeight <> -1 Then
        For Each Ctrl In Frm.Controls
            If StrComp(slStr, "cmcSearch", vbTextCompare) = 0 Then
                Ctrl.Height = llHeight
                Exit For
            End If
        Next Ctrl
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gCenterForm                     *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Center form within Traffic Form *
'*                                                     *
'*******************************************************
Sub gCenterFormModal(FrmName As Form)
'
'   gCenterForm FrmName
'   Where:
'       FrmName (I)- Name of modeless form to be centered within Traffic form
'
    Dim flLeft As Single
    Dim flTop As Single
    flLeft = EngrMain.Left + (EngrMain.Width - EngrMain.ScaleWidth) / 2 + (EngrMain.ScaleWidth - FrmName.Width) / 2
    flTop = EngrMain.Top + (EngrMain.Height - FrmName.Height) / 2 + 510
    FrmName.Move flLeft, flTop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gCenterForm                     *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Center form within Traffic Form *
'*                                                     *
'*******************************************************
Sub gCenterForm(FrmName As Form)
'
'   gCenterForm FrmName
'   Where:
'       FrmName (I)- Name of modeless form to be centered within Traffic form
'
    Dim flLeft As Single
    Dim flTop As Single
    flLeft = EngrMain.Left + (EngrMain.Width - EngrMain.ScaleWidth) / 2 + (EngrMain.ScaleWidth - FrmName.Width) / 2
    flTop = EngrMain.Top + (EngrMain.Height - FrmName.Height) / 2 - 510
    FrmName.Move flLeft, flTop
End Sub
Public Sub gGetSchDates()
    Dim slEDate As String
    Dim slLDate As String
    slEDate = gGetEarlestSchdDate(False)
    If slEDate <> "" Then
        slLDate = gGetLatestSchdDate(False)
        EngrMain!imcTask(SCHEDULEJOB).Caption = "SCHEDULE: " & slEDate & "-" & slLDate
    Else
        EngrMain!imcTask(SCHEDULEJOB).Caption = "SCHEDULE: None"
    End If
End Sub

