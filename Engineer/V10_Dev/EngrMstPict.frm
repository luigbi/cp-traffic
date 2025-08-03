VERSION 5.00
Begin VB.Form EngrMstPict 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   ClientHeight    =   6585
   ClientLeft      =   0
   ClientTop       =   1035
   ClientWidth     =   9435
   Enabled         =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6585
   ScaleWidth      =   9435
   Visible         =   0   'False
End
Attribute VB_Name = "EngrMstPict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_GotFocus()
Dim i As Integer
i = 0
End Sub
Private Sub Form_Load()
    Dim ilRet As Integer
    
    If igTestSystem Then
        EngrMstPict.BackColor = &HC0C0C0
        EngrMstPict.Picture = LoadPicture()  'Set Picture to None
    Else
        On Error GoTo mRetryPicture:
        'MstPict.Picture = LoadPicture(sgLogoDirectory & "CSIBack.Bmp")
        If Screen.Height / Screen.TwipsPerPixelY <= 480 Then
            'MstPict.Picture = LoadPicture(sgLogoDirectory & "CSI640T.Bmp")
            ilRet = 0
            EngrMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI640e.jpg")
            If ilRet <> 0 Then
                ilRet = 0
                EngrMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI640e.gif")
                If ilRet <> 0 Then
                    On Error GoTo mNoPicture:
                    ilRet = 0
                    EngrMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI640e.bmp")
                End If
            End If
        ElseIf Screen.Height / Screen.TwipsPerPixelY <= 600 Then
            'EngrMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI800T.Bmp")
            ilRet = 0
            EngrMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI800e.jpg")
            If ilRet <> 0 Then
                ilRet = 0
                EngrMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI800e.gif")
                If ilRet <> 0 Then
                    On Error GoTo mNoPicture:
                    ilRet = 0
                    EngrMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI800e.bmp")
                End If
            End If
        ElseIf Screen.Height / Screen.TwipsPerPixelY <= 768 Then
            'EngrMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1024T.Bmp")
            ilRet = 0
            EngrMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1024e.jpg")
            If ilRet <> 0 Then
                ilRet = 0
                EngrMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1024e.gif")
                If ilRet <> 0 Then
                    On Error GoTo mNoPicture:
                    ilRet = 0
                    EngrMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1024e.bmp")
                End If
            End If
        Else
            'EngrMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1024T.Bmp")
            ilRet = 0
            EngrMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1280e.jpg")
            If ilRet <> 0 Then
                ilRet = 0
                EngrMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1280e.gif")
                If ilRet <> 0 Then
                    On Error GoTo mNoPicture:
                    ilRet = 0
                    EngrMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1280e.bmp")
                End If
            End If
        End If
    End If
    EngrMstPict.Move 0, 0, Screen.Width / 1.05, Screen.Height / 1.2
    Exit Sub
mRetryPicture:
    ilRet = 1
    Resume Next
mNoPicture:
    'gFadeForm EngrMstPict, False, False, True
    On Error GoTo 0
    Resume Next
End Sub
Private Sub Form_Paint()
    Dim ilFlip As Integer
    Dim ilX As Integer
    
    ilX = EngrMstPict.TextWidth("Test System") \ 4
    If igTestSystem Then
        ilFlip = 0
        EngrMstPict.CurrentX = ilX
        EngrMstPict.CurrentY = EngrMstPict.TextHeight("Tj")
        Do
            Do
                If EngrMstPict.CurrentX + EngrMstPict.TextWidth("Test System") > EngrMstPict.Width Then
                    Exit Do
                End If
                EngrMstPict.Print "Test System";
                EngrMstPict.CurrentX = EngrMstPict.CurrentX + EngrMstPict.TextWidth("Test System") \ 3
            Loop
            If ilFlip = 0 Then
                EngrMstPict.CurrentX = ilX + EngrMstPict.TextWidth("Test System") \ 2
                ilFlip = 1
            Else
                EngrMstPict.CurrentX = ilX
                ilFlip = 0
            End If
            EngrMstPict.CurrentY = EngrMstPict.CurrentY + 3 * EngrMstPict.TextHeight("Tj")
            If EngrMstPict.CurrentY > EngrMstPict.Height Then
                Exit Do
            End If
        Loop
    End If
End Sub
