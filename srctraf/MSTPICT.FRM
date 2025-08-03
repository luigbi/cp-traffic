VERSION 5.00
Begin VB.Form MstPict 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
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
   Begin VB.Image imcPict 
      Height          =   420
      Left            =   4695
      Top             =   3645
      Visible         =   0   'False
      Width           =   1020
   End
End
Attribute VB_Name = "MstPict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of MSTPICT.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Private Sub Form_GotFocus()
Dim i As Integer
i = 0
End Sub
Private Sub Form_Load()
    Dim ilRet As Integer

    'If igTestSystem Then
    If (igTestSystem) Or (igShowVersionNo = 1) Or (igShowVersionNo = 2) Or (Trim$(sgWallpaper) <> "") Then
        MstPict.BackColor = GRAY
        MstPict.Picture = LoadPicture()  'Set Picture to None
    Else
        If Not igRemoteUser Then
            On Error GoTo mRetryPicture:
            If igTaskLoc = 0 Then
                MstPict.Move 0, 0, Traffic!pbcMsgArea.Width, Screen.Height
            Else
                MstPict.Move 0, 0, Traffic!pbcMsgArea.Width, Traffic!pbcMsgArea.Top
            End If
            'MstPict.Picture = LoadPicture(sgLogoPath & "CSIBack.Bmp")
            If Screen.Height / Screen.TwipsPerPixelY <= 480 Then
                'MstPict.Picture = LoadPicture(sgLogoPath & "CSI640T.Bmp")
                'ilRet = 0
                If gFileExist(sgLogoPath & "CSI640T.jpg") = 0 Then
                    'MstPict.Picture = LoadPicture(sgLogoPath & "CSI640T.jpg")
                    mLoadPicture "CSI640T.jpg"
                ElseIf gFileExist(sgLogoPath & "CSI640T.gif") = 0 Then
                    'MstPict.Picture = LoadPicture(sgLogoPath & "CSI640T.gif")
                    mLoadPicture "CSI640T.gif"
                ElseIf gFileExist(sgLogoPath & "CSI640T.Bmp") = 0 Then
                    MstPict.Picture = LoadPicture(sgLogoPath & "CSI640T.Bmp")
                End If
            ElseIf Screen.Height / Screen.TwipsPerPixelY <= 600 Then
                'MstPict.Picture = LoadPicture(sgLogoPath & "CSI800T.Bmp")
                If gFileExist(sgLogoPath & "CSI800T.jpg") = 0 Then
                    'MstPict.Picture = LoadPicture(sgLogoPath & "CSI800T.jpg")
                    mLoadPicture "CSI800T.jpg"
                ElseIf gFileExist(sgLogoPath & "CSI800T.gif") = 0 Then
                    'MstPict.Picture = LoadPicture(sgLogoPath & "CSI800T.gif")
                    mLoadPicture "CSI800T.gif"
                ElseIf gFileExist(sgLogoPath & "CSI800T.Bmp") = 0 Then
                    MstPict.Picture = LoadPicture(sgLogoPath & "CSI800T.Bmp")
                End If
            ElseIf Screen.Height / Screen.TwipsPerPixelY <= 768 Then
                If gFileExist(sgLogoPath & "CSI1024T.jpg") = 0 Then
                    'MstPict.Picture = LoadPicture(sgLogoPath & "CSI1024T.jpg")
                    mLoadPicture "CSI1024T.jpg"
                ElseIf gFileExist(sgLogoPath & "CSI1024T.gif") = 0 Then
                    'MstPict.Picture = LoadPicture(sgLogoPath & "CSI1024T.gif")
                    mLoadPicture "CSI1024T.gif"
                ElseIf gFileExist(sgLogoPath & "CSI1024T.Bmp") = 0 Then
                    MstPict.Picture = LoadPicture(sgLogoPath & "CSI1024T.Bmp")
                End If
            Else
                If gFileExist(sgLogoPath & "CSI1280T.jpg") = 0 Then
                    'MstPict.Picture = LoadPicture(sgLogoPath & "CSI1280T.jpg")
                    mLoadPicture "CSI1280T.jpg"
                ElseIf gFileExist(sgLogoPath & "CSI1280T.gif") = 0 Then
                    'MstPict.Picture = LoadPicture(sgLogoPath & "CSI1280T.gif")
                    mLoadPicture "CSI1280T.gif"
                ElseIf gFileExist(sgLogoPath & "CSI1280T.Bmp") = 0 Then
                    MstPict.Picture = LoadPicture(sgLogoPath & "CSI1280T.Bmp")
                End If
            End If
        Else
            gFadeForm MstPict, False, False, True
        End If
    End If
    Exit Sub
mRetryPicture:
    ilRet = 1
    Resume Next
mNoPicture:
    gFadeForm MstPict, False, False, True
    On Error GoTo 0
    Resume Next
End Sub
Private Sub Form_Paint()
'    Dim ilFlip As Integer
'    Dim slVersion As String
'    Dim ilPos As Integer
'
'    If (igTestSystem) And (igShowVersionNo <> -1) Then
'        ilFlip = 0
'        MstPict.CurrentX = 0
'        MstPict.CurrentY = MstPict.TextHeight("Tj")
'        Do
'            Do
'                If MstPict.CurrentX + MstPict.TextWidth("Test System") > MstPict.Width Then
'                    Exit Do
'                End If
'                MstPict.Print "Test System";
'                MstPict.CurrentX = MstPict.CurrentX + MstPict.TextWidth("Test System") \ 3
'            Loop
'            If ilFlip = 0 Then
'                MstPict.CurrentX = MstPict.TextWidth("Test System") \ 2
'                ilFlip = 1
'            Else
'                MstPict.CurrentX = 0
'                ilFlip = 0
'            End If
'            MstPict.CurrentY = MstPict.CurrentY + 3 * MstPict.TextHeight("Tj")
'            If MstPict.CurrentY > MstPict.Height Then
'                Exit Do
'            End If
'        Loop
'        Exit Sub
'    End If
'    If (igShowVersionNo = 1) Or (igShowVersionNo = 2) Then
'        slVersion = sgCSVersion
'        If igShowVersionNo = 2 Then
'            ilPos = InStr(1, slVersion, " Created", vbTextCompare)
'            If ilPos > 0 Then
'                slVersion = Left$(slVersion, ilPos) & " Debug"
'            End If
'        End If
'        ilFlip = 0
'        MstPict.CurrentX = 0
'        MstPict.CurrentY = MstPict.TextHeight("Tj")
'        Do
'            Do
'                If MstPict.CurrentX + MstPict.TextWidth(slVersion) > MstPict.Width Then
'                    Exit Do
'                End If
'                MstPict.Print slVersion;
'                MstPict.CurrentX = MstPict.CurrentX + MstPict.TextWidth(slVersion) \ 3
'            Loop
'            If ilFlip = 0 Then
'                MstPict.CurrentX = MstPict.TextWidth(slVersion) \ 2
'                ilFlip = 1
'            Else
'                MstPict.CurrentX = 0
'                ilFlip = 0
'            End If
'            MstPict.CurrentY = MstPict.CurrentY + 3 * MstPict.TextHeight("Tj")
'            If MstPict.CurrentY > MstPict.Height Then
'                Exit Do
'            End If
'        Loop
'        Exit Sub
'    End If
'    If (sgWallpaper <> "") Then
'        slVersion = Trim$(sgWallpaper)
'        ilFlip = 0
'        MstPict.CurrentX = 0
'        MstPict.CurrentY = MstPict.TextHeight("Tj")
'        Do
'            Do
'                If MstPict.CurrentX + MstPict.TextWidth(slVersion) > MstPict.Width Then
'                    Exit Do
'                End If
'                MstPict.Print slVersion;
'                MstPict.CurrentX = MstPict.CurrentX + MstPict.TextWidth(slVersion) \ 3
'            Loop
'            If ilFlip = 0 Then
'                MstPict.CurrentX = MstPict.TextWidth(slVersion) \ 2
'                ilFlip = 1
'            Else
'                MstPict.CurrentX = 0
'                ilFlip = 0
'            End If
'            MstPict.CurrentY = MstPict.CurrentY + 3 * MstPict.TextHeight("Tj")
'            If MstPict.CurrentY > MstPict.Height Then
'                Exit Do
'            End If
'        Loop
'        Exit Sub
'    End If
    gWriteBkgd
End Sub

Private Sub mLoadPicture(slFileName As String)
    imcPict.Stretch = False
    imcPict.Picture = LoadPicture(sgLogoPath & slFileName)
    imcPict.Stretch = True
    imcPict.Move 0, 0, MstPict.Width, MstPict.Height
    imcPict.Visible = True
End Sub
