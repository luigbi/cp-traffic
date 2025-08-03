VERSION 5.00
Begin VB.Form frmMstPict 
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
   Begin VB.Label lacMsg 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   645
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   825
      Visible         =   0   'False
      Width           =   9105
   End
End
Attribute VB_Name = "frmMstPict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bFirstActive As Boolean

Private Sub Form_Activate()
    If bFirstActive = True Then Exit Sub
    If (igTestSystem) Or (igShowVersionNo = 1) Or (igShowVersionNo = 2) Or (Trim$(sgWallpaper) <> "") Then
        mMstPictSetMsg
    End If
    bFirstActive = True
End Sub

Private Sub Form_GotFocus()
Dim i As Integer
i = 0
End Sub
Private Sub Form_Load()
    Dim ilRet As Integer
    
    If (igTestSystem) Or (igShowVersionNo = 1) Or (igShowVersionNo = 2) Or (Trim$(sgWallpaper) <> "") Then
        frmMstPict.BackColor = &HC0C0C0
        frmMstPict.Picture = LoadPicture()  'Set Picture to None
        lacMsg(0).BackColor = &HC0C0C0
    Else
        On Error GoTo mRetryPicture:
        'MstPict.Picture = LoadPicture(sgLogoDirectory & "CSIBack.Bmp")
        If Screen.Height / Screen.TwipsPerPixelY <= 480 Then
            'MstPict.Picture = LoadPicture(sgLogoDirectory & "CSI640T.Bmp")
            If gFileExist(sgLogoDirectory & "CSI640A.jpg") Then
                frmMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI640A.jpg")
            ElseIf gFileExist(sgLogoDirectory & "CSI640A.gif") Then
                frmMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI640A.gif")
            ElseIf gFileExist(sgLogoDirectory & "CSI640A.Bmp") Then
                frmMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI640A.Bmp")
            End If
        ElseIf Screen.Height / Screen.TwipsPerPixelY <= 600 Then
            'frmMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI800T.Bmp")
            If gFileExist(sgLogoDirectory & "CSI800A.jpg") Then
                frmMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI800A.jpg")
            ElseIf gFileExist(sgLogoDirectory & "CSI800A.gif") Then
                frmMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI800A.gif")
            ElseIf gFileExist(sgLogoDirectory & "CSI800A.Bmp") Then
                frmMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI800A.Bmp")
            End If
        ElseIf Screen.Height / Screen.TwipsPerPixelY <= 768 Then
            'frmMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1024T.Bmp")
            If gFileExist(sgLogoDirectory & "CSI1024A.jpg") Then
                frmMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1024A.jpg")
            ElseIf gFileExist(sgLogoDirectory & "CSI1024A.gif") Then
                frmMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1024A.gif")
            ElseIf gFileExist(sgLogoDirectory & "CSI1024A.Bmp") Then
                frmMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1024A.Bmp")
            End If
        Else
            'frmMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1024T.Bmp")
            If gFileExist(sgLogoDirectory & "CSI1280A.jpg") Then
                frmMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1280A.jpg")
            ElseIf gFileExist(sgLogoDirectory & "CSI1280A.gif") Then
                frmMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1280A.gif")
            ElseIf gFileExist(sgLogoDirectory & "CSI1280A.Bmp") Then
                frmMstPict.Picture = LoadPicture(sgLogoDirectory & "CSI1280A.Bmp")
            End If
        End If
    End If
'Setting frmMain.BackColor = &HC0C0C0
'    If (igShowVersionNo = 1) Or (igShowVersionNo = 2) Then
'        If Screen.Height / Screen.TwipsPerPixelY <= 480 Then
'            frmMstPict.Move 0, 0, Screen.Width / 1.05, Screen.Height / 1.2
'        ElseIf Screen.Height / Screen.TwipsPerPixelY <= 600 Then
'            frmMstPict.Move 0, 0, Screen.Width / 1.016, Screen.Height / 1.16
'        ElseIf Screen.Height / Screen.TwipsPerPixelY <= 768 Then
'            frmMstPict.Move 0, 0, Screen.Width / 1.008, Screen.Height / 1.12
'        Else
'            frmMstPict.Move 0, 0, Screen.Width / 1.004, Screen.Height / 1.086
'        End If
'    Else
        frmMstPict.Move 0, 0, Screen.Width - 120, Screen.Height - 2000
'    End If
'    If (igTestSystem) Or (igShowVersionNo = 1) Or (igShowVersionNo = 2) Or (Trim$(sgWallpaper) <> "") Then
'        mMstPictSetMsg
'    End If
    Exit Sub
mRetryPicture:
    ilRet = 1
    Resume Next
mNoPicture:
    'gFadeForm frmMstPict, False, False, True
    On Error GoTo 0
    Resume Next
End Sub

Public Sub mMstPictSetMsg()
'    Dim slVersion As String
'    Dim llWidth As Long
'    Dim ilFlip As Integer
'    Dim ilIndex As Integer
'    Dim ilCycle As Integer
'    Dim ilLoop As Integer
'
'    If igTestSystem Then
'        slVersion = "Test System"
'    ElseIf (igShowVersionNo = 1) Or (igShowVersionNo = 2) Then
'        slVersion = "Version " & App.Major & "." & App.Minor
'        If igShowVersionNo = 2 Then
'            slVersion = slVersion & " Debug"
'        End If
'        If Trim$(sgWallpaper) <> "" Then
'            slVersion = slVersion & " " & sgWallpaper
'        End If
'    ElseIf (Trim$(sgWallpaper) <> "") Then
'        slVersion = sgWallpaper
'    End If
'    slVersion = slVersion & "          "
'    llWidth = frmMstPict.TextWidth(slVersion)
'    slVersion = Trim$(slVersion)
'    ilCycle = frmMstPict.Width \ llWidth
'    lacMsg(0).Width = frmMstPict.Width
'    lacMsg(0).Top = frmMstPict.TextHeight(slVersion)
'    ilFlip = 0
'    ilIndex = 0
'    Do
'        If ilFlip = 0 Then
'            slVersion = slVersion & "          "
'            ilFlip = 1
'        Else
'            slVersion = "          " & slVersion
'            ilFlip = 0
'        End If
'        lacMsg(ilIndex).Caption = ""
'        For ilLoop = 1 To ilCycle Step 1
'            lacMsg(ilIndex).Caption = lacMsg(ilIndex).Caption & slVersion
'        Next ilLoop
'        lacMsg(ilIndex).Visible = True
'        slVersion = Trim$(slVersion)
'        If lacMsg(ilIndex).Top + 4 * frmMstPict.TextHeight(slVersion) > frmMstPict.Height Then
'            Exit Do
'        End If
'        ilIndex = ilIndex + 1
'        Load lacMsg(ilIndex)
'        lacMsg(ilIndex).Top = lacMsg(ilIndex - 1).Top + 3 * frmMstPict.TextHeight(slVersion)
'    Loop
    gWriteBkgd
End Sub
