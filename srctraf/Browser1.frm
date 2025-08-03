VERSION 2.00
Begin Form Browser 
   BackColor       =   &H8000000F&
   BorderStyle     =   3  'Fixed Double
   ClientHeight    =   5175
   ClientLeft      =   1155
   ClientTop       =   1800
   ClientWidth     =   7125
   ControlBox      =   0   'False
   FontBold        =   -1  'True
   FontItalic      =   0   'False
   FontName        =   "Arial"
   FontSize        =   8.25
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   5580
   Left            =   1095
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7125
   Top             =   1455
   Width           =   7245
   Begin PictureBox plcBrowser 
      Alignment       =   0  'Left Justify - TOP
      BevelWidth      =   2
      Height          =   4380
      Left            =   210
      TabIndex        =   4
      Top             =   285
      Width           =   6690
      Begin FileListBox lbcBrowserFile 
         Archive         =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Case            =   0  'Unchanged
         DividerStyle    =   0  'None
         FileTypePictures=   3  'Show for all items
         FixedHeight     =   17
         Font3D          =   0  'None
         Height          =   1560
         Hidden          =   0   'False
         Index           =   1
         IntegralSize    =   -1  'True
         Left            =   420
         ListStyle       =   1  '3D (BackColor ignored)
         MultiColumn     =   0   'False
         Normal          =   -1  'True
         Pattern         =   "*.bmp"
         ReadOnly        =   -1  'True
         ScrollHorizontal=   0   'False
         ScrollVertical  =   -1  'True
         SelectionType   =   0  'Single
         ShadowColor     =   0  'Dark Grey
         System          =   0   'False
         TabIndex        =   10
         Top             =   360
         Width           =   2505
         WndStyle        =   1151336787
      End
      Begin ListBox lbcBrowser 
         Prop47          =   BROWSER.FRX:0000
         BorderStyle     =   1  'Fixed Single
         Case            =   0  'Unchanged
         DividerStyle    =   3  'Raised
         FixedHeight     =   14
         Font3D          =   0  'None
         Height          =   1710
         IntegralSize    =   -1  'True
         Left            =   300
         ListStyle       =   0  '2D (BackColor used)
         MultiColumn     =   0   'False
         ReFreshOnUpdate =   -1  'True
         ScrollHorizontal=   0   'False
         ScrollVertical  =   -1  'True
         SelectionType   =   0  'Single
         ShadowColor     =   0  'Dark Grey
         Sorted          =   0   'False
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2850
         Width           =   6330
         WndStyle        =   1151336785
      End
      Begin Grid gdcBrowser 
         Height          =   1710
         HighLight       =   0   'False
         Left            =   285
         Rows            =   19
         TabIndex        =   15
         Top             =   2865
         Width           =   6345
      End
      Begin DirListBox lbcBrowserPath 
         BackColor       =   &H00FFFFFF&
         Height          =   1605
         Left            =   3465
         TabIndex        =   12
         Top             =   720
         Width           =   2565
      End
      Begin DriveListBox cbcBrowserDrive 
         Height          =   315
         Left            =   3465
         TabIndex        =   11
         Top             =   360
         Width           =   2565
      End
      Begin FileListBox lbcBrowserFile 
         Archive         =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Case            =   0  'Unchanged
         DividerStyle    =   0  'None
         FileTypePictures=   3  'Show for all items
         FixedHeight     =   17
         Font3D          =   0  'None
         Height          =   1560
         Hidden          =   0   'False
         Index           =   0
         IntegralSize    =   -1  'True
         Left            =   420
         ListStyle       =   1  '3D (BackColor ignored)
         MultiColumn     =   0   'False
         Normal          =   -1  'True
         Pattern         =   "*.bmp"
         ReadOnly        =   -1  'True
         ScrollHorizontal=   0   'False
         ScrollVertical  =   -1  'True
         SelectionType   =   0  'Single
         ShadowColor     =   0  'Dark Grey
         System          =   0   'False
         TabIndex        =   9
         Top             =   720
         Width           =   2505
         WndStyle        =   1151336787
      End
      Begin TextBox edcBrowserFile 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   420
         TabIndex        =   8
         Top             =   360
         Width           =   2505
      End
      Begin PictureBox pbcPicBrowser 
         Height          =   1725
         Index           =   0
         Left            =   195
         ScaleHeight     =   1695
         ScaleWidth      =   6075
         TabIndex        =   6
         Top             =   2565
         Width           =   6105
         Begin PictureBox pbcPicBrowser 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1515
            Index           =   1
            Left            =   15
            ScaleHeight     =   1515
            ScaleWidth      =   5940
            TabIndex        =   7
            Top             =   0
            Width           =   5940
         End
      End
      Begin VScrollBar vbcPicBrowser 
         Height          =   1545
         Left            =   6315
         TabIndex        =   5
         Top             =   2550
         Width           =   240
      End
      Begin Label lacFileName 
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   195
         TabIndex        =   17
         Top             =   2310
         Width           =   6105
      End
      Begin Label lacBrowserPath 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Caption         =   "File Path"
         Height          =   210
         Left            =   3480
         TabIndex        =   13
         Top             =   120
         Width           =   2550
      End
      Begin Label lacBrowserFile 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Caption         =   "File Name"
         Height          =   210
         Left            =   420
         TabIndex        =   14
         Top             =   120
         Width           =   2490
      End
   End
   Begin CommandButton cmcCancel 
      Caption         =   "&Cancel"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3825
      TabIndex        =   3
      Top             =   4875
      Width           =   945
   End
   Begin PictureBox pbcClickFocus 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   1770
      Width           =   75
   End
   Begin PictureBox plcScreen 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   0  'None
      BevelWidth      =   3
      BorderWidth     =   2
      BorderStyle      = 0
      Caption         =   "Browser"
      Font3D          =   4  'Inset w/heavy shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      BackColor        = &H8000000F&
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ShadowColor     =   1  'Black
      TabIndex        =   0
      Top             =   -15
      Width           =   885
   End
   Begin CommandButton cmcDone 
      Caption         =   "&Ok"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2460
      TabIndex        =   1
      Top             =   4875
      Width           =   945
   End
   Begin Image imcHelp 
      Height          =   345
      Left            =   345
      Top             =   4755
      Width           =   360
   End
End
' Copyright 1993 Counterpoint Software®, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Browser.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Browser input screen code
'
'   Shift8 added if multi-selection required to igBrowserType
'
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim hmFrom As Integer
Dim smFieldValues() As String
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imBrowserIndex As Integer
Sub cbcBrowserDrive_Change ()
    Screen.MousePointer = vbHourGlass
    lbcBrowserPath.Path = cbcBrowserDrive.Drive
    lbcBrowserFile(imBrowserIndex).Clear
    Screen.MousePointer = vbDefault
End Sub
Sub cmcCancel_Click ()
    igBrowserReturn = 0
    mTerminate
End Sub
Sub cmcDone_Click ()
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slName As String
    Dim ilPos As Integer
    Dim ilLoop As Integer
    igBrowserReturn = 1
    slName = Trim$(edcBrowserFile.Text)
    If Len(slName) <= 0 Then
        Beep
        edcBrowserFile.SetFocus
        Exit Sub
    End If
    ilPos = InStr(slName, "*")
    If ilPos > 0 Then
        Beep
        lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\" & slName
        If imBrowserIndex = 0 Then
            edcBrowserFile.SetFocus
        Else
            lbcBrowserFile(imBrowserIndex).SetFocus
        End If
        Exit Sub
    End If
    ilPos = InStr(slName, "?")
    If ilPos > 0 Then
        Beep
        lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\" & slName
        If imBrowserIndex = 0 Then
            edcBrowserFile.SetFocus
        Else
            lbcBrowserFile(imBrowserIndex).SetFocus
        End If
        Exit Sub
    End If
    If imBrowserIndex = 0 Then
        sgBrowserFile = slName
        If InStr(sgBrowserFile, ":") = 0 Then
            sgBrowserFile = lbcBrowserPath.Path & "\" & sgBrowserFile
        End If
    Else
        sgBrowserFile = ""
        For ilLoop = 0 To lbcBrowserFile(1).ListCount - 1 Step 1
            If lbcBrowserFile(1).Selected(ilLoop) Then
                sgBrowserFile = sgBrowserFile & lbcBrowserPath.Path & "\" & lbcBrowserFile(1).List(ilLoop) & "|"
            End If
        Next ilLoop
    End If
    mTerminate
End Sub
Sub cmcDone_GotFocus ()
    gCtrlGotFocus ActiveControl
End Sub
Sub edcBrowserFile_Change ()
    Dim ilPos As Integer
    Dim slName As String
    Dim slFromFile As String
    Dim ilRet As Integer
    slName = Trim$(edcBrowserFile.Text)
    If Len(slName) <= 0 Then
        If imBrowserIndex = 0 Then
            edcBrowserFile.SetFocus
        Else
            lbcBrowserFile(imBrowserIndex).SetFocus
        End If
        Exit Sub
    End If
    ilPos = InStr(slName, "*")
    If ilPos > 0 Then
        If (Right$(slName, 1) <> ".") Then
            lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\" & slName
            If imBrowserIndex = 0 Then
                edcBrowserFile.SetFocus
            Else
                lbcBrowserFile(imBrowserIndex).SetFocus
            End If
        End If
        Exit Sub
    End If
    ilPos = InStr(slName, "?")
    If ilPos > 0 Then
        If (Right$(slName, 1) <> ".") Then
            lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\" & slName
            If imBrowserIndex = 0 Then
                edcBrowserFile.SetFocus
            Else
                lbcBrowserFile(imBrowserIndex).SetFocus
            End If
        End If
        Exit Sub
    End If
    slFromFile = slName
    If InStr(slFromFile, ":") = 0 Then
        slFromFile = lbcBrowserPath.Path & "\" & slFromFile
    End If
    If (igBrowserType And Not SHIFT8) = 0 Then
        'Read file
        If InStr(slFromFile, ".") = 0 Then
            slFromFile = slFromFile & ".bmp"
        End If
        Screen.MousePointer = vbHourGlass
        pbcPicBrowser(1).Picture = LoadPicture(slFromFile)
        'vbcPicBrowser.Top = 0
        vbcPicBrowser.Max = pbcPicBrowser(1).Height - pbcPicBrowser(0).Height
        vbcPicBrowser.Enabled = (pbcPicBrowser(0).Height < pbcPicBrowser(1).Height)
        If vbcPicBrowser.Enabled Then
            vbcPicBrowser.SmallChange = pbcPicBrowser(0).Height
            vbcPicBrowser.LargeChange = pbcPicBrowser(0).Height
        End If
        Screen.MousePointer = vbDefault
    ElseIf (igBrowserType And Not SHIFT8) = 1 Then
        If InStr(slFromFile, ".") = 0 Then
            slFromFile = slFromFile & ".csv"
        End If
        Screen.MousePointer = vbHourGlass
        ilRet = mReadFile(slFromFile)
        Screen.MousePointer = vbDefault
    ElseIf (igBrowserType And Not SHIFT8) = 2 Then
        If InStr(slFromFile, ".") = 0 Then
            slFromFile = slFromFile & ".txt"
        End If
        Screen.MousePointer = vbHourGlass
        ilRet = mReadFile(slFromFile)
        Screen.MousePointer = vbDefault
    ElseIf (igBrowserType And Not SHIFT8) = 3 Then
        If InStr(slFromFile, ".") = 0 Then
            slFromFile = slFromFile & ".rec"
        End If
        Screen.MousePointer = vbHourGlass
        ilRet = mReadFile(slFromFile)
        Screen.MousePointer = vbDefault
    ElseIf (igBrowserType And Not SHIFT8) = 4 Then
        If InStr(slFromFile, ".") = 0 Then
            slFromFile = slFromFile & ".rt?"
        End If
        Screen.MousePointer = vbHourGlass
        ilRet = mReadFile(slFromFile)
        Screen.MousePointer = vbDefault
    ElseIf (igBrowserType And Not SHIFT8) = 5 Then
        If InStr(slFromFile, ".") = 0 Then
            slFromFile = slFromFile & ".ct?"
        End If
        Screen.MousePointer = vbHourGlass
        ilRet = mReadFile(slFromFile)
        Screen.MousePointer = vbDefault
    ElseIf (igBrowserType And Not SHIFT8) = 6 Then
        'If InStr(slFromFile, ".") = 0 Then
        '    slFromFile = slFromFile & ".txt"
        'End If
        Screen.MousePointer = vbHourGlass
        ilRet = mReadFile(slFromFile)
        Screen.MousePointer = vbDefault
    End If
    On Error GoTo 0
    Exit Sub
edcBrowserFileErr:
    ilRet = Err
    Resume Next
End Sub
Sub Form_Activate ()
    Dim ilPos As Integer
    Dim slDrive As String
    Dim slPath As String
'    gShowBranner
    If sgImportPath <> "" Then
        ilPos = InStr(sgImportPath, ":")
        If ilPos > 0 Then
            slDrive = Left$(sgImportPath, ilPos)
            slPath = Mid$(sgImportPath, ilPos + 1)
            If Right$(slPath, 1) = "/" Then
                slPath = Left$(slPath, Len(slPath) - 1)
            End If
            cbcBrowserDrive.Drive = slDrive
            lbcBrowserPath.Path = slPath
            If (igBrowserType And Not SHIFT8) = 0 Then
                If edcBrowserFile.Text <> "*.bmp" Then
                    edcBrowserFile.Text = "*.bmp"
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\*.bmp"
            ElseIf (igBrowserType And Not SHIFT8) = 1 Then
                If edcBrowserFile.Text <> "*.csv" Then
                    edcBrowserFile.Text = "*.csv"
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\*.csv"
            ElseIf (igBrowserType And Not SHIFT8) = 2 Then
                If edcBrowserFile.Text <> "*.txt" Then
                    edcBrowserFile.Text = "*.txt"
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\*.txt"
            ElseIf (igBrowserType And Not SHIFT8) = 3 Then
                If edcBrowserFile.Text <> "*.rec" Then
                    edcBrowserFile.Text = "*.rec"
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\*.rec"
            ElseIf (igBrowserType And Not SHIFT8) = 4 Then
                If edcBrowserFile.Text <> "*.rt?" Then
                    edcBrowserFile.Text = "*.rt?"
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\*.rt?"
            ElseIf (igBrowserType And Not SHIFT8) = 5 Then
                If edcBrowserFile.Text <> "*.ct?" Then
                    edcBrowserFile.Text = "*.ct?"
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\*.ct?"
            ElseIf (igBrowserType And Not SHIFT8) = 6 Then
                If edcBrowserFile.Text <> "Tape*.*" Then
                    edcBrowserFile.Text = "Tape*.*"
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\Tape*.*"
            End If
        End If
    End If
    If (igBrowserType And Not SHIFT8) = 0 Then
        pbcPicBrowser(0).Visible = True
        pbcPicBrowser(1).Visible = True
        vbcPicBrowser.Visible = True
        gdcBrowser.Visible = False
        lbcBrowser.Visible = False
    ElseIf (igBrowserType And Not SHIFT8) = 1 Then
        pbcPicBrowser(0).Visible = False
        pbcPicBrowser(1).Visible = False
        vbcPicBrowser.Visible = False
        lbcBrowser.Visible = False
        gdcBrowser.Visible = True
    ElseIf (igBrowserType And Not SHIFT8) = 2 Then
        pbcPicBrowser(0).Visible = False
        pbcPicBrowser(1).Visible = False
        vbcPicBrowser.Visible = False
        gdcBrowser.Visible = False
        lbcBrowser.Visible = True
    ElseIf (igBrowserType And Not SHIFT8) = 3 Then
        pbcPicBrowser(0).Visible = False
        pbcPicBrowser(1).Visible = False
        vbcPicBrowser.Visible = False
        gdcBrowser.Visible = False
        lbcBrowser.Visible = True
    ElseIf (igBrowserType And Not SHIFT8) = 4 Then
        pbcPicBrowser(0).Visible = False
        pbcPicBrowser(1).Visible = False
        vbcPicBrowser.Visible = False
        gdcBrowser.Visible = False
        lbcBrowser.Visible = True
    ElseIf (igBrowserType And Not SHIFT8) = 5 Then
        pbcPicBrowser(0).Visible = False
        pbcPicBrowser(1).Visible = False
        vbcPicBrowser.Visible = False
        gdcBrowser.Visible = False
        lbcBrowser.Visible = True
    ElseIf (igBrowserType And Not SHIFT8) = 6 Then
        pbcPicBrowser(0).Visible = False
        pbcPicBrowser(1).Visible = False
        vbcPicBrowser.Visible = False
        gdcBrowser.Visible = False
        lbcBrowser.Visible = True
    End If
End Sub
Sub Form_Click ()
    pbcClickFocus.SetFocus
End Sub
Sub Form_Load ()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Sub imcHelp_Click ()
    Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    Traffic!cdcSetup.Action = 6
End Sub
Sub lbcBrowserFile_Click (Index As Integer)
    Dim ilPos As Integer
    Dim slName As String
    Dim slFromFile As String
    Dim ilRet As Integer
    If lbcBrowserFile(imBrowserIndex).ListIndex < 0 Then
        Exit Sub
    End If
    edcBrowserFile.Text = Trim$(lbcBrowserFile(imBrowserIndex).List(lbcBrowserFile(imBrowserIndex).ListIndex))
    ''Read file
    'slFromFile = slName
    'If InStr(slFromFile, ":") = 0 Then
    '    slFromFile = lbcBrowserPath.Path & "\" & slFromFile
    'End If
    'Screen.MousePointer = vbHourGlass
    'ilRet = 0
    'On Error GoTo lbcBrowserFileErr:
    'pbcPicBrowser(1).Picture = LoadPicture(slFromFile)
    'Screen.MousePointer = vbDefault
    'Exit Sub
'lbcBrowserFileErr:
    'ilRet = Err
    'Resume Next
End Sub
Sub lbcBrowserPath_Change ()
    If (igBrowserType And Not SHIFT8) = 0 Then
        edcBrowserFile.Text = "*.bmp"
        lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\*.bmp"
    ElseIf (igBrowserType And Not SHIFT8) = 1 Then
        edcBrowserFile.Text = "*.csv"
        lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\*.csv"
    ElseIf (igBrowserType And Not SHIFT8) = 2 Then
        edcBrowserFile.Text = "*.txt"
        lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\*.txt"
    ElseIf (igBrowserType And Not SHIFT8) = 3 Then
        edcBrowserFile.Text = "*.rec"
        lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\*.rec"
    ElseIf (igBrowserType And Not SHIFT8) = 4 Then
        edcBrowserFile.Text = "*.rt?"
        lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\*.rt?"
    ElseIf (igBrowserType And Not SHIFT8) = 5 Then
        edcBrowserFile.Text = "*.ct?"
        lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\*.ct?"
    ElseIf (igBrowserType And Not SHIFT8) = 6 Then
        edcBrowserFile.Text = "Tape*.*"
        lbcBrowserFile(imBrowserIndex).FileName = lbcBrowserPath.Path & "\Tape*.*"
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Sub mInit ()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    imTerminate = False
    
    Screen.MousePointer = vbHourGlass
    Browser.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone Browser
    'Browser.Show
    Screen.MousePointer = vbHourGlass
'    mInitDDE
    pbcPicBrowser(0).Move 195, 2565
    vbcPicBrowser.Move pbcPicBrowser(0).Left + pbcPicBrowser(0).Width, pbcPicBrowser(0).Top, vbcPicBrowser.Width, pbcPicBrowser(0).Height
    gdcBrowser.Move 195, 2565
    lbcBrowser.Move 195, 2565
    If (igBrowserType And SHIFT8) = SHIFT8 Then
        imBrowserIndex = 1
        edcBrowserFile.Visible = False
        lbcBrowserFile(0).Visible = False
    Else
        imBrowserIndex = 0
        lbcBrowserFile(1).Visible = False
        lacFileName.Visible = False
    End If
    imcHelp.Picture = Traffic!imcHelp.Picture
'    gCenterModalForm Browser
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadFile                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File                      *
'*                                                     *
'*******************************************************
Function mReadFile (slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim ilMaxCol As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    lacFileName.Caption = slFromFile
    If (igBrowserType And Not SHIFT8) = 1 Then   'Grid
    ElseIf (igBrowserType And Not SHIFT8) = 2 Then   'Text
        lbcBrowser.Clear
    ElseIf (igBrowserType And Not SHIFT8) = 3 Then   'Text
        lbcBrowser.Clear
    ElseIf (igBrowserType And Not SHIFT8) = 4 Then   'Text
        lbcBrowser.Clear
    ElseIf (igBrowserType And Not SHIFT8) = 5 Then   'Text
        lbcBrowser.Clear
    ElseIf (igBrowserType And Not SHIFT8) = 6 Then   'Text
        lbcBrowser.Clear
    End If
    ilRet = 0
    On Error GoTo mReadFileErr:
    hmFrom = FreeFile
    Open slFromFile For Input Access Read As hmFrom
    If ilRet <> 0 Then
        Close hmFrom
        MsgBox "Open " & slFromFile, vbExclamation, "Open Error"
        cmcCancel.SetFocus
        mReadFile = False
        Exit Function
    End If
    If (igBrowserType And Not SHIFT8) = 1 Then
        ilRow = 1
        ilMaxCol = 0
    End If
    Do
        On Error GoTo mReadFileErr:
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            Exit Do
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                If (igBrowserType And Not SHIFT8) = 1 Then   'Grid
                    ReDim smFieldValues(1 To 100) As String
                    gParseCDFields slLine, False, smFieldValues()
                    For ilLoop = UBound(smFieldValues) To LBound(smFieldValues) Step -1
                        If smFieldValues(ilLoop) <> "" Then
                            ilCol = ilLoop
                            If ilLoop + 1 > ilMaxCol Then
                                ilMaxCol = ilLoop + 1
                                gdcBrowser.Cols = ilMaxCol
                                gdcBrowser.Row = 0
                                For ilIndex = 1 To ilCol Step 1
                                    gdcBrowser.Col = ilIndex
                                    gdcBrowser.Text = Trim$(Str$(ilIndex))
                                Next ilIndex
                            End If
                            Exit For
                        End If
                    Next ilLoop
                    ilRow = ilRow + 1
                    gdcBrowser.Rows = ilRow
                    gdcBrowser.Row = ilRow - 1
                    gdcBrowser.Col = 0
                    gdcBrowser.Text = Trim$(Str$(ilRow - 1))
                    For ilLoop = LBound(smFieldValues) To ilCol Step 1
                        gdcBrowser.Row = ilRow - 1
                        gdcBrowser.Col = ilLoop
                        gdcBrowser.Text = smFieldValues(ilLoop)
                    Next ilLoop
                    'gdcBrowser.Cols = ilCol    '2
                    'gdcBrowser.Col = ilCol - 1
                    'gdcBrowser.Row = ilRow - 1
                    'gdcBrowser.Text = slLine
                    If ilRow > 100 Then
                        ilRow = ilRow + 1
                        gdcBrowser.Rows = ilRow
                        gdcBrowser.Row = ilRow - 1
                        gdcBrowser.Col = 0
                        gdcBrowser.Text = Trim$(Str$(ilRow - 1))
                        gdcBrowser.Row = ilRow - 1
                        gdcBrowser.Col = 1
                        gdcBrowser.Text = "..."
                        Exit Do
                    End If
                ElseIf (igBrowserType And Not SHIFT8) = 2 Then   'Text
                    lbcBrowser.AddItem slLine
                    If lbcBrowser.ListCount > 100 Then
                        lbcBrowser.AddItem "...."
                        Exit Do
                    End If
                ElseIf (igBrowserType And Not SHIFT8) = 3 Then   'Text
                    lbcBrowser.AddItem slLine
                    If lbcBrowser.ListCount > 100 Then
                        lbcBrowser.AddItem "...."
                        Exit Do
                    End If
                ElseIf (igBrowserType And Not SHIFT8) = 4 Then   'Text
                    lbcBrowser.AddItem slLine
                    If lbcBrowser.ListCount > 100 Then
                        lbcBrowser.AddItem "...."
                        Exit Do
                    End If
                ElseIf (igBrowserType And Not SHIFT8) = 5 Then   'Text
                    lbcBrowser.AddItem slLine
                    If lbcBrowser.ListCount > 100 Then
                        lbcBrowser.AddItem "...."
                        Exit Do
                    End If
                ElseIf (igBrowserType And Not SHIFT8) = 6 Then   'Text
                    lbcBrowser.AddItem slLine
                    If lbcBrowser.ListCount > 100 Then
                        lbcBrowser.AddItem "...."
                        Exit Do
                    End If
                End If
            End If
        End If
    Loop Until ilEof
    Close hmFrom
    mReadFile = True
    MousePointer = vbDefault
    Exit Function
mReadFileErr:
    ilRet = Err
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Sub mTerminate ()
'
'   mTerminate
'   Where:
'
    Dim ilRet As Integer
    Erase smFieldValues
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Browser
    Set Browser = Nothing   'Remove data segment
    igManUnload = NO
End Sub
Sub pbcClickFocus_KeyUp (KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        Traffic!cdcSetup.Action = 6
    End If
End Sub
Sub plcScreen_Click ()
    pbcClickFocus.SetFocus
End Sub
Sub vbcPicBrowser_Change ()
    pbcPicBrowser(1).Top = -vbcPicBrowser.Value
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Browser"
End Sub
