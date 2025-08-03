VERSION 5.00
Begin VB.Form Browser 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5175
   ClientLeft      =   1155
   ClientTop       =   1800
   ClientWidth     =   11520
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5175
   ScaleWidth      =   11520
   Begin VB.PictureBox plcBrowser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   90
      ScaleHeight     =   4320
      ScaleWidth      =   11205
      TabIndex        =   3
      Top             =   285
      Width           =   11265
      Begin VB.CheckBox ckcAll 
         Caption         =   "All"
         Height          =   210
         Left            =   180
         TabIndex        =   18
         Top             =   2025
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.FileListBox lbcBrowserFile 
         Appearance      =   0  'Flat
         Height          =   1290
         Index           =   1
         Left            =   195
         MultiSelect     =   2  'Extended
         Pattern         =   "*.bmp"
         TabIndex        =   9
         Top             =   345
         Width           =   5160
      End
      Begin VB.ListBox lbcBrowser 
         Appearance      =   0  'Flat
         Height          =   1500
         Left            =   195
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2430
         Width           =   10425
      End
      Begin VB.DirListBox lbcBrowserPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1290
         Left            =   6000
         TabIndex        =   11
         Top             =   705
         Width           =   5040
      End
      Begin VB.DriveListBox cbcBrowserDrive 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5985
         TabIndex        =   10
         Top             =   360
         Width           =   5040
      End
      Begin VB.FileListBox lbcBrowserFile 
         Appearance      =   0  'Flat
         Height          =   1290
         Index           =   0
         Left            =   195
         Pattern         =   "*.bmp"
         TabIndex        =   8
         Top             =   720
         Width           =   5160
      End
      Begin VB.TextBox edcBrowserFile 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   195
         TabIndex        =   7
         Top             =   360
         Width           =   5160
      End
      Begin VB.PictureBox pbcPicBrowser 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1725
         Index           =   0
         Left            =   195
         ScaleHeight     =   1695
         ScaleWidth      =   10140
         TabIndex        =   5
         Top             =   2565
         Width           =   10170
         Begin VB.PictureBox pbcPicBrowser 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1515
            Index           =   1
            Left            =   255
            ScaleHeight     =   1515
            ScaleWidth      =   5940
            TabIndex        =   6
            Top             =   -150
            Width           =   5940
         End
      End
      Begin VB.VScrollBar vbcPicBrowser 
         Height          =   1545
         Left            =   10380
         TabIndex        =   4
         Top             =   2550
         Width           =   240
      End
      Begin VB.ListBox lbcGridBrowser 
         Appearance      =   0  'Flat
         Height          =   1500
         Left            =   195
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2325
         Width           =   10425
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   195
         TabIndex        =   15
         Top             =   2310
         Width           =   6105
      End
      Begin VB.Label lacBrowserPath 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "File Path"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6000
         TabIndex        =   12
         Top             =   120
         Width           =   4560
      End
      Begin VB.Label lacBrowserFile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "File Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   420
         TabIndex        =   13
         Top             =   120
         Width           =   2490
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   6060
      TabIndex        =   2
      Top             =   4875
      Width           =   945
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   1770
      Width           =   75
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Left            =   4665
      TabIndex        =   0
      Top             =   4875
      Width           =   945
   End
   Begin VB.Label lacScreen 
      Height          =   210
      Left            =   30
      TabIndex        =   17
      Top             =   15
      Width           =   8700
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   345
      Top             =   4755
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Browser.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

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
Dim imFirstActivate As Integer
Dim hmFrom As Integer
Dim smFieldValues() As String
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBrowserIndex As Integer
Dim imAllClicked As Integer
Dim imSetAll As Integer


Private Sub cbcBrowserDrive_Change()
    Screen.MousePointer = vbHourglass
    lbcBrowserPath.Path = cbcBrowserDrive.Drive
    'lbcBrowserFile(imBrowserIndex).Cls
    If (igBrowserType And Not SHIFT8) = 8 Then
        If Not imAllClicked Then
            imSetAll = False
            ckcAll.Value = vbUnchecked
            imSetAll = True
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim ilValue As Integer
    Dim llRet As Long
    Dim llRg As Long
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        If lbcBrowserFile(1).ListCount > 0 Then
            llRg = CLng(lbcBrowserFile(1).ListCount - 1) * &H10000 + 0
            'llRet = SendMessageByNum(lbcLines.hwnd, &H400 + 28, ilValue, llRg)
            llRet = SendMessageByNum(lbcBrowserFile(1).hwnd, LB_SELITEMRANGE, ilValue, llRg)
        End If
        imAllClicked = False
    End If

End Sub

Private Sub cmcCancel_Click()
    igBrowserReturn = 0
    mTerminate
End Sub
Private Sub cmcDone_Click()
    Dim slName As String
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim blFound As Boolean

    igBrowserReturn = 1
    If (igBrowserType And Not SHIFT8) = 8 Then
        blFound = False
        For ilLoop = 0 To lbcBrowserFile(1).ListCount - 1 Step 1
            If lbcBrowserFile(1).Selected(ilLoop) Then
                blFound = True
                Exit For
            End If
        Next ilLoop
        If Not blFound Then
            Beep
            lbcBrowserFile(1).SetFocus
            Exit Sub
        End If
    Else
        slName = Trim$(edcBrowserFile.Text)
        If Len(slName) <= 0 Then
            Beep
            edcBrowserFile.SetFocus
            Exit Sub
        End If
    End If
    slStr = lbcBrowserPath.Path
    If right$(slStr, 1) <> "\" Then
        slStr = slStr & "\"
    End If
    ilPos = InStr(slName, "*")
    If ilPos > 0 Then
        Beep
        lbcBrowserFile(imBrowserIndex).fileName = slStr & slName
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
        lbcBrowserFile(imBrowserIndex).fileName = slStr & slName
        If imBrowserIndex = 0 Then
            edcBrowserFile.SetFocus
        Else
            lbcBrowserFile(imBrowserIndex).SetFocus
        End If
        Exit Sub
    End If
    If imBrowserIndex = 0 Then
        sgBrowserFile = slName
        'If InStr(sgBrowserFile, ":") = 0 Then
        If (InStr(sgBrowserFile, ":") = 0) And (Left$(sgBrowserFile, 2) <> "\\") Then
            sgBrowserFile = slStr & sgBrowserFile
        End If
    Else
        If (igBrowserType And Not SHIFT8) = 8 Then
            sgBrowserDrivePath = slStr
            slStr = ""
        End If
        sgBrowserFile = ""
        For ilLoop = 0 To lbcBrowserFile(1).ListCount - 1 Step 1
            If lbcBrowserFile(1).Selected(ilLoop) Then
                sgBrowserFile = sgBrowserFile & slStr & lbcBrowserFile(1).List(ilLoop) & "|"
            End If
        Next ilLoop
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcBrowserFile_Change()
    Dim ilPos As Integer
    Dim slName As String
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilClearListBox As Integer

    '1-6-05 Loop thru the entries gathered from the folder.  For each selected entry,
    'show the data in the lbcBrowser list box
    ilClearListBox = -1
    sgBrowserFile = ""
    If (igBrowserType And Not SHIFT8) = 8 Then
        Exit Sub
    End If
    For ilLoop = 0 To lbcBrowserFile(imBrowserIndex).ListCount - 1 Step 1
        If lbcBrowserFile(imBrowserIndex).Selected(ilLoop) Then
            slName = sgBrowserFile & slStr & lbcBrowserFile(imBrowserIndex).List(ilLoop)

            If ilClearListBox = -1 Or imBrowserIndex = 0 Then      'first time, or only single selection allowed (index = 0)
                ilClearListBox = 0
            Else
                ilClearListBox = 1
            End If

    'slName = Trim$(edcBrowserFile.Text)
            If Len(slName) <= 0 Then
                If imBrowserIndex = 0 Then
                    edcBrowserFile.SetFocus
                Else
                    lbcBrowserFile(imBrowserIndex).SetFocus
                End If
                Exit Sub
            End If
            slStr = lbcBrowserPath.Path
            If right$(slStr, 1) <> "\" Then
                slStr = slStr & "\"
            End If
            ilPos = InStr(slName, "*")
            If ilPos > 0 Then
                If (right$(slName, 1) <> ".") Then
                    lbcBrowserFile(imBrowserIndex).fileName = slStr & slName
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
                If (right$(slName, 1) <> ".") Then
                    lbcBrowserFile(imBrowserIndex).fileName = slStr & slName
                    If imBrowserIndex = 0 Then
                        edcBrowserFile.SetFocus
                    Else
                        lbcBrowserFile(imBrowserIndex).SetFocus
                    End If
                End If
                Exit Sub
            End If
            slFromFile = slName
            'If InStr(slFromFile, ":") = 0 Then
            If (InStr(slFromFile, ":") = 0) And (Left$(slFromFile, 2) <> "\\") Then
                slFromFile = slStr & slFromFile
            End If
            If (igBrowserType And Not SHIFT8) = 0 Then
                'Read file
                If InStr(slFromFile, ".") = 0 Then
                    slFromFile = slFromFile & ".bmp"
                End If
                Screen.MousePointer = vbHourglass
                If gFileExist(slFromFile) = 0 Then
                    pbcPicBrowser(1).Picture = LoadPicture(slFromFile)
                End If
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
                Screen.MousePointer = vbHourglass
                ilRet = mReadFile(slFromFile, ilClearListBox)
                Screen.MousePointer = vbDefault
            ElseIf (igBrowserType And Not SHIFT8) = 2 Then
                If InStr(slFromFile, ".") = 0 Then
                    slFromFile = slFromFile & ".txt"
                End If
                Screen.MousePointer = vbHourglass
                ilRet = mReadFile(slFromFile, ilClearListBox)
                Screen.MousePointer = vbDefault
            ElseIf (igBrowserType And Not SHIFT8) = 3 Then
                If InStr(slFromFile, ".") = 0 Then
                    slFromFile = slFromFile & ".rec"
                End If
                Screen.MousePointer = vbHourglass
                ilRet = mReadFile(slFromFile, ilClearListBox)
                Screen.MousePointer = vbDefault
            ElseIf (igBrowserType And Not SHIFT8) = 4 Then
                If InStr(slFromFile, ".") = 0 Then
                    slFromFile = slFromFile & ".rt?"
                End If
                Screen.MousePointer = vbHourglass
                ilRet = mReadFile(slFromFile, ilClearListBox)
                Screen.MousePointer = vbDefault
            ElseIf (igBrowserType And Not SHIFT8) = 5 Then
                If InStr(slFromFile, ".") = 0 Then
                    slFromFile = slFromFile & ".ct?"
                End If
                Screen.MousePointer = vbHourglass
                ilRet = mReadFile(slFromFile, ilClearListBox)
                Screen.MousePointer = vbDefault
            ElseIf (igBrowserType And Not SHIFT8) = 6 Then
                'If InStr(slFromFile, ".") = 0 Then
                '    slFromFile = slFromFile & ".txt"
                'End If
                Screen.MousePointer = vbHourglass
                ilRet = mReadFile(slFromFile, ilClearListBox)
                Screen.MousePointer = vbDefault
            ElseIf (igBrowserType And Not SHIFT8) = 7 Then
                'If InStr(slFromFile, ".") = 0 Then
                '    slFromFile = slFromFile & ".txt"
                'End If
                Screen.MousePointer = vbHourglass
                ilRet = mReadFile(slFromFile, ilClearListBox)
                Screen.MousePointer = vbDefault
            End If
        End If          'not selected
    Next ilLoop         'next entry in list box, see if selected
    On Error GoTo 0
    Exit Sub
End Sub
Private Sub Form_Activate()
    Dim ilPos As Integer
    Dim slDrive As String
    Dim slPath As String
    Dim slStr As String
    Dim slImportPath As String
    Dim slDrivePath As String

    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
'    gShowBranner

    '1-5-05 use general import path unless coming from Automation
    slImportPath = sgImportPath
    If igProphetImportPathFlag = 1 Then      'change to automation import path
        slImportPath = sgProphetImportPath
    ElseIf igProphetImportPathFlag = 2 Then     '1-10-12 its wide orbit
        slImportPath = sgWideOrbitImportPath
    ElseIf igProphetImportPathFlag = 3 Then     '6-27-12 Jelli
        slImportPath = sgJelliImportPath
    ElseIf igProphetImportPathFlag = 4 Then     '1-7-16 zetta
        slImportPath = sgZettaImportPath
    End If


    If (igBrowserType And Not SHIFT8) = 8 Then
        slImportPath = sgStationInvoiceImportPath
    End If

    If slImportPath <> "" Then
        ilPos = InStr(slImportPath, ":")
        If ilPos > 0 Then
            slDrive = Left$(slImportPath, ilPos)
            slPath = Mid$(slImportPath, ilPos + 1)
            If right$(slPath, 1) = "/" Then
                slPath = Left$(slPath, Len(slPath) - 1)
            End If
            cbcBrowserDrive.Drive = slDrive
            lbcBrowserPath.Path = slPath
            slStr = lbcBrowserPath.Path
            If right$(slStr, 1) <> "\" Then
                slStr = slStr & "\"
            End If
            If (igBrowserType And Not SHIFT8) = 0 Then
                If edcBrowserFile.Text <> "*.bmp" Then
                    edcBrowserFile.Text = "*.bmp"
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.bmp"
            ElseIf (igBrowserType And Not SHIFT8) = 1 Then
                If edcBrowserFile.Text <> "*.csv" Then
                    edcBrowserFile.Text = "*.csv"
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.csv"
            ElseIf (igBrowserType And Not SHIFT8) = 2 Then
                If edcBrowserFile.Text <> "*.txt" Then
                    edcBrowserFile.Text = "*.txt"
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.txt"
            ElseIf (igBrowserType And Not SHIFT8) = 3 Then
                If edcBrowserFile.Text <> "*.rec" Then
                    edcBrowserFile.Text = "*.rec"
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.rec"
            ElseIf (igBrowserType And Not SHIFT8) = 4 Then
                If edcBrowserFile.Text <> "*.rt?" Then
                    edcBrowserFile.Text = "*.rt?"
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.rt?"
            ElseIf (igBrowserType And Not SHIFT8) = 5 Then
                If edcBrowserFile.Text <> "*.ct?" Then
                    edcBrowserFile.Text = "*.ct?"
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.ct?"
            ElseIf (igBrowserType And Not SHIFT8) = 6 Then
                If edcBrowserFile.Text <> "Tape*.*" Then
                    edcBrowserFile.Text = "Tape*.*"
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).fileName = slStr & "Tape*.*"
            ElseIf (igBrowserType And Not SHIFT8) = 7 Then
                If edcBrowserFile.Text <> sgBrowseMaskFile Then
                    edcBrowserFile.Text = sgBrowseMaskFile
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).fileName = slStr & sgBrowseMaskFile
            ElseIf (igBrowserType And Not SHIFT8) = 8 Then
                If edcBrowserFile.Text <> "*.pdf;*.Txt;*.EDI" Then
                    edcBrowserFile.Text = "*.pdf;*.Txt;*.EDI"
                Else
                    edcBrowserFile_Change
                End If
                lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.pdf;*.Txt;*.EDI"
            End If
        End If
    End If
    If (igBrowserType And Not SHIFT8) = 0 Then
        pbcPicBrowser(0).Visible = True
        pbcPicBrowser(1).Visible = True
        vbcPicBrowser.Visible = True
        lbcGridBrowser.Visible = False
        lbcBrowser.Visible = False
    ElseIf (igBrowserType And Not SHIFT8) = 1 Then
        pbcPicBrowser(0).Visible = False
        pbcPicBrowser(1).Visible = False
        vbcPicBrowser.Visible = False
        lbcBrowser.Visible = False
        lbcGridBrowser.Visible = True
    ElseIf (igBrowserType And Not SHIFT8) = 2 Then
        pbcPicBrowser(0).Visible = False
        pbcPicBrowser(1).Visible = False
        vbcPicBrowser.Visible = False
        lbcGridBrowser.Visible = False
        lbcBrowser.Visible = True
    ElseIf (igBrowserType And Not SHIFT8) = 3 Then
        pbcPicBrowser(0).Visible = False
        pbcPicBrowser(1).Visible = False
        vbcPicBrowser.Visible = False
        lbcGridBrowser.Visible = False
        lbcBrowser.Visible = True
    ElseIf (igBrowserType And Not SHIFT8) = 4 Then
        pbcPicBrowser(0).Visible = False
        pbcPicBrowser(1).Visible = False
        vbcPicBrowser.Visible = False
        lbcGridBrowser.Visible = False
        lbcBrowser.Visible = True
    ElseIf (igBrowserType And Not SHIFT8) = 5 Then
        pbcPicBrowser(0).Visible = False
        pbcPicBrowser(1).Visible = False
        vbcPicBrowser.Visible = False
        lbcGridBrowser.Visible = False
        lbcBrowser.Visible = True
    ElseIf (igBrowserType And Not SHIFT8) = 6 Then
        pbcPicBrowser(0).Visible = False
        pbcPicBrowser(1).Visible = False
        vbcPicBrowser.Visible = False
        lbcGridBrowser.Visible = False
        lbcBrowser.Visible = True
    ElseIf (igBrowserType And Not SHIFT8) = 7 Then
        pbcPicBrowser(0).Visible = False
        pbcPicBrowser(1).Visible = False
        vbcPicBrowser.Visible = False
        lbcGridBrowser.Visible = False
        lbcBrowser.Visible = True
    ElseIf (igBrowserType And Not SHIFT8) = 8 Then
        pbcPicBrowser(0).Visible = False
        pbcPicBrowser(1).Visible = False
        vbcPicBrowser.Visible = False
        lbcGridBrowser.Visible = False
        lbcBrowser.Visible = False
    End If
    Me.KeyPreview = True
    Me.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase smFieldValues
    
    Set Browser = Nothing   'Remove data segment
    
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub


Private Sub lacScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub lbcBrowserFile_Click(Index As Integer)
    If (igBrowserType And Not SHIFT8) = 8 Then
        If Not imAllClicked Then
            imSetAll = False
            ckcAll.Value = vbUnchecked
            imSetAll = True
        End If
    End If
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
Private Sub lbcBrowserPath_Change()
    Dim slStr As String
    slStr = lbcBrowserPath.Path
    If right$(slStr, 1) <> "\" Then
        slStr = slStr & "\"
    End If
    If (igBrowserType And Not SHIFT8) = 0 Then
        edcBrowserFile.Text = "*.bmp"
        lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.bmp"
    ElseIf (igBrowserType And Not SHIFT8) = 1 Then
        edcBrowserFile.Text = "*.csv"
        lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.csv"
    ElseIf (igBrowserType And Not SHIFT8) = 2 Then
        edcBrowserFile.Text = "*.txt"
        lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.txt"
    ElseIf (igBrowserType And Not SHIFT8) = 3 Then
        edcBrowserFile.Text = "*.rec"
        lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.rec"
    ElseIf (igBrowserType And Not SHIFT8) = 4 Then
        edcBrowserFile.Text = "*.rt?"
        lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.rt?"
    ElseIf (igBrowserType And Not SHIFT8) = 5 Then
        edcBrowserFile.Text = "*.ct?"
        lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.ct?"
    ElseIf (igBrowserType And Not SHIFT8) = 6 Then
        edcBrowserFile.Text = "Tape*.*"
        lbcBrowserFile(imBrowserIndex).fileName = slStr & "Tape*.*"
    ElseIf (igBrowserType And Not SHIFT8) = 7 Then
        edcBrowserFile.Text = sgBrowseMaskFile
        lbcBrowserFile(imBrowserIndex).fileName = slStr & sgBrowseMaskFile
    ElseIf (igBrowserType And Not SHIFT8) = 8 Then
        edcBrowserFile.Text = "*.pdf;*.Txt;*.EDI"
        lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.pdf;*.Txt;*.EDI"
        If Not imAllClicked Then
            imSetAll = False
            ckcAll.Value = vbUnchecked
            imSetAll = True
        End If
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
Private Sub mInit()
'
'   mInit
'   Where:
'
    imTerminate = False
    imFirstActivate = True
    
    imSetAll = True
        
    Screen.MousePointer = vbHourglass
    If Trim$(sgBrowserTitle) <> "" Then
        lacScreen.Caption = sgBrowserTitle
    Else
        lacScreen.Caption = "Browser"
    End If
    Browser.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone Browser
    'Browser.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    pbcPicBrowser(0).Move 195, 2565
    vbcPicBrowser.Move pbcPicBrowser(0).Left + pbcPicBrowser(0).Width, pbcPicBrowser(0).Top, vbcPicBrowser.Width, pbcPicBrowser(0).Height
    lbcGridBrowser.Move 195, 2565
    lbcBrowser.Move 195, 2565
    If (igBrowserType And Not SHIFT8) = 8 Then
        imBrowserIndex = 1
        edcBrowserFile.Visible = False
        lbcBrowserFile(0).Visible = False
        ckcAll.Visible = True
        'ckcAll.Top = lbcBrowserFile(1).Top + lbcBrowserFile(1).Height + 60
        'plcBrowser.Height = ckcAll.Top + ckcAll.Height + lbcBrowserFile(1).Top
        'cmcDone.Top = 2 * plcBrowser.Top + plcBrowser.Height
        'Browser.Height = cmcDone.Top + 5 * cmcDone.Height / 3
        ckcAll.Top = plcBrowser.Height - (3 * ckcAll.Height / 2)
        lbcBrowserFile(1).Height = ckcAll.Top - lbcBrowserFile(1).Top
        lbcBrowserPath.Height = lbcBrowserFile(1).Height - cbcBrowserDrive.Height - 60
        cmcCancel.Top = cmcDone.Top
        gCenterStdAlone Browser
    Else
        imBrowserIndex = 0
        lbcBrowserFile(1).Visible = False
        lacFileName.Visible = False
    End If
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    gCenterModalForm Browser
    Screen.MousePointer = vbDefault
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
'*
'*      1-6-05 add flag to tell whether to intialize the
'*             list box of data because multiple files
'*             selection is allowed.
'*******************************************************
Private Function mReadFile(slFromFile As String, ilClearListBox As Integer) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilRow As Integer
    Dim ilMaxCol As Integer
    If (igBrowserType And Not SHIFT8) = 8 Then
        mReadFile = True
        Exit Function
    End If
    lacFileName.Caption = slFromFile
    If ilClearListBox = 0 Then           'first time thru, init the list box; otherwise append to it
        If (igBrowserType And Not SHIFT8) = 1 Then   'Grid
            'Currently just using list box
            lbcGridBrowser.Clear
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
        ElseIf (igBrowserType And Not SHIFT8) = 7 Then   'Text
            lbcBrowser.Clear
        ElseIf (igBrowserType And Not SHIFT8) = 8 Then   'Text
        End If
    End If
    ilRet = 0
    'On Error GoTo mReadFileErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        MsgBox "Open " & slFromFile & ", Error #" & str$(ilRet), vbExclamation, "Open Error"
        cmcCancel.SetFocus
        mReadFile = False
        Exit Function
    End If
    If (igBrowserType And Not SHIFT8) = 1 Then
        ilRow = 1
        ilMaxCol = 0
    End If
    Err.Clear
    Do
        'On Error GoTo mReadFileErr:
        If EOF(hmFrom) Then
            Exit Do
        End If
        Line Input #hmFrom, slLine
        On Error GoTo 0
        ilRet = Err.Number
        If ilRet = 62 Then
            Exit Do
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                If (igBrowserType And Not SHIFT8) = 1 Then   'Grid
''                    ReDim smFieldValues(1 To 100) As String
'                    ReDim smFieldValues(0 To 99) As String
'                    gParseCDFields slLine, False, smFieldValues()
'                    For ilLoop = UBound(smFieldValues) To LBound(smFieldValues) Step -1
'                        If smFieldValues(ilLoop) <> "" Then
'                            ilCol = ilLoop
'                            If ilLoop + 1 > ilMaxCol Then
'                                ilMaxCol = ilLoop + 1
'                                lbcGridBrowser.Cols = ilMaxCol
'                                lbcGridBrowser.Row = 0
'                                For ilIndex = 1 To ilCol Step 1
'                                    lbcGridBrowser.Col = ilIndex
'                                    lbcGridBrowser.Text = Trim$(Str$(ilIndex))
'                                Next ilIndex
'
'                            End If
'                            Exit For
'                        End If
'                    Next ilLoop
'                    ilRow = ilRow + 1
'                    lbcGridBrowser.Rows = ilRow
'                    lbcGridBrowser.Row = ilRow - 1
'                    lbcGridBrowser.Col = 0
'                    lbcGridBrowser.Text = Trim$(Str$(ilRow - 1))
'                    For ilLoop = LBound(smFieldValues) To ilCol Step 1
'                        lbcGridBrowser.Row = ilRow - 1
'                        lbcGridBrowser.Col = ilLoop
'                        lbcGridBrowser.Text = smFieldValues(ilLoop)
'                    Next ilLoop
'                    'lbcGridBrowser.Cols = ilCol    '2
'                    'lbcGridBrowser.Col = ilCol - 1
'                    'lbcGridBrowser.Row = ilRow - 1
'                    'lbcGridBrowser.Text = slLine
'                    If ilRow > 100 Then
'                        ilRow = ilRow + 1
'                        lbcGridBrowser.Rows = ilRow
'                        lbcGridBrowser.Row = ilRow - 1
'                        lbcGridBrowser.Col = 0
'                        lbcGridBrowser.Text = Trim$(Str$(ilRow - 1))
'                        lbcGridBrowser.Row = ilRow - 1
'                        lbcGridBrowser.Col = 1
'                        lbcGridBrowser.Text = "..."
'                        Exit Do
'                    End If
                    'Not using grid at this time- just show in list box
                    lbcGridBrowser.AddItem slLine
                    If lbcGridBrowser.ListCount > 100 Then
                        lbcGridBrowser.AddItem "...."
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
                ElseIf (igBrowserType And Not SHIFT8) = 7 Then   'Text
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
'mReadFileErr:
'    ilRet = Err.Number
'    Resume Next
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
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Browser
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub vbcPicBrowser_Change()
    pbcPicBrowser(1).Top = -vbcPicBrowser.Value
End Sub
