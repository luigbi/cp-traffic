VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBrowse 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   9360
   ControlBox      =   0   'False
   Icon            =   "AffBrowse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pbcArial 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7605
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   5055
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3285
      TabIndex        =   7
      Top             =   4875
      Width           =   1245
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4860
      TabIndex        =   6
      Top             =   4875
      Width           =   1245
   End
   Begin VB.FileListBox lbcBrowserFile 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   4260
      TabIndex        =   2
      Top             =   465
      Width           =   3780
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
      Left            =   270
      TabIndex        =   1
      Top             =   465
      Width           =   3675
   End
   Begin VB.DirListBox lbcBrowserPath 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   990
      Left            =   270
      TabIndex        =   0
      Top             =   825
      Width           =   3675
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8415
      Top             =   4995
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5430
      FormDesignWidth =   9360
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdBrowser 
      Height          =   2355
      Left            =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2235
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   4154
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      Rows            =   10
      Cols            =   1
      FixedCols       =   0
      BackColorSel    =   12632256
      ForeColorSel    =   16711680
      BackColorUnpopulated=   16777215
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.Label lacErrorMsg 
      Alignment       =   2  'Center
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   225
      TabIndex        =   10
      Top             =   4605
      Width           =   8760
   End
   Begin VB.Label lacSample 
      Alignment       =   2  'Center
      Caption         =   "Sample"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1980
      Width           =   8760
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
      Left            =   4260
      TabIndex        =   4
      Top             =   195
      Width           =   3765
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
      Left            =   270
      TabIndex        =   3
      Top             =   195
      Width           =   3600
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private hmFrom As Integer
Const CALLLETTERS = 1



Private Sub cbcBrowserDrive_Change()
    Screen.MousePointer = vbHourglass
    lbcBrowserPath.Path = cbcBrowserDrive.Drive
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmcCancel_Click()
    igBrowseReturn = 0
    sgBrowseFile = ""
    Unload frmBrowse
End Sub

Private Sub cmcDone_Click()

    Dim slName As String
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim slStr As String

    igBrowseReturn = 1
    slStr = lbcBrowserPath.Path
    If right$(slStr, 1) <> "\" Then
        slStr = slStr & "\"
    End If
    sgBrowseFile = slStr & lbcBrowserFile.List(lbcBrowserFile.ListIndex)
    
    Unload frmBrowse
    
End Sub

Private Sub Form_Activate()
    Dim ilPos As Integer
    Dim slDrive As String
    Dim slPath As String
    
    ilPos = InStr(sgImportDirectory, ":")
    If ilPos > 0 Then
        slDrive = Left$(sgImportDirectory, ilPos)
        slPath = Mid$(sgImportDirectory, ilPos + 1)
        If right$(slPath, 1) = "/" Then
            slPath = Left$(slPath, Len(slPath) - 1)
        End If
        cbcBrowserDrive.Drive = slDrive
        lbcBrowserPath.Path = slPath
    End If
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / 1.1
    Me.Height = (Screen.Height) / 1.3
    Me.Top = (Screen.Height - Me.Height) / 1.4
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts Me
    gCenterForm Me
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    'Me.Width = (Screen.Width) / 2
    'Me.Height = (Screen.Height) / 4
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    Screen.MousePointer = vbDefault

End Sub

Private Function mGetFile() As Integer
    Dim slFromFile As String
    Dim slLine As String
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slChar As String
    Dim llRow As Long
    Dim ilCol As Integer
    Dim blError1 As Boolean
    Dim blError2 As Boolean
    Dim blError3 As Boolean
    Dim blTitleFound As Boolean
    Dim blFound As Boolean
    'Dim slFields(1 To 90) As String
    Dim slFields(0 To 89) As String
    
    'On Error GoTo mTrapFileOpenError:
    mGetFile = False
    lacErrorMsg.Caption = ""
    Select Case igBrowseType
        Case 1  'Import station Update data
            grdBrowser.Redraw = False
            gGrid_Clear grdBrowser, True
            llRow = grdBrowser.FixedRows
    End Select
    slStr = lbcBrowserPath.Path
    If right$(slStr, 1) <> "\" Then
        slStr = slStr & "\"
    End If
    sgBrowseFile = slStr & lbcBrowserFile.List(lbcBrowserFile.ListIndex)
    slFromFile = sgBrowseFile
    'ilRet = 0
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        cmcDone.Enabled = False
        Close hmFrom
        gMsgBox "Unable to open file. Error = " & Trim$(Str$(ilRet))
        Exit Function
    End If
    blError1 = False
    blError2 = False
    blError3 = False
    blTitleFound = False
    Do While Not EOF(hmFrom)
        ilRet = 0
        'Line Input #hmFrom, slLine
        slLine = ""
        Do While Not EOF(hmFrom)
            slChar = Input(1, #hmFrom)
            If slChar = sgLF Then
                Exit Do
            ElseIf slChar <> sgCR Then
                slLine = slLine & slChar
            End If
        Loop
        If ilRet <> 0 Then
            Exit Do
        End If
        Select Case igBrowseType
            Case 1  'Import station Update data
                gParseCDFields slLine, False, slFields()
                If Trim$(slFields(UBound(slFields) - 1)) <> "" Then
                    blError1 = True
                End If
                If Trim$(slFields(UBound(slFields))) <> "" Then
                    blError1 = False
                End If
                If llRow >= grdBrowser.Rows Then
                    grdBrowser.AddItem ""
                End If
                If Not blTitleFound Then
                    'If StrComp(UCase$(Trim$(slFields(1))), UCase$(Trim$(sgStationImportTitles(CALLLETTERS))), vbTextCompare) = 0 Then
                    If StrComp(UCase$(Trim$(slFields(0))), UCase$(Trim$(sgStationImportTitles(CALLLETTERS))), vbTextCompare) = 0 Then
                        For ilCol = LBound(slFields) To UBound(slFields) Step 1
                            If slFields(ilCol) = "" Then
                                Exit For
                            End If
                            'Test Name
                            blFound = False
                            For ilLoop = 1 To UBound(sgStationImportTitles) Step 1
                                If StrComp(UCase$(Trim$(slFields(ilCol))), UCase$(Trim$(sgStationImportTitles(ilLoop))), vbTextCompare) = 0 Then
                                    blFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            'grdBrowser.Col = ilCol '- 1
                            If ilCol >= grdBrowser.Cols Then
                                grdBrowser.Cols = grdBrowser.Cols + 1
                                'grdBrowser.Col = ilCol '- 1
                            End If
                            grdBrowser.Col = ilCol
                            'Incresing cols, causes row to be reset to grdBrowser.FixedRow
                            grdBrowser.Row = 0
                            If blFound Then
                                grdBrowser.CellForeColor = vbBlack
                            Else
                                blError2 = True
                                grdBrowser.CellForeColor = vbRed
                            End If
                            'grdBrowser.TextMatrix(0, ilCol - 1) = slFields(ilCol)
                            grdBrowser.TextMatrix(0, ilCol) = slFields(ilCol)
                        Next ilCol
                        blTitleFound = True
                    End If
                Else
                    'For ilCol = 0 To grdBrowser.Cols - 1 Step 1
                    '    If (ilCol = SUIDINDEX) Or (ilCol = SUFREQINDEX) Or (ilCol = SUDMARANKINDEX) Or (ilCol = SUMSARANKINDEX) Then
                    '        If (Trim$(slFields(ilCol + 1)) <> "") And (Trim$(slFields(ilCol + 1)) <> "N/A") Then
                    '            If (Asc(slFields(ilCol + 1)) < Asc("0")) Or (Asc(slFields(ilCol + 1)) > Asc("9")) Then
                    '                blError3 = True
                    '            End If
                    '        End If
                    '    End If
                    '    grdBrowser.TextMatrix(llRow, ilCol) = Trim$(slFields(ilCol + 1))
                    'Next ilCol
                    If llRow >= grdBrowser.Rows Then
                        grdBrowser.AddItem ""
                    End If
                    For ilCol = 0 To grdBrowser.Cols - 1 Step 1
                        'grdBrowser.TextMatrix(llRow, ilCol) = Trim$(slFields(ilCol + 1))
                        grdBrowser.TextMatrix(llRow, ilCol) = Trim$(slFields(ilCol))
                    Next ilCol
                    llRow = llRow + 1
                    If llRow > 20 Then
                        Exit Do
                    End If
                End If
        End Select
    Loop
    Close hmFrom
    Select Case igBrowseType
        Case 1  'Import station Update data
            lacErrorMsg.Caption = ""
            If Not blTitleFound Then
                lacErrorMsg.Caption = "Titles Not Found"
            End If
            If blError1 Then
                lacErrorMsg.Caption = "Too many fields defined"
            End If
            If blError2 Then
                lacErrorMsg.Caption = "Titles Not Valid"
            End If
            If blError3 Then
                If lacErrorMsg.Caption = "" Then
                    lacErrorMsg.Caption = "Numeric Field in Error"
                Else
                    lacErrorMsg.Caption = lacErrorMsg.Caption & ", " & "Numeric Field in Error"
                End If
            End If
            If (Not blTitleFound) Or (blError1) Or (blError2) Or (blError3) Then
                cmcDone.Enabled = False
            Else
                cmcDone.Enabled = True
            End If
            grdBrowser.Redraw = True
    End Select
    Screen.MousePointer = vbDefault
    
    mGetFile = True
    Exit Function

'mTrapFileOpenError:
'    ilRet = Err.Number
'    Resume Next
ErrHand:
    gMsgBox "A general error occured in mGetFile."
    cmcDone.Enabled = False
End Function

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    lbcBrowserPath.Top = cbcBrowserDrive.Top + cbcBrowserDrive.Height + 120
    lbcBrowserPath.Height = lbcBrowserFile.Top + lbcBrowserFile.Height - lbcBrowserPath.Top
    Select Case igBrowseType
        Case 1  'Import station Update data
            mSetStationUpdateGridColumns
            mSetStationUpdateGridTitles
            gGrid_Clear grdBrowser, True
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBrowse = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub lbcBrowserFile_Click()
    Screen.MousePointer = vbHourglass
    If lbcBrowserFile.ListIndex >= 0 Then
        mGetFile
    Else
        'Clear Grid
        Select Case igBrowseType
            Case 1  'Import station Update data
                gGrid_Clear grdBrowser, True
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub lbcBrowserPath_Change()
    Dim slStr As String
    slStr = lbcBrowserPath.Path
    If right$(slStr, 1) <> "\" Then
        slStr = slStr & "\"
    End If
    lbcBrowserFile.fileName = slStr & sgBrowseMaskFile
End Sub

Private Sub mSetStationUpdateGridColumns()
    Dim ilCol As Integer
    
'    grdBrowser.ColWidth(SUSTATIONINDEX) = pbcArial.TextWidth("WWWW-AM")    'grdBrowser.Width * 0.07
'    grdBrowser.ColWidth(SUIDINDEX) = pbcArial.TextWidth("999999999")    'grdBrowser.Width * 0.16
'    grdBrowser.ColWidth(SUFREQINDEX) = pbcArial.TextWidth("FREQUENCY")    'grdBrowser.Width * 0.1
'    grdBrowser.ColWidth(SUFORMATINDEX) = pbcArial.TextWidth("FORMATFORMATFORMAT")    'grdBrowser.Width * 0.1
'    grdBrowser.ColWidth(SUDMARANKINDEX) = pbcArial.TextWidth("DMA RANK ")    'grdBrowser.Width * 0.06
'    grdBrowser.ColWidth(SUDMANAMEINDEX) = pbcArial.TextWidth("DMANAMEDMANAMEDMANAME")    'grdBrowser.Width * 0.03
'    grdBrowser.ColWidth(SUSTATEINDEX) = pbcArial.TextWidth("STATE")    'grdBrowser.Width * 0.13
'    grdBrowser.ColWidth(SULICENSEINDEX) = pbcArial.TextWidth("CITYLICENSECITYLICENSE")    'grdBrowser.Width * 0.13
'    grdBrowser.ColWidth(SUOWNERINDEX) = pbcArial.TextWidth("OWNERNAMEOWNERNAME")    'grdBrowser.Width * 0.13
'    grdBrowser.ColWidth(SUMSARANKINDEX) = pbcArial.TextWidth("MSA RANK ")    'grdBrowser.Width * 0.13
'    grdBrowser.ColWidth(SUMSANAMEINDEX) = pbcArial.TextWidth("MSANAMEMSANAMEMSANAME")    'grdBrowser.Width * 0.13
    
    'Align columns to left
    gGrid_AlignAllColsLeft grdBrowser
End Sub

Private Sub mSetStationUpdateGridTitles()
    'Set column titles
'    grdBrowser.TextMatrix(0, SUSTATIONINDEX) = "Station"
'    grdBrowser.TextMatrix(0, SUIDINDEX) = "ID #"
'    grdBrowser.TextMatrix(0, SUFREQINDEX) = "Frequency"
'    grdBrowser.TextMatrix(0, SUFORMATINDEX) = "Format"
'    grdBrowser.TextMatrix(0, SUDMARANKINDEX) = "DMA Rank"
'    grdBrowser.TextMatrix(0, SUDMANAMEINDEX) = "DMA Name"
'    grdBrowser.TextMatrix(0, SUSTATEINDEX) = "State"
'    grdBrowser.TextMatrix(0, SULICENSEINDEX) = "License"
'    grdBrowser.TextMatrix(0, SUOWNERINDEX) = "Owner"
'    grdBrowser.TextMatrix(0, SUMSARANKINDEX) = "MSA Rank"
'    grdBrowser.TextMatrix(0, SUMSANAMEINDEX) = "MSA Name"

End Sub
