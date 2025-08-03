VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form fCrViewerExport 
   Caption         =   "Export Report"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frcFile 
      Caption         =   "Save to File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2070
      Left            =   990
      TabIndex        =   1
      Top             =   150
      Width           =   5505
      Begin VB.ComboBox cbcFileType 
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
         Left            =   825
         TabIndex        =   6
         Top             =   570
         Width           =   2955
      End
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4005
         TabIndex        =   4
         Top             =   1140
         Width           =   1005
      End
      Begin VB.TextBox edcFileName 
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
         Height          =   300
         Left            =   810
         TabIndex        =   3
         Top             =   1155
         Width           =   2925
      End
      Begin VB.CommandButton cmcUserChoice 
         Appearance      =   0  'Flat
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2145
         TabIndex        =   2
         Top             =   1635
         Width           =   1320
      End
      Begin VB.Label lacType 
         Caption         =   "Format"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   150
         TabIndex        =   7
         Top             =   585
         Width           =   630
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
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
         Height          =   225
         Left            =   150
         TabIndex        =   5
         Top             =   1185
         Width           =   645
      End
   End
   Begin VB.CommandButton cmcUserChoice 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   6660
      Top             =   1575
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".pdf"
      Filter          =   "Pdf|*.pdf|Xls|*.xls|Doc|.Doc|Txt|*.txt|Csv|*.csv|Rtf|*.rtf|All Files|*.*"
      FilterIndex     =   1
      FontSize        =   0
      MaxFileSize     =   256
   End
End
Attribute VB_Name = "fCrViewerExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim smFileName As String
Dim imChgMode As Integer
Dim imBSMode As Integer
Dim imFTSelectedIndex As Integer
Dim imComboBoxIndex As Integer
Public bmContinue As Boolean
Const EXPORTSHOW = 8

Private Sub Form_Load()
    mInit
End Sub
Private Sub mInit()
    Dim vlArray As Variant
    
    gPopExportTypes cbcFileType, True
    'see more options
    mAddType
    If Not ogReport Is Nothing Then
        vlArray = ogReport.Reports.Keys
        smFileName = Replace(vlArray(0), ".rpt", " ", 1, 1, vbTextCompare)
        edcFileName = smFileName
        mSetCommands
    Else
        MsgBox "Problem with exporting", vbExclamation, "Error"
        Unload fCrViewerExport
    End If

End Sub
Private Sub mAddType()
  cbcFileType.AddItem "See Advanced options..."

End Sub
Private Sub mSetCommands()

    Dim ilEnable As Integer
    ilEnable = True
    If ilEnable Then
        If (imFTSelectedIndex >= 0) And (edcFileName.Text <> "") Then
            ilEnable = True
        Else
            ilEnable = False
        End If
    End If
    cmcUserChoice(0).Enabled = ilEnable
End Sub
Private Sub cmcBrowse_Click()
'Pdf|*.pdf|Xls|*.xls|Doc|.Doc|Txt|*.txt|Csv|*.csv|Rtf|*.rtf|All Files|*.*
    Dim slSavePath As String
    Dim myConnections As New CCsiSystemConnection
    Dim ilSelection As Integer
    
    Set myConnections = New CCsiSystemConnection
    slSavePath = myConnections.GlobalReportSavePath
    cdcSetup.Flags = cdlOFNPathMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
    Select Case cbcFileType.ListIndex
        Case 0
            ilSelection = 1
        Case 1, 2, 3
            ilSelection = 2
        Case 4
            ilSelection = 3
        Case 5
            ilSelection = 4
        Case 6
            ilSelection = 5
        Case 7
            ilSelection = 6
        Case Else
            ilSelection = 7
    End Select
    cdcSetup.FilterIndex = ilSelection
    cdcSetup.fileName = edcFileName.Text
    cdcSetup.InitDir = Left$(slSavePath, Len(slSavePath) - 1)
    cdcSetup.Action = 2 'DLG_FILE_SAVE
    edcFileName.Text = cdcSetup.fileName
    gChDrDir
    mSetCommands
    Set myConnections = Nothing
End Sub

Private Sub cmcUserChoice_Click(Index As Integer)
    Dim ilRet As Integer
    Select Case Index
        Case 0  'go' button
            If cbcFileType.ListIndex <> EXPORTSHOW Then
                smFileName = Trim(edcFileName.Text)
                'PDF?
                If cbcFileType.ListIndex = 0 And ogReport.Reports.Count = 1 Then
                    'bmcontinue set to true here
                    ExportInfo.Show vbModal
                Else
                    bmContinue = True
                End If
                    If bmContinue Then
                        ilRet = ogReport.Export(smFileName, cbcFileType.ListIndex, False)
                    If ilRet = 0 Then
                        MsgBox "File was not saved", vbExclamation + vbOKOnly, "Error"
                    End If
                End If
            End If
            bmContinue = False
    End Select  'currently, cancel button simply takes you back to display.
    Unload fCrViewerExport
End Sub

Private Sub edcFileName_Change()
 mSetCommands
End Sub
Private Sub cbcFileType_Change()
    If imChgMode = False Then
        imChgMode = True
        If cbcFileType.Text <> "" Then
            gManLookAhead cbcFileType, imBSMode, imComboBoxIndex
        End If
        imFTSelectedIndex = cbcFileType.ListIndex
        imChgMode = False
    End If
    mSetCommands
End Sub

Private Sub cbcFileType_Click()

    If cbcFileType.ListIndex = EXPORTSHOW Then
        cmcUserChoice(0).Caption = "Continue"
        Report.bmShowExportForm = True
    Else
        Report.bmShowExportForm = False
        cmcUserChoice(0).Caption = "Save"
        imComboBoxIndex = cbcFileType.ListIndex
        imFTSelectedIndex = cbcFileType.ListIndex
        mSetCommands
    End If
End Sub
Private Sub cbcFileType_GotFocus()
    If cbcFileType.Text = "" Then
        cbcFileType.ListIndex = 0
    End If
    imComboBoxIndex = cbcFileType.ListIndex
    gCtrlGotFocus cbcFileType
End Sub
Private Sub cbcFileType_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcFileType_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcFileType.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
