VERSION 5.00
Begin VB.Form FileMakerToOrder 
   Caption         =   "FileMaker to Order"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
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
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   0
      Width           =   75
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   2700
      TabIndex        =   4
      Top             =   360
      Width           =   2700
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2160
      TabIndex        =   3
      Top             =   3840
      Width           =   1425
   End
   Begin VB.CommandButton cmcImport 
      Appearance      =   0  'Flat
      Caption         =   "Continue to &Orders"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   360
      TabIndex        =   2
      Top             =   3840
      Width           =   1425
   End
   Begin VB.PictureBox plcModel 
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
      Height          =   2670
      Left            =   480
      ScaleHeight     =   2610
      ScaleWidth      =   2955
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   3015
      Begin VB.ListBox lbcFileMaker 
         Appearance      =   0  'Flat
         Height          =   2370
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.Label lblResult 
      Caption         =   "Result"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   3360
      Width           =   2775
   End
End
Attribute VB_Name = "FileMakerToOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const FILEPATHMAIN As String = "Relevant"
Const FILEPATHIN As String = "Input"
Const FILEPATHOUT As String = "Output"
Const FORMNAME As String = "FileMakerToOrder"
Const MODIFICATIONS As String = "Modifications"
Const ISSUESFILENAME As String = "FileMakerToOrdersIssues"
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim omFile As FileSystemObject
Dim smFullPathIn As String
Dim smFullPathOut As String
Dim myIssues As CLogger

Private Sub mInit()
    Dim ilFileCount As Integer
    
On Error GoTo ERRORBOX

    imTerminate = False
    imFirstActivate = True
    lblResult.Caption = ""
    Screen.MousePointer = vbHourglass
    'Me.Height = cmcImport.Top + 5 * cmcImport.Height / 3
    gCenterStdAlone Me
    Me.ZOrder vbBringToFront
    Screen.MousePointer = vbHourglass
    ogContractCreator.CreationUser = Filemaker
    Set omFile = New FileSystemObject
    If mBuildFolderPaths(sgExePath) Then
        ilFileCount = mLoadFileMakerFiles()
    Else
        mSetControlsToAnIssue "Could not build folder paths"
        Exit Sub
    End If
    Set myIssues = New CLogger
    myIssues.LogPath = myIssues.CreateLogName(sgDBPath & "Messages\" & ISSUESFILENAME)
    If Len(myIssues.ErrorMessage) = 0 Then
        'continue with previous
        If Len(ogContractCreator.fileName) > 0 Then
            If ogContractCreator.ProcessResult = Success Then
                lblResult.Caption = ogContractCreator.fileName & " processed as order #" & ogContractCreator.ContractNumber
                If mFileProcessFinish(ogContractCreator.fileName) Then
                   ilFileCount = mLoadFileMakerFiles()
                Else
                    mSetControlsToAnIssue "Could not finish processing file " & ogContractCreator.fileName, False
                    myIssues.WriteWarning lblResult.Caption
                End If
            ElseIf ogContractCreator.ProcessResult = Cancelled Then
                If ogContractCreator.DeleteOrder() Then
                    mSetControlsToAnIssue ogContractCreator.fileName & " was not saved", False
                Else
                    mSetControlsToAnIssue "There was an issue handling this unsaved file", False
                    myIssues.WriteWarning "Contact Counterpoint.  Could not delete Contract #" & ogContractCreator.ContractNumber & " " & ogContractCreator.ErrorMessage
                End If
            Else
                mSetControlsToAnIssue ogContractCreator.fileName & " on the orders screen.", False
                myIssues.WriteWarning lblResult.Caption & ogContractCreator.ErrorMessage
                If Not ogContractCreator.DeleteOrder() Then
                    myIssues.WriteWarning "Contact Counterpoint.  Could not delete Contract #" & ogContractCreator.ContractNumber & " " & ogContractCreator.ErrorMessage
                End If

            End If
            ogContractCreator.Clear
        'first time in (if they'd chosen cancelled previously, they'd never come back in)
        Else
            If ilFileCount < 0 Then
                mSetControlsToAnIssue "Could not load files"
                myIssues.WriteWarning "Could not load files: " & ogContractCreator.ErrorMessage
            End If
        End If
    Else
        mSetControlsToAnIssue "Could not create 'issues' file"
        gLogMsg FORMNAME & "-mInit Could not create 'issues' file: " & myIssues.ErrorMessage, "TrafficErrors.txt", False
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
ERRORBOX:
    gHandleError "", FORMNAME & "-mInit"
End Sub

Private Sub mTerminate()
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Me
    igManUnload = NO
End Sub





Private Sub lbcFileMaker_DblClick()
    cmcImport_Click
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Pending FileMaker Orders"
End Sub
Private Function mBuildFolderPaths(slSiblingPath As String) As Boolean
    Dim slParentPath As String
    Dim blRet As Boolean
    Dim slPathToMainFolder As String
    blRet = True
    On Error GoTo ERRORBOX
    slParentPath = omFile.GetParentFolderName(slSiblingPath)
    If (omFile.FolderExists(slParentPath)) Then
        slPathToMainFolder = omFile.BuildPath(slParentPath, FILEPATHMAIN)
        If Not omFile.FolderExists(slPathToMainFolder) Then
            omFile.CreateFolder (slPathToMainFolder)
        End If
        smFullPathIn = omFile.BuildPath(slPathToMainFolder, FILEPATHIN)
        smFullPathOut = omFile.BuildPath(slPathToMainFolder, FILEPATHOUT)
        If (Not omFile.FolderExists(smFullPathIn)) Then
            omFile.CreateFolder (smFullPathIn)
        End If
        If (Not omFile.FolderExists(smFullPathOut)) Then
            omFile.CreateFolder (smFullPathOut)
        End If
    Else
        blRet = False
    End If
    mBuildFolderPaths = blRet
    Exit Function
ERRORBOX:
    gHandleError "", FORMNAME & "-mBuildFolderPaths"
    mBuildFolderPaths = False
End Function
Private Function mLoadFileMakerFiles() As Integer
    Dim ilRet As Integer
    Dim slFileName As String
    Dim olFolder As Folder
    Dim olFile As file
    
    ilRet = 0
On Error GoTo ERRORBOX
    lbcFileMaker.Clear
    lbcFileMaker.AddItem MODIFICATIONS
    Set olFolder = omFile.GetFolder(smFullPathIn)
    For Each olFile In olFolder.Files
        slFileName = olFile.Name
        If InStr(UCase(slFileName), ".XML") Then
            ilRet = ilRet + 1
            lbcFileMaker.AddItem slFileName
        End If
    Next
    mLoadFileMakerFiles = ilRet
    Exit Function
ERRORBOX:
    mLoadFileMakerFiles = -1
    gHandleError "", FORMNAME & "-mLoadFileMakerFiles"
End Function
Private Function mFileProcessFinish(slFileMakerCurrentFile As String) As Boolean
    Dim slPathAndFileIn As String
    Dim slPathAndFileOut As String
    
On Error GoTo ERRBOX
    If Len(slFileMakerCurrentFile) > 0 Then
        slPathAndFileIn = omFile.BuildPath(smFullPathIn, slFileMakerCurrentFile)
        slPathAndFileOut = omFile.BuildPath(smFullPathOut, slFileMakerCurrentFile)
        If omFile.FileExists(slPathAndFileIn) Then
            If omFile.FileExists(slPathAndFileOut) Then
                omFile.DeleteFile slPathAndFileOut
            End If
            omFile.MoveFile slPathAndFileIn, slPathAndFileOut
        Else
            lblResult.ForeColor = vbRed
            lblResult.Caption = "could not move file!"
        End If
    End If
    mFileProcessFinish = True
    Exit Function
ERRBOX:
    mFileProcessFinish = False
End Function
Private Sub mSetControlsToAnIssue(slErrorMessage As String, Optional blBlockImport As Boolean = True)
    lblResult.Caption = slErrorMessage
    lblResult.ForeColor = vbRed
    If blBlockImport Then
        cmcImport.Enabled = False
    End If
End Sub
Private Sub cmcCancel_Click()
    ogContractCreator.Clear True
    mTerminate
End Sub

Private Sub cmcImport_Click()
    mImport
End Sub
Private Sub mImport()
    Dim blContinue As Boolean
    
    blContinue = True
    lblResult.Caption = ""
    If lbcFileMaker.ListCount = 1 Then
        lbcFileMaker.Text = MODIFICATIONS
    End If
    If Len(lbcFileMaker.Text) > 0 Then
        If lbcFileMaker.Text = MODIFICATIONS Then
            ogContractCreator.ContractToProcess = NoProcess
        Else
            'test name first
            If ogContractCreator.LoadXml(smFullPathIn, lbcFileMaker.Text) Then
                If ogContractCreator.ValidateXML() Then
                    ogContractCreator.ContractToProcess = Order
                    blContinue = ogContractCreator.CreateOrder()
                    If Not blContinue Then
                        mSetControlsToAnIssue "Error in validation", False
                        myIssues.WriteWarning "Contact Counterpoint. " & lbcFileMaker.Text & " failed validation:" & ogContractCreator.ErrorMessage, True
                        blContinue = False
                    End If
                Else
                    mSetControlsToAnIssue "Failed validation", False
                    myIssues.WriteWarning lbcFileMaker.Text & " failed validation:" & ogContractCreator.ErrorMessage, True
                    blContinue = False
                End If
            Else
                    mSetControlsToAnIssue "Failed to load", False
                    myIssues.WriteWarning lbcFileMaker.Text & " failed to load:" & ogContractCreator.ErrorMessage, True
                    blContinue = False
            End If
        End If
        If blContinue Then
            Basic10!tmcFilemakerToOrder.Enabled = True
            mTerminate
        End If
    Else
        MsgBox "Please select an option before continuing to orders", vbInformation, "No Option Chosen"
    End If
End Sub
Private Sub cmcImport_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub Form_Activate()
    mInit
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
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
   ' mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Set FileMakerToOrder = Nothing

End Sub

