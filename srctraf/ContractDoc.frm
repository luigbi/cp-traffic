VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form ContractDoc 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5205
   ClientLeft      =   18855
   ClientTop       =   2625
   ClientWidth     =   10080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5205
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pbcPermissions_AddEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9270
      ScaleHeight     =   165
      ScaleWidth      =   495
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   510
      Width           =   525
   End
   Begin VB.PictureBox pbcPermissions_Remove 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9270
      ScaleHeight     =   165
      ScaleWidth      =   495
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   735
      Width           =   525
   End
   Begin VB.Timer tmrColumns 
      Interval        =   50
      Left            =   8625
      Top             =   4590
   End
   Begin VB.Frame Frame1 
      Height          =   630
      Left            =   2685
      TabIndex        =   11
      Top             =   5370
      Visible         =   0   'False
      Width           =   3720
      Begin VB.TextBox edcDescription 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         MaxLength       =   100
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   150
         Visible         =   0   'False
         Width           =   3270
      End
   End
   Begin VB.TextBox txtRecordCount 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1755
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   4290
      Width           =   630
   End
   Begin MSComDlg.CommonDialog cdcAttachments 
      Left            =   9165
      Top             =   4545
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvcAttachments 
      Height          =   3195
      Left            =   165
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   5636
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imglstListview"
      SmallIcons      =   "imglstListview"
      ColHdrIcons     =   "imglstListview"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmcRemove 
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6150
      TabIndex        =   4
      Top             =   4605
      Width           =   2010
   End
   Begin VB.CommandButton cmcAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3885
      TabIndex        =   3
      Top             =   4605
      Width           =   2010
   End
   Begin VB.CommandButton cmcDone 
      Cancel          =   -1  'True
      Caption         =   "&Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1590
      TabIndex        =   2
      Top             =   4605
      Width           =   2010
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   30
      ScaleHeight     =   165
      ScaleWidth      =   105
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3600
      Width           =   105
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   1765
      ScaleHeight     =   30
      ScaleWidth      =   15
      TabIndex        =   5
      Top             =   135
      Width           =   15
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   45
      TabIndex        =   6
      Top             =   1695
      Width           =   45
   End
   Begin VB.Label Label1 
      Caption         =   "Permissions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7455
      TabIndex        =   17
      Top             =   165
      Width           =   1575
   End
   Begin VB.Label lbcPermissions_AddEdit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Add/Edit Attachments"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7575
      TabIndex        =   16
      Top             =   510
      Width           =   1575
   End
   Begin VB.Label lbcPermissions_Remove 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Remove Attachments"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7575
      TabIndex        =   15
      Top             =   720
      Width           =   1605
   End
   Begin VB.Shape shpPermissions 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   585
      Left            =   7440
      Top             =   405
      Width           =   2445
   End
   Begin VB.Label lblContract 
      Caption         =   "Contract #:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   12
      Top             =   495
      Width           =   3330
   End
   Begin VB.Image imgAttachmentIcon 
      Enabled         =   0   'False
      Height          =   420
      Left            =   135
      Picture         =   "ContractDoc.frx":0000
      Stretch         =   -1  'True
      Top             =   165
      Width           =   555
   End
   Begin VB.Label lacRecordCount 
      Caption         =   "Total Record Count:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   9
      Top             =   4290
      Width           =   1440
   End
   Begin VB.Label lacContractAttachments 
      Caption         =   "Available Contract Attachments"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   870
      TabIndex        =   8
      Top             =   195
      Width           =   3330
   End
End
Attribute VB_Name = "ContractDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Form created Feb 2024 - JJB
'   Contract Attachments SOW

Option Explicit
Option Compare Text

Dim tmContractDocs() As ACFList
Dim lmCntrNo As Long
Dim bmViewOnly As Boolean
Dim bmAttachments_Add As Boolean
Dim bmAttachments_Remove As Boolean

Dim imTotalRecords As Integer
Dim imContractDocs_Count As Integer
Dim imContractDocsAdded_Count As Integer
Dim imUserCount As Integer

Dim smUsers(5000, 1) As String
  
Private Const COLUMN_FILENAME = 0
Private Const COLUMN_MODIFYDATE = 1
Private Const COLUMN_DESCRIPTION = 2
Private Const COLUMN_USERNAME = 3
Private Const COLUMN_CODE = 4
Private Const COLUMN_PATH = 5
Private Const COLUMN_CHANGED = 6

Private Const COLUMNWIDTH_FILENAME = 2425
Private Const COLUMNWIDTH_MODIFYDATE = 1750
Private Const COLUMNWIDTH_DESCRIPTION = 3725
Private Const COLUMNWIDTH_USERNAME = 1800
Private Const COLUMNWIDTH_CODE = 0
Private Const COLUMNWIDTH_PATH = 0
Private Const COLUMNWIDTH_CHANGED = 0
           
Private Const CHANGETYPE_NONE = 0
Private Const CHANGETYPE_NEW = 1
Private Const CHANGETYPE_DELETED = 2
Private Const CHANGETYPE_DESCRIPTIONCHANGED = 3

Private Const CSIDL_PERSONAL = &H5
Private Const NOERROR = 0

Private Type SHTEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHTEMID
End Type

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
'Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Property Let CONTRACTNO(ByVal llCntrNo As Long)
    lmCntrNo = llCntrNo
End Property

Public Property Let isViewOnly(ByVal blIsView As Boolean)
    bmViewOnly = blIsView
End Property

Sub mGetUsers()
    
    Dim rs As ADODB.Recordset
    Dim i As Integer

    SQLQuery = " select distinct"
    SQLQuery = SQLQuery & vbCrLf & "     urfCode, "
    SQLQuery = SQLQuery & vbCrLf & "     urfName "
    SQLQuery = SQLQuery & vbCrLf & " from "
    SQLQuery = SQLQuery & vbCrLf & "     URF_User_Options"
  
    Set rs = gSQLSelectCall(SQLQuery)

    Erase smUsers
    imUserCount = 0
    
    i = 0
    While Not rs.EOF
        smUsers(i, 0) = rs!urfCode
        smUsers(i, 1) = Trim$(gDecryptField(Trim$(rs!urfName)))
        i = i + 1
        imUserCount = imUserCount + 1
        rs.MoveNext
    Wend
    
    rs.Close
    
End Sub

Sub mSetEditDescription()

    Dim i As Integer
    Dim ItemSel As ListItem
    
    If Not lvcAttachments.SelectedItem Is Nothing And Not bmViewOnly Then
        
        i = COLUMN_DESCRIPTION
        
        With Frame1
            .Visible = True
            .Top = lvcAttachments.SelectedItem.Top + lvcAttachments.Top
            .Left = lvcAttachments.ColumnHeaders(i + 1).Left + lvcAttachments.Left + 5
            .Width = lvcAttachments.ColumnHeaders(i + 1).Width
            .height = lvcAttachments.SelectedItem.height
            .ZOrder 0
        End With
        
        With edcDescription
            .Visible = True
            .Tag = Trim(lvcAttachments.SelectedItem.SubItems(i))
            .Text = Trim(lvcAttachments.SelectedItem.SubItems(i))
            .SetFocus
            .SelStart = 0
            .Left = 0
            .Top = 0
            .Width = lvcAttachments.ColumnHeaders(i + 1).Width
            .height = lvcAttachments.SelectedItem.height
            .SelLength = Len(.Text)
        End With
    End If
    
End Sub

Sub mSetRecordCount()

    txtRecordCount.Text = lvcAttachments.ListItems.Count
    
    If lvcAttachments.ListItems.Count = 0 Or bmViewOnly = True Then
        cmcRemove.Enabled = False
    Else
        If bmViewOnly = False And bmAttachments_Remove = True Then
            cmcRemove.Enabled = True
        Else
            cmcRemove.Enabled = False
        End If
    End If
            
End Sub

Sub SetPermissions()
    
    Dim sView As String
    
    bmAttachments_Add = IIF(tgUrf(0).sAddAttach = "Y" Or Trim(tgUrf(0).sName) = "Guide", True, False)
    bmAttachments_Remove = IIF(tgUrf(0).sRemoveAttach = "Y" Or Trim(tgUrf(0).sName) = "Guide", True, False)

    If Not bmViewOnly = True Then
        cmcAdd.Enabled = IIF(bmAttachments_Add, True, False)
        edcDescription.Locked = IIF(bmAttachments_Add, False, True)
        pbcPermissions_AddEdit.BackColor = IIF(bmAttachments_Add, vbGreen, vbRed)

        cmcRemove.Enabled = IIF(bmAttachments_Remove, True, False)
        pbcPermissions_Remove.BackColor = IIF(bmAttachments_Remove, vbGreen, vbRed)
        
        If bmAttachments_Add = True Then
            lbcPermissions_AddEdit.ToolTipText = "You have permissions to add attachments or modify descriptions.  This was inherited from the User Options screen."
            pbcPermissions_AddEdit.ToolTipText = "You have permissions to add attachments or modify descriptions.  This was inherited from the User Options screen."
        Else
            lbcPermissions_AddEdit.ToolTipText = "You do NOT have permissions to add attachments or modify descriptions.  This was inherited from the User Options screen."
            pbcPermissions_AddEdit.ToolTipText = "You do NOT have permissions to add attachments or modify descriptions.  This was inherited from the User Options screen."
        End If
        
        If bmAttachments_Remove = True Then
            lbcPermissions_Remove.ToolTipText = "You have permissions to remove attachments.  This was inherited from the User Options screen."
            pbcPermissions_Remove.ToolTipText = "You have permissions to remove attachments.  This was inherited from the User Options screen."
        Else
            lbcPermissions_Remove.ToolTipText = "You do NOT have permissions to remove attachments.  This was inherited from the User Options screen."
            pbcPermissions_Remove.ToolTipText = "You do NOT have permissions to remove attachments.  This was inherited from the User Options screen."
        End If
        
    Else
        cmcAdd.Enabled = False
        cmcRemove.Enabled = False
        edcDescription.Locked = True
        pbcPermissions_AddEdit.BackColor = vbRed
        pbcPermissions_Remove.BackColor = vbRed
        lbcPermissions_AddEdit.ToolTipText = "You do NOT have permissions to add attachments or modify descriptions.  This was inherited from the previous screen."
        pbcPermissions_AddEdit.ToolTipText = "You do NOT have permissions to add attachments or modify descriptions.  This was inherited from the previous screen."
        lbcPermissions_Remove.ToolTipText = "You do NOT have permissions to remove attachments.  This was inherited from the previous screen."
        pbcPermissions_Remove.ToolTipText = "You do NOT have permissions to remove attachments.  This was inherited from the previous screen."
    End If
    
End Sub

'Private Function GetSpecialfolder(CSIDL As Long) As String
'
'    Dim r As Long
'    Dim IDL As ITEMIDLIST
'    Dim Path$
'
'    'Get the special folder
'    r = SHGetSpecialFolderLocation(100, CSIDL, IDL)
'    If r = NOERROR Then
'        'Create a buffer
'        Path$ = Space$(512)
'        'Get the path from the IDList
'        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
'        'Remove the unnecessary chr$(0)'s
'        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
'        Exit Function
'    End If
'    GetSpecialfolder = ""
'
'End Function

Private Sub mInit()
    
    gCenterModalForm ContractDoc
    tmContractDocs = tgContractDocuments
    lblContract.Caption = "Contract #:  " & IIF(lmCntrNo = 0, "NEW", lmCntrNo)
    
    SetPermissions
    mAttachments_Populate
    Set_Tooltips
    
End Sub

Sub mAttachments_Populate()
  
    Dim i As Integer
    Dim X As Integer
    Dim iListviewRecord As Integer
    Dim iNumberOfAttachments As Integer
    
    mGetUsers
    
    With lvcAttachments
        .View = lvwReport
              
        With .ColumnHeaders
            .Add , , "Filename", COLUMNWIDTH_FILENAME
            .Add , , "Modified Date", COLUMNWIDTH_MODIFYDATE, lvwColumnCenter
            .Add , , "Description", COLUMNWIDTH_DESCRIPTION
            .Add , , "User", COLUMNWIDTH_USERNAME
            .Add , , "Code", COLUMNWIDTH_CODE
            .Add , , "Path", COLUMNWIDTH_PATH
            .Add , , "Changed", COLUMNWIDTH_CHANGED
        End With
    End With
    
    With lvcAttachments
        iListviewRecord = 0
        For i = 0 To UBound(tmContractDocs) - 1
            If tmContractDocs(i).bDeleted = False Then
                iListviewRecord = iListviewRecord + 1
                
                .ListItems.Add , , Trim$(tmContractDocs(i).sFileName)
                .ListItems(iListviewRecord).SubItems(COLUMN_MODIFYDATE) = tmContractDocs(i).sEnteredDate
                .ListItems(iListviewRecord).SubItems(COLUMN_DESCRIPTION) = tmContractDocs(i).sDescription
                For X = 0 To imUserCount - 1
                    If tmContractDocs(i).iUrfCode = smUsers(X, 0) Then
                        .ListItems(iListviewRecord).SubItems(COLUMN_USERNAME) = smUsers(X, 1)
                        Exit For
                    End If
                Next X
                .ListItems(iListviewRecord).SubItems(COLUMN_CODE) = tmContractDocs(i).lCode
                .ListItems(iListviewRecord).SubItems(COLUMN_PATH) = tmContractDocs(i).sTempPath
                .ListItems(iListviewRecord).SubItems(COLUMN_CHANGED) = CHANGETYPE_NONE
            End If
        Next i
        
    End With
    
    imContractDocs_Count = UBound(tmContractDocs)
    mSetRecordCount
    
End Sub

Sub mAttachments_PreSave()

    Dim i As Integer
    Dim X As Integer
    Dim iUpper As Integer
    
    iUpper = imContractDocs_Count + imContractDocsAdded_Count - 1
    
    ' Loop through at the listview records.  Lets update the tmContractDocs array based on whether
    ' it is a NEW record, DELETED record or a CHANGED description record.
    
    ' Check for DELETED records or a CHANGED description records.  No changes do nothing.
    For X = 1 To lvcAttachments.ListItems.Count
        For i = 0 To iUpper
            With tmContractDocs(i)
                If tmContractDocs(i).lCode = Val(lvcAttachments.ListItems(X).SubItems(COLUMN_CODE)) Or Trim(tmContractDocs(i).sFileName) = Trim(lvcAttachments.ListItems(X).Text) Then
                    Select Case Val(lvcAttachments.ListItems(X).SubItems(COLUMN_CHANGED))
                        Case CHANGETYPE_NONE    ' No changes so do no processing
                        Case CHANGETYPE_DELETED ' Attachment was deleted
                            .bChanged = True
                            .bDeleted = True
                            .sDeletedDate = DateValue(gNow)
                            .sDeletedTime = TimeValue(gNow)
                            .sTempPath = ""
                            bgContractAttachments_Changed = True
                        Case CHANGETYPE_DESCRIPTIONCHANGED  'Attachment description was changed
                            .bChanged = True
                            .sDescription = lvcAttachments.ListItems(X).SubItems(COLUMN_DESCRIPTION)
                            bgContractAttachments_Changed = True
                    End Select
                    Exit For
                End If
            End With
        Next i
    Next X

    ' Check for ADDED attachments
    For X = 1 To lvcAttachments.ListItems.Count
        If Val(lvcAttachments.ListItems(X).SubItems(COLUMN_CHANGED)) = CHANGETYPE_NEW Then
            ' Attachment was added"
            Call mAttachments_PreSave_NewAttachments(lvcAttachments.ListItems(X).Text, lvcAttachments.ListItems(X).SubItems(COLUMN_PATH), lvcAttachments.ListItems(X).SubItems(COLUMN_DESCRIPTION))
            imContractDocsAdded_Count = imContractDocsAdded_Count + 1
            
            lvcAttachments.ListItems(X).SubItems(COLUMN_CHANGED) = 0  '  Presaved it once and we don't want to do it again
            bgContractAttachments_Changed = True
        End If
    Next X
            
End Sub

Public Sub mAttachments_PreSave_NewAttachments(sFileName As String, sPath As String, sDescription As String)

    Static bFisrtPass As Boolean
    Static iRow As Integer
    
    iRow = iRow + 1

    With tmContractDocs(UBound(tgContractDocuments) + iRow - 1)
        .bChanged = True
        .lCode = 0                     'Dan needs this to know it's a new record
        .sTempPath = sPath
        .lCntrNo = lmCntrNo
        .sFileName = sFileName
        .sDescription = sDescription
        .iUrfCode = tgUrf(0).iCode
        .bDeleted = False
        .sEnteredDate = DateValue(gNow)
        .sEnteredTime = TimeValue(gNow)
        .sDeletedDate = "1970-01-01"    'need this: means 'no date'
        .sDeletedTime = "00:00:00"
    End With

End Sub

Sub mAttachments_Remove()

    If lvcAttachments.SelectedItem.SubItems(COLUMN_CHANGED) = CHANGETYPE_DELETED Then
        MsgBox "This attachment is already marked for deletion.  It will be removed when you exit the Contract Attachments screen", vbInformation, "Delete Attachment Issue"
    Else
        If MsgBox("Please confirm to remove this record:  '" + lvcAttachments.SelectedItem.Text + "'", vbYesNoCancel + vbDefaultButton2 + vbQuestion, "Remove Attachment") = vbYes Then
            If lvcAttachments.SelectedItem.SubItems(COLUMN_MODIFYDATE) = "<Not Committed>" Or lvcAttachments.SelectedItem.SubItems(COLUMN_CODE) = 0 Then
                lvcAttachments.ListItems.Remove (lvcAttachments.SelectedItem.Index)
                imContractDocsAdded_Count = imContractDocsAdded_Count - 1
            Else
                lvcAttachments.SelectedItem.SubItems(COLUMN_CHANGED) = CHANGETYPE_DELETED
                lvcAttachments.SelectedItem.SubItems(COLUMN_DESCRIPTION) = "<Marked For Deletion>"
                lvcAttachments.SelectedItem.ForeColor = vbRed
            End If
            mSetRecordCount
        End If
    End If
    
End Sub

Sub Set_Tooltips()
    
    'cmcDone.ToolTipText = "Returns to the previous screen.  Any changes made with attachments will be saved from the previous screen.  If you cancel out of the previous screen then all attachment changes will be discarded."
    'cmcAdd.ToolTipText = "Add NEW attachments.  Acceptable file types are Word (*.docx;*.doc), Excel (*.xlsx;*.xls), PDF (*.pdf) and Email (*.msg).  Multiple files may be selected at one time."
    'cmcRemove.ToolTipText = "Removes the currently selected attachment.  Removed attachments are archived."
    lblContract.ToolTipText = "The contract number that is associated with the current attachments."
    lvcAttachments.ToolTipText = ""
 
End Sub

Private Sub cmcAdd_Click()
    mAttachments_Add
End Sub

Private Sub cmcDone_Click()

    Dim iUpper As Integer
    
    If bmViewOnly = False Then
        iUpper = UBound(tgContractDocuments) + imContractDocsAdded_Count
        ReDim Preserve tmContractDocs(0 To iUpper) As ACFList
        mAttachments_PreSave
    
        tgContractDocuments = tmContractDocs
    End If
    
    Unload Me
    
End Sub

Private Sub cmcRemove_Click()

    mAttachments_Remove
    mSetEditDescription
    
End Sub

Private Sub edcDescription_Change()
    'ChangeDescription
End Sub

Sub ChangeDescription()

    With lvcAttachments.SelectedItem
        If .SubItems(COLUMN_CHANGED) <> CHANGETYPE_DELETED Then
            .SubItems(COLUMN_DESCRIPTION) = edcDescription.Text
            If .SubItems(COLUMN_CHANGED) <> CHANGETYPE_NEW And edcDescription.Text <> edcDescription.Tag Then
                .SubItems(COLUMN_CHANGED) = CHANGETYPE_DESCRIPTIONCHANGED
                bgContractAttachments_Changed = True
            End If
        End If
    End With
    
End Sub

Private Sub edcDescription_KeyPress(KeyAscii As Integer)
    'ChangeDescription
End Sub


Private Sub edcDescription_KeyUp(KeyCode As Integer, Shift As Integer)
    ChangeDescription
End Sub

Private Sub Form_Load()
    mInit
End Sub

Private Sub lvcAttachments_Click()
    mSetEditDescription
End Sub

Private Sub lvcAttachments_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    'sort/reverse clicked column
    lvcAttachments.SortKey = ColumnHeader.Index - 1
    lvcAttachments.SortOrder = Abs(lvcAttachments.SortOrder - 1)
    lvcAttachments.Sorted = True
    
    mSetEditDescription
    
End Sub

Private Sub lvcAttachments_DblClick()

    Dim lResult As Long
    Dim sFileName As String
    Dim sPath As String
    
    sFileName = lvcAttachments.SelectedItem.Text
    If lvcAttachments.SelectedItem.SubItems(COLUMN_CHANGED) = CHANGETYPE_NEW Or lmCntrNo = 0 Then
        sPath = Replace(lvcAttachments.SelectedItem.SubItems(COLUMN_PATH), lvcAttachments.SelectedItem.Text, "")
        lResult = ShellExecute(Me.HWnd, "open", sFileName, "", sPath, 1)
    Else
        sPath = sgContractAttachmentPath & "\"
        lResult = ShellExecute(Me.HWnd, "open", sFileName, "", sPath & lmCntrNo & "\", 1)
    End If
        
    If lResult = 2 Then
        MsgBox "File not found", vbCritical, "Problem"
    End If
    
End Sub

Private Sub mAttachments_Add()
 
    Dim sFilter As String
    Dim i As Integer
    Dim myFiles() As String
    Dim myPath As String
    
    
    sFilter = sgAttachment_Types & "|All Files (*.*)|*.*"
        
    With cdcAttachments
        .DialogTitle = "Select New Contract Attachment"
        .MaxFileSize = 32000
        .CancelError = False
        .DefaultExt = "docx"
        .FilterIndex = 1        '1 = Save and 2 = Open
       ' .InitDir = GetSpecialfolder(CSIDL_PERSONAL)
        .Filter = sFilter
        .fileName = ""
        .flags = cdlOFNAllowMultiselect + cdlOFNExplorer + cdlOFNLongNames
        .ShowOpen
        
        myFiles = Split(.fileName, vbNullChar) 'the Filename returned is delimeted by a null character because we selected the cdlOFNLongNames flag
        
        Select Case UBound(myFiles)
            Case 0 'if only one was selected we are done
                If mbValidFilename(Trim(right(myFiles(0), Len(myFiles(0)) - InStrRev(myFiles(0), "\")))) = True Then
                    lvcAttachments.ListItems.Add , , sSanitizedFilename(Trim(right(myFiles(0), Len(myFiles(0)) - InStrRev(myFiles(0), "\"))))
                    lvcAttachments.ListItems(lvcAttachments.ListItems.Count).SubItems(COLUMN_CODE) = 0
                    lvcAttachments.ListItems(lvcAttachments.ListItems.Count).SubItems(COLUMN_MODIFYDATE) = "<Not Committed>"
                    lvcAttachments.ListItems(lvcAttachments.ListItems.Count).SubItems(COLUMN_USERNAME) = sgUserName
                    lvcAttachments.ListItems(lvcAttachments.ListItems.Count).SubItems(COLUMN_PATH) = myFiles(0)
                    lvcAttachments.ListItems(lvcAttachments.ListItems.Count).SubItems(COLUMN_CHANGED) = CHANGETYPE_NEW
                    imContractDocsAdded_Count = imContractDocsAdded_Count + 1
                End If
            Case Is > 0 'if more than one, we need to loop through it and append the root directory
                For i = 1 To UBound(myFiles)
                    myPath = myFiles(0) & IIF(right(myFiles(0), 1) <> "\", "\", "") & myFiles(i)
                    If mbValidFilename(Trim(right(myPath, Len(myPath) - InStrRev(myPath, "\")))) = True Then
                        lvcAttachments.ListItems.Add , , sSanitizedFilename(Trim(right(myPath, Len(myPath) - InStrRev(myPath, "\"))))
                        lvcAttachments.ListItems(lvcAttachments.ListItems.Count).SubItems(COLUMN_CODE) = 0
                        lvcAttachments.ListItems(lvcAttachments.ListItems.Count).SubItems(COLUMN_MODIFYDATE) = "<Not Committed>"
                        lvcAttachments.ListItems(lvcAttachments.ListItems.Count).SubItems(COLUMN_USERNAME) = sgUserName
                        lvcAttachments.ListItems(lvcAttachments.ListItems.Count).SubItems(COLUMN_PATH) = myPath
                        lvcAttachments.ListItems(lvcAttachments.ListItems.Count).SubItems(COLUMN_CHANGED) = CHANGETYPE_NEW
                        imContractDocsAdded_Count = imContractDocsAdded_Count + 1
                    End If
                Next i
        End Select
    End With

    mSetRecordCount
    gChDrDir
    
End Sub

Public Function sSanitizedFilename(sFileName As String) As String
    'INVALID CHARACTERS
    '<   less than 60
    '>   greater than 62
    ':   colon 58
    '"   double quote 34
    '/   forward slash 47
    '\   backslash 92
    '|   vertical bar or pipe 124
    '?   question mark 63
    '*   asterisk 42
    ''   apostrophe

    Const sInvalidChars As String = ":\/?*<>|'"""
    Dim lThisChar As Long
   
    sSanitizedFilename = sFileName
    'Loop over each invalid character, removing any instances found
    For lThisChar = 1 To Len(sInvalidChars)
        sSanitizedFilename = Replace$(sSanitizedFilename, Mid(sInvalidChars, lThisChar, 1), "")
    Next
    
End Function


Private Function mbValidFilename(sFileName As String) As Boolean

    Dim i As Integer
    
    ' Does filename already exist?
'    For i = 1 To lvcAttachments.ListItems.Count
'        If lvcAttachments.ListItems(i).Text = sFileName Then
'            MsgBox "The selected file (" & sFileName & ") already exists.  Please select a different file.", vbCritical, "Process Failed!"
'            mbValidFilename = False
'            GoTo ex_function
'        End If
'    Next i
    
    If bValidExtension(sFileName) = True Then
        mbValidFilename = True
    Else
        MsgBox "Valid file types are Word (*.docx;*.doc), Excel (*.xlsx;*.xls), PDF (*.pdf) and Email (*.msg).  Please select a different file.", vbCritical, "Invalid File Type Selected"
        mbValidFilename = False
    End If
  
ex_function:

End Function

Private Function bValidExtension(sFileName As String) As Boolean
    'Function returns true or false if it is a valid attachment extension.
    
    Dim iDotLocation As Integer
    
    iDotLocation = InStr(sFileName, ".")
    
    If iDotLocation = 0 Then 'sFileName doesn't have an extension
        bValidExtension = False
    Else                     'sFileName does have an extension, so retrieve it.
        If InStr(1, sgAttachment_Types, right$(sFileName, Len(sFileName) - iDotLocation)) > 1 Then
            bValidExtension = True
        Else
            bValidExtension = False
        End If
    End If
    
End Function

Private Sub tmrColumns_Timer()

    'This is to keep the user from changing the column widths for the visible columns (except last visible).  This is done entirely for
    'cosmestic reasons.  The listview has limitations where there is no event for when columns are resized so this leaves the description edit textbox
    'un-resized when columns are resized.  Yes, there may be a more simple or elegant way to do this but this method worked for now.  :) - JJB
    
    Static bFirstpass As Boolean
    Static iCOLUMNWIDTH_FILENAME As Double
    Static iCOLUMNWIDTH_MODIFYDATE As Double
    Static iCOLUMNWIDTH_DESCRIPTION As Double
    Static iCOLUMNWIDTH_USERNAME As Double
    
    If Not bFirstpass Then
        bFirstpass = True
        iCOLUMNWIDTH_FILENAME = lvcAttachments.ColumnHeaders.item(COLUMN_FILENAME + 1).Width
        iCOLUMNWIDTH_MODIFYDATE = lvcAttachments.ColumnHeaders.item(COLUMN_MODIFYDATE + 1).Width
        iCOLUMNWIDTH_DESCRIPTION = lvcAttachments.ColumnHeaders.item(COLUMN_DESCRIPTION + 1).Width
        iCOLUMNWIDTH_USERNAME = lvcAttachments.ColumnHeaders.item(COLUMN_USERNAME + 1).Width
    End If
    
    If lvcAttachments.ColumnHeaders.item(COLUMN_FILENAME + 1).Width <> iCOLUMNWIDTH_FILENAME Then lvcAttachments.ColumnHeaders.item(COLUMN_FILENAME + 1).Width = iCOLUMNWIDTH_FILENAME
    If lvcAttachments.ColumnHeaders.item(COLUMN_MODIFYDATE + 1).Width <> iCOLUMNWIDTH_MODIFYDATE Then lvcAttachments.ColumnHeaders.item(COLUMN_MODIFYDATE + 1).Width = iCOLUMNWIDTH_MODIFYDATE
    If lvcAttachments.ColumnHeaders.item(COLUMN_DESCRIPTION + 1).Width <> iCOLUMNWIDTH_DESCRIPTION Then lvcAttachments.ColumnHeaders.item(COLUMN_DESCRIPTION + 1).Width = iCOLUMNWIDTH_DESCRIPTION
    If lvcAttachments.ColumnHeaders.item(COLUMN_USERNAME + 1).Width <> iCOLUMNWIDTH_USERNAME Then lvcAttachments.ColumnHeaders.item(COLUMN_USERNAME + 1).Width = iCOLUMNWIDTH_USERNAME
      
End Sub
