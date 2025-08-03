VERSION 5.00
Begin VB.Form frmBIACheck 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   10200
   ControlBox      =   0   'False
   Icon            =   "AffBIACheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frcStructure 
      Enabled         =   0   'False
      Height          =   3240
      Index           =   0
      Left            =   2340
      TabIndex        =   0
      Top             =   135
      Visible         =   0   'False
      Width           =   9540
      Begin VB.CommandButton cmdBIACheck 
         Caption         =   "Yes"
         Height          =   345
         Index           =   0
         Left            =   3495
         TabIndex        =   16
         Top             =   2835
         Width           =   1110
      End
      Begin VB.CommandButton cmdBIACheck 
         Caption         =   "No"
         Height          =   345
         Index           =   1
         Left            =   4920
         TabIndex        =   15
         Top             =   2835
         Width           =   1110
      End
      Begin VB.Label lacBIACheck 
         Alignment       =   2  'Center
         Caption         =   "Sample Import Results shown below for WXXXX-AM:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   18
         Top             =   150
         Width           =   9075
      End
      Begin VB.Label lacStructureMsg 
         Alignment       =   2  'Center
         Caption         =   "Proceed with Importing of all BIA Data?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   885
         TabIndex        =   17
         Top             =   2445
         Width           =   7830
      End
      Begin VB.Label lacCurrent 
         Caption         =   "Current:"
         Height          =   240
         Left            =   1590
         TabIndex        =   14
         Top             =   510
         Width           =   1725
      End
      Begin VB.Label lacChangeTo 
         Caption         =   "Change To:"
         Height          =   240
         Left            =   5700
         TabIndex        =   13
         Top             =   525
         Width           =   1725
      End
      Begin VB.Label lacMarket 
         Caption         =   "Market Name:"
         Height          =   240
         Index           =   0
         Left            =   15
         TabIndex        =   12
         Top             =   825
         Width           =   1215
      End
      Begin VB.Label lacMarket 
         Height          =   270
         Index           =   1
         Left            =   1575
         TabIndex        =   11
         Top             =   825
         Width           =   3810
      End
      Begin VB.Label lacMarket 
         Height          =   270
         Index           =   2
         Left            =   5700
         TabIndex        =   10
         Top             =   825
         Width           =   3810
      End
      Begin VB.Label lacFormat 
         Caption         =   "Format:"
         Height          =   195
         Index           =   0
         Left            =   15
         TabIndex        =   9
         Top             =   2040
         Width           =   960
      End
      Begin VB.Label lacFormat 
         Height          =   195
         Index           =   1
         Left            =   1575
         TabIndex        =   8
         Top             =   2040
         Width           =   3810
      End
      Begin VB.Label lacFormat 
         Height          =   195
         Index           =   2
         Left            =   5700
         TabIndex        =   7
         Top             =   2040
         Width           =   3810
      End
      Begin VB.Label lacRank 
         Caption         =   "Rank:"
         Height          =   225
         Index           =   0
         Left            =   15
         TabIndex        =   6
         Top             =   1230
         Width           =   690
      End
      Begin VB.Label lacRank 
         Height          =   225
         Index           =   1
         Left            =   1575
         TabIndex        =   5
         Top             =   1230
         Width           =   3810
      End
      Begin VB.Label lacRank 
         Height          =   225
         Index           =   2
         Left            =   5700
         TabIndex        =   4
         Top             =   1230
         Width           =   3810
      End
      Begin VB.Label lacOwner 
         Caption         =   "Owner:"
         Height          =   165
         Index           =   0
         Left            =   15
         TabIndex        =   3
         Top             =   1635
         Width           =   1095
      End
      Begin VB.Label lacOwner 
         Height          =   165
         Index           =   1
         Left            =   1575
         TabIndex        =   2
         Top             =   1635
         Width           =   3810
      End
      Begin VB.Label lacOwner 
         Height          =   165
         Index           =   2
         Left            =   5700
         TabIndex        =   1
         Top             =   1635
         Width           =   3810
      End
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8940
      Top             =   2805
   End
   Begin VB.PictureBox ReSize1 
      Height          =   480
      Left            =   8160
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   22
      Top             =   2940
      Width           =   1200
   End
   Begin VB.Frame frcStructure 
      Enabled         =   0   'False
      Height          =   3240
      Index           =   1
      Left            =   330
      TabIndex        =   19
      Top             =   150
      Width           =   9585
      Begin VB.CommandButton cmdBIACheck 
         Caption         =   "Cancel Import"
         Height          =   345
         Index           =   2
         Left            =   3975
         TabIndex        =   21
         Top             =   2805
         Width           =   1815
      End
      Begin VB.TextBox edcMsg 
         Alignment       =   2  'Center
         Height          =   2385
         Left            =   825
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   285
         Width           =   8085
      End
   End
End
Attribute VB_Name = "frmBIACheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private hmFrom As Integer
Private smCallLetters As String
Private smBand As String
Private smCallLettersPlusBand As String
Private smMarketName As String
Private smRank As String
Private smOwnerName As String
Private smFormat As String


Private Sub cmdBIACheck_Click(Index As Integer)

    If cmdBIACheck(0).Value = True Then
        igBIARetStatus = 0
    ElseIf cmdBIACheck(1).Value = True Then
        igBIARetStatus = 1
    Else
        igBIARetStatus = 2
    End If
    
    Unload frmBIACheck
    
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    frcStructure(0).Move 330, 135
    frcStructure(0).BorderStyle = vbBSNone
    frcStructure(1).Move frcStructure(0).Left, frcStructure(0).Top
    frcStructure(1).BorderStyle = vbBSNone
    edcMsg.Text = "Checking BIA Import Form." & sgCRLF & "Comma-Delimited columns should be:" & sgCRLF & " ""Call Letters"", ""Band"", ""Market Name"", Rank, ""Owner Name"", ""Station Format"""
    cmdBIACheck(2).Visible = False
    Me.Width = (Screen.Width) / 2
    Me.Height = (Screen.Height) / 4
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    tmcStart.Enabled = True

End Sub

Private Function mCheckFile()
    Dim slFromFile As String
    Dim slLine As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilShttIndex As Integer
    Dim llMktIndex As Long
    Dim ilStationMatch As Integer
    Dim ilBIAChgd As Integer
    Dim llOwnerIndex As Long
    Dim llFormatIndex As Long
    'Dim slFields(1 To 6) As String
    Dim slFields(0 To 5) As String
    Dim temp_rst As ADODB.Recordset
    
    On Error GoTo mTrapFileOpenError:
    mCheckFile = False
    ilStationMatch = False
    slFromFile = sgBIAFileName
    'ilRet = 0
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        gMsgBox "Unable to open file. Error = " & Trim$(Str$(ilRet))
        Exit Function
    End If
    ilRet = 0
    Line Input #hmFrom, slLine
    If ilRet <> 0 Then
        Close hmFrom
        gMsgBox "Unable to read from file. Error = " & Trim$(Str$(ilRet))
        Exit Function
    End If
    ilBIAChgd = False
    On Error GoTo ErrHand:
    Do While Not EOF(hmFrom)
        ilRet = 0
        Line Input #hmFrom, slLine
        If ilRet <> 0 Then
            gMsgBox "Unable to read from file. Error = " & Trim$(Str$(ilRet))
            Exit Function
        End If
        gParseCDFields slLine, False, slFields()
        'smCallLetters = Trim$(slFields(1))
        smCallLetters = Trim$(slFields(0))
        'smBand = Trim$(slFields(2))
        smBand = Trim$(slFields(1))
        smCallLettersPlusBand = Trim(Replace(smCallLetters, "-", "")) & "-" & Trim(Replace(smBand, "-", ""))
        'smMarketName = Trim$(slFields(3))
        smMarketName = Trim$(slFields(2))
        'smRank = Trim$(slFields(4))
        smRank = Trim$(slFields(3))
        'smOwnerName = Trim$(slFields(5))
        smOwnerName = Trim$(slFields(4))
        'smFormat = Trim$(slFields(6))
        smFormat = Trim$(slFields(5))
        If smMarketName <> "" Then
            ilShttIndex = gBinarySearchStation(smCallLettersPlusBand)
            If ilShttIndex <> -1 Then
                ilStationMatch = True
                If smMarketName <> "" Then
                    llMktIndex = LookupMarketByName(smMarketName)
                    If llMktIndex = -1 Then
                        'New market
                        ilBIAChgd = True
                    Else
                        If tgStationInfo(ilShttIndex).iMktCode <> tgMarketInfo(llMktIndex).iCode Then
                            'Market changed
                            ilBIAChgd = True
                        Else
                            If tgMarketInfo(llMktIndex).iRank <> Val(smRank) Then
                                ilBIAChgd = True
                            End If
                        End If
                    End If
                End If
                If smOwnerName <> "" Then
                    llOwnerIndex = LookupOwnerByName(smOwnerName)
                    If llOwnerIndex = -1 Then
                        ilBIAChgd = True
                    Else
                        If tgOwnerInfo(llOwnerIndex).iCode <> tgStationInfo(ilShttIndex).iOwnerCode Then
                            ilBIAChgd = True
                        End If
                    End If
                End If
                If smFormat <> "" Then
                    llFormatIndex = LookupFormatByName(smFormat)
                    If llFormatIndex = -1 Then
                        ilBIAChgd = True
                    Else
                        If tgFormatInfo(llFormatIndex).iCode <> tgStationInfo(ilShttIndex).iFormatCode Then
                            ilBIAChgd = True
                        End If
                    End If
                End If
                If ilBIAChgd Then
                
                    Close hmFrom
                    
                    lacBIACheck.Caption = "Sample Change, " & smCallLettersPlusBand
                    lacMarket(1).Caption = Trim$(tgStationInfo(ilShttIndex).sMarket)
                    lacMarket(2).Caption = smMarketName
                    lacRank(1).Caption = ""
                    If tgStationInfo(ilShttIndex).iMktCode > 0 Then
                        SQLQuery = "Select mktRank from Mkt where mktCode = " & tgStationInfo(ilShttIndex).iMktCode
                        Set temp_rst = cnn.Execute(SQLQuery)
                        If Not temp_rst.EOF Then
                            lacRank(1).Caption = temp_rst!mktRank
                        End If
                    End If
                    lacRank(2).Caption = smRank
                    
                    lacOwner(1).Caption = ""
                    If tgStationInfo(ilShttIndex).iOwnerCode > 0 Then
                        SQLQuery = "Select arttLastName from Artt where arttCode = " & tgStationInfo(ilShttIndex).iOwnerCode
                        Set temp_rst = cnn.Execute(SQLQuery)
                        If Not temp_rst.EOF Then
                            lacOwner(1).Caption = Trim$(temp_rst!arttLastName)
                        End If
                    End If
                    lacOwner(2).Caption = smOwnerName
                    
                    lacFormat(1).Caption = ""
                    If tgStationInfo(ilShttIndex).iFormatCode > 0 Then
                        SQLQuery = "Select fmtName from FMT_Station_Format where fmtCode = " & tgStationInfo(ilShttIndex).iFormatCode
                        Set temp_rst = cnn.Execute(SQLQuery)
                        If Not temp_rst.EOF Then
                            lacFormat(1).Caption = Trim$(temp_rst!fmtName)
                        End If
                    End If
                    lacFormat(2).Caption = smFormat
                    
                    frcStructure(0).Visible = True
                    frcStructure(0).Enabled = True
                    frcStructure(1).Visible = False
                    frcStructure(1).Enabled = False
                    Screen.MousePointer = vbDefault
                    mCheckFile = True
                    Exit Function
                End If
            End If
        End If
    Loop
    Close hmFrom
    frcStructure(1).Visible = True
    frcStructure(1).Enabled = True
    If ilStationMatch Then
        'No change in data
        cmdBIACheck(2).Caption = "Done"
        edcMsg.Text = "No Data Changes Required"
    Else
        'No Stations matched
        edcMsg.Text = "No Stations Found.  The import form should be a comma-delimited file with the following fields:" & sgCRLF & " ""Call Letters"", ""Band"", ""Market Name"", Rank, ""Owner Name"", ""Station Format"""
    End If
    cmdBIACheck(2).Visible = True
    Screen.MousePointer = vbDefault
    
    mCheckFile = True
    Exit Function

mTrapFileOpenError:
    ilRet = Err.Number
    Resume Next
ErrHand:
    gMsgBox "A general error occured in mCheckFile."
End Function

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub tmcStart_Timer()
    Dim ilRet As Integer
    tmcStart.Enabled = False
    ilRet = mCheckFile()
    If ilRet Then
    Else
        igBIARetStatus = 1
        Unload frmBIACheck
    End If
End Sub

'***************************************************************************
'
'***************************************************************************
Private Function LookupMarketByName(sMarketName As String) As Long
    Dim llLoop As Long
    Dim llIdx As Long

    LookupMarketByName = -1
    On Error GoTo ErrHandler
    llIdx = UBound(tgMarketInfo)
    If llIdx < 1 Then
        Exit Function
    End If
    For llLoop = 0 To llIdx
        If StrComp(Trim(tgMarketInfo(llLoop).sName), Trim(sMarketName), vbTextCompare) = 0 Then
            LookupMarketByName = llLoop
            Exit Function
        End If
    Next
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function LookupOwnerByName(sOwnerName As String) As Long
    Dim llLoop As Long
    Dim llIdx As Long

    LookupOwnerByName = -1
    On Error GoTo ErrHandler
    llIdx = UBound(tgOwnerInfo)
    If llIdx < 1 Then
        Exit Function
    End If
    For llLoop = 0 To llIdx
        If StrComp(Trim(tgOwnerInfo(llLoop).sName), Trim(sOwnerName), vbTextCompare) = 0 Then
            LookupOwnerByName = llLoop
            Exit Function
        End If
    Next
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function LookupFormatByName(sFormatName As String) As Long
    Dim llLoop As Long
    Dim llIdx As Long

    LookupFormatByName = -1
    On Error GoTo ErrHandler
    llIdx = UBound(tgFormatInfo)
    If llIdx < 1 Then
        Exit Function
    End If
    For llLoop = 0 To llIdx
        If StrComp(Trim(tgFormatInfo(llLoop).sName), Trim(sFormatName), vbTextCompare) = 0 Then
            LookupFormatByName = llLoop
            Exit Function
        End If
    Next
ErrHandler:
End Function

