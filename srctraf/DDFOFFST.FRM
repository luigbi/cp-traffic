VERSION 5.00
Begin VB.Form DDFOffst 
   Appearance      =   0  'Flat
   Caption         =   "DDF Offset"
   ClientHeight    =   5805
   ClientLeft      =   1740
   ClientTop       =   2145
   ClientWidth     =   6945
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5805
   ScaleWidth      =   6945
   Begin VB.PictureBox pbcGauge 
      Height          =   210
      Left            =   3825
      ScaleHeight     =   150
      ScaleWidth      =   2865
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4470
      Width           =   2925
      Begin VB.Line lncGauge 
         BorderColor     =   &H80000002&
         BorderStyle     =   3  'Dot
         X1              =   60
         X2              =   2790
         Y1              =   75
         Y2              =   75
      End
   End
   Begin VB.PictureBox plcPack 
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
      Height          =   300
      Left            =   60
      ScaleHeight     =   300
      ScaleWidth      =   3495
      TabIndex        =   15
      Top             =   4290
      Visible         =   0   'False
      Width           =   3495
      Begin VB.OptionButton rbcPack 
         Caption         =   "Spy Work"
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
         Height          =   270
         Index           =   0
         Left            =   1335
         TabIndex        =   17
         Top             =   0
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton rbcPack 
         Caption         =   "VBHlp"
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
         Height          =   270
         Index           =   1
         Left            =   2520
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Width           =   885
      End
   End
   Begin VB.PictureBox plcResult 
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
      Height          =   3105
      Left            =   4230
      ScaleHeight     =   3045
      ScaleWidth      =   2265
      TabIndex        =   6
      Top             =   855
      Width           =   2325
      Begin VB.ListBox lbcResult 
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
         Height          =   2955
         Left            =   15
         TabIndex        =   7
         Top             =   45
         Width           =   2220
      End
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Files"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   75
      TabIndex        =   2
      Top             =   3975
      Width           =   1350
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6150
      Top             =   5100
   End
   Begin VB.PictureBox plcChkFix 
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
      Height          =   3120
      Left            =   60
      ScaleHeight     =   3060
      ScaleWidth      =   3150
      TabIndex        =   1
      Top             =   840
      Width           =   3210
      Begin VB.ListBox lbcDDFNames 
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
         Height          =   1590
         Left            =   45
         MultiSelect     =   2  'Extended
         TabIndex        =   13
         Top             =   1440
         Width           =   3075
      End
      Begin VB.DriveListBox cbcDDF 
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
         Left            =   2115
         TabIndex        =   11
         Top             =   60
         Width           =   1020
      End
      Begin VB.DirListBox lbcDDF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   45
         TabIndex        =   10
         Top             =   420
         Width           =   3075
      End
      Begin VB.Label lacDDF 
         Appearance      =   0  'Flat
         Caption         =   "DDF Table Names"
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
         Left            =   105
         TabIndex        =   12
         Top             =   120
         Width           =   1665
      End
   End
   Begin VB.CommandButton cmcChk 
      Appearance      =   0  'Flat
      Caption         =   "&Generate"
      Height          =   285
      Left            =   2130
      TabIndex        =   4
      Top             =   5115
      Width           =   1140
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3930
      TabIndex        =   5
      Top             =   5115
      Width           =   1140
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   30
      Width           =   6435
   End
   Begin VB.Label lacResults 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   30
      TabIndex        =   14
      Top             =   495
      Width           =   6750
   End
   Begin VB.Label lacTotal 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   4155
      TabIndex        =   9
      Top             =   4770
      Width           =   2430
   End
   Begin VB.Label lacResults 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   30
      TabIndex        =   8
      Top             =   300
      Width           =   6750
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2730
      TabIndex        =   3
      Top             =   4140
      Width           =   3960
   End
End
Attribute VB_Name = "DDFOffst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'**********************************************************
'                Btrieve Match MODULE DEFINITIONS
'
'   Created : 4/25/94       By : D. LeVine
'   Modified :              By :
'
'**********************************************************
Option Explicit
Option Compare Text
Dim hmMsg As Integer   'From file hanle
Dim hmFile As Integer           'file handle
Dim hmOffset As Integer
Dim hmPack As Integer
Dim smDDFPath As String
Dim lmActualRecLen As Long
Dim lmBaseRecLen As Long
Dim smFieldValue As String
Dim tmDDFFile() As DDFFILE
Dim tmDDFNames() As DDFNAMES
Dim tmDDFField() As DDFFIELD    'DDF File- Field Names
Dim tmFieldSrchKey1 As DDFFIELD1
Dim tmDDFIndex() As DDFINDEX    'DDF File- Index Names
Dim tmIndexSrchKey0 As DDFINDEX0

'Files to exclude
'10/31/11: Include all files
'Dim smExludeFiles(0 To 6) As String
Dim smExludeFiles(0 To 0) As String

'Module Status Flags
Dim imTerminate As Integer      'True = terminating task, False= OK
Dim imChgMode As Integer        'Change mode status (so change not entered when in change)
Dim imBSMode As Integer         'Backspace flag
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim smNowDate As String   'Todays date
Dim imFirstTime As Integer
Dim imCancelled As Integer
Dim imProcessing As Integer
Dim smFilePath As String
Dim lmTotalNoFiles As Long
Dim lmProcessedNoFiles As Long
Dim imRun As Integer
Dim imExit As Integer
Dim hmDB As Integer
Dim tmRec As BASEREC
Dim tmGetDDFIndex As DDFINDEX
Dim tmGetDDFFile As DDFFILE
Dim tmGetDDFFIELD As DDFFIELD
Dim smTVIDateTime As String



'*******************************************************
'*                                                     *
'*      Procedure Name:mGetField                       *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get field from record           *
'*                                                     *
'*******************************************************
Private Sub mGetField(tlDDFField As DDFFIELD)
'
'   ilRecType: 0= tmRec
'
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilSize As Integer
    Dim ilValue As Integer
    Dim llValue As Long
    Dim ilDyMn As Integer
    Dim ilYear As Integer
    Dim ilHsSec As Integer
    Dim ilMinHr As Integer
    Dim slMoney As String
    'Dim slHex As String
    'Dim slInt2 As INT2
    'Dim slInt4 As INT4
    'Used with ePDNToStr
    'Dim slResult As String * 20
    'Dim ilPos As Integer
    ilIndex = tlDDFField.iOffset + 1
    ilSize = tlDDFField.iSize
    smFieldValue = "Error"
    Select Case Asc(tlDDFField.sDataType)
        Case 0  'String
            smFieldValue = ""
            For ilLoop = 1 To ilSize Step 1
                'If Asc(tmRec.sChar(ilIndex)) = 0 Then
                If (tmRec.sChar(ilIndex)) = 0 Then
                    smFieldValue = smFieldValue & " "
                Else
                    'If Asc(tmRec.sChar(ilIndex)) > 127 Then
                    '    smFieldValue = smFieldValue & Chr(Asc(tmRec.sChar(ilIndex)) - 128)
                    'Else
                        smFieldValue = smFieldValue & Chr(tmRec.sChar(ilIndex))
                    'End If
                End If
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 1  'Integer
            If ilSize = 1 Then
                ilValue = tmRec.sChar(ilIndex) 'Asc(tmRec.sChar(ilIndex))
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 2 Then
                ''If ilRecType = 0 Then
                ''    slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
                ''Else
                ''    slInt2.sInt2 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1)
                ''End If
                ''HMemCpy ilValue, slInt2, 2
                '''ilValue = (256 * CLng(Asc(tmRec.sChar(ilIndex + 1))) + Asc(tmRec.sChar(ilIndex))) And &HFFFF&
                'ilValue = 256 * (tmRec.sChar(ilIndex + 1)) + (tmRec.sChar(ilIndex))
                HMemCpy ilValue, tmRec.sChar(ilIndex), 2
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 4 Then
                ''If ilRecType = 0 Then
                ''    slInt4.sInt4 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1) & tmRec.sChar(ilIndex + 2) & tmRec.sChar(ilIndex + 3)
                ''Else
                ''    slInt4.sInt4 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1) & tgKeyRec.sChar(ilIndex + 2) & tgKeyRec.sChar(ilIndex + 3)
                ''End If
                ''HMemCpy llValue, slInt4, 4
                '''llValue = (256 * CLng(Asc(tmRec.sChar(ilIndex + 3))) + Asc(tmRec.sChar(ilIndex + 2))) And &HFFFF&
                '''llValue = 32768 * llValue + (256 * CLng(Asc(tmRec.sChar(ilIndex + 1))) + Asc(tmRec.sChar(ilIndex))) And &HFFFF&
                'llValue = 256 * CLng(tmRec.sChar(ilIndex + 3)) + tmRec.sChar(ilIndex + 2)
                'llValue = 65536 * llValue + 256 * CLng(tmRec.sChar(ilIndex + 1)) + tmRec.sChar(ilIndex)
                HMemCpy llValue, tmRec.sChar(ilIndex), 4
                smFieldValue = Trim$(Str$(llValue))
            End If
        Case 2  'IEEE Float
        Case 3  'Btrieve Date
            'ilDyMn = 256 * Asc(tmRec.sChar(ilIndex + 1)) + Asc(tmRec.sChar(ilIndex))
            'ilYear = 256 * Asc(tmRec.sChar(ilIndex + 3)) + Asc(tmRec.sChar(ilIndex + 2))
            ilDyMn = 256 * (tmRec.sChar(ilIndex + 1)) + (tmRec.sChar(ilIndex))
            ilYear = 256 * (tmRec.sChar(ilIndex + 3)) + (tmRec.sChar(ilIndex + 2))
            gUnpackDate ilDyMn, ilYear, smFieldValue
        Case 4  'Btrieve Time
            'ilHsSec = 256 * Asc(tmRec.sChar(ilIndex + 1)) + Asc(tmRec.sChar(ilIndex))
            'ilMinHr = 256 * Asc(tmRec.sChar(ilIndex + 3)) + Asc(tmRec.sChar(ilIndex + 2))
            ilHsSec = 256 * (tmRec.sChar(ilIndex + 1)) + (tmRec.sChar(ilIndex))
            ilMinHr = 256 * (tmRec.sChar(ilIndex + 3)) + (tmRec.sChar(ilIndex + 2))
            gUnpackTime ilHsSec, ilMinHr, "A", "1", smFieldValue
        Case 5  'COBOL Decimal COMP-3
            slMoney = ""
            For ilLoop = 1 To ilSize Step 1
                'slMoney = slMoney & tmRec.sChar(ilIndex)
                slMoney = slMoney & Chr(tmRec.sChar(ilIndex))
                ilIndex = ilIndex + 1
            Next ilLoop
            'ePDNToStr slMoney, ilSize, Asc(tlDDFField.sDec), slResult
            'ilPos = InStr(slResult, Chr(0))
            'If ilPos > 0 Then
            '    smFieldValue = Left$(slResult, ilPos - 1)
            'Else
            '    smFieldValue = Trim$(slResult)
            'End If
            gPDNToStr slMoney, Asc(tlDDFField.sDec), smFieldValue
        Case 6  'COBOL Money
            slMoney = ""
            For ilLoop = 1 To ilSize Step 1
                'slMoney = slMoney & tmRec.sChar(ilIndex)
                slMoney = slMoney & Chr(tmRec.sChar(ilIndex))
                ilIndex = ilIndex + 1
            Next ilLoop
            'ePDNToStr slMoney, ilSize, Asc(tlDDFField.sDec), slResult
            'ilPos = InStr(slResult, Chr(0))
            'If ilPos > 0 Then
            '    smFieldValue = Left$(slResult, ilPos - 1)
            'Else
            '    smFieldValue = Trim$(slResult)
            'End If
            gPDNToStr slMoney, Asc(tlDDFField.sDec), smFieldValue
        Case 7  'Logical
            If ilSize = 1 Then
                'ilValue = Asc(tmRec.sChar(ilIndex))
                ilValue = (tmRec.sChar(ilIndex))
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 2 Then
                ''If ilRecType = 0 Then
                ''    slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
                ''Else
                ''    slInt2.sInt2 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1)
                ''End If
                ''HMemCpy ilValue, slInt2, 2
                '''ilValue = (256 * CLng(Asc(tmRec.sChar(ilIndex + 1))) + Asc(tmRec.sChar(ilIndex))) And &HFFFF&
                'ilValue = (256 * (tmRec.sChar(ilIndex + 1))) + (tmRec.sChar(ilIndex))
                HMemCpy ilValue, tmRec.sChar(ilIndex), 2
                smFieldValue = Trim$(Str$(ilValue))
            End If
        Case 8  'COBOL Numeric
        Case 9  'BASIC Float
        Case 10 'Pascal LString
            'ilValue = Asc(tmRec.sChar(ilIndex))
            ilValue = tmRec.sChar(ilIndex)
            smFieldValue = ""
            ilIndex = ilIndex + 1
            For ilLoop = 1 To ilValue Step 1
                'If Asc(tmRec.sChar(ilIndex)) = 0 Then
                If (tmRec.sChar(ilIndex)) = 0 Then
                    smFieldValue = smFieldValue & " "
                Else
                    'smFieldValue = smFieldValue & tmRec.sChar(ilIndex)
                    smFieldValue = smFieldValue & Chr(tmRec.sChar(ilIndex))
                End If
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 11 'C ZString
            smFieldValue = ""
            'Do While Asc(tmRec.sChar(ilIndex)) <> 0
            '    smFieldValue = smFieldValue & tmRec.sChar(ilIndex)
            '    ilIndex = ilIndex + 1
            'Loop
            Do While (tmRec.sChar(ilIndex)) <> 0
                smFieldValue = smFieldValue & Chr(tmRec.sChar(ilIndex))
                ilIndex = ilIndex + 1
            Loop
        Case 12 'Variable Length Note
        Case 13 'LVar (length, Characters, 0, 0); Length does not include zeros
            ''If ilRecType = 0 Then
            ''    slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
            ''Else
            ''    slInt2.sInt2 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1)
            ''End If
            ''HMemCpy ilValue, slInt2, 2
            'ilValue = (256 * (tmRec.sChar(ilIndex + 1))) + (tmRec.sChar(ilIndex))
            HMemCpy ilValue, tmRec.sChar(ilIndex), 2
            smFieldValue = ""
            ilIndex = ilIndex + 2
            For ilLoop = 1 To ilValue Step 1
                'If Asc(tmRec.sChar(ilIndex)) = 0 Then
                '    smFieldValue = smFieldValue & " "
                'Else
                '    smFieldValue = smFieldValue & tmRec.sChar(ilIndex)
                'End If
                If (tmRec.sChar(ilIndex)) = 0 Then
                    smFieldValue = smFieldValue & " "
                Else
                    smFieldValue = smFieldValue & Chr(tmRec.sChar(ilIndex))
                End If
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 14 'Unsigned Binary
            ''If ilRecType = 0 Then
            ''    slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
            ''Else
            ''    slInt2.sInt2 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1)
            ''End If
            ''HMemCpy ilValue, slInt2, 2
            'ilValue = (256 * (tmRec.sChar(ilIndex + 1))) + (tmRec.sChar(ilIndex))
            HMemCpy ilValue, tmRec.sChar(ilIndex), 2
            llValue = gUnsignIntToLong(ilValue)
            smFieldValue = Trim$(Str$(llValue))
        Case 15 'AutoIncrement
            If ilSize = 2 Then
                ''If ilRecType = 0 Then
                ''    slInt2.sInt2 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1)
                ''Else
                ''    slInt2.sInt2 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1)
                ''End If
                ''HMemCpy ilValue, slInt2, 2
                '''ilValue = (256 * CLng(Asc(tmRec.sChar(ilIndex + 1))) + Asc(tmRec.sChar(ilIndex))) And &HFFFF&
                'ilValue = (256 * (tmRec.sChar(ilIndex + 1))) + (tmRec.sChar(ilIndex))
                HMemCpy ilValue, tmRec.sChar(ilIndex), 2
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 4 Then
                ''If ilRecType = 0 Then
                ''    slInt4.sInt4 = tmRec.sChar(ilIndex) & tmRec.sChar(ilIndex + 1) & tmRec.sChar(ilIndex + 2) & tmRec.sChar(ilIndex + 3)
                ''Else
                ''    slInt4.sInt4 = tgKeyRec.sChar(ilIndex) & tgKeyRec.sChar(ilIndex + 1) & tgKeyRec.sChar(ilIndex + 2) & tgKeyRec.sChar(ilIndex + 3)
                ''End If
                ''HMemCpy llValue, slInt4, 4
                '''llValue = (256 * CLng(Asc(tmRec.sChar(ilIndex + 3))) + Asc(tmRec.sChar(ilIndex + 2))) And &HFFFF&
                '''llValue = 32768 * llValue + (256 * CLng(Asc(tmRec.sChar(ilIndex + 1))) + Asc(tmRec.sChar(ilIndex))) And &HFFFF&
                'llValue = 256 * CLng(tmRec.sChar(ilIndex + 3)) + tmRec.sChar(ilIndex + 2)
                'llValue = 65536 * llValue + 256 * CLng(tmRec.sChar(ilIndex + 1)) + tmRec.sChar(ilIndex)
                HMemCpy llValue, tmRec.sChar(ilIndex), 4
                smFieldValue = Trim$(Str$(llValue))
            Else
                smFieldValue = "Error"
            End If
        Case 16 'Bit
        Case 17 'COBOL Numeric STS
    End Select
End Sub


Public Sub mGetIndexFields()
    Dim tlDDFField As DDFFIELD
    
    tlDDFField.iOffset = 0
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFIndex.iFileID = Val(smFieldValue)
    tlDDFField.iOffset = 2
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFIndex.iFieldID = Val(smFieldValue)
    tlDDFField.iOffset = 4
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFIndex.iNumber = Val(smFieldValue)
    tlDDFField.iOffset = 6
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFIndex.iPart = Val(smFieldValue)
    tlDDFField.iOffset = 8
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFIndex.iFlag = Val(smFieldValue)
End Sub

Public Sub mGetFileFields()
    Dim tlDDFField As DDFFIELD
    
    tlDDFField.iOffset = 0
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFFile.iFileID = Val(smFieldValue)
    tlDDFField.iOffset = 2
    tlDDFField.iSize = 20
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFFile.sName = smFieldValue
    tlDDFField.iOffset = 22
    tlDDFField.iSize = 64
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFFile.sLocation = smFieldValue
    tlDDFField.iOffset = 86
    tlDDFField.iSize = 1
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFFile.sFlags = smFieldValue
    'tmGetDDFFile.sLocation = smFieldValue
    tlDDFField.iOffset = 87
    tlDDFField.iSize = 10
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFFile.sReserved = smFieldValue
End Sub

Public Sub mGetFieldFields()
    Dim tlGDDFField As DDFFIELD
    
    tlGDDFField.iOffset = 0
    tlGDDFField.iSize = 2
    tlGDDFField.sDataType = Chr(1)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField
    tmGetDDFFIELD.iFieldID = Val(smFieldValue)
    tlGDDFField.iOffset = 2
    tlGDDFField.iSize = 2
    tlGDDFField.sDataType = Chr(1)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField
    tmGetDDFFIELD.iFileID = Val(smFieldValue)
    tlGDDFField.iOffset = 4
    tlGDDFField.iSize = 20
    tlGDDFField.sDataType = Chr(0)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField
    tmGetDDFFIELD.sName = smFieldValue
    'tlGDDFField.iOffset = 24
    'tlGDDFField.iSize = 1
    'tlGDDFField.sDataType = Chr(1)
    'tlGDDFField.sDec = ""
    'mGetField tlGDDFField, 0
    'tmGetDDFFIELD.sDataType = smFieldValue
    tmGetDDFFIELD.sDataType = Chr(tmRec.sChar(25))
    tlGDDFField.iOffset = 25
    tlGDDFField.iSize = 2
    tlGDDFField.sDataType = Chr(1)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField
    tmGetDDFFIELD.iOffset = Val(smFieldValue)
    tlGDDFField.iOffset = 27
    tlGDDFField.iSize = 2
    tlGDDFField.sDataType = Chr(1)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField
    tmGetDDFFIELD.iSize = Val(smFieldValue)
    tlGDDFField.iOffset = 29
    tlGDDFField.iSize = 1
    tlGDDFField.sDataType = Chr(0)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField
    tmGetDDFFIELD.sDec = smFieldValue
    tlGDDFField.iOffset = 30
    tlGDDFField.iSize = 2
    tlGDDFField.sDataType = Chr(1)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField
    tmGetDDFFIELD.iFlags = Val(smFieldValue)
End Sub

Private Sub cbcDDF_Change()
    lbcDDF.Path = cbcDDF.Drive
End Sub
Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of Coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcDDFNames.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcDDFNames.hwnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    
    End If
End Sub
Private Sub cmcCancel_Click()
    If imProcessing Then
        imCancelled = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    'gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcChk_Click()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim slName As String
    Dim slToName As String
    Dim slToOffsetFile As String
    Dim slToPackFile As String
    Dim llPercent As Long
    Dim slError As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slChar As String * 1
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilPos As Integer
    Dim slOldName As String
    Dim ilCount As Integer
    Dim llNoRead As Long
    Dim slFileName As String
    Dim ilField As Integer
    Dim slStr As String
    Dim slDDFFile As String
    Dim slDDFDateTime As String
    Screen.MousePointer = vbHourglass
    imCancelled = False
    ilRet = 0
    On Error GoTo cmcChkErr:
    smFilePath = lbcDDF.Path & "\"
    slToOffsetFile = smFilePath & "DDFOffSt.csi"
    slDateTime = FileDateTime(slToOffsetFile)
    If ilRet = 0 Then
        Kill slToOffsetFile
    End If
    On Error GoTo 0
    ilRet = 0
    On Error GoTo cmcChkErr:
    hmOffset = FreeFile
    Open slToOffsetFile For Output As hmOffset
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Open " & slToOffsetFile, vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        cmcCancel.SetFocus
        Exit Sub
    End If
    ilRet = 0
    On Error GoTo cmcChkErr:
    smFilePath = lbcDDF.Path & "\"
    slDDFFile = smFilePath & "Field.DDF"
    slDDFDateTime = FileDateTime(slDDFFile)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Missing " & slDDFFile, vbOkOnly + vbCritical + vbApplicationModal, "DDF File"
        cmcCancel.SetFocus
        Exit Sub
    End If
    
    ilRet = 0
    On Error GoTo cmcChkErr:
    
    slToPackFile = smFilePath & "DDFPack.csi"
    slDateTime = FileDateTime(slToPackFile)
    If ilRet = 0 Then
        Kill slToPackFile
    End If
    On Error GoTo 0
    ilRet = 0
    On Error GoTo cmcChkErr:
    hmPack = FreeFile
    Open slToPackFile For Output As hmPack
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Open " & slToPackFile, vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        cmcCancel.SetFocus
        Exit Sub
    End If
    
    lacResults(0).Caption = "Results Stored into: " & slToOffsetFile
    lacResults(1).Caption = "Results Stored into: " & slToPackFile
    On Error GoTo 0
    slStr = Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmOffset, "'DDF Offset " & slStr
    If smTVIDateTime <> "" Then
        Print #hmOffset, "'DDF Date " & smTVIDateTime
    Else
        Print #hmOffset, "'DDF Date " & slDDFDateTime
    End If
    Print #hmOffset, ""
    
    Print #hmPack, "'DDF Pack " & slStr
    If smTVIDateTime <> "" Then
        Print #hmPack, "'DDF Date " & smTVIDateTime
    Else
        Print #hmPack, "'DDF Date " & slDDFDateTime
    End If
    Print #hmPack, ""
    lmTotalNoFiles = 0
    For ilLoop = 0 To lbcDDFNames.ListCount - 1 Step 1
        If lbcDDFNames.Selected(ilLoop) Then
            lmTotalNoFiles = lmTotalNoFiles + 1
        End If
    Next ilLoop
    lbcResult.Clear
    imProcessing = True
    lmProcessedNoFiles = 0
    lncGauge.X2 = lncGauge.X1
    lncGauge.Visible = True
    For ilLoop = 0 To lbcDDFNames.ListCount - 1 Step 1
        If lbcDDFNames.Selected(ilLoop) Then
            slName = lbcDDFNames.List(ilLoop)
            lacProcessing.Caption = slName
            ilPos = InStr(1, Trim$(tmDDFNames(ilLoop).tDDFFile.sLocation), ".")
            slFileName = left$(Trim$(tmDDFNames(ilLoop).tDDFFile.sLocation), ilPos - 1)
            mFieldBuild tmDDFNames(ilLoop).tDDFFile.iFileID, tmDDFField()
            For ilField = LBound(tmDDFField) To UBound(tmDDFField) - 1 Step 1
                Print #hmOffset, slFileName & ", " & Trim$(tmDDFField(ilField).sName) & "," & Str$(tmDDFField(ilField).iOffset)
            Next ilField
            DoEvents
            If imTerminate Then
                Close hmOffset
                mTerminate
                Exit Sub
            End If
            'Packing
            slStr = ""
            For ilField = LBound(tmDDFField) To UBound(tmDDFField) - 1 Step 1
                slStr = slStr & mGetFieldPack(tmDDFField(ilField))
            Next ilField
            Print #hmPack, slFileName & ", " & slStr
            mIndexBuild tmDDFNames(ilLoop).tDDFFile.iFileID, tmDDFIndex()
            mIndexPrint slFileName, tmDDFIndex(), tmDDFField()
            lmProcessedNoFiles = lmProcessedNoFiles + 1
            llPercent = 100 * (lmProcessedNoFiles / lmTotalNoFiles)
            If llPercent >= 100 Then
                If lmProcessedNoFiles < lmTotalNoFiles Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If llPercent > 0 Then
                lncGauge.X2 = ((pbcGauge.Width - 2 * lncGauge.X1) * llPercent) / 100
                lncGauge.Visible = True
                DoEvents
            End If
        End If
    Next ilLoop
    Print #hmOffset, ""
    'Print #hmOffset, "Completed " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Close #hmOffset
    lacProcessing.Caption = "Processing- Done"
    imProcessing = False
    cmcCancel.Caption = "&Done"
    cmcChk.Enabled = False
    If imExit Then
        mTerminate
    Else
        cmcCancel.SetFocus
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
cmcChkErr:
    ilRet = Err
    Resume Next
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
    DoEvents
    If imRun Then
        cmcChk_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    'Reset used instead of Close to cause # Clients on network to be decrement
    'ilRet = btrReset(hgHlf)
    'btrDestroy hgHlf
    End
End Sub
Private Sub lbcDDF_Change()
    Dim slStr As String
    
    'mFileBuild tmDDFFile()
    slStr = lbcDDF.Path
    If right$(slStr, 1) <> "\" Then
        smDDFPath = lbcDDF.Path & "\"
    Else
        smDDFPath = lbcDDF.Path
    End If
    sgDBPath = smDDFPath
    mFileBuild tmDDFFile()
End Sub
Private Sub lbcResult_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gBtrvErrorMsg                   *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Show error btrieve message      *
'*                                                     *
'*******************************************************
Private Sub mBtrvErrorMsg(ilErrCode As Integer, slMess As String)
'
'   slMess = "Open Arf"
'   gBtrvErrorMsg ilRet, slMess
'   Where:
'       ilErrCode (I)- Error number from btrieve
'       slMess (I) - Message to be added to the btrieve error message
'       MainForm (I)- Name of Form to unload if error exists
'
    Dim slErrMsg As String
    If ilErrCode = BTRV_ERR_NONE Then
        Exit Sub
    End If
    slErrMsg = "Btrieve: "
    Select Case ilErrCode
        Case BTRV_ERR_INVALID_OP    '1
            slErrMsg = slErrMsg & "Invalid Operation (1)"
        Case BTRV_ERR_IO_ERR        '2
            slErrMsg = slErrMsg & "I/O Error (2)"
        Case BTRV_ERR_NOT_OPEN      '3
            slErrMsg = slErrMsg & "File not Opened (3)"
        Case BTRV_ERR_KEY_NOT_FOUND '4
            slErrMsg = slErrMsg & "Key Value not Found (4)"
        Case BTRV_ERR_DUPLICATE_KEY '5
            slErrMsg = slErrMsg & "Duplicate Key Value (5)"
        Case BTRV_ERR_INVALID_KEY   '6
            slErrMsg = slErrMsg & "Invalid Key Number (6)"
        Case BTRV_ERR_DIFF_KEY      '7
            slErrMsg = slErrMsg & "Different Key Number (7)"
        Case BTRV_ERR_INVALID_POS '8
            slErrMsg = slErrMsg & "Invalid Positioning (8)"
        Case BTRV_ERR_END_OF_FILE '9
            slErrMsg = slErrMsg & "End of File (9)"
        Case BTRV_ERR_MOD_KEY_VALUE '10
            slErrMsg = slErrMsg & "Modifiable Key Value Error (10)"
        Case BTRV_ERR_INVALID_FNAME '11
            slErrMsg = slErrMsg & "Invalid File Name (11)"
        Case BTRV_ERR_FILE_NOT_FOUND    '12
            slErrMsg = slErrMsg & "File Not Found (12)"
        Case BTRV_ERR_EXT_FILE      '13
            slErrMsg = slErrMsg & "Extended File Error (13)"
        Case BTRV_ERR_PREIMAGE_OPEN '14
            slErrMsg = slErrMsg & "Per-image Open Error (14)"
        Case BTRV_ERR_PREIMAGE_IO   '15
            slErrMsg = slErrMsg & "Pre-image I/O Error (15)"
        Case BTRV_ERR_EXPANSION '16
            slErrMsg = slErrMsg & "Expansion Error (16)"
        Case BTRV_ERR_CLOSE '17
            slErrMsg = slErrMsg & "Close Error (17)"
        Case BTRV_ERR_DISK_FULL '18
            slErrMsg = slErrMsg & "Disk Full (18)"
        Case BTRV_ERR_UNRECOVERABLE '19
            slErrMsg = slErrMsg & "Unrecoverable Error (19)"
        Case BTRV_ERR_MGR_INACTIVE  '20
            slErrMsg = slErrMsg & "Record Manager Inactive (20)"
        Case BTRV_ERR_KEYBUF_LENGTH '21
            slErrMsg = slErrMsg & "Key Buffer Too Short (21)"
        Case BTRV_ERR_DATABUF_LENGTH    '22
            slErrMsg = slErrMsg & "Data Buffer Length (22)"
        Case BTRV_ERR_POSBLK_LENGTH '23
            slErrMsg = slErrMsg & "Position Block Length (23)"
        Case BTRV_ERR_PAGESIZE  '24
            slErrMsg = slErrMsg & "Page Size Error (24)"
        Case BTRV_ERR_CREATE_IO '25
            slErrMsg = slErrMsg & "Create I/O Error"
        Case BTRV_ERR_NUMKEYS   '26
            slErrMsg = slErrMsg & "Number of Keys (26)"
        Case BTRV_ERR_INVALID_KEYPOS    '27
            slErrMsg = slErrMsg & "Invalid Key Position (27)"
        Case BTRV_ERR_REC_LENGTH    '28
            slErrMsg = slErrMsg & "Invalid Record Length (28)"
        Case BTRV_ERR_KEY_LENGTH    '29
            slErrMsg = slErrMsg & "Invalid Key Length (29)"
        Case BTRV_ERR_NOT_BTRV_FILE '30
            slErrMsg = slErrMsg & "Not a Btrieve File (30)"
        Case BTRV_ERR_ALREADY_EXTD  '31
            slErrMsg = slErrMsg & "File Already Extended (31)"
        Case BTRV_ERR_EXTD_IO   '32
            slErrMsg = slErrMsg & "Extend I/O Error (32)"
        Case BTRV_ERR_INVALID_EXT_NAME  '34
            slErrMsg = slErrMsg & "Invalid Extension Name (34)"
        Case BTRV_ERR_DIRECTORY '35
            slErrMsg = slErrMsg & "Directory Error (35)"
        Case BTRV_ERR_TRANSACTION   '36
            slErrMsg = slErrMsg & "Transaction Error"
        Case BTRV_ERR_TRANS_ACTIVE  '37
            slErrMsg = slErrMsg & "Transaction in Active (37)"
        Case BTRV_ERR_TRANS_FILE_IO '38
            slErrMsg = slErrMsg & "Transaction Control File I/O Error (38)"
        Case BTRV_ERR_END_ABORT_TRANS   '39
            slErrMsg = slErrMsg & "End/Abort Transaction Error (39)"
        Case BTRV_ERR_TRANS_MAX_FILES   '40
            slErrMsg = slErrMsg & "Transaction Max Files (40)"
        Case BTRV_ERR_OP_NOT_ALLOWED    '41
            slErrMsg = slErrMsg & "Operation Not allowed (41)"
        Case BTRV_ERR_ACCEL_ACCESS  '42
            slErrMsg = slErrMsg & "Incomplete Accelerated Access (42)"
        Case BTRV_ERR_INVALID_REC_ADDR  '43
            slErrMsg = slErrMsg & "Invalid Record Address (43)"
        Case BTRV_ERR_NULL_KEY_PATH '44
            slErrMsg = slErrMsg & "Null Key Path (44)"
        Case BTRV_ERR_INCON_KEY_FLAGS   '45
            slErrMsg = slErrMsg & "Inconsistent Key Flags (45)"
        Case BTRV_ERR_ACCESS_DENIED '46
            slErrMsg = slErrMsg & "Access to File Denied (46)"
        Case BTRV_ERR_MAX_OPEN_FILES    '47
            slErrMsg = slErrMsg & "Maximum Open Files (47)"
        Case BTRV_ERR_INVALID_ALT_SEQ   '48
            slErrMsg = slErrMsg & "Invalid Alternate Sequence Definition (48)"
        Case BTRV_ERR_KEY_TYPE  '49
            slErrMsg = slErrMsg & "Key Type Error"
        Case BTRV_ERR_OWNER_SET '50
            slErrMsg = slErrMsg & "Owner Already Set (50)"
        Case BTRV_ERR_INVALID_OWNER '51
            slErrMsg = slErrMsg & "Invalid Owner (51)"
        Case BTRV_ERR_WRITE_CACHE   '52
            slErrMsg = slErrMsg & "Error Writing Cache (52)"
        Case BTRV_ERR_INVALID_INTF  '53
            slErrMsg = slErrMsg & "Invalid Interface (53)"
        Case BTRV_ERR_VARIABLE_PAGE '54
            slErrMsg = slErrMsg & "Variable Page Error (54)"
        Case 55
            slErrMsg = slErrMsg & "Autoincrement Error (55)"
        Case BTRV_ERR_INCOMPLT_INDEX    '56
            slErrMsg = slErrMsg & "Incomplete Index"
        Case BTRV_ERR_EXPAND_MEM    '57
            slErrMsg = slErrMsg & "Expanded Memory Error (57)"
        Case BTRV_ERR_COMPBUF_SIZE  '58
            slErrMsg = slErrMsg & "Compression Buffer Too Short (58)"
        Case BTRV_ERR_FILE_EXISTS   '59
            slErrMsg = slErrMsg & "File Already Exists (59)"
        Case BTRV_ERR_REJECT_COUNT  '60
            slErrMsg = slErrMsg & "Reject Count Reached (60)"
        Case BTRV_ERR_WORK_SPACE_SIZE   '61
            slErrMsg = slErrMsg & "Work Space Too Small (61)"
        Case BTRV_ERR_INCORRECT_DESCP   '62
            slErrMsg = slErrMsg & "Incorrect Description (62)"
        Case BTRV_ERR_INVALID_EXTDBUF   '63
            slErrMsg = slErrMsg & "Invalid Extended Insert Buffer (63)"
        Case BTRV_ERR_FILTER_LIMIT  '64
            slErrMsg = slErrMsg & "Filter Limit Reached (64)"
        Case BTRV_ERR_INCOR_FLD_OFFSET  '65
            slErrMsg = slErrMsg & "Incorrect Field Offset (65)"
        Case BTRV_ERR_AUTO_TRANS_ABORT  '74
            slErrMsg = slErrMsg & "Automatic Transaction Abort (74)"
        Case BTRV_ERR_DEADLOCK_DETECT   '78
            slErrMsg = slErrMsg & "Deadlock Detected (78)"
        Case BTRV_ERR_CONFLICT  '80
            slErrMsg = slErrMsg & "Conflict Error (80)"
        Case BTRV_ERR_LOCK  '81
            slErrMsg = slErrMsg & "Lock Error (81)"
        Case BTRV_ERR_LOST_POS  '82
            slErrMsg = slErrMsg & "Lost Position (82)"
        Case BTRV_ERR_READ_TRANS    '83
            slErrMsg = slErrMsg & "Read Outside Transaction (83)"
        Case BTRV_ERR_REC_LOCKED    '84
            slErrMsg = slErrMsg & "Record Locked (84)"
        Case BTRV_ERR_FILE_LOCKED   '85
            slErrMsg = slErrMsg & "File Locked (85)"
        Case BTRV_ERR_FILE_TBL_FULL '86
            slErrMsg = slErrMsg & "File Table Full (86)"
        Case BTRV_ERR_HNDL_TBL_FULL '87
            slErrMsg = slErrMsg & "Handle Table Full (87)"
        Case BTRV_ERR_INCOM_MODE    '88
            slErrMsg = slErrMsg & "Incompatible Mode Error (88)"
        Case BTRV_ERR_REDIR_DEV_FULL    '90
            slErrMsg = slErrMsg & "Redirection Device Table Full (90)"
        Case BTRV_ERR_SERVER    '91
            slErrMsg = slErrMsg & "Serve Error (91)"
        Case BTRV_ERR_TRANS_TBL_FULL    '92
            slErrMsg = slErrMsg & "Transaction Table Full (92)"
        Case BTRV_ERR_INCOM_LOCK    '93
            slErrMsg = slErrMsg & "Incompatible Lock Type (93)"
        Case BTRV_ERR_PERMISSION    '94
            slErrMsg = slErrMsg & "Permission Error (94)"
        Case BTRV_ERR_SESSION   '95
            slErrMsg = slErrMsg & "Session No Longer Valid (95)"
        Case BTRV_ERR_COMM_ENV  '96
            slErrMsg = slErrMsg & "Communications Environment Error (96)"
        Case BTRV_ERR_DATA_MSGSIZE  '97
            slErrMsg = slErrMsg & "Data Message too Small (97)"
        Case BTRV_ERR_INTERNAL_TRANS    '98
            slErrMsg = slErrMsg & "Internal Transaction Error (98)"
    End Select
    slErrMsg = slErrMsg & ": " & slMess
    MsgBox slErrMsg, vbOkOnly + vbCritical + vbApplicationModal, "Btrieve Error"
    mTerminate
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFieldBuild                     *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list box with field    *
'*                     names                           *
'*                                                     *
'*******************************************************
Private Sub mFieldBuild(ilFileID As Integer, tlDDFField() As DDFFIELD)
'    Dim ilRet As Integer
'    Dim ilUpper As Integer
'    smDDFPath = lbcDDF.Path & "\"
'    Screen.MousePointer = vbHourglass
'    ReDim tlDDFField(0 To 0) As DDFFIELD
'    ilUpper = UBound(tlDDFField)
'    hmFile = CBtrvTable(ONEHANDLE)    'CBtrvObj()
'    ilRet = btrOpen(hmFile, "", smDDFPath & "Field.ddf", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmFile)
'        btrDestroy hmFile
'        Screen.MousePointer = 0 'DEFAULT
'        'mBtrvErrorMsg ilRet, "btrOpen:" & lbcDDF(0).List(lbcDDF(0).ListIndex)
'        Exit Sub
'    End If
'    imActualRecLen = Len(tlDDFField(ilUpper)) 'Len(tmRec)
'    imBaseRecLen = imActualRecLen
'    tmFieldSrchKey1.iFileID = ilFileID
'    ilRet = btrGetGreaterOrEqual(hmFile, tlDDFField(ilUpper), imActualRecLen, tmFieldSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        Screen.MousePointer = 0 'DEFAULT
'        mBtrvErrorMsg ilRet, "btrGetGreaterOrEqual: " & "FIELD.DDF"
'        ilRet = btrClose(hmFile)
'        btrDestroy hmFile
'        Exit Sub
'    End If
'    Do While (ilRet = BTRV_ERR_NONE) And (tlDDFField(ilUpper).iFileID = ilFileID)
'        If imTerminate Then
'            Screen.MousePointer = 0 'DEFAULT
'            ilRet = btrClose(hmFile)
'            btrDestroy hmFile
'            mTerminate
'            Exit Sub
'        End If
'        If Asc(tlDDFField(ilUpper).sDataType) <= 17 Then
'            ilUpper = ilUpper + 1
'            ReDim Preserve tlDDFField(0 To ilUpper) As DDFFIELD
'        End If
'        imActualRecLen = imBaseRecLen
'        ilRet = btrGetNext(hmFile, tlDDFField(ilUpper), imActualRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'    Loop
'    ilRet = btrClose(hmFile)
'    btrDestroy hmFile
'    Screen.MousePointer = 0 'DEFAULT
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim ilMax As Integer
    Dim ilOffset As Integer
    ReDim tlDDFField(0 To 0) As DDFFIELD
    'smDDFPath = lbcPath.Path & "\"
    Screen.MousePointer = vbHourglass
    ReDim tlField(0 To 0) As DDFFIELD
    ilUpper = UBound(tlField)
    hmFile = CBtrvTable()    'CBtrvObj()
    ilRet = btrOpen(hmFile, "", smDDFPath & "Field.ddf", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Screen.MousePointer = 0 'DEFAULT
        'mBtrvErrorMsg ilRet, "btrOpen:" & lbcDDF(0).List(lbcDDF(0).ListIndex)
        Exit Sub
    End If
    lmActualRecLen = Len(tlField(ilUpper)) 'Len(tmRec)
    lmBaseRecLen = lmActualRecLen
    tmFieldSrchKey1.iFileID = ilFileID
    'ilRet = btrGetGreaterOrEqual(hmFile, tlField(ilUpper), lmActualRecLen, tmFieldSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    ilRet = btrGetGreaterOrEqual(hmFile, tmRec, lmActualRecLen, tmFieldSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = 0 'DEFAULT
        mBtrvErrorMsg ilRet, "btrGetGreaterOrEqual: " & "FIELD.DDF"
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Exit Sub
    End If
    Do While (ilRet = BTRV_ERR_NONE)
        If imTerminate Then
            Screen.MousePointer = 0 'DEFAULT
            ilRet = btrClose(hmFile)
            btrDestroy hmFile
            mTerminate
            Exit Sub
        End If
        mGetFieldFields
        tlField(ilUpper) = tmGetDDFFIELD
        If (tlField(ilUpper).iFileID <> ilFileID) Then
            Exit Do
        End If
        If Asc(tlField(ilUpper).sDataType) <= 17 Then
            ilUpper = ilUpper + 1
            ReDim Preserve tlField(0 To ilUpper) As DDFFIELD
        End If
        lmActualRecLen = lmBaseRecLen
        ilRet = btrGetNext(hmFile, tmRec, lmActualRecLen, BTRV_LOCK_NONE)
    Loop
    ilRet = btrClose(hmFile)
    btrDestroy hmFile
    'Sort by offset
    ilMax = 0
    For ilLoop = 0 To UBound(tlField) - 1 Step 1
        If tlField(ilLoop).iOffset > ilMax Then
            ilMax = tlField(ilLoop).iOffset
        End If
    Next ilLoop
    ilUpper = UBound(tlDDFField)
    ilOffset = 0
    Do
        For ilLoop = 0 To UBound(tlField) - 1 Step 1
            If tlField(ilLoop).iOffset = ilOffset Then
                tlDDFField(ilUpper) = tlField(ilLoop)
                ilUpper = ilUpper + 1
                ReDim Preserve tlDDFField(0 To ilUpper) As DDFFIELD
                Exit For
            End If
        Next ilLoop
        ilOffset = ilOffset + 1
    Loop While ilOffset <= ilMax
    Screen.MousePointer = 0 'DEFAULT
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFilePop                        *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build Table of file names       *
'*                                                     *
'*******************************************************
Private Sub mFileBuild(tlDDFFile() As DDFFILE)
'    Dim ilRet As Integer
'    Dim ilUpper As Integer
'    Dim ilLoop As Integer
'    smDDFPath = lbcDDF.Path & "\"
'    Screen.MousePointer = vbHourglass
'    ReDim tlDDFFile(0 To 0) As DDFFILE
'    ReDim tmDDFNames(0 To 0) As DDFNAMES
'    lbcDDFNames.Clear
'    ilUpper = UBound(tlDDFFile)
'    hmFile = CBtrvTable(ONEHANDLE)    'CBtrvObj()
'    ilRet = btrOpen(hmFile, "", smDDFPath & "File.ddf", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmFile)
'        btrDestroy hmFile
'        Screen.MousePointer = 0 'DEFAULT
'        'mBtrvErrorMsg ilRet, "btrOpen:" & lbcDDF(0).List(lbcDDF(0).ListIndex)
'        Exit Sub
'    End If
'    imActualRecLen = Len(tlDDFFile(ilUpper)) 'Len(tmRec)
'    imBaseRecLen = imActualRecLen
'    ilRet = btrGetFirst(hmFile, tlDDFFile(ilUpper), imActualRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'    If ilRet <> BTRV_ERR_NONE Then
'        Screen.MousePointer = 0 'DEFAULT
'        mBtrvErrorMsg ilRet, "btrGetFirst: " & "FILE.DDF"
'        ilRet = btrClose(hmFile)
'        btrDestroy hmFile
'        Exit Sub
'    End If
'    Do While ilRet = BTRV_ERR_NONE
'        If imTerminate Then
'            Screen.MousePointer = 0 'DEFAULT
'            ilRet = btrClose(hmFile)
'            btrDestroy hmFile
'            mTerminate
'            Exit Sub
'        End If
'        If (StrComp(Trim$(tlDDFFile(ilUpper).sName), "X$Field", 1) <> 0) And (StrComp(Trim$(tlDDFFile(ilUpper).sName), "X$File", 1) <> 0) And (StrComp(Trim$(tlDDFFile(ilUpper).sName), "X$Index", 1) <> 0) Then
'            ilUpper = ilUpper + 1
'            ReDim Preserve tlDDFFile(0 To ilUpper) As DDFFILE
'        End If
'        imActualRecLen = imBaseRecLen
'        ilRet = btrGetNext(hmFile, tlDDFFile(ilUpper), imActualRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'    Loop
'    ilRet = btrClose(hmFile)
'    btrDestroy hmFile
'    ReDim tmDDFNames(0 To UBound(tlDDFFile)) As DDFNAMES
'    For ilLoop = 0 To UBound(tlDDFFile) - 1 Step 1
'        tmDDFNames(ilLoop).sKey = tlDDFFile(ilLoop).sName
'        tmDDFNames(ilLoop).tDDFFile = tlDDFFile(ilLoop)
'    Next ilLoop
'    ilUpper = UBound(tmDDFNames)
'    If ilUpper > 0 Then
'        ArraySortTyp fnAV(tmDDFNames(), 0), ilUpper, 0, LenB(tmDDFNames(0)), 0, LenB(tmDDFNames(0).sKey), 0
'    End If
'    For ilLoop = 0 To UBound(tmDDFNames) - 1 Step 1
'        lbcDDFNames.AddItem Trim$(tmDDFNames(ilLoop).tDDFFile.sName)
'    Next ilLoop
'    Screen.MousePointer = 0 'DEFAULT
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilExcludeFile As Integer
    Dim ilTest As Integer
    
    
    'slStr = lbcPath.Path
    'If right$(slStr, 1) <> "\" Then
    '    smDDFPath = lbcPath.Path & "\"
    'Else
    '    smDDFPath = lbcPath.Path
    'End If
    Screen.MousePointer = vbHourglass
    ReDim tlDDFFile(0 To 0) As DDFFILE
    ReDim tmDDFNames(0 To 0) As DDFNAMES
    smTVIDateTime = ""
    lbcDDFNames.Clear
    ilUpper = UBound(tlDDFFile)
    hmFile = CBtrvTable()    'CBtrvObj()
    ilRet = btrOpen(hmFile, "", smDDFPath & "File.ddf", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Screen.MousePointer = 0 'DEFAULT
        'mBtrvErrorMsg ilRet, "btrOpen:" & lbcDDF(0).List(lbcDDF(0).ListIndex)
        Exit Sub
    End If
    lmActualRecLen = Len(tlDDFFile(ilUpper)) 'Len(tmRec)
    lmBaseRecLen = lmActualRecLen
    'ilRet = btrGetFirst(hmFile, tlDDFFile(ilUpper), lmActualRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    ilRet = btrGetFirst(hmFile, tmRec, lmActualRecLen, INDEXKEY0, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = 0 'DEFAULT
        mBtrvErrorMsg ilRet, "btrGetFirst: " & "FILE.DDF"
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Exit Sub
    End If
    Do While ilRet = BTRV_ERR_NONE
        If imTerminate Then
            Screen.MousePointer = 0 'DEFAULT
            ilRet = btrClose(hmFile)
            btrDestroy hmFile
            mTerminate
            Exit Sub
        End If
        mGetFileFields
        tlDDFFile(ilUpper) = tmGetDDFFile
        'If (StrComp(Trim$(tlDDFFile(ilUpper).sName), "X$Field", 1) <> 0) And (StrComp(Trim$(tlDDFFile(ilUpper).sName), "X$File", 1) <> 0) And (StrComp(Trim$(tlDDFFile(ilUpper).sName), "X$Index", 1) <> 0) Then
        ilExcludeFile = False
        If (InStr(1, Trim$(tlDDFFile(ilUpper).sName), "X$", 1) = 1) Then
            ilExcludeFile = True
        ElseIf StrComp(left$(tlDDFFile(ilUpper).sName, 3), "TVI", vbTextCompare) = 0 Then
            smTVIDateTime = Trim$(tlDDFFile(ilUpper).sName)
            ilExcludeFile = True
        ElseIf StrComp(UCase$(left$(tlDDFFile(ilUpper).sName, 4)), "DDF_", vbTextCompare) = 0 Then
            ilExcludeFile = True
        Else
            '10/31/11: Include all files
            'For ilTest = LBound(smExludeFiles) To UBound(smExludeFiles) Step 1
            '    If StrComp(Trim$(tlDDFFile(ilUpper).sLocation), smExludeFiles(ilTest), vbTextCompare) = 0 Then
            '        ilExcludeFile = True
            '        Exit For
            '    End If
            'Next ilTest
        End If
        If Not ilExcludeFile Then
            ilUpper = ilUpper + 1
            ReDim Preserve tlDDFFile(0 To ilUpper) As DDFFILE
        End If
        lmActualRecLen = lmBaseRecLen
        'ilRet = btrGetNext(hmFile, tlDDFFile(ilUpper), lmActualRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilRet = btrGetNext(hmFile, tmRec, lmActualRecLen, BTRV_LOCK_NONE)
    Loop
    ilRet = btrClose(hmFile)
    btrDestroy hmFile
    ReDim tmDDFNames(0 To UBound(tlDDFFile)) As DDFNAMES
    For ilLoop = 0 To UBound(tlDDFFile) - 1 Step 1
        tmDDFNames(ilLoop).sKey = tlDDFFile(ilLoop).sName
        tmDDFNames(ilLoop).tDDFFile = tlDDFFile(ilLoop)
    Next ilLoop
    ilUpper = UBound(tmDDFNames)
    If ilUpper > 0 Then
        ArraySortTyp fnAV(tmDDFNames(), 0), ilUpper, 0, LenB(tmDDFNames(0)), 0, LenB(tmDDFNames(0).sKey), 0
    End If
    For ilLoop = 0 To UBound(tmDDFNames) - 1 Step 1
        lbcDDFNames.AddItem Trim$(tmDDFNames(ilLoop).tDDFFile.sName)
    Next ilLoop
    If ckcAll.Value = vbChecked Then
        ckcAll.Value = vbUnchecked
        ckcAll.Value = vbChecked
    Else
        ckcAll.Value = vbChecked
    End If
    Screen.MousePointer = 0 'DEFAULT
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetFieldPack                   *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Translate field to Pack String  *
'*                                                     *
'*******************************************************
Private Function mGetFieldPack(tlDDFField As DDFFIELD) As String
    Dim slStr As String
    Dim ilSize As Integer
    Dim ilLoop As Integer
    slStr = ""
    ilSize = tlDDFField.iSize
    Select Case Asc(tlDDFField.sDataType)
        Case 0  'String
            If ilSize = 1 Then
                slStr = "b"
            Else
                If rbcPack(0).Value Then
                    slStr = "b" & Trim$(Str$(ilSize))
                Else
                    For ilLoop = 1 To ilSize Step 1
                        slStr = slStr & "b"
                    Next ilLoop
                End If
            End If
        Case 1  'Integer
            If ilSize = 1 Then
                slStr = "b"
            ElseIf ilSize = 2 Then
                slStr = "i"
            ElseIf ilSize = 4 Then
                slStr = "l"
            End If
        Case 2  'IEEE Float
        Case 3  'Btrieve Date
            slStr = "ii"
        Case 4  'Btrieve Time
            slStr = "ii"
        Case 5  'COBOL Decimal COMP-3
            If rbcPack(0).Value Then
                slStr = "b" & Trim$(Str$(ilSize))
            Else
                For ilLoop = 1 To ilSize Step 1
                    slStr = slStr & "b"
                Next ilLoop
            End If
        Case 6  'COBOL Money
            If rbcPack(0).Value Then
                slStr = "b" & Trim$(Str$(ilSize))
            Else
                For ilLoop = 1 To ilSize Step 1
                    slStr = slStr & "b"
                Next ilLoop
            End If
        Case 7  'Logical
            If ilSize = 1 Then
                slStr = "b"
            ElseIf ilSize = 2 Then
                slStr = "i"
            End If
        Case 8  'COBOL Numeric
        Case 9  'BASIC Float
        Case 10 'Pascal LString
        Case 11 'C ZString
            If ilSize = 1 Then
                slStr = "b"
            Else
                If rbcPack(0).Value Then
                    slStr = "b" & Trim$(Str$(ilSize))
                Else
                    For ilLoop = 1 To ilSize Step 1
                        slStr = slStr & "b"
                    Next ilLoop
                End If
            End If
        Case 12 'Variable Length Note
            'Only SSF using this field type
            If rbcPack(0).Value Then
                slStr = "b" & Trim$(Str$(ilSize))
            Else
                For ilLoop = 1 To ilSize Step 1
                    slStr = slStr & "b"
                Next ilLoop
            End If
        Case 13 'LVar (length, Characters, 0, 0); Length does not include zeros
            'Used my comments
            'slStr = "IB" & Trim$(Str$(ilSize - 2))
            If rbcPack(0).Value Then
                slStr = "ib" & Trim$(Str$(ilSize - 2))
            Else
                slStr = "i"
                For ilLoop = 1 To ilSize - 2 Step 1
                    slStr = slStr & "b"
                Next ilLoop
            End If
        Case 14 'Unsigned Binary
            slStr = "i"
        Case 15 'AutoIncrement
            If ilSize = 2 Then
                slStr = "i"
            ElseIf ilSize = 4 Then
                slStr = "l"
            Else
            End If
        Case 16 'Bit
        Case 17 'COBOL Numeric STS
    End Select
    If rbcPack(0).Value Then
        slStr = UCase$(slStr)
    End If
    mGetFieldPack = slStr
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mIndexBuild                     *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build Index table               *
'*                                                     *
'*******************************************************
Private Sub mIndexBuild(ilFileID As Integer, tlDDFIndex() As DDFINDEX)
'    Dim ilRet As Integer
'    Dim ilUpper As Integer
'    smDDFPath = lbcDDF.Path & "\"
'    Screen.MousePointer = vbHourglass
'    ReDim tlDDFIndex(0 To 0) As DDFINDEX
'    ilUpper = UBound(tlDDFIndex)
'    hmFile = CBtrvTable(ONEHANDLE)    'CBtrvObj()
'    ilRet = btrOpen(hmFile, "", smDDFPath & "Index.ddf", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmFile)
'        btrDestroy hmFile
'        Screen.MousePointer = 0 'DEFAULT
'        'mBtrvErrorMsg ilRet, "btrOpen:" & lbcDDF(0).List(lbcDDF(0).ListIndex)
'        Exit Sub
'    End If
'    imActualRecLen = Len(tlDDFIndex(ilUpper)) 'Len(tmRec)
'    imBaseRecLen = imActualRecLen
'    tmIndexSrchKey0.iFileID = ilFileID
'    ilRet = btrGetGreaterOrEqual(hmFile, tlDDFIndex(ilUpper), imActualRecLen, tmIndexSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        Screen.MousePointer = 0 'DEFAULT
'        mBtrvErrorMsg ilRet, "btrGetGreaterOrEqual: " & "FIELD.DDF"
'        ilRet = btrClose(hmFile)
'        btrDestroy hmFile
'        Exit Sub
'    End If
'    Do While (ilRet = BTRV_ERR_NONE) And (tlDDFIndex(ilUpper).iFileID = ilFileID)
'        If imTerminate Then
'            Screen.MousePointer = 0 'DEFAULT
'            ilRet = btrClose(hmFile)
'            btrDestroy hmFile
'            mTerminate
'            Exit Sub
'        End If
'        ilUpper = ilUpper + 1
'        ReDim Preserve tlDDFIndex(0 To ilUpper) As DDFINDEX
'        imActualRecLen = imBaseRecLen
'        ilRet = btrGetNext(hmFile, tlDDFIndex(ilUpper), imActualRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'    Loop
'    ilRet = btrClose(hmFile)
'    btrDestroy hmFile
'    Screen.MousePointer = 0 'DEFAULT
    Dim ilRet As Integer
    Dim ilUpper As Integer
    'smDDFPath = lbcPath.Path & "\"
    Screen.MousePointer = vbHourglass
    ReDim tlDDFIndex(0 To 0) As DDFINDEX
    ilUpper = UBound(tlDDFIndex)
    hmFile = CBtrvTable()    'CBtrvObj()
    ilRet = btrOpen(hmFile, "", smDDFPath & "Index.ddf", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Screen.MousePointer = 0 'DEFAULT
        'mBtrvErrorMsg ilRet, "btrOpen:" & lbcDDF(0).List(lbcDDF(0).ListIndex)
        Exit Sub
    End If
    lmActualRecLen = Len(tlDDFIndex(ilUpper)) 'Len(tmRec)
    lmBaseRecLen = lmActualRecLen
    tmIndexSrchKey0.iFileID = ilFileID
    'ilRet = btrGetGreaterOrEqual(hmFile, tlDDFIndex(ilUpper), lmActualRecLen, tmIndexSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
    ilRet = btrGetGreaterOrEqual(hmFile, tmRec, lmActualRecLen, tmIndexSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = 0 'DEFAULT
        mBtrvErrorMsg ilRet, "btrGetGreaterOrEqual: " & "INDEX.DDF"
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Exit Sub
    End If
    Do While (ilRet = BTRV_ERR_NONE)
        If imTerminate Then
            Screen.MousePointer = 0 'DEFAULT
            ilRet = btrClose(hmFile)
            btrDestroy hmFile
            mTerminate
            Exit Sub
        End If
        mGetIndexFields
        tlDDFIndex(ilUpper) = tmGetDDFIndex
        If (tlDDFIndex(ilUpper).iFileID <> ilFileID) Then
            Exit Do
        End If
        ilUpper = ilUpper + 1
        ReDim Preserve tlDDFIndex(0 To ilUpper) As DDFINDEX
        lmActualRecLen = lmBaseRecLen
        ilRet = btrGetNext(hmFile, tmRec, lmActualRecLen, BTRV_LOCK_NONE)
    Loop
    ilRet = btrClose(hmFile)
    btrDestroy hmFile
    Screen.MousePointer = 0 'DEFAULT
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mIndexPop                       *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list box with Index    *
'*                     names                           *
'*                                                     *
'*******************************************************
Private Sub mIndexPrint(slFileName As String, tlDDFIndex() As DDFINDEX, tlDDFField() As DDFFIELD)
    Dim ilLoop As Integer
    Dim ilMaxIndex As Integer
    Dim ilMaxSeg As Integer
    Dim ilIndex As Integer
    Dim ilSeg As Integer
    Dim slName As String
    Dim ilField As Integer
    ilMaxIndex = 0
    For ilLoop = 0 To UBound(tlDDFIndex) - 1 Step 1
        If tlDDFIndex(ilLoop).iNumber > ilMaxIndex Then
            ilMaxIndex = tlDDFIndex(ilLoop).iNumber
        End If
    Next ilLoop
    For ilIndex = 0 To ilMaxIndex Step 1
        ilMaxSeg = 0
        For ilLoop = 0 To UBound(tlDDFIndex) - 1 Step 1
            If (tlDDFIndex(ilLoop).iNumber = ilIndex) And (tlDDFIndex(ilLoop).iPart > ilMaxSeg) Then
                ilMaxSeg = tlDDFIndex(ilLoop).iPart
            End If
        Next ilLoop
        slName = slFileName & "Key" & Trim$(Str$(ilIndex)) & ","
        For ilSeg = 0 To ilMaxSeg Step 1
            For ilLoop = 0 To UBound(tlDDFIndex) - 1 Step 1
                If (tlDDFIndex(ilLoop).iNumber = ilIndex) And (tlDDFIndex(ilLoop).iPart = ilSeg) Then
                    For ilField = 0 To UBound(tlDDFField) - 1 Step 1
                        If tlDDFField(ilField).iFieldID = tlDDFIndex(ilLoop).iFieldID Then
                            slName = slName & mGetFieldPack(tlDDFField(ilField))
                        End If
                    Next ilField
                End If
            Next ilLoop
        Next ilSeg
        Print #hmPack, slName
    Next ilIndex
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:4/17/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
    
    Dim ilRet As Integer   'Return from btrieve calls
    Dim slDate As String
    Dim slCommand As String
    Dim ilSpace As Integer
    Dim ilPos As Integer
    Dim slStr As String
    Dim slDrive As String
    Dim slPath As String
    
    imTerminate = False
    imSetAll = True
    imAllClicked = False
    imCancelled = False
    imProcessing = False
    imRun = False
    imExit = False
    smDDFPath = ""
    lncGauge.Visible = False
    Screen.MousePointer = vbHourglass
    If Command$ <> "Debug" Then
        'smCmmdFileName = Command$   'Selection File
        slCommand = Command$
        ilPos = InStr(1, slCommand, "/F:", 1)
        If ilPos > 0 Then
            slStr = ""
            ilSpace = InStr(ilPos, slCommand, " ")
            If ilSpace = 0 Then
                slStr = Trim$(Mid$(slCommand, ilPos + 3))
                If ilPos - 1 > 0 Then
                    slCommand = Trim$(left$(slCommand, ilPos - 1))
                Else
                    slCommand = ""
                End If
            Else
                slStr = Mid$(slCommand, ilPos + 3, ilSpace - ilPos - 3)
                If ilPos - 1 > 0 Then
                    slCommand = left$(slCommand, ilPos - 1) & Mid$(slCommand, ilSpace)
                Else
                    slCommand = Mid$(slCommand, ilSpace)
                End If
            End If
            smDDFPath = slStr   'Selection File
        End If
        ilPos = InStr(1, slCommand, "/Exit", 1)
        If ilPos > 0 Then
            imExit = True
        End If
        ilPos = InStr(1, slCommand, "/Run", 1)
        If (ilPos > 0) And (smDDFPath <> "") Then
            imRun = True
        End If
    End If
    'hmDB = CBtrvMngrInit(0, "", "", "", 0) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    hmDB = CBtrvMngrInit(0) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    
    ilPos = InStr(smDDFPath, ":")
    If ilPos > 0 Then
        slDrive = left(smDDFPath, ilPos)
        slPath = Mid$(smDDFPath, ilPos + 1)
        If right$(slPath, 1) = "\" Then
            slPath = left$(slPath, Len(slPath) - 1)
        Else
            smDDFPath = smDDFPath & "\"
        End If
        On Error Resume Next
        cbcDDF.Drive = slDrive
        slStr = lbcDDF.Path
        ilPos = InStr(slStr, ":")
        If ilPos > 0 Then
            slStr = Mid$(slStr, ilPos + 1)
        End If
        If StrComp(slStr, slPath, 1) = 0 Then
            lbcDDF_Change
        Else
            lbcDDF.Path = slPath
        End If
        On Error GoTo 0
    End If
    
    '10/31/11: Include all files
    'smExludeFiles(0) = "aet.mkd"
    'smExludeFiles(1) = "clt.mkd"
    'smExludeFiles(2) = "cmt.mkd"
    'smExludeFiles(3) = "eit.mkd"
    'smExludeFiles(4) = "mgt.mkd"
    'smExludeFiles(5) = "pcr.mkd"
    'smExludeFiles(6) = "tnt.mkd"
    
    DDFOffst.Height = cmcChk.top + 7 * cmcChk.Height / 2
    'gCenterStdAlone BtrMatch
    DDFOffst.Move (Screen.Width - DDFOffst.Width) \ 2, (Screen.Height - DDFOffst.Height) \ 2 + 175 '+ Screen.Height \ 10
    DDFOffst.Show
    
    Screen.MousePointer = vbHourglass
    imChgMode = False
    imBSMode = False
    imFirstTime = True
    smNowDate = Format$(Now, "m/d/yy")
    
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    mTerminate
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:4/17/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: terminate Links                *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
    Dim ilRet As Integer
    Erase tmDDFFile
    Erase tmDDFNames
    Erase tmDDFField
    Erase tmDDFIndex
    Screen.MousePointer = vbDefault
    Unload DDFOffst
    Set DDFOffst = Nothing
    End
End Sub
Private Sub plcPack_Paint()
    plcPack.CurrentX = 0
    plcPack.CurrentY = 0
    plcPack.Print "Pack by"
End Sub
Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    Screen.MousePointer = vbHourglass
    'lbcASCFile.Clear
    Screen.MousePointer = vbDefault
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Generate DDF Offset used by CSI_OS32 and Packing used by CSI_IO32"
End Sub
