VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form csiArchive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Archive"
   ClientHeight    =   6075
   ClientLeft      =   4245
   ClientTop       =   3630
   ClientWidth     =   7560
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
   Icon            =   "csiArchive.frx":0000
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.PictureBox plcSignon 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Left            =   -15
      Picture         =   "csiArchive.frx":08CA
      ScaleHeight     =   6015
      ScaleWidth      =   7485
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7545
      Begin VB.TextBox edcPayRevHistEnd 
         Height          =   315
         Left            =   4230
         TabIndex        =   12
         Top             =   1995
         Width           =   1050
      End
      Begin VB.TextBox edcCntrEnd 
         Height          =   315
         Left            =   3690
         TabIndex        =   10
         Top             =   1575
         Width           =   1050
      End
      Begin VB.TextBox edcTrafficEnd 
         Height          =   315
         Left            =   4440
         TabIndex        =   8
         Top             =   1155
         Width           =   1050
      End
      Begin VB.TextBox edcTrafficStart 
         Height          =   315
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1155
         Width           =   1050
      End
      Begin VB.TextBox edcAffiliateEnd 
         Height          =   315
         Left            =   4560
         TabIndex        =   5
         Top             =   735
         Width           =   1050
      End
      Begin VB.TextBox edcAffiliateStart 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   735
         Width           =   1050
      End
      Begin VB.Timer tmcStart 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   750
         Top             =   5220
      End
      Begin VB.PictureBox pbcClickFocus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   45
         ScaleHeight     =   165
         ScaleWidth      =   105
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   930
         Width           =   105
      End
      Begin ComctlLib.ProgressBar plcGauge 
         Height          =   225
         Left            =   1560
         TabIndex        =   15
         Top             =   2970
         Visible         =   0   'False
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   397
         _Version        =   327682
         Appearance      =   1
      End
      Begin RichTextLib.RichTextBox edcRTF 
         Height          =   345
         Left            =   1875
         TabIndex        =   28
         Top             =   5055
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   609
         _Version        =   393217
         TextRTF         =   $"csiArchive.frx":A217
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment and Revenue Removed through:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   11
         Top             =   2010
         Width           =   3975
      End
      Begin VB.Label lacCntr 
         BackStyle       =   0  'Transparent
         Caption         =   "Expired Contract Removed through:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   9
         Top             =   1590
         Width           =   3525
      End
      Begin VB.Label lacPayRevHistStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment and Revenue Removal Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   180
         TabIndex        =   24
         Top             =   4425
         Width           =   3825
      End
      Begin VB.Label lacPayRevHistStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "0 of 0 Months Processed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   4050
         TabIndex        =   25
         Top             =   4425
         Visible         =   0   'False
         Width           =   3270
      End
      Begin VB.Label lacCntrStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Expired Contract Removal Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   180
         TabIndex        =   22
         Top             =   4125
         Width           =   3360
      End
      Begin VB.Label lacCntrStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "0 of 0 Months Processed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   3465
         TabIndex        =   23
         Top             =   4110
         Visible         =   0   'False
         Width           =   3555
      End
      Begin VB.Label lacMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   150
         TabIndex        =   14
         Top             =   2625
         Width           =   7170
      End
      Begin VB.Label lacCopyStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "0 of 0 Months Processed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   3540
         TabIndex        =   21
         Top             =   3825
         Visible         =   0   'False
         Width           =   3555
      End
      Begin VB.Label lacTrafficSpotsStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "0 of 0 Months Processed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   4140
         TabIndex        =   19
         Top             =   3525
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label lacAffiliateSpotsStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "0 of 0 Months processed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   4245
         TabIndex        =   17
         Top             =   3225
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Label lacCopyStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Copy Archive and Removal Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   3825
         Width           =   3360
      End
      Begin VB.Label lacTrafficSpotsStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Traffic Spot Archive and Removal Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   3525
         Width           =   3975
      End
      Begin VB.Label lacAffiliateSpotsStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Affiliate Spot Archive and Removal Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   3225
         Width           =   4080
      End
      Begin VB.Label lacLastDatePurged 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Date Purged ran Successfully:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   13
         Top             =   2385
         Width           =   5610
      End
      Begin VB.Label lacStop 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3075
         TabIndex        =   27
         Top             =   5355
         Width           =   1335
      End
      Begin VB.Label lacMin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Minimize"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3075
         TabIndex        =   26
         Top             =   4845
         Width           =   1350
      End
      Begin VB.Label lacTraffic 
         BackStyle       =   0  'Transparent
         Caption         =   "Traffic Purge Dates- Start:                     End:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   6
         Top             =   1170
         Width           =   5595
      End
      Begin VB.Label lacAffiliate 
         BackStyle       =   0  'Transparent
         Caption         =   "Affiliate Purge Dates- Start:                     End:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   750
         Width           =   5595
      End
      Begin VB.Image cmcCSLogo 
         Height          =   510
         Left            =   60
         Top             =   60
         Width           =   3210
      End
      Begin VB.Label imcOutline 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   420
         Left            =   4785
         TabIndex        =   2
         Top             =   4875
         Visible         =   0   'False
         Width           =   1410
      End
   End
End
Attribute VB_Name = "csiArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: csiArchive.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Set Advertiser/Agency Credit code
Option Explicit
Option Compare Text

Private smStartDateTime As String
Private smEndDateTime As String

Dim imClosed As Integer
Dim imCancelled As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim smNowDate As String
Dim lmNowDate As Long
Dim imGenArch As Integer
Dim imUserInput As Integer
Dim smArchivePath As String

Dim smInvEndStdMnth As String
Dim smAffStartDate As String
Dim smTrafStartDate As String
Dim smAffEndDate As String
Dim smTrafEndDate As String
Dim smProjEndDate As String
Dim smPropEndDate As String
Dim smCntrEndDate As String
Dim smPayRevHistEndDate As String
Dim smLastArchRunDate As String
Dim imAffStatus As Integer      '-1:Unable to get Affiliate start date; don't allow archive;
                                '0:Affiliate dates obtained and archive;
                                '1:Affilaite dates obatined but no archive required
Dim imTrafStatus As Integer     '-1:Unable to get Traffic start date; don't allow archive;
                                '0:Traffic dates obtained and archive;
'Dim tmCASGVVefCode() As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
Dim tmLstVefCode() As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
Dim tmCpttVefCode() As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
Dim tmLcfVefCode() As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
Dim tmSsfVefCode() As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
Dim tmVcfVefCode() As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
Dim tmVlfVefCode() As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
Dim tmDlfVefCode() As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
Dim tmEgfVefCode() As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
Dim tmStfVefCode() As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
Dim tmPVefCode() As DBVEHICLE     'Package vehicle
Dim tmRVefCode() As DBVEHICLE     'Rep vehicles
Dim hmSpf As Integer
Dim imSpfRecLen As Integer
Dim tmSpf As SPF
Dim hmSaf As Integer
Dim tmSaf As SAF            'Schedule Attributes record image
Dim tmSafSrchKey As INTKEY0    'Vef key record image
Dim tmSafSrchKey1 As SAFKEY1    'Vef key record image
Dim imSafRecLen As Integer

Dim hmArchive As Integer

Dim tmATTDateInfo As ATTDATEINFO
Dim tmATTDateArchive() As ATTDATEARCHIVE
Dim tmCrfArchiveInfo() As CRFARCHIVEINFO
Dim tmLangMnf() As MNF
Dim lmRetainCifCode() As Long

Dim imBackupDB As Integer   'True or False (NoBackup)
Dim imArchiveMethod As Integer  '1=Fast, 0=Slow (NoDLL)
Dim imNoOutput As Integer   'True or False

Dim tmTeam() As MNF
Dim smTeamTag As String






'
'
'       Remove CCT Comments using the AST's latest date to remove
'       ie. removing ast thru 12/31/07.  Remove any CCT whose action date
'       is  prior to 12/31/07.
Public Function mRemoveCCT(llRetainDate As Long, llCctCount As Long)
    Dim llDate As Long
    Dim llCCT As Long
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim llLastCode As Long
    
    If llRetainDate <= 0 Then
        mRemoveCCT = True
        Exit Function
    End If

    ReDim llCCTCode(0 To 1000) As Long
    ilUpper = 0
    llLastCode = 0
    Do
        DoEvents
        ''go thru entire file, no matching key to use
        'ilRet = btrGetFirst(hgDBCct, tgDBCct, igDBCctRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        tgDBCctSrchKey0.lCode = llLastCode
        ilRet = btrGetGreaterOrEqual(hgDBCct, tgDBCct, igDBCctRecLen, tgDBCctSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get lSdf current record to obtain date
        Do While ilRet = BTRV_ERR_NONE
            llLastCode = tgDBCct.lCode
            gUnpackDateLong tgDBCct.iActionDate(0), tgDBCct.iActionDate(1), llDate
            If (llDate < llRetainDate) And (llDate > 0) Then
                'llCCTCode(UBound(llCCTCode)) = tgDBCct.lCode
                'ReDim Preserve llCCTCode(0 To UBound(llCCTCode) + 1) As Long
                llCCTCode(ilUpper) = tgDBCct.lCode
                ilUpper = ilUpper + 1
                If ilUpper >= UBound(llCCTCode) Then
                    Exit Do
                End If
            End If
            ilRet = btrGetNext(hgDBCct, tgDBCct, igDBCctRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
        DoEvents
        If ilUpper = 0 Then
            Exit Do
        End If
        llCCT = 0
        Do While llCCT < ilUpper    'UBound(llCCTCode)
            tgDBCctSrchKey0.lCode = llCCTCode(llCCT)
            ilRet = btrGetEqual(hgDBCct, tgDBCct, igDBCctRecLen, tgDBCctSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hgDBCct)
                If ilRet = BTRV_ERR_NONE Then
                    llCctCount = llCctCount + 1
                    llCCT = llCCT + 1
                Else
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        llCCT = llCCT + 1
                    End If
                End If
            Else
                llCCT = llCCT + 1
            End If
        Loop
        ilUpper = 0
    Loop
    Erase llCCTCode
    mRemoveCCT = True
End Function







Private Sub edcAffiliateEnd_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcTrafficEnd_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
'    gShowBranner
End Sub
Private Sub Form_Load()
    If App.PrevInstance Then
        MsgBox "Only one copy of csiArchive can be run at a time, sorry", vbInformation + vbOKOnly, "Counterpoint"
        gLogMsg "Second copy of csiArchive path: " & App.Path & " from " & Trim$(gGetComputerName()), "csiArchive.txt", False
        End
    End If
    gSetBkgdMode
    'igBkgdProg = 6
    mInit
    If imUserInput Then
        Me.WindowState = vbNormal
    Else
        Me.WindowState = vbMinimized
    End If
    tmcStart.Enabled = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ilRes As Integer
    Dim ilRet As Integer
    If imClosed = True Then
        Exit Sub
    End If
    Cancel = 0
    tmcStart.Enabled = False
    
    'Doug
    'See if web is locked.  If so transfer is occuring, do you still want stop or not?  I so then unlock and get out.
    'if not locked just show the message below.
    'If imUserInput Then
    '    ilRet = MsgBox("Stop the DB Archive", vbQuestion + vbYesNo, "Stop Service")
    '    If ilRet = vbNo Then
    '        Cancel = 1
    '        imCancelled = False
    '        'tmcRestartTask.Enabled = True
    '        Exit Sub
    '    End If
    'End If
    imClosed = True
    imCancelled = True
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then
        Me.Left = Screen.Width / 2 - Me.Width / 2
        Me.Top = Screen.Height / 2 - Me.Height / 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    Erase tmTeam
    
    'Erase tmCASGVVefCode
    Erase tmLstVefCode
    Erase tmCpttVefCode
    Erase tmLcfVefCode
    Erase tmSsfVefCode
    Erase tmVcfVefCode
    Erase tmVlfVefCode
    Erase tmDlfVefCode
    Erase tmEgfVefCode
    Erase tmStfVefCode
    Erase tmPVefCode
    Erase tmRVefCode
    Erase tmATTDateArchive
    Erase tmCrfArchiveInfo
    Erase tmLangMnf
    Erase lmRetainCifCode
    
    btrDestroy hgDBGhf
    btrDestroy hgDBIhf
    btrDestroy hgDBIsf
    
    btrDestroy hgAuf
    
    btrDestroy hmSaf
    btrDestroy hmSpf
    
    'Affiliate Spot Archive
    btrDestroy hgDBWebl
    btrDestroy hgDBEdf
    btrDestroy hgDBEsf
    btrDestroy hgDBPet
    btrDestroy hgDBVat
    btrDestroy hgDBEpt
    btrDestroy hgDBDat
    btrDestroy hgDBShtt
    btrDestroy hgDBAtt
    btrDestroy hgDBCptt
    btrDestroy hgDBLst
    btrDestroy hgDBAet
    btrDestroy hgDBAst
    btrDestroy hgDBCct      '7-24-09
    
    'Proposal Archive
    btrDestroy hgDBMgf
    btrDestroy hgDBMsf
    btrDestroy hgDBCgf
    btrDestroy hgDBEff
    btrDestroy hgDBVsf
    btrDestroy hgDBCxf
    btrDestroy hgDBSbf
    btrDestroy hgDBCff
    btrDestroy hgDBClf
    btrDestroy hgDBChf
    
    'Traffic Spot Archive Archive
    btrDestroy hgDBSdf
    btrDestroy hgDBSmf
    btrDestroy hgDBStf
    btrDestroy hgDBMtf
    btrDestroy hgDBPsf
    btrDestroy hgDBRsf
    btrDestroy hgDBTzf
    
    'Program Archive
    btrDestroy hgDBLcf
    btrDestroy hgDBSsf
    btrDestroy hgDBVlf
    btrDestroy hgDBVcf
    btrDestroy hgDBDlf
    btrDestroy hgDBEgf
    btrDestroy hgDBLtf
    btrDestroy hgDBLvf
    btrDestroy hgDBLef
    btrDestroy hgDBCef
    btrDestroy hgDBGsf
    
    'Copy Archive
    btrDestroy hgDBSef
    btrDestroy hgDBRaf
    btrDestroy hgDBCrf
    btrDestroy hgDBCnf
    btrDestroy hgDBCaf
    btrDestroy hgDBCif
    btrDestroy hgDBCpf
    btrDestroy hgDBCsf
    btrDestroy hgDBCyf
    
    'Advertiser and Agency Archive
    btrDestroy hgDBAgf
    btrDestroy hgDBAdf
    btrDestroy hgDBPrf
    btrDestroy hgDBPnf
    btrDestroy hgDBRvf
    btrDestroy hgDBPhf
    btrDestroy hgDBCdf
    
    'Blackout Archive
    btrDestroy hgDBBof
    
    'Projection Archive
    btrDestroy hgDBPjf
    
    'Unused
    btrDestroy hgDBSlf
    btrDestroy hgDBRcf
    btrDestroy hgDBRif
    btrDestroy hgDBVef
    btrDestroy hgDBVff
    btrDestroy hgDBScf
    btrDestroy hgDBBsf
    btrDestroy hgDBBvf
    btrDestroy hgDBMcf
    btrDestroy hgDBDnf
    btrDestroy hgDBDrf
    btrDestroy hgDBDpf
    btrDestroy hgDBDef
    btrDestroy hgDBVpf
    btrDestroy hgDBVof
    btrDestroy hgDBVbf
    btrDestroy hgDBVaf
    btrDestroy hgDBPif
    btrDestroy hgDBPvf
    btrDestroy hgDBAuf
    btrDestroy hgDBAlf
    btrDestroy hgDBRqf
    btrDestroy hgDBRff
    Set csiArchive = Nothing
    btrStopAppl
    Set csiArchive = Nothing   'Remove data segment
    End
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
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim llAdfNoRecs As Integer
    Dim llAgfNoRecs As Integer
    Dim slStr As String
    Dim llLoop As Long
    
    imTerminate = False
    
    'Screen.MousePointer = vbHourGlass
    mParseCmmdLine
    If Not gCheckDDFDates() Then
        imTerminate = True
        Exit Sub
    End If
    
    If Not mCheckRecLength() Then
        imTerminate = True
        Exit Sub
    End If
    If imUserInput Then
        lacMin.Caption = "Archive"
        edcAffiliateEnd.Locked = False
        edcTrafficEnd.Locked = False
    Else
        lacMin.Caption = "Minimize"
        edcAffiliateEnd.Locked = True
        edcTrafficEnd.Locked = True
    End If
    imGenArch = False
    imCancelled = False
    imClosed = False
    smNowDate = Format$(Now, "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    
    If imNoOutput Then
        lacAffiliateSpotsStatus(0).Caption = "Affiliate Spot Removal Status:"
        lacTrafficSpotsStatus(0).Caption = "Traffic Spot Removal Status:"
        lacCopyStatus(0).Caption = "Copy Removal Status:"
    End If
    
    'Affiliate Spots
    hgDBAst = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBAst, "", sgDBPath & "Ast.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ast.Mkd)", csiArchive
    On Error GoTo 0
    igDBAstRecLen = Len(tgDBAst)
    hgDBAet = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBAet, "", sgDBPath & "Aet.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Aet.Mkd)", csiArchive
    On Error GoTo 0
    igDBAetRecLen = Len(tgDBAet)
    hgDBLst = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBLst, "", sgDBPath & "Lst.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lst.Mkd)", csiArchive
    On Error GoTo 0
    igDBLstRecLen = Len(tgDBLst)
    hgDBCptt = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBCptt, "", sgDBPath & "Cptt.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cptt.Mkd)", csiArchive
    On Error GoTo 0
    igDBCpttRecLen = Len(tgDBCptt)
    hgDBAtt = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBAtt, "", sgDBPath & "Att.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Att.Mkd)", csiArchive
    On Error GoTo 0
    igDBAttRecLen = Len(tgDBAtt)
    hgDBShtt = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBShtt, "", sgDBPath & "Shtt.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Shtt.Mkd)", csiArchive
    On Error GoTo 0
    igDBShttRecLen = Len(tgDBShtt)
    hgDBDat = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBDat, "", sgDBPath & "Dat.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dat.Mkd)", csiArchive
    On Error GoTo 0
    igDBDatRecLen = Len(tgDBDat)
    
    hgDBEpt = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBEpt, "", sgDBPath & "Ept.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ept.Mkd)", csiArchive
    On Error GoTo 0
    igDBEptRecLen = Len(tgDBEpt)
    
    hgDBPet = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBPet, "", sgDBPath & "Pet.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pet.Mkd)", csiArchive
    On Error GoTo 0
    igDBPetRecLen = Len(tgDBPet)
    
    hgDBVat = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBVat, "", sgDBPath & "Vat.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vat.Mkd)", csiArchive
    On Error GoTo 0
    igDBVatRecLen = Len(tgDBVat)
    
    
    hgDBEsf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBEsf, "", sgDBPath & "Esf.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Esf.Mkd)", csiArchive
    On Error GoTo 0
    igDBEsfRecLen = Len(tgDBEsf)
    hgDBEdf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBEdf, "", sgDBPath & "Edf.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Edf.Mkd)", csiArchive
    On Error GoTo 0
    igDBEdfRecLen = Len(tgDBEdf)
    hgDBWebl = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBWebl, "", sgDBPath & "Webl.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Webl.Mkd)", csiArchive
    On Error GoTo 0
    igDBWeblRecLen = Len(tgDBWebl)
    
    '7-24-09 CTT comments
    hgDBCct = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBCct, "", sgDBPath & "CCT.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CCT.Mkd)", csiArchive
    On Error GoTo 0
    igDBCctRecLen = Len(tgDBCct)
    
    'Traffic spot Archive
    hgDBSdf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", csiArchive
    On Error GoTo 0
    igDBSdfRecLen = Len(tgDBSdf)
    hgDBSmf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf.Btr)", csiArchive
    On Error GoTo 0
    igDBSmfRecLen = Len(tgDBSmf)
    hgDBStf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBStf, "", sgDBPath & "Stf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Stf.Btr)", csiArchive
    On Error GoTo 0
    igDBStfRecLen = Len(tgDBStf)
    hgDBMtf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBMtf, "", sgDBPath & "Mtf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mtf.Btr)", csiArchive
    On Error GoTo 0
    igDBMtfRecLen = Len(tgDBMtf)
    hgDBPsf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBPsf, "", sgDBPath & "Psf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Psf.Btr)", csiArchive
    On Error GoTo 0
    igDBPsfRecLen = Len(tgDBPsf)
    hgDBRsf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBRsf, "", sgDBPath & "Rsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rsf.Btr)", csiArchive
    On Error GoTo 0
    igDBRsfRecLen = Len(tgDBRsf)
    hgDBTzf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Tzf.Btr)", csiArchive
    On Error GoTo 0
    igDBTzfRecLen = Len(tgDBTzf)
    
    'Proposal Archive
    hgDBChf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", csiArchive
    On Error GoTo 0
    igDBChfRecLen = Len(tgDBChf)
    hgDBClf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", csiArchive
    On Error GoTo 0
    igDBClfRecLen = Len(tgDBClf)
    hgDBCff = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", csiArchive
    On Error GoTo 0
    igDBCffRecLen = Len(tgDBCff)
    hgDBCxf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cxf.Btr)", csiArchive
    On Error GoTo 0
    igDBCxfRecLen = Len(tgDBCxf)
    hgDBVsf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vsf.Btr)", csiArchive
    On Error GoTo 0
    igDBVsfRecLen = Len(tgDBVsf)
    hgDBSbf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sbf.Btr)", csiArchive
    On Error GoTo 0
    igDBSbfRecLen = Len(tgDBSbf)
    hgDBEff = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBEff, "", sgDBPath & "Eff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Eff.Btr)", csiArchive
    On Error GoTo 0
    igDBEffRecLen = Len(tgDBEff)
    hgDBCgf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBCgf, "", sgDBPath & "Cgf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cgf.Btr)", csiArchive
    On Error GoTo 0
    igDBCgfRecLen = Len(tgDBCgf)
    hgDBMsf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBMsf, "", sgDBPath & "Msf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Msf.Btr)", csiArchive
    On Error GoTo 0
    igDBMsfRecLen = Len(tgDBMsf)
    hgDBMgf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBMgf, "", sgDBPath & "Mgf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mgf.Btr)", csiArchive
    On Error GoTo 0
    igDBMgfRecLen = Len(tgDBMgf)
        
    'Programming Archive
    hgDBLcf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lcf.Btr)", csiArchive
    On Error GoTo 0
    igDBLcfRecLen = Len(tgDBLcf)
    hgDBSsf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ssf.Btr)", csiArchive
    On Error GoTo 0
    igDBSsfRecLen = Len(tgDBSsf)
    hgDBVlf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBVlf, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vlf.Btr)", csiArchive
    On Error GoTo 0
    igDBVlfRecLen = Len(tgDBVlf)
    hgDBVcf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBVcf, "", sgDBPath & "Vcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vcf.Btr)", csiArchive
    On Error GoTo 0
    igDBVcfRecLen = Len(tgDBVcf)
    hgDBDlf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBDlf, "", sgDBPath & "Dlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dlf.Btr)", csiArchive
    On Error GoTo 0
    igDBDlfRecLen = Len(tgDBDlf)
    hgDBEgf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBEgf, "", sgDBPath & "Egf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Egf.Btr)", csiArchive
    On Error GoTo 0
    igDBEgfRecLen = Len(tgDBEgf)
    hgDBLtf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBLtf, "", sgDBPath & "Ltf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ltf.Btr)", csiArchive
    On Error GoTo 0
    igDBLtfRecLen = Len(tgDBLtf)
    hgDBLvf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBLvf, "", sgDBPath & "Lvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lvf.Btr)", csiArchive
    On Error GoTo 0
    igDBLvfRecLen = Len(tgDBLvf)
    hgDBLef = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBLef, "", sgDBPath & "Lef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lef.Btr)", csiArchive
    On Error GoTo 0
    igDBLefRecLen = Len(tgDBLef)
    hgDBCef = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cef.Btr)", csiArchive
    On Error GoTo 0
    igDBCefRecLen = Len(tgDBCef)
    hgDBGhf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBGhf, "", sgDBPath & "ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ghf.Btr)", csiArchive
    On Error GoTo 0
    igDBGhfRecLen = Len(tgDBGhf)
    hgDBGsf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Gsf.Btr)", csiArchive
    On Error GoTo 0
    igDBGsfRecLen = Len(tgDBGsf)
    hgDBIhf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBIhf, "", sgDBPath & "Ihf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ihf.Btr)", csiArchive
    On Error GoTo 0
    igDBIhfRecLen = Len(tgDBIhf)
    hgDBIsf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBIsf, "", sgDBPath & "Isf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Isf.Btr)", csiArchive
    On Error GoTo 0
    igDBIsfRecLen = Len(tgDBIsf)
    
    'Copy Archive
    hgDBCrf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Crf.Btr)", csiArchive
    On Error GoTo 0
    igDBCrfRecLen = Len(tgDBCrf)
    hgDBCnf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBCnf, "", sgDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cnf.Btr)", csiArchive
    On Error GoTo 0
    igDBCnfRecLen = Len(tgDBCnf)
    hgDBCaf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBCaf, "", sgDBPath & "Caf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Caf.Btr)", csiArchive
    On Error GoTo 0
    igDBCafRecLen = Len(tgDBCaf)
    hgDBCif = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", csiArchive
    On Error GoTo 0
    igDBCifRecLen = Len(tgDBCif)
    hgDBCpf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cpf.Btr)", csiArchive
    On Error GoTo 0
    igDBCpfRecLen = Len(tgDBCpf)
    hgDBCyf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBCyf, "", sgDBPath & "Cyf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cyf.Btr)", csiArchive
    On Error GoTo 0
    igDBCyfRecLen = Len(tgDBCyf)
    hgDBCsf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBCsf, "", sgDBPath & "Csf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Csf.Btr)", csiArchive
    On Error GoTo 0
    igDBCsfRecLen = Len(tgDBCsf)
    hgDBRaf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgDBRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Raf.Btr)", csiArchive
    On Error GoTo 0
    igDBRafRecLen = Len(tgDBRaf)
    hgDBSef = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgDBSef, "", sgDBPath & "Sef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sef.Btr)", csiArchive
    On Error GoTo 0
    igDBSefRecLen = Len(tgDBSef)
    
    'Advertiser and Agency Archive
    hgDBAdf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", csiArchive
    On Error GoTo 0
    igDBAdfRecLen = Len(tgDBAdf)
    hgDBAgf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Agf.Btr)", csiArchive
    On Error GoTo 0
    igDBAgfRecLen = Len(tgDBAgf)
    hgDBPrf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Prf.Btr)", csiArchive
    On Error GoTo 0
    igDBPrfRecLen = Len(tgDBPrf)
    hgDBPnf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBPnf, "", sgDBPath & "Pnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pnf.Btr)", csiArchive
    On Error GoTo 0
    igDBPnfRecLen = Len(tgDBPnf)
    hgDBRvf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rvf.Btr)", csiArchive
    On Error GoTo 0
    igDBRvfRecLen = Len(tgDBRvf)
    hgDBPhf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBPhf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Phf.Btr)", csiArchive
    On Error GoTo 0
    hgDBCdf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBCdf, "", sgDBPath & "Cdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cdf.Btr)", csiArchive
    On Error GoTo 0
    igDBCdfRecLen = Len(tgDBCdf)
    
    hgDBBof = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBBof, "", sgDBPath & "Bof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Bof.Btr)", csiArchive
    On Error GoTo 0
    igDBBofRecLen = Len(tgDBBof)
    
    hgDBPjf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBPjf, "", sgDBPath & "Pjf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pjf.Btr)", csiArchive
    On Error GoTo 0
    igDBPjfRecLen = Len(tgDBPjf)
    
    'Unused
    hgDBSlf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Slf.Btr)", csiArchive
    On Error GoTo 0
    igDBSlfRecLen = Len(tgDBSlf)
    hgDBRcf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBRcf, "", sgDBPath & "Rcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rcf.Btr)", csiArchive
    On Error GoTo 0
    igDBRcfRecLen = Len(tgDBRcf)
    hgDBRif = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBRif, "", sgDBPath & "Rif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rif.Btr)", csiArchive
    On Error GoTo 0
    igDBRifRecLen = Len(tgDBRif)
    hgDBVef = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", csiArchive
    On Error GoTo 0
    igDBVefRecLen = Len(tgDBVef)
    hgDBVff = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBVff, "", sgDBPath & "Vff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vff.Btr)", csiArchive
    On Error GoTo 0
    igDBVffRecLen = Len(tgDBVff)
    hgDBScf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBScf, "", sgDBPath & "Scf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Scf.Btr)", csiArchive
    On Error GoTo 0
    igDBScfRecLen = Len(tgDBScf)
    hgDBBsf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBBsf, "", sgDBPath & "Bsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Bsf.Btr)", csiArchive
    On Error GoTo 0
    igDBBsfRecLen = Len(tgDBBsf)
    hgDBBvf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBBvf, "", sgDBPath & "Bvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Bvf.Btr)", csiArchive
    On Error GoTo 0
    igDBBvfRecLen = Len(tgDBBvf)
    hgDBMcf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mcf.Btr)", csiArchive
    On Error GoTo 0
    igDBMcfRecLen = Len(tgDBMcf)
    hgDBDnf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dnf.Btr)", csiArchive
    On Error GoTo 0
    igDBDnfRecLen = Len(tgDBDnf)
    hgDBDrf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Drf.Btr)", csiArchive
    On Error GoTo 0
    igDBDrfRecLen = Len(tgDBDrf)
    hgDBDpf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dpf.Btr)", csiArchive
    On Error GoTo 0
    igDBDpfRecLen = Len(tgDBDpf)
    hgDBDef = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Def.Btr)", csiArchive
    On Error GoTo 0
    igDBDefRecLen = Len(tgDBDef)
    
    hgDBVpf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBVpf, "", sgDBPath & "Vpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vpf.Btr)", csiArchive
    On Error GoTo 0
    igDBVpfRecLen = Len(tgDBVpf)
    hgDBVof = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBVof, "", sgDBPath & "Vof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vof.Btr)", csiArchive
    On Error GoTo 0
    igDBVofRecLen = Len(tgDBVof)
    hgDBVbf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBVbf, "", sgDBPath & "Vbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vbf.Btr)", csiArchive
    On Error GoTo 0
    igDBVbfRecLen = Len(tgDBVbf)
    hgDBVaf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBVaf, "", sgDBPath & "Vaf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vaf.Btr)", csiArchive
    On Error GoTo 0
    igDBVafRecLen = Len(tgDBVaf)
    hgDBPif = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBPif, "", sgDBPath & "Pif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pif.Btr)", csiArchive
    On Error GoTo 0
    igDBPifRecLen = Len(tgDBPif)
    hgDBPvf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBPvf, "", sgDBPath & "Pvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pvf.Btr)", csiArchive
    On Error GoTo 0
    igDBPvfRecLen = Len(tgDBPvf)
    hgDBAlf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBAlf, "", sgDBPath & "Alf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Alf.Btr)", csiArchive
    On Error GoTo 0
    igDBAlfRecLen = Len(tgDBAlf)
    hgDBAuf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBAuf, "", sgDBPath & "Auf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Auf.Btr)", csiArchive
    On Error GoTo 0
    igDBAufRecLen = Len(tgDBAuf)
    
    hgDBRqf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBRqf, "", sgDBPath & "Rqf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rqf.Btr)", csiArchive
    On Error GoTo 0
    igDBRqfRecLen = Len(tgDBRqf)
    hgDBRff = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgDBRff, "", sgDBPath & "Rff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rff.Btr)", csiArchive
    On Error GoTo 0
    igDBRffRecLen = Len(tgDBRff)
    
    
    mSetAffStatus
    
    ilRet = gObtainVef()        'Stored into tgMVef, use gBinarySearchVef to find match by code
    'mSplitVefCodes
    ilRet = gObtainAdvt()       'Stored into tgCommAdf, use gBinarySearchAdf to find match by code
    ilRet = gObtainAgency()     'Stored into tgCommAgf, use gBinarySearchAgf to find match by code
    ilRet = gObtainAvail()      'Stored into tgAvailAnf, Use mBinarySearchAvail to find match by code
    ilRet = gObtainMCF()        'Stored into tgMCF, Use mBinarySearchMcf to find match by code
    ilRet = gObtainUrf()        'Stored into tgPopUrf, Use mBinarySearchUrf to find match by code
    ilRet = btrGetLast(hgDBCif, tgDBCif, igDBCifRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        ReDim lmRetainCifCode(0 To tgDBCif.lCode) As Long
    Else
        ReDim lmRetainCifCode(0 To 0) As Long
    End If
    For llLoop = 0 To UBound(lmRetainCifCode) Step 1
        lmRetainCifCode(llLoop) = 0
    Next llLoop
    'Missed Reason
    slStr = ""
    'ReDim tgMRMnf(1 To 1) As MNF
    ReDim tgMRMnf(0 To 0) As MNF
    ilRet = gObtainMnfForType("M", slStr, tgMRMnf())
    slStr = ""
    'ReDim tmLangMnf(1 To 1) As MNF
    ReDim tmLangMnf(0 To 0) As MNF
    ilRet = gObtainMnfForType("L", slStr, tmLangMnf())

    hmSpf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSpf, "", sgDBPath & "Spf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Spf.Btr)", csiArchive
    On Error GoTo 0
    imSpfRecLen = Len(tmSpf)
    ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)  'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        imTerminate = True
        Exit Sub
    End If
    hmSaf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Saf.Btr)", csiArchive
    On Error GoTo 0
    imSafRecLen = Len(tmSaf)
    tmSafSrchKey1.iVefCode = 0
    ilRet = btrGetEqual(hmSaf, tmSaf, imSafRecLen, tmSafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        imTerminate = True
        Exit Sub
    End If
    If tmSpf.sGUseAffSys <> "Y" Then
        lacAffiliate.Visible = False
        edcAffiliateStart.Visible = False
        edcAffiliateEnd.Visible = False
        lacAffiliateSpotsStatus(0).Visible = False
        lacAffiliateSpotsStatus(1).Visible = False
    End If
    If tmSpf.sGUsePropSys <> "Y" Then 'Using Proposal System
    End If
    If (Asc(tmSpf.sOptionFields) And &H1) <> &H1 Then
    End If
    
    
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim ilPos As Integer
    Dim ilSpace As Integer
    
    slCommand = Command$
    sgCommandStr = "debug"
    igDirectCall = 0
    sgIniPath = ""
    'slCommand = sgCommandStr  'Command$
    'If (StrComp(slCommand, "Debug", 1) = 0) Or (StrComp(slCommand, "CS", 1) = 0) Then
        igStdAloneMode = True 'Switch from/to stand alone mode
        sgCallAppName = ""
        slStr = "Guide"
        ilTestSystem = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
    '    ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
    '    If Trim$(slStr) = "" Then
    '        MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
    '        End
    '    End If
    '    ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
    '    ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
    '    If StrComp(slTestSystem, "Test", 1) = 0 Then
    '        ilTestSystem = True
    '    Else
    '        ilTestSystem = False
    '    End If
    '    ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    ilPos = InStr(1, slCommand, "/IniLoc:", vbTextCompare)
    If ilPos > 0 Then
        igDirectCall = -1
        ilSpace = InStr(ilPos, slCommand, " ")
        If ilSpace = 0 Then
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8))
        Else
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8, ilSpace - ilPos - 8))
        End If
        If right$(sgIniPath, 1) <> "\" Then
            sgIniPath = sgIniPath & "\"
        End If
    End If
    imUserInput = False
    ilPos = InStr(1, slCommand, "/UserInput", 1)
''Remove this line of code
'ilPos = 1
    If ilPos > 0 Then
        imUserInput = True
        igBkgdProg = 0
        'Debug
        'lMsgBox "Current Mode: UserInput.", vbExclamation, ""
      '  gLogMsg "Current Mode: UserInput.", "csiArchive.txt", False
    Else
        igBkgdProg = 7
        'Debug
        'lMsgBox "Current Mode: Background.", vbExclamation, ""
       ' gLogMsg "Current Mode: Background.", "csiArchive.txt", False

    End If
    imBackupDB = True
    If InStr(1, slCommand, "/NoBackup", vbTextCompare) > 0 Then
        imBackupDB = False
    End If
    imArchiveMethod = 1
    If InStr(1, slCommand, "/NoDLL", vbTextCompare) > 0 Then
        imArchiveMethod = 0
    End If
    imNoOutput = False
    If InStr(1, slCommand, "/NoOutput", vbTextCompare) > 0 Then
        imNoOutput = True
        imArchiveMethod = 0
    End If
    gInitStdAlone csiArchive, slStr, ilTestSystem
    If igBkgdProg = 0 Then
        gLogMsg "Current Mode: UserInput.", "csiArchive.txt", False
    Else
        gLogMsg "Current Mode: Background.", "csiArchive.txt", False
    End If
End Sub
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

        
    igParentRestarted = False
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Traffic
    Unload csiArchive
    igManUnload = NO
End Sub


Private Function mCheckRecLength() As Integer
    If Not gRecLengthOk("Chf.Btr", Len(tgDBChf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Clf.Btr", Len(tgDBClf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Cff.Btr", Len(tgDBCff)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Adf.Btr", Len(tgDBAdf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Agf.Btr", Len(tgDBAgf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    'If Not gRecLengthOk("Mnf.Btr", Len(tgMRMnf)) Then
    '    mCheckRecLength = False
    '    Exit Function
    'End If
    mCheckRecLength = True
End Function

Private Sub lacMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move lacMin.Left - 30, lacMin.Top - 45
    imcOutline.Visible = True
End Sub

Private Sub lacStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move lacStop.Left - 30, lacStop.Top - 45
    imcOutline.Visible = True
End Sub

Private Sub pbcClickFocus_Click()
    imcOutline.Visible = False
End Sub

Private Sub plcSignon_Click()
    imcOutline.Visible = False
End Sub

Private Sub tmcStart_Timer()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim ilPosSv As Integer
    Dim llDate As Long
    
    tmcStart.Enabled = False
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
    mCheckGG
    If igGGFlag = 0 Then
        imTerminate = True
        mTerminate
        Exit Sub
    End If
    
    
    'smArchivePath = Left$(sgDBPath, Len(sgDBPath) - 1)
    'ilPosSv = 0
    'ilPos = InStr(1, smArchivePath, "\", vbTextCompare)
    'Do While ilPos > 0
    '    ilPosSv = ilPos
    '    ilPos = InStr(ilPos + 1, smArchivePath, "\", vbTextCompare)
    'Loop
    'If ilPosSv > 0 Then
    '    On Error Resume Next
    '    smArchivePath = Left$(smArchivePath, ilPosSv) & "Archive"
    '    MkDir smArchivePath
    'Else
    '    mTerminate
    '    Exit Sub
    'End If
    On Error Resume Next
    smArchivePath = Left$(sgArchivePath, Len(sgArchivePath) - 1)
    MkDir smArchivePath
    Screen.MousePointer = vbHourglass
    On Error GoTo 0
    imAffStatus = 0
    imTrafStatus = 0
    smAffStartDate = ""
    smTrafStartDate = ""
    smAffEndDate = ""
    smTrafEndDate = ""
    smPropEndDate = ""
    smProjEndDate = ""
    smCntrEndDate = ""
    smPayRevHistEndDate = ""
    lacMsg.Caption = "Obtaining Archive Dates"
    DoEvents
    gUnpackDate tmSpf.iBLastStdMnth(0), tmSpf.iBLastStdMnth(1), smInvEndStdMnth
    If smInvEndStdMnth <> "" Then
        If tmSpf.sGUseAffSys = "Y" Then
            If tmSpf.iRetainAffSpot > 0 Then
                If tmSpf.iRetainAffSpot < 500 Then
                    llDate = gGetEarliestAffSpotDate(hgDBAst, hgDBAtt, -1)
                    If llDate > 0 Then
                        smAffStartDate = Format$(llDate, "m/d/yy")
                        smAffEndDate = gObtainEndStd(DateAdd("d", -31 * tmSpf.iRetainAffSpot - 7, smInvEndStdMnth))
                        If gDateValue(smAffStartDate) <= gDateValue(smAffEndDate) Then
                            If imUserInput Then
                                edcAffiliateStart.Text = smAffStartDate
                                edcAffiliateEnd.Text = smAffEndDate
                            Else
                                'Restrict removal to one month
                                smAffEndDate = gObtainEndStd(smAffStartDate)
                                edcAffiliateStart.Text = smAffStartDate
                                edcAffiliateEnd.Text = smAffEndDate
                            End If
                            imAffStatus = 0
                        Else
                            'Done required ro be removed
                            lacAffiliate.Caption = "Affiliate Purge Dates- None required"
                            edcAffiliateStart.Visible = False
                            edcAffiliateEnd.Visible = False
                            imAffStatus = 1
                        End If
                    Else
                        If llDate < 0 Then
                            lacAffiliate.Caption = "Affiliate Purge Dates- Unable to Get"
                            edcAffiliateStart.Visible = False
                            edcAffiliateEnd.Visible = False
                            imAffStatus = -1
                        Else
                            'Done required to be removed
                            lacAffiliate.Caption = "Affiliate Purge Dates- None required"
                            edcAffiliateStart.Visible = False
                            edcAffiliateEnd.Visible = False
                            imAffStatus = 1
                        End If
                    End If
                Else
                    'Done required to be removed
                    lacAffiliate.Caption = "Affiliate Purge Dates- Bypass"
                    edcAffiliateStart.Visible = False
                    edcAffiliateEnd.Visible = False
                    imAffStatus = 1
                End If
            Else
                lacAffiliate.Caption = "Retain Affiliate Spots Number of Months Not Set"
                edcAffiliateStart.Visible = False
                edcAffiliateEnd.Visible = False
                imAffStatus = -1
            End If
        End If
        If tmSpf.iRetainTrafSpot > 0 Then
            If tmSpf.iRetainTrafSpot < 500 Then
                llDate = gGetEarliestTrafSpotDate(hgDBSdf, -1)
                smTrafStartDate = Format$(llDate, "m/d/yy")
                smTrafEndDate = gObtainEndStd(DateAdd("d", -31 * tmSpf.iRetainTrafSpot - 7, smInvEndStdMnth))
                If tmSpf.sGUseAffSys = "Y" Then
                    If (tmSpf.iRetainAffSpot > 0) And (smAffEndDate <> "") Then
                        If gDateValue(smTrafEndDate) > gDateValue(smAffEndDate) Then
                            smTrafEndDate = smAffEndDate
                        End If
                    End If
                End If
                If gDateValue(smTrafStartDate) <= gDateValue(smTrafEndDate) Then
                    If imUserInput Then
                        edcTrafficStart.Text = smTrafStartDate
                        edcTrafficEnd.Text = smTrafEndDate
                    Else
                        'Restrict removal to one month
                        smTrafEndDate = gObtainEndStd(smTrafStartDate)
                        edcTrafficStart.Text = smTrafStartDate
                        edcTrafficEnd.Text = smTrafEndDate
                    End If
                    imTrafStatus = 0
                Else
                    'Done required to be removed
                    lacTraffic.Caption = "Traffic Purge Dates- None required"
                    edcTrafficStart.Visible = False
                    edcTrafficEnd.Visible = False
                    imTrafStatus = 1
                End If
            Else
                'Done required to be removed
                lacTraffic.Caption = "Traffic Purge Dates- Bypass"
                edcTrafficStart.Visible = False
                edcTrafficEnd.Visible = False
                imTrafStatus = 1
            End If
        Else
            lacTraffic.Caption = "Retain Traffic Spots Number of Months Not Set"
            edcTrafficStart.Visible = False
            edcTrafficEnd.Visible = False
            imTrafStatus = -1
        End If
        If tgSpf.sGUsePropSys = "Y" Then 'Using Proposal System
            If (tmSpf.iRetainTrafProp > 0) And (tmSpf.iRetainTrafProp < 500) Then
                smPropEndDate = gObtainEndStd(DateAdd("d", -31 * tmSpf.iRetainTrafProp - 7, smInvEndStdMnth))
            Else
                smPropEndDate = ""
            End If
        End If
        If (Asc(tmSpf.sOptionFields) And &H1) = &H1 Then
            If (tmSaf.iRetainTrafProj > 0) And (tmSaf.iRetainTrafProj < 500) Then
                smProjEndDate = gObtainEndStd(DateAdd("d", -31 * tmSaf.iRetainTrafProj - 7, smInvEndStdMnth))
            Else
                smProjEndDate = ""
            End If
        End If
        If (tmSaf.iRetainCntr > 0) And (imTrafStatus >= 0) Then
            If tmSaf.iRetainCntr < 500 Then
                smCntrEndDate = gObtainEndStd(DateAdd("d", -31 * tmSaf.iRetainCntr - 7, smInvEndStdMnth))
                If gDateValue(smTrafEndDate) <= gDateValue(smCntrEndDate) Then
                    smCntrEndDate = smTrafEndDate
                End If
                edcCntrEnd.Text = smCntrEndDate
            Else
                edcCntrEnd.Text = "Bypass"
            End If
        End If
        If (tmSaf.iRetainPayRevHist > 0) And (imTrafStatus >= 0) Then
            If tmSaf.iRetainPayRevHist < 500 Then
                smPayRevHistEndDate = gObtainEndStd(DateAdd("d", -31 * tmSaf.iRetainPayRevHist - 7, smInvEndStdMnth))
                If gDateValue(smTrafEndDate) <= gDateValue(smPayRevHistEndDate) Then
                    smPayRevHistEndDate = smTrafEndDate
                End If
                edcPayRevHistEnd.Text = smPayRevHistEndDate
            Else
                edcPayRevHistEnd.Text = "Bypass"
            End If
        End If
    
        gUnpackDate tmSaf.iLastArchRunDate(0), tmSaf.iLastArchRunDate(1), smLastArchRunDate
        If smLastArchRunDate <> "" Then
            lacLastDatePurged.Caption = lacLastDatePurged.Caption & " " & smLastArchRunDate
        Else
            lacLastDatePurged.Caption = lacLastDatePurged.Caption & " " & "Never Archived"
        End If
    Else
        imAffStatus = -1
        imTrafStatus = -1
    End If
    lacMsg.Caption = ""
    'lacMsg.Visible = False
    DoEvents
    Screen.MousePointer = vbDefault
    If Not imUserInput Then
        If (imAffStatus <> -1) And (imTrafStatus <> -1) Then
            Screen.MousePointer = vbHourglass
            ilRet = mGenerateArchive()
            Screen.MousePointer = vbDefault
        End If
        mTerminate
    Else
        If (imAffStatus = -1) Or (imTrafStatus = -1) Then
            lacMin.Enabled = False
        End If
    End If
    Exit Sub
End Sub
Private Sub lacMin_Click()
    Dim ilRet As Integer
    If imUserInput Then
        If Not imGenArch Then
            Screen.MousePointer = vbHourglass
    smStartDateTime = Now
            ilRet = mGenerateArchive()
    smEndDateTime = Now
            Screen.MousePointer = vbDefault
            If ilRet Then
                lacStop.Caption = "Done"
                lacMin.Enabled = False
            End If
        End If
    Else
        csiArchive.WindowState = vbMinimized
    End If
End Sub

Private Sub lacStop_Click()
    If Not imGenArch Then
        mTerminate
    Else
        imCancelled = True
    End If
End Sub





'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchAvail              *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Avail Name index into    *
'*                     tgAvailAnf                      *
'*                                                     *
'*******************************************************
Public Function mBinarySearchAvail(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tgAvailAnf)
    ilMax = UBound(tgAvailAnf) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tgAvailAnf(ilMiddle).iCode Then
            'found the match
            mBinarySearchAvail = ilMiddle
            Exit Function
        ElseIf ilCode < tgAvailAnf(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchAvail = -1
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchMcf                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Media Code index into    *
'*                     tgMcf                           *
'*                                                     *
'*******************************************************
Public Function mBinarySearchMcf(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tgMCF)
    ilMax = UBound(tgMCF) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tgMCF(ilMiddle).iCode Then
            'found the match
            mBinarySearchMcf = ilMiddle
            Exit Function
        ElseIf ilCode < tgMCF(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchMcf = -1
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchUrf                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain User index into tgPopUrf *
'*                                                     *
'*******************************************************
Public Function mBinarySearchUrf(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tgPopUrf)
    ilMax = UBound(tgPopUrf) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tgPopUrf(ilMiddle).iCode Then
            'found the match
            mBinarySearchUrf = ilMiddle
            Exit Function
        ElseIf ilCode < tgPopUrf(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchUrf = -1
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchMRMnf              *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Missed Reason index into *
'*                     tgMRMnf                         *
'*                                                     *
'*******************************************************
Public Function mBinarySearchMRMnf(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tgMRMnf)
    ilMax = UBound(tgMRMnf) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tgMRMnf(ilMiddle).iCode Then
            'found the match
            mBinarySearchMRMnf = ilMiddle
            Exit Function
        ElseIf ilCode < tgMRMnf(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchMRMnf = -1
End Function

Public Function mBinarySearchLangMnf(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tmLangMnf)
    ilMax = UBound(tmLangMnf) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tmLangMnf(ilMiddle).iCode Then
            'found the match
            mBinarySearchLangMnf = ilMiddle
            Exit Function
        ElseIf ilCode < tmLangMnf(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchLangMnf = -1
End Function


Private Function mGenerateArchive() As Integer
    Dim llDate As Long
    Dim llAffStartDate As Long
    Dim llAffEndDate As Long
    Dim llTrafStartDate As Long
    Dim llTrafEndDate As Long
    Dim slDate As String
    Dim ilNoMonths As Integer
    Dim ilMonth As Integer
    Dim llAtt As Long
    Dim llTestDate As Long
    Dim ilAstError As Integer
    Dim ilLstError As Integer
    Dim ilCpttError As Integer
    Dim ilAttError As Integer
    Dim ilWeblError As Integer
    Dim ilESFError As Integer
    Dim ilLcfError As Integer
    Dim ilSsfError As Integer
    Dim ilVcfError As Integer
    Dim ilVlfError As Integer
    Dim ilDlfError As Integer
    Dim ilEgfError As Integer
    Dim ilSdfError As Integer
    Dim ilPsfError As Integer
    Dim ilStfError As Integer
    Dim ilCrfError As Integer
    Dim ilRqfError As Integer
    Dim ilRet As Integer
    Dim llAffEarliestAstDate As Long
    Dim slArchiveFolderName As String
    Dim slTitle As String
    Dim llLstAstCount As Long
    Dim llSdfAstCount As Long
    Dim llLcfCount As Long
    Dim llSsfCount As Long
    Dim llVcfCount As Long
    Dim llVlfCount As Long
    Dim llDlfCount As Long
    Dim llEgfCount As Long
    Dim llSdfCount As Long
    Dim llPsfCount As Long
    Dim llStfCount As Long
    Dim llCrfCount As Long
    Dim llTotalCCTCount As Long '7-24-09 count of CCT comments removed
    Dim llTotalLstAstCount As Long
    Dim llTotalSdfAstCount As Long
    Dim llTotalLstCount As Long
    Dim llTotalCpttCount As Long
    Dim llTotalAttCount As Long
    Dim llTotalWeblCount As Long
    Dim llTotalLcfCount As Long
    Dim llTotalSsfCount As Long
    Dim llTotalVcfCount As Long
    Dim llTotalVlfCount As Long
    Dim llTotalDlfCount As Long
    Dim llTotalEgfCount As Long
    Dim llTotalSdfCount As Long
    Dim llTotalPsfCount As Long
    Dim llTotalStfCount As Long
    Dim llTotalCrfCount As Long
    Dim llTotalCifCount As Long
    Dim llTotalBofCount As Long
    Dim llTotalPropCount As Long
    Dim llTotalPjfCount As Long
    Dim llTotalAdfCount As Long
    Dim llTotalAgfCount As Long
    Dim llTotalEsfCount As Long
    Dim llTotalRqfCount As Long
    Dim ilVef As Integer
    Dim llDropDate As Long
    Dim llOffAirDate As Long
    Dim llLstEarliestDate As Long
    Dim llCpttEarliestDate As Long
    Dim llLstCpttStartDate As Long
    Dim llTotalsNoDays As Long
    Dim llNoDays As Long
    Dim llPercent As Long
    Dim llEarliestTrafSpot As Long
    Dim llFirstAllowedCopyDate As Long
    Dim llCrf As Long
    Dim llCrfEarliestDate As Long
    Dim llCrfLatestDate As Long
    Dim llCopyStartDate As Long
    Dim llCopyEndDate As Long
    Dim ilMissingDate As Integer
    Dim llAdfAgfRetainDate As Long
    Dim llEsfDate As Long
    Dim llRqfDate As Long
    
    llEarliestTrafSpot = 0
    llAffEarliestAstDate = 0
    'imArchiveMethod = 1 '0=Archive and delete within this code; 1=Call DLL for archive and delete
    'Check if valid data
    If imAffStatus = 0 Then
        smAffStartDate = edcAffiliateStart.Text
        slDate = edcAffiliateEnd.Text
        If StrComp(slDate, "Bypass", vbTextCompare) = 0 Then
            slDate = ""
        End If
    Else
        smAffStartDate = ""
        slDate = ""
        smAffEndDate = ""
    End If
    If tmSpf.sGUseAffSys = "Y" Then
        If slDate <> "" Then
            If gDateValue(slDate) > gDateValue(smAffEndDate) Then
                If imUserInput Then
                    MsgBox "Affiliate End Date can't be advanced"
                    edcAffiliateEnd.Text = smAffEndDate
                Else
                    gLogMsg "Archive not run because Affiliate End Date advanced", "csiArchive.txt", False
                End If
                mGenerateArchive = False
                Exit Function
            End If
            If gDateValue(gObtainEndStd(slDate)) <> gDateValue(slDate) Then
                If imUserInput Then
                    MsgBox "Affiliate End Date must be the End of a Standard Broadcast Month"
                Else
                    gLogMsg "Archive not run because Affiliate End Date is not the End Of a Standard Broadcast Month", "csiArchive.txt", False
                End If
                mGenerateArchive = False
                Exit Function
            End If
        End If
    End If
    If imAffStatus = 0 Then
        smAffEndDate = Trim$(edcAffiliateEnd.Text)
        If StrComp(smAffEndDate, "Bypass", vbTextCompare) = 0 Then
            smAffEndDate = ""
        End If
    Else
        smAffEndDate = ""
    End If
    If imTrafStatus = 0 Then
        smTrafStartDate = edcTrafficStart.Text
        slDate = edcTrafficEnd.Text
        If StrComp(slDate, "Bypass", vbTextCompare) = 0 Then
            slDate = ""
        End If
    Else
        smTrafStartDate = ""
        slDate = ""
    End If
    If slDate <> "" Then
        If gDateValue(slDate) > gDateValue(smTrafEndDate) Then
            If imUserInput Then
                MsgBox "Traffic End Date can't be advanced"
                edcTrafficEnd.Text = smTrafEndDate
            Else
                gLogMsg "Archive not run because Traffic End Date advanced", "csiArchive.txt", False
            End If
            mGenerateArchive = False
            Exit Function
        End If
        If gDateValue(gObtainEndStd(slDate)) <> gDateValue(slDate) Then
            If imUserInput Then
                MsgBox "Traffic End Date must be the End of a Standard Broadcast Month"
            Else
                gLogMsg "Archive not run because Traffic End Date is not the End Of a Standard Broadcast Month", "csiArchive.txt", False
            End If
            mGenerateArchive = False
            Exit Function
        End If
        If smAffEndDate <> "" Then
            If gDateValue(slDate) > gDateValue(smAffEndDate) Then
                If imUserInput Then
                    MsgBox "Traffic End Date can't be newer than Affiliate End Date"
                    edcTrafficEnd.Text = smTrafEndDate
                Else
                    gLogMsg "Archive not run because Traffic End Date newer than Affiliate End Date", "csiArchive.txt", False
                End If
                mGenerateArchive = False
                Exit Function
            End If
        End If
    End If
    If imTrafStatus = 0 Then
        smTrafEndDate = edcTrafficEnd.Text
        If StrComp(slDate, "Bypass", vbTextCompare) = 0 Then
            smTrafEndDate = ""
        End If
    Else
        smTrafEndDate = ""
    End If
    If Not mCheckUnusedDates() Then
        mGenerateArchive = False
        Exit Function
    End If
    If imUserInput And imNoOutput Then
        ilRet = MsgBox("Removing Information without creating any Archive files, Ok to continue?", vbQuestion + vbYesNo, "Warning")
        If ilRet = vbNo Then
            mGenerateArchive = False
            Exit Function
        End If
    End If
    If imArchiveMethod = 1 Then
        If Not csiArchiveInit(sgDBPath) Then
            If imUserInput Then
                MsgBox "Unable to Initialize csi_Utils"""
            Else
                gLogMsg "Unable to Initialize csi_Utils", "csiArchive.txt", False
            End If
            mGenerateArchive = False
            Exit Function
        End If
    End If
    gLogMsg "", "csiArchive.txt", False
    gLogMsg "Archive started, Method = " & imArchiveMethod, "csiArchive.txt", False

    If Not mBackupDatabase() Then
        mGenerateArchive = False
        Exit Function
    End If

    mSplitVefCodes
    plcGauge.Visible = True
    plcGauge.Value = 0
    imGenArch = True
    mGenerateArchive = True
    llTotalCCTCount = 0
    llTotalLstAstCount = 0
    llTotalSdfAstCount = 0
    llTotalLstCount = 0
    llTotalCpttCount = 0
    llTotalAttCount = 0
    llTotalWeblCount = 0
    llTotalLcfCount = 0
    llTotalSsfCount = 0
    llTotalVcfCount = 0
    llTotalVlfCount = 0
    llTotalDlfCount = 0
    llTotalEgfCount = 0
    llTotalSdfCount = 0
    llTotalPsfCount = 0
    llTotalStfCount = 0
    llTotalCrfCount = 0
    llTotalCifCount = 0
    llTotalBofCount = 0
    llTotalPropCount = 0
    llTotalPjfCount = 0
    llTotalAdfCount = 0
    llTotalAgfCount = 0
    llTotalEsfCount = 0
    llTotalRqfCount = 0
    ilAstError = False
    ilLstError = False
    ilCpttError = False
    ilAttError = False
    ilWeblError = False
    ilLcfError = False
    ilSsfError = False
    ilVcfError = False
    ilVlfError = False
    ilDlfError = False
    ilEgfError = False
    ilSdfError = False
    ilPsfError = False
    ilStfError = False
    ilCrfError = False
    ilESFError = False
    ilRqfError = False
    If tmSpf.sGUseAffSys = "Y" Then
        If smAffEndDate <> "" Then
            If gDateValue(smAffStartDate) <= gDateValue(smAffEndDate) Then
                If imNoOutput Then
                    lacMsg.Caption = "Removing Affiliate Spots"
                Else
                    lacMsg.Caption = "Archiving and Removing Affiliate Spots"
                End If
                DoEvents
                llTotalsNoDays = gDateValue(smAffEndDate) - gDateValue(smAffStartDate) + 1
                llNoDays = 0
                llLstEarliestDate = mGetEarliestLstDate(-1)
                llCpttEarliestDate = mGetEarliestCpttDate(-1)
                llLstCpttStartDate = -1
                gLogMsg "Affiliate Spots Archive date range- From: " & smAffStartDate & " to " & smAffEndDate, "csiArchive.Txt", False
                slDate = gObtainEndStd(smAffStartDate)
                ilNoMonths = DateDiff("m", slDate, smAffEndDate) + 1
                'Gather agreements
                mGetAtt
                llAffStartDate = gDateValue(smAffStartDate)
                llAffEndDate = gDateValue(slDate)
                lacAffiliateSpotsStatus(1).Caption = "0" & " of " & Trim$(Str$(ilNoMonths)) & " Months Processed"
                lacAffiliateSpotsStatus(1).Visible = True
                DoEvents
                For ilMonth = 1 To ilNoMonths Step 1
                    llLstAstCount = 0
                    llSdfAstCount = 0
                    DoEvents
                    If imCancelled Then
                        lacAffiliateSpotsStatus(1).Caption = "Cancelled"
                        Exit For
                    End If
                    slDate = gObtainEndStd(Format$(llAffEndDate, "m/d/yy"))
                    slArchiveFolderName = smArchivePath & "\Archive" & Format$(gAdjYear(slDate), "yyyymmdd")
                    On Error Resume Next
                    MkDir slArchiveFolderName
                    If imArchiveMethod = 0 Then
                        ilRet = mOpenArchiveFile(slArchiveFolderName & "\AffSpots" & Format$(gAdjYear(slDate), "yyyymmdd") & ".Csv")
                        If Not ilRet Then
                            plcGauge.Visible = False
                            imGenArch = False
                            mGenerateArchive = False
                            Exit Function
                        End If
                        slTitle = "Traf/Log Spot,Contract #,Line #,Call Letters,Line Vehicle,Agency,City,Advertiser,Addr ID,Product,Net Region,Feed Date,Feed Time,Pledge Date,Pledge Start Time,Pledge End Time,Spot Len,Game #,Copy Region,Cart,Product,ISCI,Creative Title,Audience,Demo,Airing vehicle,Air Date,Air Time,CP Posting Status,Air Status" & "," & Now
                        mPrintTitle slTitle 'Print #hmArchive, slTitle
                        For llDate = llAffStartDate To llAffEndDate Step 1
                            DoEvents
                            If imCancelled Then
                                lacAffiliateSpotsStatus(1).Caption = "Cancelled"
                                Exit For
                            End If
                            For llAtt = 0 To UBound(tmATTDateArchive) - 1 Step 1
                                ilAstError = Not mArchiveAndRemoveAst(tmATTDateArchive(llAtt).lCode, llDate, tmATTDateArchive(llAtt).iExportType, llLstAstCount, llSdfAstCount)
                                DoEvents
                                If (ilAstError) Or (imCancelled) Then
                                    Exit For
                                End If
                            Next llAtt
                            If (ilAstError) Or (imCancelled) Then
                                Exit For
                            End If
                        Next llDate
                        mClose  'Close #hmArchive
                        btrDestroy hgDBAst
                        ilRet = btrClose(hgDBAst)
                        DoEvents
                        hgDBAst = CBtrvTable(TWOHANDLES)    'CBtrvObj()
                        ilRet = btrOpen(hgDBAst, "", sgDBPath & "Ast.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                    Else
                        If Not csiOpenArchiveFileName(slArchiveFolderName & "\AffSpots" & Format$(gAdjYear(slDate), "yyyymmdd") & ".Csv") Then
                            plcGauge.Visible = False
                            imGenArch = False
                            mGenerateArchive = False
                            Exit Function
                        End If
                        For llDate = llAffStartDate To llAffEndDate Step 1
                            DoEvents
                            If imCancelled Then
                                lacAffiliateSpotsStatus(1).Caption = "Cancelled"
                                Exit For
                            End If
                            For llAtt = 0 To UBound(tmATTDateArchive) - 1 Step 1
                                ilAstError = Not csiArchiveAndRemoveAst(tmATTDateArchive(llAtt).lCode, Format$(llDate, "mm/dd/yyyy"), Format$(llDate, "mm/dd/yyyy"), tmATTDateArchive(llAtt).iExportType, llLstAstCount, llSdfAstCount)
                                DoEvents
                                If (ilAstError) Or (imCancelled) Then
                                    Exit For
                                End If
                            Next llAtt
                            If (ilAstError) Or (imCancelled) Then
                                Exit For
                            End If
                        Next llDate
                        ilRet = csiCloseArchiveFile()
                    End If
                    gLogMsg Trim$(Str$(llLstAstCount)) & " Affiliate Spots from Affiliate Log Spot Archived and Removed for month " & Format$(llAffEndDate, "m/d/yy"), "csiArchive.Txt", False
                    gLogMsg Trim$(Str$(llSdfAstCount)) & " Affiliate Spots from Traffic Spot Archived and Removed for month " & Format$(llAffEndDate, "m/d/yy"), "csiArchive.Txt", False
                    gLogMsg "Total of " & Trim$(Str$(llLstAstCount + llSdfAstCount)) & " Affiliate Spots Archived and Removed for month " & Format$(llAffEndDate, "m/d/yy"), "csiArchive.Txt", False
                    llTotalLstAstCount = llTotalLstAstCount + llLstAstCount
                    llTotalSdfAstCount = llTotalSdfAstCount + llSdfAstCount
                    If (ilAstError) Or (imCancelled) Then
                        Exit For
                    End If
                    
                    If (ilAstError) Or (imCancelled) Then
                        llAffEarliestAstDate = gGetEarliestAffSpotDate(hgDBAst, hgDBAtt, -1)
                    Else
                        llAffEarliestAstDate = llAffEndDate + 1
                    End If
                    gPackDateLong llAffEarliestAstDate, tmSaf.iEarliestAffSpot(0), tmSaf.iEarliestAffSpot(1)
                    ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
                    If (imCancelled) Then
                        lacAffiliateSpotsStatus(1).Caption = "Cancelled"
                        DoEvents
                        Exit For
                    End If
                    If (ilAstError) Then
                        lacAffiliateSpotsStatus(1).Caption = "AST removal error"
                        DoEvents
                        Exit For
                    End If
                    llNoDays = llNoDays + (llAffEndDate - llAffStartDate) + 1
                    lacAffiliateSpotsStatus(1).Caption = Trim$(Str$(ilMonth)) & " of " & Trim$(Str$(ilNoMonths)) & " Months Processed"
                    DoEvents
                    llPercent = ((CLng(llNoDays) + 1) * CSng(100)) / llTotalsNoDays
                    If llPercent >= 100 Then
                        llPercent = 100
                    End If
                    plcGauge.Value = llPercent
                    DoEvents
                    If (imCancelled) Then
                        lacAffiliateSpotsStatus(1).Caption = "Cancelled"
                        DoEvents
                        Exit For
                    End If
                    llAffStartDate = llAffEndDate + 1
                    llAffEndDate = gDateValue(gObtainEndStd(Format$(llAffStartDate, "m/d/yy")))
                Next ilMonth
                If ilNoMonths > 1 And Not imCancelled Then
                    gLogMsg "Total of " & Trim$(Str$(llTotalLstAstCount)) & " Affiliate Spots from Affiliate Log Spot Archived and Removed", "csiArchive.Txt", False
                    gLogMsg "Total of " & Trim$(Str$(llTotalSdfAstCount)) & " Affiliate Spots from Traffic Spot Archived and Removed", "csiArchive.Txt", False
                    gLogMsg "Grand Total of " & Trim$(Str$(llTotalLstAstCount + llTotalSdfAstCount)) & " Affiliate Spots Archived and Removed", "csiArchive.Txt", False
                End If
                DoEvents
                If (Not imCancelled) And (Not ilAstError) Then
                    'Remove LST, CPTT
                    'For ilVef = 0 To UBound(tmCASGVVefCode) - 1 Step 1
                    '    ilLstError = Not mRemoveLst(tmCASGVVefCode(ilVef).iVefCode, llAffEarliestAstDate, llTotalLstCount)
                    '    If (ilLstError) Or (imCancelled) Then
                    '        Exit For
                    '    End If
                    '    ilCpttError = Not mRemoveCptt(tmCASGVVefCode(ilVef).iVefCode, llAffEarliestAstDate, llTotalCpttCount)
                    '    If (ilCpttError) Or (imCancelled) Then
                    '        Exit For
                    '    End If
                    'Next ilVef
                    For ilVef = 0 To UBound(tmLstVefCode) - 1 Step 1
                        DoEvents
                        If imCancelled Then
                            lacAffiliateSpotsStatus(1).Caption = "Cancelled"
                            Exit For
                        End If
                        ilLstError = Not mRemoveLst(tmLstVefCode(ilVef).iVefCode, llAffEarliestAstDate, llTotalLstCount)
                        If (ilLstError) Then
                            Exit For
                        End If
                    Next ilVef
                    For ilVef = 0 To UBound(tmCpttVefCode) - 1 Step 1
                        DoEvents
                        If imCancelled Then
                            lacAffiliateSpotsStatus(1).Caption = "Cancelled"
                            Exit For
                        End If
                        ilCpttError = Not mRemoveCptt(tmCpttVefCode(ilVef).iVefCode, llAffEarliestAstDate, llTotalCpttCount)
                        If (ilCpttError) Then
                            Exit For
                        End If
                    Next ilVef
                    
                    '7-24-09 Remove the CCT comments
                    ilRet = mRemoveCCT(llAffEarliestAstDate, llTotalCCTCount)
                    gLogMsg "Total of " & Trim$(Str$(llTotalCCTCount)) & " Affiliate CP Comments Removed through " & Format$(llAffEarliestAstDate - 1, "m/d/yy"), "csiArchive.Txt", False
                    gLogMsg "Total of " & Trim$(Str$(llTotalLstCount)) & " Affiliate Log Spots Removed through " & Format$(llAffEarliestAstDate - 1, "m/d/yy"), "csiArchive.Txt", False
                    gLogMsg "Total of " & Trim$(Str$(llTotalCpttCount)) & " Affiliate C.P. Posting Days Removed through " & Format$(llAffEarliestAstDate - 1, "m/d/yy"), "csiArchive.Txt", False
                    If (imCancelled) Then
                        lacAffiliateSpotsStatus(1).Caption = "Cancelled"
                        DoEvents
                    End If
                    If (ilLstError) Or (ilCpttError) Then
                        lacAffiliateSpotsStatus(1).Caption = "CPTT or LST Removal Error"
                        DoEvents
                    End If
                    If (Not ilLstError) And (Not ilCpttError) And (Not imCancelled) Then
                        ilAttError = Not mRemoveAtt(llAffEarliestAstDate, llTotalAttCount)
                    End If
                    gLogMsg "Total of " & Trim$(Str$(llTotalAttCount)) & " Affiliate Agreements Removed through " & Format$(llAffEarliestAstDate - 1, "m/d/yy"), "csiArchive.Txt", False
                    If (Not ilLstError) And (Not ilCpttError) And (Not imCancelled) And (Not ilAttError) Then
                        ilWeblError = Not mRemoveWebl(llAffEarliestAstDate, llTotalWeblCount)
                    End If
                    gLogMsg "Total of " & Trim$(Str$(llTotalWeblCount)) & " Web Posting Information Removed through " & Format$(llAffEarliestAstDate - 1, "m/d/yy"), "csiArchive.Txt", False
                    '5/22/18: Moved as part of the affiliate logic
                    If (Not ilLstError) And (Not ilCpttError) And (Not imCancelled) And (Not ilAttError) Then
                        ilESFError = mRemoveEsf(llAffEarliestAstDate, llTotalEsfCount, llEsfDate)
                    End If
                    gLogMsg "Total of " & Trim$(Str$(llTotalEsfCount)) & " Affiliate Export Information records Removed through " & Format$(llEsfDate, "m/d/yy"), "csiArchive.Txt", False

                End If
                gLogMsg "Affiliate Spots Archive last completed date " & Format$(llAffEarliestAstDate - 1, "m/d/yy"), "csiArchive.Txt", False
                If imCancelled Then
                    gLogMsg "Archive cancelled by User", "csiArchive.Txt", False
                    lacMsg.Caption = "User Terminate Archive"
                    DoEvents
                Else
                    If (ilAstError) Or (ilLstError) Or (ilCpttError) Or (ilAttError) Then
                        gLogMsg "Affiliate Archive cancelled because of an Error", "csiArchive.Txt", False
                        lacMsg.Caption = "Archive Terminate because of an Error"
                        DoEvents
                    Else
                        lacAffiliateSpotsStatus(1).Caption = "Completed"
                        DoEvents
                    End If
                End If
            Else
                llAffEarliestAstDate = gGetEarliestAffSpotDate(hgDBAst, hgDBAtt, -1)
            End If
        Else
            llAffEarliestAstDate = gGetEarliestAffSpotDate(hgDBAst, hgDBAtt, -1)
            lacAffiliateSpotsStatus(1).Caption = "Bypassed, as date not defined"
            lacAffiliateSpotsStatus(1).Visible = True
        End If
    Else
        llAffEarliestAstDate = gDateValue("12/31/2069")
    End If
    If (imCancelled) Then
        If imArchiveMethod = 1 Then
            ilRet = csiArchiveDeInit()
        End If
        plcGauge.Visible = False
        imGenArch = False
        mGenerateArchive = False
        Exit Function
    End If
    plcGauge.Value = 0
    'Programming, Traffic Spots, Copy, Blackouts
    If smTrafEndDate <> "" Then
        If gDateValue(smTrafStartDate) <= gDateValue(smTrafEndDate) Then
            If imNoOutput Then
                lacMsg.Caption = "Removing Traffic Spots"
            Else
                lacMsg.Caption = "Archiving and Removing Traffic Spots"
            End If
            DoEvents
            llTotalsNoDays = gDateValue(smTrafEndDate) - gDateValue(smTrafStartDate) + 1
            llNoDays = 0
            gLogMsg "Traffic Spots Archive date range- From: " & smTrafStartDate & " to " & smTrafEndDate, "csiArchive.Txt", False
            slDate = gObtainEndStd(smTrafStartDate)
            ilNoMonths = DateDiff("m", slDate, smTrafEndDate) + 1
            llTrafStartDate = gDateValue(smTrafStartDate)
            llTrafEndDate = gDateValue(slDate)
            lacTrafficSpotsStatus(1).Caption = "0" & " of " & Trim$(Str$(ilNoMonths)) & " Months Processed"
            lacTrafficSpotsStatus(1).Visible = True
            DoEvents
            For ilMonth = 1 To ilNoMonths Step 1
                llLcfCount = 0
                llSsfCount = 0
                llVcfCount = 0
                llVlfCount = 0
                llDlfCount = 0
                llEgfCount = 0
                llStfCount = 0
                llSdfCount = 0
                llPsfCount = 0
                DoEvents
                If imCancelled Then
                    lacTrafficSpotsStatus(1).Caption = "Cancelled"
                    Exit For
                End If
                slDate = gObtainEndStd(Format$(llTrafEndDate, "m/d/yy"))
                slArchiveFolderName = smArchivePath & "\Archive" & Format$(gAdjYear(slDate), "yyyymmdd")
                On Error Resume Next
                MkDir slArchiveFolderName
                ilRet = mOpenArchiveFile(slArchiveFolderName & "\TrafSpots" & Format$(gAdjYear(slDate), "yyyymmdd") & ".Csv")
                If Not ilRet Then
                    plcGauge.Visible = False
                    imGenArch = False
                    mGenerateArchive = False
                    Exit Function
                End If
                slTitle = "Spot/Pkg,Contract #,Line #,Agency,City,Advertiser,Addr ID,Product,Net Region,Line Vehicle,Airing Vehicle,Date,Time,Len,Status,Tracer,PLchg,C/T,%Trade,Actual Price,Rate Card Price,Spot Type,Billed,User,Game #,Miss Date,Miss Time,Reason,Region Copy Names,COPY,Zone;Cart #;Product;ISCI;Creative;Script" & "," & Now
                mPrintTitle slTitle 'Print #hmArchive, slTitle
                
                llDate = llTrafEndDate
                If (llDate >= llAffEarliestAstDate) Then
                    llDate = llAffEarliestAstDate - 1
                End If
                'For ilVef = 0 To UBound(tmCASGVVefCode) - 1 Step 1
                '    DoEvents
                '    If imCancelled Then
                '        lacTrafficSpotsStatus(1).Caption = "Cancelled"
                '        Exit For
                '    End If
                '    'Log Calendar
                '    ilLcfError = Not mRemoveLcf(tmCASGVVefCode(ilVef).iVefCode, llDate, llLcfCount)
                '    'Spot Summary
                '    ilSsfError = Not mRemoveSsf(tmCASGVVefCode(ilVef).iVefCode, llDate, llSsfCount)
                '    'Vehicle links and Conflicts
                '    If tmCASGVVefCode(ilVef).sType = "S" Then
                '        ilVcfError = Not mRemoveVcf(tmCASGVVefCode(ilVef).iVefCode, llDate, llVcfCount)
                '        ilVlfError = Not mRemoveVlf(tmCASGVVefCode(ilVef).iVefCode, llDate, llVlfCount)
                '    End If
                '    'Should be A only but S and C added
                '    If (tmCASGVVefCode(ilVef).sType = "A") Or (tmCASGVVefCode(ilVef).sType = "S") Or (tmCASGVVefCode(ilVef).sType = "C") Then
                '        ilDlfError = Not mRemoveDlf(hgDBDlf, tmCASGVVefCode(ilVef).iVefCode, llDate, llDlfCount)
                '    End If
                '    'Should be C and A only but S added as some found
                '    If (tmCASGVVefCode(ilVef).sType = "C") Or (tmCASGVVefCode(ilVef).sType = "A") Or (tmCASGVVefCode(ilVef).sType = "S") Then
                '        ilEgfError = Not mRemoveDlf(hgDBEgf, tmCASGVVefCode(ilVef).iVefCode, llDate, llEgfCount)
                '    End If
                '    If (tmCASGVVefCode(ilVef).sType = "C") Or (tmCASGVVefCode(ilVef).sType = "S") Or (tmCASGVVefCode(ilVef).sType = "V") Then
                '        ilStfError = Not mRemoveStf(tmCASGVVefCode(ilVef).iVefCode, llDate, llStfCount)
                '    End If
                '    If (ilLcfError) Or (ilSsfError) Or (imCancelled) Then
                '        Exit For
                '    End If
                'Next ilVef
                For ilVef = 0 To UBound(tmLcfVefCode) - 1 Step 1
                    DoEvents
                    If imCancelled Then
                        lacTrafficSpotsStatus(1).Caption = "Cancelled"
                        Exit For
                    End If
                    'Log Calendar
                    ilLcfError = Not mRemoveLcf(tmLcfVefCode(ilVef).iVefCode, llDate, llLcfCount)
                    If (ilLcfError) Then
                        Exit For
                    End If
                Next ilVef
                For ilVef = 0 To UBound(tmSsfVefCode) - 1 Step 1
                    DoEvents
                    If imCancelled Then
                        lacTrafficSpotsStatus(1).Caption = "Cancelled"
                        Exit For
                    End If
                    'Spot Summary
                    ilSsfError = Not mRemoveSsf(tmSsfVefCode(ilVef).iVefCode, llDate, llSsfCount)
                    If (ilSsfError) Then
                        Exit For
                    End If
                Next ilVef
                If (Not ilLcfError) And (Not ilSsfError) Then
                    For ilVef = 0 To UBound(tmVcfVefCode) - 1 Step 1
                        DoEvents
                        If imCancelled Then
                            lacTrafficSpotsStatus(1).Caption = "Cancelled"
                            Exit For
                        End If
                        'Vehicle conflict
                        ilVcfError = Not mRemoveVcf(tmVcfVefCode(ilVef).iVefCode, llDate, llVcfCount)
                        If (ilVcfError) Then
                            Exit For
                        End If
                    Next ilVef
                    For ilVef = 0 To UBound(tmVlfVefCode) - 1 Step 1
                        DoEvents
                        If imCancelled Then
                            lacTrafficSpotsStatus(1).Caption = "Cancelled"
                            Exit For
                        End If
                        'Vehicle link
                        ilVlfError = Not mRemoveVlf(tmVlfVefCode(ilVef).iVefCode, llDate, llVlfCount)
                        If (ilVlfError) Then
                            Exit For
                        End If
                    Next ilVef
                    For ilVef = 0 To UBound(tmDlfVefCode) - 1 Step 1
                        DoEvents
                        If imCancelled Then
                            lacTrafficSpotsStatus(1).Caption = "Cancelled"
                            Exit For
                        End If
                        'Delivery link
                        ilDlfError = Not mRemoveDlf(hgDBDlf, tmDlfVefCode(ilVef).iVefCode, llDate, llDlfCount)
                        If (ilDlfError) Then
                            Exit For
                        End If
                    Next ilVef
                    For ilVef = 0 To UBound(tmEgfVefCode) - 1 Step 1
                        DoEvents
                        If imCancelled Then
                            lacTrafficSpotsStatus(1).Caption = "Cancelled"
                            Exit For
                        End If
                        'Engineering link
                        ilEgfError = Not mRemoveDlf(hgDBEgf, tmEgfVefCode(ilVef).iVefCode, llDate, llEgfCount)
                        If (ilEgfError) Then
                            Exit For
                        End If
                    Next ilVef
                    For ilVef = 0 To UBound(tmStfVefCode) - 1 Step 1
                        DoEvents
                        If imCancelled Then
                            lacTrafficSpotsStatus(1).Caption = "Cancelled"
                            Exit For
                        End If
                        'Spot Tracking
                        ilStfError = Not mRemoveStf(tmStfVefCode(ilVef).iVefCode, llDate, llStfCount)
                        If (ilStfError) Then
                            Exit For
                        End If
                    Next ilVef
                End If
                gLogMsg Trim$(Str$(llLcfCount)) & " Traffic Log Calendar Days Removed for month " & Format$(llDate, "m/d/yy"), "csiArchive.Txt", False
                llTotalLcfCount = llTotalLcfCount + llLcfCount
                gLogMsg Trim$(Str$(llSsfCount)) & " Traffic Spot Summary Days Removed for month " & Format$(llDate, "m/d/yy"), "csiArchive.Txt", False
                llTotalSsfCount = llTotalSsfCount + llSsfCount
                gLogMsg Trim$(Str$(llVlfCount)) & " Traffic Selling/Airing Links Removed for month " & Format$(llDate, "m/d/yy"), "csiArchive.Txt", False
                llTotalVlfCount = llTotalVlfCount + llVlfCount
                gLogMsg Trim$(Str$(llVcfCount)) & " Traffic Selling Conflicts Removed for month " & Format$(llDate, "m/d/yy"), "csiArchive.Txt", False
                llTotalVcfCount = llTotalVcfCount + llVcfCount
                gLogMsg Trim$(Str$(llDlfCount)) & " Traffic Delivery Links Removed for month " & Format$(llDate, "m/d/yy"), "csiArchive.Txt", False
                llTotalDlfCount = llTotalDlfCount + llDlfCount
                gLogMsg Trim$(Str$(llEgfCount)) & " Traffic Engineering Links Removed for month " & Format$(llDate, "m/d/yy"), "csiArchive.Txt", False
                llTotalEgfCount = llTotalEgfCount + llEgfCount
                gLogMsg Trim$(Str$(llStfCount)) & " Traffic Spot Tracking Removed for month " & Format$(llDate, "m/d/yy"), "csiArchive.Txt", False
                llTotalStfCount = llTotalStfCount + llStfCount
                For llDate = llTrafStartDate To llTrafEndDate Step 1
                    DoEvents
                    If imCancelled Then
                        lacTrafficSpotsStatus(1).Caption = "Cancelled"
                        Exit For
                    End If
                    If (ilLcfError) Or (ilSsfError) Then
                        lacTrafficSpotsStatus(1).Caption = "Lcf or Ssf removal error"
                        DoEvents
                        Exit For
                    End If
                    slDate = Format(llDate, "m/d/yy")
                    If tmSpf.sGUseAffSys = "Y" Then
                        If (llDate >= llAffEarliestAstDate) Then
                            Exit For
                        End If
                    End If
                    ilSdfError = Not mArchiveAndRemoveSdf("S", llDate, llSdfCount)
                    If (imCancelled) Then
                        lacTrafficSpotsStatus(1).Caption = "Cancelled"
                        DoEvents
                        Exit For
                    End If
                    If (ilSdfError) Then
                        lacTrafficSpotsStatus(1).Caption = "Sdf removal error"
                        DoEvents
                        Exit For
                    End If
                    ilPsfError = Not mArchiveAndRemoveSdf("P", llDate, llPsfCount)
                    If (imCancelled) Then
                        lacTrafficSpotsStatus(1).Caption = "Cancelled"
                        DoEvents
                        Exit For
                    End If
                    If (ilPsfError) Then
                        lacTrafficSpotsStatus(1).Caption = "Psf removal error"
                        DoEvents
                        Exit For
                    End If
                Next llDate
                mClose  'Close #hmArchive
                gLogMsg Trim$(Str$(llSdfCount)) & " Traffic Spots Archived and Removed for month " & Format$(llTrafEndDate, "m/d/yy"), "csiArchive.Txt", False
                llTotalSdfCount = llTotalSdfCount + llSdfCount
                gLogMsg Trim$(Str$(llPsfCount)) & " Traffic Package Spots Archived and Removed for month " & Format$(llTrafEndDate, "m/d/yy"), "csiArchive.Txt", False
                llTotalPsfCount = llTotalPsfCount + llPsfCount
                llNoDays = llNoDays + (llTrafEndDate - llTrafStartDate) + 1
                lacTrafficSpotsStatus(1).Caption = Trim$(Str$(ilMonth)) & " of " & Trim$(Str$(ilNoMonths)) & " Months Processed"
                DoEvents
                llPercent = ((CLng(llNoDays) + 1) * CSng(100)) / llTotalsNoDays
                If llPercent >= 100 Then
                    llPercent = 100
                End If
                plcGauge.Value = llPercent
                DoEvents
                If (imCancelled) Then
                    Exit For
                End If
                If (ilLcfError) Or (ilSsfError) Or (ilSdfError) Or (ilPsfError) Then
                    Exit For
                End If
                llTrafStartDate = llTrafEndDate + 1
                llTrafEndDate = gDateValue(gObtainEndStd(Format$(llTrafStartDate, "m/d/yy")))
            Next ilMonth
            If ilNoMonths > 1 Then
                If tmSpf.sGUseAffSys = "Y" Then
                    If (llTrafStartDate >= llAffEarliestAstDate) Then
                        llTrafStartDate = llAffEarliestAstDate
                    End If
                End If
                gLogMsg "Total of " & Trim$(Str$(llTotalLcfCount)) & " Traffic Log Calendar Days Removed through " & Format$(llTrafStartDate - 1, "m/d/yy"), "csiArchive.Txt", False
                gLogMsg "Total of " & Trim$(Str$(llTotalSsfCount)) & " Traffic Spot Summary Days Removed through " & Format$(llTrafStartDate - 1, "m/d/yy"), "csiArchive.Txt", False
                gLogMsg "Total of " & Trim$(Str$(llTotalVlfCount)) & " Traffic Selling/Airing Links Removed through " & Format$(llTrafStartDate - 1, "m/d/yy"), "csiArchive.Txt", False
                gLogMsg "Total of " & Trim$(Str$(llTotalVcfCount)) & " Traffic Selling Conflicts Removed through " & Format$(llTrafStartDate - 1, "m/d/yy"), "csiArchive.Txt", False
                gLogMsg "Total of " & Trim$(Str$(llTotalDlfCount)) & " Traffic Delivery Links Removed through " & Format$(llTrafStartDate - 1, "m/d/yy"), "csiArchive.Txt", False
                gLogMsg "Total of " & Trim$(Str$(llTotalEgfCount)) & " Traffic Engineering Links Removed through " & Format$(llTrafStartDate - 1, "m/d/yy"), "csiArchive.Txt", False
                gLogMsg "Total of " & Trim$(Str$(llTotalStfCount)) & " Traffic Spot Tracking Removed through " & Format$(llTrafStartDate - 1, "m/d/yy"), "csiArchive.Txt", False
                gLogMsg "Total of " & Trim$(Str$(llTotalSdfCount)) & " Traffic Spots Archived and Removed through " & Format$(llTrafStartDate - 1, "m/d/yy"), "csiArchive.Txt", False
                gLogMsg "Total of " & Trim$(Str$(llTotalPsfCount)) & " Traffic Package Spots Archived and Removed through " & Format$(llTrafStartDate - 1, "m/d/yy"), "csiArchive.Txt", False
            End If
            If imCancelled Then
                gLogMsg "Archive cancelled by User", "csiArchive.Txt", False
                lacMsg.Caption = "User Terminate Archive"
                DoEvents
            Else
                If (ilLcfError) Or (ilSsfError) Or (ilSdfError) Or (ilPsfError) Then
                    gLogMsg "Traffic Archive cancelled because of an Error", "csiArchive.Txt", False
                    lacMsg.Caption = "Archive Terminated because of an Error"
                    DoEvents
                Else
                    lacTrafficSpotsStatus(1).Caption = "Completed"
                    DoEvents
                End If
            End If
        End If
    Else
        lacTrafficSpotsStatus(1).Caption = "Bypassed, as date not defined"
        lacTrafficSpotsStatus(1).Visible = True
        llEarliestTrafSpot = gGetEarliestTrafSpotDate(hgDBSdf, -1)
    End If
    If (imCancelled) Then
        If imArchiveMethod = 1 Then
            ilRet = csiArchiveDeInit()
        End If
        plcGauge.Visible = False
        imGenArch = False
        mGenerateArchive = False
        Exit Function
    End If
    'Purge Copy
    If smTrafEndDate <> "" Then
        If imNoOutput Then
            lacMsg.Caption = "Removing Copy Rotations"
        Else
            lacMsg.Caption = "Archiving and Removing Copy Rotations"
        End If
        DoEvents
        ReDim tmCifDelete(0 To 0) As CIF
        llEarliestTrafSpot = gGetEarliestTrafSpotDate(hgDBSdf, -1)
        llFirstAllowedCopyDate = llEarliestTrafSpot
        If tmSpf.sGUseAffSys = "Y" Then
            If llAffEarliestAstDate < llFirstAllowedCopyDate Then
                llFirstAllowedCopyDate = llAffEarliestAstDate
            End If
        End If
        ilRet = mGatherAndSortCrf(llFirstAllowedCopyDate, llCrfEarliestDate, llCrfLatestDate, ilMissingDate)
        If (imCancelled) Then
            If imArchiveMethod = 1 Then
                ilRet = csiArchiveDeInit()
            End If
            plcGauge.Visible = False
            imGenArch = False
            mGenerateArchive = False
            Exit Function
        End If
        If UBound(tmCrfArchiveInfo) - 1 > 0 Then
            ArraySortTyp fnAV(tmCrfArchiveInfo(), 0), UBound(tmCrfArchiveInfo), 0, LenB(tmCrfArchiveInfo(0)), 0, LenB(tmCrfArchiveInfo(0).sKey), 0
        End If
        If UBound(tmCrfArchiveInfo) > 0 Then
            If llCrfEarliestDate <= llCrfLatestDate Then
                llCrf = 0
                If ilMissingDate Then
                    slArchiveFolderName = smArchivePath & "\Archive" & "CopyRotNoDate"
                    On Error Resume Next
                    MkDir slArchiveFolderName
                    ilRet = mOpenArchiveFile(slArchiveFolderName & "\CopyRotNoDate" & ".Csv")
                    If Not ilRet Then
                        plcGauge.Visible = False
                        imGenArch = False
                        mGenerateArchive = False
                        Exit Function
                    End If
                    slTitle = "Advertiser,Addr ID,Contract #,Product,Assign Vehicle,Rot Type,Rot #,Start Date,End Date,Start Time,End Time,Days,Len,In/Out,Avail,Comment,Entry Date,Zone,User,Earliest Spot Date,Latest Spot Date,Short Title,Copy Region,Live Copy,MG Assign,Active,Language,Games or Teams,INSTRUCTIONS,Cart#;Product;ISCI;Creative title" & "," & Now
                    mPrintTitle slTitle 'Print #hmArchive, slTitle
                    Do While llCrf < UBound(tmCrfArchiveInfo)
                        If tmCrfArchiveInfo(llCrf).lDate > 0 Then
                            Exit Do
                        End If
                        ilCrfError = Not mArchiveAndRemoveCrf(llCopyStartDate, llCopyEndDate, tmCrfArchiveInfo(llCrf).lCrfCode, llCrfCount)
                        If (imCancelled) Then
                            lacCopyStatus(1).Caption = "Cancelled"
                            lacCopyStatus(1).Visible = True
                            DoEvents
                            Exit Do
                        End If
                        If (ilCrfError) Then
                            lacCopyStatus(1).Caption = "Crf removal error"
                            lacCopyStatus(1).Visible = True
                            DoEvents
                            Exit Do
                        End If
                        llCrf = llCrf + 1
                    Loop
                End If
                mClose  'Close #hmArchive
                gLogMsg Trim$(Str$(llCrfCount)) & " Copy Rotations Archived and Removed for month " & "No Dates", "csiArchive.Txt", False
                llTotalCrfCount = llTotalCrfCount + llCrfCount
                If (Not imCancelled) And (Not ilCrfError) Then
                    llTotalsNoDays = llCrfLatestDate - llCrfEarliestDate + 1
                    llNoDays = 0
                    gLogMsg "Copy Archive date range- From: " & Format$(llCrfEarliestDate, "m/d/yy") & " to " & Format$(llCrfLatestDate, "m/d/yy"), "csiArchive.Txt", False
                    slDate = gObtainEndStd(Format$(llCrfEarliestDate, "m/d/yy"))
                    ilNoMonths = DateDiff("m", slDate, Format$(llCrfLatestDate, "m/d/yy")) + 1
                    llCopyStartDate = llCrfEarliestDate
                    llCopyEndDate = gDateValue(slDate)
                    lacCopyStatus(1).Caption = "0" & " of " & Trim$(Str$(ilNoMonths)) & " Months Processed"
                    lacCopyStatus(1).Visible = True
                    DoEvents
                    For ilMonth = 1 To ilNoMonths Step 1
                        llCrfCount = 0
                        DoEvents
                        If imCancelled Then
                            lacTrafficSpotsStatus(1).Caption = "Cancelled"
                            Exit For
                        End If
                        slDate = gObtainEndStd(Format$(llCopyEndDate, "m/d/yy"))
                        slArchiveFolderName = smArchivePath & "\Archive" & Format$(gAdjYear(slDate), "yyyymmdd")
                        On Error Resume Next
                        MkDir slArchiveFolderName
                        ilRet = mOpenArchiveFile(slArchiveFolderName & "\CopyRot" & Format$(gAdjYear(slDate), "yyyymmdd") & ".Csv")
                        If Not ilRet Then
                            plcGauge.Visible = False
                            imGenArch = False
                            mGenerateArchive = False
                            Exit Function
                        End If
                        slTitle = "Advertiser,Addr ID,Contract #,Product,Assign Vehicle,Rot Type,Rot #,Start Date,End Date,Start Time,End Time,Days,Len,In/Out,Avail,Comment,Entry Date,Zone,User,Earliest Spot Date,Latest Spot Date,Short Title,Copy Region,Live Copy,MG Assign,Active,Language,Games or Teams,INSTRUCTIONS,Cart#;Product;ISCI;Creative title" & "," & Now
                        mPrintTitle slTitle 'Print #hmArchive, slTitle
                        'For llDate = llCopyStartDate To llCopyEndDate Step 1
                            'For llCrf = 0 To UBound(tmCrfArchiveInfo) - 1 Step 1
                            Do While llCrf < UBound(tmCrfArchiveInfo)
                                If tmCrfArchiveInfo(llCrf).lDate > llCopyEndDate Then
                                    Exit Do
                                End If
                                ilCrfError = Not mArchiveAndRemoveCrf(llCopyStartDate, llCopyEndDate, tmCrfArchiveInfo(llCrf).lCrfCode, llCrfCount)
                                If (imCancelled) Then
                                    lacCopyStatus(1).Caption = "Cancelled"
                                    DoEvents
                                    Exit Do
                                End If
                                If (ilCrfError) Then
                                    lacCopyStatus(1).Caption = "Crf removal error"
                                    DoEvents
                                    Exit Do
                                End If
                                llCrf = llCrf + 1
                            'Next llCrf
                            Loop
                            If (imCancelled) Then
                                lacTrafficSpotsStatus(1).Caption = "Cancelled"
                                DoEvents
                                Exit For
                            End If
                            If (ilCrfError) Then
                                lacCopyStatus(1).Caption = "Crf removal error"
                                DoEvents
                                Exit For
                            End If
                        'Next llDate
                        mClose  'Close #hmArchive
                        gLogMsg Trim$(Str$(llCrfCount)) & " Copy Rotations Archived and Removed for month " & Format$(llCopyEndDate, "m/d/yy"), "csiArchive.Txt", False
                        llTotalCrfCount = llTotalCrfCount + llCrfCount
                        llNoDays = llNoDays + (llCopyEndDate - llCopyStartDate) + 1
                        lacCopyStatus(1).Caption = Trim$(Str$(ilMonth)) & " of " & Trim$(Str$(ilNoMonths)) & " Months Processed"
                        DoEvents
                        llPercent = ((CLng(llNoDays) + 1) * CSng(100)) / llTotalsNoDays
                        If llPercent >= 100 Then
                            llPercent = 100
                        End If
                        plcGauge.Value = llPercent
                        DoEvents
                        If (imCancelled) Then
                            Exit For
                        End If
                        If (ilCrfError) Then
                            Exit For
                        End If
                        llCopyStartDate = llCopyEndDate + 1
                        llCopyEndDate = gDateValue(gObtainEndStd(Format$(llCopyStartDate, "m/d/yy")))
                    Next ilMonth
                    If ilNoMonths > 1 Then
                        gLogMsg "Total of " & Trim$(Str$(llTotalCrfCount)) & " Copy Rotations Archived and Removed through " & Format$(llCopyStartDate - 1, "m/d/yy"), "csiArchive.Txt", False
                    End If
                    If Not imCancelled Then
                        ilRet = mRemoveCif(llCrfLatestDate, llTotalCifCount)
                        gLogMsg "Total of " & Trim$(Str$(llTotalCifCount)) & " Copy Instructions Removed through " & Format$(llCrfLatestDate - 1, "m/d/yy"), "csiArchive.Txt", False
                    End If
                End If
                If imCancelled Then
                    gLogMsg "Archive cancelled by User", "csiArchive.Txt", False
                    lacMsg.Caption = "User Terminate Archive"
                    DoEvents
                Else
                    If (ilCrfError) Then
                        gLogMsg "Copy Archive cancelled because of an Error", "csiArchive.Txt", False
                        lacMsg.Caption = "Archive Terminated because of an Error"
                        DoEvents
                    Else
                        lacCopyStatus(1).Caption = "Completed"
                        DoEvents
                    End If
                End If
            End If
        End If
        plcGauge.Visible = False
        If (imCancelled) Then
            If imArchiveMethod = 1 Then
                ilRet = csiArchiveDeInit()
            End If
            plcGauge.Visible = False
            imGenArch = False
            mGenerateArchive = False
            Exit Function
        End If
        'Purge Blackout
        lacMsg.Caption = "Removing Old Blackouts"
        DoEvents
        ilRet = mRemoveBof(llFirstAllowedCopyDate, llTotalBofCount)
        gLogMsg "Total of " & Trim$(Str$(llTotalBofCount)) & " Traffic Blackouts Removed through " & Format$(llFirstAllowedCopyDate - 1, "m/d/yy"), "csiArchive.Txt", False
        If (imCancelled) Then
            If imArchiveMethod = 1 Then
                ilRet = csiArchiveDeInit()
            End If
            plcGauge.Visible = False
            imGenArch = False
            mGenerateArchive = False
            Exit Function
        End If
    End If
    'Purge Projection
    If smProjEndDate <> "" Then
        lacMsg.Caption = "Removing Old Projections"
        DoEvents
        ilRet = mRemovePjf(gDateValue(smProjEndDate) + 1, llTotalPjfCount)
        gLogMsg "Total of " & Trim$(Str$(llTotalPjfCount)) & " Traffic Projections Removed through " & Format$(smProjEndDate, "m/d/yy"), "csiArchive.Txt", False
        If (imCancelled) Then
            If imArchiveMethod = 1 Then
                ilRet = csiArchiveDeInit()
            End If
            plcGauge.Visible = False
            imGenArch = False
            mGenerateArchive = False
            Exit Function
        End If
    End If
    'Purge Proposals
    If smPropEndDate <> "" Then
        lacMsg.Caption = "Removing Old Proposals"
        DoEvents
        ilRet = mRemoveProposals(gDateValue(smPropEndDate) + 1, llTotalPropCount)
        gLogMsg "Total of " & Trim$(Str$(llTotalPropCount)) & " Traffic Proposals Removed through " & Format$(smPropEndDate, "m/d/yy"), "csiArchive.Txt", False
        If (imCancelled) Then
            If imArchiveMethod = 1 Then
                ilRet = csiArchiveDeInit()
            End If
            plcGauge.Visible = False
            imGenArch = False
            mGenerateArchive = False
            Exit Function
        End If
    End If
    
    '5/22/18: Moved into th affiliate logic
    'lacMsg.Caption = "Removing Affiliate Export Information"
    'DoEvents
    'ilRet = mRemoveEsf(llTotalEsfCount, llEsfDate)
    'gLogMsg "Total of " & Trim$(Str$(llTotalEsfCount)) & " Affiliate Export Information records Removed through " & Format$(llEsfDate, "m/d/yy"), "csiArchive.Txt", False
    'If (imCancelled) Then
    '    If imArchiveMethod = 1 Then
    '        ilRet = csiArchiveDeInit()
    '    End If
    '    plcGauge.Visible = False
    '    imGenArch = False
    '    mGenerateArchive = False
    '    Exit Function
    'End If
     
    lacMsg.Caption = "Removing Report Queue Information"
    DoEvents
    ilRet = mRemoveRqf(llTotalRqfCount, llRqfDate)
    gLogMsg "Total of " & Trim$(Str$(llTotalRqfCount)) & " Report Queue Information records Removed through " & Format$(llRqfDate, "m/d/yy"), "csiArchive.Txt", False
    If (imCancelled) Then
        If imArchiveMethod = 1 Then
            ilRet = csiArchiveDeInit()
        End If
        plcGauge.Visible = False
        imGenArch = False
        mGenerateArchive = False
        Exit Function
    End If
   
    
    mCleanupUnused llEarliestTrafSpot, llAffEarliestAstDate
    If (imCancelled) Then
        If imArchiveMethod = 1 Then
            ilRet = csiArchiveDeInit()
        End If
        plcGauge.Visible = False
        imGenArch = False
        mGenerateArchive = False
        Exit Function
    End If
    
    'Purge Advertiser and Agencies
    lacMsg.Caption = "Removing Unused Advertisers"
    DoEvents
    llAdfAgfRetainDate = gDateValue(gObtainEndStd(DateAdd("d", -31 * 2 - 7, Format$(Now, "m/d/yy")))) + 1
    ilRet = mRemoveAdf(llAdfAgfRetainDate, llTotalAdfCount)
    gLogMsg "Total of " & Trim$(Str$(llTotalAdfCount)) & " Unused Advertisers Removed through " & Format$(llAdfAgfRetainDate - 1, "m/d/yy"), "csiArchive.Txt", False
    If (imCancelled) Then
        If imArchiveMethod = 1 Then
            ilRet = csiArchiveDeInit()
        End If
        plcGauge.Visible = False
        imGenArch = False
        mGenerateArchive = False
        Exit Function
    End If
    lacMsg.Caption = "Removing Unused Agencies"
    DoEvents
    ilRet = mRemoveAgf(llAdfAgfRetainDate, llTotalAgfCount)
    gLogMsg "Total of " & Trim$(Str$(llTotalAgfCount)) & " Unused Agencies Removed through " & Format$(llAdfAgfRetainDate - 1, "m/d/yy"), "csiArchive.Txt", False
    If (imCancelled) Then
        If imArchiveMethod = 1 Then
            ilRet = csiArchiveDeInit()
        End If
        plcGauge.Visible = False
        imGenArch = False
        mGenerateArchive = False
        Exit Function
    End If
    gPackDate Format$(Now, "m/d/yy"), tmSaf.iLastArchRunDate(0), tmSaf.iLastArchRunDate(1)
    ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
    gLogMsg "Archive Finished", "csiArchive.txt", False
    lacMsg.Caption = "Completed"
    DoEvents
    If imArchiveMethod = 1 Then
        ilRet = csiArchiveDeInit()
    End If
    plcGauge.Visible = False
    imGenArch = False
    mGenerateArchive = True
    Exit Function
End Function

Private Sub mGetAtt()
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim llUpper As Long
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim llDropDate As Long
    Dim llOffAirDate As Long
    Dim llLastAtt As Long
    Dim llAttTest As Long
    Dim llAtt As Long
    Dim llEndDate As Long
    Dim llDate As Long
    Dim llMaxAtt As Long
    Dim llAttTop As Long
    Dim ilFound As Integer
    
    btrExtClear hgDBAtt   'Clear any previous extend operation
    ReDim tmATTDateArchive(0 To 0) As ATTDATEARCHIVE
    ilExtLen = Len(tmATTDateInfo)  'Extract operation record size
    ilRet = btrGetFirst(hgDBAtt, tgDBAtt, igDBAttRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
    Call btrExtSetBounds(hgDBAtt, llNoRec, -1, "UC", "ATTDATEEXTPK", ATTDATEEXTPK) 'Set extract limits (all records)
    gPackDate smAffEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Att", "attOnAir")
    ilRet = btrExtAddLogicConst(hgDBAtt, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
    ilOffSet = gFieldOffset("Att", "attCode")
    ilRet = btrExtAddField(hgDBAtt, ilOffSet, 4) 'Extract start/end time, and days
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Att", "attOnAir")
    ilRet = btrExtAddField(hgDBAtt, ilOffSet, 4) 'Extract start/end time, and days
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Att", "attOffAir")
    ilRet = btrExtAddField(hgDBAtt, ilOffSet, 4) 'Extract start/end time, and days
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Att", "attDropDate")
    ilRet = btrExtAddField(hgDBAtt, ilOffSet, 4) 'Extract start/end time, and days
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Att", "attExportType")
    ilRet = btrExtAddField(hgDBAtt, ilOffSet, 2) 'Extract start/end time, and days
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    llMaxAtt = 0
    llUpper = UBound(tmATTDateArchive)
    ilRet = btrExtGetNext(hgDBAtt, tmATTDateInfo, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        'ilRet = btrExtGetFirst(hmClf, tlClfExt, ilExtLen, llRecPos)
        ilExtLen = Len(tmATTDateInfo)  'Extract operation record size
        If ilRet = BTRV_ERR_REJECT_COUNT Then
            ilRet = btrExtGetNext(hgDBAtt, tmATTDateInfo, ilExtLen, llRecPos)
        End If
        Do While ilRet = BTRV_ERR_NONE
            If tmATTDateInfo.lCode > llMaxAtt Then
                llMaxAtt = tmATTDateInfo.lCode
            End If
            tmATTDateArchive(llUpper).lCode = tmATTDateInfo.lCode
            gUnpackDateLong tmATTDateInfo.iOnAir(0), tmATTDateInfo.iOnAir(1), tmATTDateArchive(llUpper).lStartDate
            gUnpackDateLong tmATTDateInfo.iOffAir(0), tmATTDateInfo.iOffAir(1), llOffAirDate
            gUnpackDateLong tmATTDateInfo.iDropDate(0), tmATTDateInfo.iDropDate(1), llDropDate
            If llDropDate < llOffAirDate Then
                tmATTDateArchive(llUpper).lEndDate = llDropDate
            Else
                tmATTDateArchive(llUpper).lEndDate = llOffAirDate
            End If
            tmATTDateArchive(llUpper).iExportType = tmATTDateInfo.iExportType
            llUpper = UBound(tmATTDateArchive) + 1
            ReDim Preserve tmATTDateArchive(0 To llUpper) As ATTDATEARCHIVE
            ilRet = btrExtGetNext(hgDBAtt, tmATTDateInfo, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hgDBAtt, tmATTDateInfo, ilExtLen, llRecPos)
            Loop
        Loop
    End If
    If UBound(tmATTDateArchive) > 0 Then
        ArraySortTyp fnAV(tmATTDateArchive(), 0), UBound(tmATTDateArchive), 0, LenB(tmATTDateArchive(0)), 0, -2, 0
    End If
    llAttTop = UBound(tmATTDateArchive) - 1
    llLastAtt = 0
    llEndDate = gDateValue(smAffEndDate)
    For llAtt = 0 To llAttTop Step 1
        If llLastAtt + 1 < tmATTDateArchive(llAtt).lCode Then
            For llAttTest = llLastAtt + 1 To tmATTDateArchive(llAtt).lCode - 1 Step 1
                tgDBAstSrchKey1.lAtfCode = llAttTest
                tgDBAstSrchKey1.iFeedDate(0) = 0
                tgDBAstSrchKey1.iFeedDate(1) = 0
                ilRet = btrGetGreaterOrEqual(hgDBAst, tgDBAst, igDBAstRecLen, tgDBAstSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get lSdf current record to obtain date
                If ilRet = BTRV_ERR_NONE Then
                    If tgDBAst.lAtfCode = llAttTest Then
                        gUnpackDateLong tgDBAst.iFeedDate(0), tgDBAst.iFeedDate(1), llDate
                        If llDate <= llEndDate Then
                            llUpper = UBound(tmATTDateArchive)
                            tmATTDateArchive(llUpper).lCode = llAttTest
                            tmATTDateArchive(llUpper).lStartDate = 0
                            tmATTDateArchive(llUpper).lEndDate = 9999999
                            tmATTDateArchive(llUpper).iExportType = 2
                            llUpper = UBound(tmATTDateArchive) + 1
                            ReDim Preserve tmATTDateArchive(0 To llUpper) As ATTDATEARCHIVE
                        End If
                    End If
                End If
            Next llAttTest
        End If
        llLastAtt = tmATTDateArchive(llAtt).lCode
    Next llAtt
    llAtt = llMaxAtt + 1
    Do
        tgDBAstSrchKey1.lAtfCode = llAtt
        tgDBAstSrchKey1.iFeedDate(0) = 0
        tgDBAstSrchKey1.iFeedDate(1) = 0
        ilRet = btrGetGreaterOrEqual(hgDBAst, tgDBAst, igDBAstRecLen, tgDBAstSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get lSdf current record to obtain date
        If ilRet = BTRV_ERR_NONE Then
            gUnpackDateLong tgDBAst.iFeedDate(0), tgDBAst.iFeedDate(1), llDate
            If llDate <= llEndDate Then
                ilFound = False
                For llAttTest = llAttTop + 1 To UBound(tmATTDateArchive) - 1 Step 1
                    If tgDBAst.lAtfCode = tmATTDateArchive(llAttTest).lCode Then
                        ilFound = True
                        Exit For
                    End If
                Next llAttTest
                If Not ilFound Then
                    llUpper = UBound(tmATTDateArchive)
                    tmATTDateArchive(llUpper).lCode = tgDBAst.lAtfCode
                    tmATTDateArchive(llUpper).lStartDate = 0
                    tmATTDateArchive(llUpper).lEndDate = 9999999
                    tmATTDateArchive(llUpper).iExportType = 2
                    llUpper = UBound(tmATTDateArchive) + 1
                    ReDim Preserve tmATTDateArchive(0 To llUpper) As ATTDATEARCHIVE
                End If
            End If
            llAtt = tgDBAst.lAtfCode + 1
        Else
            Exit Do
        End If

    Loop
End Sub

Private Function mOpenArchiveFile(slFolderFile As String) As Integer
    Dim ilRet As Integer
    
    If imNoOutput Then
        mOpenArchiveFile = True
        Exit Function
    End If
    
    'On Error GoTo mOpenArchiveFileErr:
    'hmArchive = FreeFile
    'Open slFolderFile For Append As hmArchive
    ilRet = gFileOpen(slFolderFile, "Append", hmArchive)
    If ilRet <> 0 Then
        mBtrErrorMsg "mOpenArchiveFile: " & slFolderFile & " Error ", err.Number
        mOpenArchiveFile = False
        Exit Function
    End If
    mOpenArchiveFile = True
    Exit Function
'mOpenArchiveFileErr:
'    mBtrErrorMsg "mOpenArchiveFile: " & slFolderFile & " Error ", Err.Number
'    mOpenArchiveFile = False
'    Exit Function
End Function
Private Sub mPrintTitle(slTitle As String)

    If imNoOutput Then
        Exit Sub
    End If
    
    Print #hmArchive, slTitle
End Sub
Private Sub mClose()
    
    If imNoOutput Then
        Exit Sub
    End If
    
    Close #hmArchive
End Sub


Private Function mArchiveAffiliateSpots(slSource As String) As Integer
    Dim ilRet As Integer
    Dim ilLnVef As Integer
    Dim ilAirVef As Integer
    Dim ilAdf As Integer
    Dim ilAgf As Integer
    Dim slStr As String
    Dim slRecord As String
    Dim ilGetRegionCopy As Integer
    Dim slCopyRegionName As String
    Dim slCart As String
    Dim slISCI As String
    Dim slCopyProduct As String
    Dim slCreativeTitle As String
    Dim slScript As String
    Dim ilVpf As Integer
    Dim slNetRegionName As String
    
    If imNoOutput Then
        mArchiveAffiliateSpots = True
        Exit Function
    End If
    
    'Contract #,Line #,Call Letters,Line Vehicle,Agency,City,Advertiser,Addr ID,Product,Net Region,Feed Date,Feed Time,
    'Pledge Start Date,Pledge Start Time,Pledge End Time,Spot Len,Game #,Copy Region,Cart,Product,ISCI,Creative Title,
    'Audience,Demo,Airing vehicle,Air Date,Air Time,CP Posting Status,Air Status
    If tgDBLst.lCode <> tgDBAst.lLsfCode Then
        tgDBLstSrchKey0.lCode = tgDBAst.lLsfCode
        ilRet = btrGetEqual(hgDBLst, tgDBLst, igDBLstRecLen, tgDBLstSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            'mBtrErrorMsg "mArchiveAffiliateSpots: Get Equal LST error ", ilRet
            'mArchiveAffiliateSpots = False
            'Exit Function
            slSource = "T"
            slRecord = slSource & ","
            tgDBSdfSrchKey3.lCode = tgDBAst.lSdfCode
            ilRet = btrGetEqual(hgDBSdf, tgDBSdf, igDBSdfRecLen, tgDBSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                'mBtrErrorMsg "mArchiveAffiliateSpots: Get Equal SDF error ", ilRet
                'mArchiveAffiliateSpots = True
                'Exit Function
                tgDBLst.lRafCode = 0
                tgDBLst.sSplitNetwork = ""
                tgDBLst.iLnVefCode = 0
                tgDBLst.iAdfCode = 0
                tgDBLst.iAgfCode = 0
                tgDBLst.lCntrNo = 0
                tgDBLst.iLineNo = 0
                tgDBLst.sProd = ""
                tgDBLst.iLen = 0
                tgDBLst.lAud = 0
                tgDBLst.sDemo = ""
                tgDBLst.lGsfCode = 0
                tgDBLst.iLogDate(0) = tgDBAst.iFeedDate(0)
                tgDBLst.iLogDate(1) = tgDBAst.iFeedDate(1)
                tgDBSdf.lChfCode = 0
                tgDBSdf.iGameNo = 0
            Else
                tgDBLst.lRafCode = 0
                tgDBLst.sSplitNetwork = ""
                tgDBLst.iLnVefCode = tgDBSdf.iVefCode
                tgDBLst.iAdfCode = tgDBSdf.iAdfCode
                tgDBLst.iAgfCode = 0
                tgDBLst.lCntrNo = 0
                tgDBLst.iLineNo = tgDBSdf.iLineNo
                tgDBLst.sProd = ""
                tgDBLst.iLen = tgDBSdf.iLen
                tgDBLst.lAud = 0
                tgDBLst.sDemo = ""
                tgDBLst.lGsfCode = 0
                tgDBLst.iLogDate(0) = tgDBAst.iFeedDate(0)
                tgDBLst.iLogDate(1) = tgDBAst.iFeedDate(1)
                tgDBChfSrchKey0.lCode = tgDBSdf.lChfCode
                ilRet = btrGetEqual(hgDBChf, tgDBChf, igDBChfRecLen, tgDBChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    tgDBLst.iAgfCode = tgDBChf.iAgfCode
                    tgDBLst.lCntrNo = tgDBChf.lCntrNo
                    tgDBLst.sProd = tgDBChf.sProduct
                    If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Or (tgDBSdf.sSchStatus = "G") Or (tgDBSdf.sSchStatus = "O") Then
                        tgDBClfSrchKey0.lChfCode = tgDBChf.lCode
                        tgDBClfSrchKey0.iLine = tgDBSdf.iLineNo
                        tgDBClfSrchKey0.iCntRevNo = tgDBChf.iCntRevNo
                        tgDBClfSrchKey0.iPropVer = tgDBChf.iPropVer
                        ilRet = btrGetGreaterOrEqual(hgDBClf, tgDBClf, igDBClfRecLen, tgDBClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        If (ilRet = BTRV_ERR_NONE) And (tgDBClf.lChfCode = tgDBChf.lChfCode) And (tgDBClf.iLine = tgDBSdf.iLineNo) Then
                            tgDBLst.lRafCode = tgDBClf.lRafCode
                            If tgDBLst.lRafCode > 0 Then
                                tgDBLst.sSplitNetwork = "P"
                            End If
                            tgDBLst.iLnVefCode = tgDBClf.iVefCode
                        End If
                    End If
                End If
            End If
        Else
            slSource = "A"
            slRecord = slSource & ","
        End If
    Else
        slRecord = slSource & ","
    End If
    slNetRegionName = ""
    If (tgDBLst.sSplitNetwork = "P") Or (tgDBLst.sSplitNetwork = "S") Then
        tgDBRafSrchKey.lCode = tgDBLst.lRafCode
        ilRet = btrGetEqual(hgDBRaf, tgDBRaf, igDBRafRecLen, tgDBRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            slNetRegionName = Trim$(tgDBRaf.sName)
        End If
    End If
    tgDBShttSrchKey0.iCode = tgDBAst.iShfCode
    ilRet = btrGetEqual(hgDBShtt, tgDBShtt, igDBShttRecLen, tgDBShttSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        'mBtrErrorMsg "mArchiveAffiliateSpots: Get Equal Shtt error ", ilRet
        'mArchiveAffiliateSpots = False
        'Exit Function
        tgDBShtt.sCallLetters = "Missing"
    End If
    ilLnVef = gBinarySearchVef(tgDBLst.iLnVefCode)
    ilAirVef = gBinarySearchVef(tgDBAst.iVefCode)
    ilAdf = gBinarySearchAdf(tgDBLst.iAdfCode)
    If tgDBLst.iAgfCode > 0 Then
        ilAgf = gBinarySearchAgf(tgDBLst.iAgfCode)
    Else
        ilAgf = -1
    End If
    'Contract #
    slRecord = slRecord & Trim$(Str$(tgDBLst.lCntrNo)) & ","
    'Line #
    slRecord = slRecord & Trim$(Str$(tgDBLst.iLineNo)) & ","
    'Call Letters
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgDBShtt.sCallLetters)) & ","
    'Line vehicle
    If ilLnVef >= LBound(tgMVef) Then
        slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgMVef(ilLnVef).sName)) & ","
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("") & ","
    End If
    'Agency and City
    If ilAgf >= LBound(tgCommAgf) Then
        slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgCommAgf(ilAgf).sName)) & "," & gAddQuotesRemoveChars(Trim$(tgCommAgf(ilAgf).sCityID)) & ","
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("") & "," & gAddQuotesRemoveChars("") & ","
    End If
    'Advertiser and ID
    If ilAdf >= LBound(tgCommAdf) Then
        slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgCommAdf(ilAdf).sName)) & "," & gAddQuotesRemoveChars(Trim$(tgCommAdf(ilAdf).sAddrID)) & ","
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("") & "," & gAddQuotesRemoveChars("") & ","
    End If
    'Contract Product
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgDBLst.sProd)) & ","
    'Net Region
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(slNetRegionName)) & ","
    'Feed Date
    gUnpackDate tgDBAst.iFeedDate(0), tgDBAst.iFeedDate(1), slStr
    slRecord = slRecord & Format(slStr, "mm/dd/yyyy") & ","
    'Feed Time
    gUnpackTime tgDBAst.iFeedTime(0), tgDBAst.iFeedTime(1), "A", "1", slStr
    slRecord = slRecord & Format(slStr, "hh:mm:ssam/pm") & ","
    tgDBDatSrchKey0.lCode = tgDBAst.lDatCode
    ilRet = btrGetEqual(hgDBDat, tgDBDat, igDBDatRecLen, tgDBDatSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        'Pledge Start Date
        gUnpackDate tgDBLst.iLogDate(0), tgDBLst.iLogDate(1), slStr
        slRecord = slRecord & Format(slStr, "mm/dd/yyyy") & ","
        'Pledge Start Time
        gUnpackTime tgDBDat.iPdStTime(0), tgDBDat.iPdStTime(1), "A", "1", slStr
        slRecord = slRecord & Format(slStr, "hh:mm:ssam/pm") & ","
        'Pledge End Time
        gUnpackTime tgDBDat.iPdEdTime(0), tgDBDat.iPdEdTime(1), "A", "1", slStr
        slRecord = slRecord & Format(slStr, "hh:mm:ssam/pm") & ","
    Else
        '4/7/15: Handle case were no dat defined. TTP 7476
        tgDBDatSrchKey1.lCode = tgDBAst.lAtfCode
        ilRet = btrGetEqual(hgDBDat, tgDBDat, igDBDatRecLen, tgDBDatSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            'Not Carried
            tgDBDat.iFdStatus = 8
            slRecord = slRecord & ","
            slRecord = slRecord & ","
            slRecord = slRecord & ","
        Else
            'Live
            tgDBDat.iFdStatus = 0
            'Pledge Start Date
            gUnpackDate tgDBAst.iFeedDate(0), tgDBAst.iFeedDate(1), slStr
            slRecord = slRecord & Format(slStr, "mm/dd/yyyy") & ","
            'Pledge Start Time
            gUnpackTime tgDBAst.iFeedTime(0), tgDBAst.iFeedTime(1), "A", "1", slStr
            slRecord = slRecord & Format(slStr, "hh:mm:ssam/pm") & ","
            'Pledge End Time: This is wrong as all spot lengths within the break should be used but this is OK for now
            slRecord = slRecord & Format(gTimeToLong(slStr, False) + tgDBAst.iLen, "hh:mm:ssam/pm") & ","
       End If
    End If
    'Spot Lebgth
    slRecord = slRecord & Trim$(Str$(tgDBLst.iLen)) & ","
    'Game #
    If slSource = "A" Then
        If tgDBLst.lGsfCode > 0 Then
            tgDBGsfSrchKey0.lCode = tgDBLst.lGsfCode
            ilRet = btrGetEqual(hgDBGsf, tgDBGsf, igDBGsfRecLen, tgDBGsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                slRecord = slRecord & Trim$(Str$(tgDBGsf.iGameNo)) & ","
            Else
                slRecord = slRecord & ","
            End If
        Else
            slRecord = slRecord & ","
        End If
        'Get copy stuff prior to region so that reion can replace it is required
        slCart = Trim$(tgDBLst.sCart)
        tgDBCpfSrchKey.lCode = tgDBLst.lcpfCode
        ilRet = btrGetEqual(hgDBCpf, tgDBCpf, igDBCpfRecLen, tgDBCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            slISCI = Trim$(tgDBCpf.sISCI)
            slCopyProduct = Trim$(tgDBCpf.sName)
            slCreativeTitle = Trim$(tgDBCpf.sCreative)
        Else
            slISCI = ""
            slCopyProduct = ""
            slCreativeTitle = ""
        End If
        slScript = ""
        If tgDBLst.lCifCode > 0 Then
            tgDBCifSrchKey.lCode = tgDBLst.lCifCode
            ilRet = btrGetEqual(hgDBCif, tgDBCif, igDBCifRecLen, tgDBCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                If tgDBCif.lCsfCode > 0 Then
                    igDBCsfRecLen = Len(tgDBCsf) '5011
                    tgDBCsfSrchKey.lCode = tgDBCif.lCsfCode
                    ilRet = btrGetEqual(hgDBCsf, tgDBCsf, igDBCsfRecLen, tgDBCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        'slScript = Trim$(Left$(tgDBCsf.sComment, tgDBCsf.iStrLen))
                        slScript = gStripChr0(tgDBCsf.sComment)
                    End If
                End If
            End If
        End If
    Else
        If tgDBSdf.iGameNo > 0 Then
            slRecord = slRecord & Trim$(Str$(tgDBSdf.iGameNo)) & ","
        Else
            slRecord = slRecord & ","
        End If
        slCart = ""
        slISCI = ""
        slCopyProduct = ""
        slCreativeTitle = ""
        slScript = ""
        mGetCopy tgDBSdf.sPtType, tgDBSdf.lCopyCode, slCart, slCopyProduct, slISCI, slCreativeTitle, slScript
    End If
    'Copy Region
    '3/2/15: Added pledge test
    'ilGetRegionCopy = False
    'If ((Asc(tgSpf.sUsingFeatures2) And SPLITCOPY) = SPLITCOPY) Then
    '    If ilAirVef >= LBound(tgMVef) Then
    '        If (tgMVef(ilAirVef).sType = "C") Or (tgMVef(ilAirVef).sType = "A") Or (tgMVef(ilAirVef).sType = "G") Then
    '            ilVpf = gBinarySearchVpf(tgMVef(ilAirVef).iCode)
    '            If ilVpf >= LBound(tgVpf) Then
    '                If tgVpf(ilVpf).sAllowSplitCopy = "Y" Then
    '                    'Get region
    '                    ilGetRegionCopy = True
    '                End If
    '            End If
    '        ElseIf (tgMVef(ilAirVef).sType = "S") Or (tgMVef(ilAirVef).sType = "P") Then
    '            'Get Region copy
    '            ilGetRegionCopy = True
    '        End If
    '    End If
    'End If
    ilGetRegionCopy = False
    If (tgDBDat.iFdStatus <> 2) And (tgDBDat.iFdStatus <> 3) And (tgDBDat.iFdStatus <> 4) And (tgDBDat.iFdStatus <> 5) And (tgDBDat.iFdStatus <> 8) Then
        '3/2/15: Added Live test
        If tgDBAst.lAtfCode <> tgDBAtt.lCode Then
            tgDBAttSrchKey0.lCode = tgDBAst.lAtfCode
            ilRet = btrGetEqual(hgDBAtt, tgDBAtt, igDBAttRecLen, tgDBAttSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                tgDBAtt.sForbidSplitLive = "N"
            End If
        End If
        If (tgDBAtt.sForbidSplitLive <> "Y") Or (tgDBDat.iFdStatus <> 0) Then
            If ((Asc(tgSpf.sUsingFeatures2) And SPLITCOPY) = SPLITCOPY) Then
                If ilAirVef >= LBound(tgMVef) Then
                    If (tgMVef(ilAirVef).sType = "C") Or (tgMVef(ilAirVef).sType = "A") Or (tgMVef(ilAirVef).sType = "G") Then
                        ilVpf = gBinarySearchVpf(tgMVef(ilAirVef).iCode)
                        If ilVpf >= LBound(tgVpf) Then
                            If tgVpf(ilVpf).sAllowSplitCopy = "Y" Then
                                'Get region
                                ilGetRegionCopy = True
                            End If
                        End If
                    ElseIf (tgMVef(ilAirVef).sType = "S") Or (tgMVef(ilAirVef).sType = "P") Then
                        'Get Region copy
                        ilGetRegionCopy = True
                    End If
                End If
            End If
        End If
    End If
    slCopyRegionName = ""
    If ilGetRegionCopy Then
        mGetRegionCopy slCopyRegionName, slCart, slCopyProduct, slISCI, slCreativeTitle, slScript
    End If
    'Region Name
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(slCopyRegionName)) & ","
    'Cart #
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(slCart)) & ","
    'Copy Prduct if Different the contract product
    If Trim$(slCopyProduct) <> "" Then
        If StrComp(Trim$(tgDBLst.sProd), Trim$(slCopyProduct), vbTextCompare) <> 0 Then
            slRecord = slRecord & gAddQuotesRemoveChars(Trim$(slCopyProduct)) & ","
        Else
            slRecord = slRecord & gAddQuotesRemoveChars("") & ","
        End If
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("") & ","
    End If
    'ISCI
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(slISCI)) & ","
    'Creative title
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(slCreativeTitle)) & ","
    'Audience
    If tgDBLst.lAud > 0 Then
        slRecord = slRecord & Trim$(Str$(tgDBLst.lAud)) & ","
    Else
        slRecord = slRecord & ","
    End If
    'Demo
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgDBLst.sDemo)) & ","
    'Airing vehicle
    If ilAirVef >= LBound(tgMVef) Then
        slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgMVef(ilAirVef).sName)) & ","
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("") & ","
    End If
    'Air date
    gUnpackDate tgDBAst.iAirDate(0), tgDBAst.iAirDate(1), slStr
    slRecord = slRecord & Format(slStr, "mm/dd/yyyy") & ","
    'Air Time
    gUnpackTime tgDBAst.iAirTime(0), tgDBAst.iAirTime(1), "A", "1", slStr
    slRecord = slRecord & Format(slStr, "hh:mm:ssam/pm") & ","
    'CP Posting Status
    If tgDBAst.iCPStatus = 0 Then
        slRecord = slRecord & gAddQuotesRemoveChars("Not Posted") & ","
    ElseIf tgDBAst.iCPStatus = 2 Then
        slRecord = slRecord & gAddQuotesRemoveChars("Posted as Not Aired") & ","
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("Posted") & ","
    End If
    'Airing Status
    tgDBAst.iStatus = mGetAirStatus(tgDBAst.iStatus)
    If tgDBAst.iStatus = 20 Or tgDBAst.iStatus = ASTEXTENDED_MG Then
        slRecord = slRecord & gAddQuotesRemoveChars("MG")
    ElseIf tgDBAst.iStatus = 21 Or tgDBAst.iStatus = ASTEXTENDED_BONUS Then
        slRecord = slRecord & gAddQuotesRemoveChars("Bonus")
    ElseIf tgDBAst.iStatus = ASTEXTENDED_REPLACEMENT Then
        slRecord = slRecord & gAddQuotesRemoveChars("Replacememt")
    ElseIf tgDBAst.iStatus = 22 Then
        slRecord = slRecord & gAddQuotesRemoveChars("Missed Part of MG")
    Else
        If tgDBAst.iStatus <= 14 Then
            slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgStatusTypes(tgDBAst.iStatus).sName))
        Else
            slRecord = slRecord & gAddQuotesRemoveChars("Status " & tgDBAst.iStatus)
        End If
    End If
    Print #hmArchive, slRecord
    mArchiveAffiliateSpots = True
End Function

Private Function mArchiveTrafficSpots(slSpotType As String) As Integer
    Dim ilRet As Integer
    Dim ilLnVef As Integer
    Dim ilAirVef As Integer
    Dim ilAdf As Integer
    Dim ilAgf As Integer
    Dim slStr As String
    Dim slRecord As String
    Dim ilGetRegionCopy As Integer
    Dim slCart As String
    Dim slISCI As String
    Dim slCopyProduct As String
    Dim slCreativeTitle As String
    Dim slScript As String
    Dim ilVpf As Integer
    Dim ilUrf As Integer
    Dim ilUrfFound As Integer
    Dim ilCffFound As Integer
    Dim llCffStartDate As Long
    Dim llCffEndDate As Long
    Dim llSdfDate As Long
    Dim slPrice As String
    ReDim ilSdfDate(0 To 1) As Integer
    Dim ilMnf As Integer
    Dim slRegionCopy As String
    Dim slNetRegionName As String
    Dim ilZone As Integer
    Dim slCopy As String
    
    If imNoOutput Then
        mArchiveTrafficSpots = True
        Exit Function
    End If
        
    'Spot/Pkg,Contract #,Line #,Agency,City,Advertiser,Addr ID,Product,Net Region,Line Vehicle,Airing Vehicle,Date,
    'Time,Len,Status,Tracer,PLchg,C/T,%Trade,Actual Price,Rate Card Price,Spot Type,Billed,User,Game #,
    'Miss Date,Miss Time,Reason,COPY,Zone,Cart #,Product,ISCI,Creative,Script,REGION,Copy Region

    tgDBChfSrchKey0.lCode = tgDBSdf.lChfCode
    ilRet = btrGetEqual(hgDBChf, tgDBChf, igDBChfRecLen, tgDBChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        'mBtrErrorMsg "mArchiveTrafficSpots: Get Equal Chf error ", ilRet
        'mArchiveTrafficSpots = True
        'Exit Function
        tgDBChf.lCode = tgDBSdf.lChfCode
        tgDBChf.iPctTrade = 0
        tgDBChf.sProduct = ""
        tgDBChf.iAgfCode = 0
        tgDBChf.lCntrNo = 0
        tgDBChf.iCntRevNo = 32000
        tgDBChf.iPropVer = 32000
    End If
    slNetRegionName = ""
    If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
        tgDBClfSrchKey0.lChfCode = tgDBChf.lCode
        tgDBClfSrchKey0.iLine = tgDBSdf.iLineNo
        tgDBClfSrchKey0.iCntRevNo = tgDBChf.iCntRevNo
        tgDBClfSrchKey0.iPropVer = tgDBChf.iPropVer
        ilRet = btrGetGreaterOrEqual(hgDBClf, tgDBClf, igDBClfRecLen, tgDBClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If (ilRet = BTRV_ERR_NONE) And (tgDBClf.lChfCode = tgDBChf.lChfCode) And (tgDBClf.iLine = tgDBSdf.iLineNo) Then
            If tgDBClf.lRafCode > 0 Then
                tgDBRafSrchKey.lCode = tgDBClf.lRafCode
                ilRet = btrGetEqual(hgDBRaf, tgDBRaf, igDBRafRecLen, tgDBRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    If tgDBRaf.sType = "N" Then
                        slNetRegionName = Trim$(tgDBRaf.sName)
                    End If
                End If
            End If
        End If
    End If
    ilSdfDate(0) = tgDBSdf.iDate(0)
    ilSdfDate(1) = tgDBSdf.iDate(1)
    If tgDBSdf.lSmfCode > 0 Then
        tgDBSmfSrchKey2.lCode = tgDBSdf.lCode
        ilRet = btrGetEqual(hgDBSmf, tgDBSmf, igDBSmfRecLen, tgDBSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilLnVef = gBinarySearchVef(tgDBSmf.iOrigSchVef)
            ilSdfDate(0) = tgDBSmf.iMissedDate(0)
            ilSdfDate(1) = tgDBSmf.iMissedDate(1)
        Else
            ilLnVef = gBinarySearchVef(tgDBSdf.iVefCode)
        End If
    Else
        ilLnVef = gBinarySearchVef(tgDBSdf.iVefCode)
    End If
    ilAirVef = gBinarySearchVef(tgDBSdf.iVefCode)
    ilAdf = gBinarySearchAdf(tgDBSdf.iAdfCode)
    If tgDBChf.iAgfCode > 0 Then
        ilAgf = gBinarySearchAgf(tgDBChf.iAgfCode)
    Else
        ilAgf = -1
    End If
    'Spot/Pkg
    slRecord = slSpotType & ","
    'Contract #
    slRecord = slRecord & Trim$(Str$(tgDBChf.lCntrNo)) & ","
    'Line #
    slRecord = slRecord & Trim$(Str$(tgDBSdf.iLineNo)) & ","
    'Agency and City
    If ilAgf >= LBound(tgCommAgf) Then
        slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgCommAgf(ilAgf).sName)) & "," & gAddQuotesRemoveChars(Trim$(tgCommAgf(ilAgf).sCityID)) & ","
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("") & "," & gAddQuotesRemoveChars("") & ","
    End If
    'Advertiser and ID
    If ilAdf >= LBound(tgCommAdf) Then
        slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgCommAdf(ilAdf).sName)) & "," & gAddQuotesRemoveChars(Trim$(tgCommAdf(ilAdf).sAddrID)) & ","
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("") & "," & gAddQuotesRemoveChars("") & ","
    End If
    'Contract Product
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgDBChf.sProduct)) & ","
    'Net Region
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(slNetRegionName)) & ","
    'Line vehicle
    If ilLnVef >= LBound(tgMVef) Then
        slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgMVef(ilLnVef).sName)) & ","
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("") & ","
    End If
    'Airing vehicle
    If ilAirVef >= LBound(tgMVef) Then
        slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgMVef(ilAirVef).sName)) & ","
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("") & ","
    End If
    If tgDBSdf.sSchStatus <> "M" Then
        'Air date
        gUnpackDate tgDBSdf.iDate(0), tgDBSdf.iDate(1), slStr
        slRecord = slRecord & Format(slStr, "mm/dd/yyyy") & ","
        'Air Time
        gUnpackTime tgDBSdf.iTime(0), tgDBSdf.iTime(1), "A", "1", slStr
        slRecord = slRecord & Format(slStr, "hh:mm:ssam/pm") & ","
    Else
        slRecord = slRecord & "" & "," & "" & ","
    End If
    'Spot Length
    slRecord = slRecord & Trim$(Str$(tgDBSdf.iLen)) & ","
    'Status
    Select Case tgDBSdf.sSchStatus
        Case "S"
            slStr = "Schd"
        Case "M"
            slStr = "Missed"
        Case "G"
            slStr = "MG"
        Case "O"
            slStr = "Outside"
        Case "C"
            slStr = "Cancel"
        Case "H"
            slStr = "Hidden"
        Case Else
            slStr = ""
    End Select
    slRecord = slRecord & gAddQuotesRemoveChars(slStr) & ","
    'Tracer
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgDBSdf.sTracer)) & ","
    'PostLog Chg
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgDBSdf.sAffChg)) & ","
    If tgDBChf.iPctTrade = 0 Then
        slRecord = slRecord & gAddQuotesRemoveChars("C") & "," & "" & ","
    ElseIf tgDBChf.iPctTrade = 100 Then
        slRecord = slRecord & gAddQuotesRemoveChars("T%") & "," & Trim$(Str$(tgDBChf.iPctTrade)) & ","
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("T%") & "," & Trim$(Str$(tgDBChf.iPctTrade)) & ","
    End If
    'Actual Price, RC Price and Spot Type
    If tgDBSdf.sSpotType = "X" Then
        slRecord = slRecord & "" & "," & "" & "," & gAddQuotesRemoveChars("Fill") & ","
    Else
        Select Case tgPriceCff.sPriceType   'tlCff.sPriceType
            Case "N"    'No Charge
                slPrice = "N/C"
            Case "M"    'MG Line
                slPrice = "MG"
            Case "B"    'Bonus
                slPrice = "Bonus"
            Case "S"    'Spinoff
                slPrice = "Spinoff"
            Case "P"    'Package
                slPrice = ".00"
            Case "R"    'Recapturable
                slPrice = "Recapturable"
            Case "A"    'ADU
                slPrice = "ADU"
            Case "X"
                slPrice = "Fill"
            Case Else
                slPrice = ""
        End Select
        ilCffFound = False
        gUnpackDateLong ilSdfDate(0), ilSdfDate(1), llSdfDate
        tgDBCffSrchKey.lChfCode = tgDBChf.lCode
        tgDBCffSrchKey.iClfLine = tgDBSdf.iLineNo
        tgDBCffSrchKey.iCntRevNo = tgDBChf.iCntRevNo
        tgDBCffSrchKey.iPropVer = tgDBChf.iPropVer
        tgDBCffSrchKey.iStartDate(0) = ilSdfDate(0)
        tgDBCffSrchKey.iStartDate(1) = ilSdfDate(1)
        ilRet = btrGetGreaterOrEqual(hgDBCff, tgDBCff, igDBCffRecLen, tgDBCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tgDBCff.lChfCode = tgDBChf.lCode) And (tgDBCff.iClfLine = tgDBSdf.iLineNo)
            gUnpackDateLong tgDBCff.iStartDate(0), tgDBCff.iStartDate(1), llCffStartDate
            gUnpackDateLong tgDBCff.iEndDate(0), tgDBCff.iEndDate(1), llCffEndDate
            If (llCffStartDate <= llSdfDate) And (llCffEndDate >= llSdfDate) Then
                ilCffFound = True
                If slPrice = "" Then
                    slPrice = gLongToStrDec(tgDBCff.lActPrice, 2)
                End If
                Exit Do
            End If
            ilRet = btrGetNext(hgDBCff, tgDBCff, igDBCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If ilCffFound Then
            slRecord = slRecord & slPrice & "," & gLongToStrDec(tgDBCff.lPropPrice, 0) & ","
        Else
            slRecord = slRecord & slPrice & "," & "" & ","
        End If
        If tgDBSdf.sSpotType = "O" Then
            slRecord = slRecord & gAddQuotesRemoveChars("Open BB") & ","
        ElseIf tgDBSdf.sSpotType = "C" Then
            slRecord = slRecord & gAddQuotesRemoveChars("Close BB") & ","
        Else
            slRecord = slRecord & gAddQuotesRemoveChars("") & ","
        End If
    End If
    'Billed
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgDBSdf.sBill)) & ","
    'User
    If tgDBSdf.iUrfCode > 2 Then
        For ilUrf = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
            DoEvents
            If tgPopUrf(ilUrf).iCode = tgDBSdf.iUrfCode Then
                slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgPopUrf(ilUrf).sRept)) & ","
                ilUrfFound = True
                Exit For
            End If
        Next ilUrf
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("Guide") & ","
        ilUrfFound = True
    End If
    If Not ilUrfFound Then
        slRecord = slRecord & gAddQuotesRemoveChars("") & ","
    End If
    'Game Number
    If tgDBSdf.iGameNo > 0 Then
        slRecord = slRecord & Trim$(Str$(tgDBSdf.iGameNo)) & ","
    Else
        slRecord = slRecord & "" & ","
    End If
    'Missed Date, Missed Time, Missed reason
    If tgDBSdf.sSpotType <> "X" Then
        If tgDBSdf.sSchStatus = "M" Then
            gUnpackDate tgDBSdf.iDate(0), tgDBSdf.iDate(1), slStr
            slRecord = slRecord & Format(slStr, "mm/dd/yyyy") & ","
            'Air Time
            gUnpackTime tgDBSdf.iTime(0), tgDBSdf.iTime(1), "A", "1", slStr
            slRecord = slRecord & Format(slStr, "hh:mm:ssam/pm") & ","
            ilMnf = mBinarySearchMRMnf(tgDBSdf.iMnfMissed)
            If ilMnf <> -1 Then
                slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgMRMnf(ilMnf).sName)) & ","
            Else
                slRecord = slRecord & gAddQuotesRemoveChars("") & ","
            End If
        ElseIf (tgDBSdf.sSchStatus = "G") Or (tgDBSdf.sSchStatus = "O") Then
            gUnpackDate tgDBSmf.iMissedDate(0), tgDBSmf.iMissedDate(1), slStr
            slRecord = slRecord & Format(slStr, "mm/dd/yyyy") & ","
            'Air Time
            gUnpackTime tgDBSmf.iMissedTime(0), tgDBSmf.iMissedTime(1), "A", "1", slStr
            slRecord = slRecord & Format(slStr, "hh:mm:ssam/pm") & ","
            ilMnf = mBinarySearchMRMnf(tgDBSdf.iMnfMissed)
            If ilMnf <> -1 Then
                slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgMRMnf(ilMnf).sName)) & ","
            Else
                slRecord = slRecord & gAddQuotesRemoveChars("") & ","
            End If
        Else
            slRecord = slRecord & "" & "," & "" & "," & gAddQuotesRemoveChars("") & ","
        End If
    Else
        slRecord = slRecord & "" & "," & "" & "," & gAddQuotesRemoveChars("") & ","
    End If
    'Region Copy Names
    ilGetRegionCopy = False
    If ((Asc(tgSpf.sUsingFeatures2) And SPLITCOPY) = SPLITCOPY) Then
        If ilAirVef >= LBound(tgMVef) Then
            If (tgMVef(ilAirVef).sType = "C") Or (tgMVef(ilAirVef).sType = "A") Or (tgMVef(ilAirVef).sType = "G") Then
                ilVpf = gBinarySearchVpf(tgMVef(ilAirVef).iCode)
                If ilVpf <> -1 Then
                    If tgVpf(ilVpf).sAllowSplitCopy = "Y" Then
                        'Get region
                        ilGetRegionCopy = True
                    End If
                End If
            ElseIf (tgMVef(ilAirVef).sType = "S") Or (tgMVef(ilAirVef).sType = "P") Then
                'Get Region copy
                ilGetRegionCopy = True
            End If
        End If
    End If
    
    slRegionCopy = ""
    If ilGetRegionCopy Then
        tgDBRsfSrchKey1.lCode = tgDBSdf.lCode
        ilRet = btrGetEqual(hgDBRsf, tgDBRsf, igDBRsfRecLen, tgDBRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        Do While (ilRet = BTRV_ERR_NONE) And (tgDBRsf.lSdfCode = tgDBSdf.lCode)
            If (tgDBRsf.sType <> "B") And (tgDBRsf.sType <> "A") Then
                tgDBRafSrchKey.lCode = tgDBRsf.lRafCode
                ilRet = btrGetEqual(hgDBRaf, tgDBRaf, igDBRafRecLen, tgDBRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    If tgDBRaf.sType = "C" Then     'Split copy
                        slRegionCopy = slRegionCopy & gAddQuotesRemoveChars(Trim$(tgDBRaf.sName)) & ";"
                    End If
                End If
            End If
            ilRet = btrGetNext(hgDBRsf, tgDBRsf, igDBRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    If slRegionCopy <> "" Then
        slRecord = slRecord & gAddQuotesRemoveChars(Left$(slRegionCopy, Len(slRegionCopy) - 1)) & ","
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("") & ","
    End If
    'Copy
    slCopy = ""
    slRecord = slRecord & "COPY" & ","
    slCart = ""
    slISCI = ""
    slCopyProduct = ""
    slCreativeTitle = ""
    slScript = ""
    If tgDBSdf.sPtType = "1" Then
        mGetCopy tgDBSdf.sPtType, tgDBSdf.lCopyCode, slCart, slCopyProduct, slISCI, slCreativeTitle, slScript
        If StrComp(Trim$(tgDBChf.sProduct), Trim$(slCopyProduct), vbTextCompare) = 0 Then
            slCopy = slCopy & "All" & ";" & Trim$(slCart) & ";" & ";" & Trim$(slISCI) & ";" & Trim$(slCreativeTitle) & ";" & Trim$(slScript) & ";"
        Else
            slCopy = slCopy & "All" & ";" & Trim$(slCart) & ";" & Trim$(slCopyProduct) & ";" & Trim$(slISCI) & ";" & Trim$(slCreativeTitle) & ";" & Trim$(slScript) & ";"
        End If
    ElseIf tgDBSdf.sPtType = "3" Then
        tgDBTzfSrchKey.lCode = tgDBSdf.lCopyCode
        ilRet = btrGetEqual(hgDBTzf, tgDBTzf, igDBTzfRecLen, tgDBTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            For ilZone = 1 To 6 Step 1
                If (tgDBTzf.lCifZone(ilZone) > 0) Then
                    slCart = ""
                    slISCI = ""
                    slCopyProduct = ""
                    slCreativeTitle = ""
                    slScript = ""
                    mGetCopy "1", tgDBTzf.lCifZone(ilZone), slCart, slCopyProduct, slISCI, slCreativeTitle, slScript
                    If (slCart <> "") Or (slISCI <> "") Then
                        If StrComp(Trim$(tgDBChf.sProduct), Trim$(slCopyProduct), vbTextCompare) = 0 Then
                            slCopy = slCopy & Trim$(tgDBTzf.sZone(ilZone)) & ";" & Trim$(slCart) & ";" & ";" & Trim$(slISCI) & ";" & Trim$(slCreativeTitle) & ";" & Trim$(slScript) & ";"
                        Else
                            slCopy = slCopy & Trim$(tgDBTzf.sZone(ilZone)) & ";" & Trim$(slCart) & ";" & Trim$(slCopyProduct) & ";" & Trim$(slISCI) & ";" & Trim$(slCreativeTitle) & ";" & Trim$(slScript) & ";"
                        End If
                    End If
                End If
            Next ilZone
        Else
            'slRecord = slRecord & "" & "," & "" & "," & "" & "," & "" & "," & "" & "," & "" & ","
            slCopy = slCopy & ""
        End If
    Else
        'slRecord = slRecord & "" & "," & "" & "," & "" & "," & "" & "," & "" & "," & "" & ","
        slCopy = slCopy & ""
    End If
    'Copy
    If Len(slCopy) > 0 Then
        slCopy = Left$(slCopy, Len(slCopy) - 1)
    End If
    slRecord = slRecord & gAddQuotesRemoveChars(slCopy)
    Print #hmArchive, slRecord
    mArchiveTrafficSpots = True
End Function

Private Sub mSetAffStatus()
'    tgStatusTypes(0).sName = "1-Aired Live"        'In Agreement and Pre_Log use 'Air Live'
'    tgStatusTypes(0).iPledged = 0
'    tgStatusTypes(0).iStatus = 0
'    tgStatusTypes(1).sName = "2-Aired Delay B'cast" '"2-Aired In Daypart"  'In Agreement and Pre-Log use 'Air In Daypart'
'    tgStatusTypes(1).iPledged = 1
'    tgStatusTypes(1).iStatus = 1
'    tgStatusTypes(2).sName = "3-Not Aired Tech Diff"
'    tgStatusTypes(2).iPledged = 2
'    tgStatusTypes(2).iStatus = 2
'    tgStatusTypes(3).sName = "4-Not Aired Blackout"
'    tgStatusTypes(3).iPledged = 2
'    tgStatusTypes(3).iStatus = 3
'    tgStatusTypes(4).sName = "5-Not Aired Other"
'    tgStatusTypes(4).iPledged = 2
'    tgStatusTypes(4).iStatus = 4
'    tgStatusTypes(5).sName = "6-Not Aired Product"
'    tgStatusTypes(5).iPledged = 2
'    tgStatusTypes(5).iStatus = 5
'    tgStatusTypes(6).sName = "7-Aired Outside Pledge"  'In Pre-Log use 'Air-Outside Pledge'
'    tgStatusTypes(6).iPledged = 3
'    tgStatusTypes(6).iStatus = 6
'    tgStatusTypes(7).sName = "8-Aired Not Pledged"  'in Pre-Log use 'Air-Not Pledged'
'    tgStatusTypes(7).iPledged = 3
'    tgStatusTypes(7).iStatus = 7
'    tgStatusTypes(8).sName = "9-Not Carried or Aired"
'    tgStatusTypes(8).iPledged = 2
'    tgStatusTypes(8).iStatus = 8
'    tgStatusTypes(9).sName = "10-Delay Cmml/Prg"  'In Agreement and Pre-Log use 'Air In Daypart'
'    tgStatusTypes(9).iPledged = 1
'    tgStatusTypes(9).iStatus = 9
'    tgStatusTypes(10).sName = "11-Air Cmml Only"  'In Agreement and Pre-Log use 'Air In Daypart'
'    tgStatusTypes(10).iPledged = 1
'    tgStatusTypes(10).iStatus = 10
'    tgStatusTypes(11).sName = "MG"
'    tgStatusTypes(11).iPledged = 3
'    tgStatusTypes(11).iStatus = 20
'    tgStatusTypes(12).sName = "Bonus"
'    tgStatusTypes(12).iPledged = 3
'    tgStatusTypes(12).iStatus = 21
    tgStatusTypes(0).sName = "1-Aired Live"        'In Agreement and Pre_Log use 'Air Live'
    tgStatusTypes(0).iPledged = 0
    tgStatusTypes(0).iStatus = 0
    tgStatusTypes(1).sName = "2-Aired Delay B'cast" '"2-Aired In Daypart"  'In Agreement and Pre-Log use 'Air In Daypart'
    tgStatusTypes(1).iPledged = 1
    tgStatusTypes(1).iStatus = 1
    tgStatusTypes(2).sName = "3-Not Aired Tech Diff"
    tgStatusTypes(2).iPledged = 2
    tgStatusTypes(2).iStatus = 2
    tgStatusTypes(3).sName = "4-Not Aired Blackout"
    tgStatusTypes(3).iPledged = 2
    tgStatusTypes(3).iStatus = 3
    tgStatusTypes(4).sName = "5-Not Aired Other"
    tgStatusTypes(4).iPledged = 2
    tgStatusTypes(4).iStatus = 4
    tgStatusTypes(5).sName = "6-Not Aired Product"
    tgStatusTypes(5).iPledged = 2
    tgStatusTypes(5).iStatus = 5
    tgStatusTypes(6).sName = "7-Aired Outside Pledge"  'In Pre-Log use 'Air-Outside Pledge'
    tgStatusTypes(6).iPledged = 3
    tgStatusTypes(6).iStatus = 6
    tgStatusTypes(7).sName = "8-Aired Not Pledged"  'in Pre-Log use 'Air-Not Pledged'
    tgStatusTypes(7).iPledged = 3
    tgStatusTypes(7).iStatus = 7
    'D.S. 11/6/08 remove the "or Aired" from the status 9 description
    'Affiliate Meeting Decisions item 5) f-iv
    'tgStatusTypes(8).sName = "9-Not Carried or Aired"
    tgStatusTypes(8).sName = "9-Not Carried"
    tgStatusTypes(8).iPledged = 2
    tgStatusTypes(8).iStatus = 8
    tgStatusTypes(9).sName = "10-Delay Cmml/Prg"  'In Agreement and Pre-Log use 'Air In Daypart'
    tgStatusTypes(9).iPledged = 1
    tgStatusTypes(9).iStatus = 9
    tgStatusTypes(10).sName = "11-Air Cmml Only"  'In Agreement and Pre-Log use 'Air In Daypart'
    tgStatusTypes(10).iPledged = 1
    tgStatusTypes(10).iStatus = 10
    tgStatusTypes(ASTEXTENDED_MG).sName = "MG"
    tgStatusTypes(ASTEXTENDED_MG).iPledged = 3
    tgStatusTypes(ASTEXTENDED_MG).iStatus = ASTEXTENDED_MG
    tgStatusTypes(ASTEXTENDED_BONUS).sName = "Bonus"
    tgStatusTypes(ASTEXTENDED_BONUS).iPledged = 3
    tgStatusTypes(ASTEXTENDED_BONUS).iStatus = ASTEXTENDED_BONUS
    tgStatusTypes(ASTEXTENDED_REPLACEMENT).sName = "Replacement"
    tgStatusTypes(ASTEXTENDED_REPLACEMENT).iPledged = 3
    tgStatusTypes(ASTEXTENDED_REPLACEMENT).iStatus = ASTEXTENDED_REPLACEMENT
    tgStatusTypes(ASTAIR_MISSED_MG_BYPASS).sName = "15-Missed MG Bypassed"
    tgStatusTypes(ASTAIR_MISSED_MG_BYPASS).iPledged = 2
    tgStatusTypes(ASTAIR_MISSED_MG_BYPASS).iStatus = ASTAIR_MISSED_MG_BYPASS
End Sub

Private Sub mGetRegionCopy(slRegionName As String, slCartNo As String, slProduct As String, slISCI As String, slCreativeTitle As String, slScript As String)
    Dim ilRet As Integer
    Dim ilMatchFound As Integer
    

'    tgDBRsfSrchKey1.lCode = tgDBAst.lRsfCode 'tgDBLst.lSdfCode
'    ilRet = btrGetEqual(hgDBRsf, tgDBRsf, igDBRsfRecLen, tgDBRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
'    Do While (ilRet = BTRV_ERR_NONE) And (tgDBRsf.lSdfCode = tgDBLst.lSdfCode)
'        If (tgDBRsf.sType <> "B") And (tgDBRsf.sType <> "A") Then
'            tgDBRafSrchKey.lCode = tgDBRsf.lRafCode
'            ilRet = btrGetEqual(hgDBRaf, tgDBRaf, igDBRafRecLen, tgDBRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'            If ilRet = BTRV_ERR_NONE Then
'                If tgDBRaf.sType = "C" Then     'Split copy
'                    tgDBSefSrchKey1.lCode = tgDBRaf.lCode
'                    ilRet = btrGetEqual(hgDBSef, tgDBSef, igDBSefRecLen, tgDBSefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
'                    Do While (ilRet = BTRV_ERR_NONE) And (tgDBSef.lRafCode = tgDBRaf.lCode)
'                        tgDBRafSrchKey.lCode = tgDBRsf.lRafCode
'                        ilRet = btrGetEqual(hgDBRaf, tgDBRaf, igDBRafRecLen, tgDBRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                        If ilRet = BTRV_ERR_NONE Then
'                            Select Case Trim$(tgDBRaf.sCategory)
'                                Case "M"    'Market
'                                    If tgDBSef.iIntCode = tgDBShtt.iMktCode Then
'                                        ilMatchFound = True
'                                        If Trim$(tgDBRaf.sInclExcl) <> "E" Then
'                                            mGetCopy tgDBRsf.sPtType, tgDBRsf.lCopyCode, slCartNo, slProduct, slISCI, slCreativeTitle, slScript
'                                        End If
'                                        slRegionName = Trim$(tgDBRaf.sName)
'                                        Exit Sub
'                                    End If
'                                Case "A"    'MSA Market
'                                    If tgDBSef.iIntCode = tgDBShtt.iMetCode Then
'                                        ilMatchFound = True
'                                        If Trim$(tgDBRaf.sInclExcl) <> "E" Then
'                                            mGetCopy tgDBRsf.sPtType, tgDBRsf.lCopyCode, slCartNo, slProduct, slISCI, slCreativeTitle, slScript
'                                        End If
'                                        slRegionName = Trim$(tgDBRaf.sName)
'                                        Exit Sub
'                                    End If
'                                Case "N"    'State Name
'                                    If StrComp(Trim$(tgDBSef.sName), Trim$(tgDBShtt.sState), vbTextCompare) = 0 Then
'                                        ilMatchFound = True
'                                        If Trim$(tgDBRaf.sInclExcl) <> "E" Then
'                                            mGetCopy tgDBRsf.sPtType, tgDBRsf.lCopyCode, slCartNo, slProduct, slISCI, slCreativeTitle, slScript
'                                        End If
'                                        slRegionName = Trim$(tgDBRaf.sName)
'                                        Exit Sub
'                                    End If
'                                Case "Z"    'Zip Code
'                                    If StrComp(Trim$(tgDBSef.sName), Trim$(tgDBShtt.sZip), vbTextCompare) = 0 Then
'                                        ilMatchFound = True
'                                        If Trim$(tgDBRaf.sInclExcl) <> "E" Then
'                                            mGetCopy tgDBRsf.sPtType, tgDBRsf.lCopyCode, slCartNo, slProduct, slISCI, slCreativeTitle, slScript
'                                        End If
'                                        slRegionName = Trim$(tgDBRaf.sName)
'                                        Exit Sub
'                                    End If
'                                'Case "O"    'Owner
'                                '    If tgDBSef.iIntCode = tgDBShtt.iOwnerArttCode Then
'                                '        ilMatchFound = True
'                                '        If Trim$(tgDBRaf.sInclExcl) <> "E" Then
'                                '            mGetCopy tgDBRsf.sPtType, tgDBRsf.lCopyCode, slCartNo, slProduct, slISCI, slCreativeTitle, slScript
'                                '        End If
'                                '        slRegionName = Trim$(tgDBRaf.sName)
'                                '        Exit Sub
'                                '    End If
'                                Case "F"    'Format
'                                    If tgDBSef.iIntCode = tgDBShtt.iFmtCode Then
'                                        ilMatchFound = True
'                                        If Trim$(tgDBRaf.sInclExcl) <> "E" Then
'                                            mGetCopy tgDBRsf.sPtType, tgDBRsf.lCopyCode, slCartNo, slProduct, slISCI, slCreativeTitle, slScript
'                                        End If
'                                        slRegionName = Trim$(tgDBRaf.sName)
'                                        Exit Sub
'                                    End If
'                                Case "S"    'Station
'                                    If tgDBSef.iIntCode = tgDBShtt.iCode Then
'                                        ilMatchFound = True
'                                        If Trim$(tgDBRaf.sInclExcl) <> "E" Then
'                                            mGetCopy tgDBRsf.sPtType, tgDBRsf.lCopyCode, slCartNo, slProduct, slISCI, slCreativeTitle, slScript
'                                        End If
'                                        slRegionName = Trim$(tgDBRaf.sName)
'                                        Exit Sub
'                                    End If
'                            End Select
'                        End If
'                        ilRet = btrGetNext(hgDBSef, tgDBSef, igDBSefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                    Loop
'                    If (Trim$(tgDBRaf.sInclExcl) = "E") Then
'                        mGetCopy tgDBRsf.sPtType, tgDBRsf.lCopyCode, slCartNo, slProduct, slISCI, slCreativeTitle, slScript
'                        Exit Sub
'                    End If
'                End If
'            End If
'        End If
'        ilRet = btrGetNext(hgDBRsf, tgDBRsf, igDBRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'    Loop

    tgDBRsfSrchKey0.lCode = tgDBAst.lRsfCode
    ilRet = btrGetEqual(hgDBRsf, tgDBRsf, igDBRsfRecLen, tgDBRsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        mGetCopy tgDBRsf.sPtType, tgDBRsf.lCopyCode, slCartNo, slProduct, slISCI, slCreativeTitle, slScript
        tgDBRafSrchKey.lCode = tgDBRsf.lRafCode
        ilRet = btrGetEqual(hgDBRaf, tgDBRaf, igDBRafRecLen, tgDBRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            slRegionName = Trim$(tgDBRaf.sName)
        End If
    End If

End Sub

Private Sub mGetCopy(slPtType As String, llCopyCode As Long, slCartNo As String, slProduct As String, slISCI As String, slCreativeTitle As String, slScript As String)
    Dim ilRet As Integer
    Dim ilMcf As Integer
    Dim slTemp As String
    Dim blOneChar As Byte
    Dim ilLoop As Integer
    Dim ilLen As Integer
    Dim ilActualLen As Integer
    
    If slPtType = "1" Then
        tgDBCifSrchKey.lCode = llCopyCode
        ilRet = btrGetEqual(hgDBCif, tgDBCif, igDBCifRecLen, tgDBCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            slCartNo = Trim$(tgDBCif.sName)
            If tgDBCif.iMcfCode > 0 Then
                ilMcf = mBinarySearchMcf(tgDBCif.iMcfCode)
                If ilMcf <> -1 Then
                    slCartNo = Trim(tgMCF(ilMcf).sName) & slCartNo
                End If
            End If
            tgDBCpfSrchKey.lCode = tgDBCif.lcpfCode
            ilRet = btrGetEqual(hgDBCpf, tgDBCpf, igDBCpfRecLen, tgDBCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                slISCI = Trim$(tgDBCpf.sISCI)
                slProduct = Trim$(tgDBCpf.sName)
                slCreativeTitle = Trim$(tgDBCpf.sCreative)
            Else
                slISCI = ""
                slProduct = ""
                slCreativeTitle = ""
            End If
            slScript = ""
            If tgDBCif.lCsfCode > 0 Then
                igDBCsfRecLen = Len(tgDBCsf) '5011
                tgDBCsfSrchKey.lCode = tgDBCif.lCsfCode
                ilRet = btrGetEqual(hgDBCsf, tgDBCsf, igDBCsfRecLen, tgDBCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    'slScript = Trim$(Left$(tgDBCsf.sComment, tgDBCsf.iStrLen))
                    slScript = gStripChr0(tgDBCsf.sComment)
                    edcRTF.TextRTF = slScript
                    slScript = Trim$(edcRTF.Text)
                End If
            End If
        End If
    End If
End Sub

Private Function mRemoveVcf(ilVefCode As Integer, llDate As Long, llVcfCount As Long) As Integer
    Dim ilDay As Integer
    Dim ilRet As Integer
    Dim llTestDate As Long
    Dim ilPass As Integer
    Dim ilDelete As Integer
    
    For ilPass = 0 To 2 Step 1
        DoEvents
        If ilPass = 0 Then
            ilDay = 0
        ElseIf ilPass = 1 Then
            ilDay = 6
        Else
            ilDay = 7
        End If
        Do
            DoEvents
            tgDBVcfSrchKey0.iSellCode = ilVefCode
            tgDBVcfSrchKey0.iSellDay = ilDay
            tgDBVcfSrchKey0.iEffDate(0) = 0
            tgDBVcfSrchKey0.iEffDate(1) = 0
            tgDBVcfSrchKey0.iSellTime(0) = 0
            tgDBVcfSrchKey0.iSellTime(1) = 0
            tgDBVcfSrchKey0.iSellPosNo = 0
            ilRet = btrGetGreaterOrEqual(hgDBVcf, tgDBVcf, igDBVcfRecLen, tgDBVcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            ilDelete = False
            If tgDBVcf.iSellCode <> ilVefCode Then
                Exit Do
            End If
            If ilDay <> tgDBVcf.iSellDay Then
                Exit Do
            End If
            gUnpackDateLong tgDBVcf.iTermDate(0), tgDBVcf.iTermDate(1), llTestDate
            If llTestDate = 0 Then
                Exit Do
            End If
            If llTestDate > llDate Then
                Exit Do
            End If
            ilRet = btrDelete(hgDBVcf)
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet <> BTRV_ERR_CONFLICT Then
                    mBtrErrorMsg "mRemoveVCF: Delete error ", ilRet
                    mRemoveVcf = False
                    Exit Function
                End If
            Else
                llVcfCount = llVcfCount + 1
            End If
        Loop
    Next ilPass
    mRemoveVcf = True
End Function

Private Function mRemoveVlf(ilVefCode As Integer, llDate As Long, llVlfCount As Long) As Integer
    Dim ilDay As Integer
    Dim ilPass As Integer
    Dim ilRet As Integer
    Dim llTestDate As Long
    
    For ilPass = 0 To 2 Step 1
        DoEvents
        If ilPass = 0 Then
            ilDay = 0
        ElseIf ilPass = 1 Then
            ilDay = 6
        Else
            ilDay = 7
        End If
        Do
            DoEvents
            tgDBVlfSrchKey0.iSellCode = ilVefCode
            tgDBVlfSrchKey0.iSellDay = ilDay
            tgDBVlfSrchKey0.iEffDate(0) = 0
            tgDBVlfSrchKey0.iEffDate(1) = 0
            tgDBVlfSrchKey0.iSellTime(0) = 0
            tgDBVlfSrchKey0.iSellTime(1) = 0
            tgDBVlfSrchKey0.iSellPosNo = 0
            tgDBVlfSrchKey0.iSellSeq = 0
            'Remove Vlf
            ilRet = btrGetGreaterOrEqual(hgDBVlf, tgDBVlf, igDBVlfRecLen, tgDBVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            If tgDBVlf.iSellCode <> ilVefCode Then
                Exit Do
            End If
            If ilDay <> tgDBVlf.iSellDay Then
                Exit Do
            End If
            gUnpackDateLong tgDBVlf.iTermDate(0), tgDBVlf.iTermDate(1), llTestDate
            If llTestDate = 0 Then
                Exit Do
            End If
            If llTestDate > llDate Then
                Exit Do
            End If
            ilRet = btrDelete(hgDBVlf)
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet <> BTRV_ERR_CONFLICT Then
                    mBtrErrorMsg "mRemoveVLF: Delete error ", ilRet
                    mRemoveVlf = False
                    Exit Function
                End If
            Else
                llVlfCount = llVlfCount + 1
            End If
        Loop
    Next ilPass
    mRemoveVlf = True
End Function

Private Function mRemoveDlf(hlFile As Integer, ilVefCode As Integer, llDate As Long, llDlfCount As Long) As Integer
    Dim ilDay As Integer
    Dim ilPass As Integer
    Dim ilRet As Integer
    Dim llTestDate As Long
    
    For ilPass = 0 To 2 Step 1
        DoEvents
        If ilPass = 0 Then
            ilDay = 0
        ElseIf ilPass = 1 Then
            ilDay = 6
        Else
            ilDay = 7
        End If
        Do
            DoEvents
            tgDBDlfSrchKey0.iVefCode = ilVefCode
            tgDBDlfSrchKey0.sAirDay = Trim$(Str$(ilDay))
            tgDBDlfSrchKey0.iStartDate(0) = 0
            tgDBDlfSrchKey0.iStartDate(1) = 0
            tgDBDlfSrchKey0.iAirTime(0) = 0
            tgDBDlfSrchKey0.iAirTime(1) = 0
            'Remove Dlf
            ilRet = btrGetGreaterOrEqual(hlFile, tgDBDlf, igDBDlfRecLen, tgDBDlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            If tgDBDlf.iVefCode <> ilVefCode Then
                Exit Do
            End If
            If Trim$(Str$(ilDay)) <> tgDBDlf.sAirDay Then
                Exit Do
            End If
            gUnpackDateLong tgDBDlf.iTermDate(0), tgDBDlf.iTermDate(1), llTestDate
            If llTestDate = 0 Then
                Exit Do
            End If
            If llTestDate > llDate Then
                Exit Do
            End If
            ilRet = btrDelete(hlFile)
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet <> BTRV_ERR_CONFLICT Then
                    mBtrErrorMsg "mRemoveDLF: Delete error ", ilRet
                    mRemoveDlf = False
                    Exit Function
                End If
            Else
                llDlfCount = llDlfCount + 1
            End If
        Loop
    Next ilPass
    mRemoveDlf = True
End Function


Private Function mRemoveLcf(ilVefCode As Integer, llInDate As Long, llLcfCount As Long) As Integer
    Dim ilRet As Integer
    Dim llDate As Long
    Dim ilVef As Integer
    Dim ilMatchDay As Integer
    
    llDate = llInDate
    'If air vehicle retain date just prior to delete date
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If tgMVef(ilVef).iCode = ilVefCode Then
            If tgMVef(ilVef).sType = "A" Then
                igDBLcfRecLen = Len(tgDBLcf)
                tgDBLcfSrchKey2.iVefCode = ilVefCode
                gPackDateLong llDate, tgDBLcfSrchKey2.iLogDate(0), tgDBLcfSrchKey2.iLogDate(1)
                ilRet = btrGetLessOrEqual(hgDBLcf, tgDBLcf, igDBLcfRecLen, tgDBLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
                If ilRet = BTRV_ERR_NONE Then
                    If tgDBLcf.iVefCode = ilVefCode Then
                        If tgDBLcf.iLogDate(1) <> 0 Then
                            'Back up to Sunday
                            gUnpackDateLong tgDBLcf.iLogDate(0), tgDBLcf.iLogDate(1), llDate
                            ilMatchDay = 0
                            Do Until gWeekDayLong(llDate) = ilMatchDay
                                llDate = llDate - 1
                            Loop
                            llDate = llDate - 1
                        End If
                    End If
                End If
            End If
            Exit For
        End If
    Next ilVef
    Do
        igDBLcfRecLen = Len(tgDBLcf)
        tgDBLcfSrchKey2.iVefCode = ilVefCode
        gPackDateLong llDate, tgDBLcfSrchKey2.iLogDate(0), tgDBLcfSrchKey2.iLogDate(1)
        ilRet = btrGetLessOrEqual(hgDBLcf, tgDBLcf, igDBLcfRecLen, tgDBLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        If tgDBLcf.iVefCode <> ilVefCode Then
            Exit Do
        End If
        'Retain TFN: LogDate(0) from 1 to 7 and LogDate(1) = 0
        If tgDBLcf.iLogDate(1) = 0 Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBLcf)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveLCF: Delete error ", ilRet
                mRemoveLcf = False
                Exit Function
            End If
        Else
            llLcfCount = llLcfCount + 1
        End If
    Loop
    mRemoveLcf = True
End Function

Private Function mRemoveSsf(ilVefCode As Integer, llInDate As Long, llSsfCount As Long) As Integer
    Dim ilRet As Integer
    Dim llDate As Long
    Dim ilVef As Integer
    Dim ilMatchDay As Integer
    
    llDate = llInDate
    'If air vehicle retain date just prior to delete date
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If tgMVef(ilVef).iCode = ilVefCode Then
            If tgMVef(ilVef).sType = "A" Then
                igDBSsfRecLen = Len(tgDBSsf)
                tgDBSsfSrchKey2.iVefCode = ilVefCode
                gPackDateLong llDate, tgDBSsfSrchKey2.iDate(0), tgDBSsfSrchKey2.iDate(1)
                ilRet = btrGetLessOrEqual(hgDBSsf, tgDBSsf, igDBSsfRecLen, tgDBSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
                If ilRet = BTRV_ERR_NONE Then
                    If tgDBSsf.iVefCode = ilVefCode Then
                        If tgDBSsf.iDate(1) <> 0 Then
                            'Back up to Sunday
                            gUnpackDateLong tgDBSsf.iDate(0), tgDBSsf.iDate(1), llDate
                            ilMatchDay = 0
                            Do Until gWeekDayLong(llDate) = ilMatchDay
                                llDate = llDate - 1
                            Loop
                            llDate = llDate - 1
                        End If
                    End If
                End If
            End If
            Exit For
        End If
    Next ilVef
    
    Do
        igDBSsfRecLen = Len(tgDBSsf)
        tgDBSsfSrchKey2.iVefCode = ilVefCode
        gPackDateLong llDate, tgDBSsfSrchKey2.iDate(0), tgDBSsfSrchKey2.iDate(1)
        ilRet = btrGetLessOrEqual(hgDBSsf, tgDBSsf, igDBSsfRecLen, tgDBSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        If tgDBSsf.iVefCode <> ilVefCode Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBSsf)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveSSF: Delete error ", ilRet
                mRemoveSsf = False
                Exit Function
            End If
        Else
            llSsfCount = llSsfCount + 1
        End If
    Loop
    mRemoveSsf = True
End Function


Private Function mRemoveAtt(llRetainDate As Long, llAttCount As Long) As Integer
    Dim llAtt As Long
    Dim llDropDate As Long
    Dim llOffAirDate As Long
    Dim ilRet As Integer
    
    llAtt = 0
    Do
        If tmATTDateArchive(llAtt).lEndDate < llRetainDate Then
            ilRet = mRemoveDat(tmATTDateArchive(llAtt).lCode)
            ilRet = mRemoveEpt(tmATTDateArchive(llAtt).lCode)
            ilRet = mRemovePet(tmATTDateArchive(llAtt).lCode)
            ilRet = mRemoveVat(tmATTDateArchive(llAtt).lCode)
            tgDBAttSrchKey0.lCode = tmATTDateArchive(llAtt).lCode
            ilRet = btrGetEqual(hgDBAtt, tgDBAtt, igDBAttRecLen, tgDBAttSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hgDBAtt)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        mBtrErrorMsg "mRemoveATT: Delete error ", ilRet
                        mRemoveAtt = False
                        Exit Function
                    End If
                Else
                    llAttCount = llAttCount + 1
                    llAtt = llAtt + 1
                End If
            Else
                llAtt = llAtt + 1
            End If
        Else
            llAtt = llAtt + 1
        End If
    Loop While llAtt < UBound(tmATTDateArchive)
    mRemoveAtt = True
End Function

Private Function mRemoveLst(ilVefCode As Integer, llRetainDate As Long, llLstCount As Long) As Integer
    Dim ilRet As Integer
    
    Do
        tgDBLstSrchKey2.iLogVefCode = ilVefCode
        gPackDateLong llRetainDate - 1, tgDBLstSrchKey2.iLogDate(0), tgDBLstSrchKey2.iLogDate(1)
        ilRet = btrGetLessOrEqual(hgDBLst, tgDBLst, igDBLstRecLen, tgDBLstSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        If tgDBLst.iLogVefCode <> ilVefCode Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBLst)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveLST: Delete error ", ilRet
                mRemoveLst = False
                Exit Function
            End If
        Else
            llLstCount = llLstCount + 1
        End If
    Loop
    mRemoveLst = True
End Function

Private Function mRemoveCptt(ilVefCode As Integer, llRetainDate As Long, llCpttCount As Long) As Integer
    Dim ilRet As Integer
    
    Do
        tgDBCpttSrchKey1.iVefCode = ilVefCode
        gPackDateLong llRetainDate - 1, tgDBCpttSrchKey1.iStartDate(0), tgDBCpttSrchKey1.iStartDate(1)
        ilRet = btrGetGreaterOrEqual(hgDBCptt, tgDBCptt, igDBCpttRecLen, tgDBCpttSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        If tgDBCptt.iVefCode <> ilVefCode Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBCptt)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveCPTT: Delete error ", ilRet
                mRemoveCptt = False
                Exit Function
            End If
        Else
            llCpttCount = llCpttCount + 1
        End If
    Loop
    mRemoveCptt = True
End Function

Private Function mArchiveAndRemoveAst(llAttCode As Long, llDate As Long, ilExportType As Integer, llLstAstCount As Long, llSdfAstCount As Long) As Integer
    Dim ilRet As Integer
    Dim llTestDate As Long
    Dim slSource As String
    
    tgDBLst.lCode = 0
    Do
        tgDBAstSrchKey1.lAtfCode = llAttCode
        gPackDateLong llDate, tgDBAstSrchKey1.iFeedDate(0), tgDBAstSrchKey1.iFeedDate(1)
        ilRet = btrGetEqual(hgDBAst, tgDBAst, igDBAstRecLen, tgDBAstSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        'gUnpackDateLong tgDBAst.iFeedDate(0), tgDBAst.iFeedDate(1), llTestDate
        'If llTestDate <> llDate Then
        '    Exit Do
        'End If
        'Produce Archive Record
        ilRet = mArchiveAffiliateSpots(slSource)
        If Not ilRet Then
            mArchiveAndRemoveAst = False
            Exit Function
        End If
        If ilExportType = 2 Then
            tgDBAetSrchKey2.lCode = tgDBAst.lCode
            ilRet = btrGetEqual(hgDBAet, tgDBAet, igDBAetRecLen, tgDBAetSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hgDBAet)
            End If
        End If
        ilRet = btrDelete(hgDBAst)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mArchiveAndRemoveAst: Delete AST error ", ilRet
                mArchiveAndRemoveAst = False
                Exit Function
            End If
        Else
            If slSource = "A" Then
                llLstAstCount = llLstAstCount + 1
            Else
                llSdfAstCount = llSdfAstCount + 1
            End If
        End If
    Loop
    mArchiveAndRemoveAst = True
End Function

Private Function mArchiveAndRemoveSdf(slType As String, llDate As Long, ilCount As Long) As Integer
    Dim ilRet As Integer
    Dim llTestDate As Long
    
    Do
        gPackDateLong llDate, tgDBSdfSrchKey4.iDate(0), tgDBSdfSrchKey4.iDate(1)
        tgDBSdfSrchKey4.lChfCode = 0
        If slType = "S" Then
            ilRet = btrGetGreaterOrEqual(hgDBSdf, tgDBSdf, igDBSdfRecLen, tgDBSdfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get lSdf current record to obtain date
        Else
            ilRet = btrGetGreaterOrEqual(hgDBPsf, tgDBSdf, igDBSdfRecLen, tgDBSdfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get lSdf current record to obtain date
        End If
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        gUnpackDateLong tgDBSdf.iDate(0), tgDBSdf.iDate(1), llTestDate
        If llTestDate <> llDate Then
            Exit Do
        End If
        'Produce Archive Record
        ilRet = mArchiveTrafficSpots(slType)
        If Not ilRet Then
            mArchiveAndRemoveSdf = False
            Exit Function
        End If
        If tgDBSdf.lSmfCode > 0 Then
            tgDBSmfSrchKey2.lCode = tgDBSdf.lCode
            ilRet = btrGetEqual(hgDBSmf, tgDBSmf, igDBSmfRecLen, tgDBSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hgDBSmf)
            End If
        End If
        If tgDBSdf.sPtType = "3" Then
            tgDBTzfSrchKey.lCode = tgDBSdf.lCopyCode
            ilRet = btrGetEqual(hgDBTzf, tgDBTzf, igDBTzfRecLen, tgDBTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hgDBTzf)
            End If
        End If
        tgDBRsfSrchKey1.lCode = tgDBSdf.lCode
        ilRet = btrGetEqual(hgDBRsf, tgDBRsf, igDBRsfRecLen, tgDBRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hgDBRsf)
        End If
        If slType = "S" Then
            ilRet = btrDelete(hgDBSdf)
        Else
            ilRet = btrDelete(hgDBPsf)
        End If
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                If slType = "S" Then
                    mBtrErrorMsg "mArchiveAndRemoveSDF: Delete SDF error ", ilRet
                Else
                    mBtrErrorMsg "mArchiveAndRemoveSDF: Delete PSF error ", ilRet
                End If
                mArchiveAndRemoveSdf = False
                Exit Function
            End If
        Else
            ilCount = ilCount + 1
        End If
    Loop
    mArchiveAndRemoveSdf = True
End Function
Private Function mGetEarliestLstDate(ilVefCode As Integer) As Long
    Dim ilRet As Integer
    Dim llDate As Long
    Dim llTestDate As Long
    Dim ilVef As Integer
    
    llDate = -1
    If ilVefCode = -1 Then
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            DoEvents
            tgDBLstSrchKey2.iLogVefCode = tgMVef(ilVef).iCode
            tgDBLstSrchKey2.iLogDate(0) = 257
            tgDBLstSrchKey2.iLogDate(1) = 1900
            ilRet = btrGetGreaterOrEqual(hgDBLst, tgDBLst, igDBLstRecLen, tgDBLstSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) And (tgDBLst.iLogVefCode = tgMVef(ilVef).iCode) Then
                gUnpackDateLong tgDBLst.iLogDate(0), tgDBLst.iLogDate(1), llTestDate
                If llDate = -1 Then
                    llDate = llTestDate
                Else
                    If llTestDate < llDate Then
                        llDate = llTestDate
                    End If
                End If
            End If
        Next ilVef
    Else
        tgDBLstSrchKey2.iLogVefCode = ilVefCode
        tgDBLstSrchKey2.iLogDate(0) = 257
        tgDBLstSrchKey2.iLogDate(1) = 1900
        ilRet = btrGetGreaterOrEqual(hgDBLst, tgDBLst, igDBLstRecLen, tgDBLstSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If (ilRet = BTRV_ERR_NONE) And (tgDBLst.iLogVefCode = ilVefCode) Then
            gUnpackDateLong tgDBLst.iLogDate(0), tgDBLst.iLogDate(1), llTestDate
            If llDate = -1 Then
                llDate = llTestDate
            Else
                If llTestDate < llDate Then
                    llDate = llTestDate
                End If
            End If
        End If
    End If
    mGetEarliestLstDate = llDate
End Function

Private Function mGetEarliestCpttDate(ilVefCode As Integer) As Long
    Dim ilRet As Integer
    Dim llDate As Long
    Dim llTestDate As Long
    Dim ilVef As Integer
    
    llDate = -1
    If ilVefCode = -1 Then
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            DoEvents
            tgDBCpttSrchKey1.iVefCode = tgMVef(ilVef).iCode
            tgDBCpttSrchKey1.iStartDate(0) = 257
            tgDBCpttSrchKey1.iStartDate(1) = 1900
            ilRet = btrGetLessOrEqual(hgDBCptt, tgDBCptt, igDBCpttRecLen, tgDBCpttSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) And (tgDBCptt.iVefCode = tgMVef(ilVef).iCode) Then
                gUnpackDateLong tgDBCptt.iStartDate(0), tgDBCptt.iStartDate(1), llTestDate
                If llDate = -1 Then
                    llDate = llTestDate
                Else
                    If llTestDate < llDate Then
                        llDate = llTestDate
                    End If
                End If
            End If
        Next ilVef
    Else
        tgDBCpttSrchKey1.iVefCode = ilVefCode
        tgDBCpttSrchKey1.iStartDate(0) = 257
        tgDBCpttSrchKey1.iStartDate(1) = 1900
        ilRet = btrGetGreaterOrEqual(hgDBCptt, tgDBCptt, igDBCpttRecLen, tgDBCpttSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If (ilRet = BTRV_ERR_NONE) And (tgDBCptt.iVefCode = ilVefCode) Then
            gUnpackDateLong tgDBCptt.iStartDate(0), tgDBCptt.iStartDate(1), llTestDate
            If llDate = -1 Then
                llDate = llTestDate
            Else
                If llTestDate < llDate Then
                    llDate = llTestDate
                End If
            End If
        End If
    End If
    mGetEarliestCpttDate = llDate
End Function

Private Function mRemoveStf(ilVefCode As Integer, llDate As Long, llStfCount As Long) As Integer
    Dim ilRet As Integer
    Dim llTestDate As Long
    
    Do
        tgDBStfSrchKey0.iVefCode = ilVefCode
        gPackDateLong llDate, tgDBStfSrchKey0.iLogDate(0), tgDBStfSrchKey0.iLogDate(1)
        gPackTime "11:59:59pm", tgDBStfSrchKey0.iLogTime(0), tgDBStfSrchKey0.iLogTime(1)
        ilRet = btrGetLessOrEqual(hgDBStf, tgDBStf, igDBStfRecLen, tgDBStfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        If tgDBStf.iVefCode <> ilVefCode Then
            Exit Do
        End If
        'gUnpackDateLong tgDBStf.iLogDate(0), tgDBStf.iLogDate(1), llTestDate
        'If llTestDate <> llDate Then
        '    Exit Do
        'End If
        ilRet = btrDelete(hgDBStf)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveSTF: Delete error ", ilRet
                mRemoveStf = False
                Exit Function
            End If
        Else
            llStfCount = llStfCount + 1
        End If
    Loop
    mRemoveStf = True
End Function

Private Function mRemoveCnf(llCrfCode As Long) As Integer
    Dim ilRet As Integer
    
    Do
        tgDBCnfSrchKey.lCrfCode = llCrfCode
        tgDBCnfSrchKey.iInstrNo = 0
        ilRet = btrGetGreaterOrEqual(hgDBCnf, tgDBCnf, igDBCnfRecLen, tgDBCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        If tgDBCnf.lCrfCode <> llCrfCode Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBCnf)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveCNF: Delete error ", ilRet
                mRemoveCnf = False
                Exit Function
            End If
        End If
    Loop
    mRemoveCnf = True
End Function

Private Function mRemoveCaf(llCrfCode As Long) As Integer
    Dim ilRet As Integer
    
    Do
        tgDBCafSrchKey1.lCrfCode = llCrfCode
        ilRet = btrGetGreaterOrEqual(hgDBCaf, tgDBCaf, igDBCafRecLen, tgDBCafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        If tgDBCaf.lCrfCode <> llCrfCode Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBCaf)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveCaf: Delete error ", ilRet
                mRemoveCaf = False
                Exit Function
            End If
        End If
    Loop
    mRemoveCaf = True
End Function


Private Sub mSplitVefCodes()
    Dim ilVef As Integer
    Dim ilRet As Integer
    
    'ReDim tmCASGVVefCode(0 To 0) As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
    ReDim tmLstVefCode(0 To 0) As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
    ReDim tmCpttVefCode(0 To 0) As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
    ReDim tmLcfVefCode(0 To 0) As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
    ReDim tmSsfVefCode(0 To 0) As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
    ReDim tmVcfVefCode(0 To 0) As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
    ReDim tmVlfVefCode(0 To 0) As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
    ReDim tmDlfVefCode(0 To 0) As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
    ReDim tmEgfVefCode(0 To 0) As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
    ReDim tmStfVefCode(0 To 0) As DBVEHICLE   'Vehicle to be checked (C=Conventional; A=Airing; S=Selling)
    ReDim tmPVefCode(0 To 0) As DBVEHICLE     'Package vehicle
    ReDim tmRVefCode(0 To 0) As DBVEHICLE     'Rep vehicles
    
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        'If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "S") Or (tgMVef(ilVef).sType = "A") Or (tgMVef(ilVef).sType = "G") Or (tgMVef(ilVef).sType = "V") Then
        '    tmCASGVVefCode(UBound(tmCASGVVefCode)).sType = tgMVef(ilVef).sType
        '    tmCASGVVefCode(UBound(tmCASGVVefCode)).iVefCode = tgMVef(ilVef).iCode
        '    ReDim Preserve tmCASGVVefCode(0 To UBound(tmCASGVVefCode) + 1) As DBVEHICLE
        'ElseIf tgMVef(ilVef).sType = "P" Then
        If tgMVef(ilVef).sType = "P" Then
            tmPVefCode(UBound(tmPVefCode)).sType = tgMVef(ilVef).sType
            tmPVefCode(UBound(tmPVefCode)).iVefCode = tgMVef(ilVef).iCode
            ReDim Preserve tmPVefCode(0 To UBound(tmPVefCode) + 1) As DBVEHICLE
        ElseIf tgMVef(ilVef).sType = "R" Then
            tmRVefCode(UBound(tmRVefCode)).sType = tgMVef(ilVef).sType
            tmRVefCode(UBound(tmRVefCode)).iVefCode = tgMVef(ilVef).iCode
            ReDim Preserve tmRVefCode(0 To UBound(tmRVefCode) + 1) As DBVEHICLE
        Else
            tgDBLstSrchKey2.iLogVefCode = tgMVef(ilVef).iCode
            tgDBLstSrchKey2.iLogDate(0) = 0
            tgDBLstSrchKey2.iLogDate(1) = 0
            ilRet = btrGetGreaterOrEqual(hgDBLst, tgDBLst, igDBLstRecLen, tgDBLstSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) And (tgDBLst.iLogVefCode = tgMVef(ilVef).iCode) Then
                tmLstVefCode(UBound(tmLstVefCode)).sType = tgMVef(ilVef).sType
                tmLstVefCode(UBound(tmLstVefCode)).iVefCode = tgMVef(ilVef).iCode
                ReDim Preserve tmLstVefCode(0 To UBound(tmLstVefCode) + 1) As DBVEHICLE
            End If
            tgDBCpttSrchKey1.iVefCode = tgMVef(ilVef).iCode
            tgDBCpttSrchKey1.iStartDate(0) = 0
            tgDBCpttSrchKey1.iStartDate(1) = 0
            ilRet = btrGetLessOrEqual(hgDBCptt, tgDBCptt, igDBCpttRecLen, tgDBCpttSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) And (tgDBCptt.iVefCode = tgMVef(ilVef).iCode) Then
                tmCpttVefCode(UBound(tmCpttVefCode)).sType = tgMVef(ilVef).sType
                tmCpttVefCode(UBound(tmCpttVefCode)).iVefCode = tgMVef(ilVef).iCode
                ReDim Preserve tmCpttVefCode(0 To UBound(tmCpttVefCode) + 1) As DBVEHICLE
            End If
            tgDBLcfSrchKey2.iVefCode = tgMVef(ilVef).iCode
            tgDBLcfSrchKey2.iLogDate(0) = 0
            tgDBLcfSrchKey2.iLogDate(1) = 0
            ilRet = btrGetGreaterOrEqual(hgDBLcf, tgDBLcf, igDBLcfRecLen, tgDBLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) And (tgDBLcf.iVefCode = tgMVef(ilVef).iCode) Then
                tmLcfVefCode(UBound(tmLcfVefCode)).sType = tgMVef(ilVef).sType
                tmLcfVefCode(UBound(tmLcfVefCode)).iVefCode = tgMVef(ilVef).iCode
                ReDim Preserve tmLcfVefCode(0 To UBound(tmLcfVefCode) + 1) As DBVEHICLE
            End If
            igDBSsfRecLen = Len(tgDBSsf)
            tgDBSsfSrchKey2.iVefCode = tgMVef(ilVef).iCode
            tgDBSsfSrchKey2.iDate(0) = 0
            tgDBSsfSrchKey2.iDate(1) = 0
            ilRet = btrGetGreaterOrEqual(hgDBSsf, tgDBSsf, igDBSsfRecLen, tgDBSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) And (tgDBSsf.iVefCode = tgMVef(ilVef).iCode) Then
                tmSsfVefCode(UBound(tmSsfVefCode)).sType = tgMVef(ilVef).sType
                tmSsfVefCode(UBound(tmSsfVefCode)).iVefCode = tgMVef(ilVef).iCode
                ReDim Preserve tmSsfVefCode(0 To UBound(tmSsfVefCode) + 1) As DBVEHICLE
            End If
            tgDBVcfSrchKey0.iSellCode = tgMVef(ilVef).iCode
            tgDBVcfSrchKey0.iSellDay = 0
            tgDBVcfSrchKey0.iEffDate(0) = 0
            tgDBVcfSrchKey0.iEffDate(1) = 0
            tgDBVcfSrchKey0.iSellTime(0) = 0
            tgDBVcfSrchKey0.iSellTime(1) = 0
            tgDBVcfSrchKey0.iSellPosNo = 0
            ilRet = btrGetGreaterOrEqual(hgDBVcf, tgDBVcf, igDBVcfRecLen, tgDBVcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
            If (ilRet = BTRV_ERR_NONE) And (tgDBVcf.iSellCode = tgMVef(ilVef).iCode) Then
                tmVcfVefCode(UBound(tmVcfVefCode)).sType = tgMVef(ilVef).sType
                tmVcfVefCode(UBound(tmVcfVefCode)).iVefCode = tgMVef(ilVef).iCode
                ReDim Preserve tmVcfVefCode(0 To UBound(tmVcfVefCode) + 1) As DBVEHICLE
            Else
                tgDBVcfSrchKey0.iSellCode = tgMVef(ilVef).iCode
                tgDBVcfSrchKey0.iSellDay = 6
                tgDBVcfSrchKey0.iEffDate(0) = 0
                tgDBVcfSrchKey0.iEffDate(1) = 0
                tgDBVcfSrchKey0.iSellTime(0) = 0
                tgDBVcfSrchKey0.iSellTime(1) = 0
                tgDBVcfSrchKey0.iSellPosNo = 0
                ilRet = btrGetGreaterOrEqual(hgDBVcf, tgDBVcf, igDBVcfRecLen, tgDBVcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                If (ilRet = BTRV_ERR_NONE) And (tgDBVcf.iSellCode = tgMVef(ilVef).iCode) Then
                    tmVcfVefCode(UBound(tmVcfVefCode)).sType = tgMVef(ilVef).sType
                    tmVcfVefCode(UBound(tmVcfVefCode)).iVefCode = tgMVef(ilVef).iCode
                    ReDim Preserve tmVcfVefCode(0 To UBound(tmVcfVefCode) + 1) As DBVEHICLE
                Else
                    tgDBVcfSrchKey0.iSellCode = tgMVef(ilVef).iCode
                    tgDBVcfSrchKey0.iSellDay = 7
                    tgDBVcfSrchKey0.iEffDate(0) = 0
                    tgDBVcfSrchKey0.iEffDate(1) = 0
                    tgDBVcfSrchKey0.iSellTime(0) = 0
                    tgDBVcfSrchKey0.iSellTime(1) = 0
                    tgDBVcfSrchKey0.iSellPosNo = 0
                    ilRet = btrGetGreaterOrEqual(hgDBVcf, tgDBVcf, igDBVcfRecLen, tgDBVcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                    If (ilRet = BTRV_ERR_NONE) And (tgDBVcf.iSellCode = tgMVef(ilVef).iCode) Then
                        tmVcfVefCode(UBound(tmVcfVefCode)).sType = tgMVef(ilVef).sType
                        tmVcfVefCode(UBound(tmVcfVefCode)).iVefCode = tgMVef(ilVef).iCode
                        ReDim Preserve tmVcfVefCode(0 To UBound(tmVcfVefCode) + 1) As DBVEHICLE
                    End If
                End If
            End If
            tgDBVlfSrchKey0.iSellCode = tgMVef(ilVef).iCode
            tgDBVlfSrchKey0.iSellDay = 0
            tgDBVlfSrchKey0.iEffDate(0) = 0
            tgDBVlfSrchKey0.iEffDate(1) = 0
            tgDBVlfSrchKey0.iSellTime(0) = 0
            tgDBVlfSrchKey0.iSellTime(1) = 0
            tgDBVlfSrchKey0.iSellPosNo = 0
            tgDBVlfSrchKey0.iSellSeq = 0
            ilRet = btrGetGreaterOrEqual(hgDBVlf, tgDBVlf, igDBVlfRecLen, tgDBVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
            If (ilRet = BTRV_ERR_NONE) And (tgDBVlf.iSellCode = tgMVef(ilVef).iCode) Then
                tmVlfVefCode(UBound(tmVlfVefCode)).sType = tgMVef(ilVef).sType
                tmVlfVefCode(UBound(tmVlfVefCode)).iVefCode = tgMVef(ilVef).iCode
                ReDim Preserve tmVlfVefCode(0 To UBound(tmVlfVefCode) + 1) As DBVEHICLE
            Else
                tgDBVlfSrchKey0.iSellCode = tgMVef(ilVef).iCode
                tgDBVlfSrchKey0.iSellDay = 6
                tgDBVlfSrchKey0.iEffDate(0) = 0
                tgDBVlfSrchKey0.iEffDate(1) = 0
                tgDBVlfSrchKey0.iSellTime(0) = 0
                tgDBVlfSrchKey0.iSellTime(1) = 0
                tgDBVlfSrchKey0.iSellPosNo = 0
                tgDBVlfSrchKey0.iSellSeq = 0
                ilRet = btrGetGreaterOrEqual(hgDBVlf, tgDBVlf, igDBVlfRecLen, tgDBVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                If (ilRet = BTRV_ERR_NONE) And (tgDBVlf.iSellCode = tgMVef(ilVef).iCode) Then
                    tmVlfVefCode(UBound(tmVlfVefCode)).sType = tgMVef(ilVef).sType
                    tmVlfVefCode(UBound(tmVlfVefCode)).iVefCode = tgMVef(ilVef).iCode
                    ReDim Preserve tmVlfVefCode(0 To UBound(tmVlfVefCode) + 1) As DBVEHICLE
                Else
                    tgDBVlfSrchKey0.iSellCode = tgMVef(ilVef).iCode
                    tgDBVlfSrchKey0.iSellDay = 7
                    tgDBVlfSrchKey0.iEffDate(0) = 0
                    tgDBVlfSrchKey0.iEffDate(1) = 0
                    tgDBVlfSrchKey0.iSellTime(0) = 0
                    tgDBVlfSrchKey0.iSellTime(1) = 0
                    tgDBVlfSrchKey0.iSellPosNo = 0
                    tgDBVlfSrchKey0.iSellSeq = 0
                    ilRet = btrGetGreaterOrEqual(hgDBVlf, tgDBVlf, igDBVlfRecLen, tgDBVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                    If (ilRet = BTRV_ERR_NONE) And (tgDBVlf.iSellCode = tgMVef(ilVef).iCode) Then
                        tmVlfVefCode(UBound(tmVlfVefCode)).sType = tgMVef(ilVef).sType
                        tmVlfVefCode(UBound(tmVlfVefCode)).iVefCode = tgMVef(ilVef).iCode
                        ReDim Preserve tmVlfVefCode(0 To UBound(tmVlfVefCode) + 1) As DBVEHICLE
                    End If
                End If
            End If
            tgDBDlfSrchKey0.iVefCode = tgMVef(ilVef).iCode
            tgDBDlfSrchKey0.sAirDay = "0"
            tgDBDlfSrchKey0.iStartDate(0) = 0
            tgDBDlfSrchKey0.iStartDate(1) = 0
            tgDBDlfSrchKey0.iAirTime(0) = 0
            tgDBDlfSrchKey0.iAirTime(1) = 0
            ilRet = btrGetGreaterOrEqual(hgDBDlf, tgDBDlf, igDBDlfRecLen, tgDBDlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
            If (ilRet = BTRV_ERR_NONE) And (tgDBDlf.iVefCode = tgMVef(ilVef).iCode) Then
                tmDlfVefCode(UBound(tmDlfVefCode)).sType = tgMVef(ilVef).sType
                tmDlfVefCode(UBound(tmDlfVefCode)).iVefCode = tgMVef(ilVef).iCode
                ReDim Preserve tmDlfVefCode(0 To UBound(tmDlfVefCode) + 1) As DBVEHICLE
            Else
                tgDBDlfSrchKey0.iVefCode = tgMVef(ilVef).iCode
                tgDBDlfSrchKey0.sAirDay = "6"
                tgDBDlfSrchKey0.iStartDate(0) = 0
                tgDBDlfSrchKey0.iStartDate(1) = 0
                tgDBDlfSrchKey0.iAirTime(0) = 0
                tgDBDlfSrchKey0.iAirTime(1) = 0
                ilRet = btrGetGreaterOrEqual(hgDBDlf, tgDBDlf, igDBDlfRecLen, tgDBDlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                If (ilRet = BTRV_ERR_NONE) And (tgDBDlf.iVefCode = tgMVef(ilVef).iCode) Then
                    tmDlfVefCode(UBound(tmDlfVefCode)).sType = tgMVef(ilVef).sType
                    tmDlfVefCode(UBound(tmDlfVefCode)).iVefCode = tgMVef(ilVef).iCode
                    ReDim Preserve tmDlfVefCode(0 To UBound(tmDlfVefCode) + 1) As DBVEHICLE
                Else
                    tgDBDlfSrchKey0.iVefCode = tgMVef(ilVef).iCode
                    tgDBDlfSrchKey0.sAirDay = "7"
                    tgDBDlfSrchKey0.iStartDate(0) = 0
                    tgDBDlfSrchKey0.iStartDate(1) = 0
                    tgDBDlfSrchKey0.iAirTime(0) = 0
                    tgDBDlfSrchKey0.iAirTime(1) = 0
                    ilRet = btrGetGreaterOrEqual(hgDBDlf, tgDBDlf, igDBDlfRecLen, tgDBDlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                    If (ilRet = BTRV_ERR_NONE) And (tgDBDlf.iVefCode = tgMVef(ilVef).iCode) Then
                        tmDlfVefCode(UBound(tmDlfVefCode)).sType = tgMVef(ilVef).sType
                        tmDlfVefCode(UBound(tmDlfVefCode)).iVefCode = tgMVef(ilVef).iCode
                        ReDim Preserve tmDlfVefCode(0 To UBound(tmDlfVefCode) + 1) As DBVEHICLE
                    End If
                End If
            End If
            tgDBDlfSrchKey0.iVefCode = tgMVef(ilVef).iCode
            tgDBDlfSrchKey0.sAirDay = "0"
            tgDBDlfSrchKey0.iStartDate(0) = 0
            tgDBDlfSrchKey0.iStartDate(1) = 0
            tgDBDlfSrchKey0.iAirTime(0) = 0
            tgDBDlfSrchKey0.iAirTime(1) = 0
            ilRet = btrGetGreaterOrEqual(hgDBEgf, tgDBDlf, igDBDlfRecLen, tgDBDlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
            If (ilRet = BTRV_ERR_NONE) And (tgDBDlf.iVefCode = tgMVef(ilVef).iCode) Then
                tmEgfVefCode(UBound(tmEgfVefCode)).sType = tgMVef(ilVef).sType
                tmEgfVefCode(UBound(tmEgfVefCode)).iVefCode = tgMVef(ilVef).iCode
                ReDim Preserve tmEgfVefCode(0 To UBound(tmEgfVefCode) + 1) As DBVEHICLE
            Else
                tgDBDlfSrchKey0.iVefCode = tgMVef(ilVef).iCode
                tgDBDlfSrchKey0.sAirDay = "6"
                tgDBDlfSrchKey0.iStartDate(0) = 0
                tgDBDlfSrchKey0.iStartDate(1) = 0
                tgDBDlfSrchKey0.iAirTime(0) = 0
                tgDBDlfSrchKey0.iAirTime(1) = 0
                ilRet = btrGetGreaterOrEqual(hgDBEgf, tgDBDlf, igDBDlfRecLen, tgDBDlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                If (ilRet = BTRV_ERR_NONE) And (tgDBDlf.iVefCode = tgMVef(ilVef).iCode) Then
                    tmEgfVefCode(UBound(tmEgfVefCode)).sType = tgMVef(ilVef).sType
                    tmEgfVefCode(UBound(tmEgfVefCode)).iVefCode = tgMVef(ilVef).iCode
                    ReDim Preserve tmEgfVefCode(0 To UBound(tmEgfVefCode) + 1) As DBVEHICLE
                Else
                    tgDBDlfSrchKey0.iVefCode = tgMVef(ilVef).iCode
                    tgDBDlfSrchKey0.sAirDay = "7"
                    tgDBDlfSrchKey0.iStartDate(0) = 0
                    tgDBDlfSrchKey0.iStartDate(1) = 0
                    tgDBDlfSrchKey0.iAirTime(0) = 0
                    tgDBDlfSrchKey0.iAirTime(1) = 0
                    ilRet = btrGetGreaterOrEqual(hgDBEgf, tgDBDlf, igDBDlfRecLen, tgDBDlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                    If (ilRet = BTRV_ERR_NONE) And (tgDBDlf.iVefCode = tgMVef(ilVef).iCode) Then
                        tmEgfVefCode(UBound(tmEgfVefCode)).sType = tgMVef(ilVef).sType
                        tmEgfVefCode(UBound(tmEgfVefCode)).iVefCode = tgMVef(ilVef).iCode
                        ReDim Preserve tmEgfVefCode(0 To UBound(tmEgfVefCode) + 1) As DBVEHICLE
                    End If
                End If
            End If
            tgDBStfSrchKey0.iVefCode = tgMVef(ilVef).iCode
            tgDBStfSrchKey0.iLogDate(0) = 0
            tgDBStfSrchKey0.iLogDate(1) = 0
            gPackTime "", tgDBStfSrchKey0.iLogTime(0), tgDBStfSrchKey0.iLogTime(1)
            ilRet = btrGetGreaterOrEqual(hgDBStf, tgDBStf, igDBStfRecLen, tgDBStfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) And (tgDBStf.iVefCode = tgMVef(ilVef).iCode) Then
                tmStfVefCode(UBound(tmStfVefCode)).sType = tgMVef(ilVef).sType
                tmStfVefCode(UBound(tmStfVefCode)).iVefCode = tgMVef(ilVef).iCode
                ReDim Preserve tmStfVefCode(0 To UBound(tmStfVefCode) + 1) As DBVEHICLE
            End If
       End If
    Next ilVef
End Sub

Private Sub mBtrErrorMsg(slMsg As String, ilInRet As Integer)
    Dim ilRet As Integer
    
    ilRet = ilInRet
    If ilRet >= 30000 Then
        ilRet = csiHandleValue(0, 7)
    End If
    gLogMsg slMsg & Trim$(Str$(ilRet)), "csiArchive.txt", False
End Sub

Private Function mRemoveDat(llAttCode As Long) As Integer
    Dim ilRet As Integer
    
    Do
        tgDBDatSrchKey1.lCode = llAttCode
        ilRet = btrGetEqual(hgDBDat, tgDBDat, igDBDatRecLen, tgDBDatSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBDat)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveDat: Delete error ", ilRet
                mRemoveDat = False
                Exit Function
            End If
        End If
    Loop
    mRemoveDat = True
End Function
Private Function mRemoveEpt(llAttCode As Long) As Integer
    Dim ilRet As Integer
    
    Do
        tgDBEptSrchKey2.lAttCode = llAttCode
        ilRet = btrGetEqual(hgDBEpt, tgDBEpt, igDBEptRecLen, tgDBEptSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBEpt)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveEpt: Delete error ", ilRet
                mRemoveEpt = False
                Exit Function
            End If
        End If
    Loop
    mRemoveEpt = True
End Function
Private Function mRemovePet(llAttCode As Long) As Integer
    Dim ilRet As Integer
    
    Do
        tgDBPetSrchKey1.lAttCode = llAttCode
        ilRet = btrGetEqual(hgDBPet, tgDBPet, igDBPetRecLen, tgDBPetSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBPet)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemovePet: Delete error ", ilRet
                mRemovePet = False
                Exit Function
            End If
        End If
    Loop
    mRemovePet = True
End Function
Private Function mRemoveVat(llAttCode As Long) As Integer
    Dim ilRet As Integer
    
    Do
        tgDBVatSrchKey0.lAttCode = llAttCode
        ilRet = btrGetEqual(hgDBVat, tgDBVat, igDBVatRecLen, tgDBVatSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBVat)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveVat: Delete error ", ilRet
                mRemoveVat = False
                Exit Function
            End If
        End If
    Loop
    mRemoveVat = True
End Function
Private Function mArchiveAndRemoveCrf(llStartDate As Long, llEndDate As Long, llCrfCode As Long, ilCount As Long) As Integer
    Dim ilRet As Integer
    Dim llTestDate As Long
    Dim llCif As Long
    Dim ilFound As Integer
    Dim slScript As String
    
    smTeamTag = ""
    mTeamPop
    
    Do
        tgDBCrfSrchKey.lCode = llCrfCode
        ilRet = btrGetEqual(hgDBCrf, tgDBCrf, igDBCrfRecLen, tgDBCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        gUnpackDateLong tgDBCrf.iEndDate(0), tgDBCrf.iEndDate(1), llTestDate
        If llTestDate < llStartDate Then
            Exit Do
        End If
        If llTestDate > llEndDate Then
            Exit Do
        End If
        
        ilRet = mArchiveCopy()
        If Not ilRet Then
            mArchiveAndRemoveCrf = False
            Exit Function
        End If
        
        If tgDBCrf.lCsfCode > 0 Then
            igDBCsfRecLen = Len(tgDBCsf) '5011
            tgDBCsfSrchKey.lCode = tgDBCrf.lCsfCode
            ilRet = btrGetEqual(hgDBCsf, tgDBCsf, igDBCsfRecLen, tgDBCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hgDBCsf)
            End If
        End If
        ilRet = mRemoveCnf(tgDBCrf.lCode)
        ilRet = btrDelete(hgDBCrf)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mArchiveAndRemoveCRF: Delete CRF error ", ilRet
                mArchiveAndRemoveCrf = False
                Exit Function
            End If
        Else
            ilCount = ilCount + 1
        End If
    Loop
    mArchiveAndRemoveCrf = True
End Function

Private Function mArchiveCopy() As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilAdf As Integer
    Dim ilAgf As Integer
    Dim slStr As String
    Dim slRecord As String
    Dim slCart As String
    Dim slISCI As String
    Dim slCopyProduct As String
    Dim slCreativeTitle As String
    Dim ilCif As Integer
    Dim ilUrf As Integer
    Dim ilMnf As Integer
    Dim ilAnf As Integer
    Dim slCopyRegionName As String
    Dim ilDay As Integer
    ReDim ilDays(0 To 6) As Integer
    ReDim slDays(0 To 6) As String * 1
    Dim slStrDays As String
    Dim slShortTitle As String
    Dim ilUrfFound As Integer
    Dim slLanguage As String
    Dim slScript As String
    Dim slCopy As String
    
    If imNoOutput Then
        mArchiveCopy = True
        Exit Function
    End If
        
    'Advertiser,Addr ID,Contract #,Product,Assign Vehicle,Rot Type,Rot #,Start Date,End Date,Start Time,End Time,Days,
    'Len,In/Out,Avail,Comment,Entry Date,Zone, User,Earliest Spot Date,Latest Spot Date,Short Title,Copy Region,
    'Live Copy,MG Assign,Active,Language,Games or Teams,INSTRUCTIONS,Cart#,Product,ISCI,Creative title

    tgDBChfSrchKey0.lCode = tgDBCrf.lChfCode
    ilRet = btrGetEqual(hgDBChf, tgDBChf, igDBChfRecLen, tgDBChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        'mBtrErrorMsg "mArchiveCopy: Get Equal Chf error ", ilRet
        'mArchiveCopy = True
        'Exit Function
        tgDBChf.lCntrNo = 0
        tgDBChf.sProduct = ""
    End If
    slScript = ""
    If tgDBCrf.lCsfCode > 0 Then
        igDBCsfRecLen = Len(tgDBCsf) '5011
        tgDBCsfSrchKey.lCode = tgDBCrf.lCsfCode
        ilRet = btrGetEqual(hgDBCsf, tgDBCsf, igDBCsfRecLen, tgDBCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            'slScript = Trim$(Left$(tgDBCsf.sComment, tgDBCsf.iStrLen))
            slScript = gStripChr0(tgDBCsf.sComment)
        End If
    End If
    slCopyRegionName = ""
    If tgDBCrf.lRafCode > 0 Then
        tgDBRafSrchKey.lCode = tgDBCrf.lRafCode
        ilRet = btrGetEqual(hgDBRaf, tgDBRaf, igDBRafRecLen, tgDBRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If tgDBRaf.sType <> "N" Then
                slCopyRegionName = Trim$(tgDBRaf.sName)
            End If
        End If
    End If
    slShortTitle = ""
    If (tgSpf.sUseProdSptScr = "P") Then
        If tgDBCrf.lSifCode > 0 Then
            tgDBSifSrchKey.lCode = tgDBCrf.lSifCode
            ilRet = btrGetEqual(hgDBSif, tgDBSif, igDBSifRecLen, tgDBSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                slShortTitle = Trim$(tgDBSif.sName)
            End If
        End If
    End If
    slLanguage = ""
    If tgDBCrf.iLangMnfCode > 0 Then
        ilMnf = mBinarySearchLangMnf(tgDBCrf.iLangMnfCode)
        If ilMnf <> -1 Then
            slLanguage = Trim$(tmLangMnf(ilMnf).sName)
        End If
    End If
    ilVef = gBinarySearchVef(tgDBCrf.iVefCode)
    ilAdf = gBinarySearchAdf(tgDBCrf.iAdfCode)
    ilAnf = mBinarySearchAvail(tgDBCrf.ianfCode)
    'Advertiser and ID
    If ilAdf >= LBound(tgCommAdf) Then
        slRecord = gAddQuotesRemoveChars(Trim$(tgCommAdf(ilAdf).sName)) & "," & gAddQuotesRemoveChars(Trim$(tgCommAdf(ilAdf).sAddrID)) & ","
    Else
        slRecord = gAddQuotesRemoveChars("") & "," & gAddQuotesRemoveChars("") & ","
    End If
    'Contract #
    slRecord = slRecord & Trim$(Str$(tgDBChf.lCntrNo)) & ","
    'Contract Product
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgDBChf.sProduct)) & ","
    'Assigned vehicle
    If ilVef >= LBound(tgMVef) Then
        slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgMVef(ilVef).sName)) & ","
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("") & ","
    End If
    'Rotation Type
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgDBCrf.sRotType)) & ","
    'Rot #
    slRecord = slRecord & Trim$(Str$(tgDBCrf.iRotNo)) & ","
    'Start Date
    gUnpackDate tgDBCrf.iStartDate(0), tgDBCrf.iStartDate(1), slStr
    slRecord = slRecord & Format(slStr, "mm/dd/yyyy") & ","
    'End Date
    gUnpackDate tgDBCrf.iEndDate(0), tgDBCrf.iEndDate(1), slStr
    slRecord = slRecord & Format(slStr, "mm/dd/yyyy") & ","
    'Start Time
    gUnpackTime tgDBCrf.iStartTime(0), tgDBCrf.iStartTime(1), "A", "1", slStr
    slRecord = slRecord & Format(slStr, "hh:mm:ssam/pm") & ","
    'End Time
    gUnpackTime tgDBCrf.iEndTime(0), tgDBCrf.iEndTime(1), "A", "1", slStr
    slRecord = slRecord & Format(slStr, "hh:mm:ssam/pm") & ","
    'Days
    slStr = ""
    For ilDay = 0 To 6 Step 1
        If tgDBCrf.sDay(ilDay) = "Y" Then
            ilDays(ilDay) = 1
        Else
            ilDays(ilDay) = 0
        End If
        slDays(ilDay) = ""
    Next ilDay
    slRecord = slRecord & gAddQuotesRemoveChars(gDayNames(ilDays(), slDays(), 2, slStrDays)) & ","
    'Len
    'Spot Length
    slRecord = slRecord & Trim$(Str$(tgDBCrf.iLen)) & ","
    'In/Out
    slRecord = slRecord & gAddQuotesRemoveChars(tgDBCrf.sInOut) & ","
    'Avail
    If ilAnf >= LBound(tgAvailAnf) Then
        slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgAvailAnf(ilAnf).sName)) & ","
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("") & ","
    End If
    'Comment
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(slScript)) & ","
    'Entry Date
    gUnpackDate tgDBCrf.iEntryDate(0), tgDBCrf.iEntryDate(1), slStr
    slRecord = slRecord & Format(slStr, "mm/dd/yyyy") & ","
    'Zone
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgDBCrf.sZone)) & ","
    'User
    ilUrfFound = False
    If tgDBSdf.iUrfCode > 2 Then
        For ilUrf = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
            DoEvents
            If tgPopUrf(ilUrf).iCode = tgDBSdf.iUrfCode Then
                slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgPopUrf(ilUrf).sRept)) & ","
                ilUrfFound = True
                Exit For
            End If
        Next ilUrf
    Else
        slRecord = slRecord & gAddQuotesRemoveChars("Guide") & ","
        ilUrfFound = True
    End If
    If Not ilUrfFound Then
        slRecord = slRecord & gAddQuotesRemoveChars("") & ","
    End If
    'Earliest Spot Date
    gUnpackDate tgDBCrf.iEarliestDateAssg(0), tgDBCrf.iEarliestDateAssg(1), slStr
    slRecord = slRecord & Format(slStr, "mm/dd/yyyy") & ","
    'Latest Spot Date
    gUnpackDate tgDBCrf.iLatestDateAssg(0), tgDBCrf.iLatestDateAssg(1), slStr
    slRecord = slRecord & Format(slStr, "mm/dd/yyyy") & ","
    'Short Title
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(slShortTitle)) & ","
    'Copy Region
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(slCopyRegionName)) & ","
    'Live Copy
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgDBCrf.sLiveCopy)) & ","
    'MG Assign
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgDBCrf.sMGCopyAssign)) & ","
    'Active
    slRecord = slRecord & gAddQuotesRemoveChars(Trim$(tgDBCrf.sState)) & ","
    'Language
    'If ilMnf <> -1 Then
        slRecord = slRecord & gAddQuotesRemoveChars(slLanguage) & ","
    'Else
    '    slRecord = slRecord & ","
    'End If
    slStr = ""
    If (tgDBCrf.sAirGameType = "G") Or (tgDBCrf.sAirGameType = "T") Then
        tgDBCafSrchKey1.lCrfCode = tgDBCrf.lCode
        ilRet = btrGetGreaterOrEqual(hgDBCaf, tgDBCaf, igDBCafRecLen, tgDBCafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            If tgDBCaf.lCrfCode <> tgDBCrf.lCode Then
                Exit Do
            End If
            If (tgDBCrf.sAirGameType = "G") And (tgDBCaf.sType = "G") Then
                If slStr = "" Then
                    slStr = "Games #:"
                End If
                slStr = slStr & " " & Trim$(Str$(tgDBCaf.iGameNo))
            ElseIf (tgDBCrf.sAirGameType = "T") And (tgDBCaf.sType = "T") Then
                If slStr = "" Then
                    slStr = "Teams:"
                End If
                For ilMnf = LBound(tmTeam) To UBound(tmTeam) - 1 Step 1
                    If tgDBCaf.iTeamMnfCode = tmTeam(ilMnf).iCode Then
                        slStr = slStr & " " & Trim$(tmTeam(ilMnf).sName)
                        Exit For
                    End If
                Next ilMnf
            End If
            ilRet = btrGetNext(hgDBCaf, tgDBCaf, igDBCafRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    slRecord = slRecord & gAddQuotesRemoveChars(slStr) & ","
    'INSTRUCTIONS
    slRecord = slRecord & "INSTRUCTIONS" & ","
    'Cart#
    'Product
    'ISCI
    'Creative title
    slCopy = ""
    tgDBCnfSrchKey.lCrfCode = tgDBCrf.lCode
    tgDBCnfSrchKey.iInstrNo = 0
    ilRet = btrGetGreaterOrEqual(hgDBCnf, tgDBCnf, igDBCnfRecLen, tgDBCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        If tgDBCnf.lCrfCode <> tgDBCrf.lCode Then
            Exit Do
        End If
        tgDBCifSrchKey.lCode = tgDBCnf.lCifCode
        ilRet = btrGetEqual(hgDBCif, tgDBCif, igDBCifRecLen, tgDBCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            slCart = ""
            slCopyProduct = ""
            slISCI = ""
            slCreativeTitle = ""
            slScript = ""
            mGetCopy "1", tgDBCif.lCode, slCart, slCopyProduct, slISCI, slCreativeTitle, slScript
            If StrComp(Trim$(tgDBChf.sProduct), Trim$(slCopyProduct), vbTextCompare) = 0 Then
                slCopy = slCopy & Trim$(slCart) & ";" & "" & ";" & Trim$(slISCI) & ";" & Trim$(slCreativeTitle) & ";"   ' & gAddQuotesRemoveChars(Trim$(slScript)) & ","
            Else
                slCopy = slCopy & Trim$(slCart) & ";" & Trim$(slCopyProduct) & ";" & Trim$(slISCI) & ";" & Trim$(slCreativeTitle) & ";" ' & gAddQuotesRemoveChars(Trim$(slScript)) & ";"
            End If
        Else
            slCopy = slCopy & ""
        End If
        ilRet = btrGetNext(hgDBCnf, tgDBCnf, igDBCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If Len(slCopy) > 0 Then
        slCopy = Left$(slCopy, Len(slCopy) - 1)
    End If
    slRecord = slRecord & gAddQuotesRemoveChars(slCopy)
    Print #hmArchive, slRecord
    mArchiveCopy = True
End Function

Private Function mRemoveCif(llRetainDate As Long, llCifCount As Long) As Integer
    Dim ilRet As Integer
    Dim ilDelete As Integer
    Dim ilMcf As Integer
    Dim llDate As Long
    
    ilRet = btrGetFirst(hgDBCif, tgDBCif, igDBCifRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        ilDelete = False
        If tgDBCif.lCode <= UBound(lmRetainCifCode) Then
            If lmRetainCifCode(tgDBCif.lCode) <= 0 Then
                gUnpackDateLong tgDBCif.iUsedDate(0), tgDBCif.iUsedDate(1), llDate
                If llDate = 0 Then
                    gUnpackDateLong tgDBCif.iDateEntrd(0), tgDBCif.iDateEntrd(1), llDate
                End If
                If (llDate > 0) And (llDate < llRetainDate) Then
                    If tgDBCif.iMcfCode > 0 Then
                        ilMcf = mBinarySearchMcf(tgDBCif.iMcfCode)
                        If ilMcf <> -1 Then
                            If tgMCF(ilMcf).sReuse = "Y" Then
                                If tgDBCif.sPurged = "H" Then
                                    ilDelete = True
                                End If
                            Else
                                ilDelete = True
                            End If
                        End If
                    Else
                        ilDelete = True
                    End If
                    If ilDelete Then
                        '10/17/08:  CPF must be retained as many cif records could be referencing the same CPF record.
                        '           To remove CPF we would need to create a map of used cifcpfcodes, then remove any not referenced
                        'Remove CPF
                        'tgDBCpfSrchKey.lCode = tgDBCif.lCpfCode
                        'ilRet = btrGetEqual(hgDBCpf, tgDBCpf, igDBCpfRecLen, tgDBCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        'If ilRet = BTRV_ERR_NONE Then
                        '    ilRet = btrDelete(hgDBCpf)
                        'End If
                        'Remove CYF
                        Do
                            tgDBCyfSrchKey.lCifCode = tgDBCif.lCode
                            tgDBCyfSrchKey.iVefCode = 0
                            tgDBCyfSrchKey.sSource = ""
                            tgDBCyfSrchKey.sTimeZone = ""
                            tgDBCyfSrchKey.lRafCode = 0
                            ilRet = btrGetGreaterOrEqual(hgDBCyf, tgDBCyf, igDBCyfRecLen, tgDBCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                            If ilRet <> BTRV_ERR_NONE Then
                                Exit Do
                            End If
                            If tgDBCyf.lCifCode <> tgDBCif.lCode Then
                                Exit Do
                            End If
                            ilRet = btrDelete(hgDBCyf)
                        Loop
                        'Remove CSF
                        If tgDBCif.lCsfCode > 0 Then
                            igDBCsfRecLen = Len(tgDBCsf) '5011
                            tgDBCsfSrchKey.lCode = tgDBCif.lCsfCode
                            ilRet = btrGetEqual(hgDBCsf, tgDBCsf, igDBCsfRecLen, tgDBCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                ilRet = btrDelete(hgDBCsf)
                            End If
                        End If
                        ilRet = btrDelete(hgDBCif)
                        If ilRet <> BTRV_ERR_NONE Then
                            If ilRet <> BTRV_ERR_CONFLICT Then
                                mBtrErrorMsg "mRemoveCIF: Delete error ", ilRet
                                mRemoveCif = False
                                Exit Function
                            End If
                        Else
                            llCifCount = llCifCount + 1
                        End If
                    End If
                End If
            End If
        End If
        If ilDelete Then
            tgDBCifSrchKey.lCode = tgDBCif.lCode + 1
            ilRet = btrGetGreaterOrEqual(hgDBCif, tgDBCif, igDBCifRecLen, tgDBCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        Else
            ilRet = btrGetNext(hgDBCif, tgDBCif, igDBCifRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        End If
    Loop
    mRemoveCif = True
End Function

Private Function mRemoveBof(llRetainDate As Long, llBofCount As Long) As Integer
    Dim llDate As Long
    Dim llBof As Long
    Dim ilRet As Integer
    
    If llRetainDate <= 0 Then
        mRemoveBof = True
        Exit Function
    End If

    ReDim llBofCode(0 To 0) As Long
    ilRet = btrGetFirst(hgDBBof, tgDBBof, igDBBofRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        gUnpackDateLong tgDBBof.iEndDate(0), tgDBBof.iEndDate(1), llDate
        If (llDate < llRetainDate) And (llDate > 0) Then
            llBofCode(UBound(llBofCode)) = tgDBBof.lCode
            ReDim Preserve llBofCode(0 To UBound(llBofCode) + 1) As Long
        End If
        ilRet = btrGetNext(hgDBBof, tgDBBof, igDBBofRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    llBof = 0
    Do While llBof < UBound(llBofCode)
        tgDBBofSrchKey0.lCode = llBofCode(llBof)
        ilRet = btrGetEqual(hgDBBof, tgDBBof, igDBBofRecLen, tgDBBofSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hgDBBof)
            If ilRet = BTRV_ERR_NONE Then
                llBofCount = llBofCount + 1
                llBof = llBof + 1
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    llBof = llBof + 1
                End If
            End If
        Else
            llBof = llBof + 1
        End If
    Loop
    Erase llBofCode
    mRemoveBof = True
End Function

Private Function mRemovePjf(llRetainDate As Long, llPjfCount As Long) As Integer
    Dim llDate As Long
    Dim llPjf As Long
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim llLastCode As Long
    
    If llRetainDate <= 0 Then
        mRemovePjf = True
        Exit Function
    End If

    ReDim llPjfCode(0 To 1000) As Long
    ilUpper = 0
    llLastCode = 0
    Do
        DoEvents
        'ilRet = btrGetFirst(hgDBPjf, tgDBPjf, igDBPjfRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        tgDBPjfSrchKey1.lCode = llLastCode
        ilRet = btrGetGreaterOrEqual(hgDBPjf, tgDBPjf, igDBPjfRecLen, tgDBPjfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get lSdf current record to obtain date
        Do While ilRet = BTRV_ERR_NONE
            llLastCode = tgDBPjf.lCode
            gUnpackDateLong tgDBPjf.iRolloverDate(0), tgDBPjf.iRolloverDate(1), llDate
            If (llDate < llRetainDate) And (llDate > 0) Then
                'llPjfCode(UBound(llPjfCode)) = tgDBPjf.lCode
                'ReDim Preserve llPjfCode(0 To UBound(llPjfCode) + 1) As Long
                llPjfCode(ilUpper) = tgDBPjf.lCode
                ilUpper = ilUpper + 1
                If ilUpper >= UBound(llPjfCode) Then
                    Exit Do
                End If
            End If
            ilRet = btrGetNext(hgDBPjf, tgDBPjf, igDBPjfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
        DoEvents
        If ilUpper = 0 Then
            Exit Do
        End If
        llPjf = 0
        Do While llPjf < ilUpper    'UBound(llPjfCode)
            tgDBPjfSrchKey1.lCode = llPjfCode(llPjf)
            ilRet = btrGetEqual(hgDBPjf, tgDBPjf, igDBPjfRecLen, tgDBPjfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hgDBPjf)
                If ilRet = BTRV_ERR_NONE Then
                    llPjfCount = llPjfCount + 1
                    llPjf = llPjf + 1
                Else
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        llPjf = llPjf + 1
                    End If
                End If
            Else
                llPjf = llPjf + 1
            End If
        Loop
        ilUpper = 0
    Loop
    Erase llPjfCode
    mRemovePjf = True
End Function

Private Function mRemoveProposals(llRetainDate As Long, llPropCount As Long) As Integer
    Dim ilPass As Integer
    Dim llDate As Long
    Dim llChf As Long
    Dim slStatus As String
    Dim ilRet As Integer
    Dim llLkVsfCode As Long
    
    If llRetainDate <= 0 Then
        mRemoveProposals = True
        Exit Function
    End If
    For ilPass = 0 To 3 Step 1
        DoEvents
        ReDim llPropChfCode(0 To 0) As Long
        If ilPass = 0 Then
            slStatus = "W"
        ElseIf ilPass = 1 Then
            slStatus = "D"
        ElseIf ilPass = 2 Then
            slStatus = "C"
        ElseIf ilPass = 3 Then
            slStatus = "I"
        End If
        tgDBChfSrchKey3.sStatus = slStatus
        tgDBChfSrchKey3.sNRProcessed = ""
        ilRet = btrGetGreaterOrEqual(hgDBChf, tgDBChf, igDBChfRecLen, tgDBChfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
        Do While ilRet = BTRV_ERR_NONE
            DoEvents
            If tgDBChf.sStatus <> slStatus Then
                Exit Do
            End If
            If tgDBChf.iCntRevNo <= 0 Then
                gUnpackDateLong tgDBChf.iPropDate(0), tgDBChf.iPropDate(1), llDate
                If llDate < llRetainDate Then
                    llPropChfCode(UBound(llPropChfCode)) = tgDBChf.lCode
                    ReDim Preserve llPropChfCode(0 To UBound(llPropChfCode) + 1) As Long
                End If
            End If
            ilRet = btrGetNext(hgDBChf, tgDBChf, igDBChfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
        llChf = 0
        Do While llChf < UBound(llPropChfCode)
            tgDBChfSrchKey0.lCode = llPropChfCode(llChf)
            ilRet = btrGetEqual(hgDBChf, tgDBChf, igDBChfRecLen, tgDBChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
'                'Remove SBF
'                Do
'                    tgDBSbfSrchKey0.lChfCode = tgDBChf.lCode
'                    tgDBSbfSrchKey0.iDate(0) = 0
'                    tgDBSbfSrchKey0.iDate(1) = 0
'                    tgDBSbfSrchKey0.sTranType = ""
'                    ilRet = btrGetGreaterOrEqual(hgDBSbf, tgDBSbf, igDBSbfRecLen, tgDBSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        Exit Do
'                    End If
'                    If tgDBSbf.lChfCode <> tgDBChf.lCode Then
'                        Exit Do
'                    End If
'                    ilRet = btrDelete(hgDBSbf)
'                Loop
'                'Remove clf, cff, cxf and cgf
'                Do
'                    tgDBClfSrchKey1.lChfCode = tgDBChf.lCode
'                    tgDBClfSrchKey1.iVefCode = 0
'                    ilRet = btrGetGreaterOrEqual(hgDBClf, tgDBClf, igDBClfRecLen, tgDBClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        Exit Do
'                    End If
'                    If tgDBClf.lChfCode <> tgDBChf.lCode Then
'                        Exit Do
'                    End If
'                    'Remove Cff
'                    Do
'                        tgDBCffSrchKey.lChfCode = tgDBChf.lCode
'                        tgDBCffSrchKey.iClfLine = tgDBClf.iLine
'                        tgDBCffSrchKey.iCntRevNo = 0
'                        tgDBCffSrchKey.iPropVer = 0
'                        tgDBCffSrchKey.iStartDate(0) = 0
'                        tgDBCffSrchKey.iStartDate(1) = 0
'                        ilRet = btrGetGreaterOrEqual(hgDBCff, tgDBCff, igDBCffRecLen, tgDBCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
'                        If ilRet <> BTRV_ERR_NONE Then
'                            Exit Do
'                        End If
'                        If tgDBCff.lChfCode <> tgDBChf.lCode Then
'                            Exit Do
'                        End If
'                        If tgDBCff.iClfLine <> tgDBClf.iLine Then
'                            Exit Do
'                        End If
'                        ilRet = btrDelete(hgDBCff)
'                    Loop
'                    'Remove cgf
'                    Do
'                        tgDBCgfSrchKey1.lClfCode = tgDBClf.lCode
'                        ilRet = btrGetGreaterOrEqual(hgDBCgf, tgDBCgf, igDBCgfRecLen, tgDBCgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
'                        If ilRet <> BTRV_ERR_NONE Then
'                            Exit Do
'                        End If
'                        If tgDBCgf.lClfCode <> tgDBClf.lCode Then
'                            Exit Do
'                        End If
'                        ilRet = btrDelete(hgDBCgf)
'                    Loop
'                    'Remove cxf
'                    If tgDBClf.lCxfCode > 0 Then
'                        igDBCxfRecLen = Len(tgDBCxf) '5011
'                        tgDBCxfSrchKey.lCode = tgDBClf.lCxfCode
'                        ilRet = btrGetEqual(hgDBCxf, tgDBCxf, igDBCxfRecLen, tgDBCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                        If ilRet = BTRV_ERR_NONE Then
'                            ilRet = btrDelete(hgDBCxf)
'                        End If
'                    End If
'                    ilRet = btrDelete(hgDBClf)
'                Loop
'                'Remove vsf
'                If tgDBChf.lVefCode < 0 Then
'                    tgDBVsfSrchKey.lCode = -tgDBChf.lVefCode
'                    ilRet = btrGetEqual(hgDBVsf, tgDBVsf, igDBVsfRecLen, tgDBVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                    If ilRet = BTRV_ERR_NONE Then
'                        llLkVsfCode = tgDBVsf.lLkVsfCode
'                        ilRet = btrDelete(hgDBVsf)
'                        Do While llLkVsfCode > 0
'                            tgDBVsfSrchKey.lCode = llLkVsfCode
'                            ilRet = btrGetEqual(hgDBVsf, tgDBVsf, igDBVsfRecLen, tgDBVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                            If ilRet <> BTRV_ERR_NONE Then
'                                Exit Do
'                            End If
'                            llLkVsfCode = tgDBVsf.lLkVsfCode
'                            ilRet = btrDelete(hgDBVsf)
'                        Loop
'                    End If
'                End If
'                'Remove msf and mgf
'                Do
'                    tgDBMsfSrchKey2.lChfCode = tgDBChf.lCode
'                    ilRet = btrGetEqual(hgDBMsf, tgDBMsf, igDBMsfRecLen, tgDBMsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        Exit Do
'                    End If
'                    'Remove Mgf
'                    Do
'                        tgDBMgfSrchKey1.lMsfCode = tgDBMsf.lCode
'                        tgDBMgfSrchKey1.iGameNo = 0
'                        ilRet = btrGetGreaterOrEqual(hgDBMgf, tgDBMgf, igDBMgfRecLen, tgDBMgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
'                        If ilRet <> BTRV_ERR_NONE Then
'                            Exit Do
'                        End If
'                        If tgDBMgf.lMsfCode <> tgDBMsf.lCode Then
'                            Exit Do
'                        End If
'                        ilRet = btrDelete(hgDBMgf)
'                    Loop
'                    ilRet = btrDelete(hgDBMsf)
'                Loop
'                'Remove cxf
'                'If tgDBChf.lCxfCanc > 0 Then
'                '    igDBCxfRecLen = Len(tgDBCxf) '5011
'                '    tgDBCxfSrchKey.lCode = tgDBChf.lCxfCanc
'                '    ilRet = btrGetEqual(hgDBCxf, tgDBCxf, igDBCxfRecLen, tgDBCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                '    If ilRet = BTRV_ERR_NONE Then
'                '        ilRet = btrDelete(hgDBCxf)
'                '    End If
'                'End If
'                ilRet = gDeleteCxf(tgDBChf.lCxfCanc)
'                'If tgDBChf.lCxfChgR > 0 Then
'                '    igDBCxfRecLen = Len(tgDBCxf) '5011
'                '    tgDBCxfSrchKey.lCode = tgDBChf.lCxfChgR
'                '    ilRet = btrGetEqual(hgDBCxf, tgDBCxf, igDBCxfRecLen, tgDBCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                '    If ilRet = BTRV_ERR_NONE Then
'                '        ilRet = btrDelete(hgDBCxf)
'                '    End If
'                'End If
'                ilRet = gDeleteCxf(tgDBChf.lCxfChgR)
'                'If tgDBChf.lCxfCode > 0 Then
'                '    igDBCxfRecLen = Len(tgDBCxf) '5011
'                '    tgDBCxfSrchKey.lCode = tgDBChf.lCxfCode
'                '    ilRet = btrGetEqual(hgDBCxf, tgDBCxf, igDBCxfRecLen, tgDBCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                '    If ilRet = BTRV_ERR_NONE Then
'                '        ilRet = btrDelete(hgDBCxf)
'                '    End If
'                'End If
'                ilRet = gDeleteCxf(tgDBChf.lCxfCode)
'                'If tgDBChf.lCxfInt > 0 Then
'                '    igDBCxfRecLen = Len(tgDBCxf) '5011
'                '    tgDBCxfSrchKey.lCode = tgDBChf.lCxfInt
'                '    ilRet = btrGetEqual(hgDBCxf, tgDBCxf, igDBCxfRecLen, tgDBCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                '    If ilRet = BTRV_ERR_NONE Then
'                '        ilRet = btrDelete(hgDBCxf)
'                '    End If
'                'End If
'                ilRet = gDeleteCxf(tgDBChf.lCxfInt)
'                'If tgDBChf.lCxfMerch > 0 Then
'                '    igDBCxfRecLen = Len(tgDBCxf) '5011
'                '    tgDBCxfSrchKey.lCode = tgDBChf.lCxfMerch
'                '    ilRet = btrGetEqual(hgDBCxf, tgDBCxf, igDBCxfRecLen, tgDBCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                '    If ilRet = BTRV_ERR_NONE Then
'                '        ilRet = btrDelete(hgDBCxf)
'                '    End If
'                'End If
'                ilRet = gDeleteCxf(tgDBChf.lCxfMerch)
'                'If tgDBChf.lCxfProm > 0 Then
'                '    igDBCxfRecLen = Len(tgDBCxf) '5011
'                '    tgDBCxfSrchKey.lCode = tgDBChf.lCxfProm
'                '    ilRet = btrGetEqual(hgDBCxf, tgDBCxf, igDBCxfRecLen, tgDBCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                '    If ilRet = BTRV_ERR_NONE Then
'                '        ilRet = btrDelete(hgDBCxf)
'                '    End If
'                'End If
'                ilRet = gDeleteCxf(tgDBChf.lCxfProm)
'                'Remove chf
'                ilRet = btrDelete(hgDBChf)
                ilRet = mDeleteChf()
                If ilRet = BTRV_ERR_NONE Then
                    llPropCount = llPropCount + 1
                    llChf = llChf + 1
                Else
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        llChf = llChf + 1
                    End If
                End If
            Else
                llChf = llChf + 1
            End If
        Loop
    Next ilPass
    Erase llPropChfCode
    mRemoveProposals = True
    Exit Function
End Function

Private Function mRemoveAdf(llRetainDate As Long, llAdfCount As Long) As Integer
    Dim ilAdf As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim llDate As Long
    Dim slType As String
    Dim ilPass As Integer
    
    ReDim ilAdfCode(0 To 0) As Integer
    ilRet = btrGetFirst(hgDBAdf, tgDBAdf, igDBAdfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        gUnpackDateLong tgDBAdf.iDateEntrd(0), tgDBAdf.iDateEntrd(1), llDate
        If llDate < llRetainDate Then
            ilFound = False
            tgDBChfSrchKey2.iCode = tgDBAdf.iCode
            ilRet = btrGetEqual(hgDBChf, tgDBChf, igDBChfRecLen, tgDBChfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilFound = True
            End If
            tgDBRvfSrchKey1.iCode = tgDBAdf.iCode
            ilRet = btrGetEqual(hgDBRvf, tgDBRvf, igDBRvfRecLen, tgDBRvfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilFound = True
            End If
            tgDBRvfSrchKey1.iCode = tgDBAdf.iCode
            ilRet = btrGetEqual(hgDBPhf, tgDBRvf, igDBRvfRecLen, tgDBRvfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilFound = True
            End If
            If Not ilFound Then
                ilFound = gIICodeRefExist(csiArchive, tgDBAdf.iCode, "Cif.Btr", "CifAdfCode")
            End If
            If Not ilFound Then
                ilFound = gIICodeRefExist(csiArchive, tgDBAdf.iCode, "Pjf.Btr", "PjfAdfCode")
            End If
            If Not ilFound Then
                ilFound = gIICodeRefExist(csiArchive, tgDBAdf.iCode, "Bof.Btr", "BofAdfCode")
            End If
            If Not ilFound Then
                ilAdfCode(UBound(ilAdfCode)) = tgDBAdf.iCode
                ReDim Preserve ilAdfCode(0 To UBound(ilAdfCode) + 1) As Integer
            End If
        End If
        ilRet = btrGetNext(hgDBAdf, tgDBAdf, igDBAdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilAdf = 0
    Do While ilAdf < UBound(ilAdfCode)
        tgDBAdfSrchKey.iCode = ilAdfCode(ilAdf)
        ilRet = btrGetEqual(hgDBAdf, tgDBAdf, igDBAdfRecLen, tgDBAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            'Remove SIF
            Do
                tgDBSifSrchKey1.iCode = tgDBAdf.iCode
                ilRet = btrGetEqual(hgDBSif, tgDBSif, igDBSifRecLen, tgDBSifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBSif)
            Loop
            'Remove CDF
            Do
                tgDBCdfSrchKey1.iCode = tgDBAdf.iCode
                igDBCdfRecLen = Len(tgDBCdf)
                ilRet = btrGetEqual(hgDBCdf, tgDBCdf, igDBCdfRecLen, tgDBCdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBCdf)
            Loop
            'Remove CRF, csf
            For ilPass = 0 To 2 Step 1
                DoEvents
                If ilPass = 0 Then
                    slType = "O"
                ElseIf ilPass = 1 Then
                    slType = "C"
                Else
                    slType = "A"
                End If
                Do
                    DoEvents
                    tgDBCrfSrchKey1.sRotType = slType
                    tgDBCrfSrchKey1.iEtfCode = 0
                    tgDBCrfSrchKey1.iEnfCode = 0
                    tgDBCrfSrchKey1.iAdfCode = tgDBAdf.iCode
                    tgDBCrfSrchKey1.lChfCode = 0
                    tgDBCrfSrchKey1.lFsfCode = 0
                    tgDBCrfSrchKey1.iVefCode = 0
                    tgDBCrfSrchKey1.iRotNo = 32000
                    ilRet = btrGetGreaterOrEqual(hgDBCrf, tgDBCrf, igDBCrfRecLen, tgDBCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
                    If (ilRet = BTRV_ERR_NONE) And (tgDBCrf.sRotType = slType) And (tgDBCrf.iEtfCode = 0) And (tgDBCrf.iEnfCode = 0) And (tgDBCrf.iAdfCode = tgDBAdf.iCode) Then
                        If tgDBCrf.lCsfCode > 0 Then
                            tgDBCsfSrchKey.lCode = tgDBCrf.lCsfCode
                            igDBCsfRecLen = Len(tgDBCsf)
                            ilRet = btrGetEqual(hgDBCsf, tgDBCsf, igDBCsfRecLen, tgDBCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet <> BTRV_ERR_NONE Then
                                Exit Do
                            End If
                            ilRet = btrDelete(hgDBCsf)
                        End If
                        ilRet = btrDelete(hgDBCrf)
                    Else
                        Exit Do
                    End If
                Loop
            Next ilPass
            ''Remove Pnf
            'Do
            '    tgDBPnfSrchKey2.iCode = tgDBAdf.iCode
            '    ilRet = btrGetEqual(hgDBPnf, tgDBPnf, igDBPnfRecLen, tgDBPnfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
            '    If ilRet <> BTRV_ERR_NONE Then
            '        Exit Do
            '    End If
            '    ilRet = btrDelete(hgDBPnf)
            'Loop
            'Remove Prf
            Do
                tgDBPrfSrchKey1.iCode = tgDBAdf.iCode
                ilRet = btrGetEqual(hgDBPrf, tgDBSif, igDBPrfRecLen, tgDBPrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBPrf)
            Loop
            ilRet = btrDelete(hgDBAdf)
            If ilRet = BTRV_ERR_NONE Then
                llAdfCount = llAdfCount + 1
                ilAdf = ilAdf + 1
                'Remove Pnf. 12/13/17-Moved here to check if two advertisers referenced the same buyer which should not happen
                Do
                    tgDBPnfSrchKey2.iCode = tgDBAdf.iCode
                    ilRet = btrGetEqual(hgDBPnf, tgDBPnf, igDBPnfRecLen, tgDBPnfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                    If Not gIICodeRefExist(csiArchive, tgDBPnf.iCode, "Adf.Btr", "adfPnfBuyer") Then
                        ilRet = btrDelete(hgDBPnf)
                    Else
                        Exit Do
                    End If
                Loop
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    ilAdf = ilAdf + 1
                End If
            End If
        Else
            ilAdf = ilAdf + 1
        End If
    Loop
    Erase ilAdfCode
    mRemoveAdf = True
End Function

Private Function mRemoveAgf(llRetainDate As Long, llAgfCount As Long) As Integer
    Dim ilAgf As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim llDate As Long
    
    ReDim ilAgfCode(0 To 0) As Integer
    ilRet = btrGetFirst(hgDBAgf, tgDBAgf, igDBAgfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        gUnpackDateLong tgDBAgf.iDateEntrd(0), tgDBAgf.iDateEntrd(1), llDate
        If llDate < llRetainDate Then
            ilFound = False
            tgDBRvfSrchKey0.iCode = tgDBAgf.iCode
            ilRet = btrGetEqual(hgDBRvf, tgDBRvf, igDBRvfRecLen, tgDBRvfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilFound = True
            End If
            tgDBRvfSrchKey0.iCode = tgDBAgf.iCode
            ilRet = btrGetEqual(hgDBPhf, tgDBRvf, igDBRvfRecLen, tgDBRvfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilFound = True
            End If
            If Not ilFound Then
                ilRet = gIICodeRefExist(csiArchive, tgDBAgf.iCode, "Adf.Btr", "ADFAGFCODE")
            End If
            If Not ilFound Then
                ilFound = gIICodeRefExist(csiArchive, tgDBAgf.iCode, "Chf.Btr", "CHFAGFCODE")
            End If
            If Not ilFound Then
                ilAgfCode(UBound(ilAgfCode)) = tgDBAgf.iCode
                ReDim Preserve ilAgfCode(0 To UBound(ilAgfCode) + 1) As Integer
            End If
        End If
        ilRet = btrGetNext(hgDBAgf, tgDBAgf, igDBAgfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilAgf = 0
    Do While ilAgf < UBound(ilAgfCode)
        tgDBAgfSrchKey.iCode = ilAgfCode(ilAgf)
        ilRet = btrGetEqual(hgDBAgf, tgDBAgf, igDBAgfRecLen, tgDBAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            'Remove CDF
            Do
                tgDBCdfSrchKey0.iCode = tgDBAgf.iCode
                igDBCdfRecLen = Len(tgDBCdf)
                ilRet = btrGetEqual(hgDBCdf, tgDBCdf, igDBCdfRecLen, tgDBCdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBCdf)
            Loop
            ''Remove Pnf
            'Do
            '    tgDBPnfSrchKey1.iCode = tgDBAgf.iCode
            '    ilRet = btrGetEqual(hgDBPnf, tgDBPnf, igDBPnfRecLen, tgDBPnfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            '    If ilRet <> BTRV_ERR_NONE Then
            '        Exit Do
            '    End If
            '    ilRet = btrDelete(hgDBPnf)
            'Loop
            ilRet = btrDelete(hgDBAgf)
            If ilRet = BTRV_ERR_NONE Then
                llAgfCount = llAgfCount + 1
                ilAgf = ilAgf + 1
                'Remove Pnf 12/13/17-Moved here to check if two agencies referenced the same buyer which should not happen
                Do
                    tgDBPnfSrchKey1.iCode = tgDBAgf.iCode
                    ilRet = btrGetEqual(hgDBPnf, tgDBPnf, igDBPnfRecLen, tgDBPnfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                    If Not gIICodeRefExist(csiArchive, tgDBPnf.iCode, "Agf.Btr", "agfPnfBuyer") Then
                        ilRet = btrDelete(hgDBPnf)
                    Else
                        Exit Do
                    End If
                Loop
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    ilAgf = ilAgf + 1
                End If
            End If
        Else
            ilAgf = ilAgf + 1
        End If
    Loop
    Erase ilAgfCode
    mRemoveAgf = True
End Function


Private Function mGatherAndSortCrf(llFirstAllowedCopyDate As Long, llCrfEarliestDate As Long, llCrfLatestDate As Long, ilMissingDate As Integer) As Integer
    Dim ilRet As Integer
    Dim llDate As Long
    Dim slDate As String
    
    llCrfEarliestDate = -1
    llCrfLatestDate = -1
    ilMissingDate = False
    ReDim tmCrfArchiveInfo(0 To 0) As CRFARCHIVEINFO
    ilRet = btrGetFirst(hgDBCrf, tgDBCrf, igDBCrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        If (imCancelled) Then
            lacCopyStatus(1).Caption = "Cancelled"
            DoEvents
            Exit Do
        End If
        gUnpackDateLong tgDBCrf.iEndDate(0), tgDBCrf.iEndDate(1), llDate
        If llDate < llFirstAllowedCopyDate Then
            If llDate <= 0 Then
                ilMissingDate = True
            Else
                If llCrfEarliestDate = -1 Then
                    llCrfEarliestDate = llDate
                    llCrfLatestDate = llDate
                End If
                If llDate < llCrfEarliestDate Then
                    llCrfEarliestDate = llDate
                End If
                If llDate > llCrfLatestDate Then
                    llCrfLatestDate = llDate
                End If
            End If
            slDate = Trim$(Str$(llDate))
            Do While Len(slDate) < 5
                slDate = "0" & slDate
            Loop
            tmCrfArchiveInfo(UBound(tmCrfArchiveInfo)).sKey = slDate
            tmCrfArchiveInfo(UBound(tmCrfArchiveInfo)).lCrfCode = tgDBCrf.lCode
            tmCrfArchiveInfo(UBound(tmCrfArchiveInfo)).lDate = llDate
            ReDim Preserve tmCrfArchiveInfo(0 To UBound(tmCrfArchiveInfo) + 1) As CRFARCHIVEINFO
        Else
            tgDBCnfSrchKey.lCrfCode = tgDBCrf.lCode
            tgDBCnfSrchKey.iInstrNo = 0
            ilRet = btrGetGreaterOrEqual(hgDBCnf, tgDBCnf, igDBCnfRecLen, tgDBCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While ilRet = BTRV_ERR_NONE
                If tgDBCnf.lCrfCode <> tgDBCrf.lCode Then
                    Exit Do
                End If
                If tgDBCnf.lCifCode <= UBound(lmRetainCifCode) Then
                    lmRetainCifCode(tgDBCnf.lCifCode) = lmRetainCifCode(tgDBCnf.lCifCode) + 1
                End If
                ilRet = btrGetNext(hgDBCnf, tgDBCnf, igDBCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
        ilRet = btrGetNext(hgDBCrf, tgDBCrf, igDBCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mGatherAndSortCrf = True
End Function

Private Function mRemoveContracts(llRetainDate As Long, llCntrCount As Long) As Integer
    Dim ilRet As Integer
    Dim llDate As Long
    Dim llCntrNo As Long
    Dim ilRetain As Integer
    Dim llChf As Long
    Dim slEndDate As String
    ReDim llRemoveCntrNo(0 To 0) As Long
    
    llCntrNo = -1
    ilRetain = False
    tgDBChfSrchKey1.lCntrNo = 0
    tgDBChfSrchKey1.iCntRevNo = 32000
    tgDBChfSrchKey1.iPropVer = 32000
    ilRet = btrGetGreaterOrEqual(hgDBChf, tgDBChf, igDBChfRecLen, tgDBChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While ilRet = BTRV_ERR_NONE
        If llCntrNo = -1 Then
            llCntrNo = tgDBChf.lCntrNo
        End If
        If llCntrNo <> tgDBChf.lCntrNo Then
            If Not ilRetain Then
                llRemoveCntrNo(UBound(llRemoveCntrNo)) = llCntrNo
                ReDim Preserve llRemoveCntrNo(0 To UBound(llRemoveCntrNo) + 1) As Long
            End If
            ilRetain = False
            llCntrNo = tgDBChf.lCntrNo
        End If
        If Not ilRetain Then
            gUnpackDateLong tgDBChf.iEndDate(0), tgDBChf.iEndDate(1), llDate
            If llDate >= llRetainDate Then
                ilRetain = True
            Else
                tgDBSdfSrchKey5.lCode = tgDBChf.lCode
                ilRet = btrGetEqual(hgDBSdf, tgDBSdf, igDBSdfRecLen, tgDBSdfSrchKey5, INDEXKEY5, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    ilRetain = True
                End If
            End If
        End If
        ilRet = btrGetNext(hgDBChf, tgDBChf, igDBChfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    If Not ilRetain Then
        llRemoveCntrNo(UBound(llRemoveCntrNo)) = llCntrNo
        ReDim Preserve llRemoveCntrNo(0 To UBound(llRemoveCntrNo) + 1) As Long
    End If
    If imCancelled Then
        llCntrCount = 0
        Erase llRemoveCntrNo
        lacCntrStatus(1).Caption = "Cancelled"
        lacCntrStatus(1).Visible = True
        mRemoveContracts = False
        Exit Function
    End If
    For llChf = 0 To UBound(llRemoveCntrNo) - 1 Step 1
        If imCancelled Then
            llCntrCount = llChf
            Erase llRemoveCntrNo
            lacCntrStatus(1).Caption = "Cancelled"
            lacCntrStatus(1).Visible = True
            mRemoveContracts = False
            Exit Function
        End If
        DoEvents
        tgDBChfSrchKey1.lCntrNo = llRemoveCntrNo(llChf)
        tgDBChfSrchKey1.iCntRevNo = 32000
        tgDBChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hgDBChf, tgDBChf, igDBChfRecLen, tgDBChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tgDBChf.lCntrNo = llRemoveCntrNo(llChf))
            DoEvents
            ilRet = mDeleteChf()
            'If imUserInput Then
                gUnpackDate tgDBChf.iEndDate(0), tgDBChf.iEndDate(1), slEndDate
                gLogMsg "Contract " & tgDBChf.lCntrNo & " Internal  Version " & tgDBChf.iCntRevNo & " End Date " & slEndDate & " chfCode " & tgDBChf.lCode, "csiArchiveRemovalDetail.txt", False
            'End If
            ilRet = btrGetNext(hgDBChf, tgDBChf, igDBChfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
    Next llChf
    llCntrCount = UBound(llRemoveCntrNo)
    Erase llRemoveCntrNo
    lacCntrStatus(1).Caption = "Completed"
    lacCntrStatus(1).Visible = True
    mRemoveContracts = True
End Function

Private Function mDeleteChf() As Integer
    Dim ilRet As Integer
    Dim llLkVsfCode As Long
    Dim ilPass As Integer
    Dim slType As String
    
    'Remove copy
    'Remove CRF, csf
    For ilPass = 0 To 2 Step 1
        DoEvents
        If ilPass = 0 Then
            slType = "O"
        ElseIf ilPass = 1 Then
            slType = "C"
        Else
            slType = "A"
        End If
        Do
            DoEvents
            tgDBCrfSrchKey1.sRotType = slType
            tgDBCrfSrchKey1.iEtfCode = 0
            tgDBCrfSrchKey1.iEnfCode = 0
            tgDBCrfSrchKey1.iAdfCode = tgDBChf.iAdfCode
            tgDBCrfSrchKey1.lChfCode = tgDBChf.lCode
            tgDBCrfSrchKey1.lFsfCode = 0
            tgDBCrfSrchKey1.iVefCode = 0
            tgDBCrfSrchKey1.iRotNo = 32000
            ilRet = btrGetGreaterOrEqual(hgDBCrf, tgDBCrf, igDBCrfRecLen, tgDBCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) And (tgDBCrf.sRotType = slType) And (tgDBCrf.iEtfCode = 0) And (tgDBCrf.iEnfCode = 0) And (tgDBCrf.iAdfCode = tgDBChf.iAdfCode) And (tgDBCrf.lChfCode = tgDBChf.lCode) Then
                ilRet = mRemoveCnf(tgDBCrf.lCode)
                If tgDBCrf.lCsfCode > 0 Then
                    tgDBCsfSrchKey.lCode = tgDBCrf.lCsfCode
                    igDBCsfRecLen = Len(tgDBCsf)
                    ilRet = btrGetEqual(hgDBCsf, tgDBCsf, igDBCsfRecLen, tgDBCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                    ilRet = btrDelete(hgDBCsf)
                End If
                ilRet = btrDelete(hgDBCrf)
            Else
                Exit Do
            End If
        Loop
    Next ilPass
    'Remove SBF
    Do
        tgDBSbfSrchKey0.lChfCode = tgDBChf.lCode
        tgDBSbfSrchKey0.iDate(0) = 0
        tgDBSbfSrchKey0.iDate(1) = 0
        tgDBSbfSrchKey0.sTranType = ""
        ilRet = btrGetGreaterOrEqual(hgDBSbf, tgDBSbf, igDBSbfRecLen, tgDBSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        If tgDBSbf.lChfCode <> tgDBChf.lCode Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBSbf)
    Loop
    'Remove clf, cff, cxf and cgf
    Do
        tgDBClfSrchKey1.lChfCode = tgDBChf.lCode
        tgDBClfSrchKey1.iVefCode = 0
        ilRet = btrGetGreaterOrEqual(hgDBClf, tgDBClf, igDBClfRecLen, tgDBClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        If tgDBClf.lChfCode <> tgDBChf.lCode Then
            Exit Do
        End If
        'Remove Cff
        Do
            tgDBCffSrchKey.lChfCode = tgDBChf.lCode
            tgDBCffSrchKey.iClfLine = tgDBClf.iLine
            tgDBCffSrchKey.iCntRevNo = 32000
            tgDBCffSrchKey.iPropVer = 32000
            tgDBCffSrchKey.iStartDate(0) = 0
            tgDBCffSrchKey.iStartDate(1) = 0
            ilRet = btrGetGreaterOrEqual(hgDBCff, tgDBCff, igDBCffRecLen, tgDBCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            If tgDBCff.lChfCode <> tgDBChf.lCode Then
                Exit Do
            End If
            If tgDBCff.iClfLine <> tgDBClf.iLine Then
                Exit Do
            End If
            ilRet = btrDelete(hgDBCff)
        Loop
        'Remove cgf
        Do
            tgDBCgfSrchKey1.lClfCode = tgDBClf.lCode
            ilRet = btrGetGreaterOrEqual(hgDBCgf, tgDBCgf, igDBCgfRecLen, tgDBCgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            If tgDBCgf.lClfCode <> tgDBClf.lCode Then
                Exit Do
            End If
            ilRet = btrDelete(hgDBCgf)
        Loop
        'Remove cxf
        If tgDBClf.lCxfCode > 0 Then
            igDBCxfRecLen = Len(tgDBCxf) '5011
            tgDBCxfSrchKey.lCode = tgDBClf.lCxfCode
            ilRet = btrGetEqual(hgDBCxf, tgDBCxf, igDBCxfRecLen, tgDBCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hgDBCxf)
            End If
        End If
        ilRet = btrDelete(hgDBClf)
    Loop
    'Remove vsf
    '12/16/08:  Must be retained as same record can be referenced by multi-contracts
    'If tgDBChf.lVefCode < 0 Then
    '    tgDBVsfSrchKey.lCode = -tgDBChf.lVefCode
    '    ilRet = btrGetEqual(hgDBVsf, tgDBVsf, igDBVsfRecLen, tgDBVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '    If ilRet = BTRV_ERR_NONE Then
    '        llLkVsfCode = tgDBVsf.lLkVsfCode
    '        ilRet = btrDelete(hgDBVsf)
    '        Do While llLkVsfCode > 0
    '            tgDBVsfSrchKey.lCode = llLkVsfCode
    '            ilRet = btrGetEqual(hgDBVsf, tgDBVsf, igDBVsfRecLen, tgDBVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '            If ilRet <> BTRV_ERR_NONE Then
    '                Exit Do
    '            End If
    '            llLkVsfCode = tgDBVsf.lLkVsfCode
    '            ilRet = btrDelete(hgDBVsf)
    '        Loop
    '    End If
    'End If
    'Remove msf and mgf
    Do
        tgDBMsfSrchKey2.lChfCode = tgDBChf.lCode
        ilRet = btrGetEqual(hgDBMsf, tgDBMsf, igDBMsfRecLen, tgDBMsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        'Remove Mgf
        Do
            tgDBMgfSrchKey1.lMsfCode = tgDBMsf.lCode
            tgDBMgfSrchKey1.iGameNo = 0
            ilRet = btrGetGreaterOrEqual(hgDBMgf, tgDBMgf, igDBMgfRecLen, tgDBMgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            If tgDBMgf.lMsfCode <> tgDBMsf.lCode Then
                Exit Do
            End If
            ilRet = btrDelete(hgDBMgf)
        Loop
        ilRet = btrDelete(hgDBMsf)
    Loop
    'Remove cxf
    ilRet = gDeleteCxf(tgDBChf.lCxfCanc)
    ilRet = gDeleteCxf(tgDBChf.lCxfChgR)
    ilRet = gDeleteCxf(tgDBChf.lCxfCode)
    ilRet = gDeleteCxf(tgDBChf.lCxfInt)
    ilRet = gDeleteCxf(tgDBChf.lCxfMerch)
    ilRet = gDeleteCxf(tgDBChf.lCxfProm)
    'Remove chf
    ilRet = btrDelete(hgDBChf)
    mDeleteChf = ilRet
End Function

Private Function mRemovePayRevHist(llRetainDate As Long, llPayRevHistCount As Long) As Integer
    Dim ilRet As Integer
    
    Do
        If imCancelled Then
            lacPayRevHistStatus(1).Caption = "Cancelled"
            lacPayRevHistStatus(1).Visible = True
            mRemovePayRevHist = False
            Exit Function
        End If
        gPackDateLong llRetainDate - 1, tgDBPhfSrchKey3.iTranDate(0), tgDBPhfSrchKey3.iTranDate(1)
        ilRet = btrGetLessOrEqual(hgDBPhf, tgDBRvf, igDBRvfRecLen, tgDBPhfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBPhf)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemovePayRevHist: Delete error ", ilRet
                mRemovePayRevHist = False
                Exit Function
            End If
        Else
            llPayRevHistCount = llPayRevHistCount + 1
        End If
    Loop
    lacPayRevHistStatus(1).Caption = "Completed"
    lacPayRevHistStatus(1).Visible = True
    mRemovePayRevHist = True
End Function

Private Function mRemoveUnusedVefRcfSlf() As Integer
    Dim ilSlf As Integer
    Dim ilVsf As Integer
    Dim ilRet As Integer
    Dim llLkVsfCode As Long
    Dim ilYear As Integer
    Dim ilVef As Integer
    Dim ilRcf As Integer
    
    lacMsg.Caption = "Checking for Unused Vehicles, Rate Cards and Salespersons"
    ilYear = Year(gAdjYear(Format$(gNow(), "m/d/yy"))) - 1
    ilRet = btrGetLast(hgDBVef, tgDBVef, igDBVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        ReDim slVefUsed(0 To tgDBVef.iCode + 1) As String * 1
    Else
        ReDim slVefUsed(0 To 0) As String * 1
    End If
    ilRet = btrGetLast(hgDBSlf, tgDBSlf, igDBSlfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        ReDim slSlfUsed(0 To tgDBSlf.iCode + 1) As String * 1
    Else
        ReDim slSlfUsed(0 To 0) As String * 1
    End If
    ilRet = btrGetLast(hgDBRcf, tgDBRcf, igDBRcfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        ReDim slRcfUsed(0 To tgDBRcf.iCode + 1) As String * 1
    Else
        ReDim slRcfUsed(0 To 0) As String * 1
    End If
    'Loop on contracts
    ilRet = btrGetFirst(hgDBChf, tgDBChf, igDBChfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If imCancelled Then
            mRemoveUnusedVefRcfSlf = False
            Exit Function
        End If
        For ilSlf = LBound(tgDBChf.iSlfCode) To UBound(tgDBChf.iSlfCode) Step 1
            DoEvents
            If tgDBChf.iSlfCode(ilSlf) > 0 Then
                If tgDBChf.iSlfCode(ilSlf) <= UBound(slSlfUsed) Then
                    slSlfUsed(tgDBChf.iSlfCode(ilSlf)) = "Y"
                End If
            End If
        Next ilSlf
        If tgDBChf.iRcfCode <= UBound(slRcfUsed) Then
            slRcfUsed(tgDBChf.iRcfCode) = "Y"
        End If
        If tgDBChf.lVefCode < 0 Then
            llLkVsfCode = -tgDBChf.lVefCode
            Do While llLkVsfCode > 0
                DoEvents
                tgDBVsfSrchKey.lCode = llLkVsfCode
                ilRet = btrGetEqual(hgDBVsf, tgDBVsf, igDBVsfRecLen, tgDBVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                For ilVsf = LBound(tgDBVsf.iFSCode) To UBound(tgDBVsf.iFSCode) Step 1
                    DoEvents
                    If tgDBVsf.iFSCode(ilVsf) > 0 Then
                        If tgDBVsf.iFSCode(ilVsf) <= UBound(slVefUsed) Then
                            slVefUsed(tgDBVsf.iFSCode(ilVsf)) = "Y"
                        End If
                    End If
                Next ilVsf
                llLkVsfCode = tgDBVsf.lLkVsfCode
            Loop
        Else
            slVefUsed(tgDBChf.lVefCode) = "Y"
        End If
        ilRet = btrGetNext(hgDBChf, tgDBChf, igDBChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    'Scan lines to catch packages lines which are never placed into vsf
    ilRet = btrGetFirst(hgDBClf, tgDBClf, igDBClfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If imCancelled Then
            mRemoveUnusedVefRcfSlf = False
            Exit Function
        End If
        If tgDBClf.iVefCode <= UBound(slVefUsed) Then
            slVefUsed(tgDBClf.iVefCode) = "Y"
        End If
        ilRet = btrGetNext(hgDBClf, tgDBClf, igDBClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    'Loop on PHF
    ilRet = btrGetFirst(hgDBPhf, tgDBRvf, igDBRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If imCancelled Then
            mRemoveUnusedVefRcfSlf = False
            Exit Function
        End If
        If (tgDBRvf.iSlfCode <= UBound(slSlfUsed)) And (tgDBRvf.iSlfCode > 0) Then
            slSlfUsed(tgDBRvf.iSlfCode) = "Y"
        End If
        If (tgDBRvf.iAirVefCode <= UBound(slVefUsed)) And (tgDBRvf.iAirVefCode > 0) Then
            slVefUsed(tgDBRvf.iAirVefCode) = "Y"
        End If
        If (tgDBRvf.iBillVefCode <= UBound(slVefUsed)) And (tgDBRvf.iBillVefCode > 0) Then
            slVefUsed(tgDBRvf.iBillVefCode) = "Y"
        End If
        ilRet = btrGetNext(hgDBPhf, tgDBRvf, igDBRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    'Loop on RVF
    ilRet = btrGetFirst(hgDBRvf, tgDBRvf, igDBRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If imCancelled Then
            mRemoveUnusedVefRcfSlf = False
            Exit Function
        End If
        If (tgDBRvf.iSlfCode <= UBound(slSlfUsed)) And (tgDBRvf.iSlfCode > 0) Then
            slSlfUsed(tgDBRvf.iSlfCode) = "Y"
        End If
        If (tgDBRvf.iAirVefCode <= UBound(slVefUsed)) And (tgDBRvf.iAirVefCode > 0) Then
            slVefUsed(tgDBRvf.iAirVefCode) = "Y"
        End If
        If (tgDBRvf.iBillVefCode <= UBound(slVefUsed)) And (tgDBRvf.iBillVefCode > 0) Then
            slVefUsed(tgDBRvf.iBillVefCode) = "Y"
        End If
        ilRet = btrGetNext(hgDBRvf, tgDBRvf, igDBRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    'Loop on Budgets
    ilRet = btrGetFirst(hgDBBvf, tgDBBvf, igDBBvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If imCancelled Then
            mRemoveUnusedVefRcfSlf = False
            Exit Function
        End If
        If (tgDBBvf.iVefCode <= UBound(slVefUsed)) And (tgDBBvf.iVefCode > 0) Then
            slVefUsed(tgDBBvf.iVefCode) = "Y"
        End If
        ilRet = btrGetNext(hgDBBvf, tgDBBvf, igDBBvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilRet = btrGetFirst(hgDBBsf, tgDBBsf, igDBBsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If imCancelled Then
            mRemoveUnusedVefRcfSlf = False
            Exit Function
        End If
        If (tgDBBsf.iSlfCode <= UBound(slSlfUsed)) And (tgDBBsf.iSlfCode > 0) Then
            slSlfUsed(tgDBBsf.iSlfCode) = "Y"
        End If
        ilRet = btrGetNext(hgDBBsf, tgDBBsf, igDBBsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    'Loop on SCF
    ilRet = btrGetFirst(hgDBScf, tgDBScf, igDBScfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If imCancelled Then
            mRemoveUnusedVefRcfSlf = False
            Exit Function
        End If
        If (tgDBScf.iSlfCode <= UBound(slSlfUsed)) And (tgDBScf.iSlfCode > 0) Then
            slSlfUsed(tgDBScf.iSlfCode) = "Y"
        End If
        If (tgDBScf.iVefCode <= UBound(slVefUsed)) And (tgDBScf.iVefCode > 0) Then
            slVefUsed(tgDBScf.iVefCode) = "Y"
        End If
        ilRet = btrGetNext(hgDBScf, tgDBScf, igDBScfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    'Loop on Media Code
    ilRet = btrGetFirst(hgDBMcf, tgDBMcf, igDBMcfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If imCancelled Then
            mRemoveUnusedVefRcfSlf = False
            Exit Function
        End If
        If (tgDBMcf.iVefCode <= UBound(slVefUsed)) And (tgDBMcf.iVefCode > 0) Then
            slVefUsed(tgDBMcf.iVefCode) = "Y"
        End If
        ilRet = btrGetNext(hgDBMcf, tgDBMcf, igDBMcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    
    'Remove unused and dormant Vehicle
    lacMsg.Caption = "Removing Unused Vehicles"
    ReDim ilDeleteCode(0 To 0) As Integer
    ilRet = btrGetFirst(hgDBVef, tgDBVef, igDBVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If tgDBVef.sState = "D" Then
            If (tgDBVef.iCode <= UBound(slVefUsed)) And (tgDBVef.iCode > 0) Then
                If slVefUsed(tgDBVef.iCode) <> "Y" Then
                    'Test if referenced by affiliate spots
                    If gGetEarliestAffSpotDate(hgDBAst, hgDBAtt, tgDBVef.iCode) = 0 Then
                        ilDeleteCode(UBound(ilDeleteCode)) = tgDBVef.iCode
                        ReDim Preserve ilDeleteCode(0 To UBound(ilDeleteCode) + 1) As Integer
                    End If
                End If
            End If
        End If
        ilRet = btrGetNext(hgDBVef, tgDBVef, igDBVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If imCancelled Then
        mRemoveUnusedVefRcfSlf = False
        Exit Function
    End If
    ilRet = mRemoveVef(ilDeleteCode())
    If imCancelled Then
        mRemoveUnusedVefRcfSlf = False
        Exit Function
    End If
    If UBound(ilDeleteCode) > 0 Then
        gLogMsg "Total of " & Trim$(Str$(UBound(ilDeleteCode))) & " Vehicles Removed", "csiArchive.Txt", False
    End If
    
    'Remove unused and marked dormant salespersons
    lacMsg.Caption = "Removing Unused Salespersons"
    ReDim ilDeleteCode(0 To 0) As Integer
    ilRet = btrGetFirst(hgDBSlf, tgDBSlf, igDBSlfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If imCancelled Then
            mRemoveUnusedVefRcfSlf = False
            Exit Function
        End If
        If tgDBSlf.sState = "D" Then
            If (tgDBSlf.iCode <= UBound(slSlfUsed)) And (tgDBSlf.iCode > 0) Then
                If slSlfUsed(tgDBSlf.iCode) <> "Y" Then
                    ilDeleteCode(UBound(ilDeleteCode)) = tgDBSlf.iCode
                    ReDim Preserve ilDeleteCode(0 To UBound(ilDeleteCode) + 1) As Integer
                End If
            End If
        End If
        ilRet = btrGetNext(hgDBSlf, tgDBSlf, igDBSlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If imCancelled Then
        mRemoveUnusedVefRcfSlf = False
        Exit Function
    End If
    'Remove slf, pjf, bsf, scf
    'Clear references
    ilRet = mClearSlfRef(ilDeleteCode())
    If imCancelled Then
        mRemoveUnusedVefRcfSlf = False
        Exit Function
    End If
    For ilSlf = 0 To UBound(ilDeleteCode) - 1 Step 1
        If imCancelled Then
            mRemoveUnusedVefRcfSlf = False
            Exit Function
        End If
        ilRet = mRemoveSlf(ilDeleteCode(ilSlf))
    Next ilSlf
    If UBound(ilDeleteCode) > 0 Then
        gLogMsg "Total of " & Trim$(Str$(UBound(ilDeleteCode))) & " Salespersons Removed", "csiArchive.Txt", False
    End If
    'Remove unused Rate card and items
    lacMsg.Caption = "Removing Unused Rate Cards"
    ReDim ilDeleteCode(0 To 0) As Integer
    ilRet = btrGetFirst(hgDBRcf, tgDBRcf, igDBRcfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If imCancelled Then
            mRemoveUnusedVefRcfSlf = False
            Exit Function
        End If
        If tgDBRcf.iYear < ilYear Then
            If (tgDBRcf.iCode <= UBound(slRcfUsed)) And (tgDBRcf.iCode > 0) Then
                If slRcfUsed(tgDBRcf.iCode) <> "Y" Then
                    ilDeleteCode(UBound(ilDeleteCode)) = tgDBRcf.iCode
                    ReDim Preserve ilDeleteCode(0 To UBound(ilDeleteCode) + 1) As Integer
                End If
            End If
        End If
        ilRet = btrGetNext(hgDBRcf, tgDBRcf, igDBRcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    'Remove Rcf, Rif
    For ilRcf = 0 To UBound(ilDeleteCode) - 1 Step 1
        If imCancelled Then
            mRemoveUnusedVefRcfSlf = False
            Exit Function
        End If
        ilRet = mRemoveRcf(ilDeleteCode(ilRcf))
    Next ilRcf
    If UBound(ilDeleteCode) > 0 Then
        gLogMsg "Total of " & Trim$(Str$(UBound(ilDeleteCode))) & " Rate Cards Removed", "csiArchive.Txt", False
    End If
    Erase ilDeleteCode
    Erase slVefUsed
    Erase slSlfUsed
    Erase slRcfUsed
    mRemoveUnusedVefRcfSlf = True
End Function

Private Function mRemoveRcf(ilRcfCode As Integer) As Integer
    Dim ilRet As Integer
    
    Do
        tgDBRifSrchKey0.iCode = ilRcfCode
        ilRet = btrGetEqual(hgDBRif, tgDBRif, igDBRifRecLen, tgDBRifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBRif)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveRcf: Delete RIF error ", ilRet
                mRemoveRcf = False
                Exit Function
            End If
        End If
    Loop
    Do
        tgDBRcfSrchKey0.iCode = ilRcfCode
        ilRet = btrGetEqual(hgDBRcf, tgDBRcf, igDBRcfRecLen, tgDBRcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBRcf)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveRcf: Delete RCF error ", ilRet
                mRemoveRcf = False
                Exit Function
            End If
        End If
        'If imUserInput Then
            gLogMsg "Rate Card " & tgDBRcf.sName & " Year " & tgDBRcf.iYear & " rcfCode " & tgDBRcf.iCode, "csiArchiveRemovalDetail.txt", False
        'End If
    Loop
    mRemoveRcf = True
End Function

Private Function mRemoveSlf(ilSlfCode As Integer) As Integer
    Dim ilRet As Integer
    
    Do
        tgDBScfSrchKey1.iCode = ilSlfCode
        ilRet = btrGetEqual(hgDBScf, tgDBScf, igDBScfRecLen, tgDBScfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBScf)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveSlf: Delete SCF error ", ilRet
                mRemoveSlf = False
                Exit Function
            End If
        End If
    Loop
    Do
        tgDBPjfSrchKey0.iSlfCode = ilSlfCode
        tgDBPjfSrchKey0.iRolloverDate(0) = 0
        tgDBPjfSrchKey0.iRolloverDate(1) = 0
        ilRet = btrGetEqual(hgDBPjf, tgDBPjf, igDBPjfRecLen, tgDBPjfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBPjf)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveSlf: Delete PJF error ", ilRet
                mRemoveSlf = False
                Exit Function
            End If
        End If
    Loop
    Do
        tgDBSlfSrchKey0.iCode = ilSlfCode
        ilRet = btrGetEqual(hgDBSlf, tgDBSlf, igDBSlfRecLen, tgDBSlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBSlf)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveSlf: Delete error ", ilRet
                mRemoveSlf = False
                Exit Function
            End If
        End If
        'If imUserInput Then
            gLogMsg "Salesperson " & tgDBSlf.sFirstName & " " & tgDBSlf.sLastName & " slfCode " & tgDBSlf.iCode, "csiArchiveRemovalDetail.txt", False
        'End If
    Loop
End Function

Private Function mBackupDatabase() As Boolean
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim ilTotalHours As Integer
    Dim slTrafficINIPathName As String
    Dim smCSIServerINIFile As String
    Dim slStatus As String
    Dim slStartDateTime As String
    Dim slCurDateTime As String
    
    'imBackupDB = True
    If Not imBackupDB Then
        mBackupDatabase = True
        Exit Function
    End If
    On Error GoTo ErrHand
    mBackupDatabase = False
    smCSIServerINIFile = sgExePath & "\CSI_Server.ini"
    lacMsg.Caption = "Backing up the database..."
    DoEvents
    
    ' Check to see if a backup is running now.
    If gIsBackupRunning() Then
        If imUserInput Then
            MsgBox "A backup is currently in progress."
        Else
            gLogMsg "A backup is currently in progress.", "csiArchive.txt", False
        End If
        Exit Function
    End If

    slTrafficINIPathName = sgDBPath & "Traffic.ini"
    ilRet = csiStartBackup(sgDBPath, slTrafficINIPathName, 3)
    If ilRet <> 0 Then
        ' Unable to communicate to CSI_Server
        If imUserInput Then
            MsgBox "Backup is not available. Error code = " & ilRet
        Else
            gLogMsg "Backup is not available. Error code = " & ilRet, "csiArchive.txt", False
        End If
        Exit Function
    End If

    ' Backup is now running. Wait here for it to finish.
    lacMsg.Caption = "Database backup is running..."
    slStartDateTime = gNow()
    Do
        DoEvents
        If Not gIsBackupRunning() Then
            ' Backup has finished. Verify it was successful.
            Call gLoadINIValue(smCSIServerINIFile, "MainSettings", "LastBackupStatus", slStatus)
            ilPos = InStr(1, slStatus, "success", vbTextCompare)
            If ilPos < 1 Then
                If imUserInput Then
                    MsgBox "Backup failed. Error Msg = " & slStatus
                Else
                    gLogMsg "Backup failed. Error Msg = " & slStatus, "csiArchive.txt", False
                End If
                Exit Function
            End If
            mBackupDatabase = True
            Exit Function
        End If
        Sleep 2000  ' Sleep 2 seconds
        
        ' To prevent the possibility of getting stuck in this function, timeout after a certain time has elapsed.
        slCurDateTime = gNow()
        ilTotalHours = DateDiff("h", slStartDateTime, slCurDateTime)
        If ilTotalHours > 6 Then
            ' Something went wrong.
            Call gLoadINIValue(smCSIServerINIFile, "MainSettings", "LastBackupStatus", slStatus)
            If imUserInput Then
                MsgBox "Backup did not complete in time. Error Msg = " & slStatus
            Else
                gLogMsg "Backup did not complete in time. Error Msg = " & slStatus, "csiArchive.txt", False
            End If
            Exit Function
        End If
    Loop
    Exit Function

ErrHand:
End Function



Private Function mRemoveBudgets(llRetainDate As Long, llBsfCount As Long, llBvfCount As Long) As Integer
    Dim ilRet As Integer
    Dim ilYear As Integer
    
    If llRetainDate <= 0 Then
        mRemoveBudgets = True
        Exit Function
    End If
    ilYear = Year(gAdjYear(Format$(llRetainDate, "m/d/yy")))
    Do
        If imCancelled Then
            mRemoveBudgets = False
            Exit Function
        End If
        tgDBBsfSrchKey0.iYear = ilYear - 1
        tgDBBsfSrchKey0.iSeqNo = 32000
        tgDBBsfSrchKey0.iSlfCode = 32000
        ilRet = btrGetLessOrEqual(hgDBBsf, tgDBBsf, igDBBsfRecLen, tgDBBsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBBsf)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveBudgets: Delete BSF error ", ilRet
                mRemoveBudgets = False
                Exit Function
            End If
        Else
            llBsfCount = llBsfCount + 1
        End If
    Loop
    Do
        If imCancelled Then
            mRemoveBudgets = False
            Exit Function
        End If
        tgDBBvfSrchKey0.iYear = ilYear - 1
        tgDBBvfSrchKey0.iSeqNo = 32000
        tgDBBvfSrchKey0.iMnfBudget = 32000
        ilRet = btrGetLessOrEqual(hgDBBvf, tgDBBvf, igDBBvfRecLen, tgDBBvfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBBvf)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveBudgets: Delete BVF error ", ilRet
                mRemoveBudgets = False
                Exit Function
            End If
        Else
            llBvfCount = llBvfCount + 1
        End If
    Loop
    mRemoveBudgets = True
End Function

Private Function mRemoveSCF(llRetainDate As Long, llScfCount As Long) As Integer
    Dim llDate As Long
    Dim llScf As Long
    Dim ilRet As Integer
    
    If llRetainDate <= 0 Then
        mRemoveSCF = True
        Exit Function
    End If

    ReDim llScfCode(0 To 0) As Long
    ilRet = btrGetFirst(hgDBScf, tgDBScf, igDBScfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        gUnpackDateLong tgDBScf.iEndDate(0), tgDBScf.iEndDate(1), llDate
        If (llDate < llRetainDate) And (llDate > 0) Then
            llScfCode(UBound(llScfCode)) = tgDBScf.lCode
            ReDim Preserve llScfCode(0 To UBound(llScfCode) + 1) As Long
        End If
        ilRet = btrGetNext(hgDBScf, tgDBScf, igDBScfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    llScf = 0
    Do While llScf < UBound(llScfCode)
        If imCancelled Then
            mRemoveSCF = False
            Exit Function
        End If
        tgDBScfSrchKey0.lCode = llScfCode(llScf)
        ilRet = btrGetEqual(hgDBScf, tgDBScf, igDBScfRecLen, tgDBScfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hgDBScf)
            If ilRet = BTRV_ERR_NONE Then
                llScfCount = llScfCount + 1
                llScf = llScf + 1
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    llScf = llScf + 1
                End If
            End If
        Else
            llScf = llScf + 1
        End If
    Loop
    Erase llScfCode
    mRemoveSCF = True
End Function


Private Function mRemoveVef(ilDeleteCode() As Integer)
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilFound As Integer
    Dim llPjf As Long
    Dim llRif As Long
    Dim llCrf As Long
    Dim llLkPvfCode As Long
    Dim llPvf As Long
    Dim ilPvf As Integer
    Dim ilLvf As Integer
    Dim ilLtf As Integer
    Dim llDrf As Long
    Dim llBof As Long
    Dim llAuf As Long
    
    ReDim llPjfCode(0 To 0) As Long
    ilRet = btrGetFirst(hgDBPjf, tgDBPjf, igDBPjfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If imCancelled Then
            mRemoveVef = False
            Exit Function
        End If
        ilFound = False
        For ilVef = 0 To UBound(ilDeleteCode) - 1 Step 1
            DoEvents
            If tgDBPjf.iVefCode = ilDeleteCode(ilVef) Then
                ilFound = True
                Exit For
            End If
        Next ilVef
        If ilFound Then
            llPjfCode(UBound(llPjfCode)) = tgDBPjf.lCode
            ReDim Preserve llPjfCode(0 To UBound(llPjfCode) + 1) As Long
        End If
        ilRet = btrGetNext(hgDBPjf, tgDBPjf, igDBPjfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    If imCancelled Then
        mRemoveVef = False
        Exit Function
    End If
    llPjf = 0
    Do While llPjf < UBound(llPjfCode)
        DoEvents
        If imCancelled Then
            mRemoveVef = False
            Exit Function
        End If
        tgDBPjfSrchKey1.lCode = llPjfCode(llPjf)
        ilRet = btrGetEqual(hgDBPjf, tgDBPjf, igDBPjfRecLen, tgDBPjfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hgDBPjf)
            If ilRet = BTRV_ERR_NONE Then
                llPjf = llPjf + 1
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    llPjf = llPjf + 1
                End If
            End If
        Else
            llPjf = llPjf + 1
        End If
    Loop
    Erase llPjfCode
    
    
    ReDim llRifCode(0 To 0) As Long
    ilRet = btrGetFirst(hgDBRif, tgDBRif, igDBRifRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If imCancelled Then
            mRemoveVef = False
            Exit Function
        End If
        ilFound = False
        For ilVef = 0 To UBound(ilDeleteCode) - 1 Step 1
            DoEvents
            If tgDBRif.iVefCode = ilDeleteCode(ilVef) Then
                ilFound = True
                Exit For
            End If
        Next ilVef
        If ilFound Then
            llRifCode(UBound(llRifCode)) = tgDBRif.lCode
            ReDim Preserve llRifCode(0 To UBound(llRifCode) + 1) As Long
        End If
        ilRet = btrGetNext(hgDBRif, tgDBRif, igDBRifRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    llRif = 0
    Do While llRif < UBound(llRifCode)
        DoEvents
        If imCancelled Then
            mRemoveVef = False
            Exit Function
        End If
        tgDBRifSrchKey1.lCode = llRifCode(llRif)
        ilRet = btrGetEqual(hgDBRif, tgDBRif, igDBRifRecLen, tgDBRifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hgDBRif)
            If ilRet = BTRV_ERR_NONE Then
                llRif = llRif + 1
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    llRif = llRif + 1
                End If
            End If
        Else
            llRif = llRif + 1
        End If
    Loop
    Erase llRifCode
    
    ReDim llCrfCode(0 To 0) As Long
    ilRet = btrGetFirst(hgDBCrf, tgDBCrf, igDBCrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        ilFound = False
        For ilVef = 0 To UBound(ilDeleteCode) - 1 Step 1
            DoEvents
            If imCancelled Then
                mRemoveVef = False
                Exit Function
            End If
            If tgDBCrf.iVefCode = ilDeleteCode(ilVef) Then
                ilFound = True
                Exit For
            End If
        Next ilVef
        If ilFound Then
            llCrfCode(UBound(llCrfCode)) = tgDBCrf.lCode
            ReDim Preserve llCrfCode(0 To UBound(llCrfCode) + 1) As Long
        End If
        ilRet = btrGetNext(hgDBCrf, tgDBCrf, igDBCrfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    llCrf = 0
    Do While llCrf < UBound(llCrfCode)
        DoEvents
        If imCancelled Then
            mRemoveVef = False
            Exit Function
        End If
        tgDBCrfSrchKey.lCode = llCrfCode(llCrf)
        ilRet = btrGetEqual(hgDBCrf, tgDBCrf, igDBCrfRecLen, tgDBCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hgDBCrf)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = mRemoveCnf(llCrfCode(llCrf))
                llCrf = llCrf + 1
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    llCrf = llCrf + 1
                End If
            End If
        Else
            llCrf = llCrf + 1
        End If
    Loop
    Erase llCrfCode
    
    ReDim llDrfCode(0 To 0) As Long
    ilRet = btrGetFirst(hgDBDrf, tgDBDrf, igDBDrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        ilFound = False
        For ilVef = 0 To UBound(ilDeleteCode) - 1 Step 1
            DoEvents
            If imCancelled Then
                mRemoveVef = False
                Exit Function
            End If
            If tgDBDrf.iVefCode = ilDeleteCode(ilVef) Then
                ilFound = True
                Exit For
            End If
        Next ilVef
        If ilFound Then
            llDrfCode(UBound(llDrfCode)) = tgDBDrf.lCode
            ReDim Preserve llDrfCode(0 To UBound(llDrfCode) + 1) As Long
        End If
        ilRet = btrGetNext(hgDBDrf, tgDBDrf, igDBDrfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    llDrf = 0
    Do While llDrf < UBound(llDrfCode)
        DoEvents
        If imCancelled Then
            mRemoveVef = False
            Exit Function
        End If
        tgDBDrfSrchKey2.lCode = llDrfCode(llDrf)
        ilRet = btrGetEqual(hgDBDrf, tgDBDrf, igDBDrfRecLen, tgDBDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
        
            Do
                DoEvents
                If imCancelled Then
                    mRemoveVef = False
                    Exit Function
                End If
                tgDBDpfSrchKey1.lDrfCode = llDrfCode(llDrf)
                tgDBDpfSrchKey1.iMnfDemo = 0
                ilRet = btrGetGreaterOrEqual(hgDBDpf, tgDBDpf, igDBDpfRecLen, tgDBDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get lSdf current record to obtain date
                If ilRet = BTRV_ERR_NONE Then
                    If tgDBDpf.lDrfCode <> llDrfCode(llDrf) Then
                        Exit Do
                    End If
                    ilRet = btrDelete(hgDBDpf)
                    If ilRet <> BTRV_ERR_NONE Then
                        If ilRet <> BTRV_ERR_CONFLICT Then
                            mRemoveVef = False
                            Exit Function
                        End If
                    End If
                Else
                    Exit Do
                End If
            Loop

            ilRet = btrDelete(hgDBDrf)
            If ilRet = BTRV_ERR_NONE Then
                llDrf = llDrf + 1
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    llDrf = llDrf + 1
                End If
            End If
        Else
            llDrf = llDrf + 1
        End If
    Loop
    Erase llDrfCode
    
    ReDim llBofCode(0 To 0) As Long
    ilRet = btrGetFirst(hgDBBof, tgDBBof, igDBBofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        ilFound = False
        For ilVef = 0 To UBound(ilDeleteCode) - 1 Step 1
            DoEvents
            If imCancelled Then
                mRemoveVef = False
                Exit Function
            End If
            If tgDBBof.iVefCode = ilDeleteCode(ilVef) Then
                ilFound = True
                Exit For
            End If
        Next ilVef
        If ilFound Then
            llBofCode(UBound(llBofCode)) = tgDBBof.lCode
            ReDim Preserve llBofCode(0 To UBound(llBofCode) + 1) As Long
        End If
        ilRet = btrGetNext(hgDBBof, tgDBBof, igDBBofRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    llBof = 0
    Do While llBof < UBound(llBofCode)
        DoEvents
        If imCancelled Then
            mRemoveVef = False
            Exit Function
        End If
        tgDBBofSrchKey0.lCode = llBofCode(llBof)
        ilRet = btrGetEqual(hgDBBof, tgDBBof, igDBBofRecLen, tgDBBofSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hgDBBof)
            If ilRet = BTRV_ERR_NONE Then
                llBof = llBof + 1
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    llBof = llBof + 1
                End If
            End If
        Else
            llBof = llBof + 1
        End If
    Loop
    Erase llBofCode
    
    ReDim ilLtfCode(0 To 0) As Integer
    ilRet = btrGetFirst(hgDBLtf, tgDBLtf, igDBLtfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        ilFound = False
        For ilVef = 0 To UBound(ilDeleteCode) - 1 Step 1
            DoEvents
            If imCancelled Then
                mRemoveVef = False
                Exit Function
            End If
            If tgDBLtf.iVefCode = ilDeleteCode(ilVef) Then
                ilFound = True
                Exit For
            End If
        Next ilVef
        If ilFound Then
            ilLtfCode(UBound(ilLtfCode)) = tgDBLtf.iCode
            ReDim Preserve ilLtfCode(0 To UBound(ilLtfCode) + 1) As Integer
        End If
        ilRet = btrGetNext(hgDBLtf, tgDBLtf, igDBLtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilLtf = 0
    Do While ilLtf < UBound(ilLtfCode)
        DoEvents
        If imCancelled Then
            mRemoveVef = False
            Exit Function
        End If
        tgDBLtfSrchKey.iCode = ilLtfCode(ilLtf)
        ilRet = btrGetEqual(hgDBLtf, tgDBLtf, igDBLtfRecLen, tgDBLtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            Do
                DoEvents
                tgDBLvfSrchKey1.iLtfCode = ilLtfCode(ilLtf)
                tgDBLvfSrchKey1.iVersion = 32000
                ilRet = btrGetGreaterOrEqual(hgDBLvf, tgDBLvf, igDBLvfRecLen, tgDBLvfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBLvf.iLtfCode <> ilLtfCode(ilLtf) Then
                    Exit Do
                End If
                Do
                    DoEvents
                    tgDBLefSrchKey1.lCode = tgDBLvf.lCode
                    ilRet = btrGetEqual(hgDBLef, tgDBLef, igDBLefRecLen, tgDBLefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                    If tgDBLef.lCefCode > 0 Then
                        tgDBCefSrchKey0.lCode = tgDBLef.lCefCode
                        igDBCefRecLen = 1009
                        ilRet = btrGetEqual(hgDBCef, tgDBCef, igDBCefRecLen, tgDBCefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet = BTRV_ERR_NONE Then
                            ilRet = btrDelete(hgDBCef)
                        End If
                    End If
                    If tgDBLef.lEvtIDCefCode > 0 Then
                        tgDBCefSrchKey0.lCode = tgDBLef.lEvtIDCefCode
                        igDBCefRecLen = 1009
                        ilRet = btrGetEqual(hgDBCef, tgDBCef, igDBCefRecLen, tgDBCefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet = BTRV_ERR_NONE Then
                            ilRet = btrDelete(hgDBCef)
                        End If
                    End If
                    ilRet = btrDelete(hgDBLef)
                Loop
                ilRet = btrDelete(hgDBLvf)
            Loop
            ilRet = btrDelete(hgDBLtf)
            If ilRet = BTRV_ERR_NONE Then
                ilLtf = ilLtf + 1
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    ilLtf = ilLtf + 1
                End If
            End If
        Else
            ilLtf = ilLtf + 1
        End If
    Loop
    Erase ilLtfCode
    
    ReDim llAufCode(0 To 0) As Long
    ilRet = btrGetFirst(hgDBAuf, tgDBAuf, igDBAufRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        ilFound = False
        For ilVef = 0 To UBound(ilDeleteCode) - 1 Step 1
            DoEvents
            If imCancelled Then
                mRemoveVef = False
                Exit Function
            End If
            If tgDBAuf.iVefCode = ilDeleteCode(ilVef) Then
                ilFound = True
                Exit For
            End If
        Next ilVef
        If ilFound Then
            llAufCode(UBound(llAufCode)) = tgDBAuf.lCode
            ReDim Preserve llAufCode(0 To UBound(llAufCode) + 1) As Long
        End If
        ilRet = btrGetNext(hgDBAuf, tgDBAuf, igDBAufRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    llAuf = 0
    Do While llAuf < UBound(llAufCode)
        DoEvents
        If imCancelled Then
            mRemoveVef = False
            Exit Function
        End If
        tgDBAufSrchKey0.lCode = llAufCode(llAuf)
        ilRet = btrGetEqual(hgDBAuf, tgDBAuf, igDBAufRecLen, tgDBAufSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hgDBAuf)
            If ilRet = BTRV_ERR_NONE Then
                llAuf = llAuf + 1
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    llAuf = llAuf + 1
                End If
            End If
        Else
            llAuf = llAuf + 1
        End If
    Loop
    Erase llAufCode
    
    For ilVef = 0 To UBound(ilDeleteCode) - 1 Step 1
        DoEvents
        mClearPackageOfVefCode ilDeleteCode(ilVef)
        If imCancelled Then
            mRemoveVef = False
            Exit Function
        End If
    Next ilVef
    
    For ilVef = 0 To UBound(ilDeleteCode) - 1 Step 1
        DoEvents
        tgDBVefSrchKey0.iCode = ilDeleteCode(ilVef)
        ilRet = btrGetEqual(hgDBVef, tgDBVef, igDBVefRecLen, tgDBVefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            'Remove library: lcf, ghf, gsf
            DoEvents
            If imCancelled Then
                mRemoveVef = False
                Exit Function
            End If
            Do
                DoEvents
                tgDBLcfSrchKey2.iVefCode = ilDeleteCode(ilVef)
                tgDBLcfSrchKey2.iLogDate(0) = 0
                tgDBLcfSrchKey2.iLogDate(1) = 0
                ilRet = btrGetGreaterOrEqual(hgDBLcf, tgDBLcf, igDBLcfRecLen, tgDBLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBLcf.iVefCode <> ilDeleteCode(ilVef) Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBLcf)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            
            Do
                DoEvents
                tgDBGsfSrchKey3.iVefCode = ilDeleteCode(ilVef)
                tgDBGsfSrchKey3.iGameNo = 0
                ilRet = btrGetGreaterOrEqual(hgDBGsf, tgDBGsf, igDBGsfRecLen, tgDBGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBGsf.iVefCode <> ilDeleteCode(ilVef) Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBGsf)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            Do
                DoEvents
                tgDBGhfSrchKey1.iVefCode = ilDeleteCode(ilVef)
                ilRet = btrGetEqual(hgDBGhf, tgDBGhf, igDBGhfRecLen, tgDBGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBGhf)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            Do
                DoEvents
                tgDBIsfSrchKey2.iVefCode = ilDeleteCode(ilVef)
                tgDBIsfSrchKey2.iGameNo = 0
                ilRet = btrGetGreaterOrEqual(hgDBIsf, tgDBIsf, igDBIsfRecLen, tgDBIsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBIsf.iVefCode <> ilDeleteCode(ilVef) Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBIsf)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            Do
                DoEvents
                tgDBIhfSrchKey2.iVefCode = ilDeleteCode(ilVef)
                ilRet = btrGetEqual(hgDBIhf, tgDBIhf, igDBIhfRecLen, tgDBLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBIhf)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            
            'Remove SSF
            Do
                DoEvents
                igDBSsfRecLen = Len(tgDBSsf)
                tgDBSsfSrchKey2.iVefCode = ilDeleteCode(ilVef)
                tgDBSsfSrchKey2.iDate(0) = 0
                tgDBSsfSrchKey2.iDate(1) = 0
                ilRet = btrGetGreaterOrEqual(hgDBSsf, tgDBSsf, igDBSsfRecLen, tgDBSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBSsf.iVefCode <> ilDeleteCode(ilVef) Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBSsf)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            'Remove links: vlf, vcf, dlf, egf
            Do
                DoEvents
                tgDBDlfSrchKey0.iVefCode = ilDeleteCode(ilVef)
                tgDBDlfSrchKey0.sAirDay = ""
                tgDBDlfSrchKey0.iStartDate(0) = 0
                tgDBDlfSrchKey0.iStartDate(1) = 0
                tgDBDlfSrchKey0.iAirTime(0) = 0
                tgDBDlfSrchKey0.iAirTime(1) = 0
                'Remove Dlf
                ilRet = btrGetGreaterOrEqual(hgDBDlf, tgDBDlf, igDBDlfRecLen, tgDBDlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBDlf.iVefCode <> ilDeleteCode(ilVef) Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBDlf)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            Do
                DoEvents
                tgDBDlfSrchKey0.iVefCode = ilDeleteCode(ilVef)
                tgDBDlfSrchKey0.sAirDay = ""
                tgDBDlfSrchKey0.iStartDate(0) = 0
                tgDBDlfSrchKey0.iStartDate(1) = 0
                tgDBDlfSrchKey0.iAirTime(0) = 0
                tgDBDlfSrchKey0.iAirTime(1) = 0
                'Remove Egf
                ilRet = btrGetGreaterOrEqual(hgDBEgf, tgDBDlf, igDBDlfRecLen, tgDBDlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBDlf.iVefCode <> ilDeleteCode(ilVef) Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBEgf)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            Do
                DoEvents
                tgDBVcfSrchKey0.iSellCode = ilDeleteCode(ilVef)
                tgDBVcfSrchKey0.iSellDay = 0
                tgDBVcfSrchKey0.iEffDate(0) = 0
                tgDBVcfSrchKey0.iEffDate(1) = 0
                tgDBVcfSrchKey0.iSellTime(0) = 0
                tgDBVcfSrchKey0.iSellTime(1) = 0
                tgDBVcfSrchKey0.iSellPosNo = 0
                ilRet = btrGetGreaterOrEqual(hgDBVcf, tgDBVcf, igDBVcfRecLen, tgDBVcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBVcf.iSellCode <> ilDeleteCode(ilVef) Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBVcf)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            Do
                DoEvents
                tgDBVlfSrchKey0.iSellCode = ilDeleteCode(ilVef)
                tgDBVlfSrchKey0.iSellDay = 0
                tgDBVlfSrchKey0.iEffDate(0) = 0
                tgDBVlfSrchKey0.iEffDate(1) = 0
                tgDBVlfSrchKey0.iSellTime(0) = 0
                tgDBVlfSrchKey0.iSellTime(1) = 0
                tgDBVlfSrchKey0.iSellPosNo = 0
                tgDBVlfSrchKey0.iSellSeq = 0
                'Remove Vlf
                ilRet = btrGetGreaterOrEqual(hgDBVlf, tgDBVlf, igDBVlfRecLen, tgDBVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBVlf.iSellCode <> ilDeleteCode(ilVef) Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBVlf)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            Do
                DoEvents
                tgDBVlfSrchKey1.iAirCode = ilDeleteCode(ilVef)
                tgDBVlfSrchKey1.iAirDay = 0
                tgDBVlfSrchKey1.iEffDate(0) = 0
                tgDBVlfSrchKey1.iEffDate(1) = 0
                tgDBVlfSrchKey1.iAirTime(0) = 0
                tgDBVlfSrchKey1.iAirTime(1) = 0
                tgDBVlfSrchKey1.iAirPosNo = 0
                tgDBVlfSrchKey1.iAirSeq = 0
                'Remove Vlf
                ilRet = btrGetGreaterOrEqual(hgDBVlf, tgDBVlf, igDBVlfRecLen, tgDBVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBVlf.iAirCode <> ilDeleteCode(ilVef) Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBVlf)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            Do
                DoEvents
                tgDBAlfSrchKey1.iVefCode = ilDeleteCode(ilVef)
                tgDBAlfSrchKey1.iDate(0) = 0
                tgDBAlfSrchKey1.iDate(1) = 0
                ilRet = btrGetGreaterOrEqual(hgDBAlf, tgDBAlf, igDBAlfRecLen, tgDBAlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBAlf.iVefCode <> ilDeleteCode(ilVef) Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBAlf)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            

            'Remove Affiliate: lst, cptt, att, dat
            Do
                DoEvents
                tgDBLstSrchKey2.iLogVefCode = ilDeleteCode(ilVef)
                tgDBLstSrchKey2.iLogDate(0) = 0
                tgDBLstSrchKey2.iLogDate(1) = 0
                ilRet = btrGetGreaterOrEqual(hgDBLst, tgDBLst, igDBLstRecLen, tgDBLstSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBLst.iLogVefCode <> ilDeleteCode(ilVef) Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBLst)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            Do
                DoEvents
                tgDBCpttSrchKey1.iVefCode = ilDeleteCode(ilVef)
                tgDBCpttSrchKey1.iStartDate(0) = 0
                tgDBCpttSrchKey1.iStartDate(1) = 0
                ilRet = btrGetLessOrEqual(hgDBCptt, tgDBCptt, igDBCpttRecLen, tgDBCpttSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBCptt.iVefCode <> ilDeleteCode(ilVef) Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBCptt)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            Do
                DoEvents
                tgDBAttSrchKey1.iCode = ilDeleteCode(ilVef)
                ilRet = btrGetEqual(hgDBAtt, tgDBAtt, igDBAttRecLen, tgDBAttSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                ilRet = mRemoveDat(tgDBAtt.lCode)
                ilRet = mRemoveEpt(tgDBAtt.lCode)
                ilRet = mRemovePet(tgDBAtt.lCode)
                ilRet = mRemoveVat(tgDBAtt.lCode)
                ilRet = btrDelete(hgDBAtt)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
        
            Do
                DoEvents
                tgDBPifSrchKey1.iVefCode = ilDeleteCode(ilVef)
                tgDBPifSrchKey1.iStartDate(0) = 0
                tgDBPifSrchKey1.iStartDate(1) = 0
                tgDBPifSrchKey1.iSeqNo = 0
                ilRet = btrGetGreaterOrEqual(hgDBPif, tgDBPif, igDBPifRecLen, tgDBPifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBPif.iVefCode <> ilDeleteCode(ilVef) Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBPif)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            Do
                DoEvents
                tgDBVbfSrchKey1.iVefCode = ilDeleteCode(ilVef)
                tgDBVbfSrchKey1.iStartDate(0) = 0
                tgDBVbfSrchKey1.iStartDate(1) = 0
                ilRet = btrGetGreaterOrEqual(hgDBVbf, tgDBVbf, igDBVbfRecLen, tgDBVbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBVbf.iVefCode <> ilDeleteCode(ilVef) Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBVbf)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            Do
                DoEvents
                tgDBVafSrchKey1.iVefCode = ilDeleteCode(ilVef)
                ilRet = btrGetEqual(hgDBVaf, tgDBVaf, igDBVafRecLen, tgDBVafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBVaf)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            Do
                DoEvents
                tgDBVofSrchKey0.iVefCode = ilDeleteCode(ilVef)
                tgDBVofSrchKey0.sType = ""
                ilRet = btrGetGreaterOrEqual(hgDBVof, tgDBVof, igDBVofRecLen, tgDBVofSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If tgDBVof.iVefCode <> ilDeleteCode(ilVef) Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBVof)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            Do
                DoEvents
                tgDBVpfSrchKey0.iCode = ilDeleteCode(ilVef)
                ilRet = btrGetEqual(hgDBVpf, tgDBVpf, igDBVpfRecLen, tgDBVpfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                ilRet = btrDelete(hgDBVpf)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        Exit Do
                    End If
                End If
            Loop
            If tgDBVef.lPvfCode > 0 Then
                tgDBPvfSrchKey0.lCode = tgDBVef.lPvfCode
                ilRet = btrGetEqual(hgDBPvf, tgDBPvf, igDBPvfRecLen, tgDBPvfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    llLkPvfCode = tgDBPvf.lLkPvfCode
                    ilRet = btrDelete(hgDBPvf)
                    Do While llLkPvfCode > 0
                        DoEvents
                        tgDBPvfSrchKey0.lCode = llLkPvfCode
                        ilRet = btrGetEqual(hgDBPvf, tgDBPvf, igDBPvfRecLen, tgDBPvfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            Exit Do
                        End If
                        llLkPvfCode = tgDBPvf.lLkPvfCode
                        ilRet = btrDelete(hgDBPvf)
                    Loop
                End If
            End If
            Do
                tgDBVffSrchKey1.iCode = ilDeleteCode(ilVef)
                ilRet = btrGetEqual(hgDBVff, tgDBVff, igDBVffRecLen, tgDBVffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hgDBVff)
                Else
                    Exit Do
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            
            ilRet = btrDelete(hgDBVef)
            'If imUserInput Then
                gLogMsg "Vehicle " & tgDBVef.sName & " Type " & tgDBVef.sType & " vefCode " & tgDBVef.iCode, "csiArchiveRemovalDetail.txt", False
            'End If
        End If
    Next ilVef
    
    'Remove library: lcf, ltf, lvf, lef, ghf, gsf
    'Remove SSF
    'Remove links: vlf, vcf, dlf, egf
    'Remove Urf
    'Remove Pvf
    'Remove Affiliate: lst, cptt, att, dat
    'Remove vehicles: vof, vpf, vbf, pif, vaf, saf
    mRemoveVef = True
End Function

Private Function mRemoveUnusedDnf(llRetainDate As Long) As Integer
'Private Function mRemoveUnusedDnf() As Integer
    Dim ilRet As Integer
    Dim ilDnf As Integer
    ''Dim ilBookYear As Integer
    ''Dim slDate As String
    ''Dim ilNowYear As Integer
    '6/15/18: Passed as parameter
    'Dim llRetainDate As Long
    Dim llDate As Long
    
    lacMsg.Caption = "Checking for Unused Research Books"
    ilRet = btrGetLast(hgDBDnf, tgDBDnf, igDBDnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        ReDim slDnfUsed(0 To tgDBDnf.iCode + 1) As String * 1
    Else
        mRemoveUnusedDnf = True
        Exit Function
    End If
    'Loop on CLF to catch all the books currently referenced including ones with adate prior to retain date
    ilRet = btrGetFirst(hgDBClf, tgDBClf, igDBClfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        If (tgDBClf.iDnfCode <= UBound(slDnfUsed)) And (tgDBClf.iDnfCode > 0) Then
            slDnfUsed(tgDBClf.iDnfCode) = "Y"
        End If
        ilRet = btrGetNext(hgDBClf, tgDBClf, igDBClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    'Retain books not referenced but with a date after the retain date.  This is required so that if the
    'book references at population source that would be removed except that this book references it
    ilRet = btrGetFirst(hgDBDnf, tgDBDnf, igDBDnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        If imCancelled Then
            mRemoveUnusedDnf = False
            Exit Function
        End If
        If (tgDBDnf.iCode <= UBound(slDnfUsed)) And (tgDBDnf.iCode > 0) Then
            gUnpackDateLong tgDBDnf.iBookDate(0), tgDBDnf.iBookDate(1), llDate
            If llDate >= llRetainDate Then
                slDnfUsed(tgDBDnf.iCode) = "Y"
            End If
        End If
        ilRet = btrGetNext(hgDBDnf, tgDBDnf, igDBDnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    
    '5/22/18: Check if book referenced as population and if so retain if that book is retained
    ilRet = btrGetFirst(hgDBDnf, tgDBDnf, igDBDnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        If imCancelled Then
            mRemoveUnusedDnf = False
            Exit Function
        End If
        If slDnfUsed(tgDBDnf.iCode) = "Y" Then
            If (tgDBDnf.iPopDnfCode <= UBound(slDnfUsed)) And (tgDBDnf.iPopDnfCode > 0) Then
                slDnfUsed(tgDBDnf.iPopDnfCode) = "Y"
            End If
            If (tgDBDnf.iQualPopDnfCode <= UBound(slDnfUsed)) And (tgDBDnf.iQualPopDnfCode > 0) Then
                slDnfUsed(tgDBDnf.iQualPopDnfCode) = "Y"
            End If
        End If
        ilRet = btrGetNext(hgDBDnf, tgDBDnf, igDBDnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    
    ''5/22/18: Use same date rule as used with advertiser and agency
    ''ilNowYear = Year(gAdjYear(Format$(gNow(), "m/d/yy")))
    'llRetainDate = gDateValue(gObtainEndStd(DateAdd("d", -31 * 2 - 7, Format$(Now, "m/d/yy")))) + 1

    lacMsg.Caption = "Removing Unused Research Books"
    ReDim ilDeleteCode(0 To 0) As Integer
    ilRet = btrGetFirst(hgDBDnf, tgDBDnf, igDBDnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        If imCancelled Then
            mRemoveUnusedDnf = False
            Exit Function
        End If
        If (tgDBDnf.iCode <= UBound(slDnfUsed)) And (tgDBDnf.iCode > 0) Then
            If slDnfUsed(tgDBDnf.iCode) <> "Y" Then
                '5/22/18: Change date test to be the same as Advertiser and agency
                'gUnpackDate tgDBDnf.iBookDate(0), tgDBDnf.iBookDate(1), slDate
                'ilBookYear = Year(gAdjYear(slDate))
                'If ilBookYear < ilNowYear Then
                gUnpackDateLong tgDBDnf.iBookDate(0), tgDBDnf.iBookDate(1), llDate
                If llDate < llRetainDate Then
                    ilDeleteCode(UBound(ilDeleteCode)) = tgDBDnf.iCode
                    ReDim Preserve ilDeleteCode(0 To UBound(ilDeleteCode) + 1) As Integer
                End If
            End If
        End If
        ilRet = btrGetNext(hgDBDnf, tgDBDnf, igDBDnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilRet = mClearDnfRef(ilDeleteCode())
    For ilDnf = 0 To UBound(ilDeleteCode) - 1 Step 1
        DoEvents
        If imCancelled Then
            mRemoveUnusedDnf = False
            Exit Function
        End If
        ilRet = mRemoveDnf(ilDeleteCode(ilDnf))
    Next ilDnf
    If UBound(ilDeleteCode) > 0 Then
        gLogMsg "Total of " & Trim$(Str$(UBound(ilDeleteCode))) & " Research Books Removed", "csiArchive.Txt", False
    End If
    Erase ilDeleteCode
    Erase slDnfUsed
    mRemoveUnusedDnf = True
End Function

Private Function mRemoveDnf(ilDnfCode As Integer) As Integer
    Dim ilRet As Integer
    Dim slBookDate As String

    Do
        DoEvents
        If imCancelled Then
            mRemoveDnf = False
            Exit Function
        End If
        tgDBDrfSrchKey0.iDnfCode = ilDnfCode
        tgDBDrfSrchKey0.iMnfSocEco = 0
        tgDBDrfSrchKey0.iRdfCode = 0
        tgDBDrfSrchKey0.iVefCode = 0
        tgDBDrfSrchKey0.sDemoDataType = ""
        tgDBDrfSrchKey0.sInfoType = ""
        ilRet = btrGetGreaterOrEqual(hgDBDrf, tgDBDrf, igDBDrfRecLen, tgDBDrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get lSdf current record to obtain date
        If ilRet = BTRV_ERR_NONE Then
            If tgDBDrf.iDnfCode <> ilDnfCode Then
                Exit Do
            End If
            ilRet = btrDelete(hgDBDrf)
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet <> BTRV_ERR_CONFLICT Then
                    mBtrErrorMsg "mRemoveDNF: Delete DRF error ", ilRet
                    mRemoveDnf = False
                    Exit Function
                End If
            End If
        Else
            Exit Do
        End If
    Loop
    Do
        DoEvents
        If imCancelled Then
            mRemoveDnf = False
            Exit Function
        End If
        tgDBDefSrchKey1.iDnfCode = ilDnfCode
        tgDBDefSrchKey1.iStartDate(0) = 0
        tgDBDefSrchKey1.iStartDate(1) = 0
        ilRet = btrGetGreaterOrEqual(hgDBDef, tgDBDef, igDBDefRecLen, tgDBDefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get lSdf current record to obtain date
        If ilRet = BTRV_ERR_NONE Then
            If tgDBDef.iDnfCode <> ilDnfCode Then
                Exit Do
            End If
            ilRet = btrDelete(hgDBDef)
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet <> BTRV_ERR_CONFLICT Then
                    mBtrErrorMsg "mRemoveDNF: Delete DEF error ", ilRet
                    mRemoveDnf = False
                    Exit Function
                End If
            End If
        Else
            Exit Do
        End If
    Loop
    Do
        DoEvents
        If imCancelled Then
            mRemoveDnf = False
            Exit Function
        End If
        tgDBDpfSrchKey2.iDnfCode = ilDnfCode
        tgDBDpfSrchKey2.iMnfDemo = 0
        ilRet = btrGetGreaterOrEqual(hgDBDpf, tgDBDpf, igDBDpfRecLen, tgDBDpfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get lSdf current record to obtain date
        If ilRet = BTRV_ERR_NONE Then
            If tgDBDpf.iDnfCode <> ilDnfCode Then
                Exit Do
            End If
            ilRet = btrDelete(hgDBDpf)
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet <> BTRV_ERR_CONFLICT Then
                    mBtrErrorMsg "mRemoveDNF: Delete DPF error ", ilRet
                    mRemoveDnf = False
                    Exit Function
                End If
            End If
        Else
            Exit Do
        End If
    Loop
    Do
        DoEvents
        If imCancelled Then
            mRemoveDnf = False
            Exit Function
        End If
        tgDBDnfSrchKey0.iCode = ilDnfCode
        ilRet = btrGetEqual(hgDBDnf, tgDBDnf, igDBDnfRecLen, tgDBDnfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        ilRet = btrDelete(hgDBDnf)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet <> BTRV_ERR_CONFLICT Then
                mBtrErrorMsg "mRemoveDnf: Delete error ", ilRet
                mRemoveDnf = False
                Exit Function
            End If
        End If
        'If imUserInput Then
            gUnpackDate tgDBDnf.iBookDate(0), tgDBDnf.iBookDate(1), slBookDate
            gLogMsg "Book " & tgDBDnf.sBookName & " Book Date " & slBookDate & " dnfCode " & tgDBDnf.iCode, "csiArchiveRemovalDetail.txt", False
        'End If
    Loop
    mRemoveDnf = True
End Function

Private Sub mCleanupUnused(llEarliestTrafSpot As Long, llAffEarliestAstDate As Long)
    Dim ilRet As Integer
    Dim llEarliestCntrDate As Long
    Dim llEarliestPayRevHistDate As Long
    Dim llTotalCntrCount As Long
    Dim llTotalBsfCount As Long
    Dim llTotalBvfCount As Long
    Dim llTotalScfCount As Long
    Dim llTotalPayRevHistCount As Long
    
    
'lacCntrStatus(1).Caption = "Bypassed"
''lacCntrStatus(1).Visible = True
'lacPayRevHistStatus(1).Caption = "Bypassed"
''lacPayRevHistStatus(1).Visible = True
'Exit Sub
    
    'If imUserInput Then
        gLogMsg "Archive Removal Details", "csiArchiveRemovalDetail.txt", True
    'End If
    llTotalCntrCount = 0
    llTotalBsfCount = 0
    llTotalBvfCount = 0
    llTotalScfCount = 0
    llTotalPayRevHistCount = 0
    
    'contract
    lacMsg.Caption = "Removing Expired Contracts"
    DoEvents
    llEarliestCntrDate = 0
    If smCntrEndDate <> "" Then
        If gValidDate(smCntrEndDate) Then
            llEarliestCntrDate = llEarliestTrafSpot
            If tmSpf.sGUseAffSys = "Y" Then
                If llAffEarliestAstDate < llEarliestTrafSpot Then
                    llEarliestCntrDate = llAffEarliestAstDate
                End If
            End If
            If gDateValue(smCntrEndDate) < llEarliestCntrDate Then
                llEarliestCntrDate = gDateValue(smCntrEndDate) + 1
            End If
            ilRet = mRemoveContracts(llEarliestCntrDate, llTotalCntrCount)
            If llEarliestCntrDate > 0 Then
                gLogMsg "Total of " & Trim$(Str$(llTotalCntrCount)) & " Expired Contracts Removed through " & Format$(llEarliestCntrDate - 1, "m/d/yy"), "csiArchive.Txt", False
            Else
                gLogMsg "Total of " & Trim$(Str$(llTotalCntrCount)) & " Expired Contracts Removed", "csiArchive.Txt", False
            End If
            If (imCancelled) Then
                Exit Sub
            End If
            lacMsg.Caption = "Removing Budgets"
            DoEvents
            ilRet = mRemoveBudgets(llEarliestCntrDate, llTotalBsfCount, llTotalBvfCount)
            If llEarliestCntrDate > 0 Then
                gLogMsg "Total of " & Trim$(Str$(llTotalBsfCount)) & " Budgets by Salesperson Removed through " & Format$(llEarliestCntrDate - 1, "m/d/yy"), "csiArchive.Txt", False
                gLogMsg "Total of " & Trim$(Str$(llTotalBvfCount)) & " Budgets by Vehicle Removed through " & Format$(llEarliestCntrDate - 1, "m/d/yy"), "csiArchive.Txt", False
            Else
                gLogMsg "Total of " & Trim$(Str$(llTotalBsfCount)) & " Budgets by Salesperson Removed", "csiArchive.Txt", False
                gLogMsg "Total of " & Trim$(Str$(llTotalBvfCount)) & " Budgets by Vehicle Removed", "csiArchive.Txt", False
            End If
            If (imCancelled) Then
                Exit Sub
            End If
            ilRet = mRemoveSCF(llEarliestCntrDate, llTotalScfCount)
            If llEarliestCntrDate > 0 Then
                gLogMsg "Total of " & Trim$(Str$(llTotalScfCount)) & " Salesperson Commission Specifications Removed through " & Format$(llEarliestCntrDate - 1, "m/d/yy"), "csiArchive.Txt", False
            Else
                gLogMsg "Total of " & Trim$(Str$(llTotalScfCount)) & " Salesperson Commission Specifications Removed", "csiArchive.Txt", False
            End If
            'Remove MG tracking records
            ilRet = mRemoveMtf()
        End If
    Else
        lacCntrStatus(1).Caption = "Bypassed"
        lacCntrStatus(1).Visible = True
    End If
    If (imCancelled) Then
        Exit Sub
    End If
    lacMsg.Caption = "Removing Payment and Revenue History"
    DoEvents
    If smPayRevHistEndDate <> "" Then
        If gValidDate(smPayRevHistEndDate) Then
            llEarliestPayRevHistDate = llEarliestTrafSpot
            If tmSpf.sGUseAffSys = "Y" Then
                If llAffEarliestAstDate < llEarliestTrafSpot Then
                    llEarliestPayRevHistDate = llAffEarliestAstDate
                End If
            End If
            If gDateValue(smPayRevHistEndDate) < llEarliestPayRevHistDate Then
                llEarliestPayRevHistDate = gDateValue(smPayRevHistEndDate) + 1
            End If
            ilRet = mRemovePayRevHist(llEarliestPayRevHistDate, llTotalPayRevHistCount)
            If llEarliestPayRevHistDate > 0 Then
                gLogMsg "Total of " & Trim$(Str$(llTotalPayRevHistCount)) & " Payment and Revenue History Removed through " & Format$(llEarliestPayRevHistDate - 1, "m/d/yy"), "csiArchive.Txt", False
            Else
                gLogMsg "Total of " & Trim$(Str$(llTotalPayRevHistCount)) & " Payment and Revenue History Removed", "csiArchive.Txt", False
            End If
        End If
    Else
        lacPayRevHistStatus(1).Caption = "Bypassed"
        lacPayRevHistStatus(1).Visible = True
    End If
    If (imCancelled) Then
        Exit Sub
    End If
    If smCntrEndDate <> "" Then
        '6/15/18: Added using traffic spot instead of 2 month rule
        ''Remove old books
        'ilRet = mRemoveUnusedDnf()
        If llEarliestCntrDate > 0 Then
            ilRet = mRemoveUnusedDnf(llEarliestCntrDate)
        End If
        'If imCancelled Then
        '    Exit Sub
        'End If
        'Remove unused Vehicles; Salesperson and Rate Cards
        ilRet = mRemoveUnusedVefRcfSlf()
    End If
End Sub

Private Function mCheckUnusedDates() As Integer
    Dim slDate As String
    Dim slStr As String
    
    slStr = Trim$(edcCntrEnd.Text)
    If StrComp(slStr, "Bypass", vbTextCompare) = 0 Then
        slStr = ""
    End If
    If (slStr <> "") Then
        slDate = Trim$(edcCntrEnd.Text)
        If gDateValue(slDate) > gDateValue(smCntrEndDate) Then
            If imUserInput Then
                MsgBox "Expired Contract End Date can't be advanced"
                edcCntrEnd.Text = smCntrEndDate
            Else
                gLogMsg "Archive not run because Expired Contract End Date advanced", "csiArchive.txt", False
            End If
            mCheckUnusedDates = False
            Exit Function
        End If
        If gDateValue(gObtainEndStd(slDate)) <> gDateValue(slDate) Then
            If imUserInput Then
                MsgBox "Expired Contract Date must be the End of a Standard Broadcast Month"
            Else
                gLogMsg "Archive not run because Expired Contract Date is not the End Of a Standard Broadcast Month", "csiArchive.txt", False
            End If
            mCheckUnusedDates = False
            Exit Function
        End If
        If (gDateValue(slDate) > gDateValue(smTrafEndDate)) And (imTrafStatus <= 0) Then
            If imUserInput Then
                MsgBox "Expired Contract Date can't be newer than Traffic End Date"
                edcTrafficEnd.Text = smTrafEndDate
            Else
                gLogMsg "Archive not run because EXpired Contract Date newer than Traffic End Date", "csiArchive.txt", False
            End If
            mCheckUnusedDates = False
            Exit Function
        End If
        smCntrEndDate = Trim$(edcCntrEnd.Text)
    Else
        smCntrEndDate = ""
    End If
    slStr = Trim$(edcPayRevHistEnd.Text)
    If StrComp(slStr, "Bypass", vbTextCompare) = 0 Then
        slStr = ""
    End If
    If slStr <> "" Then
        slDate = Trim$(edcPayRevHistEnd.Text)
        If gDateValue(slDate) > gDateValue(smPayRevHistEndDate) Then
            If imUserInput Then
                MsgBox "Payment and Revenue Date can't be advanced"
                edcPayRevHistEnd.Text = smPayRevHistEndDate
            Else
                gLogMsg "Archive not run because Payment and Revenue Date advanced", "csiArchive.txt", False
            End If
            mCheckUnusedDates = False
            Exit Function
        End If
        If gDateValue(gObtainEndStd(slDate)) <> gDateValue(slDate) Then
            If imUserInput Then
                MsgBox "Payment and Revenue History Date must be the End of a Standard Broadcast Month"
            Else
                gLogMsg "Archive not run because Payment and Revenue History Date is not the End Of a Standard Broadcast Month", "csiArchive.txt", False
            End If
            mCheckUnusedDates = False
            Exit Function
        End If
        If (gDateValue(slDate) > gDateValue(smTrafEndDate)) And (imTrafStatus <= 0) Then
            If imUserInput Then
                MsgBox "Payment and Revenue History Date can't be newer than Traffic End Date"
                edcTrafficEnd.Text = smTrafEndDate
            Else
                gLogMsg "Archive not run because Payment and Revenue History Date newer than Traffic End Date", "csiArchive.txt", False
            End If
            mCheckUnusedDates = False
            Exit Function
        End If
        smPayRevHistEndDate = Trim$(edcPayRevHistEnd.Text)
    Else
        smPayRevHistEndDate = ""
    End If
    mCheckUnusedDates = True
End Function

Private Function mClearSlfRef(ilDeleteCode() As Integer) As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim ilAdf As Integer
    Dim ilAgf As Integer
    Dim ilSlf As Integer
    
    ReDim ilAdfCode(0 To 0) As Integer
    ilRet = btrGetFirst(hgDBAdf, tgDBAdf, igDBAdfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        ilFound = False
        For ilSlf = 0 To UBound(ilDeleteCode) - 1 Step 1
            DoEvents
            If tgDBAdf.iSlfCode = ilDeleteCode(ilSlf) Then
                ilFound = True
                Exit For
            End If
        Next ilSlf
        If ilFound Then
            ilAdfCode(UBound(ilAdfCode)) = tgDBAdf.iCode
            ReDim Preserve ilAdfCode(0 To UBound(ilAdfCode) + 1) As Integer
        End If
        ilRet = btrGetNext(hgDBAdf, tgDBAdf, igDBAdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilAdf = 0
    Do While ilAdf < UBound(ilAdfCode)
        DoEvents
        tgDBAdfSrchKey.iCode = ilAdfCode(ilAdf)
        ilRet = btrGetEqual(hgDBAdf, tgDBAdf, igDBAdfRecLen, tgDBAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            tgDBAdf.iSlfCode = 0
            ilRet = btrUpdate(hgDBAdf, tgDBAdf, igDBAdfRecLen)
            If ilRet = BTRV_ERR_NONE Then
                ilAdf = ilAdf + 1
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    ilAdf = ilAdf + 1
                End If
            End If
        Else
            ilAdf = ilAdf + 1
        End If
    Loop
    Erase ilAdfCode

    ReDim ilAgfCode(0 To 0) As Integer
    ilRet = btrGetFirst(hgDBAgf, tgDBAgf, igDBAgfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        ilFound = False
        For ilSlf = 0 To UBound(ilDeleteCode) - 1 Step 1
            DoEvents
            If tgDBAgf.iSlfCode = ilDeleteCode(ilSlf) Then
                ilFound = True
                Exit For
            End If
        Next ilSlf
        If ilFound Then
            ilAgfCode(UBound(ilAgfCode)) = tgDBAgf.iCode
            ReDim Preserve ilAgfCode(0 To UBound(ilAgfCode) + 1) As Integer
        End If
        ilRet = btrGetNext(hgDBAgf, tgDBAgf, igDBAgfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilAgf = 0
    Do While ilAgf < UBound(ilAgfCode)
        DoEvents
        tgDBAgfSrchKey.iCode = ilAgfCode(ilAgf)
        ilRet = btrGetEqual(hgDBAgf, tgDBAgf, igDBAgfRecLen, tgDBAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            tgDBAgf.iSlfCode = 0
            ilRet = btrUpdate(hgDBAgf, tgDBAgf, igDBAgfRecLen)
            If ilRet = BTRV_ERR_NONE Then
                ilAgf = ilAgf + 1
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    ilAgf = ilAgf + 1
                End If
            End If
        Else
            ilAgf = ilAgf + 1
        End If
    Loop
    Erase ilAgfCode
    
    mClearSlfRef = True
End Function

Private Function mClearDnfRef(ilDeleteCode() As Integer) As Integer

    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim ilVef As Integer
    Dim ilAgf As Integer
    
    ReDim ilVefCode(0 To 0) As Integer
    ilRet = btrGetFirst(hgDBVef, tgDBVef, igDBVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        ilFound = False
        For ilVef = 0 To UBound(ilDeleteCode) - 1 Step 1
            DoEvents
            If tgDBVef.iDnfCode = ilDeleteCode(ilVef) Then
                ilFound = True
                Exit For
            End If
        Next ilVef
        If ilFound Then
            ilVefCode(UBound(ilVefCode)) = tgDBVef.iCode
            ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
        End If
        ilRet = btrGetNext(hgDBVef, tgDBVef, igDBVefRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilVef = 0
    Do While ilVef < UBound(ilVefCode)
        DoEvents
        tgDBVefSrchKey0.iCode = ilVefCode(ilVef)
        ilRet = btrGetEqual(hgDBVef, tgDBVef, igDBVefRecLen, tgDBVefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            tgDBVef.iDnfCode = 0
            ilRet = btrUpdate(hgDBVef, tgDBVef, igDBVefRecLen)
            If ilRet = BTRV_ERR_NONE Then
                ilVef = ilVef + 1
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    ilVef = ilVef + 1
                End If
            End If
        Else
            ilVef = ilVef + 1
        End If
    Loop
    Erase ilVefCode
    mClearDnfRef = True
End Function

Private Function mRemoveMtf() As Integer
    Dim llMtf As Long
    Dim ilRet As Integer
    
    ReDim llMtfCode(0 To 0) As Long
    ilRet = btrGetFirst(hgDBMtf, tgDBMtf, igDBMtfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        tgDBChfSrchKey0.lCode = tgDBSdf.lChfCode
        ilRet = btrGetEqual(hgDBChf, tgDBChf, igDBChfRecLen, tgDBChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            llMtfCode(UBound(llMtfCode)) = tgDBMtf.lCode
            ReDim Preserve llMtfCode(0 To UBound(llMtfCode) + 1) As Long
        End If
        ilRet = btrGetNext(hgDBMtf, tgDBMtf, igDBMtfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    llMtf = 0
    Do While llMtf < UBound(llMtfCode)
        tgDBMtfSrchKey.lCode = llMtfCode(llMtf)
        ilRet = btrGetEqual(hgDBMtf, tgDBMtf, igDBMtfRecLen, tgDBMtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hgDBMtf)
            If ilRet = BTRV_ERR_NONE Then
                llMtf = llMtf + 1
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    llMtf = llMtf + 1
                End If
            End If
        Else
            llMtf = llMtf + 1
        End If
    Loop
    Erase llMtfCode
    mRemoveMtf = True
End Function


Private Function mRemoveWebl(llRetainDate As Long, llWeblCount As Long) As Integer
'    Dim llDate As Long
'    Dim llWebl As Long
'    Dim ilRet As Integer
'
'    If llRetainDate <= 0 Then
'        mRemoveWebl = True
'        Exit Function
'    End If
'
'    ReDim llWeblCode(0 To 0) As Long
'    ilRet = btrGetFirst(hgDBWebl, tgDBWebl, igDBWeblRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'    Do While ilRet = BTRV_ERR_NONE
'        gUnpackDateLong tgDBWebl.iPostDay(0), tgDBWebl.iPostDay(1), llDate
'        If (llDate < llRetainDate) And (llDate > 0) Then
'            llWeblCode(UBound(llWeblCode)) = tgDBWebl.lCode
'            ReDim Preserve llWeblCode(0 To UBound(llWeblCode) + 1) As Long
'        End If
'        ilRet = btrGetNext(hgDBWebl, tgDBWebl, igDBWeblRecLen, BTRV_LOCK_NONE, SETFORWRITE)
'    Loop
'    llWebl = 0
'    Do While llWebl < UBound(llWeblCode)
'        tgDBWeblSrchKey0.lCode = llWeblCode(llWebl)
'        ilRet = btrGetEqual(hgDBWebl, tgDBWebl, igDBWeblRecLen, tgDBWeblSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet = BTRV_ERR_NONE Then
'            ilRet = btrDelete(hgDBWebl)
'            If ilRet = BTRV_ERR_NONE Then
'                llWeblCount = llWeblCount + 1
'                llWebl = llWebl + 1
'            Else
'                If ilRet <> BTRV_ERR_CONFLICT Then
'                    llWebl = llWebl + 1
'                End If
'            End If
'        Else
'            llWebl = llWebl + 1
'        End If
'    Loop
'    Erase llWeblCode
'    mRemoveWebl = True
    Dim llDate As Long
    Dim llWebl As Long
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim llLastCode As Long

    
    If llRetainDate <= 0 Then
        mRemoveWebl = True
        Exit Function
    End If

    ReDim llWeblCode(0 To 1000) As Long
    ilUpper = 0
    llLastCode = 0
    Do
        DoEvents
        tgDBWeblSrchKey0.lCode = llLastCode
        ilRet = btrGetGreaterOrEqual(hgDBWebl, tgDBWebl, igDBWeblRecLen, tgDBWeblSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get lSdf current record to obtain date
        Do While ilRet = BTRV_ERR_NONE
            llLastCode = tgDBWebl.lCode
            gUnpackDateLong tgDBWebl.iPostDay(0), tgDBWebl.iPostDay(1), llDate
            If (llDate < llRetainDate) And (llDate > 0) Then
                llWeblCode(ilUpper) = tgDBWebl.lCode
                ilUpper = ilUpper + 1
                If ilUpper >= UBound(llWeblCode) Then
                    Exit Do
                End If
                'ReDim Preserve llWeblCode(0 To UBound(llWeblCode) + 1) As Long
            End If
            ilRet = btrGetNext(hgDBWebl, tgDBWebl, igDBWeblRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
        DoEvents
        If ilUpper = 0 Then
            Exit Do
        End If
        llWebl = 0
        Do While llWebl < ilUpper   'UBound(llWeblCode)
            tgDBWeblSrchKey0.lCode = llWeblCode(llWebl)
            ilRet = btrGetEqual(hgDBWebl, tgDBWebl, igDBWeblRecLen, tgDBWeblSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hgDBWebl)
                If ilRet = BTRV_ERR_NONE Then
                    llWeblCount = llWeblCount + 1
                    llWebl = llWebl + 1
                Else
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        llWebl = llWebl + 1
                    End If
                End If
            Else
                llWebl = llWebl + 1
            End If
        Loop
        ilUpper = 0
    Loop
    Erase llWeblCode
    mRemoveWebl = True
End Function

Private Function mRemoveEsf(llRetainDate As Long, llEsfCount As Long, llOutDate As Long) As Integer
    Dim llDate As Long
    Dim llEsf As Long
    Dim ilRet As Integer
    'Dim slDate As String
    'Dim llRetainDate As Long
    Dim ilUpper As Integer
    Dim llLastCode As Long
    
    '5/22/18: Changed to use the affiliate retain date
    'slDate = Format(Now, "m/d/yy")
    'slDate = gObtainStartStd(slDate)
    'slDate = DateAdd("d", -32, slDate)
    'slDate = gObtainStartStd(slDate)
    'llRetainDate = gDateValue(slDate)
    llOutDate = llRetainDate - 1
    If llRetainDate <= 0 Then
        mRemoveEsf = True
        Exit Function
    End If

    ReDim llEsfCode(0 To 1000) As Long
    ilUpper = 0
    llLastCode = 0
    Do
        DoEvents
        'ilRet = btrGetFirst(hgDBEsf, tgDBEsf, igDBEsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        tgDBEsfSrchKey0.lCode = llLastCode
        ilRet = btrGetGreaterOrEqual(hgDBEsf, tgDBEsf, igDBEsfRecLen, tgDBEsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get lSdf current record to obtain date
        Do While ilRet = BTRV_ERR_NONE
            llLastCode = tgDBEsf.lCode
            gUnpackDateLong tgDBEsf.iExpDate(0), tgDBEsf.iExpDate(1), llDate
            llDate = llDate + tgDBEsf.iNumDays - 1
            If (llDate < llRetainDate) And (llDate > 0) Then
                'llEsfCode(UBound(llEsfCode)) = tgDBEsf.lCode
                'ReDim Preserve llEsfCode(0 To UBound(llEsfCode) + 1) As Long
                llEsfCode(ilUpper) = tgDBEsf.lCode
                ilUpper = ilUpper + 1
                If ilUpper >= UBound(llEsfCode) Then
                    Exit Do
                End If
            End If
            ilRet = btrGetNext(hgDBEsf, tgDBEsf, igDBEsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
        DoEvents
        If ilUpper = 0 Then
            Exit Do
        End If
        llEsf = 0
        Do While llEsf < ilUpper    'UBound(llEsfCode)
            tgDBEdfSrchKey1.lCode = llEsfCode(llEsf)
            ilRet = btrGetEqual(hgDBEdf, tgDBEdf, igDBEdfRecLen, tgDBEdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hgDBEdf)
                'If ilRet = BTRV_ERR_NONE Then
                '    llEsf = llEsf + 1
                'Else
                '    If ilRet <> BTRV_ERR_CONFLICT Then
                '        llEsf = llEsf + 1
                '    End If
                'End If
            Else
                llEsf = llEsf + 1
            End If
        Loop
        llEsf = 0
        Do While llEsf < ilUpper    'UBound(llEsfCode)
            tgDBEsfSrchKey0.lCode = llEsfCode(llEsf)
            ilRet = btrGetEqual(hgDBEsf, tgDBEsf, igDBEsfRecLen, tgDBEsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hgDBEsf)
                If ilRet = BTRV_ERR_NONE Then
                    llEsfCount = llEsfCount + 1
                    llEsf = llEsf + 1
                Else
                    If ilRet <> BTRV_ERR_CONFLICT Then
                        llEsf = llEsf + 1
                    End If
                End If
            Else
                llEsf = llEsf + 1
            End If
        Loop
        ilUpper = 0
    Loop
    Erase llEsfCode
    mRemoveEsf = True
End Function

Private Function mRemoveRqf(llRqfCount As Long, llOutDate As Long) As Integer
    Dim llDate As Long
    Dim llRqf As Long
    Dim ilRet As Integer
    Dim slDate As String
    Dim llRetainDate As Long
    
    slDate = Format(Now, "m/d/yy")
    llRetainDate = gDateValue(slDate) - 1
    llOutDate = llRetainDate - 1
    If llRetainDate <= 0 Then
        mRemoveRqf = True
        Exit Function
    End If

    ReDim llRqfCode(0 To 0) As Long
    ilRet = btrGetFirst(hgDBRqf, tgDBRqf, igDBRqfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        llDate = tgDBRqf.lEnteredDate
        If (llDate < llRetainDate) And (llDate > 0) Then
            llRqfCode(UBound(llRqfCode)) = tgDBRqf.lCode
            ReDim Preserve llRqfCode(0 To UBound(llRqfCode) + 1) As Long
        End If
        ilRet = btrGetNext(hgDBRqf, tgDBRqf, igDBRqfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    llRqf = 0
    Do While llRqf < UBound(llRqfCode)
        tgDBRffSrchKey1.lRqfCode = llRqfCode(llRqf)
        tgDBRffSrchKey1.sType = ""
        tgDBRffSrchKey1.iSequenceNumber = 0
        ilRet = btrGetGreaterOrEqual(hgDBRff, tgDBRff, igDBRffRecLen, tgDBRffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If (ilRet = BTRV_ERR_NONE) And (tgDBRff.lRqfCode = llRqfCode(llRqf)) Then
            ilRet = btrDelete(hgDBRff)
            'If ilRet = BTRV_ERR_NONE Then
            '    llRqf = llRqf + 1
            'Else
            '    If ilRet <> BTRV_ERR_CONFLICT Then
            '        llRqf = llRqf + 1
            '    End If
            'End If
        Else
            llRqf = llRqf + 1
        End If
    Loop
    llRqf = 0
    Do While llRqf < UBound(llRqfCode)
        tgDBRqfSrchKey0.lCode = llRqfCode(llRqf)
        ilRet = btrGetEqual(hgDBRqf, tgDBRqf, igDBRqfRecLen, tgDBRqfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hgDBRqf)
            If ilRet = BTRV_ERR_NONE Then
                llRqfCount = llRqfCount + 1
                llRqf = llRqf + 1
            Else
                If ilRet <> BTRV_ERR_CONFLICT Then
                    llRqf = llRqf + 1
                End If
            End If
        Else
            llRqf = llRqf + 1
        End If
    Loop
    Erase llRqfCode
    mRemoveRqf = True
End Function

Private Sub mTeamPop()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    ilRet = gObtainMnfForType("Z", smTeamTag, tmTeam())
    Exit Sub
End Sub

Private Sub mCheckGG()
    Dim ilRet As Integer
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim llNow As Long
    Dim llDate As Long
    Dim slStr As String
    Dim ilLoop As Integer
    Dim hlSaf As Integer
    
    On Error Resume Next
    
    'If imLastHourGGChecked = Hour(Now) Then
    '    Exit Sub
    'End If
    'imLastHourGGChecked = Hour(Now)
    igGGFlag = 1
    hlSaf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSaf
        Exit Sub
    End If
    
    imSafRecLen = Len(tmSaf)
    ilRet = btrGetFirst(hlSaf, tmSaf, imSafRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSaf
        Exit Sub
    End If
    
    ilField1 = Asc(tmSaf.sName)
    slStr = Mid$(tmSaf.sName, 2, 5)
    llDate = Val(slStr)
    ilField2 = Asc(Mid$(tmSaf.sName, 11, 1))
    llNow = gDateValue(Format$(Now, "m/d/yy"))
    If (ilField1 = 0) And (ilField2 = 1) Then
        If llDate <= llNow Then
            ilField2 = 0
        End If
    End If
    If (ilField1 = 0) And (ilField2 = 0) Then
        igGGFlag = 0
    End If
    'gSetRptGGFlag tmSaf
    btrDestroy hlSaf
End Sub

Function mGetAirStatus(ilAstStatus) As Integer
    mGetAirStatus = ilAstStatus Mod 100
End Function

Private Sub mClearPackageOfVefCode(ilVefCode As Integer)
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llPvf As Long
    Dim blUpdate As Boolean
    ReDim llPvfCode(0 To 0) As Long
    
    ilRet = btrGetFirst(hgDBPvf, tgDBPvf, igDBPvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If imCancelled Then
            Exit Sub
        End If
        For ilLoop = LBound(tgDBPvf.iVefCode) To UBound(tgDBPvf.iVefCode) Step 1
            If tgDBPvf.iVefCode(ilLoop) = ilVefCode Then
                llPvfCode(UBound(llPvfCode)) = tgDBPvf.lCode
                ReDim Preserve llPvfCode(0 To UBound(llPvfCode) + 1) As Long
                Exit For
            End If
        Next ilLoop
        ilRet = btrGetNext(hgDBPvf, tgDBPvf, igDBPvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    For llPvf = 0 To UBound(llPvfCode) - 1 Step 1
        DoEvents
        If imCancelled Then
            Exit Sub
        End If
        tgDBPvfSrchKey0.lCode = llPvfCode(llPvf)
        ilRet = btrGetEqual(hgDBPvf, tgDBPvf, igDBPvfRecLen, tgDBPvfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            blUpdate = False
            For ilLoop = LBound(tgDBPvf.iVefCode) To UBound(tgDBPvf.iVefCode) Step 1
                If tgDBPvf.iVefCode(ilLoop) = ilVefCode Then
                    tgDBPvf.iVefCode(ilLoop) = 0
                    blUpdate = True
                End If
            Next ilLoop
            If blUpdate Then
                ilRet = btrUpdate(hgDBPvf, tgDBPvf, igDBPvfRecLen)
            End If
        End If
    Next llPvf
End Sub
