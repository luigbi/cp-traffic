VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ImptGen 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6435
   ClientLeft      =   1035
   ClientTop       =   2790
   ClientWidth     =   9270
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
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6435
   ScaleWidth      =   9270
   Begin VB.Timer tmcCancel 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   285
      Top             =   2895
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   180
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.Frame frcImportOptions 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   -120
      Width           =   9135
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "RCS"
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
         Height          =   210
         Index           =   4
         Left            =   4815
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   735
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "Prophet NexGen"
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
         Height          =   210
         Index           =   2
         Left            =   4815
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   495
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "Dalet"
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
         Height          =   210
         Index           =   0
         Left            =   4815
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "Drake"
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
         Height          =   210
         Index           =   1
         Left            =   6735
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "Scott"
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
         Height          =   210
         Index           =   5
         Left            =   2895
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   975
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "Prophet Wizard"
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
         Height          =   210
         Index           =   3
         Left            =   975
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   735
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "Prophet MediaStar"
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
         Height          =   210
         Index           =   6
         Left            =   6735
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   495
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "iMediaTouch"
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
         Height          =   210
         Index           =   7
         Left            =   975
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "Audio Vault"
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
         Height          =   210
         Index           =   8
         Left            =   975
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "WireReady"
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
         Height          =   210
         Index           =   9
         Left            =   6735
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   975
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "Audio Vault Sat"
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
         Height          =   210
         Index           =   10
         Left            =   2895
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "RCS-5"
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
         Height          =   210
         Index           =   11
         Left            =   6735
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   735
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "Rivendell"
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
         Height          =   210
         Index           =   12
         Left            =   975
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   975
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "Wide Orbit"
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
         Height          =   210
         Index           =   13
         Left            =   4815
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   975
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "Jelli"
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
         Height          =   210
         Index           =   14
         Left            =   2895
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   495
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "Zetta"
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
         Height          =   210
         Index           =   15
         Left            =   975
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1215
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.OptionButton rbcAutoType 
         Caption         =   "RadioMan"
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
         Height          =   210
         Index           =   16
         Left            =   2895
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   735
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.Label lblPlcScreen 
         Appearance      =   0  'Flat
         Caption         =   "Import"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   0
         TabIndex        =   18
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.PictureBox plcControls 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4695
      ScaleWidth      =   9135
      TabIndex        =   19
      Top             =   1560
      Width           =   9135
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
         Height          =   285
         Left            =   4725
         TabIndex        =   33
         Top             =   4305
         Width           =   1050
      End
      Begin VB.CommandButton cmcImport 
         Appearance      =   0  'Flat
         Caption         =   "&Import"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3405
         TabIndex        =   32
         Top             =   4305
         Width           =   1050
      End
      Begin VB.TextBox edcLinkSrceDoneMsg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6150
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox edcLinkDestHelpMsg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5880
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3495
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox edcLinkDestDoneMsg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6495
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   3495
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.PictureBox plcFrom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   930
         ScaleHeight     =   345
         ScaleWidth      =   5865
         TabIndex        =   26
         Top             =   0
         Width           =   5895
         Begin VB.TextBox edcFrom 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   45
            TabIndex        =   27
            Top             =   105
            Width           =   5715
         End
      End
      Begin VB.CommandButton cmcBrowser 
         Appearance      =   0  'Flat
         Caption         =   "Import Browser.."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7350
         TabIndex        =   25
         Top             =   30
         Width           =   1605
      End
      Begin VB.ListBox lbcMsg 
         Appearance      =   0  'Flat
         Height          =   2565
         ItemData        =   "IMPTGEN.frx":0000
         Left            =   930
         List            =   "IMPTGEN.frx":0002
         MultiSelect     =   2  'Extended
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   780
         Width           =   6975
      End
      Begin VB.PictureBox plcPLMove 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   5850
         TabIndex        =   21
         Top             =   450
         Visible         =   0   'False
         Width           =   5850
         Begin VB.OptionButton rbcPLMove 
            Caption         =   "Outsides"
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
            Height          =   210
            Index           =   1
            Left            =   4290
            TabIndex        =   23
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton rbcPLMove 
            Caption         =   "MG's"
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
            Height          =   210
            Index           =   0
            Left            =   3420
            TabIndex        =   22
            Top             =   0
            Width           =   825
         End
      End
      Begin VB.CommandButton cmcBrowse 
         Caption         =   ".."
         Height          =   285
         Left            =   6905
         TabIndex        =   20
         Top             =   30
         Width           =   375
      End
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1080
         TabIndex        =   35
         Top             =   3705
         Width           =   6975
      End
      Begin VB.Label lacMsg 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1080
         TabIndex        =   34
         Top             =   3945
         Width           =   6975
      End
      Begin VB.Label lacFrom 
         Appearance      =   0  'Flat
         Caption         =   "From File"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   0
         TabIndex        =   28
         Top             =   75
         Width           =   810
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   195
      Top             =   3780
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "ImptGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of IMPTGEN.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  imLLf                         imGsf                                                   *
'******************************************************************************************

' Copyright 1993 Counterpoint Software®, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ImptGen.Frm
'
' Release: 5.1: implement Prophet import.  Any spot not in import file is flagged as missed
' Release 5.1 2-6-03 Implement another verionn of Prophet called Wizard (original is NexGen)for the export.
'          The import for both version of Prophet will be the same
' Release 5.5 9-11-06 Wire ready import
'             5-1-20 RadioMan import
' Release 5.5 10-5-06 Audio Vault Sat import
' Description:
'   This file contains the Import RCS input screen code
Option Explicit
Option Compare Text

Dim imFirstActivate As Integer
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim hmFrom As Integer   'From file hanle
Dim hmMsg As Integer   'From file hanle
'Spots
Dim hmSdf As Integer
Dim tmSdf As SDF
Dim tmSdfSrchKey3 As LONGKEY0
Dim tmSdfSrchKey6 As SDFKEY6
Dim imSdfRecLen As Integer
Dim tmImpSdf() As SDF
Dim tmReconVehicles() As RECONVEHICLES

'Spots
Dim hmSmf As Integer
Dim tmSmf As SMF
Dim imSmfRecLen As Integer
Dim hmCHF As Integer
Dim hmClf As Integer
Dim hmCff As Integer
Dim hmRdf As Integer
Dim hmCrf As Integer
Dim hmCgf As Integer
Dim tmSsf As SSF                'SSF record image
Dim hmSsf As Integer
Dim imSsfRecLen As Integer
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmSsfSrchKey1 As SSFKEY1
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS

Dim hmLlf As Integer
Dim tmLlf As LLF
Dim imLlfRecLen As Integer
Dim tmLLFSrchKey1 As LLFKEY1

Dim hmLdf As Integer
Dim tmLdf As LDF
Dim imLdfRecLen As Integer
Dim tmLDFSrchKey1 As LDFKEY1

Dim hmGsf As Integer
Dim tmGsf As GSF
Dim imGsfRecLen As Integer
Dim tmGsfSrchKey4 As GSFKEY4

Dim hmGhf As Integer

Dim hmSxf As Integer

'LCF Calendar Log
Dim hmLcf As Integer        'Log Calendar handle
Dim tmLcf As LCF
Dim imLcfRecLen As Integer
Dim tmLcfSrchKey0 As LCFKEY0

'ANF named avails
Dim hmAnf As Integer
Dim tmAnf As ANF
Dim imAnfRecLen As Integer
Dim tmAnfSrchKey As INTKEY0

Dim imTerminate As Integer
Dim imBypassFocus As Integer
Dim imImporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim lmNowDate As Long
Dim lmSpotDate As Long
Dim lmSpotTime As Long
Dim smFieldValues() As String
Dim smTempFieldValues() As String   '6-27-12 for log vehicles (Wide Orbit and Jelli)
Dim imUsingLiveLog As Integer    '5-22-06
Dim tmIgnoreTimes() As IGNORETIMES
Dim tmPreemptVeh() As PREEMPTVEH
Dim smParsedValues(0 To 14) As String   'for variable length/fields records; this array contains 1 records. index zero ignored
Const AUTOIMPORT_JELLI = 14
Const AUTOIMPORT_ZETTA = 15
Const AUTOIMPORT_RADIOMAN = 16           '5-1-20

                                    'fields that have been parsed
                                    '
'           Build table of the vehicles/games processed to update
'           LCF Day is Complete flag
'           Save the vehicle/game only when a spot is processed
'
Public Sub mSaveVefProc()
Dim ilLoop As Integer
Dim ilFound As Integer
Dim ilUpper As Integer

        ilFound = False
        ilUpper = UBound(tmReconVehicles)
        For ilLoop = LBound(tmReconVehicles) To UBound(tmReconVehicles) - 1
            If tmSdf.iVefCode = tmReconVehicles(ilLoop).iVefCode And tmSdf.iGameNo = tmReconVehicles(ilLoop).iGameNo Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            tmReconVehicles(ilUpper).iVefCode = tmSdf.iVefCode
            tmReconVehicles(ilUpper).iGameNo = tmSdf.iGameNo
            ReDim Preserve tmReconVehicles(0 To ilUpper + 1) As RECONVEHICLES
        End If
        Exit Sub
End Sub


Private Sub cmcBrowse_Click()
    'TTP 10339 - Automation Import - use windows File Browser (to see Local Drives and mapped drives from RDP)
    CMDialogBox.InitDir = sgImportPath
    
    If rbcAutoType(0).Value Then        'Dalet (No Import Defined)
        CMDialogBox.DialogTitle = "Import Dalet File"
        CMDialogBox.Filter = "Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
        CMDialogBox.DefaultExt = "Text (*.txt)"
    
    ElseIf rbcAutoType(1).Value Then    'Drake (No Import Defined)
        CMDialogBox.DialogTitle = "Import From Drake TXT File"
        CMDialogBox.Filter = "Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
        CMDialogBox.DefaultExt = "Text (*.txt)"
        
    ElseIf rbcAutoType(2).Value Then    'Prophet nexgen
        CMDialogBox.DialogTitle = "Import From Prophet nexgen REC File"
        CMDialogBox.Filter = "rec (*.rec)|*.rec|Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
        CMDialogBox.DefaultExt = "rec (*.rec)"
    
    ElseIf rbcAutoType(3).Value Then    'Prophet wizard (2-5-03)
        CMDialogBox.DialogTitle = "Import From Prophet wizard REC File"
        CMDialogBox.Filter = "rec (*.rec)|*.rec|Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
        CMDialogBox.DefaultExt = "rec (*.rec)"
    
    ElseIf rbcAutoType(4).Value Then    'RCS-4
        CMDialogBox.DialogTitle = "Import From RCS-4 REC File"
        CMDialogBox.Filter = "rec (*.rec)|*.rec|Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
        CMDialogBox.DefaultExt = "rec (*.rec)"
    
    ElseIf rbcAutoType(5).Value Then    'Scott (No Import Defined)
        CMDialogBox.DialogTitle = "Import From Scott TXT File"
        CMDialogBox.Filter = "Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
        CMDialogBox.DefaultExt = "Text (*.txt)"
    
    ElseIf rbcAutoType(6).Value Then    'Prophet MediaStar (2-5-03)
        CMDialogBox.DialogTitle = "Import From Prophet MediaStar 003 File"
        CMDialogBox.Filter = "003 (*.003)|*.003|Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
        CMDialogBox.DefaultExt = "003 (*.003)"
    
    ElseIf rbcAutoType(7).Value Then    'IMediaTouch (No Import Defined)
    
    ElseIf rbcAutoType(8).Value Then    'Audio Vault (Nothing seems to enable this - see Audio Vault Sat)
    
    ElseIf rbcAutoType(9).Value Then    'Wire ready (9-11-06)
        CMDialogBox.DialogTitle = "Import From Wire ready REC File"
        CMDialogBox.Filter = "rec (*.rec)|*.rec|Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
        CMDialogBox.DefaultExt = "rec (*.rec)"
    
    ElseIf rbcAutoType(10).Value Then   'Audio vault sat (10-05-06)
        If (Asc(tgSpf.sUsingFeatures3) And INCMEDIACODEAUDIOVAULT) = INCMEDIACODEAUDIOVAULT Then
            CMDialogBox.DialogTitle = "Import From audio vault Log File"
            CMDialogBox.Filter = "log (*.log)|*.log|Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
            CMDialogBox.DefaultExt = "log (*.log)" 'using media codes
        Else
            CMDialogBox.DialogTitle = "Import From audio vault REC File"
            CMDialogBox.Filter = "rec (*.rec)|*.rec|Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
            CMDialogBox.DefaultExt = "rec (*.rec)" 'not using media codes
        End If
   
   ElseIf rbcAutoType(11).Value Then    'RCS-5
        CMDialogBox.DialogTitle = "Import From RCS-5 REC File"
        CMDialogBox.Filter = "rec (*.rec)|*.rec|Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
        CMDialogBox.DefaultExt = "rec (*.rec)"
    
    ElseIf rbcAutoType(12).Value Then   'Rivendell
        CMDialogBox.DialogTitle = "Import From Rivendell cpo File"
        CMDialogBox.Filter = "cpo (*.cpo)|*.air|Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
        CMDialogBox.DefaultExt = "cpo (*.cpo)"
        CMDialogBox.InitDir = sgWideOrbitImportPath
        
    ElseIf rbcAutoType(13).Value Then   'WideOrbit (1-10-12)
        CMDialogBox.DialogTitle = "Import From WideOrbit air File"
        CMDialogBox.Filter = "air (*.air)|*.air|Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
        CMDialogBox.DefaultExt = "air (*.air)"
        CMDialogBox.InitDir = sgWideOrbitImportPath
    
    ElseIf rbcAutoType(14).Value Then   'Jelli (6-26-12) - 11/12/21 - nobody uses,  no site option to enable
        CMDialogBox.DialogTitle = "Import From jelli TXT File"
        CMDialogBox.Filter = "Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
        CMDialogBox.DefaultExt = "Text (*.txt)"
        CMDialogBox.InitDir = sgJelliImportPath
        
    ElseIf rbcAutoType(15).Value Then      'zetta (1-7-16)
        CMDialogBox.DialogTitle = "Import From zetta REC File"
        CMDialogBox.Filter = "rec (*.rec)|*.rec|Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
        CMDialogBox.DefaultExt = "rec (*.rec)"
        CMDialogBox.InitDir = sgZettaImportPath
    
    ElseIf rbcAutoType(16).Value Then      'RadioMan
        CMDialogBox.DialogTitle = "Import From RadioMan REC File"
        CMDialogBox.Filter = "rec (*.rec)|*.rec|Text (*.txt)|*.txt|All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc"
        CMDialogBox.DefaultExt = "rec (*.rec)"
    
    End If
    
    CMDialogBox.Action = 1 'Open dialog
    edcFrom.Text = CMDialogBox.fileName
    
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
End Sub

'' MsgBox parameters
'Const vbOkOnly = 0                 ' OK button only
'Const vbCritical = 16          ' Critical message
'Const vbApplicationModal = 0
'Const INDEXKEY0 = 0
Private Sub cmcBrowser_Click()
    'CMDialogBox.DialogTitle = "Import from File"
    'CMDialogBox.Filter = "Comma|*.CSV|ASC|*.Asc|Text|*.Txt|All|*.*"
    'CMDialogBox.InitDir = Left$(sgImportPath, Len(sgImportPath) - 1)
    'CMDialogBox.DefaultExt = ".Csv"
    'CMDialogBox.Flags = cdlOFNCreatePrompt
    'CMDialogBox.Action = 1 'Open dialog
    'edcFrom.Text = CMDialogBox.Filename
    'ChDir sgCurDir
    
    igProphetImportPathFlag = 1      'set to change path in browser to automation import path

    If rbcAutoType(0).Value Then        'Dalet (No Import Defined)
        sgBrowseMaskFile = "*.*"
    ElseIf rbcAutoType(1).Value Then    'Drake (No Import Defined)
        sgBrowseMaskFile = "*.*"
    ElseIf rbcAutoType(2).Value Then    'Prophet nexgen
        sgBrowseMaskFile = "*.rec"          '12-30-02 chg from .txt to .rec
    ElseIf rbcAutoType(3).Value Then    'Prophet wizard (2-5-03)
        sgBrowseMaskFile = "*.rec"
    ElseIf rbcAutoType(4).Value Then    'RCS-4
        sgBrowseMaskFile = "*.rec"
    ElseIf rbcAutoType(5).Value Then    'Scott (No Import Defined)
        sgBrowseMaskFile = "*.*"
    ElseIf rbcAutoType(6).Value Then    'Prophet MediaStar (2-5-03)
        sgBrowseMaskFile = "*.003"
    ElseIf rbcAutoType(7).Value Then    'iMediaTouch (No Import Defined)
        sgBrowseMaskFile = "*.*"
    ElseIf rbcAutoType(8).Value Then    'Audio Vault (Nothing seems to enable this)
        sgBrowseMaskFile = "*.*"
    ElseIf rbcAutoType(9).Value Then    'Wire ready (9-11-06)
        sgBrowseMaskFile = "*.rec"
    ElseIf rbcAutoType(10).Value Then   'audio vault sat (10-05-06)
        If (Asc(tgSpf.sUsingFeatures3) And INCMEDIACODEAUDIOVAULT) = INCMEDIACODEAUDIOVAULT Then
            sgBrowseMaskFile = "*.log"  'using media codes
        Else
            sgBrowseMaskFile = "*.rec"  'not using media codes
        End If
    ElseIf rbcAutoType(11).Value Then   'RCS-5
        sgBrowseMaskFile = "*.rec"
    ElseIf rbcAutoType(12).Value Then   'Rivendell
        sgBrowseMaskFile = "*.cpo"
    ElseIf rbcAutoType(13).Value Then   'WideOrbit (1-10-12)
        sgBrowseMaskFile = "*.air"
        igProphetImportPathFlag = 2     'set to change path in browser to sgWideOrbitImportPath import path
    ElseIf rbcAutoType(14).Value Then   'jelli (6-26-12) - nobody uses,  no site option to enable
        sgBrowseMaskFile = "*.txt"
        igProphetImportPathFlag = 3     'set to change path in browser to Jelli import path
    ElseIf rbcAutoType(15).Value Then   'zetta (1-7-16)
        sgBrowseMaskFile = "*.rec"
        igProphetImportPathFlag = 4     'set to change path in browser to zetta import path
    ElseIf rbcAutoType(16).Value Then   'RadioMan
        sgBrowseMaskFile = "*.rec"
    End If
    
    '10-23-06 send the browser mask in global field from the init module
    'If rbcAutoType(6).Value Then
    '    sgBrowseMaskFile = "*.003"
        igBrowserType = SHIFT8 + 7
    'Else
    '    igBrowserType = SHIFT8 + 3
    'End If

    'igProphetImportPathFlag = 1      '1-10-12 move to above,set to change path in browser to automation import path for prophet and wide orbit
    Browser.Show vbModal
    igProphetImportPathFlag = 0      'reset to change path back to general import path
    If igBrowserReturn = 1 Then
        edcFrom.Text = sgBrowserFile
        edcFrom.SetFocus
    End If
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
End Sub
Private Sub cmcBrowser_GotFocus()
    lacInfo.Visible = False
End Sub
Private Sub cmcCancel_Click()
    If imImporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
'
'   9-15-06 change prophet import to remove or bump spot if time period is covereed
'   by a game.  Previously, these spots were bypassed.
'
Private Sub cmcImport_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llLLfDate                     llTempTime                    ilLoopPrevCurrent         *
'*                                                                                        *
'******************************************************************************************

    Dim slFromFile As String
    Dim slFileName As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slFYear As String
    Dim slFMonth As String
    Dim slFDay As String
    Dim slDate As String
    Dim slLetter As String
    Dim ilFound As Integer
    Dim slChar As String
    Dim ilVef As Integer
    Dim ilPos As Integer
    Dim ilAutomationType As Integer     '1 = dalet, 2 = prophet (NextGen), 3 = scott, 4=drake, 5 = rcs-4, 6 = Prophet wizard,7=prophet media star,
                                        '8 = unused, 9 =audio vault, 10 = wire ready, 11=rcs-5, 12 = rivendell,  13 = wide orbit, 14 = jelli, 15 = zetta, 16 = radioman
    Dim slAutoDesc As String
    Dim slAutoExt As String
    Dim ilLen As Integer
    Dim llFilePos As Long
    Dim ilUpperSelected As Integer
    Dim ilSelected As Integer
    Dim ilStationCode As Integer
    ReDim slselectedimport(0 To 0) As String
    Dim llImportDate As Long        '5-22-06
    Dim ilImportDate(0 To 1) As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilLoopTime As Integer
    Dim ilTempTime(0 To 1) As Integer
    Dim ilLoopOnVeh As Integer
    Dim ilAutomationType2 As Integer
    Dim ilPass As Integer  '0=Import vehicle; 1=Multi-log vehicle
    Dim ilSvVef As Integer
    Dim ilLogVefCode As Integer
    Dim ilDate(0 To 1) As Integer
    Dim ilLoopOnVEhicles As Integer
    
    cmcCancel.Caption = "Cancel"
    ilAutomationType2 = Asc(tgSpf.sAutoType2)    'continuation of automation equipment types

    ilAutomationType = Asc(tgSpf.sAutoType)
    If rbcAutoType(0).Value Then        'Dalet (No Import Defined)
        ilAutomationType = 1
        slAutoDesc = "Dalet"
        slAutoExt = ".txt"
    ElseIf rbcAutoType(1).Value Then    'Drake (No Import Defined)
        ilAutomationType = 4
        slAutoDesc = "Drake"
        slAutoExt = ".txt"
    ElseIf rbcAutoType(2).Value Then    'Prophet nexgen
        ilAutomationType = 2
        slAutoDesc = "Prophet NexGen"
        slAutoExt = ".rec"              '12-30-02 chg from .txt to .rec
    ElseIf rbcAutoType(3).Value Then    'Prophet wizard (2-5-03 )
        ilAutomationType = 6
        slAutoDesc = "Prophet Wizard"
        slAutoExt = ".rec"
    ElseIf rbcAutoType(4).Value Then    'RCS-4
        ilAutomationType = 5
        slAutoDesc = "RCS_4"
        slAutoExt = ".rec"
    ElseIf rbcAutoType(5).Value Then    'Scott (No Import Defined)
        ilAutomationType = 3
        slAutoDesc = "Scott"
        slAutoExt = ".txt"
    ElseIf rbcAutoType(6).Value Then    'Prophet MediaStar (2-5-03)
        ilAutomationType = 7
        slAutoDesc = "Prophet MediaStar"
        slAutoExt = ".003"
        
    ElseIf rbcAutoType(7).Value Then    'iMediaTouch
    
    ElseIf rbcAutoType(8).Value Then    'Audio Vault
    
    ElseIf rbcAutoType(9).Value Then    'Wire ready (9-11-06)
        ilAutomationType = 10
        slAutoDesc = "WireReady"
        slAutoExt = ".rec"
    ElseIf rbcAutoType(10).Value Then   'Audio Vault sat (10-05-06)
        ilAutomationType = 9
        slAutoDesc = "Audio Vault"
        If (Asc(tgSpf.sUsingFeatures3) And INCMEDIACODEAUDIOVAULT) = INCMEDIACODEAUDIOVAULT Then
            slAutoExt = ".log"          'using media codes
        Else
            slAutoExt = ".rec"          'not using media codes
        End If
    ElseIf rbcAutoType(11).Value Then   'RCS-5
        ilAutomationType = 11
        slAutoDesc = "RCS_5"
        slAutoExt = ".rec"
     ElseIf rbcAutoType(12).Value Then  'Rivendell
        ilAutomationType = 12
        slAutoDesc = "Rivendell"
        slAutoExt = ".cpo"
     ElseIf rbcAutoType(13).Value Then  'Wide Orbit (1-10-12)
        ilAutomationType = 13
        slAutoDesc = "Wide Orbit"
        slAutoExt = ".air"
    ElseIf rbcAutoType(14).Value Then   'Jelli (6-26-12) - 11/12/21 nobody uses,  no site option to enable
        ilAutomationType = AUTOIMPORT_JELLI
        slAutoDesc = "Jelli"
        slAutoExt = ".txt"
    ElseIf rbcAutoType(15).Value Then   'Zetta (1-7-16)
        ilAutomationType = AUTOIMPORT_ZETTA
        slAutoDesc = "Zetta"
        slAutoExt = ".rec"
    ElseIf rbcAutoType(16).Value Then   'Radio Man (5-1-20)
        ilAutomationType = AUTOIMPORT_RADIOMAN
        slAutoDesc = "RadioMan"
        slAutoExt = ".rec"
    End If

    'ilAutomationType = Val(tgspf.sAutoType) 'which interface, 1 = dalet, 2 = prophet, 4 = drake, 3 = scott , 5 = rcs, 13 = wide orbit
    lacInfo.Visible = False
    lbcMsg.Clear
    If imImporting Then
        Exit Sub
    End If
    slFromFile = Trim$(edcFrom.Text)
    ilLen = Len(slFromFile)
    If ilLen = 0 Then
        Beep
        edcFrom.SetFocus
        Exit Sub
    End If
    ilPos = 1
    ilUpperSelected = UBound(slselectedimport)
    Do While ilPos <= ilLen
        llFilePos = InStr(ilPos, slFromFile, "|", 1)
        If llFilePos > 0 Then
            slselectedimport(ilUpperSelected) = Mid$(slFromFile, CLng(ilPos), llFilePos - ilPos)
            ReDim Preserve slselectedimport(0 To ilUpperSelected + 1) As String
            ilUpperSelected = ilUpperSelected + 1
            ilPos = CInt(llFilePos + 1)
        Else
            slselectedimport(ilUpperSelected) = Mid$(slFromFile, CLng(ilPos))
            ReDim Preserve slselectedimport(0 To ilUpperSelected + 1) As String
            ilUpperSelected = ilUpperSelected + 1
            Exit Do
        End If
    Loop


    For ilSelected = 0 To ilUpperSelected - 1
        slFromFile = Trim$(slselectedimport(ilSelected))
        'If InStr(slFromFile, ":") = 0 Then
        If (InStr(slFromFile, ":") = 0) And (Left$(slFromFile, 2) <> "\\") Then
            slFromFile = sgProphetImportPath & slFromFile    '1-6-05 chg from sgImportPath to sgProphetImportpath
        End If
        'Obtain file name- Form is MMDDYYVV.REC where MM=Month; DD=Day; YY=Year; VV=Vehicle Station code
        ilPos = InStr(1, slFromFile, Trim$(slAutoExt), 1)
        If ilPos <= 0 Then
            If (ilAutomationType = 5) Or (ilAutomationType = 11) Then         'rcs
                ilRet = MsgBox("File Name Not in the Form MMDDYYSS" & Trim$(slAutoExt) & " as Required", vbOKOnly + vbCritical, "Import " & Trim$(slAutoDesc))
            ElseIf ilAutomationType = 7 Then         'mediastar
                ilRet = MsgBox("File Name Not in the Form SSMMDDYY" & Trim$(slAutoExt) & " as Required", vbOKOnly + vbCritical, "Import " & Trim$(slAutoDesc))
            ElseIf ilAutomationType = 12 Then         'Rivendell
                ilRet = MsgBox("File Name Not in the Form YYYYMMDD" & Trim$(slAutoExt) & " as Required", vbOKOnly + vbCritical, "Import " & Trim$(slAutoDesc))
            ElseIf ilAutomationType = 9 Or ilAutomationType = AUTOIMPORT_ZETTA Then           'audio vault or Zetta (1-7-16)
                ilRet = MsgBox("File Name Not in the Form YYMMDDSS" & Trim$(slAutoExt) & " as Required", vbOKOnly + vbCritical, "Import " & Trim$(slAutoDesc))
            Else                                 'all automation types are YYMMDDSS except RCS which is MMDDYYSS
                ilRet = MsgBox("File Name Not in the Form YYDDMMSS" & Trim$(slAutoExt) & " as Required", vbOKOnly + vbCritical, "Import " & Trim$(slAutoDesc))
            End If
            edcFrom.SetFocus
            Exit Sub
        End If
        ilPos = ilPos - 6
        Do
            ilPos = ilPos - 1
            slChar = Mid$(slFromFile, ilPos, 1)
        Loop While (slChar <> "\") And (slChar <> ":") And (ilPos > 1)
        If ilPos <= 1 Then
            If (ilAutomationType = 5) Or (ilAutomationType = 11) Then        'rcs
                ilRet = MsgBox("File Name Not in the Form MMDDYYSS.REC as Required", vbOKOnly + vbCritical, "Import " & Trim$(slAutoDesc))
            ElseIf ilAutomationType = 7 Then        'media star
                ilRet = MsgBox("File Name Not in the Form SSMMDDYY.003 as Required", vbOKOnly + vbCritical, "Import " & Trim$(slAutoDesc))
            ElseIf ilAutomationType = 12 Then        'Rivendell
                ilRet = MsgBox("File Name Not in the Form SSYYYYMMDD.cpo as Required", vbOKOnly + vbCritical, "Import " & Trim$(slAutoDesc))
            Else                                'all automation types are YYMMDDSS except RCS which is MMDDYYSS
                ilRet = MsgBox("File Name Not in the Form YYMMDDSS.REC as Required", vbOKOnly + vbCritical, "Import " & Trim$(slAutoDesc))
            End If
            edcFrom.SetFocus
            Exit Sub
        End If

        lbcMsg.Clear
    Next ilSelected
    'end of checking for validity on location

    If Not mOpenMsgFile(slAutoDesc) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    For ilSelected = 0 To ilUpperSelected - 1
        'loop on files toimport

        Screen.MousePointer = vbHourglass
        slFromFile = Trim$(slselectedimport(ilSelected))
        'If InStr(slFromFile, ":") = 0 Then
        If (InStr(slFromFile, ":") = 0) And (Left$(slFromFile, 2) <> "\\") Then
            slFromFile = sgProphetImportPath & slFromFile
        End If

        ilPos = InStr(1, slFromFile, Trim$(slAutoExt), 1)
        ilPos = ilPos - 6
        Do
            ilPos = ilPos - 1
            slChar = Mid$(slFromFile, ilPos, 1)
        Loop While (slChar <> "\") And (slChar <> ":") And (ilPos > 1)

        'Retrieve the date of the import based on filename
        ilPos = ilPos + 1
        slFileName = Mid$(slFromFile, ilPos)
        'ilAutomationType  '1 = dalet, 2 = prophet (NextGen), 3 = scott, 4=drake, 5 = rcs-4, 6 = Prophet wizard,7=prophet media star,
                           '8 = unused, 9 =audio vault, 10 = wire ready, 11=rcs-5, 12 = rivendell,  13 = wide orbit, 14 = jelli, 15 = zetta, 16 = radidoman
        'automationtype not rcs-4 and rcs5, prophet media star and rivendell and jelli
        If (ilAutomationType <> 5) And (ilAutomationType <> 7) And (ilAutomationType <> 11) And (ilAutomationType <> 12) And ilAutomationType <> AUTOIMPORT_JELLI Then            'all automation types are YYMMDDSS except RCS which is MMDDYYSS
            ilPos = 1
            slFYear = Mid$(slFileName, ilPos, 2)
            ilPos = ilPos + 2
            slFMonth = Mid$(slFileName, ilPos, 2)
            ilPos = ilPos + 2
            slFDay = Mid$(slFileName, ilPos, 2)

        ElseIf ilAutomationType = 7 Then        'SSMMDDYY
            ilPos = InStr(slFileName, ".") - 6
            slFMonth = Mid$(slFileName, ilPos, 2)
            ilPos = ilPos + 2
            slFDay = Mid$(slFileName, ilPos, 2)
            ilPos = ilPos + 2
            slFYear = Mid$(slFileName, ilPos, 2)
        ElseIf ilAutomationType = 12 Then        'SSYYYYMMDD
            ilPos = InStr(slFileName, ".") - 8
            slFYear = Mid$(slFileName, ilPos, 4)
            ilPos = ilPos + 4
            slFMonth = Mid$(slFileName, ilPos, 2)
            ilPos = ilPos + 2
            slFDay = Mid$(slFileName, ilPos, 2)
        ElseIf ilAutomationType = AUTOIMPORT_JELLI Then          '6-27-12
            ilPos = InStr(slFileName, ".") - 9
            slFMonth = Mid$(slFileName, ilPos, 2)
            ilPos = ilPos + 2
            slFDay = Mid$(slFileName, ilPos, 2)
            ilPos = ilPos + 2
            slFYear = Mid$(slFileName, ilPos, 2)
        Else            'yymmddssss
            ilPos = 1
            slFMonth = Mid$(slFileName, ilPos, 2)
            ilPos = ilPos + 2
            slFDay = Mid$(slFileName, ilPos, 2)
            ilPos = ilPos + 2
            slFYear = Mid$(slFileName, ilPos, 2)
        End If

        slDate = slFMonth & "/" & slFDay & "/" & slFYear
        If ilAutomationType = 7 Then        'SSMMDDYY
            ilPos = 1
            ilStationCode = InStr(slFileName, ".")
            ilStationCode = ilStationCode - 7
            slLetter = Mid$(slFileName, ilPos, ilStationCode)
       ElseIf ilAutomationType = 12 Then        'SSYYYYMMDD
            ilPos = 1
            ilStationCode = InStr(slFileName, ".")
            ilStationCode = ilStationCode - 9
            slLetter = Mid$(slFileName, ilPos, ilStationCode)
        ElseIf ilAutomationType = 13 Then       '1-10-12 wide orbit; get the station code from foldername
            'need to find the subfolder name that contains the station code to determine the vehicle to process
            ilPos = InStr(1, slFromFile, Trim$(slAutoExt), 1)
            ilPos = ilPos - 8
            slLetter = ""
            Do
                slLetter = Mid$(slFromFile, ilPos, 1) & slLetter
                ilPos = ilPos - 1
            Loop While (Mid$(slFromFile, ilPos, 1) <> "\") And (ilPos > 1)
        ElseIf ilAutomationType = AUTOIMPORT_JELLI Then            '6-27-12
            'need to find the subfolder name that contains the station code to determine the vehicle to process
            ilPos = InStr(1, slFromFile, Trim$(slAutoExt), 1)
            ilPos = ilPos - 17     'backup past the filename (jelli-mmddyyair.txt)
            slLetter = ""
            Do
                slLetter = Mid$(slFromFile, ilPos, 1) & slLetter
                ilPos = ilPos - 1
            Loop While (Mid$(slFromFile, ilPos, 1) <> "\") And (ilPos > 1)
        Else
            ilPos = ilPos + 2
            ilStationCode = InStr(slFileName, ".")
            ilStationCode = ilStationCode - ilPos
            slLetter = Mid$(slFileName, ilPos, ilStationCode)
        End If
        'Search the vehicle table for the matching station code.  This is the vehicle to be updated
        ilVef = -1
        ilFound = False
        For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            '6/15/10: 30 was matching against 300
            'If StrComp(slLetter, Trim$(Left$(tgMVef(ilLoop).sCodeStn, ilStationCode)), 1) = 0 Then
            If StrComp(slLetter, Trim$(tgMVef(ilLoop).sCodeStn), 1) = 0 Then
                ilVef = ilLoop
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If ilFound Then         'valid station code

            '5-22-06 if sports, find the Live Log times that should be ignored
            ReDim tmIgnoreTimes(0 To 0) As IGNORETIMES
            ReDim tmPreemptVeh(0 To 0) As PREEMPTVEH
            ilLoopTime = 0
            If imUsingLiveLog Then
                llImportDate = gDateValue(slDate)       'find all live logs for this import date
                'build array of their times so that the imported spots can be ignored and the
                'existing spots in SDF can be ignore.  Test for any live log crossing midnight.
                'for any logs crossing midnight, only go to 12M.
                'go to previous date to see if any logs crossed midnight there.  then
                'create times in array from 12m to the end time of the live log.
                'i.e.     Import Sun 6/2/06.  Games goes from 10p - 1pm.
                '         Build array from 10p - 12m.
                '         12m -1p will be build when Mon 6/3/06 is imported
                '         Then for import date of 6/2/06, look at previous day to see
                '         if anything went past 12m.  Continue building array from 12m to that
                '         live logs end time.  If live log aired on 6/1/06 and went into Monday 6/2
                '         to 2AM, continue building array from 12M - 2a for the 6/2 import date
                '
                gPackDateLong llImportDate, ilImportDate(0), ilImportDate(1)        'need btrieve import date
                tmGsfSrchKey4.iAirDate(0) = ilImportDate(0)
                tmGsfSrchKey4.iAirDate(1) = ilImportDate(1)
                tmGsfSrchKey4.iAirVefCode = tgMVef(ilVef).iCode
                ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)
                'gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llLLfDate
                Do While (ilRet = BTRV_ERR_NONE) And (ilImportDate(0) = tmGsf.iAirDate(0) And ilImportDate(1) = tmGsf.iAirDate(1) And tgMVef(ilVef).iCode = tmGsf.iAirVefCode)
                    tmPreemptVeh(ilLoopTime).iVefCode = tmGsf.iVefCode
                    tmPreemptVeh(ilLoopTime).iAirDate(0) = tmGsf.iAirDate(0)
                    tmPreemptVeh(ilLoopTime).iAirDate(1) = tmGsf.iAirDate(1)
                    tmPreemptVeh(ilLoopTime).iPrevDay = False
                    tmPreemptVeh(ilLoopTime).iGameNo = tmGsf.iGameNo          '9-15-06
                    ReDim Preserve tmPreemptVeh(0 To UBound(tmPreemptVeh) + 1)
                    ilLoopTime = ilLoopTime + 1
                    ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop

                gPackDateLong llImportDate - 1, ilImportDate(0), ilImportDate(1)      'need btrieve import date
                tmGsfSrchKey4.iAirDate(0) = ilImportDate(0)
                tmGsfSrchKey4.iAirDate(1) = ilImportDate(1)
                tmGsfSrchKey4.iAirVefCode = tgMVef(ilVef).iCode
                ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)
                'gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llLLfDate
                Do While (ilRet = BTRV_ERR_NONE) And (ilImportDate(0) = tmGsf.iAirDate(0) And ilImportDate(1) = tmGsf.iAirDate(1) And tgMVef(ilVef).iCode = tmGsf.iAirVefCode)
                    tmPreemptVeh(ilLoopTime).iVefCode = tmGsf.iVefCode
                    tmPreemptVeh(ilLoopTime).iAirDate(0) = tmGsf.iAirDate(0)
                    tmPreemptVeh(ilLoopTime).iAirDate(1) = tmGsf.iAirDate(1)
                    tmPreemptVeh(ilLoopTime).iPrevDay = True
                    tmPreemptVeh(ilLoopTime).iGameNo = tmGsf.iGameNo            '9-15-06
                    ReDim Preserve tmPreemptVeh(0 To UBound(tmPreemptVeh) + 1)
                    ilLoopTime = ilLoopTime + 1
                    ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop

                ilLoopTime = 0
                'Go thru the preempt vehicle table to find which airing vehicle for the  vehicle that is preempted (which is being imported)
                For ilLoopOnVeh = 0 To UBound(tmPreemptVeh) - 1
                    'setup date and vehicle key to LLF
                    tmLLFSrchKey1.iAirDate(0) = tmPreemptVeh(ilLoopOnVeh).iAirDate(0)
                    tmLLFSrchKey1.iAirDate(1) = tmPreemptVeh(ilLoopOnVeh).iAirDate(1)
                    tmLLFSrchKey1.iVefCode = tmPreemptVeh(ilLoopOnVeh).iVefCode
                    ilRet = btrGetEqual(hmLlf, tmLlf, imLlfRecLen, tmLLFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                    'gUnpackDateLong tmLlf.iAirDate(0), tmLlf.iAirDate(1), llLLfDate
                    Do While (ilRet = BTRV_ERR_NONE) And (tmPreemptVeh(ilLoopOnVeh).iAirDate(0) = tmLlf.iAirDate(0) And tmPreemptVeh(ilLoopOnVeh).iAirDate(1) = tmLlf.iAirDate(1) And tmPreemptVeh(ilLoopOnVeh).iVefCode = tmLlf.iVefCode)
                        If tmLlf.sLogCompleted = "Y" And tmLlf.sReplaceStatus = "N" And tmPreemptVeh(ilLoopOnVeh).iGameNo = tmLlf.iGameNo Then       'log must be completed before using it to determine if sports ran
                            llEndTime = 99999
                            '4/3/12: Get start time if exist, use instead of LCF start time
                            llStartTime = 99999
                            tmLDFSrchKey1.lLlfCode = tmLlf.lCode
                            ilRet = btrGetEqual(hmLdf, tmLdf, imLdfRecLen, tmLDFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                            Do While (ilRet = BTRV_ERR_NONE) And (tmLdf.lLlfCode = tmLlf.lCode)
                                If (tmLdf.sType = "G") Or (tmLdf.sType = "P") Then
                                    '4/3/12: Get start time if exist, use instead of LCF start time
                                    If tmLdf.sSubType = "S" Then
                                        gUnpackTimeLong tmLdf.iTime(0), tmLdf.iTime(1), False, llStartTime
                                    End If
                                    If tmLdf.sSubType = "E" Then
                                        gUnpackTimeLong tmLdf.iTime(0), tmLdf.iTime(1), False, llEndTime
                                        'Exit Do
                                    End If
                                    '4/3/12: Get start time if exist, use instead of LCF start time
                                    If (llStartTime <> 99999) And (llEndTime <> 99999) Then
                                        Exit Do
                                    End If
                                End If
                                ilRet = btrGetNext(hmLdf, tmLdf, imLdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                            '11/1/11: Bypass Ignore time if Prog/Game end time not defined
                            If (llEndTime <> 99999) Then
                                '4/3/12: Get start time if exist, use instead of LCF start time
                                If (llStartTime = 99999) Then
                                    gUnpackTimeLong tmLlf.iLcfStartTime(0), tmLlf.iLcfStartTime(1), False, llStartTime
                                End If
                                
                                If llEndTime < llStartTime Then       'this program cross midnight to next day
                                                                            'create times only to 12M for this day to import
                                    If tmPreemptVeh(ilLoopOnVeh).iPrevDay = False Then  'importing current day (not
                                        tmIgnoreTimes(ilLoopTime).lStartTime = llStartTime
                                        gPackTime "12M", ilTempTime(0), ilTempTime(1)       'consider this the start of day
                                        gUnpackTimeLong ilTempTime(0), ilTempTime(1), True, tmIgnoreTimes(ilLoopTime).lEndTime
                                    Else                        'prev day, setup start of day cross midnighttimes
                                        gPackTime "12M", ilTempTime(0), ilTempTime(1)       'consider this the start of day
                                        gUnpackTimeLong ilTempTime(0), ilTempTime(1), False, tmIgnoreTimes(ilLoopTime).lEndTime
                                        tmIgnoreTimes(ilLoopTime).lEndTime = llEndTime
                                    End If
                                    ReDim Preserve tmIgnoreTimes(0 To UBound(tmIgnoreTimes) + 1)
                                    ilLoopTime = ilLoopTime + 1
                                Else                            'no crossing midnight
                                    If tmPreemptVeh(ilLoopOnVeh).iPrevDay = False Then      'current day
                                        tmIgnoreTimes(ilLoopTime).lStartTime = llStartTime
                                        tmIgnoreTimes(ilLoopTime).lEndTime = llEndTime
                                        ReDim Preserve tmIgnoreTimes(0 To UBound(tmIgnoreTimes) + 1)
                                        ilLoopTime = ilLoopTime + 1
                                    End If
                                End If
                                '4/3/12: Moved with If above
                                'ReDim Preserve tmIgnoreTimes(0 To UBound(tmIgnoreTimes) + 1)
                                'ilLoopTime = ilLoopTime + 1
                            End If
                        End If
                        ilRet = btrGetNext(hmLlf, tmLlf, imLlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                Next ilLoopOnVeh            'find next vehicles livelog times
            End If                      'endif using live log

            imImporting = True
            'Output date record
            lacInfo.Visible = True
            lacInfo.Caption = "Reading Import File " & slFileName
            DoEvents
            Print #hmMsg, ""
            Print #hmMsg, "** Reading Import File " & Trim$(slFileName) & ", Updating " & Trim$(tgMVef(ilVef).sName) & " for " & slDate & " **"
            ilRet = mReadFile(slFromFile)
            If ilRet Then
                gPackDate slDate, ilDate(0), ilDate(1)
                'Update files
                For ilPass = 0 To 1 Step 1
                    If tgMVef(ilVef).sType = "L" Then
                        ilLogVefCode = tgMVef(ilVef).iCode
                        For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            If ilLogVefCode = tgMVef(ilLoop).iVefCode Then
                                lacInfo.Caption = "Updating " & Trim$(tgMVef(ilLoop).sName) & " for " & slDate
                                Print #hmMsg, "** Updating " & Trim$(tgMVef(ilLoop).sName) & " for " & slDate
                                DoEvents
                                ilRet = mSaveRec(slDate, tgMVef(ilLoop).iCode, ilAutomationType, tgMVef(ilLoop).iCombineVefCode, tgMVef(ilLoop).sCodeStn)
                                If ilRet Then
                                    '4-27-11 update day is complete flag on successful import
                                    For ilLoopOnVEhicles = LBound(tmReconVehicles) To UBound(tmReconVehicles) - 1
                                        gUpdateLCFCompleteFlag hmLcf, tmLcf, ilDate(0), ilDate(1), tmReconVehicles(ilLoopOnVEhicles).iGameNo, tmReconVehicles(ilLoopOnVEhicles).iVefCode, "C"
                                    Next ilLoopOnVEhicles

                                    Print #hmMsg, "** Import Complete: " & Trim$(tgMVef(ilLoop).sName) & " for " & slDate; " on " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
                                    lacInfo.Caption = "Import Complete " & Trim$(tgMVef(ilLoop).sName) & " for " & slDate
                                    DoEvents
                                Else
                                    Print #hmMsg, "** Import Failed " & Trim$(tgMVef(ilLoop).sName) & " for " & slDate
                                    lacInfo.Caption = "Import Failed " & Trim$(tgMVef(ilLoop).sName) & " for " & slDate
                                    DoEvents
                                End If
                            End If
                        Next ilLoop
                        Exit For
                    Else
                        lacInfo.Caption = "Updating " & Trim$(tgMVef(ilVef).sName) & " on " & slDate
                        DoEvents
                        ilRet = mSaveRec(slDate, tgMVef(ilVef).iCode, ilAutomationType, tgMVef(ilVef).iCombineVefCode, tgMVef(ilLoop).sCodeStn)
                        If ilRet Then
                            '4-27-11 update day is complete flag on successful import
                            For ilLoopOnVEhicles = LBound(tmReconVehicles) To UBound(tmReconVehicles) - 1
                                gUpdateLCFCompleteFlag hmLcf, tmLcf, ilDate(0), ilDate(1), tmReconVehicles(ilLoopOnVEhicles).iGameNo, tmReconVehicles(ilLoopOnVEhicles).iVefCode, "C"
                            Next ilLoopOnVEhicles

                            Print #hmMsg, "** Import Complete: " & Trim$(tgMVef(ilVef).sName) & " for " & slDate; " on " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
                            lacInfo.Caption = "Import Complete " & Trim$(tgMVef(ilVef).sName) & " for " & slDate
                            DoEvents
                        Else
                            Print #hmMsg, "** Import Failed " & Trim$(tgMVef(ilVef).sName) & " for " & slDate
                            lacInfo.Caption = "Import Failed " & Trim$(tgMVef(ilVef).sName) & " for " & slDate
                            DoEvents
                        End If
                        If (tgMVef(ilVef).iCombineVefCode > 0) And (ilPass = 0) Then
                            'ilVef = 0
                            ilFound = False
                            ilSvVef = ilVef
                            For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                If tgMVef(ilVef).iCombineVefCode = tgMVef(ilLoop).iCode Then
                                    ilVef = ilLoop
                                    'Set Combine so that spots will be ignored in pass 1
                                    tgMVef(ilVef).iCombineVefCode = tgMVef(ilSvVef).iCode
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then         'valid station code
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    End If
                Next ilPass
            Else
                Print #hmMsg, "** Import Failed or Terminated " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
                lacInfo.Caption = "Import Failed"
            End If
        Else                'no station code found
            Print #hmMsg, "** Import Failed: Missing Station Code for " & slFileName
            lacInfo.Caption = "Import Failed: Missing Station Code for " & slFileName
        End If
        If imTerminate Then
            ilSelected = ilUpperSelected    'force to get out of loop
        End If
    Next ilSelected
    If imTerminate Then
        lacInfo.Caption = "** Import Terminated " & Trim$(slAutoDesc) & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Print #hmMsg, "** Import Terminated " & Trim$(slAutoDesc) & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Print #hmMsg, ""
        Close #hmMsg
    Else
        Print #hmMsg, "** Import Complete " & Trim$(slAutoDesc) & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
        Print #hmMsg, ""
        Close #hmMsg
    End If

    lacMsg.Caption = "Messages sent to " & sgDBPath & "Messages\" & "ImptGen.Txt"

    cmcCancel.Caption = "Done"
    lacInfo.Visible = True
    Screen.MousePointer = vbDefault
    imImporting = False
    Exit Sub

    ilRet = err.Number
    Resume Next
End Sub
Private Sub edcFrom_Change()
    'slName = Trim$(edcFrom.Text)
    'If Len(slName) <= 0 Then
    '    edcFrom.SetFocus
    '    Exit Sub
    'End If
End Sub
Private Sub edcFrom_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
        'mInitDDE
        'mSendHelpMsg "BT"
    End If
    lacInfo.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub Form_Activate()
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

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        'plcCalendar.Visible = False
        gFunctionKeyBranch KeyCode
    End If

End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        'cmcCancel_Click
        tmcCancel.Enabled = True
        Me.Left = 2 * Screen.Width      'move off the screen so screen won't flash
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase smFieldValues
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    btrDestroy hmCHF
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmRdf
    btrDestroy hmCrf
    btrDestroy hmCgf
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    btrDestroy hmGhf
    btrDestroy hmSxf
    ilRet = btrClose(hmAnf)
    btrDestroy hmAnf

    If (imUsingLiveLog) Then      '5-22--06 using sports
        ilRet = btrClose(hmLlf)
        btrDestroy hmLlf
        ilRet = btrClose(hmLdf)
        btrDestroy hmLdf
    End If
    
    Set ImptGen = Nothing   'Remove data segment

End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
'
'
'           Convert SDF Internal Code from string to Long
'
'           <input> Recd to process from smFieldValues
'                   ilStartPos - start position of field to convert
'                   illen -length of field to convert
'           <return> error return code: True = found SDF, else False
'                   llLong - Field converted to a long variable
Private Function mConvertLong(tlFieldValues As String, ilStartPos, ilLen) As Long
Dim slStr As String
Dim llLong As Long

    slStr = Mid$(tlFieldValues, ilStartPos, ilLen)
    llLong = Val(Trim$(slStr))
    mConvertLong = llLong
End Function
'
'
'           Convert HHMMSS (military time to 12hr time
'
'           <input> tlFieldValues - record to process from smFieldValue array
'                   HrStartPos - starting location of Hour field (relative to 1)
'                   MinStartPos - starting location of minutes field (relative to 1)
'                   SecStartPos - starting location of seconds field (relative to 1)
'                   (assume all are 2 characters)
'           <output> slXMid - Y/N
'           <return> air time as Long
'
'       3-28-07 when vehicle cross midnight, automation import comes in as a time past 24:00.
'               Need to adjust for x-midnight
Private Function mConvertTime(tlFieldValues As String, ilHrStartPos As Integer, ilMinStartPos As Integer, ilSecStartPos As Integer, slXMid As String) As Long
Dim slHour As String
Dim slMin As String
Dim slSec As String
Dim slTime As String
Dim llAirTime As Long

        slXMid = "N"
        mConvertTime = 0
        slHour = Mid$(tlFieldValues, ilHrStartPos, 2)
        slMin = Mid$(tlFieldValues, ilMinStartPos, 2)
        If Trim(slMin) = "" Then
            slMin = "00"
        End If
        slSec = Mid$(tlFieldValues, ilSecStartPos, 2)
        If Trim(slSec) = "" Then
            slSec = "00"
        End If

        If Val(slHour) >= 24 Then       'cross midnight, subtract 24hr
            slTime = Trim$(str$(Val(slHour) - 24)) & ":" & slMin & ":" & slSec & "AM"
            slXMid = "Y"
        ElseIf Val(slHour) = 0 Then
            slTime = "12" & ":" & slMin & ":" & slSec & "AM"
        ElseIf Val(slHour) <= 11 Then
            slTime = slHour & ":" & slMin & ":" & slSec & "AM"
        ElseIf Val(slHour) = 12 Then
            slTime = slHour & ":" & slMin & ":" & slSec & "PM"
        Else
            slTime = Trim$(str$(Val(slHour) - 12)) & ":" & slMin & ":" & slSec & "PM"
        End If
        llAirTime = gTimeToLong(slTime, False)
        slTime = gFormatTimeLong(llAirTime, "A", "1")

        mConvertTime = llAirTime            'return converted time to caller
End Function
'
'
'           mFindSDF
'           <input> llSdfCode - Spot code to obtain
'                   slXMid = "Y" if cross midnight spot (set flag in SDF)
'           <return> True if spot found, else false
'
Private Function mFindSDF(llSdfCode As Long, slXMid As String) As Integer
Dim ilRet As Integer
    mFindSDF = False
    tmSdfSrchKey3.lCode = llSdfCode
    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, 3, BTRV_LOCK_NONE, SETFORREADONLY)  'position record so update works
    If ilRet <> BTRV_ERR_NONE Then
        'write error log in main program so that the record can be dumped into error log
        'If llSdfCode = 0 Then
         '   lbcMsg.AddItem "Internal spot code missing for: " & slAdvt
        '    Print #hmMsg, "Internal spot code missing for: "
        'Else
        '    lbcMsg.AddItem "Spot missing for internal code " & Str$(llSdfCode) & ", " & slAdvt
        '    Print #hmMsg, "Spot missing for internal code " & Str$(llSdfCode) & ", " & slAdvt
        'End If
    Else
        'convert date & time
        gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, lmSpotTime
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), lmSpotDate
        tmSdf.sXCrossMidnight = slXMid
        mFindSDF = True
    End If
End Function
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
    Dim slDate As String
    Dim ilHowMany As Integer
    Dim ilCols As Integer
    Dim ilRows As Integer
    Dim ilStartPos As Integer
    Dim ilStartRow As Integer
    Dim ilAutomationType As Integer
    Dim ilAutomationType2 As Integer
    Dim ilAutomationType3 As Integer
    Dim ilLoop As Integer
    imTerminate = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    imImporting = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoBytes = 0
    lmProcessedNoBytes = 0
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    hmSdf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptGen
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    hmSmf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptGen
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptGen
    On Error GoTo 0
    imSsfRecLen = Len(tmSsf)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptGen
    On Error GoTo 0
    'imChfRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptGen
    On Error GoTo 0
    'imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptGen
    On Error GoTo 0
    'imCffRecLen = Len(tmCff)
    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptGen
    On Error GoTo 0
    'imRdfRecLen = Len(tmRdf)
    hmCrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptGen
    On Error GoTo 0
    'imCrfRecLen = Len(tmCrf)
    hmCgf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCgf, "", sgDBPath & "Cgf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptGen
    On Error GoTo 0
    
    '4-26-11 Lcf table to set "Day is Not Complete flag"
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:LCF)", ImptGen
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf)

    hmGsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imGsfRecLen = Len(tmGsf)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptGen
    On Error GoTo 0
    
    hmGhf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptGen
    On Error GoTo 0
    
    hmSxf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmSxf, "", sgDBPath & "Sxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptGen
    On Error GoTo 0
    
    '10-4-12 ANF table to determine excluded avails
    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imAnfRecLen = Len(tmAnf)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptGen
    On Error GoTo 0

    If (Asc(tgSpf.sUsingFeatures) And USINGLIVELOG) = USINGLIVELOG Then   '5-22-06
        imUsingLiveLog = True
        hmLlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmLlf, "", sgDBPath & "Llf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        imLlfRecLen = Len(tmLlf)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptGen
        On Error GoTo 0

        hmLdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmLdf, "", sgDBPath & "Ldf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        imLdfRecLen = Len(tmLdf)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptGen
        On Error GoTo 0

    Else
        imUsingLiveLog = False
    End If

    'Populate arrays to determine if records exist
    ilRet = gObtainVef()    'Populate tgMVef
    
    '10-4-12 get the named avails for excluded ones
    ilRet = gObtainAvail()     'tgavailanf
    If ilRet = False Then
        'all avails will beincluded
    End If

    For ilLoop = 0 To 16
        rbcAutoType(ilLoop).Enabled = True
        rbcAutoType(ilLoop).Visible = False
    Next ilLoop

    ilHowMany = 0
    ilCols = 1
    ilRows = 1
    ilStartPos = 780
    ilStartRow = 240
    ilAutomationType = Asc(tgSpf.sAutoType)
    ilAutomationType2 = Asc(tgSpf.sAutoType2)       'continuation of automation types
    ilAutomationType3 = Asc(tgSpf.sAutoType3)        'continuation of automation types
    
    
    '----------------------------------------------------------------------------------------
    'Put these in Alphabetical order, and they will automatically arrange in 4 column Grid
    'Audio Vault
    If (ilAutomationType2 And (AUDIOVAULT)) = (AUDIOVAULT) Then        '10-05-6
        rbcAutoType(10).Left = ilStartPos
        ilStartPos = ilStartPos + rbcAutoType(10).Width + 60
        rbcAutoType(10).Top = ilStartRow
        ilHowMany = ilHowMany + 1
        ilCols = ilCols + 1
        If ilCols > 4 Then
            ilCols = 1
            ilStartRow = ilStartRow + 240
            ilRows = ilRows + 1
            ilStartPos = 780
        End If
        If ilHowMany = 1 Then
            rbcAutoType(10).Value = True 'Audio Vault sat
        End If
        rbcAutoType(10).Visible = True
        '10-23-06 determine the extension to look for Audio Vault based on if using Media definitions because
        'a version of Audio Vault already existted using the *.rec extension.  Vendors are not flexible in their
        'naming conventions.
        If (Asc(tgSpf.sUsingFeatures3) And INCMEDIACODEAUDIOVAULT) = INCMEDIACODEAUDIOVAULT Then
            sgBrowseMaskFile = "*.log"      'using media codes
        Else
            sgBrowseMaskFile = "*.rec"      'not using media codes
        End If
    End If
    
    'Dalet - no import interface yet
    '    If (ilAutomationType And DALET) = DALET Then
    '        'plcScreen.Caption = "Import Dalet"
    '        lblPlcScreen.Caption = "Import Dalet"
    '        rbcAutoType(0).Left = ilStartPos
    '        'rbcAutoType(0).Top = 45
    '        rbcAutoType(0).Top = ilStartRow
    '        ilStartPos = ilStartPos + rbcAutoType(0).Width + 60
    '        ilHowMany = ilHowMany + 1
    '        ilCols = ilCols + 1
    '        If ilCols > 4 Then
    '            ilRows =ilRows +1
    '            ilCols = 1
    '            ilStartRow = ilStartRow + 240
    '            ilStartPos = 780
    '        End If
    '        rbcAutoType(0).Visible = True
    '        rbcAutoType(0).Enabled = False  'no import interface yet
    '        sgBrowseMaskFile = "*.txt"          '10-23-06
    '    End If
    
    'Drake - no import interface yet
    '    If (ilAutomationType And DRAKE) = DRAKE Then
    '        'plcScreen.Caption = "Import Prophet"
    '        lblPlcScreen.Caption = "Import Drake"
    '        rbcAutoType(1).Left = ilStartPos
    '        'rbcAutoType(1).Top = 45
    '        ilStartPos = ilStartPos + rbcAutoType(1).Width + 60
    '        rbcAutoType(1).Top = ilStartRow
    '        ilHowMany = ilHowMany + 1
    '        ilCols = ilCols + 1
    '        If ilCols > 4 Then
    '            ilRows =ilRows +1
    '            ilCols = 1
    '            ilStartRow = ilStartRow + 240
    '            ilStartPos = 780
    '        End If
    '        If ilHowMany = 1 Then
    '            rbcAutoType(1).Value = True
    '        End If
    '        rbcAutoType(1).Visible = True
    '        rbcAutoType(1).Enabled = False  'no import interface yet
    '        sgBrowseMaskFile = "*.txt"          '10-23-06
    '    End If

    'IMediaTouch - No Import Defined
    '    If (ilAutomationType And IMEDIATOUCH) = IMEDIATOUCH Then  'IMediaTouch
    '        rbcAutoType(7).Left = ilStartPos
    '        'rbcAutoType(14).Top = 45
    '        ilStartPos = ilStartPos + rbcAutoType(7).Width + 60
    '        rbcAutoType(7).Top = ilStartRow
    '        ilHowMany = ilHowMany + 1
    '        ilCols = ilCols + 1
    '        If ilCols > 4 Then
    '            ilRows = ilRows + 1
    '            ilCols = 1
    '            ilStartRow = ilStartRow + 240
    '            ilStartPos = 780
    '        End If
    '        If ilHowMany = 1 Then
    '            rbcAutoType(7).Value = True
    '        End If
    '        rbcAutoType(7).Visible = True
    '        rbcAutoType(7).Enabled = False  'no import interface yet
    '        'sgBrowseMaskFile = "*.txt"          '???
    '    End If

    'Jelli - nobody uses,  no site option to enable
    If (ilAutomationType3 And JELLI) = JELLI Then      '6-26-12
        rbcAutoType(14).Left = ilStartPos
        ilStartPos = ilStartPos + rbcAutoType(14).Width + 60
        rbcAutoType(14).Top = ilStartRow
        ilHowMany = ilHowMany + 1
        ilCols = ilCols + 1
        If ilCols > 4 Then
            ilRows = ilRows + 1
            ilCols = 1
            ilStartRow = ilStartRow + 240
            ilStartPos = 780
        End If
        If ilHowMany = 1 Then
            rbcAutoType(14).Value = True
        End If
        rbcAutoType(14).Visible = True
        sgBrowseMaskFile = "*.txt"          '6-26-12
    End If
    
    'Prophet NexGen
    If (ilAutomationType And PROPHETNEXGEN) = PROPHETNEXGEN Then
        rbcAutoType(2).Left = ilStartPos
        ilStartPos = ilStartPos + rbcAutoType(2).Width + 60
        rbcAutoType(2).Top = ilStartRow
        ilHowMany = ilHowMany + 1
        ilCols = ilCols + 1
        If ilCols > 4 Then
            ilRows = ilRows + 1
            ilCols = 1
            ilStartRow = ilStartRow + 240
            ilStartPos = 780
        End If
        If ilHowMany = 1 Then
            rbcAutoType(2).Value = True 'Prophet nexgen
        End If
        rbcAutoType(2).Visible = True
        sgBrowseMaskFile = "*.rec"          '10-23-06
    End If
    
    'Prophet MediaStar
    If (ilAutomationType And PROPHETMEDIASTAR) = PROPHETMEDIASTAR Then      '2-5-03
        rbcAutoType(6).Left = ilStartPos
        ilStartPos = ilStartPos + rbcAutoType(6).Width + 60
        rbcAutoType(6).Top = ilStartRow
        ilHowMany = ilHowMany + 1
        ilCols = ilCols + 1
        If ilCols > 4 Then
            ilRows = ilRows + 1
            ilCols = 1
            ilStartRow = ilStartRow + 240
            ilStartPos = 780
        End If
        If ilHowMany = 1 Then
            rbcAutoType(6).Value = True 'Prophet MediaStar
        End If
        rbcAutoType(6).Visible = True
        sgBrowseMaskFile = "*.003"          '10-23-06
    End If
    
    'Prophet wizard
    If (ilAutomationType And PROPHETWIZARD) = PROPHETWIZARD Then
        rbcAutoType(3).Left = ilStartPos
        ilStartPos = ilStartPos + rbcAutoType(3).Width + 60
        rbcAutoType(3).Top = ilStartRow
        ilHowMany = ilHowMany + 1
        ilCols = ilCols + 1
        If ilCols > 4 Then
            ilRows = ilRows + 1
            ilCols = 1
            ilStartRow = ilStartRow + 240
            ilStartPos = 780
        End If
        If ilHowMany = 1 Then
            rbcAutoType(3).Value = True 'Prophet wizard
        End If
        rbcAutoType(3).Visible = True
        sgBrowseMaskFile = "*.rec"          '10-23-06
    End If
    
    'RadioMan
    If (ilAutomationType3 And RADIOMAN) = RADIOMAN Then      '5-1-20
        'lblPlcScreen.Caption = "Import RadioMan"
        rbcAutoType(16).Left = ilStartPos
        'rbcAutoType(16).Top = 45
        ilStartPos = ilStartPos + rbcAutoType(16).Width + 60
        rbcAutoType(16).Top = ilStartRow
        ilHowMany = ilHowMany + 1
        ilCols = ilCols + 1
        If ilCols > 4 Then
            ilRows = ilRows + 1
            ilCols = 1
            ilStartRow = ilStartRow + 240
            ilStartPos = 780
        End If
        If ilHowMany = 1 Then
            rbcAutoType(16).Value = True
        End If
        rbcAutoType(16).Visible = True
        sgBrowseMaskFile = "*.rec"
    End If
    
    'RCS-4
    If (ilAutomationType And RCS4DIGITCART) = RCS4DIGITCART Then
        rbcAutoType(4).Left = ilStartPos
        ilStartPos = ilStartPos + rbcAutoType(4).Width + 60
        rbcAutoType(4).Top = ilStartRow
        ilHowMany = ilHowMany + 1
        ilCols = ilCols + 1
        If ilCols > 4 Then
            ilRows = ilRows + 1
            ilCols = 1
            ilStartRow = ilStartRow + 240
            ilStartPos = 780
        End If
        If ilHowMany = 1 Then
            rbcAutoType(4).Value = True 'RCS-4
        End If
        rbcAutoType(4).Visible = True
        sgBrowseMaskFile = "*.rec"          '10-23-06
    End If
    
    'RCS-5
    If (ilAutomationType2 And RCS5DIGITCART) = RCS5DIGITCART Then
        rbcAutoType(11).Left = ilStartPos
        ilStartPos = ilStartPos + rbcAutoType(11).Width + 60
        rbcAutoType(11).Top = ilStartRow
        ilHowMany = ilHowMany + 1
        ilCols = ilCols + 1
        If ilCols > 4 Then
            ilRows = ilRows + 1
            ilCols = 1
            ilStartRow = ilStartRow + 240
            ilStartPos = 780
        End If
        If ilHowMany = 1 Then
            rbcAutoType(11).Value = True
        End If
        rbcAutoType(11).Visible = True
        sgBrowseMaskFile = "*.rec"          '10-23-06
    End If
    
    'Rivendell
    If (Asc(tgSpf.sUsingFeatures8) And RIVENDELLEXPORT) = RIVENDELLEXPORT Then
        rbcAutoType(12).Left = ilStartPos
        ilStartPos = ilStartPos + rbcAutoType(12).Width + 60
        rbcAutoType(12).Top = ilStartRow
        ilHowMany = ilHowMany + 1
        ilCols = ilCols + 1
        If ilCols > 4 Then
            ilRows = ilRows + 1
            ilCols = 1
            ilStartRow = ilStartRow + 240
            ilStartPos = 780
        End If
        If ilHowMany = 1 Then
            rbcAutoType(12).Value = True
        End If
        rbcAutoType(12).Visible = True
        sgBrowseMaskFile = "*.cpo"          '10-23-06
    End If
    
    'Scott doesnt have an import
    '    If (ilAutomationType And SCOTT) = SCOTT Then      '2-5-03
    '        'plcScreen.Caption = "Import RCS"
    '        lblPlcScreen.Caption = "Import Scott"
    '        rbcAutoType(5).Left = ilStartPos
    '        ilStartPos = ilStartPos + rbcAutoType(5).Width + 60
    '        ilHowMany = ilHowMany + 1
    '        ilCols = ilCols + 1
    '        If ilCols > 4 Then
    '            ilRows =ilRows +1
    '            ilCols = 1
    '            ilStartRow = ilStartRow + 240
    '            ilStartPos=780
    '        End If
    '        If ilHowMany = 1 Then
    '            rbcAutoType(5).Value = True
    '        End If
    '        rbcAutoType(5).Visible = True
    '        rbcAutoType(5).Enabled = False  'no import interface yet
    '        sgBrowseMaskFile = "*.rec"          '10-23-06
    '    End If
    
    'Wide Orbit
    If (ilAutomationType3 And WIDEORBIT) = WIDEORBIT Then      '1-10-12
        rbcAutoType(13).Left = ilStartPos
        ilStartPos = ilStartPos + rbcAutoType(13).Width + 60    '6-27-12
        rbcAutoType(13).Top = ilStartRow
        ilHowMany = ilHowMany + 1
        ilCols = ilCols + 1
        If ilCols > 4 Then
            ilRows = ilRows + 1
            ilCols = 1
            ilStartRow = ilStartRow + 240
            ilStartPos = 780
        End If
        If ilHowMany = 1 Then
            rbcAutoType(13).Value = True
        End If
        rbcAutoType(13).Visible = True
        sgBrowseMaskFile = "*.air"          '1-10-12
    End If
    
    'WireReady
    If (ilAutomationType2 And WIREREADY) = WIREREADY Then      '9-11-06
        rbcAutoType(9).Left = ilStartPos
        ilStartPos = ilStartPos + rbcAutoType(9).Width + 60
        rbcAutoType(9).Top = ilStartRow
        ilHowMany = ilHowMany + 1
        ilCols = ilCols + 1
        If ilCols > 4 Then
            ilRows = ilRows + 1
            ilCols = 1
            ilStartRow = ilStartRow + 240
            ilStartPos = 780
        End If
        If ilHowMany = 1 Then
            rbcAutoType(9).Value = True 'WireReady
        End If
        rbcAutoType(9).Visible = True
        sgBrowseMaskFile = "*.rec"          '10-23-06
    End If
    
    'Zetta
    If (ilAutomationType3 And ZETTA) = ZETTA Then      '1-7-16
        rbcAutoType(15).Left = ilStartPos
        ilStartPos = ilStartPos + rbcAutoType(15).Width + 60
        rbcAutoType(15).Top = ilStartRow
        ilHowMany = ilHowMany + 1
        ilCols = ilCols + 1
        If ilCols > 4 Then
            ilRows = ilRows + 1
            ilCols = 1
            ilStartRow = ilStartRow + 240
            ilStartPos = 780
        End If
        If ilHowMany = 1 Then
            rbcAutoType(15).Value = True
        End If
        rbcAutoType(15).Visible = True
        sgBrowseMaskFile = "*.rec"          '1-7-16
    End If
    
    'Dynamically make room for multiple columns of Import buttons
    Select Case ilRows
        Case 1: frcImportOptions.height = 495
        Case 2: frcImportOptions.height = 735
        Case 3: frcImportOptions.height = 975
        Case 4: frcImportOptions.height = 1215
        Case Else: frcImportOptions.height = 1455
    End Select
    
    'Move the rest of the controls down if needed
    plcControls.Top = frcImportOptions.Top + frcImportOptions.height + 120
    
    'Make form taller if needed
    Me.height = plcControls.Top + plcControls.height + 120
    
    If ilHowMany = 0 Then       'at least one import type must be defined
        MsgBox "No Automation type defined or Import disallowed ", vbOKOnly + vbCritical + vbApplicationModal, "mInit"
        imTerminate = True
        Exit Sub
    End If
    
    lblPlcScreen.Caption = "Import"
    lblPlcScreen.Visible = True
        
    'Determine how spots should be set if makegood/outside
    If tgSpf.sPLMove = "M" Then
        rbcPLMove(0).Value = True
        plcPLMove.Visible = False
    ElseIf tgSpf.sPLMove = "O" Then
        rbcPLMove(1).Value = True
        plcPLMove.Visible = False
    Else
        plcPLMove.Visible = True
    End If

    gCenterStdAlone ImptGen
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    mSetCommand
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*      <INPUT> slAutoDesc - Dalet, Prophet, Scott, Scott
'                            or Drake
'*******************************************************
Private Function mOpenMsgFile(slAutoDesc As String)
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    'On Error GoTo mOpenMsgFileErr:
    'slToFile = sgImportPath & "ImptGen.Txt"
    slToFile = sgDBPath & "Messages\" & "ImptGen.Txt"

    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        If gDateValue(slFileDate) = lmNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Output As hmMsg
            ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    '
    '***  Test for Automation type to show message here  ***
    '
    '
    Print #hmMsg, "** Import " & slAutoDesc & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"

    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function
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

    slCommand = sgCommandStr  'Command$
'    If StrComp(slCommand, "Debug", 1) = 0 Then
'        igStdAloneMode = True 'Switch from/to stand alone mode
'        sgCallAppName = ""
'        slStr = "Guide"
'        ilTestSystem = False
'        imShowHelpMsg = False
'    Else
'        igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            'End
            imTerminate = True
            Exit Sub
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
'        imShowHelpMsg = True
'        ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
'        If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
'            imShowHelpMsg = False
'        End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone ImptGen, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igCmmCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igCmmCallSource = CALLNONE
    'End If
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
Private Function mReadFile(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    ilRet = 0
    'On Error GoTo mReadFileErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        mReadFile = False
        Exit Function
    End If
    '6-27-12 Change all occurences to read into smTempFieldVAlues array instead of smFieldValues
    'in mSaveRec they will be transferred into smFieldValues.  Wide Orbit and Jelli have their vehicle station codes
    'stored within the record and are using log vehicles
    
    ReDim smTempFieldValues(0 To 0) As String
    err.Clear
    Do
        ilRet = 0
        'On Error GoTo mReadFileErr:
        If EOF(hmFrom) Then
            Exit Do
        End If
        Line Input #hmFrom, slLine
        On Error GoTo 0
        ilRet = err.Number
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Close hmFrom
            mReadFile = False
            Exit Function
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                smTempFieldValues(UBound(smTempFieldValues)) = slLine
                ReDim Preserve smTempFieldValues(0 To UBound(smTempFieldValues) + 1) As String
            End If
        End If
    Loop Until ilEof
    Close hmFrom
    If ilRet <> 0 Then
        mReadFile = False
    Else
        mReadFile = True
    End If
    MousePointer = vbDefault
    Exit Function
'mReadFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function
'***************************************************************************************
'*
'*      Procedure Name:mSaveRec
'*
'       <input> slDate - import date
'               ilVefCode - vehicle internal code
'               ilautomationType - which import (prophet, rivendell, wide orbit, etc)
'               ilCombineVefCode
'               slVehicleStationCode - 5 character station code from vehicle
'
'       Array of smFieldValues contains 1 record per entry
'
'*             Created:6/29/93       By:D. LeVine
'*            Modified:              By:
'*
'*            Comments:Update or added record
'*      9-11-06 Add WireReady import
'
'   RCS-4 rturn record format:
'     Column   length   Description
'       1        8      Actual Time aired:  Military HH:mm:ss
'       9        2      Blank
'       11       14     Actual Length MM:ss
'       15       17     blank
'       18       1      edit code
'       19       1      Blank
'       20       21     Discrepancy code
'       22       1      Blank
'       Export record duplicated
'       23       1      Always "C"
'       24       7      Time sent to be aired HHMM:SS (no : between HH & SS)
'       31       4      Cart #
'       35       24     Advt/prod name
'       59       3      blank
'       62       4      spot length MM:SS
'       66       4      blank
'       70       6      first part of internal spot code (sdf.lcode)
'       76       2      2nd part of internal spot code (sdf.lcode)
'       78       12     blank
'       90        1     c/r
'       91        1     l/f
'
'   RCS-5 rturn record format:
'     Column   length   Description
'       1        8      Actual Time aired:  Military HH:mm:ss
'       9        2      Blank
'       11       14     Actual Length MM:ss
'       15       17     blank
'       18       1      edit code
'       19       1      Blank
'       20       21     Discrepancy code
'       22       1      Blank
'       Export record duplicated
'       23       1      Always "C"
'       24       7      Time sent to be aired HHMM:SS (no : between HH & SS)
'       31       15      Cart #
'       46       24     Advt/prod name
'       60       3      blank
'       73       4      spot length MM:SS
'       77       4      blank
'       81       6      first part of internal spot code (sdf.lcode)
'       87       2      2nd part of internal spot code (sdf.lcode)
'       89       12     blank
'       90  More blanks
'                 1     c/r
'                 1     l/f
'
'   Prophet & Dalet return record format:
'
'   Prophet return  (1-6-03):
'          Column      length   Description
'      Start    end
'       1        8       8      Date Aired:  MM/DD/YY
'       9        9       1      Blank
'       10       17      8      Actual Time aired:  Military HH:mm:ss
'       18       18      1      Blank
'       19       63      45     Event Description Advt Name (30)/Prod(15)
'       64       64      1      Blank
'       65       72      8      Length HH:MM:SS
'       73       73      1      Blank
'       74       82      9      Cart #/ Cut #
'       83       83      1      Blank
'       84       93      10     Counterpoint Internal Spot Code
'       94       94      1      Blank
'       95       96      2      Prophet event status (01= spot played, all others did not air)
'       97       97      1      C/R
'       98       98      1      Line Feed
'
'
'   Prophet return  (MediaStar 11-4-03):
'          Column      length   Description
'      Start    end
'       1        8       8      Date Aired:  MM/DD/YY
'       9        16      8      Actual Time aired:  Military HH:mm:ss
'       17       45      29     Event Description Advt Name/Prod for 29 characters
'       46       46      1      *
'       62       66      5      Cart # without media code
'       67       68      2      Cut #
'       69       72      2      Prophet event status (01= spot played, all others did not air)
'                        1      C/R
'                        1      Line Feed
'
'   WireReady Return 9/11/06:
'   RadioMan return : 5-1-20
'
'      Start    end
'       1        8       8      Sch Time aired HH:MM:SS in military tie
'       9        28      20     Cart # (w/material code) or ISCI (info only)
'       29       75      46     Event Description Advt Name (max 30)/Prod (info only)
'       75       94      20     Creative Title (information only)
'       95       98      4      spot length MMSS
'       99       108     10     internal spot code
'       109      116     8      Date Aired XX/xx/xx
'       117      124     8      Time Aired (HH:MM:SS) military time
'       125      125     1      Discrep code 1 = aired, 2 thru N = not aired.
'       126      126     1      C/R
'       127      127     1      Line Feed
'
'   Rivendell
'          Column      length   Description
'      Start    end
'       1       8      8        Scheduled event start time HH:MM:SS
'      10       17     8        Actual event start time HH:MM:SS
'      19       26     8        Schedule event Length HH:MM:SS
'      28       35     8        Actual event playout length HH:MM:SS
'      37       42     6        Cart Number
'      44       77     34       Cart Title
'      79       110    32       GUID  sdfCode:vefcode:vehicle name
'
'
'   Wide Orbit aired log (1-10-12), Comma delimited
'   Field 1     Format          Description
'   Date        YY-MM-DD        Date of airing
'   Air Time    HH:MM:SS        Actual air time
'   Event Type  "on-air"        anything but "on-air" will be missed
'               "JUMPBY"
'               "SHORT:"
'                               if "on-air" the next 5 fields are:
'   Copy        DA+xxxx         Cart #, reference only
'   Category                    Unused, reference only
'   Title       50char          Title (25 char), Partner ID (25 char)- this is the internal spot code preceded by a ":", then the Station code (for reference) preceded by a ":"
'                               i.e. Tops Baseball cards          :2500:MNS
'   Spot Length MM:SS           length of spot. spotlength compared, if different spot will be missed
'   sched time  HH:MM:SS        unused, reference only

'   Audio Vault import format:
'   Parsed values:
'   field 1         ! for comment code, used to test and bypass event
'   field 2         error code:  W = warning, e = error, s = stopped early.  Blank assumes OK, otherwise not aired
'   field 3         n/a
'   field 4         air time (hh:mm:ss military time)
'   field 5         duration
'   field 6         align code (:), empty code ("), jump code (J) , all these records should be ignored
'   field 7         n/a description of cut or cart.
'   field 8         n/a
'   field 9         description of advt name &prod (30 char) last 10 char is the internal spot code, right justified
'   field 10        n/a
'   field 11        n/a
'   field 12        n/a
'   field 13        n/a
'   field 14        n/a

'6-27-12 Jelli Comma delimited
'   Field Name       Format          Max Length
'   ISCI Code        AN              20
'   Start Date/Time  YYYY-mm-dd hh:mm:ss    military 19
'   Length           nnn             3
'   Creative Title   AN              30
'   Veh STation Code AN              5
'   contract         N               9
'   Advt Name        AN              30
'   Agency Name      AN              40
'   Internal Code    nnn             10
'   Status code      n               1 (0 missed, 1=aired)
'
'
'        ilAutomationType = 1           'Dalet
'        ilAutomationType = 2           'prophet nexgen (import)
'        ilAutomationType = 3           'scott
'        ilAutomationType = 4           'drake
'        ilAutomationType = 5           'rcs-4  (import)
'        ilAutomationType = 6           'prophet wizard (import)
'        ilAutomationType = 7           'prophet mediastar  (import)
'        ilAutomationType = 9           'audio vault (import)
'        ilAutomationType = 10          'wire ready  (import)
'        ilAutomationType = 11          'rcs-5 (import)
'        ilAutomationType = 12          'Rivendell (import)
'        ilAutomationType =  13         'Wide Orbit (import) 1-10-12
'        ilAutomationType = 14          'Jelli 6-27-12
'        ilAutomationType = 15          'Zetta 1-7-16
'        ilAutomationType = 16          'Radio Man 5-1-20
'************************************************************************************
Private Function mSaveRec(slDate As String, ilVefCode As Integer, ilAutomationType As Integer, ilCombineVefCode As Integer, slVehicleStationCode As String)
    Dim ilRet As Integer
    Dim slStr As String
    Dim llSdfCode As Long
    Dim ilECode As Integer
    Dim ilDCode As Integer
    Dim llSsfMemDate As Long
    Dim llSsfRecPos As Long
    Dim slTime As String
    Dim ilLoop As Integer
    ReDim ilSchDate(0 To 1) As Integer
    Dim ilChgSpot As Integer    '0=Ok; 1= Update Sdf failed; 2= Missed to Sch disallowed
    Dim ilFoundSDF As Integer
    Dim llTime As Long
    Dim slSchTime As String
    Dim llAirTime As Long
    Dim slMOrO As String * 1
    Dim llDate As Long
    Dim slSchDate As String
    Dim ilSdfLoop As Integer
    Dim slAdvt As String
    Dim ilStatus As Integer
    Dim slSpotType As String * 1
    Dim ilAdvt As Integer
    Dim ilFoundLL As Integer        '5-22-06
    Dim ilLoopLL As Integer
    Dim slXMid As String * 1         '3-28-07 flag (Y/N) to indicate if spot has crossed midnight from import spot
    Dim llAiredSpotLen As Long      '10-2-12, was integer
    Dim slStationRef As String
    Dim ilUpper  As Integer
    Dim ilMatchOnGame As Boolean
    
    ReDim tmReconVehicles(0 To 0) As RECONVEHICLES  'retain list of vehicles/games that have been imported
    gPackDate slDate, ilSchDate(0), ilSchDate(1)
    llDate = gDateValue(slDate)
    
    ReDim tmImpSdf(0 To 0) As SDF
    ilMatchOnGame = False
     
    mBuildSDFArray llDate, ilVefCode
    
'    ilFoundSDF = gObtainSDFbyVehDate(ImptGen, hmSdf, ilSchDate(), ilVefCode, tmImpSdf())
'    If Not (ilFoundSDF) Then
'        mSaveRec = False
'        Screen.MousePointer = vbDefault    'Default
'        ilRet = MsgBox(slDate & " & vehicle code " & Str$(ilVefCode) & " not found", vbOkOnly + vbExclamation, "Import ")
'        Exit Function
'    End If
    
    mGetMergedGameSpots slDate, ilVefCode
    ilRet = btrBeginTrans(hmSdf, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrAbortTrans(hmSdf)
        Screen.MousePointer = vbDefault    'Default
        ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Import ")
        mSaveRec = False
        Exit Function
    End If
    llSsfMemDate = 0
    
    
    If ilAutomationType = AUTOIMPORT_JELLI Or ilAutomationType = 13 Then        'file doesnt contain a vehicle station code, process for log vehicles
        'Gather all the matching spots by station code within the import file for the vehicle that is being processed.  Wide orbit and Jelli have the
        'vehicle station codes stored within the import file.  Build them in a temporary array, then copy it back to the array that processes them
        '
        ilUpper = 0
 
        ReDim smFieldValues(0 To 0) As String
        'Wide Orbit and Jelli are using log vehicles and its vehicle station code is defined within the record.  Filter out only the
        'matching vehicle station codes of the vehicle that it is processing
        For ilLoop = LBound(smTempFieldValues) To UBound(smTempFieldValues) - 1
             If ilAutomationType = 13 Then                          'wide orbit
                gParseDelimitedFields ",", smTempFieldValues(ilLoop), False, smParsedValues()   'parse out all 14 fields in record 6-27-12 chg to get from col 1, not 2
                mAgeParsedValues
                slStationRef = Mid$(smParsedValues(6), 1, 50)
                ilRet = gParseItem(slStationRef, 2, ":", slStr)
                    If Trim$(slStr) = Trim$(slVehicleStationCode) Then      'ignore the records that do not have a matching station code
                        smFieldValues(ilUpper) = smTempFieldValues(ilLoop)
                        ilUpper = ilUpper + 1
                        ReDim Preserve smFieldValues(0 To ilUpper) As String
                    End If
             ElseIf ilAutomationType = AUTOIMPORT_JELLI Then
                gParseDelimitedFields ",", smTempFieldValues(ilLoop), False, smParsedValues()   'parse out all 14 fields in record
                mAgeParsedValues
                If Trim$(smParsedValues(5)) = Trim$(slVehicleStationCode) Then
                    smFieldValues(ilUpper) = smTempFieldValues(ilLoop)
                    ilUpper = ilUpper + 1
                    ReDim Preserve smFieldValues(0 To ilUpper) As String
                End If
             End If
        Next ilLoop
    Else            'not wide orbit or jelli, copy all events from temp into processing array
        ReDim smFieldValues(0 To 0) As String
        For ilLoop = 0 To UBound(smTempFieldValues) - 1
            smFieldValues(ilLoop) = smTempFieldValues(ilLoop)
            ReDim Preserve smFieldValues(0 To ilLoop + 1) As String
        Next ilLoop
    End If
    
    For ilLoop = LBound(smFieldValues) To UBound(smFieldValues) - 1 Step 1
        '3-6-03 implement import of Prophet Wizard the same as Prophet NexGen.  Exports are slight different
        '9-11-06 WireReady added (ilAutomationType = 10)
        'ilAutomationType (Imports): 1=Dalet, 2 = prophet nexgen,3 =scott,4 = drake,5=rcs-4, 6 =prophet wizard,7=prophet mediastar
        '                       8 = unused,9 =audio vault,10 =wire ready,11=rcs-5,12 =Rivendell,13 =Wide Orbit ,14 =Jelli, 15=Zetta, 16 = RaidoMan
        If ilAutomationType = 2 Or ilAutomationType = 6 Or ilAutomationType = 7 Or ilAutomationType = 9 Or ilAutomationType = 10 Or ilAutomationType = 12 Or ilAutomationType = 13 Or ilAutomationType = AUTOIMPORT_JELLI Or ilAutomationType = AUTOIMPORT_ZETTA Or ilAutomationType = AUTOIMPORT_RADIOMAN Then          'prophet (nexGen or Wizard, MediaStar, Audio Vault, 1-10-12 Wide Orbit
        'The import record is read one line at a time and if the air status is a "01", the time will be updated.
        'If the aired status is other than a 01, it will be set to missed
        'All spots not processed from the table of vehicles that are scheduled (and not in import file), are flagged as missed.
        '
            If ilAutomationType = 7 Then        'media star
                'Original time shifted 5 characters right, copy area exanded by 5
                llAirTime = mConvertTime(smFieldValues(ilLoop), 9, 12, 15, slXMid)
                slTime = gFormatTimeLong(llAirTime, "A", "1")                'Air Time of spot returned
                slAdvt = Mid$(smFieldValues(ilLoop), 17, 29)            'get the advert/prod desc
                slStr = Mid$(smFieldValues(ilLoop), 46, 1)
                If slStr <> "*" Then
                    llSdfCode = -1
                Else
                    llSdfCode = mConvertLong(smFieldValues(ilLoop), 71, 10) 'convert SDF Code and read SDF
                End If
                ilStatus = Val(Mid$(smFieldValues(ilLoop), 69, 2))     'status if aired (01 aired, all others missed)
            ElseIf ilAutomationType = 9 Then            'audio vault: vertical bar delimited record.  need to parse out fields
                gParseDelimitedFields "|", Mid$(smFieldValues(ilLoop), 2), False, smParsedValues()   'parse out all 14 fields in record
                mAgeParsedValues
                If Trim$(smParsedValues(1)) = "!" Or Trim$(smParsedValues(6)) = ":" Or Trim$(smParsedValues(6)) = "J" Or Trim$(smParsedValues(6)) = """" Or Trim$(smParsedValues(9)) = "" Then      'should this be bypass because its a comment, an align, empty or jump record
                    'ignore this record:  comment or event other than a spot
                    ilStatus = 0
                    llAiredSpotLen = 0
                    llSdfCode = 0
                    llAirTime = 0
                    slAdvt = "Ignore Event"
                Else
                    llAirTime = mConvertTime(smParsedValues(4), 1, 4, 7, slXMid)
                    slTime = gFormatTimeLong(llAirTime, "A", "1")
                    slAdvt = Mid$(smParsedValues(9), 1, 30)       'advt prod name for descriptions
                    llSdfCode = mConvertLong(smParsedValues(9), 31, 10) 'convert SDF Code and read SDF
                    slStr = Trim$(smParsedValues(1))                'status
                    If slStr = "" And llSdfCode > 0 Then
                        ilStatus = 1                                'assume aired
                    Else
                        ilStatus = 0
                    End If
                    '3-22-12 do not add 00: prior to the spot lenght ; it is already hr:mm:ss.  code was erroneously copied from Wide orbit since its time length input is mm:ss
                    slStr = smParsedValues(5)   '"00:" & smParsedValues(5)           '1-10-12 if spot lengths do not match, miss the spot
                    'ilAiredSpotLen = gLengthToLong(slStr)      '10-2-12        use long to be able to ignore illegal records
                    llAiredSpotLen = gLengthToLong(slStr)       '10-2-12
                End If
            ElseIf ilAutomationType = 10 Or ilAutomationType = AUTOIMPORT_RADIOMAN Then                       'WireReady, Radioman added 5-1-20
                llAirTime = mConvertTime(smFieldValues(ilLoop), 117, 120, 123, slXMid)
                slTime = gFormatTimeLong(llAirTime, "A", "1")                'Air Time of spot returned
                slAdvt = Mid$(smFieldValues(ilLoop), 29, 45)           'get the advert/prod desc
                llSdfCode = mConvertLong(smFieldValues(ilLoop), 99, 10) 'convert SDF Code and read SDF
                ilStatus = Val(Mid$(smFieldValues(ilLoop), 125, 1))     'status if aired (01 aired, all others missed)
            ElseIf ilAutomationType = 12 Then        'Rivendell
                'Original time shifted 5 characters right, copy area exanded by 5
                llAirTime = mConvertTime(smFieldValues(ilLoop), 10, 13, 16, slXMid)
                slTime = gFormatTimeLong(llAirTime, "A", "1")                'Air Time of spot returned
                slAdvt = Mid$(smFieldValues(ilLoop), 44, 34)            'get the cart title desc
                gParseDelimitedFields ":", Mid$(smFieldValues(ilLoop), 79), False, smParsedValues()   'parse out 3 fields in record
                mAgeParsedValues
                llSdfCode = Val(smParsedValues(1)) 'convert SDF Code and read SDF
                If Val(smParsedValues(2)) = ilVefCode Then
                    ilStatus = 1
                Else
                    llSdfCode = -1
                End If
            ElseIf ilAutomationType = 13 Then                   '1-10-12 wide orbit
                llAirTime = mConvertTime(smFieldValues(ilLoop), 10, 13, 16, slXMid)
                slTime = gFormatTimeLong(llAirTime, "A", "1")                'Air Time of spot returned
                gParseDelimitedFields ",", Mid$(smFieldValues(ilLoop), 2), False, smParsedValues()   'parse out all 14 fields in record
                mAgeParsedValues
                If smParsedValues(3) = "on-air" Then        'anything other than "on-air" didnt air
                    ilStatus = 1            'aired status
                Else
                    ilStatus = 0
                End If
                slStationRef = Mid$(smParsedValues(6), 26, 25)
                ilRet = gParseItem(slStationRef, 1, ":", slStr)
                
                llSdfCode = -1
                If Trim$(slStr) <> "" Then
                    llSdfCode = Val(slStr)
                End If
              
                slStr = "00:" & smParsedValues(7)
                'ilAiredSpotLen = gLengthToLong(slStr)          '10-2-12 change to use long for records with illegal formatting
                llAiredSpotLen = gLengthToLong(slStr)
           ElseIf ilAutomationType = AUTOIMPORT_JELLI Then     '6-27-12
                gParseDelimitedFields ",", smFieldValues(ilLoop), False, smParsedValues()   'parse out all 14 fields in record
                mAgeParsedValues
                llAirTime = mConvertTime(smParsedValues(2), 12, 15, 18, slXMid)
                slTime = gFormatTimeLong(llAirTime, "A", "1")                'Air Time of spot returned
                llSdfCode = mConvertLong(smParsedValues(9), 1, 10)   'convert SDF Code and read SDF
                slAdvt = Trim$(smParsedValues(4))              'get the advert/prod desc
                'ilAiredSpotLen = Val(smParsedValues(3))
                llAiredSpotLen = Val(smParsedValues(3))
                ilStatus = Val(smParsedValues(10))      'status if aired (01 aired, all others missed)
'                If Trim$(smParsedValues(5)) <> Trim$(slVehicleStationCode) Then
'                    llSdfCode = 0               'only process the vehicle that matches since we are processing all vehicles within the log vehicle, and they are all intermixed
'                End If
            ElseIf ilAutomationType = AUTOIMPORT_ZETTA Then         '1-7-16
                llAirTime = mConvertTime(smFieldValues(ilLoop), 1, 4, 7, slXMid)
                slTime = gFormatTimeLong(llAirTime, "A", "1")                'Air Time of spot returned
                slAdvt = Mid$(smFieldValues(ilLoop), 103, 24)           'get the advert desc
                llSdfCode = mConvertLong(smFieldValues(ilLoop), 83, 10) 'convert SDF Code and read SDF
                ilStatus = Val(Mid$(smFieldValues(ilLoop), 21, 3))     'status if aired (0 aired, all others missed)
                If ilStatus = 0 And llSdfCode > 0 Then
                    ilStatus = 1                                'assume aired
                Else
                    ilStatus = 0
                End If
            Else                            'prophet nextgen or prophet wizard
                llAirTime = mConvertTime(smFieldValues(ilLoop), 10, 13, 16, slXMid)
                slTime = gFormatTimeLong(llAirTime, "A", "1")                'Air Time of spot returned
                llSdfCode = mConvertLong(smFieldValues(ilLoop), 84, 10)   'convert SDF Code and read SDF
                slAdvt = Mid$(smFieldValues(ilLoop), 19, 45)              'get the advert/prod desc
                ilStatus = Val(Mid$(smFieldValues(ilLoop), 95, 2))       'status if aired (01 aired, all others missed)
            End If
            'if not media star or its media star with an sdf code, continue
                
            If ((ilAutomationType <> 7) And (ilAutomationType <> 12)) Or ((ilAutomationType = 7 Or ilAutomationType = 12 Or ilAutomationType = AUTOIMPORT_JELLI) And (llSdfCode <> -1)) Then
            '1-10-12 Prophet Nextgen, Prophet Wizard, Audio Vault, Wire Ready or RadioMan, Prophet MediaStart if valid sdf code, or Rivendell and valid sdfcode goes thru this
                'ilFoundSDF = mFindSDF(llSdfCode)   'convert SDF Code and read SDF
                If llSdfCode > 0 Then               'dont even try to read if no code exists from import
                    ilFoundSDF = mFindSDF(llSdfCode, slXMid)   'read SDF
                Else
                    ilFoundSDF = False                  'spot doesnt exist
                End If
                If ilFoundSDF Then
                    If (ilCombineVefCode <= 0) Or ((ilCombineVefCode > 0) And (tmSdf.iVefCode = ilVefCode)) Then
                        '5-22-06 if this spot coming in from import is within the live log times, bypass it
                        ilFoundLL = False
                        For ilLoopLL = 0 To UBound(tmIgnoreTimes) - 1
                            If llAirTime >= tmIgnoreTimes(ilLoopLL).lStartTime And llAirTime <= tmIgnoreTimes(ilLoopLL).lEndTime Then   'import spot falls within a live log time
                                ilFoundLL = True
                                Exit For
                            End If
                        Next ilLoopLL
                        '5-22-06 Cant be a spot that live log has already posted (testing the live log times)

                        If Not ilFoundLL Then           'didnt find an ignore time, continue to process the incoming spot
                        '1-3-06 found the spot, if it already has been changed in some way, ignore it as it has already been posted
                        'Can never post twice.
                            If ((tmSdf.sAffChg <> "Y" And tmSdf.sAffChg <> "A" And tmSdf.sAffChg <> "B" And tmSdf.sAffChg <> "C" And tmSdf.sAffChg <> "T") And (tmSdf.sSchStatus <> "M")) Then       'OK to post this spot if never changed or its not flagged as missed
                                
                                '1-10-12 if non-matching import spot lengths for Wide Orbit(13) or AudioVault(9), bump the spot
                                '6-27-12 Jelli
                                '3-21-12 allow 3"variance of spot length difference
                                'If (ilAutomationType = 13 Or ilAutomationType = 9 Or ilAutomationType = AUTOIMPORT_JELLI) And ((tmSdf.iLen - ilAiredSpotLen) > 3 Or (ilAiredSpotLen - tmSdf.iLen) > 3) Then    '(ilAiredSpotLen <> tmSdf.iLen) Then
                                'ilAutomationType (Imports): 1=Dalet, 2 = prophet nexgen,3 =scott,4 = drake,5=rcs-4, 6 =prophet wizard,7=prophet mediastar
                                '                       8 = unused,9 =audio vault,10 =wire ready,11=rcs-5,12 =Rivendell,13 =Wide Orbit ,14 =Jelli, 15=Zetta, 16 = RadidoMan
                                If (ilAutomationType = 13 Or ilAutomationType = 9 Or ilAutomationType = AUTOIMPORT_JELLI) And ((tmSdf.iLen - llAiredSpotLen) > 3 Or (llAiredSpotLen - tmSdf.iLen) > 3) Then
                                    ilStatus = 0
                                End If
                                If ilStatus = 1 Then        'aired status
                                    'gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                                    slSchTime = gFormatTimeLong(lmSpotTime, "A", "1")
                                    'gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                                    slSchDate = Format$(lmSpotDate, "m/d/yy")
                                    tmSdf.sXCrossMidnight = slXMid      '3-28-07

                                    If (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
                                        If rbcPLMove(0).Value Then
                                            slMOrO = "G"
                                        Else
                                            slMOrO = "O"
                                        End If
                                        ilChgSpot = gProcessSpot(hmSsf, hmSdf, hmSmf, hmCHF, hmClf, hmCff, hmRdf, hmCrf, hmCgf, tmSdf, ilVefCode, slSchDate, slSchTime, slDate, slTime, slMOrO, hmSxf, hmGsf, hmGhf)
                                    Else
                                        tmSdfSrchKey3.lCode = tmSdf.lCode
                                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, 3, BTRV_LOCK_NONE, SETFORWRITE)  'position record so update works
                                        If ilRet = BTRV_ERR_NONE Then
                                            gPackDate slDate, tmSdf.iDate(0), tmSdf.iDate(1)
                                            gPackTime slTime, tmSdf.iTime(0), tmSdf.iTime(1)
                                            tmSdf.sAffChg = "T"     '5-27-05  indicate a time change on spt
                                            tmSdf.sXCrossMidnight = slXMid      '3-28-07

                                            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                                            If ilRet <> BTRV_ERR_NONE Then
                                                ilChgSpot = 1
                                            Else
                                                ilChgSpot = 0    '5-27-05  no error, spot was updated
                                            End If
                                        Else
                                            ilChgSpot = 1
                                        End If
                                    End If
                                    If ilChgSpot = 1 Then
                                        '6-10-05 continue with remainder of file
                                        'ilRet = btrAbortTrans(hmSdf)
                                        lbcMsg.AddItem "Call Counterpoint & Run checks: Cannot find contract or SMF for " & smFieldValues(ilLoop)
                                        Print #hmMsg, "Call Counterpoint & Run checks: Cannot find contract or SMF for " & smFieldValues(ilLoop)

                                        'lbcMsg.AddItem "Error in Processing, Cancelled" & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                                        'Print #hmMsg, "Error in Processing, Cancelled at " & Format$(Now, "h:mm:ssAM/PM") & ", " & smFieldValues(ilLoop)
                                        'Close #hmMsg
                                        'Screen.MousePointer = vbDefault
                                        'Exit Function
                                    ElseIf ilChgSpot = 2 Then
                                        lbcMsg.AddItem "Missed To Scheduled Disallowed " & smFieldValues(ilLoop)
                                        Print #hmMsg, "Missed To Scheduled Disallowed " & smFieldValues(ilLoop)
                                    ElseIf ilChgSpot = 3 Then
                                        lbcMsg.AddItem "Date change Disallowed " & smFieldValues(ilLoop)
                                        Print #hmMsg, "Date change Disallowed " & smFieldValues(ilLoop)
                                    Else
                                        mSaveVefProc            'keep track of vehicles/games processed
                                    End If
                                Else
                                    'see if spot already Missed
                                    If tmSdf.sSchStatus <> "M" Then
                                        slSpotType = tmSdf.sSpotType
                                        If (slSpotType <> "O") And (slSpotType <> "C") Then
                                            tmSdf.sXCrossMidnight = slXMid      '3-28-07

                                            ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, tmSdf.iGameNo, tmSmf, hmSsf, tmSsf, llSsfMemDate, llSsfRecPos, hmSxf, hmGsf, hmGhf)
                                            If Not ilRet Then
                                                lbcMsg.AddItem "Spot Not Set to Missed for Internal code " & str$(llSdfCode)
                                                Print #hmMsg, "Spot Not Set to Missed for: " & smFieldValues(ilLoop)
                                            Else                'successful update of spot
                                                mSaveVefProc            'keep track of vehicles/games processed
                                                If slSpotType = "X" Then           'if fill spot, it has been removed.  Note that in log
                                                    lbcMsg.AddItem "Fill Spot Missed and Removed for " & Trim(slAdvt) & " on " & Trim(slDate) & " at " & Trim(slTime)
                                                    Print #hmMsg, "Fill Spot Missed and Removed on " & Trim(slDate) & " at " & Trim(slTime) & ",  " & smFieldValues(ilLoop)
                                                End If
                                            End If
                                        Else
                                            tmSdfSrchKey3.lCode = tmSdf.lCode
                                            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, 3, BTRV_LOCK_NONE, SETFORWRITE)  'position record so update works
                                            If ilRet = BTRV_ERR_NONE Then
                                                ilRet = btrDelete(hmSdf)
                                                lbcMsg.AddItem "Open/Close Spot Missed and Removed for " & Trim(slAdvt) & " on " & Trim(slDate) & " at " & Trim(slTime)
                                                Print #hmMsg, "Open/Close Spot Missed and Removed on " & Trim(slDate) & " at " & Trim(slTime) & ",  " & smFieldValues(ilLoop)
                                                mSaveVefProc            'keep track of vehicles/games processed
                                            End If
                                        End If
                                    End If
                                End If
                            End If                      'sdfaffchg <> any change and sdfschstatus <> "M"
                            'Found the matching spot from the import file, mark it as processed
                            For ilSdfLoop = LBound(tmImpSdf) To UBound(tmImpSdf) - 1
                                If llSdfCode = tmImpSdf(ilSdfLoop).lCode Then   'found the matching entry within the table of spots for this vehicle & date
                                    tmImpSdf(ilSdfLoop).lCode = -tmImpSdf(ilSdfLoop).lCode      'negate it to indicate it has been processed.  All non-processed spots will be assumed to be missed
                                    Exit For
                                End If
                            Next ilSdfLoop
                        End If                          'not ilFoundll
                        '9-15-06 move above so a processed spot is set; do not mark spots
                        'that was found to be inside a game.  Those spots should be set to  missed or removed if a fill
                        'Found the matching spot from the import file, mark it as processed
                        'For ilSdfLoop = LBound(tmImpSdf) To UBound(tmImpSdf) - 1
                        '    If llSdfCode = tmImpSdf(ilSdfLoop).lCode Then   'found the matching entry within the table of spots for this vehicle & date
                        '        tmImpSdf(ilSdfLoop).lCode = -tmImpSdf(ilSdfLoop).lCode      'negate it to indicate it has been processed.  All non-processed spots will be assumed to be missed
                        '        Exit For
                        '    End If
                        'Next ilSdfLoop
                    End If
                Else                        'no SDF found from the import record
                    If ilAutomationType = 9 Then
                        'spot wasnt found, is this a record to bypass because its not a spot
                        If Trim$(smParsedValues(1)) = "!" Or Trim$(smParsedValues(6)) = ":" Or Trim$(smParsedValues(6)) = "J" Or Trim$(smParsedValues(6)) = """" Or Trim$(smParsedValues(9)) = "" Then      'should this be bypass because its a comment, an align, empty or jump record
                            'bypass, do nothing
                        Else
                            'indicate spot in Error log
                            If llSdfCode <= 0 Then           'no internal code defined
                                lbcMsg.AddItem "Internal spot code missing for " & smFieldValues(ilLoop)    ' slAdvt
                                Print #hmMsg, "Internal spot code missing: " & smFieldValues(ilLoop)
                            Else
                                lbcMsg.AddItem "Spot missing for Internal code " & str$(llSdfCode)
                                Print #hmMsg, "Spot missing for Internal code: " & smFieldValues(ilLoop)
                            End If
                        End If
'                    ElseIf ilAutomationType = AUTOIMPORT_JELLI Then             'processing 1 vehicle within the log vehicle at a time; but must match on the proper matching row in the file
'                        If Trim$(smParsedValues(5)) = Trim$(slVehicleStationCode) Then    'if processing the matching vehicle based on station code, its truly an illegal internal code;
'                            lbcMsg.AddItem "Internal spot code missing for " & smFieldValues(ilLoop)    'slAdvt
'                            Print #hmMsg, "Internal spot code missing: " & smFieldValues(ilLoop)
'                        End If         'the vehicle within the log vehicle isnt the one supposed tobe processing, ignore
                    Else
                        'indicate spot in Error log
                        If llSdfCode <= 0 Then           'no internal code defined
                            lbcMsg.AddItem "Internal spot code missing for " & smFieldValues(ilLoop)    'slAdvt
                            Print #hmMsg, "Internal spot code missing: " & smFieldValues(ilLoop)
                        Else
                            lbcMsg.AddItem "Spot missing for Internal code " & str$(llSdfCode)
                            Print #hmMsg, "Spot missing for Internal code: " & smFieldValues(ilLoop)
                        End If
                    End If
                End If
            End If
        Else                            'rcs
            llAirTime = mConvertTime(smFieldValues(ilLoop), 1, 3, 6, slXMid)        'hhmm:ss
            slTime = gFormatTimeLong(llAirTime, "A", "1")
            'slAdvt = Mid$(smFieldValues(ilLoop), 13, 24)           '3-9-17
            'Obtain spot ID, rt justified
            If ilAutomationType = 11 Then   'rcs-5
                slStr = Mid$(smFieldValues(ilLoop), 22 + 59, 8)
                slAdvt = Mid$(smFieldValues(ilLoop), 22 + 24, 24)   '3-9-17
            Else                            'rcs-4
                slStr = Mid$(smFieldValues(ilLoop), 22 + 48, 8)
                slAdvt = Mid$(smFieldValues(ilLoop), 22 + 13, 24)   '3-9-17
            End If
            llSdfCode = Val(Trim$(slStr))
            tmSdfSrchKey3.lCode = llSdfCode
            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, 3, BTRV_LOCK_NONE, SETFORREADONLY)  'position record so update works
            If ilRet <> BTRV_ERR_NONE Then
                lbcMsg.AddItem "Sdf Missing for" & str$(llSdfCode) & ", " & slAdvt
                Print #hmMsg, "Sdf Missing for " & smFieldValues(ilLoop) & ", " & Trim$(slAdvt)
            Else
                If (ilCombineVefCode <= 0) Or ((ilCombineVefCode > 0) And (tmSdf.iVefCode = ilVefCode)) Then
                    '6/9/10:  Add clearing of spots not found for RCS
                    'Found the matching spot from the import file, mark it as processed
                    For ilSdfLoop = LBound(tmImpSdf) To UBound(tmImpSdf) - 1
                        If llSdfCode = tmImpSdf(ilSdfLoop).lCode Then   'found the matching entry within the table of spots for this vehicle & date
                            tmImpSdf(ilSdfLoop).lCode = -tmImpSdf(ilSdfLoop).lCode      'negate it to indicate it has been processed.  All non-processed spots will be assumed to be missed
                            Exit For
                        End If
                    Next ilSdfLoop
                    gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                    slSchTime = gFormatTimeLong(llTime, "A", "1")

                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                    slSchDate = Format$(llDate, "m/d/yy")

                    tmSdf.sXCrossMidnight = slXMid      '3-28-07

                    slStr = Mid$(smFieldValues(ilLoop), 18, 1)
                    ilECode = Val(Trim$(slStr))
                    slStr = Mid$(smFieldValues(ilLoop), 20, 2)
                    ilDCode = Val(Trim$(slStr))
                    If (ilECode = 3) Or (ilECode = 6) Or (ilDCode = 51) Then  'Missed
                        If (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
                            ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, tmSdf.iGameNo, tmSmf, hmSsf, tmSsf, llSsfMemDate, llSsfRecPos, hmSxf, hmGsf, hmGhf)
                            If Not ilRet Then
                                lbcMsg.AddItem "Sdf Not Set to Missed for" & str$(llSdfCode)
                                Print #hmMsg, "Sdf Not Set to Missed for " & smFieldValues(ilLoop)
                            Else
                                mSaveVefProc            'keep track of vehicles/games processed
                            End If
                        Else
                            tmSdfSrchKey3.lCode = tmSdf.lCode
                            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, 3, BTRV_LOCK_NONE, SETFORWRITE)  'position record so update works
                            If ilRet = BTRV_ERR_NONE Then
                                ilRet = btrDelete(hmSdf)
                                lbcMsg.AddItem "Open/Close Spot Missed and Removed for " & Trim(slAdvt) & " on " & Trim(slDate) & " at " & Trim(slTime)
                                Print #hmMsg, "Open/Close Spot Missed and Removed on " & Trim(slDate) & " at " & Trim(slTime) & ",  " & smFieldValues(ilLoop)
                                mSaveVefProc            'keep track of vehicles/games processed
                            End If
                        End If
                    Else
                        'gPackTime slTime, ilAirTime0, ilAirTime1
                        'If (tmSdf.iTime(0) <> ilAirTime0) Or (tmSdf.iTime(1) <> ilAirTime1) Then
                        '    'Get time and set into Sdf
                        '    Do
                        '        tmSdfSrchKey3.lCode = llSdfCode
                        '        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, 3, BTRV_LOCK_NONE, SETFORWRITE)  'position record so update works
                        '        If ilRet <> BTRV_ERR_NONE Then
                        '            lbcMsg.AddItem "Sdf Not Updated for" & Str$(llSdfCode)
                        '            Print #hmMsg, "Sdf Not Updated for " & smFieldValues(ilLoop)
                        '            Exit Do
                        '        End If
                        '        tmSdf.iTime(0) = ilAirTime0
                        '        tmSdf.iTime(1) = ilAirTime1
                        '        tmSdf.sAffChg = "Y"
                        '        ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                        '    Loop While ilRet = BTRV_ERR_CONFLICT
                        'End If
                        If (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
                            If rbcPLMove(0).Value Then
                                slMOrO = "G"
                            Else
                                slMOrO = "O"
                            End If
                            ilChgSpot = gProcessSpot(hmSsf, hmSdf, hmSmf, hmCHF, hmClf, hmCff, hmRdf, hmCrf, hmCgf, tmSdf, ilVefCode, slSchDate, slSchTime, slDate, slTime, slMOrO, hmSxf, hmGsf, hmGhf)
                        Else
                            tmSdfSrchKey3.lCode = tmSdf.lCode
                            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, 3, BTRV_LOCK_NONE, SETFORWRITE)  'position record so update works
                            If ilRet = BTRV_ERR_NONE Then
                                tmSdf.sXCrossMidnight = slXMid      '3-28-07
                                gPackDate slDate, tmSdf.iDate(0), tmSdf.iDate(1)
                                gPackTime slTime, tmSdf.iTime(0), tmSdf.iTime(1)
                                ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                                If ilRet <> BTRV_ERR_NONE Then
                                    ilChgSpot = 1
                                Else
                                    ilChgSpot = 0
                                End If
                            Else
                                ilChgSpot = 1
                            End If
                        End If
                        If ilChgSpot = 1 Then
                            '6-10-05 process remainder of file
                            'ilRet = btrAbortTrans(hmSdf)
                            lbcMsg.AddItem "Call Counterpoint & Run checks: Cannot find contract or SMF for " & smFieldValues(ilLoop)
                            Print #hmMsg, "Call Counterpoint & Run checks: Cannot find contract or SMF for " & smFieldValues(ilLoop)

                            'lbcMsg.AddItem "Error in Processing, Cancelled" & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                            'Print #hmMsg, "Error in Processing, Cancelled" & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")

                            'Close #hmMsg
                            'Screen.MousePointer = vbDefault
                            'Exit Function
                        ElseIf ilChgSpot = 2 Then
                            lbcMsg.AddItem "Missed To Scheduled Disallowed " & smFieldValues(ilLoop)
                            Print #hmMsg, "Missed To Scheduled Disallowed " & smFieldValues(ilLoop)
                        ElseIf ilChgSpot = 3 Then
                            lbcMsg.AddItem "Date change Disallowed " & smFieldValues(ilLoop)
                            Print #hmMsg, "Date change Disallowed " & smFieldValues(ilLoop)
                        Else
                            mSaveVefProc            'keep track of vehicles/games processed
                        End If

                    End If
                End If
            End If
        End If
    Next ilLoop
    '6/9/10:  Add clearing of spots not found for any import
    'If ilAutomationType = 2 Or ilAutomationType = 6 Or ilAutomationType = 7 Or ilAutomationType = 10 Or ilAutomationType = 9 Or ilAutomationType = 12 Then            'prophet (NexGen & Wizard) or WireReady or RadioMan:  Loop thru all the spots for the vehicle & date to see if anything didnt come in via Import.
                                            'These are are assumed to be missed
        'Found the matching spot from the import file, mark it as processed
        For ilSdfLoop = LBound(tmImpSdf) To UBound(tmImpSdf) - 1
            If tmImpSdf(ilSdfLoop).lCode > 0 Then   'not processed with the import, flag as missed

                gUnpackTimeLong tmImpSdf(ilSdfLoop).iTime(0), tmImpSdf(ilSdfLoop).iTime(1), False, llTime
                slSchTime = gFormatTimeLong(llTime, "A", "1")
                gUnpackDateLong tmImpSdf(ilSdfLoop).iDate(0), tmImpSdf(ilSdfLoop).iDate(1), llDate
                slSchDate = Format$(llDate, "m/d/yy")
                tmSdf = tmImpSdf(ilSdfLoop)
                slSpotType = tmSdf.sSpotType        'need to detrmine if this is a fill spot for message in log
                ilAdvt = gBinarySearchAdf(tmSdf.iAdfCode)
                'If ilAdvt > 0 Then
                If ilAdvt <> -1 Then
                    slAdvt = Trim$(tgCommAdf(ilAdvt).sName)
                Else
                    slAdvt = "Advertiser" & Trim$(str$(tmSdf.iAdfCode)) & " missing"
                End If
                'see if spot already Missed
                If tmSdf.sSchStatus <> "M" Then

                    '5-22-06 see if this spot that was not found with the imported spots should be bypassed
                    'because it was from a live log
                    ilFoundLL = False
                    For ilLoopLL = 0 To UBound(tmIgnoreTimes) - 1
                        If llTime >= tmIgnoreTimes(ilLoopLL).lStartTime And llAirTime <= tmIgnoreTimes(ilLoopLL).lEndTime Then   'import spot falls within a live log time
                            ilFoundLL = True
                            Exit For
                        End If
                    Next ilLoopLL

                    'If Not ilFoundLL Then         'no live timeperiod  found for this spot
                    If (slSpotType <> "O") And (slSpotType <> "C") Then      'not open or close
                        'temporarily do not set all to missed
                        ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, tmSdf.iGameNo, tmSmf, hmSsf, tmSsf, llSsfMemDate, llSsfRecPos, hmSxf, hmGsf, hmGhf)
                        If Not ilRet Then
                            lbcMsg.AddItem "Sdf Not Set to Missed for" & slAdvt & " @" & slSchTime & ", Spot code = " & str$(tmSdf.lCode)
                            Print #hmMsg, "Sdf Not Set to Missed for" & slAdvt & " @" & slSchTime & ", Spot code = " & str$(tmSdf.lCode)
                        Else                'successful update of spot
                            If ilFoundLL Then               '9-15-06 show diff mesg if spot is flagged as missed due to game
                                If slSpotType = "X" Then           'if fill spot, it has been removed.  Note that in log
                                    lbcMsg.AddItem "Fill Spot Pre-empted and Removed due to Event for " & Trim(slAdvt) & " on " & Trim(slSchDate) & " at " & Trim(slTime)
                                    Print #hmMsg, "Fill Spot Pre-empted and Removed due to Event for " & Trim(slAdvt) & " on " & Trim(slSchDate) & " at " & Trim(slTime)     ' & ",  " & smFieldValues(ilLoop)
                                Else    '4-26-05 schedule spot had no import information, set as missed
                                    lbcMsg.AddItem "Spot pre-empted due to Event, set to Missed for " & Trim(slAdvt) & " on " & Trim(slDate) & " at " & Trim(slTime)
                                    Print #hmMsg, "Spot pre-empted due to Event, set to Missed for " & Trim(slAdvt) & " on " & Trim(slDate) & " at " & Trim(slTime)
                                End If
                            Else
                                If slSpotType = "X" Then           'if fill spot, it has been removed.  Note that in log
                                    lbcMsg.AddItem "Fill Spot Missed and Removed for " & Trim(slAdvt) & " on " & Trim(slSchDate) & " at " & Trim(slTime)
                                    Print #hmMsg, "Fill Spot Missed and Removed for " & Trim(slAdvt) & " on " & Trim(slSchDate) & " at " & Trim(slTime)     ' & ",  " & smFieldValues(ilLoop)
                                Else    '4-26-05 schedule spot had no import information, set as missed
                                    '4-19-12 message incorrect
                                    'lbcMsg.AddItem "No scheduled spot for import data, set to Missed for " & Trim(slAdvt) & " on " & Trim(slSchDate) & " at " & Trim(slSchTime)
                                    lbcMsg.AddItem "No import data for scheduled spot, set to Missed for " & Trim(slAdvt) & " on " & Trim(slSchDate) & " at " & Trim(slSchTime)
                                    'Print #hmMsg, "No scheduled spot for import data, set to Missed for " & Trim(slAdvt) & " on " & Trim(slSchDate) & " at " & Trim(slSchTime)
                                    Print #hmMsg, "No import data for scheduled spot, set to Missed for " & Trim(slAdvt) & " on " & Trim(slSchDate) & " at " & Trim(slSchTime)
                                End If
                            End If
                            mSaveVefProc            'keep track of vehicles/games processed
                        End If
                    Else
                        tmSdfSrchKey3.lCode = tmSdf.lCode
                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, 3, BTRV_LOCK_NONE, SETFORWRITE)  'position record so update works
                        If ilRet = BTRV_ERR_NONE Then
                            ilRet = btrDelete(hmSdf)
                            lbcMsg.AddItem "Open/Close Spot Missed and Removed for " & Trim(slAdvt) & " on " & Trim(slSchDate) & " at " & Trim(slSchTime)
                            Print #hmMsg, "Open/Close Spot Missed and Removed for " & Trim(slAdvt) & " on " & Trim(slSchDate) & " at " & Trim(slSchTime)   ' & ",  " & smFieldValues(ilLoop)
                            mSaveVefProc            'keep track of vehicles/games processed
                        End If
                    End If
                    'End If
                End If

            End If
        Next ilSdfLoop
    'End If
    ilRet = btrEndTrans(hmSdf)
    mSaveRec = True
    Screen.MousePointer = vbDefault    'Default
    Exit Function

    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function
Private Sub mSetCommand()
    If rbcPLMove(0).Value Or rbcPLMove(1).Value Then
        cmcImport.Enabled = True
    Else
        cmcImport.Enabled = False
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

    Screen.MousePointer = vbDefault

    sgDoneMsg = ""
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ImptGen
    igManUnload = NO
End Sub

Private Sub rbcPLMove_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcPLMove(Index).Value
    'End of Coded added
    mSetCommand
End Sub
Private Sub plcPLMove_Paint()
    plcPLMove.CurrentX = 0
    plcPLMove.CurrentY = 0
    plcPLMove.Print "Set Discrepant Post Log Moved Spots as"
End Sub

Private Sub tmcCancel_Timer()
    tmcCancel.Enabled = False       'screen has now been focused to show
    cmcCancel_Click         'simulate clicking of cancen button
End Sub

Private Sub mGetMergedGameSpots(slDate As String, ilVefCode As Integer)
    Dim llDate As Long
    Dim llTestDate As Long
    Dim ilRet As Integer
    Dim ilVpf As Integer
    Dim ilMatchOnGame As Boolean
    Dim llSsfRecPos As Long
    Dim ilEvt As Integer
    Dim ilExclAnf As Boolean
    Dim ilAnf As Integer
    Dim ilSpot As Integer
    Dim tlSdf As SDF
    
    If (Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS Then
        ilMatchOnGame = True
        llDate = gDateValue(slDate)
        tmGsfSrchKey4.iAirVefCode = ilVefCode
        gPackDate slDate, tmGsfSrchKey4.iAirDate(0), tmGsfSrchKey4.iAirDate(1)
        ilRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iAirVefCode = ilVefCode)
            gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llTestDate
            If (llTestDate = llDate) Then
                ilVpf = gBinarySearchVpf(tmGsf.iVefCode)
                If ilVpf > 0 Then
                    If ((tgVpf(ilVpf).sGenLog <> "L") And (tgVpf(ilVpf).sGenLog <> "A")) Or ((tgVpf(ilVpf).sGenLog = "A") And (tmGsf.sLiveLogMerge = "M")) Then
                    
                        '10-4-12 obtain the SSF for the merged vehicle
                        
                        'Retrieve the SSF so the Named AVail exclusion flag can be tested
'                        tmSdfSrchKey6.iVefCode = tmGsf.iVefCode
'                        tmSdfSrchKey6.iGameNo = tmGsf.iGameNo
'                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey6, INDEXKEY6, BTRV_LOCK_NONE, SETFORWRITE)
'                        Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = tmGsf.iVefCode) And (tmSdf.iGameNo = tmGsf.iGameNo)
'                            tmImpSdf(UBound(tmImpSdf)) = tmSdf
'                            ReDim Preserve tmImpSdf(0 To UBound(tmImpSdf) + 1) As SDF
'                            ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
'                        Loop
                                                                          
                        'tmSsfSrchKey1.iType = tmGsf.iGameNo
                        'tmSsfSrchKey1.iVefCode = tmGsf.iVefCode
                        'ilRet = gSSFGetGreaterOrEqualKey1(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                        ''ilType = tmSsf.iType
                        imSsfRecLen = Len(tmSsf)
                        tmSsfSrchKey2.iVefCode = tmGsf.iVefCode
                        tmSsfSrchKey2.iDate(0) = tmGsf.iAirDate(0)
                        tmSsfSrchKey2.iDate(1) = tmGsf.iAirDate(1)
                        ilRet = gSSFGetEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                        Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = tmGsf.iVefCode)
                            If (tmSsf.iDate(0) = tmGsf.iAirDate(0)) And (tmSsf.iDate(1) = tmGsf.iAirDate(1)) Then
                                If tmSsf.iType = tmGsf.iGameNo Then
                                    Exit Do
                                End If
                            End If
                            imSsfRecLen = Len(tmSsf)
                            ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                        Loop
                         
                        'Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = tmGsf.iGameNo) And (tmSsf.iVefCode = tmGsf.iVefCode)
                        If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = tmGsf.iGameNo) And (tmSsf.iVefCode = tmGsf.iVefCode) And (tmSsf.iDate(0) = tmGsf.iAirDate(0)) And (tmSsf.iDate(1) = tmGsf.iAirDate(1)) Then
                            ilRet = gSSFGetPosition(hmSsf, llSsfRecPos)
                            ilEvt = 1
                            Do While ilEvt <= tmSsf.iCount
                               LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Avail
                                    '10-4-12 should the avail be included
                                    ilExclAnf = False
                                    For ilAnf = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1        'find the matching avail and determine if it should be incl/excl
                                        If (tgAvailAnf(ilAnf).sAutomationExport = "N") And (tgAvailAnf(ilAnf).iCode = tmAvail.ianfCode) Then        'avail should be excluded
                                            ilExclAnf = True
                                            Exit For
                                        End If
                                    Next ilAnf
                                    
                                    If Not ilExclAnf Then           'include this avail
                                         For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                             ilEvt = ilEvt + 1
                                            LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                             
                                             tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                             ilRet = btrGetEqual(hmSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                             If ilRet = BTRV_ERR_NONE Then
                                                'build into sdf array
                                                tmImpSdf(UBound(tmImpSdf)) = tlSdf           'save entire record
                                                ReDim Preserve tmImpSdf(LBound(tmImpSdf) To UBound(tmImpSdf) + 1) As SDF
                                             End If
                                    
                                         Next ilSpot
                                    Else          'bypass all the spots in avail
                                        For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                             ilEvt = ilEvt + 1
                                         Next ilSpot
                                    End If
                                End If
                                ilEvt = ilEvt + 1
                            Loop
'                            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
'                            ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                            If ilMatchOnGame Then
'                                ilType = tmSsf.iType
'                            End If
                        'Loop
                        End If
                    End If
                End If
            Else
                Exit Do
            End If
            ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
    End If
    Exit Sub
End Sub
'
'           Get the ssf and gather all spots for each avail that are to be included in automation export
'           Avail has a property that indicates the avail can be excluded from automation export;
'           The import must ignore those spots
'
'            mBuildSdfAray - Build array of spots to process.  As each spot is proces, its audo code is negated
'           if at the end of procesing, the spot is not negated, it will be bumped because it wasnt found in import
'           <input> llDate - date to process
'                   ilToVefCode - vehicle to process
'
Public Sub mBuildSDFArray(llDate As Long, ilToVefCode As Integer)
 
    Dim ilDay As Integer
    Dim ilEvt As Integer
    Dim ilSpot As Integer
    Dim ilRet As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim slDate As String
    Dim ilType As Integer
    Dim llSsfDate As Long
    Dim llSsfRecPos As Long
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llTestDate As Long
    Dim tlSdf As SDF
    Dim llLockRecCode As Long
    Dim slUserName As String
    Dim ilVef As Integer
    Dim ilVefIndex As Integer
    Dim ilAnf As Integer
    Dim ilExclAnf As Boolean
    
        ilType = 0
        ilVefIndex = gBinarySearchVef(ilToVefCode)
        If ilVefIndex = -1 Then
            Exit Sub
        End If
        ilDay = gWeekDayLong(llDate)
        slDate = Format$(llDate, "m/d/yy")
        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
        gPackDate slDate, ilLogDate0, ilLogDate1
        
        If tgMVef(ilVefIndex).sType <> "G" Then
            tmSsfSrchKey.iType = ilType
            tmSsfSrchKey.iVefCode = ilToVefCode
            tmSsfSrchKey.iDate(0) = ilLogDate0
            tmSsfSrchKey.iDate(1) = ilLogDate1
            tmSsfSrchKey.iStartTime(0) = 0
            tmSsfSrchKey.iStartTime(1) = 0
            ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        Else
            'gather all games for the day
            tmSsfSrchKey2.iVefCode = ilToVefCode
            tmSsfSrchKey2.iDate(0) = ilLogDate0
            tmSsfSrchKey2.iDate(1) = ilLogDate1
            ilRet = gSSFGetGreaterOrEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
            ilType = tmSsf.iType
                
        End If
        
        Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilToVefCode) And (tmSsf.iDate(0) = ilLogDate0) And (tmSsf.iDate(1) = ilLogDate1)
            ilRet = gSSFGetPosition(hmSsf, llSsfRecPos)
            ilEvt = 1
            Do While ilEvt <= tmSsf.iCount
               LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Avail
                    '10-4-12 should the avail be included
                    ilExclAnf = False
                    For ilAnf = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1        'find the matching avail and determine if it should be incl/excl
                        If (tgAvailAnf(ilAnf).sAutomationExport = "N") And (tgAvailAnf(ilAnf).iCode = tmAvail.ianfCode) Then        'avail should be excluded
                            ilExclAnf = True
                            Exit For
                        End If
                    Next ilAnf
                    
                    If Not ilExclAnf Then           'include this avail
                         For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                             ilEvt = ilEvt + 1
                            LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                             
                             tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                             ilRet = btrGetEqual(hmSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                             If ilRet = BTRV_ERR_NONE Then
                                'build into sdf array
                                tmImpSdf(UBound(tmImpSdf)) = tlSdf           'save entire record
                                ReDim Preserve tmImpSdf(LBound(tmImpSdf) To UBound(tmImpSdf) + 1) As SDF
                             End If
                    
                         Next ilSpot
                    Else          'bypass all the spots in avail
                        For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                             ilEvt = ilEvt + 1
                         Next ilSpot
                    End If
                End If
                ilEvt = ilEvt + 1
            Loop
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            If tgMVef(ilVefIndex).sType = "G" Then
                ilType = tmSsf.iType
            End If
        Loop
     
    Exit Sub
End Sub

Private Sub mAgeParsedValues()
    Dim ilAge As Integer
    
    For ilAge = UBound(smParsedValues) - 1 To LBound(smParsedValues) Step 1
        smParsedValues(ilAge + 1) = smParsedValues(ilAge)
    Next ilAge
End Sub
