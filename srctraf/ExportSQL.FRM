VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ExportSQL 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5730
   ClientLeft      =   435
   ClientTop       =   1920
   ClientWidth     =   7095
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
   Icon            =   "ExportSQL.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5730
   ScaleWidth      =   7095
   Begin VB.CheckBox ckcFile 
      Caption         =   "File"
      Height          =   195
      Left            =   6045
      TabIndex        =   5
      Top             =   390
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CheckBox ckcSQLServer 
      Caption         =   "SQL Server"
      Height          =   225
      Left            =   4575
      TabIndex        =   4
      Top             =   390
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lbcLogMsg 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "ExportSQL.frx":08CA
      Left            =   5310
      List            =   "ExportSQL.frx":08CC
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   135
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6390
      Top             =   4575
   End
   Begin VB.TextBox edcID 
      Height          =   330
      Left            =   1230
      TabIndex        =   2
      Top             =   360
      Width           =   1395
   End
   Begin VB.CheckBox ckcSpots 
      Caption         =   "Export Spots"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   810
      Value           =   1  'Checked
      Width           =   1470
   End
   Begin VB.CheckBox ckcDaypart 
      Caption         =   "Export Dayparts"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1110
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox ckcRegionCopy 
      Caption         =   "Export Regional Copy"
      Height          =   255
      Left            =   2010
      TabIndex        =   10
      Top             =   1110
      Value           =   1  'Checked
      Width           =   2385
   End
   Begin VB.CheckBox ckcVehicle 
      Caption         =   "Export Vehicles"
      Height          =   255
      Left            =   4410
      TabIndex        =   8
      Top             =   810
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox ckcAdvertiser 
      Caption         =   "Export Advertisers"
      Height          =   255
      Left            =   2010
      TabIndex        =   7
      Top             =   810
      Value           =   1  'Checked
      Width           =   2025
   End
   Begin VB.TextBox edcEndDate 
      Height          =   330
      Left            =   1230
      TabIndex        =   14
      Top             =   2370
      Width           =   1395
   End
   Begin VB.TextBox edcStartDate 
      Height          =   330
      Left            =   1230
      TabIndex        =   12
      Top             =   1845
      Width           =   1395
   End
   Begin VB.ListBox lbcVehicle 
      Height          =   2400
      Left            =   2865
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   1500
      Width           =   3840
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   285
      Left            =   2085
      TabIndex        =   17
      Top             =   4485
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "&Export"
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
      Left            =   2385
      TabIndex        =   18
      Top             =   4815
      Width           =   1050
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
      Height          =   285
      Left            =   3720
      TabIndex        =   19
      Top             =   4815
      Width           =   1050
   End
   Begin VB.Label lacSendTo 
      Caption         =   "Send Export To:"
      Height          =   180
      Left            =   3030
      TabIndex        =   3
      Top             =   390
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lacID 
      Caption         =   "Session ID"
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   390
      Width           =   1035
   End
   Begin VB.Label lacEndDate 
      Caption         =   "End Date"
      Height          =   270
      Left            =   105
      TabIndex        =   13
      Top             =   2430
      Width           =   915
   End
   Begin VB.Label lacStartDate 
      Caption         =   "Start Date"
      Height          =   270
      Left            =   105
      TabIndex        =   11
      Top             =   1905
      Width           =   915
   End
   Begin VB.Label lacScreen 
      Caption         =   "Export to SQL Server"
      Height          =   285
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   2250
   End
   Begin VB.Label lacTo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   150
      TabIndex        =   20
      Top             =   5295
      Width           =   5820
   End
   Begin VB.Label lacInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   300
      TabIndex        =   16
      Top             =   4095
      Visible         =   0   'False
      Width           =   6435
   End
End
Attribute VB_Name = "ExportSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software®, Do not copy
'
' File Name: ExportSQL.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Export Selling to Airing Links screen code
Option Explicit
Option Compare Text

'SQL Server
Dim oADOConn As Object
Dim Rs As ADODB.Recordset
Dim RsVeh As ADODB.Recordset
Dim smDBConnectStr As String
Dim smDSN As String
Dim smUID As String
Dim smPWD As String

'Job Status Fields
Dim imJobType As Integer
Dim smStartDate As String
Dim smEndDate As String
Dim smSessionID As String

Dim smIniPathFileName As String
Dim lmTotalNoRecs As Long
Dim lmProcessedNoRecs As Long
Public sgSQLDateForm As String          'Default: yyyy-mm-dd
Public sgShowDateForm As String

Dim imUserInput As Integer
Dim smCurDir As String
Dim smStartIn As String
Dim imFirstTime As Integer
Dim imGenError As Integer
Dim imTestSystem As Integer
Dim lmInputStartDate As Long
Dim lmInputEndDate As Long
'5/17/11: Export all non region rotations
Dim lmAllCrfCode() As Long
Dim lmRafCrfCode() As Long
Dim lmRafCode() As Long
Dim tmAirSellLink() As AIRSELLLINK

'Contract record information
Dim hmCHF As Integer        'Contract line file handle
Dim tmChfSrchKey As LONGKEY0 'CLF key record image
Dim imCHFRecLen As Integer  'CLF record length
Dim tmChf As CHF            'CLF record image

Dim hmClf As Integer        'Contract line file handle
Dim tmClfSrchKey As CLFKEY0 'CLF key record image
Dim imClfRecLen As Integer  'CLF record length
Dim tmClf As CLF            'CLF record image

Dim hmSdf As Integer        'Spot Detail
Dim imSdfRecLen As Integer  'SDF record length
Dim tmSdf As SDF            'SDF record image
Dim tmSdfSrchKey3 As LONGKEY0

'Copy Rotation
Dim tmCrfSrchKey As LONGKEY0  'CRF key record image
Dim tmCrfSrchKey1 As CRFKEY1  'CRF key record image
Dim hmCrf As Integer        'CRF Handle
Dim imCrfRecLen As Integer      'CRF record length
Dim tmCrf As CRF
'Copy/Product
Dim hmCpf As Integer
Dim tmCpfSrchKey As LONGKEY0  'CPF key record image
Dim tmCpf As CPF
Dim imCpfRecLen As Integer  'CPF record length
'Copy instruction record information
Dim hmCnf As Integer        'Copy instruction file handle
Dim tmCnfSrchKey As CNFKEY0 'CNF key record image
Dim imCnfRecLen As Integer  'CNF record length
Dim tmCnf As CNF            'CNF record image
Dim tmCnfRot() As CNF
'Copy inventory record information
Dim hmCif As Integer        'Copy line file handle
Dim tmCifSrchKey As LONGKEY0  'CIF key record image
Dim imCifRecLen As Integer  'CIF record length
Dim tmCif As CIF            'CIF record image
'Regions
Dim tmRafSrchKey As LONGKEY0  'CRF key record image
Dim hmRaf As Integer        'CRF Handle
Dim imRafRecLen As Integer      'CRF record length
Dim tmRaf As RAF
'Comment
Dim tmCsf As CSF            'CSF record image
Dim tmCsfSrchKey As LONGKEY0  'CSF key record image
Dim hmCsf As Integer        'CSF Handle
Dim imCsfRecLen As Integer      'CSF record length

'Comment
Dim tmCef As CEF            'CEF record image
Dim tmCefSrchKey As LONGKEY0  'CEF key record image
Dim hmCef As Integer        'CEF Handle
Dim imCefRecLen As Integer      'CEF record length

'One day file (ODF)
Dim hmOdf As Integer        'One day file
Dim imOdfRecLen As Integer  'ODF record length
Dim tmOdf As ODF            'ODF record image
Dim tmOdfSrchKey2 As ODFKEY2 'ODF key record image

Dim hmAnf As Integer            'Named avail file handle
Dim tmAnfSrchKey As INTKEY0    'ANF record image
Dim imAnfRecLen As Integer      'ANF record length
Dim tmAnf As ANF

'Media code record information
Dim hmMcf As Integer        'Contract line file handle
Dim tmMcfSrchKey As INTKEY0 'MCF key record image
Dim imMcfRecLen As Integer  'MCF record length
Dim tmMcf As MCF            'MCF record image

'Log Spot record
Dim hmLst As Integer        'Log Spots file

Dim hmBof As Integer
Dim hmPrf As Integer
Dim hmSif As Integer

Dim hmVLF As Integer

Dim hmRsf As Integer
Dim hmCvf As Integer

Dim tmVef As VEF

'Prefeed
Dim hmPff As Integer
Dim tmPFF As PFF        'GSF record image
Dim imPffRecLen As Integer        'GSF record length
Dim tmPffSrchKey0 As LONGKEY0
Dim tmPffSrchKey1 As PFFKEY1
Dim tmPrefeedExpt() As PREFEEDEXPT

'Delivery file (DLF)
Dim hmDlf As Integer        'Delivery link file
Dim imDlfRecLen As Integer  'DLF record length
Dim tmDlfSrchKey As DLFKEY0 'DLF key record image
Dim tmDlf As DLF            'DLF record image

Dim imCreating As Integer
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim lmNowDate As Long
Dim lmStartDate As Long
Dim lmEndDate As Long
Dim smSyncDate As String
Dim smSyncTime As String
Dim imATTExist As Integer
Dim smLogType As String         'P=Preliminary; F=Final; R=Reprint
Dim smNewLines(0 To 0) As String * 52   'required as parameter to mBlackoutTest
Dim hmMsg As Integer
Dim lmComment As Long
Dim lmRestrict As Long
Dim lmISCI As Long
Dim lmRegionDef As Long
Dim lmTotalSpots As Long
Dim lmTotalZeroSpots As Long
Dim lmTotalSomeSpots As Long

Dim hmFile As Integer



Private Sub ckcSpots_Click()

    If imUserInput Then
        If ckcSpots.Value = vbChecked Then
            ckcAdvertiser.Value = vbChecked
            ckcAdvertiser.Enabled = False
            ckcVehicle.Value = vbChecked
            ckcVehicle.Enabled = False
            ckcDaypart.Value = vbChecked
            ckcDaypart.Enabled = False
        Else
            ckcAdvertiser.Enabled = True
            ckcVehicle.Enabled = True
            ckcDaypart.Enabled = True
        End If
        If (ckcSpots.Value = vbChecked) Or (ckcRegionCopy.Value = vbChecked) Then
            edcStartDate.Enabled = True
            edcEndDate.Enabled = True
            lbcVehicle.Enabled = True
        Else
            edcStartDate.Enabled = False
            edcEndDate.Enabled = False
            lbcVehicle.Enabled = False
        End If
    End If
End Sub

Private Sub cmcCancel_Click()
    If imCreating Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
'
'
Private Sub cmcExport_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilVef As Integer
    Dim llPercent As Long
    Dim slDate As String
    Dim ilVefCode As Integer
    Dim llDate As Long
    Dim ilAnyVehSelected As Integer
    Dim slSql As String
    Dim llAllCrfCount As Long
        
    lacInfo.Visible = False
    If imCreating Then
        Exit Sub
    End If
    sgSQLDateForm = "yyyy-mm-dd"
    sgShowDateForm = "m/d/yyyy"
    'lMsgBox "Starting Export Process", vbExclamation, ""
    gLogMsg "Starting Export Process", "ExportSQL.txt", False
    
    gLogMsg "", "ExportSQL.txt", False
    
    
    If (ckcSpots.Value = vbChecked) Or (ckcRegionCopy.Value = vbChecked) Then
        'slDate = edcStartDate.Text
        slDate = gAdjYear(Format$(edcStartDate.Text, sgShowDateForm))
        If Not gValidDate(slDate) Then
            imGenError = True
            If imUserInput Then
                Beep
                'lMsgBox "ERROR: Start Date not a Legal Date", vbExclamation, "Error"
                gLogMsg "ERROR: Start Date not a Legal Date", "ExportSQL.txt", False
                edcStartDate.SetFocus
            Else
                'lMsgBox "ERROR: Start Date invalid", vbExclamation, "ERROR"
                gLogMsg "ERROR: Start Date invalid [" & slDate & "]", "ExportSQL.txt", False
            End If
            Exit Sub
        End If
        lmInputStartDate = gDateValue(slDate)
        'lMsgBox "Start Date: " & Format(slDate, "mm/dd/yyyy"), vbExclamation, "Info"
        gLogMsg "Start Date: " & Format(slDate, "mm/dd/yyyy"), "ExportSQL.txt", False
        'slDate = edcEndDate.Text
        slDate = gAdjYear(Format$(edcEndDate.Text, sgShowDateForm))
        If Not gValidDate(slDate) Then
            imGenError = True
            If imUserInput Then
                Beep
                'lMsgBox "ERROR: End Date not a Legal Date", vbExclamation, "Error"
                gLogMsg "ERROR: End Date not a Legal Date", "ExportSQL.txt", False
                edcStartDate.SetFocus
            Else
                'lMsgBox "ERROR: End Date invalid", vbExclamation, "Error"
                gLogMsg "ERROR: End Date invalid", "ExportSQL.txt", False
            End If
            Exit Sub
        End If
        'lMsgBox "End Date: " & Format(slDate, "mm/dd/yyyy"), vbExclamation, "Info"
        gLogMsg "End Date: " & Format(slDate, "mm/dd/yyyy"), "ExportSQL.txt", False
        lmInputEndDate = gDateValue(slDate)
        If lmInputEndDate < lmInputStartDate Then
            imGenError = True
            If imUserInput Then
                Beep
                'lMsgBox "ERROR: Start Date must be prior to End Date", vbExclamation, "Error"
                gLogMsg "ERROR: Start Date must be prior to End Date", "ExportSQL.txt", False
                edcStartDate.SetFocus
            Else
                'lMsgBox "ERROR: Start Date must be prior to End Date", vbExclamation, "Error"
                gLogMsg "ERROR: Start Date must be prior to End Date", "ExportSQL.txt", False
            End If
            Exit Sub
        End If
        ilAnyVehSelected = False
        For ilVef = 0 To lbcVehicle.ListCount - 1 Step 1
            If lbcVehicle.Selected(ilVef) Then
                ilAnyVehSelected = True
                Exit For
            End If
        Next ilVef
                
        'If it's being ran by the scheduler or SBS and NO vehicles were selected
        'then select ALL vehicles
        If ilAnyVehSelected = False And imUserInput = False Then
            For ilVef = 0 To lbcVehicle.ListCount - 1 Step 1
                lbcVehicle.Selected(ilVef) = True
            Next ilVef
            ilAnyVehSelected = True
            'lMsgBox "No vehicles were selected so selecting ALL vehicles.", vbExclamation, "Info"
            gLogMsg "No vehicles were selected so selecting ALL vehicles.", "ExportSQL.txt", False
        End If
        
        'This can only occur if imUserInput = true
        If Not ilAnyVehSelected Then
            imGenError = True
            If imUserInput Then
                Beep
                'lMsgBox "ERROR: No Vehicles Selected", vbExclamation, "Error"
                gLogMsg "ERROR: No Vehicles Selected", "ExportSQL.txt", False
                lbcVehicle.SetFocus
            Else
                'lMsgBox "ERROR: No Vehicles Selected", vbExclamation, "Error"
                gLogMsg "ERROR: No Vehicles Selected", "ExportSQL.txt", False
            End If
            Exit Sub
        End If
    End If
    
    If imUserInput = True Then
        If ckcSQLServer.Value = vbChecked Then
            If Not mVerifyINIValues() Then
                imTerminate = True
                Exit Sub
            End If
            smDBConnectStr = "dsn=" & smDSN & ";uid=" & smUID & ";pwd=" & smPWD & ";"
            If Not ConnectToDB(smDBConnectStr) Then
                Exit Sub
            End If
            slSql = "Delete From SpotOrigDateList"
            oADOConn.Execute slSql
        End If
        If ckcFile.Value = vbChecked Then
            gLogMsgWODT "ON", hmFile, sgDBPath & "Messages\" & "ExportSQLFile.txt"
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    cmcCancel.Caption = "&Cancel"
    imCreating = True
    lacInfo.Caption = ""
    lacInfo.Visible = True
 
    If Not mGetAndExportVehicle() Then
        imGenError = True
        'lMsgBox "ERROR: Export Vehicle Failed", vbExclamation, "Error"
        gLogMsg "ERROR: Export Vehicle Failed", "ExportSQL.txt", False
        mCloseFile
        Exit Sub
    End If

    If Not mGetAndExportAdvertiser() Then
        imGenError = True
        'lMsgBox "ERROR: Export Advertiser Failed", vbExclamation, "Error"
        gLogMsg "ERROR: Export Advertiser Failed", "ExportSQL.txt", False
        mCloseFile
        imGenError = True
        Exit Sub
    End If

    If Not mGetAndExportDaypart() Then
        imGenError = True
        'lMsgBox "ERROR: Export Daypart Failed", vbExclamation, "Error"
        gLogMsg "ERROR: Export Daypart Failed", "ExportSQL.txt", False
        mCloseFile
        Exit Sub
    End If

    
    If (ckcSpots.Value = vbChecked) Or (ckcRegionCopy.Value = vbChecked) Then
        hmMcf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open MCF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open MCF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        '8/2/06:  Required when getting cart number
        imMcfRecLen = Len(tmMcf) 'btrRecordLength(hmSdf)    'Get Cff size

'        hmLst = CBtrvTable(ONEHANDLE)
'        ilRet = btrOpen(hmLst, "", sgDBPath & "Lst.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'        If ilRet <> BTRV_ERR_NONE Then
'            imGenError = True
'            lMsgBox "ERROR: Open LST: " & Str$(ilRet), vbExclamation, "Error"
'            Exit Sub
'        End If
        
        hmRsf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmRsf, "", sgDBPath & "Rsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open RSF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open RSF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        
        hmCvf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open RSF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open CVF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        
        hmVLF = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open VLF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open VLF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        
        hmBof = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmBof, "", sgDBPath & "Bof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open BOF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open BOF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        hmPrf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open PRF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open PRF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        hmSif = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open SIF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open SIF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        
        hmCrf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open CRF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open CRF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        imCrfRecLen = Len(tmCrf) 'btrRecordLength(hmSdf)    'Get Cff size
       
        hmRaf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open RAF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open RAF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        imRafRecLen = Len(tmRaf) 'btrRecordLength(hmSdf)    'Get Cff size
        hmCHF = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open CHF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open CHF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        imCHFRecLen = Len(tmChf) 'btrRecordLength(hmSdf)    'Get Cff size
        
        hmClf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open CLF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open CLF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        imClfRecLen = Len(tmClf) 'btrRecordLength(hmSdf)    'Get Cff size
       
        hmSdf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open SDF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open SDF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        imSdfRecLen = Len(tmSdf) 'btrRecordLength(hmSdf)    'Get Cff size
       
        hmPff = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmPff, "", sgDBPath & "Pff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open SDF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open PFF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        imPffRecLen = Len(tmPFF) 'btrRecordLength(hmSdf)    'Get Cff size
       
        hmDlf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmDlf, "", sgDBPath & "Dlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open SDF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open DLF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        imDlfRecLen = Len(tmDlf) 'btrRecordLength(hmSdf)    'Get Cff size
       
        hmCsf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmCsf, "", sgDBPath & "Csf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open CSF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open CSF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        imCsfRecLen = Len(tmCsf) 'btrRecordLength(hmSdf)    'Get Cff size
       
        hmCef = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open CEF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open CEF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        imCefRecLen = Len(tmCef) 'btrRecordLength(hmSdf)    'Get Cff size
       
       
        hmCnf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmCnf, "", sgDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open CNF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open CNF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        imCnfRecLen = Len(tmCnf) 'btrRecordLength(hmSdf)    'Get Cff size
       
       
        hmCif = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open CIF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open CIF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        imCifRecLen = Len(tmCif) 'btrRecordLength(hmSdf)    'Get Cff size
       
       
        hmCpf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open CPF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open CPF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        imCpfRecLen = Len(tmCpf) 'btrRecordLength(hmSdf)    'Get Cff size
       
        hmOdf = CBtrvTable(TWOHANDLES)
        ilRet = btrOpen(hmOdf, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open ODF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open ODF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        imOdfRecLen = Len(tmOdf) 'btrRecordLength(hmSdf)    'Get Cff size
       
        hmAnf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            imGenError = True
            'lMsgBox "ERROR: Open ANF: " & Str$(ilRet), vbExclamation, "Error"
            gLogMsg "ERROR: Open ANF: " & Str$(ilRet), "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
        imAnfRecLen = Len(tmAnf) 'btrRecordLength(hmSdf)    'Get Cff size
       
        If Not mGetAndExportRegions() Then
            imGenError = True
            'lMsgBox "ERROR: Export Region Failed", vbExclamation, "Error"
            gLogMsg "ERROR: Export Region Failed", "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        End If
    
        If Not mGetAndExportSpots() Then
            imGenError = True
            'lMsgBox "ERROR: Export Spots Failed", vbExclamation, "Error"
            gLogMsg "ERROR: Export Spots Failed", "ExportSQL.txt", False
            mCloseFile
            Exit Sub
        Else
            gLogMsg "** " & CStr(lmTotalZeroSpots) & " *** Vehicle Count With No Spots To Export ***", "ExportSQL.txt", False
            gLogMsg "** " & CStr(lmTotalSomeSpots) & " *** Vehicle Count With Spots To Export ***", "ExportSQL.txt", False
            gLogMsg "** " & CStr(lmTotalSomeSpots + lmTotalZeroSpots) & " *** Total Count Vehicles Tested for Export ***", "ExportSQL.txt", False
            gLogMsg "** " & CStr(lmTotalSpots) & " *** Total Spots Exported ***", "ExportSQL.txt", False
        End If

        If (ckcSpots.Value = vbChecked) Then
            If Not mExportAllCrf(llAllCrfCount) Then
                imGenError = True
                'lMsgBox "ERROR: Export Region Failed", vbExclamation, "Error"
                gLogMsg "ERROR: Export Copy Rotation not associated with Spot Failed", "ExportSQL.txt", False
                mCloseFile
                Exit Sub
            Else
                gLogMsg "** " & CStr(llAllCrfCount) & " *** Total Rotations not tied to Spot Exported ***", "ExportSQL.txt", False
            End If
        Else
            gLogMsg "Copy Rotations not tied to Spots Not Exported as Spot Export Not Requested", "ExportSQL.txt", False
        End If

        ilRet = btrClose(hmAnf)
        btrDestroy hmAnf
        ilRet = btrClose(hmOdf)
        btrDestroy hmOdf
        ilRet = btrClose(hmCpf)
        btrDestroy hmCpf
        ilRet = btrClose(hmCif)
        btrDestroy hmCif
        ilRet = btrClose(hmCnf)
        btrDestroy hmCnf
        ilRet = btrClose(hmCef)
        btrDestroy hmCef
        ilRet = btrClose(hmCsf)
        btrDestroy hmCsf
        ilRet = btrClose(hmDlf)
        btrDestroy hmDlf
        ilRet = btrClose(hmPff)
        btrDestroy hmPff
        ilRet = btrClose(hmSdf)
        btrDestroy hmSdf
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        ilRet = btrClose(hmRaf)
        btrDestroy hmRaf
        ilRet = btrClose(hmLst)
        btrDestroy hmLst
        ilRet = btrClose(hmMcf)
        btrDestroy hmMcf
        ilRet = btrClose(hmBof)
        btrDestroy hmBof
        ilRet = btrClose(hmPrf)
        btrDestroy hmPrf
        ilRet = btrClose(hmSif)
        btrDestroy hmSif
        ilRet = btrClose(hmVLF)
        btrDestroy hmVLF
        ilRet = btrClose(hmRsf)
        btrDestroy hmRsf
        ilRet = btrClose(hmCvf)
        btrDestroy hmCvf
        ilRet = btrClose(hmCrf)
        btrDestroy hmCrf
    End If

    mCloseFile

    If igBkgdProg <> 3 Then
        lacInfo.Visible = True
        imCreating = False
        cmcCancel.Caption = "&Done"
        cmcCancel.SetFocus
        cmcExport.Enabled = False
    End If

    If (Not imTerminate) And (Not imGenError) Then
        'lMsgBox "** Completed Successfully ** ", vbExclamation, ""
        gLogMsg "** Completed Successfully ** ", "ExportSQL.txt", False
    End If
    Screen.MousePointer = vbDefault
    If imTerminate Then
        cmcCancel_Click
    End If
    Exit Sub
cmcExportErr:
    imTerminate = True
    ilRet = Err
    Resume Next
End Sub

Private Sub edcEndDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcID_GotFocus()
    If imFirstTime Then
        If Not tmcStart.Enabled Then
            tmcStart.Enabled = True
        End If
        Exit Sub
    End If
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcStartDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Load()
    
    mInit
    If imTerminate Then
        tmcStart.Enabled = True
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    Erase tmAirSellLink
    '5/17/11: All Copy rotations for date span
    Erase lmAllCrfCode
    Erase lmRafCrfCode
    Erase lmRafCode
    Erase tmCnfRot
    Erase tmPrefeedExpt
    
    'lMsgBox "Closing Files and SQL Connection ", vbExclamation, ""
    gLogMsg "Closing Files and SQL Connection ", "ExportSQL.txt", False
    ilRet = btrClose(hmAnf)
    btrDestroy hmAnf
    ilRet = btrClose(hmOdf)
    btrDestroy hmOdf
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmCnf)
    btrDestroy hmCnf
    ilRet = btrClose(hmCef)
    btrDestroy hmCef
    ilRet = btrClose(hmCsf)
    btrDestroy hmCsf
    ilRet = btrClose(hmDlf)
    btrDestroy hmDlf
    ilRet = btrClose(hmPff)
    btrDestroy hmPff
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmRaf)
    btrDestroy hmRaf
    ilRet = btrClose(hmLst)
    btrDestroy hmLst
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    ilRet = btrClose(hmBof)
    btrDestroy hmBof
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    ilRet = btrClose(hmVLF)
    btrDestroy hmVLF
    ilRet = btrClose(hmRsf)
    btrDestroy hmRsf
    ilRet = btrClose(hmCvf)
    btrDestroy hmCvf
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    
    gLogMsg "** Ending Program **", "ExportSQL.txt", False

    gEraseGlobalVar True
    
    Set ExportSQL = Nothing   'Remove data segment

    btrStopAppl
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
'
'      2-02-01 Move Opens to gcmcExport; right before the
'           generation of summary records.  The timer
'           wasnt activated until after opening; therefore
'           files could remain open and the backup
'           aborted
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDate As String
    Dim hlLogFile As Integer
    Dim slSql As String
    
    imTerminate = False
    imFirstTime = True
    imGenError = False
    lmTotalSpots = 0
    lmTotalZeroSpots = 0
    lmTotalSomeSpots = 0
    
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    
    If Not gCheckDDFDates() Then
        imTerminate = True
        Exit Sub
    End If
    
    If Not mCheckRecLength() Then
        imTerminate = True
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    imCreating = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoRecs = 0
    lmProcessedNoRecs = 0
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    lacTo.Caption = sgDBPath
    gSpfRead
    ilRet = gObtainAdvt()
    ilRet = gObtainVef()
    ilRet = gVpfRead()
    ilRet = gObtainRdf(sgMRdfStamp, tgMRdf())
    ReDim tgNtrSortInfo(0 To 0) As NTRSORTINFO              'array of vehicles with start/end indices pointing to TGNTRInfo array of SBF records
    mPopulate
    gCenterStdAlone ExportSQL
    tmcStart.Enabled = True
    Screen.MousePointer = vbDefault
    
    'D.S. Crank up SQL Server
    
    If imUserInput = False Then
        smDBConnectStr = "dsn=" & smDSN & ";uid=" & smUID & ";pwd=" & smPWD & ";"
        If Not ConnectToDB(smDBConnectStr) Then
            Exit Sub
        End If
    End If
    
    'If imUserInput Then
    '    slSQL = "Delete From SpotOrigDateList"
    '    oADOConn.Execute slSQL
    'End If
    
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    Dim flTextHeight As Single  'Standard text height
    Dim ilLoop As Integer
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
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim slTestSystem As String
    Dim slHelpSystem As String
    Dim slChar As String
    
    slCommand = Command$
    igStdAloneMode = True 'False  'Switch from/to stand alone mode
    sgCallAppName = ""
    igDirectCall = 0
    sgIniPath = ""
    imTestSystem = False
    ilPos = InStr(1, slCommand, "/S:Test", 1)
    If ilPos > 0 Then
        imTestSystem = True
    End If
    ilPos = InStr(1, slCommand, "/S:Prod", 1)
    If ilPos > 0 Then
        imTestSystem = False
    End If
    imUserInput = False
    ilPos = InStr(1, slCommand, "/UserInput", 1)
    If ilPos > 0 Then
        imUserInput = True
        igBkgdProg = 0
        'Debug
        'lMsgBox "Current Mode: UserInput.", vbExclamation, ""
'        gLogMsg "Current Mode: UserInput.", "ExportSQL.txt", False
    Else
        igBkgdProg = 3
        'Debug
        'lMsgBox "Current Mode: Background.", vbExclamation, ""
'        gLogMsg "Current Mode: Background.", "ExportSQL.txt", False

    End If
    smCurDir = CurDir$
    'Debug
    'lMsgBox "CurDir = " & smCurDir, vbExclamation, ""
'    gLogMsg "CurDir = " & smCurDir, "ExportSQL.txt", False
    smStartIn = ""
    ilPos = InStr(1, slCommand, "/StartIn:", 1)
    If ilPos > 0 Then
        smStartIn = Trim$(Mid$(slCommand, ilPos + Len("/StartIn:")))
        'Debug
        'lMsgBox "smStartIn " & smStartIn, vbExclamation, ""
        'gLogMsg "smStartIn " & smStartIn, "ExportSQL.txt", False
        If smStartIn <> "" Then
            'Debug
            'lMsgBox "smStartIn <> ", vbExclamation, """"
         '   gLogMsg "smStartIn <> ", "ExportSQL.txt", False
            
            ChDrive Left$(smStartIn, 1)
            ChDir$ smStartIn
            smCurDir = CurDir$
        End If
    End If
    
    smIniPathFileName = smCurDir & "\ExportSQL.ini"
    If imUserInput = False Then
        If Not mVerifyINIValues() Then
            imTerminate = True
            Exit Sub
        End If
        lacSendTo.Visible = False
        ckcSQLServer.Visible = False
        ckcFile.Visible = False
    Else
        lacSendTo.Visible = True
        ckcSQLServer.Visible = True
        ckcFile.Visible = True
    End If
    slStr = "Guide"
    slCommand = "Traffic\Guide"
    gInitStdAlone ExportSQL, slStr, imTestSystem
    If igBkgdProg = 0 Then
        gLogMsg "Current Mode: UserInput.", "ExportSQL.txt", False
    Else
        gLogMsg "Current Mode: Background.", "ExportSQL.txt", False
    End If
    gLogMsg "CurDir = " & smCurDir, "ExportSQL.txt", False
    If ilPos > 0 Then
        gLogMsg "smStartIn " & smStartIn, "ExportSQL.txt", False
        If smStartIn <> "" Then
            'Debug
            gLogMsg "smStartIn <> ", "ExportSQL.txt", False
        End If
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
    Dim ilRet As Integer

    
    Screen.MousePointer = vbDefault
    igParentRestarted = False
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'lMsgBox "** Ending Program **", vbExclamation, ""
    Unload ExportSQL
    igManUnload = NO
End Sub


Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilVpf As Integer
    
    'If tgSpf.sMktBase = "Y" Then
    '    llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH + VEHBYMKT + VEHBYPASSNOLOG ' Airing and all conventional vehicles (except with Log) and Log
    '    ilRet = gPopUserVehicleByMkt(Logs, llFilter, igLogMktCode(), lbcVehicle, tgLogVehicle(), sgLogVehicleTag)
    'Else
    '    llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH + VEHBYPASSNOLOG ' Airing and all conventional vehicles (except with Log) and Log
    '    ilRet = gPopUserVehicleBox(ExportSQL, llFilter, lbcVehicle, tgLogVehicle(), sgLogVehicleTag)
    'End If
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        ilVpf = gBinarySearchVpf(tgMVef(ilVef).iCode)
        If ilVpf <> -1 Then
            If tgVpf(ilVpf).sExportSQL = "Y" Then
                If ((tgMVef(ilVef).sType = "C") And (tgMVef(ilVef).iVefCode <= 0)) Or (tgMVef(ilVef).sType = "A") Or (tgMVef(ilVef).sType = "L") Or (tgMVef(ilVef).sType = "G") Then
                    lbcVehicle.AddItem Trim$(tgMVef(ilVef).sName)
                    lbcVehicle.ItemData(lbcVehicle.NewIndex) = tgMVef(ilVef).iCode
                End If
            End If
        End If
    Next ilVef
End Sub

Private Function mGetAndExportAdvertiser() As Integer
    Dim ilRet As Integer
    Dim ilAdf As Integer
    Dim slName As String
    Dim slSql As String
    Dim ilAdfCode As Integer
    
    mGetAndExportAdvertiser = True
    lmProcessedNoRecs = 0
    If (ckcAdvertiser.Value = vbUnchecked) And (ckcSpots.Value = vbUnchecked) And (ckcRegionCopy.Value = vbUnchecked) Then
        'lMsgBox "Advertiser Export Not Requested", vbExclamation, ""
        gLogMsg "Advertiser Export Not Requested", "ExportSQL.txt", False
        Exit Function
    End If
    'lMsgBox "Advertiser Export Started", vbExclamation, ""
    gLogMsg "Advertiser Export Started", "ExportSQL.txt", False
    ilRet = gObtainAdvt()
    If Not ilRet Then
        imGenError = True
        'lMsgBox "Advertiser Export Started", vbExclamation, ""
        gLogMsg "Advertiser Export Started", "ExportSQL.txt", False
        mGetAndExportAdvertiser = False
        Exit Function
    End If
    For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
        If imTerminate Then
            'lMsgBox "Export Terminated", vbExclamation, ""
            gLogMsg "Export Terminated", "ExportSQL.txt", False
            mGetAndExportAdvertiser = False
            Exit Function
        End If
        ilAdfCode = tgCommAdf(ilAdf).iCode
        slName = Trim$(tgCommAdf(ilAdf).sName)
        
        'Doug:  Call Store Procedure at this point for Advertiser table:
                'advtCSIID=ilAdfCode, advtName=slName
        slName = gFixQuote(slName)
        slSql = ilAdfCode & ", "
        slSql = slSql & "'" & slName & "'"

        If Not mRunSP("sp_InsertAdvertiser", slSql) Then
            'lMsgBox "ERROR: executing sp_InsertAdvertiser statement.", vbCritical, "Error"
            gLogMsg "ERROR: executing sp_InsertAdvertiser statement.", "ExportSQL.txt", False
            mGetAndExportAdvertiser = False
            Exit Function
        End If
        lmProcessedNoRecs = lmProcessedNoRecs + 1
    Next ilAdf
                
    'lMsgBox "Advertiser Export Completed. " & CStr(lmProcessedNoRecs) & " Records Were Exported", vbExclamation, ""
    gLogMsg CStr(lmProcessedNoRecs) & " Records Were Exported. Advertiser Export Completed. ", "ExportSQL.txt", False
End Function

Private Function mGetAndExportVehicle() As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim slName As String
    '8/2/06:  Include Station code with vehicle
    Dim slStationCode As String
    Dim ilVefCode As Integer
    Dim slType As String
    Dim slLLD As String
    Dim ilVpf As Integer
    Dim slSql As String
    Dim slAllowSelection As String
    Dim llVehCount As Long
    
    mGetAndExportVehicle = True
    lmProcessedNoRecs = 0
    If (ckcVehicle.Value = vbUnchecked) And (ckcSpots.Value = vbUnchecked) And (ckcRegionCopy.Value = vbUnchecked) Then
        'lMsgBox "Vehicle Export Not Requested", vbExclamation, ""
        gLogMsg "Vehicle Export Not Requested", "ExportSQL.txt", False
        Exit Function
    End If
    'lMsgBox "Vehicle Export Started", vbExclamation, ""
    gLogMsg "Vehicle Export Started", "ExportSQL.txt", False
    ilRet = gObtainVef()
    If Not ilRet Then
        imGenError = True
        'lMsgBox "ERROR: Obtain Vehicles Failed", vbExclamation, "Error"
        gLogMsg "ERROR: Obtain Vehicles Failed", "ExportSQL.txt", False
        mGetAndExportVehicle = False
        Exit Function
    End If
    ilRet = gVpfRead()
    If Not ilRet Then
        imGenError = True
        'lMsgBox "ERROR: Obtain Vehicle Options Failed", vbExclamation, "Error"
        gLogMsg "ERROR: Obtain Vehicle Options Failed", "ExportSQL.txt", False
        mGetAndExportVehicle = False
        Exit Function
    End If
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If imTerminate Then
            'lMsgBox "Export Terminated", vbExclamation, ""
            gLogMsg "Export Terminated", "ExportSQL.txt", False
            mGetAndExportVehicle = False
            Exit Function
        End If
        If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "A") Or (tgMVef(ilVef).sType = "S") Then
            ilVefCode = tgMVef(ilVef).iCode
            slName = Trim$(tgMVef(ilVef).sName)
            slType = tgMVef(ilVef).sType
            '8/2/06:  Get station code
            slStationCode = Trim$(tgMVef(ilVef).sCodeStn)
            ilVpf = gBinarySearchVpf(ilVefCode)
            If ilVpf <> -1 Then
                gUnpackDate tgVpf(ilVpf).iLLD(0), tgVpf(ilVpf).iLLD(1), slLLD
                slLLD = Format(slLLD, "yyyy-mm-dd")
                slAllowSelection = "N"
                If tgVpf(ilVpf).sExportSQL = "Y" Then
                    slAllowSelection = "Y"
                End If
            Else
                slLLD = ""
                slAllowSelection = "N"
            End If
                    
            'Doug:  Call Store Procedure at this point for vehicle table:
                    'vehCSIID=ilVefCode, vehName=slName, vehType=slType, vehLLD=slLLD,
                    'vehAllowSelection = slAllowSelection
'8/2/06-Doug:  Add slStationCode to stored procedure
            slName = gFixQuote(slName)
            slSql = ilVefCode & ", "
            slSql = slSql & "'" & gFixQuote(slName) & "', "
            slSql = slSql & "'" & gFixQuote(slType) & "', "
            slSql = slSql & "'" & slLLD & "', "
            slSql = slSql & "'" & slAllowSelection & "', "
            slSql = slSql & "'" & slStationCode & "'"
            
            If Not mRunSP("sp_InsertVehicles", slSql) Then
                'lMsgBox "ERROR: executing sp_InsertVehicles statement.", vbCritical, "Error"
                gLogMsg "ERROR: executing sp_InsertVehicles statement.", "ExportSQL.txt", False
                mGetAndExportVehicle = False
                Exit Function
            End If
            lmProcessedNoRecs = lmProcessedNoRecs + 1
        End If
    Next ilVef
    'lMsgBox "Vehicle Export Completed. " & CStr(lmProcessedNoRecs) & " Records Were Exported", vbExclamation, ""
    gLogMsg CStr(lmProcessedNoRecs) & " Records Were Exported. Vehicle Export Completed.", "ExportSQL.txt", False

End Function


Private Function mGetAndExportDaypart() As Integer
    Dim ilRet As Integer
    Dim ilRdf As Integer
    Dim ilIndex As Integer
    Dim slName As String
    Dim ilRdfCode As Integer
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slDays As String
    Dim ilDay As Integer
    Dim slSql As String
    Dim llDP_Def As Long
    Dim llDP_Pat As Long
    
    mGetAndExportDaypart = True
    llDP_Def = 0
    If (ckcDaypart.Value = vbUnchecked) And (ckcSpots.Value = vbUnchecked) Then
        'lMsgBox "Daypart Export Not Requested", vbExclamation, ""
        gLogMsg "Daypart Export Not Requested", "ExportSQL.txt", False
        Exit Function
    End If
    'lMsgBox "Daypart Export Started", vbExclamation, ""
    gLogMsg "Daypart Export Started", "ExportSQL.txt", False
    ilRet = gObtainRdf(sgMRdfStamp, tgMRdf())
    If Not ilRet Then
        imGenError = True
        'lMsgBox "ERROR: Obtain Dayparts Failed", vbExclamation, "Error"
        gLogMsg "ERROR: Obtain Dayparts Failed", "ExportSQL.txt", False
        mGetAndExportDaypart = False
        Exit Function
    End If
    For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
        If imTerminate Then
            'lMsgBox "Export Terminated", vbExclamation, ""
            gLogMsg "Export Terminated", "ExportSQL.txt", False
            mGetAndExportDaypart = False
            Exit Function
        End If
        ilRdfCode = tgMRdf(ilRdf).iCode
        slName = Trim$(tgMRdf(ilRdf).sName)
        
        'Doug:  Call Store Procedure at this point for Daypart Definition table:
                'dayCSIID=ilRdfCode,dayName=slName
        slName = gFixQuote(slName)
        slSql = ilRdfCode & ", "
        slSql = slSql & "'" & slName & "'"
        

        If Not mRunSP("sp_InsertDaypart_Definition", slSql) Then
            'lMsgBox "ERROR: executing sp_InsertDaypart_Definition.", vbCritical, "Error"
            gLogMsg "ERROR: executing sp_InsertDaypart_Definition.", "ExportSQL.txt", False
            Exit Function
        End If
        llDP_Def = llDP_Def + 1
        For ilIndex = 1 To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
            If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                gUnpackTime tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), "A", "1", slStartTime
                slStartTime = Format$(gConvertTime(slStartTime), "hh:mm:ss")
                gUnpackTime tgMRdf(ilRdf).iEndTime(0, ilIndex), tgMRdf(ilRdf).iEndTime(1, ilIndex), "A", "1", slEndTime
                slEndTime = Format$(gConvertTime(slEndTime), "hh:mm:ss")
                slDays = ""
                'For ilDay = 1 To 7 Step 1
                For ilDay = 0 To 6 Step 1
                    slDays = slDays & tgMRdf(ilRdf).sWkDays(ilIndex, ilDay)
                Next ilDay
                
                'Doug:  Call Store Procedure at this point for Daypart Pattern table:
                        'patDayCSIID=ilRdfCode, patStartTime=slStartTime, patEndTime=slEndTime,
                        'patDays=slDays
                slSql = ilRdfCode & ", "
                slSql = slSql & "'" & slStartTime & "', "
                slSql = slSql & "'" & slEndTime & "', "
                slSql = slSql & "'" & slDays & "'"
                
                If Not mRunSP("sp_InsertDaypart_Pattern", slSql) Then
                    'lMsgBox "ERROR: executing sp_InsertDaypart_Pattern.", vbCritical, "Error"
                    gLogMsg "ERROR: executing sp_InsertDaypart_Pattern.", "ExportSQL.txt", False
                    mGetAndExportDaypart = False
                    Exit Function
                End If
                llDP_Pat = llDP_Pat + 1
            End If
        Next ilIndex
        
    Next ilRdf
    'lMsgBox "Daypart Export Completed. " & CStr(llDP_Def) & " Definition And " & CStr(llDP_Pat) & " Pattern Records Were Exported", vbExclamation, """"
    gLogMsg CStr(llDP_Def) & " Definition And " & CStr(llDP_Pat) & " Pattern Records Were Exported. Daypart Export Completed.", "ExportSQL.txt", False

End Function

Private Sub tmcStart_Timer()
    Dim ilRet As Integer
    Dim slStr As String
    Dim slSql As String
    Dim ilIdx As Integer
    Dim slJobType As String
    Dim slVehName As String
    Dim ilVehCode As Integer
    Dim ilLoop As Integer
    
    tmcStart.Enabled = False
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    If imFirstTime Then
        imFirstTime = False
    End If
    'Job separator line
    'lMsgBox "", vbExclamation, ""
    gLogMsg "", "ExportSQL.txt", False
    
    'lMsgBox "** Starting Program **", vbExclamation, ""
    gLogMsg "** Starting Program **", "ExportSQL.txt", False
    'lMsgBox "Start-In Folder: " & smCurDir, vbExclamation, ""
    gLogMsg "Start-In Folder: " & smCurDir, "ExportSQL.txt", False
    'lMsgBox "INI Folder: " & smIniPathFileName, vbExclamation, ""
    gLogMsg "INI Folder: " & smIniPathFileName, "ExportSQL.txt", False

    If imUserInput Then
        'lMsgBox "SQL Export Started by User", vbExclamation, ""
        gLogMsg "SQL Export Started by User", "ExportSQL.txt", False
    Else
        'lMsgBox "SQL Export Started by SQL Server.", vbExclamation, ""
        gLogMsg "SQL Export Started by SQL Server.", "ExportSQL.txt", False
        Set Rs = CreateObject("ADODB.Recordset")
        slSql = "Select * From JobRequestStatus"
        Rs.Open slSql, oADOConn
        If Rs.EOF Then
            gLogMsg "tmcStart_Timer: JobRequestStatus is EOF.", "ExportSQL.txt", False
            Exit Sub
        End If
        
        'Set the module vars to the JobRequestStatus
        imJobType = Rs!JobType
        smStartDate = Rs!StartDate
        smEndDate = Rs!EndDate
        smSessionID = Rs!SessionID
        'lMsgBox "Session ID: " & smSessionID, vbExclamation, ""
        gLogMsg "Session ID: " & smSessionID, "ExportSQL.txt", False
        
        Set RsVeh = CreateObject("ADODB.Recordset")
        slSql = "SELECT * From vw_RequestedVehicles WHERE SessionID = " & "'" & Rs!SessionID & "'"
        RsVeh.Open slSql, oADOConn
        
        'Set the Form vars to the JobRequestStatus
        edcID.Text = smSessionID
        edcStartDate.Text = smStartDate
        edcEndDate.Text = smEndDate
        
        'Init the Form check boxes to unChecked
        ckcVehicle.Value = vbUnchecked
        ckcAdvertiser.Value = vbUnchecked
        ckcDaypart.Value = vbUnchecked
        ckcRegionCopy = vbUnchecked
        ckcSpots = vbUnchecked
        
        'If we need Spots or Regional Copy the we get the selected vehicles
        If imJobType = 1 Or imJobType = 4 Then
            Do While Not RsVeh.EOF
                slVehName = RsVeh!vehName
                ilVehCode = RsVeh!vehCSIID
                
                For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
                    If lbcVehicle.ItemData(ilLoop) = ilVehCode Then
                        lbcVehicle.Selected(ilLoop) = True
                        Exit For
                    End If
                Next ilLoop
                Call RsVeh.MoveNext
            Loop
        End If
        
        Select Case imJobType
            Case 1  'Spots
                slJobType = "Spots"
                ckcVehicle.Value = vbChecked
                ckcAdvertiser.Value = vbChecked
                ckcDaypart.Value = vbChecked
                ckcRegionCopy = vbChecked
                ckcSpots = vbChecked
                'lMsgBox "Job Type: " & CStr(imJobType) & " = Spots - Spots, Dayparts, Veh and Adv.", vbExclamation, """"
                gLogMsg "Job Type: " & CStr(imJobType) & " = Spots - Spots, Dayparts, Veh and Adv.", "ExportSQL.txt", False
            
            Case 4  'Regional Copy
                slJobType = "Regional Copy"
                ckcRegionCopy = vbChecked
                ckcVehicle.Value = vbChecked
                ckcAdvertiser.Value = vbChecked
                ckcDaypart.Value = vbChecked
                'lMsgBox "Job Type: " & CStr(imJobType) & " = Regional Copy - Region Definitions, Comments, Region ISCI pattern, Adv., and Veh.", vbExclamation, ""
                gLogMsg "Job Type: " & CStr(imJobType) & " = Regional Copy - Region Definitions, Comments, Region ISCI pattern, Adv., and Veh.", "ExportSQL.txt", False
            
            Case 2  'Vehicles
                slJobType = "Vehicles"
                ckcVehicle.Value = vbChecked
                'lMsgBox "Job Type: " & CStr(imJobType) & " = Vehicles", vbExclamation, ""
                gLogMsg "Job Type: " & CStr(imJobType) & " = Vehicles", "ExportSQL.txt", False
            
            Case 3  'Advertisers
                slJobType = "Advertisers"
                ckcAdvertiser.Value = vbChecked
                'lMsgBox "Job Type: " & CStr(imJobType) & " = Advertisers", vbExclamation, ""
                gLogMsg "Job Type: " & CStr(imJobType) & " = Advertisers", "ExportSQL.txt", False
            
                    
            Case 5  'Dayparts
                slJobType = "Dayparts"
                ckcDaypart.Value = vbChecked
                'lMsgBox "Job Type: " & CStr(imJobType) & " = DayParts -  Daypart Definitions and the Daypart Patterns", vbExclamation, ""
                gLogMsg "Job Type: " & CStr(imJobType) & " = DayParts -  Daypart Definitions and the Daypart Patterns", "ExportSQL.txt", False
            
            Case Else
                'OOPS something is wrong
                'lMsgBox "Job Type: UNKNOWN", vbExclamation, ""
                gLogMsg "Job Type: UNKNOWN", "ExportSQL.txt", False
            End Select

            slSql = "'" & Rs!SessionID & "'," & 1 & ", '" & "Gathering " & slJobType & "'"
            If Not mRunSP("sp_UpdateJobRequestStatus", slSql) Then
                'lMsgBox "ERROR: executing sp_UpdateJobRequestStatus.", vbCritical, "Error"
                gLogMsg "ERROR: executing sp_UpdateJobRequestStatus.", "ExportSQL.txt", False
            End If

        
        'Doug:  At this point you will need to add the code to get what should be generated
        '       Get Session ID and store into edcID
        
        '       Get if Spots are to be generated.  Set ckcSpots (Note: If spots selected, then ckcAdvertiser, ckcDaypart, and ckcVehicle must be checked
        
        '       Get if Advertiser are to be generated.  Set ckcAdvertiser
        
        '       Get if Vehicles are to be generated.  Set ckcVehicle
        
        '       Get if Daypart are to be generated.  Set ckcDaypart
        
        '       Get if Regional Copy is to be generated.  Set ckcRegionCopy
        
        '       Get Start Date if Spots or Regional Copy is to be generated
        '       Get End Date if Spots or Regional Copy is to be generated
        
        '       Get Vehicle names to generate.  Set lbcVehicle
        
        
        'Minimize window
       ' ExportSQL.WindowState = vbMinimized
        
        'Start Processing
        cmcExport_Click
        If imGenError Then
            'Doug-  Export Failed
            slSql = "'" & Rs!SessionID & "'," & 2 & ", '" & "Gathering " & slJobType & " Failed" & "'"
            If Not mRunSP("sp_UpdateJobRequestStatus", slSql) Then
                'lMsgBox "ERROR: executing sp_UpdateJobRequestStatus.", vbCritical, "Error"
                gLogMsg "ERROR: executing sp_UpdateJobRequestStatus.", "ExportSQL.txt", False
            End If
        Else
            slSql = "'" & Rs!SessionID & "'," & 0 & ", '" & "Gathering " & slJobType & " Completed Successfully" & "'"
            If Not mRunSP("sp_UpdateJobRequestStatus", slSql) Then
                'lMsgBox "ERROR: executing sp_UpdateJobRequestStatus.", vbCritical, "Error"
                gLogMsg "ERROR: executing sp_UpdateJobRequestStatus.", "ExportSQL.txt", False
            End If
        End If
        'D.S. Shut SQL Server connection down
        Call CloseDB
        Set Rs = Nothing
        Set RsVeh = Nothing
        'Exist
        imCreating = False
        cmcCancel_Click
    End If
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mRotPop                         *
'*                                                     *
'*             Created:8/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain rotation specifications *
'*                      Same code is in BulkFeed.Frm   *
'*                                                     *
'*******************************************************
Private Sub mRotPop()
'
'   iRet = mRotPop
'   Where:
'
    Dim ilRet As Integer    'Return status
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilOffSet As Integer
    Dim ilExtLen As Integer
    Dim ilVefSelected As Integer
    Dim llRotStartDate As Long
    Dim llRotEndDate As Long
    Dim ilRotOk As Integer
    Dim llTstStartDate As Long
    Dim llTstEndDate As Long
    Dim ilLink As Integer
    Dim ilStartIndex As Integer
    Dim ilSelVefCode As Integer
    Dim ilRaf As Integer
    Dim ilCrf As Integer
    Dim ilFound As Integer
    Dim tlIntTypeBuff As POPICODE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record

    '5/17/11: All Copy rotations for date span
    ReDim lmAllCrfCode(0 To 0) As Long
    ReDim lmRafCrfCode(0 To 0) As Long
    ReDim lmRafCode(0 To 0) As Long

    
    'lMsgBox "Obtaining Rotation Specifications.", vbExclamation, ""
    gLogMsg "Obtaining Rotation Specifications.", "ExportSQL.txt", False
    btrExtClear hmCrf   'Clear any previous extend operation
    ilExtLen = Len(tmCrf)  'Extract operation record size
    tmCrfSrchKey1.sRotType = "A"
    tmCrfSrchKey1.iEtfCode = 0
    tmCrfSrchKey1.iEnfCode = 0
    tmCrfSrchKey1.iAdfCode = 0
    tmCrfSrchKey1.lChfCode = 0
    tmCrfSrchKey1.lFsfCode = 0
    tmCrfSrchKey1.iVefCode = 0
    tmCrfSrchKey1.iRotNo = 32000
    ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
    Call btrExtSetBounds(hmCrf, llNoRec, -1, "UC", "CRF", "") 'Set extract limits (all records)
    '5/17/11: All Copy rotations for date span
    'ilOffset = gFieldOffset("Crf", "CrfZone")
    'ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, ByVal "R", 1)
    gPackDateLong lmInputEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Crf", "CrfStartDate")
    ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    gPackDateLong lmInputStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Crf", "CrfEndDate")
    ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
    ilOffSet = 0
    ilRet = btrExtAddField(hmCrf, ilOffSet, ilExtLen)  'Extract start/end time, and days
    On Error GoTo mRotPopErr
    gBtrvErrorMsg ilRet, "mRotPop (btrExtAddField):" & "Crf.Btr", ExportSQL
    On Error GoTo 0
    'ilRet = btrExtGetNextExt(hmClf)    'Extract record
    ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If ilRet = BTRV_ERR_REJECT_COUNT Then
            ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
        End If
        Do While ilRet = BTRV_ERR_NONE
            If imTerminate Then
                Exit Sub
            End If
            ilRotOk = True
            gUnpackDateLong tmCrf.iStartDate(0), tmCrf.iStartDate(1), llRotStartDate
            gUnpackDateLong tmCrf.iEndDate(0), tmCrf.iEndDate(1), llRotEndDate
            If ((llRotEndDate >= lmInputStartDate) And (llRotStartDate <= lmInputEndDate)) Then 'And (tmCrf.sState <> "D") Then
                '5/17/11: All Copy rotations for date span
                'If Trim$(tmCrf.sZone) <> "R" Then
                '    ilRotOk = False
                'End If
                '5/27/11: Bypass Dormant rotations
                'If tmCrf.sState <> "D" Then
                '1/20/11
                If tmCrf.sState = "D" Then
                    ilRotOk = False
                End If
            Else
                ilRotOk = False
            End If
            ilStartIndex = 0
            If ilRotOk Then
                Do
                    ilVefSelected = False
                    For ilLoop = ilStartIndex To lbcVehicle.ListCount - 1 Step 1
                        If lbcVehicle.Selected(ilLoop) Then
                            ilSelVefCode = Val(lbcVehicle.ItemData(ilLoop))
                            If tmCrf.iVefCode = ilSelVefCode Then
                                ilStartIndex = ilLoop
                                ilVefSelected = True
                                Exit For
                            End If
                        End If
                    Next ilLoop
                    If Not ilVefSelected Then
                        For ilLoop = ilStartIndex To lbcVehicle.ListCount - 1 Step 1
                            If lbcVehicle.Selected(ilLoop) Then
                                ilSelVefCode = Val(lbcVehicle.ItemData(ilLoop))
                                ilRet = gBinarySearchVef(ilSelVefCode)
                                If ilRet <> -1 Then
                                    If tgMVef(ilRet).sType = "A" Then
                                         For ilLink = 0 To UBound(tmAirSellLink) - 1 Step 1
                                            If tmAirSellLink(ilLink).iAirVefCode = ilSelVefCode Then
                                                If tmAirSellLink(ilLink).iSellVefCode = tmCrf.iVefCode Then
                                                    ilStartIndex = ilLoop
                                                    ilVefSelected = True
                                                    Exit For
                                                End If
                                            End If
                                        Next ilLink
                                        If ilVefSelected = True Then
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        Next ilLoop
                    End If
                    If ilVefSelected Then
                        '5/17/11: All Copy rotations for date span
                        If Trim$(tmCrf.sZone) <> "R" Then
                            ilFound = False
                            For ilCrf = 0 To UBound(lmAllCrfCode) - 1 Step 1
                                If lmAllCrfCode(ilCrf) = tmCrf.lCode Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilCrf
                            If Not ilFound Then
                                lmAllCrfCode(UBound(lmAllCrfCode)) = tmCrf.lCode
                                ReDim Preserve lmAllCrfCode(0 To UBound(lmAllCrfCode) + 1) As Long
                            End If
                        Else
                            ilFound = False
                            For ilRaf = 0 To UBound(lmRafCode) - 1 Step 1
                                If lmRafCode(ilRaf) = tmCrf.lRafCode Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilRaf
                            If Not ilFound Then
                                lmRafCode(UBound(lmRafCode)) = tmCrf.lRafCode
                                ReDim Preserve lmRafCode(0 To UBound(lmRafCode) + 1) As Long
                            End If
                            ilFound = False
                            For ilCrf = 0 To UBound(lmRafCrfCode) - 1 Step 1
                                If lmRafCrfCode(ilCrf) = tmCrf.lCode Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilCrf
                            If Not ilFound Then
                                lmRafCrfCode(UBound(lmRafCrfCode)) = tmCrf.lCode
                                ReDim Preserve lmRafCrfCode(0 To UBound(lmRafCrfCode) + 1) As Long
                            End If
                        End If
                    End If
                    ilStartIndex = ilStartIndex + 1
                Loop While ilVefSelected
            End If
            ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
            If ilRet = BTRV_ERR_REJECT_COUNT Then
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
            End If
        Loop
    End If
    'lMsgBox "Obtaining Rotation Specifications Completed.", vbExclamation, ""
    gLogMsg "Obtaining Rotation Specifications Completed.", "ExportSQL.txt", False
    Exit Sub
mRotPopErr:
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub mBuildLinkTables()
    Dim slStartDate As String
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilAirVefCode As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilIndex As Integer
    Dim ilTest As Integer
    Dim blFound As Boolean
    ReDim ilVefCode(0 To 0) As Integer
    ReDim tmAirSellLink(0 To 0) As AIRSELLLINK
    
    'lMsgBox "Start Building Table Links.", vbExclamation, ""
    gLogMsg "Start Building Table Links.", "ExportSQL.txt", False
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If imTerminate Then
            Exit Sub
        End If
        If lbcVehicle.Selected(ilLoop) Then
            ilAirVefCode = Val(lbcVehicle.ItemData(ilLoop))
            ilIndex = gBinarySearchVef(ilAirVefCode)
            If ilIndex <> -1 Then
                tmVef = tgMVef(ilIndex)
                If tmVef.sType = "A" Then
                    If imTerminate Then
                        Exit Sub
                    End If
                    slStartDate = Format$(lmInputStartDate, "m/d/yy")
                    '7/20/12: Add loop on date range
                    Do
                        gBuildLinkArray hmVLF, tmVef, slStartDate, ilVefCode()
                        For ilVef = 0 To UBound(ilVefCode) - 1 Step 1
                            blFound = False
                            For ilTest = 0 To UBound(tmAirSellLink) - 1 Step 1
                                If tmAirSellLink(ilTest).iSellVefCode = ilVefCode(ilVef) Then
                                    blFound = True
                                End If
                            Next ilTest
                            If Not blFound Then
                                tmAirSellLink(UBound(tmAirSellLink)).iAirVefCode = ilAirVefCode
                                tmAirSellLink(UBound(tmAirSellLink)).iSellVefCode = ilVefCode(ilVef)
                                ReDim Preserve tmAirSellLink(0 To UBound(tmAirSellLink) + 1) As AIRSELLLINK
                            End If
                        Next ilVef
                        slStartDate = DateAdd("d", 7, slStartDate)
                    Loop While gDateValue(slStartDate) <= lmInputEndDate
                End If
            End If
        End If
    Next ilLoop
    'Debug
    'lMsgBox "Building Table Links Completed.", vbExclamation, ""
    gLogMsg "Building Table Links Completed.", "ExportSQL.txt", False
End Sub


Private Function mGetAndExportRegions() As Integer
    Dim llCrf As Long
    Dim llRaf As Long
    Dim ilRet As Integer
    Dim llRafCode As Long
    Dim llRegionCode As Long
    Dim slRafName As String
    Dim ilRafAdfCode As Integer
    Dim slSql As String
    
    mGetAndExportRegions = True
    lmRegionDef = 0
    lmRestrict = 0
    lmComment = 0
    lmISCI = 0
    '5/17/11: Moved below mRotPop so that lmAllCrfCode is build
    'If ckcRegionCopy.Value = vbUnchecked Then
    '    'lMsgBox "Region Export Not Requested", vbExclamation, ""
    '    gLogMsg "Region Export Not Requested", "ExportSQL.txt", False
    '    Exit Function
    'End If
    'lMsgBox "Region Export Started", vbExclamation, ""
    If ckcRegionCopy.Value = vbChecked Then
        gLogMsg "Region Export Started", "ExportSQL.txt", False
    End If
    mBuildLinkTables
    If imTerminate Then
        'lMsgBox "Export Terminated", vbExclamation, ""
        gLogMsg "Export Terminated", "ExportSQL.txt", False
        mGetAndExportRegions = False
        Exit Function
    End If
    mRotPop
    If imTerminate Then
        'lMsgBox "Export Terminated", vbExclamation, ""
        gLogMsg "Export Terminated", "ExportSQL.txt", False
        mGetAndExportRegions = False
        Exit Function
    End If
    
    If ckcRegionCopy.Value = vbUnchecked Then
        'lMsgBox "Region Export Not Requested", vbExclamation, ""
        gLogMsg "Region Export Not Requested", "ExportSQL.txt", False
        Exit Function
    End If
    
    For llRaf = 0 To UBound(lmRafCode) - 1 Step 1
        If imTerminate Then
            'lMsgBox "Export Terminated", vbExclamation, ""
            gLogMsg "Export Terminated", "ExportSQL.txt", False
            mGetAndExportRegions = False
            Exit Function
        End If
        tmRafSrchKey.lCode = lmRafCode(llRaf)
        ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            llRafCode = tmRaf.lCode
            llRegionCode = tmRaf.lRegionCode
            slRafName = Trim$(tmRaf.sName)
            ilRafAdfCode = tmRaf.iAdfCode
            
            'Doug:  Call Store Procedure at this point for Region Definition table:
                    'regCSIID=llRafCode, regRegionCode=llRegionCode, regName=slRafName, regAdvtCSIID=ilRafAdfCode
            slRafName = gFixQuote(slRafName)
            slSql = llRafCode & ", "
            slSql = slSql & llRegionCode & ", "
            slSql = slSql & "'" & slRafName & "', "
            slSql = slSql & ilRafAdfCode
            

            If Not mRunSP("sp_InsertRegion_Definition", slSql) Then
                'lMsgBox "ERROR: executing sp_InsertRegion_Definition.", vbCritical, "Error"
                gLogMsg "ERROR: executing sp_InsertRegion_Definition.", "ExportSQL.txt", False
                mGetAndExportRegions = False
                Exit Function
            End If
            lmRegionDef = lmRegionDef + 1
            
        Else
            imGenError = True
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            'lMsgBox "ERROR: Get Region Definition Failed: " & Str$(ilRet) & " RafCode: " & Str$(lmRafCode(llRaf)), vbExclamation, "Error"
            gLogMsg "ERROR: Get Region Definition Failed: " & Str$(ilRet) & " RafCode: " & Str$(lmRafCode(llRaf)), "ExportSQL.txt", False
            mGetAndExportRegions = False
        End If
    Next llRaf
    'lMsgBox CStr(lmRegionDef) & " Region Definitions Exported.", vbExclamation, ""
    gLogMsg CStr(lmRegionDef) & " Region Definitions Exported.", "ExportSQL.txt", False
    For llCrf = 0 To UBound(lmRafCrfCode) - 1 Step 1
        If imTerminate Then
            'lMsgBox "Export Terminated", vbExclamation, ""
            gLogMsg "Export Terminated", "ExportSQL.txt", False
            mGetAndExportRegions = False
            Exit Function
        End If
        tmCrfSrchKey.lCode = lmRafCrfCode(llCrf)
        ilRet = btrGetEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            ilRet = mCRFExport(True, True)
            If Not ilRet Then
                mGetAndExportRegions = False
            End If
        Else
            imGenError = True
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            'lMsgBox "ERROR: Get Copy Rotation for Region Failed: " & Str$(ilRet) & " CrfCode: " & Str$(lmRafCrfCode(llCrf)), vbExclamation, "Error"
            gLogMsg "ERROR: Get Copy Rotation for Region Failed: " & Str$(ilRet) & " CrfCode: " & Str$(lmRafCrfCode(llCrf)), "ExportSQL.txt", False
            mGetAndExportRegions = False
        End If
    Next llCrf
    
    'lMsgBox "Region Restrictions: " & CStr(lmRestrict) & ", Region Comments: " & CStr(lmComment) & ", ISCI Records: " & CStr(lmISCI) & " Exported, ", vbExclamation, ""
    gLogMsg "Region Restrictions: " & CStr(lmRestrict) & ", Region Comments: " & CStr(lmComment) & ", ISCI Records: " & CStr(lmISCI) & " Exported, ", "ExportSQL.txt", False
    'lMsgBox "Region Export Completed", vbExclamation, ""
    gLogMsg "Region Export Completed", "ExportSQL.txt", False
End Function

Private Function mCRFExport(ilIncludeComments As Integer, ilIncludeCopyPattern As Integer) As Integer
    Dim ilRet As Integer
    Dim llCrfCode As Long
    Dim llRafCode As Long
    Dim ilAdfCode As Integer
    Dim llCntrNo As Long
    Dim ilSoldVefCode As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slDays As String
    Dim ilDay As Integer
    Dim ilRotNo As Integer
    Dim ilSpotLen As Integer
    Dim llCsfCode As Long
    Dim slComment As String
    Dim ilTest As Integer
    Dim ilFound As Integer
    Dim ilTotalNoInstr As Integer
    '8/2/06: Add Cart number to export
    Dim slCartNo As String
    Dim slISCI As String
    Dim slProduct As String
    Dim slCreativeTitle As String
    Dim ilPercent As Integer
    Dim slSql As String
    
            
    mCRFExport = True
    lmRegionDef = 0
    llCrfCode = tmCrf.lCode
    
    
    llRafCode = tmCrf.lRafCode
    ilAdfCode = tmCrf.iAdfCode
    tmChfSrchKey.lCode = tmCrf.lChfCode
    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE Then
        llCntrNo = tmChf.lCntrNo
        ilSoldVefCode = tmCrf.iVefCode
        gUnpackDate tmCrf.iStartDate(0), tmCrf.iStartDate(1), slStartDate
        slStartDate = Format$(slStartDate, "yyyy-mm-dd")
        gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slEndDate
        If slEndDate = "" Then
            slEndDate = "12/31/2069"
        End If
        slEndDate = Format$(slEndDate, "yyyy-mm-dd")
        gUnpackTime tmCrf.iStartTime(0), tmCrf.iStartTime(1), "A", "1", slStartTime
        slStartTime = Format$(gConvertTime(slStartTime), "hh:mm:ss")
        gUnpackTime tmCrf.iEndTime(0), tmCrf.iEndTime(1), "A", "1", slEndTime
        slEndTime = Format$(gConvertTime(slEndTime), "hh:mm:ss")
        slDays = ""
        For ilDay = 0 To 6 Step 1
            slDays = slDays & tmCrf.sDay(ilDay)
        Next ilDay
        ilRotNo = tmCrf.iRotNo
        ilSpotLen = tmCrf.iLen
        If ilIncludeComments Then
            llCsfCode = tmCrf.lCsfCode
        Else
            llCsfCode = 0
        End If
        
        If llCsfCode > 0 Then
            tmCsf.sComment = ""
            imCsfRecLen = Len(tmCsf) '5011
            tmCsfSrchKey.lCode = llCsfCode
            ilRet = btrGetEqual(hmCsf, tmCsf, imCsfRecLen, tmCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                slComment = gStripChr0(tmCsf.sComment)
                If slComment <> "" Then
                'If tmCsf.iStrLen > 0 Then
                    'slComment = Trim$(Left$(tmCsf.sComment, tmCsf.iStrLen))
                    slComment = gFixQuote(slComment)
                    'Doug:  Call Stored Procedure at this point for Comment table:
                            'commCSIID=llCsfCode, commComment=slComment
                    slSql = llCsfCode & ", "
                    slSql = slSql & "'" & slComment & "'"
                                
                    If Not mRunSP("sp_InsertComment", slSql) Then
                        'lMsgBox "ERROR: executing sp_InsertComment.", vbCritical, "Error"
                        gLogMsg "ERROR: executing sp_InsertComment.", "ExportSQL.txt", False
                        Exit Function
                    End If
                    lmComment = lmComment + 1
                End If
            Else
                llCsfCode = 0
            End If
            
        End If
        
        'Doug:  Call Stored Procedure at this point for Rotation Restrictions table:
                'rotCSIID=llCrfCode, rotRegCSIID=llRafCode, rotAdvtCSIID=ilAdfCode, rotCntrNo=llCntrNo,
                'rotVehCSIID=ilSoldVefCode, rotStartDate=slStartDate, rotEndDate=slEndDate,
                'rotStartTime=slStartTime, rotEndTime=slEndTime, rotDays=slDays, rotRotationNo=ilRotNo,
                'rotSpotLength=ilSpotLen, rotCommCSIID=llCsfCode
            
            
        slSql = llCrfCode & ", "
        slSql = slSql & llRafCode & ", "
        slSql = slSql & ilAdfCode & ", "
        slSql = slSql & llCntrNo & ", "
        slSql = slSql & ilSoldVefCode & ", "
        slSql = slSql & "'" & slStartDate & "', "
        slSql = slSql & "'" & slEndDate & "', "
        slSql = slSql & "'" & slStartTime & "', "
        slSql = slSql & "'" & slEndTime & "', "
        slSql = slSql & "'" & slDays & "', "
        slSql = slSql & ilRotNo & ", "
        slSql = slSql & ilSpotLen & ", "
        slSql = slSql & llCsfCode
        
        'Remove restraint to avoid this condition. There does not have to be a comment
        'associated with each rotation.  If we get here from Spots there will not be a Regional Definition
        'because the Rotation Restrictions will be for the non-regional assigned copy.
        
'        If llCsfCode = 0 Then
'            slSQL = slSQL & "NULL"
'        Else
'            slSQL = slSQL & llCsfCode
'        End If
        

        If Not mRunSP("sp_InsertRotation_Restrictions", slSql) Then
            'lMsgBox "ERROR: executing sp_InsertRotation_Restrictions.", vbCritical, "Error"
            gLogMsg "ERROR: executing sp_InsertRotation_Restrictions.", "ExportSQL.txt", False
            Exit Function
        End If
        lmRestrict = lmRestrict + 1
        If ilIncludeCopyPattern Then
            ReDim tmCnfRot(0 To 0) As CNF
            tmCnfSrchKey.lCrfCode = tmCrf.lCode
            tmCnfSrchKey.iInstrNo = 0
            ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
                ilFound = False
                For ilTest = 0 To UBound(tmCnfRot) - 1 Step 1
                    If tmCnf.lCifCode = tmCnfRot(ilTest).lCifCode Then
                        ilFound = True
                        tmCnfRot(ilTest).iInstrNo = tmCnfRot(ilTest).iInstrNo + 1
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    tmCnfRot(UBound(tmCnfRot)) = tmCnf
                    tmCnfRot(UBound(tmCnfRot)).iInstrNo = 1
                    ReDim Preserve tmCnfRot(0 To UBound(tmCnfRot) + 1) As CNF
                End If
                ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            ilTotalNoInstr = 0
            For ilTest = 0 To UBound(tmCnfRot) - 1 Step 1
                ilTotalNoInstr = ilTotalNoInstr + tmCnfRot(ilTest).iInstrNo
            Next ilTest
            For ilTest = 0 To UBound(tmCnfRot) - 1 Step 1
                tmCnf = tmCnfRot(ilTest)
                tmCifSrchKey.lCode = tmCnf.lCifCode
                ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    tmCpfSrchKey.lCode = tmCif.lcpfCode
                    ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        llCrfCode = tmCrf.lCode
                        '8/2/06:  Get cart (Media code plus cart number)
                        slCartNo = mGetCartNo()
                        slISCI = Trim$(tmCpf.sISCI)
                        slISCI = gFixQuote(slISCI)
                        slProduct = Trim$(tmCpf.sName)
                        slProduct = gFixQuote(slProduct)
                        slCreativeTitle = Trim$(tmCpf.sCreative)
                        slCreativeTitle = gFixQuote(slCreativeTitle)
                        ilSpotLen = tmCif.iLen
                        ilPercent = Val(gDivStr(gMulStr(Trim$(Str$(tmCnf.iInstrNo)), "100"), Str$(ilTotalNoInstr)))
                        'Doug:  Call Stored Procedure at this point for region_isci_pattern:
                                'copyRotCSIID=llCrfCode, copyISCI=slISCI, copyProduct=slProduct,
                                'copyCreative=slCreativeTitle, copySpotLength=ilSpotLen, copyPercent=ilPercent
 '8/2/06-Doug:  You need to add slCartNo to the stored prcedure call
                    
                        slSql = llCrfCode & ", "
                        slSql = slSql & "'" & slISCI & "', "
                        slSql = slSql & "'" & slProduct & "', "
                        slSql = slSql & "'" & slCreativeTitle & "', "
                        slSql = slSql & ilSpotLen & ", "
                        slSql = slSql & ilPercent & ", "
                        slSql = slSql & "'" & slCartNo & "'"

                        If Not mRunSP("sp_InsertRegion_ISCI_Pattern", slSql) Then
                            'lMsgBox "ERROR: executing sp_InsertRegion_ISCI_Pattern.", vbCritical, "Error"
                            gLogMsg "ERROR: executing sp_InsertRegion_ISCI_Pattern.", "ExportSQL.txt", False
                            Exit Function
                        End If
                        lmISCI = lmISCI + 1
                    Else
                        imGenError = True
                        If ilRet >= 30000 Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        'lMsgBox "ERROR: Get Copy Product/ISCI Failed: " & Str$(ilRet) & " CpfCode: " & Str$(tmCif.lCpfCode), vbExclamation, "Error"
                        gLogMsg "ERROR: Get Copy Product/ISCI Failed: " & Str$(ilRet) & " CpfCode: " & Str$(tmCif.lcpfCode), "ExportSQL.txt", False
                        mCRFExport = False
                        Exit Function
                    End If
                End If
            Next ilTest
        End If
    Else
        imGenError = True
        If ilRet >= 30000 Then
            ilRet = csiHandleValue(0, 7)
        End If
        'lMsgBox "ERROR: Get Contract Failed for Copy Rotation: " & Str$(ilRet) & " CrfCode/ChfCode: " & Str$(tmCrf.lcode) & "/" & Str$(tmCrf.lChfCode), vbExclamation, "Error"
        gLogMsg "ERROR: Get Contract Failed for Copy Rotation: " & Str$(ilRet) & " CrfCode/ChfCode: " & Str$(tmCrf.lCode) & "/" & Str$(tmCrf.lChfCode), "ExportSQL.txt", False
        mCRFExport = False
        Exit Function
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mGenLog                         *
'*                                                     *
'*      This code was taken from Traffic mGenLog       *
'*******************************************************
Private Function mGenLog() As Integer
    Dim ilVpfIndex As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slSuEndDate As String
    'D.S. 06/13/13
    Dim slMoStartDate As String
    Dim ilGenLST As Integer
    Dim ilODFVefCode As Integer
    Dim ilType As Integer
    Dim sLCP As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilExportType As Integer
    Dim llLoop As Long
    Dim ilVef As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim llCheck As Long
    Dim llCrf As Long
    Dim slCBlackoutLog As String
    ReDim ilEvtAllowed(0 To 14) As Integer
    
    ReDim lmRafCrfCode(0 To 0) As Long
    mGenLog = True
    ilType = 0
    sLCP = "C"
    slStartTime = "12M"
    slEndTime = "12M"
    For llLoop = LBound(ilEvtAllowed) To UBound(ilEvtAllowed) Step 1
        ilEvtAllowed(llLoop) = True
    Next llLoop
    ilEvtAllowed(0) = False      'Must have 1 as true to get output
    'ilEvtAllowed(2) = True
    ilVpfIndex = gVpfFind(ExportSQL, tmVef.iCode)
    ilExportType = 0
    slStartDate = Format$(lmInputStartDate, "m/d/yy")
    ilRet = gReadBofRec(1, hmBof, hmCif, hmPrf, hmSif, hmCHF, "B", slStartDate, 1)
    igStartBofIndex = LBound(tgRBofRec) - 1
    ig30StartBofIndex = igStartBofIndex
    ig60StartBofIndex = igStartBofIndex
    ReDim tgSpotSum(0 To 0) As SPOTSUM
    '3/13/15
    ReDim tgOdfSdfCodes(0 To 0) As ODFSDFCODES
    lgStartIndex = UBound(tgSpotSum)
    sgGenDate = Format$(gNow(), "m/d/yy")
    sgGenTime = Format$(gNow(), "h:mm:ssAM/PM")
    gPackDate sgGenDate, igGenDate(0), igGenDate(1)
    gPackTime sgGenTime, igGenTime(0), igGenTime(1)
    '10-9-01
    gUnpackTimeLong igGenTime(0), igGenTime(1), False, lgGenTime
    slStartDate = Format$(lmInputStartDate, "m/d/yy")
    slEndDate = Format$(lmInputEndDate, "m/d/yy")
    slSuEndDate = gObtainNextSunday(slEndDate)
    'D.S. 06/13/13
    slMoStartDate = gObtainPrevMonday(slStartDate)
    
    gBuildExportDates hmDlf, hmPff, tmVef.iCode, "D", lmInputStartDate, 1, tmPrefeedExpt()
    
    ilGenLST = False
    imATTExist = False
    ilODFVefCode = 0
    smLogType = "P"

    '6/5/14: Temporarily set to Y so that tgSpotSum will be created
    slCBlackoutLog = tgSpf.sCBlackoutLog
    tgSpf.sCBlackoutLog = "Y"

    If tmVef.sType = "L" Then
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If (tgMVef(ilVef).sType = "C") And (tgMVef(ilVef).iVefCode = tmVef.iCode) Then
                'ilRet = gBuildODFSpotDay("L", ilType, sLCP, tgMVef(ilVef).iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, 0)
                'D.S. 06/13/13
                'ilRet = gBuildODFSpotDay("L", ilType, sLCP, tgMVef(ilVef).iCode, slStartDate, slSuEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, 0)
                ilRet = gBuildODFSpotDay("L", ilType, sLCP, tgMVef(ilVef).iCode, slMoStartDate, slSuEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, 0)

                If Not ilRet Then
                    'lMsgBox "mGenLog failed at tmVef.sType = L", vbCritical, "Error"
                    gLogMsg "mGenLog failed at tmVef.sType = L", "ExportSQL.txt", False
                    mGenLog = False
                    Exit Function
                End If
            End If
        Next ilVef
    ElseIf tmVef.sType = "A" Then
        'ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmVef.iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, 0)
        'D.S. 06/13/13
        'ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmVef.iCode, slStartDate, slSuEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, 0)
        ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmVef.iCode, slMoStartDate, slSuEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, 0)

        If Not ilRet Then
            'lMsgBox "mGenLog failed at tmVef.sType = A", vbCritical, "Error"
            gLogMsg "mGenLog failed at tmVef.sType = A", "ExportSQL.txt", False
            mGenLog = False
            Exit Function
        End If
    Else
        'ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmVef.iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, 0)
        'D.S. 06/13/13
        'ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmVef.iCode, slStartDate, slSuEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, 0)
        ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmVef.iCode, slMoStartDate, slSuEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, smLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, "L", 0, 0, 0)

        If Not ilRet Then
            'lMsgBox "mGenLog failed at tmVef.sType NOT L or A", vbCritical, "Error"
            gLogMsg "mGenLog failed at tmVef.sType NOT L or A", "ExportSQL.txt", False
            mGenLog = False
            Exit Function
        End If
    End If
     
    '6/5/14: reset
    tgSpf.sCBlackoutLog = slCBlackoutLog
    
    ilRet = mBuildBlackouts()
    'Read ODF and create spots
    ilRet = mODFExport()
    gDeleteOdf "G", ilType, sLCP, tmVef.iVefCode
    For llLoop = LBound(tgSpotSum) To UBound(tgSpotSum) - 1 Step 1
        If imTerminate Then
            'lMsgBox "Export Terminated", vbExclamation, ""
            gLogMsg "Export Terminated", "ExportSQL.txt", False
            mGenLog = False
            Exit Function
        End If
        If tgSpotSum(llLoop).lCrfCode > 0 Then
            '5/17/11: All Copy rotations for date span
            For llCrf = 0 To UBound(lmAllCrfCode) - 1 Step 1
                If tgSpotSum(llLoop).lCrfCode = lmAllCrfCode(llCrf) Then
                    lmAllCrfCode(llCrf) = -lmAllCrfCode(llCrf)
                    Exit For
                End If
            Next llCrf
            
            ilFound = False
            For llCheck = LBound(tgSpotSum) To llLoop - 1 Step 1
                If tgSpotSum(llLoop).lCrfCode = tgSpotSum(llCheck).lCrfCode Then
                    ilFound = True
                    Exit For
                End If
            Next llCheck
            If Not ilFound Then
                tmCrfSrchKey.lCode = tgSpotSum(llLoop).lCrfCode
                ilRet = btrGetEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    '8/2/06:  Include pattern
                    'ilRet = mCRFExport(False, False)
                    ilRet = mCRFExport(True, True)
                    If Not ilRet Then
                        mGenLog = False
                    End If
                Else
                    imGenError = True
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    'lMsgBox "ERROR: Get Copy Rotation: " & Str$(ilRet) & " CrfCode: " & Str$(tgSpotSum(llLoop).lCrfCode), vbExclamation, "Error"
                    gLogMsg "ERROR: Get Copy Rotation: " & Str$(ilRet) & " CrfCode: " & Str$(tgSpotSum(llLoop).lCrfCode), "ExportSQL.txt", False
                    mGenLog = False
                End If
            End If
        End If
    Next llLoop
End Function

Private Function mGetAndExportSpots() As Integer
    Dim ilVef As Integer
    Dim ilVefCode As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer
    
    mGetAndExportSpots = True
    If ckcSpots.Value = vbUnchecked Then
        'lMsgBox "Region Export Not Requested", vbExclamation, ""
        gLogMsg "Spot Export Not Requested", "ExportSQL.txt", False
        Exit Function
    End If
    'lMsgBox "Spot Export Selected Vehicles Started", vbExclamation, ""
    gLogMsg "Spot Export Selected Vehicles Started", "ExportSQL.txt", False
    For ilVef = 0 To lbcVehicle.ListCount - 1 Step 1
        If imTerminate Then
            'lMsgBox "Export Terminated", vbExclamation, ""
            gLogMsg "Export Terminated", "ExportSQL.txt", False
            mGetAndExportSpots = False
            Exit Function
        End If
        If lbcVehicle.Selected(ilVef) Then
            ilVefCode = Val(lbcVehicle.ItemData(ilVef))
            ilIndex = gBinarySearchVef(ilVefCode)
            If ilIndex <> -1 Then
                tmVef = tgMVef(ilIndex)
                'lMsgBox "Exporting " & Trim$(tgMVef(ilIndex).sName) & ", CSIID = " & ilVefCode, vbExclamation, ""
                gLogMsg Trim$(tgMVef(ilIndex).sName) & ", CSIID = " & ilVefCode, "ExportSQL.txt", False
                ilRet = mGenLog()
                If Not ilRet Then
                    mGetAndExportSpots = False
                End If
            End If
        End If
    Next ilVef
    If mGetAndExportSpots = False Then
        'lMsgBox "Spot Export Completed, But Errors Occured", vbExclamation, ""
        gLogMsg "Spot Export Completed, But Errors Occured", "ExportSQL.txt", False
    Else
        'lMsgBox "Spot Export Completed Successfully", vbExclamation, ""
        gLogMsg "Spot Export Completed Successfully", "ExportSQL.txt", False
    End If
    
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mBuildBlackouts                 *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:5/4/94       By:D. Hannifan     *
'*                                                     *
'*            Comments: Build ODF and suppress/replace *
'*                      as required
'*                                                     *
'*******************************************************
Private Function mBuildBlackouts() As Integer
    Dim ilRet As Integer
    If tgSpf.sCBlackoutLog = "Y" Then
        lgEndIndex = UBound(tgSpotSum)
        gBlackoutTest 1, hmCif, hmMcf, hmOdf, hmRsf, hmCpf, hmCrf, hmCnf, hmClf, hmLst, hmCvf, smNewLines(), hmMsg, lbcLogMsg
    End If
    mBuildBlackouts = True
    Exit Function
End Function

Private Function mODFExport() As Integer
    Dim ilRet As Integer
    Dim llSdfCode As Long
    Dim ilAdfCode As Integer
    Dim llCntrNo As Long
    Dim ilLineNo As Integer
    Dim ilAirVefCode As Integer
    Dim ilSoldVefCode As Integer
    Dim slAirDate As String
    Dim llAirDate As Long
    Dim slAirTime As String
    Dim llAirTime As Long
    Dim llLoop As Long
    Dim ilVef As Integer
    Dim slZone As String
    Dim ilSeqNo As Integer
    Dim llTime As Long
    Dim ilLnVefCode As Integer
    Dim ilVefCode As Integer
    Dim llDate As Long
    Dim llCrfCode As Long
    '8/2/06: Add Cart number to export
    Dim slCartNo As String
    Dim slISCI As String
    Dim slProduct As String
    Dim slShortTitle As String
    Dim slAvailTime As String
    Dim llAvailTime As Long
    Dim slAvailEventID As String
    Dim ilPositionNo As Integer
    Dim ilBreakNo As Integer
    Dim slSoldTime As String
    Dim ilRdfCode As Integer
    Dim ilRdf As Integer
    Dim ilPos As Integer
    Dim slOvTimes As String
    Dim slOvStartTime As String
    Dim slOvEndTime As String
    Dim slOvDays As String
    Dim slLocalNetwork As String
    Dim ilSpotLen As Integer
    Dim slOrigAirDate As String
    Dim slSql As String
    Dim llSpotCount As Long
    Dim ilPff As Integer
    Dim llToAirDate As Long
    Dim llToSpotTime As Long
    Dim llSpotTime As Long

    mODFExport = True
    llSpotCount = 0
    tmOdfSrchKey2.iGenDate(0) = igGenDate(0)   'ilLogDate0
    tmOdfSrchKey2.iGenDate(1) = igGenDate(1)
    tmOdfSrchKey2.lGenTime = lgGenTime
    ilRet = btrGetEqual(hmOdf, tmOdf, imOdfRecLen, tmOdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tmOdf.iGenDate(0) = igGenDate(0)) And (tmOdf.iGenDate(1) = igGenDate(1)) And (tmOdf.lGenTime = lgGenTime)
        If tmOdf.iType = 4 Then
            'Create Output
            gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), slAirDate
            llAirDate = gDateValue(slAirDate)
            gUnpackTime tmOdf.iLocalTime(0), tmOdf.iLocalTime(1), "A", "1", slAirTime
            llAirTime = gTimeToLong(slAirTime, False)
            For llLoop = LBound(tgSpotSum) To UBound(tgSpotSum) - 1 Step 1
                ilVefCode = tgSpotSum(llLoop).iVefCode
                llDate = tgSpotSum(llLoop).lDate
                slZone = tgSpotSum(llLoop).sZone
                ilSeqNo = tgSpotSum(llLoop).iSeqNo
                llTime = tgSpotSum(llLoop).lTime
                ilLnVefCode = tgSpotSum(llLoop).iLnVefCode
                If (tmOdf.iVefCode = ilVefCode) And (llDate = llAirDate) And (llTime = llAirTime) And (Trim$(slZone) = Trim$(tmOdf.sZone)) And (ilSeqNo = tmOdf.iSeqNo) Then
            
                    tmSdfSrchKey3.lCode = tgSpotSum(llLoop).lSdfCode
                    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        ilPff = LBound(tmPrefeedExpt)
                        llSdfCode = tgSpotSum(llLoop).lSdfCode
                        ilAdfCode = tmOdf.iAdfCode
                        llCntrNo = tmOdf.lCntrNo
                        ilLineNo = tmSdf.iLineNo
                        ilSpotLen = tmSdf.iLen
                        ilAirVefCode = tmVef.iCode
                        ilSoldVefCode = tgSpotSum(llLoop).iLnVefCode
                        gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), slAirDate
                        slAirDate = Format$(slAirDate, "yyyy-mm-dd")
                        gUnpackTime tmOdf.iLocalTime(0), tmOdf.iLocalTime(1), "A", "1", slAirTime
                        slAirTime = Format$(gConvertTime(slAirTime), "hh:mm:ss")
                        llCrfCode = tgSpotSum(llLoop).lCrfCode
                        '6/14/07:  add position and break
                        ilPositionNo = tmOdf.iPositionNo
                        ilBreakNo = tmOdf.iBreakNo
                        '8/2/06:  Initialize cart
                        slCartNo = ""
                        slISCI = ""
                        slProduct = tmOdf.sProduct
                        If tmOdf.lCifCode > 0 Then
                            tmCifSrchKey.lCode = tmOdf.lCifCode
                            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                '8/2/06:  Get cart (Media code plus cart number)
                                slCartNo = mGetCartNo()
                                tmCpfSrchKey.lCode = tmCif.lcpfCode
                                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    slISCI = Trim$(tmCpf.sISCI)
                                End If
                            End If
                        End If
                        slShortTitle = tmOdf.sShortTitle
                        gUnpackTime tmOdf.iAirTime(0), tmOdf.iAirTime(1), "A", "1", slAvailTime
                        slAvailTime = Format$(gConvertTime(slAvailTime), "hh:mm:ss")
                        gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slSoldTime
                        slSoldTime = Format$(gConvertTime(slSoldTime), "hh:mm:ss")
                        'Returns the Avail ID
                        tmCefSrchKey.lCode = tmOdf.lEvtIDCefCode
                        imCefRecLen = Len(tmCef)
                        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            'slAvailEventID = Trim$(Left$(tmCef.sComment, tmCef.iStrLen)) 'tlCef.sComment
                            slAvailEventID = gStripChr0(tmCef.sComment)
                        Else
                            slAvailEventID = ""
                        End If
                        slZone = Left$(tmOdf.sZone, 1)
                        tmClfSrchKey.lChfCode = tmSdf.lChfCode
                        tmClfSrchKey.iLine = tmSdf.iLineNo
                        tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                        tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                        Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And (tmClf.sSchStatus = "A")
                            ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        slOvStartTime = ""
                        slOvEndTime = ""
                        slOvDays = ""
                        ilRdfCode = 0
                        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
                            ilRdfCode = tmClf.iRdfCode
                            ilRdf = gBinarySearchRdf(ilRdfCode)
                            If ilRdf <> -1 Then
                                If StrComp(Trim$(tgMRdf(ilRdf).sName), Trim$(tmOdf.sDPDesc), vbTextCompare) <> 0 Then
                                    ilPos = InStr(1, tmOdf.sDPDesc, " ", vbTextCompare)
                                    If ilPos > 0 Then
                                        slOvTimes = Mid$(tmOdf.sDPDesc, ilPos + 1)
                                        slOvDays = tgSpotSum(llLoop).sDays
                                        If ((tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0)) And ((tmClf.iEndTime(0) <> 1) Or (tmClf.iEndTime(1) <> 0)) Then
                                            gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slOvStartTime
                                            gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slOvEndTime
                                            slOvStartTime = Format$(gConvertTime(slOvStartTime), "hh:mm:ss")
                                            slOvEndTime = Format$(gConvertTime(slOvEndTime), "hh:mm:ss")
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        slLocalNetwork = "L"
                        tmAnfSrchKey.iCode = tmOdf.ianfCode
                        ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If (ilRet = BTRV_ERR_NONE) Then
                            If Left$(tmAnf.sName, 1) = "N" Then
                                slLocalNetwork = "N"
                            End If
                        End If
                        gUnpackDate tgSpotSum(llLoop).iOrigAirDate(0), tgSpotSum(llLoop).iOrigAirDate(1), slOrigAirDate
                        slOrigAirDate = Format$(slOrigAirDate, "yyyy-mm-dd")
                        'Doug:  Call Store Procedure at this point for Airing Spots table:
                                
                                'spotCSIID=llSdfCode, spotAdsvtCSIID = ilAdfCode, spotCntrNo = llCntrNo, spotLineNo=ilLineNo
                                'spotLength = ilSpotLen, spotAirVehCSIID = ilAirVefCode, spotSoldVehCSIID = ilSoldVefCode
                                'spotDate = slAirDate, spotTime = slAirTime, spotRotCSIID = llCrfCode, spotISCI = slISCI
                                'spotProduct=slProduct, spotShortTitle=slShortTitle, spotSoldAvailTime = slSoldTime
                                'spotAvailEvtID = slAvailEventID, spotZone = slZone, spotDayCSIID=ilRdfCode,
                                'spotOvStartTime=slOvStartTime, spotOvEndTime = slOvEndTime, spotOvDays = slOvDays
                                'spotLocalNetwork=slLocalNetwork
 '8/2/06-Doug:  You need to add slCartNo to the stored prcedure call
 
 '6/14/07-Doug:  You will need to add ilPositionNo and ilBreakNo to the store procedure call
 '               Also, Jeff will have to add the fields to the sql server database
                        llSpotTime = gTimeToLong(slAirTime, False)
                        llToAirDate = gDateValue(slAirDate)
                        llToSpotTime = gTimeToLong(slAirTime, False)
                        Do
                            If (llToAirDate >= lmInputStartDate) And (llToAirDate <= lmInputEndDate) Then
                                slSql = llSdfCode & ", "
                                slSql = slSql & ilAdfCode & ", "
                                slSql = slSql & llCntrNo & ", "
                                slSql = slSql & ilLineNo & ", "
                                slSql = slSql & ilSpotLen & ", "
                                slSql = slSql & ilAirVefCode & ", "
                                slSql = slSql & ilSoldVefCode & ", "
                                slSql = slSql & "'" & slOrigAirDate & "', "
                                'slSQL = slSQL & "'" & slAirDate & "', "
                                'slSQL = slSQL & "'" & slAirTime & "', "
                                slSql = slSql & "'" & Format$(llToAirDate, "yyyy-mm-dd") & "', "
                                'slSQL = slSQL & "'" & Format$(gFormatTimeLong(llToSpotTime, "A", "1"), "hh:mm:ss") & "', "
                                If llToSpotTime = 0 Then
                                    slSql = slSql & "'" & "00:00:00" & "', "
                                Else
                                    slSql = slSql & "'" & Format$(gFormatTimeLong(llToSpotTime, "A", "1"), "hh:mm:ss") & "', "
                                End If
                                slSql = slSql & llCrfCode & ", "
                                slSql = slSql & "'" & gFixQuote(slISCI) & "', "
                                slSql = slSql & "'" & gFixQuote(slProduct) & "', "
                                slSql = slSql & "'" & gFixQuote(slShortTitle) & "', "
                                slSql = slSql & "'" & slSoldTime & "', "
                                slSql = slSql & "'" & gFixQuote(slAvailEventID) & "', "
                                slSql = slSql & "'" & slZone & "', "
                                slSql = slSql & ilRdfCode & ", "
                                'slSQL = slSQL & "'" & slOvStartTime & "', "
                                If UCase(slOvStartTime) = "12M" Then
                                    slSql = slSql & "'" & "00:00:00" & "', "
                                Else
                                    slSql = slSql & "'" & slOvStartTime & "', "
                                End If
                                'slSQL = slSQL & "'" & slOvEndTime & "', "
                                If UCase(slOvEndTime) = "12M" Then
                                    slSql = slSql & "'" & "00:00:00" & "', "
                                Else
                                    slSql = slSql & "'" & slOvEndTime & "', "
                                End If
                                slSql = slSql & "'" & slOvDays & "', "
                                slSql = slSql & "'" & gFixQuote(slLocalNetwork) & "', "
                                slSql = slSql & "'" & slCartNo & "', "
                                slSql = slSql & "'" & ilPositionNo & "', "
                                slSql = slSql & "'" & ilBreakNo & "'"
                                
                                
                                'Debug Only
                                'lMsgBox "About to try to insert: " & slSQL, vbCritical, "Error"
                                'gLogMsg "About to try to insert: " & slSQL, "ExportSQL.txt", False
        
                                If Not mRunSP("sp_InsertAiringSpots", slSql) Then
                                    'lMsgBox "ERROR: executing sp_Airing_Spots.", vbCritical, "Error"
                                    gLogMsg "ERROR: executing sp_InsertAiringSpots.", "ExportSQL.txt", False
                                    mODFExport = False
                                    Exit Function
                                End If
                                llSpotCount = llSpotCount + 1
                            End If
                            Do
                                If Not mPreFeed(llAirDate, llSpotTime, ilPff, tmOdf.sZone, llToAirDate, llToSpotTime) Then
                                    Exit For
                                Else
                                    Exit Do
                                End If
                            Loop
                        Loop
                    Else
                        imGenError = True
                        If ilRet >= 30000 Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        'lMsgBox "ERROR: Get Spot: " & Str$(ilRet) & " SdfCode: " & Str$(tgSpotSum(llLoop).lSdfCode), vbExclamation, "Error"
                        gLogMsg "ERROR: Get Spot: " & Str$(ilRet) & " SdfCode: " & Str$(tgSpotSum(llLoop).lSdfCode), "ExportSQL.txt", False
                        mODFExport = False
                    End If
                    Exit For
                End If
            Next llLoop
        End If
        ilRet = btrGetNext(hmOdf, tmOdf, imOdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
    Loop
    gLogMsg "   " & CStr(llSpotCount) & " Spots exported", "ExportSQL.txt", False
    'did this vehicle have any spots to export?
    If llSpotCount > 0 Then
        lmTotalSomeSpots = lmTotalSomeSpots + 1
    Else
        lmTotalZeroSpots = lmTotalZeroSpots + 1
    End If
    lmTotalSpots = lmTotalSpots + llSpotCount
End Function

Public Sub CloseDB()
    oADOConn.Close
    Set oADOConn = Nothing
End Sub

Public Function ConnectToDB(sConnectionStr As String) As Boolean
    On Error GoTo ErrHand
    ConnectToDB = False
    Set oADOConn = CreateObject("ADODB.Connection")
    oADOConn.Open sConnectionStr
    ConnectToDB = True
    Exit Function

ErrHand:
    Call LogADOError
    Exit Function
End Function

Public Function mRunSP(StoredProcedureName As String, SQLStatement As String) As Boolean
    Dim slSql As String
    

    On Error GoTo ErrHand1
    mRunSP = False
    If (imUserInput = False) Or (ckcSQLServer.Value = vbChecked) Then
        slSql = StoredProcedureName & " " & SQLStatement
        oADOConn.Execute slSql
    End If
    On Error GoTo ErrHand2
    If (imUserInput = True) And (ckcFile.Value = vbChecked) Then
        gLogMsgWODT "W", hmFile, SQLStatement
    End If
    mRunSP = True
    Exit Function

ErrHand1:
    Call LogADOError
    Exit Function
ErrHand2:
    Exit Function
End Function

Public Sub LogADOError()
On Error Resume Next
Dim slErrMsg As String
Dim errLoop As ADODB.Error

    slErrMsg = ""
    ' Loop through each Error object in Errors collection.
    For Each errLoop In oADOConn.Errors

        slErrMsg = slErrMsg & "Desc:" & errLoop.Description & "Err#:" & errLoop.NativeError
        ' slErrMsg = slErrMsg & vbCrLf & vbCrLf
    Next
'    strError(0) = "Error Number: " & errLoop.Number
'    strError(1) = "  Description: " & errLoop.Description
'    strError(2) = "  Source: " & errLoop.Source
'    strError(3) = "  SQL State: " & errLoop.SQLState
'    strError(4) = "  Native Error: " & errLoop.NativeError

    'lMsgBox "ERROR: " & slErrMsg, vbExclamation, ""
    gLogMsg "ERROR: " & slErrMsg, "ExportSQL.txt", False
    oADOConn.Errors.Clear
End Sub

'***************************************************************************************
'*
'* Procedure Name: gLoadOption
'*
'* Created: 8/22/03 - J. Dutschke
'*
'* Modified:              By:
'*
'* Comments: This function loads a string value from the ini file.
'*           It relies on the global variable sgIniPathFileName to
'*           contain the path and name of the ini file to use.
'*
'***************************************************************************************
Public Function gLoadOption(Section As String, Key As String, sValue As String) As Boolean
    On Error GoTo ERR_gLoadOption
    Dim BytesCopied As Integer
    Dim sBuffer As String * 128

    gLoadOption = False
    BytesCopied = GetPrivateProfileString(Section, Key, "Not Found", sBuffer, 128, smIniPathFileName)
    If BytesCopied > 0 Then
        If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
            sValue = Left(sBuffer, BytesCopied)
            gLoadOption = True
        End If
    End If
    Exit Function

ERR_gLoadOption:
    ' return now if an error occurs
    'lMsgBox "ERROR: Loading INI Inforamtion", vbExclamation, ""
    gLogMsg "ERROR: Loading INI Inforamtion", "ExportSQL.txt", False
End Function


'***************************************************************************************
'*
'* Function Name: lLogMsg
'*
'* Created: August, 2004  By: D. Smith
'*
'* Modified:              By:
'*
'* Comments: A general file routine that shows: Date and Time followed by a message
'*
'***************************************************************************************

'Public Sub lLogMsg(sMsg As String, sFileName As String, iKill As Integer)
'
'    'Params
'    'sMsg is the string to be written out
'    'sFileName is the name of the file to be written to in the Messages directory
'    'iKill = True then delete the file first, iKill = False then append to the file
'
'    Dim slFullMsg As String
'    Dim hlLogFile As Integer
'    Dim ilRet As Integer
'    Dim slDateTime As String
'    Dim slToFile As String
'
'    slToFile = "D:\CSI\" & sFileName
'    On Error GoTo Error
'
'    If iKill = True Then
'        ilRet = 0
'        slDateTime = FileDateTime(slToFile)
'        If ilRet = 0 Then
'            Kill slToFile
'        End If
'    End If
'
'    hlLogFile = FreeFile
'    Open slToFile For Append As hlLogFile
'    slFullMsg = Format$(Now, "mm/dd/yyyy") & " " & Format$(Now, "hh:mm:ssam/pm") & " " & sMsg
'    Print #hlLogFile, slFullMsg
'    Close hlLogFile
'    Exit Sub
'
'Error:
'    ilRet = 1
'    Resume Next
'
'End Sub

'8/2/06:   Get cart number from tmCif
Private Function mGetCartNo() As String
    Dim slCartNo As String
    Dim ilRet As Integer
    
    slCartNo = ""
    If (tmMcf.iCode <> tmCif.iMcfCode) And (tmCif.iMcfCode > 0) Then
        tmMcfSrchKey.iCode = tmCif.iMcfCode
        ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmMcf.sName = "C"
            tmMcf.sPrefix = "C"
        End If
    End If
    If Trim$(tmCif.sCut) = "" Then
        slCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName)
    Else
        slCartNo = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut)
    End If
    mGetCartNo = slCartNo
End Function

Private Function mCheckRecLength() As Integer
    Dim ilSize As Integer
    
    If Not gRecLengthOk("Chf.Btr", Len(tmChf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Clf.Btr", Len(tmClf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    
    If Not gRecLengthOk("Sdf.Btr", Len(tmSdf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Crf.Btr", Len(tmCrf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    
    If Not gRecLengthOk("Cpf.Btr", Len(tmCpf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Cnf.Btr", Len(tmCnf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Cif.Btr", Len(tmCif)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Raf.Btr", Len(tmRaf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    ilSize = Len(tmCsf)
    'ilSize = ilSize - Len(tmCsf.iStrLen) - Len(tmCsf.sComment)
    If Not gRecLengthOk("Csf.Btr", ilSize) Then
        mCheckRecLength = False
        Exit Function
    End If
    ilSize = Len(tmCef)
    'ilSize = ilSize - Len(tmCef.iStrLen) - Len(tmCef.sComment)
    If Not gRecLengthOk("Cef.Btr", ilSize) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Odf.Btr", Len(tmOdf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Anf.Btr", Len(tmAnf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Mcf.Btr", Len(tmMcf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    
    If Not gRecLengthOk("Vef.Btr", Len(tmVef)) Then
        mCheckRecLength = False
        Exit Function
    End If
    mCheckRecLength = True
End Function

Private Function mPreFeed(llFromAirDate As Long, llFromSpotTime As Long, ilPrefeed As Integer, slZone As String, llToAirDate As Long, llToSpotTime As Long) As Integer
    Dim ilLoop As Integer
    Dim ilDay As Integer
    
    mPreFeed = True
    ilDay = gWeekDayLong(llFromAirDate)
    For ilLoop = ilPrefeed To UBound(tmPrefeedExpt) - 1 Step 1
        If tmPrefeedExpt(ilLoop).iFDay <> -1 Then
            If tmPrefeedExpt(ilLoop).iFDay = ilDay Then
                If (llFromSpotTime >= tmPrefeedExpt(ilLoop).lFStartTime) And (llFromSpotTime <= tmPrefeedExpt(ilLoop).lFEndTime) Then
                    If (Left$(slZone, 1) = tmPrefeedExpt(ilLoop).sFZone) Or (tmPrefeedExpt(ilLoop).sFZone = "A") Then
                        'llToAirDate = llFromAirDate - (tmPrefeedExpt(ilLoop).iFDay - tmPrefeedExpt(ilLoop).iTDay)
                        llToAirDate = llFromAirDate - (gWeekDayLong(llFromAirDate) - tmPrefeedExpt(ilLoop).iTDay)
                        llToSpotTime = llFromSpotTime + tmPrefeedExpt(ilLoop).lAdjTime
                        If llToSpotTime < 0 Then
                            llToAirDate = llToAirDate - 1
                            llToSpotTime = 86400 + llToSpotTime
                        ElseIf llToSpotTime > 86400 Then
                            llToAirDate = llToAirDate + 1
                            llToSpotTime = llToSpotTime - 86400
                        End If
                        ilPrefeed = ilLoop + 1
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ilLoop
    mPreFeed = False
End Function

Private Function mVerifyINIValues() As Integer
    mVerifyINIValues = False
    'Debug
    lLogMsg "INI location = " & smIniPathFileName, "ExportSQL.Txt", False
    
    If Not gLoadOption("Locations", "DSN", smDSN) Then
        'lMsgBox "ERROR - ExportSQL.ini [Locations] 'DSN' key is missing.", vbCritical, "Error"
        gLogMsg "ERROR - ExportSQL.ini [Locations] 'DSN' key is missing.", "ExportSQL.txt", False
        If igBkgdProg = 0 Then
            MsgBox "ERROR - ExportSQL.ini [Locations] 'DSN' key is missing.", vbCritical + vbOKOnly
        End If
        Exit Function
    End If
    'Debug
    lLogMsg "DSN = " & smDSN, "ExportSQL.Txt", False
    If Not gLoadOption("Locations", "UID", smUID) Then
        'lMsgBox "ERROR: ExportSQL.ini [Locations] 'UID' key is missing.", vbCritical, "Error"
        gLogMsg "ERROR: ExportSQL.ini [Locations] 'UID' key is missing.", "ExportSQL.txt", False
        If igBkgdProg = 0 Then
            MsgBox "ERROR: ExportSQL.ini [Locations] 'UID' key is missing.", vbCritical + vbOKOnly
        End If
        Exit Function
    End If
    'Debug
    lLogMsg "UID = " & smUID, "ExportSQL.Txt", False
    If Not gLoadOption("Locations", "PWD", smPWD) Then
        'lMsgBox "ERROR: ExportSQL.ini [Locations] 'PWD' key is missing.", vbCritical, "Error"
        gLogMsg "ERROR: ExportSQL.ini [Locations] 'PWD' key is missing.", "ExportSQL.txt", False
        If igBkgdProg = 0 Then
            MsgBox "ERROR: ExportSQL.ini [Locations] 'PWD' key is missing.", vbCritical + vbOKOnly
        End If
        Exit Function
    End If
    'Debug
    lLogMsg "PWD = " & smPWD, "ExportSQL.Txt", False
    If smPWD = "BLANK" Then
        smPWD = ""
    End If
    mVerifyINIValues = True
End Function


Private Sub mCloseFile()
    If imUserInput = True Then
        If ckcFile.Value = vbChecked Then
            gLogMsgWODT "C", hmFile, ""
        End If
    End If
End Sub

'5/17/11: All Copy rotations for date span
Private Function mExportAllCrf(llCount As Long) As Integer
    Dim llCrf As Long
    Dim ilRet As Integer
    
    mExportAllCrf = True
    llCount = 0
    For llCrf = 0 To UBound(lmAllCrfCode) - 1 Step 1
        If imTerminate Then
            'lMsgBox "Export Terminated", vbExclamation, ""
            gLogMsg "Export Terminated", "ExportSQL.txt", False
            mExportAllCrf = False
            Exit Function
        End If
        If lmAllCrfCode(llCrf) > 0 Then
            tmCrfSrchKey.lCode = lmAllCrfCode(llCrf)
            ilRet = btrGetEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                ilRet = mCRFExport(True, True)
                If Not ilRet Then
                    mExportAllCrf = False
                End If
                llCount = llCount + 1
            Else
                imGenError = True
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                'lMsgBox "ERROR: Get Copy Rotation for Region Failed: " & Str$(ilRet) & " CrfCode: " & Str$(lmRafCrfCode(llCrf)), vbExclamation, "Error"
                gLogMsg "ERROR: Get Copy Rotation not tied with Spot Failed: " & Str$(ilRet) & " CrfCode: " & Str$(lmAllCrfCode(llCrf)), "ExportSQL.txt", False
                mExportAllCrf = False
            End If
        End If
    Next llCrf
    gLogMsg "Export Copy Rotations not tied with Spots Completed", "ExportSQL.txt", False
End Function

