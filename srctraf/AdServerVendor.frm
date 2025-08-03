VERSION 5.00
Begin VB.Form AdServerVendor 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6750
   ClientLeft      =   4185
   ClientTop       =   3630
   ClientWidth     =   10470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FF0000&
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
   ScaleHeight     =   6750
   ScaleWidth      =   10470
   Begin VB.Frame frcBasic 
      Caption         =   "Basic Information"
      Height          =   1095
      Left            =   5400
      TabIndex        =   34
      Top             =   960
      Width           =   4215
      Begin VB.Frame frcPriority 
         Height          =   375
         Left            =   0
         TabIndex        =   40
         Top             =   700
         Width           =   4215
         Begin VB.TextBox txtPriority 
            Height          =   285
            Index           =   1
            Left            =   3120
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox txtPriority 
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lbcPriorityEnd 
            Caption         =   "End"
            Height          =   255
            Left            =   2500
            TabIndex        =   20
            Top             =   0
            Width           =   495
         End
         Begin VB.Label lbcPriority 
            Caption         =   "Priority Start"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.TextBox txtVendorName 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   300
         Width           =   2175
      End
      Begin VB.Label lbcVendorName 
         Caption         =   "Vendor Name"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame frcBucket 
      Caption         =   "Amazon Bucket Information"
      Height          =   3735
      Left            =   240
      TabIndex        =   28
      Top             =   2040
      Width           =   4215
      Begin VB.TextBox txtBucket 
         Height          =   525
         Index           =   4
         Left            =   1240
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   12
         Text            =   "AdServerVendor.frx":0000
         Top             =   2925
         Width           =   2535
      End
      Begin VB.TextBox txtBucket 
         Height          =   525
         Index           =   3
         Left            =   1240
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   11
         Text            =   "AdServerVendor.frx":0006
         Top             =   2270
         Width           =   2535
      End
      Begin VB.TextBox txtBucket 
         Height          =   525
         Index           =   2
         Left            =   1240
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   10
         Text            =   "AdServerVendor.frx":000C
         Top             =   1615
         Width           =   2535
      End
      Begin VB.TextBox txtBucket 
         Height          =   525
         Index           =   1
         Left            =   1240
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   9
         Text            =   "AdServerVendor.frx":0012
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtBucket 
         Height          =   525
         Index           =   0
         Left            =   1240
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   8
         Text            =   "AdServerVendor.frx":0018
         Top             =   300
         Width           =   2535
      End
      Begin VB.Label lbcBucketPrivate 
         Caption         =   "Private Key"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2925
         Width           =   1215
      End
      Begin VB.Label lbcBucketName 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   300
         Width           =   975
      End
      Begin VB.Label lbcBucketFolder 
         Caption         =   "Folder"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lbcBucketAccess 
         Caption         =   "Access Key"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2270
         Width           =   1215
      End
      Begin VB.Label lbcBucketRegion 
         Caption         =   "Region"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1615
         Width           =   975
      End
   End
   Begin VB.Frame frcAPI 
      Caption         =   "API Information"
      Height          =   3735
      Left            =   5400
      TabIndex        =   25
      Top             =   2040
      Width           =   4215
      Begin VB.TextBox txtAPI 
         Height          =   525
         Index           =   1
         Left            =   1240
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   5
         Text            =   "AdServerVendor.frx":001E
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtAPI 
         Height          =   525
         Index           =   0
         Left            =   1240
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   4
         Text            =   "AdServerVendor.frx":0024
         Top             =   300
         Width           =   2535
      End
      Begin VB.Frame frcVendor 
         Height          =   2175
         Left            =   0
         TabIndex        =   37
         Top             =   1560
         Width           =   4215
         Begin VB.ComboBox cboTimeZone 
            Height          =   315
            ItemData        =   "AdServerVendor.frx":002A
            Left            =   1200
            List            =   "AdServerVendor.frx":002C
            TabIndex        =   42
            Top             =   1490
            Width           =   1455
         End
         Begin VB.TextBox txtAPI 
            Height          =   525
            Index           =   3
            Left            =   1240
            MultiLine       =   -1  'True
            ScrollBars      =   1  'Horizontal
            TabIndex        =   7
            Text            =   "AdServerVendor.frx":002E
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox txtAPI 
            Height          =   525
            Index           =   2
            Left            =   1240
            MultiLine       =   -1  'True
            ScrollBars      =   1  'Horizontal
            TabIndex        =   6
            Text            =   "AdServerVendor.frx":0034
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label lbcTimeZone 
            Caption         =   "Zone"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label lbcOrganizationID 
            Caption         =   "Org ID"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lbcURL 
            Caption         =   "URL"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label lbcPassword 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbcUserName 
         Caption         =   "User Name"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.OptionButton rbcStatus 
      Caption         =   "Dormant"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.OptionButton rbcStatus 
      Caption         =   "Active"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmcPosition 
      Appearance      =   0  'Flat
      Caption         =   "&Position"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6465
      TabIndex        =   16
      Top             =   6135
      Width           =   945
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   5205
      TabIndex        =   15
      Top             =   6120
      Width           =   945
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Left            =   11640
      Top             =   240
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   3960
      TabIndex        =   14
      Top             =   6120
      Width           =   945
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Done"
      Height          =   285
      Left            =   2745
      TabIndex        =   13
      Top             =   6120
      Width           =   945
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      ItemData        =   "AdServerVendor.frx":003A
      Left            =   6600
      List            =   "AdServerVendor.frx":003C
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   3000
   End
   Begin VB.Label lbcCode 
      Caption         =   "0"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   9840
      TabIndex        =   36
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lbcImpressionMethod 
      Caption         =   "Label4"
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lbcContractMethod 
      Caption         =   "Label4"
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Impressions"
      Height          =   255
      Left            =   480
      TabIndex        =   24
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Contracts"
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Status"
      Height          =   255
      Left            =   480
      TabIndex        =   22
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "AdServerVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const NODATE As String = "01/01/1970"
Private Const METHODNONE = "N/A"
Private Const METHODMANUAL = "Manual"
Private Const METHODVENDOR = "Vendor API"
Private Const METHODCSI = "CSI API"
Private Const METHODAMAZON = "Amazon Bucket"
Private Const EASTERN = 2
Private Const CENTRAL = 1
Private Const MOUNTAIN = 3
Private Const PACIFIC = 4
Private Const NOTIMEZONE = 0
'be careful changing these!  I'm looping by the number
'I'm being sneaky with name- as it's for both api and bucket
Private Const NAMEINDEX = 0
Private Const PASSWORDINDEX = 1
Private Const URLINDEX = 2
Private Const ORGANIZATIONINDEX = 3
Private Const TIMEZONEINDEX = 4
Private Const FOLDERINDEX = 1
Private Const REGIONINDEX = 2
Private Const ACCESSINDEX = 3
Private Const PRIVATEINDEX = 4
Private Const PSTART As Integer = 0
Private Const PEND As Integer = 1
'I'm being sneaky with these...using as true/false
Private Const ACTIVE As Integer = 1
Private Const DORMANT As Integer = 0
'change of vendor name?  Change here. These are vendors I test directly
'Private Const VENDORTESTMEGAPHONE = "Megaphone"
'for Dick's lookahead
Private Const CHOICEYES As Integer = 0
Private Const CHOICENONE As Integer = 1
Private Const CHOICENEW As Integer = 2

Public imPassedID As Integer
'Added a vendor?
Public bmNeedRefresh As Boolean
'all these are used cbcSelect
Dim tmVendorInfo() As VendorInfo
Dim imCurrentVendorInfoIndex As Integer
Dim tmNewVendorInfo As VendorInfo
'from Dick's code
Dim imChgMode As Integer        'only in cbcSelect. to stop infinite looping
Dim imBSMode As Integer         'only in cbcSelect for look ahead and typing, keyselect, keypress
Dim imBypassSetting As Integer  'stop functions and clicking from running during change
Dim imTerminate As Integer
Private Sub mInit()
    imTerminate = False
    imFirstActivate = True
    Screen.MousePointer = vbHourglass
    If imTerminate Then
        Exit Sub
    End If
    mInitBox
    gCenterStdAlone Me
    Screen.MousePointer = vbHourglass
    imBSMode = False
    imCurrentVendorInfoIndex = -1
    imBypassSetting = False
    mPopMain True
    'Dan: I don't think this really works
    If imTerminate Then
        Exit Sub
    End If
    imChgMode = False
    lbcContractMethod.BackColor = vbYellow
    lbcImpressionMethod.BackColor = vbYellow
    cmcUpdate.Enabled = False
    lbcCode.Visible = bgInternalGuide
    
    mSetTabbing
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub mPopMain(blLoadFromDatabase As Boolean)
    mLoadAdVendors blLoadFromDatabase
    mLoadPassedChoice
End Sub
Private Sub mSetTabbing()
    Dim ilLoop As Integer
    Dim ilCurrentTab As Integer
    
    cbcSelect.TabIndex = 0
    For ilLoop = ACTIVE To DORMANT Step -1
        ilCurrentTab = ilCurrentTab + 1
        rbcStatus(ilLoop).TabIndex = ilCurrentTab
    Next
    ilCurrentTab = ilCurrentTab + 1
    txtVendorName.TabIndex = ilCurrentTab
    For ilLoop = PSTART To PEND
        ilCurrentTab = ilCurrentTab + 1
        txtPriority(ilLoop).TabIndex = ilCurrentTab
    Next
    For ilLoop = NAMEINDEX To ORGANIZATIONINDEX
        ilCurrentTab = ilCurrentTab + 1
        txtAPI(ilLoop).TabIndex = ilCurrentTab
    Next
    ilCurrentTab = ilCurrentTab + 1
    cboTimeZone.TabIndex = ilCurrentTab
    For ilLoop = NAMEINDEX To PRIVATEINDEX
        ilCurrentTab = ilCurrentTab + 1
        txtBucket(ilLoop).TabIndex = ilCurrentTab
    Next
    cmcDone.TabIndex = ilCurrentTab + 1
    cmcCancel.TabIndex = ilCurrentTab + 2
    cmcUpdate.TabIndex = ilCurrentTab + 3
    cmcPosition.TabIndex = ilCurrentTab + 4
End Sub
Private Function mSaveRec() As Boolean
    Dim slSql As String
    Dim ilVendorId As Integer
    Dim slStatus As String
    Dim llCount As Long
    Dim tlNewVendorInfo As VendorInfo

    If Not mValidateAll() Then
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    'strings are trimmed and made sql safe
    tlNewVendorInfo = mGatherControlValues()
With tlNewVendorInfo
    If .oStatus = ACTIVE Then
        slStatus = "A"
    Else
        slStatus = "D"
    End If
    ilVendorId = tlNewVendorInfo.iCode
    'update
    If ilVendorId > 0 Then
        slSql = "Update avf_AdVendor set avfName = '" & gSqlSafeAndTrim(.sName) & "', avfPriorityStart = " & .iPriorityStart & ", avfPriorityEnd = " & .iPriorityEnd & ", avfStatus = '" & slStatus & "'"
        If frcAPI.Visible Then
            slSql = slSql & ", avfVendorUserName = '" & gSqlSafeAndTrim(.sVendorUserName) & "', avfVendorPassword = '" & gSqlSafeAndTrim(.sVendorPassword) & "', avfTimeZone = '" & .sTimeZone & "'"
        End If
        If frcVendor.Visible Then
            slSql = slSql & ", avfVendorURL = '" & gSqlSafeAndTrim(.sVendorURL) & "', avfOrganizationID = '" & gSqlSafeAndTrim(.sOrgID) & "'"
        End If
        If frcBucket.Visible Then
            slSql = slSql & ", avfbucketname = '" & gSqlSafeAndTrim(.sBucketName) & "', avfBucketFolder = '" & gSqlSafeAndTrim(.sBucketFolder) & "', avfBucketRegion = '" & gSqlSafeAndTrim(.sBucketRegion) & "', avfBucketAccessKey = '" & gSqlSafeAndTrim(.sBucketAccessKey) & "', avfBucketPrivateKey = '" & gSqlSafeAndTrim(.sBucketPrivateKey) & "'"
        End If
        slSql = slSql & " WHERE avfCode = " & ilVendorId
        If gSQLAndReturn(slSql, False, llCount) <> 0 Then
            gHandleError "TrafficErrors.txt", Me.Name & "-mSaveRec"
            GoTo mSaveRecErr
        End If
        'just for safety.  Can this happen?
        If imCurrentVendorInfoIndex > -1 Then
            If Trim$(tmVendorInfo(imCurrentVendorInfoIndex).sName) <> Trim$(txtVendorName) Then
                bmNeedRefresh = True
            End If
            tmVendorInfo(imCurrentVendorInfoIndex) = tlNewVendorInfo
        End If
    Else
        'new
        slSql = "Insert into avf_AdVendor (avfName,avfStatus,avfPriorityStart,avfPriorityEnd, avfLastImportDate,avfVendorUserName,avfVendorPassword,avfVendorUrl,avfOrganizationID,avfTimeZone,avfBucketName,avfBucketFolder,avfBucketRegion,avfBucketAccessKey,avfBucketPrivateKey) "
        slSql = slSql & " VALUES('" & gSqlSafeAndTrim(.sName) & "','" & slStatus & "'," & .iPriorityStart & "," & .iPriorityEnd & ",'" & Format(NODATE, sgSQLDateForm) & "','" & gSqlSafeAndTrim(.sVendorUserName) & "','" & gSqlSafeAndTrim(.sVendorPassword) & "','" & gSqlSafeAndTrim(.sVendorURL) & "','" & gSqlSafeAndTrim(.sOrgID) & "','" & .sTimeZone & "'"
        slSql = slSql & ",'" & gSqlSafeAndTrim(.sBucketName) & "','" & gSqlSafeAndTrim(.sBucketFolder) & "','" & gSqlSafeAndTrim(.sBucketRegion) & "','" & gSqlSafeAndTrim(.sBucketAccessKey) & "','" & gSqlSafeAndTrim(.sBucketPrivateKey) & "')"
        ilVendorId = gInsertAndReturnCode(slSql, "avf_AdVendor", "avfCode", "DaNXDICk")
        If ilVendorId <= 0 Then
            GoTo mSaveRecErr
        End If
        tlNewVendorInfo.iCode = ilVendorId
        'hard-coded...already in array
        If imCurrentVendorInfoIndex > -1 Then
            tmVendorInfo(imCurrentVendorInfoIndex) = tlNewVendorInfo
        Else
            tmVendorInfo(UBound(tmVendorInfo)) = tlNewVendorInfo
            ReDim Preserve tmVendorInfo(UBound(tmVendorInfo) + 1)
        End If
        '"Hey!  We added a vendor!"
        bmNeedRefresh = True
    End If
End With
    imPassedID = ilVendorId
    mPopMain False
    mSaveRec = True
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function
Private Function mGatherControlValues() As VendorInfo
    Dim tlVendor As VendorInfo
    
    tlVendor.iCode = 0
    'I set these hard-coded values to be complete.  I don't use them in saving
    tlVendor.oContractMethod = ContractMethodType.Manual
    tlVendor.oImpressionMethod = ImpressionMethodType.Manual
    tlVendor.bIsCustom = True
    If imCurrentVendorInfoIndex > -1 Then
        tlVendor.iCode = tmVendorInfo(imCurrentVendorInfoIndex).iCode
        tlVendor.bIsCustom = tmVendorInfo(imCurrentVendorInfoIndex).bIsCustom
        tlVendor.bSecondaryID = tmVendorInfo(imCurrentVendorInfoIndex).bSecondaryID
        tlVendor.oContractMethod = tmVendorInfo(imCurrentVendorInfoIndex).oContractMethod
        tlVendor.oImpressionMethod = tmVendorInfo(imCurrentVendorInfoIndex).oImpressionMethod
    End If
    tlVendor.sName = txtVendorName
    tlVendor.iPriorityStart = txtPriority(PSTART)
    tlVendor.iPriorityEnd = txtPriority(PEND)
    tlVendor.sVendorUserName = txtAPI(NAMEINDEX)
    tlVendor.sVendorPassword = txtAPI(PASSWORDINDEX)
    tlVendor.sVendorURL = txtAPI(URLINDEX)
    tlVendor.sOrgID = txtAPI(ORGANIZATIONINDEX)
    tlVendor.sBucketName = txtBucket(NAMEINDEX)
    tlVendor.sBucketFolder = txtBucket(FOLDERINDEX)
    tlVendor.sBucketRegion = txtBucket(REGIONINDEX)
    tlVendor.sBucketAccessKey = txtBucket(ACCESSINDEX)
    tlVendor.sBucketPrivateKey = txtBucket(PRIVATEINDEX)
    tlVendor.sTimeZone = mTimeZoneConvertFromCombo()
    If rbcStatus(ACTIVE).Value Then
        tlVendor.oStatus = ACTIVE
    Else
        tlVendor.oStatus = DORMANT
    End If
    mGatherControlValues = tlVendor
End Function
Private Sub mMoveRecToCtrl()
    Dim tlVendor As VendorInfo
    Dim slMethod As String
    
    If imCurrentVendorInfoIndex > -1 Then
        tlVendor = tmVendorInfo(imCurrentVendorInfoIndex)
    Else
        mClearForNew
        tlVendor = tmNewVendorInfo
    End If
    With tlVendor
         lbcCode = tlVendor.iCode
         txtVendorName = Trim$(.sName)
        'careful!  tied the array index to the active dormant value of enum
        rbcStatus(.oStatus).Value = True
        Select Case .oContractMethod
            Case ContractMethodType.None
                slMethod = METHODNONE
            Case ContractMethodType.Manual
                slMethod = METHODMANUAL
            Case ContractMethodType.CSI
                slMethod = METHODCSI
            Case ContractMethodType.Vendor
                slMethod = METHODVENDOR
        End Select
        lbcContractMethod.Caption = slMethod
        Select Case .oImpressionMethod
            Case ImpressionMethodType.None
                slMethod = METHODNONE
            Case ImpressionMethodType.Manual
                slMethod = METHODMANUAL
            Case ImpressionMethodType.CSI
                slMethod = METHODCSI
            Case ImpressionMethodType.Amazon
                slMethod = METHODAMAZON
            Case ImpressionMethodType.Vendor
                slMethod = METHODVENDOR
        End Select
        lbcImpressionMethod.Caption = slMethod
        txtPriority(PSTART) = .iPriorityStart
        txtPriority(PEND) = .iPriorityEnd
        txtAPI(NAMEINDEX) = Trim$(.sVendorUserName)
        txtAPI(PASSWORDINDEX) = Trim$(.sVendorPassword)
        txtAPI(ORGANIZATIONINDEX) = Trim$(.sOrgID)
        txtAPI(URLINDEX) = Trim$(.sVendorURL)
        cboTimeZone.ListIndex = mTimeZoneConvertToCombo(.sTimeZone)
        txtBucket(NAMEINDEX) = Trim$(.sBucketName)
        txtBucket(FOLDERINDEX) = Trim$(.sBucketFolder)
        txtBucket(REGIONINDEX) = Trim$(.sBucketRegion)
        txtBucket(ACCESSINDEX) = Trim$(.sBucketAccessKey)
        txtBucket(PRIVATEINDEX) = Trim$(.sBucketPrivateKey)
    End With
End Sub
Private Sub mTerminate()
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Me
    igManUnload = NO
End Sub
Private Sub mSetCommands()
    Dim blAltered As Boolean
    
    If imBypassSetting Then
        Exit Sub
    End If
    blAltered = mAnyFieldChanged()
    If blAltered Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    cmcPosition.Enabled = False
    'must have an avfcode and not an 'information vendor' (contract method = none)
    If imCurrentVendorInfoIndex > -1 Then
        If tmVendorInfo(imCurrentVendorInfoIndex).iCode > 0 And tmVendorInfo(imCurrentVendorInfoIndex).oContractMethod <> ContractMethodType.None Then
            cmcPosition.Enabled = True
        End If
    End If
    '5/7/21 fix of 'bad focus' from validation message
    If blAltered = YES Then
        cbcSelect.Enabled = False
    Else
        cbcSelect.Enabled = True
    End If
End Sub
Private Function mAnyFieldChanged() As Boolean
    Dim blRet As Boolean
    Dim tlVendor As VendorInfo
    
    blRet = False
    If imCurrentVendorInfoIndex > -1 Then
        tlVendor = tmVendorInfo(imCurrentVendorInfoIndex)
    Else
        tlVendor = tmNewVendorInfo
    End If
    With tlVendor
        If Trim$(txtVendorName) <> Trim$(.sName) Then
            blRet = True
            GoTo CONTINUE
        End If
        If (.oStatus = ACTIVE And rbcStatus(DORMANT).Value) Or .oStatus = DORMANT And rbcStatus(ACTIVE).Value Then
            blRet = True
            GoTo CONTINUE
        End If
        If txtPriority(PSTART) <> .iPriorityStart Or txtPriority(PEND) <> .iPriorityEnd Then
            blRet = True
            GoTo CONTINUE
        End If
        If frcAPI.Visible Then
            If txtAPI(NAMEINDEX) <> Trim$(.sVendorUserName) Or txtAPI(PASSWORDINDEX) <> Trim$(.sVendorPassword) Or txtAPI(URLINDEX) <> Trim$(.sVendorURL) Then
                blRet = True
                GoTo CONTINUE
            End If
        End If
        If frcVendor.Visible Then
            If txtAPI(ORGANIZATIONINDEX) <> Trim$(.sOrgID) Or mTimeZoneConvertFromCombo() <> .sTimeZone Then
                blRet = True
                GoTo CONTINUE
            End If
        End If
        If frcBucket.Visible Then
            If txtBucket(NAMEINDEX) <> Trim$(.sBucketName) Or txtBucket(FOLDERINDEX) <> Trim$(.sBucketFolder) Or txtBucket(REGIONINDEX) <> Trim$(.sBucketRegion) Or txtBucket(ACCESSINDEX) <> Trim$(.sBucketAccessKey) Or txtBucket(PRIVATEINDEX) <> Trim$(.sBucketPrivateKey) Then
                blRet = True
                GoTo CONTINUE
            End If
        End If
    End With
CONTINUE:
    mAnyFieldChanged = blRet
End Function
Private Function mSaveRecChg(blAsk As Boolean) As Boolean

'       blAsk (I)- True = Ask if changed records should be updated;
'                 False= Update record if required without asking user
    Dim blRet As Boolean
    Dim slMess As String
    
    blRet = True
    If mAnyFieldChanged() Then
        If blAsk Then
            slMess = "Add " & txtVendorName
            If imCurrentVendorInfoIndex > -1 Then
                If tmVendorInfo(imCurrentVendorInfoIndex).iCode > 0 Then
                    slMess = "Save Changes to " & txtVendorName
                End If
            End If
            Select Case MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                Case vbCancel
                    blRet = False
                Case vbYes
                    'we'll handle below
                Case vbNo
                    blRet = False
                    imTerminate = True
            End Select
        End If
        If blRet Then
            blRet = mSaveRec()
        End If
    End If
    mSaveRecChg = blRet
End Function
Private Sub mLoadPassedChoice()
    Dim ilLoop As Integer
    Dim ilFound As Integer
    
    ilFound = -1
    If imPassedID > 0 Then
        For ilLoop = 0 To UBound(tmVendorInfo) - 1
            If tmVendorInfo(ilLoop).iCode = imPassedID Then
                ilFound = ilLoop
                Exit For
            End If
        Next ilLoop
    End If
    If ilFound > -1 Then
        'we're storing the array index # with name
         For ilLoop = 0 To cbcSelect.ListCount - 1
            If cbcSelect.ItemData(ilLoop) = ilFound Then
                cbcSelect.ListIndex = ilLoop
                Exit For
            End If
         Next
    Else
        cbcSelect.ListIndex = 0
    End If
End Sub
Private Sub mLoadAdVendors(blFromDatabase As Boolean)
    Dim ilLoop As Integer
    Dim slSql As String
    
    cbcSelect.Clear
    If blFromDatabase Then
        tmVendorInfo = gGetDigitalVendors(True)
    End If
    For ilLoop = 0 To UBound(tmVendorInfo) - 1
        cbcSelect.AddItem Trim$(tmVendorInfo(ilLoop).sName)
        'add one because we're going to add 'new' at beginning
        cbcSelect.ItemData(cbcSelect.NewIndex) = ilLoop
    Next
    cbcSelect.AddItem "[New]", 0
    cbcSelect.ItemData(0) = -1
End Sub
Private Sub mClearForNew()
     With tmNewVendorInfo
        .bIsCustom = True
        .iCode = 0
        .iPriorityEnd = 0
        .iPriorityStart = 0
        .oContractMethod = ContractMethodType.Manual
        .oImpressionMethod = ImpressionMethodType.Manual
        .oStatus = ACTIVE
        .sVendorURL = ""
        .sBucketAccessKey = ""
        .sBucketFolder = ""
        .sBucketName = ""
        .sBucketPrivateKey = ""
        .sBucketRegion = ""
        .sLastImportDate = NODATE
        .sName = ""
        .sOrgID = ""
        .sVendorPassword = ""
        .sVendorUserName = ""
        .sTimeZone = ""
        .bSecondaryID = False
    End With
End Sub
Private Sub mInitBox()
    frcVendor.BorderStyle = 0
    frcVendor.Left = 10
    frcVendor.Width = frcAPI.Width - 40
    
    frcPriority.BorderStyle = 0
    frcPriority.Left = 10
    frcPriority.Width = frcBasic.Width - 40
    
    frcVendor.Visible = False
    frcBucket.Visible = False
    
    txtAPI(NAMEINDEX).MaxLength = 255
    txtAPI(PASSWORDINDEX).MaxLength = 255
    txtAPI(URLINDEX).MaxLength = 255
    txtAPI(ORGANIZATIONINDEX).MaxLength = 70
    txtBucket(NAMEINDEX).MaxLength = 63
    txtBucket(FOLDERINDEX).MaxLength = 80
    txtBucket(REGIONINDEX).MaxLength = 16
    txtBucket(ACCESSINDEX).MaxLength = 20
    txtBucket(PRIVATEINDEX).MaxLength = 40
    
    mInitTimeZone
End Sub
Private Sub mInitTimeZone()
cboTimeZone.AddItem "Not Defined", NOTIMEZONE
cboTimeZone.AddItem "Central", CENTRAL
cboTimeZone.AddItem "Eastern", EASTERN
cboTimeZone.AddItem "Mountain", MOUNTAIN
cboTimeZone.AddItem "Pacific", PACIFIC
cboTimeZone.ListIndex = NOTIMEZONE
End Sub
Private Sub mSetDisplayFields()
    frcPriority.Visible = True
    If imCurrentVendorInfoIndex > -1 Then
        With tmVendorInfo(imCurrentVendorInfoIndex)
            'information vendors don't set priorities
            If .oContractMethod = ContractMethodType.None Then
                frcPriority.Visible = False
            End If
            If .oContractMethod = ContractMethodType.CSI Or .oContractMethod = ContractMethodType.Vendor Or .oContractMethod = ContractMethodType.None Or .oImpressionMethod = ImpressionMethodType.CSI Or .oImpressionMethod = ImpressionMethodType.Vendor Then
                frcAPI.Visible = True
                If .oContractMethod = ContractMethodType.Vendor Then
                    frcVendor.Visible = True
                Else
                    frcVendor.Visible = False
                End If
            Else
                frcAPI.Visible = False
            End If
            If .oImpressionMethod = Amazon Then
                frcBucket.Visible = True
            Else
                frcBucket.Visible = False
            End If
            'since only Megaphone is using org id and time zone, just use this secondaryid for both Can change in future
            txtAPI(ORGANIZATIONINDEX).Visible = .bSecondaryID
            lbcOrganizationID.Visible = .bSecondaryID
            cboTimeZone.Visible = .bSecondaryID
            lbcTimeZone.Visible = .bSecondaryID
            mEnableFields (.oStatus)
        End With
    Else
        'new
        frcAPI.Visible = False
        frcBucket.Visible = False
        rbcStatus(ACTIVE).Value = True
        mEnableFields False
    End If
End Sub
Private Sub mEnableFields(blEnable As Boolean)
    Dim ilLoop As Integer
    
    txtVendorName.Enabled = False
    'new
    If imCurrentVendorInfoIndex = -1 Then
        blEnable = False
        txtVendorName.Enabled = True
        txtPriority(PSTART).Enabled = True
        txtPriority(PEND).Enabled = True
    Else
        If tmVendorInfo(imCurrentVendorInfoIndex).bIsCustom Then
            txtVendorName.Enabled = True
        End If
        For ilLoop = PSTART To PEND
            txtPriority(ilLoop).Enabled = blEnable
        Next
    End If
    For ilLoop = NAMEINDEX To ORGANIZATIONINDEX
        txtAPI(ilLoop).Enabled = blEnable
    Next
    For ilLoop = NAMEINDEX To PRIVATEINDEX
        txtBucket(ilLoop).Enabled = blEnable
    Next
End Sub
Private Function mCheckKeyIsNumber(ilKeyAscii) As Boolean
    '< 13 CR LF Tab  48 to 57 are 0-9
    If ilKeyAscii > 13 And (ilKeyAscii < 48 Or ilKeyAscii > 57) Then
        Beep
        mCheckKeyIsNumber = False
        Exit Function
    End If
    mCheckKeyIsNumber = True
End Function
Private Function mValidateAll() As Boolean
    Dim blRet As Boolean
    
    blRet = False
    If mValidateVendorName() Then
        If mValidatePriorities() Then
            If mValidateUserName() Then
                blRet = True
            End If
        End If
    End If
    mValidateAll = blRet
End Function
Private Function mValidateVendorName() As Boolean
    Dim blRet As Boolean
    Dim ilLoop As Integer
    Dim slNewVendorName As String
    Dim ilVendorCode As Integer
    
    
    blRet = True
    ilVendorCode = 0
    If imCurrentVendorInfoIndex > -1 Then
        ilVendorCode = tmVendorInfo(imCurrentVendorInfoIndex).iCode
    End If
    'only enabled for custom
    If txtVendorName.Enabled Then
        slNewVendorName = Trim$(txtVendorName)
        If Len(slNewVendorName) > 0 Then
            For ilLoop = 0 To UBound(tmVendorInfo) - 1
                If Trim$(tmVendorInfo(ilLoop).sName) = slNewVendorName And tmVendorInfo(ilLoop).iCode <> ilVendorCode Then
                    blRet = False
                    gMsgBox "Vendor Name must be unique", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    Exit For
                End If
            Next
        Else
            blRet = False
            gMsgBox "VendorName required", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
        End If
        If Not blRet Then
            txtVendorName.SetFocus
        End If
    End If
    mValidateVendorName = blRet
End Function
Private Function mValidatePriorityAmount(Index As Integer) As Boolean
    Dim blRet As Boolean
    Dim ilValue As Integer
    blRet = False
    If IsNumeric(txtPriority(Index)) Then
        ilValue = txtPriority(Index)
        If ilValue >= 0 And ilValue <= 20 Then
            blRet = True
        Else
            gMsgBox "Priority must be less than 21", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
            ilValue = 0
            If imCurrentVendorInfoIndex > -1 Then
                If Index = PSTART Then
                    ilValue = tmVendorInfo(imCurrentVendorInfoIndex).iPriorityStart
                Else
                    ilValue = tmVendorInfo(imCurrentVendorInfoIndex).iPriorityEnd
                End If
            End If
            txtPriority(Index).Text = ilValue
        End If
    Else
        gMsgBox "Priority must be a number", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
    End If
    mValidatePriorityAmount = blRet
End Function
Private Function mValidatePriorities() As Boolean
    Dim blRet As Boolean
    Dim ilStart As Integer
    Dim ilEnd As Integer
    
    blRet = True
    If IsNumeric(txtPriority(PSTART)) Then
        ilStart = txtPriority(PSTART)
    Else
        blRet = False
        Beep
        gMsgBox "Priority Start must be a number", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
        txtPriority(PSTART).SetFocus
    End If
    If IsNumeric(txtPriority(PEND)) Then
        ilEnd = txtPriority(PEND)
    Else
        blRet = False
        Beep
        gMsgBox "Priority End must be less than or equal to Priority End", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
        txtPriority(PEND).SetFocus
    End If
    If ilStart > ilEnd Then
        blRet = False
        Beep
        gMsgBox "Priority Start must be less than or equal to Priority End", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
        txtPriority(PSTART).SetFocus
    End If
    If blRet Then
        blRet = mValidatePriorityAmount(PSTART)
    End If
    If blRet Then
        blRet = mValidatePriorityAmount(PEND)
        If Not blRet Then
            txtPriority(PEND).SetFocus
        End If
    Else
        txtPriority(PSTART).SetFocus
    End If
    mValidatePriorities = blRet
End Function
Private Function mValidateUserName() As Boolean
    Dim blRet As Boolean
    Dim ilLoop As Integer
    Dim slNewUsername As String
    Dim ilVendorCode As Integer
    Dim olContractMethod As ContractMethodType
        
    blRet = True
    slNewUsername = Trim$(txtAPI(NAMEINDEX))
    If imCurrentVendorInfoIndex > -1 Then
        ilVendorCode = tmVendorInfo(imCurrentVendorInfoIndex).iCode
        olContractMethod = tmVendorInfo(imCurrentVendorInfoIndex).oContractMethod
    Else
        ilVendorCode = tmNewVendorInfo.iCode
        olContractMethod = tmNewVendorInfo.oContractMethod
    End If
    'test username only when not blank
    If (olContractMethod = ContractMethodType.CSI Or olContractMethod = ContractMethodType.Vendor) And slNewUsername <> "" Then
        For ilLoop = 0 To UBound(tmVendorInfo) - 1
            If Trim$(tmVendorInfo(ilLoop).sVendorUserName) = slNewUsername And tmVendorInfo(ilLoop).iCode <> ilVendorCode Then
                blRet = False
                gMsgBox "Username must be unique", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                txtAPI(NAMEINDEX).SetFocus
                Exit For
            End If
        Next
    End If
    mValidateUserName = blRet
End Function
Private Function mTimeZoneConvertToCombo(slTimeZone As String) As Integer
    Dim ilRet As Integer
    
    Select Case slTimeZone
        Case "E"
            ilRet = EASTERN
        Case "P"
            ilRet = PACIFIC
        Case "M"
            ilRet = MOUNTAIN
        Case "C"
            ilRet = CENTRAL
        Case Else
            ilRet = NOTIMEZONE
        End Select
    mTimeZoneConvertToCombo = ilRet
End Function
Private Function mTimeZoneConvertFromCombo() As String
    Dim slRet As String
    
    Select Case cboTimeZone.ListIndex
        Case EASTERN
            slRet = "E"
        Case PACIFIC
            slRet = "P"
        Case MOUNTAIN
            slRet = "M"
        Case CENTRAL
            slRet = "C"
        Case Else
            slRet = ""
        End Select
    mTimeZoneConvertFromCombo = slRet
End Function

Private Sub cboTimeZone_LostFocus()
    mSetCommands
End Sub

Private Sub rbcStatus_Click(Index As Integer)
    If Not imBypassSetting Then
        mSetCommands
        mEnableFields rbcStatus(ACTIVE).Value
    End If
End Sub
Private Sub txtPriority_GotFocus(Index As Integer)
   With txtPriority(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub
Private Sub txtPriority_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not mCheckKeyIsNumber(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub txtPriority_Validate(Index As Integer, Cancel As Boolean)
    If Not mValidatePriorityAmount(Index) Then
        Cancel = True
        With txtPriority(Index)
           .SelStart = 0
           .SelLength = Len(.Text)
        End With
    End If
End Sub
Private Sub txtPriority_LostFocus(Index As Integer)
    mSetCommands
End Sub
Private Sub txtAPI_GotFocus(Index As Integer)
   With txtAPI(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub
Private Sub txtAPI_LostFocus(Index As Integer)
    mSetCommands
End Sub
Private Sub txtBucket_GotFocus(Index As Integer)
   With txtBucket(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub
Private Sub txtBucket_LostFocus(Index As Integer)
    mSetCommands
End Sub
Private Sub txtVendorName_GotFocus()
    With txtVendorName
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub
Private Sub txtVendorName_LostFocus()
    mSetCommands
End Sub
Private Sub cmcPosition_Click()
    Dim slStr As String
    
    'I shouldn't need this test---only enabled if not new
    If imCurrentVendorInfoIndex > -1 Then
        sgMnfCallType = "6"
        igMNmCallSource = CALLSOURCEVEHICLE
        sgMNmName = Trim$(txtVendorName) & "/" & tmVendorInfo(imCurrentVendorInfoIndex).iCode
        If igTestSystem Then
            slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\1"
        Else
            slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\1"
        End If
        sgCommandStr = slStr
        MultiNm.Show vbModal
    End If
End Sub
'This is the last of Dick's code.  Really just to set KeyPreview which is no longer needed.  Remove?
'Private Sub Form_Activate()
'    If Not imFirstActivate Then
'        DoEvents    'Process events so pending keys are not sent to this
'                    'form when keypreview turn on
'       ' gShowBranner imUpdateAllowed
'        Me.KeyPreview = True
'        Exit Sub
'    End If
'    imFirstActivate = False
'    Me.KeyPreview = True
'    AdServerVendor.Refresh
'End Sub
'Private Sub Form_Deactivate()
'    Me.KeyPreview = False
'End Sub
'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    Dim ilReSet As Integer
'
'    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
'        If cbcSelect.Enabled Then
'            cbcSelect.Enabled = False
'            ilReSet = True
'        Else
'            ilReSet = False
'        End If
'        gFunctionKeyBranch KeyCode
'        If ilReSet Then
'            cbcSelect.Enabled = True
'        End If
'    End If
'End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        If Not igManUnload Then
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
End Sub
Private Sub cbcSelect_Change()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    If ilRet = CHOICEYES Then
        imCurrentVendorInfoIndex = cbcSelect.ItemData(cbcSelect.ListIndex)
    Else
        If ilRet = CHOICENONE Then
            If cbcSelect.ListCount > 0 Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.ListIndex = -1
            End If
        End If
        imCurrentVendorInfoIndex = -1
    End If
    'will clear as needed and creates/clears tmNewVendorInfo if needed
    mMoveRecToCtrl
    mSetDisplayFields
    Screen.MousePointer = vbDefault
    imChgMode = False
    imBypassSetting = False
    mSetCommands
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    Exit Sub
End Sub
Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cmcUpdate_Click()
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            mTerminate
            Exit Sub
        End If
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    mSetCommands
    If cbcSelect.Enabled Then
        cbcSelect.SetFocus
    Else
        cmcCancel.SetFocus
    End If
End Sub
Private Sub cmcDone_Click()
    If imCurrentVendorInfoIndex > -1 Then
        imPassedID = tmVendorInfo(imCurrentVendorInfoIndex).iCode
    End If
    'changed a field, even if not tabbed off? ask if they want to save
    If Not cmcUpdate.Enabled Then
        cmcCancel_Click
        Exit Sub
    End If
    If mSaveRecChg(True) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_Click()
    If cmcUpdate.Enabled Then
        If imCurrentVendorInfoIndex = -1 Then
            mClearForNew
        End If
        imBypassSetting = True
        mMoveRecToCtrl
        imBypassSetting = False
        mSetCommands
    Else
        mTerminate
    End If
End Sub


