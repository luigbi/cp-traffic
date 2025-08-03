VERSION 5.00
Begin VB.Form ExpMK 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2805
   ClientLeft      =   825
   ClientTop       =   2400
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2805
   ScaleWidth      =   7095
   Begin VB.TextBox edcYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   0
      Top             =   570
      Width           =   615
   End
   Begin VB.TextBox edcContract 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   3
      Top             =   1080
      Width           =   1200
   End
   Begin VB.TextBox edcPeriods 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   5160
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "1"
      Top             =   570
      Width           =   615
   End
   Begin VB.ListBox lbcVehicles 
      Appearance      =   0  'Flat
      Height          =   420
      ItemData        =   "ExpMK.frx":0000
      Left            =   570
      List            =   "ExpMK.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   9
      Top             =   1500
      Visible         =   0   'False
      Width           =   4380
   End
   Begin VB.Timer tmcCancel 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6480
      Top             =   0
   End
   Begin VB.TextBox edcMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   2955
      MaxLength       =   3
      TabIndex        =   1
      Top             =   570
      Width           =   615
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
      Left            =   2400
      TabIndex        =   4
      Top             =   2355
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
      TabIndex        =   6
      Top             =   2355
      Width           =   1050
   End
   Begin VB.Label lacSelCFrom 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Year"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   195
      TabIndex        =   13
      Top             =   623
      Width           =   870
   End
   Begin VB.Label lacContract 
      Appearance      =   0  'Flat
      Caption         =   "Contract #"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   195
      TabIndex        =   12
      Top             =   1140
      Width           =   1065
   End
   Begin VB.Label lacSelCFrom1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   2220
      TabIndex        =   11
      Top             =   623
      Width           =   540
   End
   Begin VB.Label lacSelCFrom1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "# of Periods"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   3960
      TabIndex        =   10
      Top             =   623
      Width           =   1050
   End
   Begin VB.Label lacTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Miller Kaplan Export"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   8
      Top             =   75
      Width           =   1710
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   420
      TabIndex        =   7
      Top             =   1770
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   420
      TabIndex        =   5
      Top             =   1485
      Visible         =   0   'False
      Width           =   5550
   End
End
Attribute VB_Name = "ExpMK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Dim imFirstActivate As Integer

Dim tmMnf As MNF
Dim imMnfRecLen As Integer
Dim hmMnf As Integer
Dim tmMnfSrchKey As INTKEY0

Dim smExportName As String
Dim lmCntrNo As Long    'for debugging purposes to filter a single contract

Dim tmExportInfo() As MKExportInfo
Dim tmRvf() As MKExportInfo
Dim tmPhf() As MKExportInfo
Private Type MKExportInfo
    sCntrKey As String * 9              'sorting key by Contact
    sAdvKey As String * 26              'sorting key by Advertiser
    iAdfCode As Integer                 'Advertiser code
    lContractNo As Long                 'Contract number
    sStation As String                  'client name & address
    slAgency As String                  'agency contact (buyer)
    sAdvertiser As String               'advertiser name
    slBrand As String                   'product name from copy (brand)
    sAEFullName As String               'salesperson name (AE Full Name)
    sAccountType As String              'account type (cash or trade)
    sProductCodeDesc As String          'product code description
    sRevenueType As String              'revenue type (Air Time / NTR)
    sDirectOffice As String
    slYearMonth As String
    lGrossAmount As Long
    sTranDate As String
    bProcessed As Boolean
End Type

Dim tmAgf() As CodeNames
Dim tmSof() As CodeNames
Dim tmMnf2() As CodeNames
Dim tmPrf() As CodeNames
Private Type CodeNames
    lCode As Long
    slName As String
End Type

Dim tmMNFCODE() As MNFCode
Private Type MNFCode
    lCode As Long
    iMnfCode As Integer
End Type

Dim tmSlf() As SLFCodeNames
Private Type SLFCodeNames
    iCode As Integer
    iSofCode As Integer
    slName As String
End Type

Dim hmMK As Integer
Dim hmMsg As Integer

Dim imTerminate As Integer
Dim imExporting As Integer
Dim lmNowDate As Long

Dim bmStdExport As Boolean
Dim smClientName As String
Dim imExportOption As Integer       'lbcExport.ItemData(lbcExport.ListIndex)
Dim smExportOptionName As String    'Miller Kaplan

Dim imFirstTime As Integer
Dim slMonthStr As String * 36

Dim llContractNo() As MKContracts
Private Type MKContracts
    sKey As String * 9
    lContractNo As Long
End Type

Dim ilAdfCode() As Integer
Private Function mAccumulateGrossAmount(ByVal lContractNo As Long, ByVal iAdfCode As Integer, ByVal sAdvertiser As String, ByVal sAgent As String, ByVal sProduct As String, ByVal slCashTrade As String, ByVal slRevenueType As String, ByVal lGrossAmount As Long, ByVal sAgentName As String) As Boolean
    Dim llCounter1 As Integer
    Dim llCounter2 As Long
    
    For llCounter1 = 0 To UBound(tmExportInfo) - 1
        mAccumulateGrossAmount = False
        If sAdvertiser = tmExportInfo(llCounter1).sAdvertiser And sAgent = tmExportInfo(llCounter1).slAgency And sProduct = tmExportInfo(llCounter1).slBrand And _
           slCashTrade = tmExportInfo(llCounter1).sAccountType And slRevenueType = tmExportInfo(llCounter1).sRevenueType And sAgentName = tmExportInfo(llCounter1).sAEFullName Then
            
            tmExportInfo(llCounter1).lGrossAmount = tmExportInfo(llCounter1).lGrossAmount + lGrossAmount
            mAccumulateGrossAmount = True
            Exit For
        Else
            For llCounter2 = 0 To UBound(tmExportInfo) - 1
                If tmExportInfo(llCounter2).lContractNo = lContractNo Then
                    If sAdvertiser = tmExportInfo(llCounter2).sAdvertiser And sAgent = tmExportInfo(llCounter2).slAgency And _
                        sProduct = tmExportInfo(llCounter2).slBrand And slCashTrade = tmExportInfo(llCounter2).sAccountType And _
                        slRevenueType = tmExportInfo(llCounter2).sRevenueType And sAgentName = tmExportInfo(llCounter2).sAEFullName Then
                        
                        tmExportInfo(llCounter2).lGrossAmount = tmExportInfo(llCounter2).lGrossAmount + lGrossAmount
                        mAccumulateGrossAmount = True
                        Exit For
                    End If
                End If
            Next llCounter2
            If mAccumulateGrossAmount = True Then Exit For
        End If
    Next llCounter1
End Function


Private Function mCreateExportInfo(ByVal llContractNo As Long, tmRvfPhf() As MKExportInfo, ByVal slStation As String) As Boolean
    
    Dim slAgfName As String
    Dim slAdfName As String
    Dim slPrfName As String
    Dim slFullName As String
    Dim slCashTrade As String
    Dim slRevenueType As String
    Dim slMnfName As String
    Dim slSofName As String
    Dim slTranDate As String
    Dim llGrossAmount As Long
    Dim llTotalGross As Long
    Dim llRVFCount As Long
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    'Dim llContractNo As Long
    
    On Error GoTo mCreateExportInfo_Err
    
    slAgfName = "": slAdfName = "": slPrfName = "": slFullName = "": slTranDate = ""
    slCashTrade = "": slRevenueType = "": slMnfName = "": slSofName = "": llGrossAmount = 0: llTotalGross = 0
    
    
    'process records for each contact in RVF
    llRVFCount = 0
    For llRVFCount = 0 To UBound(tmRvfPhf) - 1
        If llContractNo = tmRvfPhf(llRVFCount).lContractNo Then
            If (slAdfName = "" And slPrfName = "" And slFullName = "" And _
                slCashTrade = "" And slRevenueType = "") Then
                slAgfName = tmRvfPhf(llRVFCount).slAgency
                slAdfName = tmRvfPhf(llRVFCount).sAdvertiser
                slPrfName = tmRvfPhf(llRVFCount).slBrand
                slFullName = tmRvfPhf(llRVFCount).sAEFullName
                'If Not IsNull(MK_rvf!slfLastName) Then slLastName = Trim$(MK_rvf!slfLastName)
                slCashTrade = tmRvfPhf(llRVFCount).sAccountType
                slRevenueType = tmRvfPhf(llRVFCount).sRevenueType
                slMnfName = tmRvfPhf(llRVFCount).sProductCodeDesc
                slSofName = tmRvfPhf(llRVFCount).sDirectOffice
                slTranDate = tmRvfPhf(llRVFCount).sTranDate '(MK_rvf!rvfTranDate, "m/d/yyyy")
                llGrossAmount = tmRvfPhf(llRVFCount).lGrossAmount
                'llContractNo = tmRvfPhf(llRVFCount).lContractNo
                llTotalGross = 0
            End If
            
            'filter agfName, adfName, prfName, slfFirstName, slfLastName, rvfCashTrade, tvfmnfItem
            If (slAgfName = tmRvfPhf(llRVFCount).slAgency And slAdfName = tmRvfPhf(llRVFCount).sAdvertiser And slPrfName = tmRvfPhf(llRVFCount).slBrand And _
                slFullName = tmRvfPhf(llRVFCount).sAEFullName And slCashTrade = tmRvfPhf(llRVFCount).sAccountType And slRevenueType = tmRvfPhf(llRVFCount).sRevenueType) Then
                
                'accumulate gross amount for the same filter: agency, adfName, advertiser, product, agent name, cash/trade, Air Time/NTR
                If Not mAccumulateGrossAmount(llContractNo, tmRvfPhf(llRVFCount).iAdfCode, slAdfName, slAgfName, slPrfName, slCashTrade, slRevenueType, tmRvfPhf(llRVFCount).lGrossAmount, slFullName) Then
                    tmExportInfo(UBound(tmExportInfo)).sStation = slStation
                    tmExportInfo(UBound(tmExportInfo)).slAgency = slAgfName
                    tmExportInfo(UBound(tmExportInfo)).sAdvertiser = slAdfName
                    tmExportInfo(UBound(tmExportInfo)).slBrand = slPrfName
                    tmExportInfo(UBound(tmExportInfo)).sAEFullName = slFullName
                    tmExportInfo(UBound(tmExportInfo)).sAccountType = slCashTrade
                    tmExportInfo(UBound(tmExportInfo)).sProductCodeDesc = slMnfName
                    tmExportInfo(UBound(tmExportInfo)).sRevenueType = slRevenueType
                    tmExportInfo(UBound(tmExportInfo)).lGrossAmount = llGrossAmount
                    tmExportInfo(UBound(tmExportInfo)).sDirectOffice = slSofName
                    tmExportInfo(UBound(tmExportInfo)).bProcessed = False
                    tmExportInfo(UBound(tmExportInfo)).lContractNo = llContractNo
                    tmExportInfo(UBound(tmExportInfo)).sTranDate = slTranDate
                    
                    gObtainYearMonthDayStr slTranDate, True, slYear, slMonth, slDay
                    tmExportInfo(UBound(tmExportInfo)).slYearMonth = Mid$(slMonthStr, (Val(slMonth) - 1) * 3 + 1, 3) & " " & slYear
                    tmExportInfo(UBound(tmExportInfo)).sAdvKey = slAdfName
                    
                    ReDim Preserve tmExportInfo(0 To UBound(tmExportInfo) + 1) As MKExportInfo
                End If
            Else
                'same contract number but different Cash/Trade, Air Time/NTR, product, etc.
                slAgfName = tmRvfPhf(llRVFCount).slAgency
                slAdfName = tmRvfPhf(llRVFCount).sAdvertiser
                slPrfName = tmRvfPhf(llRVFCount).slBrand
                slFullName = tmRvfPhf(llRVFCount).sAEFullName
                slCashTrade = tmRvfPhf(llRVFCount).sAccountType
                slRevenueType = tmRvfPhf(llRVFCount).sRevenueType
                slMnfName = tmRvfPhf(llRVFCount).sProductCodeDesc
                slSofName = tmRvfPhf(llRVFCount).sDirectOffice
                slTranDate = tmRvfPhf(llRVFCount).sTranDate '(MK_rvf!rvfTranDate, "m/d/yyyy")
                llGrossAmount = tmRvfPhf(llRVFCount).lGrossAmount
                'llContractNo = tmRvfPhf(llRVFCount).lContractNo
                llTotalGross = llGrossAmount
                
                'accumulate gross amount for the same filter: agency, adfName, advertiser, product, agent name, cash/trade, Air Time/NTR
                If Not mAccumulateGrossAmount(llContractNo, tmRvfPhf(llRVFCount).iAdfCode, slAdfName, slAgfName, slPrfName, slCashTrade, slRevenueType, tmRvfPhf(llRVFCount).lGrossAmount, slFullName) Then
                    tmExportInfo(UBound(tmExportInfo)).sStation = slStation
                    tmExportInfo(UBound(tmExportInfo)).slAgency = slAgfName
                    tmExportInfo(UBound(tmExportInfo)).sAdvertiser = slAdfName
                    tmExportInfo(UBound(tmExportInfo)).slBrand = slPrfName
                    tmExportInfo(UBound(tmExportInfo)).sAEFullName = slFullName
                    tmExportInfo(UBound(tmExportInfo)).sAccountType = slCashTrade
                    tmExportInfo(UBound(tmExportInfo)).sProductCodeDesc = slMnfName
                    tmExportInfo(UBound(tmExportInfo)).sRevenueType = slRevenueType
                    tmExportInfo(UBound(tmExportInfo)).lGrossAmount = llGrossAmount
                    tmExportInfo(UBound(tmExportInfo)).sDirectOffice = slSofName
                    tmExportInfo(UBound(tmExportInfo)).bProcessed = False
                    tmExportInfo(UBound(tmExportInfo)).lContractNo = llContractNo
                    tmExportInfo(UBound(tmExportInfo)).sTranDate = slTranDate

                    gObtainYearMonthDayStr slTranDate, True, slYear, slMonth, slDay
                    tmExportInfo(UBound(tmExportInfo)).slYearMonth = Mid$(slMonthStr, (Val(slMonth) - 1) * 3 + 1, 3) & " " & slYear
                    tmExportInfo(UBound(tmExportInfo)).sAdvKey = slAdfName

                    ReDim Preserve tmExportInfo(0 To UBound(tmExportInfo) + 1) As MKExportInfo
                End If
            End If
        End If
    Next llRVFCount
    mCreateExportInfo = True
    Exit Function
    
mCreateExportInfo_Err:
    mCreateExportInfo = False
End Function

Private Function mGetMNFCode(ByVal lContract As Long) As Integer
Dim rst As Recordset
    Dim iCode As Integer
    Dim llCounter As Long
    Dim slSql As String
    
    iCode = 0
    If UBound(tmMNFCODE) = 0 Then
        slSql = "Select distinct chfCntrNo, chfmnfComp1, chfType  From CHF_Contract_Header where chfType not in ('S','M')"
        Set rst = gSQLSelectCall(slSql)
        If Not rst.EOF Then
            Do While Not rst.EOF
                tmMNFCODE(UBound(tmMNFCODE)).lCode = rst!chfCntrno
                tmMNFCODE(UBound(tmMNFCODE)).iMnfCode = rst!chfmnfComp1
                
                If lContract = rst!chfCntrno Then iCode = rst!chfmnfComp1
                
                ReDim Preserve tmMNFCODE(0 To UBound(tmMNFCODE) + 1)
                rst.MoveNext
            Loop
        End If
    End If
    If iCode = 0 Then
        For llCounter = 0 To UBound(tmMNFCODE) - 1
            If tmMNFCODE(llCounter).lCode = lContract Then
                iCode = tmMNFCODE(llCounter).iMnfCode
                Exit For
            End If
        Next llCounter
    End If
    mGetMNFCode = iCode
    Set rst = Nothing
    
End Function

Private Function mGetExportName() As String
    Dim slStartMonth
    Dim slEndMonth
    Dim slStartYear As String
    Dim slEndYear As String
    
    slStartYear = edcYear.Text
    slEndYear = edcYear.Text
    
    If Val(edcMonth.Text) = 0 Then
        Select Case edcMonth.Text
        Case "Jan"
            slStartMonth = "01"
        Case "Feb"
            slStartMonth = "02"
        Case "Mar"
            slStartMonth = "03"
        Case "Apr"
            slStartMonth = "04"
        Case "May"
            slStartMonth = "05"
        Case "Jun"
            slStartMonth = "06"
        Case "Jul"
            slStartMonth = "07"
        Case "Aug"
            slStartMonth = "08"
        Case "Sep"
            slStartMonth = "09"
        Case "Oct"
            slStartMonth = "10"
        Case "Nov"
            slStartMonth = "11"
        Case "Dec"
            slStartMonth = "12"
        End Select
    Else
        slStartMonth = Format$(edcMonth.Text, "0#")
    End If
    
    slEndMonth = Val(slStartMonth) + IIF(Val(edcPeriods.Text) > 1, Val(edcPeriods.Text) - 1, 0)
    If Val(edcPeriods.Text) = 24 Then
        slEndYear = slStartYear + 2
        slEndMonth = Format((Val(slStartMonth) + 24) Mod 12, "0#")
    ElseIf slEndMonth > 12 Then
        slEndYear = CStr(Val(slEndYear) + 1)
        slEndMonth = Format((Val(slStartMonth) + 12) Mod 12, "0#")
    Else
        slEndMonth = Format(slEndMonth, "0#")
    End If
    
    mGetExportName = "MillerKaplan " & smClientName & " " & slStartMonth & slStartYear & "-" & slEndMonth & slEndYear & ".csv"
    
End Function

Private Function mGetSLFName(ByVal lSlfCode As Long) As String
    Dim rst As Recordset
    Dim slName As String
    Dim llCounter As Long
    Dim slSql As String
    
    If UBound(tmSlf) = 0 Then
        slSql = "Select slfCode, slfsofCode, ltrim(rtrim(slfFirstName)) + ' ' + ltrim(rtrim(slfLastName)) as FullName From SLF_Salespeople"
        Set rst = gSQLSelectCall(slSql)
        If Not rst.EOF Then
            Do While Not rst.EOF
                tmSlf(UBound(tmSlf)).iCode = rst!slfcode
                tmSlf(UBound(tmSlf)).iSofCode = rst!slfsofCode
                tmSlf(UBound(tmSlf)).slName = rst!FullName
                
                If lSlfCode = rst!slfcode Then slName = rst!slfsofCode & "," & rst!FullName
                
                ReDim Preserve tmSlf(0 To UBound(tmSlf) + 1)
                rst.MoveNext
            Loop
        End If
    End If
    If slName = "" Then
        For llCounter = 0 To UBound(tmSlf) - 1
            If tmSlf(llCounter).iCode = lSlfCode Then
                slName = tmSlf(llCounter).iSofCode & "," & tmSlf(llCounter).slName
                Exit For
            End If
        Next llCounter
    End If
    mGetSLFName = slName
    Set rst = Nothing
    
End Function

Private Function mGetMNFName(ByVal lMnfCode As Long) As String
    Dim rst As Recordset
    Dim slName As String
    Dim llCounter As Long
    Dim slSql As String
    
    If UBound(tmMnf2) = 0 Then
        slSql = "Select mnfCode, ltrim(rtrim(mnfName)) as mnfName From MNF_Multi_Names"
        Set rst = gSQLSelectCall(slSql)
        If Not rst.EOF Then
            Do While Not rst.EOF
                tmMnf2(UBound(tmMnf2)).lCode = rst!MNFCode
                tmMnf2(UBound(tmMnf2)).slName = rst!mnfname
                
                If lMnfCode = rst!MNFCode Then slName = rst!mnfname
                
                ReDim Preserve tmMnf2(0 To UBound(tmMnf2) + 1)
                rst.MoveNext
            Loop
        End If
    End If
    If slName = "" Then
        For llCounter = 0 To UBound(tmMnf2) - 1
            If tmMnf2(llCounter).lCode = lMnfCode Then
                slName = tmMnf2(llCounter).slName
                Exit For
            End If
        Next llCounter
    End If
    mGetMNFName = slName
    Set rst = Nothing
    
End Function


Private Function mGetPRFName(ByVal lPrfCode As Long) As String
    Dim rst As Recordset
    Dim slName As String
    Dim llCounter As Long
    Dim slSql As String
    
    If UBound(tmPrf) = 0 Then
        slSql = "Select prfCode, ltrim(rtrim(prfName)) as prfName From PRF_Product_Names"
        Set rst = gSQLSelectCall(slSql)
        If Not rst.EOF Then
            Do While Not rst.EOF
                tmPrf(UBound(tmPrf)).lCode = rst!prfCode
                tmPrf(UBound(tmPrf)).slName = rst!prfname
                
                If lPrfCode = rst!prfCode Then slName = rst!prfname
                
                ReDim Preserve tmPrf(0 To UBound(tmPrf) + 1)
                rst.MoveNext
            Loop
        End If
    End If
    If slName = "" Then
        For llCounter = 0 To UBound(tmPrf) - 1
            If tmPrf(llCounter).lCode = lPrfCode Then
                slName = tmPrf(llCounter).slName
                Exit For
            End If
        Next llCounter
    End If
    mGetPRFName = slName
    Set rst = Nothing
    
End Function


Private Function mGetSOFName(ByVal lSofCode As Integer) As String
    Dim rst As Recordset
    Dim slName As String
    Dim llCounter As Long
    Dim slSql As String
    
    If UBound(tmSof) = 0 Then
        slSql = "Select sofCode, ltrim(rtrim(sofName)) as sofName From SOF_Sales_Offices"
        Set rst = gSQLSelectCall(slSql)
        If Not rst.EOF Then
            Do While Not rst.EOF
                tmSof(UBound(tmSof)).lCode = rst!sofcode
                tmSof(UBound(tmSof)).slName = rst!sofName
                
                If lSofCode = rst!sofcode Then slName = rst!sofName
                
                ReDim Preserve tmSof(0 To UBound(tmSof) + 1)
                rst.MoveNext
            Loop
        End If
    End If
    If slName = "" Then
        For llCounter = 0 To UBound(tmSof) - 1
            If tmSof(llCounter).lCode = lSofCode Then
                slName = tmSof(llCounter).slName
                Exit For
            End If
        Next llCounter
    End If
    mGetSOFName = slName
    Set rst = Nothing
    
End Function

Private Sub cmcCancel_Click()
       If imExporting Then
           imTerminate = True
           Exit Sub
       End If
       mTerminate
End Sub

Private Sub cmcExport_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDateTime As String
    Dim slMonthHdr As String * 36
    Dim ilSaveMonth As Integer
    Dim ilYear As Integer
    Dim llStdStartDates(0 To 25) As Long   '2 years standard start dates, index zero ignored
    Dim llStartDates(0 To 25) As Long       'max 2 years, index zero ignored
    Dim llLastBilled As Long
    Dim ilLastBilledInx As Integer
    Dim slStart As String
    Dim slTimeStamp As String
    Dim ilHowManyDefined As Integer
    Dim ilHowMany As Integer
    Dim slCode As String
    Dim slNameCode As String
    
    Dim slStdStart As String
    Dim slStdEnd As String
    Dim llStdStart As Long
    Dim llStdEnd As Long
    Dim ilFirstProjInx As Integer
    Dim slRepeat As String
    
    ReDim tmSlf(0)
    ReDim tmMnf2(0)
    ReDim tmPrf(0)
    ReDim tmSof(0)
    ReDim tmAgf(0)
    ReDim tmMNFCODE(0)
    
    lacInfo(0).Visible = False
    lacInfo(1).Visible = False
    lacInfo(0).Caption = ""

    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError

    'Verify data input
    slStr = ExpMK!edcYear.Text
    ilYear = gVerifyYear(slStr)
    If ilYear = 0 Then
        ExpMK!edcYear.SetFocus                      'invalid year
        ''MsgBox "Year is Not Valid", vbOkOnly + vbApplicationModal, "Start Year"
        gAutomationAlertAndLogHandler "Year is Not Valid", vbOkOnly + vbApplicationModal, "Start Year"
        Exit Sub
    End If

    slMonthHdr = "JanFebMarAprMayJunJulAugSepOctNovDec"
    slStr = ExpMK!edcMonth.Text                     'month in text form (jan..dec, or 1-12
    gGetMonthNoFromString slStr, ilSaveMonth        'getmonth #
    If ilSaveMonth = 0 Then                         'input isn't text month name, try month #
        ilSaveMonth = Val(slStr)
        ilRet = gVerifyInt(slStr, 1, 12)
        If ilRet = -1 Then
            ExpMK!edcMonth.SetFocus                 'invalid month entry
            ''MsgBox "Month is Not Valid", vbOkOnly + vbApplicationModal, "Start Month"
            gAutomationAlertAndLogHandler "Monthis Not Valid", vbOkOnly + vbApplicationModal, "Start Month"
            Exit Sub
        End If
    End If
    
    slStr = ExpMK!edcPeriods.Text                   '#periods
    igPeriods = Val(slStr)
    ilRet = gVerifyInt(slStr, 1, 24)
    If ilRet = -1 Then
        ExpMK!edcPeriods.SetFocus
        ''MsgBox "# months must be between 1 and 24", vbOkOnly + vbApplicationModal, "Number Months"
        gAutomationAlertAndLogHandler "# months must be between 1 and 24", vbOkOnly + vbApplicationModal, "Number Months"
        Exit Sub
    End If

    lmCntrNo = 0                                    'this is for debugging on a single contract
    slStr = ExpMK!edcContract
    If slStr <> "" Then
        If Val(slStr) = 0 Then
            ExpMK!edcContract.Text = ""
            ExpMK!edcContract.SetFocus
            ''MsgBox "Contract # must be numeric", vbOkOnly + vbApplicationModal, "Contract Number"
            gAutomationAlertAndLogHandler "Contract # must be numeric", vbOkOnly + vbApplicationModal, "Contract Number"
            Exit Sub
        Else
            lmCntrNo = Val(slStr)
        End If
    End If

    'create export file name
    smExportName = mGetExportName
    'smExportName = "MillerKaplan" & smClientName & "201901-201902.csv"
    If (InStr(smExportName, ":") = 0) And (Left$(smExportName, 2) <> "\\") Then
        smExportName = Trim$(sgExportPath) & smExportName
        ilRet = gFileExist(smExportName)            'this function is not working right; should reverse the returned value
        If ilRet = 0 Then
            slRepeat = "A"
            Do
                ilRet = 0
                'On Error GoTo cmcExportDupNameErr:
                If slRepeat = "A" Then
                    smExportName = Left$(smExportName, Len(smExportName) - 4)
                Else
                    smExportName = Left$(smExportName, Len(smExportName) - 5)
                End If
                smExportName = Trim$(smExportName) & " " & slRepeat & ".csv"
                'slDateTime = FileDateTime(smExportName)
                ilRet = gFileExist(smExportName)
                If ilRet = 0 Then                   'if went to mOpenErr , there was a filename that existed with same name. Increment the letter
                    slRepeat = Chr(Asc(slRepeat) + 1)
                End If
            Loop While ilRet = 0
        End If
    End If

    If Not mOpenMsgFile() Then                      'open message file
         cmcCancel.SetFocus
         Exit Sub
    End If
    On Error GoTo 0
    ilRet = 0
    
    'Open smExportName For Output As hmMK
    ilRet = gFileOpen(smExportName, "Output", hmMK)
    If ilRet <> 0 Then
        'Print #hmMsg, "** Terminated **"
        gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Close #hmMK
        imExporting = False
        Screen.MousePointer = vbDefault
        'TTP 10011 - Error.Numner prevents MsgBox.  Additionally the Error # is stored in ilRet.
        'MsgBox "Open Error #" & str$(Error.Numner) & smExportName, vbOkOnly, "Open Error"
        ''MsgBox "Open Error #" & str$(ilRet) & " - " & smExportName, vbOkOnly, "Open Error"
        gAutomationAlertAndLogHandler "Open Error #" & str$(ilRet) & " - " & smExportName, vbOkOnly, "Open Error"
        Exit Sub
    End If
    'Print #hmMsg, "** Storing Output into " & smExportName & " **"
    gAutomationAlertAndLogHandler "* Storing Output into " & smExportName
    Screen.MousePointer = vbHourglass
    imExporting = True
    imFirstTime = True

    'assume standard (broadcast calendar) exporting
    slStart = str$(ilSaveMonth) & "/15/" & str$(ilYear)
    gBuildStartDates slStart, 1, igPeriods + 1, llStdStartDates()       'build array of std start & end dates

    ReDim llContractNo(0) As MKContracts
    
    ilRet = 0                                                           'indicates successful export file creation
    
    'loop thru the number of months (periods)
    For ilLoop = 1 To igPeriods Step 1
        If llStdStartDates(ilLoop) > 0 Then
            slStdStart = Format$(llStdStartDates(ilLoop), "m/d/yyyy")       'assume first date of proj is the quarter entered
            slStdEnd = Format$(llStdStartDates(ilLoop + 1), "m/d/yyyy")
            
            'slStdEnd = Format$(llStdStartDates(igPeriods + 1), "m/d/yyyy") 'force a 24 period -- only for testing
            If Not mObtainPhfAndRvf(slStdStart, slStdEnd) Then
                ilRet = 1
                Exit For
            End If
        End If
    Next ilLoop
    
    Close #hmMK
    mCloseMKFiles
    Erase llStdStartDates
    Screen.MousePointer = vbDefault

    If ilRet = 0 Then           'true is successful
        lacInfo(0).Caption = "Export " & Trim$(smExportOptionName) & " Successfully Completed"
        'Print #hmMsg, "** Export " & Trim$(smExportOptionName) & " Successfully completed : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        gAutomationAlertAndLogHandler "** Export " & Trim$(smExportOptionName) & " Successfully completed : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    Else
        lacInfo(0).Caption = "Export Failed"
        'Print #hmMsg, "** Export Failed **"
        gAutomationAlertAndLogHandler "Export Failed"
    End If
    lacInfo(0).Visible = True
    Close #hmMsg
    cmcCancel.Caption = "&Done"
    If igExportType <= 1 Then       'ok to set focus if manual mode
        cmcCancel.SetFocus
    End If
    Screen.MousePointer = vbDefault
    imExporting = False
    Exit Sub
    
ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    
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
'*******************************************************
Private Function mOpenMsgFile()
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    Dim slNTR  As String
    Dim slCntr As String
    Dim slMissed As String
    Dim slMonthType As String
    Dim slAdj As String

    ilRet = 0
    'On Error GoTo mOpenMsgFileErr:
    slToFile = sgDBPath & "\Messages\" & "Exp" & Trim$(smExportOptionName) & ".Txt"
    sgMessageFile = slToFile
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
            'ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbApplicationModal, "Open Error"
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
            'ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
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
        'ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    'Print #hmMsg, ""
    If edcContract.Text = "" Then
        slCntr = "All contracts"
    Else
        slCntr = "Cntr # " & edcContract.Text
    End If
    
    'Print #hmMsg, "** Export Matrix" & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " for " & edcMonth.Text & " " & edcYear.Text & " "; edcNoMonths.Text & " months, " & slNTR & ", " & slMissed & ", "  & slAdj & ", " & slCntr & " **"
    'Print #hmMsg, "** Export " & Trim$(smExportOptionName) & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " for " & slMonthType & " " & edcMonth.Text & " " & edcYear.Text & " "; edcPeriods.Text & " months, " & slNTR & ", " & slMissed & ", " & slCntr & " **"
    gAutomationAlertAndLogHandler "** Export " & Trim$(smExportOptionName) & " **" '& ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " for " & slMonthType & " " & edcMonth.Text & " " & edcYear.Text & " " & edcPeriods.Text & " months, " & slNTR & ", " & slMissed & ", " & slCntr & " **"
    gAutomationAlertAndLogHandler "* MonthType = " & slMonthType
    gAutomationAlertAndLogHandler "* Month = " & edcMonth.Text
    gAutomationAlertAndLogHandler "* Year = " & edcYear.Text
    gAutomationAlertAndLogHandler "* # Periods = " & edcPeriods.Text
    'gAutomationAlertAndLogHandler "* NTR = " & IIF(slNTR = "", "False", "True")
    'gAutomationAlertAndLogHandler "* Missed = " & IIF(slMissed = "", "False", "True")
    gAutomationAlertAndLogHandler "* Contract = " & slCntr
    
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function


Sub mCloseMKFiles()
    Dim ilRet As Integer
    
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
End Sub
Private Function mObtainPhfAndRvf(ByVal slStartPeriod As String, ByVal slEndPeriod As String) As Boolean
    Dim MK_rst As ADODB.Recordset
    Dim MK_rvf As ADODB.Recordset
    Dim MK_phf As ADODB.Recordset

    Dim slSql As String
    Dim llCounter As Long
    Dim slClientName As String
    Dim slStation As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slAgfName As String
    Dim slAdfName  As String
    Dim slPrfName  As String
    Dim slFullName As String
    'Dim slLastName  As String
    Dim slCashTrade  As String
    Dim slRevenueType  As String
    Dim llGrossAmount As Long
    Dim llTotalGross As Long
    Dim slMnfName As String
    Dim slSofName As String
    Dim blProcessed As Boolean
    Dim blFoundContract As Boolean
    Dim slTranDate As String
    Dim slTemp As String
    Dim iMnfCode As Integer
    Dim ilRet As Integer
    Dim ilInx As Integer
    Dim slStr As String
    
    On Error GoTo mObtainPhfAndRvf_Err
    
    ReDim tmExportInfo(0) As MKExportInfo
    ReDim tmRvf(0) As MKExportInfo
    ReDim tmPhf(0) As MKExportInfo
    ReDim llContractNo(0) As MKContracts
    
    ReDim ilAdfCode(0) As Integer

    mObtainPhfAndRvf = False

    slSql = "Select spfGClient From SPF_Site_Options"
    
    Set MK_rst = gSQLSelectCall(slSql)
    If Not MK_rst.EOF Then
        Do While Not MK_rst.EOF
            slStation = MK_rst!spfGClient
            MK_rst.MoveNext
        Loop
    End If
    
    'PSA and Promo contracts are ignored
    'generate two records for part cash / part trade
    slSql = "Select rvf.rvfCntrNo, adf.adfCode, rvf.rvfTranType, " & _
            "rvf.rvfCashTrade, rvf.rvfInvNo, ltrim(rtrim(adf.adfName)) as adfName, " & _
            "if (rvf.rvfCashTrade = 'C', 'Cash', 'Trade') as " & Chr(34) & "CashTrade" & Chr(34) & ", " & _
            "if (rvf.rvfmnfItem = 0, 'Air Time', 'NTR') as " & Chr(34) & "RevenueType" & Chr(34) & ", " & _
            "rvf.rvfTranDate, rvf.rvfGross, rvfagfCode, rvfPrfCode, rvfslfCode " & _
            "From RVF_Receivables rvf " & _
            "left join ADF_Advertisers adf on rvf.rvfadfCode = adf.adfCode " & _
            "Where rvf.rvfTranDate >= '" & CStr(Year(slStartPeriod)) & "-" & CStr(Month(slStartPeriod)) & "-" & CStr(Day(slStartPeriod)) & "' and rvf.rvfTranDate < '" & _
                                         CStr(Year(slEndPeriod)) & "-" & CStr(Month(slEndPeriod)) & "-" & CStr(Day(slEndPeriod)) & "' " & _
            "and rvf.rvfTranType in ('IN', 'AN','HI') " & _
            "and rvf.rvfCashTrade not in ('P','M') "
    
    'for testing one advertiser
    'slSql = slSql & " and adf.adfName = 'Boll & Branch' "
    
    'for testing one contract
    If ExpMK!edcContract <> "" Then slSql = slSql & " and rvf.rvfCntrNo = " & ExpMK!edcContract
    
    slSql = slSql & " order by rvf.rvfCntrNo "

    'RVF contract collection
    Set MK_rvf = Nothing
    Set MK_rvf = gSQLSelectCall(slSql)
    If Not MK_rvf.EOF Then
        MK_rvf.MoveFirst
        Do While Not MK_rvf.EOF
            tmRvf(UBound(tmRvf)).slAgency = ""
            ilInx = gBinarySearchAgf(MK_rvf!rvfagfcode)
            If ilInx > 0 Then
                tmRvf(UBound(tmRvf)).slAgency = IIF(ilInx >= 0, Trim$(tgCommAgf(ilInx).sName), "")
            End If
            If Not IsNull(MK_rvf!adfName) Then tmRvf(UBound(tmRvf)).sAdvertiser = Trim$(MK_rvf!adfName)
            tmRvf(UBound(tmRvf)).slBrand = mGetPRFName(MK_rvf!rvfprfcode)
            
            slTemp = mGetSLFName(MK_rvf!rvfslfcode)
            tmRvf(UBound(tmRvf)).sAEFullName = Mid(slTemp, InStr(1, slTemp, ",") + 1) 'IIF(ilInx >= 0, Trim$(tgMSlf(ilInx).sFirstName) & " " & Trim$(tgMSlf(ilInx).sLastName), "")
            tmRvf(UBound(tmRvf)).sDirectOffice = mGetSOFName(Left(slTemp, InStr(1, slTemp, ",") - 1))
            
            If Not IsNull(MK_rvf!CashTrade) Then tmRvf(UBound(tmRvf)).sAccountType = MK_rvf!CashTrade
            If Not IsNull(MK_rvf!RevenueType) Then tmRvf(UBound(tmRvf)).sRevenueType = MK_rvf!RevenueType
            tmRvf(UBound(tmRvf)).sProductCodeDesc = mGetMNFName(mGetMNFCode(MK_rvf!rvfCntrNo))

            'Date: 8/26/2019 - use end broacast end date to display in report instead of transaction date   FYM
            tmRvf(UBound(tmRvf)).sTranDate = Format(DateAdd("d", -1, slEndPeriod), "m/d/yyyy") 'Format(MK_rvf!rvfTranDate, "m/d/yyyy")
            
            'llGrossAmount = MK_rvf!rvfGross
            tmRvf(UBound(tmRvf)).lGrossAmount = gStrDecToLong(MK_rvf!rvfGross, 2)
            
            tmRvf(UBound(tmRvf)).lContractNo = MK_rvf!rvfCntrNo
            tmRvf(UBound(tmRvf)).iAdfCode = MK_rvf!adfCode
            
            'pad contract number with zeroes for sorting
            slStr = Trim$(str$(MK_rvf!rvfCntrNo))
            Do While Len(slStr) < 8
                slStr = "0" & slStr
            Loop
            
            tmRvf(UBound(tmRvf)).sCntrKey = slStr
            tmRvf(UBound(tmRvf)).sAdvKey = MK_rvf!adfCode
            
            ReDim Preserve tmRvf(0 To UBound(tmRvf) + 1)
            
            'make sure contract number zero is included
            If MK_rvf!rvfCntrNo = 0 Then
                For llCounter = 0 To UBound(llContractNo)
                    blFoundContract = False
                    If UBound(llContractNo) = 0 Then
                        Exit For
                    ElseIf llContractNo(llCounter).lContractNo = 0 And llCounter >= 0 Then
                        blFoundContract = True
                        Exit For
                    ElseIf llContractNo(llCounter).lContractNo = 0 Then
                        blFoundContract = True
                        Exit For
                    End If
                Next llCounter
            Else
                For llCounter = 0 To UBound(llContractNo)
                    blFoundContract = False
                    If llContractNo(llCounter).lContractNo = MK_rvf!rvfCntrNo Then
                        blFoundContract = True
                        Exit For
                    End If
                Next llCounter
            End If
            If Not blFoundContract Then
                llContractNo(UBound(llContractNo)).sKey = slStr
                llContractNo(UBound(llContractNo)).lContractNo = MK_rvf!rvfCntrNo
                ReDim Preserve llContractNo(0 To UBound(llContractNo) + 1) As MKContracts
            End If
            
            MK_rvf.MoveNext
        Loop
    End If
    
    'PSA and Promo contracts are ignored
    'generate two records for part cash / part trade
    slSql = "Select phf.phfCntrNo, adf.adfCode, phf.phfTranType, " & _
            "phf.phfCashTrade, phf.phfInvNo, ltrim(rtrim(adf.adfName)) as adfName, " & _
            "if (phf.phfCashTrade = 'C', 'Cash', 'Trade') as " & Chr(34) & "CashTrade" & Chr(34) & ", " & _
            "if (phf.phfmnfItem = 0, 'Air Time', 'NTR') as " & Chr(34) & "RevenueType" & Chr(34) & ", " & _
            "phf.phfTranDate, phf.phfGross, phfagfCode, phfPrfCode, phfslfCode " & _
            "From PHF_Payment_History phf " & _
            "left join ADF_Advertisers adf on phf.phfadfCode = adf.adfCode " & _
            "Where phf.phfTranDate >= '" & CStr(Year(slStartPeriod)) & "-" & CStr(Month(slStartPeriod)) & "-" & CStr(Day(slStartPeriod)) & "' and phf.phfTranDate < '" & _
                                         CStr(Year(slEndPeriod)) & "-" & CStr(Month(slEndPeriod)) & "-" & CStr(Day(slEndPeriod)) & "' " & _
            "and phf.phfTranType in ('IN', 'AN','HI') " & _
            "and phf.phfCashTrade not in ('P','M') "
            
    'for testing one advertiser only
    'slSql = slSql & " and adf.adfName = 'Boll & Branch' "

    'for testing one contract
    If ExpMK!edcContract <> "" Then slSql = slSql & " and phf.phfCntrNo = " & ExpMK!edcContract

    slSql = slSql & " order by phf.phfCntrNo"

    'PHF contract collection
    Set MK_phf = Nothing
    Set MK_phf = gSQLSelectCall(slSql)
    If Not MK_phf.EOF Then
        MK_phf.MoveFirst
        Do While Not MK_phf.EOF
            tmRvf(UBound(tmRvf)).slAgency = ""
            ilInx = gBinarySearchAgf(MK_phf!phfagfcode)
            If ilInx > 0 Then
                tmRvf(UBound(tmRvf)).slAgency = IIF(ilInx >= 0, Trim$(tgCommAgf(ilInx).sName), "")
            End If

            If Not IsNull(MK_phf!adfName) Then tmRvf(UBound(tmRvf)).sAdvertiser = Trim$(MK_phf!adfName)
            tmRvf(UBound(tmRvf)).slBrand = mGetPRFName(MK_phf!phfprfcode)

            slTemp = mGetSLFName(MK_phf!phfslfcode)
            tmRvf(UBound(tmRvf)).sAEFullName = Mid(slTemp, InStr(1, slTemp, ",") + 1) 'IIF(ilInx >= 0, Trim$(tgMSlf(ilInx).sFirstName) & " " & Trim$(tgMSlf(ilInx).sLastName), "")
            tmRvf(UBound(tmRvf)).sDirectOffice = mGetSOFName(Left(slTemp, InStr(1, slTemp, ",") - 1))

            If Not IsNull(MK_phf!CashTrade) Then tmRvf(UBound(tmRvf)).sAccountType = MK_phf!CashTrade
            If Not IsNull(MK_phf!RevenueType) Then tmRvf(UBound(tmRvf)).sRevenueType = MK_phf!RevenueType
            tmRvf(UBound(tmRvf)).sProductCodeDesc = mGetMNFName(mGetMNFCode(MK_phf!phfCntrNo))

            'Date: 8/26/2019 - use end broacast end date to display in report instead of transaction date   FYM
            tmRvf(UBound(tmRvf)).sTranDate = Format(DateAdd("d", -1, slEndPeriod), "m/d/yyyy") 'Format(MK_phf!phfTranDate, "m/d/yyyy")

            'llGrossAmount = tmRvf(UBound(tmRvf)).lGrossAmount
            tmRvf(UBound(tmRvf)).lGrossAmount = gStrDecToLong(MK_phf!phfGross, 2)
            tmRvf(UBound(tmRvf)).lContractNo = MK_phf!phfCntrNo
            tmRvf(UBound(tmRvf)).iAdfCode = MK_phf!adfCode
            
            'pad contract number with zeroes for sorting
            slStr = Trim$(str$(MK_phf!phfCntrNo))
            Do While Len(slStr) < 8
                slStr = "0" & slStr
            Loop
            tmRvf(UBound(tmRvf)).sCntrKey = slStr
            tmRvf(UBound(tmRvf)).sAdvKey = MK_phf!adfCode

            ReDim Preserve tmRvf(0 To UBound(tmRvf) + 1)

            'make sure contract number zero is included
            If MK_phf!phfCntrNo = 0 Then
                For llCounter = 0 To UBound(llContractNo) - 1
                    blFoundContract = False
                    If UBound(llContractNo) = 0 Then
                        Exit For
                    ElseIf llContractNo(llCounter).lContractNo = 0 And llCounter >= 0 Then
                        blFoundContract = True
                        Exit For
                    ElseIf llContractNo(llCounter).lContractNo = 0 Then
                        blFoundContract = True
                        Exit For
                    End If
                Next llCounter
            Else
                blFoundContract = False
                For llCounter = 0 To UBound(llContractNo)
                    If llContractNo(llCounter).lContractNo = MK_phf!phfCntrNo Then
                        blFoundContract = True
                        Exit For
                    End If
                Next llCounter
            End If
            If Not blFoundContract Then
                llContractNo(UBound(llContractNo)).sKey = slStr
                llContractNo(UBound(llContractNo)).lContractNo = MK_phf!phfCntrNo
                ReDim Preserve llContractNo(0 To UBound(llContractNo) + 1) As MKContracts
            End If

            MK_phf.MoveNext
        Loop
    End If
    
    If UBound(tmRvf) - 1 > 0 Then
        ArraySortTyp fnAV(tmRvf(), 0), UBound(tmRvf), 0, LenB(tmRvf(0)), 0, LenB(tmRvf(0).sCntrKey), 0
    End If
    
    If UBound(llContractNo) - 1 > 0 Then
        ArraySortTyp fnAV(llContractNo(), 0), UBound(llContractNo), 0, LenB(llContractNo(0)), 0, LenB(llContractNo(0).sKey), 0
    End If
    
    'loop thru the array of contracts and create the export records
    For llCounter = 0 To UBound(llContractNo) - 1
        If UBound(tmRvf) > 0 Then
            mCreateExportInfo llContractNo(llCounter).lContractNo, tmRvf, slStation
        End If
    Next llCounter
    
    'write records to export file
    If UBound(tmExportInfo) > 0 Then ilRet = mWriteExportRec
    
    If ilRet = 0 Then mObtainPhfAndRvf = True
    
    Set MK_rst = Nothing
    Set MK_rvf = Nothing
    Set MK_phf = Nothing
    Exit Function
    
mObtainPhfAndRvf_Err:
    ilRet = 1
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
    Dim slDay As String
    Dim slMonth As String
    Dim slYear As String
    Dim ilMonth As Integer
    Dim ilYear As Integer
'    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilVff As Integer
    Dim ilLoop As Integer
    Dim slLocation As String
    Dim slReturn As String * 130
    Dim slFileName As String
    Dim lmNowDate As Long

    slMonthStr = "JanFebMarAprMayJunJulAugSepOctNovDec"
    imTerminate = False
    imFirstActivate = True
    Screen.MousePointer = vbHourglass
    imExporting = False
    lmNowDate = gDateValue(Format$(gNow(), "m/d/yy"))


    gCenterStdAlone ExpMK
    
    imExportOption = ExportList!lbcExport.ItemData(ExportList!lbcExport.ListIndex)
    smExportOptionName = "MillerKaplan"
    
    ilRet = gObtainAgency()         'Build into tgCommAgf
    If ilRet = False Then
        imTerminate = True
    End If
    
    ilRet = gObtainSalesperson()    'Build into tgMSlf
    If ilRet = False Then
        imTerminate = True
    End If
    
    ilRet = gBuildAcqCommInfo(ExpMK)
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf)", ExpMK

    imMnfRecLen = Len(tmMnf)

    'determine default month year
    slDate = Format$(lmNowDate, "m/d/yy")
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    'Default to last month, based on today's date
    If Val(slMonth) = 1 Then
        ilMonth = 12
        ilYear = Val(slYear) - 1
    Else
        ilMonth = Val(slMonth) - 1
        ilYear = Val(slYear)
    End If

    edcMonth.Text = Mid$(slMonthStr, (ilMonth - 1) * 3 + 1, 3)
    edcYear.Text = Trim$(str$(ilYear))
     
     On Error GoTo mObtainIniValuesErr
    'find exports.ini
     sgIniPath = gSetPathEndSlash(sgIniPath, True)
     If igDirectCall = -1 Then
         slFileName = sgIniPath & "Exports.Ini"
     Else
         slFileName = CurDir$ & "\Exports.Ini"
     End If
     
     On Error Resume Next
     ilRet = GetPrivateProfileString(sgExportIniSectionName, "Months", "Not Found", slReturn, 128, slFileName)
     If Left$(slReturn, ilRet) = "Not Found" Then
         edcPeriods.Text = 1
     Else
         slCode = Trim$(gStripChr0(slReturn))
         If Val(slCode) = 0 Or Val(slCode) > 24 Then     'max 24 months
             edcPeriods.Text = 1                         'invalid input , take default of 1 week
         Else
             edcPeriods.Text = Trim$(gStripChr0(slReturn))
         End If
     End If
     
    On Error Resume Next
    ilRet = GetPrivateProfileString(sgExportIniSectionName, "Export", "Not Found", slReturn, 128, slFileName)
    If Left$(slReturn, ilRet) = "Not Found" Then
        'default to the export path
        sgExportPath = sgExportPath
    Else
        sgExportPath = Trim$(gStripChr0(slReturn))
    End If
    sgExportPath = gSetPathEndSlash(sgExportPath, True)
    
    smClientName = Trim$(tgSpf.sGClient)
    If tgSpf.iMnfClientAbbr > 0 Then
        tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            smClientName = Trim$(tmMnf.sName)
        End If
    End If

    Screen.MousePointer = vbDefault
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
    
    Exit Sub

mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub

mObtainIniValuesErr:
    Resume Next

End Sub





'
'
'           mWriteExportRec - gather all the information for a month and write
'           a record to the export .csv file
'
'           <input> tmExportInfo() - structure containing all the info required to write up to 24 months of data from
'                                    either the receivables or contract files (PSA/Promo excluded)
'           Return - true if error, otherwise false
Private Function mWriteExportRec() As Integer
    Dim slStation As String
    Dim slAgency As String
    Dim slAdvertiser As String
    Dim slBrand As String
    Dim slFormat As String
    Dim slAEFullName As String
    Dim slAccountType As String
    Dim slProductCodeDesc As String
    Dim slRevenueType As String
    Dim slDirectOffice As String
    Dim slYearMonth As String
    Dim lGrossAmount As Long
    Dim slStr As String
    
    Dim ilError As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    Dim slPrimComp As String
    Dim slSecComp As String
    Dim slSlsp As String
    Dim slOffice As String
    Dim slSS As String
    Dim ilOfficeInx As Integer
    
    Dim ilRemainder As Integer
    Dim slStripCents As String
    Dim llCounter As Long
    
    ilError = False
    If imFirstTime Then         'create the header record
        slStr = "ClientName, Agency, Advertiser, Product, Salesperson, Cash/Trade, Product Code Description, AirTime/NTR, Sales Office, Standard Month, Gross"
        
        On Error GoTo mWriteExportRecErr
        Print #hmMK, slStr      'write header description
        On Error GoTo 0

'        slStr = "As of " & Format$(gNow(), "mm/dd/yy") & " "
'        slStr = slStr & Format$(gNow(), "h:mm:ssAM/PM")
'
'        On Error GoTo mWriteExportRecErr
'        Print #hmMK, slStr      'write header description
'        On Error GoTo 0
        imFirstTime = False     'do the heading and time stamp only once
    End If

    For llCounter = 0 To UBound(tmExportInfo) - 1
        slStation = ""
        slAgency = ""
        slAdvertiser = ""
        slBrand = ""
        slFormat = ""
        slAEFullName = ""
        slAccountType = ""
        slProductCodeDesc = ""
        slRevenueType = ""
        slDirectOffice = ""
        slYearMonth = ""
        lGrossAmount = 0
        
        'ClientName, Agency, Advertiser, Product, Salesperson, Cash/Trade, Product Code Description, AirTime/NTR, Standard Month, Gross Amount
        slStation = tmExportInfo(llCounter).sStation
        slAgency = tmExportInfo(llCounter).slAgency
        slAdvertiser = tmExportInfo(llCounter).sAdvertiser
        slBrand = tmExportInfo(llCounter).slBrand
        'slFormat = tmExportInfo(llCounter).`
        slAEFullName = tmExportInfo(llCounter).sAEFullName
        slAccountType = tmExportInfo(llCounter).sAccountType
        slProductCodeDesc = tmExportInfo(llCounter).sProductCodeDesc
        slRevenueType = tmExportInfo(llCounter).sRevenueType
        slDirectOffice = tmExportInfo(llCounter).sDirectOffice
        slYearMonth = tmExportInfo(llCounter).slYearMonth
        lGrossAmount = tmExportInfo(llCounter).lGrossAmount
        
        slStr = Chr(34) & Trim$(slStation) & Chr(34) & ","
        slStr = slStr & Chr(34) & Trim$(slAgency) & Chr(34) & ","
        slStr = slStr & Chr(34) & Trim$(slAdvertiser) & Chr(34) & ","
        slStr = slStr & Chr(34) & Trim$(slBrand) & Chr(34) & ","
        'slStr = slStr & Chr(34) & Trim$(slFormat) & Chr(34) & ","
        slStr = slStr & Chr(34) & Trim$(slAEFullName) & Chr(34) & ","
        slStr = slStr & Chr(34) & Trim$(slAccountType) & Chr(34) & ","
        slStr = slStr & Chr(34) & Trim$(slProductCodeDesc) & Chr(34) & ","
        slStr = slStr & Chr(34) & Trim$(slRevenueType) & Chr(34) & ","
        slStr = slStr & Chr(34) & Trim$(slDirectOffice) & Chr(34) & ","
        slStr = slStr & Chr(34) & Trim$(slYearMonth) & Chr(34) & ","

        'check if pennies are present in the amount
        
        ilRemainder = lGrossAmount Mod 100
        If ilRemainder = 0 Then         'strip off the pennies if whole number
            slStripCents = Trim$(gLongToStrDec(lGrossAmount, 2))
            slStr = slStr & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3))
        Else
            slStr = slStr & Trim$(gLongToStrDec(lGrossAmount, 2))
        End If
               
        On Error GoTo mWriteExportRecErr
        Print #hmMK, slStr
        On Error GoTo 0
    Next llCounter

    mWriteExportRec = ilError
    Exit Function

mWriteExportRecErr:
    ilError = True
    Resume Next

End Function

Private Sub edcYear_GotFocus()
    gCtrlGotFocus edcYear
End Sub

Private Sub edcMonth_GotFocus()
    gCtrlGotFocus edcMonth
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    DoEvents    'Process events so pending keys are not sent to this
    Me.KeyPreview = True
    Me.Refresh
    edcYear.SetFocus
End Sub
Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
'        gFunctionKeyBranch KeyCode
'    End If
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
    
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf

    Erase tmExportInfo()
    Erase tmRvf()
    Erase tmPhf()
    Erase tmAgf()
    Erase tmSof()
    Erase tmMnf2()
    Erase tmPrf()
    Erase tmMNFCODE()

    Set ExpMK = Nothing   'Remove data segment

End Sub

Private Sub tmcCancel_Timer()
    tmcCancel.Enabled = False       'screen has now been focused to show
    cmcCancel_Click         'simulate clicking of cancen button
End Sub
Private Sub mTerminate()
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExpMK
    igManUnload = NO
End Sub
