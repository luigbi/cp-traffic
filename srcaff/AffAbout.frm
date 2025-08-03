VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Affiliate"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   ClipControls    =   0   'False
   Icon            =   "AffAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Tag             =   "About Affiliate"
   Begin VB.PictureBox pbcIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   120
      Picture         =   "AffAbout.frx":08CA
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   510
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3930
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   3360
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   2010
      TabIndex        =   1
      Tag             =   "&System Info..."
      Top             =   3360
      Width           =   1365
   End
   Begin VB.Label lblAffRequired 
      BackColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   2040
      TabIndex        =   10
      Tag             =   "Version"
      Top             =   2760
      Width           =   1980
   End
   Begin VB.Label lblAffExpectedVersion 
      BackColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Tag             =   "Version"
      Top             =   2760
      Width           =   1605
   End
   Begin VB.Label lacWebInfo 
      BackColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   1980
      TabIndex        =   8
      Tag             =   "Version"
      Top             =   2280
      Width           =   1980
   End
   Begin VB.Label lblWebVersion 
      BackColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Tag             =   "Version"
      Top             =   2280
      Width           =   1605
   End
   Begin VB.Label lacDDFInfo 
      BackColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   1980
      TabIndex        =   6
      Tag             =   "Version"
      Top             =   1800
      Width           =   2100
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Copyright 2003-2015                       Counterpoint Software, Inc."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   120
      TabIndex        =   5
      Tag             =   "App Description"
      Top             =   840
      Width           =   6960
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Affiliate Relations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   720
      TabIndex        =   4
      Tag             =   "Application Title"
      Top             =   120
      Width           =   2445
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Index           =   1
      X1              =   120
      X2              =   7050
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Tag             =   "Version"
      Top             =   1800
      Width           =   1605
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dan test
' Reg Key Security Options...
Const KEY_ALL_ACCESS = &H2003F
                                          

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub Form_Load()
    Dim slRevision As String
    frmAbout.Caption = "About Affiliiate - " & sgClientName
    slRevision = App.Revision
    If Len(slRevision) = 3 Then
        slRevision = "0" & slRevision
    End If
    lblVersion.Caption = "Exe Version " & App.Major & "." & App.Minor & " B" & slRevision & " " & App.FileDescription
    lblDescription.Caption = "Copyrighted ©2003-" & Year(Now) & " Counterpoint Software, Inc. ®" & sgCR & sgLF & "All Rights Reserved"
    
    If Not gIsUsingNovelty Then
        If gUsingWeb Then
            lblWebVersion.Caption = "Web Version " & sgWebSiteVersion
            lacWebInfo.Caption = Format(sgWebSiteDate, "ddddd ttttt")
        Else
            lblWebVersion.Caption = "Web Version - N/A"
        End If
        lblAffExpectedVersion.Caption = sgWebSiteExpectedByAffiliate
    End If
    lblAffRequired.Caption = ""
    lblTitle.Caption = App.Title
    lacDDFInfo.Caption = "DDF Info: " & Format(sgDDFDateInfo, "ddddd ttttt")
End Sub





Private Sub cmdSysInfo_Click()
        Call StartSysInfo
End Sub


Private Sub cmdOk_Click()
        Unload Me
End Sub


Public Sub StartSysInfo()
     
        Dim rc As Long
        Dim SysInfoPath As String
        
        On Error GoTo SysInfoErr
        
        ' Try To Get System Info Program Path\Name From Registry...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' Validate Existance Of Known 32 Bit File Version
                '8886
                'If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                If gFileExist(SysInfoPath & "\MSINFO32.EXE") = FILEEXISTS Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                ' Error - File Can Not Be Found...
                Else
                        GoTo SysInfoErr
                End If
        ' Error - Registry Entry Can Not Be Found...
        Else
                GoTo SysInfoErr
        End If
        
        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        gMsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Loop Counter
        Dim rc As Long                                          ' Return Code
        Dim hKey As Long                                        ' Handle To An Open Registry Key
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Data Type Of A Registry Key
        Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
        Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
        '------------------------------------------------------------
        ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
        

        tmpVal = String$(1024, 0)                             ' Allocate Variable Space
        KeyValSize = 1024                                       ' Mark Variable Size
        

        '------------------------------------------------------------
        ' Retrieve Registry Key Value...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
        

        If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
                tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
        Else                                                    ' WinNT Does NOT Null Terminate String...
                tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
        End If
        '------------------------------------------------------------
        ' Determine Key Value Type For Conversion...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' Search Data Types...
        Case REG_SZ                                             ' String Registry Key Data Type
                KeyVal = tmpVal                                     ' Copy String Value
        Case REG_DWORD                                          ' Double Word Registry Key Data Type
                For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
        End Select
        

        GetKeyValue = True                                      ' Return Success
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
        Exit Function                                           ' Exit
        

GetKeyError:    ' Cleanup After An Error Has Occured...
        KeyVal = ""                                             ' Set Return Val To Empty String
        GetKeyValue = False                                     ' Return Failure
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmAbout = Nothing
End Sub

