VERSION 5.00
Begin VB.Form ClosestBooks 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5760
   ClientLeft      =   2640
   ClientTop       =   2610
   ClientWidth     =   10095
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5760
   ScaleWidth      =   10095
   Begin VB.CommandButton cmcSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   6495
      TabIndex        =   14
      Top             =   5355
      Width           =   1050
   End
   Begin VB.PictureBox pbcIncludeToExclude 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   180
      Left            =   5280
      Picture         =   "ClosestBooks.frx":0000
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2310
      Width           =   180
   End
   Begin VB.PictureBox pbcExcludeToInclude 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   180
      Left            =   4710
      Picture         =   "ClosestBooks.frx":00DA
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3360
      Width           =   180
   End
   Begin VB.CommandButton cmcIncludeToExclude 
      Appearance      =   0  'Flat
      Caption         =   "M&ove   "
      Height          =   300
      Left            =   4590
      TabIndex        =   13
      Top             =   2250
      Width           =   945
   End
   Begin VB.CommandButton cmcExcludeToInclude 
      Appearance      =   0  'Flat
      Caption         =   "    Mo&ve"
      Height          =   300
      Left            =   4590
      TabIndex        =   12
      Top             =   3300
      Width           =   945
   End
   Begin VB.PictureBox plcExclude 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4170
      Left            =   6045
      ScaleHeight     =   4110
      ScaleWidth      =   3660
      TabIndex        =   8
      Top             =   960
      Width           =   3720
      Begin VB.ListBox lbcExclude 
         Appearance      =   0  'Flat
         Height          =   4020
         ItemData        =   "ClosestBooks.frx":01B4
         Left            =   30
         List            =   "ClosestBooks.frx":01B6
         MultiSelect     =   2  'Extended
         TabIndex        =   9
         Top             =   45
         Width           =   3585
      End
   End
   Begin VB.PictureBox plcInclude 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   360
      ScaleHeight     =   4080
      ScaleWidth      =   3660
      TabIndex        =   6
      Top             =   990
      Width           =   3720
      Begin VB.ListBox lbcInclude 
         Appearance      =   0  'Flat
         Height          =   4020
         ItemData        =   "ClosestBooks.frx":01B8
         Left            =   30
         List            =   "ClosestBooks.frx":01BA
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   30
         Width           =   3585
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4455
      TabIndex        =   2
      Top             =   5355
      Width           =   1050
   End
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
      Left            =   45
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1710
      Width           =   120
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   75
      ScaleHeight     =   270
      ScaleWidth      =   3195
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   3195
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2265
      TabIndex        =   1
      Top             =   5355
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Define which Books should be bypassed when applying the ReRate ""Closest to Air Date"" selections"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   360
      TabIndex        =   15
      Top             =   315
      Width           =   9360
   End
   Begin VB.Label lacInclude 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Books to Include"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   420
      TabIndex        =   5
      Top             =   645
      Width           =   3645
   End
   Begin VB.Label lacExclude 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Books to Exclude"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   6000
      TabIndex        =   4
      Top             =   645
      Width           =   3735
   End
End
Attribute VB_Name = "ClosestBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ClosestBooks.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract number screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim imUpdateAllowed As Integer
Dim bmChgd As Boolean
Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor

Private Type CLOSESTBOOKINFO
    iDnfCode As Integer
    sName As String * 30
    lBookDate As Long
End Type
Dim tmClosestBookInfo() As BOOKINFO


Dim rst_eff As ADODB.Recordset
Dim dnf_rst As ADODB.Recordset


Private Sub cmcExcludeToInclude_Click()
    mMoveName lbcExclude, lbcInclude
End Sub


Private Sub cmcIncludeToExclude_Click()
    mMoveName lbcInclude, lbcExclude
End Sub

Private Sub cmcCancel_Click()
    igTerminateReturn = 0   '0=Canceled
    sgPassValue = ""
    mTerminate
End Sub



Private Sub cmcDone_Click()
    Dim ilRes As Integer
    Dim blRet As Boolean
    
    If bmChgd Then
        ilRes = MsgBox("Save Changes?", vbYesNoCancel + vbQuestion, "Update")
        If ilRes = vbCancel Then
            Exit Sub
        End If
    
        If ilRes = vbYes Then
            Screen.MousePointer = vbHourglass
            blRet = mSaveRec()
            Screen.MousePointer = vbDefault
        End If
    End If
    igTerminateReturn = 1
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    Dim ilRet As Integer
    Dim slStr As String

    gCtrlGotFocus cmcDone
End Sub


Private Sub cmcSave_Click()
    Dim blRet As Boolean
    
    Screen.MousePointer = vbHourglass
    blRet = mSaveRec()
    mPopBooks
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    If (igWinStatus(COPYJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
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
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        'fmAdjFactorW = (((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        fmAdjFactorW = (((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        If fmAdjFactorW < 1# Then
            fmAdjFactorW = 1#
        Else
            'Me.Width = ((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
            Me.Width = ((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        End If
        'fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        'Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
        fmAdjFactorH = 1#
    End If
    mInit
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
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
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    bmChgd = False
    plcInclude.Width = fmAdjFactorW * plcInclude.Width
    lbcInclude.Width = plcInclude.Width - 120    'fmAdjFactorW * lbcInclude.Width
    lacInclude.Width = plcInclude.Width
    lacInclude.Left = plcInclude.Left
    cmcIncludeToExclude.Left = plcInclude.Left + plcInclude.Width + 120
    pbcIncludeToExclude.Left = cmcIncludeToExclude.Left + cmcIncludeToExclude.Width - (3 * pbcIncludeToExclude.Width) / 2
    pbcIncludeToExclude.Top = cmcIncludeToExclude.Top + 60
    cmcExcludeToInclude.Left = cmcIncludeToExclude.Left
    pbcExcludeToInclude.Left = cmcExcludeToInclude.Left + pbcExcludeToInclude.Width
    pbcExcludeToInclude.Top = cmcExcludeToInclude.Top + 60
    plcExclude.Width = plcInclude.Width    'fmAdjFactorW * plcExclude.Width
    lbcExclude.Width = lbcInclude.Width    'fmAdjFactorW * lbcExclude.Width
    plcExclude.Left = cmcIncludeToExclude.Left + cmcIncludeToExclude.Width + 120
    lacExclude.Width = plcExclude.Width
    lacExclude.Left = plcExclude.Left
    ClosestBooks.Width = plcExclude.Left + plcExclude.Width + plcInclude.Left
    cmcCancel.Left = cmcExcludeToInclude.Left
    cmcDone.Left = cmcCancel.Left - 2 * cmcDone.Width
    cmcSave.Left = cmcCancel.Left + 2 * cmcCancel.Width
    mPopBooks
    ClosestBooks.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterModalForm ClosestBooks
    Screen.MousePointer = vbDefault
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
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
    Unload ClosestBooks
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ClosestBooks = Nothing   'Remove data segment
End Sub

Private Sub lbcExclude_Click()
    mSetCommands
End Sub

Private Sub lbcInclude_Click()
    mSetCommands
End Sub


Private Sub pbcExcludeToInclude_Click()
    mMoveName lbcExclude, lbcInclude
End Sub


Private Sub pbcIncludeToExclude_Click()
    mMoveName lbcInclude, lbcExclude
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub


Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Closest Book"
End Sub

Private Sub mPopBooks()
    Dim slSQLQuery As String
    Dim llLoop As Long
    Dim blFd As Boolean
    
    lbcExclude.Clear
    lbcInclude.Clear
    ReDim tmClosestBookInfo(0 To 0) As BOOKINFO
    
    slSQLQuery = "Select effLong1 from EFF_Extra_Fields"
    slSQLQuery = slSQLQuery & " Where effType = 'E'"
    slSQLQuery = slSQLQuery & " Order by effLong1"
    Set rst_eff = gSQLSelectCall(slSQLQuery)
    Do While Not rst_eff.EOF
        tmClosestBookInfo(UBound(tmClosestBookInfo)).iDnfCode = rst_eff!effLong1
        ReDim Preserve tmClosestBookInfo(0 To UBound(tmClosestBookInfo) + 1) As BOOKINFO
        rst_eff.MoveNext
    Loop
    
    slSQLQuery = "Select dnfCode, dnfBookName, dnfBookDate from DNF_Demo_Rsrch_Names "
    slSQLQuery = slSQLQuery & " Order By dnfBookDate Desc, dnfBookName"
    Set dnf_rst = gSQLSelectCall(slSQLQuery)
    Do While Not dnf_rst.EOF
        blFd = False
        For llLoop = 0 To UBound(tmClosestBookInfo) - 1 Step 1
            If tmClosestBookInfo(llLoop).iDnfCode = dnf_rst!dnfCode Then
                tmClosestBookInfo(llLoop).sName = dnf_rst!dnfBookName
                tmClosestBookInfo(llLoop).lBookDate = gDateValue(dnf_rst!dnfBookDate)
                blFd = True
                Exit For
            End If
        Next llLoop
        If Not blFd Then
            lbcInclude.AddItem Trim$(dnf_rst!dnfBookName) & ": " & dnf_rst!dnfBookDate
            lbcInclude.ItemData(lbcInclude.NewIndex) = dnf_rst!dnfCode
        Else
            lbcExclude.AddItem Trim$(dnf_rst!dnfBookName) & ": " & dnf_rst!dnfBookDate
            lbcExclude.ItemData(lbcExclude.NewIndex) = dnf_rst!dnfCode
        End If
        dnf_rst.MoveNext
    Loop
    
    
End Sub

Private Sub mSetCommands()
    If Not imUpdateAllowed Then
        cmcIncludeToExclude.Enabled = False
        pbcIncludeToExclude.Enabled = False
        
        cmcExcludeToInclude.Enabled = False
        pbcExcludeToInclude.Enabled = False
        Exit Sub
    End If
    
    
    If lbcInclude.SelCount > 0 Then
        cmcIncludeToExclude.Enabled = True
        pbcIncludeToExclude.Enabled = True
    Else
        cmcIncludeToExclude.Enabled = False
        pbcIncludeToExclude.Enabled = False
    End If
    If lbcExclude.SelCount > 0 Then
        cmcExcludeToInclude.Enabled = True
        pbcExcludeToInclude.Enabled = True
    Else
        cmcExcludeToInclude.Enabled = False
        pbcExcludeToInclude.Enabled = False
    End If
    
End Sub

Private Sub mMoveName(lbcFrom As ListBox, lbcTo As ListBox)
    Dim slName As String
    Dim slItemData As String
    Dim llLoop As Long
    
    For llLoop = 0 To lbcFrom.ListCount - 1 Step 1
        If lbcFrom.Selected(llLoop) Then
            bmChgd = True
            slName = lbcFrom.List(llLoop)
            slItemData = lbcFrom.ItemData(llLoop)
            lbcTo.AddItem slName
            lbcTo.ItemData(lbcTo.NewIndex) = slItemData
        End If
    Next llLoop
    For llLoop = lbcFrom.ListCount - 1 To 0 Step -1
        If lbcFrom.Selected(llLoop) Then
            lbcFrom.RemoveItem lbcFrom.ListIndex
        End If
    Next llLoop


End Sub

Private Function mSaveRec() As Boolean
    Dim slSQLQuery As String
    Dim llRet As Long
    Dim ilLoop As Integer
    Dim ilRes As Integer
    Dim llEffCode As Long
    Dim llItemData As Long


    'Clear EFF
    slSQLQuery = "Delete from EFF_Extra_Fields"
    slSQLQuery = slSQLQuery & " Where effType = 'E'"
    llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    'Add to Eff
    For ilLoop = 0 To lbcExclude.ListCount - 1 Step 1
        llItemData = Val(lbcExclude.ItemData(ilLoop))
        slSQLQuery = "Insert Into EFF_Extra_Fields ( "
        slSQLQuery = slSQLQuery & "effCode, "
        slSQLQuery = slSQLQuery & "effType, "
        slSQLQuery = slSQLQuery & "effString1, "
        slSQLQuery = slSQLQuery & "effString2, "
        slSQLQuery = slSQLQuery & "effString3, "
        slSQLQuery = slSQLQuery & "effLong1, "
        slSQLQuery = slSQLQuery & "effLong2, "
        slSQLQuery = slSQLQuery & "effLong3, "
        slSQLQuery = slSQLQuery & "effUnused "
        slSQLQuery = slSQLQuery & ") "
        slSQLQuery = slSQLQuery & "Values ( "
        slSQLQuery = slSQLQuery & "Replace-X" & ", "
        slSQLQuery = slSQLQuery & "'" & gFixQuote("E") & "', "
        slSQLQuery = slSQLQuery & "'" & "" & "', "
        slSQLQuery = slSQLQuery & "'" & "" & "', "
        slSQLQuery = slSQLQuery & "'" & "" & "', "
        slSQLQuery = slSQLQuery & llItemData & ", "
        slSQLQuery = slSQLQuery & 0 & ", "
        slSQLQuery = slSQLQuery & 0 & ", "
        slSQLQuery = slSQLQuery & "'" & "" & "' "
        slSQLQuery = slSQLQuery & ") "
        llEffCode = gInsertAndReturnCode(slSQLQuery, "EFF_Extra_Fields", "EffCode", "Replace-X")
        
    Next ilLoop
    bmChgd = False
End Function
