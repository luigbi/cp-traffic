VERSION 5.00
Begin VB.Form FeedName 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4080
   ClientLeft      =   4185
   ClientTop       =   3630
   ClientWidth     =   6675
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
   ScaleHeight     =   4080
   ScaleWidth      =   6675
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   6
      Left            =   6045
      Top             =   3315
   End
   Begin VB.TextBox edcPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      IMEMode         =   3  'DISABLE
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.ListBox lbcDays 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      ItemData        =   "FeedName.frx":0000
      Left            =   750
      List            =   "FeedName.frx":0010
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2865
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   990
      TabIndex        =   17
      Top             =   2010
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.TextBox edcInterval 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   4440
      MaxLength       =   4
      TabIndex        =   16
      Top             =   1950
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox edcEndHour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   3375
      MaxLength       =   8
      TabIndex        =   15
      Top             =   1875
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox edcStartHour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   2175
      MaxLength       =   8
      TabIndex        =   14
      Top             =   1650
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox edcFTP 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   975
      MaxLength       =   70
      TabIndex        =   12
      Top             =   1425
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ListBox lbcMedia 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      ItemData        =   "FeedName.frx":0031
      Left            =   3405
      List            =   "FeedName.frx":0033
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2700
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.ListBox lbcAvailName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      ItemData        =   "FeedName.frx":0035
      Left            =   840
      List            =   "FeedName.frx":0037
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2670
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.ListBox lbcNetName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      ItemData        =   "FeedName.frx":0039
      Left            =   810
      List            =   "FeedName.frx":003B
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2460
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.ListBox lbcProducer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      ItemData        =   "FeedName.frx":003D
      Left            =   3390
      List            =   "FeedName.frx":003F
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2385
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.CommandButton cmcDropDown 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3720
      Picture         =   "FeedName.frx":0041
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   2850
      TabIndex        =   1
      Top             =   270
      Width           =   3000
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   90
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   570
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   45
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   885
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   90
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   225
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   60
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2730
      Width           =   60
   End
   Begin VB.PictureBox pbcPledgeTime 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4395
      ScaleHeight     =   210
      ScaleWidth      =   1395
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox edcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   885
      MaxLength       =   40
      TabIndex        =   5
      Top             =   945
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
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
      Left            =   3960
      TabIndex        =   24
      Top             =   3600
      Width           =   1050
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
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
      Left            =   2835
      TabIndex        =   23
      Top             =   3600
      Width           =   1050
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
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
      Left            =   1710
      TabIndex        =   22
      Top             =   3600
      Width           =   1050
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
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
      Left            =   3960
      TabIndex        =   21
      Top             =   3225
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
      Left            =   2835
      TabIndex        =   20
      Top             =   3225
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   1710
      TabIndex        =   19
      Top             =   3225
      Width           =   1050
   End
   Begin VB.PictureBox pbcFdNm 
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
      ForeColor       =   &H80000008&
      Height          =   1785
      Left            =   855
      Picture         =   "FeedName.frx":013B
      ScaleHeight     =   1785
      ScaleWidth      =   4965
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   795
      Width           =   4965
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   360
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   765
      Width           =   45
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   15
      TabIndex        =   18
      Top             =   1695
      Width           =   15
   End
   Begin VB.PictureBox plcFdNm 
      ForeColor       =   &H00000000&
      Height          =   1875
      Left            =   795
      ScaleHeight     =   1815
      ScaleWidth      =   5010
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   735
      Width           =   5070
   End
   Begin VB.Label plcScreen 
      Caption         =   "Feed Names"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   1260
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   255
      Top             =   3270
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "FeedName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of FeedName.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: FeedName.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Avail name input screen code
Option Explicit
Option Compare Text
'Avail Name Field Areas
Dim imFirstActivate As Integer
Dim tmCtrls(0 To 12)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imMaxBoxNo As Integer
Dim imBoxNo As Integer   'Current Feed Name Box
Dim hmFnf As Integer 'Feed name file handle
Dim tmFnf As FNF        'Fnf record image
Dim tmFnfSrchKey As INTKEY0    'Fnf key record image
Dim imFnfRecLen As Integer        'Fnf record length
Dim smSave(0 To 14) As String    '1=Name; 2=Pledge Time; 3=FTP Address; 4=Password;
                                 '5=Days; 6=Start Hour; 7=End Hour; 8=Interval
                                 '9=Unused; 10=Unused; 11=Network Name;
                                 '12=Producer Name; 13=Avail Name; 14=Media
Dim imSave(0 To 2) As Integer   '1=Pledge Time(0=Pre-converted Log, 1=Log Needs Conversion, 2=Insertion Order); 2=Days

Dim smOrigSave(0 To 14) As String    '1=Name; 2=Pledge Time; 3=FTP Address; 4=Password;
                                 '5=Days; 6=Start Hour; 7=End Hour; 8=Interval
                                 '9=Unused; 10=Unused; 11=Network Name;
                                 '12=Producer Name; 13=Avail Name; 14=Media

Dim tmFeedNameCode() As SORTCODE
Dim smFeedNameCodeTag As String
Dim tmProducerCode() As SORTCODE
Dim smProducerCodeTag As String
Dim tmNetRepCode() As SORTCODE
Dim smNetRepCodeTag As String
Dim tmAvailCode() As SORTCODE
Dim smAvailCodeTag As String
Dim tmMediaCode() As SORTCODE
Dim smMediaCodeTag As String


Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer
Dim imLbcArrowSetting As Integer 'True = Processing arrow key- retain current list box visibly
                                'False= Make list box invisible
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imComboBoxIndex As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imUpdateAllowed As Integer    'User can update records
Const NAMEINDEX = 1     'Name control/field
Const PLEDGETIMEINDEX = 2  'Pledge Time control/field
Const FTPADDRESSINDEX = 3  'FTP or URL Address control/field
Const PASSWORDINDEX = 4
Const CHECKDAYSINDEX = 5
Const CHECKSTARTINDEX = 6
Const CHECKENDINDEX = 7
Const CHECKINTERVALINDEX = 8
Const NETNAMEINDEX = 9
Const PRODUCERINDEX = 10
Const BOOKAVAILINDEX = 11
Const MEDIAINDEX = 12

'*******************************************************
'*                                                     *
'*      Procedure Name:mMediaPop                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the media combo       *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mMediaPop()
'
'   mMediaPop
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim slName As String
    Dim ilIndex As Integer
    If sgUseCartNo = "N" Then
        ReDim tmMediaCode(0 To 0) As SORTCODE
        Exit Sub
    End If
    ilIndex = lbcMedia.ListIndex
    If ilIndex >= 0 Then
        slName = lbcMedia.List(ilIndex)
    End If
    ilfilter(0) = NOFILTER
    slFilter(0) = ""
    ilOffSet(0) = 0
    'ilRet = gIMoveListBox(CopyInv, lbcMedia, lbcMediaCode, "Mcf.btr", gFieldOffset("Mcf", "McfName"), 6, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(CopyInv, lbcMedia, tmMediaCode(), smMediaCodeTag, "Mcf.btr", gFieldOffset("Mcf", "McfName"), 6, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mMediaPopErr
        gCPErrorMsg ilRet, "mMediaPop (gIMoveListBox)", CopyInv
        On Error GoTo 0
        lbcMedia.AddItem "[None]", 0  'Force as first item on list
        If ilIndex >= 0 Then
            gFindMatch slName, 1, lbcMedia
            If gLastFound(lbcMedia) >= 1 Then
                lbcMedia.ListIndex = gLastFound(lbcMedia)
            Else
                lbcMedia.ListIndex = -1
            End If
        Else
            lbcMedia.ListIndex = ilIndex
        End If
    End If
    Exit Sub
mMediaPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mAvailBranch                 *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to Avail  *
'*                      names and process communication*
'*                      back from avail names          *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mAvailBranch()
'
'   ilRet = mAvailBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    'Test if [New] Or new name specified
    ilRet = gOptionalLookAhead(edcDropDown, lbcAvailName, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcDropDown.Text = "[No]") Then
        imDoubleClickName = False
        mAvailBranch = False
        Exit Function
    End If
    igANmCallSource = CALLSOURCEFEED
    If edcDropDown.Text = "[New]" Then
        sgANmName = ""
    Else
        sgANmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "FeedName^Test\" & sgUserName & "\" & Trim$(str$(igANmCallSource)) & "\" & sgANmName
        Else
            slStr = "FeedName^Prod\" & sgUserName & "\" & Trim$(str$(igANmCallSource)) & "\" & sgANmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "FeedName^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igANmCallSource)) & "\" & sgANmName
    '    Else
    '        slStr = "FeedName^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igANmCallSource)) & "\" & sgANmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "AName.Exe " & slStr, 1)
    'FeedName.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    AName.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgANmName)
    igANmCallSource = Val(sgANmName)
    ilParse = gParseItem(slStr, 2, "\", sgANmName)
    'FeedName.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mAvailBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igANmCallSource = CALLDONE Then  'Done
        igANmCallSource = CALLNONE
        lbcAvailName.Clear
        smAvailCodeTag = ""
        sgAvailAnfStamp = ""
        mAvailPop
        If imTerminate Then
            mAvailBranch = False
            Exit Function
        End If
        gFindMatch sgANmName, 1, lbcAvailName
        If gLastFound(lbcAvailName) > 0 Then
            imChgMode = True
            lbcAvailName.ListIndex = gLastFound(lbcAvailName)
            edcDropDown.Text = lbcAvailName.List(lbcAvailName.ListIndex)
            imChgMode = False
            mAvailBranch = False
        Else
            imChgMode = True
            lbcAvailName.ListIndex = 1
            edcDropDown.Text = lbcAvailName.List(1)
            imChgMode = False
            edcDropDown.SetFocus
            sgANmName = ""
            Exit Function
        End If
        sgANmName = ""
    End If
    If igANmCallSource = CALLCANCELLED Then  'Cancelled
        igANmCallSource = CALLNONE
        sgANmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igANmCallSource = CALLTERMINATED Then
        igANmCallSource = CALLNONE
        sgANmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mAvailPop                    *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Avail Pop the selection Avail  *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mAvailPop()
'
'   mAvailPop
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilLp As Integer
    Dim slStr As String
    ilIndex = lbcAvailName.ListIndex
    If ilIndex > 0 Then
        slName = lbcAvailName.List(ilIndex)
    End If
    ilfilter(0) = CHARFILTER
    slFilter(0) = "F"
    ilOffSet(0) = gFieldOffset("Anf", "AnfBookLocalFeed") '2
    'ilRet = gIMoveListBox(FeedName, lbcAvailName, lbcAvailNameCode, "Anf.btr", gFieldOffset("Anf", "AnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(FeedName, lbcAvailName, tmAvailCode(), smAvailCodeTag, "Anf.btr", gFieldOffset("Anf", "AnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        'Remove "Post Log" avail name
        For ilLoop = 0 To lbcAvailName.ListCount - 1 Step 1
            slStr = Trim$(lbcAvailName.List(ilLoop))
            If StrComp(slStr, "Post Log", 1) = 0 Then
                lbcAvailName.RemoveItem ilLoop
                For ilLp = ilLoop To UBound(tmAvailCode) - 1 Step 1
                    tmAvailCode(ilLp) = tmAvailCode(ilLp + 1)
                Next ilLp
                ReDim Preserve tmAvailCode(LBound(tmAvailCode) To UBound(tmAvailCode) - 1) As SORTCODE
                Exit For
            End If
        Next ilLoop
        On Error GoTo mAvailPopErr
        gCPErrorMsg ilRet, "mAvailPop (gIMoveListBox)", FeedName
        On Error GoTo 0
        lbcAvailName.AddItem "[No]", 0  'Force as first item on list
        lbcAvailName.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 2, lbcAvailName
            If gLastFound(lbcAvailName) > 1 Then
                lbcAvailName.ListIndex = gLastFound(lbcAvailName)
            Else
                lbcAvailName.ListIndex = -1
            End If
        Else
            lbcAvailName.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mAvailPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mNetRepPop             *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Lock Box list box     *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mNetRepPop()
'
'   mContentProvPop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slProgName As String
    Dim ilProgIndex As Integer

    ilProgIndex = lbcNetName.ListIndex
    If ilProgIndex >= 1 Then
        slProgName = lbcNetName.List(ilProgIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilfilter(0) = CHARFILTER
    slFilter(0) = "N"
    ilOffSet(0) = gFieldOffset("Arf", "ArfType") '2
    'ilRet = gIMoveListBox(Agency, lbcNetName, lbcNetNameCode, "Arf.Btr", gFieldOffset("Arf", "ArfID"), 10, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(FeedName, lbcNetName, tmNetRepCode(), smNetRepCodeTag, "Arf.Btr", gFieldOffset("Arf", "ArfName"), 40, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mNetRepPopErr
        gCPErrorMsg ilRet, "mNetRepPop (gIMoveListBox)", FeedName
        On Error GoTo 0
        lbcNetName.AddItem "[None]", 0  'Force as first item on list
        lbcNetName.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilProgIndex > 1 Then
            gFindMatch slProgName, 2, lbcNetName
            If gLastFound(lbcNetName) > 1 Then
                lbcNetName.ListIndex = gLastFound(lbcNetName)
            Else
                lbcNetName.ListIndex = -1
            End If
        Else
            lbcNetName.ListIndex = ilProgIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mNetRepPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mNetRepBranch          *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to Lock   *
'*                      Box and process                *
'*                      communication back from Lock   *
'*                      Box                            *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mNetRepBranch() As Integer
'
'   ilRet = mNetRepBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
'    ilRet = gOptionalLookAhead(edcDropDown, lbcNetName, imBSMode, slStr)
    ilRet = gOptionalLookAhead(edcDropDown, lbcNetName, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mNetRepBranch = False
        Exit Function
    End If
    If igWinStatus(VEHICLESLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mNetRepBranch = True
        lbcNetName.SetFocus
        Exit Function
    End If
    sgArfCallType = "N"
    igArfCallSource = CALLSOURCEVEHOPT
    If edcDropDown.Text = "[New]" Then
        sgArfName = ""
    Else
        sgArfName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "FeedName^Test\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource)) & "\" & sgArfName
        Else
            slStr = "FeedName^Prod\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource)) & "\" & sgArfName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "FeedName^Test^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    Else
    '        slStr = "FeedName^Prod^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "NmAddr.Exe " & slStr, 1)
    'Agency.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    NmAddr.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgArfName)
    igArfCallSource = Val(sgArfName)
    ilParse = gParseItem(slStr, 2, "\", sgArfName)
    'Agency.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mNetRepBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igArfCallSource = CALLDONE Then  'Done
        igArfCallSource = CALLNONE
        lbcNetName.Clear
        smNetRepCodeTag = ""
        mNetRepPop
        If imTerminate Then
            mNetRepBranch = False
            Exit Function
        End If
        gFindMatch sgArfName, 1, lbcNetName
        sgArfName = ""
        If gLastFound(lbcNetName) > 0 Then
            imChgMode = True
            lbcNetName.ListIndex = gLastFound(lbcNetName)
            edcDropDown.Text = lbcNetName.List(lbcNetName.ListIndex)
            imChgMode = False
            mNetRepBranch = False
            mSetCommands
        Else
            imChgMode = True
            lbcNetName.Height = gListBoxHeight(lbcNetName.ListCount, 6)
            lbcNetName.ListIndex = 1
            edcDropDown.Text = lbcNetName.List(lbcNetName.ListIndex)
            imChgMode = False
            mSetCommands
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igArfCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igArfCallSource = CALLNONE
        sgArfName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igArfCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igArfCallSource = CALLNONE
        sgArfName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mProducerBranch                 *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to Lock   *
'*                      Box and process                *
'*                      communication back from Lock   *
'*                      Box                            *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mProducerBranch() As Integer
'
'   ilRet = mProducerBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcProducer, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mProducerBranch = False
        Exit Function
    End If
    If igWinStatus(VEHICLESLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mProducerBranch = True
        lbcProducer.SetFocus
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(LOCKBOXESLIST)) Then
    '    imDoubleClickName = False
    '    mProducerBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass
    sgArfCallType = "K"
    igArfCallSource = CALLSOURCEVEHOPT
    If edcDropDown.Text = "[New]" Then
        sgArfName = ""
    Else
        sgArfName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "FeedName^Test\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource)) & "\" & sgArfName
        Else
            slStr = "FeedName^Prod\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(str$(igArfCallSource)) & "\" & sgArfName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "FeedName^Test^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    Else
    '        slStr = "FeedName^Prod^NOHELP\" & sgUserName & "\" & sgArfCallType & "\" & Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "NmAddr.Exe " & slStr, 1)
    'Agency.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    NmAddr.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgArfName)
    igArfCallSource = Val(sgArfName)
    ilParse = gParseItem(slStr, 2, "\", sgArfName)
    'Agency.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mProducerBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igArfCallSource = CALLDONE Then  'Done
        igArfCallSource = CALLNONE
'        gSetMenuState True
        lbcProducer.Clear
        smProducerCodeTag = ""
        mProducerPop
        If imTerminate Then
            mProducerBranch = False
            Exit Function
        End If
'        mProducerPop
        gFindMatch sgArfName, 1, lbcProducer
        sgArfName = ""
        If gLastFound(lbcProducer) > 0 Then
            imChgMode = True
            lbcProducer.ListIndex = gLastFound(lbcProducer)
            edcDropDown.Text = lbcProducer.List(lbcProducer.ListIndex)
            imChgMode = False
            mProducerBranch = False
            mSetCommands
        Else
            imChgMode = True
            lbcProducer.Height = gListBoxHeight(lbcProducer.ListCount, 6)
            lbcProducer.ListIndex = 1
            edcDropDown.Text = lbcProducer.List(lbcProducer.ListIndex)
            imChgMode = False
            mSetCommands
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igArfCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igArfCallSource = CALLNONE
        sgArfName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igArfCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igArfCallSource = CALLNONE
        sgArfName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mProducerPop                    *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Lock Box list box     *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mProducerPop()
'
'   mLkBoxPop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcProducer.ListIndex
    If ilIndex > 1 Then
        slName = lbcProducer.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilfilter(0) = CHARFILTER
    slFilter(0) = "K"
    ilOffSet(0) = gFieldOffset("Arf", "ArfType") '2
    ilRet = gIMoveListBox(FeedName, lbcProducer, tmProducerCode(), smProducerCodeTag, "Arf.Btr", gFieldOffset("Arf", "ArfName"), 40, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mProducerPopErr
        gCPErrorMsg ilRet, "mProducerPop (gIMoveListBox)", FeedName
        On Error GoTo 0
        lbcProducer.AddItem "[None]", 0  'Force as first item on list
        lbcProducer.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcProducer
            If gLastFound(lbcProducer) > 1 Then
                lbcProducer.ListIndex = gLastFound(lbcProducer)
            Else
                lbcProducer.ListIndex = -1
            End If
        Else
            lbcProducer.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mProducerPopErr:
    On Error GoTo 0
    imTerminate = True
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
    If ilRet = 0 Then
        ilIndex = cbcSelect.ListIndex
        If Not mReadRec(ilIndex, SETFORREADONLY) Then
            GoTo cbcSelectErr
        End If
    Else
        If ilRet = 1 Then
            cbcSelect.ListIndex = 0
        End If
        ilRet = 1   'Clear fields as no match name found
    End If
    pbcFdNm.Cls
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            edcName.Text = slStr
        End If
    End If
    For ilLoop = imLBCtrls To imMaxBoxNo Step 1   'UBound(tmCtrls) Step 1
        mSetShow ilLoop  'Set show strings
    Next ilLoop
    pbcFdNm_Paint
    Screen.MousePointer = vbDefault
    imChgMode = False
    imBypassSetting = False
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    imChgMode = False
    Exit Sub
End Sub
Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcSelect_DropDown()
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
End Sub
Private Sub cbcSelect_GotFocus()
    Dim slSvText As String   'Save so list box can be reset
    If imTerminate Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        If igFdNmCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgFdNmName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgFdNmName  'New Name
            End If
            cbcSelect_Change
            If sgFdNmName <> "" Then
                mSetCommands
                gFindMatch sgFdNmName, 1, cbcSelect
                If gLastFound(cbcSelect) > 0 Then
                    cmcDone.SetFocus
                    Exit Sub
                End If
            End If
            If pbcSTab.Enabled Then
                pbcSTab.SetFocus
            Else
                cmcCancel.SetFocus
            End If
            Exit Sub
        End If
    End If
    slSvText = cbcSelect.Text
'    ilSvIndex = cbcSelect.ListIndex
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    gCtrlGotFocus cbcSelect
    If cbcSelect.ListCount <= 1 Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields
        If pbcSTab.Enabled Then
            pbcSTab.SetFocus
        Else
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If (slSvText = "") Or (slSvText = "[New]") Then
        cbcSelect.ListIndex = 0
        cbcSelect_Change    'Call change so picture area repainted
    Else
        gFindMatch slSvText, 1, cbcSelect
        If gLastFound(cbcSelect) > 0 Then
'            If (ilSvIndex <> gLastFound(cbcSelect)) Or (ilSvIndex <> cbcSelect.ListIndex) Then
            If (slSvText <> cbcSelect.List(gLastFound(cbcSelect))) Or imPopReqd Then
                cbcSelect.ListIndex = gLastFound(cbcSelect)
                cbcSelect_Change    'Call change so picture area repainted
                imPopReqd = False
            End If
        Else
            cbcSelect.ListIndex = 0
            mClearCtrlFields
            cbcSelect_Change    'Call change so picture area repainted
        End If
    End If
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
Private Sub cmcCancel_Click()
    If igFdNmCallSource <> CALLNONE Then
        igFdNmCallSource = CALLCANCELLED
        mTerminate    'Placed after setfocus no make sure which window gets focus
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If igFdNmCallSource <> CALLNONE Then
        sgFdNmName = edcName.Text 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgFdNmName = "[New]"
            If Not imTerminate Then
                mEnableBox imBoxNo
                Exit Sub
            Else
                cmcCancel_Click
                Exit Sub
            End If
        End If
    Else
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    If igFdNmCallSource <> CALLNONE Then
        If sgFdNmName = "[New]" Then
            igFdNmCallSource = CALLCANCELLED
        Else
            igFdNmCallSource = CALLDONE
        End If
        mTerminate
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    Dim ilLoop As Integer
    If imBoxNo = -1 Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If Not cmcUpdate.Enabled Then
        'Cycle to first unanswered mandatory
        For ilLoop = imLBCtrls To imMaxBoxNo Step 1   'UBound(tmCtrls) Step 1
            If mTestFields(ilLoop, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                imBoxNo = ilLoop
                mEnableBox imBoxNo
                Exit Sub
            End If
        Next ilLoop
    End If
    gCtrlGotFocus cmcDone
End Sub

Private Sub cmcDropDown_Click()
    Select Case imBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
        Case PLEDGETIMEINDEX
        Case FTPADDRESSINDEX 'Name
        Case PASSWORDINDEX 'Name
        Case CHECKDAYSINDEX
            lbcDays.Visible = Not lbcDays.Visible
        Case CHECKSTARTINDEX 'Name
        Case CHECKENDINDEX 'Name
        Case CHECKINTERVALINDEX 'Name
        Case NETNAMEINDEX
            lbcNetName.Visible = Not lbcNetName.Visible
        Case PRODUCERINDEX
            lbcProducer.Visible = Not lbcProducer.Visible
        Case BOOKAVAILINDEX
            lbcAvailName.Visible = Not lbcAvailName.Visible
        Case MEDIAINDEX
            lbcMedia.Visible = Not lbcMedia.Visible
    End Select
End Sub

Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcErase_Click()
    Dim ilRet As Integer
    Dim slStamp As String   'Date/Time stamp for file
    Dim slMsg As String
    If imSelectedIndex > 0 Then
        Screen.MousePointer = vbHourglass
        'Check that record is not referenced-Code missing
        ilRet = gIICodeRefExist(FeedName, tmFnf.iCode, "Fsf.Btr", "FsfFnfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Feed Spot name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gLICodeRefExist(FeedName, tmFnf.iCode, "Fpf.Btr", "FpfFnfCode")  'lefanfCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Feed Pledges name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & tmFnf.sName, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        slStamp = gFileDateTime(sgDBPath & "Fnf.btr")
        ilRet = btrDelete(hmFnf)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", FeedName
        On Error GoTo 0
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        'If lbcNameCode.Tag <> "" Then
        '    If slStamp = lbcNameCode.Tag Then
        '        lbcNameCode.Tag = FileDateTime(sgDBPath & "Fnf.btr")
        '    End If
        'End If
        If smFeedNameCodeTag <> "" Then
            If slStamp = smFeedNameCodeTag Then
                smFeedNameCodeTag = gFileDateTime(sgDBPath & "Fnf.btr")
            End If
        End If
        'lbcNameCode.RemoveItem imSelectedIndex - 1
        gRemoveItemFromSortCode imSelectedIndex - 1, tmFeedNameCode()
        cbcSelect.RemoveItem imSelectedIndex
    End If
    'Remove focus from control and make invisible
    mSetShow imBoxNo
    imBoxNo = -1
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcFdNm.Cls
    cbcSelect.ListIndex = 0
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
    Exit Sub
cmcEraseErr:
    On Error GoTo 0
    imTerminate = True
    Resume Next
End Sub
Private Sub cmcErase_GotFocus()
    gCtrlGotFocus cmcErase
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcReport_Click()
    Dim slStr As String
    'If Not gWinRoom(igNoExeWinRes(RPTNOSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = AVAILNAMESLIST
    ''Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "FeedName^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        Else
            slStr = "FeedName^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "FeedName^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    Else
    '        slStr = "FeedName^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptNoSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'FeedName.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    RptList.Show vbModal
    slStr = sgDoneMsg
    'FeedName.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
End Sub
Private Sub cmcReport_GotFocus()
    gCtrlGotFocus cmcReport
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcUndo_Click()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilIndex As Integer
    If imSelectedIndex > 0 Then
        ilIndex = imSelectedIndex
        If Not mReadRec(ilIndex, SETFORREADONLY) Then
            GoTo cmcUndoErr
        End If
        mMoveRecToCtrl
        For ilLoop = imLBCtrls To imMaxBoxNo Step 1   'UBound(tmCtrls) Step 1
            mSetShow ilLoop  'Set show strings
        Next ilLoop
        pbcFdNm.Cls
        pbcFdNm_Paint
        mSetCommands
        imBoxNo = -1
        pbcSTab.SetFocus
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcFdNm.Cls
    cbcSelect.ListIndex = 0
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
    Exit Sub
cmcUndoErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub cmcUndo_GotFocus()
    gCtrlGotFocus cmcUndo
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcUpdate_Click()
    Dim imSvSelectedIndex As Integer
    Dim ilCode As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLoop As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
'    slName = edcDropdown.Text   'Save name
    imSvSelectedIndex = imSelectedIndex
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    imBoxNo = -1
'    'Must reset display so altered flag is cleared and setcommand will turn select on
'    If imSvSelectedIndex <> 0 Then
'        cbcSelect.Text = slName
'    Else
'        cbcSelect.ListIndex = 0
'    End If
'    cbcSelect_Change    'Call change so picture area repainted
    ilCode = tmFnf.iCode
    cbcSelect.Clear
    smFeedNameCodeTag = ""
    mPopulate
    If imSvSelectedIndex <> 0 Then
        For ilLoop = 0 To UBound(tmFeedNameCode) - 1 Step 1 'lbcDPNameCode.ListCount - 1 Step 1
            slNameCode = tmFeedNameCode(ilLoop).sKey 'lbcDPNameCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = ilCode Then
                If cbcSelect.ListIndex = ilLoop + 1 Then
                    cbcSelect_Change
                Else
                    cbcSelect.ListIndex = ilLoop + 1
                End If
                Exit For
            End If
        Next ilLoop
    Else
        cbcSelect.ListIndex = 0
    End If
    mSetCommands
    cbcSelect.SetFocus
End Sub
Private Sub cmcUpdate_GotFocus()
    gCtrlGotFocus cmcUpdate
    mSetShow imBoxNo
    imBoxNo = -1
End Sub

Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo 'Branch on box type (control)
            Case NAMEINDEX 'Name
            Case PLEDGETIMEINDEX
            Case FTPADDRESSINDEX 'Name
            Case PASSWORDINDEX 'Name
            Case CHECKDAYSINDEX
                    gProcessArrowKey Shift, KeyCode, lbcDays, imLbcArrowSetting
            Case CHECKSTARTINDEX 'Name
            Case CHECKENDINDEX 'Name
            Case CHECKINTERVALINDEX 'Name
            Case NETNAMEINDEX
                gProcessArrowKey Shift, KeyCode, lbcNetName, imLbcArrowSetting
            Case PRODUCERINDEX
                gProcessArrowKey Shift, KeyCode, lbcProducer, imLbcArrowSetting
            Case BOOKAVAILINDEX
                gProcessArrowKey Shift, KeyCode, lbcAvailName, imLbcArrowSetting
            Case MEDIAINDEX
                gProcessArrowKey Shift, KeyCode, lbcMedia, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub

Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imBoxNo 'Branch on box type (control)
            Case NAMEINDEX 'Name
            Case PLEDGETIMEINDEX
            Case FTPADDRESSINDEX 'Name
            Case PASSWORDINDEX 'Name
            Case CHECKDAYSINDEX
            Case CHECKSTARTINDEX 'Name
            Case CHECKENDINDEX 'Name
            Case CHECKINTERVALINDEX 'Name
            Case NETNAMEINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
            Case PRODUCERINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
            Case BOOKAVAILINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
        End Select
    End If
End Sub

Private Sub edcEndHour_Change()
    If (gValidTime(edcEndHour.Text)) Or (Trim$(edcEndHour.Text) = "") Then
        mSetChg CHECKENDINDEX
    End If
End Sub

Private Sub edcEndHour_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcEndHour_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer

    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        ilFound = False
        For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
            If KeyAscii = igLegalTime(ilLoop) Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub edcFTP_Change()
    mSetChg FTPADDRESSINDEX   'can't use imBoxNo as not set when edcDropdown set via cbcSelect- altered flag set so field is saved
End Sub

Private Sub edcFTP_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcInterval_Change()
    mSetChg CHECKINTERVALINDEX 'Use NAMEINDEX instead of imBoxNo to handle calling from another function- altered flag set so field is saved
End Sub

Private Sub edcInterval_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcInterval_KeyPress(KeyAscii As Integer)
    Dim slStr As String

    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcInterval.Text
    slStr = Left$(slStr, edcInterval.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcInterval.SelStart - edcInterval.SelLength)
    If gCompNumberStr(slStr, "1440") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
        Case PLEDGETIMEINDEX
        Case FTPADDRESSINDEX 'Name
        Case PASSWORDINDEX 'Name
        Case CHECKDAYSINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcDays, imBSMode, imComboBoxIndex
        Case CHECKSTARTINDEX 'Name
        Case CHECKENDINDEX 'Name
        Case CHECKINTERVALINDEX 'Name
        Case NETNAMEINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcNetName, imBSMode, slStr)
            If ilRet = 1 Then
                lbcNetName.ListIndex = 1
            End If
        Case PRODUCERINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcProducer, imBSMode, slStr)
            If ilRet = 1 Then
                lbcProducer.ListIndex = 1
            End If
        Case BOOKAVAILINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcAvailName, imBSMode, slStr)
            If ilRet = 1 Then
                lbcAvailName.ListIndex = 1
            End If
        Case MEDIAINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcMedia, imBSMode, slStr)
            If ilRet = 1 Then
                lbcMedia.ListIndex = 1
            End If
    End Select
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
        Case PLEDGETIMEINDEX
        Case FTPADDRESSINDEX 'Name
        Case PASSWORDINDEX 'Name
        Case CHECKDAYSINDEX
        Case CHECKSTARTINDEX 'Name
        Case CHECKENDINDEX 'Name
        Case CHECKINTERVALINDEX 'Name
        Case NETNAMEINDEX
        Case PRODUCERINDEX
        Case BOOKAVAILINDEX
        Case MEDIAINDEX
    End Select
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDropDown_LostFocus()
    Select Case imBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
        Case PLEDGETIMEINDEX
        Case FTPADDRESSINDEX 'Name
        Case PASSWORDINDEX 'Name
        Case CHECKDAYSINDEX
        Case CHECKSTARTINDEX 'Name
        Case CHECKENDINDEX 'Name
        Case CHECKINTERVALINDEX 'Name
        Case NETNAMEINDEX
        Case PRODUCERINDEX
        Case BOOKAVAILINDEX
        Case MEDIAINDEX
    End Select
End Sub

Private Sub edcName_Change()
    mSetChg NAMEINDEX 'Use NAMEINDEX instead of imBoxNo to handle calling from another function- altered flag set so field is saved
End Sub

Private Sub edcName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPassword_Change()
    mSetChg PASSWORDINDEX 'Use NAMEINDEX instead of imBoxNo to handle calling from another function- altered flag set so field is saved
End Sub

Private Sub edcPassword_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcStartHour_Change()
    If (gValidTime(edcStartHour.Text)) Or (Trim$(edcStartHour.Text) = "") Then
        mSetChg CHECKSTARTINDEX
    End If
End Sub

Private Sub edcStartHour_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcStartHour_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer

    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        ilFound = False
        For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
            If KeyAscii = igLegalTime(ilLoop) Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If

End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
'    Dim ilLoop As Integer
    If (igWinStatus(FEEDNAMELIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcFdNm.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
        'Show branner
'        edcLinkDestHelpMsg.LinkExecute "BF"
    Else
        pbcFdNm.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
        'Show branner
'        edcLinkDestHelpMsg.LinkExecute "BT"
    End If
    gShowBranner imUpdateAllowed
    mSetCommands
'    DoEvents
    'This loop is required to prevent a timing problem- if calling
    'with sg----- = "", then loss GotFocus to first control
    'without this loop
'    For ilLoop = 1 To igDDEDelay Step 1
'        DoEvents
'    Next ilLoop
'    gShowBranner
    Me.KeyPreview = True
    FeedName.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        If (cbcSelect.Enabled) And (imBoxNo > 0) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        If ilReSet Then
            cbcSelect.Enabled = True
        End If
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
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    If Not igManUnload Then
        mSetShow imBoxNo
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            mEnableBox imBoxNo
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If

    smFeedNameCodeTag = ""
    Erase tmFeedNameCode
    smProducerCodeTag = ""
    Erase tmProducerCode
    smNetRepCodeTag = ""
    Erase tmNetRepCode
    smAvailCodeTag = ""
    Erase tmAvailCode
    smMediaCodeTag = ""
    Erase tmMediaCode

    btrExtClear hmFnf   'Clear any previous extend operation
    ilRet = btrClose(hmFnf)
    btrDestroy hmFnf
    
    Set FeedName = Nothing   'Remove data segment

End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear each control on the      *
'*                      screen                         *
'*                                                     *
'*******************************************************
Private Sub mClearCtrlFields()
'
'   mClearCtrlFields
'   Where:
'
    Dim ilLoop As Integer
    For ilLoop = LBound(smSave) To UBound(smSave) Step 1
        smSave(ilLoop) = ""
    Next ilLoop
    For ilLoop = LBound(imSave) To UBound(imSave) Step 1
        imSave(ilLoop) = -1
    Next ilLoop
    For ilLoop = LBound(smOrigSave) To UBound(smOrigSave) Step 1
        smOrigSave(ilLoop) = ""
    Next ilLoop
    edcDropDown.Text = ""
    edcName.Text = ""
    edcPassword.Text = ""
    edcFTP.Text = ""
    edcStartHour.Text = ""
    edcEndHour.Text = ""
    edcInterval.Text = ""
    lbcDays.ListIndex = -1
    lbcNetName.ListIndex = -1
    lbcProducer.ListIndex = -1
    lbcAvailName.ListIndex = -1
    lbcMedia.ListIndex = -1
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
        tmCtrls(ilLoop).sShow = ""
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxBoxNo) Then   'UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            'mSendHelpMsg "Enter avail name"
            edcName.Width = tmCtrls(ilBoxNo).fBoxW
            edcName.MaxLength = 40
            gMoveFormCtrl pbcFdNm, edcName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcName.Visible = True  'Set visibility
            edcName.SetFocus
        Case PLEDGETIMEINDEX
            If imSave(1) < 0 Then
                imSave(1) = 0    'Active
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcPledgeTime.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcFdNm, pbcPledgeTime, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcPledgeTime_Paint
            pbcPledgeTime.Visible = True
            pbcPledgeTime.SetFocus
        Case FTPADDRESSINDEX 'Name
            'mSendHelpMsg "Enter avail name"
            edcFTP.Width = tmCtrls(ilBoxNo).fBoxW
            edcFTP.MaxLength = 70
            gMoveFormCtrl pbcFdNm, edcFTP, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcFTP.Visible = True  'Set visibility
            edcFTP.SetFocus
        Case PASSWORDINDEX 'Name
            'mSendHelpMsg "Enter avail name"
            edcPassword.Width = tmCtrls(ilBoxNo).fBoxW
            edcPassword.MaxLength = 10
            gMoveFormCtrl pbcFdNm, edcPassword, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcPassword.Visible = True  'Set visibility
            edcPassword.SetFocus
        Case CHECKDAYSINDEX
            lbcDays.Height = gListBoxHeight(lbcDays.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 7  'tgSpf.iAProd
            gMoveFormCtrl pbcFdNm, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcDays.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            gFindMatch smSave(5), 0, lbcDays
            If gLastFound(lbcDays) >= 0 Then
                imChgMode = True
                lbcDays.ListIndex = gLastFound(lbcDays)
                edcDropDown.Text = lbcDays.List(lbcDays.ListIndex)
                imChgMode = False
            Else
                If smSave(5) <> "" Then
                    imChgMode = True
                    lbcDays.ListIndex = -1
                    edcDropDown.Text = smSave(5)
                    imChgMode = False
                Else
                    imChgMode = True
                    lbcDays.ListIndex = 0   '[None]
                    edcDropDown.Text = lbcDays.List(lbcDays.ListIndex)
                    imChgMode = False
                End If
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case CHECKSTARTINDEX 'Name
            'mSendHelpMsg "Enter avail name"
            edcStartHour.Width = tmCtrls(ilBoxNo).fBoxW
            edcStartHour.MaxLength = 8
            gMoveFormCtrl pbcFdNm, edcStartHour, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcStartHour.Visible = True  'Set visibility
            edcStartHour.SetFocus
        Case CHECKENDINDEX 'Name
            'mSendHelpMsg "Enter avail name"
            edcEndHour.Width = tmCtrls(ilBoxNo).fBoxW
            edcEndHour.MaxLength = 8
            gMoveFormCtrl pbcFdNm, edcEndHour, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcEndHour.Visible = True  'Set visibility
            edcEndHour.SetFocus
        Case CHECKINTERVALINDEX 'Name
            'mSendHelpMsg "Enter avail name"
            edcInterval.Width = tmCtrls(ilBoxNo).fBoxW
            edcInterval.MaxLength = 4
            gMoveFormCtrl pbcFdNm, edcInterval, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcInterval.Visible = True  'Set visibility
            edcInterval.SetFocus
        Case NETNAMEINDEX
            mNetRepPop
            If imTerminate Then
                Exit Sub
            End If
            lbcNetName.Height = gListBoxHeight(lbcNetName.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 40  'tgSpf.iAProd
            gMoveFormCtrl pbcFdNm, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcNetName.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            gFindMatch smSave(11), 1, lbcNetName
            If gLastFound(lbcNetName) >= 1 Then
                imChgMode = True
                lbcNetName.ListIndex = gLastFound(lbcNetName)
                edcDropDown.Text = lbcNetName.List(lbcNetName.ListIndex)
                imChgMode = False
            Else
                If smSave(11) <> "" Then
                    imChgMode = True
                    lbcNetName.ListIndex = -1
                    edcDropDown.Text = smSave(11)
                    imChgMode = False
                Else
                    imChgMode = True
                    lbcNetName.ListIndex = 1   '[None]
                    edcDropDown.Text = lbcNetName.List(lbcNetName.ListIndex)
                    imChgMode = False
                End If
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PRODUCERINDEX
            mProducerPop
            If imTerminate Then
                Exit Sub
            End If
            lbcProducer.Height = gListBoxHeight(lbcProducer.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 40  'tgSpf.iAProd
            gMoveFormCtrl pbcFdNm, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcProducer.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            gFindMatch smSave(12), 1, lbcProducer
            If gLastFound(lbcProducer) >= 1 Then
                imChgMode = True
                lbcProducer.ListIndex = gLastFound(lbcProducer)
                edcDropDown.Text = lbcProducer.List(lbcProducer.ListIndex)
                imChgMode = False
            Else
                If smSave(12) <> "" Then
                    imChgMode = True
                    lbcProducer.ListIndex = -1
                    edcDropDown.Text = smSave(12)
                    imChgMode = False
                Else
                    imChgMode = True
                    lbcProducer.ListIndex = 1   '[None]
                    edcDropDown.Text = lbcProducer.List(lbcProducer.ListIndex)
                    imChgMode = False
                End If
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case BOOKAVAILINDEX
            mAvailPop
            If imTerminate Then
                Exit Sub
            End If
            lbcAvailName.Height = gListBoxHeight(lbcAvailName.ListCount, 5)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20  'tgSpf.iAProd
            gMoveFormCtrl pbcFdNm, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcAvailName.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            gFindMatch smSave(13), 1, lbcAvailName
            If gLastFound(lbcAvailName) >= 1 Then
                imChgMode = True
                lbcAvailName.ListIndex = gLastFound(lbcAvailName)
                edcDropDown.Text = lbcAvailName.List(lbcAvailName.ListIndex)
                imChgMode = False
            Else
                If smSave(13) <> "" Then
                    imChgMode = True
                    lbcAvailName.ListIndex = -1
                    edcDropDown.Text = smSave(13)
                    imChgMode = False
                Else
                    imChgMode = True
                    lbcAvailName.ListIndex = 1   '[None]
                    edcDropDown.Text = lbcAvailName.List(lbcAvailName.ListIndex)
                    imChgMode = False
                End If
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case MEDIAINDEX
            mMediaPop
            If imTerminate Then
                Exit Sub
            End If
            lbcMedia.Height = gListBoxHeight(lbcMedia.ListCount, 5)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 6  'tgSpf.iAProd
            gMoveFormCtrl pbcFdNm, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcMedia.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            gFindMatch smSave(14), 1, lbcMedia
            If gLastFound(lbcMedia) >= 1 Then
                imChgMode = True
                lbcMedia.ListIndex = gLastFound(lbcMedia)
                edcDropDown.Text = lbcMedia.List(lbcMedia.ListIndex)
                imChgMode = False
            Else
                If smSave(14) <> "" Then
                    imChgMode = True
                    lbcMedia.ListIndex = -1
                    edcDropDown.Text = smSave(14)
                    imChgMode = False
                Else
                    imChgMode = True
                    lbcMedia.ListIndex = 1   '[None]
                    edcDropDown.Text = lbcMedia.List(lbcMedia.ListIndex)
                    imChgMode = False
                End If
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus

    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
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
    imTerminate = False
    imFirstActivate = True
    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    mInitBox
    FeedName.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    gCenterStdAlone FeedName
    'FeedName.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imPopReqd = False
    imFirstFocus = True
    imSelectedIndex = -1
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imTabDirection = 0  'Left to right movement
    imFnfRecLen = Len(tmFnf)  'Get and save ARF record length
    imBoxNo = -1 'Initialize current Box to N/A
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    hmFnf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmFnf, "", sgDBPath & "Fnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", FeedName
    On Error GoTo 0
'    gCenterModalForm FeedName
'    Traffic!plcHelp.Caption = ""
    lbcAvailName.Clear 'Force list box to be populated
    mAvailPop
    If imTerminate Then
        Exit Sub
    End If
    lbcMedia.Clear 'Force list box to be populated
    mMediaPop
    If imTerminate Then
        Exit Sub
    End If
    lbcNetName.Clear 'Force list box to be populated
    mNetRepPop
    If imTerminate Then
        Exit Sub
    End If
    lbcProducer.Clear 'Force list box to be populated
    mProducerPop
    If imTerminate Then
        Exit Sub
    End If
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If Not imTerminate Then
        cbcSelect.ListIndex = 0
    End If
    'cbcSelect.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    flTextHeight = pbcFdNm.TextHeight("1") - 35
    imMaxBoxNo = 12
    'Position panel and picture areas with panel
    plcFdNm.Move 795, 735, pbcFdNm.Width + fgPanelAdj, pbcFdNm.Height + fgPanelAdj
    pbcFdNm.Move plcFdNm.Left + fgBevelX, plcFdNm.Top + fgBevelY
    'Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2955, fgBoxStH
    'Pledge Times
    gSetCtrl tmCtrls(PLEDGETIMEINDEX), 3000, tmCtrls(NAMEINDEX).fBoxY, 1935, fgBoxStH
    'FTP
    gSetCtrl tmCtrls(FTPADDRESSINDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 3870, fgBoxStH
    'Password
    gSetCtrl tmCtrls(PASSWORDINDEX), 3915, tmCtrls(FTPADDRESSINDEX).fBoxY, 1020, fgBoxStH
    'Days to Check
    gSetCtrl tmCtrls(CHECKDAYSINDEX), 30, tmCtrls(FTPADDRESSINDEX).fBoxY + fgStDeltaY, 1260, fgBoxStH
    'Start Hour
    gSetCtrl tmCtrls(CHECKSTARTINDEX), 1305, tmCtrls(CHECKDAYSINDEX).fBoxY, 1170, fgBoxStH
    'End Hour
    gSetCtrl tmCtrls(CHECKENDINDEX), 2490, tmCtrls(CHECKDAYSINDEX).fBoxY, 915, fgBoxStH
    'Interval
    gSetCtrl tmCtrls(CHECKINTERVALINDEX), 3420, tmCtrls(CHECKDAYSINDEX).fBoxY, 1515, fgBoxStH
    'Net/Rep name
    gSetCtrl tmCtrls(NETNAMEINDEX), 30, tmCtrls(CHECKDAYSINDEX).fBoxY + fgStDeltaY, 2445, fgBoxStH
    'Producer name
    gSetCtrl tmCtrls(PRODUCERINDEX), 2490, tmCtrls(NETNAMEINDEX).fBoxY, 2445, fgBoxStH
    'Avail name
    gSetCtrl tmCtrls(BOOKAVAILINDEX), 30, tmCtrls(NETNAMEINDEX).fBoxY + fgStDeltaY, 2445, fgBoxStH
    'Media
    gSetCtrl tmCtrls(MEDIAINDEX), 2490, tmCtrls(BOOKAVAILINDEX).fBoxY, 2445, fgBoxStH
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                      and set defaults               *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec(ilTestChg As Integer)
'
'   mMoveCtrlToRec iTest
'   Where:
'       iTest (I)- True = only move if field changed
'                  False = move regardless of change state
'
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStr As String

    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
        tmFnf.sName = edcName.Text
    End If
    If Not ilTestChg Or tmCtrls(PLEDGETIMEINDEX).iChg Then
        If imSave(1) = 1 Then
            tmFnf.sPledgeTime = "L"
        ElseIf imSave(1) = 2 Then
            tmFnf.sPledgeTime = "I"
        Else
            tmFnf.sPledgeTime = "P"
        End If
    End If
    If Not ilTestChg Or tmCtrls(FTPADDRESSINDEX).iChg Then
        tmFnf.sFTP = edcFTP.Text
    End If
    If Not ilTestChg Or tmCtrls(PASSWORDINDEX).iChg Then
        tmFnf.sPW = edcPassword.Text
    End If
    If Not ilTestChg Or tmCtrls(CHECKDAYSINDEX).iChg Then
        gFindMatch smSave(5), 0, lbcDays    'Determine if name exist
        If gLastFound(lbcDays) <> -1 Then   'Name found
            If gLastFound(lbcDays) = 1 Then
                slStr = "YYYYYNN"
            ElseIf gLastFound(lbcDays) = 2 Then
                slStr = "YYYYYYN"
            ElseIf gLastFound(lbcDays) = 3 Then
                slStr = "YYYYYYY"
            Else
                slStr = ""
            End If
            tmFnf.sChkDays = slStr
        Else
            tmFnf.sChkDays = ""
        End If
    End If
    If Not ilTestChg Or tmCtrls(CHECKSTARTINDEX).iChg Then
        slStr = edcStartHour.Text
        gPackTime slStr, tmFnf.iChkStartHr(0), tmFnf.iChkStartHr(1)
    End If
    If Not ilTestChg Or tmCtrls(CHECKENDINDEX).iChg Then
        slStr = edcEndHour.Text
        gPackTime slStr, tmFnf.iChkEndHr(0), tmFnf.iChkEndHr(1)
    End If
    If Not ilTestChg Or tmCtrls(CHECKINTERVALINDEX).iChg Then
        tmFnf.lChkInterval = edcInterval.Text
    End If
    If Not ilTestChg Or tmCtrls(NETNAMEINDEX).iChg Then
        gFindMatch smSave(11), 2, lbcNetName    'Determine if name exist
        If gLastFound(lbcNetName) >= 2 Then   'Name found
            slNameCode = tmNetRepCode(gLastFound(lbcNetName) - 2).sKey   'lbcAnnCode.List(imSave(9) - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", FeedName
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmFnf.iNetArfCode = CInt(slCode)
        Else
            tmFnf.iNetArfCode = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(PRODUCERINDEX).iChg Then
        gFindMatch smSave(12), 2, lbcProducer    'Determine if name exist
        If gLastFound(lbcProducer) >= 2 Then   'Name found
            slNameCode = tmProducerCode(gLastFound(lbcProducer) - 2).sKey   'lbcAnnCode.List(imSave(9) - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", FeedName
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmFnf.iProdArfCode = CInt(slCode)
        Else
            tmFnf.iProdArfCode = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(BOOKAVAILINDEX).iChg Then
        gFindMatch smSave(13), 2, lbcAvailName    'Determine if name exist
        If gLastFound(lbcAvailName) >= 2 Then   'Name found
            slNameCode = tmAvailCode(gLastFound(lbcAvailName) - 2).sKey   'lbcAnnCode.List(imSave(9) - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", FeedName
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmFnf.ianfCode = CInt(slCode)
        Else
            tmFnf.ianfCode = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(MEDIAINDEX).iChg Then
        gFindMatch smSave(14), 1, lbcMedia    'Determine if name exist
        If gLastFound(lbcMedia) >= 1 Then   'Name found
            slNameCode = tmMediaCode(gLastFound(lbcMedia) - 1).sKey 'lbcAnnCode.List(imSave(9) - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", FeedName
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmFnf.iMcfCode = CInt(slCode)
        Else
            tmFnf.iMcfCode = 0
        End If
    End If
    Exit Sub
mMoveCtrlToRecErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()
'
'   mMoveRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilFound As Integer

    edcName.Text = Trim$(tmFnf.sName)
    smOrigSave(1) = edcName.Text

    Select Case tmFnf.sPledgeTime
        Case "P"
            imSave(1) = 0
        Case "L"
            imSave(1) = 1
        Case "I"
            imSave(1) = 2
        Case Else
            imSave(1) = -1
    End Select
    smOrigSave(2) = tmFnf.sPledgeTime

    edcFTP.Text = Trim$(tmFnf.sFTP)
    smOrigSave(3) = edcFTP.Text

    edcPassword.Text = Trim$(tmFnf.sPW)
    smOrigSave(4) = edcPassword.Text

    If StrComp(tmFnf.sChkDays, "YYYYYNN", vbTextCompare) = 0 Then
        lbcDays.ListIndex = 1
    ElseIf StrComp(tmFnf.sChkDays, "YYYYYYN", vbTextCompare) = 0 Then
        lbcDays.ListIndex = 2
    ElseIf StrComp(tmFnf.sChkDays, "YYYYYYY", vbTextCompare) = 0 Then
        lbcDays.ListIndex = 3
    Else
        lbcDays.ListIndex = 0
    End If
    imSave(2) = lbcDays.ListIndex
    smOrigSave(5) = Trim$(tmFnf.sChkDays)

    gUnpackTime tmFnf.iChkStartHr(0), tmFnf.iChkStartHr(1), "A", "2", smSave(6)
    edcStartHour.Text = smSave(6)
    gUnpackTime tmFnf.iChkEndHr(0), tmFnf.iChkEndHr(1), "A", "2", smSave(7)
    edcEndHour.Text = smSave(7)

    edcInterval.Text = str(tmFnf.lChkInterval)
    smOrigSave(8) = edcInterval.Text



    ilFound = False
    For ilLoop = 0 To UBound(tmNetRepCode) - 1 Step 1 'lbcFeedCode.ListCount - 1 Step 1
        slNameCode = tmNetRepCode(ilLoop).sKey   'lbcFeedCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If ilRet = CP_MSG_NONE Then
            If Val(slCode) = tmFnf.iNetArfCode Then
                ilFound = True
                Exit For
            End If
        Else
            ilFound = False
        End If
    Next ilLoop
    If ilFound Then
        lbcNetName.ListIndex = ilLoop + 2
    Else
        lbcNetName.ListIndex = 1   '[None]
    End If
    If lbcNetName.ListIndex <= 0 Then
        smOrigSave(11) = ""
    Else
        smOrigSave(11) = lbcNetName.List(lbcNetName.ListIndex)
    End If

    ilFound = False
    For ilLoop = 0 To UBound(tmProducerCode) - 1 Step 1 'lbcFeedCode.ListCount - 1 Step 1
        slNameCode = tmProducerCode(ilLoop).sKey   'lbcFeedCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If ilRet = CP_MSG_NONE Then
            If Val(slCode) = tmFnf.iProdArfCode Then
                ilFound = True
                Exit For
            End If
        Else
            ilFound = False
        End If
    Next ilLoop
    If ilFound Then
        lbcProducer.ListIndex = ilLoop + 2
    Else
        lbcProducer.ListIndex = 1   '[None]
    End If
    If lbcProducer.ListIndex <= 0 Then
        smOrigSave(12) = ""
    Else
        smOrigSave(12) = lbcProducer.List(lbcProducer.ListIndex)
    End If

    ilFound = False
    For ilLoop = 0 To UBound(tmAvailCode) - 1 Step 1 'lbcFeedCode.ListCount - 1 Step 1
        slNameCode = tmAvailCode(ilLoop).sKey   'lbcFeedCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If ilRet = CP_MSG_NONE Then
            If Val(slCode) = tmFnf.ianfCode Then
                ilFound = True
                Exit For
            End If
        Else
            ilFound = False
        End If
    Next ilLoop
    If ilFound Then
        lbcAvailName.ListIndex = ilLoop + 2
    Else
        lbcAvailName.ListIndex = 1   '[None]
    End If
    If lbcAvailName.ListIndex <= 0 Then
        smOrigSave(13) = ""
    Else
        smOrigSave(13) = lbcAvailName.List(lbcAvailName.ListIndex)
    End If


    ilFound = False
    For ilLoop = 0 To UBound(tmMediaCode) - 1 Step 1 'lbcFeedCode.ListCount - 1 Step 1
        slNameCode = tmMediaCode(ilLoop).sKey   'lbcFeedCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If ilRet = CP_MSG_NONE Then
            If Val(slCode) = tmFnf.iMcfCode Then
                ilFound = True
                Exit For
            End If
        Else
            ilFound = False
        End If
    Next ilLoop
    If ilFound Then
        lbcMedia.ListIndex = ilLoop + 1
    Else
        lbcMedia.ListIndex = 0   '[None]
    End If
    If lbcMedia.ListIndex < 0 Then
        smOrigSave(14) = ""
    Else
        smOrigSave(14) = lbcMedia.List(lbcMedia.ListIndex)
    End If


    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mOKName                         *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test that name is unique        *
'*                                                     *
'*******************************************************
Private Function mOKName()
    Dim slStr As String
    If edcName.Text <> "" Then    'Test name
        slStr = Trim$(edcName.Text)
        gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                If Trim$(edcName.Text) = cbcSelect.List(gLastFound(cbcSelect)) Then
                    Beep
                    MsgBox "Feed Name already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcName.Text = Trim$(tmFnf.sName) 'Reset text
                    mSetShow imBoxNo
                    mSetChg imBoxNo
                    imBoxNo = 1
                    mEnableBox imBoxNo
                    mOKName = False
                    Exit Function
                End If
            End If
        End If
    End If
    mOKName = True
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
    slCommand = sgCommandStr    'Command$
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False
    '    imShowHelpMsg = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
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
    '    imShowHelpMsg = True
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpMsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone FeedName, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igFdNmCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igANmCallSource = CALLNONE
    'End If
    If igFdNmCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgFdNmName = slStr
        Else
            sgFdNmName = ""
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer

    imPopReqd = False
    ilfilter(0) = NOFILTER
    slFilter(0) = ""
    ilOffSet(0) = 0
    'ilRet = gIMoveListBox(FeedName, cbcSelect, lbcNameCode, "Fnf.btr", gFieldOffset("Fnf", "FnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(FeedName, cbcSelect, tmFeedNameCode(), smFeedNameCodeTag, "Fnf.btr", gFieldOffset("Fnf", "FnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", FeedName
        On Error GoTo 0
        cbcSelect.AddItem "[New]", 0  'Force as first item on list
        imPopReqd = True
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec(ilSelectIndex As Integer, ilForUpdate As Integer) As Integer
'
'   iRet = ENmRead(ilSelectIndex)
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status

    slNameCode = tmFeedNameCode(ilSelectIndex - 1).sKey    'lbcNameCode.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", FeedName
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmFnfSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmFnf, tmFnf, imFnfRecLen, tmFnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", FeedName
    On Error GoTo 0
    mReadRec = True
    Exit Function
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update or added record          *
'*                                                     *
'*******************************************************
Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slStamp As String   'Date/Time stamp for file
    mSetShow imBoxNo
    If mTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        mSaveRec = False
        Exit Function
    End If
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "Fnf.btr")
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
        mMoveCtrlToRec True
        If imSelectedIndex = 0 Then 'New selected
            tmFnf.iCode = 0  'Autoincrement
            ilRet = btrInsert(hmFnf, tmFnf, imFnfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert)"
        Else 'Old record-Update
            ilRet = btrUpdate(hmFnf, tmFnf, imFnfRecLen)
            slMsg = "mSaveRec (btr(Update)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, FeedName
    On Error GoTo 0
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
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRecChg                      *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if record altered and*
'*                      requires updating              *
'*                                                     *
'*******************************************************
Private Function mSaveRecChg(ilAsk As Integer) As Integer
'
'   iAsk = True
'   iRet = mSaveRecChg(iAsk)
'   Where:
'       iAsk (I)- True = Ask if changed records should be updated;
'                 False= Update record if required without asking user
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRes As Integer
    Dim slMess As String
    Dim ilAltered As Integer
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    If mTestFields(TESTALLCTRLS, ALLMANBLANK + NOMSG) = NO Then
        If ilAltered = YES Then
            If ilAsk Then
                If imSelectedIndex > 0 Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    slMess = "Add " & edcName.Text
                End If
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    pbcFdNm_Paint
                    Exit Function
                End If
                If ilRes = vbYes Then
                    ilRes = mSaveRec()
                    mSaveRecChg = ilRes
                    Exit Function
                End If
                If ilRes = vbNo Then
                    cbcSelect.ListIndex = 0
                End If
            Else
                ilRes = mSaveRec()
                mSaveRecChg = ilRes
                Exit Function
            End If
        End If
    End If
    mSaveRecChg = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetChg                         *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if value for a       *
'*                      control is different from the  *
'*                      record                         *
'*                                                     *
'*******************************************************
Private Sub mSetChg(ilBoxNo As Integer)
'
'   mSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'
    Dim slStr As String
    If ilBoxNo < imLBCtrls Or ilBoxNo > imMaxBoxNo Then   'UBound(tmCtrls) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            gSetChgFlag smOrigSave(1), edcName, tmCtrls(ilBoxNo)
        Case PLEDGETIMEINDEX
        Case FTPADDRESSINDEX 'Name
            gSetChgFlag smOrigSave(3), edcFTP, tmCtrls(ilBoxNo)
        Case PASSWORDINDEX 'Name
            gSetChgFlag smOrigSave(4), edcPassword, tmCtrls(ilBoxNo)
        Case CHECKDAYSINDEX
            If lbcDays.ListIndex = 1 Then
                slStr = "YYYYYNN"
            ElseIf lbcDays.ListIndex = 2 Then
                slStr = "YYYYYYN"
            ElseIf lbcDays.ListIndex = 3 Then
                slStr = "YYYYYYY"
            Else
                slStr = ""
            End If
            gSetChgFlagStr smOrigSave(5), slStr, tmCtrls(ilBoxNo)
        Case CHECKSTARTINDEX 'Name
            gSetChgFlag smOrigSave(6), edcStartHour, tmCtrls(ilBoxNo)
        Case CHECKENDINDEX 'Name
            gSetChgFlag smOrigSave(7), edcEndHour, tmCtrls(ilBoxNo)
        Case CHECKINTERVALINDEX 'Name
            gSetChgFlag smOrigSave(8), edcInterval, tmCtrls(ilBoxNo)
        Case NETNAMEINDEX
            gSetChgFlag smOrigSave(11), edcDropDown, tmCtrls(ilBoxNo)
        Case PRODUCERINDEX
            gSetChgFlag smOrigSave(12), edcDropDown, tmCtrls(ilBoxNo)
        Case BOOKAVAILINDEX
            gSetChgFlag smOrigSave(13), edcDropDown, tmCtrls(ilBoxNo)
        Case MEDIAINDEX
            gSetChgFlag smOrigSave(14), edcDropDown, tmCtrls(ilBoxNo)
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set command buttons (enable or *
'*                      disabled)                      *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
'
'   mSetCommands
'   Where:
'
    Dim ilAltered As Integer
    If imBypassSetting Then
        Exit Sub
    End If
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (ilAltered = YES) And (imUpdateAllowed) Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    'Revert button set if any field changed
    If ilAltered Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
    'Erase button set if any field contains a value or change mode
    If (imSelectedIndex > 0) Or (mTestFields(TESTALLCTRLS, ALLBLANK + NOMSG) = NO) And (imUpdateAllowed) Then
        cmcErase.Enabled = True
    Else
        cmcErase.Enabled = False
    End If
    If Not ilAltered Then
        cbcSelect.Enabled = True
    Else
        cbcSelect.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocusx                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       imBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxBoxNo) Then   'UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.Visible = True
            edcName.SetFocus
        Case PLEDGETIMEINDEX
            pbcPledgeTime.Visible = True
            pbcPledgeTime.SetFocus
        Case FTPADDRESSINDEX 'Name
            edcFTP.Visible = True
            edcFTP.SetFocus
        Case PASSWORDINDEX 'Name
            edcPassword.Visible = True
            edcPassword.SetFocus
        Case CHECKDAYSINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case CHECKSTARTINDEX 'Name
            edcStartHour.Visible = True  'Set visibility
            edcStartHour.SetFocus
        Case CHECKENDINDEX 'Name
            edcEndHour.Visible = True  'Set visibility
            edcEndHour.SetFocus
        Case CHECKINTERVALINDEX 'Name
            edcInterval.Visible = True  'Set visibility
            edcInterval.SetFocus
        Case NETNAMEINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PRODUCERINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case BOOKAVAILINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case MEDIAINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus

    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxBoxNo) Then   'UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.Visible = False  'Set visibility
            slStr = edcName.Text
            smSave(1) = slStr
            gSetShow pbcFdNm, slStr, tmCtrls(ilBoxNo)
        Case PLEDGETIMEINDEX
            pbcPledgeTime.Visible = False  'Set visibility
            If imSave(1) = 0 Then
                slStr = "Pre-Converted Log"
                smSave(1) = "P"
            ElseIf imSave(1) = 1 Then
                slStr = "Log Needs Conversion"
                smSave(1) = "L"
            ElseIf imSave(1) = 2 Then
                slStr = "Insertion Order"
                smSave(1) = "I"
            Else
                slStr = ""
                smSave(1) = ""
            End If
            gSetShow pbcFdNm, slStr, tmCtrls(ilBoxNo)
        Case FTPADDRESSINDEX 'Name
            edcFTP.Visible = False  'Set visibility
            slStr = edcFTP.Text
            smSave(3) = slStr
            gSetShow pbcFdNm, slStr, tmCtrls(ilBoxNo)
        Case PASSWORDINDEX 'Name
            edcPassword.Visible = False  'Set visibility
            slStr = edcPassword.Text
            smSave(4) = slStr
            gSetShow pbcFdNm, slStr, tmCtrls(ilBoxNo)
        Case CHECKDAYSINDEX 'announcer
            lbcDays.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            imSave(2) = lbcDays.ListIndex
            If lbcDays.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcDays.List(lbcDays.ListIndex)
            End If
            smSave(5) = slStr
            gSetShow pbcFdNm, slStr, tmCtrls(ilBoxNo)
        Case CHECKSTARTINDEX 'Name
            edcStartHour.Visible = False  'Set visibility
            slStr = edcStartHour.Text
            smSave(6) = slStr
            gSetShow pbcFdNm, slStr, tmCtrls(ilBoxNo)
        Case CHECKENDINDEX 'Name
            edcEndHour.Visible = False  'Set visibility
            slStr = edcEndHour.Text
            smSave(7) = slStr
            gSetShow pbcFdNm, slStr, tmCtrls(ilBoxNo)
        Case CHECKINTERVALINDEX 'Name
            edcInterval.Visible = False  'Set visibility
            slStr = edcInterval.Text
            smSave(8) = slStr
            gSetShow pbcFdNm, slStr, tmCtrls(ilBoxNo)
        Case NETNAMEINDEX
            lbcNetName.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcNetName.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcNetName.List(lbcNetName.ListIndex)
            End If
            smSave(11) = slStr
            gSetShow pbcFdNm, slStr, tmCtrls(ilBoxNo)
        Case PRODUCERINDEX
            lbcProducer.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcProducer.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcProducer.List(lbcProducer.ListIndex)
            End If
            smSave(12) = slStr
            gSetShow pbcFdNm, slStr, tmCtrls(ilBoxNo)
        Case BOOKAVAILINDEX
            lbcAvailName.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcAvailName.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcAvailName.List(lbcAvailName.ListIndex)
            End If
            smSave(13) = slStr
            gSetShow pbcFdNm, slStr, tmCtrls(ilBoxNo)
        Case MEDIAINDEX
            lbcMedia.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcMedia.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcMedia.List(lbcMedia.ListIndex)
            End If
            smSave(14) = slStr
            gSetShow pbcFdNm, slStr, tmCtrls(ilBoxNo)
    End Select
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
    sgDoneMsg = Trim$(str$(igFdNmCallSource)) & "\" & Trim$(tmFnf.sName)
    igManUnload = YES
    Unload FeedName
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mTestFields(iTest, iState)
'   Where:
'       iTest (I)- Test all controls or control number specified
'       iState (I)- Test one of the following:
'                  ALLBLANK=All fields blank
'                  ALLMANBLANK=All mandatory
'                    field blank
'                  ALLMANDEFINED=All mandatory
'                    fields have data
'                  Plus
'                  NOMSG=No error message shown
'                  SHOWMSG=show error message
'       iRet (O)- True if all mandatory fields blank, False if not all blank
'
'
    Dim slStr As String
    If (ilCtrlNo = NAMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcName, "", "Name must be specified", tmCtrls(NAMEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = NAMEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = PLEDGETIMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imSave(1) = 0 Then
            slStr = "Pre-Converted Log"
        ElseIf imSave(1) = 1 Then
            slStr = "Log Needs Conversion"
        ElseIf imSave(1) = 2 Then
            slStr = "Insertion Order"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Pledge Time must be specified", tmCtrls(PLEDGETIMEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = PLEDGETIMEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If

    mTestFields = YES
End Function

Private Sub lbcAvailName_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcAvailName, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcAvailName_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcAvailName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcAvailName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcAvailName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcAvailName, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcDays_Click()
    gProcessLbcClick lbcDays, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcDays_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcMedia_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcMedia, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcMedia_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcMedia_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcMedia_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcMedia_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcMedia, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcNetName_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcNetName, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcNetName_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcNetName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcNetName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcNetName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcNetName, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcProducer_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcProducer, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcProducer_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcProducer_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcProducer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcProducer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcProducer, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub


Private Sub pbcFdNm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    For ilBox = imLBCtrls To imMaxBoxNo Step 1    'UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus imBoxNo
End Sub
Private Sub pbcFdNm_Paint()
    Dim ilBox As Integer
    pbcFdNm.Cls
    For ilBox = imLBCtrls To imMaxBoxNo Step 1    'UBound(tmCtrls) Step 1
        pbcFdNm.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcFdNm.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcFdNm.Print tmCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    If imBoxNo = NETNAMEINDEX Then
        If mNetRepBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = PRODUCERINDEX Then
        If mProducerBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = BOOKAVAILINDEX Then
        If mAvailBranch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= imMaxBoxNo) Then    'UBound(tmCtrls)) Then
        If (imBoxNo <> NAMEINDEX) Or (Not cbcSelect.Enabled) Then
            If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
    End If
    imTabDirection = -1  'Set-right to left
    Select Case imBoxNo
        Case -1
            imTabDirection = 0  'Set-Left to right
            If (imSelectedIndex = 0) And (cbcSelect.Text = "[New]") Then
                ilBox = 1
                mSetCommands
            Else
                mSetChg 1
                ilBox = 2
            End If
        Case 1 'Name (first control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
                Exit Sub
            End If
            ilBox = 1
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    If imBoxNo = NETNAMEINDEX Then
        If mNetRepBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = PRODUCERINDEX Then
        If mProducerBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = BOOKAVAILINDEX Then
        If mAvailBranch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= imMaxBoxNo) Then 'UBound(tmCtrls)) Then
        If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    imTabDirection = 0  'Set-Left to right
    Select Case imBoxNo
        Case -1
            imTabDirection = -1  'Set-Right to left
            ilBox = imMaxBoxNo  'UBound(tmCtrls)
        Case imMaxBoxNo 'UBound(tmCtrls) 'Suppress (last control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            If (cmcUpdate.Enabled) And (igFdNmCallSource = CALLNONE) Then
                cmcUpdate.SetFocus
            Else
                cmcDone.SetFocus
            End If
            Exit Sub
        Case Else
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcPledgeTime_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcPledgeTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("P") Or (KeyAscii = Asc("p")) Then
        If imSave(1) <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imSave(1) = 0
        pbcPledgeTime_Paint
    ElseIf KeyAscii = Asc("L") Or (KeyAscii = Asc("l")) Then
        If imSave(1) <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imSave(1) = 1
        pbcPledgeTime_Paint
    ElseIf KeyAscii = Asc("I") Or (KeyAscii = Asc("i")) Then
        If imSave(1) <> 2 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imSave(1) = 2
        pbcPledgeTime_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imSave(1) = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imSave(1) = 1
            pbcPledgeTime_Paint
        ElseIf imSave(1) = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imSave(1) = 2
            pbcPledgeTime_Paint
        ElseIf imSave(1) = 2 Then
            tmCtrls(imBoxNo).iChg = True
            imSave(1) = 0
            pbcPledgeTime_Paint
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcPledgeTime_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imSave(1) = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imSave(1) = 1
    ElseIf imSave(1) = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imSave(1) = 2
    ElseIf imSave(1) = 2 Then
        tmCtrls(imBoxNo).iChg = True
        imSave(1) = 0
    End If
    pbcPledgeTime_Paint
    mSetCommands
End Sub
Private Sub pbcPledgeTime_Paint()
    pbcPledgeTime.Cls
    pbcPledgeTime.CurrentX = fgBoxInsetX
    pbcPledgeTime.CurrentY = 0 'fgBoxInsetY
    If imSave(1) = 0 Then
        pbcPledgeTime.Print "Pre-Converted Log"
    ElseIf imSave(1) = 1 Then
        pbcPledgeTime.Print "Log Needs Conversion"
    ElseIf imSave(1) = 2 Then
        pbcPledgeTime.Print "Insertion Order"
    Else
        pbcPledgeTime.Print "   "
    End If
End Sub
Private Sub plcFdNm_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
        Case PLEDGETIMEINDEX
        Case FTPADDRESSINDEX 'Name
        Case PASSWORDINDEX 'Name
        Case CHECKDAYSINDEX
        Case CHECKSTARTINDEX 'Name
        Case CHECKENDINDEX 'Name
        Case CHECKINTERVALINDEX 'Name
        Case NETNAMEINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcNetName, edcDropDown, imChgMode, imLbcArrowSetting
        Case PRODUCERINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcProducer, edcDropDown, imChgMode, imLbcArrowSetting
        Case BOOKAVAILINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcAvailName, edcDropDown, imChgMode, imLbcArrowSetting
        Case MEDIAINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcMedia, edcDropDown, imChgMode, imLbcArrowSetting
    End Select
End Sub
