VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form EngrUserOption 
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   8340
   ControlBox      =   0   'False
   Icon            =   "EngrUserOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8340
   Begin VB.CommandButton cmcErase 
      Caption         =   "&Erase"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6030
      TabIndex        =   48
      Top             =   5715
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8100
      Top             =   4845
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6225
      FormDesignWidth =   8340
   End
   Begin VB.CommandButton cmcSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4335
      TabIndex        =   47
      Top             =   5700
      Width           =   1335
   End
   Begin VB.Frame frcDefine 
      Caption         =   "Step 2: Define User Options Properties"
      Height          =   4680
      Left            =   330
      TabIndex        =   2
      Top             =   885
      Width           =   7530
      Begin VB.TextBox edcEMail 
         Height          =   285
         Left            =   1665
         MaxLength       =   70
         TabIndex        =   10
         Top             =   1620
         Width           =   5685
      End
      Begin VB.Frame frcTab 
         Caption         =   "Alert"
         Height          =   630
         Index           =   3
         Left            =   5970
         TabIndex        =   40
         Top             =   2100
         Visible         =   0   'False
         Width           =   2685
         Begin VB.Label lacSchdAlert 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Schedule Not Retrieved"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   90
            TabIndex        =   41
            Top             =   255
            Width           =   2205
         End
      End
      Begin VB.Frame frcTab 
         Caption         =   "Notification"
         Height          =   1395
         Index           =   2
         Left            =   5190
         TabIndex        =   36
         Top             =   60
         Visible         =   0   'False
         Width           =   2760
         Begin VB.Label lacCommTest 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Commercial Test Error"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   75
            TabIndex        =   39
            Top             =   975
            Width           =   2205
         End
         Begin VB.Label lacSchdNotice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Schedule Not Retrieved"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   90
            TabIndex        =   38
            Top             =   600
            Width           =   2205
         End
         Begin VB.Label lacMerge 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Merge Errors"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   90
            TabIndex        =   37
            Top             =   240
            Width           =   2205
         End
      End
      Begin VB.Frame frcTab 
         Caption         =   "List"
         Height          =   1755
         Index           =   1
         Left            =   1455
         TabIndex        =   17
         Top             =   3900
         Visible         =   0   'False
         Width           =   5445
         Begin VB.Label lacBusControls 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Bus Controls"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2640
            TabIndex        =   24
            Top             =   435
            Width           =   1125
         End
         Begin VB.Label lacAudioControls 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Audio Controls"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   1125
         End
         Begin VB.Label lacBusGroups 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Bus Groups"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3900
            TabIndex        =   25
            Top             =   435
            Width           =   1125
         End
         Begin VB.Label lacAudioTypes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Audio Types"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3900
            TabIndex        =   21
            Top             =   120
            Width           =   1125
         End
         Begin VB.Label lacAudioNames 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Audio Names"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1380
            TabIndex        =   19
            Top             =   120
            Width           =   1125
         End
         Begin VB.Label lacUserOption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "User Options"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2640
            TabIndex        =   35
            Top             =   1410
            Width           =   1125
         End
         Begin VB.Label lacSiteOption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Site Options"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1365
            TabIndex        =   34
            Top             =   1410
            Width           =   1125
         End
         Begin VB.Label lacTimeTypes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Time Types"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3885
            TabIndex        =   33
            Top             =   1080
            Width           =   1125
         End
         Begin VB.Label lacSilence 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Silences"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2640
            TabIndex        =   32
            Top             =   1080
            Width           =   1125
         End
         Begin VB.Label lacRelay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Relays"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1365
            TabIndex        =   31
            Top             =   1080
            Width           =   1125
         End
         Begin VB.Label lacNetcues 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Netcues"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   1125
         End
         Begin VB.Label lacMaterial 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Material Types"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3900
            TabIndex        =   29
            Top             =   750
            Width           =   1125
         End
         Begin VB.Label lacFollows 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Follows"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2640
            TabIndex        =   28
            Top             =   750
            Width           =   1125
         End
         Begin VB.Label lacEventTypes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Event Types"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1380
            TabIndex        =   27
            Top             =   750
            Width           =   1125
         End
         Begin VB.Label lacComments 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Comments"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   750
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lacBuses 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Buses"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1380
            TabIndex        =   23
            Top             =   435
            Width           =   1125
         End
         Begin VB.Label lacAutomation 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Automation"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   435
            Width           =   1125
         End
         Begin VB.Label lacAudio 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Audio Sources"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2640
            TabIndex        =   20
            Top             =   120
            Width           =   1125
         End
      End
      Begin VB.Frame frcTab 
         Caption         =   "Jobs"
         Height          =   1410
         Index           =   0
         Left            =   270
         TabIndex        =   12
         Top             =   2685
         Width           =   3705
         Begin VB.Label lacSchedules 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Schedules"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1200
            TabIndex        =   16
            Top             =   1065
            Width           =   1095
         End
         Begin VB.Label lacTemplates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Templates"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1890
            TabIndex        =   15
            Top             =   705
            Width           =   1095
         End
         Begin VB.Label lacLibraries 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Libraries"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   495
            TabIndex        =   14
            Top             =   705
            Width           =   1095
         End
         Begin VB.Label lacResource 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Time Finder"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   900
            TabIndex        =   13
            Top             =   315
            Width           =   1650
         End
      End
      Begin VB.TextBox edcDisplayName 
         Height          =   285
         Left            =   1665
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1185
         Width           =   3600
      End
      Begin VB.Frame frcState 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   270
         Left            =   120
         TabIndex        =   42
         Top             =   4320
         Width           =   2100
         Begin VB.OptionButton rbcState 
            Caption         =   "Active"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton rbcState 
            Caption         =   "Dormant"
            Height          =   255
            Index           =   1
            Left            =   975
            TabIndex        =   44
            Top             =   0
            Width           =   990
         End
      End
      Begin VB.TextBox edcName 
         Height          =   285
         Left            =   1665
         MaxLength       =   40
         TabIndex        =   4
         Top             =   330
         Width           =   3600
      End
      Begin VB.TextBox edcPassword 
         Height          =   285
         Left            =   1665
         MaxLength       =   10
         TabIndex        =   6
         Top             =   735
         Width           =   1230
      End
      Begin ComctlLib.TabStrip tabUser 
         Height          =   2160
         Left            =   120
         TabIndex        =   11
         Top             =   2085
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   3810
         ShowTips        =   0   'False
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "&Jobs"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "&List"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lacEMail 
         Caption         =   "E-Mail:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1620
         Width           =   1380
      End
      Begin VB.Label lacDisplayName 
         Caption         =   "Display Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1185
         Width           =   1380
      End
      Begin VB.Label lacName 
         Caption         =   "Sign On Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   330
         Width           =   1545
      End
      Begin VB.Label lacPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   735
         Width           =   1380
      End
   End
   Begin VB.Frame frcSelect 
      Caption         =   "Step 1: Select User"
      Height          =   660
      Left            =   330
      TabIndex        =   0
      Top             =   120
      Width           =   3555
      Begin VB.ComboBox cbcSelect 
         BackColor       =   &H00FFFF80&
         Height          =   315
         ItemData        =   "EngrUserOption.frx":030A
         Left            =   150
         List            =   "EngrUserOption.frx":030C
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   3180
      End
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2595
      TabIndex        =   46
      Top             =   5700
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   810
      TabIndex        =   45
      Top             =   5715
      Width           =   1335
   End
End
Attribute VB_Name = "EngrUserOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrUserOption - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFieldChgd As Integer
Private imInChg As Integer
Private imBSMode As Integer
Private imUieCode As Integer
Private smUsedFlag As String
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private smCurrUTEStamp As String
Private smLastDatePWSet As String

Private tmUIE As UIE
Private tmCurrUTE() As UTE


Private imTabIndex As Integer

Private Sub mSetTab()
    'tabUser.Left = frcSelect.Left
    'tabUser.Height = cmcCancel.Top - (frcSelect.Top + frcSelect.Height + 300)  'tabUser.ClientTop - tabUser.Top + (10 * frcTab(0).Height) / 9
    frcTab(0).Move tabUser.ClientLeft, tabUser.ClientTop ', tabUser.ClientWidth, tabUser.ClientHeight
    frcTab(1).Move tabUser.ClientLeft, tabUser.ClientTop ', tabUser.ClientWidth, tabUser.ClientHeight
    frcTab(2).Move tabUser.ClientLeft, tabUser.ClientTop ', tabUser.ClientWidth, tabUser.ClientHeight
    frcTab(3).Move tabUser.ClientLeft, tabUser.ClientTop ', tabUser.ClientWidth, tabUser.ClientHeight
    frcTab(0).BorderStyle = 0
    frcTab(1).BorderStyle = 0
    frcTab(2).BorderStyle = 0
    frcTab(3).BorderStyle = 0
End Sub

Private Sub mClearControls(slDefault As String)
    imVersion = -1
    smUsedFlag = "N"
    gClearControls EngrUserOption
    rbcState(0).Value = False
    rbcState(1).Value = False
    ReDim tmCurrUTE(0 To 0) As UTE
    mSetCtrl slDefault, lacResource
    mSetCtrl slDefault, lacLibraries
    mSetCtrl slDefault, lacTemplates
    mSetCtrl slDefault, lacSchedules
    
    mSetCtrl slDefault, lacAutomation
    mSetCtrl slDefault, lacBuses
    mSetCtrl slDefault, lacBusControls
    mSetCtrl slDefault, lacBusGroups
    mSetCtrl slDefault, lacEventTypes
    mSetCtrl slDefault, lacTimeTypes
    mSetCtrl slDefault, lacMaterial
    mSetCtrl slDefault, lacAudio
    mSetCtrl slDefault, lacAudioControls
    mSetCtrl slDefault, lacAudioNames
    mSetCtrl slDefault, lacAudioTypes
    mSetCtrl slDefault, lacRelay
    mSetCtrl slDefault, lacFollows
    mSetCtrl slDefault, lacNetcues
    mSetCtrl slDefault, lacComments
    mSetCtrl slDefault, lacSilence
    mSetCtrl slDefault, lacSiteOption
    mSetCtrl slDefault, lacUserOption
    
    mSetCtrl slDefault, lacSchdAlert
    
    mSetCtrl slDefault, lacMerge
    mSetCtrl slDefault, lacSchdNotice
    mSetCtrl slDefault, lacCommTest
    
    imFieldChgd = False
End Sub
Private Sub mMoveRecToCtrls(ilUIEIndex As Integer)
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer
    
    edcName.text = Trim$(tgCurrUIE(ilUIEIndex).sSignOnName)
    edcPassword.text = Trim$(tgCurrUIE(ilUIEIndex).sPassword)
    edcDisplayName.text = Trim$(tgCurrUIE(ilUIEIndex).sShowName)
    edcEMail.text = Trim$(tgCurrUIE(ilUIEIndex).sEMail)
    Erase tmCurrUTE
    ilRet = gGetRecs_UTE_UserTasks(smCurrUTEStamp, tgCurrUIE(ilUIEIndex).iCode, "User Option-mMoveRecToCtrls", tmCurrUTE())
    For ilLoop = LBound(tmCurrUTE) To UBound(tmCurrUTE) - 1 Step 1
        For ilIndex = 0 To UBound(tgJobTaskNames) Step 1
            If tmCurrUTE(ilLoop).iTneCode = tgJobTaskNames(ilIndex).iCode Then
                If StrComp(Trim$(tgJobTaskNames(ilIndex).sName), lacResource.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacResource
                ElseIf StrComp(Trim$(tgJobTaskNames(ilIndex).sName), lacLibraries.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacLibraries
                ElseIf StrComp(Trim$(tgJobTaskNames(ilIndex).sName), lacTemplates.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacTemplates
                ElseIf StrComp(Trim$(tgJobTaskNames(ilIndex).sName), lacSchedules.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacSchedules
                End If
                Exit For
            End If
        Next ilIndex
        For ilIndex = 0 To UBound(tgListTaskNames) Step 1
            If tmCurrUTE(ilLoop).iTneCode = tgListTaskNames(ilIndex).iCode Then
                If StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacAutomation.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacAutomation
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacBuses.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacBuses
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacBusControls.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacBusControls
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacBusGroups.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacBusGroups
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacEventTypes.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacEventTypes
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacTimeTypes.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacTimeTypes
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacMaterial.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacMaterial
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacAudio.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacAudio
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacAudioControls.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacAudioControls
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacAudioNames.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacAudioNames
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacAudioTypes.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacAudioTypes
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacRelay.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacRelay
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacFollows.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacFollows
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacNetcues.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacNetcues
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacComments.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacComments
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacSilence.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacSilence
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacSiteOption.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacSiteOption
                ElseIf StrComp(Trim$(tgListTaskNames(ilIndex).sName), lacUserOption.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacUserOption
                End If
                Exit For
            End If
        Next ilIndex
        For ilIndex = 0 To UBound(tgAlertTaskNames) Step 1
            If tmCurrUTE(ilLoop).iTneCode = tgAlertTaskNames(ilIndex).iCode Then
                If StrComp(Trim$(tgAlertTaskNames(ilIndex).sName), lacSchdAlert.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacSchdAlert
                End If
                Exit For
            End If
        Next ilIndex
        For ilIndex = 0 To UBound(tgNoticeTaskNames) Step 1
            If tmCurrUTE(ilLoop).iTneCode = tgNoticeTaskNames(ilIndex).iCode Then
                If StrComp(Trim$(tgNoticeTaskNames(ilIndex).sName), lacMerge.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacMerge
                ElseIf StrComp(Trim$(tgNoticeTaskNames(ilIndex).sName), lacSchdNotice.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacSchdNotice
                ElseIf StrComp(Trim$(tgNoticeTaskNames(ilIndex).sName), lacCommTest.Caption, vbTextCompare) = 0 Then
                    mSetCtrl tmCurrUTE(ilLoop).sTaskStatus, lacCommTest
                End If
                Exit For
            End If
        Next ilIndex
    Next ilLoop
    If tgCurrUIE(ilUIEIndex).sState = "D" Then
        rbcState(1).Value = True
    Else
        rbcState(0).Value = True
    End If
    smLastDatePWSet = tgCurrUIE(ilUIEIndex).sLastDatePWSet
    imVersion = tgCurrUIE(ilUIEIndex).iVersion
    smUsedFlag = tgCurrUIE(ilUIEIndex).sUsedFlag
End Sub

Private Sub cbcSelect_Change()
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilLen As Integer
    Dim ilSel As Integer
    Dim llRow As Long
    
    If imInChg Then
        Exit Sub
    End If
    imInChg = True
    Screen.MousePointer = vbHourglass
    slName = LTrim$(cbcSelect.text)
    ilLen = Len(slName)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(cbcSelect.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        cbcSelect.ListIndex = llRow
        cbcSelect.SelStart = ilLen
        cbcSelect.SelLength = Len(cbcSelect.text)
        imUieCode = cbcSelect.ItemData(cbcSelect.ListIndex)
        If imUieCode <= 0 Then
            mClearControls "E"
            rbcState(0).Value = True
        Else
            'Load existing data
            mClearControls ""
            For ilLoop = LBound(tgCurrUIE) To UBound(tgCurrUIE) - 1 Step 1
                If imUieCode = tgCurrUIE(ilLoop).iCode Then
                    mMoveRecToCtrls ilLoop
                    Exit For
                End If
            Next ilLoop
        End If
    End If
    imInChg = False
    imFieldChgd = False
    mSetCommands
    Screen.MousePointer = vbDefault
    Exit Sub

End Sub

Private Sub cbcSelect_Click()
    cbcSelect_Change
End Sub

Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cbcSelect.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cmcCancel_Click()
    Unload EngrUserOption
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    On Error GoTo ErrHand
    If imFieldChgd = False Then
        Unload EngrUserOption
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        ilRet = mSave()
        If Not ilRet Then
            Exit Sub
        End If
        'Force reading of all current users
        sgCurrUIEStamp = ""
        ilRet = gGetTypeOfRecs_UIE_UserInfo("C", sgCurrUIEStamp, "EngrUser-mPopulate", tgCurrUIE())
    End If
    
    Screen.MousePointer = vbDefault
    Unload EngrUserOption
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Bus Definition-cmcDone: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    cnn.RollbackTrans
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Bus Definition-cmcDone: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub

Private Sub cmcErase_Click()
    Dim slStr As String
    Dim slMsg As String
    Dim ilRet As Integer
    
    If cmcDone.Enabled = False Then
        Exit Sub
    End If
    slStr = Trim$(edcName.text)
    If smUsedFlag <> "N" Then
        MsgBox slStr & " used or was used, unable to delete", vbInformation + vbOKCancel, "Erase"
        Exit Sub
    End If
    slMsg = "Delete " & slStr
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ilRet = gPutDelete_UIE_UserInfo(imUieCode, "EngrUserOption- Delete")
    sgCurrUIEStamp = ""
    mPopulate
    If cbcSelect.ListCount >= 1 Then
        cbcSelect.ListIndex = 0
    End If
    imFieldChgd = False
    mSetCommands
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim slName As String
    
    If imFieldChgd = True Then
        slName = Trim$(edcName.text)
        ilRet = mSave()
        If Not ilRet Then
            Exit Sub
        End If
        sgCurrUIEStamp = ""
        mPopulate
        cbcSelect.text = slName
    End If
End Sub

Private Sub edcDisplayName_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcDisplayName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcEMail_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcEMail_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPassword_Change()
    imFieldChgd = True
    smLastDatePWSet = Format$(gNow(), sgShowDateForm)
    mSetCommands
End Sub

Private Sub edcPassword_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
    mSetTab
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrUserOption
    gCenterFormModal EngrUserOption
End Sub

Private Sub Form_Load()
    Dim sName As String
    Dim sAffRepFN As String
    Dim sAffRepLN As String
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    mInit
    imTabIndex = 1
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Bus Definition-Form Load: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Bus Definition-Form Load: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase tmCurrUTE
    Set EngrUserOption = Nothing
End Sub

Private Sub lacAudio_Click()
    mSetNewStatusGYR lacAudio
End Sub

Private Sub lacAudioControls_Click()
    mSetNewStatusGYR lacAudioControls
End Sub

Private Sub lacAudioNames_Click()
    mSetNewStatusGYR lacAudioNames
End Sub

Private Sub lacAudioTypes_Click()
    mSetNewStatusGYR lacAudioTypes
End Sub

Private Sub lacAutomation_Click()
    mSetNewStatusGYR lacAutomation
End Sub

Private Sub lacBusControls_Click()
    mSetNewStatusGYR lacBusControls
End Sub

Private Sub lacBuses_Click()
    mSetNewStatusGYR lacBuses
End Sub

Private Sub lacBusGroups_Click()
    mSetNewStatusGYR lacBusGroups
End Sub

Private Sub lacComments_Click()
    mSetNewStatusGYR lacComments
End Sub

Private Sub lacCommTest_Click()
    mSetNewStatusGR lacCommTest
End Sub

Private Sub lacEMail_Click()
    mSetNewStatusGYR lacEMail
End Sub

Private Sub lacEventTypes_Click()
    mSetNewStatusGYR lacEventTypes
End Sub

Private Sub lacFollows_Click()
    mSetNewStatusGYR lacFollows
End Sub

Private Sub lacLibraries_Click()
    mSetNewStatusGYR lacLibraries
End Sub

Private Sub lacMaterial_Click()
    mSetNewStatusGYR lacMaterial
End Sub

Private Sub lacMerge_Click()
    mSetNewStatusGR lacMerge
End Sub

Private Sub lacNetcues_Click()
    mSetNewStatusGYR lacNetcues
End Sub

Private Sub lacRelay_Click()
    mSetNewStatusGYR lacRelay
End Sub

Private Sub lacResource_Click()
    mSetNewStatusGYR lacResource
End Sub

Private Sub lacSchdAlert_Click()
    mSetNewStatusGR lacSchdAlert
End Sub

Private Sub lacSchdNotice_Click()
    mSetNewStatusGR lacSchdNotice
End Sub

Private Sub lacSchedules_Click()
    mSetNewStatusGYR lacSchedules
End Sub

Private Sub lacSilence_Click()
    mSetNewStatusGYR lacSilence
End Sub

Private Sub lacSiteOption_Click()
    mSetNewStatusGYR lacSiteOption
End Sub

Private Sub lacTemplates_Click()
    mSetNewStatusGYR lacTemplates
End Sub

Private Sub lacTimeTypes_Click()
    mSetNewStatusGYR lacTimeTypes
End Sub

Private Sub lacUserOption_Click()
    mSetNewStatusGYR lacUserOption
End Sub

Private Sub rbcState_Click(Index As Integer)
    If rbcState(Index).Value Then
        imFieldChgd = True
        mSetCommands
    End If
End Sub

Private Sub edcName_Change()
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub edcName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub




Private Sub tabUser_Click()
    If imTabIndex = tabUser.SelectedItem.Index Then
        Exit Sub
    End If
    frcTab(tabUser.SelectedItem.Index - 1).Visible = True
    frcTab(imTabIndex - 1).Visible = False
    imTabIndex = tabUser.SelectedItem.Index
End Sub

Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    
    ilRet = gGetTypeOfRecs_UIE_UserInfo("C", sgCurrUIEStamp, "EngrUser-mPopulate", tgCurrUIE())
    
    cbcSelect.Clear
    cbcSelect.text = ""
    For ilLoop = 0 To UBound(tgCurrUIE) - 1 Step 1
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Then   'Or (igListStatus(USERLIST) = 2) Then
            cbcSelect.AddItem Trim$(tgCurrUIE(ilLoop).sSignOnName)
            cbcSelect.ItemData(cbcSelect.NewIndex) = tgCurrUIE(ilLoop).iCode
        Else
            If (StrComp(sgUserName, Trim$(tgCurrUIE(ilLoop).sSignOnName), vbTextCompare) = 0) Then
                cbcSelect.AddItem Trim$(tgCurrUIE(ilLoop).sSignOnName)
                cbcSelect.ItemData(cbcSelect.NewIndex) = tgCurrUIE(ilLoop).iCode
            End If
        End If
    Next ilLoop
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Then   'Or (igListStatus(USERLIST) = 2) Then
        cbcSelect.AddItem "[New]", 0
        cbcSelect.ItemData(cbcSelect.NewIndex) = 0
    End If
    
End Sub

Private Sub mInit()
    Dim ilLoop As Integer
    
    imTabIndex = 1
    imInChg = False
    imUieCode = 0
    mPopulate
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Then   'Or (igListStatus(USERLIST) = 2) Then
        cmcDone.Enabled = True
        For ilLoop = frcTab.LBound To frcTab.UBound Step 1
            frcTab(ilLoop).Enabled = True
        Next ilLoop
    Else
        frcState.Enabled = False
        For ilLoop = frcTab.LBound To frcTab.UBound Step 1
            frcTab(ilLoop).Enabled = False
        Next ilLoop
        If (igListStatus(USERLIST) <> 2) Then
            cmcDone.Enabled = False
            edcDisplayName.Enabled = False
            edcEMail.Enabled = False
            edcName.Enabled = False
            edcPassword.Enabled = False
        End If
    End If
    If cbcSelect.ListCount >= 1 Then
        cbcSelect.ListIndex = 0
    End If
    imFieldChgd = False
    mSetCommands
End Sub
Private Sub mPutCtrl(lacCtrl As Label, slTaskStatus As String)
    If lacCtrl.BackColor = vbGreen Then
        slTaskStatus = "E"
    ElseIf lacCtrl.BackColor = vbYellow Then
        slTaskStatus = "V"
    Else
        slTaskStatus = "D"
    End If

End Sub

Private Sub mSetCtrl(slTaskStatus As String, lacCtrl As Label)
    If slTaskStatus = "E" Then
        lacCtrl.BackColor = vbGreen
    ElseIf slTaskStatus = "V" Then
        lacCtrl.BackColor = vbYellow
    ElseIf slTaskStatus = "D" Then
        lacCtrl.BackColor = vbRed
    Else
        lacCtrl.BackColor = vbWhite
    End If

End Sub

Private Sub mSetCommands()
    Dim ilRet As Integer
    If imInChg Then
        Exit Sub
    End If
    If cmcDone.Enabled = False Then
        Exit Sub
    End If
    If imFieldChgd Then
        'Check that all mandatory answered
        ilRet = mCheckFields(False)
        If ilRet Then
            cmcSave.Enabled = True
            cbcSelect.Enabled = False
            cmcErase.Enabled = False
        Else
            cmcSave.Enabled = False
            cbcSelect.Enabled = False
            cmcErase.Enabled = False
        End If
    Else
        cmcSave.Enabled = False
        cbcSelect.Enabled = True
        If (cbcSelect.ListCount <= 1) Or (imUieCode <= 0) Or (smUsedFlag <> "N") Then
            cmcErase.Enabled = False
        Else
            cmcErase.Enabled = True
        End If
    End If
End Sub

Private Sub mSetNewStatusGYR(lacCtrl As Label)
    If lacCtrl.BackColor = vbGreen Then
        lacCtrl.BackColor = vbRed
    ElseIf lacCtrl.BackColor = vbYellow Then
        lacCtrl.BackColor = vbGreen
    Else
        lacCtrl.BackColor = vbYellow
    End If
    imFieldChgd = True
    mSetCommands
End Sub
Private Sub mSetNewStatusGR(lacCtrl As Label)
    If lacCtrl.BackColor = vbGreen Then
        lacCtrl.BackColor = vbRed
    Else
        lacCtrl.BackColor = vbGreen
    End If
    imFieldChgd = True
    mSetCommands
End Sub

Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim tlUte As UTE
    
    Screen.MousePointer = vbHourglass
    If Not mCheckFields(True) Then
        Screen.MousePointer = vbDefault
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        Screen.MousePointer = vbDefault
        mSave = False
        Exit Function
    End If
    mMoveCtrlsToRec
    If imUieCode <= 0 Then
        ilRet = gPutInsert_UIE_UserInfo(0, tmUIE, "User Option-mSave: UIE")
    Else
        ilRet = gPutUpdate_UIE_UserInfo(1, tmUIE, "User Option-mSave: UIE")
    End If
    'Insert Task setting
    For ilLoop = 0 To UBound(tgJobTaskNames) Step 1
        tlUte.iCode = 0
        tlUte.iUieCode = tmUIE.iCode
        tlUte.iTneCode = tgJobTaskNames(ilLoop).iCode
        If StrComp(Trim$(tgJobTaskNames(ilLoop).sName), lacResource.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacResource, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgJobTaskNames(ilLoop).sName), lacLibraries.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacLibraries, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgJobTaskNames(ilLoop).sName), lacTemplates.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacTemplates, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgJobTaskNames(ilLoop).sName), lacSchedules.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacSchedules, tlUte.sTaskStatus
        End If
        tlUte.sUnused = ""
        ilRet = gPutInsert_UTE_UserTasks(tlUte, "User Option-mSave: UTE-Job")
    Next ilLoop
     'Add Task Lists
    For ilLoop = 0 To UBound(tgListTaskNames) Step 1
        tlUte.iCode = 0
        tlUte.iUieCode = tmUIE.iCode
        tlUte.iTneCode = tgListTaskNames(ilLoop).iCode
        If StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacAutomation.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacAutomation, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacBuses.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacBuses, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacBusControls.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacBusControls, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacBusGroups.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacBusGroups, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacEventTypes.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacEventTypes, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacTimeTypes.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacTimeTypes, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacMaterial.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacMaterial, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacAudio.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacAudio, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacAudioControls.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacAudioControls, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacAudioNames.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacAudioNames, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacAudioTypes.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacAudioTypes, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacRelay.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacRelay, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacFollows.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacFollows, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacNetcues.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacNetcues, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacComments.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacComments, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacSilence.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacSilence, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacSiteOption.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacSiteOption, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgListTaskNames(ilLoop).sName), lacUserOption.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacUserOption, tlUte.sTaskStatus
        End If
        tlUte.sUnused = ""
        ilRet = gPutInsert_UTE_UserTasks(tlUte, "User Option-mSave: UTE-List")
    Next ilLoop
    For ilLoop = 0 To UBound(tgAlertTaskNames) Step 1
        tlUte.iCode = 0
        tlUte.iUieCode = tmUIE.iCode
        tlUte.iTneCode = tgAlertTaskNames(ilLoop).iCode
        If StrComp(Trim$(tgAlertTaskNames(ilLoop).sName), lacSchdAlert.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacSchdAlert, tlUte.sTaskStatus
        End If
        tlUte.sUnused = ""
        ilRet = gPutInsert_UTE_UserTasks(tlUte, "User Option-mSave: UTE-Job")
    Next ilLoop
    For ilLoop = 0 To UBound(tgNoticeTaskNames) Step 1
        tlUte.iCode = 0
        tlUte.iUieCode = tmUIE.iCode
        tlUte.iTneCode = tgNoticeTaskNames(ilLoop).iCode
        If StrComp(Trim$(tgNoticeTaskNames(ilLoop).sName), lacMerge.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacMerge, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgNoticeTaskNames(ilLoop).sName), lacSchdNotice.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacSchdNotice, tlUte.sTaskStatus
        ElseIf StrComp(Trim$(tgNoticeTaskNames(ilLoop).sName), lacCommTest.Caption, vbTextCompare) = 0 Then
            mPutCtrl lacCommTest, tlUte.sTaskStatus
        End If
        tlUte.sUnused = ""
        ilRet = gPutInsert_UTE_UserTasks(tlUte, "User Option-mSave: UTE-Job")
    Next ilLoop
    If tgUIE.iCode = tmUIE.iCode Then
        LSet tgUIE = tmUIE
    End If
    sgCurrUIEStamp = ""
    ilRet = gGetTypeOfRecs_UIE_UserInfo("C", sgCurrUIEStamp, "EngrUser-mPopulate", tgCurrUIE())
    imFieldChgd = False
    mSetCommands
    mSave = True
    Screen.MousePointer = vbDefault
End Function

Private Function mCheckFields(ilShowMsg As Integer) As Integer
    Dim slStr As String
    
    mCheckFields = True
    slStr = Trim$(edcName.text)
    If slStr = "" Then
        If ilShowMsg Then
            MsgBox "Sign On Names must be Defined", vbCritical + vbOKOnly, "Save not Completed"
            edcName.SetFocus
        End If
        mCheckFields = False
    End If
    slStr = Trim$(edcPassword.text)
    If slStr = "" Then
        If ilShowMsg Then
            MsgBox "Password must be Defined", vbCritical + vbOKOnly, "Save not Completed"
            edcPassword.SetFocus
        End If
        mCheckFields = False
    End If
    slStr = Trim$(edcDisplayName.text)
    If slStr = "" Then
        If ilShowMsg Then
            MsgBox "Display Name must be Defined", vbCritical + vbOKOnly, "Save not Completed"
            edcDisplayName.SetFocus
        End If
        mCheckFields = False
    End If
    If (rbcState(0).Value = False) And (rbcState(1).Value = False) Then
        If ilShowMsg Then
            MsgBox "Active or Dormant must be Defined", vbCritical + vbOKOnly, "Save not Completed"
            rbcState(0).SetFocus
        End If
        mCheckFields = False
    End If
End Function

Private Sub mMoveCtrlsToRec()
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    tmUIE.iCode = imUieCode
    tmUIE.sSignOnName = edcName.text
    tmUIE.sPassword = edcPassword.text
    tmUIE.sLastDatePWSet = smLastDatePWSet
    tmUIE.sShowName = edcDisplayName.text
    tmUIE.sEMail = edcEMail.text
    If rbcState(1).Value Then
        tmUIE.sState = "D"
    Else
        tmUIE.sState = "A"
    End If
    tmUIE.sLastSignOnDate = smNowDate
    tmUIE.sLastSignOnTime = smNowTime
    tmUIE.sUsedFlag = smUsedFlag
    tmUIE.iVersion = imVersion + 1
    tmUIE.iOrigUieCode = imUieCode
    tmUIE.sCurrent = "Y"
    'tmUIE.sEnteredDate = smNowDate
    'tmUIE.sEnteredTime = smNowTime
    tmUIE.sEnteredDate = Format(Now, sgShowDateForm)
    tmUIE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmUIE.iUieCode = tgUIE.iCode
    tmUIE.sUnused = ""
End Sub


Private Function mNameOk() As Integer
    Dim slName As String
    Dim ilLoop As Integer
    
    slName = Trim$(edcName.text)
    For ilLoop = 0 To UBound(tgCurrUIE) - 1 Step 1
        If (StrComp(slName, Trim$(tgCurrUIE(ilLoop).sSignOnName), vbTextCompare) = 0) Then
            If imUieCode <> tgCurrUIE(ilLoop).iCode Then
               MsgBox "Name previously used", vbOKOnly + vbExclamation, "Name Used"
               edcName.SetFocus
               mNameOk = False
            End If
        End If
    Next ilLoop
    mNameOk = True
    
End Function

