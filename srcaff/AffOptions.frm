VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   Icon            =   "AffOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Tag             =   "Options"
   Begin VB.Frame frcTab 
      Caption         =   "Restrictions"
      Height          =   4065
      Index           =   2
      Left            =   240
      TabIndex        =   38
      Top             =   1200
      Visible         =   0   'False
      Width           =   8955
      Begin VB.Frame frcGeneral 
         Caption         =   "General"
         Height          =   1530
         Left            =   4560
         TabIndex        =   44
         Top             =   360
         Width           =   4065
         Begin VB.Label lacGeneral 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Export Specifications"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   62
            Top             =   1050
            Width           =   3615
         End
         Begin VB.Label lacGeneral 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Override Export Queue Priority"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   61
            Top             =   660
            Width           =   3615
         End
         Begin VB.Label lacGeneral 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Activity Log"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   45
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame frcComment 
         Caption         =   "Comment"
         Height          =   1095
         Left            =   4560
         TabIndex        =   51
         Top             =   2760
         Width           =   4065
         Begin VB.Label lacComment 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Change"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   52
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label lacComment 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Delete"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   53
            Top             =   645
            Width           =   3615
         End
      End
      Begin VB.Frame frcUtils 
         Caption         =   "Utility"
         Height          =   1095
         Left            =   360
         TabIndex        =   48
         Top             =   2760
         Width           =   4065
         Begin VB.Label lacUtils 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Delete Posted Spots"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   50
            Top             =   645
            Width           =   3615
         End
         Begin VB.Label lacUtils 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Clear Posted Spots"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   49
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame frcAlerts 
         Caption         =   "Alerts"
         Height          =   2385
         Left            =   360
         TabIndex        =   39
         Top             =   360
         Width           =   4065
         Begin VB.Label lacAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Web Vendor"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   5
            Left            =   240
            TabIndex        =   65
            Top             =   1770
            Width           =   3615
         End
         Begin VB.Label lacAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Initiate Shutdown"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   4
            Left            =   240
            TabIndex        =   43
            Top             =   1395
            Width           =   3615
         End
         Begin VB.Label lacAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Traffic Reprint Logs/CP's"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   240
            TabIndex        =   42
            Top             =   1020
            Width           =   3615
         End
         Begin VB.Label lacAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Export ISCI"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   41
            Top             =   615
            Width           =   3615
         End
         Begin VB.Label lacAlerts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Export Spots"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   40
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame frcPledge 
         Caption         =   "Agreement"
         Height          =   735
         Left            =   4560
         TabIndex        =   46
         Top             =   1950
         Width           =   4065
         Begin VB.Label lacRestriction 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Define Pledge Information"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   47
            Top             =   240
            Width           =   3615
         End
      End
   End
   Begin VB.Frame frcTab 
      Caption         =   "Directory"
      Height          =   2205
      Index           =   1
      Left            =   510
      TabIndex        =   23
      Top             =   1590
      Visible         =   0   'False
      Width           =   7950
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Export Center"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   15
         Left            =   2580
         TabIndex        =   60
         Top             =   1215
         Width           =   2100
      End
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Post-Buy Planning"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   14
         Left            =   4950
         TabIndex        =   59
         Top             =   765
         Width           =   2100
      End
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Affiliate Management"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   225
         TabIndex        =   24
         Top             =   1215
         Width           =   2100
      End
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Log"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   735
         TabIndex        =   36
         Top             =   2475
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C.P."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   1230
         TabIndex        =   37
         Top             =   2460
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pre Log"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   270
         TabIndex        =   35
         Top             =   2445
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Affiliate A/E"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   13
         Left            =   1635
         TabIndex        =   32
         Top             =   2445
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RADAR"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   12
         Left            =   4950
         TabIndex        =   31
         Top             =   1650
         Width           =   2100
      End
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Emails by Vehicle"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   11
         Left            =   4950
         TabIndex        =   30
         Top             =   315
         Width           =   2100
      End
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Site Options"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   225
         TabIndex        =   33
         Top             =   1665
         Width           =   2100
      End
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "User Options"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   2580
         TabIndex        =   34
         Top             =   1665
         Width           =   2100
      End
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contact"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   3135
         TabIndex        =   29
         Top             =   2505
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Affiliate Affidavits"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   2580
         TabIndex        =   28
         Top             =   765
         Width           =   2100
      End
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Network Log"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   225
         TabIndex        =   27
         Top             =   765
         Width           =   2100
      End
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Affiliate Agreements"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   2580
         TabIndex        =   26
         Top             =   315
         Width           =   2100
      End
      Begin VB.Label lacWin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stations"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   225
         TabIndex        =   25
         Top             =   315
         Width           =   2100
      End
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      Caption         =   "Name Information"
      ForeColor       =   &H80000008&
      Height          =   3600
      Index           =   0
      Left            =   9075
      TabIndex        =   6
      Top             =   1290
      Width           =   9270
      Begin VB.ComboBox cboSalesSource 
         Height          =   315
         ItemData        =   "AffOptions.frx":08CA
         Left            =   1785
         List            =   "AffOptions.frx":08CC
         TabIndex        =   64
         Top             =   2880
         Width           =   4230
      End
      Begin V81Affiliate.CSI_ComboBoxList cbcDepartment 
         Height          =   330
         Left            =   1785
         TabIndex        =   22
         Top             =   2400
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   529
         BackColor       =   -2147483639
         ForeColor       =   -2147483639
         BorderStyle     =   1
      End
      Begin VB.TextBox edcInitials 
         Height          =   285
         Left            =   6990
         MaxLength       =   3
         TabIndex        =   10
         Top             =   180
         Width           =   555
      End
      Begin VB.CommandButton cmdErase 
         Caption         =   "Erase Password"
         Height          =   315
         Left            =   6960
         TabIndex        =   58
         Top             =   1050
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox edcEMail 
         Height          =   285
         Left            =   1785
         MaxLength       =   80
         TabIndex        =   20
         Top             =   2025
         Width           =   7425
      End
      Begin VB.TextBox edcCity 
         Height          =   285
         Left            =   1785
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1665
         Width           =   5595
      End
      Begin VB.TextBox edcPhone 
         Height          =   285
         Left            =   1785
         MaxLength       =   25
         TabIndex        =   16
         Top             =   1305
         Width           =   3810
      End
      Begin VB.TextBox txtPW 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1785
         MaxLength       =   20
         TabIndex        =   14
         Top             =   930
         Width           =   3810
      End
      Begin VB.TextBox txtRptName 
         Height          =   285
         Left            =   1785
         MaxLength       =   20
         TabIndex        =   12
         Top             =   540
         Width           =   3810
      End
      Begin VB.TextBox txtSignName 
         Height          =   285
         Left            =   1785
         MaxLength       =   20
         TabIndex        =   8
         Top             =   180
         Width           =   3810
      End
      Begin VB.Label Label3 
         Caption         =   "Report Filter:"
         Height          =   255
         Left            =   180
         TabIndex        =   63
         Top             =   3000
         Width           =   1635
      End
      Begin VB.Label lacDepartment 
         Caption         =   "Department:"
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   2430
         Width           =   1635
      End
      Begin VB.Label Label4 
         Caption         =   "Initials:"
         Height          =   255
         Left            =   6300
         TabIndex        =   9
         Top             =   195
         Width           =   690
      End
      Begin VB.Label lacEMail 
         Caption         =   "E-Mail Address:"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   2040
         Width           =   1635
      End
      Begin VB.Label lacCity 
         Caption         =   "City:"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label lacPhone 
         Caption         =   "Phone:"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Label labNewPW 
         Caption         =   "Password:"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   945
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Name on Report:"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   555
         Width           =   1635
      End
      Begin VB.Label Label6 
         Caption         =   "Sign on Name:"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   195
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   555
      Left            =   5250
      TabIndex        =   3
      Top             =   150
      Width           =   1275
      Begin VB.OptionButton optState 
         Caption         =   "Dormant"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   285
         Width           =   1095
      End
      Begin VB.OptionButton optState 
         Caption         =   "Active"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   15
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4260
      TabIndex        =   55
      Tag             =   "OK"
      Top             =   5415
      Width           =   1335
   End
   Begin VB.Frame frcSelect 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   4965
      Begin VB.ComboBox cboSelect 
         Height          =   315
         ItemData        =   "AffOptions.frx":08CE
         Left            =   630
         List            =   "AffOptions.frx":08D0
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   195
         Width           =   4230
      End
      Begin VB.Label Label1 
         Caption         =   "User:"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   225
         Width           =   465
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   195
      Top             =   5595
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5865
      FormDesignWidth =   9915
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   2745
      TabIndex        =   54
      Tag             =   "OK"
      Top             =   5415
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5775
      TabIndex        =   56
      Tag             =   "Cancel"
      Top             =   5415
      Width           =   1335
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4530
      Left            =   150
      TabIndex        =   57
      Top             =   810
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   7990
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Name"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Jobs"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Restrictions"
            Key             =   ""
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
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmOptions - enter site and user global options for Affiliate program
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Dim imFirstTime As Integer
Dim imUstCode As Integer
Dim smPassword As String
Dim smEmail As String
Dim lmEMailCefCode As Long
Dim imTabIndex As Integer
Private bmInCancel As Boolean
Private imClickCount As Integer
Private imDntIndex As Integer
Private imInChg As Integer
Private imBSMode As Integer
Private IsOptionsDirty As Boolean
Private bmChangeMade As Boolean     'change by user; must save or cancel to leave record
Private bmIgnoreChange As Boolean   'change set by changing records; ignore
Private Const LOGFILE As String = ""
Private Const FORMNAME As String = "FrmOptions"
Private imSSCode As Integer


Private Sub ClearControls()
    'Dan M
    bmIgnoreChange = True
    imUstCode = 0
    'cboUser.Text = "New"
    optState(0).Enabled = True
    optState(1).Enabled = True
    optState(0).Value = True
    txtSignName.Enabled = True
    txtSignName.Text = ""
    txtRptName.Text = ""
    txtPW = ""
'    txtCurrPW.text = ""
'    txtNewPW.text = ""
'    txtVerPW.text = ""
    edcPhone.Text = ""
    edcCity.Text = ""
    edcEMail.Text = ""
    smEmail = ""
    lmEMailCefCode = 0
    edcInitials.Text = ""
    optState(0).Value = True
    If bgRemoteExport Then
        lacWin(0).BackColor = vbRed   'vbRed 'Affiliate CRM
        lacWin(1).BackColor = vbGreen
        lacWin(2).BackColor = vbRed
    '    lacWin(3).BackColor = vbGreen
    '    lacWin(4).BackColor = vbGreen
        lacWin(5).BackColor = vbRed
    '    lacWin(6).BackColor = vbGreen
        '12/15/09:  Request by Anna and Mary to only have one person setup as allowed
        lacWin(7).BackColor = vbRed
        lacWin(8).BackColor = vbRed
        lacWin(9).BackColor = vbGreen
        lacWin(10).BackColor = vbGreen
        lacWin(11).BackColor = vbRed
        lacWin(12).BackColor = vbRed
        'Affiliate A/E replaced by Post Buy
        lacWin(13).BackColor = vbRed    'vbGreen
        'Post Buy
        lacWin(14).BackColor = vbRed
        lacWin(15).BackColor = vbRed     'Export
        lacRestriction(0).BackColor = vbRed
        lacGeneral(0).BackColor = vbRed
        lacGeneral(1).BackColor = vbRed
        lacGeneral(2).BackColor = vbRed
        lacAlerts(1).BackColor = vbRed
        lacAlerts(2).BackColor = vbRed
        lacAlerts(3).BackColor = vbRed
        lacAlerts(4).BackColor = vbRed
        lacUtils(0).BackColor = vbRed
        lacUtils(1).BackColor = vbRed
        lacComment(0).BackColor = vbRed
        lacComment(1).BackColor = vbRed
    Else
        lacWin(0).BackColor = vbGreen   'vbRed 'Affiliate CRM
        lacWin(1).BackColor = vbGreen
        lacWin(2).BackColor = vbGreen
    '    lacWin(3).BackColor = vbGreen
    '    lacWin(4).BackColor = vbGreen
        lacWin(5).BackColor = vbGreen
    '    lacWin(6).BackColor = vbGreen
        '12/15/09:  Request by Anna and Mary to only have one person setup as allowed
        lacWin(7).BackColor = vbYellow
        lacWin(8).BackColor = vbGreen
        lacWin(9).BackColor = vbGreen
        lacWin(10).BackColor = vbGreen
        If ((Asc(sgSpfUsingFeatures5) And STATIONINTERFACE) <> STATIONINTERFACE) Then
            lacWin(11).BackColor = vbRed
        Else
            lacWin(11).BackColor = vbGreen
        End If
        If bgNoRADAR Then
            lacWin(12).BackColor = vbRed
        Else
            lacWin(12).BackColor = vbGreen
        End If
        'Affiliate A/E replaced by Post Buy
        lacWin(13).BackColor = vbRed    'vbGreen
        'Post Buy
        lacWin(14).BackColor = vbGreen
        lacWin(15).BackColor = vbYellow     'Export
        lacRestriction(0).BackColor = vbGreen
        lacGeneral(0).BackColor = vbRed
        lacGeneral(1).BackColor = vbRed
        lacGeneral(2).BackColor = vbRed
        lacAlerts(1).BackColor = vbGreen
        lacAlerts(2).BackColor = vbGreen
        lacAlerts(3).BackColor = vbGreen
        lacAlerts(4).BackColor = vbGreen
        lacUtils(0).BackColor = vbGreen
        lacUtils(1).BackColor = vbGreen
        lacComment(0).BackColor = vbRed
        lacComment(1).BackColor = vbRed
    End If
    '8273
    lacAlerts(5).BackColor = vbGreen
    cbcDepartment.SetListIndex = -1
    If (StrComp(sgUserName, "Guide", 1) = 0) Then
        If bgLimitedGuide Then
            mSetForLimitedGuide
        Else
            mSetForInternalGuide
        End If
    End If
    cboSalesSource.ListIndex = 0
    imSSCode = 0
    bmIgnoreChange = False
End Sub

Private Sub BindControls()
    Dim iLoop As Integer
    Dim ilRet As Integer
    Dim ilDnt As Integer
    
    'Dan M
    bmIgnoreChange = True
    imUstCode = rst!ustCode  '(0).Value
    'cboUser.Text = Trim$(rst!ustName)   '(1).Value)
    txtSignName.Text = Trim$(rst!ustname)   '(1).Value)
    If StrComp(Trim$(rst!ustname), "Guide", 1) = 0 Then
        txtSignName.Enabled = False
    Else
        txtSignName.Enabled = True
    End If
    txtRptName.Text = Trim$(rst!ustReportName)
    ' Dan is this needed anymore?
    smPassword = Trim$(rst!ustpassword)
    txtPW.Text = Trim$(rst!ustpassword)
    If (StrComp(sgUserName, "Guide", 1) = 0) Then
        If bgLimitedGuide Then
            mSetForLimitedGuide
        Else
            mSetForInternalGuide
        End If
    End If
   ' If (StrComp(sgUserName, "Counterpoint", 1) = 0) Or (StrComp(sgUserName, "Guide", 1) = 0) Then
'        txtCurrPW.text = ""
'    Else
'        txtCurrPW.text = Trim$(rst!ustpassword)
'    End If
'    txtNewPW.text = ""
'    txtVerPW.text = ""
    edcPhone.Text = Trim$(rst!ustPhoneNo)
    edcCity.Text = Trim$(rst!ustCity)
    lmEMailCefCode = rst!ustEmailcefcode
    smEmail = ""
    If lmEMailCefCode > 0 Then
        ilRet = mGetCefComment(lmEMailCefCode, smEmail)
    End If
    edcEMail.Text = smEmail
    If rst!ustState = 1 Then
        optState(1).Value = True
    Else
        optState(0).Value = True
    End If
    'changed values, moved to mSet...
'    If StrComp(Trim$(rst!ustName), "Guide", 1) = 0 Then
'        optState(0).Enabled = False
'        optState(1).Enabled = False
'    Else
'        optState(0).Enabled = True
'        optState(1).Enabled = True
'    End If
    'Affiliate Management
    If rst!ustWin16 = "I" Then
        lacWin(0).BackColor = vbGreen
    ElseIf rst!ustWin16 = "V" Then
        lacWin(0).BackColor = vbYellow
    Else
        lacWin(0).BackColor = vbRed
    End If
    'Post By
    If rst!ustWin13 = "I" Then
        lacWin(14).BackColor = vbGreen
    ElseIf rst!ustWin13 = "V" Then
        lacWin(14).BackColor = vbYellow
    Else
        lacWin(14).BackColor = vbRed
    End If
    If rst!ustWin1 = "I" Then
        lacWin(1).BackColor = vbGreen
    ElseIf rst!ustWin1 = "V" Then
        lacWin(1).BackColor = vbYellow
    Else
        lacWin(1).BackColor = vbRed
    End If
    If rst!ustWin2 = "I" Then
        lacWin(2).BackColor = vbGreen
    ElseIf rst!ustWin2 = "V" Then
        lacWin(2).BackColor = vbYellow
    Else
        lacWin(2).BackColor = vbRed
    End If
'    If rst!ustWin3 = "I" Then
'        lacWin(3).BackColor = vbGreen
'    ElseIf rst!ustWin3 = "V" Then
'        lacWin(3).BackColor = vbYellow
'    Else
'        lacWin(3).BackColor = vbRed
'    End If
'    If rst!ustWin4 = "I" Then
'        lacWin(4).BackColor = vbGreen
'    ElseIf rst!ustWin4 = "V" Then
'        lacWin(4).BackColor = vbYellow
'    Else
'        lacWin(4).BackColor = vbRed
'    End If
    If rst!ustWin5 = "I" Then
        lacWin(5).BackColor = vbGreen
    ElseIf rst!ustWin5 = "V" Then
        lacWin(5).BackColor = vbYellow
    Else
        lacWin(5).BackColor = vbRed
    End If
'    If rst!ustWin6 = "I" Then
'        lacWin(6).BackColor = vbGreen
'    ElseIf rst!ustWin6 = "V" Then
'        lacWin(6).BackColor = vbYellow
'    Else
'        lacWin(6).BackColor = vbRed
'    End If
    If rst!ustWin7 = "I" Then
        lacWin(7).BackColor = vbGreen
    ElseIf rst!ustWin7 = "V" Then
        lacWin(7).BackColor = vbYellow
    Else
        lacWin(7).BackColor = vbRed
    End If
    If rst!ustWin8 = "I" Then
        lacWin(8).BackColor = vbGreen
    ElseIf rst!ustWin8 = "V" Then
        lacWin(8).BackColor = vbYellow
    Else
        lacWin(8).BackColor = vbRed
    End If
    If rst!ustWin9 = "I" Then
        lacWin(9).BackColor = vbGreen
    ElseIf rst!ustWin9 = "V" Then
        lacWin(9).BackColor = vbYellow
    Else
        lacWin(9).BackColor = vbRed
    End If
    If rst!ustWin10 = "I" Then
        lacWin(10).BackColor = vbGreen
    ElseIf rst!ustWin10 = "V" Then
        lacWin(10).BackColor = vbYellow
    Else
        lacWin(10).BackColor = vbRed
    End If
    If rst!ustWin11 = "I" Then
        lacWin(11).BackColor = vbGreen
    ElseIf rst!ustWin11 = "V" Then
        lacWin(11).BackColor = vbYellow
    Else
        lacWin(11).BackColor = vbRed
    End If
    ' Dan M added no radar flag
    If bgNoRADAR Then
        lacWin(12).BackColor = vbRed
    ElseIf rst!ustWin12 = "I" Then  'RADAR
        lacWin(12).BackColor = vbGreen
    ElseIf rst!ustWin12 = "V" Then
        lacWin(12).BackColor = vbYellow
    Else
        lacWin(12).BackColor = vbRed
    End If
    If rst!ustWin17 = "I" Then
        lacWin(15).BackColor = vbGreen
    ElseIf rst!ustWin17 = "V" Then
        lacWin(15).BackColor = vbYellow
    Else
        lacWin(15).BackColor = vbRed
    End If
    'Affiliate A/E: Replaced by Post Buy
    'If rst!ustWin13 = "I" Then
    '    lacWin(13).BackColor = vbGreen
    'ElseIf rst!ustWin13 = "V" Then
    '    lacWin(13).BackColor = vbYellow
    'Else
    '    lacWin(13).BackColor = vbRed
    'End If
    If rst!ustPledge = "Y" Then
        lacRestriction(0).BackColor = vbGreen
    Else
        lacRestriction(0).BackColor = vbRed
    End If
    If rst!ustActivityLog = "V" Then
        lacGeneral(0).BackColor = vbYellow
    Else
        lacGeneral(0).BackColor = vbRed
    End If
    If rst!ustChgExptPriority = "Y" Then
        lacGeneral(1).BackColor = vbGreen
    Else
        lacGeneral(1).BackColor = vbRed
    End If
    If rst!ustExptSpec = "Y" Then
        lacGeneral(2).BackColor = vbGreen
    Else
        lacGeneral(2).BackColor = vbRed
    End If
    If rst!ustExptSpotAlert = "Y" Then
        lacAlerts(1).BackColor = vbGreen
    Else
        lacAlerts(1).BackColor = vbRed
    End If
    If rst!ustExptISCIAlert = "Y" Then
        lacAlerts(2).BackColor = vbGreen
    Else
        lacAlerts(2).BackColor = vbRed
    End If
    If rst!ustTrafLogAlert = "Y" Then
        lacAlerts(3).BackColor = vbGreen
    Else
        lacAlerts(3).BackColor = vbRed
    End If
    If rst!ustAllowedToBlock = "Y" Then
        lacAlerts(4).BackColor = vbGreen
    Else
        lacAlerts(4).BackColor = vbRed
    End If
    '8273
    If rst!ustVendorAlert = "N" Then
        lacAlerts(5).BackColor = vbRed
    Else
        lacAlerts(5).BackColor = vbGreen
    End If
    If rst!ustWin14 = "Y" Then
        lacUtils(0).BackColor = vbGreen
    Else
        lacUtils(0).BackColor = vbRed
    End If
    
    If rst!ustWin15 = "Y" Then
        lacUtils(1).BackColor = vbGreen
    Else
        lacUtils(1).BackColor = vbRed
    End If
    If rst!ustAllowCmmtChg = "Y" Then
        lacComment(0).BackColor = vbGreen
    Else
        lacComment(0).BackColor = vbRed
    End If
    If rst!ustAllowCmmtDelete = "Y" Then
        lacComment(1).BackColor = vbGreen
    Else
        lacComment(1).BackColor = vbRed
    End If
    edcInitials.Text = Trim$(rst!ustUserInitials)
    cbcDepartment.SetListIndex = -1
    If rst!ustDntCode > 0 Then
        For ilDnt = 0 To cbcDepartment.ListCount - 1 Step 1
            If cbcDepartment.GetItemData(ilDnt) = rst!ustDntCode Then
                cbcDepartment.SetListIndex = ilDnt
                Exit For
            End If
        Next ilDnt
    End If
    cboSalesSource.ListIndex = 0
    imSSCode = 0
    If rst!ustSSMnfCode > 0 Then
        For ilDnt = 0 To cboSalesSource.ListCount - 1 Step 1
            If cboSalesSource.ItemData(ilDnt) = rst!ustSSMnfCode Then
                cboSalesSource.ListIndex = ilDnt
                imSSCode = cboSalesSource.ItemData(ilDnt)
                Exit For
            End If
        Next ilDnt
    End If
    
    bmIgnoreChange = False
    Exit Sub
End Sub

Private Sub cbcDepartment_LostFocus()
    Dim ilDnt As Integer
    
    imDntIndex = cbcDepartment.ListIndex
    If imDntIndex = 0 Then
        If Not frmDepartment.Visible Then
            sgDepartmentName = ""
            frmDepartment.Show vbModal
            mPopDepartment
            If igDepartmentReturn Then
                For ilDnt = 0 To cbcDepartment.ListCount - 1 Step 1
                    If cbcDepartment.GetItemData(ilDnt) = igDepartmentReturnCode Then
                        cbcDepartment.SetListIndex = ilDnt
                        Exit Sub
                    End If
                Next ilDnt
            End If
            cbcDepartment.SetListIndex = 1
        End If
    End If

End Sub

Private Sub cbcDepartment_ReSetLoc()
    cbcDepartment.Top = lacDepartment.Top - 30 + edcEMail.Height - cbcDepartment.Height
End Sub

Private Sub cboSalesSource_Change()

    Dim sName As String
    Dim iLen As Integer
    Dim iSel As Integer
    Dim lRow As Long
    
    
    If imInChg Then
        Exit Sub
    End If
    imInChg = True
    Screen.MousePointer = vbHourglass
    sName = LTrim$(cboSalesSource.Text)
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    lRow = SendMessageByString(cboSalesSource.hwnd, CB_FINDSTRING, -1, sName)
    If lRow >= 0 Then
        cboSalesSource.ListIndex = lRow
        cboSalesSource.SelStart = iLen
        cboSalesSource.SelLength = Len(cboSalesSource.Text)
        imSSCode = cboSalesSource.ItemData(cboSalesSource.ListIndex)
    End If
    Screen.MousePointer = vbDefault
    imInChg = False
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError LOGFILE, FORMNAME & "-cboSalesSource_Change"
End Sub



Private Sub cboSalesSource_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboSalesSource.SelLength <> 0 Then
            imBSMode = True
        End If
    End If

End Sub

Private Sub cboSelect_Change()
    Dim iLoop As Integer
    Dim sName As String
    Dim iLen As Integer
    Dim iSel As Integer
    Dim lRow As Long
    Dim iZone As Integer
    
    If imInChg Then
        Exit Sub
    End If
    imInChg = True
    Screen.MousePointer = vbHourglass
    sName = LTrim$(cboSelect.Text)
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    lRow = SendMessageByString(cboSelect.hwnd, CB_FINDSTRING, -1, sName)
    If lRow >= 0 Then
        cboSelect.ListIndex = lRow
        cboSelect.SelStart = iLen
        cboSelect.SelLength = Len(cboSelect.Text)
        imUstCode = cboSelect.ItemData(cboSelect.ListIndex)
        If imUstCode <= 0 Then
            ClearControls
            IsOptionsDirty = False
        Else                                                                'Load existing station data
            SQLQuery = "SELECT * "
            SQLQuery = SQLQuery + " FROM ust"
            SQLQuery = SQLQuery + " WHERE (ust.ustCode = " & imUstCode & ")"
            
            Set rst = gSQLSelectCall(SQLQuery)
            If rst.EOF Then
                gMsgBox "No matching records were found", vbOKOnly
                ClearControls
            Else
                BindControls
            End If
            IsOptionsDirty = True
        End If
    End If
    Screen.MousePointer = vbDefault
    imInChg = False
    Exit Sub
ErrHand:
    'ttp 5217
    Screen.MousePointer = vbDefault
    gHandleError LOGFILE, FORMNAME & "-cboSelect_Change"
End Sub

Private Sub cboSalesSource_Click()
    cboSalesSource_Change
End Sub

Private Sub cboSalesSource_GotFocus()
    cboSalesSource.ZOrder
End Sub

Private Sub cboSalesSource_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cbocboSalesSource_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboSalesSource.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub


Private Sub cboSelect_Click()
    cboSelect_Change
End Sub

Private Sub cboSelect_GotFocus()
    cboSelect.ZOrder
End Sub

Private Sub cboSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboSelect.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim ilRet As Integer
    
    ilRet = gPopSalesPeopleInfo()
    Unload Me
End Sub

Private Sub cmdCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bmInCancel = True
End Sub

Private Sub cmdDone_Click()
    Dim iRet As Integer
    
    Screen.MousePointer = vbHourglass
    iRet = mSave(True)
    iRet = gPopSalesPeopleInfo()
    Screen.MousePointer = vbDefault
    Unload frmOptions
    
    Exit Sub
End Sub

Private Sub cmdErase_Click()
If InStr(1, cmdErase.Caption, "Erase") > 0 Then
    txtPW.Text = ""
    mChangeOccured
Else 'change password
    bgShowCurrentPassword = True
    sgPassUserName = Trim(txtSignName.Text)
    AffNewPW.Show vbModal
    If igExitAff = False Then
        txtPW.Text = sgPasswordPasser
    End If
    'show that name is not enabled: changes user made have not been saved.
    cboSelect.SelLength = 0
End If
End Sub

Private Sub cmdSave_Click()
    Dim iRet As Integer
    Dim iIndex As Integer
    
    Screen.MousePointer = vbHourglass
    iRet = mSave(False)
    If iRet = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If IsOptionsDirty = False Then
        cboSelect.Clear
        'If (StrComp(sgUserName, "Counterpoint", 1) = 0) Or (StrComp(sgUserName, "Guide", 1) = 0) Then
        '    cboUser.AddItem "New, -1", 0
        '    SQLQuery = "SELECT ustName, ustCode FROM ust ORDER BY ustName"
        '    Set rst = gSQLSelectCall(SQLQuery)
        'Else
        '    SQLQuery = "SELECT ustName, ustCode FROM ust WHERE (ustName = '" & sgUserName & "')"
        '    Set rst = gSQLSelectCall(SQLQuery)
        'End If
        'While Not rst.EOF
        '    cboUser.AddItem "" & Trim$(rst!ustName) & ", " & rst!ustCode & """"
        '    rst.MoveNext
        'Wend
        If (StrComp(sgUserName, "Counterpoint", 1) = 0) Or (StrComp(sgUserName, "Guide", 1) = 0) Then
            'cboSelect.AddItem "[New]"
            'cboSelect.ItemData(cboSelect.NewIndex) = 0
            SQLQuery = "SELECT ustName, ustCode FROM ust ORDER BY ustName"
            Set rst = gSQLSelectCall(SQLQuery)
        Else
            SQLQuery = "SELECT ustName, ustCode FROM ust WHERE (ustName = '" & sgUserName & "')"
            Set rst = gSQLSelectCall(SQLQuery)
        End If
        While Not rst.EOF
            cboSelect.AddItem Trim$(rst!ustname)
            cboSelect.ItemData(cboSelect.NewIndex) = rst!ustCode
            rst.MoveNext
        Wend
        If (StrComp(sgUserName, "Counterpoint", 1) = 0) Or (StrComp(sgUserName, "Guide", 1) = 0) Then
            cboSelect.AddItem "[New]", 0
            cboSelect.ItemData(cboSelect.NewIndex) = 0
            imInChg = True
            cboSelect.SelText = "[New]"
            imInChg = False
        End If
        ClearControls
    End If

    Screen.MousePointer = vbDefault
End Sub

Private Sub cbcDepartment_DblClick()
    Dim ilDnt As Integer
    sgDepartmentName = cbcDepartment.Text
    frmDepartment.Show vbModal
    mPopDepartment
    If igDepartmentReturn Then
        For ilDnt = 0 To cbcDepartment.ListCount - 1 Step 1
            If cbcDepartment.GetItemData(ilDnt) = igDepartmentReturnCode Then
                cbcDepartment.SetListIndex = ilDnt
                Exit Sub
            End If
        Next ilDnt
    End If
    cbcDepartment.SetListIndex = 1
End Sub

Private Sub edcCity_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcEMail_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcInitials_Change()
    If Not bmIgnoreChange Then
          mChangeOccured
    End If
End Sub

Private Sub edcInitials_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcPhone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
    If imFirstTime Then
        bgUserVisible = True

    '    cboUser.Columns(0).Width = cboUser.Width
    '    cboUser.DroppedDown = True
    '    cboUser.DroppedDown = False
    '    If (StrComp(sgUserName, "Counterpoint", 1) <> 0) And (StrComp(sgUserName, "Guide", 1) <> 0) Then
    '        cboUser.MoveFirst
    '        cboUser_Click
    '    End If
        'If (Not igDemoMode) And (Len(sgSpecialPassword) <> 4) Then
        '    lacWin(15).Visible = False
        'Else
            lacWin(5).Width = lacWin(15).Width
            lacWin(7).Width = lacWin(15).Width
            lacWin(7).Left = lacWin(2).Left + lacWin(2).Width - lacWin(7).Width
        'End If
        imFirstTime = False
    End If
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.4
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    cmdDone.Top = TabStrip1.Top + TabStrip1.Height + 180
    cmdSave.Top = cmdDone.Top
    cmdCancel.Top = cmdDone.Top
    cmdErase.Left = txtPW.Left
    cmdErase.Top = txtPW.Top
    
    gSetFonts frmOptions
    gCenterForm frmOptions
    cbcDepartment.ReSizeFont = "A"
    cbcDepartment.SetDropDownWidth cbcDepartment.Width
    cbcDepartment.PopUpListDirection "A"
    '8273
'    frcAlerts.Top = 90
'    frcAlerts.Left = frcGeneral.Left
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iIndex As Integer
    Dim ilRet As Integer
    On Error GoTo ErrHand

    Screen.MousePointer = vbHourglass
    frmOptions.Caption = "Options - " & sgClientName
    ilRet = mOpenCEFFile()
    
    bmInCancel = False
    imTabIndex = 1
    imBSMode = False
    imInChg = False
    mPopDepartment
    cboSelect.Clear
    ilRet = mFillSalesSource()
    
    If (StrComp(sgUserName, "Counterpoint", 1) = 0) Or (StrComp(sgUserName, "Guide", 1) = 0) Then
        'cboSelect.AddItem "[New]"
        'cboSelect.ItemData(cboSelect.NewIndex) = 0
        SQLQuery = "SELECT ustName, ustCode FROM ust ORDER BY ustName"
        Set rst = gSQLSelectCall(SQLQuery)
      '  txtCurrPW.PasswordChar = ""
    Else
        SQLQuery = "SELECT ustName, ustCode FROM ust WHERE (ustName = '" & sgUserName & "')"
        Set rst = gSQLSelectCall(SQLQuery)
       ' txtCurrPW.PasswordChar = "*"
      ' Dan M allow user to change password without dipslaying
        mSetForGeneralUser
    End If
   While Not rst.EOF
        cboSelect.AddItem Trim$(rst!ustname)
        cboSelect.ItemData(cboSelect.NewIndex) = rst!ustCode
        rst.MoveNext
    Wend
    If (StrComp(sgUserName, "Counterpoint", 1) = 0) Or (StrComp(sgUserName, "Guide", 1) = 0) Then
        cboSelect.AddItem "[New]", 0
        cboSelect.ItemData(cboSelect.NewIndex) = 0
        txtPW.Visible = True
    Else    'Dan M user cannot make himself dormant
        frcTab(1).Enabled = False
        frcTab(2).Enabled = False
        optState(0).Enabled = False
        optState(1).Enabled = False
    End If
    ' Dan M 4/14/09 select 'new' or user's name.
    cboSelect.ListIndex = 0
    imFirstTime = True
    '8273
    lacAlerts(5).Visible = gAllowVendorAlerts(False)
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    'ttp 5217
    Screen.MousePointer = vbDefault
    gHandleError LOGFILE, FORMNAME & "-mLoad"
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    TabStrip1.Left = frcSelect.Left
    TabStrip1.Height = cmdCancel.Top - (frcSelect.Top + frcSelect.Height + 300)  'TabStrip1.ClientTop - TabStrip1.Top + (10 * frcTab(0).Height) / 9
    'TabStrip1.Width = frcSelect.Width
    frcTab(0).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    frcTab(1).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    frcTab(2).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    'frcTab(3).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    frcTab(0).BorderStyle = 0
    frcTab(1).BorderStyle = 0
    frcTab(2).BorderStyle = 0
    'frcTab(3).BorderStyle = 0
    cbcDepartment.Height = edcEMail.Height
    lacDepartment.Top = cbcDepartment.Top + 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    bgUserVisible = False

    ilRet = mCloseCEFFile()
    Set frmOptions = Nothing
End Sub


Private Sub lacAlerts_Click(Index As Integer)
    If bgRemoteExport Then
        lacAlerts(Index).BackColor = vbRed
        Exit Sub
    End If
    If lacAlerts(Index).BackColor = vbGreen Then
        lacAlerts(Index).BackColor = vbRed
    Else
        lacAlerts(Index).BackColor = vbGreen
    End If
    mChangeOccured
End Sub



Private Sub lacComment_Click(Index As Integer)
    If lacComment(Index).BackColor = vbGreen Then
        lacComment(Index).BackColor = vbRed
    Else
        lacComment(Index).BackColor = vbGreen
    End If
    mChangeOccured
End Sub

Private Sub lacGeneral_Click(Index As Integer)
    If bgRemoteExport Then
        lacGeneral(Index).BackColor = vbRed
        Exit Sub
    End If
    If Index = 0 Then
        If lacGeneral(Index).BackColor = vbYellow Then
            lacGeneral(Index).BackColor = vbRed
        Else
            lacGeneral(Index).BackColor = vbYellow
        End If
    Else
        If lacGeneral(Index).BackColor = vbGreen Then
            lacGeneral(Index).BackColor = vbRed
        Else
            lacGeneral(Index).BackColor = vbGreen
        End If
    End If
End Sub

Private Sub lacRestriction_Click(Index As Integer)
    If bgRemoteExport Then
        lacRestriction(Index).BackColor = vbRed
        Exit Sub
    End If
    If lacRestriction(Index).BackColor = vbGreen Then
        lacRestriction(Index).BackColor = vbRed
    Else
        lacRestriction(Index).BackColor = vbGreen
    End If
    mChangeOccured
End Sub

Private Sub lacUtils_Click(Index As Integer)
    If bgRemoteExport Then
        lacUtils(Index).BackColor = vbRed
        Exit Sub
    End If
    If lacUtils(Index).BackColor = vbGreen Then
        lacUtils(Index).BackColor = vbRed
    Else
        lacUtils(Index).BackColor = vbGreen
    End If
    mChangeOccured
End Sub
Private Sub lacWin_Click(Index As Integer)
    If bgRemoteExport Then
        If Index = 1 Or Index = 9 Or Index = 10 Then
            If lacWin(Index).BackColor = vbGreen Then
                lacWin(Index).BackColor = vbRed
            ElseIf lacWin(Index).BackColor = vbRed Then
                If (Index = 0) Or (Index = 14) Then
                    lacWin(Index).BackColor = vbGreen
                Else
                    lacWin(Index).BackColor = vbYellow
                End If
            Else
                lacWin(Index).BackColor = vbGreen
            End If
        Else
            lacWin(Index).BackColor = vbRed
        End If
    Else
        If (Not bgNoRADAR) Or (Index <> 12) Then
            If lacWin(Index).BackColor = vbGreen Then
                lacWin(Index).BackColor = vbRed
            ElseIf lacWin(Index).BackColor = vbRed Then
                If (Index = 0) Or (Index = 14) Then
                    lacWin(Index).BackColor = vbGreen
                Else
                    lacWin(Index).BackColor = vbYellow
                End If
            Else
                lacWin(Index).BackColor = vbGreen
            End If
        End If
    End If
    mChangeOccured
End Sub

Private Sub optState_Click(Index As Integer)
If Not bmIgnoreChange Then
    mChangeOccured
End If

End Sub

Private Sub TabStrip1_Click()
    If imTabIndex = TabStrip1.SelectedItem.Index Then
        Exit Sub
    End If
    frcTab(TabStrip1.SelectedItem.Index - 1).Visible = True
    frcTab(imTabIndex - 1).Visible = False
    imTabIndex = TabStrip1.SelectedItem.Index
End Sub

'Private Sub txtCurrPW_GotFocus()
'    gCtrlGotFocus ActiveControl
'End Sub

Private Sub txtPW_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtRptName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtSignName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

'Private Sub txtVerPW_GotFocus()
'    gCtrlGotFocus ActiveControl
'End Sub

Private Function mSave(iAsk As Integer) As Integer
    Dim iState As Integer
    Dim iIndex As Integer
    Dim sName As String
    Dim sRptName As String
    Dim sCurrPW As String
    Dim sNewPW As String
    Dim sVerPW As String
    ReDim sWin(0 To 14) As String * 1
    Dim sPledge As String * 1
    Dim CurDate As String
    Dim i As Integer
    Dim iLoop As Integer
    Dim sExptSpot As String
    Dim sExptISCI As String
    Dim sTrafLog As String
    Dim slPhone As String
    Dim slCity As String
    Dim slAllowedBlock As String
    Dim slUstClear As String
    Dim slUstDelete As String
    Dim ilLoop As Integer
    Dim ilDntCode As Integer
    Dim slAllowCmmtChg As String
    Dim slAllowCmmtDelete As String
    Dim slActivityLog As String
    Dim slChgExptPriority As String
    Dim slExptSpec As String
    Dim slChgRptPriority As String
    '8272
    Dim slVendorAlert As String
    
    On Error GoTo ErrHand
        
    mSave = False
    
    CurDate = Format(gNow(), sgShowDateForm)

    sName = Trim$(txtSignName.Text)
    If sName = "" Then
        If Not iAsk Then    '"Not iAsk" is Save button
            gMsgBox "Sign on Name must be Defined.", vbOKOnly
        End If
        Exit Function
    End If
    SQLQuery = "SELECT ustCode FROM ust Where (upper(ustName) = '" & UCase(sName) & "')"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsOptionsDirty = False Then
        If Not rst.EOF Then
            gMsgBox "Name Previously Defined.", vbOKOnly
            Exit Function
        End If
    Else
        If Not rst.EOF Then
            If rst!ustCode <> imUstCode Then
                gMsgBox "Name Previously Defined.", vbOKOnly
                Exit Function
            End If
        End If
    End If
    sRptName = Trim$(txtRptName.Text)
    sNewPW = Trim$(txtPW.Text)
   ' sCurrPW = Trim$(txtCurrPW.text)
   ' sNewPW = Trim$(txtNewPW.text)
   ' sVerPW = Trim$(txtVerPW.text)
    slPhone = Trim$(edcPhone.Text)
    slCity = Trim$(edcCity.Text)
    smEmail = Trim$(edcEMail.Text)
    lmEMailCefCode = mPutCefComment(lmEMailCefCode, smEmail)
    For iLoop = 0 To 12 Step 1
        If lacWin(iLoop).BackColor = vbGreen Then
            sWin(iLoop) = "I"
        ElseIf lacWin(iLoop).BackColor = vbYellow Then
            sWin(iLoop) = "V"
        Else
            sWin(iLoop) = "H"
        End If
    Next iLoop
    'Post By: replaced Affiliate A/E
    If lacWin(14).BackColor = vbGreen Then
        sWin(13) = "I"
    ElseIf lacWin(14).BackColor = vbYellow Then
        sWin(13) = "V"
    Else
        sWin(13) = "H"
    End If

    'Export
    If lacWin(15).BackColor = vbGreen Then
        sWin(14) = "I"
    ElseIf lacWin(15).BackColor = vbYellow Then
        sWin(14) = "V"
    Else
        sWin(14) = "H"
    End If

    If lacRestriction(0).BackColor = vbGreen Then
        sPledge = "Y"
    Else
        sPledge = "N"
    End If
    If lacGeneral(0).BackColor = vbYellow Then
        slActivityLog = "V"
    Else
        slActivityLog = "H"
    End If
    If lacGeneral(1).BackColor = vbGreen Then
        slChgExptPriority = "Y"
    Else
        slChgExptPriority = "N"
    End If
    If lacGeneral(2).BackColor = vbGreen Then
        slExptSpec = "Y"
    Else
        slExptSpec = "N"
    End If
    If lacAlerts(1).BackColor = vbGreen Then
        sExptSpot = "Y"
    Else
        sExptSpot = "N"
    End If
    If lacAlerts(2).BackColor = vbGreen Then
        sExptISCI = "Y"
    Else
        sExptISCI = "N"
    End If
    If lacAlerts(3).BackColor = vbGreen Then
        sTrafLog = "Y"
    Else
        sTrafLog = "N"
    End If
    If lacAlerts(4).BackColor = vbGreen Then
        slAllowedBlock = "Y"
    Else
        slAllowedBlock = "N"
    End If
    '8273
    If gAllowVendorAlerts(False) Then
        If lacAlerts(5).BackColor = vbGreen Then
            slVendorAlert = "Y"
        Else
            slVendorAlert = "N"
        End If
    Else
        slVendorAlert = ""
    End If
    If lacUtils(0).BackColor = vbGreen Then
        slUstClear = "Y"
    Else
        slUstClear = "N"
    End If
    
    If lacUtils(1).BackColor = vbGreen Then
        slUstDelete = "Y"
    Else
        slUstDelete = "N"
    End If
    
    slChgRptPriority = "N"
    
    ilDntCode = 0
    imDntIndex = cbcDepartment.ListIndex
    If imDntIndex > 1 Then
        ilDntCode = cbcDepartment.GetItemData(imDntIndex)
    End If
    
    If lacComment(0).BackColor = vbGreen Then
        slAllowCmmtChg = "Y"
    Else
        slAllowCmmtChg = "N"
    End If
    If lacComment(1).BackColor = vbGreen Then
        slAllowCmmtDelete = "Y"
    Else
        slAllowCmmtDelete = "N"
    End If
    
 ' passwords can now be blank
'    If IsOptionsDirty = False Then
'        If sNewPW = "" Then
'            If Not iAsk Then    '"Not iAsk" is Save button
'                gMsgBox "New Password must be Defined.", vbOKOnly
'            End If
'            Exit Function
'        End If
'        If StrComp(sNewPW, sVerPW, 1) <> 0 Then
'            If Not iAsk Then    '"Not iAsk" is Save button
'                gMsgBox "New and Verify Passwords must Match.", vbOKOnly
'            End If
'            Exit Function
'        End If
'    End If
    If bgStrongPassword And Not gStrongPassword(sNewPW) And Not LenB(sNewPW) = 0 Then  'allow blank password
        '5608  changed to strong password and this password is not strong (should only happen if guide is making changes)? then blank password
        sNewPW = ""
       ' txtPW.Text = smPassword
        'ttp 5608 various strong password errors
       ' txtPW.SetFocus
        'Exit Function
    End If
    mSave = True
    'Determine state of rep (active or dormant)
    iState = -1
    For i = 0 To 1
        If optState(i).Value Then
            iState = i
            Exit For
        End If
    Next i
    
    'Add new user
    If IsOptionsDirty = False Then
        SQLQuery = "INSERT INTO ust(ustName, ustReportName, ustPassword, "
        SQLQuery = SQLQuery & "ustState, ustPassDate, ustWin1, "
        SQLQuery = SQLQuery & "ustWin2, ustWin3, ustWin4, "
        SQLQuery = SQLQuery & "ustWin5, ustWin6, ustWin7, "
        SQLQuery = SQLQuery & "ustWin8, ustWin9, ustPledge, "
        SQLQuery = SQLQuery & "ustExptSpotAlert, ustExptISCIAlert, ustTrafLogAlert, "
        SQLQuery = SQLQuery & "ustWin10, ustWin11, ustWin12, ustWin13, "
        SQLQuery = SQLQuery & "ustWin14, ustWin15, ustPhoneNo, ustCity, ustEMailCefCode, ustAllowedToBlock,"
        SQLQuery = SQLQuery & "ustWin16, "
        SQLQuery = SQLQuery & "ustUserInitials, "
        SQLQuery = SQLQuery & "ustDntCode, "
        SQLQuery = SQLQuery & "ustAllowCmmtChg, "
        SQLQuery = SQLQuery & "ustAllowCmmtDelete, "
        SQLQuery = SQLQuery & "ustActivityLog, "
        SQLQuery = SQLQuery & "ustWin17, "
        SQLQuery = SQLQuery & "ustChgExptPriority, "
        SQLQuery = SQLQuery & "ustExptSpec, "
        SQLQuery = SQLQuery & "ustChgRptPriority, "
        SQLQuery = SQLQuery & "ustSSMnfCode, "
        '8273
        SQLQuery = SQLQuery & "ustVendoralert, "
        SQLQuery = SQLQuery & "ustUnused "
        SQLQuery = SQLQuery & ") "
        SQLQuery = SQLQuery & "VALUES ('" & gFixQuote(sName) & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(sRptName) & "', '" & sNewPW & "', "
        SQLQuery = SQLQuery & iState & ", '" & Format$(CurDate, sgSQLDateForm) & "', '" & sWin(1) & "', "
        SQLQuery = SQLQuery & "'" & sWin(2) & "', '" & sWin(3) & "', '" & sWin(4) & "', "
        SQLQuery = SQLQuery & "'" & sWin(5) & "', '" & sWin(6) & "', '" & sWin(7) & "', "
        SQLQuery = SQLQuery & "'" & sWin(8) & "', '" & sWin(9) & "', '" & sPledge & "', "
        SQLQuery = SQLQuery & "'" & sExptSpot & "', '" & sExptISCI & "', '" & sTrafLog & "', "
        SQLQuery = SQLQuery & "'" & sWin(10) & "', '" & sWin(11) & "', '" & sWin(12) & "', '" & sWin(13) & "', "
        SQLQuery = SQLQuery & "'" & slUstClear & "', '" & slUstDelete & "', "
        SQLQuery = SQLQuery & "'" & slPhone & "', '" & slCity & "', " & lmEMailCefCode & ", '" & slAllowedBlock & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(sWin(0)) & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(edcInitials.Text) & "', "
        SQLQuery = SQLQuery & ilDntCode & ", "
        SQLQuery = SQLQuery & "'" & gFixQuote(slAllowCmmtChg) & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(slAllowCmmtDelete) & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(slActivityLog) & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(sWin(14)) & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(slChgExptPriority) & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(slExptSpec) & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(slChgRptPriority) & "', "
        SQLQuery = SQLQuery & imSSCode & ", "
        '8273
        SQLQuery = SQLQuery & "'" & slVendorAlert & "', "
        SQLQuery = SQLQuery & "'" & "" & "' "
        SQLQuery = SQLQuery & ") "

        If iAsk Then
            If gMsgBox("Save new user?", vbYesNo) = vbYes Then
                cnn.BeginTrans
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError LOGFILE, FORMNAME & "-mSave"
                    cnn.RollbackTrans
                    mSave = False
                    Exit Function
                End If
                cnn.CommitTrans
            End If
        Else
            cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError LOGFILE, FORMNAME & "-mSave"
                cnn.RollbackTrans
                mSave = False
                Exit Function
            End If
            cnn.CommitTrans
        End If
           
    Else
        'UPDATE existing rep
        SQLQuery = "UPDATE ust"
        SQLQuery = SQLQuery & " SET ustName = '" & gFixQuote(sName) & "', "
        SQLQuery = SQLQuery & "ustReportName = '" & gFixQuote(sRptName) & "', "
        'Dan M won't allow guide to change password
       ' If (StrComp(smPassword, sCurrPW, 1) = 0) Then
           ' If StrComp(sNewPW, sVerPW, 1) = 0 Then
        SQLQuery = SQLQuery & "ustPassword = '" & sNewPW & "', "
        SQLQuery = SQLQuery & "ustPassDate = '" & Format$(CurDate, sgSQLDateForm) & "', "
           ' End If
       ' End If
        SQLQuery = SQLQuery & "ustState = " & iState & ", "
        SQLQuery = SQLQuery & "ustWin1 = '" & sWin(1) & "', "
        SQLQuery = SQLQuery & "ustWin2 = '" & sWin(2) & "', "
        SQLQuery = SQLQuery & "ustWin3 = '" & sWin(3) & "', "
        SQLQuery = SQLQuery & "ustWin4 = '" & sWin(4) & "', "
        SQLQuery = SQLQuery & "ustWin5 = '" & sWin(5) & "', "
        SQLQuery = SQLQuery & "ustWin6 = '" & sWin(6) & "', "
        SQLQuery = SQLQuery & "ustWin7 = '" & sWin(7) & "', "
        SQLQuery = SQLQuery & "ustWin8 = '" & sWin(8) & "', "
        SQLQuery = SQLQuery & "ustWin9 = '" & sWin(9) & "', "
        SQLQuery = SQLQuery & "ustPledge = '" & sPledge & "', "
        SQLQuery = SQLQuery & "ustExptSpotAlert = '" & sExptSpot & "', "
        SQLQuery = SQLQuery & "ustExptISCIAlert = '" & sExptISCI & "', "
        SQLQuery = SQLQuery & "ustTrafLogAlert = '" & sTrafLog & "', "
        SQLQuery = SQLQuery & "ustWin10 = '" & sWin(10) & "', "
        SQLQuery = SQLQuery & "ustWin11 = '" & sWin(11) & "', "
        SQLQuery = SQLQuery & "ustWin12 = '" & sWin(12) & "', "
        SQLQuery = SQLQuery & "ustWin13 = '" & sWin(13) & "', "
        SQLQuery = SQLQuery & "ustWin14 = '" & slUstClear & "', "
        SQLQuery = SQLQuery & "ustWin15 = '" & slUstDelete & "', "
        SQLQuery = SQLQuery & "ustEMailCefCode = " & lmEMailCefCode & ", "
        SQLQuery = SQLQuery & "ustPhoneNo = '" & slPhone & "', "
        SQLQuery = SQLQuery & "ustCity = '" & slCity & "', "
        SQLQuery = SQLQuery & "ustAllowedToBlock = '" & slAllowedBlock & "', "
        SQLQuery = SQLQuery & "ustWin16 = '" & gFixQuote(sWin(0)) & "', "
        SQLQuery = SQLQuery & "ustUserInitials = '" & gFixQuote(edcInitials.Text) & "', "
        SQLQuery = SQLQuery & "ustDntCode = " & ilDntCode & ", "
        SQLQuery = SQLQuery & "ustAllowCmmtChg = '" & gFixQuote(slAllowCmmtChg) & "', "
        SQLQuery = SQLQuery & "ustAllowCmmtDelete = '" & gFixQuote(slAllowCmmtDelete) & "', "
        SQLQuery = SQLQuery & "ustActivityLog = '" & slActivityLog & "', "
        SQLQuery = SQLQuery & "ustWin17 = '" & sWin(14) & "', "
        SQLQuery = SQLQuery & "ustChgExptPriority = '" & slChgExptPriority & "', "
        SQLQuery = SQLQuery & "ustExptSpec = '" & slExptSpec & "', "
        SQLQuery = SQLQuery & "ustSSMnfCode = " & imSSCode
        '8273
        SQLQuery = SQLQuery & ", ustVendorAlert = '" & slVendorAlert & "'"
        SQLQuery = SQLQuery & " WHERE (ustCode = " & imUstCode & ")"
        'Dan M 4/4/09 changes to users other than current are being lost.
        If iAsk Then
            If bmChangeMade Then
                If gMsgBox("Save changes to " & sName & " ?", vbYesNo) = vbYes Then
                    cnn.BeginTrans
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError LOGFILE, FORMNAME & "-mSave"
                        cnn.RollbackTrans
                        mSave = False
                        Exit Function
                    End If
                    cnn.CommitTrans
                End If
            End If
        Else
            cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError LOGFILE, FORMNAME & "-mSave"
                cnn.RollbackTrans
                mSave = False
                Exit Function
            End If
            cnn.CommitTrans
        End If
        
'        If iAsk Then
'            If gMsgBox("Save all changes?", vbYesNo) = vbYes Then
'                cnn.BeginTrans
'                'cnn.Execute SQLQuery, rdExecDirect
'                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                    GoSub ErrHand:
'                End If
'                cnn.CommitTrans
'            End If
'        Else
'            cnn.BeginTrans
'            'cnn.Execute SQLQuery, rdExecDirect
'            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                GoSub ErrHand:
'            End If
'            cnn.CommitTrans
'        End If
    End If
    If igUstCode = imUstCode Then
            sgUserName = sName
            sgReportName = sRptName
            sgPhoneNo = slPhone
            sgCity = slCity
            sgEMail = smEmail
    'Dan M 1/12/10 no one can change these setting but guide (who cannot change himself),  so all this is unnecessary
'            For ilLoop = 1 To 13 Step 1
'                sgUstWin(ilLoop) = sWin(ilLoop)
'            Next ilLoop
'            sgUstClear = slUstClear
'            sgUstDelete = slUstDelete
'            sgUstPledge = sPledge
'            sgExptSpotAlert = sExptSpot
'            sgExptISCIAlert = sExptISCI
'            sgTrafLogAlert = sTrafLog
'            sgAllowedToBlock = slAllowedBlock
   End If
    'Dan M staying on form
    If Not iAsk Then
        bmChangeMade = False
        cboSelect.Enabled = True
        cboSelect_Change
    End If
    Exit Function

ErrHand:
    'ttp 5217
    Screen.MousePointer = vbDefault
    gHandleError LOGFILE, FORMNAME & "-mSave"
    mSave = False
End Function
Private Sub mSetForLimitedGuide()
    txtPW.Visible = False
    If StrComp(txtSignName.Text, "Guide", vbTextCompare) = 0 Then
        optState(0).Enabled = False
        optState(1).Enabled = False
        frcTab(1).Enabled = False
        frcTab(2).Enabled = False
        mChangeEraseToCancel
    Else
        optState(0).Enabled = True
        optState(1).Enabled = True
        frcTab(1).Enabled = True
        frcTab(2).Enabled = True
        If txtSignName = "" Then
            cmdErase.Visible = False
        Else
            mChangeCancelToErase
        End If
    End If
End Sub
Private Sub mSetForInternalGuide()
'only difference from setforlimitedguide is other user's passwords are visible
    txtPW.Visible = True
    If StrComp(txtSignName.Text, "Guide", vbTextCompare) = 0 Then
        optState(0).Enabled = False
        optState(1).Enabled = False
        frcTab(1).Enabled = False
        frcTab(2).Enabled = False
    Else
        optState(0).Enabled = True
        optState(1).Enabled = True
        frcTab(1).Enabled = True
        frcTab(2).Enabled = True
        If txtSignName = "" Then
           txtPW.Visible = False
        End If
    End If

End Sub
Private Sub mSetForGeneralUser()
    txtPW.Visible = False
    mChangeEraseToCancel
End Sub
Private Sub mChangeEraseToCancel()

    cmdErase.Caption = "Change Password"
    cmdErase.Visible = True

End Sub
Private Sub mChangeCancelToErase()
    cmdErase.Caption = "Erase Password"
    cmdErase.Visible = True

End Sub
Private Sub mChangeOccured()
    bmChangeMade = True
    cboSelect.Enabled = False
End Sub
'Dan M auto generated _change 5/4/09

Private Sub edcCity_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcEMail_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub edcPhone_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub


Private Sub txtPW_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub txtRptName_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub txtSignName_Change()
If Not bmIgnoreChange Then
      mChangeOccured
End If
End Sub

Private Sub lacEMail_Click()
 mChangeOccured
End Sub

Private Sub lacCity_Click()
 mChangeOccured
End Sub

Private Sub lacPhone_Click()
 mChangeOccured
End Sub

Private Sub mPopDepartment()

    On Error GoTo ErrHand
    cbcDepartment.Clear
    cbcDepartment.AddItem ("[New]")
    cbcDepartment.SetItemData = -1
    cbcDepartment.AddItem ("[None]")
    cbcDepartment.SetItemData = 0
    SQLQuery = "SELECT * FROM dnt ORDER BY dntName"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        cbcDepartment.AddItem Trim$(rst!dntName)
        cbcDepartment.SetItemData = rst!dntCode
        rst.MoveNext
    Loop
    
    Exit Sub
ErrHand:
    'ttp 5217
    Screen.MousePointer = vbDefault
    gHandleError LOGFILE, FORMNAME & "-mPopDepartment"
End Sub

Private Function mFillSalesSource() As Boolean

    'D.S. 06/11/14
    Dim ilIdx As Integer
    Dim rstSalesSource As ADODB.Recordset
    
    On Error GoTo ErrHand
    mFillSalesSource = False
    
    SQLQuery = "SELECT DISTINCT MnfCode, mnfName"
    SQLQuery = SQLQuery + " From MNF_Multi_Names"
    SQLQuery = SQLQuery + " WHERE mnfType = " & "'" & "S" & "'"
    SQLQuery = SQLQuery + " ORDER BY mnfName"
    Set rstSalesSource = gSQLSelectCall(SQLQuery)
    cboSalesSource.Clear
    cboSalesSource.AddItem ("[All Sales Sources]"), 0
    cboSalesSource.ItemData(cboSalesSource.NewIndex) = 0
    While Not rstSalesSource.EOF
        cboSalesSource.AddItem Trim$(rstSalesSource!mnfName)
        cboSalesSource.ItemData(cboSalesSource.NewIndex) = rstSalesSource!mnfCode
        rstSalesSource.MoveNext
    Wend
    cboSalesSource.ListIndex = 0
    mFillSalesSource = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError LOGFILE, FORMNAME & "-mFillSalesSource"
    Exit Function
End Function
