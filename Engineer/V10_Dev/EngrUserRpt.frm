VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form EngrUserRpt 
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   7065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7065
   Begin VB.Frame frcOption 
      Caption         =   "Report Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   240
      TabIndex        =   6
      Top             =   1860
      Width           =   6495
      Begin VB.CheckBox ckcInclOther 
         Caption         =   "Include Audio Sources"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   255
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.TextBox edcTo 
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox edcFrom 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame frcOldNew 
         Caption         =   "Show"
         Height          =   615
         Left            =   2415
         TabIndex        =   10
         Top             =   1350
         Visible         =   0   'False
         Width           =   2535
         Begin VB.OptionButton optOldNew 
            Caption         =   "History"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optOldNew 
            Caption         =   "Current"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Label lacChangeTo 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lacChangeDates 
         Caption         =   "Enter change dates- From"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3240
      Top             =   960
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   4290
      FormDesignWidth =   7065
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4455
      TabIndex        =   9
      Top             =   1200
      Width           =   1920
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4275
      TabIndex        =   8
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4050
      TabIndex        =   7
      Top             =   240
      Width           =   2685
   End
   Begin VB.Frame frcOutput 
      Caption         =   "Report Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1065
         TabIndex        =   4
         Top             =   690
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   720
         Width           =   870
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   480
         Width           =   2190
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2310
      End
   End
End
Attribute VB_Name = "EngrUserRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  EngrUserRpt - a current/history report of List  tables:
'   Audio Source, Audio Name, Audio Type, Bus Definitions,
'   Bus Groups, Comments, Control Names, Follow, Material Types,
'   Netcue, Relay, Silence, Time Types, User
'
'
'
'*
'*  Created September,  2004
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private Sub cmdDone_Click()
    Unload EngrUserRpt
End Sub

Private Sub cmdReport_Click()
    Dim iType As Integer
    Dim sOutput As String
    Dim ilRet As Integer
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim slRptName As String
    Dim slExportName As String
    Dim SQLQuery As String
    Dim ilListIndex As Integer
    Dim slSQLFromDate As String
    Dim slSQLToDAte As String
    Dim slDate As String
    Dim ilLoop As Integer
  
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass


    If optRptDest(0).Value = True Then
       ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        ilRptDest = 2
        ilExportType = cboFileType.ListIndex
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    
    If optOldNew(0).Value = True Then       'current only
        sgCrystlFormula1 = "'N'"
        slSQLFromDate = "1/1/1970"
        slSQLToDAte = "12/31/2069"
    Else
        sgCrystlFormula1 = "'Y'"            'show history
        slDate = edcFrom.text
        If edcFrom.text = "" Then
            slSQLFromDate = "1/1/1970"
        Else
            If Not gIsDate(slDate) Then
                Beep
                MsgBox "Invalid From Date"
                edcFrom.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            slSQLFromDate = slDate
        End If
   
        sgCrystlFormula2 = "Date(" + Format$(slSQLFromDate, "yyyy") + "," + Format$(slSQLFromDate, "mm") + "," + Format$(slSQLFromDate, "dd") + ")"

        slDate = edcTo.text
        If edcTo.text = "" Then
            slSQLToDAte = "12/31/2069"
        Else
            If Not gIsDate(slDate) Then
                Beep
                MsgBox "Invalid To Date"
                edcTo.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            slSQLToDAte = slDate
        End If

        sgCrystlFormula3 = "Date(" + Format$(slSQLToDAte, "yyyy") + "," + Format$(slSQLToDAte, "mm") + "," + Format$(slSQLToDAte, "dd") + ")"

    End If
    
    gObtainReportforCrystal slRptName, slExportName     'determine which .rpt to call and setup an export name is user selected output to export

    If sgCrystlFormula1 = "'N'" Then        'Get CURRENT ONLY
        slRptName = slRptName & ".rpt"      'concatenate the crystal report name plus extension
        If igRptIndex = USER_RPT Then
            SQLQuery = "Select * from UIE_User_Info where uieCurrent = 'Y' Order by uieSignOnName"
        ElseIf igRptIndex = RELAY_RPT Then
            SQLQuery = "Select * from RNE_Relay_Name , UIE_User_Info where rneCurrent = 'Y' and rneuiecode = uiecode  Order by rneName "
        ElseIf igRptIndex = MATTYPE_RPT Then
            SQLQuery = "Select * from MTE_Material_Type, UIE_User_Info  where mteCurrent = 'Y' AND mteuiecode = uiecode Order by mteName"
         ElseIf igRptIndex = FOLLOW_RPT Then
            SQLQuery = "Select * from FNE_Follow_Name , UIE_User_Info where fneCurrent = 'Y' and fneuiecode = uiecode Order by fneName "
        ElseIf igRptIndex = TIMETYPE_RPT Then
            SQLQuery = "Select * from TTE_Time_Type , UIE_User_Info  where tteCurrent = 'Y' and tteuiecode = uiecode Order by tteType desc, tteName "
        ElseIf igRptIndex = SILENCE_RPT Then
            SQLQuery = "Select * from SCE_Silence_Char , UIE_User_Info where sceCurrent = 'Y' and sceuiecode = uiecode Order by sceAutoChar  "
        ElseIf igRptIndex = AUDIONAME_RPT Then
            SQLQuery = "Select * from ANE_Audio_Name , ATE_Audio_type, UIE_User_Info where aneCurrent = 'Y' and  aneuiecode = uiecode  and aneatecode = atecode Order by aneName "
        ElseIf igRptIndex = AUDIOTYPE_RPT Then
            SQLQuery = "Select * from ATE_Audio_Type ,UIE_User_Info where ateCurrent = 'Y' and ateuiecode = uiecode Order by ateName  "
        ElseIf igRptIndex = AUDIOSOURCE_RPT Then
            SQLQuery = "Select  * from (((((((ase_audio_source  Inner Join ane_audio_name ane_audio_name on aseprianecode = ane_audio_name.anecode) "
            SQLQuery = SQLQuery & " left outer join ane_audio_name ane_buaudio_name on asebkupanecode = ane_buaudio_name.anecode) "
            SQLQuery = SQLQuery & " left outer join ane_audio_name ane_protaudio_name on aseprotanecode = ane_protaudio_name.anecode) "
            SQLQuery = SQLQuery & " left outer join cce_control_char cce_control_char on asepriccecode = cce_control_char.ccecode) "
            SQLQuery = SQLQuery & " left outer join cce_control_char cce_bucontrol_char on asebkupccecode = cce_bucontrol_char.ccecode) "
            SQLQuery = SQLQuery & " left outer join cce_control_char cce_protcontrol_char on aseprotccecode = cce_protcontrol_char.ccecode) "
            SQLQuery = SQLQuery & " inner join uie_user_info on aseuiecode = uiecode) "
            SQLQuery = SQLQuery & " where asecurrent = 'Y'"
          ElseIf igRptIndex = SITE_RPT Then
            SQLQuery = "Select * from SOE_Site_Option, UIE_User_Info where soeCurrent = 'Y' and soeuiecode = uiecode"
        ElseIf igRptIndex = BUSGROUP_RPT Then
            SQLQuery = "Select * from BGE_Bus_Group, UIE_User_Info where bgeCurrent = 'Y' and bgeuiecode = uiecode order by bgename"
        ElseIf igRptIndex = BUS_RPT Then
            SQLQuery = "select * from (((bde_bus_definition inner join uie_user_info on bdeuiecode = uiecode) left outer join cce_control_char on bdeccecode = ccecode) left outer join ase_audio_source Ase on bdeasecode = ase.asecode) left outer join ane_audio_name on ase.aseprianecode = anecode where bdeCurrent = 'Y'"
        ElseIf igRptIndex = NETCUE_RPT Then
            SQLQuery = "Select * from UIE_User_Info,  NNE_Netcue_Name left outer join DNE_Day_Name on nnednecode = dnecode where nneCurrent = 'Y' and nneuiecode = uiecode"
        ElseIf igRptIndex = CONTROL_RPT Then
            SQLQuery = "Select * from CCE_Control_Char ,UIE_User_Info where cceCurrent = 'Y' and cceuiecode = uiecode Order by cceType, cceAutoChar  "
        ElseIf igRptIndex = COMMENT_RPT Then
            SQLQuery = "Select * from CTE_Commts_And_Title ,UIE_User_Info where cteCurrent = 'Y' and cteuiecode = uiecode Order by cteComment  "
        ElseIf igRptIndex = EVENT_RPT Then
            SQLQuery = "Select * from ETE_Event_Type, EPE_Event_Properties, UIE_User_info where "
            SQLQuery = SQLQuery & " epeetecode = etecode and eteuiecode = uiecode and eteCurrent = 'Y'"
            SQLQuery = SQLQuery & " Order by eteName, eteOrigeteCode, eteVersion desc"
        ElseIf igRptIndex = AUTOMATION_RPT Then
            SQLQuery = "Select * from AEE_Auto_Equip,  UIE_User_info where "
            SQLQuery = SQLQuery & " aeeuiecode = uiecode and aeeCurrent = 'Y'"
            SQLQuery = SQLQuery & " Order by aeeName"
        End If
        
    Else       'GET HISTORY ONLY
        slRptName = slRptName & "Hist.rpt"      'all history reports will have same name as non-history reports but with "Hist" appended
        If igRptIndex = USER_RPT Then
            SQLQuery = "Select * from AIE_Active_Info, UIE_User_Info UIE_User_Info, UIE_User_Info UIE_To, UIE_User_Info UIE_From "
            SQLQuery = SQLQuery & " where aietofilecode = uie_To.uiecode and aiefromfilecode = uie_from.uiecode and "
            SQLQuery = SQLQuery & " aieuiecode = UIE_User_Info.uiecode "
            SQLQuery = SQLQuery & " and AIE_Active_Info.aieRefFileName = 'UIE'"
            SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by UIE_User_Info.uieSignOnName"
        ElseIf igRptIndex = RELAY_RPT Then
            SQLQuery = "Select * from AIE_Active_Info, UIE_User_Info UIE_User_Info, "
            SQLQuery = SQLQuery & " RNE_Relay_Name RNE_Relay_Name, RNE_Relay_Name RNE_ToRelay_Name, "
            SQLQuery = SQLQuery & " UIE_User_info UIE_FromUser_Info, UIE_User_info UIE_ToUser_Info "
            SQLQuery = SQLQuery & " where aieuiecode = UIE_User_Info.uiecode and aiefromfilecode = RNE_Relay_Name.rnecode and "
            SQLQuery = SQLQuery & " aietofilecode = RNE_ToRelay_Name.rnecode and RNE_Relay_Name.rneuiecode = UIE_FromUser_Info.uiecode and "
            SQLQuery = SQLQuery & " RNE_ToRelay_Name.rneuiecode = UIE_ToUser_Info.uiecode"
            SQLQuery = SQLQuery & " and AIE_Active_Info.aieRefFileName = 'RNE'"
            SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by RNE_ToRelay_Name.rneName "
        ElseIf igRptIndex = MATTYPE_RPT Then
            SQLQuery = "Select * from AIE_Active_Info, UIE_User_Info UIE_User_Info, "
            SQLQuery = SQLQuery & " MTE_MATERIAL_TYPE MTE_MATERIAL_TYPE, MTE_MATERIAL_TYPE MTE_TO, "
            SQLQuery = SQLQuery & " UIE_User_info UIE_FromUser_Info, UIE_User_info UIE_ToUser_Info "
            SQLQuery = SQLQuery & " where aieuiecode = UIE_User_Info.uiecode and aiefromfilecode = MTE_MATERIAL_TYPE.mtecode and "
            SQLQuery = SQLQuery & " aietofilecode = MTE_to.mtecode and MTE_MATERIAL_TYPE.mteuiecode = UIE_FromUser_Info.uiecode and "
            SQLQuery = SQLQuery & " MTE_to.mteuiecode = UIE_ToUser_Info.uiecode"
            SQLQuery = SQLQuery & " and AIE_Active_Info.aieRefFileName = 'MTE'"
            SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by MTE_to.mteName "
        ElseIf igRptIndex = FOLLOW_RPT Then
            SQLQuery = "Select * from AIE_Active_Info, UIE_User_Info UIE_User_Info, "
            SQLQuery = SQLQuery & " FNE_Follow_Name FNE_Follow_Name, FNE_Follow_Name FNE_To, "
            SQLQuery = SQLQuery & " UIE_User_info UIE_From, UIE_User_info UIE_To "
            SQLQuery = SQLQuery & " where aieuiecode = UIE_User_Info.uiecode and aiefromfilecode = FNE_Follow_Name.fnecode and "
            SQLQuery = SQLQuery & " aietofilecode = FNE_To.fnecode and FNE_Follow_Name.fneuiecode = UIE_From.uiecode and "
            SQLQuery = SQLQuery & " FNE_To.fneuiecode = UIE_To.uiecode"
            SQLQuery = SQLQuery & " and AIE_Active_Info.aieRefFileName = 'FNE'"
            SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by FNE_to.fneName "
        ElseIf igRptIndex = TIMETYPE_RPT Then
            SQLQuery = "Select * from AIE_Active_Info, UIE_User_Info UIE_User_Info, "
            SQLQuery = SQLQuery & " TTE_Time_Type TTE_Time_Type, TTE_Time_Type TTE_ToTime_Type, "
            SQLQuery = SQLQuery & " UIE_User_info UIE_FromUser_Info, UIE_User_info UIE_ToUser_Info "
            SQLQuery = SQLQuery & " where aieuiecode = UIE_User_Info.uiecode and aiefromfilecode = TTE_Time_Type.ttecode and "
            SQLQuery = SQLQuery & " aietofilecode = TTE_ToTime_Type.ttecode and TTE_Time_Type.tteuiecode = UIE_FromUser_Info.uiecode and "
            SQLQuery = SQLQuery & " TTE_ToTime_Type.tteuiecode = UIE_ToUser_Info.uiecode"
            SQLQuery = SQLQuery & " and AIE_Active_Info.aieRefFileName = 'TTE'"
            SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by TTE_ToTime_Type.tteType desc, TTE_ToTime_Type.tteName, TTE_ToTime_Type.tteorigttecode, TTE_ToTime_Type.tteversion desc "
        ElseIf igRptIndex = SILENCE_RPT Then
            SQLQuery = "Select * from AIE_Active_Info, UIE_User_Info UIE_User_Info, "
            SQLQuery = SQLQuery & " SCE_Silence_Char SCE_Silence_Char, SCE_Silence_Char SCE_ToSilence_Char, "
            SQLQuery = SQLQuery & " UIE_User_info UIE_FromUser_Info, UIE_User_info UIE_ToUser_Info "
            SQLQuery = SQLQuery & " where aieuiecode = UIE_User_Info.uiecode and aiefromfilecode = SCE_Silence_Char.scecode and "
            SQLQuery = SQLQuery & " aietofilecode = SCE_ToSilence_Char.scecode and SCE_Silence_Char.sceuiecode = UIE_FromUser_Info.uiecode and "
            SQLQuery = SQLQuery & " SCE_ToSilence_Char.sceuiecode = UIE_ToUser_Info.uiecode"
            SQLQuery = SQLQuery & " and AIE_Active_Info.aieRefFileName = 'SCE'"
            SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by SCE_ToSilence_Char.sceAutoChar "
        ElseIf igRptIndex = AUDIONAME_RPT Then
            SQLQuery = "Select * from AIE_Active_Info, UIE_User_Info UIE_User_Info, "
            SQLQuery = SQLQuery & " ANE_Audio_Name ANE_Audio_Name, ANE_Audio_Name ANE_To, "
            SQLQuery = SQLQuery & " UIE_User_info UIE_From, UIE_User_info UIE_To, "
            SQLQuery = SQLQuery & " ATE_Audio_Type ATE_Audio_Type, ATE_Audio_Type ATE_To "
            SQLQuery = SQLQuery & " where aieuiecode = UIE_User_Info.uiecode and aiefromfilecode = ANE_Audio_Name.anecode and "
            SQLQuery = SQLQuery & " aietofilecode = ANE_To.anecode and ANE_Audio_Name.aneuiecode = UIE_From.uiecode and "
            SQLQuery = SQLQuery & " ANE_To.aneuiecode = UIE_To.uiecode and "
            SQLQuery = SQLQuery & " ANE_Audio_Name.aneatecode = ATE_Audio_Type.atecode and ANE_To.aneatecode = ATE_To.atecode "
            SQLQuery = SQLQuery & " and AIE_Active_Info.aieRefFileName = 'ANE'"
            SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by ANE_To.aneName "
        ElseIf igRptIndex = AUDIOTYPE_RPT Then
            SQLQuery = "Select * from AIE_Active_Info, UIE_User_Info UIE_User_Info, "
            SQLQuery = SQLQuery & " ATE_Audio_Type ATE_Audio_Type, ATE_Audio_Type ATE_To, "
            SQLQuery = SQLQuery & " UIE_User_info UIE_From, UIE_User_info UIE_To "
            SQLQuery = SQLQuery & " where aieuiecode = UIE_User_Info.uiecode and aiefromfilecode = ATE_Audio_Type.atecode and "
            SQLQuery = SQLQuery & " aietofilecode = ATE_To.atecode and ATE_Audio_Type.ateuiecode = UIE_From.uiecode and "
            SQLQuery = SQLQuery & " ATE_To.ateuiecode = UIE_To.uiecode"
            SQLQuery = SQLQuery & " and AIE_Active_Info.aieRefFileName = 'ATE'"
            SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by ATE_To.ateName "
        ElseIf igRptIndex = AUDIOSOURCE_RPT Then
            SQLQuery = "Select  * from  (((((((((((((((((aie_active_info inner join  uie_user_info uie_user_info on aieuiecode = uie_user_info.uiecode) "
            SQLQuery = SQLQuery & " inner join ASE_audio_source ase_audio_source on aiefromfilecode = ase_audio_source.asecode) "
            SQLQuery = SQLQuery & " inner join ane_audio_name ane_audio_name on ase_audio_source.aseprianecode = ane_audio_name.anecode) "
            SQLQuery = SQLQuery & " left outer join cce_control_char cce_control_char on ase_audio_source.asepriccecode = cce_control_char.ccecode) "
            SQLQuery = SQLQuery & " left outer join ane_audio_name ane_buaudio_name on ase_audio_source.asebkupanecode = ane_buaudio_name.anecode) "
            SQLQuery = SQLQuery & " left outer join cce_control_char cce_bucontrol_char on ase_audio_source.asebkupccecode = cce_bucontrol_char.ccecode) "
            SQLQuery = SQLQuery & " left outer join ane_audio_name ane_protaudio_name on ase_audio_source.aseprotanecode = ane_protaudio_name.anecode) "
            SQLQuery = SQLQuery & " left outer join cce_control_char cce_protcontrol_char on ase_audio_source.aseprotccecode = cce_protcontrol_char.ccecode) "
            SQLQuery = SQLQuery & " inner join uie_user_info uie_from on ase_audio_source.aseuiecode = uie_from.uiecode) "
            SQLQuery = SQLQuery & " inner join ASE_audio_source ase_toprim on aietofilecode = ase_toprim.asecode) "
            SQLQuery = SQLQuery & " inner join ane_audio_name ane_topri on ase_toprim.aseprianecode = ane_topri.anecode) "
            SQLQuery = SQLQuery & " left outer join cce_control_char cce_topri on ase_toprim.asepriccecode = cce_topri.ccecode)"
            SQLQuery = SQLQuery & " left outer join ane_audio_name ane_tobu on ase_toprim.asebkupanecode = ane_tobu.anecode) "
            SQLQuery = SQLQuery & " left outer join cce_control_char cce_tobu on ase_toprim.asebkupccecode = cce_tobu.ccecode) "
            SQLQuery = SQLQuery & " left outer join ane_audio_name ane_toprot on ase_toprim.aseprotanecode = ane_toprot.anecode) "
            SQLQuery = SQLQuery & " left outer join cce_control_char cce_toprot on ase_toprim.aseprotccecode = cce_toprot.ccecode) "
            SQLQuery = SQLQuery & " inner join uie_user_info uie_to on ase_toprim.aseuiecode = uie_to.uiecode) "
            SQLQuery = SQLQuery & " where AIE_Active_Info.aieRefFileName = 'ASE'"
            SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by ANE_ToPri.anename, ASE_ToPrim.aseorigasecode, ase_toprim.aseversion desc"
          
        ElseIf igRptIndex = SITE_RPT Then           'no history report for SITE
        ElseIf igRptIndex = BUSGROUP_RPT Then
            SQLQuery = "Select * from AIE_Active_Info, BGE_Bus_Group BGE_Bus_Group, BGE_Bus_Group BGE_To, UIE_User_Info UIE_User_Info, UIE_User_Info UIE_From, UIE_User_Info UIE_To "
            SQLQuery = SQLQuery & " where aieuiecode = UIE_User_Info.uiecode and  "
            SQLQuery = SQLQuery & " aiefromfilecode = bge_Bus_Group.bgecode and "
            SQLQuery = SQLQuery & " aietofilecode = bge_To.bgecode and "
            SQLQuery = SQLQuery & " BGE_Bus_Group.bgeuiecode = UIE_From.uiecode and "
            SQLQuery = SQLQuery & " bge_To.bgeuiecode = UIE_To.uiecode and "
            SQLQuery = SQLQuery & " aie_active_info.aieRefFileName = 'BGE' and "
            SQLQuery = SQLQuery & "(aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by BGE_To.bgeName"
        ElseIf igRptIndex = BUS_RPT Then
            SQLQuery = "select * from (((((((((((AIE_Active_Info inner join UIE_User_Info UIE_User_info on aieuiecode = uie_user_info.uiecode) "
            SQLQuery = SQLQuery & " inner join BDE_Bus_Definition BDE_To on aietofilecode = bde_to.bdecode) "
            SQLQuery = SQLQuery & " inner join BDE_Bus_Definition BDE_Bus_Definition on aiefromfilecode = BDE_Bus_Definition.bdecode) "
            SQLQuery = SQLQuery & " inner join Uie_User_Info uie_to on bde_to.bdeuiecode = uie_to.uiecode) "
            SQLQuery = SQLQuery & " inner join Uie_User_Info uie_from on BDE_Bus_Definition.bdeuiecode = uie_from.uiecode) "
            SQLQuery = SQLQuery & " left outer join ase_audio_source ase_to on bde_to.bdeasecode = ase_to.asecode)"
            SQLQuery = SQLQuery & " left outer join ase_audio_source ase_from on BDE_Bus_Definition.bdeasecode = ase_from.asecode) "
            SQLQuery = SQLQuery & " left outer join ane_audio_name ane_to on ase_to.aseprianecode = ane_to.anecode) "
            SQLQuery = SQLQuery & " left outer join ane_audio_name ane_from on ase_from.aseprianecode = ane_from.anecode) "
            SQLQuery = SQLQuery & " left outer join cce_control_char cce_to on bde_to.bdeccecode = cce_to.ccecode) "
            SQLQuery = SQLQuery & " left outer join cce_control_char cce_from on BDE_Bus_Definition.bdeccecode = cce_from.ccecode) "
            SQLQuery = SQLQuery & " where aiereffilename = 'BDE'"
            SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by bde_to.bdename"
        ElseIf igRptIndex = NETCUE_RPT Then
            SQLQuery = "Select * from  (((((((AIE_Active_Info Inner join UIE_User_Info UIE_User_Info on aieuiecode = uie_user_info.uiecode) "
            SQLQuery = SQLQuery & " inner join  nne_netcue_Name NNE_Netcue_Name on aiefromfilecode = nne_netcue_Name.nnecode) "
            SQLQuery = SQLQuery & " inner join NNE_Netcue_Name NNE_To on aietofilecode = nne_to.nnecode) "
            SQLQuery = SQLQuery & " left outer join dne_day_name dne_day_name on nne_netcue_name.nnednecode = dne_day_name.dnecode) "
            SQLQuery = SQLQuery & " left outer join  dne_day_name dne_to on nne_to.nnednecode = dne_to.dnecode) "
            SQLQuery = SQLQuery & " inner join uie_user_info uie_fromUser_info on nne_netcue_name.nneuiecode = uie_fromUser_info.uiecode) "
            SQLQuery = SQLQuery & " inner join uie_user_info uie_toUser_info on nne_to.nneuiecode = uie_toUser_info.uiecode) "
            SQLQuery = SQLQuery & " where aieRefFileName = 'NNE'"
            SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by NNE_Netcue_Name.nneName "
        ElseIf igRptIndex = CONTROL_RPT Then
            SQLQuery = "Select * from AIE_Active_Info, UIE_User_Info UIE_User_Info, "
            SQLQuery = SQLQuery & " CCE_Control_char CCE_Control_char, CCE_Control_char CCE_to, "
            SQLQuery = SQLQuery & " UIE_User_info UIE_FromUser_Info, UIE_User_info UIE_ToUser_Info "
            SQLQuery = SQLQuery & " where aieuiecode = UIE_User_Info.uiecode and aiefromfilecode = CCE_Control_char.ccecode and "
            SQLQuery = SQLQuery & " aietofilecode = CCE_to.ccecode and CCE_Control_char.cceuiecode = UIE_FromUser_Info.uiecode and "
            SQLQuery = SQLQuery & " CCE_to.cceuiecode = UIE_ToUser_Info.uiecode"
            SQLQuery = SQLQuery & " and AIE_Active_Info.aieRefFileName = 'CCE'"
            SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by CCE_To.cceType, CCE_to.cceAutoChar "
        ElseIf igRptIndex = COMMENT_RPT Then
            SQLQuery = "Select * from AIE_Active_Info, UIE_User_Info UIE_User_Info, "
            SQLQuery = SQLQuery & " CTE_Commts_And_Title CTE_Commts_And_Title, CTE_Commts_And_Title CTE_to, "
            SQLQuery = SQLQuery & " UIE_User_info UIE_FromUser_Info, UIE_User_info UIE_ToUser_Info "
            SQLQuery = SQLQuery & " where aieuiecode = UIE_User_Info.uiecode and aiefromfilecode = CTE_Commts_And_Title.ctecode and "
            SQLQuery = SQLQuery & " aietofilecode = CTE_to.ctecode and CTE_Commts_And_Title.cteuiecode = UIE_FromUser_Info.uiecode and "
            SQLQuery = SQLQuery & " CTE_to.cteuiecode = UIE_ToUser_Info.uiecode"
            SQLQuery = SQLQuery & " and AIE_Active_Info.aieRefFileName = 'CTE' and CTE_to.cteType = 'T2' "
            SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by CTE_to.cteComment"
        ElseIf igRptIndex = EVENT_RPT Then
            SQLQuery = "select * from aie_Active_Info, UIE_User_Info UIE_User_Info, EPE_Event_Properties EPE_Event_Properties, "
            SQLQuery = SQLQuery & " ETE_Event_Type ETE_Event_Type, UIE_User_Info UIE_FromETE, EPE_Event_Properties EPE_To, "
            SQLQuery = SQLQuery & " ETE_Event_Type ETE_To where  "
            SQLQuery = SQLQuery & " aieuiecode = UIE_User_Info.uiecode and aiefromfilecode = epe_Event_Properties.epecode and "
            SQLQuery = SQLQuery & " EPE_Event_Properties.epeetecode = ete_Event_type.etecode and ete_Event_type.eteuiecode = uie_fromete.uiecode and "
            SQLQuery = SQLQuery & " aietofilecode = epe_to.epecode and EPE_to.epeetecode = ete_to.etecode "
            SQLQuery = SQLQuery & " and (AIE_Active_Info.aieRefFileName = 'ETE' or  AIE_Active_Info.aieRefFileName = 'EPE') "
            SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by ETE_Event_Type.etename, aieorigfilecode, ETE_To.eteversion desc"
        ElseIf igRptIndex = AUTOMATION_RPT Then
            SQLQuery = "select * from aie_Active_Info, UIE_User_Info UIE_User_Info, AEE_Auto_Equip aEE_Auto_Equip, "
            SQLQuery = SQLQuery & " AEE_Auto_Equip AEE_To, UIE_User_Info UIE_From, UIE_User_Info UIE_To where "
            SQLQuery = SQLQuery & " aieuiecode = UIE_User_Info.uiecode and aiefromfilecode = AEE_Auto_Equip.aeecode and AEE_Auto_Equip.aeeuiecode = uie_From.uiecode and "
            SQLQuery = SQLQuery & " aietofilecode = aee_to.aeecode and aee_to.aeeuiecode = uie_to.uiecode and"
            SQLQuery = SQLQuery & " AIE_Active_Info.aieRefFileName = 'AEE'   "
            SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by AEE_to.aeeName"
        End If
    End If
    
    EngrCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName
    Screen.MousePointer = vbDefault
    

    If igRptSource = vbModal Then
        Unload EngrUserRpt
    End If
    
    Exit Sub
    

    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in User Rpt-cmdReport: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
            Screen.MousePointer = vbDefault
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in User Rpt-cmdReport: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub cmdReturn_Click()
    EngrReports.Show
    Unload EngrUserRpt
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    gSetFonts EngrUserRpt
    gCenterForm EngrUserRpt
End Sub

Private Sub Form_Load()
    'EngrUserRpt.Caption = "User - " & sgClientName
    gPopExportTypes cboFileType
    cboFileType.Enabled = False
    gChangeCaption frcOption
    mInit
    End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set EngrUserRpt = Nothing
End Sub
Private Sub optOldNew_Click(Index As Integer)
    If Index = 0 Then               'current
        lacChangeDates.Visible = False
        lacChangeTo.Visible = False
        edcFrom.Visible = False
        edcTo.Visible = False
        'If Audio Type report, additional option to see the sources that reference it
        If igRptIndex = AUDIOTYPE_RPT Then
            ckcInclOther.Caption = "Include Audio Source"
            'ckcInclOther.Move lacChangeDates.Left, edcFrom.Top
            ckcInclOther.Visible = True
        'if Bus Group report, additional option to see the buses that reference it
        ElseIf igRptIndex = BUSGROUP_RPT Then
            ckcInclOther.Caption = "Include Bus Definitions"
            'ckcInclOther.Move lacChangeDates.Left, edcFrom.Top
            ckcInclOther.Visible = True
        Else
            ckcInclOther.Visible = False
        End If
    Else
        lacChangeDates.Visible = True
        lacChangeTo.Visible = True
        edcFrom.Visible = True
        edcTo.Visible = True
        ckcInclOther.Visible = False
        
    End If
End Sub
Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       'default to adobe
    Else
        cboFileType.Enabled = False
    End If
End Sub
'
'           Initialize any screen selectivity coordinates
'
Public Sub mInit()
    'activate click event positions on screen
    optOldNew_Click 0
    If igRptSource = vbModal Then
        cmdReturn.Enabled = False
    Else
        cmdReturn.Enabled = True
    End If
End Sub
