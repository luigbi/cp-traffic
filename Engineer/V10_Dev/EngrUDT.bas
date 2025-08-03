Attribute VB_Name = "EngrUDT"
'
' Release: 1.0
'
' Description:
'   This file contains the User Defined Types (UDT)

Option Explicit

Type FILECHECK
    iByte(1 To 20000) As Byte
End Type

'As Aired import information
Type AAE
    lCode              As Long               ' Internal Code reference
    lSheCode           As Long               ' Schedule Header reference
    lSeeCode           As Long               ' Schedule Event reference
    sAirDate           As String * 10        ' Air Date of event
    lAirTime           As Long               ' Air Time (in tenths of a second)
    sAutoOff           As String * 1         ' Auto-Off Error: A=Auto-Off;
                                             ' M=Master Auto-Off; S=Semi-Auto
    sData              As String * 1         ' Data error: D=Error
    sSchedule          As String * 1         ' Schedule error: S=Sequence;
                                             ' G=Gap; O=Overlap
    sTrueTime          As String * 1         ' True Time error: T=True-time
                                             ' error
    sSourceConflict    As String * 1         ' Source Conflict error:C=Conflict
                                             ' error
    sSourceUnavail     As String * 1         ' Source Not Available: A=Not
                                             ' available error
    sSourceItem        As String * 1         ' Source Item Not Available: I=Item
                                             ' not available; K-Item not yet
                                             ' checked; E=Item exists
    sBkupSrceUnavail   As String * 1         ' Backup Source Not Available:
                                             ' A=Not available
    sBkupSrceItem      As String * 1         ' Backup Source Item Not Available:
                                             ' I=Item not available; K=Item not
                                             ' yet checked; E=Item exist
    sProtSrceUnavail   As String * 1         ' Protection Source Not Available:
                                             ' A=Not available
    sProtSrceItem      As String * 1         ' Protection Source Item Not
                                             ' Available: I=Item not available;
                                             ' K=Item not yet checked; E=Item
                                             ' exist
    sDate              As String * 10
    lEventID           As Long               ' Event ID
    sBusName           As String * 8         ' Bus name
    sBusControl        As String * 1         ' Bus Control
    sEventType         As String * 1         ' Event Type
    sStartTime         As String * 10        ' Start time in Tenths(xx:xx:xx.x)
    sStartType         As String * 3         ' Start Type
    sFixedTime         As String * 1         ' Fixed Time
    sEndType           As String * 3         ' End Type
    sDuration          As String * 10        ' Duration in Tenths (xx:xx:xx.x)
    sOutTime           As String * 10        ' Out Time in Tenths
    sMaterialType      As String * 3         ' Material type
    sAudioName         As String * 8         ' Audio source
    sAudioItemID       As String * 32        ' Audio commercial Cart # or
                                             ' Program Item ID
    sAudioISCI         As String * 20        ' ISCI Code
    sAudioCrtlChar     As String * 1         ' Audio Control Character
    sBkupAudioName     As String * 8         ' Backup Audio Name
    sBkupCtrlChar      As String * 1         ' Backup Control Character
    sProtAudioName     As String * 8         ' Protection Audio Name
    sProtItemID        As String * 32        ' Protection Item ID
    sProtISCI          As String * 20        ' Protection ISCI
    sProtCtrlChar      As String * 1         ' Protection Control Character
    sRelay1            As String * 8         ' Relay Name # 1
    sRelay2            As String * 8         ' Relay Name # 2
    sFollow            As String * 19        ' Follow
    sSilenceTime       As String * 7         ' Silence time length (xx:xx)
    sSilence1          As String * 1         ' Silence # 1
    sSilence2          As String * 1         ' Silence # 2
    sSilence3          As String * 1         ' Silence # 3
    sSilence4          As String * 1         ' Silence # 4
    sNetcueStart       As String * 3         ' Netcue Start
    sNetcueEnd         As String * 3         ' Netcue End
    sTitle1            As String * 66        ' Title # 1
    sTitle2            As String * 90        ' Title # 2
    sABCFormat         As String * 1         ' ABC Format.  Default value zero
                                             ' (0)
    sABCPgmCode        As String * 25        ' ABC Program Code
    sABCXDSMode        As String * 2         ' ABC XDS Mode.  Default value *
    sABCRecordItem     As String * 5         ' ABC Record item
    sEnteredDate       As String * 10        ' Entered Date
    sEnteredTime       As String * 11        ' Entered Time
    sUnused            As String * 20        ' Unused buffer
End Type

'Automation Contact
Type ACE
    iCode              As Integer         ' Internal Reference Code
    iAeeCode           As Integer         ' Auto Equip Reference
    sType              As String * 1      ' Type: P=Primary; S=Secondary (backup
                                          ' )
    sContact           As String * 40     ' Contact Name
    sPhone             As String * 20     ' Contact phone number
    sFax               As String * 20     ' Contact fax number
    sEMail             As String * 70     ' Contact E-Maill address
    sUnused            As String * 20     ' Unused buffer
End Type

'Automation Data Flags
Type ADE
    iCode              As Integer         ' Internal Reference Code
    iAeeCode           As Integer         ' Automation Equipment Reference
    iScheduleData      As Integer         ' Returned Schedule Data Start Column
    iDate              As Integer         ' Date field
    iDateNoChar        As Integer         ' Date number of characters
    iTime              As Integer         ' Time field
    iTimeNoChar        As Integer         ' Time number of characters
    iAutoOff           As Integer         ' Auto-Off Error
    iData              As Integer         ' Data error
    iSchedule          As Integer         ' Schedule error
    iTrueTime          As Integer         ' True Time error
    iSourceConflict    As Integer         ' Source conflict error
    iSourceUnavail     As Integer         ' Source Unavailable error
    iSourceItem        As Integer         ' Source Item Not Available error
    iBkupSrceUnavail   As Integer         ' Backup source Unavailable error
    iBkupSrceItem      As Integer         ' Backup source item not available err
                                          ' or
    iProtSrceUnavail   As Integer         ' Protection source unavailable error
    iProtSrceItem      As Integer         ' Protection item unavailable error
    sUnused            As String * 20     ' Unused buffer
End Type

'Automation Equipment
Type AEE
    iCode              As Integer         ' Internal Reference Code
    sName              As String * 20     ' Automation Equipment Name
    sDescription       As String * 50     ' Automation Equipment Description
    sManufacture       As String * 50     ' Manufacture Name of Automation Equip
                                          ' ment
    sFixedTimeChar     As String * 1      ' Fixed Time Export Character
    lAlertSchdDelay    As Long            ' Delay time after Schedule sent to te
                                          ' st if removed (hh:mm:ss)
    sState             As String * 1      ' Stae: A=Active; D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    iOrigAeeCode       As Integer         ' Original Automation Equipment Code u
                                          ' sed to tie all versions together
    sCurrent           As String * 1      ' Current Version: Y=Yes; N=No.  Y sho
                                          ' uld only be with the highest version
                                          '  #
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

'Automation Format
Type AFE
    iCode              As Integer            ' Internal Reference Code
    iAeeCode           As Integer            ' Automation Equipment Reference
    sType              As String * 1         ' Type:S=Schedule Export
    sSubType           As String * 1         ' SubType:S=Start Column; N=Number
                                             ' of Characters
    iBus               As Integer            ' Bus Field
    iBusControl        As Integer            ' Bus Control Field
    iEventType         As Integer            ' Event Type
    iTime              As Integer            ' Time Field
    iStartType         As Integer            ' Start Type for Time Field
    iFixedTime         As Integer            ' Fixed Time Field
    iEndType           As Integer            ' End Type for Time Field
    iDuration          As Integer            ' Duration Field
    iEndTime           As Integer            ' End Time
    iMaterialType      As Integer            ' Material Type Field
    iAudioName         As Integer            ' Audio Source Name Field
    iAudioItemID       As Integer            ' Audio Item ID Field
    iAudioISCI         As Integer            ' Audio ISCI
    iAudioControl      As Integer            ' Audio Control Field
    iBkupAudioName     As Integer            ' Backup Audio Source Name Field
    iBkupAudioControl  As Integer            ' Backup Audio Control Field
    iProtAudioName     As Integer            ' Protection (Backup of Backup)
                                             ' Audio Source Name Field
    iProtItemID        As Integer            ' Protection Audio Item ID Field
    iProtISCI          As Integer            ' Protection Audio ISCI
    iProtAudioControl  As Integer            ' Protection Audio Control Field
    iRelay1            As Integer            ' Relay 1 Field
    iRelay2            As Integer            ' Relay 2 Field
    iFollow            As Integer            ' Follow
    iSilenceTime       As Integer            ' Silence Time Field
    iSilence1          As Integer            ' Silence 1 Field
    iSilence2          As Integer            ' Silence 2 Field
    iSilence3          As Integer            ' Silence 3 Field
    iSilence4          As Integer            ' Silence 4 Field
    iStartNetcue       As Integer            ' Start Netcue Field
    iStopNetcue        As Integer            ' Stop Netcue Field
    iTitle1            As Integer            ' Title 1 Field
    iTitle2            As Integer            ' Title 2 Field
    iEventID           As Integer            ' Event ID
    iDate              As Integer            ' Date field
    iABCFormat         As Integer            ' ABC format field.
    iABCPgmCode        As Integer            ' ABC Program Code
    iABCXDSMode        As Integer            ' ABC XDS Mode.
    iABCRecordItem     As Integer            ' ABC Record Item
    sUnused            As String * 20        ' Unused buffer
End Type

'Active Information
Type AIE
    lCode              As Long            ' Internal Reference Code
    iUieCode           As Integer         ' User Reference
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    sRefFileName       As String * 8      ' Reference File Name
    lToFileCode        As Long            ' To File Code reference
    lFromFileCode      As Long            ' From File Code reference
    lOrigFileCode      As Long            ' Original File Code to tie all together
    sUnused            As String * 20     ' Unused buffer
End Type

'Audio Names
Type ANE
    iCode              As Integer         ' Internal Reference Code
    sName              As String * 8      ' Audio Name
    sDescription       As String * 50     ' Audio Name Description
    iCceCode           As Integer         ' Default Control Reference
    iAteCode           As Integer         ' Audio Type Reference
    sState             As String * 1      ' State: A=Active; D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    iOrigAneCode       As Integer         ' Original Audio Name Code used to tie
                                          '  version together
    sCurrent           As String * 1      ' Current version: Y=Yes; N=No
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (used that altered or
                                          '  added record)
    sCheckConflicts    As String * 1      ' Check for Conflicts(Y or N; default
                                          ' to Y; Test for N)
    sUnused            As String * 19     ' Unused
End Type

'Automation Path
Type APE
    iCode              As Integer         ' Internal Reference Code
    iAeeCode           As Integer         ' Automation Equipment Reference
    sType              As String * 2      ' Type: SE=Server Export; SI=Server Im
                                          ' port; CE=Client Export; CI=Client Im
                                          ' port
    sSubType           As String * 1         ' P or Blank = Production; T=Test
    sNewFileName       As String * 20     ' File name
    sChgFileName       As String * 20     ' Change file name
    sDelFileName       As String * 20     ' Delete file name
    sNewFileExt        As String * 3      ' New File Extension
    sChgFileExt        As String * 3      ' Change File Extension
    sDelFileExt        As String * 3      ' Deleted File Extension
    sPath              As String * 100    ' File Path
    sDateFormat        As String * 20     ' Date format associated with file nam
                                          ' e (or blank)
    sTimeFormat        As String * 20     ' Time format associated with file nam
                                          ' e
    sUnused            As String * 20     ' Unused buffer
End Type

'Advertiser Reference
Type ARE
    lCode              As Long            ' Internal Reference Code
    sName              As String * 35     ' Advertiser Name
    sUnusued           As String * 20     ' Unused buffer
End Type

'Audio Source
Type ASE
    iCode              As Integer         ' Internal Reference Code
    iPriAneCode        As Integer         ' Primary Audio Source Reference
    iPriCceCode        As Integer         ' Default Primary Audio Source Control
                                          '  Character reference
    sDescription       As String * 50     ' Audio Source Description
    iBkupAneCode       As Integer         ' Default Backup Audio Source Name Ref
                                          ' erence
    iBkupCceCode       As Integer         ' Default Backup Audio Source Control
                                          ' Character reference
    iProtAneCode       As Integer         ' Default Protection (Backup of the Ba
                                          ' ckup) Audio Source Name Reference
    iProtCceCode       As Integer         ' Default Protection Audio Source Cont
                                          ' rol Character reference
    sState             As String * 1      ' State: A=Active; D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    iOrigAseCode       As Integer         ' Original Audio Source Code used to t
                                          ' ie all versions together
    sCurrent           As String * 1      ' Current: Y=Yes; N=No.  Y should be o
                                          ' nly with the highest version #
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

'Audio Type
Type ATE
    iCode              As Integer         ' Internal Reference Code
    sName              As String * 20     ' Audio Type Name
    sDescription       As String * 50     ' Audio Type Description
    sState             As String * 1      ' State: A=Active; D=Dormant
    sTestItemID        As String * 1      ' Test Item ID (Y or N)
    lPreBufferTime     As Long            ' Pre-Buffer Time (in tenths of a seco
                                          ' nd). Used to check conflicts
    lPostBufferTime    As Long            ' Post-buffer Time (in tenths of a sec
                                          ' ond).  Used to check conflicts
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    iOrigAteCode       As Integer         ' Original Audio Type Code used to tie
                                          '  all versions together
    sCurrent           As String * 1      ' Current version: Y=Yes; N=No.  Y sho
                                          ' uld only be with the highest version
                                          '  #
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

'Bus Definition
Type BDE
    iCode              As Integer         ' Internal Reference Code
    sName              As String * 8      ' Bus Name
    sDescription       As String * 50     ' Bus Description
    sChannel           As String * 30     ' Channel Name
    iAseCode           As Integer         ' Default Commercial Autio Source Refe
                                          ' rence
    sState             As String * 1      ' State: A=Active; D=Dormant
    iCceCode           As Integer         ' Default Bus Control Reference
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    iOrigBdeCode       As Integer         ' Original Bus Code used to tie all ve
                                          ' rsions together
    sCurrent           As String * 1      ' Current Version: Y=Yes; N=No. Y shou
                                          ' ld be only with highest version #
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

'Bus Groups
Type BGE
    iCode              As Integer         ' Internal Reference Code
    sName              As String * 20     ' Bus Group Name
    sDescription       As String * 50     ' Bus Group Description
    sState             As String * 1      ' State: A=Active; D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    iOrigBgeCode       As Integer         ' Original Bus Group Code used to tie
                                          ' version together
    sCurrent           As String * 1      ' Current: Y=Yes; N=No.  Only the late
                                          ' st version should be Y
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

'Bus Selection GroupType BSE
Type BSE
    iCode              As Integer         ' Internal Reference Code
    iBdeCode           As Integer         ' Bus Definition Reference Code
    iBgeCode           As Integer         ' Bus Group Reference
    sUnused            As String * 20     ' Unused Buffer
End Type

'Control Character
Type CCE
    iCode              As Integer         ' Internal Reference Code
    sType              As String * 1      ' Type: A=Audio Source Control Charact
                                          ' er; B=Bus Control Character
    sAutoChar          As String * 1      ' Automation System Character
    sDescription       As String * 50     ' Bus Control Description
    sState             As String * 1      ' State: A=Active; D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    iOrigCceCode       As Integer         ' Original Bus Control Code used to ti
                                          ' e version together
    sCurrent           As String * 1      ' Current version: Y=Yes; N=No
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

'Conflict Events
Type CEE
    lCode              As Long               ' Auto Increment
    lGenDate           As Long               ' Generated Date
    lGenTime           As Long               ' Generated time
    sEvtType           As String * 1         ' Event Type: B=Bus; A=Audio Name
    iBdeCode           As Integer            ' BDECode if ceeEvtType is B and A;
    iANECode           As Integer            ' ANECode if ceeEvtType is A
    lStartDate         As Long               ' Start Date
    lEndDate           As Long               ' End Date
    sDay               As String * 2         ' Day: Mo; Tu; We; Th; Fr; Sa; Su
    lStartTime         As Long               ' Start Time in tenths of a second
    lEndTime           As Long               ' End Time in tenths of a second
    lGridEventRow      As Long               ' Grid Event Row
    iGridEventCol      As Integer            ' Grid event column
    sUnused            As String * 8         ' Unused
End Type

'Conflict Master Events
Type CME
    lCode              As Long               ' Auto Increment
    sSource            As String * 1         ' Source: S=Schedule; L=Library;
                                             ' T=Template
    lSHEDHECode        As Long               ' SHECode if cmeSource is S;
                                             ' DHECode if cmeSource is L or T
    lDseCode           As Long               ' Subname
    lDeeCode           As Long               ' DEECode for all cmeSource
    lSeeCode           As Long               ' SEECode if cmeSource = "S",
                                             ' otherwise it is zero
    sEvtType           As String * 1         ' Event Type: B=Bus; A=Audio Name
    iBdeCode           As Integer            ' BDECode if ceeEvtType is B and A;
    iANECode           As Integer            ' ANECode if ceeEvtType is A
    lStartDate         As Long               ' Start Date
    lEndDate           As Long               ' End Date
    sDay               As String * 2         ' Day: Mo; Tu; We; Th; Fr; Sa; Su
    lStartTime         As Long               ' Start Time in tenth of a second
    lEndTime           As Long               ' End Time in tenths of a second
    sItemID            As String * 32        ' Item ID used after compare to
                                             ' determine if not in error (same
                                             ' Item ID are Ok if times match)
    sXMidNight         As String * 1         ' Event cross midnight (Y or N)
    sUnused            As String * 8         ' Unused
End Type

'Comment and Title
Type CTE
    lCode              As Long            ' Internal Reference Code
    sType              As String * 2      ' T1=Title 1; T2=Title 2; DH=Day Heade
                                          ' r Comment
    sComment           As String * 66     ' Type: T1=66 char; T2=66 Char; DH=66
                                          '
    sState             As String * 1      ' State: A=Active; D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    lOrigCteCode       As Long            ' Ortiginal Comment Code used to tie a
                                          ' ll versions together
    sCurrent           As String * 1      ' Current: Y=Yes; N=No.  Y should only
                                          '  be with the highest version #
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

'Comment and Title
Type CTEAPI
    lCode              As Long            ' Internal Reference Code
    sType              As String * 2      ' T1=Title 1; T2=Title 2; DH=Day Heade
                                          ' r Comment
    sComment           As String * 66     ' Type: T1=66 char; T2=66 Char; DH=66
                                          '
    sState             As String * 1      ' State: A=Active; D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    lOrigCteCode       As Long            ' Ortiginal Comment Code used to tie a
                                          ' ll versions together
    sCurrent           As String * 1      ' Current: Y=Yes; N=No.  Y should only
                                          '  be with the highest version #
    iEneteredDate(0 To 1) As Integer      ' Entered Date
    iEnteredTime(0 To 1) As Integer       ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

'Day Bus Selection
Type DBE
    lCode              As Long            ' Internal Reference
    sType              As String * 1      ' Type: G=Group; B=Bus
    lDheCode           As Long            ' Day Name Reference
    iBdeCode           As Integer         ' Bus Reference
    iBgeCode           As Integer         ' Bus Group Reference
    sUnused            As String * 20     ' Unused
End Type

'Day Events
Type DEE
    lCode              As Long               ' Internal Reference Code
    lDheCode           As Long               ' Day Header Reference
    iCceCode           As Integer            ' Bus Control Reference
    iEteCode           As Integer            ' Event Type reference
    lTime              As Long               ' Event Time (in tenths of a
                                             ' second)
    iStartTteCode      As Integer            ' Start Time Type reference
    sFixedTime         As String * 1         ' Fixed Time: Y=Yes; N=No
    iEndTteCode        As Integer            ' End Time Type reference
    lDuration          As Long               ' Duration (in tenths of a second)
    sHours             As String * 24        ' Hours to Run
    sDays              As String * 7         ' Days to Air (run), Y or N, first
                                             ' character(left most) is Monday,
                                             ' second is Tuesday,...
    iMteCode           As Integer            ' Material Type reference
    iAudioAseCode      As Integer            ' Audio Source Reference
    sAudioItemID       As String * 32        ' Audio Commercial Cart # or
                                             ' Program Item ID
    sAudioISCI         As String * 20        ' ISCI code
    iAudioCceCode      As Integer            ' Audio Source Control reference
    iBkupAneCode       As Integer            ' Backup Audio Name Reference
    iBkupCceCode       As Integer            ' Backup Audio Control Reference
    iProtAneCode       As Integer            ' Protection (backup of the Backup)
                                             ' Audio Name Reference
    sProtItemID        As String * 32        ' Protection Commercial Cart # or
                                             ' Program Item ID
    sProtISCI          As String * 20        ' Protection ISCI code
    iProtCceCode       As Integer            ' Protect Control Character
                                             ' Reference
    i1RneCode          As Integer            ' Relay 1 Reference
    i2RneCode          As Integer            ' Relay 2 Reference
    iFneCode           As Integer            ' Follow Name Reference
    lSilenceTime       As Long               ' Silence time length (mm:ss)
    i1SceCode          As Integer            ' Silence Character reference
    i2SceCode          As Integer            ' Silence Character Reference
    i3SceCode          As Integer            ' Silence Character Reference
    i4SceCode          As Integer            ' Silence Character Reference
    iStartNneCode      As Integer            ' Start Netcue Reference
    iEndNneCode        As Integer            ' End Netcue Reference
    l1CteCode          As Long               ' Comment Title 1 reference
    l2CteCode          As Long               ' Comment Title 2 Reference
    lEventID           As Long               ' Event ID (Obtained from Site and
                                             ' used as external reference)
    sIgnoreConflicts   As String * 1         ' A=Ignore Audio Conflicts;
                                             ' B=Ignore Bus Conflicts; I=Ignore
                                             ' Bus and Audio Conflicts
    sABCFormat         As String * 1         ' ABC Format.  Default value zero
                                             ' (0)
    sABCPgmCode        As String * 25        ' ABC Program Code
    sABCXDSMode        As String * 2         ' ABC XDS Mode.  Default value *
    sABCRecordItem     As String * 5         ' ABC Record item
    sUnused            As String * 19        ' Unused buffer
End Type

Type DEEAPI
    lCode                 As Long            ' Internal Reference Code
    lDheCode              As Long            ' Day Header Reference
    iCceCode              As Integer         ' Bus Control Reference
    iEteCode              As Integer         ' Event Type reference
    lTime                 As Long            ' Event Time (in tenths of a
                                             ' second)
    iStartTteCode         As Integer         ' Start Time Type reference
    sFixedTime            As String * 1      ' Fixed Time: Y=Yes; N=No
    iEndTteCode           As Integer         ' End Time Type reference
    lDuration             As Long            ' Duration (in tenths of a second)
    sHours                As String * 24     ' Hours to Run
    sDays                 As String * 7      ' Days to Air (run), Y or N, first
                                             ' character(left most) is Monday,
                                             ' second is Tuesday,...
    iMteCode              As Integer         ' Material Type reference
    iAudioAseCode         As Integer         ' Audio Source Reference
    sAudioItemID          As String * 32     ' Audio Commercial Cart # or
                                             ' Program Item ID
    sAudioISCI            As String * 20     ' ISCI code
    iAudioCceCode         As Integer         ' Audio Source Control reference
    iBkupAneCode          As Integer         ' Backup Audio Name Reference
    iBkupCceCode          As Integer         ' Backup Audio Control Reference
    iProtAneCode          As Integer         ' Protection (backup of the Backup)
                                             ' Audio Name Reference
    sProtItemID           As String * 32     ' Protection Commercial Cart # or
                                             ' Program Item ID
    sProtISCI             As String * 20     ' Protection ISCI code
    iProtCceCode          As Integer         ' Protect Control Character
                                             ' Reference
    i1RneCode             As Integer         ' Relay 1 Reference
    i2RneCode             As Integer         ' Relay 2 Reference
    iFneCode              As Integer         ' Follow Name Reference
    lSilenceTime          As Long            ' Silence time length (mm:ss)
    i1SceCode             As Integer         ' Silence Character reference
    i2SceCode             As Integer         ' Silence Character Reference
    i3SceCode             As Integer         ' Silence Character Reference
    i4SceCode             As Integer         ' Silence Character Reference
    iStartNneCode         As Integer         ' Start Netcue Reference
    iEndNneCode           As Integer         ' End Netcue Reference
    l1CteCode             As Long            ' Comment Title 1 reference
    l2CteCode             As Long            ' Comment Title 2 Reference
    lEventID              As Long            ' Event ID (Obtained from Site and
                                             ' used as external reference)
    sIgnoreConflicts      As String * 1      ' A=Ignore Audio Conflicts;
                                             ' B=Ignore Bus Conflicts; I=Ignore
                                             ' Bus and Audio Conflicts
    sABCFormat            As String * 1      ' ABC Format.  Default value zero
                                             ' (0)
    sABCPgmCode           As String * 25     ' ABC Program Code
    sABCXDSMode           As String * 2      ' ABC XDS Mode.  Default value *
    sABCRecordItem        As String * 5      ' ABC Record item
    sUnused               As String * 19     ' Unused buffer
End Type

'Day Header
Type DHE
    lCode              As Long            ' Internal Reference Code
    sType              As String * 1      ' Type: L=Library; T=Template
    lDneCode           As Long            ' Day Name Reference
    lDseCode           As Long            ' Day Sub-Name reference
    sStartTime         As String * 11     ' Type =L: Library Start Time
    lLength            As Long            ' Library or Template length (in secon
                                          ' ds)
    sHours             As String * 24     ' Library Default run Hours
    sStartDate         As String * 10     ' Library or Template Start Date
    sEndDate           As String * 10     ' Library or Template End Date
    sDays              As String * 7      ' Library Default Days, Y or N, First
                                          ' character (left most) is Monday, Sec
                                          ' ond Character is Tuesday,...
    lCteCode           As Long            ' Comment Reference
    sState             As String * 1      ' State: A=Active; D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    lOrigDHECode       As Long            ' Original Day Header Code used to tie
                                          '  all versions together
    sCurrent           As String * 1      ' Current version: Y=Yes; N=No.  Y sho
                                          ' uld only be with highest version #
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sIgnoreConflicts   As String * 1      ' A=Ignore Audio Conflicts; B=Ignore B
                                          ' us Conflicts; I=Ignore Bus and Audio
                                          '  Conflicts
    sBusNames          As String * 50     ' Buses separated by comma's.  User to speed-up
                                          ' showing the bus names on the selection screen
    sUnused            As String * 19     ' Unused buffer
End Type

Type DHEAPI
    lCode                 As Long            ' Internal Reference Code
    sType                 As String * 1      ' Type: L=Library; T=Template
    lDneCode              As Long            ' Day Name Reference
    lDseCode              As Long            ' Day Sub-Name reference
    iStartTime(0 To 1)    As Integer         ' Type =L: Library Start Time
    lLength               As Long            ' Library or Template length (in
                                             ' seconds)
    sHours                As String * 24     ' Library Default run Hours
    iStartDate(0 To 1)    As Integer         ' Library or Template Start Date
    iEndDate(0 To 1)      As Integer         ' Library or Template End Date
    sDays                 As String * 7      ' Library Default Days, Y or N,
                                             ' First character (left most) is
                                             ' Monday, Second Character is
                                             ' Tuesday,...
    lCteCode              As Long            ' Comment Reference
    sState                As String * 1      ' State: A=Active; D=Dormant
    sUsedFlag             As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion              As Integer         ' Version (Starting at 0)
    lOrigDHECode          As Long            ' Original Day Header Code used to
                                             ' tie all versions together
    sCurrent              As String * 1      ' Current version: Y=Yes; N=No.  Y
                                             ' should only be with highest
                                             ' version #
    iEnteredDate(0 To 1)  As Integer         ' Entered Date
    iEnteredTime(0 To 1)  As Integer         ' Entered Time
    iUieCode              As Integer         ' User Reference (User that altered
                                             ' or added record)
    sIgnoreConflicts      As String * 1      ' A=Ignore Audio Conflicts;
                                             ' B=Ignore Bus Conflicts; I=Ignore
                                             ' Bus and Audio Conflicts
    sBusNames             As String * 50     ' Buses separated by comma's.  User to speed-up
                                             ' showing the bus names on the selection screen
    sUnused               As String * 19     ' Unused buffer
End Type



'Day Name
Type DNE
    lCode              As Long            ' Internal Reference Code
    sType              As String * 1      ' Type:  L=Library; T=Template
    sName              As String * 20     ' Day Name
    sDescription       As String * 50     ' Description
    sState             As String * 1      ' State: A=Active; D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    lOrigDneCode       As Long            ' Original Day Name used to tie all ve
                                          ' rsions together
    sCurrent           As String * 1      ' Current: Y=Yes; N=No,  Y should only
                                          '  be with the highest version #
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

'Day Subname
Type DSE
    lCode              As Long            ' Internal Reference Code
    sName              As String * 20     ' Day Sub-Name
    sDescription       As String * 50
    sState             As String * 1      ' State: A=Active, D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    lOrigDseCode       As Long            ' Original Day Sub-Name Code used to t
                                          ' ie all versions together
    sCurrent           As String * 1      ' Current: Y=Yes; N=No.  Y should only
                                          '  be with the highest version #
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

'Day Event Buses
Type EBE
    lCode              As Long            ' Internal Reference
    lDeeCode           As Long            ' Day Event Reference
    iBdeCode           As Integer         ' Bus Reference
    sUnused            As String * 20     ' Unused
End Type

'Event Properities
Type EPE
    iCode              As Integer            ' Internal Reference Code
    iEteCode           As Integer            ' Event Type Reference
    sType              As String * 1         ' Type: U=Used; M=Mandatory
    sBus               As String * 1         ' Bus Field: Y=Yes; N=No
    sBusControl        As String * 1         ' Bus Control Field: Y=Yes; N=No
    sTime              As String * 1         ' Time Field: Y=Yes; N=No
    sStartType         As String * 1         ' Start Type Field: Y=Yes; N=No
    sFixedTime         As String * 1         ' Time Fixed Field: Y=Yes; N=No
    sEndType           As String * 1         ' End Type Field: Y=Yes; N=No
    sDuration          As String * 1         ' Duration Feild: Y=Yes; N=No
    sMaterialType      As String * 1         ' MaterialType field: Y=Yes; N=No
    sAudioName         As String * 1         ' Audio Source Name Field: Y=Yes;
                                             ' N=No
    sAudioItemID       As String * 1         ' Audio Item ID Field: Y=Yes; N=No
    sAudioISCI         As String * 1         ' Audio ISCI Field: Y=Yes; N=No
    sAudioControl      As String * 1         ' Audio Control Field: Y=Yes; N=No
    sBkupAudioName     As String * 1         ' Backup Audio Source Field: Y=Yes;
                                             ' N=No
    sBkupAudioControl  As String * 1         ' Backup Audio Control Field:
                                             ' Y=Yes; N=No
    sProtAudioName     As String * 1         ' Protection (Backup of the Backup)
                                             ' Audio Source Name Field: Y=Yes;
                                             ' N=No
    sProtAudioItemID   As String * 1         ' Protection Audio Item ID Field:
                                             ' Y=Yes; No=No
    sProtAudioISCI     As String * 1         ' Protection Audio ISCI Field:
                                             ' Y=Yes; N=No
    sProtAudioControl  As String * 1         ' Protection Audio Control Field:
                                             ' Y=Yes; N=No
    sRelay1            As String * 1         ' Relay 1 Field: Y=Yes; N=No
    sRelay2            As String * 1         ' Relay 2 Field: Y=Yes; N=No
    sFollow            As String * 1         ' Follow Field: Y=Yes; N=No
    sSilenceTime       As String * 1         ' Sielence Time Field: Y=Yes; N=No
    sSilence1          As String * 1         ' Silence 1 Field: Y=Yes; N=No
    sSilence2          As String * 1         ' Silence 2 Field: Y=Yes; N=No
    sSilence3          As String * 1         ' Silence 3 Field: Y=Yes; N=No
    sSilence4          As String * 1         ' Silence 4 Field: Y=Yes; N=No
    sStartNetcue       As String * 1         ' Start Netcue Field: Y=Yes; N=No
    sStopNetcue        As String * 1         ' Stop Netcue Field: Y=Yes; N=No
    sTitle1            As String * 1         ' Title 1 Field: Y=Yes; N=No
    sTitle2            As String * 1         ' Title 2 Field: Y=Yes; N=No
    sABCFormat         As String * 1         ' ABC Format Field: Y=Yes; N=No
    sABCPgmCode        As String * 1         ' ABC Program Code Field: Y=Yes;
                                             ' N=No
    sABCXDSMode        As String * 1         ' ABC XDS Mode Field: Y=Yes; N=No
    sABCRecordItem     As String * 1         ' ABC Record Item Field: Y=Yes;
                                             ' N=No
    sUnused            As String * 20        ' Unused buffer
End Type

'Event Type
Type ETE
    iCode              As Integer         ' Internal Reference Code
    sName              As String * 20     ' Event Type Name
    sDescription       As String * 50     ' Event Type description
    sCategory          As String * 1      ' Category: P=Program; A=Avail
    sAutoCodeChar      As String * 1      ' Automation System Code Character
    sState             As String * 1      ' State: A=Active; D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    iOrigEteCode       As Integer         ' Original Event Type Code used to tie
                                          '  all version together
    sCurrent           As String * 1      ' Current: Y=Yes; N=No.  Y should only
                                          '  be with highest version #
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

'Follow Name
Type FNE
    iCode              As Integer         ' Internal Reference Code
    sName              As String * 19     ' Master/Follow Name
    sDescription       As String * 50     ' Follow Description
    sState             As String * 1      ' State: A=Active; D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    iOrigFneCode       As Integer         ' Original Follow Name Code used to ti
                                          ' e all versions together
    sCurrent           As String * 1      ' Current version: Y=Yes; N=No.  Y sho
                                          ' uld only be with the highest version
                                          '  #
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

'ITE_Item_Test
Type ITE
    iCode              As Integer            ' Internal Reference Code
    iSoeCode           As Integer            ' Site Option Reference
    sType              As String * 1         ' Type: P=Primary; S=Secondary
                                             ' (Backup)
    sName              As String * 20        ' Name of Equipment
    iDataBits          As Integer            ' Data Bits on serial port
    sParity            As String * 1         ' Parity: E=Even;O=Odd; M=Mark;
                                             ' N=None;S=Space
    sStopBit           As String * 1         ' Stop Bit Required: 1=1 Stop bit;
                                             ' 2=2 stop bits
    iBaud              As Integer            ' Baud Rate
    sMachineID         As String * 2         ' Machine ID
    sStartCode         As String * 1         ' Start Code character
    sReplyCode         As String * 1         ' End Code Character
    iMinMgsID          As Integer            ' Min message ID
    iMaxMgsID          As Integer            ' Max Message ID
    iCurrMgsID         As Integer            ' Current Message ID
    sMgsType           As String * 2         ' Message Type
    sCheckSum          As String * 1         ' Check Sum required: Y=Yes; N=No
    sCmmdSeq           As String * 20        ' Command Sequence
    sMgsEndCode        As String * 10        ' Message end code
                                             ' characters(lf,cr,13)
    sTitleID           As String * 2         ' Title Command ID
    sLengthID          As String * 2         ' Length Command ID
    sConnectSeq        As String * 10        ' Connection Sequence command.
                                             ' Also used to wake up connection
    sMgsErrType        As String * 2         ' Message Error Type
    sUnused            As String * 20        ' Unused buffer
End Type

'Message Type
Type MIE
    lCode                 As Long            ' Internal Reference Code
    sType                 As String * 1      ' Message Type: M=Import Merge;
                                             ' S=Creation of Schedule;
                                             ' E=Generation of Automation
                                             ' Export; A=As Air Import
    lID                   As Long            ' Message ID.  Used to retain
                                             ' message generation order
    sMessage              As String * 100    ' Message
    sEnteredDate          As String * 10     ' Entered Date
    sEnteredTime          As String * 11     ' Entered Time
    iUieCode              As Integer         ' User Reference (User that was
                                             ' running when message created)
    sUnused               As String * 20     ' Unused buffer
End Type

'Material Type
Type MTE
    iCode              As Integer         ' Internal Reference Code
    sName              As String * 3      ' Material Type Name
    sDescription       As String * 50     ' Material Type Description
    sState             As String * 1      ' State: Y=Yes; N=No
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    iOrigMteCode       As Integer         ' Original Material Type Code used to
                                          ' tie all version together
    sCurrent           As String * 1      ' Current Version: Y=Yes; N=No
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

'Netcue Name
Type NNE
    iCode              As Integer         ' Internal Reference Code
    sName              As String * 3      ' Netcue Name
    sDescription       As String * 50     ' Netcue Description
    lDneCode           As Long            ' Day Name to associate Netcue with or
                                          '  0 if none
    sState             As String * 1      ' State: Active; D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    iOrigNneCode       As Integer         ' Original Netcue Name Code used to ti
                                          ' e all versions together
    sCurrent           As String * 1      ' Current Version: Y=Yes; N=No
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

Type RLE
    lCode              As Long               ' Internal Reference Code
    iUieCode           As Integer            ' User Reference (User that Locked
                                             ' record)
    sFileName          As String * 3         ' File Name
    lRecCode           As Long               ' Record Locked
    sEnteredDate       As String * 10        ' Entered Date
    sEnteredTime       As String * 11        ' Entered Time
    sUnused            As String * 20        ' Unused buffer
End Type

Type RLEAPI
    lCode                 As Long            ' Internal Reference Code
    iUieCode              As Integer         ' User Reference (User that Locked
                                             ' record)
    sFileName             As String * 3      ' File Name
    lRecCode              As Long            ' Record Locked
    iEnteredDate(0 To 1)  As Integer         ' Entered Date
    iEnteredTime(0 To 1)  As Integer         ' Entered Time
    sUnused               As String * 20     ' Unused buffer
End Type

'Relay
Type RNE
    iCode              As Integer         ' Internal Reference Code
    sName              As String * 8      ' Relay Name
    sDescription       As String * 50     ' Relay Description
    sState             As String * 1      ' State: A=Active; D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    iOrigRneCode       As Integer         ' Original Relay Name Code used to tie
                                          '  all versions together
    sCurrent           As String * 1      ' Current Version: Y=Yes; N=No
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

'Silence Character
Type SCE
    iCode              As Integer         ' Internal Reference Code
    sAutoChar          As String * 1      ' Automation System Character
    sDescription       As String * 50     ' Silence Description
    sState             As String * 1      ' State: A=Active; D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    iOrigSceCode       As Integer         ' Original Silence Character Code used
                                          '  to tie all versions together
    sCurrent           As String * 1      ' Current Version: Y=Yes; N=No.  Y sho
                                          ' uld only be with highest version #
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

'Schedule Events
Type SEE
    lCode              As Long               ' Internal Reference Code
    lSheCode           As Long               ' Schedule Header Reference
    sAction            As String * 1         ' Type of Action: N=New; C=Changed;
                                             ' D=Deleted(if sent); U=Unchanged; R=Removed(if Unsent)
    lDeeCode           As Long               ' Day Event used to created this
                                             ' event Reference
    iBdeCode           As Integer            ' Bus Reference Code
    iBusCceCode        As Integer            ' Bus Control Character reference
    sSchdType          As String * 1         ' Schedule Event Type: S=Spot;
                                             ' A=Avail; L=Library; T=Template
    iEteCode           As Integer            ' Event Type reference
    lTime              As Long               ' Event Time (in tenths of a
                                             ' second)
    iStartTteCode      As Integer            ' Start Time Type reference
    sFixedTime         As String * 1         ' Fixed Time: Y=Yes; N=No
    iEndTteCode        As Integer            ' End Time Type reference
    lDuration          As Long               ' Duration (in tenths of a second)
    iMteCode           As Integer            ' Material Code
    iAudioAseCode      As Integer            ' Audio Source reference
    sAudioItemID       As String * 32        ' Audio Commercial Cart # or
                                             ' Program Item ID
    sAudioItemIDChk    As String * 1         ' Audio Item ID checked on Primary
                                             ' device: N=Not Checked; O=Checked
                                             ' and OK; F=Checked and Failed
    sAudioISCI         As String * 20        ' Audio ISCI
    iAudioCceCode      As Integer            ' Audio Source Control reference
    iBkupAneCode       As Integer            ' Backup Audio Name reference
    iBkupCceCode       As Integer            ' Backup Audio Source Control
                                             ' reference
    iProtAneCode       As Integer            ' Protection (Backup of the backup)
                                             ' Audio Name reference
    sProtItemID        As String * 32        ' Protection Commercial Cart # or
                                             ' Program Item ID
    sProtItemIDChk     As String * 1         ' Protection Item ID checked on
                                             ' Secondary device: N=Not Checked;
                                             ' O=Checked and OK; F=Checked and
                                             ' Failed
    sProtISCI          As String * 20        ' Protection Audio ISCI
    iProtCceCode       As Integer            ' Protection Audio Control
                                             ' reference
    i1RneCode          As Integer            ' Relay 1 reference
    i2RneCode          As Integer            ' Relay 2 reference
    iFneCode           As Integer            ' Follow Name reference
    lSilenceTime       As Long               ' Silence time length (mm:ss)
    i1SceCode          As Integer            ' Silence Character reference
    i2SceCode          As Integer            ' Silence Character reference
    i3SceCode          As Integer            ' Silence Character reference
    i4SceCode          As Integer            ' Slience Character reference
    iStartNneCode      As Integer            ' Start Netcue reference
    iEndNneCode        As Integer            ' End Netcue reference
    l1CteCode          As Long               ' Comment Title 1 reference
    l2CteCode          As Long               ' Comment Title 2 reference
    lAreCode           As Long               ' Advertiser Name reference
    lSpotTime          As Long               ' Spot time (or -1 if avail or
                                             ' program)
    lEventID           As Long               ' Unique Event ID (Range set in
                                             ' Site)
    sAsAirStatus       As String * 1         ' As Aired Status: P=Posted; N=Not
                                             ' posted
    sSentStatus        As String * 1         ' Sent to automation status: N=Not
                                             ' Send; S=Sent
    sSentDate          As String * 10        ' Sent Date
    sIgnoreConflicts   As String * 1         ' A=Ignore Audio Conflicts;
                                             ' B=Ignore Bus Conflicts; I=Ignore
                                             ' Bus and Audio Conflicts
    lDheCode           As Long               ' Used to access DHECode instead of
                                             ' reading in DEE to get the value
                                             ' (helps speed up conflict testing)
    lOrigDHECode       As Long               ' Used to test DHECode in Conflict
                                             ' testing (help speed up code so
                                             ' that accessing DEE, then DHE not
                                             ' required)
    sInsertFlag        As String * 1         ' Temporary flag used only in
                                             ' Schedule Definition to know is
                                             ' row inserted
    sABCFormat         As String * 1         ' ABC Format.  Default value zero
                                             ' (0).
    sABCPgmCode        As String * 25        ' ABC Program Code
    sABCXDSMode        As String * 2         ' ABC XDS Mode.  Default value *
    sABCRecordItem     As String * 5         ' ABC Record Item
    sUnused            As String * 10        ' Unused Buffer
    'The follow field is not part of the definition of SEE
    'It is used for the remaining Avail Time
    lAvailLength As Long
End Type

Type SEETIMESORT
    sKey As String * 30 'Event Time or Spot Time; BDECode
    tSEE As SEE
End Type

Type SEEAPI
    lCode                 As Long            ' Internal Reference Code
    lSheCode              As Long            ' Schedule Header Reference
    sAction               As String * 1      ' Type of Action: N=New; C=Changed;
                                             ' D=Deleted(if Sent); U=Unchanged;
                                             ' R=Removed(if not sent)
    lDeeCode              As Long            ' Day Event used to created this
                                             ' event Reference
    iBdeCode              As Integer         ' Bus Reference Code
    iBusCceCode           As Integer         ' Bus Control Character reference
    sSchdType             As String * 1      ' Schedule Event Type: S=Spot;
                                             ' A=Avail; L=Library; T=Template
    iEteCode              As Integer         ' Event Type reference
    lTime                 As Long            ' Event Time (in tenths of a
                                             ' second)
    iStartTteCode         As Integer         ' Start Time Type reference
    sFixedTime            As String * 1      ' Fixed Time: Y=Yes; N=No
    iEndTteCode           As Integer         ' End Time Type reference
    lDuration             As Long            ' Duration (in tenths of a second)
    iMteCode              As Integer         ' Material Code
    iAudioAseCode         As Integer         ' Audio Source reference
    sAudioItemID          As String * 32     ' Audio Commercial Cart # or
                                             ' Program Item ID
    sAudioItemIDChk       As String * 1      ' Audio Item ID checked on Primary
                                             ' device: N=Not Checked; O=Checked
                                             ' and OK; F=Checked and Failed
    sAudioISCI            As String * 20     ' Audio ISCI
    iAudioCceCode         As Integer         ' Audio Source Control reference
    iBkupAneCode          As Integer         ' Backup Audio Name reference
    iBkupCceCode          As Integer         ' Backup Audio Source Control
                                             ' reference
    iProtAneCode          As Integer         ' Protection (Backup of the backup)
                                             ' Audio Name reference
    sProtItemID           As String * 32     ' Protection Commercial Cart # or
                                             ' Program Item ID
    sProtItemIDChk        As String * 1      ' Protection Item ID checked on
                                             ' Secondary device: N=Not Checked;
                                             ' O=Checked and OK; F=Checked and
                                             ' Failed
    sProtISCI             As String * 20     ' Protection Audio ISCI
    iProtCceCode          As Integer         ' Protection Audio Control
                                             ' reference
    i1RneCode             As Integer         ' Relay 1 reference
    i2RneCode             As Integer         ' Relay 2 reference
    iFneCode              As Integer         ' Follow Name reference
    lSilenceTime          As Long            ' Silence time length (mm:ss)
    i1SceCode             As Integer         ' Silence Character reference
    i2SceCode             As Integer         ' Silence Character reference
    i3SceCode             As Integer         ' Silence Character reference
    i4SceCode             As Integer         ' Slience Character reference
    iStartNneCode         As Integer         ' Start Netcue reference
    iEndNneCode           As Integer         ' End Netcue reference
    l1CteCode             As Long            ' Comment Title 1 reference
    l2CteCode             As Long            ' Comment Title 2 reference
    lAreCode              As Long            ' Advertiser Name reference
    lSpotTime             As Long            ' Spot time (or -1 if avail or
                                             ' program)
    lEventID              As Long            ' Unique Event ID (Range set in
                                             ' Site)
    sAsAirStatus          As String * 1      ' As Aired Status: P=Posted; N=Not
                                             ' posted
    sSentStatus           As String * 1      ' Sent to automation status: N=Not
                                             ' Send; S=Sent
    sSentDate             As String * 10     ' Sent Date
    sIgnoreConflicts      As String * 1      ' A=Ignore Audio Conflicts;
                                             ' B=Ignore Bus Conflicts; I=Ignore
                                             ' Bus and Audio Conflicts
    lDheCode              As Long            ' Used to access DHECode instead of
                                             ' reading in DEE to get the value
                                             ' (helps speed up conflict testing)
    lOrigDHECode          As Long            ' Used to test DHECode in Conflict
                                             ' testing (help speed up code so
                                             ' that accessing DEE, then DHE not
                                             ' required)
    sInsertFlag           As String * 1      ' Temporary flag used only in
                                             ' Schedule Definition to know is
                                             ' row inserted
    sABCFormat            As String * 1      ' ABC Format.  Default value zero
                                             ' (0).
    sABCPgmCode           As String * 25     ' ABC Program Code
    sABCXDSMode           As String * 2      ' ABC XDS Mode.  Default value *
    sABCRecordItem        As String * 5      ' ABC Record Item
    sUnused               As String * 10     ' Unused Buffer
End Type

Type SEEKEY4
    lSheCode              As Long
    lTime                 As Long
End Type

Type SEEKEY5
    lSheCode              As Long
    iBdeCode              As Integer
End Type

Type SEESORT
    sKey As String * 30 'Time; BDECode; SpotTime
    tSEEAPI As SEEAPI
End Type

Type SEEBRACKET
    sSource As String * 1
    lIndex As Long
End Type

'SGE_Site_Gen_Schd
Type SGE
    iCode              As Integer         ' Internal Reference Code
    iSoeCode           As Integer         ' Site Option Reference
    sType              As String * 1      ' Type of Creation:S=Schedule; A=Autom
                                          ' ation Export
    sSubType           As String * 1      ' P or Blank = Production; T=Test
    iGenMo             As Integer         ' # of lead days to create Mondays ima
                                          ' ge
    iGenTu             As Integer         ' # of lead days to create Tuesdays im
                                          ' age
    iGenWe             As Integer         ' # of lead days to create Wednesdays
                                          ' image
    iGenTh             As Integer         ' # of lead days to create Thursdays i
                                          ' mage
    iGenFr             As Integer         ' # of lead days to create Fridays ima
                                          ' ge
    iGenSa             As Integer         ' # of lead days to create Saturdays i
                                          ' mage
    iGenSu             As Integer         ' # of lead days to create Sundays ima
                                          ' ge
    sGenTime           As String * 11     ' Generation Time
    sPurgeAfterGen     As String * 1      ' Purge After Generation: Y=Yes; N=No
                                          ' (Use Purge Time)
    sPurgeTime         As String * 11     ' sgeType = S: Purge Time
    lAlertInterval     As Long            ' sgeType=A:  Creat Alert after xx min
                                          ' utes if Automation file nor removed
    sUnused            As String * 20     ' Unused buffer
End Type

'Schedule Header
Type SHE
    lCode              As Long            ' Internal Reference Code
    iAeeCode           As Integer         ' Automation System Reference
    sAirDate           As String * 10     ' Air date
    sLoadedAutoStatus  As String * 1      ' Load Automation System Status: N=Not
                                          '  Loaded; L=Loaded
    sLoadedAutoDate    As String * 10     ' Loaded Automation System Date
    iChgSeqNo          As Integer         ' Last Change Sequence Number used
    sAsAirStatus       As String * 1      ' As Aired Status: N=Not imported; I=I
                                          ' mported (As Air)
    sLoadedAsAirDate   As String * 10     ' Loaded date for As Aired
    sLastDateItemChk   As String * 10     ' Last date Item ID's checked
    sCreateLoad        As String * 1      ' Create Load (call from Engineering t
                                          ' o Service) (Y/N)
    iVersion           As Integer         ' Version (Starting at 0)
    lOrigSheCode       As Long            ' Original Schedule Header Code used t
                                          ' o tie all versions together
    sCurrent           As String * 1      ' Current version: Y=Yes; N=No.  Y sho
                                          ' uld only be with highest version
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sConflictExist     As String * 1      ' Y=Yes; N=No.
    sSpotMergeStatus   As String * 1      ' E=Error(Overbooked, Underbooked, No Avail); M= Spot Merge without Errors; N=Not Merged.
    sLoadStatus        As String * 1      ' E=Error; N=No Error.
    sXDayGenOnly       As String * 1      ' Cross Day Generated Only (Y/N).
                                          ' i.e. Y=Template events that
                                          ' crossed midnight generated into
                                          ' next day schedule only
    sUnused            As String * 16     ' Unused buffer
End Type

'SOE_Site_Option
Type SOE
    iCode              As Integer         ' Internal Reference Code
    sClientName        As String * 40     ' Client Name
    sAddr1             As String * 60     ' Client Address Line 1
    sAddr2             As String * 60     ' Client Address Line 2
    sAddr3             As String * 60     ' Client Addess Line 3
    sPhone             As String * 20     ' Phone
    sFax               As String * 20     ' Fax Number
    iDaysRetainAsAir   As Integer         ' Number of days to retain schedule an
                                          ' d As Air
    iDaysRetainActive  As Integer         ' Number of days to retain Active Log
    lChgInterval       As Long            ' number of seconds to allow changes w
                                          ' ith current time
    sMergeDateFormat   As String * 20     ' Merge Date Format assocaited with Me
                                          ' rge file name
    sMergeTimeFormat   As String * 20     ' Merge time format assocaited with Fi
                                          ' le Name
    sMergeFileFormat   As String * 20     ' Merge File Name Format
    sMergeFileExt      As String * 3      ' Merge File Name Extension
    sMergeStartTime    As String * 11     ' Merge Start Time
    sMergeEndTime      As String * 11     ' Merge End Time
    iMergeChkInterval  As Integer         ' Merge Interval for Checking (in Minu
                                          ' tes
    sMergeStopFlag     As String * 1      ' Merge Stop Flag (Y = Yes, N= No)
    iAlertInterval     As Integer         ' Aert Interval (In Minutes)
    sSchAutoGenSeq     As String * 1      ' Generate Order of Schedule and Autom
                                          ' ation: I=Independent (each has its o
                                          ' wn time): S=Sch, then Auto; A=Auto,
                                          ' then Sch
    lMinEventID        As Long            ' Min value to be used for Event ID
    lMaxEventID        As Long            ' Max Event ID
    lCurrEventID       As Long            ' Current Event ID
    iNoDaysRetainPW    As Integer         ' Number of days password can be retai
                                          ' ned.
    iVersion           As Integer         ' Version (Starting at 0)
    iOrigSoeCode       As Integer         ' Original Site Option Code used to ti
                                          ' e all versions together
    sCurrent           As String * 1      ' Current Version: Y=Yes; N=No.  Y sho
                                          ' uld only be with the highest version
                                          '  #
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    iSpotItemIDWindow  As Integer         ' Spot Item ID Test window in
                                          ' Milliseconds
    lTimeTolerance     As Long            ' +,- Time Tolerance for 'As Aired'
                                          ' compare (mm:ss.t)
    lLengthTolerance   As Long            ' +,- Length Tolerance for 'As
                                          ' Aired' compare (mm:ss.t)
    sMatchATNotB       As String * 1      'Conflict Test: Match Audio, Time but not Bus- Y or N.  Default is Y
    sMatchATBNotI      As String * 1      'Conflict Test: Match Audio, Time, Bus but not Item ID- Y or N.  Default is Y
    sMatchANotT        As String * 1      'Conflict Test: Match Audio But not Time but times overlap- Y or N.  Default is Y
    sMatchBNotT        As String * 1      'Conflict Test: Match Bus but not times and times overlap- Y or N.  Default is Y
    sSchAutoGenSeqTst  As String * 1      ' Generate Order of Schedule and Autom
                                          ' ation: I=Independent (each has its o
                                          ' wn time): S=Sch, then Auto; A=Auto,
                                          ' then Sch
    sMergeStopFlagTst  As String * 1      ' Merge Stop Flag (Y = Yes, N= No)
    sUnused            As String * 4      ' Unused buffer
End Type

'SPE_Site_Path
Type SPE
    iCode              As Integer         ' Internal Reference Code
    iSoeCode           As Integer         ' Site Option Reference
    sType              As String * 2      ' Commercial Merge: SP=Server-Primary;
                                          '  SB-Server-Backup; CP=Client-Primary
                                          ' ; CB=Client Backup
    sSubType           As String * 1      ' P or Blank = Production; T=Test
    sPath              As String * 100    ' Merge Path
    sUnused            As String * 20     ' Unused buffer
End Type

Type SSE
    iCode              As Integer            ' Internal Reference Code
    iSoeCode           As Integer            ' Site Option Reference
    sEMailHost         As String * 80        ' HostName/SMTP server address or
                                             ' URL
    iEMailPort         As Integer            ' Port number for SMTP server
    sEMailAcctName     As String * 80        ' Account name for SMTP credentials
    sEMailPassword     As String * 80        ' Password for the SMTP credentials
    sEMailTLS          As String * 1         ' Transport Layer sercurity (Y/N)
    sUnused            As String * 10        ' Unused
End Type

Type TNE
    iCode              As Integer         ' Internal Reference Code
    sType              As String * 1      ' Type: J=Job; L=List; A=Alert; N=Noti
                                          ' fication
    sName              As String * 30     ' Task Name
    sUnused            As String * 20     ' Unused Buffer
End Type

Type TSE
    lCode              As Long            ' Internal Reference Code
    lDheCode           As Long            ' Day Event Header reference
    iBdeCode           As Integer         ' Bus Definition reference
    sLogDate           As String * 10     ' Log date
    sStartTime         As String * 11     ' Start Time of Template library
    sDescription       As String * 50     ' Template Date/Time description
    sState             As String * 1      ' State: A=Active; D=Dormant
    lCteCode           As Long            ' Comment (100 characters)
    iVersion           As Integer         ' Version (Starting at 0)
    lOrigTseCode       As Long            ' Original Template Schedule Code used
                                          '  to tie all versions together
    sCurrent           As String * 1      ' Current: Y=Yes; N=No.  Y should only
                                          '  be with the highest version #
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

Type TTE
    iCode              As Integer         ' Internal Reference Code
    sType              As String * 1      ' Type: S=Start Type; E=End Type
    sName              As String * 3      ' Time Type Name
    sDescription       As String * 50     ' Time Type Description
    sState             As String * 1      ' State: A=Active; D=Dormant
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Starting at 0)
    iOrigTteCode       As Integer         ' Original Time Type used to tie all v
                                          ' ersions together
    sCurrent           As String * 1      ' Current Version: Y=Yes; N=No
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type

Type UIE
    iCode              As Integer         ' Internal Reference Code
    sSignOnName        As String * 40     ' Sign On Name
    sPassword          As String * 10     ' Password
    sLastDatePWSet     As String * 10     ' Date Password last set
    sShowName          As String * 40     ' Show Name
    sState             As String * 1      ' State: A=Active; D=Dormant
    sEMail             As String * 70     ' E-Mail Address
    sLastSignOnDate    As String * 10     ' Last Sign On Date
    sLastSignOnTime    As String * 11     ' Last Sign On Time
    sUsedFlag          As String * 1      ' Used Flag: Y=Yes; N=No
    iVersion           As Integer         ' Version (Start at 0)
    iOrigUieCode       As Integer         ' Original User Info Code used to tie
                                          ' all versions together
    sCurrent           As String * 1      ' Current Version: Y=Yes; N=No
    sEnteredDate       As String * 10     ' Entered Date
    sEnteredTime       As String * 11     ' Entered Time
    iUieCode           As Integer         ' User Reference (User that altered or
                                          '  added record)
    sUnused            As String * 20     ' Unused buffer
End Type


Type UTE
    iCode              As Integer         ' Internal Reference Code
    iUieCode           As Integer         ' User Information Reference
    iTneCode           As Integer         ' Task Name Reference
    sTaskStatus        As String * 1      ' Job/List: E=Edit; V=View; D=Disable.
                                          '   Alerts/Notications: Y=Yes; N=No
    sUnused            As String * 20     ' Unused buffer
End Type

Type DDFFILENAMES
    sShortName As String * 4                'file name such as vef, shtt, etc.
    sLongName As String * 20                'full file name i.e. vef_vehicles.  when converting from btrieve to odbc drivers,
                                            'the full filename is required in locations field
End Type

Type REPORTNAMES
    sRptName As String                      'report name in list box
    sCrystalName As String
    iRptIndex As Integer                    'report index
    sRptPicture As String                 'report .bmp, .jpg showing sample of report
    sRptDesc As String                    'report description
End Type

'Filter
Type FILTERVALUES
    sFieldName As String * 20
    iOperator As Integer        '1=Equal; 2=Not Equal; 3=Greater Than; 4=Less Than; 5=Greater than or equal to; 6=Less than or equal to
    sValue As String
    lCode As Long               'File Code Value
    iUsed As Integer            'True=Field used when checking filter condition
End Type

Type FIELDSELECTION          'FILTERFIELDS
    sFieldName As String * 20
    iFieldType As Integer '1=Number; 2=String; 3=Date; 4=Time; 5=List; 6=Time in Tengths; 7=Length; 8=Length in Tenths; 9=Match List
    iMaxNoChar As Integer   'Max number of characters
    sListFile As String * 4 'File Name is iFilterType = 5
    sMandatory As String * 1
End Type

Type MATCHLIST
    sValue As String
    lValue As Long
End Type

Type SCHDREPLACEVALUES
    sFieldName As String * 20
    sOldValue As String
    lOldCode As Long               'File Code Value
    sNewValue As String
    lNewCode As Long               'File Code Value
End Type

Type LIBREPLACEVALUES
    sFieldName As String * 20
    sBuses As String
    sHours As String
    sOldValue As String
    lOldCode As Long               'File Code Value
    sNewValue As String
    lNewCode As Long               'File Code Value
End Type

Type ITEMIDCHK
    sItemID As String * 32
    sTitle As String * 35
    lLength As Long
    sAudioStatus As String * 1
    sProtStatus As String * 1
    sPriResult As String * 35
    sPriLen As String * 5
    sProtResult As String * 35
    sProtLen As String * 5
    lSeeCode As Long
End Type

Type LOADUNCHGDEVENT
    iBdeCode As Integer
    lFirstSEECode As Long   'First unchanged event prior to first changed event
    lLastSEECode As Long   'First unchanged event after last changed event
    sSendStatus As String * 1   'N=Not at First event to send; S=Send; F=Finished with Send
    iLastMsgGen As Integer  'Unable to find last unchanged event message generated
End Type

Type SCHDCHGINFO
    lNewChgDHE As Long      'New or Changed DHE (Check or Create DEE)
    lCheckDHE As Long       'Check DEE for valid dates, if invalid see if part of the Split
    lSplitDHE As Long       'DHE created from a Split Overlap, change DEE to reference this DHE instead of the lCheckDHE
    lExpandDHE As Long      'Add Events because DHE dates expanded
    lDEEDHE As Long         'DHE value DEE converted to when NewChg is Chg and updated
End Type

Type DHETSE
    tDHE As DHE
    tTSE As TSE
End Type

Type CONFLICTLIST
    sType As String * 1 'S=Schedule; L=Library; T=Template; E=Event
    lSheCode As Long
    lSeeCode As Long
    lDheCode As Long
    lDseCode As Long
    lDeeCode As Long
    lIndex As Long       'sType = E only. In SchdDef this is CurrSeeIndex; in LibDef and TempDef this is RowIndex
    sStartDate As String * 10
    sEndDate As String * 10
    iNextIndex As Integer
End Type

Type MERGEINFO
    lAvailRunTime As Long
    iSpotSoldTime As Integer
    lLastSpotAddedIndex As Long     'Used to remove end netque
End Type

Type SCHDEXTRACT
    sEventType As String * 1
    sBus As String * 8
    sBusCtrl As String * 1
    sTime As String * 11
    sStartType As String * 3
    sEndType As String * 3
    sDuration As String * 11
    sMaterialType As String * 3
    sAudioName As String * 8
    sAudioID As String * 32
    sAudioISCI As String * 20
    sAudioCtrl As String * 1
    sBackupName As String * 8
    sBackupCtrl As String * 1
    sProtName As String * 8
    sProtItemID As String * 32
    sProtISCI As String * 20
    sProtCtrl As String * 1
    sRelay1 As String * 8
    sRelay2 As String * 8
    sFollow As String * 19
    sSilenceTime As String * 5
    sSilence1 As String * 1
    sSilence2 As String * 1
    sSilence3 As String * 1
    sSilence4 As String * 1
    sNetcue1 As String * 3
    sNetcue2 As String * 3
    sTitle1 As String * 66
    sTitle2 As String * 90
    sABCFormat As String * 1
    sABCPgmCode As String * 25
    sABCXDSMode As String * 2
    sABCRecordItem As String * 5
    sFixedTime As String * 1
    sDate As String * 10
    sEndTime As String * 11
    sEventID As String * 10
    sHours As String * 24
    sDays As String * 7
    sOffset As String * 7
    lOffset As Long
    lRunningTime As Long
    lLinkBus As Long
End Type

Type CONFLICTRESULTS
    tCEE As CEE
    tCME As CME
End Type

Type INTKEY0
    iCode As Integer
End Type

Type LONGKEY0
    lCode As Long
End Type

Type SCHDSORT
    sKey As String * 20     'Time(6)|Bus(8)|category with Program first
    lRow As Long
End Type

Type NAMESORT
    sKey As String * 20
    lCode As Long
    iCode As Integer
End Type

Type CTESORT
    sKey As String * 66
    lCode As Long
End Type

Type DEECTE
    lDeeCode As Long
    lCteCode As Long
    lDheCode As Long
    sComment As String * 66
End Type
