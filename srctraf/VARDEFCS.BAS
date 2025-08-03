Attribute VB_Name = "VARDEFCS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of VARDEFCS.BAS on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: VarDefCS.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Global variables for contract scheduling
Option Explicit
Public tgSsfSrchKey As SSFKEY0      'SSF key record image
Public tgSsfSrchKey1 As SSFKEY1      'SSF key record image
Public tgSsfSrchKey2 As SSFKEY2      'SSF key record image
Public lgSsfDate(0 To 6) As Long    'Dates of the days stored into tgSsf
Public lgSsfRecPos(0 To 6) As Long  'Record positions
Public tgSsf(0 To 6) As SSF         'Spot summary for one week (0 index for monday;
                                    '1 for tuesday;...; 6 for sunday)
Public tgSdfSrchKey As SDFKEY0      'SDF key record image
Type SDFSCHEXT
    iDate(0 To 1) As Integer        'Schedule or missed Date of spot
    iTime(0 To 1) As Integer        'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sSpotType As String * 1
End Type
Public Const SDFSCHEXTPK As String = "IIIIB"
'Spots to be added as if part of the avail and spots that are to be ignored
'This type is used during shuffling of spots and testing if a spot requires
'preempting because of a line change
Type SPOTMOVE
                                    'From
    iSpotIndex As Integer           'Index of spot within Ssf (to be bypasses)
    lSpotSsfRecPos As Long          'Ssf Record position
    lSdfCode As Long                'Sdf auto code
                                    'To
    iAvailIndex As Integer          'Index of avail where spot might be moved (-1=no avail index)
    lAvailSsfRecPos As Long         'Ssf record position of the above avail
    iMnfComp(0 To 1) As Integer     'Competitives of spot being considered of moving
    iAdfCode As Integer             'Advertiser of spot being considered for move
End Type
Public tgSpotMove() As SPOTMOVE     'array of spots that are to be bypasses during spot testing
                                    'and/or moved
'Spots requiring scheduling
Public lgUnschIndex As Long
'Public lgUnschRecPos() As Long  'Record position for spots created and require scheduling
Public lgUnschSdfCode() As Long  'Record position for spots created and require scheduling
Public lgReschIndex As Long
Public lgReschSdfCode() As Long  'Record code for spots preempted that require rescheduling
Public sgSSFErrorMsg() As String  'Error messages from gMakeSSF and gSynchSdfSsf
'Scheduling parameters
Public igRnd As Integer        'True=Seed random; False=Repeat number sequence
'Buy 6-10; Mo-We; Currently 3 spots schedule
'  Hour   Count      Day   Count        Q-H   Count
'   6       2        Mo      1          1st     2
'   7       0        Tu      0          2nd     0
'   8       0        We      2          3rd     1
'   9       1                           4th     0
'Cycle pattern 1
'Hour 8 8 7 7 8 7 8 7 8 8 7 7 8 7 8 7 8 8 7 7 8 7 8 7 -- 6
'Day  T T T T T T T T M M M M M M M M W W W W W W W W    W
'Q-H  4 2 4 2 3 3 1 1 4 2 4 2 3 3 1 1 4 2 4 2 3 3 1 1    1
'     A A A A A A A A A A A A A A A A A A A A A A A A    A
'           B   B   B       B   B   B       B   B   B    B
'                   C               C               C    C
'                                                   D    D
'                                                        E
'
'Cycle pattern 2
'Hour 8 8 8 8 8 8 8 8 8 8 8 8 7 7 7 7 -- 6
'Day  T T T T M M M M W W W W T T T T    W
'Q-H  4 2 3 1 4 2 3 1 4 2 3 1 4 2 3 1    1
'     A A A A A A A A A A A A A A A A    A
'       B B B   B B B   B B B   B B B    B
'           C       C       C       C    C
'                           D            C
'                                        E
Public sgMajor As String * 1   'H=Hours; D=Days; Q=Quarter-Hours
Public sgMiddle As String * 1  'H=Hours; D=Days; Q=Quarter-Hours
Public sgMinor As String * 1   'H=Hours; D=Days; Q=Quarter-Hours
Public igCyclePattern As Integer   '1=Cycle pattern 1
                                '2=Cycle pattern 2
'Which test is to be performed during which iternation
'A is at each iternation of major, middle and minor
'B is at each time the count changes within minor
'C is at each minor iternation
'D is at each middle iternation
'E is at each major iternation
Public sgInsertPass As String * 1  'Point: A
Public sgMovePass As String * 1    'Point: A or B or C or D or E or N
Public sgCompPass As String * 1    'Point: A or B or C or D or E or N
Public sgPreemptPass As String * 1 'Point: A or B or C or D or E or N
Public igReschNoPasses As Integer    'When rescheduling preempted or missed spots, schedule in one pass(1) or two(2)
Public lgNoBonusRemoved As Long   'Number of bonus spots when removed
Public igSetEarliestDateAsToday As Integer  '1=No; 2=Yes
Public igUsePreferred As Integer    'True=Use preferred days/times when rescheduling spots (this is only set to true in Rectify)
'If two, pass 1 look for open avails only, pass2 allow preemption
'Sort; Randomize- N=No sorting; No randomizing
'                 R=No sorting; randomize all counts
'                 S=Sorting counts lowest to highest; No randomizing of matched counts
'                 B=Sorting counts lowest to highest; Randomize matching counts
Public sgSortRndHour As String * 1 'Sort hours (lowest count to highest count with matched count randomily mixed)
Public sgSortRndDay As String * 1  'Sort days (lowest count to highest count with matched count randomily mixed)
Public sgSortRndQH As String * 1   'Sort quarter hours (lowest count to highest count with matched count randomily mixed)
'Program scheduling
Public tgAvail As AVAILSS
Public tgSpot As CSPOTSS

Public tgSaf() As SAF   'Array of SAF records

Public tgPdf() As PDF

