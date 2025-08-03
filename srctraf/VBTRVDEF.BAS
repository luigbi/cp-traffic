'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of VBTRVDEF.BAS on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Constants (Marked)                                                              *
'*  BTRV_LOCK_MULT_NOWAIT         BTRV_UNLOCK_MULTIPLE                                    *
'******************************************************************************************

' Copyright 1992, 1993 Classic Software, Inc. All rights reserved.

'
' File Name: VBTRVDEF.TXT
'
' Release: 2.0
'
' Description:
'   This provides a lot of useful definitions for use with Btrieve.




' File Flag Defines

'Global Const BTRV_FF_VARIABLE = 1
'Global Const BTRV_FF_BLANK_TRUNC = 2
'Global Const BTRV_FF_PREALLOCATION = 4
'Global Const BTRV_FF_DATA_COMPRESS = 8
'Global Const BTRV_FF_KEY_ONLY = 16
'Global Const BTRV_FF_FREE_SPACE10 = 64
'Global Const BTRV_FF_FREE_SPACE20 = 128
'Global Const BTRV_FF_FREE_SPACE30 = 192

' Key Flag Defines

'Global Const BTRV_KF_DUPLICATE = 1
'Global Const BTRV_KF_MODIFIABLE = 2
'Global Const BTRV_KF_BINARY = 4
'Global Const BTRV_KF_NULL = 8
'Global Const BTRV_KF_SEGMENTED = 16
'Global Const BTRV_KF_ALT_COL_SEQ = 32
'Global Const BTRV_KF_DESCENDING = 64
'Global Const BTRV_KF_SUPPLEMENTAL = 128
'Global Const BTRV_KF_EXTENDED = 256
'Global Const BTRV_KF_MANUAL = 512

' Key type defines

Global Const BTRV_KT_STRING = 0
Global Const BTRV_KT_INT = 1
'Global Const BTRV_KT_FLOAT = 2
Global Const BTRV_KT_DATE = 3
Global Const BTRV_KT_TIME = 4
'Global Const BTRV_KT_DECIMAL = 5
'Global Const BTRV_KT_LOGICAL = 7
'Global Const BTRV_KT_BFLOAT = 9
'Global Const BTRV_KT_UBINARY = 14

' Lock Types

Global Const BTRV_LOCK_NONE = 0
Global Const BTRV_LOCK_MULT_NOWAIT = 400 'VBC NR

Global Const BTRV_UNLOCK_MULTIPLE = -2 'VBC NR

' Open lock constants for the control.

' File creation flag

'Global Const BTRV_CRE_OVERWRITE = 0

' Open file Modes

Global Const BTRV_OPEN_NORMAL = 1
Global Const BTRV_OPEN_VERIFY = -3

' Open mode constants for the control.
'Global Const BTRCTL_OPEN_NORMAL = 5
'Global Const BTRCTL_OPEN_NONE = 4
'Global Const BTRCTL_OPEN_ACCEL = 3
'Global Const BTRCTL_OPEN_RDONLY = 2
'Global Const BTRCTL_OPEN_VERIFY = 1
'Global Const BTRCTL_OPEN_EXCLUSIVE = 0

' Open shareable const. When you open a file you can specify that
'   it be registered as shareable with BTRV_OPEN_SHARE.
Global Const BTRV_OPEN_NONSHARE = 0
Global Const BTRV_OPEN_SHARE = 1

' Owner Access and Encryption Codes

'Global Const BTRV_OWNER_NO_ENCRYPT = 0
'Global Const BTRV_RDONLY_NO_ENCRYPT = 1
'Global Const BTRV_OWNER_ENCRYPT = 2
'Global Const BTRV_RDONLY_ENCRYPT = 3

' Extended operations defines

' Extended operation comparison codes

Global Const BTRV_EXT_EQUAL = 1
Global Const BTRV_EXT_GT = 2
Global Const BTRV_EXT_NOT_EQUAL = 4
Global Const BTRV_EXT_GTE = 5
Global Const BTRV_EXT_LTE = 6

' Extended operation comparison  code bias values


' Extended operation logic expressions

Global Const BTRV_EXT_LAST_TERM = 0
Global Const BTRV_EXT_AND = 1
Global Const BTRV_EXT_OR = 2

