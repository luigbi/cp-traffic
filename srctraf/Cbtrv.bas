Attribute VB_Name = "CBTRV"
Declare Sub vbPackUDT Lib "VBHLP32.DLL" (pUDT As Any, ppResult As Long, ByVal pszFields As String)
Declare Sub vbUnpackUDT Lib "VBHLP32.DLL" (pUDT As Any, ppResult As Long)
Declare Function vbPackUDTGetSize Lib "VBHLP32.DLL" (ppResult As Long) As Long
Declare Sub vbPackUDTFree Lib "VBHLP32.DLL" (ppResult As Long)


' Copyright 1992, 1993 Classic Software, Inc. All rights reserved.
' File Name: CBTRV.BAS
' Release: 2.1
' Description:
' This file contains the Basic interface to the cbtrv432.dll
' C library for Btrieve. This library is a standard Windows
' DLL callable from any process capable of calling a Windows DLL.
' This includes both VB and MS Access. This library is used by the
' VBX controls for their Database Access.
'
' The functions detailed
' within this library can be called by accessing the "Ohnd" property
' of the File, Table, and List VBX controls.
' For example assume "VBtrv1" is the name of a control. then
' ret = dbGetFirst(VBtrv1.Ohnd, 0, 0) will call the Get first
' function and fill the internal record buffer of the control.
' This is because the Ohnd property is a handle to the live object in use
' by the control for access to the library.
'
' The Library is Object oriented with multiple inheritance.
' Access is provided via "Object Handles". This is simply a
' "Majic Number" that acts as an intelligent position block.
' Construction of a CBtrvTable Handle will provide access to every
' function in this library.
' See the CBtrv documentation for further details.
Public Const DDF_NAMESIZE = 20
Public Const FILE_LOCATIONSIZE = 64
' DDF Record structure definitions
Type rFILEREC
   Id As Integer
   Name As String * 20
   Location As String * 64
   Flags As String * 1
   Reserved As String * 10
End Type
Type rFIELDREC
   Id As Integer
   FileId As Integer
   Name As String * 20
   DataType As String * 1
   Offset As Integer
   Size As Integer
   DecPlaces As String * 1
   Flags As Integer
End Type
Type rINDEXREC
   FileId  As Integer
   FieldId As Integer
   Number As Integer
   Part As Integer
   Flags As Integer
End Type
' btrStopAppl cleans up all handles for your application. It forces files
' closed, ends transactions and destroys all handles in use.
' You must destroy all handles on program termination.
' This function will perform all required cleanup.
' The Alternative is to perform a btrDestroy on each handle individually.
Declare Sub btrStopAppl Lib "cbtrv432.dll" ()
Declare Sub btrDestroy Lib "cbtrv432.dll" (ByVal Ohnd%)
' CBtrvOapi provides all of the basic Btrieve functionality.
'Declare Function CBtrvOapi Lib "cbtrv432.dll" () As Integer
'Declare Function CBtrvOapiOpen Lib "cbtrv432.dll" (ByVal OwnerName$, ByVal FileName$, ByVal fOpenMode%, ByVal Shareable%, ByVal LockBias%) As Integer
' CBtrvObj adds file creation and extended operations.
'Declare Function CBtrvObj Lib "cbtrv432.dll" () As Integer
'Declare Function CBtrvObjOpen Lib "cbtrv432.dll" (ByVal OwnerName$, ByVal FileName$, ByVal fOpenMode%, ByVal Shareable%, ByVal LockBias%) As Integer
Declare Sub btrCreClear Lib "cbtrv432.dll" (ByVal Ohnd%)
Declare Sub btrCreClearCollate Lib "cbtrv432.dll" (ByVal Ohnd%)
Declare Sub btrCreClearFile Lib "cbtrv432.dll" (ByVal Ohnd%)
Declare Sub btrCreClearKey Lib "cbtrv432.dll" (ByVal Ohnd%)
Declare Function btrCreCollate Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal SequenceName$, CollateSeq As Any) As Integer
Declare Function btrCreCollateFile Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal CollateFileName$) As Integer
Declare Function btrCreFile Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal RecLength&, ByVal PageSize%, ByVal NumIndexes%, ByVal uFileFlags%, ByVal uAllocation%) As Integer
Declare Function btrCreKey Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyPosition&, ByVal KeyLength&, ByVal KeyFlags%, ByVal ExtendKeyType%, ByVal NullValue%) As Integer
Declare Sub btrExtClear Lib "cbtrv432.dll" (ByVal Ohnd%)
'Declare Sub btrExtClearExtend Lib "cbtrv432.dll" (ByVal Ohnd%)
'Declare Sub btrExtClearFields Lib "cbtrv432.dll" (ByVal Ohnd%)
'Declare Sub btrExtClearInsert Lib "cbtrv432.dll" (ByVal Ohnd%)
'Declare Sub btrExtClearLogic Lib "cbtrv432.dll" (ByVal Ohnd%)
Declare Function btrExtAddField Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldOffset&, ByVal FieldLength&) As Integer
Declare Function btrExtAddInsert Lib "cbtrv432.dll" (ByVal Ohnd%, RecordImage As Any, ByVal RecordLength&) As Integer
Declare Function btrExtAddLogic Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal DataType%, ByVal FieldOffset&, ByVal FieldLength&, ByVal ComparisonCode%, ByVal AndOrLogic%, ByVal SecondFieldOffset&) As Integer
Declare Function btrExtAddLogicConst Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal DataType%, ByVal FieldOffset&, ByVal FieldLength&, ByVal ComparisonCode%, ByVal AndOrLogic%, ConstField As Any, ByVal ConstSize&) As Integer
Declare Function btrExtAddLogicString Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal DataType%, ByVal FieldOffset&, ByVal FieldLength&, ByVal ComparisonCode%, ByVal AndOrLogic%, ByVal ConstField$, ByVal ConstSize&) As Integer
'Declare Sub btrExtGetBounds Lib "cbtrv432.dll" (ByVal Ohnd%, MaxRetrieved As Integer, MaxSkipped As Integer, ExtractSize As Integer)
Declare Function btrExtGetCurrent Lib "cbtrv432.dll" (ByVal Ohnd%, Record As Any, RecordSize As Long, RecordPosition As Any) As Integer
Declare Function btrExtGetFirst Lib "cbtrv432.dll" (ByVal Ohnd%, Record As Any, RecordSize As Long, RecordPosition As Any) As Integer
Declare Function btrExtGetLast Lib "cbtrv432.dll" (ByVal Ohnd%, Record As Any, RecordSize As Long, RecordPosition As Any) As Integer
Declare Function btrExtGetNext Lib "cbtrv432.dll" (ByVal Ohnd%, Record As Any, RecordSize As Long, RecordPosition As Any) As Integer
'Declare Function btrExtGetPosition Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
Declare Function btrExtGetPrevious Lib "cbtrv432.dll" (ByVal Ohnd%, Record As Any, RecordSize As Long, RecordPosition As Any) As Integer
'Declare Function btrExtGetRecord Lib "cbtrv432.dll" (ByVal Ohnd%, Record As Any, RecordSize As Integer, RecordPosition As Any, ByVal RecordNumber%) As Integer
'Declare Function btrExtRestorePosition Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrExtSavePosition Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrExtUpdate Lib "cbtrv432.dll" (ByVal Ohnd%, Record As Any, RecordSize As Integer) As Integer
'Declare Function btrExtRecsInserted Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
Declare Function btrExtRecsReturned Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
' Use either EG or UC for header control. EG excludes current record. UC includes current record.
Declare Sub btrExtSetBounds Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal MaxRetrieved%, ByVal MaxSkipped%, ByVal HeaderControl$)
'Declare Function btrCurrIndex Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Sub btrFilePath Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FileName$)
'Declare Function btrIndexes Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrKeySize Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyNumber%) As Integer
Declare Function btrOpenFiles Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrOpenMode Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Sub btrOwnerName Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal OwnName$)
Declare Function btrRecordLength Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
Declare Function btrRecords Lib "cbtrv432.dll" (ByVal Ohnd%) As Long
'Declare Function btrTransInProgress Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
Declare Function btrAbortTrans Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
Declare Function btrBeginTrans Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal LockBias%) As Integer
Declare Function btrClear Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrClearOwner Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
Declare Function btrClone Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal NewFileName$, ByVal FileFlag%) As Integer
Declare Function btrClose Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrCloseAll Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrContinuous Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Names$, ByVal Mode%) As Integer
'Declare Function btrCreate Lib "cbtrv432.dll" (ByVal Ohnd%, FileStructure As Any, StructLength As Integer, ByVal FileName$, ByVal fOverWrite%) As Integer
Declare Function btrCreCreate Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FileName$, ByVal fOverWrite%) As Integer
'Declare Function btrCreateSupplIndex Lib "cbtrv432.dll" (ByVal Ohnd%, KeySpecs As Any, SpecLength As Integer, KeyNumber As Integer) As Integer
'Declare Function btrCreCreateSupplIndex Lib "cbtrv432.dll" (ByVal Ohnd%, KeyNumber As Integer) As Integer
Declare Function btrDelete Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrDropSupplIndex Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyNumber%) As Integer
Declare Function btrEndTrans Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
Declare Function btrExtend Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal ExtendFilePath$, ByVal fStoreFlag%) As Integer
Declare Function btrFFlush Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrFindPercentage Lib "cbtrv432.dll" (ByVal Ohnd%, Percentage As Long, KeyBuffer As Any, ByVal KeyNumber%) As Integer
'Declare Function btrFindCurrPercentage Lib "cbtrv432.dll" (ByVal Ohnd%, Percentage As Long, ByVal KeyNumber%) As Integer
'Declare Function btrGetChunk Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long) As Integer
Declare Function btrGetCurrent Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long) As Integer
Declare Function btrGetCurrentKey Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, ByVal KeyNumber%, ByVal LockBias%) As Integer
Declare Function btrGetData Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Operation%, DataBuffer As Any, DataLength As Long, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
Declare Function btrGetDirect Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, RecordPosition As Any, ByVal KeyNumber%, ByVal LockBias%) As Integer
'Declare Function btrGetDirectory Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal DirPath$, ByVal DriveNumber%) As Integer
Declare Function btrGetEqual Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
Declare Function btrGetFirst Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
Declare Function btrGetGreater Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
Declare Function btrGetGreaterOrEqual Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
Declare Function btrGetKey Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Operation%, KeyBuffer As Any, ByVal KeyNumber%) As Integer
Declare Function btrGetLast Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
Declare Function btrGetLess Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
Declare Function btrGetLessOrEqual Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
Declare Function btrGetNext Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, ByVal Lock_GetKey%) As Integer
Declare Function btrGetNextExt Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long) As Integer
Declare Function btrExtGetNextExt Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
Declare Function btrExtGetRecPos Lib "cbtrv432.dll" (ByVal Ohnd%) As Long
'Declare Function btrGetPercentage Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Percentage&, DataBuffer As Any, DataLength As Long, ByVal KeyNumber%) As Integer
Declare Function btrGetPosition Lib "cbtrv432.dll" (ByVal Ohnd%, RecordPosition As Any) As Integer
Declare Function btrGetPrevious Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, ByVal Lock_GetKey%) As Integer
'Declare Function btrGetPreviousExt Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long) As Integer
'Declare Function btrExtGetPreviousExt Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrGetPrivateDir Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal DirPath$) As Integer
'Declare Function btrGetStep Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Operation%, DataBuffer As Any, DataLength As Long) As Integer
Declare Function btrInsert Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, ByVal KeyNumber%) As Integer
'Declare Function btrInsertExt Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, ByVal KeyNumber%) As Integer
'Declare Function btrExtInsertExt Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyNumber%) As Integer
'Declare Function btrMakeCurrKey Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyNumber%) As Integer
'Declare Function btrOBTRV Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Operation%, DataBuffer As Any, DataLength As Any, KeyBuffer As Any, ByVal KeyNumber%) As Integer
Declare Function btrOpen Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal OwnerName$, ByVal FileName$, ByVal fOpenMode%, ByVal Shareable%, ByVal LockBias%) As Integer
'Declare Sub btrPrint Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal AppendData%)
Declare Function btrReset Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrResetStation Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Workstation%, ByVal fReset%) As Integer
'Declare Function btrRestartApi Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrRestorePosition Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrRestoreKeyPosition Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyNumber%) As Integer
'Declare Function btrSavePosition Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrSetDirectory Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal DirPath$) As Integer
'Declare Function btrSetOwner Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal OwnerName$, ByVal SameOwnerName$, ByVal fAccess%) As Integer
'Declare Function btrSetPrivateDir Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal DirPath$) As Integer
Declare Function btrShutdownApi Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Abort%) As Integer
'Declare Function btrStat Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long) As Integer
'Declare Function btrStatExtension Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, ByVal ExtensionFile$) As Integer
'Declare Function btrStepFirst Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, ByVal LockBias%) As Integer
'Declare Function btrStepLast Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, ByVal LockBias%) As Integer
'Declare Function btrStepNext Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, ByVal LockBias%) As Integer
'Declare Function btrStepNextExt Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long) As Integer
'Declare Function btrExtStepNextExt Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrStepPrevious Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long, ByVal LockBias%) As Integer
'Declare Function btrStepPreviousExt Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long) As Integer
'Declare Function btrExtStepPreviousExt Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
Declare Function btrStop Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
Declare Function btrUnlock Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal fLock%) As Integer
Declare Function btrUnlockRecord Lib "cbtrv432.dll" (ByVal Ohnd%, RecordPosition As Any) As Integer
Declare Function btrUpdate Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long) As Integer
'Declare Function btrUpdateChunk Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long) As Integer
'Update and maintain logical position in file. Useful for modifiable keys.
'Declare Function btrUpdateLogical Lib "cbtrv432.dll" (ByVal Ohnd%, DataBuffer As Any, DataLength As Long) As Integer
Declare Function btrVersion Lib "cbtrv432.dll" (ByVal Ohnd%, VersionData As Any) As Integer
' CBtrvMngr provides for Record Manager initialization and shutdown.
Declare Function CBtrvMngr Lib "cbtrv432.dll" (ByVal OptionString$) As Integer
Declare Function CBtrvMngrInit Lib "cbtrv432.dll" (ByVal InitType%) As Integer
Declare Function btrMngrsInUse Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
Declare Function btrInit Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal OptionString$) As Integer
'Declare Function btrLRType Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Sub btrOptionStr Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal OptionStr$)
'Declare Function btrReqInit Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal OptionString$) As Integer
'Declare Function btrRestart Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function btrRestartOptions Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal OptionString$) As Integer
Declare Function btrShutdown Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Abort%) As Integer
' CBtrvDB takes on management of your record buffer internally. It will also
' build key buffers of of the internal record buffer for Keyed operations
' like Get Equal.
'Declare Function CBtrvDB Lib "cbtrv432.dll" () As Integer
'Declare Function dbGetCurrent Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function dbGetEqual Lib "cbtrv432.dll" (ByVal Ohnd%, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
'Declare Function dbkGetEqual Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
'Declare Function dbGetFirst Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
'Declare Function dbGetGreater Lib "cbtrv432.dll" (ByVal Ohnd%, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
'Declare Function dbkGetGreater Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
'Declare Function dbGetGreaterOrEqual Lib "cbtrv432.dll" (ByVal Ohnd%, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
'Declare Function dbkGetGreaterOrEqual Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
'Declare Function dbGetLast Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
'Declare Function dbGetLess Lib "cbtrv432.dll" (ByVal Ohnd%, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
'Declare Function dbkGetLess Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
'Declare Function dbGetLessOrEqual Lib "cbtrv432.dll" (ByVal Ohnd%, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
'Declare Function dbkGetLessOrEqual Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
'Declare Function dbGetNext Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Lock_GetKey%) As Integer
'Declare Function dbGetPercentage Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Percentage&, ByVal KeyNumber%) As Integer
'Declare Function dbGetPrevious Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Lock_GetKey%) As Integer
'Declare Function dbInsert Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyNumber%) As Integer
'Declare Function dbStepFirst Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal LockBias%) As Integer
'Declare Function dbStepLast Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal LockBias%) As Integer
'Declare Function dbStepNext Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal LockBias%) As Integer
'Declare Function dbStepPrevious Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal LockBias%) As Integer
'Declare Function dbUpdate Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function dbGetRecord Lib "cbtrv432.dll" (ByVal Ohnd%, Dest As Any, ByVal dSize%) As Integer
'Declare Function dbSetField Lib "cbtrv432.dll" (ByVal Ohnd%, Src As Any, ByVal sSize%, ByVal dOffset%) As Integer
'Declare Function dbSetKey Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyNumber%, KeyValue As Any) As Integer
'Declare Function dbSetKeySeg Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyNumber%, ByVal SegNumber%, KeyValue As Any) As Integer
'Declare Function dbSetRecord Lib "cbtrv432.dll" (ByVal Ohnd%, sSource As Any, ByVal sSize%) As Integer
'Declare Function dbExtAddInsert Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
' CBtrvTable adds field based access to your data via DDF Files.
' The functions use the internal record buffer for getting and setting
' data by field name. Use the CBtrvDB (db) function for reading data
' to and from the database into the internal record buffer.
Declare Function CBtrvTable Lib "cbtrv432.dll" () As Integer
'Declare Function ddOpenDDFTable Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal DDFPath$, ByVal TableName$, ByVal OwnerName$, ByVal FileName$, ByVal fOpenMode%, ByVal Shareable%, ByVal LockBias%) As Integer
'Declare Function ddExtAddField Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$) As Integer
'Declare Function ddExtAddLogic Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$, ByVal ComparisonCode%, ByVal AndOrLogic%, ByVal SecondField$) As Integer
'Declare Function ddExtAddLogicConst Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$, ByVal ComparisonCode%, ByVal AndOrLogic%, ConstField As Any, ByVal ConstSize%) As Integer
'Declare Function ddExtAddLogicString Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$, ByVal ComparisonCode%, ByVal AndOrLogic%, ByVal ConstField$) As Integer
'Declare Function ddSetField Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$, FieldValue As Any) As Integer
'Declare Function ddSetFieldString Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$, ByVal InputMask$, ByVal FieldValue$) As Integer
'Declare Function ddGetField Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$, FieldValue As Any) As Integer
'Declare Function ddGetFieldString Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$, ByVal OutputMask$, FieldValue As String) As Integer
' The following functions position to the first, next, last, or previous
' record in the virtual list provided by the extended operations.
' Use the ddExtGetField functions for retrieving data. You can also use
' the "btr" versions for getting the entire record.
'Declare Function ddExtGetCurrent Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function ddExtGetFirst Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function ddExtGetLast Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function ddExtGetNext Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function ddExtGetPrevious Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function ddExtInsert Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function ddExtUpdate Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function ddExtDelete Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function ddExtGetField Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$, FieldValue As Any) As Integer
'Declare Function ddExtSetField Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$, FieldValue As Any) As Integer
'Declare Function ddExtGetFieldString Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$, ByVal OutputMask$, FieldValue As String) As Integer
'Declare Function ddExtSetFieldString Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$, ByVal InputMask$, ByVal FieldValue$) As Integer
' Have extended operation track logical record number.
'Declare Function ddSetTrackExtRecord Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal TrackFlag%) As Integer
'Declare Function ddGetTrackExtRecord Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
' Create File from DDF Definition. Use FilePathName to overide file location in DDF.
' Set to "" if you want to use File Loacation from DDF.
'Declare Function ddCreateDDFTable Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal DDFPath$, ByVal TableName$, ByVal FilePathName$, ByVal fOverWrite%) As Integer
'Declare Function ddCreateTable Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal TableName$, ByVal fOverWrite%) As Integer
' The following to statements are used to provide query capability.
' The ExtAddFields allows you to specify fields by Name and comma delimited.
' For Example "DeptNum, DeptName". Unlike the ExtAddField function this will
' add more than one field with a single call. Use "*" for all fields.
' The ExtAddLogicQuery allows you to use a query statement to extract your
' Data. For example "DeptNum = 200 and LastName = Jones".
' Both the fields and logic statement are converted into a Btrieve Extended
' operation definition. To get your Data use the ExtGetFirst, ExtGetNext,
' ExtGetPrevious and ExtGetLast family of functions. These functions will automatically
' Call the Get Next Extended and Get Previous Extended functions where appropriate.
' The result of the querey are a virtual list. You can get a count on the number of
' elements in the list with the ddExtRecords and ddExtRecordCount
' functions. The ExtRecordCount function forces a recount whereas the ddExtRecords
' will return the count on the first call and the value of the last call on
' subsequent calls.
'Declare Function ddExtAddFields Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Fields$) As Integer
'Declare Function ddExtAddLogicQuery Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Query$) As Integer
' The following two functions allow you to specify a SQL query:
'
'   SELECT [ Field Names ]
'   FROM [ Table Name ]
'   WHERE [ Logical Statement ]
'
'   The WHERE clause is optional. The Field Names can be "*"
'   in which case all fields are selected.
'
'   For Example
'
'   "Select * From Department"
'
'    Or
'
'    "Select DeptName, DeptNum From Department where DeptNum = 200"
'
'   logical statements using field names are seperated by AND and OR.
'   For example:
'
'   "A < 10 and B > 5"
'
'    Where a and b are field names in the file. Only simple logical
'    statements are supported. You are restricted to only the same
'    logical statements supported by Btrieves Extended operations.
'
' The FilePath is retrieved from the DDF and the file is opened
' and the Extended operation built. Use the ExtGetFirst, ExtGetNext,
' etc for iterating through your data. The result of the operation
' is a virtual list.
'
' The ddExtQuery function uses the file name in the DDF. The
' ddExtQueryFile function uses specified file path. It's behavior will be the
' same as ddExtQuery if the file name is "".
'Declare Function ddExtQuery Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Query$) As Integer
'Declare Function ddExtQueryFile Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Query$, ByVal FileName$) As Integer
' Allows overiding record count.
'Declare Sub ddSetExtRecordCount Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal RecCount&)
' Position to a logical record number in a set from extended operation.
' For Example if the ExtRecordCount returned 500 then you can position to the
' 250th record with the following.
' Dim ret As integer
' ret = ddExtSetToRecord(Ohnd, 250)
' You can then use ExtGetField to get any field or btrExtGetCurrent to get the entire
' record extracted.
'Declare Function ddExtSetToRecord Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal RecordNumber&) As Integer
' Get refreshed count of records in extended operation set.
' If a logic filter has not been specified then this functionwill
' return the same value as btrRecords
Declare Function ddExtRecordCount Lib "cbtrv432.dll" (ByVal Ohnd%) As Long
' Returns result of last count of records.
'Declare Function ddExtRecords Lib "cbtrv432.dll" (ByVal Ohnd%) As Long
' CTableDef add in DDF definition information. Read Only.
' A CTableDef is built into the CBtrvTable class. So you can call
' any of these functions with a handle constructed from the CBtrvTable
' constructor function.
'Declare Function CTableDef Lib "cbtrv432.dll" () As Integer
'Declare Function ddfOpenDDF Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal DDFPath$) As Integer
'Declare Function ddfFirstTable Lib "cbtrv432.dll" (ByVal Ohnd%, FileRec As rFILEREC) As Integer
'Declare Function ddfNextTable Lib "cbtrv432.dll" (ByVal Ohnd%, FileRec As rFILEREC) As Integer
'Declare Function ddfPreviousTable Lib "cbtrv432.dll" (ByVal Ohnd%, FileRec As rFILEREC) As Integer
'Declare Function ddfLastTable Lib "cbtrv432.dll" (ByVal Ohnd%, FileRec As rFILEREC) As Integer
'Declare Function ddfFirstTableName Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal TableName$) As Integer
'Declare Function ddfNextTableName Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal TableName$) As Integer
'Declare Function ddfPreviousTableName Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal TableName$) As Integer
'Declare Function ddfLastTableName Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal TableName$) As Integer
' Get definition for a particular table in DDF.
'Declare Function ddfTable Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal TableName$, FileRec As rFILEREC) As Integer
'Declare Function ddfOpenDDFTableDefs Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal DDFPath$, ByVal TableName$) As Integer
'Declare Function ddfOpenTableDefs Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal TableName$) As Integer
' Controls loading and unloading of DDF definitions from memory.
' With the definitions in memory the speed is increased but at the price of
' increased memory use.
'Declare Function ddfUnloadFieldsFromMem Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function ddfUnloadIndexesFromMem Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function ddfLoadFieldsFromDisk Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function ddfLoadIndexesFromDisk Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function ddfFieldCount Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function ddfSegmentCount Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function ddfIndexCount Lib "cbtrv432.dll" (ByVal Ohnd%) As Integer
'Declare Function ddfSegmentCountOnIndex Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Index%) As Integer
' DDF Definition functions.
' Find a specified field definition.
'Declare Function ddfField Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$, FieldRec As rFIELDREC) As Integer
' Determines whether a particular field exists on an index. Returns
' Segment definition for field.
'Declare Function ddfIndexOnFieldSegment Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$, ByVal Index%, IndexRec As rINDEXREC) As Integer
' Iterate through all Field definitions. These functions affect
' the current field definition in memory or on disk.
'Declare Function ddfFirstField Lib "cbtrv432.dll" (ByVal Ohnd%, FieldRec As rFIELDREC) As Integer
'Declare Function ddfNextField Lib "cbtrv432.dll" (ByVal Ohnd%, FieldRec As rFIELDREC) As Integer
'Declare Function ddfPreviousField Lib "cbtrv432.dll" (ByVal Ohnd%, FieldRec As rFIELDREC) As Integer
'Declare Function ddfLastField Lib "cbtrv432.dll" (ByVal Ohnd%, FieldRec As rFIELDREC) As Integer
' Iterate all FieldNames.
'Declare Function ddfFirstFieldName Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$) As Integer
'Declare Function ddfNextFieldName Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$) As Integer
'Declare Function ddfPreviousFieldName Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$) As Integer
'Declare Function ddfLastFieldName Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$) As Integer
' Iterate through all of the segments that a field exists on.
' This will affect both the current Field and current segment position
' in either memory or disk.
'Declare Function ddfFirstSegmentOnField Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal FieldName$, IndexRec As rINDEXREC) As Integer
'Declare Function ddfNextSegmentOnField Lib "cbtrv432.dll" (ByVal Ohnd%, IndexRec As rINDEXREC) As Integer
' Iterate though all the segments. Thes functions affect the
' position of the current segment definition either on memory or
' disk.
'Declare Function ddfFirstSegmentOnTable Lib "cbtrv432.dll" (ByVal Ohnd%, IndexRec As rINDEXREC) As Integer
'Declare Function ddfNextSegmentOnTable Lib "cbtrv432.dll" (ByVal Ohnd%, IndexRec As rINDEXREC) As Integer
' Functions for accessing index and segments within an index.
' An index contains one or more segments ordered 0 to N where
' N is the last segment.
' Get a particular indexes segment definition.
'Declare Function ddfIndexSegOnTable Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Index%, ByVal Seg%, IndexRec As rINDEXREC) As Integer
' Get the next segment in the current index.
'Declare Function ddfNextSegOnIndex Lib "cbtrv432.dll" (ByVal Ohnd%, IndexRec As rINDEXREC) As Integer
' Gets the associated field definition for the specified Index Segment.
'Declare Function ddfFieldOnIndexSegment Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Index%, ByVal Seg%, FieldRec As rFIELDREC) As Integer
' Iterate through all of the fields defined on an index. This will
' affect both the current Field and segment in memory.
'Declare Function ddfFirstFieldOnIndex Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal Index%, FieldRec As rFIELDREC, IndexRec As rINDEXREC) As Integer
'Declare Function ddfNextFieldOnIndex Lib "cbtrv432.dll" (ByVal Ohnd%, FieldRec As rFIELDREC, IndexRec As rINDEXREC) As Integer
