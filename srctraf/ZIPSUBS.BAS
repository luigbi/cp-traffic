Attribute VB_Name = "ZIPSUBS"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ZipSubs.BAS
'
' Release: 1.0
'
' Description:
'  This file contains the Zip subs and functions
Option Explicit
Option Compare Text

Public Const NO_ACTION = 0
Public Const UE_OK = 0
' the "function"s that are supported for unzipping
Public Const UNZIP_COUNTALLZIPMEMBERS = 1
Public Const UNZIP_GETNEXTZIPINFO = 2
Public Const UNZIP_COUNTNAMEDZIPMEMBERS = 3
Public Const UNZIP_GETNEXTNAMEDZIPINFO = 4
Public Const UNZIP_GETCOMMENTSIZE = 5
Public Const UNZIP_GETCOMMENT = 6
Public Const UNZIP_GETINDEXEDZIPINFO = 7
Public Const UNZIP_EXTRACT = 8
Public Const UNZIP_FILETOMEM = 9
' the "function"s that are supported for zipping
Public Const ZIP_FRESHEN = 1
Public Const ZIP_DELETE = 2
Public Const ZIP_UPDATE = 3
Public Const ZIP_ADD = 4
Public Const ZIP_MEMTOFILE = 5
' **************************************************************************************
'
'  Procedure:  initUNZIPCmdStruct()
'
'  Purpose:  make the inputted UNZIP control properties "plain vanilla"
'
'  NOTE: These settings are default settings only.  The differ somewhat from the
'        example code.
' **************************************************************************************
Sub initUNZIPCmdStruct(zpcName As Control) 'IMI_DUNZIPVBXCONTROL)
'  zpcName.Action = 0
'  zpcName.BackgroundProcessFlag = False
'  zpcName.ConvertLFtoCRLFFlag = False
'  zpcName.DecryptCode = ""
'  zpcName.DecryptFlag = False
'  zpcName.Destination = ""
'  zpcName.DiagnosticFlag = False
'  zpcName.ErrorCode = 0
'  zpcName.Filespec = ""
'  zpcName.FreshenFlag = False
'  zpcName.MajorStatusFlag = False
'  zpcName.MessageCallbackFlag = False
'  zpcName.MinorStatusFlag = False
'  zpcName.NoDirectoryItemsFlag = True
'  zpcName.NoDirectoryNamesFlag = True
'  zpcName.OverwriteFlag = False
'  zpcName.QuietFlag = False
'  zpcName.RecurseFlag = True
'  zpcName.ReturnCount = 0
'  zpcName.ReturnString = ""
'  zpcName.TestFlag = False
'  zpcName.UnZIPIndex = -1
'  zpcName.UnZipSubOptions = 0
'  zpcName.UpdateFlag = False
'  zpcName.zi_attr = 0
'  zpcName.zi_cMethod = 0
'  zpcName.zi_cPathType = 0
'  zpcName.zi_crc_32 = 0
'  zpcName.zi_cSize = 0
'  zpcName.zi_DateTime = ""
'  zpcName.zi_FileName = ""
'  zpcName.zi_index = 0
'  zpcName.zi_oSize = 0
'  zpcName.ZIPFile = ""
'
'  ' added for rev 3.00
  zpcName.RenameCallbackFlag = True
'  zpcName.ExtProgTitle = ""
'  zpcName.UnZIPStringSize = 0
'  zpcName.UnZIPStringOffset = 0
  zpcName.ActionDZ = 0
  zpcName.BackgroundProcessFlag = False
  zpcName.ConvertLFtoCRLFFlag = False
  zpcName.DecryptCode = ""
  zpcName.DecryptFlag = False
  zpcName.Destination = ""
  zpcName.DiagnosticFlag = False
  zpcName.ErrorCode = 0
  zpcName.Filespec = ""
  zpcName.FreshenFlag = False
  zpcName.MajorStatusFlag = False
  zpcName.MessageCallbackFlag = False
  zpcName.MinorStatusFlag = False
  zpcName.NoDirectoryItemsFlag = True
  zpcName.NoDirectoryNamesFlag = True
  zpcName.OverwriteFlag = False
  zpcName.QuietFlag = False
  zpcName.RecurseFlag = True
  zpcName.ReturnCount = 0
  zpcName.ReturnString = ""
  zpcName.TestFlag = False
  zpcName.UnZIPIndex = -1
  zpcName.UnZipSubOptions = 0
  zpcName.UpdateFlag = False

  zpcName.zi_attr = 0
  zpcName.zi_cMethod = 0
  zpcName.zi_cPathType = 0
  zpcName.zi_crc_32 = 0
  zpcName.zi_cSize = 0
  zpcName.zi_DateTime = ""
  zpcName.zi_FileName = ""
  zpcName.zi_index = 0
  zpcName.zi_oSize = 0
  zpcName.ZIPFile = ""
  
  ' added for rev 3.00
  zpcName.RenameCallbackFlag = True
  zpcName.ExtProgTitle = ""
  zpcName.UnZIPStringSize = 0
  zpcName.UnZIPStringOffset = 0

End Sub
' **************************************************************************************
'
'  Procedure:  initZIPCmdStruct()
'
'  Purpose:  make the ZIP control's properties "plain vanilla"
'
'  NOTE: These settings are default settings only.  The differ somewhat from the
'        example code.
' **************************************************************************************
Sub initZIPCmdStruct(zpcName As Control) 'IMI_DZIPVBXCONTROL)
'  zpcName.Action = NO_ACTION
'  zpcName.AddCommentFlag = False
'  zpcName.AfterDateFlag = False
'  zpcName.BackgroundProcessFlag = False
'  zpcName.Comment = ""
'  zpcName.CompressionFactor = 5
'  zpcName.ConvertLFtoCRLFFlag = False
'  zpcName.Date = ""
'  zpcName.DeleteOriginalFlag = False
'  zpcName.DiagnosticFlag = False
'  zpcName.DontCompressTheseSuffixesFlag = False
'  zpcName.DosifyFlag = False
'  zpcName.EncryptCode = ""
'  zpcName.EncryptFlag = False
'  zpcName.ExcludeFollowing = ""
'  zpcName.ExcludeFollowingFlag = False
'  zpcName.FixFlag = False
'  zpcName.FixHarderFlag = False
'  zpcName.GrowExistingFlag = False
'  zpcName.IncludeFollowing = ""
'  zpcName.IncludeOnlyFollowingFlag = False
'  zpcName.IncludeSysandHiddenFlag = False
'  zpcName.IncludeVolumeFlag = False
'  zpcName.ItemList = ""
'  zpcName.MajorStatusFlag = False
'  zpcName.MessageCallbackFlag = False
'  zpcName.MinorStatusFlag = False
'  zpcName.MultiVolumeControl = 0
'  zpcName.NoDirectoryEntriesFlag = False
'  zpcName.NoDirectoryNamesFlag = False
'  zpcName.OldAsLatestFlag = False
'  zpcName.PathForTempFlag = False
'  zpcName.QuietFlag = False
'  zpcName.RecurseFlag = False
'  zpcName.StoreSuffixes = ""
'  zpcName.TempPath = ""
'  zpcName.ZIPFile = ""
'  zpcName.ZipSubOptions = 0
'  ' added for rev 3.00
  zpcName.RenameCallbackFlag = True
'  zpcName.ExtProgTitle = ""
'  zpcName.ZIPString = ""
  zpcName.ActionDZ = NO_ACTION
  zpcName.AddCommentFlag = False
  zpcName.AfterDateFlag = False
  zpcName.BackgroundProcessFlag = False
  zpcName.Comment = ""
  zpcName.CompressionFactor = 5
  zpcName.ConvertLFtoCRLFFlag = False
  zpcName.Date = ""
  zpcName.DeleteOriginalFlag = False
  zpcName.DiagnosticFlag = False
  zpcName.DontCompressTheseSuffixesFlag = False
  zpcName.DosifyFlag = False
  zpcName.EncryptCode = ""
  zpcName.EncryptFlag = False
  zpcName.ExcludeFollowing = ""
  zpcName.ExcludeFollowingFlag = False
  zpcName.FixFlag = False
  zpcName.FixHarderFlag = False
  zpcName.GrowExistingFlag = False
  zpcName.IncludeFollowing = ""
  zpcName.IncludeOnlyFollowingFlag = False
  zpcName.IncludeSysandHiddenFlag = False
  zpcName.IncludeVolumeFlag = False
  zpcName.ItemList = ""
  zpcName.MajorStatusFlag = False
  zpcName.MessageCallbackFlag = False
  zpcName.MinorStatusFlag = False
  zpcName.MultiVolumeControl = 0
  zpcName.NoDirectoryEntriesFlag = False
  zpcName.NoDirectoryNamesFlag = False
  zpcName.OldAsLatestFlag = False
  zpcName.PathForTempFlag = False
  zpcName.QuietFlag = False
  zpcName.RecurseFlag = False
  zpcName.StoreSuffixes = ""
  zpcName.TempPath = ""
  zpcName.ZIPFile = ""
  zpcName.ZipSubOptions = 0

  ' added for rev 3.00
  zpcName.RenameCallbackFlag = True
  zpcName.ExtProgTitle = ""
  zpcName.ZIPString = ""

End Sub

Public Function gUnzipError(ilError As Integer) As String
    Select Case ilError
        Case 2
            gUnzipError = "Unexpected end of Zip File"
        Case 3
            gUnzipError = "Zip file structure invalid"
        Case 4
            gUnzipError = "Out of Memory"
        Case 5
            gUnzipError = "Out of Memory"
        Case 9
            gUnzipError = "File not found"""
        Case 11
            gUnzipError = "No Files to test extract"
        Case 12
            gUnzipError = "Destination same drive as Zip"
        Case 25
            gUnzipError = "Index out of bounds"
        Case 28
            gUnzipError = "Error creating output file"
        Case 29
            gUnzipError = "Error opening archive"
        Case 39
            gUnzipError = "Bad CRC"
        Case 40
            gUnzipError = "Application canceled operation"
        Case 41
            gUnzipError = "File skipped, encrypted"
        Case 42
            gUnzipError = "File skipped, unknown compression method"
        Case 44
            gUnzipError = "Bad or missing decrypt code"
        Case 45
            gUnzipError = "Busy, can't enter now"
        Case 46
            gUnzipError = "Can't extract a Volumn ID item"
        Case 47
            gUnzipError = "Bad command structure"
        Case 48
            gUnzipError = "Operation canceled"
        Case 49
            gUnzipError = "Operation skipped- password does not match"
        Case 50
            gUnzipError = "Disk Full"
        Case Else
            gUnzipError = "Undocument error"
    End Select
End Function

Public Function gZipError(ilError As Integer) As String
    Select Case ilError
        Case 1
            gZipError = "Busy, can't enter now"
        Case 2
            gZipError = "Unexpected end of Zip file"
        Case 3
            gZipError = "Zip file structure invalid"
        Case 4
            gZipError = "Out of Memory"
        Case 5
            gZipError = "Internal logic error"
        Case 6
            gZipError = "Entry too big to split"
        Case 7
            gZipError = "Invalid comment format"
        Case 8
            gZipError = "Zip file invalid or insufficient memory"
        Case 9
            gZipError = "Application canceled operation"
        Case 10
            gZipError = "Temporary file failure"
        Case 11
            gZipError = "Input file read failure"
        Case 12
            gZipError = "Nothing to do- no matching file to zip"
        Case 13
            gZipError = "Missing or empty ZIP file"
        Case 14
            gZipError = "Output file write failure, possible disk full"
        Case 15
            gZipError = "Could not create output file"
        Case 16
            gZipError = "Invalid combination of control parameters"
        Case 17
            gZipError = "Could not complete operation"
        Case 18
            gZipError = "File not found or no read permission"
        Case 19
            gZipError = "Media error, disk not ready, or hardware error"
        Case 20
            gZipError = "Invalid combination of multi-volumn control parameters"
        Case 21
            gZipError = "Improper use of multi-volumn ZIP file"
        Case Else
            gZipError = "Undocument error"
    End Select
            
End Function
