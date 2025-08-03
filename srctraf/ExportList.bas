Attribute VB_Name = "ReportListSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Budget.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExportList.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Date/Time subs and functions
Option Explicit
Option Compare Text

'Used to get current setting of color, vertical and horizontal resoluation
Public Declare Function GetDeviceCaps Lib "gdi32" _
   (ByVal hdc As Long, _
    ByVal nIndex As Long) As Long


'*******************************************************
'*                                                     *
'*      Procedure Name:gCenterModalForm                *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Center modal form within        *
'*                     Traffic Form                    *
'*                                                     *
'*******************************************************
Sub gCenterModalForm(FrmName As Form)
'
'   gCenterModalForm FrmName
'   Where:
'       FrmName (I)- Name of modal form to be centered within Traffic form
'
        gCenterStdAlone FrmName
    'Dim flLeft As Single
    'Dim flTop As Single
    'flLeft = Traffic.Left + (Traffic.Width - Traffic.ScaleWidth) / 2 + (Traffic.ScaleWidth - FrmName.Width) / 2
    'flTop = Traffic.Top + (Traffic.Height - FrmName.Height + 2 * Traffic.cmcTask(0).Height - 60) / 2 + Traffic.cmcTask(0).Height
    'FrmName.Move flLeft, flTop
End Sub
