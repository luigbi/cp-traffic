Attribute VB_Name = "INVSTDALONE"


' Proprietary Software, Do not copy
'
' File Name: InvStdAlone.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice support functions
Option Explicit
Option Compare Text

Public tmRcf As RCF


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
    Dim flLeft As Single
    Dim flTop As Single
    flLeft = Invoice.Left + (Invoice.Width - Invoice.ScaleWidth) / 2 + (Invoice.ScaleWidth - FrmName.Width) / 2
    flTop = Invoice.Top + (Invoice.Height - FrmName.Height + 2 * Invoice.cmcCancel.Height - 60) / 2 + Invoice.cmcCancel.Height
    FrmName.Move flLeft, flTop
End Sub

Sub gShowBranner(ilUpdateAllowed As Integer)
'
'   gShowBranner
'   Where:
'

End Sub
