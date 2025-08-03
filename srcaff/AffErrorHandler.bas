Attribute VB_Name = "modErrorHandler"

Option Explicit

Public sgCallStack(0 To 9) As String
Public lgUlfCode As Long
Public hgUaf As Integer
Public tgUaf As UAF
Private tmUafSrchKey3 As UAFKEY3
Public igUAFRecLen As Integer
Public sgReportListName As String
Public igLogActivityStatus As Integer
Public igNoDaysRetainUAF As Integer
Type UAFSTACK
    sName As String * 50
    lUafCode As Long
End Type
Public tgUafStack() As UAFSTACK
Type TASKNAMEMAP
    sFormName As String * 20
    sUafName As String * 50
End Type
Public tgTaskNameMap() As TASKNAMEMAP

Public sgTmfStatus As String    'S=Start; C=Completed; E=Error; A=Aborted

Public Sub gDbg_HandleError(ModuleAndFunction As String)
    Dim slMsg As String
    Dim slToFile As String
    Dim slDateTime As String
    Dim ilTo As Integer
    Dim ilLine As Integer
    Dim slDesc As String
    Dim ilErrNo As Integer
    Dim slAppName As String
    Dim ilPos As Integer
    Dim slPCName As String
    Dim slMACAddr As String
    Dim ilRet As Integer
    
   
    ' Get the error information now to preserve it.
    'It must be prior to Resume Next as that clears the values
    ilLine = Erl
    ilErrNo = Err.Number
    slDesc = Err.Description
    
    On Error Resume Next
    
    slAppName = App.EXEName
    ilPos = InStr(1, slAppName, ".", 1)
    If ilPos > 0 Then
        slAppName = Left$(slAppName, ilPos - 1)
    End If
    slAppName = slAppName & ".exe"
    
    slMsg = Format$(Now, "mm-dd-yyyy") & " " & Format$(Now, "hh:mm:ssam/pm") & " User: " & sgUserName & " "
    slMsg = slMsg & vbCrLf & vbCrLf & "Module: " & ModuleAndFunction & vbCrLf
    slMsg = slMsg & "Line No: " & ilLine & vbCrLf
    slMsg = slMsg & "Error: " & Str(ilErrNo) & vbCrLf
    slMsg = slMsg & "Desc: " & slDesc & vbCrLf & vbCrLf
    'slMsg = slMsg & slAppName & ": " & Format$(FileDateTime(sgExeDirectory & slAppName), "m/d/yy") & " at " & Format$(FileDateTime(sgExeDirectory & slAppName), "h:mm:ssAM/PM") & vbCrLf & vbCrLf
    slMsg = slMsg & slAppName & ": " & Format$(gFileDateTime(sgExeDirectory & slAppName), "m/d/yy") & " at " & Format$(gFileDateTime(sgExeDirectory & slAppName), "h:mm:ssAM/PM") & vbCrLf & vbCrLf
    slMsg = slMsg & "The System will now shut down."

    gMsgBox slMsg, vbCritical, "Application Error"

    ' Reformat and Log the error message as well
    'slMsg = slMsg = Format$(Now, "mm-dd-yyyy") & " " & Format$(Now, "hh:mm:ssam/pm") & " User: " & sgUserName & " "
    slMsg = Format$(Now, "mm-dd-yyyy") & " " & Format$(Now, "hh:mm:ssam/pm") & " User: " & sgUserName & " " & _
          ", Module: " & ModuleAndFunction & _
          ", Line No: " & ilLine & _
          ", Error: " & Str(ilErrNo) & _
          ", Desc: " & slDesc & _
          ", " & slAppName & ": " & Format$(gFileDateTime(sgExeDirectory & slAppName), "m/d/yy") & " at " & Format$(gFileDateTime(sgExeDirectory & slAppName), "h:mm:ssAM/PM")

    slPCName = Trim$(gGetComputerName())
    slMACAddr = Trim$(gGetMACs_AdaptInfo())
    slMsg = slMsg & " PC Name: " & slPCName
    'D.S. 4/1/09
    'If Trim$(sgReportName) <> "" Then
    '    slMsg = slMsg & "   User: " & Trim$(sgReportName)
    'Else
    '    slMsg = slMsg & " User: " & Trim$(sgUserName)
    'End If
    
    If right$(sgMsgDirectory, 1) <> "\" Then
        slToFile = sgMsgDirectory & "\AffErrorLog.Txt"
    Else
        slToFile = sgMsgDirectory & "AffErrorLog.Txt"
    End If
    'ilTo = FreeFile
    'Open slToFile For Append As ilTo
    ilRet = gFileOpen(slToFile, "Append", ilTo)
    If ilRet = 0 Then
        Print #ilTo, slMsg
    End If
    Close #ilTo
    
    gSaveStackTrace slToFile
   
'    slAppName = App.EXEName
'    If InStr(1, slAppName, ".", 1) > 0 Then
'        slAppName = Left$(slAppName, ilPos - 1)
'    End If

    'Unload Traffic
        
    sgTmfStatus = "A"
    
    btrStopAppl
    End
End Sub



Public Sub gErrorApplStop(ModuleAndFunction As String)
    Dim slMsg As String
    Dim slToFile As String
    Dim slDateTime As String
    Dim ilTo As Integer
    Dim ilErrNo As Integer
    Dim slAppName As String
    Dim ilPos As Integer
    Dim ilRet As Integer
    On Error Resume Next
    
    slAppName = App.EXEName
    ilPos = InStr(1, slAppName, ".", 1)
    If ilPos > 0 Then
        slAppName = Left$(slAppName, ilPos - 1)
    End If
    slAppName = slAppName & ".exe"

    slDateTime = Format$(Now(), "ddddd h:mm:ssAM/PM")
    slMsg = slDateTime & vbCrLf & vbCrLf & "Module: " & ModuleAndFunction & vbCrLf & _
          slAppName & ": " & Format$(gFileDateTime(sgExeDirectory & slAppName), "m/d/yy") & " at " & Format$(gFileDateTime(sgExeDirectory & slAppName), "h:mm:ssAM/PM") & vbCrLf & vbCrLf & _
          "The System will now shut down."

    gMsgBox slMsg, vbCritical, "Application Error"

    ' Reformat and Log the error message as well
    slMsg = slDateTime & _
          ", Module: " & ModuleAndFunction & _
          ", " & slAppName & ": " & Format$(gFileDateTime(sgExeDirectory & slAppName), "m/d/yy") & " at " & Format$(gFileDateTime(sgExeDirectory & slAppName), "h:mm:ssAM/PM")

    slToFile = sgDBPath & "Messages\TrafficErrors.Txt"
    'ilTo = FreeFile
    'Open slToFile For Append As ilTo
    ilRet = gFileOpen(slToFile, "Append", ilTo)
    If ilRet = 0 Then
        Print #ilTo, slMsg
    End If
    Close #ilTo
    
    sgTmfStatus = "E"
    
    gSaveStackTrace slToFile
    
    
    btrStopAppl
    End

End Sub

Public Sub gAddCallToStack(slCallToAdd As String)
    Dim ilLoop As Integer
    
    For ilLoop = UBound(sgCallStack) To LBound(sgCallStack) + 1 Step -1
        sgCallStack(ilLoop) = sgCallStack(ilLoop - 1)
    Next ilLoop
    sgCallStack(LBound(sgCallStack)) = slCallToAdd
End Sub

Public Sub gRemoveCallFromStack()
    Dim ilLoop As Integer
    
    For ilLoop = LBound(sgCallStack) To UBound(sgCallStack) - 1 Step 1
        sgCallStack(ilLoop) = sgCallStack(ilLoop + 1)
    Next ilLoop
    sgCallStack(UBound(sgCallStack)) = ""
End Sub

Public Sub gSaveStackTrace(slLogFileName As String)
    Dim ilTo As Integer
    Dim ilLoop As Integer
    Dim ilTotalLen As Integer
    Dim slMethodName As String
    Dim slCallStack As String
    Dim ilRet As Integer
    
    ' Verify there is at least one item on the call stack. Otherwise don't print anything.
    ilTotalLen = 0
    For ilLoop = LBound(sgCallStack) To UBound(sgCallStack)
        slMethodName = sgCallStack(ilLoop)
        ilTotalLen = ilTotalLen + Len(slMethodName)
    Next ilLoop
    If ilTotalLen < 1 Then
        Exit Sub
    End If

    'ilTo = FreeFile
    'Open slLogFileName For Append As ilTo
    ilRet = gFileOpen(slLogFileName, "Append", ilTo)
    If ilRet = 0 Then
        slCallStack = "Call Stack Trace : "
        For ilLoop = LBound(sgCallStack) To UBound(sgCallStack)
            slMethodName = sgCallStack(ilLoop)
            If Len(slMethodName) > 0 Then
                slCallStack = slCallStack + slMethodName + ", "
            End If
        Next ilLoop
        Print #ilTo, slCallStack
    End If
    Close #ilTo
End Sub

Public Sub gUserActivityLog(slFunction As String, slInName As String)
    'slFunction (I): L = Load Traffic or Affiliate form;
    '                S = Start report
    '                U = Unload Traffic or Affiliate form
    '                E = End report
    Dim ilRet As Integer
    Dim slAppName As String
    Dim ilLowLimit As Integer
    Dim ilStack As Integer
    Dim slName As String
    Dim ilLoop As Integer
    Dim tlUafSrchKey As LONGKEY0
    Dim slNowDate As String
    Dim llDeleteDate As Long
    Dim slDate As String
    
    If (igLogActivityStatus <> 32123) And (igLogActivityStatus <> -32123) Then
        Exit Sub
    End If
    If igNoDaysRetainUAF <= -1 Then
        Exit Sub
    End If
    If igNoDaysRetainUAF = 0 Then
        igNoDaysRetainUAF = 5
    End If
    ilRet = 0
    'On Error GoTo ErrHandle
    'ilLowLimit = LBound(tgUafStack)
    If PeekArray(tgUafStack).Ptr <> 0 Then
        ilLowLimit = LBound(tgUafStack)
    Else
        ilRet = 1
        ilLowLimit = 0
    End If
    If ilRet = 1 Then
        If (igLogActivityStatus = -32123) Then
            igLogActivityStatus = 0
            Exit Sub
        End If
        'Initialize arrays
        slAppName = UCase(App.EXEName)
        If InStr(1, slAppName, UCase("Traffic"), vbBinaryCompare) >= 1 Then
            mTrafficFormNames
        ElseIf InStr(1, slAppName, UCase("Reports"), vbBinaryCompare) >= 1 Then
            mReportsFormNames
        ElseIf InStr(1, slAppName, UCase("Affiliat"), vbBinaryCompare) >= 1 Then
            mAffiliateFormNames
        Else
            'igLogActivityStatus = 0
            'Exit Sub
            mAffiliateFormNames
        End If
        'Open file
        hgUaf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hgUaf, "", sgDBPath & "UAF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            MsgBox "Unable to Open User Activity Log File, Error = " & Str$(ilRet), vbOKOnly + vbInformation, "Warning"
            igLogActivityStatus = 0
            Exit Sub
        End If
        igUAFRecLen = Len(tgUaf)
        ReDim tgUafStack(0 To 0) As UAFSTACK
        slNowDate = Format(Now, "m/d/yy")
        llDeleteDate = gDateValue(DateAdd("d", -(igNoDaysRetainUAF + 1), slNowDate))
        gPackDate Format$(llDeleteDate, "m/d/yy"), tmUafSrchKey3.iStartDate(0), tmUafSrchKey3.iStartDate(1)
        ilRet = btrGetLessOrEqual(hgUaf, tgUaf, igUAFRecLen, tmUafSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
        Do While ilRet = BTRV_ERR_NONE
            gUnpackDate tgUaf.iStartDate(0), tgUaf.iStartDate(1), slDate
            If gDateValue(slDate) > llDeleteDate Then
                Exit Do
            End If
            ilRet = btrDelete(hgUaf)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            ilRet = btrGetLessOrEqual(hgUaf, tgUaf, igUAFRecLen, tmUafSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
        Loop
    End If
    If (InStr(1, UCase(slInName), ".FRM", vbBinaryCompare) > 0) Then
        slName = mBinarySearchTaskName(Trim$(UCase(slInName)))
    Else
        slName = slInName
    End If
    slName = Trim$(slName)
    If (igLogActivityStatus = 32123) Then
        If (slFunction = "L") Or (slFunction = "S") Then     'Form Load or Start Report
            'Create UAF and place in Stack
            tgUaf.lCode = 0
            tgUaf.sSystemType = UCase(Left$(App.EXEName, 1))
            If tgUaf.sSystemType = "R" Then         '5-23-11 if reports mode, its still a traffic system type
                tgUaf.sSystemType = "T"
            End If
            If (InStr(1, UCase(slInName), ".FRM", vbBinaryCompare) > 0) Then
                tgUaf.sSubType = "T"
            Else
                tgUaf.sSubType = "R"
            End If
            tgUaf.lUlfCode = lgUlfCode
            tgUaf.iUserCode = igUstCode
            tgUaf.sName = slName
            tgUaf.sStatus = "I"
            gPackDate Format$(Now, "m/d/yy"), tgUaf.iStartDate(0), tgUaf.iStartDate(1)
            gPackTime Format$(Now, "h:mm:ssAM/PM"), tgUaf.iStartTime(0), tgUaf.iStartTime(1)
            gPackDate "12/31/2069", tgUaf.iEndDate(0), tgUaf.iEndDate(1)
            gPackTime "12:00:00 AM", tgUaf.iEndTime(0), tgUaf.iEndTime(1)
            gPackDate Format$(gNow(), "m/d/yy"), tgUaf.iCSIDate(0), tgUaf.iCSIDate(1)
            ilRet = btrInsert(hgUaf, tgUaf, igUAFRecLen, INDEXKEY0)
            If ilRet = BTRV_ERR_NONE Then
                'Add to Stack
                tgUafStack(UBound(tgUafStack)).lUafCode = tgUaf.lCode
                tgUafStack(UBound(tgUafStack)).sName = tgUaf.sName
                ReDim Preserve tgUafStack(0 To UBound(tgUafStack) + 1) As UAFSTACK
            End If
        ElseIf (slFunction = "U") Or (slFunction = "E") Then    'Form Unload or Report End
            'Search for Match in Stack and Update
            For ilStack = UBound(tgUafStack) - 1 To 0 Step -1
                If StrComp(slName, Trim$(tgUafStack(ilStack).sName), vbBinaryCompare) = 0 Then
                    tlUafSrchKey.lCode = tgUafStack(ilStack).lUafCode
                    ilRet = btrGetEqual(hgUaf, tgUaf, igUAFRecLen, tlUafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        gPackDate Format$(Now, "m/d/yy"), tgUaf.iEndDate(0), tgUaf.iEndDate(1)
                        gPackTime Format$(Now, "h:mm:ssAM/PM"), tgUaf.iEndTime(0), tgUaf.iEndTime(1)
                        tgUaf.sStatus = "C"
                        ilRet = btrUpdate(hgUaf, tgUaf, igUAFRecLen)
                        If ilRet = BTRV_ERR_NONE Then
                            For ilLoop = ilStack To UBound(tgUafStack) - 1 Step 1
                                tgUafStack(ilLoop) = tgUafStack(ilLoop + 1)
                            Next ilLoop
                            ReDim Preserve tgUafStack(0 To UBound(tgUafStack) - 1) As UAFSTACK
                        End If
                    End If
                End If
            Next ilStack
        End If
    ElseIf (igLogActivityStatus = -32123) Then  'Unload all
        For ilStack = 0 To UBound(tgUafStack) - 1 Step 1
            tlUafSrchKey.lCode = tgUafStack(ilStack).lUafCode
            ilRet = btrGetEqual(hgUaf, tgUaf, igUAFRecLen, tlUafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                gPackDate Format$(Now, "m/d/yy"), tgUaf.iEndDate(0), tgUaf.iEndDate(1)
                gPackTime Format$(Now, "h:mm:ssAM/PM"), tgUaf.iEndTime(0), tgUaf.iEndTime(1)
                tgUaf.sStatus = "C"
                ilRet = btrUpdate(hgUaf, tgUaf, igUAFRecLen)
            End If
        Next ilStack
        ilRet = btrClose(hgUaf)
        btrDestroy hgUaf
        igLogActivityStatus = 0
        ReDim tgUafStack(0 To 0) As UAFSTACK
    End If
    Exit Sub
ErrHandle:
    ilRet = 1
    Resume Next
End Sub


Private Sub mAddTaskName(ilIndex As Integer, slFormName As String, slUafName As String)
    If ilIndex >= UBound(tgTaskNameMap) Then
        ReDim Preserve tgTaskNameMap(0 To UBound(tgTaskNameMap) + 20) As TASKNAMEMAP
    End If
    tgTaskNameMap(ilIndex).sFormName = slFormName
    tgTaskNameMap(ilIndex).sUafName = slUafName
    ilIndex = ilIndex + 1
End Sub

Private Sub mTrafficFormNames()
    Dim ilNumber As Integer
    Dim ilLoop As Integer
    
    ReDim tgTaskNameMap(0 To 212) As TASKNAMEMAP
    ilNumber = 0
    mAddTaskName ilNumber, "Advt.Frm", "Advertiser"
    mAddTaskName ilNumber, "AdvtProd.Frm", "Advertiser Product"
    mAddTaskName ilNumber, "AffCopy.Frm", "Affiliate Copy"
    mAddTaskName ilNumber, "Agency.Frm", "Agency"
    mAddTaskName ilNumber, "AlertVw.Frm", "Alert Viewer"
    mAddTaskName ilNumber, "AName.Frm", "Avail Names"
    mAddTaskName ilNumber, "ARInvNo.Frm", "A/R Invoice Number"
    mAddTaskName ilNumber, "ARReconc.Frm", "A/R Reconcile"
    mAddTaskName ilNumber, "AuxWnd.Frm", "List"
    mAddTaskName ilNumber, "Basic10.Frm", "Jobs"
    mAddTaskName ilNumber, "Blackout.Frm", "Blackout"
    mAddTaskName ilNumber, "BlkDate.Frm", "Rollover Block Date"
    mAddTaskName ilNumber, "BlockVw.Frm", "Lock View"
    mAddTaskName ilNumber, "BPlate.Frm", "Boiler Plate"
    mAddTaskName ilNumber, "Browser.Frm", "Browser"
    mAddTaskName ilNumber, "BrSnap.Frm", "Contract Snap Shot"
    mAddTaskName ilNumber, "BtrCheck.Frm", "File Check"
    mAddTaskName ilNumber, "Buf12Mo.Frm", "Budget by 12 Months"
    mAddTaskName ilNumber, "BudActA.Frm", "Budgets by Actuals"
    mAddTaskName ilNumber, "BudActB.Frm", "Budget by Budget"
    mAddTaskName ilNumber, "BudAdd.Frm", "Add Vehicle to Budget"
    mAddTaskName ilNumber, "BudAdvt.Frm", "Budget by Advertiser"
    mAddTaskName ilNumber, "BUDataToTest.Frm", "Copy Prod Data to Test"
    mAddTaskName ilNumber, "Budget.Frm", "Budget"
    mAddTaskName ilNumber, "BudModel.Frm", "Budget Model"
    mAddTaskName ilNumber, "BudResch.Frm", "Budget by Research"
    mAddTaskName ilNumber, "BudScale.Frm", "Scale Budget"
    mAddTaskName ilNumber, "BudTrend.Frm", "Budget Trend"
    mAddTaskName ilNumber, "BulkFeed.Frm", "Bulk Feed Resend"
    mAddTaskName ilNumber, "BUZip.Frm", "Backup Database"
    mAddTaskName ilNumber, "Calendar.Frm", "Calendar"
    mAddTaskName ilNumber, "CAvail.Frm", "Contract Avails View"
    mAddTaskName ilNumber, "CCancel.Frm", "Contract Line Cancel"
    mAddTaskName ilNumber, "CClone.Frm", "Contract Line Clone"
    mAddTaskName ilNumber, "CCntrNo.Frm", "Contract Number"
    mAddTaskName ilNumber, "CffCheck.Frm", "Contract Flight Check"
    mAddTaskName ilNumber, "CGameSch.Frm", "Contract Event Schedule"
    mAddTaskName ilNumber, "CModel.Frm", "Contract Model"
    mAddTaskName ilNumber, "CntrFind.Frm", "Contract Find"
    mAddTaskName ilNumber, "CntrProj.Frm", "Contract Projection"
    mAddTaskName ilNumber, "CntrSch.Frm", "Contract Schedule"
    mAddTaskName ilNumber, "Collect.Frm", "Receivable"
    mAddTaskName ilNumber, "Contract.Frm", "Contract"
    mAddTaskName ilNumber, "Copy.Frm", "Copy Rotation"
    mAddTaskName ilNumber, "CopyAsgn.Frm", "Copy Assigning"
    mAddTaskName ilNumber, "CopyDupl.Frm", "Copy Inventory Duplication"
    mAddTaskName ilNumber, "CopyInv.Frm", "Copy Inventory"
    mAddTaskName ilNumber, "CopyRato.Frm", "Copy Ratio"
    mAddTaskName ilNumber, "CopyRegn.Frm", "Copy Region"
    mAddTaskName ilNumber, "CopySplit.Frm", "Copy Split"
    mAddTaskName ilNumber, "CorpCal.Frm", "Corporate Calendar"
    mAddTaskName ilNumber, "CPackage.Frm", "Dynamic Package"
    mAddTaskName ilNumber, "CRBPkg.Frm", "Dynamic Rate Card Package"
    mAddTaskName ilNumber, "CRevNo.Frm", "Contract Revision Number"
    mAddTaskName ilNumber, "CScale.Frm", "Contract Dollar Scale"
    mAddTaskName ilNumber, "CShift.Frm", "Contract Line Shift"
    mAddTaskName ilNumber, "CShTitle.Frm", "Contract Short Title"
    mAddTaskName ilNumber, "CSIAbout.Frm", "CSI About"
    mAddTaskName ilNumber, "CSINewPW.Frm", "New Password"
    mAddTaskName ilNumber, "CSnapPrt.Frm", "Contract Snap Shot Selection"
    mAddTaskName ilNumber, "CSPWord.Frm", "Password"
    mAddTaskName ilNumber, "DateRange.Frm", "Receivable Date Range"
    mAddTaskName ilNumber, "Daypart.Frm", "Daypart"
    mAddTaskName ilNumber, "EName.Frm", "Event Name"
    mAddTaskName ilNumber, "EType.Frm", "Event Type"
    mAddTaskName ilNumber, "ExpACC.Frm", "Export Accounting"
    mAddTaskName ilNumber, "ExpBkCpy.Frm", "Export Bulk Copy"
    mAddTaskName ilNumber, "ExpCmChg.Frm", "Export Commercial Changes"
    mAddTaskName ilNumber, "ExpCnCAP.Frm", "Export CnC Advertiser/Product"
    mAddTaskName ilNumber, "ExpCnCNI.Frm", "Export CnC Network Inventory"
    mAddTaskName ilNumber, "ExpCnCSA.Frm", "Export CnC Selling To Airing Links"
    mAddTaskName ilNumber, "ExpCnCSS.Frm", "Export CnC Schedule Spots"
    mAddTaskName ilNumber, "ExpDall.Frm", "Export Dallas"
    mAddTaskName ilNumber, "ExpGP.Frm", "Export Great Plains G/L"
    mAddTaskName ilNumber, "ExpGPBarter.Frm", "Export Great Plains Barter"
    mAddTaskName ilNumber, "ExpInv.Frm", "Export Invoices"
    mAddTaskName ilNumber, "ExpISCIXRef.Frm", "Export ISCI Cross Reference"
    mAddTaskName ilNumber, "ExpMatrix.Frm", "Export Matrix"
    mAddTaskName ilNumber, "ExpNY.Frm", "Export New York"
    mAddTaskName ilNumber, "ExpPhnx.Frm", "Export Phoenix"
    mAddTaskName ilNumber, "ExpProj.Frm", "Export Projections"
    mAddTaskName ilNumber, "ExpRevenue.Frm", "Export Revenue"
    mAddTaskName ilNumber, "ExpStnFd.Frm", "Export Station Feed"
    mAddTaskName ilNumber, "ExpEnco.Frm", "Export Enco"
    mAddTaskName ilNumber, "ExptGen.Frm", "Export To Automation Equipment"
    mAddTaskName ilNumber, "ExptMP2.Frm", "Export MP2"
    mAddTaskName ilNumber, "ExpEfficioRev.frm", "Export Efficio Revenue"
    mAddTaskName ilNumber, "ExpMatrix.Frm", "Export Tableau"                '7-8-15   added tableau export, same as matrix format
    mAddTaskName ilNumber, "ExpMatrix.frm", "Export RAB"                    '2-3-20 added RAB CRM export, similar to matrix/tableau
    mAddTaskName ilNumber, "FdCartNo.Frm", "Find Cart Number"
    mAddTaskName ilNumber, "FeedName.Frm", "Feed Name"
    mAddTaskName ilNumber, "FeedPlge.Frm", "Feed Pledge"
    mAddTaskName ilNumber, "FeedSpot.Frm", "Feed Spot"
    mAddTaskName ilNumber, "GameInv.Frm", "Event Inventory"
    mAddTaskName ilNumber, "GameLib.Frm", "Event Library"
    mAddTaskName ilNumber, "GameSchd.Frm", "Event Schedule"
    mAddTaskName ilNumber, "GenMsg.Frm", "General Message Box"
    mAddTaskName ilNumber, "GenSch.Frm", "General Progress Message"
    mAddTaskName ilNumber, "GetGames.Frm", "Event Selection"
    mAddTaskName ilNumber, "IconTraf.Frm", "General Icons"
    mAddTaskName ilNumber, "ImptCntr.Frm", "Import Contracts"
    'mAddTaskName ilNumber, "ImptCopy.Frm", "Import Copy inventory"
    mAddTaskName ilNumber, "ImptGen.Frm", "Import Automation Times"
    mAddTaskName ilNumber, "ImptIMS.Frm", "Import IMS Research"
    mAddTaskName ilNumber, "ImptMark.Frm", "Import Act1 Research"
    mAddTaskName ilNumber, "ImptRad.Frm", "Import RADAR Research"
    mAddTaskName ilNumber, "ImptSat.Frm", "Import Satellite Research"
    mAddTaskName ilNumber, "InvCheck.Frm", "Invoice Check"
    mAddTaskName ilNumber, "InvItem.Frm", "Event Inventory Item Name"
    mAddTaskName ilNumber, "Invoice.Frm", "Invoice"
    mAddTaskName ilNumber, "InvType.Frm", "Event Inventory Type"
    mAddTaskName ilNumber, "LinkDlvy.Frm", "Delivery Links"
    mAddTaskName ilNumber, "Links.Frm", "Link Selection"
    mAddTaskName ilNumber, "LinksDef.Frm", "Selling/Airing Links"
    mAddTaskName ilNumber, "LiveLog.Frm", "Live Log"
    mAddTaskName ilNumber, "LLAddSpt.Frm", "Live Log Add Spot"
    mAddTaskName ilNumber, "LLFeed.Frm", "Live Log Feed"
    mAddTaskName ilNumber, "LLHelp.Frm", "Live Log Help"
    mAddTaskName ilNumber, "LLSignOn.Frm", "Live Log Sign On"
    mAddTaskName ilNumber, "Locks.Frm", "Lock Avails/Spots"
    mAddTaskName ilNumber, "LogChk.Frm", "Lock Check"
    mAddTaskName ilNumber, "LogMkt.Frm", "Market Selection"
    mAddTaskName ilNumber, "Logs.Frm", "Logs"
    mAddTaskName ilNumber, "MathCalc.Frm", "Calculator"
    mAddTaskName ilNumber, "Media.Frm", "Media"
    mAddTaskName ilNumber, "Merge.Frm", "Merge"
    mAddTaskName ilNumber, "Messages.Frm", "Message Folder Info"
    mAddTaskName ilNumber, "MstPict.Frm", "Background Picture"
    mAddTaskName ilNumber, "MultiNm.Frm", "List Functions"
    mAddTaskName ilNumber, "NetworkSplit.Frm", "Network Split Region"
    mAddTaskName ilNumber, "NmAddr.Frm", "Lock Box Address"
    mAddTaskName ilNumber, "Password.Frm", "User Password"
    mAddTaskName ilNumber, "PBEMail.Frm", "E-Mail"
    mAddTaskName ilNumber, "PdModel.Frm", "Feed Pledge Model"
    mAddTaskName ilNumber, "Persnnel.Frm", "Personnel"
    mAddTaskName ilNumber, "PEvent.Frm", "Program Library Events"
    mAddTaskName ilNumber, "PLogMkt.Frm", "Post Log Selection"
    mAddTaskName ilNumber, "PLogTime.Frm", "Post Log Time"
    mAddTaskName ilNumber, "POModel.Frm", "Transfer PO Transaction"
    mAddTaskName ilNumber, "PostAdjt.Frm", "Invoice Post Adjustment"
    mAddTaskName ilNumber, "PostItem.Frm", "Invoice Post Item"
    mAddTaskName ilNumber, "PostLog.Frm", "Post Log"
    mAddTaskName ilNumber, "PostRep.Frm", "Post Rep"
    mAddTaskName ilNumber, "PostRepTimes.Frm", "Post Rep Spot Times"
    mAddTaskName ilNumber, "PreFeed.Frm", "Spot Pre-Feed Specification"
    mAddTaskName ilNumber, "PrgDates.Frm", "Program Library Air Dates/Times"
    mAddTaskName ilNumber, "PrgDel.Frm", "Program Library Delete Specifications"
    mAddTaskName ilNumber, "PrgDupl.Frm", "Program Library Duplication"
    mAddTaskName ilNumber, "PrgSch.Frm", "Program Library Scheduling"
    mAddTaskName ilNumber, "Program.Frm", "Program Layout"
    mAddTaskName ilNumber, "Purge.Frm", "Copy Inventory Purge"
    mAddTaskName ilNumber, "RateCard.Frm", "Rate Card"
    mAddTaskName ilNumber, "RCImpact.Frm", "Rate Card Proposal Impact"
    mAddTaskName ilNumber, "RCModel.Frm", "Rate Card Model"
    mAddTaskName ilNumber, "RCReallo.Frm", "Rate Card Dollar Reallocation"
    mAddTaskName ilNumber, "RCSplit.Frm", "Rate Card Split"
    mAddTaskName ilNumber, "RCTerms.Frm", "Rate Card Terms"
    mAddTaskName ilNumber, "RepNet.Frm", "Rep-Network Link"
    mAddTaskName ilNumber, "Report.Frm", "Crystal Report Interface"
    mAddTaskName ilNumber, "Research.Frm", "Research"
    mAddTaskName ilNumber, "Rollover.Frm", "Rollover Dates"
    mAddTaskName ilNumber, "RptList.Frm", "Report List"
    mAddTaskName ilNumber, "RptNoSel.Frm", "Report With No Selections"
    mAddTaskName ilNumber, "RptSel.Frm", "General Accounting, Copy, Lists Selection"
    mAddTaskName ilNumber, "RptSelBR.Frm", "Contract: Report Selection"
    mAddTaskName ilNumber, "RptSelCt.Frm", "General Proposal, Order, Sales: Report Selection"
    mAddTaskName ilNumber, "RptSelDB.Frm", "Demo Bar Summary: Report Selection"
    mAddTaskName ilNumber, "RptSelEx.Frm", "Enco Export: Report Selection"
    mAddTaskName ilNumber, "RptSelIA.Frm", "Selection: IA"
    mAddTaskName ilNumber, "RptSelIn.Frm", "Invoice: Report Selection"
    mAddTaskName ilNumber, "RptSelLg.Frm", "Log: Report Selection"
    mAddTaskName ilNumber, "RptSelOA.Frm", "Order Audit: Report Selection"
    mAddTaskName ilNumber, "RptSelPr.Frm", "Proposal Research Recap: Report Selection"
    mAddTaskName ilNumber, "RptSelRI.Frm", "Remote Invoice Worksheet: Report Selection"
    mAddTaskName ilNumber, "RptSelSS.Frm", "Spot Screen Snap Shot: Report Selection"
    mAddTaskName ilNumber, "RptSelTx.Frm", "Text Dump: Report Selection"
    mAddTaskName ilNumber, "RptSets.Frm", "Report Sets"
    mAddTaskName ilNumber, "RschByCust.Frm", "Custom Research"
    mAddTaskName ilNumber, "RSModel.Frm", "Research Book Model"
    mAddTaskName ilNumber, "SaleHist.Frm", "Build BackLog"
    mAddTaskName ilNumber, "SetModel.Frm", "Report Set Model"
    mAddTaskName ilNumber, "ShoCrdit.Frm", "Show Credit Info"
    mAddTaskName ilNumber, "ShtTitle.Frm", "Short Title"
    mAddTaskName ilNumber, "Signon.Frm", "Sign On"
    mAddTaskName ilNumber, "SiteOpt.Frm", "Site Option"
    mAddTaskName ilNumber, "SlspComm.Frm", "Salesperson Commission"
    mAddTaskName ilNumber, "SlspCrte.Frm", "Create Salesperson Commission"
    mAddTaskName ilNumber, "SlspMod.Frm", "Salesperson Model"
    mAddTaskName ilNumber, "SMFCheck.Frm", "SMF Check"
    mAddTaskName ilNumber, "SOffice.Frm", "Sales Office"
    mAddTaskName ilNumber, "SPerson.Frm", "Salesperson"
    mAddTaskName ilNumber, "SplitModel.Frm", "Split Model"
    mAddTaskName ilNumber, "SPModel.Frm", "Standard Package Model"
    mAddTaskName ilNumber, "SportChk.Frm", "Sport Check"
    mAddTaskName ilNumber, "Sports.Frm", "Sports Task"
    mAddTaskName ilNumber, "SpotAction.Frm", "Spot Action"
    mAddTaskName ilNumber, "SpotFill.Frm", "Spot Fill"
    mAddTaskName ilNumber, "SpotLine.Frm", "Extra Bonus Line Selection"
    mAddTaskName ilNumber, "SpotMG.Frm", "1 For 1 MG"
    mAddTaskName ilNumber, "Spots.Frm", "Spots"
    mAddTaskName ilNumber, "SpotWks.Frm", "Extend Spot Weeks"
    mAddTaskName ilNumber, "SsfCheck.Frm", "SSF Check"
    mAddTaskName ilNumber, "StdPkg.Frm", "Standard Package"
    mAddTaskName ilNumber, "StnFdCpy.Frm", "Station Feed Copy Export"
    mAddTaskName ilNumber, "StnFdUnd.Frm", "Station Feed Undo"
    mAddTaskName ilNumber, "TaxTable.Frm", "Tax Table"
    mAddTaskName ilNumber, "Traffic.Frm", "Traffic"
    mAddTaskName ilNumber, "UndoBkFd.Frm", "Bulk Feed Undo"
    mAddTaskName ilNumber, "UnSchd.Frm", "Rectify"
    mAddTaskName ilNumber, "UserLogEMail.Frm", "User Log E-Mail"
    mAddTaskName ilNumber, "UserOpt.Frm", "User Option"
    mAddTaskName ilNumber, "UsersLog.Frm", "Users Log"
    mAddTaskName ilNumber, "Vehicle.Frm", "Vehicle"
    mAddTaskName ilNumber, "VehModel.Frm", "Vehicle Model"
    mAddTaskName ilNumber, "VehOpt.Frm", "Vehicle Options"
    mAddTaskName ilNumber, "ViewList.Frm", "Days Not Market Complete"
    ReDim Preserve tgTaskNameMap(0 To ilNumber) As TASKNAMEMAP
    For ilLoop = 0 To UBound(tgTaskNameMap) - 1 Step 1
        tgTaskNameMap(ilLoop).sFormName = UCase(tgTaskNameMap(ilLoop).sFormName)
    Next ilLoop
    ArraySortTyp fnAV(tgTaskNameMap(), 0), UBound(tgTaskNameMap), 0, LenB(tgTaskNameMap(0)), 0, LenB(tgTaskNameMap(0).sFormName), 0
End Sub

Private Sub mAffiliateFormNames()
    Dim ilNumber As Integer
    Dim ilLoop As Integer
    
    ReDim tgTaskNameMap(0 To 113) As TASKNAMEMAP
    ilNumber = 0
    mAddTaskName ilNumber, "AffNewPW.Frm", "User New Password"
    mAddTaskName ilNumber, "BUZip.Frm", "Backup Database"
    mAddTaskName ilNumber, "CSPWord.Frm", "Password"
    mAddTaskName ilNumber, "AffAbout.Frm", "About"
    mAddTaskName ilNumber, "AffAddBonus.Frm", "Add Bonus"
    mAddTaskName ilNumber, "AffAddMG.Frm", "Add MG"
    mAddTaskName ilNumber, "AffAdvFulfillRpt.Frm", "Advertiser Fulfillmen: Report Selection"
    mAddTaskName ilNumber, "AffAdvPlaceRpt.Frm", "Advertiser Placement: Report Selection"
    mAddTaskName ilNumber, "AffDP.Frm", "Agreement Daypart"
    mAddTaskName ilNumber, "AffiliateRpt.Frm", "Spot Aired: Report Selection"
    mAddTaskName ilNumber, "AffRep.Frm", "Affiliate A/E"
    mAddTaskName ilNumber, "AffAgmnt.Frm", "Agreement"
    mAddTaskName ilNumber, "AffAiredRpt.Frm", "Spot Aired: Report Selection"
    mAddTaskName ilNumber, "AffAlertRpt.Frm", "Alert Status: Report Selection"
    mAddTaskName ilNumber, "AffAlertVw.Frm", "Alert Viewer"
    mAddTaskName ilNumber, "AffAstCheckUtil.Frm", "AST Check Utility"
    mAddTaskName ilNumber, "AffAvRemap.Frm", "Avail Remap for Agreements"
    mAddTaskName ilNumber, "AffBrowse.Frm", "Browse"
    mAddTaskName ilNumber, "AffCategoryMatching.Frm", "Station Import Category Matching"
    mAddTaskName ilNumber, "AffCDStartTime.Frm", "CD Start Time"
    mAddTaskName ilNumber, "AffClrRpt.Frm", "Clearance: Report Selection"
    mAddTaskName ilNumber, "AffContact.Frm", "Contact"
    mAddTaskName ilNumber, "AffContactCopy.Frm", "Contact"
    mAddTaskName ilNumber, "AffContactEMail.Frm", "Contact E-Mail"
    mAddTaskName ilNumber, "AffContactGrid.Frm", "Contact Grid"
    mAddTaskName ilNumber, "AffCPLog.Frm", "C.P./Log"
    mAddTaskName ilNumber, "AffCPCount.Frm", "C.P. Count"
    mAddTaskName ilNumber, "AffCPRetStatus.Frm", "C.P. Return Status"
    mAddTaskName ilNumber, "AffCPReturns.Frm", "C.P. Return Posting Selection"
    mAddTaskName ilNumber, "AffCPTTCheck.Frm", "C.P. Check"
    mAddTaskName ilNumber, "AffCrystal.Frm", "Crystal Interface"
    mAddTaskName ilNumber, "AffCPDateTimes.Frm", "C.P. Posting Date/Time"
    mAddTaskName ilNumber, "AffDelqRpt.Frm", "Delinquent: Report Selection"
    mAddTaskName ilNumber, "AffDirectory.Frm", "Directory"
    mAddTaskName ilNumber, "AffDuplCPTTFix.Frm", "Fix Duplicate C.P."
    mAddTaskName ilNumber, "AffDuplSHTTFix.Frm", "Fix Duplicate Stations"
    mAddTaskName ilNumber, "AffEMail.Frm", "E-Mail"
    mAddTaskName ilNumber, "AffEMailConv.Frm", "E-Mail Conversion"
    mAddTaskName ilNumber, "AffExpMonRpt.Frm", "Export Monitoring: Report Selection"
    mAddTaskName ilNumber, "AffExportCncSpots.Frm", "Export CnC Spots"
    mAddTaskName ilNumber, "AffExportISCI.Frm", "Export ISCI"
    mAddTaskName ilNumber, "AffExportISCIRef.Frm", "Export ISCI Cross Reference"
    mAddTaskName ilNumber, "AffExportLabelInfo.Frm", "Export Label Info"
    mAddTaskName ilNumber, "AffExportMarketron.Frm", "Export To Marketron"
    mAddTaskName ilNumber, "AffExportOLA.Frm", "Export To OLA"
    mAddTaskName ilNumber, "AffExportRCS.Frm", "Export To RCS"
    If Not gUsingUnivision Then
        mAddTaskName ilNumber, "AffExportSchdSpot.Frm", "Export To Univision"
    End If
    mAddTaskName ilNumber, "AffExportStarGuide.Frm", "Export To Star Guide"
    mAddTaskName ilNumber, "AffExportStationInformation.Frm", "Export Station Information"
    mAddTaskName ilNumber, "AffExportWegener.Frm", "Export To Wegener"
    mAddTaskName ilNumber, "AffExportXDigital.Frm", "Export To X-Digital"
    mAddTaskName ilNumber, "AffFastAdd.Frm", "Fast Add"
    mAddTaskName ilNumber, "AffFastAddWarning.Frm", "Fast Add Warning"
    mAddTaskName ilNumber, "AffFastEnd.Frm", "Fast End"
    mAddTaskName ilNumber, "AffGenMsg.Frm", "General Message Box"
    mAddTaskName ilNumber, "AffGetGame.Frm", "Event Selection"
    mAddTaskName ilNumber, "AffGetPath.Frm", "Get Path/Folder"
    mAddTaskName ilNumber, "AffGroupNameFormat.Frm", "Group Name Format"
    mAddTaskName ilNumber, "AffGroupNameMarket.Frm", "Group Name DMA Market"
    mAddTaskName ilNumber, "AffGroupNameMSAMarket.Frm", "Group Name MSA Market"
    mAddTaskName ilNumber, "AffGroupNameState.Frm", "Group Name State"
    mAddTaskName ilNumber, "AffGroupNameTimeZone.Frm", "Group Name Time Zone"
    mAddTaskName ilNumber, "AffGroupNameVehicle.Frm", "Group Name Vehicle"
    mAddTaskName ilNumber, "AffGroupRpt.Frm", "Group: Report Selection"
    mAddTaskName ilNumber, "AffHistory.Frm", "Call Letter History Question"
    mAddTaskName ilNumber, "AffImportAE.Frm", "Import A/E"
    If Not gUsingUnivision Then
        mAddTaskName ilNumber, "AffImportAiredSpot.Frm", "Import Univision Spots"
    End If
    mAddTaskName ilNumber, "AffImportCSISpot.Frm", "Import CSI Spots"
    mAddTaskName ilNumber, "AffImptCSV.Frm", "Import CSV Files"
    mAddTaskName ilNumber, "AffImportMarketron.Frm", "Import Marketron"
    mAddTaskName ilNumber, "AffImportUpdateStations.Frm", "Import Station Information"
    mAddTaskName ilNumber, "AffJournalRpt.Frm", "Journal: Report Selection"
    mAddTaskName ilNumber, "AffLabelRpt.Frm", "Label: Report Selection"
    mAddTaskName ilNumber, "AffLogActivityRpt.Frm", "Log Activity: Report Selection"
    mAddTaskName ilNumber, "AffLogin.Frm", "Log In"
    mAddTaskName ilNumber, "AffLogInactivityRpt.Frm", "Log Inactivity: Report Selection"
    mAddTaskName ilNumber, "AffMain.Frm", "Main"
    mAddTaskName ilNumber, "AffMarkAssignRpt.Frm", "Market Assignment: Report Selection"
    mAddTaskName ilNumber, "AffMerge.Frm", "Merge"
    mAddTaskName ilNumber, "AffMessages.Frm", "Message Folder Info"
    mAddTaskName ilNumber, "AffModel.Frm", "Model"
    mAddTaskName ilNumber, "AffMstPict.Frm", "Background Picture"
    mAddTaskName ilNumber, "AffOptions.Frm", "User Options"
    mAddTaskName ilNumber, "AffPgmClrRpt.Frm", "Program Clearance: Report Selection"
    mAddTaskName ilNumber, "AffPldgAirRpt.Frm", "Pledge vs Aired: Report Selection"
    mAddTaskName ilNumber, "AffPledgeRpt.Frm", "Pledge: Report Selection"
    mAddTaskName ilNumber, "AffPostActivityRpt.Frm", "Web Log Activity: Report Selection"
    mAddTaskName ilNumber, "AffPostLog.Frm", "Post Log"
    mAddTaskName ilNumber, "frmProgressMsg.Frm", "Progress Message"
    mAddTaskName ilNumber, "AffRadarExport.Frm", "RADAR Export"
    mAddTaskName ilNumber, "AffRadarProgSchd.Frm", "RADAR Program Schedule"
    mAddTaskName ilNumber, "AffReports.Frm", "Report List"
    mAddTaskName ilNumber, "AffRgAssgnRpt.Frm", "Regional Copy Assignment: Report Selection"
    mAddTaskName ilNumber, "frmSelRemap.Frm", "Re-mapped Affiliate"
    mAddTaskName ilNumber, "AffSendReport.Frm", "Crystal Interface"
    mAddTaskName ilNumber, "frmSiteOptions.Frm", "Site Options"
    mAddTaskName ilNumber, "AffSpotUtil.Frm", "Spot Utility"
    mAddTaskName ilNumber, "AffStation.Frm", "Station"
    mAddTaskName ilNumber, "AffStationMktInfo.Frm", "DMA Market Information"
    mAddTaskName ilNumber, "AffStationMSAMktInfo.Frm", "MSA Market Information"
    mAddTaskName ilNumber, "AffStationOwnerInfo.Frm", "Owner Information"
    mAddTaskName ilNumber, "AffStationRpt.Frm", "Station: Report Selection"
    mAddTaskName ilNumber, "AffStationSearch.Frm", "Affiliate Management"
    mAddTaskName ilNumber, "AffStationSearchFilter.Frm", "Affiliate Management Filter"
    mAddTaskName ilNumber, "AffStatioZone.Frm", "Station Agreement Update"
    mAddTaskName ilNumber, "AffTitle.Frm", "Personnel Title"
    mAddTaskName ilNumber, "AffUserLogEMail.Frm", "User Log E-Mail"
    mAddTaskName ilNumber, "AffUsersLog.Frm", "Users Log"
    mAddTaskName ilNumber, "AffVehAffRpt.Frm", "Agreement: Report Selection"
    mAddTaskName ilNumber, "AffVerifyRpt.Frm", "Feed Verification: Report Selection"
    mAddTaskName ilNumber, "AffViewReport.Frm", "Preview Crystal Report Interface"
    mAddTaskName ilNumber, "AffWebEMail.Frm", "Web E-Mail"
    mAddTaskName ilNumber, "AffWebExportSchdSpot.Frm", "Web Export Schedule Spots"
    mAddTaskName ilNumber, "AffWebImportAiredSpot.Frm", "Web Import Aired Spots"
    mAddTaskName ilNumber, "AffWebIniOptions.Frm", "Web INI Options"
    mAddTaskName ilNumber, "AffXMLTestMode.Frm", "XML Test Mode"
    ReDim Preserve tgTaskNameMap(0 To ilNumber) As TASKNAMEMAP
    For ilLoop = 0 To UBound(tgTaskNameMap) - 1 Step 1
        tgTaskNameMap(ilLoop).sFormName = UCase(tgTaskNameMap(ilLoop).sFormName)
    Next ilLoop
    ArraySortTyp fnAV(tgTaskNameMap(), 0), UBound(tgTaskNameMap), 0, LenB(tgTaskNameMap(0)), 0, LenB(tgTaskNameMap(0).sFormName), 0
End Sub

Private Sub mReportsFormNames()
    Dim ilNumber As Integer
    Dim ilLoop As Integer
    
    ReDim tgTaskNameMap(0 To 62) As TASKNAMEMAP
    ilNumber = 0
    mAddTaskName ilNumber, "AffNewPW.Frm", "Password"
    mAddTaskName ilNumber, "BUZip.Frm", "Backup Database"
    mAddTaskName ilNumber, "CSPWord.Frm", "Password"
    mAddTaskName ilNumber, "AffAbout.Frm", "About"
    mAddTaskName ilNumber, "AffAddBonus.Frm", "Add Bonus"
    mAddTaskName ilNumber, "AffAddMG.Frm", "Add MG"
    mAddTaskName ilNumber, "AffAdvFulfillRpt.Frm", "Advertiser Fulfillmen: Report Selection"
    mAddTaskName ilNumber, "AffAdvPlaceRpt.Frm", "Advertiser Placement: Report Selection"
    mAddTaskName ilNumber, "AffDP.Frm", "Agreement Daypart"
    mAddTaskName ilNumber, "AffiliateRpt.Frm", "Spot Aired: Report Selection"
    mAddTaskName ilNumber, "AffRep.Frm", "Affiliate A/E"
    mAddTaskName ilNumber, "AffAgmnt.Frm", "Agreement"
    mAddTaskName ilNumber, "AffAgmntPledgeSpec.Frm", "Auto-Fill"
    mAddTaskName ilNumber, "AffAiredRpt.Frm", "Spot Aired: Report Selection"
    mAddTaskName ilNumber, "AffAlertRpt.Frm", "Alert Status: Report Selection"
    mAddTaskName ilNumber, "AffAlertVw.Frm", "Alert Viewer"
    mAddTaskName ilNumber, "AffAstCheckUtil.Frm", "AST Check Utility"
    mAddTaskName ilNumber, "AffAvRemap.Frm", "Avail Remap for Agreements"
    mAddTaskName ilNumber, "AffBrowse.Frm", "Browse"
    mAddTaskName ilNumber, "AffCategoryMatching.Frm", "Station Import Category Matching"
    mAddTaskName ilNumber, "AffCDStartTime.Frm", "CD Start Time"
    mAddTaskName ilNumber, "AffClrRpt.Frm", "Clearance: Report Selection"
    mAddTaskName ilNumber, "AffCmmtSrc.Frm", "Comment Source"
    mAddTaskName ilNumber, "AffCommentRpt.Frm", "Contact Comment: Report Selection"
    mAddTaskName ilNumber, "AffCompliantTracer.Frm", "Compliance Tracer"
    mAddTaskName ilNumber, "AffContact.Frm", "Contact"
    'mAddTaskName ilNumber, "AffContactCopy.Frm", "Contact"
    mAddTaskName ilNumber, "AffContactEMail.Frm", "Contact E-Mail"
    mAddTaskName ilNumber, "AffContactGrid.Frm", "Contact Grid"
    mAddTaskName ilNumber, "AffCPLog.Frm", "C.P./Log"
    mAddTaskName ilNumber, "AffCPCount.Frm", "C.P. Count"
    mAddTaskName ilNumber, "AffCPRetStatus.Frm", "C.P. Return Status"
    mAddTaskName ilNumber, "AffCPReturns.Frm", "C.P. Return Posting Selection"
    mAddTaskName ilNumber, "AffCPTTCheck.Frm", "C.P. Check"
    mAddTaskName ilNumber, "AffCrystal.Frm", "Crystal Interface"
    mAddTaskName ilNumber, "AffCPDateTimes.Frm", "C.P. Posting Date/Time"
    mAddTaskName ilNumber, "AffDelqRpt.Frm", "Delinquent: Report Selection"
    mAddTaskName ilNumber, "AffDepartment.Frm", "Department"
    mAddTaskName ilNumber, "AffDirectory.Frm", "Directory"
    mAddTaskName ilNumber, "AffDuplCPTTFix.Frm", "Fix Duplicate C.P."
    mAddTaskName ilNumber, "AffDuplSHTTFix.Frm", "Fix Duplicate Stations"
    mAddTaskName ilNumber, "AffEMail.Frm", "E-Mail"
    mAddTaskName ilNumber, "AffEMailConv.Frm", "E-Mail Conversion"
    mAddTaskName ilNumber, "AffExpMonRpt.Frm", "Export Monitoring: Report Selection"
    mAddTaskName ilNumber, "AffExportCncSpots.Frm", "Export CnC Spots"
    mAddTaskName ilNumber, "AffExportISCI.Frm", "Export ISCI"
    mAddTaskName ilNumber, "AffExportISCIRef.Frm", "Export ISCI Cross Reference"
    mAddTaskName ilNumber, "AffExportLabelInfo.Frm", "Export Label Info"
    mAddTaskName ilNumber, "AffExportMarketron.Frm", "Export To Marketron"
    mAddTaskName ilNumber, "AffExportOLA.Frm", "Export To OLA"
    mAddTaskName ilNumber, "AffExportRCS.Frm", "Export To RCS"
    If Not gUsingUnivision Then
        mAddTaskName ilNumber, "AffExportSchdSpot.Frm", "Export To Univision"
    End If
    mAddTaskName ilNumber, "AffExportStarGuide.Frm", "Export To Star Guide"
    mAddTaskName ilNumber, "AffExportStationInformation.Frm", "Export Station Information"
    mAddTaskName ilNumber, "AffExportWegener.Frm", "Export To Wegener"
    mAddTaskName ilNumber, "AffExportXDigital.Frm", "Export To X-Digital"
    mAddTaskName ilNumber, "AffFastAdd.Frm", "Fast Add"
    mAddTaskName ilNumber, "AffFastAddWarning.Frm", "Fast Add Warning"
    mAddTaskName ilNumber, "AffFastEnd.Frm", "Fast End"
    mAddTaskName ilNumber, "AffGenMsg.Frm", "General Message Box"
    mAddTaskName ilNumber, "AffGetGame.Frm", "Event Selection"
    mAddTaskName ilNumber, "AffGetPath.Frm", "Get Path/Folder"
    mAddTaskName ilNumber, "AffGroupNameFormat.Frm", "Group Name Format"
    mAddTaskName ilNumber, "AffGroupNameMarket.Frm", "Group Name DMA Market"
    mAddTaskName ilNumber, "AffGroupNameMSAMarket.Frm", "Group Name MSA Market"
    mAddTaskName ilNumber, "AffGroupNameState.Frm", "Group Name State"
    mAddTaskName ilNumber, "AffGroupNameTimeZone.Frm", "Group Name Time Zone"
    mAddTaskName ilNumber, "AffGroupNameVehicle.Frm", "Group Name Vehicle"
    mAddTaskName ilNumber, "AffGroupRpt.Frm", "Group: Report Selection"
    mAddTaskName ilNumber, "AffHistory.Frm", "Call Letter History Question"
    mAddTaskName ilNumber, "AffImportAE.Frm", "Import A/E"
    If Not gUsingUnivision Then
        mAddTaskName ilNumber, "AffImportAiredSpot.Frm", "Import Univision Spots"
    End If
    mAddTaskName ilNumber, "AffImportCSISpot.Frm", "Import CSI Spots"
    mAddTaskName ilNumber, "AffImptCSV.Frm", "Import CSV Files"
    mAddTaskName ilNumber, "AffImportMarketron.Frm", "Import Marketron"
    mAddTaskName ilNumber, "AffImportUpdateStations.Frm", "Import Station Information"
    mAddTaskName ilNumber, "AffJournalRpt.Frm", "Journal: Report Selection"
    mAddTaskName ilNumber, "AffLabelRpt.Frm", "Label: Report Selection"
    mAddTaskName ilNumber, "AffLogActivityRpt.Frm", "Log Activity: Report Selection"
    mAddTaskName ilNumber, "AffLogin.Frm", "Log In"
    mAddTaskName ilNumber, "AffLogInactivityRpt.Frm", "Log Inactivity: Report Selection"
    mAddTaskName ilNumber, "AffMain.Frm", "Main"
    mAddTaskName ilNumber, "AffMarkAssignRpt.Frm", "Market Assignment: Report Selection"
    mAddTaskName ilNumber, "AffMerge.Frm", "Merge"
    mAddTaskName ilNumber, "AffMessages.Frm", "Message Folder Info"
    mAddTaskName ilNumber, "AffModel.Frm", "Model"
    mAddTaskName ilNumber, "AffMstPict.Frm", "Background Picture"
    mAddTaskName ilNumber, "AffMultiName.Frm", "Names"
    mAddTaskName ilNumber, "AffOptions.Frm", "User Options"
    mAddTaskName ilNumber, "AffPgmClrRpt.Frm", "Program Clearance: Report Selection"
    mAddTaskName ilNumber, "AffPldgAirRpt.Frm", "Pledge vs Aired: Report Selection"
    mAddTaskName ilNumber, "AffPledgeRpt.Frm", "Pledge: Report Selection"
    mAddTaskName ilNumber, "AffPostActivityRpt.Frm", "Web Log Activity: Report Selection"
    mAddTaskName ilNumber, "AffPostLog.Frm", "Post Log"
    mAddTaskName ilNumber, "AffProgressMsg.Frm", "Progress Message"
    mAddTaskName ilNumber, "AffRadarExport.Frm", "RADAR Export"
    mAddTaskName ilNumber, "AffRadarProgSchd.Frm", "RADAR Program Schedule"
    mAddTaskName ilNumber, "AffReports.Frm", "Report List"
    mAddTaskName ilNumber, "AffRgAssgnRpt.Frm", "Regional Copy Assignment: Report Selection"
    mAddTaskName ilNumber, "AffRptNoSel.Frm", "Report with No Selection"
    mAddTaskName ilNumber, "AffSelRemap.Frm", "Re-mapped Affiliate"
    mAddTaskName ilNumber, "AffSendReport.Frm", "Crystal Interface"
    mAddTaskName ilNumber, "AffSetDDFFields.Frm", "Set New Fields"
    mAddTaskName ilNumber, "AffSiteOptions.Frm", "Site Options"
    mAddTaskName ilNumber, "AffSpotUtil.Frm", "Spot Utility"
    mAddTaskName ilNumber, "AffStation.Frm", "Station"
    mAddTaskName ilNumber, "AffStationMktInfo.Frm", "DMA Market Information"
    mAddTaskName ilNumber, "AffStationMSAMktInfo.Frm", "MSA Market Information"
    mAddTaskName ilNumber, "AffStationOwnerInfo.Frm", "Owner Information"
    mAddTaskName ilNumber, "AffStationRpt.Frm", "Station: Report Selection"
    mAddTaskName ilNumber, "AffStationSearch.Frm", "Affiliate Management"
    mAddTaskName ilNumber, "AffStationSearchFilter.Frm", "Affiliate Management Filter"
    mAddTaskName ilNumber, "AffStatioZone.Frm", "Station Agreement Update"
    mAddTaskName ilNumber, "AffTitle.Frm", "Personnel Title"
    mAddTaskName ilNumber, "AffUserLogEMail.Frm", "User Log E-Mail"
    mAddTaskName ilNumber, "AffUserOptionsRpt.Frm", "User Options: Report Selection"
    mAddTaskName ilNumber, "AffUsersLog.Frm", "Users Log"
    mAddTaskName ilNumber, "AffVehAffRpt.Frm", "Agreement: Report Selection"
    mAddTaskName ilNumber, "AffVerifyRpt.Frm", "Feed Verification: Report Selection"
    'mAddTaskName ilNumber, "AffViewReport.Frm", "Preview Crystal Report Interface"
    mAddTaskName ilNumber, "AffWebEMail.Frm", "Web E-Mail"
    mAddTaskName ilNumber, "AffWebExportSchdSpot.Frm", "Web Export Schedule Spots"
    mAddTaskName ilNumber, "AffWebImportAiredSpot.Frm", "Web Import Aired Spots"
    mAddTaskName ilNumber, "AffWebIniOptions.Frm", "Web INI Options"
    mAddTaskName ilNumber, "AffXMLTestMode.Frm", "XML Test Mode"
    ReDim Preserve tgTaskNameMap(0 To ilNumber) As TASKNAMEMAP
    For ilLoop = 0 To UBound(tgTaskNameMap) - 1 Step 1
        tgTaskNameMap(ilLoop).sFormName = UCase(tgTaskNameMap(ilLoop).sFormName)
    Next ilLoop
    ArraySortTyp fnAV(tgTaskNameMap(), 0), UBound(tgTaskNameMap), 0, LenB(tgTaskNameMap(0)), 0, LenB(tgTaskNameMap(0).sFormName), 0
End Sub

Private Function mBinarySearchTaskName(slFormName As String) As String

    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim ilResult As Integer
    
    mBinarySearchTaskName = slFormName
    llMin = LBound(tgTaskNameMap)
    llMax = UBound(tgTaskNameMap) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        ilResult = StrComp(Trim(tgTaskNameMap(llMiddle).sFormName), slFormName, vbBinaryCompare)
        Select Case ilResult
            Case 0:
                mBinarySearchTaskName = tgTaskNameMap(llMiddle).sUafName  ' Found it !
                Exit Function
            Case 1:
                llMax = llMiddle - 1
            Case -1:
                llMin = llMiddle + 1
        End Select
    Loop
    Exit Function
End Function


