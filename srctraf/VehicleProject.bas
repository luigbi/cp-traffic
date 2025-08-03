Attribute VB_Name = "VehicleProject"
Public imFirstActivate As Integer
Public imPopReqd As Integer         'Flag indicating if lbcStep(2) was populated


'*******************************************************
'*                                                     *
'*      Procedure Name:gShowBranner                    *
'*                                                     *
'*             Created:4/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Show branner in main title bar  *
'*                                                     *
'*******************************************************
Sub gShowBranner(ilUpdateAllowed As Integer)
'
'   gShowBranner
'   Where:
'
    Dim sAllowed As String
    Dim slName As String
    Dim slDateTime As String

        If ilUpdateAllowed Then
            sAllowed = sAllowed & ", Input OK"
        Else
            sAllowed = sAllowed & ", View Only"
        End If
        If Trim$(tgUrf(0).sRept) <> "" Then
            slName = Trim$(tgUrf(0).sRept)
        Else
            slName = sgUserName
        End If
End Sub
