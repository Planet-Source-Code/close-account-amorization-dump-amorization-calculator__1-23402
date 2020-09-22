Attribute VB_Name = "modMain"
'Loan amorization calculator by David Fiala
'djf1010@aol.com - May 24 2001

Public Sub AddNewRow(lngMonth As Long, sinBalance As Single, sinTotalPrincipal As Single, sinToPrincipal As Single, sinToInterest As Single)
    On Error GoTo ErrorMan
    sinBalance = CustomConvert(sinBalance)
    sinTotalPrincipal = CustomConvert(sinTotalPrincipal)
    sinToPrincipal = CustomConvert(sinToPrincipal)
    sinToInterest = CustomConvert(sinToInterest)
    Dim lstitmAmor As ListItem
    Set lstitmAmor = frmAmorization.lstAmor.ListItems.Add(, , lngMonth)
    With lstitmAmor
        .SubItems(1) = "$" & sinBalance
        .SubItems(2) = "$" & sinTotalPrincipal
        .SubItems(3) = "$" & sinToPrincipal
        .SubItems(4) = "$" & sinToInterest
    End With
    Exit Sub
ErrorMan:
    Call NormalErrMan
End Sub

Public Function CustomConvert(sinNumber As Single) As Currency
    On Error GoTo ErrorMan
    If InStr(1, sinNumber, ".") = 0 Then
        CustomConvert = sinNumber
        Exit Function
    End If
    CustomConvert = Mid(sinNumber, 1, InStr(1, sinNumber, ".", vbTextCompare) + 2)
    Exit Function
ErrorMan:
    Call NormalErrMan
End Function

Public Sub NormalErrMan()
    On Error GoTo KillApp
    Dim frm As Form
    Select Case Err.Number
        Case 6
            MsgBox "This application's calculations have gone over thier maximum limit. Ending application..."
        Case 13
            MsgBox "This application has received the wrong kind of data. Ending application..."
        Case Else
            MsgBox "Unexpected error. Ending application..."
    End Select
    For Each frm In Forms
        Unload frm
    Next
KillApp:
    MsgBox "An error occurred while attempting to shutdown. This application will now forcefully end.", vbCritical, "FATAL ERROR"
    End
End Sub
