Sub Delete_Critical()
On Error Resume Next
Application.DisplayAlerts = False

Sheets("Critical_Repeated").Delete

Application.DisplayAlerts = True

Application.DisplayAlerts = False

Sheets("Critical_VA_Review_Analysis").Delete

Application.DisplayAlerts = True

Application.DisplayAlerts = False

Sheets("Critical_Repeated_Analysis").Delete

Application.DisplayAlerts = True
End Sub
