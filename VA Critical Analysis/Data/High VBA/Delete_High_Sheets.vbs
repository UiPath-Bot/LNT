Sub Delete_High()
On Error Resume Next
Application.DisplayAlerts = False

Sheets("High_Repeated").Delete

Application.DisplayAlerts = True

Application.DisplayAlerts = False

Sheets("High_VA_Review_Analysis").Delete

Application.DisplayAlerts = True

Application.DisplayAlerts = False

Sheets("High_Repeated_Analysis").Delete

Application.DisplayAlerts = True
End Sub


