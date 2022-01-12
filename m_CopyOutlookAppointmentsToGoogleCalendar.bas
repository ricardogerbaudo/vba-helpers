'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Purpose:  This module loops through all appointments in your Outlook Calendar and
'           creates a HTTP Request that fires an IFTTT webhooker which creates a 
'           Google Calendar event with the parameters provided in the URL query string.
' Author:   Ricardo Gerbaudo
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub ListAppointments()

    Dim oAppointments As Object
    Dim oAppointmentItem As Outlook.AppointmentItem
    
    Set oAppointments = Outlook.Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar)
    
    For Each oAppointmentItem In oAppointments.Items
        Debug.Print oAppointmentItem.Start & vbTab & oAppointmentItem.Subject
        SendRequest oAppointmentItem.Subject, oAppointmentItem.Start, oAppointmentItem.End
        DoEvents
    Next

End Sub

Sub SendRequest(strEventTitle As String, startDate As Date, endDate As Date)
        
    Dim strStartDate As String
    Dim strEndDate As String
    
    strStartDate = Format(startDate, "mm/dd/yyyy hh:mm")
    strEndDate = Format(endDate, "mm/dd/yyyy hh:mm")
    
    Dim strUrl As String
    strUrl = "https://maker.ifttt.com/trigger/{YOUR_EVENT_NAME_HERE}/with/key/{YOUR_KEY_HERE}?value1=" & strEventTitle & "&value2=" & strStartDate & "&value3=" & strEndDate & ""
    
    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")

    With httpRequest
        .Open "GET", strUrl, False
        .Send
        Debug.Print .ResponseText
    End With
    
End Sub
