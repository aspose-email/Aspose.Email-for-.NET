Imports Aspose.Email.Exchange
Imports Aspose.Email.Mail

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http:www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http:www.aspose.com/community/forums/default.aspx            

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange

    Public Class SecondaryCalendarEvents
        Public Shared Sub Run()
            Using client As IEWSClient = EWSClient.GetEWSClient("https:outlook.office365.com/ews/exchange.asmx", "your.username", "your.password")
                Try
                    'Create an appointmenta that will be added to secondary calendar folder
                    Dim [date] As DateTime = DateTime.Now
                    Dim startTime As New DateTime([date].Year, [date].Month, [date].Day, [date].Hour, 0, 0)
                    Dim endTime As DateTime = startTime.AddHours(1)
                    Dim timeZone As String = "America/New_York"
                    Dim listAppointments As Appointment()
                    Dim appointment As New Appointment("Room 121", startTime, endTime, "from@domain.com", "attendee@domain.com")
                    appointment.SetTimeZone(timeZone)
                    appointment.Summary = "EMAILNET-35198 - " + Guid.NewGuid().ToString()
                    appointment.Description = "EMAILNET-35198 Ability to add event to Secondary Calendar of Office 365"

                    'Verify that the new folder has been created
                    Dim calendarSubFolders As ExchangeFolderInfoCollection = client.ListSubFolders(client.MailboxInfo.CalendarUri)

                    Dim getfolderName As String = ""
                    Dim setFolderName As String = "New Sub Calendar"
                    Dim alreadyExits As Boolean = False
                    For i As Integer = 0 To calendarSubFolders.Count - 1
                        getfolderName = calendarSubFolders(i).DisplayName

                        If getfolderName.Equals(setFolderName) Then
                            alreadyExits = True
                        End If
                    Next

                    If alreadyExits Then
                        Console.WriteLine("Folder Already Exists")
                    Else
                        ' create new calendar folder
                        client.CreateFolder(client.MailboxInfo.CalendarUri, setFolderName, Nothing, "IPF.Appointment")

                        Console.WriteLine(calendarSubFolders.Count)

                        'Get the created folder URI
                        Dim newCalendarFolderUri As String = calendarSubFolders(0).Uri
                        ' appointment api with calendar folder uri
                        ' create
                        client.CreateAppointment(appointment, newCalendarFolderUri)
                        appointment.Location = "Room 122"
                        ' update
                        client.UpdateAppointment(appointment, newCalendarFolderUri)
                        ' list
                        listAppointments = client.ListAppointments(newCalendarFolderUri)

                        ' list default calendar folder
                        listAppointments = client.ListAppointments(client.MailboxInfo.CalendarUri)

                        ' cancel
                        client.CancelAppointment(appointment, newCalendarFolderUri)
                        listAppointments = client.ListAppointments(newCalendarFolderUri)

                        ' appointment api with context current calendar folder uri
                        client.CurrentCalendarFolderUri = newCalendarFolderUri
                        ' create
                        client.CreateAppointment(appointment)
                        appointment.Location = "Room 122"
                        ' update
                        client.UpdateAppointment(appointment)
                        ' list
                        listAppointments = client.ListAppointments()

                        ' list default calendar folder
                        listAppointments = client.ListAppointments(client.MailboxInfo.CalendarUri)

                        ' cancel
                        client.CancelAppointment(appointment)

                        listAppointments = client.ListAppointments()

                    End If
                Finally
                End Try
            End Using
        End Sub
    End Class
End Namespace