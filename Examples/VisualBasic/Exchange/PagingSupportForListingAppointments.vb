Imports Aspose.Email.Exchange
Imports Aspose.Email.Mail

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
    Public Class PagingSupportForListingAppointments
        Private Shared Sub Run()
            'ExStart: PagingSupportForListingAppointments
            Using client As IEWSClient = EWSClient.GetEWSClient("exchange.domain.com", "username", "password")
                Try
                    Dim appts As Appointment() = client.ListAppointments()
                    Console.WriteLine(appts.Length)
                    Dim [date] As DateTime = DateTime.Now
                    Dim startTime As New DateTime([date].Year, [date].Month, [date].Day, [date].Hour, 0, 0)
                    Dim endTime As DateTime = startTime.AddHours(1)
                    Dim appNumber As Integer = 10
                    Dim appointmentsDict As New Dictionary(Of String, Appointment)()
                    For i As Integer = 0 To appNumber - 1
                        startTime = startTime.AddHours(1)
                        endTime = endTime.AddHours(1)
                        Dim timeZone As String = "America/New_York"
                        Dim appointment As New Appointment("Room 112", startTime, endTime, "from@domain.com", "to@domain.com")
                        appointment.SetTimeZone(timeZone)
                        appointment.Summary = "NETWORKNET-35157_3 - " + Guid.NewGuid().ToString()
                        appointment.Description = "EMAILNET-35157 Move paging parameters to separate class"
                        Dim uid As String = client.CreateAppointment(appointment)
                        appointmentsDict.Add(uid, appointment)
                    Next
                    Dim totalAppointmentCol As AppointmentCollection = client.ListAppointments()

                    ' LISTING APPOINTMENTS WITH PAGING SUPPORT /
                    Dim itemsPerPage As Integer = 2
                    Dim pages As New List(Of AppointmentPageInfo)()
                    Dim pagedAppointmentCol As AppointmentPageInfo = client.ListAppointmentsByPage(itemsPerPage)
                    Console.WriteLine(pagedAppointmentCol.Items.Count)
                    pages.Add(pagedAppointmentCol)
                    While Not pagedAppointmentCol.LastPage
                        pagedAppointmentCol = client.ListAppointmentsByPage(itemsPerPage, pagedAppointmentCol.PageOffset + 1)
                        pages.Add(pagedAppointmentCol)
                    End While
                    Dim retrievedItems As Integer = 0
                    For Each folderCol As AppointmentPageInfo In pages
                        retrievedItems += folderCol.Items.Count
                    Next
                Finally
                End Try
            End Using
            'ExEnd: PagingSupportForListingAppointments
        End Sub
    End Class
End Namespace