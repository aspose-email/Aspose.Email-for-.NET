Imports System.Threading
Imports Aspose.Email.Exchange
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.CSharp.Email.Exchange
    Class GetContactsExtendedProperties
        Public Shared Sub Run()
            Try
                ' ExStart:GetContactsExtendedProperties
                Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

                ' The required extended properties must be added in order to create or read them from the Exchange Server
                Dim extraFields As String() = New String() {"User Field 1", "User Field 2", "User Field 3", "User Field 4"}
                For Each extraField As String In extraFields
                    client.ContactExtendedPropertiesDefinition.Add(extraField)
                Next

                ' Create a test contact on the Exchange Server
                Dim contact As New Contact()
                contact.DisplayName = "EMAILNET-38433 - " + Guid.NewGuid().ToString()
                For Each extraField As String In extraFields
                    contact.ExtendedProperties.Add(extraField, extraField)
                Next

                Dim contactId As String = client.CreateContact(contact)

                ' retrieve the contact back from the server after some time
                Thread.Sleep(5000)
                contact = client.GetContact(contactId)

                ' Parse the extended properties of contact 
                For Each extraField As String In extraFields
                    If contact.ExtendedProperties.ContainsKey(extraField) Then
                        Console.WriteLine(contact.ExtendedProperties(extraField).ToString())
                    End If
                    ' ExEnd:GetContactsExtendedProperties
                Next
            Catch ex As Exception
                Console.Write(ex.Message)
                Throw
            End Try
        End Sub
    End Class
End Namespace