Imports System.IO
Imports Aspose.Email.AntiSpam
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
    Class BayesianSpamAnalyzer
        Public Shared Sub Run()
            ' ExStart:PrintHeaderUsingMailMessageSaveOptions
            Dim hamFolder As String = RunExamples.GetDataDir_Email() + "/hamFolder"
            Dim spamFolder As String = RunExamples.GetDataDir_Email() + "/Spam"
            Dim testFolder As String = RunExamples.GetDataDir_Email()
            Dim dataBaseFile As String = RunExamples.GetDataDir_Email() + "SpamFilterDatabase.txt"

            TeachAndCreateDatabase(hamFolder, spamFolder, dataBaseFile)
            Dim testFiles As String() = Directory.GetFiles(testFolder, "*.eml")
            Dim analyzer As New SpamAnalyzer(dataBaseFile)
            For Each file As String In testFiles
                Dim msg As MailMessage = MailMessage.Load(file)
                Console.WriteLine(msg.Subject)
                Dim probability As Double = analyzer.Test(msg)
                PrintResult(probability)
            Next
            ' ExEnd:PrintHeaderUsingMailMessageSaveOptions
        End Sub

        Private Shared Sub TeachAndCreateDatabase(hamFolder As String, spamFolder As String, dataBaseFile As String)
            Dim hamFiles As String() = Directory.GetFiles(hamFolder, "*.eml")
            Dim spamFiles As String() = Directory.GetFiles(spamFolder, "*.eml")

            Dim analyzer As New SpamAnalyzer()

            For i As Integer = 0 To hamFiles.Length - 1
                Dim hamMailMessage As MailMessage
                Try
                    hamMailMessage = MailMessage.Load(hamFiles(i))
                Catch generatedExceptionName As Exception
                    Continue For
                End Try

                Console.WriteLine(i)
                analyzer.TrainFilter(hamMailMessage, False)
            Next

            For i As Integer = 0 To spamFiles.Length - 1
                Dim spamMailMessage As MailMessage
                Try
                    spamMailMessage = MailMessage.Load(hamFiles(i))
                Catch generatedExceptionName As Exception
                    Continue For
                End Try

                Console.WriteLine(i)
                analyzer.TrainFilter(spamMailMessage, True)
            Next

            analyzer.SaveDatabase(dataBaseFile)
        End Sub

        Private Shared Sub PrintResult(probability As Double)
            If probability < 0.05 Then
                Console.WriteLine("This is ham)")
            ElseIf probability > 0.95 Then
                Console.WriteLine("This is spam)")
            Else
                Console.WriteLine("Maybe spam)")
            End If
        End Sub
    End Class
End Namespace