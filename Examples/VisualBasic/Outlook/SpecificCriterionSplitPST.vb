Imports System.Collections.Generic
Imports Aspose.Email.Outlook.Pst

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class SpecificCriterionSplitPST
        Public Shared Sub Run()
            ' The path to the File directory.
            ' ExStart:SpecificCriterionSplitPST
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim criteria As IList(Of MailQuery) = New List(Of MailQuery)()
            Dim pstQueryBuilder As New PersonalStorageQueryBuilder()
            pstQueryBuilder.SentDate.Since(New DateTime(2005, 4, 1))
            pstQueryBuilder.SentDate.Before(New DateTime(2005, 4, 7))
            criteria.Add(pstQueryBuilder.GetQuery())
            pstQueryBuilder = New PersonalStorageQueryBuilder()
            pstQueryBuilder.SentDate.Since(New DateTime(2005, 4, 7))
            pstQueryBuilder.SentDate.Before(New DateTime(2005, 4, 13))
            criteria.Add(pstQueryBuilder.GetQuery())

            Using personalStorage__1 As PersonalStorage = PersonalStorage.FromFile(dataDir & Convert.ToString("PersonalStorage_New.pst"))
                personalStorage__1.SplitInto(criteria, dataDir & Convert.ToString("pathToPst"))
            End Using
            ' ExEnd:SpecificCriterionSplitPST
        End Sub
    End Class
End Namespace
