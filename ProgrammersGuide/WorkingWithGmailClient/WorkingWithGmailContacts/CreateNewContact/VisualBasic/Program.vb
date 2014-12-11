'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Email. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Email
Imports Aspose.Email.Exchange.Schema
Imports Aspose.Email.Tests
Imports Aspose.Email.Google
Imports Aspose.Email.Mail
Imports System

Namespace CreateNewContactExample
	Public Class Program
		Public Shared Sub Main()
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")
			Dim username As String = "aspose.examples"
			Dim email As String = "aspose.examples@gmail.com"
			Dim password As String = "aspose123"
			Dim clientId As String = "491198589534-c08fdgnpu3ausn9pbdjjjh9r2s4vlla0.apps.googleusercontent.com"
			Dim clientSecret As String = "Km9G7a6PivZv4dDxZ3-ZVPk2"
			Dim refresh_token As String = "1/UlvQ1pvWN5DQLgHd8LmcNchPs50s7E96sTnjdwHKVS8"

			'The refresh_token is to be used in below.
			Dim user As New GoogleTestUser(username, email, password, clientId, clientSecret, refresh_token) 'refresh token

			'Gmail Client
			Dim client As IGmailClient = Aspose.Email.Google.GmailClient.GetInstance(user.ClientId, user.ClientSecret, user.RefreshToken, user.EMail)

			'Create a Contact
			Dim contact As New Contact()
			contact.Prefix = "Prefix"
			contact.GivenName = "GivenName"
			contact.Surname = "Surname"
			contact.MiddleName = "MiddleName"
			contact.DisplayName = "Test User 1"
			contact.Suffix = "Suffix"

			contact.JobTitle = "JobTitle"
			contact.DepartmentName = "DepartmentName"
			contact.CompanyName = "CompanyName"
			contact.Profession = "Profession"

			contact.Notes = "Notes"

			Dim address As New PostalAddress()
			address.Category = PostalAddressCategory.Work
			address.Address = "Address"
			address.Street = "Street"
			address.PostOfficeBox = "PostOfficeBox"
			address.City = "City"
			address.StateOrProvince = "StateOrProvince"
			address.PostalCode = "PostalCode"
			address.Country = "Country"
			contact.PhysicalAddresses.Add(address)

			Dim pnWork As New PhoneNumber()
			pnWork.Number = "323423423423"
			pnWork.Category = PhoneNumberCategory.Work
			contact.PhoneNumbers.Add(pnWork)
			Dim pnHome As New PhoneNumber()
			pnHome.Number = "323423423423"
			pnHome.Category = PhoneNumberCategory.Home
			contact.PhoneNumbers.Add(pnHome)
			Dim pnMobile As New PhoneNumber()
			pnMobile.Number = "323423423423"
			pnMobile.Category = PhoneNumberCategory.Mobile
			contact.PhoneNumbers.Add(pnMobile)

			contact.Urls.Blog = "Blog.ru"
			contact.Urls.BusinessHomePage = "BusinessHomePage.ru"
			contact.Urls.HomePage = "HomePage.ru"
			contact.Urls.Profile = "Profile.ru"

			contact.Events.Birthday = DateTime.Now.AddYears(-30)
			contact.Events.Anniversary = DateTime.Now.AddYears(-10)

			contact.InstantMessengers.AIM = "AIM"
			contact.InstantMessengers.GoogleTalk = "GoogleTalk"
			contact.InstantMessengers.ICQ = "ICQ"
			contact.InstantMessengers.Jabber = "Jabber"
			contact.InstantMessengers.MSN = "MSN"
			contact.InstantMessengers.QQ = "QQ"
			contact.InstantMessengers.Skype = "Skype"
			contact.InstantMessengers.Yahoo = "Yahoo"

			contact.AssociatedPersons.Spouse = "Spouse"
			contact.AssociatedPersons.Sister = "Sister"
			contact.AssociatedPersons.Relative = "Relative"
			contact.AssociatedPersons.ReferredBy = "ReferredBy"
			contact.AssociatedPersons.Partner = "Partner"
			contact.AssociatedPersons.Parent = "Parent"
			contact.AssociatedPersons.Mother = "Mother"
			contact.AssociatedPersons.Manager = "Manager"

			'Email Address
			Dim eAddress As New Aspose.Email.Mail.EmailAddress()
			eAddress.Address = "email@gmail.com"
			contact.EmailAddresses.Add(eAddress)

			Dim contactUri As String = client.CreateContact(contact)

		End Sub
	End Class
End Namespace