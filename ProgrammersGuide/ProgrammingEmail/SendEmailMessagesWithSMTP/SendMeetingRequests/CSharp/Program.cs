//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Email;
using Aspose.Email.Mail;
using System;

namespace SendMeetingRequests
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Create an instance of the MailMessage class
            MailMessage msg = new MailMessage();

            //set the sender
            msg.From = "newcustomeronnet@gmail.com"; 

            //Set the recipient, who will receive the meeting request.
            //Basically, the recipient is the same as the meeting attendees.
            msg.To = "person1@domain.com, person2@domain.com, person3@domain.com, asposetest123@gmail.com";

            //create Appointment instance
            Appointment app = new Appointment(
            "Room 112", //location
            new DateTime(2006, 7, 17, 13, 0, 0), //start time
            new DateTime(2006, 7, 17, 14, 0, 0), //end time
            msg.From,//organizer
            msg.To //attendee
            );

            app.Summary = "Release Meetting";
            app.Description = "Discuss for the next release";

            //add appointment to the message
            msg.AddAlternateView(app.RequestApointment());

            //Create an instance of SmtpClient class
            SmtpClient client = GetSmtpClient();

            try
            {
                //Client.Send will send this message
                client.Send(msg);
                //Show Message SentÅE only if message sent successfully
                Console.WriteLine("Message sent");
            }

            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
            }
            
        }
        private static SmtpClient GetSmtpClient()
        {
            SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "asposetest123@gmail.com", "F123456f");
            client.SecurityMode = SmtpSslSecurityMode.Explicit;
            client.EnableSsl = true;

            return client;
        }
    }
}