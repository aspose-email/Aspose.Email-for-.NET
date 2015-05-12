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
using Aspose.Email.Outlook;

namespace DraftAppointmentRequest
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");
            Directory.CreateDirectory(dataDir);

            string sender = "test@gmail.com";
            string recipient = "test@email.com";

            MailMessage message = new MailMessage(sender, recipient, string.Empty, string.Empty);

            Appointment app = new Appointment(string.Empty, DateTime.Now, DateTime.Now, sender, recipient);
            app.Method = AppointmentMethodType.Publish;

            message.AddAlternateView(app.RequestApointment());

            MapiMessage msg = MapiMessage.FromMailMessage(message);

            // Save the appointment as draft.
            msg.Save(dataDir + "draftAppointment.msg");
        }
    }
}