using Aspose.Email.Calendar;
using Aspose.Email.Clients.Smtp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    public partial class SendMeetingRequests : Form
    {
        public SendMeetingRequests()
        {
            InitializeComponent();
            LoadAttendees();
        }

        private void LoadAttendees()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("FirstName", typeof(string));
            dt.Columns.Add("LastName", typeof(string));
            dt.Columns.Add("EmailAddress", typeof(string));

            dt.Rows.Add("Attendee", "1", "attendee1@gmail.com");
            dt.Rows.Add("Attendee", "2", "attendee2@gmail.com");
            dt.Rows.Add("Attendee", "3", "attendee3@gmail.com");

            dgAttendees.DataSource = dt;
        }

        // ExStart:SendMeetingRequests
        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                // Create an instance of SMTPClient
                SmtpClient client = new SmtpClient(txtMailServer.Text, txtUsername.Text, txtPassword.Text);
                // Get the attendees
                MailAddressCollection attendees = new MailAddressCollection();
                foreach (DataGridViewRow row in dgAttendees.Rows)
                {
                    if (row.Cells["EmailAddress"].Value != null)
                    {
                        // Get names and addresses from the grid and add to MailAddressCollection
                        attendees.Add(new MailAddress(row.Cells["EmailAddress"].Value.ToString(),
                        row.Cells["FirstName"].Value.ToString() + " " + row.Cells["LastName"].Value.ToString()));
                    }
                }

                // Create an instance of MailMessage for sending the invitation
                MailMessage msg = new MailMessage();

                // Set from address, attendees
                msg.From = txtFrom.Text;
                msg.To = attendees;

                // Create am instance of Appointment
                Appointment app = new Appointment(txtLocation.Text, dtTimeFrom.Value, dtTimeTo.Value, txtFrom.Text, attendees);
                app.Summary = "Monthly Meeting";
                app.Description = "Please confirm your availability.";
                msg.AddAlternateView(app.RequestApointment());

                // Save the info to the database
                if (SaveIntoDB(msg, app) == true)
                {
                    // Save the message and Send the message with the meeting request
                    msg.Save(msg.MessageId + ".eml", SaveOptions.DefaultEml);
                    client.Send(msg);
                    MessageBox.Show("message sent");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }    
        }

        private bool SaveIntoDB(MailMessage msg, Appointment app)
        {
            try
            {
                // Save message and Appointment information
                string strSQLInsertIntoMessage = @"INSERT INTO Message (MessageID, Subject, FromEmailAddress, FromDisplayName, Body,
		CalUniqueID, CalSequenceID, CalLocation, CalStartTime, CalEndTime, Caldescription, CalSummary, MeetingStatus)
		VALUES ('" + msg.MessageId + "' , '" + msg.Subject + "' , '" + msg.From[0].Address + "' , '" +
                msg.From[0].DisplayName + "' , '" + msg.Body + "' , '" + app.UniqueId + "' , '" + app.SequenceId + @"' ,
		'" + app.Location + "' , '" + app.StartDate + "' , '" + app.EndDate + "' , '" + app.Description + @"' ,
		'" + app.Summary + "' , 1); ";

                ExecuteNonQuery(strSQLInsertIntoMessage);

                // Save attendee information
                string strSQLInsertIntoAttendees = "";
                foreach (DataGridViewRow row in dgAttendees.Rows)
                {
                    if (row.Cells["EmailAddress"].Value != null)
                    {
                        strSQLInsertIntoAttendees += "INSERT INTO Attendent (MessageID, EmailAddress, DisplayName) VALUES (" + "'" + msg.MessageId + "' , '" + row.Cells["EmailAddress"].Value + "' , '" + row.Cells["FirstName"].Value.ToString() + " " + row.Cells["LastName"].Value.ToString() + "' );";
                        ExecuteNonQuery(strSQLInsertIntoAttendees);
                    }
                }
                
                // If records are successfully inserted into the database, return true
                return true;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message); // Return false in case of exception
                return false;
            }
        }
        // ExEnd:SendMeetingRequests

        public void ExecuteNonQuery(string stringCommand)
        {
            try
            {
                string connectionString = "Data Source =(local)\\SQLEXPRESS;Initial Catalog = Test; Integrated Security = true";
                SqlConnection connection = new SqlConnection(connectionString);
                SqlCommand command = new SqlCommand(stringCommand, connection);
                command.CommandType = CommandType.Text;

                connection.Open();
                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
