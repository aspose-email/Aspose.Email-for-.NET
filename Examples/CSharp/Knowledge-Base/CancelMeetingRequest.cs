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
    public partial class CancelMeetingRequest : Form
    {
        SqlConnection connection;

        public CancelMeetingRequest()
        {
            InitializeComponent();

            string connectionString = "Data Source =(local)\\SQLEXPRESS;Initial Catalog = Test; Integrated Security = true";
            connection = new SqlConnection(connectionString);

            LoadAppointments();
        }

        private void LoadAppointments()
        {
            try
            {
                // Get the message and calendar information from the database get the attendee information
                string strSQLGetMessages = "SELECT * FROM Message";

                // Get the attendee information in reader
                SqlCommand command = new SqlCommand(strSQLGetMessages, connection);
                command.CommandType = CommandType.Text;

                OpenConnection();
                SqlDataReader reader = command.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(reader);
                dgMeetings.DataSource = dt;
                CloseConnection();              
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // ExStart:CancelMeetingRequest
            try
            {
                OpenConnection();

                // Get the message ID of the selected row
                string strMessageID = dgMeetings.SelectedRows[0].Cells["MessageID"].Value.ToString();

                // Get the message and calendar information from the database get the attendee information
                string strSQLGetAttendees = "SELECT * FROM Attendent WHERE MessageID = '" + strMessageID + "' ";

                // Get the attendee information in reader
                SqlDataReader rdAttendees = ExecuteReader(strSQLGetAttendees);

                // Create a MailAddressCollection from the attendees found in the database
                MailAddressCollection attendees = new MailAddressCollection();
                while (rdAttendees.Read())
                {
                    attendees.Add(new MailAddress(rdAttendees["EmailAddress"].ToString(), rdAttendees["DisplayName"].ToString()));
                }

                rdAttendees.Close();

                // Get message and calendar information
                string strSQLGetMessage = "SELECT * FROM [Message] WHERE MessageID = '" + strMessageID + "' ";

                // Get the reader for the above query
                SqlDataReader rdMessage = ExecuteReader(strSQLGetMessage);
                rdMessage.Read();

                // Create the Calendar object from the database information
                Appointment app = new Appointment(rdMessage["CalLocation"].ToString(), rdMessage["CalSummary"].ToString(), rdMessage["CalDescription"].ToString(),
                DateTime.Parse(rdMessage["CalStartTime"].ToString()), DateTime.Parse(rdMessage["CalEndTime"].ToString()), new MailAddress(rdMessage["FromEmailAddress"].ToString(), rdMessage["FromDisplayName"].ToString()), attendees);

                // Create message and Set from and to addresses and Add the cancel meeting request
                MailMessage msg = new MailMessage();
                msg.From = new MailAddress(rdMessage["FromEmailAddress"].ToString(), "");
                msg.To = attendees;
                msg.Subject = "Cencel meeting";
                msg.AddAlternateView(app.CancelAppointment());
                SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587, "user@gmail.com", "password");
                smtp.Send(msg);
                MessageBox.Show("cancellation request successfull");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                CloseConnection();
            }
            // ExEnd:CancelMeetingRequest
        }


        public SqlDataReader ExecuteReader(string stringCommand)
        {
            SqlDataReader reader = null;
            try
            {
                
                SqlCommand command = new SqlCommand(stringCommand, connection);
                command.CommandType = CommandType.Text;
                reader = command.ExecuteReader();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return reader;
        }

        public void OpenConnection()
        {            
            connection.Open();
        }

        public void CloseConnection()
        {
            if (connection != null && connection.State == System.Data.ConnectionState.Open)
            {
                connection.Close();
            }
        }
    }
}
