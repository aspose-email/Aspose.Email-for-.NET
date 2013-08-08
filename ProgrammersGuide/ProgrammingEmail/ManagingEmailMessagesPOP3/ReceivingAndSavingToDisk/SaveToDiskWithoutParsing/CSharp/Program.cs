//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Email;
using Aspose.Email.Pop3;
using System;

namespace SaveToDiskWithoutParsing
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Create an instance of the Pop3Client class
            Pop3Client client = new Pop3Client();

            //Specify host, username and password for your client
            client.Host = "pop.gmail.com";

            // Set username
            client.Username = "asposetest123@gmail.com";

            // Set password
            client.Password = "F123456f";

            // Set the port to 995. This is the SSL port of POP3 server
            client.Port = 995;

            // Enable SSL
            client.EnableSsl = true;

            try
            {
                //Connect and log in to the POP3 mail server
                client.Connect(true);

                //Save message to disk by message sequence number
                client.SaveMessage(19,dataDir + "1234.eml");

                client.Disconnect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}