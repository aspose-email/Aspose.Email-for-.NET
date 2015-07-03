//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using System;
using Aspose.Email;
using Aspose.Email.Exchange;
using Aspose.Email.Mail;
using System.Net;

namespace CSharp.Exchange
{
    class SaveExchangeTaskToDisc
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Exchange();
            string dstEmail = dataDir + "Message.msg";

            ExchangeTask task = new ExchangeTask();
            task.Subject = "TASK-ID - " + Guid.NewGuid();
            task.Status = ExchangeTaskStatus.InProgress;
            task.StartDate = DateTime.Now;
            task.DueDate = task.StartDate.AddDays(3);
            task.Save(dstEmail);

            Console.WriteLine(Environment.NewLine + "Task saved on disk successfully at " + dstEmail);
        }
    }
}
