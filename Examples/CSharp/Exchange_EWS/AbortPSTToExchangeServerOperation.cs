using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Common.Delegate;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CSharp.Exchange_EWS
{
    public class CustomAbortRestoreException : Exception { }

    public class AbortPSTToExchangeServerOperation
    {
        public static void Run()
        {
            //ExStart: AbortPSTToExchangeServerOperation
            using (IEWSClient client = EWSClient.GetEWSClient("https://exchange.office365.com/ews/exchange.asmx", "username", "password"))
            {
                DateTime startTime = DateTime.Now;
                TimeSpan maxRestoreTime = TimeSpan.FromSeconds(15);
                int processedItems = 0;

                BeforeItemCallback callback = delegate
                {
                    if (DateTime.Now >= startTime.Add(maxRestoreTime))
                    {
                        throw new CustomAbortRestoreException();
                    }

                    processedItems++;
                };

                try
                {
                    //create a test pst and add some test messages to it
                    var pst = PersonalStorage.Create(new MemoryStream(), FileFormatVersion.Unicode);
                    var folder = pst.RootFolder.AddSubFolder("My test folder");
                    for (int i = 0; i < 20; i++)
                    {
                        var message = new MapiMessage("from@gmail.com", "to@gmail.com", "subj", new string('a', 10000));
                        folder.AddMessage(message);
                    }

                    //now restore the PST with callback
                    client.Restore(pst, new Aspose.Email.Clients.Exchange.WebService.RestoreSettings
                    {
                        BeforeItemCallback = callback
                    });
                    Console.WriteLine("Success!");
                }
                catch (CustomAbortRestoreException)
                {
                    Console.WriteLine($"Timeout! {processedItems}");
                }
                //ExEnd: AbortPSTToExchangeServerOperation
            }
        }
    }
}
