// Demonstrates how to keep watching a folder after monitoring has failed.
//
// StartMonitoring reports new and deleted messages through a callback. When something
// goes wrong - most often the connection drops - monitoring stops and the error
// callback receives the exception together with a MonitoringState. Handing that state
// to ResumeMonitoring picks up where monitoring stopped, and the changes that happened
// in between are reported too, so none are lost.

using System;
using System.Threading;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ResumeMonitoringAfterError
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            using (var failed = new ManualResetEventSlim(false))
            {
                IImapMonitoringState lastState = null;

                ImapMonitoringEventHandler onChange = (sender, e) =>
                {
                    foreach (var info in e.NewMessages)
                        Console.WriteLine($"{DateTime.Now:T}  new in {e.FolderName}: {info.Subject}");
                    if (e.DeletedMessages.Length > 0)
                    {
                        Console.WriteLine($"{DateTime.Now:T}  {e.DeletedMessages.Length} message(s) " +
                                          $"deleted from {e.FolderName}");
                    }
                };

                ImapMonitoringErrorEventHandler onError = (sender, e) =>
                {
                    Console.WriteLine($"{DateTime.Now:T}  monitoring of {e.FolderName} stopped: {e.Error.Message}");
                    lastState = e.MonitoringState;
                    failed.Set();
                };

                client.StartMonitoring(onChange, onError, ImapFolderInfo.InBox);
                Console.WriteLine($"{DateTime.Now:T}  watching the Inbox for 60 seconds - send yourself a message.");

                if (failed.Wait(TimeSpan.FromSeconds(60)) && lastState != null)
                {
                    failed.Reset();
                    client.ResumeMonitoring(onChange, onError, lastState);
                    Console.WriteLine($"{DateTime.Now:T}  monitoring of {lastState.FolderName} resumed for 60 seconds.");

                    failed.Wait(TimeSpan.FromSeconds(60));
                }

                client.StopMonitoring();
                Console.WriteLine($"{DateTime.Now:T}  monitoring stopped.");
            }
        }
    }
}
