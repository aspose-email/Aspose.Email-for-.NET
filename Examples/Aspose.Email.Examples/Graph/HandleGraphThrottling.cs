// Demonstrates how to survive Microsoft Graph throttling.
//
// Graph answers a client that asks too often with HTTP 429 and a Retry-After header.
// A retry policy makes the client wait and try again on its own; without one, the call
// surfaces a GraphThrottlingException that the caller has to handle.

using System;
using Aspose.Email.Clients.Graph;

namespace Aspose.Email.Examples.Graph
{
    internal static class HandleGraphThrottling
    {
        public static void Run()
        {
            if (!ClientBuilder.IsGraphConfigured)
            {
                GraphExampleInfo.PrintNotConfigured();
                return;
            }

            using (var client = ClientBuilder.Graph(AuthType.ModernWithAppPermission))
            {
                client.RetryPolicy = new GraphRetryPolicy
                {
                    Enabled = true,
                    MaxRetryAttempts = 5,

                    // The first wait, doubled on each further attempt.
                    BaseDelay = TimeSpan.FromSeconds(1),

                    // A ceiling, so backoff cannot grow without bound.
                    MaxDelay = TimeSpan.FromSeconds(30),

                    // Honour the server's Retry-After when it sends one, and fall back to
                    // exponential backoff with jitter when it does not. The jitter keeps
                    // many clients from retrying in lockstep.
                    DelayStrategy = GraphRetryDelayStrategy.RetryAfterThenExponentialBackoffWithJitter
                };

                Console.WriteLine($"Retry policy: {client.RetryPolicy.MaxRetryAttempts} attempt(s), " +
                                  $"{client.RetryPolicy.BaseDelay.TotalSeconds}s base, " +
                                  $"{client.RetryPolicy.MaxDelay.TotalSeconds}s max, " +
                                  $"{client.RetryPolicy.DelayStrategy}");

                try
                {
                    var messages = client.ListMessages(KnownFolders.Inbox, null);
                    Console.WriteLine($"Listed {messages.Count} message(s).");
                }
                catch (GraphThrottlingException ex)
                {
                    // Reached only when the retries were used up, or when the policy is off.
                    Console.WriteLine("Still throttled after the retries:");
                    Console.WriteLine($"  status code: {ex.StatusCode}");
                    Console.WriteLine($"  retry after: {ex.RetryAfter}");
                    Console.WriteLine($"  request id:  {ex.RequestId}");
                }

                // GraphRetryPolicy.Default is the library's own setting, and assigning a
                // policy with Enabled = false turns retrying off entirely.
                Console.WriteLine($"\nLibrary default: enabled={GraphRetryPolicy.Default.Enabled}, " +
                                  $"attempts={GraphRetryPolicy.Default.MaxRetryAttempts}");
            }
        }
    }
}
