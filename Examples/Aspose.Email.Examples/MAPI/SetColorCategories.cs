// Demonstrates how to add, list, remove and clear the colour categories of an
// Outlook message.

using System;
using System.Collections.Generic;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetColorCategories
    {
        public static void Run()
        {
            var msg = MapiMessage.Load(Data.Mapi/"message1.msg");

            FollowUpManager.AddCategory(msg, "Purple Category");
            FollowUpManager.AddCategory(msg, "Red Category");
            Print("After adding two categories", FollowUpManager.GetCategories(msg));

            FollowUpManager.RemoveCategory(msg, "Red Category");
            Print("After removing the red one", FollowUpManager.GetCategories(msg));

            FollowUpManager.ClearCategories(msg);
            Print("After clearing all", FollowUpManager.GetCategories(msg));

            var outputPath = Data.Out/"SetColorCategories_out.msg";
            msg.Save(outputPath);
            Console.WriteLine($"\nSaved to {outputPath}");
        }

        private static void Print(string stage, IList<string> categories)
        {
            var value = categories.Count == 0 ? "(none)" : string.Join(", ", categories);
            Console.WriteLine($"{stage,-28}: {value}");
        }
    }
}
