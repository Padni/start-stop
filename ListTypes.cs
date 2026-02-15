using System;
using System.Reflection;
using Microsoft.Toolkit.Uwp.Notifications;

class Program
{
    static void Main()
    {
        var asm = typeof(ToastNotificationActivatedEventArgsCompat).Assembly;
        foreach (var t in asm.GetTypes())
        {
            if (t.Namespace == "Microsoft.Toolkit.Uwp.Notifications")
            {
                Console.WriteLine(t.FullName);
            }
        }
    }
}