using System;

namespace ConsoleColorsClean
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.Error.WriteLine(
                    "Usage: {0} dump | remove | flags FILE.lnk",
                    typeof(Program).Assembly.GetName().Name);
                return 1;
            }

            try
            {
                if (args[0].Equals("dump"))
                {
                    Shortcut.DumpConsoleInfo(Console.Out, args[1]);
                }
                else if (args[0].Equals("flags"))
                {
                    Console.WriteLine(Shortcut.GetFlags(args[1]));
                }
                else if (args[0].Equals("remove"))
                {
                    Shortcut.RmProps(args[1]);
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
                return ex.HResult;
            }
        }
    }
}
