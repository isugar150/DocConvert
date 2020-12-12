using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocConvert_Console.Common
{
    public class LogMgr
    {
        public static void ConsoleWrite(string text, ConsoleColor color)
        {
            ConsoleColor originalColor = Console.ForegroundColor;
            Console.ForegroundColor = color;

            Console.WriteLine(text);
            Console.Out.Flush();

            Console.ForegroundColor = originalColor;
        }
    }
}
