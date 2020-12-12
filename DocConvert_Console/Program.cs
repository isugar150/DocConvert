using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocConvert_Console.Common;

namespace DocConvert_Console
{
    class Program
    {
        static void Main(string[] args)
        {
            LogMgr.ConsoleWrite(@"
          _____              _____                          _   
         |  __ \            / ____|                        | |  
         | |  | | ___   ___| |     ___  _ ____   _____ _ __| |_ 
         | |  | |/ _ \ / __| |    / _ \| '_ \ \ / / _ \ '__| __|
         | |__| | (_) | (__| |___| (_) | | | \ V /  __/ |  | |_ 
         |_____/ \___/ \___|\_____\___/|_| |_|\_/ \___|_|   \__|
          / ____|                                               
         | (___   ___ _ ____   _____ _ __                       
          \___ \ / _ \ '__\ \ / / _ \ '__|                      
          ____) |  __/ |   \ V /  __/ |                         
         |_____/ \___|_|    \_/ \___|_|  
                ", ConsoleColor.Green);
        }
    }
}
