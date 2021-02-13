using System;
using System.IO;
using DocConvert_Console.Common;
using DocConvert_Console.define;
using DocConvert_Core.IniLib;

namespace DocConvert_Console
{
    class Program
    {
        public static iniProperties IniProperties = new iniProperties(); // Ini 설정값 변수
        static void Main(string[] args)
        {
            // 제품명 로그
            LogMgr.Write(@"
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

                ", ConsoleColor.Green, LOG_LEVEL.WARN);
            // 저작권 정보
            LogMgr.Write("         Copyright© 2021. Jm's Corp. All rights reserved.\r\n\r\n", ConsoleColor.White, LOG_LEVEL.WARN);
            LogMgr.Write("Starting DocConvert Server..", ConsoleColor.White, LOG_LEVEL.WARN);
            LogMgr.Write("Program Directory: " + Environment.CurrentDirectory, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write(".Net Framework Version: " + Environment.Version.ToString(), ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("OS Version: " + Environment.OSVersion.ToString(), ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("PC(Domain) Name: " + Environment.UserDomainName, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("User Name: " + Environment.UserName, ConsoleColor.White, LOG_LEVEL.INFO);
            #region parse Ini File
            LogMgr.Write("------------ Parsing Ini File ------------", ConsoleColor.White, LOG_LEVEL.INFO);
            try
            {
                IniFile pairs = new IniFile(); // Ini 변수 초기화
                if (new FileInfo(Environment.CurrentDirectory + "/DocConvert_Console.ini").Exists)
                {
                    pairs.Load(Environment.CurrentDirectory + "/DocConvert_Console.ini");
                    IniProperties.Bind_IP = pairs["Common"]["Bind IP"].ToString2().Trim();
                    IniProperties.HTTP_Port = int.Parse(pairs["Common"]["HTTP Port"].ToString2().Trim());
                    IniProperties.Client_KEY = pairs["Common"]["Client KEY"].ToString2().Trim();
                    IniProperties.Workspace_Directory = pairs["Common"]["Workspace Directory"].ToString2().Trim();

                    IniProperties.SchedulerTime = pairs["Scheduler"]["SchedulerTime"].ToString2().Trim();
                    IniProperties.CleanWorkspaceSchedulerYn = pairs["Scheduler"]["CleanWorkspaceSchedulerYn"].ToString2().Trim().Equals("Y");
                    IniProperties.CleanWorkspaceDay = int.Parse(pairs["Scheduler"]["CleanWorkspaceDay"].ToString2().Trim());
                    IniProperties.CleanLogSchedulerYn = pairs["Scheduler"]["CleanLogSchedulerYn"].ToString2().Trim().Equals("Y");
                    IniProperties.CleanLogDay = int.Parse(pairs["Scheduler"]["CleanLogDay"].ToString2().Trim());

                    IniProperties.DRM_useYn = pairs["DRM Setting"]["DRM useYn"].ToString2().Trim().Equals("Y");
                    IniProperties.DRM_Path = pairs["DRM Setting"]["DRM Path"].ToString2().Trim();
                    IniProperties.DRM_Result = pairs["DRM Setting"]["DRM Result"].ToString2().Trim();
                    IniProperties.DRM_Args = pairs["DRM Setting"]["DRM Args"].ToString2().Trim();
                }
                else
                {
                    Setting.createSetting();
                    LogMgr.Write("The configuration file is created in " + Environment.CurrentDirectory + @"\DocConvert_Server.ini", ConsoleColor.Yellow, LOG_LEVEL.INFO);
                }
            } catch (NullReferenceException e1)
            {
                LogMgr.Write("ERROR CODE: " + define.define.PARSING_INI_ERROR.ToString(), ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write(e1.Message, ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write(e1.StackTrace, ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write("There is a problem with the config file and continues with default settings", ConsoleColor.Red, LOG_LEVEL.ERROR);
            } catch (Exception e1)
            {
                LogMgr.Write("ERROR CODE: " + define.define.UNDEFINE_ERROR.ToString(), ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write(e1.Message, ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write(e1.StackTrace, ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write("There is a problem with the config file and continues with default settings", ConsoleColor.Red, LOG_LEVEL.ERROR);
            }

            // 파싱한 설정 출력
            LogMgr.Write("- Bind IP: " + IniProperties.Bind_IP, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- HTTP Port: " + IniProperties.HTTP_Port, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- Client KEY: " + IniProperties.Client_KEY, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- Workspace Directory: " + IniProperties.Workspace_Directory, ConsoleColor.White, LOG_LEVEL.INFO);
            
            LogMgr.Write("- SchedulerTime: " + IniProperties.SchedulerTime, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- CleanWorkspaceSchedulerYn: " + IniProperties.CleanWorkspaceSchedulerYn, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- CleanWorkspaceDay: " + IniProperties.CleanWorkspaceDay, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- CleanLogSchedulerYn: " + IniProperties.CleanLogSchedulerYn, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- CleanLogDay: " + IniProperties.CleanLogDay, ConsoleColor.White, LOG_LEVEL.INFO);
                          
            LogMgr.Write("- DRM useYn: " + IniProperties.DRM_useYn, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- DRM Path: " + IniProperties.DRM_Path, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- DRM Result: " + IniProperties.DRM_Result, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- DRM Args: " + IniProperties.DRM_Args, ConsoleColor.White, LOG_LEVEL.INFO);

            LogMgr.Write("------------------------------------------", ConsoleColor.White, LOG_LEVEL.INFO);
            #endregion

            LogMgr.Write("DocConvert Server Started", ConsoleColor.White, LOG_LEVEL.WARN);

            //콘솔 안꺼지게
            Console.ReadKey();
        }
    }
}
