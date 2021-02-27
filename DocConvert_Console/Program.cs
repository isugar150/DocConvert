using System;
using System.IO;
using System.Threading;
using DocConvert_Console.Common;
using DocConvert_Core.IniLib;
using Microsoft.Win32;
using NLog;

//TODO 파일 관리 스케줄러 만들어야됨.
//TODO DRM 사용관련 업데이트해야함.


namespace DocConvert_Console
{
    public class Program
    {
        public static iniProperties IniProperties = new iniProperties(); // Ini 설정값 변수

        public static ManualResetEvent manualEvent = new ManualResetEvent(false);

        public static bool isHwpConverting = false;
        
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

                ", ConsoleColor.Green, LOG_LEVEL.FATAL, true);
            // 저작권 정보
            LogMgr.Write("         Copyright© 2021. Jm's Corp. All rights reserved.\r\n\r\n", ConsoleColor.White, LOG_LEVEL.FATAL, true);
            LogMgr.Write("Starting DocConvert Server..", ConsoleColor.White, LOG_LEVEL.WARN);
            LogMgr.Write("Program Directory: " + Environment.CurrentDirectory, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write(".Net Framework Version: " + Environment.Version.ToString(), ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("OS Version: " + Environment.OSVersion.ToString(), ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("PC(Domain) Name: " + Environment.UserDomainName, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("User Name: " + Environment.UserName, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("DocConvert_Core LogLevel: " + LogMgr.getLogLevel("DocConvert_Core_Log"), ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("DocConvert_Console LogLevel: " + LogMgr.getLogLevel("DocConvert_Console_Log"), ConsoleColor.White, LOG_LEVEL.INFO);
            #region parse Ini File
            LogMgr.Write("------------ Parsing Ini File ------------", ConsoleColor.White, LOG_LEVEL.INFO);
            try
            {
                IniFile pairs = new IniFile(); // Ini 변수 초기화
                if (new FileInfo(Environment.CurrentDirectory + "/DocConvert_Console.ini").Exists)
                {
                    pairs.Load(Environment.CurrentDirectory + "/DocConvert_Console.ini");
                    IniProperties.Bind_IP = pairs["Common"]["Bind IP"].ToString2().Trim();
                    IniProperties.Socket_Port = int.Parse(pairs["Common"]["Socket Port"].ToString2().Trim());
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
                LogMgr.Write("ERROR CODE: " + define.PARSING_INI_ERROR.ToString(), ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write(e1.Message, ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write(e1.StackTrace, ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write("There is a problem with the config file and continues with default settings", ConsoleColor.Red, LOG_LEVEL.ERROR);
            } catch (Exception e1)
            {
                LogMgr.Write("ERROR CODE: " + define.UNDEFINE_ERROR.ToString(), ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write(e1.Message, ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write(e1.StackTrace, ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write("There is a problem with the config file and continues with default settings", ConsoleColor.Red, LOG_LEVEL.ERROR);
            }

            // 파싱한 설정 출력
            LogMgr.Write("- Bind IP: " + IniProperties.Bind_IP, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- Socket Port: " + IniProperties.Socket_Port, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- Client KEY: " + IniProperties.Client_KEY, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- Workspace Directory: " + IniProperties.Workspace_Directory, ConsoleColor.White, LOG_LEVEL.INFO);
            
            LogMgr.Write("- SchedulerTime: " + IniProperties.SchedulerTime, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- CleanWorkspaceSchedulerYn: " + IniProperties.CleanWorkspaceSchedulerYn, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- CleanWorkspaceDay: " + IniProperties.CleanWorkspaceDay, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- CleanLogSchedulerYn: " + IniProperties.CleanLogSchedulerYn, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("- CleanLogDay: " + IniProperties.CleanLogDay, ConsoleColor.White, LOG_LEVEL.INFO);
                          
            LogMgr.Write("- DRM useYn: " + IniProperties.DRM_useYn, ConsoleColor.White, LOG_LEVEL.INFO);
            if (IniProperties.DRM_useYn)
            {
                LogMgr.Write("- DRM Path: " + IniProperties.DRM_Path, ConsoleColor.White, LOG_LEVEL.INFO);
                LogMgr.Write("- DRM Result: " + IniProperties.DRM_Result, ConsoleColor.White, LOG_LEVEL.INFO);
                LogMgr.Write("- DRM Args: " + IniProperties.DRM_Args, ConsoleColor.White, LOG_LEVEL.INFO);
            }

            LogMgr.Write("------------------------------------------", ConsoleColor.White, LOG_LEVEL.INFO);
            #endregion

            #region Workspace 초기화
            DirectoryInfo tmpDir = new DirectoryInfo(IniProperties.Workspace_Directory + @"\tmp");
            DirectoryInfo dataDir = new DirectoryInfo(IniProperties.Workspace_Directory + @"\data");
            if (!tmpDir.Exists)
            {
                tmpDir.Create();
                LogMgr.Write("Created tmp Directory " + tmpDir.FullName, ConsoleColor.Yellow, LOG_LEVEL.INFO);
            }
            if (!dataDir.Exists)
            {
                dataDir.Create();
                LogMgr.Write("Created data Directory " + dataDir.FullName, ConsoleColor.Yellow, LOG_LEVEL.INFO);
            }
            #endregion

            #region init Socket Server
            Thread socket_Thread = new Thread(() =>
            {
                new Network.AsyncSocketServer(IniProperties.Bind_IP, IniProperties.Socket_Port);
            });
            socket_Thread.Start();
            LogMgr.Write("Socket Server Listen On IP: " + IniProperties.Bind_IP + " PORT: " + IniProperties.Socket_Port, ConsoleColor.Green, LOG_LEVEL.WARN);
            #endregion

            #region 한글 DLL 레지스트리 등록
            if (File.Exists(Environment.CurrentDirectory + @"\FilePathCheckerModuleExample.dll"))
            {
                RegistryKey regKey = Registry.CurrentUser.CreateSubKey(@"Software\HNC\HwpCtrl\Modules", RegistryKeyPermissionCheck.ReadWriteSubTree);
                regKey.SetValue("FilePathCheckerModuleExample", Environment.CurrentDirectory + @"\FilePathCheckerModuleExample.dll", RegistryValueKind.String);
                LogMgr.Write("Hangul DLL has been registered in the registry.", LOG_LEVEL.INFO);
            }
            else
            {
                LogMgr.Write("Hangul DLL does not exist in the executed path.", LOG_LEVEL.ERROR);
            }
            #endregion

            LogMgr.Write("DocConvert Server Started in " + DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss"), ConsoleColor.Green, LOG_LEVEL.WARN);


            Console.WriteLine("Press the exit key to exit.");
            while (true)
            {
                string k = Console.ReadLine();
                if ("exit".Equals(k, StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }
            }
        }
    }
}
