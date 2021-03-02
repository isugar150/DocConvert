using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using DocConvert.Common;
using DocConvert_Core.IniLib;
using Microsoft.Win32;

namespace DocConvert
{
    public class Program
    {
        public static iniProperties IniProperties = new iniProperties(); // Ini 설정값 변수

        public static ManualResetEvent manualEvent = new ManualResetEvent(false);

        public static bool isHwpConverting = false;

        private static Timer _Scheduler;

        [DllImport("kernel32")]
        public static extern Int32 GetCurrentProcessId();

        static void Main(string[] args)
        {
            #region 프로세스 실행중이면 종료
            Process[] process = Process.GetProcesses();
            string pidPath = @".\DocConvert.pid";
            int myPID = 0;
            try
            {
                try
                {
                    int oldPID = int.Parse(File.ReadAllText(pidPath));
                    foreach (Process prs in process)
                    {
                        if (prs.Id == oldPID)
                        {
                            prs.Kill();
                            LogMgr.Write("I ended a program I had run before.", ConsoleColor.Yellow, LOG_LEVEL.WARN);
                            break;
                        }
                    }
                }
                catch (Exception) { }

                myPID = GetCurrentProcessId(); // 현재 프로그램 PID 가져오기.
                File.WriteAllText(pidPath, myPID.ToString());
            }
            catch (Exception) { }
            #endregion

            #region 제품명 및 기본 정보 출력
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

                ", ConsoleColor.Gray, LOG_LEVEL.FATAL, true);
            // 저작권 정보
            LogMgr.Write("         Copyright© 2021. Jm's Corp. All rights reserved.\r\n\r\n", ConsoleColor.Gray, LOG_LEVEL.FATAL, true);
            LogMgr.Write("Starting DocConvert Server..", ConsoleColor.White, LOG_LEVEL.WARN);
            LogMgr.Write("Program Directory: " + Environment.CurrentDirectory, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write(".Net Framework Version: " + Environment.Version.ToString(), ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("OS Version: " + Environment.OSVersion.ToString(), ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("PC(Domain) Name: " + Environment.UserDomainName, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("User Name: " + Environment.UserName, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("DocConvert Version: " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(), ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("DocConvert PID: " + myPID, ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("DocConvert LogLevel: " + LogMgr.getLogLevel("DocConvert_Log"), ConsoleColor.White, LOG_LEVEL.INFO);
            LogMgr.Write("DocConvert_Core LogLevel: " + LogMgr.getLogLevel("DocConvert_Core_Log"), ConsoleColor.White, LOG_LEVEL.INFO);
            #endregion

            #region parse Ini File
            LogMgr.Write("------------ Parsing Ini File ------------", ConsoleColor.White, LOG_LEVEL.INFO);
            try
            {
                for (int i = 0; i < 3; i++)
                {
                    IniFile pairs = new IniFile(); // Ini 변수 초기화
                    if (new FileInfo(Environment.CurrentDirectory + @".\DocConvert.ini").Exists)
                    {
                        pairs.Load(Environment.CurrentDirectory + @".\DocConvert.ini");
                        IniProperties.Bind_IP = pairs["Common"]["Bind IP"].ToString2().Trim();
                        IniProperties.Socket_Port = int.Parse(pairs["Common"]["Socket Port"].ToString2().Trim());
                        IniProperties.Client_KEY = pairs["Common"]["Client KEY"].ToString2().Trim();
                        IniProperties.Workspace_Directory = pairs["Common"]["Workspace Directory"].ToString2().Trim();
                        IniProperties.Product_Name = pairs["Common"]["Product Name"].ToString2().Trim();

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
                        LogMgr.Write("The configuration file is created in " + Environment.CurrentDirectory + @"\DocConvert.ini", ConsoleColor.Yellow, LOG_LEVEL.INFO);
                    }
                }
            }
            catch (NullReferenceException e1)
            {
                LogMgr.Write("ERROR CODE: " + define.PARSING_INI_ERROR.ToString(), ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write(e1.Message, ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write(e1.StackTrace, ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write("There is a problem with the config file and continues with default settings", ConsoleColor.Red, LOG_LEVEL.ERROR);
            }
            catch (Exception e1)
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
            LogMgr.Write("- Product Name: " + IniProperties.Product_Name, ConsoleColor.White, LOG_LEVEL.INFO);

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

            #region 프로그램 기본 설정
            Console.Title = IniProperties.Product_Name;
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

            #region init Socket Server
            Thread socket_Thread = new Thread(() =>
            {
                new Network.AsyncSocketServer(IniProperties.Bind_IP, IniProperties.Socket_Port);
            });
            socket_Thread.Start();
            LogMgr.Write("Socket Server Listen On " + IniProperties.Bind_IP + ":" + IniProperties.Socket_Port, ConsoleColor.Green, LOG_LEVEL.WARN);
            #endregion

            #region 파일정리 스케줄러 등록
            if (IniProperties.CleanWorkspaceSchedulerYn || IniProperties.CleanLogSchedulerYn)
            {
                int workDateMilliseconds = CalculateTimerInterval();

                _Scheduler = new Timer(Callback, null, workDateMilliseconds, Timeout.Infinite);
                LogMgr.Write("============ [Scheduler] ============", ConsoleColor.Yellow, LOG_LEVEL.INFO);
                LogMgr.Write("Next scheduler operation time: " + (DateTime.Now + TimeSpan.FromMilliseconds(workDateMilliseconds)).ToString("yyyy/MM/dd HH:mm"), ConsoleColor.Yellow, LOG_LEVEL.INFO);

                if (IniProperties.CleanWorkspaceSchedulerYn)
                {
                    LogMgr.Write("If the workspace last modified date is past " + (DateTime.Now.AddDays(-IniProperties.CleanWorkspaceDay) + TimeSpan.FromMilliseconds(workDateMilliseconds)).ToString("yyyy/MM/dd HH:mm"), ConsoleColor.Yellow, LOG_LEVEL.INFO);
                }

                if (IniProperties.CleanLogSchedulerYn)
                {
                    LogMgr.Write("If the log last modified date is past " + (DateTime.Now.AddDays(-IniProperties.CleanLogDay) + TimeSpan.FromMilliseconds(workDateMilliseconds)).ToString("yyyy/MM/dd HH:mm"), ConsoleColor.Yellow, LOG_LEVEL.INFO);
                }

                LogMgr.Write("=====================================", ConsoleColor.Yellow, LOG_LEVEL.INFO);

                LogMgr.Write("Cleanup scheduler is working", ConsoleColor.Green, LOG_LEVEL.INFO);
            }
            #endregion

            LogMgr.Write("DocConvert Server Started in " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), ConsoleColor.Green, LOG_LEVEL.WARN);

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

        #region 스케줄러 관련 메소드

        /// <summary>
        /// 스케줄러 등록시 지정할 타이머
        /// </summary>
        /// <returns>해당 시간까지 Milliseconds</returns>
        private static int CalculateTimerInterval()
        {
            string[] timeStr = IniProperties.SchedulerTime.Split(',');
            int[] time = new int[3];
            for (int i = 0; i < 2; i++)
                time[0] = int.Parse(timeStr[0]); // 시
            time[1] = int.Parse(timeStr[1]); // 분
            DateTime timeTaken = new DateTime();
            if (new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, time[0], time[1], 0) < DateTime.Now)
                timeTaken = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, time[0], time[1], 1).AddDays(1);
            else
                timeTaken = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, time[0], time[1], 1);

            TimeSpan curTime = timeTaken - DateTime.Now;

            return (int)curTime.TotalMilliseconds;
        }

        private static void Callback(Object state)
        {
            Thread.Sleep(3000);
            try
            {
                int workDateMilliseconds = CalculateTimerInterval();
                _Scheduler.Change(CalculateTimerInterval(), Timeout.Infinite);
                LogMgr.Write("============ [Scheduler] ============", ConsoleColor.Yellow, LOG_LEVEL.INFO);
                LogMgr.Write("Next scheduler operation time: " + (DateTime.Now + TimeSpan.FromMilliseconds(workDateMilliseconds)).ToString("yyyy/MM/dd HH:mm"), ConsoleColor.Yellow, LOG_LEVEL.INFO);
                if (IniProperties.CleanWorkspaceSchedulerYn)
                {
                    LogMgr.Write("If the workspace last modified date is past " + (DateTime.Now.AddDays(-IniProperties.CleanWorkspaceDay) + TimeSpan.FromMilliseconds(workDateMilliseconds)).ToString("yyyy/MM/dd HH:mm"), ConsoleColor.Yellow, LOG_LEVEL.INFO);
                }
                if (IniProperties.CleanLogSchedulerYn)
                {
                    LogMgr.Write("If the log last modified date is past " + (DateTime.Now.AddDays(-IniProperties.CleanLogDay) + TimeSpan.FromMilliseconds(workDateMilliseconds)).ToString("yyyy/MM/dd HH:mm"), ConsoleColor.Yellow, LOG_LEVEL.INFO);
                }
                LogMgr.Write("=====================================", ConsoleColor.Yellow, LOG_LEVEL.INFO);
            }
            catch (Exception) { _Scheduler.Dispose(); LogMgr.Write("An error occurred while the scheduler was running, so it was disabled.", LOG_LEVEL.ERROR); }
            if (IniProperties.CleanWorkspaceSchedulerYn)
            {
                LogMgr.Write("Workspace cleanup scheduler has been run", LOG_LEVEL.WARN);
                DirectoryInfo workspaceDir = new DirectoryInfo(IniProperties.Workspace_Directory);
                deleteFolder(workspaceDir.FullName, IniProperties.CleanWorkspaceDay);
                if (!workspaceDir.Exists)
                    workspaceDir.Create();
                DirectoryInfo dataDir = new DirectoryInfo(IniProperties.Workspace_Directory + @"\data");
                if (!dataDir.Exists)
                    dataDir.Create();
                DirectoryInfo tmpDir = new DirectoryInfo(IniProperties.Workspace_Directory + @"\tmp");
                if (!tmpDir.Exists)
                    tmpDir.Create();
            }
            DirectoryInfo logDir = new DirectoryInfo(IniProperties.Workspace_Directory);
            deleteFolder(logDir.FullName, IniProperties.CleanLogDay);
            if (!logDir.Exists)
                logDir.Create();
            if (IniProperties.CleanLogSchedulerYn)
            {
                LogMgr.Write("Logs cleanup scheduler has been run", LOG_LEVEL.WARN);
                deleteFolder(Environment.CurrentDirectory + @"\Log", IniProperties.CleanLogDay);
            }
        }


        /// <summary>
        /// 수정날짜가 설정한 일수가 지났으면 디렉토리와 안의 파일까지 삭제하는 로직
        /// </summary>
        /// <param name="strPath">대상 경로</param>
        /// <param name="DeletionCycle">지난 일수</param>
        public static void deleteFolder(string strPath, int DeletionCycle)
        {
            foreach (string Folder in Directory.GetDirectories(strPath))
            {
                deleteFolder(Folder, DeletionCycle); //재귀함수 호출
            }

            foreach (string file in Directory.GetFiles(strPath))
            {
                FileInfo fi = new FileInfo(file);
                if (fi.LastWriteTime <= DateTime.Now.AddDays(-DeletionCycle).AddSeconds(-4))  // -4초는 Sleep
                {
                    fi.Delete();
                    LogMgr.Write("File " + fi.FullName + " deleted");
                }
                if (!isFiles(fi.DirectoryName))
                {
                    DirectoryInfo di = new DirectoryInfo(fi.DirectoryName);
                    LogMgr.Write("Folder " + fi.DirectoryName + " deleted");
                    di.Delete(true);
                }
            }
        }

        /// <summary>
        /// 파일인지 디렉토리인지 확인하는 로직.
        /// </summary>
        /// <param name="dir"></param>
        /// <returns></returns>
        private static bool isFiles(string dir)
        {
            string[] Directories = Directory.GetDirectories(dir);   // Defalut Folder
            {
                string[] Files = Directory.GetFiles(dir);   // File list Search
                if (Files.Length != 0)
                {
                    return true;
                }

                foreach (string nodeDir in Directories)   // Folder list Search
                {
                    isFiles(nodeDir);   // reSearch
                }
            }
            return false;
        }

        #endregion
    }
}
