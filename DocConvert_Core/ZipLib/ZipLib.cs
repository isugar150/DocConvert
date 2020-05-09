using DocConvert_Core.FileLib;
using ICSharpCode.SharpZipLib.Zip;
using NLog;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace DocConvert_Core.ZipLib
{
    public class ZipLib
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Engine_Log");
        public static bool CreateZipFile(string[] filenames, string outPath)
        {
            try {
            foreach (string str in filenames)
            {
                LockFile.UnLock_File(str);
            }
                using (ZipOutputStream s = new ZipOutputStream(File.Create(outPath)))
                {
                    s.SetLevel(9); // 압축 레벨
                    byte[] buffer = new byte[4096];
                    foreach (string file in filenames)
                    {
                        ZipEntry entry = new ZipEntry(Path.GetFileName(file));
                        entry.DateTime = DateTime.Now;
                        s.PutNextEntry(entry);
                        using (FileStream fs = File.OpenRead(file))
                        {
                            int sourceBytes;
                            do
                            {
                                sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                s.Write(buffer, 0, sourceBytes);
                            } while (sourceBytes > 0);
                        }
                    }
                    s.Finish();
                    s.Close();

                    return true;
                } 
            }
            catch(Exception e1)
            {

                logger.Error("======= Method: " + MethodBase.GetCurrentMethod().Name + " =======");
                logger.Error(new StackTrace(e1, true).ToString());
                logger.Error("변환 실패: " + e1.Message);
                logger.Error("================ End ================");
                return false;
            }
        }

        public static bool UnZipFile(string sourceZip)
        {
            try {
                using (ZipInputStream s = new ZipInputStream(File.OpenRead(sourceZip)))
                {
                    ZipEntry theEntry;
                    while ((theEntry = s.GetNextEntry()) != null)
                    {

                        Console.WriteLine(theEntry.Name);

                        string directoryName = Path.GetDirectoryName(theEntry.Name);
                        string fileName = Path.GetFileName(theEntry.Name);

                        if (directoryName.Length > 0)
                        {
                            Directory.CreateDirectory(directoryName);
                        }

                        if (fileName != String.Empty)
                        {
                            using (FileStream streamWriter = File.Create(theEntry.Name))
                            {

                                int size = 2048;
                                byte[] data = new byte[2048];
                                while (true)
                                {
                                    size = s.Read(data, 0, data.Length);
                                    if (size > 0)
                                    {
                                        streamWriter.Write(data, 0, size);
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception e1)
            {

                logger.Error("======= Method: " + MethodBase.GetCurrentMethod().Name + " =======");
                logger.Error(new StackTrace(e1, true).ToString());
                logger.Error("변환 실패: " + e1.Message);
                logger.Error("================ End ================");
                return false;
            }
        }
    }
}
