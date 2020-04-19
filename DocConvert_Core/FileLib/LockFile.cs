using System;

namespace DocConvert_Core.FileLib
{
    class LockFile
    {
        /// <summary>
        /// File의 잠금을 해제합니다.
        /// </summary>
        /// <param name="FilePath"></param>
        public static void UnLock_File(string FilePath)
        {
            string adminUserName = Environment.UserName;
            System.Security.AccessControl.FileSecurity ds = System.IO.File.GetAccessControl(FilePath);
            System.Security.AccessControl.FileSystemAccessRule fsa = new System.Security.AccessControl.FileSystemAccessRule(adminUserName,
                System.Security.AccessControl.FileSystemRights.FullControl, System.Security.AccessControl.AccessControlType.Deny);

            ds.RemoveAccessRule(fsa);
            System.IO.File.SetAccessControl(FilePath, ds);
        }

        /// <summary>
        /// File를 잠급니다.
        /// </summary>
        /// <param name="FilePath"></param>
        public static void Lock_File(string FilePath)
        {
            string adminUserName = Environment.UserName;// getting your adminUserName
            System.Security.AccessControl.FileSecurity ds = System.IO.File.GetAccessControl(FilePath);
            System.Security.AccessControl.FileSystemAccessRule fsa = new System.Security.AccessControl.FileSystemAccessRule(adminUserName,
                System.Security.AccessControl.FileSystemRights.FullControl, System.Security.AccessControl.AccessControlType.Deny);

            ds.AddAccessRule(fsa);
            System.IO.File.SetAccessControl(FilePath, ds);
        }
    }
}
