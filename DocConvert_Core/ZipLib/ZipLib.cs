using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocConvert_Core.ZipLib
{
    public class ZipLib
    {
		public static bool CreateZipFile(string[] filenames, string outPath)
		{
			using (ZipOutputStream s = new ZipOutputStream(File.Create(outPath)))
			{
				s.SetLevel(9); // 압축 레벨
				byte[] buffer = new byte[4096];
				foreach (string file in filenames)
				{
					var entry = new ZipEntry(Path.GetFileName(file));
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

		public static bool UnZipFile(string sourceZip)
		{
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
    }
}
