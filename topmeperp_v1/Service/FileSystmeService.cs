using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ICSharpCode.SharpZipLib.Zip;
using log4net;
using System.IO;
using System.IO.Compression;
using topmeperp.Models;
using System.Collections;

namespace topmeperp.Service
{
    public class ZipFileCreator
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //讀取目錄下所有檔案
        private static ArrayList GetFiles(string path)
        {
            ArrayList files = new ArrayList();

            if (Directory.Exists(path))
            {
                files.AddRange(Directory.GetFiles(path));
            }

            return files;
        }
        //建立目錄

        public static void CreateDirectory(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
        public static void ClearDirectory(string path)
        {
            Directory.Delete(path, true);
        }
        public void ZipFiles(string path, string password, string comment)
        {
            ZipOutputStream zos = null;
            try
            {
                string zipPath = path + @"\" + Path.GetFileName(path) + ".zip";
                ArrayList files = GetFiles(path);
                zos = new ZipOutputStream(File.Create(zipPath));
                if (password != null && password != string.Empty) zos.Password = password;
                if (comment != null && comment != "") zos.SetComment(comment);
                zos.SetLevel(9);//Compression level 0-9 (9 is highest)
                byte[] buffer = new byte[4096];

                foreach (string f in files)
                {
                    ZipEntry entry = new ZipEntry(Path.GetFileName(f));
                    entry.DateTime = DateTime.Now;
                    zos.PutNextEntry(entry);
                    FileStream fs = File.OpenRead(f);
                    int sourceBytes;

                    do
                    {
                        sourceBytes = fs.Read(buffer, 0, buffer.Length);
                        zos.Write(buffer, 0, sourceBytes);
                    } while (sourceBytes > 0);

                    fs.Close();
                    fs.Dispose();
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            finally
            {
                zos.Finish();
                zos.Close();
                zos.Dispose();
            }
        }
    }
}