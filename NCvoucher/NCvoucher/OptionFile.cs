using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NCvoucher
{
    class OptionFile
    {
        /// <summary>
        /// 根据路径删除文件
        /// </summary>
        /// <param name="path"></param>
        public static void DeleteFile(string path)
        {
            if (File.Exists(path))
            {
                FileAttributes attr = File.GetAttributes(path);
                if (attr == FileAttributes.Directory)
                {
                    Directory.Delete(path, true);
                }
                else
                {
                    File.Delete(path);
                }
            }
        }
        /// <summary>
        /// 根据路径创建文件
        /// </summary>
        /// <param name="path"></param>
        public static void CreateFile(string path)
        {
            if (!File.Exists(path))
            {
                FileStream stream = File.Create(path);
                stream.Close();
                stream.Dispose();
            }
        }
    }
}
