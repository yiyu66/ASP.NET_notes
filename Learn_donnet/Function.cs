using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Learn_donnet
{
    class Function
    {
        /// <summary>
        /// 删除文件夹中的所有内容
        /// </summary>
        /// <param name="srcPath">文件夹地址</param>
        static public void DelectDir(string srcPath) //删除文件夹下的文件
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(srcPath);
                FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();  //返回目录中所有文件和子目录
                foreach (FileSystemInfo i in fileinfo)
                {
                    if (i is DirectoryInfo)            //判断是否文件夹
                    {
                        DirectoryInfo subdir = new DirectoryInfo(i.FullName);
                        subdir.Delete(true);          //删除子目录和文件
                    }
                    else
                    {
                        File.Delete(i.FullName);      //删除指定文件
                    }
                }
            }
            catch (Exception e)
            {
                throw;
            }
        }
    }
}
