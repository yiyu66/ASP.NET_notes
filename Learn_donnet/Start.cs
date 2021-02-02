using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Learn_donnet
{
    class Start
    {
        static void Main()
        {
            string fileName = "E:\\Excel2003.xls";//定义要创建表格的位置及名称；
            NPOIExcle.CreatExcel(fileName);
            NPOIExcle.ReadExcel(fileName);
        }

    }
}

