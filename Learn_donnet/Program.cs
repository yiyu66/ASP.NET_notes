using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Learn_donnet
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = "E:\\Excel2003.xls";//定义要创建表格的位置及名称；
            //CreatExcel(fileName);
            ReadExcel(fileName);
        }
        
        public static void CreatExcel(string fileName)
        {
            Console.WriteLine("新建Excel文件测试");
            Console.ReadKey();
            Console.WriteLine("创建中...");
            HSSFWorkbook workbook2003 = new HSSFWorkbook(); //新建xls工作簿
            workbook2003.CreateSheet("Sheet1");  //新建3个Sheet工作表
            workbook2003.CreateSheet("Sheet2");
            workbook2003.CreateSheet("Sheet3");
            //FileStream file2003 = new FileStream(@"D:\华理文件\C#\Learn_donnet\Learn_donnet\Excel2003.xls", FileMode.Create);
            //workbook2003.Write(file2003);
            //file2003.Close();  //关闭文件流
            //workbook2003.Close();

            //XSSFWorkbook workbook2007 = new XSSFWorkbook();  //新建xlsx工作簿
            //workbook2007.CreateSheet("Sheet1");
            //workbook2007.CreateSheet("Sheet2");
            //workbook2007.CreateSheet("Sheet3");
            //FileStream file2007 = new FileStream(@"D:\华理文件\C#\Learn_donnet\Learn_donnet\Excel2007.xlsx", FileMode.Create);
            //workbook2007.Write(file2007);
            //file2007.Close();
            //workbook2007.Close();
            HSSFSheet SheetOne = (HSSFSheet)workbook2003.GetSheet("Sheet1"); //获取名称为Sheet1的工作表
            for (int i = 0; i < 10; i++)
            {
                SheetOne.CreateRow(i);   //创建10行
            }

            //对每一行创建10个单元格
            HSSFRow SheetRow = (HSSFRow)SheetOne.GetRow(0);  //获取Sheet1工作表的首行
            HSSFCell[] SheetCell = new HSSFCell[10];
            for (int i = 0; i < 10; i++)
            {
                SheetCell[i] = (HSSFCell)SheetRow.CreateCell(i);  //为第一行创建10个单元格
            }
            //创建之后就可以赋值了
            SheetCell[0].SetCellValue(true); //赋值为bool型         
            SheetCell[1].SetCellValue(0.000001); //赋值为浮点型
            SheetCell[2].SetCellValue("Excel2003"); //赋值为字符串
            SheetCell[3].SetCellValue("123456789987654321");//赋值为长字符串
            for (int i = 4; i < 10; i++)
            {
                SheetCell[i].SetCellValue(i);    //循环赋值为整形
            }
            FileStream file2003 = new FileStream(@fileName, FileMode.Create);
            //FileStream file2003 = new FileStream(@"D:\华理文件\C#\Learn_donnet\Learn_donnet\Excel2003.xls", FileMode.Create);
            workbook2003.Write(file2003);
            file2003.Close();  //关闭文件流
            workbook2003.Close();

            Console.WriteLine("创建成功");
            Console.ReadKey();
        }

        public static void ReadExcel(string fileName)
        {

            IWorkbook workbook = null;  //新建IWorkbook对象
            FileStream fileStream = new FileStream(@fileName, FileMode.Open, FileAccess.Read);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
            {
                workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook
            }
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
            {
                workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook
            }
            ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表
            IRow row;// = sheet.GetRow(0);            //新建当前工作表行数据
            for (int i = 0; i < sheet.LastRowNum; i++)  //对工作表每一行
            {
                row = sheet.GetRow(i);   //row读入第i行数据
                if (row != null)
                {
                    for (int j = 0; j < row.LastCellNum; j++)  //对工作表每一列
                    {
                        string cellValue = row.GetCell(j).ToString(); //获取i行j列数据
                        Console.WriteLine(cellValue);
                    }
                }
            }
            Console.ReadLine();
            fileStream.Close();
            workbook.Close();

        }
    }
}

