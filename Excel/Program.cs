using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo fileInfo = new FileInfo(@"sample.xlsx");
            if (fileInfo.Exists)
            {
                fileInfo.Delete();
                fileInfo = new FileInfo(@"sample.xlsx");
            }
            using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet sht = excelPackage.Workbook.Worksheets.Add("test");
                sht.Cells[1, 1].Value = "第一章";
                sht.Cells[7, 1].Value = "第二章";
                sht.Cells[2, 1].Value = 1;
                sht.Cells[5, 1].Value = 2;
                sht.Cells[6, 1].Value = 3;
                sht.Cells[8, 1].Value = 1;
                sht.Cells[10, 1].Value = 2;
                sht.Cells[11, 1].Value = 3;
                sht.Cells[3, 2].Value = 1.1;
                sht.Cells[4, 2].Value = 1.2;
                sht.Cells[9, 2].Value = 1.1;
                sht.Cells[12, 2].Value = 3.1;
                sht.Cells[13,2].Value = 3.2;
                sht.Cells[14, 2].Value = 3.3;
                sht.Row(1).OutlineLevel = 1;
                sht.Row(7).OutlineLevel = 1;
                sht.Row(2).OutlineLevel = 2;
                sht.Row(5).OutlineLevel = 2;
                sht.Row(6).OutlineLevel = 2;
                sht.Row(8).OutlineLevel = 2;
                sht.Row(10).OutlineLevel = 2;
                sht.Row(11).OutlineLevel = 2;
                sht.Row(3).OutlineLevel = 3;
                sht.Row(4).OutlineLevel = 3;
                sht.Row(9).OutlineLevel = 3;
                sht.Row(12).OutlineLevel = 3;
                sht.Row(13).OutlineLevel = 3;
                sht.Row(14).OutlineLevel = 3;
                excelPackage.Save();
            }

        }
    }
}
