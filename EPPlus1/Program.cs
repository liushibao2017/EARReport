using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus1
{
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo newFile = new FileInfo(@"test.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(@"test.xlsx");
            }
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("test");
                worksheet.Cells.Style.WrapText = true;//自动换行属性
                worksheet.View.ShowGridLines = false;

                worksheet.Cells[1,1].Value = "名称";
                worksheet.Cells[1,2].Value = "价格";
                worksheet.Cells[1,3].Value = "销量";

                worksheet.Cells[2, 1].Value = "大米";
                worksheet.Cells[2, 2].Value = 56;
                worksheet.Cells[2, 3].Value = 100;

                worksheet.Cells[3, 1].Value = "玉米";
                worksheet.Cells[3, 2].Value = 45;
                worksheet.Cells[3, 3].Value = 150;

                worksheet.Cells[4, 1].Value = "小米";
                worksheet.Cells[4, 2].Value = 38;
                worksheet.Cells[4, 3].Value = 130;

                worksheet.Cells[5, 1].Value = "糯米";
                worksheet.Cells[5, 2].Value = 22;
                worksheet.Cells[5, 3].Value = 200;

                using (ExcelRange range = worksheet.Cells[1, 1, 5, 3])
                {
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }
                using (ExcelRange range = worksheet.Cells[1, 1, 1, 3])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Font.Color.SetColor(Color.White);
                    range.Style.Font.Name="微软雅黑";
                    range.Style.Font.Size = 12;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.Orange);
                }
                worksheet.Cells[1, 1, 5, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin,Color.Blue);

                ExcelChart chart = worksheet.Drawings.AddChart("chart",eChartType.Line);
                ExcelChartSerie serie = chart.Series.Add(worksheet.Cells[2,3,5,3],worksheet.Cells[2,1,5,1]);
                serie.HeaderAddress = worksheet.Cells[1, 3];


                chart.SetPosition(200,50);
                chart.SetSize(500,300);
                chart.Title.Text = "销量走势";
                chart.Title.Font.Color = Color.Black;
                chart.Title.Font.Size = 15;
                chart.Title.Font.Bold = true;
                chart.Style = eChartStyle.Style30;
                chart.Legend.Border.LineStyle = eLineStyle.SystemDot;
                chart.Legend.Border.Fill.Color = Color.Blue;


                package.Save();
            }

        }
    }
}
