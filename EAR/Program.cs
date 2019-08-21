using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System.Data;
using System.IO;
using OfficeOpenXml.Style;
using System.Drawing;
using OfficeOpenXml.Drawing;

namespace EAR
{
    class EPHelper
    {
        static void Main(string[] args)
        {
            #region Functions
            //EPplus操作Excel文件的属性和方法
            EPHelper.Helper();                                       //取消注释本行，即可执行方法

            //导出附有数据图的Excel文件
            //EPHelper.GenerateExcelReportColumnClustered();           //取消注释本行，即可执行方法

            //DataTable导出为Excel文件
            //DataTable dtTest = new DataTable();
            //dtTest.Columns.Add("Name", typeof(string));
            //dtTest.Columns.Add("Century", typeof(int));
            //dtTest.Columns.Add("Money", typeof(float));
            //dtTest.Columns.Add("Time", typeof(DateTime));
            //dtTest.Columns.Add("Test", typeof(string));

            //for (int i = 0; i < 400; i++)
            //{
            //    DataRow dr = dtTest.NewRow();
            //    dr[0] = "Name" + i;
            //    dr[1] = i;
            //    dr[2] = (float)i;
            //    dr[3] = DateTime.Now.ToString();
            //    dr[4] = "HelloWorld";
            //    dtTest.Rows.Add(dr);
            //}

            //string outputPath = "D:\\ExcelTest\\" + DateTime.Now.ToString("yyyy_MM_dd_HHmmss") + ".xlsx";
            //string sheetName = "Test1";
           // EPHelper.DataTableToExcel(dtTest, outputPath, sheetName, true);     //取消注释本行和上方DataTable定义，即可执行方法

            //读取Excel文件到DataTable
            //EPHelper.ExcelToDataTable(InputPath, sheetName, firstRowAsColumnName);   //取消注释本行，即可执行方法
            #endregion
        }

        #region EPplusHelper 
        private static void Helper()
        {
            using (ExcelPackage package = new ExcelPackage(
                new FileInfo("D:\\ExcelTest\\" + DateTime.Now.ToString("yyyy_MM_dd_HHmmss") + ".xlsx")))
            {
                //创建一个名为Name的sheet对象
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Name");

                #region 样本数据
                //提示：EPplus中操作Excel文件时，表格的行和列都是从索引1开始的
                worksheet.Cells[1, 1].Value = "名称";
                worksheet.Cells[1, 2].Value = "价格";
                worksheet.Cells[1, 3].Value = "销量";

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
                #endregion

                #region 公式计算:方法一
                worksheet.Cells[1, 4].Value = "总价";
                //求积：如下表示>>从D2到D5单元格的值，是以B2*C2开始并依次到B5*C5的乘积结果,结果是数字类型
                //若要相乘的两列的值是字符串类型的数字，也可以进行相乘运算
                worksheet.Cells["D2:D5"].Formula = "B2*C2";
                //求和
                worksheet.Cells[1, 5].Value = "总量";
                worksheet.Cells["E2:E5"].Formula = "C2+10";
                #endregion

                #region 公式计算:方法二
                worksheet.Cells[6, 1].Value = "总计";
                //SUBTOTAL(1,{0}):格式1包含隐藏，计算单元格值的平均值             101格式忽略隐藏
                //SUBTOTAL(2,{0}):格式2包含隐藏，计算单元格数值的项数             102格式忽略隐藏
                //SUBTOTAL(3,{0}):格式3包含隐藏，计算单元格非空数值的项数         103格式忽略隐藏
                //SUBTOTAL(4,{0}):格式4包含隐藏，获取单元格的值的最大值           104格式忽略隐藏
                //SUBTOTAL(5,{0}):格式5包含隐藏，获取单元格的值的最小值           105格式忽略隐藏
                //SUBTOTAL(6,{0}):格式6包含隐藏，计算单元格的乘积                 106格式忽略隐藏
                //SUBTOTAL(7,{0}):格式7包含隐藏，计算单元格的值的标准偏差         107格式忽略隐藏
                //SUBTOTAL(8,{0}):格式8包含隐藏，计算整个样本总体的标准偏差       108格式忽略隐藏
                //SUBTOTAL(9,{0}):格式9包含隐藏，计算单元格值的和                 109格式忽略隐藏
                //SUBTOTAL(10,{0}):格式10包含隐藏，计算单元格的值的方差           110格式忽略隐藏
                //SUBTOTAL(11,{0}):格式11包含隐藏，计算整个样本总体的方差         111格式忽略隐藏
                worksheet.Cells[6, 2, 6, 5].Formula = string.Format("SUBTOTAL(9,{0})", new ExcelAddress(2, 2, 5, 2).Address);
                #endregion

                #region 单元格文本格式
                worksheet.Cells[5, 3].Style.Numberformat.Format = "#,##0.00";   //设置数值类型的内容保留两位小数,"#,##0.000"保留三位小数
                worksheet.Cells[5, 4].Style.Numberformat.Format = " @";         //限制该单元格的内容为文本格式
                #endregion

                #region 单元格对齐方式
                worksheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                worksheet.Cells[1, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;//不指定单元格，表示所有单元格内容向下对齐
                #endregion

                #region 合并单元格
                //合并单元格，如下表示:从Excel表格的第一行的第四列向右合并第一行的第五列的值，合并后的单元格位置是第一行的第四列
                //若在合并单元格之前，单元格已经填充数值，则合并后只保留第一个，本例中即，第一行的第四列的值
                worksheet.Cells[1, 4, 1, 5].Merge = true;

                //若是对先进行合并的单元格进行赋值，单元格也只会保留第一个值
                //我这里对第一第二的理解是，从左向右，合并的过程就像是左边的单元格向右吃掉单元格，所以值留下的始终的左边的
                //而被吃掉的单元格，就算赋值(无论先后)都没有任何意义了
                worksheet.Cells[7, 1, 7, 2].Merge = true;
                worksheet.Cells[7, 1].Value = "合并后7_1";
                worksheet.Cells[7, 2].Value = "合并后7_2";
                worksheet.Cells[7, 3].Value = "合并后7_3";

                //worksheet.Cells[8, 2, 8, 1].Merge = true;    这行代码想要从右向左合并，明显是造反，因此ERROR了 = =||
                #endregion

                #region 单元格字体样式
                worksheet.Cells[1, 1].Style.Font.Bold = true;//字体为粗体
                worksheet.Cells[1, 2].Style.Font.Color.SetColor(Color.Red);//字体颜色
                worksheet.Cells[1, 3].Style.Font.Name = "微软雅黑";//字体
                worksheet.Cells[1, 4].Style.Font.Size = 12;//字体大小(这个字体的大小就是Excel表格里面可以设置的字体大小)
                worksheet.Cells.Style.ShrinkToFit = true;//单元格内容自适应单元格的大小
                #endregion

                #region 单元格背景样式
                worksheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;//这一行代码是设置单元格背景颜色必须的
                worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(128, 128, 128));//设置单元格背景色
                #endregion

                #region 单元格边框
                //设置单元格所有边框
                worksheet.Cells[1, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));

                //单独设置单元格底部边框样式和颜色（上下左右均可分开设置）
                worksheet.Cells[2, 3].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;//这一行代码是设置单元格颜色必须的
                worksheet.Cells[2, 3].Style.Border.Bottom.Color.SetColor(Color.FromArgb(191, 191, 191));

                //单元格自适应
                //worksheet.Cells.Style.WrapText = true;
                #endregion

                #region 单元格行高和列宽
                //worksheet.Row(1).Height = 24;//设置行高
                worksheet.Row(1).CustomHeight = true;//自动调整行高
                worksheet.Column(3).Width = 24;//设置列宽
                #endregion

                #region 设置sheet页
                //worksheet.View.ShowGridLines = false;//去掉sheet的网格线

                //看到这里相信看客已经发现了，当对Excel进行颜色的操作时，都会需要一行设置颜色的代码
                //当设置sheet的背景颜色时，背景图片就会被覆盖掉
                //worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                //worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.LightGray);//设置背景色

                //设置sheet背景图片
                //worksheet.BackgroundImage.Image = Image.FromFile(@"E:\Learn_contents\其他\壁纸\f6d5450e0f29c3e8ef5ea81db18efd43.jpg");
                #endregion

                #region 插入图片
                //注意：这个“插入”的意思是在Excel的sheet上浮动添加图片，并不是在一个单元格里面添加图片

                //插入图片
                ExcelPicture picture = worksheet.Drawings.AddPicture("logo", Image.FromFile(@"E:\Learn_contents\其他\壁纸\f6d5450e0f29c3e8ef5ea81db18efd43.jpg"));
                picture.SetPosition(100, 100);//设置图片的位置
                picture.SetSize(100, 100);//设置图片的大小
                #endregion

                #region 插入形状

                //插入形状
                //eShapeStyle.Rect        艺术字
                ExcelShape shape = worksheet.Drawings.AddShape("shape", eShapeStyle.Rect);

                shape.Font.Color = Color.Red;//设置形状的字体颜色
                shape.Font.Size = 15;//字体大小
                shape.Font.Bold = true;//字体粗细
                shape.Fill.Style = eFillStyle.NoFill;//设置形状的填充样式
                shape.Border.Fill.Style = eFillStyle.NoFill;//边框样式
                shape.SetPosition(200, 300);//形状的位置
                shape.SetSize(80, 30);//形状的大小
                shape.Text = "test";//形状的内容
                #endregion

                #region 给图片插入超链接
                //这里使用的图片和上面的插入图片是一样的，因为FromFile()方法，所以会报错，实验的时候把上面的“插入图片”的代码注释就好
                //ExcelPicture picture2 = worksheet.Drawings.AddPicture("logo", Image.FromFile(@"E:\Learn_contents\其他\壁纸\f6d5450e0f29c3e8ef5ea81db18efd43.jpg"), new ExcelHyperLink("http:\\www.baidu.com", UriKind.Relative));
                #endregion

                #region 隐藏sheet、列
                worksheet.Hidden = eWorkSheetHidden.Hidden;//隐藏sheet

                //这里是sheet的对象worksheet调用的方法，所以隐藏的是当前sheet
                worksheet.Column(2).Hidden = true;//隐藏指定列
                worksheet.Row(2).Hidden = true;//隐藏指定行
                #endregion

                #region Excel文件加密
                worksheet.Protection.IsProtected = true;//设置是否进行锁定
                worksheet.Protection.SetPassword("admin");//设置密码
                worksheet.Protection.AllowAutoFilter = false;//下面是一些锁定时权限的设置
                worksheet.Protection.AllowDeleteColumns = false;
                worksheet.Protection.AllowDeleteRows = false;
                worksheet.Protection.AllowEditScenarios = false;
                worksheet.Protection.AllowEditObject = false;
                worksheet.Protection.AllowFormatCells = false;
                worksheet.Protection.AllowFormatColumns = false;
                worksheet.Protection.AllowFormatRows = false;
                worksheet.Protection.AllowInsertColumns = false;
                worksheet.Protection.AllowInsertHyperlinks = false;
                worksheet.Protection.AllowInsertRows = false;
                worksheet.Protection.AllowPivotTables = false;
                worksheet.Protection.AllowSelectLockedCells = false;
                worksheet.Protection.AllowSelectUnlockedCells = false;
                worksheet.Protection.AllowSort = false;
                #endregion

                #region Excel文件属性设置
                //windows操作系统中，在Excel文件属性的详细信息中可以查看编辑的各个属性信息
                package.Workbook.Properties.Category = "类别";
                package.Workbook.Properties.Author = "作者";
                package.Workbook.Properties.Comments = "备注";
                package.Workbook.Properties.Company = "公司";
                package.Workbook.Properties.Keywords = "关键字";
                package.Workbook.Properties.Manager = "管理者";
                package.Workbook.Properties.Status = "内容状态";
                package.Workbook.Properties.Subject = "主题";
                package.Workbook.Properties.Title = "标题";
                package.Workbook.Properties.LastModifiedBy = "最后一次保存者";
                #endregion

                #region 设置下拉框
                //设置下拉框时，首先设置下拉框显示的数据区域并为其命名
                var val = worksheet.DataValidations.AddListValidation(worksheet.Cells[7, 8].Address);//设置下拉框显示的数据区域
                val.Formula.ExcelFormula = "=parameter";//数据区域的名称

                val.PromptTitle = "数值";//设置下拉框提示的标题
                val.Prompt = "下拉选择参数";//下拉框提示
                val.ShowInputMessage = true;//是否显示提示内容
                val.Formula.Values.Add("123");//为下拉框赋值(下拉框的赋值使用的方法，是使用list集合添加的)
                val.Formula.Values.Add("465");
                val.Formula.Values.Add("789");
                #endregion

                #region 嵌入vba代码(VBA代码不熟悉，因此没有实例)
                //创建工程对象，方法一：
                //package.Workbook.CreateVBAProject();//创建工程对象
                //worksheet.CodeModule.Name = "Name";

                //创建工程对象，方法二：
                //OfficeOpenXml.VBA.ExcelVbaProject proj = package.Workbook.VbaProject;
                //OfficeOpenXml.VBA.ExcelVBAModule sheetModule = proj.Modules["Name"];
                //worksheet.CodeModule.Code = File.ReadAllText(@"D:\ExcelTest\vba.txt", Encoding.Default);
                #endregion

                #region 其他
                //？？这个属性没能看出是什么作用，在社区API中也只是说明该属性是布尔类型，单没有解释
                //不过，这个属性的true和false并没有体现出什么影响，所以这个权当看过就行。
                //当让，若是看客中有人知道，请一定要在评论留言中说明，多谢。
                worksheet.Cells.Style.WrapText = true;

                int sheetCount = package.Workbook.Worksheets.Count;//获取总的sheet数量

                ExcelWorksheet worksheetOther = package.Workbook.Worksheets[1];//选定指定sheet页

                int maxColumnNum = worksheet.Dimension.End.Column;//设置最大列数
                int minColumnNum = worksheet.Dimension.Start.Column;//设置最小列数
                int maxRowNum = worksheet.Dimension.End.Row;//设置最大行数
                int minRowNum = worksheet.Dimension.Start.Row;//设置最小行数

                var range = worksheet.Cells[1, 1, 4, 4];//获取某一个区域

                //ExcelPackage ep = new ExcelPackage(file, OpenPassword)带参构造函数，创建一个需要密码的Excel文件
                //ep.Save(OpenPassword)保存时，用这行语句保存
                #endregion

                //保存Excel文件
                package.Save();
            }
        }
        #endregion

        #region 导出附有数据图的Excel文件
        public static void GenerateExcelReportColumnClustered()
        {
            string fileName = DateTime.Now.ToString("yyyy_MM_dd_HHmmss") + ".xlsx";
            FileInfo file = new FileInfo("D:\\ExcelTest\\" + fileName);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("ColumnClustered");
                worksheet.Cells.Style.WrapText = true;

                #region 样本数据
                worksheet.Cells[1, 1].Value = "名称";
                worksheet.Cells[1, 2].Value = "价格";
                worksheet.Cells[1, 3].Value = "销量";

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
                #endregion

                #region 样本数据格式
                using (ExcelRange range = worksheet.Cells[1, 1, 5, 3])
                {
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }

                using (ExcelRange range = worksheet.Cells[1, 1, 1, 3])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Font.Color.SetColor(Color.White);
                    range.Style.Font.Name = "微软雅黑";
                    range.Style.Font.Size = 12;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(128, 128, 128));
                }



                worksheet.Cells[1, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet.Cells[1, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet.Cells[1, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));

                worksheet.Cells[2, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet.Cells[2, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet.Cells[2, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));

                worksheet.Cells[3, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet.Cells[3, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet.Cells[3, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));

                worksheet.Cells[4, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet.Cells[4, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet.Cells[4, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));

                worksheet.Cells[5, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet.Cells[5, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet.Cells[5, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                #endregion

                #region 设置数据图
                //创建图表样式,并将图表命名为"chart"
                ExcelChart chart = worksheet.Drawings.AddChart("chartColumnClustered", eChartType.ColumnClustered);//柱形图
                //ExcelChart chart = worksheet.Drawings.AddChart("chart", eChartType.Line);//折线图
                //ExcelChart chart = worksheet.Drawings.AddChart("chart", eChartType.Pie3D);//扇形图

                #region 若需要一个Excel文件中导出多个数据图，撤销下面的这段代码就好了
                //ExcelChart chart2 = worksheet.Drawings.AddChart("chartPie3D", eChartType.Pie3D);//扇形图
                //ExcelChartSerie serie11 = chart2.Series.Add(worksheet.Cells[2, 3, 5, 3], worksheet.Cells[2, 1, 5, 1]);
                //serie11.HeaderAddress = worksheet.Cells[1, 3];
                //ExcelChartSerie serie23 = chart2.Series.Add(worksheet.Cells[2, 2, 5, 2], worksheet.Cells[2, 1, 5, 1]);
                //serie23.HeaderAddress = worksheet.Cells[1, 2];
                //chart2.SetPosition(150, 400);//设置位置，top:150px; left:10px
                //chart2.SetSize(300, 300);//设置大小，length:300px; height:300px
                //chart2.Title.Text = "销量走势";//设置图表的标题
                //chart2.Title.Font.Color = Color.FromArgb(89, 89, 89);//设置标题的颜色
                //chart2.Title.Font.Size = 15;//标题的大小
                //chart2.Title.Font.Bold = true;//标题的粗体
                //chart2.Style = eChartStyle.Style15;//设置图表的样式
                //chart2.Legend.Border.LineStyle = eLineStyle.Solid;
                //chart2.Legend.Border.Fill.Color = Color.FromArgb(217, 217, 217);//设置图例的颜色

                //ExcelChart chart3 = worksheet.Drawings.AddChart("chartLine", eChartType.Line);//折线图
                //ExcelChartSerie serie111 = chart3.Series.Add(worksheet.Cells[2, 3, 5, 3], worksheet.Cells[2, 1, 5, 1]);
                //serie111.HeaderAddress = worksheet.Cells[1, 3];
                //ExcelChartSerie serie3 = chart3.Series.Add(worksheet.Cells[2, 2, 5, 2], worksheet.Cells[2, 1, 5, 1]);
                //serie3.HeaderAddress = worksheet.Cells[1, 2];
                //chart3.SetPosition(150, 810);//设置位置，top:150px; left:10px
                //chart3.SetSize(300, 300);//设置大小，length:500px; height:300px
                //chart3.Title.Text = "销量走势";//设置图表的标题
                //chart3.Title.Font.Color = Color.FromArgb(89, 89, 89);//设置标题的颜色
                //chart3.Title.Font.Size = 15;//标题的大小
                //chart3.Title.Font.Bold = true;//标题的粗体
                //chart3.Style = eChartStyle.Style15;//设置图表的样式
                //chart3.Legend.Border.LineStyle = eLineStyle.Solid;
                //chart3.Legend.Border.Fill.Color = Color.FromArgb(217, 217, 217);//设置图例的颜色
                #endregion

                //选择数据
                //chart.Series.Add()方法所需参数为：chart.Series.Add(Y轴数据区,X轴数据区) 
                ExcelChartSerie serie = chart.Series.Add(worksheet.Cells[2, 3, 5, 3], worksheet.Cells[2, 1, 5, 1]);
                serie.HeaderAddress = worksheet.Cells[1, 3];//设置当前柱形图对象代表的主题
                ExcelChartSerie serie2 = chart.Series.Add(worksheet.Cells[2, 2, 5, 2], worksheet.Cells[2, 1, 5, 1]);
                serie2.HeaderAddress = worksheet.Cells[1, 2];

                //设置图表样式
                chart.SetPosition(150, 10);//设置位置，top:150px; left:10px
                chart.SetSize(300, 300);//设置大小，length:300px; height:300px
                chart.Title.Text = "销量走势";//设置图表的标题
                chart.Title.Font.Color = Color.FromArgb(89, 89, 89);//设置标题的颜色
                chart.Title.Font.Size = 15;//标题的大小
                chart.Title.Font.Bold = true;//标题的粗体
                chart.Style = eChartStyle.Style15;//设置图表的样式
                chart.Legend.Border.LineStyle = eLineStyle.Solid;
                chart.Legend.Border.Fill.Color = Color.FromArgb(217, 217, 217);//设置图例的颜色
                #endregion

                package.Save();//保存文件  
            }
        }
        #endregion

        #region DataTable导出为Excel文件
        public static void DataTableToExcel(DataTable InputDataTable, string OutputPath, string SheetName, bool IsInNeedFirstRow)
        {
            try
            {
                //创建文件IO流对象
                FileInfo fileInfo = new FileInfo(OutputPath);

                //创建Excel工程对象
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    //创建一个指定名称的sheet页对象
                    ExcelWorksheet workSheet = package.Workbook.Worksheets.Add(SheetName);

                    //将DataTable的数据填充到Excel文件(包含列名)
                    //workSheet.Cells[]索引中，用来指定数据填充开始的位置
                    workSheet.Cells[1, 1].LoadFromDataTable(InputDataTable, true);

                    //设置列宽
                    workSheet.Column(4).Width = 14;
                    workSheet.Column(5).Width = 14;

                    //设置表格样式
                    for (int i = 0; i <= InputDataTable.Rows.Count; i++)
                    {
                        for (int j = 1; j <= InputDataTable.Columns.Count; j++)
                        {
                            if (i == 0)
                            {
                                //设置Excel文件第一行的单元格背景色
                                workSheet.Cells[1, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                workSheet.Cells[1, j].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(128, 128, 128));
                                workSheet.Cells[1, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                                workSheet.Cells[1, j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            }
                            //设置Excel单元格中有数据的单元格边框
                            workSheet.Cells[i + 1, j].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(0, 0, 0));
                        }
                    }

                    //合并第五列相同内容的单元格
                    int startColumn = 1;
                    for (int i = 2; i <= workSheet.Dimension.End.Row; i++)
                    {
                        if (workSheet.Cells[i, 5].Value.ToString() == workSheet.Cells[startColumn, 5].Value.ToString())
                        {
                            if (i == workSheet.Dimension.End.Row)
                            {
                                workSheet.Cells[startColumn, 5, i, 5].Merge = true;
                                workSheet.Cells[startColumn, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                                workSheet.Cells[startColumn, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                            }
                        }
                        else
                        {
                            workSheet.Cells[startColumn, 5, i - 1, 5].Merge = true;
                            startColumn = i;
                        }
                    }

                    //根据是否需要第一行，删除第一行数据
                    if (!IsInNeedFirstRow)
                        workSheet.DeleteRow(1);

                    //根据DataTable第一行的列类型，设置Excel文件中对应的类型
                    for (int j = 0; j < InputDataTable.Columns.Count; j++)
                    {
                        string dtcolumntype = InputDataTable.Columns[j].DataType.Name.ToLower();
                        switch (dtcolumntype)
                        {
                            case "datetime":
                                workSheet.Column(j + 1).Style.Numberformat.Format = "yyyy/m/d";
                                break;
                            case "single":
                                workSheet.Column(j + 1).Style.Numberformat.Format = "#,##0.00";
                                break;
                            case "int":
                            case "int16":
                            case "int32":
                            case "int64":
                                workSheet.Column(j + 1).Style.Numberformat.Format = "#,##0";
                                break;
                            default:
                                workSheet.Column(j + 1).Style.Numberformat.Format = "@";
                                break;
                        }
                    }
                    workSheet.InsertRow(1, 1);
                    //保存Excel文件
                    package.Save();

                    //释放资源
                    workSheet.Dispose();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
        }
        #endregion

        #region 读取Excel文件到DataTable
        public static DataTable ExcelToDataTable(string InputPath, string sheetName, bool firstRowAsColumnName)
        {
            try
            {
                //创建文件IO流对象
                FileInfo fileInfo = new FileInfo(InputPath);

                //创建Excel工程对象
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    //创建一个指定名称的sheet页对象
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

                    int maxColumnNum = worksheet.Dimension.End.Column;//最大列数
                    int maxRowNum = worksheet.Dimension.End.Row;//最小行数

                    //创建目标表
                    DataTable targetTable = new DataTable();
                    DataColumn dtColumn;

                    //是否把Excel文件的第一行作为目标表名称
                    int startRowNum = 1;//根据是否把Excel文件第一行作为表名称，来确定填充数据时开始的行数
                    if (firstRowAsColumnName)
                    {
                        startRowNum = 2;
                        for (int i = 1; i <= maxColumnNum; i++)
                        {
                            dtColumn = new DataColumn(worksheet.Cells[1, i].Value.ToString(), typeof(string));
                            targetTable.Columns.Add(dtColumn);
                        }
                    }
                    else
                    {
                        for (int j = 1; j <= maxColumnNum; j++)
                        {
                            dtColumn = new DataColumn("Column_" + j, typeof(string));
                            targetTable.Columns.Add(dtColumn);
                        }
                    }

                    //为目标表填充Excel文件的数据
                    for (int i = startRowNum; i <= maxRowNum; i++)
                    {
                        DataRow vRow = targetTable.NewRow();
                        for (int m = 1; m <= maxColumnNum; m++)
                        {
                            vRow[m - 1] = worksheet.Cells[i, m].Value;
                        }
                        targetTable.Rows.Add(vRow);
                    }

                    //代码编辑到这里，有一个问题：NPOI导入到DataTable，希望DataTable中的数据时Excel文件执行公式后的最终结果集，
                    //可以使用worksheet.ForceFormulaRecalculation = true这句代码，来激活Excel文件中的公式
                    //但是，很遗憾UP主没能找到在EPplus中如何激活Excel文件里的公式。
                    //若有看客找到了这个问题的解决方法，请一定要评论区留言说明！！！！非常感谢。

                    //OfficeOpenXml.ExcelCalcMode.Automatic = true;

                    //释放资源
                    worksheet.Dispose();

                    //返回目标表
                    return targetTable;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
        }
        #endregion
    }
}
