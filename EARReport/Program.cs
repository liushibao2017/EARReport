using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Text;
using System.Threading.Tasks;
using EPPlusSamples;
using System.Reflection;
using THC.FinancialSimulation;
using System.Drawing;
using OfficeOpenXml.Style;
using Newtonsoft.Json;

namespace EARReport
{
    class Program
    {

        List<EaRLayoutSetting> settings = new List<EaRLayoutSetting>();


        /// <summary>
        /// 读取json文件
        /// </summary>
        /// <param name="filename">json文件名</param>
        /// <returns></returns>
        FinProjection ReadFromJson(string filename)
        {
            FinProjection finsim = null;
            using (StreamReader reader = File.OpenText(filename))
            {
                JsonSerializer serializer = JsonSerializer.Create();
                finsim = serializer.Deserialize(reader, typeof(FinProjection)) as FinProjection;
            }
            return finsim;

        }
       
    
        /// <summary>
        /// 树的写入
        /// </summary>
        /// <param name="sht"></param>
        /// <param name="finsim"></param>
        /// <param name="irow"></param>
        /// <returns></returns>
        protected int WriteCOAData(ExcelWorksheet sht, COAFinSim finsim, List<EaRLayoutSetting> settings, int irow)
        {

            irow = WriteCOALines(sht, finsim.Name, finsim.Value, settings, irow);
            if (finsim.Childs != null)
            {
                foreach (var v in finsim.Childs)
                {
                    // sht.Cells[irow, 2].Value = v.Value.Name;
                    irow = WriteCOAData(sht, v.Value, settings, irow);
                }
            }
            return irow;
        }
        /// <summary>
        /// 树value写入
        /// </summary>
        /// <param name="sht"></param>
        /// <param name="lines"></param>
        /// <param name="irow"></param>
        /// <returns></returns>
        protected int WriteCOALines(ExcelWorksheet sht, string coaname, FinSimCOALine lines, List<EaRLayoutSetting> settings, int irow)
        {
            if (lines == null) return irow;
            sht.Cells[irow, 2].Value = coaname;

            Dictionary<EaRLineItemEnum, string> coalines = null;
            foreach (var v in settings)
            {
                if (coaname == v.COALineName)
                {
                    coalines = v.Lines;
                    break;
                }
            }
            if (coalines != null)
            {
                Dictionary<EaRLineItemEnum, string>.Enumerator en = coalines.GetEnumerator();
                for (int j = 0; j < coalines.Count; j++)
                {
                    if (en.MoveNext())
                    {
                        sht.Cells[irow + 1 + j, 2].Value = "  " + en.Current.Value;
                    }
                    for (int i = 0; i <= 12; i++)
                    {
                        sht.Cells[irow + 1 + j, i + 3].Value = lines.CashFlows[i].RemainingBalance;
                    }
                }
            }
            return irow + 10;
        }
       
        void PrepareSampleEaRLayoutSettings()
        {
            EaRLayoutSetting asssets = new EaRLayoutSetting();
            asssets.COALineName = "ASSETS";
            Dictionary<EaRLineItemEnum, string> assetslines = new Dictionary<EaRLineItemEnum, string>
            {
                { EaRLineItemEnum.Accounting, "Book/Market" },
                { EaRLineItemEnum.Interest, "Interests" }
            };
            asssets.Lines = assetslines;
            settings.Add(asssets);

            EaRLayoutSetting federal = new EaRLayoutSetting();
            federal.COALineName = "Federal";
            Dictionary<EaRLineItemEnum, string> federallines = new Dictionary<EaRLineItemEnum, string>
            {
                { EaRLineItemEnum.Balance, "Balance" },
                { EaRLineItemEnum.Interest, "Interest" },
                { EaRLineItemEnum.ReinvestingRate, "WAC(%)" }
            };
            federal.Lines = federallines;
            settings.Add(federal);

            EaRLayoutSetting loan = new EaRLayoutSetting();
            loan.COALineName = "LOAN";
            Dictionary<EaRLineItemEnum, string> loanlines = new Dictionary<EaRLineItemEnum, string>
            {
                { EaRLineItemEnum.Book, "Book" },
                { EaRLineItemEnum.Prepayment, "Prem/Disc Amort" },
                { EaRLineItemEnum.Balance, "Balance" },
                { EaRLineItemEnum.Interest, "Interest" },
                { EaRLineItemEnum.NIC, "Non-Interest Cost" },
                { EaRLineItemEnum.Amortization, "Prin.Amort" },
                { EaRLineItemEnum.Recovery, "Prin.Recovery" }
            };
            loan.Lines = loanlines;
            settings.Add(loan);

            EaRLayoutSetting investment = new EaRLayoutSetting();
            investment.COALineName = "INVESTMENT";
            Dictionary<EaRLineItemEnum, string> investmentlines = new Dictionary<EaRLineItemEnum, string>
            {
                { EaRLineItemEnum.Book, "Book" },
                { EaRLineItemEnum.Amortization, "Prem/Disc Amort" },
                { EaRLineItemEnum.Prepayment, "Prin.Prepay" },
                { EaRLineItemEnum.Recovery, "Prin.Recovery" }
            };
            investment.Lines = investmentlines;
            settings.Add(investment);

            EaRLayoutSetting other = new EaRLayoutSetting();
            other.COALineName = "Other-Asset";
            Dictionary<EaRLineItemEnum, string> otherlines = new Dictionary<EaRLineItemEnum, string>();
            otherlines.Add(EaRLineItemEnum.Balance, "Balance");
            otherlines.Add(EaRLineItemEnum.Interest, "Interest");
            other.Lines = otherlines;
            settings.Add(other);

            EaRLayoutSetting li = new EaRLayoutSetting();
            li.COALineName = "LIABILITIES";
            Dictionary<EaRLineItemEnum, string> lilines = new Dictionary<EaRLineItemEnum, string>
            {
                { EaRLineItemEnum.Accounting, "Book/Market" },
                { EaRLineItemEnum.Interest, "Interests" }
            };
            li.Lines = lilines;
            settings.Add(li);

            EaRLayoutSetting cd = new EaRLayoutSetting();
            cd.COALineName = "CD";
            Dictionary<EaRLineItemEnum, string> cdlines = new Dictionary<EaRLineItemEnum, string>
            {
                { EaRLineItemEnum.Balance, "Balance" },
                { EaRLineItemEnum.Interest, "Interest" },
                { EaRLineItemEnum.NIC, "Non-Interest Cost" },
                { EaRLineItemEnum.ReinvestingRate, "Implied Rate(%)" },
                { EaRLineItemEnum.Market, "Matured CD" },
                { EaRLineItemEnum.Withdrawal, "Withdrawal" },
                { EaRLineItemEnum.PB, "Perf.Bal" },
                { EaRLineItemEnum.NA, "New Account" },
                { EaRLineItemEnum.Offer, "Offer-Rate of new CD Account(%)" }
            };
            cd.Lines = cdlines;
            settings.Add(cd);
        }

        static void Main(string[] args)
        {
            FileInfo newFile = new FileInfo(@"EAR.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(@"EAR.xlsx");
            }
            string dir = Path.GetFullPath("../..");
            string filePath = dir + "\\THC NII Report Template.xlsx";
            FileInfo templateFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                //EaR summary
                ExcelWorksheet shtES = package.Workbook.Worksheets["EaR summary"];
                shtES.Cells[11, 4].Value = "Dn 200BP";
                shtES.Cells[11, 5].Value = "Dn 100BP";
                shtES.Cells[11, 6].Value = "Base";
                shtES.Cells[11, 7].Value = "Up 100BP";
                shtES.Cells[11, 8].Value = "Up 200BP";
                shtES.Cells[11, 9].Value = "Up 300BP";
                shtES.Cells[11, 10].Value = "Up 400BP ";
                shtES.Cells[11, 11].Value = "Flattener";
                shtES.Cells[11, 12].Value = "Ramp Up";
                shtES.Cells[21, 4].Value = "Dn 200BP";
                shtES.Cells[21, 5].Value = "Dn 100BP";
                shtES.Cells[21, 6].Value = "Base";
                shtES.Cells[21, 7].Value = "Up 100BP";
                shtES.Cells[21, 8].Value = "Up 200BP";
                shtES.Cells[21, 9].Value = "Up 300BP";
                shtES.Cells[21, 10].Value = "Up 400BP ";
                shtES.Cells[21, 11].Value = "Flattener";
                shtES.Cells[21, 12].Value = "Ramp Up";
                //1st year projection
                ExcelWorksheet sht1year = package.Workbook.Worksheets["1st year projection"];
                sht1year.Cells.Style.Font.Name = "Calibri";
                sht1year.Cells.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                Program pro = new Program();
                FinProjection finProjection = pro.ReadFromJson("azf201906.json");
                sht1year.Cells[7, 2].Value = "Scenario:Base Case";
                sht1year.Cells[8, 2].Value = "ASSETS INTEREST INCOME";
                sht1year.Cells[9, 2].Value = "LIABILITIES INTEREST COST";
                sht1year.Cells[10, 2].Value = "NET INTEREST INCOME";
                for (int i = 1; i <= 12; i++)
                {
                    sht1year.Cells[8, i + 3].Value = finProjection.TotalLines.NetIncome.InterestIncome[i];
                    sht1year.Cells[9, i + 3].Value = finProjection.TotalLines.NetIncome.InterestCost[i];
                    sht1year.Cells[10, i + 3].Value = finProjection.TotalLines.NetIncome.NetInterestIncome[i];
                    double dblNIC = finProjection.TotalLines.NetIncome.NonInterestExpense.Value[i];
                    double dblNII = finProjection.TotalLines.NetIncome.NonInterestIncome.Value[i];
                    sht1year.Cells[11, i + 3].Value = dblNIC - dblNII;
                    sht1year.Cells[12, i + 3].Value = finProjection.TotalLines.NetIncome.LoanLossProvision[i];
                    sht1year.Cells[14, i + 3].Value = finProjection.TotalLines.NetIncome.TaxPayments[i];
                    sht1year.Cells[15, i + 3].Value = finProjection.TotalLines.NetIncome.NI[i];
                    sht1year.Cells[16, i + 3].Value = finProjection.TotalLines.NetIncome.DividendPayment[i];
                    sht1year.Cells[18, i + 3].Value = finProjection.TotalLines.NetIncome.THCNetChangeUnRealizedGain[i];
                }

                sht1year.Cells[11, 2].Value = "Non Interest Expense(income)";
                sht1year.Cells[12, 2].Value = "Provision of losses";
                sht1year.Cells[13, 2].Value = "Profit before taxes";
                sht1year.Cells[14, 2].Value = "Tax";

                sht1year.Cells[15, 2].Value = "Net Income";
                sht1year.Cells[16, 2].Value = "Dividend Payment";
                sht1year.Cells[17, 2].Value = "Retained Earning chg";
                sht1year.Cells[18, 2].Value = "Unrealized G/L";
                sht1year.Cells[6, 2].Value = "Date";
                using (ExcelRange range = sht1year.Cells[7, 2, 10, 2])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Font.Name = "微软雅黑";
                    sht1year.Cells[7, 2].Value = "Scenario:Base Case";
                    sht1year.Cells[8, 2].Value = "ASSETS INTEREST INCOME";
                    sht1year.Cells[9, 2].Value = "LIABILITIES INTEREST COST";
                    sht1year.Cells[10, 2].Value = "NET INTEREST INCOME";
                }
                using (ExcelRange range = sht1year.Cells["B7"])
                {
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(128, 128, 128));
                }
                sht1year.Cells[5, 15].Value = "Currency: USD .Amounts in 000s";
                using (ExcelRange range = sht1year.Cells[11, 2, 18, 2])
                {
                    range.Style.Font.Italic = true;
                    range.Style.Font.Name = "Calibri";
                    sht1year.Cells[11, 2].Value = "Non Interest Expense(income)";
                    sht1year.Cells[12, 2].Value = "Provision of losses";
                    sht1year.Cells[13, 2].Value = "Profit before taxes";
                    sht1year.Cells[14, 2].Value = "Tax";
                    sht1year.Cells[15, 2].Value = "Net Income";
                    sht1year.Cells[16, 2].Value = "Dividend Payment";
                    sht1year.Cells[17, 2].Value = "Retained Earning chg";
                    sht1year.Cells[18, 2].Value = "Unrealized G/L";

                }
                using (ExcelRange range = sht1year.Cells[14, 2, 14, 15])
                {
                    range.Style.Font.UnderLine = true;
                }
                using (ExcelRange range = sht1year.Cells[18, 2, 18, 15])
                {
                    range.Style.Font.UnderLine = true;
                }
                sht1year.Cells[19, 2].Value = "Equity";
                sht1year.Cells[19, 2].Style.Font.Name = "Calibri";
                sht1year.Cells[19, 2].Style.Font.Bold = true;
                sht1year.Cells[19, 2].Style.Font.Italic = true;
                sht1year.Cells[19, 2].Style.Font.Size = 11;
                pro.PrepareSampleEaRLayoutSettings();
                pro.WriteCOAData(sht1year, finProjection.COA, pro.settings, 21);//写入树的相关数据
                package.SaveAs(newFile);
            }

        }
    }
}



