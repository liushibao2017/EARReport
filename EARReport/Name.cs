using Nest;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using THC.FinancialSimulation;

namespace EARReport
{
    
    #region 树value枚举项
    public enum EaRLineItemEnum
    {
        Balance = 1,
        Book = 2,
        Market = 3,
        Accounting = 4,
        Interest = 5,
        RemainingBalance = 6,
        Amortization = 7,
        Prepayment = 8,
        Recovery = 9,
        Loss = 10,
        NIC = 11,
        Reinvested = 12,
        ReinvestingRate = 13,
        Withdrawal=14,
        PB=15,
        NA=16,
        Offer=17,
    }
    #endregion

    #region Ear report 树布局类

    public class EaRLayoutSetting
    {
        public string COALineName { get; set; }
        public Dictionary<EaRLineItemEnum, string> Lines { get; set; }

    } 
    #endregion
}
